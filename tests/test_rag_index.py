"""Tests für rag_index.py – Store-Dateien, Embedding-Aufrufe (gemockt) und inkrementeller Aufbau.

Es wird nie echtes Ollama angesprochen: requests.post bzw. rag_index.embed
werden durch deterministische Fakes ersetzt.
"""

import sys
import json
import sqlite3
import hashlib

import numpy as np
import pytest
import requests

import corpus
import rag_index


# --------------------------------------------------------------------------
# Hilfen: gefälschte Ollama-Antworten und deterministische Vektoren
# --------------------------------------------------------------------------
DIM = 8


class FakeResp:
    """Minimaler Ersatz für requests.Response."""

    def __init__(self, payload, status=200):
        self.status_code = status
        self._payload = payload

    def raise_for_status(self):
        if self.status_code >= 400:
            raise requests.exceptions.HTTPError(f"HTTP {self.status_code}")

    def json(self):
        return self._payload


def det_vec(text):
    """Deterministischer, textabhängiger Pseudo-Embedding-Vektor (nie Nullvektor)."""
    digest = hashlib.sha1(text.encode("utf-8")).digest()
    return [b / 255.0 + 0.01 for b in digest[:DIM]]


def fake_embed_factory(calls):
    """Ersatz für rag_index.embed, der jeden Batch in `calls` mitschreibt."""
    def fake_embed(texts, model, url, timeout=600):
        calls.append(list(texts))
        return [det_vec(t) for t in texts]
    return fake_embed


def make_chunk(uid="outlook:inbox/mail.eml:0", seq=0, **kw):
    """Chunk-Dict, wie es corpus.chunk_records + chunk_hash liefern würden."""
    c = {"uid": uid, "cid": f"{uid}#{seq}", "src": "outlook", "root": "outlook",
         "rel": "inbox/mail.eml", "who": "Alice Example", "ppl": "alice example",
         "ts": 1751875200.0, "date": "2025-07-07 10:00", "title": "Testmail",
         "ctx": "inbox", "text": "Inhalt"}
    c.update(kw)
    c.setdefault("hash", corpus.chunk_hash(c))
    return c


# --------------------------------------------------------------------------
# _chunk_row / _people_rows
# --------------------------------------------------------------------------
def test_chunk_row_zerlegt_cid_und_uid():
    c = make_chunk(uid="teams:1on1/a.html:7", seq=2, src="teams", root="teams",
                   rel="1on1/a.html")
    row = rag_index._chunk_row(4, c)
    assert row[0] == 5                                  # id = Index + 1
    assert row[1] == "teams:1on1/a.html:7"
    assert row[2] == 2                                  # seq aus cid "...#2"
    assert row[3] == 7                                  # msg_idx aus uid "...:7"
    assert row[4:7] == ("teams", "teams", "1on1/a.html")
    assert row[14] == c["hash"]


def test_chunk_row_nicht_numerischer_nachrichtenindex_faellt_auf_null():
    c = make_chunk(uid="outlook:inbox/mail.eml:x", seq=0)
    assert rag_index._chunk_row(0, c)[3] == 0


def test_people_rows_zaehlt_jede_nachricht_nur_einmal():
    chunks = [
        make_chunk(uid="u:1", seq=0, src="teams", who="Alice", ppl="alice alpha"),
        make_chunk(uid="u:1", seq=1, src="teams", who="Alice", ppl="alice alpha"),
        make_chunk(uid="u:2", seq=0, src="teams", who="Alice", ppl="alice beta"),
        make_chunk(uid="u:3", seq=0, src="outlook", who="Bob ", ppl="bob"),
    ]
    rows = {(src, who): (cnt, ppl)
            for src, who, cnt, ppl in rag_index._people_rows(chunks)}
    assert rows[("teams", "Alice")][0] == 2             # Folge-Chunk #1 zählt nicht extra
    assert rows[("teams", "Alice")][1] == "alice alpha beta"
    assert rows[("outlook", "Bob")] == (1, "bob")       # who wird getrimmt


def test_people_rows_begrenzt_personen_tokens():
    viele = " ".join(f"t{i:03d}" for i in range(100))
    chunks = [
        make_chunk(uid="u:1", seq=0, who="Alice", ppl=viele),
        make_chunk(uid="u:2", seq=0, who="Alice", ppl="zzz"),
    ]
    rows = rag_index._people_rows(chunks)
    assert len(rows) == 1
    _, _, cnt, ppl = rows[0]
    assert cnt == 2
    toks = ppl.split()
    assert len(toks) == rag_index.PPL_TOKEN_CAP         # nur die ersten 60 Tokens
    assert "zzz" not in toks                            # Kappe erreicht → nichts mehr dazu


def test_people_rows_ohne_who_und_ppl():
    rows = rag_index._people_rows([make_chunk(who=None, ppl=None)])
    assert rows == [("outlook", "", 1, "")]


# --------------------------------------------------------------------------
# Store schreiben: corpus.db, vectors.npy, info.json
# --------------------------------------------------------------------------
def test_write_db_schreibt_schema_und_inhalte(tmp_path):
    chunks = [
        make_chunk(uid="teams:1on1/a.html:0", seq=0, src="teams", root="teams",
                   rel="1on1/a.html", who="Alice", title="Projekt Alpha",
                   text="Bericht ist fertig."),
        make_chunk(uid="teams:1on1/a.html:0", seq=1, src="teams", root="teams",
                   rel="1on1/a.html", who="Alice", title="Projekt Alpha",
                   text="Zweiter Teil."),
        make_chunk(text="hier die neue Nachricht."),
    ]
    rag_index.write_db(tmp_path, chunks)
    assert (tmp_path / "corpus.db").exists()
    assert not (tmp_path / "corpus.db.tmp").exists()    # atomarer Tausch

    con = sqlite3.connect(tmp_path / "corpus.db")
    rows = list(con.execute(
        "SELECT id, uid, seq, msg_idx, src, text, hash FROM chunks ORDER BY id"))
    assert [r[0] for r in rows] == [1, 2, 3]
    assert rows[0][1:5] == ("teams:1on1/a.html:0", 0, 0, "teams")
    assert rows[1][2] == 1                              # zweiter Chunk derselben Nachricht
    assert rows[2][5] == "hier die neue Nachricht."
    assert all(r[6] for r in rows)                      # Hashes gespeichert

    people = set(con.execute("SELECT src, who, messages FROM people"))
    assert people == {("teams", "Alice", 1), ("outlook", "Alice Example", 1)}

    # FTS5-Volltext: Text und Titel sind durchsuchbar, rowid == chunks.id
    hit = [r[0] for r in con.execute(
        "SELECT rowid FROM chunks_fts WHERE chunks_fts MATCH 'bericht'")]
    assert hit == [1]
    hit = [r[0] for r in con.execute(
        "SELECT rowid FROM chunks_fts WHERE chunks_fts MATCH 'testmail'")]
    assert hit == [3]
    con.close()


def test_write_db_ersetzt_vorhandene_db(tmp_path):
    rag_index.write_db(tmp_path, [make_chunk(text="alt")])
    rag_index.write_db(tmp_path, [make_chunk(text="neu"),
                                  make_chunk(uid="u:1", text="zwei")])
    con = sqlite3.connect(tmp_path / "corpus.db")
    assert {r[0] for r in con.execute("SELECT text FROM chunks")} == {"neu", "zwei"}
    con.close()


def test_save_vectors_normalisiert_und_speichert_float16(tmp_path):
    V = np.array([[3.0, 4.0], [0.0, 0.0]], dtype="float32")
    out = rag_index.save_vectors(tmp_path, V)
    stored = np.load(tmp_path / "vectors.npy")
    assert stored.dtype == np.float16
    assert np.allclose(stored[0].astype("float32"), [0.6, 0.8], atol=1e-3)
    assert np.all(stored[1] == 0)                       # Nullvektor: keine Division durch 0
    assert np.array_equal(out, stored)                  # Rückgabe == gespeicherte Matrix
    assert not (tmp_path / "vectors.npy.tmp").exists()


def test_write_info(tmp_path):
    rag_index.write_info(tmp_path, "bge-m3", np.int64(1024), 7)
    info = json.loads((tmp_path / "info.json").read_text(encoding="utf-8"))
    assert info == {"model": "bge-m3", "dim": 1024, "chunks": 7,
                    "dtype": "float16", "format": rag_index.FORMAT}


# --------------------------------------------------------------------------
# Alten Store lesen (inkrementelle Läufe)
# --------------------------------------------------------------------------
def test_load_old_store_und_vectors_roundtrip(tmp_path):
    chunks = [make_chunk(uid="u:0", text="eins"), make_chunk(uid="u:1", text="zwei")]
    rag_index.write_db(tmp_path, chunks)
    rag_index.save_vectors(tmp_path, np.array([[1.0, 0.0], [0.0, 2.0]], dtype="float32"))

    hashes, V = rag_index._load_old_store(tmp_path)
    assert hashes == [c["hash"] for c in chunks]        # in id-Reihenfolge
    assert V.shape == (2, 2)

    old = rag_index.load_old_vectors(tmp_path)
    assert set(old) == {chunks[0]["hash"], chunks[1]["hash"]}
    assert np.allclose(np.asarray(old[chunks[0]["hash"]], dtype="float32"),
                       [1.0, 0.0], atol=1e-3)


def test_load_old_store_leerer_ordner(tmp_path):
    assert rag_index._load_old_store(tmp_path) == ([], None)
    assert rag_index.load_old_vectors(tmp_path) == {}


def test_load_old_vectors_ignoriert_hashes_ohne_vektorzeile(tmp_path):
    chunks = [make_chunk(uid="u:0", text="eins"), make_chunk(uid="u:1", text="zwei")]
    rag_index.write_db(tmp_path, chunks)
    rag_index.save_vectors(tmp_path, np.array([[1.0, 0.0]], dtype="float32"))
    assert set(rag_index.load_old_vectors(tmp_path)) == {chunks[0]["hash"]}


def test_load_old_vectors_unlesbarer_store(tmp_path, capsys):
    (tmp_path / "corpus.db").write_bytes(b"kein sqlite")
    rag_index.save_vectors(tmp_path, np.ones((1, 2), dtype="float32"))
    assert rag_index.load_old_vectors(tmp_path) == {}
    assert "komplett neu" in capsys.readouterr().out


# --------------------------------------------------------------------------
# embed() – Ollama-Aufrufe (requests.post gemockt)
# --------------------------------------------------------------------------
def test_embed_sendet_batch_und_liefert_embeddings(monkeypatch):
    seen = {}

    def fake_post(url, json=None, timeout=None):
        seen["url"], seen["json"], seen["timeout"] = url, json, timeout
        return FakeResp({"embeddings": [[1.0, 2.0], [3.0, 4.0]]})

    monkeypatch.setattr(requests, "post", fake_post)
    out = rag_index.embed(["a", "b"], "bge-m3", "http://ollama.test")
    assert out == [[1.0, 2.0], [3.0, 4.0]]
    assert seen["url"] == "http://ollama.test/api/embed"
    assert seen["json"] == {"model": "bge-m3", "input": ["a", "b"]}


def test_embed_akzeptiert_alte_single_form(monkeypatch):
    monkeypatch.setattr(requests, "post",
                        lambda *a, **k: FakeResp({"embedding": [1.0, 2.0]}))
    assert rag_index.embed(["a"], "m", "http://x") == [[1.0, 2.0]]


def test_embed_404_meldet_fehlendes_modell(monkeypatch):
    monkeypatch.setattr(requests, "post", lambda *a, **k: FakeResp({}, status=404))
    with pytest.raises(SystemExit, match="ollama pull bge-m3"):
        rag_index.embed(["a"], "bge-m3", "http://x")


def test_embed_ohne_verbindung_bricht_ab(monkeypatch):
    def fail(*a, **k):
        raise requests.exceptions.ConnectionError("down")

    monkeypatch.setattr(requests, "post", fail)
    with pytest.raises(SystemExit, match="Keine Verbindung zu Ollama"):
        rag_index.embed(["a"], "m", "http://x")


def test_embed_unerwartete_antwort_bricht_ab(monkeypatch):
    monkeypatch.setattr(requests, "post", lambda *a, **k: FakeResp({"foo": 1}))
    with pytest.raises(SystemExit, match="Unerwartete Embedding-Antwort"):
        rag_index.embed(["a"], "m", "http://x")


def test_embed_serverfehler_wird_durchgereicht(monkeypatch):
    monkeypatch.setattr(requests, "post", lambda *a, **k: FakeResp({}, status=500))
    with pytest.raises(requests.exceptions.HTTPError):
        rag_index.embed(["a"], "m", "http://x")


# --------------------------------------------------------------------------
# build_index – Ende-zu-Ende mit Mini-Export (embed gemockt)
# --------------------------------------------------------------------------
TEAMS_HTML = """<html><body>
<h1>Projekt Alpha</h1>
<div class="msg">
  <span class="name">Alice Example</span>
  <span class="time">2025-06-01 09:30</span>
  <div class="body">Bericht ist fertig.</div>
</div>
<div class="msg">
  <span class="name">Bob</span>
  <span class="time">2025-06-01 09:35</span>
  <div class="body">Danke!</div>
</div>
</body></html>"""


def _eml(subject="Testmail", body="hier die neue Nachricht."):
    return (f"From: Alice Example <alice@example.com>\n"
            f"To: Bob Builder <bob@example.com>\n"
            f"Subject: {subject}\n"
            f"Date: Mon, 07 Jul 2025 10:00:00 +0000\n"
            f"Content-Type: text/plain; charset=utf-8\n"
            f"\n{body}\n").encode()


def _make_exports(tmp_path):
    teams = tmp_path / "teams_export"
    (teams / "1on1").mkdir(parents=True)
    (teams / "1on1" / "alice__abc123.html").write_text(TEAMS_HTML, encoding="utf-8")
    outlook = tmp_path / "outlook_export"
    (outlook / "inbox").mkdir(parents=True)
    (outlook / "inbox" / "mail.eml").write_bytes(_eml())


def _build(tmp_path, monkeypatch, **kw):
    calls = []
    monkeypatch.setattr(rag_index, "embed", fake_embed_factory(calls))
    store = tmp_path / "store"
    result = rag_index.build_index(str(tmp_path / "teams_export"),
                                   str(tmp_path / "outlook_export"),
                                   str(store), "test-modell", "http://ollama.test", **kw)
    return result, calls, store


def test_build_index_erzeugt_kompletten_store(tmp_path, monkeypatch):
    _make_exports(tmp_path)
    (n, neu, dim), calls, store = _build(tmp_path, monkeypatch)
    assert (n, neu, dim) == (3, 3, DIM)                 # 2 Teams-Nachrichten + 1 Mail
    assert sum(len(c) for c in calls) == 3

    V = np.load(store / "vectors.npy")
    assert V.dtype == np.float16 and V.shape == (3, DIM)
    norms = np.linalg.norm(V.astype("float32"), axis=1)
    assert np.allclose(norms, 1.0, atol=1e-2)           # Zeilen sind normalisiert

    # Zeile 0 gehört zum ersten Teams-Chunk (chunks.id = 1)
    erwartet = np.asarray(det_vec("Projekt Alpha\nBericht ist fertig."), dtype="float32")
    erwartet /= np.linalg.norm(erwartet)
    assert np.allclose(V[0].astype("float32"), erwartet, atol=1e-2)

    con = sqlite3.connect(store / "corpus.db")
    rows = list(con.execute("SELECT id, src, who, title FROM chunks ORDER BY id"))
    con.close()
    assert [r[0] for r in rows] == [1, 2, 3]
    assert rows[0][1:] == ("teams", "Alice Example", "Projekt Alpha")
    assert rows[2][1:3] == ("outlook", "Alice Example")

    info = json.loads((store / "info.json").read_text(encoding="utf-8"))
    assert info["chunks"] == 3 and info["dim"] == DIM and info["model"] == "test-modell"


def test_build_index_inkrementell_und_nach_aenderung(tmp_path, monkeypatch):
    _make_exports(tmp_path)
    (_, neu1, _), _, store = _build(tmp_path, monkeypatch)
    assert neu1 == 3
    V1 = np.load(store / "vectors.npy").astype("float32")

    # Zweiter Lauf ohne Änderung: nichts wird neu eingebettet
    (_, neu2, _), calls2, _ = _build(tmp_path, monkeypatch)
    assert neu2 == 0
    assert calls2 == []                                 # kein einziger embed-Aufruf
    V2 = np.load(store / "vectors.npy").astype("float32")
    assert np.allclose(V1, V2, atol=1e-3)

    # Eine Mail ändern: nur dieser eine Chunk wird neu eingebettet
    (tmp_path / "outlook_export" / "inbox" / "mail.eml").write_bytes(
        _eml(body="komplett neuer Inhalt."))
    (_, neu3, _), calls3, _ = _build(tmp_path, monkeypatch)
    assert neu3 == 1
    assert sum(len(c) for c in calls3) == 1
    assert calls3[0] == ["Testmail\nkomplett neuer Inhalt."]
    V3 = np.load(store / "vectors.npy").astype("float32")
    assert np.allclose(V3[:2], V1[:2], atol=1e-3)       # Teams-Zeilen wiederverwendet

    con = sqlite3.connect(store / "corpus.db")
    rows = list(con.execute("SELECT text FROM chunks WHERE src = 'outlook'"))
    con.close()
    assert rows == [("komplett neuer Inhalt.",)]


def test_build_index_bettet_identische_texte_nur_einmal_ein(tmp_path, monkeypatch):
    outlook = tmp_path / "outlook_export"
    (outlook / "inbox").mkdir(parents=True)
    (outlook / "inbox" / "a.eml").write_bytes(_eml())
    (outlook / "inbox" / "b.eml").write_bytes(_eml())   # identischer Inhalt, andere Datei
    (n, neu, _), calls, store = _build(tmp_path, monkeypatch)
    assert (n, neu) == (2, 2)
    assert sum(len(c) for c in calls) == 1              # nur ein eindeutiger Text
    V = np.load(store / "vectors.npy")
    assert np.array_equal(V[0], V[1])                   # Vektor auf beide Chunks verteilt


def test_build_index_batcht_embedding_aufrufe(tmp_path, monkeypatch):
    outlook = tmp_path / "outlook_export"
    (outlook / "inbox").mkdir(parents=True)
    for i in range(5):
        (outlook / "inbox" / f"m{i}.eml").write_bytes(
            _eml(subject=f"Mail {i}", body=f"Inhalt {i}."))
    (n, neu, _), calls, _ = _build(tmp_path, monkeypatch, batch=2)
    assert (n, neu) == (5, 5)
    assert sorted(len(c) for c in calls) == [1, 2, 2]   # 5 Texte in Batches zu 2


def test_build_index_ohne_inhalte_bricht_ab(tmp_path):
    with pytest.raises(SystemExit, match="Keine Inhalte"):
        rag_index.build_index(str(tmp_path / "fehlt"), str(tmp_path / "auch_fehlt"),
                              str(tmp_path / "store"), "m", "http://x")


# --------------------------------------------------------------------------
# main() – Argument-Verdrahtung
# --------------------------------------------------------------------------
def test_main_reicht_argumente_an_build_index_weiter(monkeypatch, capsys):
    seen = {}

    def fake_build(teams, outlook, store, model, url, batch):
        seen["args"] = (teams, outlook, store, model, url, batch)
        return 3, 1, 8

    monkeypatch.setattr(rag_index, "build_index", fake_build)
    monkeypatch.setattr(sys, "argv",
                        ["rag_index.py", "t_dir", "o_dir", "--store", "s", "--batch", "7"])
    rag_index.main()
    assert seen["args"] == ("t_dir", "o_dir", "s", rag_index.DEFAULT_MODEL,
                            rag_index.DEFAULT_OLLAMA, 7)
    assert "3 Chunks" in capsys.readouterr().out
