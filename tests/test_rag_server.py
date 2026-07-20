"""Tests für rag_server.py – Retrieval-Logik, Ollama-Anbindung (gemockt) und HTTP-API.

Ollama wird nie angesprochen: requests.post ist überall gemockt. Für die
HTTP-Schicht startet ein eigener Wegwerf-Server auf 127.0.0.1 (ephemerer Port),
gegen den mit http.client echte Requests laufen.
"""

import json
import threading
import http.client
from http.server import ThreadingHTTPServer

import numpy as np
import pytest
import requests

import rag_server


# --------------------------------------------------------------------------
# Hilfen: Meta-Daten, gefälschte Ollama-Antworten
# --------------------------------------------------------------------------
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


def _chunk(**kw):
    base = {"src": "teams", "root": "teams", "rel": "1on1/a.html", "who": "Alice Example",
            "ppl": "alice example bob projekt alpha", "ts": 100.0,
            "date": "2025-06-01 09:30", "title": "Projekt Alpha", "ctx": "1:1-Chat",
            "text": "Bericht ist fertig."}
    base.update(kw)
    return base


def _meta():
    """Vier Chunks über alle Quellen; Vektorzeile i == Einheitsvektor e_i."""
    return [
        _chunk(),
        _chunk(src="outlook", root="outlook", rel="inbox/mail.eml", who="Bob Builder",
               ppl="bob builder bob@example.com", ts=200.0, date="2025-07-07 10:00",
               title="Testmail", ctx="inbox", text="m" * 700),
        _chunk(src="kalender", root="outlook", rel="kalender/Arbeit/t.ics",
               who="Alice Example", ppl="alice example alice@example.com", ts=300.0,
               date="2025-06-01 12:00", title="Planung", ctx="Kalender: Arbeit",
               text="Ort: Raum 42. Agenda folgt"),
        _chunk(src="kontakte", root="outlook", rel="kontakte/Team/c.vcf",
               who="Firma GmbH", ppl="carol carol@example.com", ts=None, date="",
               title="Carol", ctx="Kontakte: Team", text=None),
    ]


def _kaputt(*a, **k):
    raise requests.exceptions.ConnectionError("down")


# --------------------------------------------------------------------------
# build_mask – Personen-/Datums-/Quellenfilter
# --------------------------------------------------------------------------
def test_build_mask_ohne_filter_laesst_alles_durch():
    assert rag_server.build_mask(_meta(), "", None, None, "all").all()
    assert rag_server.build_mask(_meta(), "", None, None, "").all()


def test_build_mask_quelle():
    mask = rag_server.build_mask(_meta(), "", None, None, "teams")
    assert mask.tolist() == [True, False, False, False]


def test_build_mask_person_case_insensitiver_teilstring():
    assert rag_server.build_mask(_meta(), "ALICE", None, None, "all").tolist() == \
        [True, False, True, False]
    assert rag_server.build_mask(_meta(), "bob@example.com", None, None, "all").tolist() == \
        [False, True, False, False]


def test_build_mask_person_mit_fehlendem_ppl():
    assert rag_server.build_mask([_chunk(ppl=None)], "alice", None, None, "all").tolist() == \
        [False]


def test_build_mask_datum_grenzen_inklusive_und_ohne_ts_raus():
    meta = _meta()                                       # ts: 100, 200, 300, None
    assert rag_server.build_mask(meta, "", 200.0, None, "all").tolist() == \
        [False, True, True, False]
    assert rag_server.build_mask(meta, "", None, 200.0, "all").tolist() == \
        [True, True, False, False]
    assert rag_server.build_mask(meta, "", 150.0, 250.0, "all").tolist() == \
        [False, True, False, False]


def test_build_mask_filter_kombinieren_sich():
    mask = rag_server.build_mask(_meta(), "alice", 150.0, None, "kalender")
    assert mask.tolist() == [False, False, True, False]


# --------------------------------------------------------------------------
# rank / browse / src_link / hit_dict
# --------------------------------------------------------------------------
def test_rank_sortiert_nach_kosinus_und_respektiert_k_und_maske():
    V = np.array([[1, 0], [0, 1], [0.6, 0.8]], dtype="float32")
    q = np.array([1, 0], dtype="float32")
    alle = np.ones(3, dtype=bool)
    hits = rag_server.rank(V, q, alle, 3)
    assert [i for i, _ in hits] == [0, 2, 1]
    assert hits[0][1] == pytest.approx(1.0)
    assert hits[1][1] == pytest.approx(0.6)
    assert [i for i, _ in rag_server.rank(V, q, alle, 2)] == [0, 2]
    mask = np.array([False, True, True])
    assert [i for i, _ in rag_server.rank(V, q, mask, 3)] == [2, 1]


def test_rank_leere_maske_liefert_nichts():
    V = np.eye(2, dtype="float32")
    q = np.array([1, 0], dtype="float32")
    assert rag_server.rank(V, q, np.zeros(2, dtype=bool), 5) == []


def test_rank_verwirft_treffer_mit_kosinus_minus_eins():
    # Dokumentiert Ist-Verhalten: sims == -1.0 ist vom Masken-Sentinel nicht
    # unterscheidbar – ein exakt entgegengesetzter Vektor fällt trotz True-Maske raus.
    V = np.array([[-1.0, 0.0]], dtype="float32")
    q = np.array([1.0, 0.0], dtype="float32")
    assert rag_server.rank(V, q, np.ones(1, dtype=bool), 5) == []


def test_browse_sortiert_neueste_zuerst_none_ans_ende():
    meta = _meta()
    mask = np.ones(len(meta), dtype=bool)
    treffer = rag_server.browse(meta, mask, 10)
    assert [i for i, _ in treffer] == [2, 1, 0, 3]       # ts absteigend, None zuletzt
    assert all(s is None for _, s in treffer)            # kein Score im Blätter-Modus
    assert [i for i, _ in rag_server.browse(meta, mask, 2)] == [2, 1]


def test_browse_beachtet_maske():
    mask = np.array([True, False, False, True])
    assert [i for i, _ in rag_server.browse(_meta(), mask, 10)] == [0, 3]


def test_src_link_kodiert_pfadsegmente():
    c = {"root": "teams", "rel": "1on1/ü b.html"}
    assert rag_server.src_link(c) == "/src/teams/1on1/%C3%BC%20b.html"


def test_hit_dict_form_und_preview_kappung():
    h = rag_server.hit_dict(_meta()[1], 0.5)
    assert set(h) == {"who", "date", "title", "ctx", "src", "link", "preview", "score"}
    assert h["link"] == "/src/outlook/inbox/mail.eml"
    assert len(h["preview"]) == 600                      # Text (700 Zeichen) gekappt
    assert h["score"] == 0.5


def test_hit_dict_ohne_text():
    h = rag_server.hit_dict(_meta()[3], None)
    assert h["preview"] == "" and h["score"] is None


# --------------------------------------------------------------------------
# STATE-Fixture (wie main() es befüllt) + Ollama-Helfer
# --------------------------------------------------------------------------
@pytest.fixture
def state(tmp_path):
    """STATE mit Mini-Index und echten Quelldateien befüllen, danach zurücksetzen."""
    teams = tmp_path / "teams_export"
    (teams / "1on1").mkdir(parents=True)
    (teams / "1on1" / "a.html").write_text("<html><body>Teams-Quelle</body></html>",
                                           encoding="utf-8")
    (teams / "1on1" / "ü b.html").write_text("<html>Umlaut</html>", encoding="utf-8")
    outlook = tmp_path / "outlook_export"
    (outlook / "inbox").mkdir(parents=True)
    (outlook / "inbox" / "mail.eml").write_bytes(b"From: alice@example.com\n\nHallo")
    (tmp_path / "geheim.txt").write_text("streng geheim", encoding="utf-8")

    alt = dict(rag_server.STATE)
    rag_server.STATE.clear()
    rag_server.STATE.update(meta=_meta(), V=np.eye(4, dtype="float32"),
                            teams_dir=str(teams), outlook_dir=str(outlook),
                            embed_model="test-embed", chat_model="test-chat",
                            ollama="http://ollama.test")
    yield rag_server.STATE
    rag_server.STATE.clear()
    rag_server.STATE.update(alt)


def test_embed_query_normalisiert_und_fragt_ollama(state, monkeypatch):
    seen = {}

    def fake_post(url, json=None, timeout=None):
        seen["url"], seen["json"] = url, json
        return FakeResp({"embeddings": [[3.0, 4.0]]})

    monkeypatch.setattr(rag_server.requests, "post", fake_post)
    v = rag_server.embed_query("frage")
    assert np.allclose(v, [0.6, 0.8])
    assert seen["url"] == "http://ollama.test/api/embed"
    assert seen["json"] == {"model": "test-embed", "input": ["frage"]}


def test_embed_query_alte_single_form_und_nullvektor(state, monkeypatch):
    monkeypatch.setattr(rag_server.requests, "post",
                        lambda *a, **k: FakeResp({"embedding": [0.0, 0.0]}))
    assert np.array_equal(rag_server.embed_query("x"), [0.0, 0.0])


def test_chat_baut_kontext_mit_quellennummern(state, monkeypatch):
    seen = {}

    def fake_post(url, json=None, timeout=None):
        seen["url"], seen["json"] = url, json
        return FakeResp({"message": {"content": "Antwort [1]"}})

    monkeypatch.setattr(rag_server.requests, "post", fake_post)
    meta = _meta()
    assert rag_server.chat("Was ist fertig?", [meta[0], meta[2]]) == "Antwort [1]"
    assert seen["url"] == "http://ollama.test/api/chat"
    payload = seen["json"]
    assert payload["model"] == "test-chat" and payload["stream"] is False
    system, user = payload["messages"]
    assert system == {"role": "system", "content": rag_server.SYSTEM_PROMPT}
    assert "Frage: Was ist fertig?" in user["content"]
    assert ("[1] (2025-06-01 09:30, Alice Example, teams – 1:1-Chat) "
            "Bericht ist fertig.") in user["content"]
    assert "[2] (2025-06-01 12:00" in user["content"]


# --------------------------------------------------------------------------
# HTTP-Schicht: Wegwerf-Server auf 127.0.0.1 mit ephemerem Port
# --------------------------------------------------------------------------
@pytest.fixture
def server(state):
    httpd = ThreadingHTTPServer(("127.0.0.1", 0), rag_server.Handler)
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    yield httpd.server_address
    httpd.shutdown()
    httpd.server_close()
    thread.join(timeout=5)


def _get(addr, path):
    conn = http.client.HTTPConnection(*addr, timeout=10)
    conn.request("GET", path)
    r = conn.getresponse()
    status, ctype, body = r.status, r.headers.get("Content-Type"), r.read()
    conn.close()
    return status, ctype, body


def _post(addr, path, payload=None, raw=None):
    conn = http.client.HTTPConnection(*addr, timeout=10)
    body = raw if raw is not None else json.dumps(payload or {}).encode("utf-8")
    conn.request("POST", path, body=body, headers={"Content-Type": "application/json"})
    r = conn.getresponse()
    status, data = r.status, r.read()
    conn.close()
    try:
        return status, json.loads(data)
    except ValueError:
        return status, data


# ---- GET: UI + Quelldateien ----
def test_get_liefert_ui(server):
    status, ctype, body = _get(server, "/")
    assert status == 200 and ctype.startswith("text/html")
    assert body.decode("utf-8") == rag_server.UI_HTML
    assert _get(server, "/index.html")[0] == 200


def test_get_unbekannter_pfad_404(server):
    assert _get(server, "/gibtsnicht")[0] == 404


def test_get_quelldatei_teams_und_outlook(server):
    status, ctype, body = _get(server, "/src/teams/1on1/a.html")
    assert status == 200 and ctype == "text/html; charset=utf-8"
    assert b"Teams-Quelle" in body
    status, ctype, _ = _get(server, "/src/outlook/inbox/mail.eml")
    assert status == 200 and ctype == "message/rfc822"


def test_get_quelldatei_ueber_src_link_mit_sonderzeichen(server):
    link = rag_server.src_link({"root": "teams", "rel": "1on1/ü b.html"})
    status, _, body = _get(server, link)
    assert status == 200 and b"Umlaut" in body


def test_get_quelldatei_pfadausbruch_wird_blockiert(server):
    assert _get(server, "/src/teams/../geheim.txt")[0] == 403


def test_get_quelldatei_fehlerfaelle(server):
    assert _get(server, "/src/teams/fehlt.html")[0] == 404   # Datei fehlt
    assert _get(server, "/src/anderswo/x.html")[0] == 404    # unbekannte Wurzel
    assert _get(server, "/src/teams")[0] == 404              # kein Dateipfad


# ---- POST /api/search ----
def test_api_search_ohne_query_blaettert_nach_datum(server):
    status, data = _post(server, "/api/search", {"query": ""})
    assert status == 200 and data["semantic"] is False
    assert [h["title"] for h in data["hits"]] == \
        ["Planung", "Testmail", "Projekt Alpha", "Carol"]
    assert all(h["score"] is None for h in data["hits"])


def test_api_search_filter_und_k(server):
    _, data = _post(server, "/api/search", {"query": "", "src": "teams"})
    assert [h["src"] for h in data["hits"]] == ["teams"]
    _, data = _post(server, "/api/search", {"query": "", "person": "Bob@Example.com"})
    assert [h["title"] for h in data["hits"]] == ["Testmail"]
    _, data = _post(server, "/api/search", {"query": "", "k": 2})
    assert len(data["hits"]) == 2


def test_api_search_semantisch(server, monkeypatch):
    monkeypatch.setattr(rag_server.requests, "post",
                        lambda *a, **k: FakeResp({"embeddings": [[0.0, 1.0, 0.0, 0.0]]}))
    status, data = _post(server, "/api/search", {"query": "testmail", "k": 2})
    assert status == 200 and data["semantic"] is True
    assert len(data["hits"]) == 2
    assert data["hits"][0]["title"] == "Testmail"        # Vektorzeile 1 == Anfrage
    assert data["hits"][0]["score"] == pytest.approx(1.0)


def test_api_search_ollama_nicht_erreichbar_502(server, monkeypatch):
    monkeypatch.setattr(rag_server.requests, "post", _kaputt)
    status, data = _post(server, "/api/search", {"query": "x"})
    assert status == 502
    assert "Ollama nicht erreichbar" in data["error"]
    assert "http://ollama.test" in data["error"]


# ---- POST /api/answer ----
def test_api_answer_ohne_frage_400(server):
    status, data = _post(server, "/api/answer", {"query": "  "})
    assert status == 400 and data["error"] == "Bitte eine Frage eingeben."


def test_api_answer_mit_quellen(server, monkeypatch):
    def fake_post(url, json=None, timeout=None):
        if url.endswith("/api/embed"):
            return FakeResp({"embeddings": [[0.0, 1.0, 0.0, 0.0]]})
        if url.endswith("/api/chat"):
            return FakeResp({"message": {"content": "Die Antwort [1]."}})
        raise AssertionError(f"unerwartete URL: {url}")

    monkeypatch.setattr(rag_server.requests, "post", fake_post)
    status, data = _post(server, "/api/answer", {"query": "Worum geht es?", "k": 2})
    assert status == 200
    assert data["answer"] == "Die Antwort [1]."
    assert len(data["sources"]) == 2
    assert data["sources"][0]["title"] == "Testmail"
    assert data["sources"][0]["score"] == pytest.approx(1.0)
    assert data["sources"][0]["link"] == "/src/outlook/inbox/mail.eml"


def test_api_answer_ohne_treffer_im_filter(server, monkeypatch):
    monkeypatch.setattr(rag_server.requests, "post",
                        lambda *a, **k: FakeResp({"embeddings": [[1.0, 0.0, 0.0, 0.0]]}))
    status, data = _post(server, "/api/answer", {"query": "x", "person": "niemand"})
    assert status == 200
    assert data["sources"] == []
    assert "Keine passenden Quellen" in data["answer"]


def test_api_answer_ollama_nicht_erreichbar_502(server, monkeypatch):
    monkeypatch.setattr(rag_server.requests, "post", _kaputt)
    status, data = _post(server, "/api/answer", {"query": "x"})
    assert status == 502 and "Ollama nicht erreichbar" in data["error"]


# ---- POST: Fehlerfälle ----
def test_post_unbekannter_pfad_404(server):
    assert _post(server, "/api/unbekannt", {})[0] == 404


def test_post_kaputter_body_400(server):
    status, data = _post(server, "/api/search", raw=b"kein json")
    assert status == 400 and data["error"] == "Ungültige Anfrage."
