"""Tests für mcp_server.py – MCP-Tools über einem kleinen, echten Store.

Der Store (corpus.db + vectors.npy) wird pro Test mit den Schreib-Helfern aus
rag_index.py in tmp_path aufgebaut – damit ist das Schema garantiert identisch
mit dem, was mcp_server.py erwartet. Es werden KEINE Netzwerkaufrufe gemacht:
_embed_query wird immer gestubbt (Standard: wirft, wie bei "Ollama down");
Tests der semantischen Suche setzen deterministische Einheitsvektoren.
"""

from datetime import datetime
from urllib.parse import quote

import numpy as np
import pytest

import corpus
import mcp_server
import rag_index

# --------------------------------------------------------------------------
# Testdaten: kleiner Korpus mit Teams-, Outlook-, Kalender- und Kontakt-Einträgen
# --------------------------------------------------------------------------
DIM = 16  # Vektor-Dimension: jeder Chunk bekommt einen eigenen Einheitsvektor

UID_T0 = "teams:1on1/alice__chat.html:0"
UID_T1 = "teams:1on1/alice__chat.html:1"
UID_T2 = "teams:1on1/alice__chat.html:2"
UID_TX = "teams:1on1/max__chat.html:0"
UID_M1 = "outlook:inbox/mail1.eml:0"
UID_M2 = "outlook:inbox/mail2.eml:0"
UID_M3 = "outlook:sent/protokoll.eml:0"
UID_CAL = "kalender:kalender/Arbeit/termin.ics:0"
UID_CON = "kontakte:kontakte/Team/alice.vcf:0"

# Lange Mail → mehrere überlappende Chunks (Test für _join_chunks/get_document)
LONG_TEXT = " ".join(
    f"Absatz {i}: die Quartalsplanung wurde ausführlich besprochen und Punkt {i} im Protokoll festgehalten."
    for i in range(50))

# Bewusst reines ASCII: die Fenster-Tests von read_source_file schneiden an
# Byte-Grenzen; Mehrbyte-Zeichen würden dort (korrekt) zu Ersatzzeichen.
TEAMS_FILE_CONTENT = "<html><body>Chatverlauf Alice und Bob - Projekt Alpha</body></html>"
MAIL_FILE_CONTENT = "From: carla@example.com\nSubject: Rechnung 4711\n\nDie Rechnung ist freigegeben.\n"


def _ts(s):
    return datetime.strptime(s, "%Y-%m-%d %H:%M").timestamp()


def _rec(uid, src, root, rel, who, ppl, ts, date, title, ctx, text):
    return {"uid": uid, "src": src, "root": root, "rel": rel, "who": who,
            "ppl": ppl, "ts": ts, "date": date, "title": title, "ctx": ctx,
            "text": text}


def _sample_records():
    return [
        _rec(UID_T0, "teams", "teams", "1on1/alice__chat.html", "Alice Beispiel",
             "alice beispiel projekt alpha", _ts("2025-06-01 09:30"),
             "2025-06-01 09:30", "Projekt Alpha", "1:1-Chat",
             "Hallo Bob, die Rechnung 4711 für Projekt Alpha ist fertig."),
        _rec(UID_T1, "teams", "teams", "1on1/alice__chat.html", "Bob Baumeister",
             "bob baumeister projekt alpha", _ts("2025-06-01 09:35"),
             "2025-06-01 09:35", "Projekt Alpha", "1:1-Chat",
             "Danke Alice, ich prüfe die Rechnung morgen früh."),
        _rec(UID_T2, "teams", "teams", "1on1/alice__chat.html", "Alice Beispiel",
             "alice beispiel projekt alpha", _ts("2025-06-01 09:40"),
             "2025-06-01 09:40", "Projekt Alpha", "1:1-Chat",
             "Perfekt, dann bis morgen im Büro!"),
        _rec(UID_TX, "teams", "teams", "1on1/max__chat.html", "(unbekannt)",
             "max mustermann", _ts("2025-06-02 10:00"),
             "2025-06-02 10:00", "Max", "1:1-Chat",
             "Kurze Notiz ohne bekannten Absender."),
        _rec(UID_M1, "outlook", "outlook", "inbox/mail1.eml", "Carla Chef",
             "carla chef carla@example.com alice beispiel alice@example.com",
             _ts("2025-06-10 08:00"), "2025-06-10 08:00",
             "Rechnung 4711 freigegeben", "inbox",
             "Hallo zusammen, die Rechnung 4711 ist freigegeben und kann verschickt werden."),
        _rec(UID_M2, "outlook", "outlook", "inbox/mail2.eml", "Alice Beispiel",
             "alice beispiel alice@example.com", _ts("2025-07-01 12:00"),
             "2025-07-01 12:00", "Urlaubsantrag August", "inbox",
             "Hiermit beantrage ich Urlaub vom 4. bis 15. August. Viele Grüße, Alice"),
        _rec(UID_M3, "outlook", "outlook", "sent/protokoll.eml", "Doris Docs",
             "doris docs doris@example.com", _ts("2025-05-20 16:00"),
             "2025-05-20 16:00", "Protokoll Quartalsplanung", "sent", LONG_TEXT),
        _rec(UID_CAL, "kalender", "outlook", "kalender/Arbeit/termin.ics",
             "Alice Beispiel", "alice beispiel bob baumeister",
             _ts("2025-06-15 14:00"), "2025-06-15 14:00", "Quartalsplanung",
             "Kalender: Arbeit", "Ort: Raum 42. Agenda folgt."),
        _rec(UID_CON, "kontakte", "outlook", "kontakte/Team/alice.vcf", "",
             "alice beispiel alice@example.com", None, "", "Alice Beispiel",
             "Kontakte: Team", "Firma GmbH · Entwicklung. E-Mail: alice@example.com"),
    ]


# Neueste zuerst; Kontakt (ts = NULL) am Ende – erwartete browse-Reihenfolge
BROWSE_ORDER = [UID_M2, UID_CAL, UID_M1, UID_TX, UID_T2, UID_T1, UID_T0,
                UID_M3, UID_CON]


def _build_store(tmp_path):
    """Store + Export-Ordner in tmp_path anlegen (Schreibpfad aus rag_index.py)."""
    store = tmp_path / "rag_store"
    store.mkdir()
    teams_dir = tmp_path / "teams_export"
    outlook_dir = tmp_path / "outlook_export"
    (teams_dir / "1on1").mkdir(parents=True)
    (outlook_dir / "inbox").mkdir(parents=True)
    (teams_dir / "1on1" / "alice__chat.html").write_text(
        TEAMS_FILE_CONTENT, encoding="utf-8")
    (outlook_dir / "inbox" / "mail1.eml").write_text(
        MAIL_FILE_CONTENT, encoding="utf-8")
    # Datei AUSSERHALB der Exporte – darf über read_source_file nie erreichbar sein
    (tmp_path / "geheim.txt").write_text("STRENG GEHEIM", encoding="utf-8")

    chunks = corpus.chunk_records(_sample_records())
    for c in chunks:
        c["hash"] = corpus.chunk_hash(c)
    assert len(chunks) <= DIM, "Testkorpus zu groß für die Vektor-Dimension"
    # Chunk i → Einheitsvektor e_i: Kosinus zum Query-Vektor q ist exakt q[i]
    V = np.zeros((len(chunks), DIM), dtype="float32")
    for i in range(len(chunks)):
        V[i, i] = 1.0
    rag_index.write_db(store, chunks)
    rag_index.save_vectors(store, V)
    rag_index.write_info(store, "test-embed", DIM, len(chunks))
    return store, chunks, teams_dir, outlook_dir


@pytest.fixture
def state(tmp_path, monkeypatch):
    """Store bauen, STATE füllen und nach dem Test wiederherstellen.

    _embed_query wirft standardmäßig (kein Netzwerk!); semantische Tests
    überschreiben den Stub mit deterministischen Vektoren.
    """
    store, chunks, teams_dir, outlook_dir = _build_store(tmp_path)
    old = dict(mcp_server.STATE)
    mcp_server.STATE.clear()
    V = np.load(store / "vectors.npy", mmap_mode="r")
    mcp_server.STATE.update(
        db=str(store / "corpus.db"), V=V, np=np, semantic=True,
        vector_dtype=str(V.dtype), teams_dir=str(teams_dir),
        outlook_dir=str(outlook_dir), embed_model="test-embed",
        ollama="http://127.0.0.1:1")

    def _kein_netz(text):
        raise RuntimeError("Embedding nicht gestubbt (Tests machen kein Netzwerk)")

    monkeypatch.setattr(mcp_server, "_embed_query", _kein_netz)
    yield {"store": store, "chunks": chunks, "tmp": tmp_path,
           "teams_dir": teams_dir, "outlook_dir": outlook_dir}
    mcp_server.STATE.clear()
    mcp_server.STATE.update(old)


@pytest.fixture
def empty_state():
    """STATE leeren (Server nicht initialisiert) und danach wiederherstellen."""
    old = dict(mcp_server.STATE)
    mcp_server.STATE.clear()
    yield
    mcp_server.STATE.clear()
    mcp_server.STATE.update(old)


def _stub_semantic(monkeypatch, chunks, weights):
    """_embed_query so stubben, dass uid → Gewicht die Kosinus-Rangfolge vorgibt."""
    q = np.zeros(DIM, dtype="float32")
    for i, c in enumerate(chunks):
        q[i] = weights.get(c["uid"], 0.0)
    nrm = np.linalg.norm(q)
    if nrm:
        q = q / nrm
    monkeypatch.setattr(mcp_server, "_embed_query", lambda text: q)


def _uids(res):
    return [h["uid"] for h in res["results"]]


# --------------------------------------------------------------------------
# Hilfsfunktionen (ohne Store)
# --------------------------------------------------------------------------
def test_to_ts_parses_and_clamps_day_end():
    assert mcp_server._to_ts("2025-06-01", False) == datetime(2025, 6, 1).timestamp()
    assert mcp_server._to_ts("2025-06-01", True) == datetime(2025, 6, 1, 23, 59, 59).timestamp()
    assert mcp_server._to_ts("01.06.2025", False) is None  # falsches Format
    assert mcp_server._to_ts("", False) is None
    assert mcp_server._to_ts(None, True) is None


def test_where_builds_fragments():
    w, p = mcp_server._where("", None, None, "all")
    assert w == "1=1" and p == []
    w, p = mcp_server._where("Alice", 1.0, 2.0, "teams")
    assert "src = ?" in w and "ppl LIKE ?" in w
    assert "ts >= ?" in w and "ts <= ?" in w
    assert p == ["teams", "%alice%", 1.0, 2.0]  # Person wird kleingeschrieben


def test_fts_match_sanitizes_query():
    # Freitext wird zu einer ODER-Liste zitierter Tokens – FTS5-Syntax
    # (AND/OR/NEAR, Klammern, Anführungszeichen) kann nicht injiziert werden.
    assert mcp_server._fts_match('Rechnung: 4711 AND "x(y)') == '"rechnung" OR "4711" OR "and" OR "x" OR "y"'
    assert mcp_server._fts_match("Größe") == '"größe"'
    assert mcp_server._fts_match("...!!!") == ""
    assert mcp_server._fts_match("") == ""


def test_rrf_merge_orders_by_reciprocal_rank():
    sem = [(1, 0.9), (2, 0.5)]
    lex = [(2, -1.0), (3, -2.0)]
    merged = mcp_server._rrf_merge(sem, lex)
    assert [cid for cid, _ in merged] == [2, 1, 3]  # 2 ist in beiden Listen
    scores = dict(merged)
    assert scores[2] == pytest.approx(1 / 61 + 1 / 62)
    assert scores[1] == pytest.approx(1 / 61)


def test_join_chunks_removes_overlap():
    rows = [{"text": "abcdef"}, {"text": "defghi"}, {"text": "xyz"}]
    assert mcp_server._join_chunks(rows) == "abcdefghixyz"
    assert mcp_server._join_chunks([{"text": ""}, {"text": "abc"}]) == "abc"
    assert mcp_server._join_chunks([]) == ""


def test_source_uri_percent_encodes_path():
    uri = mcp_server._source_uri("teams", "1on1/alice chat.html")
    assert uri == "o365://teams/1on1%2Falice%20chat.html"


def test_read_window_replaces_clipped_utf8(tmp_path):
    f = tmp_path / "umlaut.txt"
    f.write_bytes("ää".encode())  # 4 Bytes
    text, total, start, more = mcp_server._read_window(f, 0, 3)
    assert total == 4 and start == 0 and more
    assert text.startswith("ä") and "�" in text  # zerschnittene Sequenz


# --------------------------------------------------------------------------
# corpus_stats
# --------------------------------------------------------------------------
def test_corpus_stats_counts_per_source(state):
    chunks = state["chunks"]
    out = mcp_server.corpus_stats()
    assert out["chunks"] == len(chunks)
    assert out["by_source"]["teams"] == {"chunks": 4, "messages": 4}
    assert out["by_source"]["kalender"] == {"chunks": 1, "messages": 1}
    assert out["by_source"]["kontakte"] == {"chunks": 1, "messages": 1}
    n_outlook = sum(1 for c in chunks if c["src"] == "outlook")
    assert n_outlook > 3  # die lange Mail wurde wirklich in mehrere Chunks geteilt
    assert out["by_source"]["outlook"] == {"chunks": n_outlook, "messages": 3}
    assert out["semantic_available"] is True
    assert out["default_backend"] == "hybrid"
    assert out["embed_model"] == "test-embed"
    assert out["vector_dtype"] == "float16"
    assert out["teams_dir"] == str(state["teams_dir"])


def test_corpus_stats_lexical_when_semantic_off(state):
    mcp_server.STATE["semantic"] = False
    out = mcp_server.corpus_stats()
    assert out["default_backend"] == "lexical"
    assert out["semantic_available"] is False
    assert out["embed_model"] is None


# --------------------------------------------------------------------------
# search_messages – lexikalischer Pfad (FTS5/BM25)
# --------------------------------------------------------------------------
def test_search_lexical_finds_and_dedupes(state):
    res = mcp_server.search_messages("Rechnung", mode="lexical")
    assert res["backend"] == "lexical"
    uids = _uids(res)
    assert set(uids) == {UID_T0, UID_T1, UID_M1}
    assert len(uids) == len(set(uids))  # eine Nachricht nur einmal
    hit = res["results"][0]
    assert hit["source_label"] in ("Teams", "Mail")
    assert hit["uri"].startswith("o365://")
    assert hit["score"] is not None
    assert "Rechnung" in hit["preview"]


def test_search_lexical_no_hits_and_empty_query(state):
    assert mcp_server.search_messages("xyzzyplugh", mode="lexical")["count"] == 0
    res = mcp_server.search_messages("", mode="lexical")
    assert res["count"] == 0 and res["results"] == []


def test_search_source_filter(state):
    res = mcp_server.search_messages("Rechnung", source="outlook", mode="lexical")
    assert _uids(res) == [UID_M1]
    res = mcp_server.search_messages("Rechnung", source="teams", mode="lexical")
    assert set(_uids(res)) == {UID_T0, UID_T1}


def test_search_person_filter(state):
    # Personenfilter läuft über die ppl-Spalte (kleingeschriebene Namen + Adressen)
    res = mcp_server.search_messages("Rechnung", person="Carla", mode="lexical")
    assert _uids(res) == [UID_M1]
    res = mcp_server.search_messages("Rechnung", person="carla@example.com",
                                     mode="lexical")
    assert _uids(res) == [UID_M1]
    res = mcp_server.search_messages("Rechnung", person="Niemand", mode="lexical")
    assert res["count"] == 0


def test_search_date_filters(state):
    # Teams-Treffer sind vom 01.06., die Mail vom 10.06.
    res = mcp_server.search_messages("Rechnung", date_from="2025-06-05",
                                     mode="lexical")
    assert _uids(res) == [UID_M1]
    res = mcp_server.search_messages("Rechnung", date_to="2025-06-05",
                                     mode="lexical")
    assert set(_uids(res)) == {UID_T0, UID_T1}
    # date_to ist inklusiv (bis 23:59:59 des Tages)
    res = mcp_server.search_messages("Rechnung", date_to="2025-06-10",
                                     mode="lexical")
    assert UID_M1 in _uids(res)


def test_search_k_and_offset_page_through_results(state):
    page1 = mcp_server.search_messages("Rechnung", k=2, offset=0, mode="lexical")
    page2 = mcp_server.search_messages("Rechnung", k=2, offset=2, mode="lexical")
    assert page1["count"] == 2 and page2["count"] == 1
    assert page1["offset"] == 0 and page2["offset"] == 2
    assert set(_uids(page1)) | set(_uids(page2)) == {UID_T0, UID_T1, UID_M1}
    assert not set(_uids(page1)) & set(_uids(page2))


def test_search_preview_chars(state):
    res = mcp_server.search_messages("Rechnung", mode="lexical", preview_chars=10)
    assert all(len(h["preview"]) <= 10 for h in res["results"])
    res = mcp_server.search_messages("Rechnung", mode="lexical", preview_chars=0)
    assert all("preview" not in h for h in res["results"])


# --------------------------------------------------------------------------
# search_messages – semantischer Pfad und Hybrid-Fusion (RRF)
# --------------------------------------------------------------------------
def test_search_semantic_ranks_by_stubbed_cosine(state, monkeypatch):
    # Query-Vektor: Urlaubsmail am ähnlichsten, Rechnungsmail auf Platz 2
    _stub_semantic(monkeypatch, state["chunks"], {UID_M2: 1.0, UID_M1: 0.5})
    res = mcp_server.search_messages("freie Tage im Sommer", mode="semantic")
    assert res["backend"] == "semantic"
    uids = _uids(res)
    assert uids[0] == UID_M2 and uids[1] == UID_M1
    scores = [h["score"] for h in res["results"]]
    assert scores == sorted(scores, reverse=True)
    # Kosinus entspricht den (normierten) Stub-Gewichten: Platz 2 = halber Score
    assert scores[1] == pytest.approx(scores[0] * 0.5, abs=0.01)
    assert scores[0] > 0.8


def test_search_hybrid_fuses_semantic_and_lexical(state, monkeypatch):
    # Semantik: M1 vor M2. Lexikalisch trifft "Rechnung 4711" M1/T0/T1, aber
    # nie M2 – M2 kann nur über den semantischen Zweig in die Liste kommen.
    _stub_semantic(monkeypatch, state["chunks"], {UID_M1: 1.0, UID_M2: 0.6})
    res = mcp_server.search_messages("Rechnung 4711", mode="hybrid")
    assert res["backend"] == "hybrid"
    uids = _uids(res)
    assert uids[0] == UID_M1        # Platz 1 in beiden Backends → RRF-Sieger
    assert UID_M2 in uids           # reiner Semantik-Treffer bleibt erhalten
    assert UID_T0 in uids           # reiner BM25-Treffer bleibt erhalten


def test_search_auto_falls_back_to_lexical_when_ollama_down(state):
    # Der Fixture-Stub für _embed_query wirft – wie ein nicht erreichbares Ollama
    res = mcp_server.search_messages("Rechnung", mode="auto")
    assert res["backend"] == "lexical"
    assert set(_uids(res)) == {UID_T0, UID_T1, UID_M1}
    assert "nicht gestubbt" in mcp_server.STATE["last_semantic_error"]


def test_search_semantic_mode_reports_error_when_ollama_down(state):
    res = mcp_server.search_messages("Rechnung", mode="semantic")
    assert set(res) == {"error"}
    assert "Semantic ranking failed" in res["error"]


def test_search_lexical_mode_never_touches_embeddings(state):
    # mode="lexical" darf _embed_query gar nicht erst aufrufen
    res = mcp_server.search_messages("Urlaub", mode="lexical")
    assert res["backend"] == "lexical"
    assert _uids(res) == [UID_M2]
    assert "last_semantic_error" not in mcp_server.STATE


# --------------------------------------------------------------------------
# browse_messages
# --------------------------------------------------------------------------
def test_browse_newest_first_nulls_last(state):
    res = mcp_server.browse_messages(k=50)
    assert _uids(res) == BROWSE_ORDER  # ts absteigend, Kontakt ohne ts am Ende
    assert res["count"] == len(BROWSE_ORDER)
    assert res["results"][0]["score"] is None  # browse hat keine Relevanzwertung


def test_browse_pagination(state):
    page1 = mcp_server.browse_messages(k=4, offset=0)
    page2 = mcp_server.browse_messages(k=4, offset=4)
    page3 = mcp_server.browse_messages(k=4, offset=8)
    assert _uids(page1) == BROWSE_ORDER[:4]
    assert _uids(page2) == BROWSE_ORDER[4:8]
    assert _uids(page3) == BROWSE_ORDER[8:]
    assert mcp_server.browse_messages(k=4, offset=100)["count"] == 0


def test_browse_filters(state):
    assert _uids(mcp_server.browse_messages(source="teams")) == \
        [UID_TX, UID_T2, UID_T1, UID_T0]
    assert _uids(mcp_server.browse_messages(source="kontakte")) == [UID_CON]
    res = mcp_server.browse_messages(person="bob", source="teams")
    assert _uids(res) == [UID_T1]
    res = mcp_server.browse_messages(date_from="2025-06-10", date_to="2025-06-30")
    assert _uids(res) == [UID_CAL, UID_M1]


def test_browse_preview_toggle(state):
    res = mcp_server.browse_messages(source="kalender")
    assert res["results"][0]["preview"].startswith("Ort: Raum 42.")
    res = mcp_server.browse_messages(source="kalender", preview_chars=0)
    assert "preview" not in res["results"][0]


# --------------------------------------------------------------------------
# get_document
# --------------------------------------------------------------------------
def test_get_document_rejoins_chunks_to_full_text(state):
    assert sum(c["uid"] == UID_M3 for c in state["chunks"]) > 1
    out = mcp_server.get_document(UID_M3)
    assert out["text"] == LONG_TEXT  # Überlappungen exakt entfernt
    assert out["title"] == "Protokoll Quartalsplanung"
    assert out["source"] == "outlook" and out["source_label"] == "Mail"
    assert out["uri"] == "o365://outlook/" + quote("sent/protokoll.eml", safe="")
    assert "context_before" not in out  # ohne Kontext-Parameter kein Kontext


def test_get_document_unknown_uid(state):
    out = mcp_server.get_document("outlook:gibtsnicht.eml:0")
    assert "error" in out and "gibtsnicht" in out["error"]


def test_get_document_conversation_context(state):
    out = mcp_server.get_document(UID_T1, context_before=1, context_after=1)
    assert [e["uid"] for e in out["context_before"]] == [UID_T0]
    assert [e["uid"] for e in out["context_after"]] == [UID_T2]
    assert out["context_before"][0]["who"] == "Alice Beispiel"
    assert "Rechnung 4711" in out["context_before"][0]["text"]
    # Kontext stammt nur aus derselben Datei – die fremde Teams-Datei fehlt
    out = mcp_server.get_document(UID_T0, context_before=5, context_after=5)
    ctx_uids = {e["uid"] for e in out["context_before"] + out["context_after"]}
    assert ctx_uids == {UID_T1, UID_T2}


# --------------------------------------------------------------------------
# list_people
# --------------------------------------------------------------------------
def test_list_people_counts_and_excludes_unknown(state):
    out = mcp_server.list_people()
    people = {p["name"]: p["messages"] for p in out["people"]}
    assert people == {"Alice Beispiel": 4, "Bob Baumeister": 1,
                      "Carla Chef": 1, "Doris Docs": 1}
    assert out["people"][0]["name"] == "Alice Beispiel"  # häufigste zuerst
    assert out["total_distinct"] == 4
    assert "(unbekannt)" not in people and "" not in people


def test_list_people_source_contains_and_limit(state):
    out = mcp_server.list_people(source="teams")
    assert {p["name"]: p["messages"] for p in out["people"]} == \
        {"Alice Beispiel": 2, "Bob Baumeister": 1}
    # contains matcht Name ODER ppl-Tokens (auch E-Mail-Adressen)
    out = mcp_server.list_people(contains="carla")
    assert [p["name"] for p in out["people"]] == ["Carla Chef"]
    out = mcp_server.list_people(contains="doris@example.com")
    assert [p["name"] for p in out["people"]] == ["Doris Docs"]
    out = mcp_server.list_people(limit=1)
    assert out["count"] == 1 and out["total_distinct"] == 4


# --------------------------------------------------------------------------
# read_source_file – inkl. Path-Traversal-Schutz (sicherheitsrelevant!)
# --------------------------------------------------------------------------
def test_read_source_file_reads_export_file(state):
    out = mcp_server.read_source_file("teams", "1on1/alice__chat.html")
    assert out["content"] == TEAMS_FILE_CONTENT
    assert out["suffix"] == ".html"
    assert out["total_bytes"] == len(TEAMS_FILE_CONTENT.encode())
    assert out["offset"] == 0 and out["truncated"] is False
    out = mcp_server.read_source_file("outlook", "inbox/mail1.eml")
    assert "Rechnung 4711" in out["content"]


def test_read_source_file_windows_with_offset(state):
    # Kleine Fenster + offset müssen die Datei lückenlos rekonstruieren
    total = len(TEAMS_FILE_CONTENT.encode())
    parts, offset = [], 0
    while True:
        out = mcp_server.read_source_file("teams", "1on1/alice__chat.html",
                                          max_chars=10, offset=offset)
        parts.append(out["content"])
        offset += 10
        if not out["truncated"]:
            break
    assert "".join(parts) == TEAMS_FILE_CONTENT
    assert len(parts) == -(-total // 10)


def test_read_source_file_rejects_path_traversal(state):
    # Die Geheimdatei liegt direkt über den Export-Ordnern
    for evil in ("../geheim.txt", "../../geheim.txt", "1on1/../../geheim.txt"):
        out = mcp_server.read_source_file("teams", evil)
        assert out == {"error": "Path outside the export directory."}
    out = mcp_server.read_source_file("outlook", "../geheim.txt")
    assert "error" in out and "GEHEIM" not in str(out)


def test_read_source_file_rejects_absolute_paths(state):
    secret = state["tmp"] / "geheim.txt"
    out = mcp_server.read_source_file("teams", str(secret))
    assert out == {"error": "Path outside the export directory."}
    out = mcp_server.read_source_file("outlook", "/etc/passwd")
    assert out == {"error": "Path outside the export directory."}


def test_read_source_file_rejects_symlink_escape(state):
    # Symlink INNERHALB des Exports, Ziel außerhalb → muss abgelehnt werden
    link = state["outlook_dir"] / "inbox" / "link.eml"
    link.symlink_to(state["tmp"] / "geheim.txt")
    out = mcp_server.read_source_file("outlook", "inbox/link.eml")
    assert out == {"error": "Path outside the export directory."}


def test_read_source_file_invalid_root_and_missing_file(state):
    out = mcp_server.read_source_file("kalender", "termin.ics")
    assert out == {"error": "source_root must be 'teams' or 'outlook'."}
    out = mcp_server.read_source_file("teams", "1on1/fehlt.html")
    assert out == {"error": "File not found: 1on1/fehlt.html"}
    out = mcp_server.read_source_file("teams", "")  # Verzeichnis, keine Datei
    assert "error" in out


# --------------------------------------------------------------------------
# MCP-Resource o365://{root}/{path}
# --------------------------------------------------------------------------
def test_source_resource_returns_file_by_encoded_uri(state):
    content = mcp_server.source_resource(
        "teams", quote("1on1/alice__chat.html", safe=""))
    assert content == TEAMS_FILE_CONTENT


def test_source_resource_rejects_traversal(state):
    with pytest.raises(ValueError, match="outside the export directory"):
        mcp_server.source_resource("teams", quote("../geheim.txt", safe=""))
    with pytest.raises(ValueError):
        mcp_server.source_resource("wurzel", "x")


# --------------------------------------------------------------------------
# Nicht initialisierter STATE
# --------------------------------------------------------------------------
def test_tools_without_initialized_state(empty_state):
    # read_source_file scheitert kontrolliert (kein Export-Verzeichnis bekannt) …
    out = mcp_server.read_source_file("teams", "x.html")
    assert out == {"error": "source_root must be 'teams' or 'outlook'."}
    # … die DB-gestützten Tools werfen mangels STATE["db"] einen KeyError
    # (aktuelles Verhalten – hier festgenagelt)
    with pytest.raises(KeyError):
        mcp_server.corpus_stats()
    with pytest.raises(KeyError):
        mcp_server.search_messages("test", mode="lexical")
    with pytest.raises(KeyError):
        mcp_server.browse_messages()


def test_list_people_contains_ist_umlaut_unabhaengig(tmp_path):
    """SQLite-LIKE ist nur ASCII-case-insensitiv – py_lower() lässt auch
    großgeschriebene Umlaut-Eingaben ("MÜLLER") den Namen finden."""
    store = tmp_path / "store_umlaut"
    store.mkdir()
    recs = [_rec("teams:x.html:0", "teams", "teams", "x.html", "Jörg Müller",
                 "jörg müller joerg@example.com", _ts("2025-06-01 10:00"),
                 "2025-06-01 10:00", "Chat", "1:1-Chat", "Servus!")]
    chunks = corpus.chunk_records(recs)
    for c in chunks:
        c["hash"] = corpus.chunk_hash(c)
    rag_index.write_db(store, chunks)

    old = dict(mcp_server.STATE)
    mcp_server.STATE.clear()
    mcp_server.STATE["db"] = str(store / "corpus.db")
    try:
        out = mcp_server.list_people(contains="MÜLLER")
        assert [p["name"] for p in out["people"]] == ["Jörg Müller"]
        out = mcp_server.list_people(contains="JOERG@EXAMPLE.COM")  # ppl-Token
        assert [p["name"] for p in out["people"]] == ["Jörg Müller"]
        out = mcp_server.list_people(contains="gibtsnicht")
        assert out["people"] == []
    finally:
        mcp_server.STATE.clear()
        mcp_server.STATE.update(old)
