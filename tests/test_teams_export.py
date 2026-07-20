"""Tests für teams_export.py – Graph-Client, Rendering, Fortschritt und Job-Aufbau.

Alle Netzwerkzugriffe sind durch Fakes ersetzt (SESSION bzw. Graph-Objekte);
es wird nie wirklich das Netz berührt. Die reinen Helfer (safe, parse_ts, …)
sind bereits in test_teams_export_helpers.py abgedeckt.
"""

import json
import re
import threading

import pytest
import requests

import teams_export as te

GRAPH = te.GRAPH
TIME_RE = re.compile(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}")   # lokale Zeit, ohne festen Wert


# --------------------------------------------------------------------------
# Gemeinsame Fakes und Fixtures
# --------------------------------------------------------------------------
class FakeResponse:
    """Minimaler Ersatz für requests.Response."""

    def __init__(self, status_code=200, payload=None, headers=None, content=b""):
        self.status_code = status_code
        self._payload = payload if payload is not None else {}
        self.headers = headers or {}
        self.content = content

    def json(self):
        return self._payload

    def raise_for_status(self):
        if self.status_code >= 400:
            raise requests.HTTPError(f"HTTP {self.status_code}")


class FakeSession:
    """Gibt vorbereitete Antworten der Reihe nach zurück und protokolliert Aufrufe."""

    def __init__(self, responses):
        self.responses = list(responses)
        self.calls = []

    def get(self, url, headers=None, params=None, timeout=None):
        self.calls.append((url, params))
        return self.responses.pop(0)


class FakeGraph:
    """Fake für Graph/TokenClient: paged/get aus vorbereiteten Daten je URL."""

    channels_enabled = True

    def __init__(self, pages=None, gets=None):
        self.pages = pages or {}
        self.gets = gets or {}
        self.paged_calls = []

    def paged(self, url, params=None):
        self.paged_calls.append(url)
        val = self.pages[url]
        if isinstance(val, Exception):
            raise val
        yield from val

    def get(self, url, params=None):
        return self.gets[url]


@pytest.fixture(autouse=True)
def _clear_stop():
    """STOP ist ein modulweites Event – nach jedem Test zurücksetzen."""
    yield
    te.STOP.clear()


@pytest.fixture
def sleeps(monkeypatch):
    """time.sleep abklemmen und die gewünschten Wartezeiten mitschreiben."""
    calls = []
    monkeypatch.setattr(te.time, "sleep", lambda s: calls.append(s))
    return calls


def _msg(name, text, ts, ctype="text", **extra):
    m = {
        "messageType": "message",
        "from": {"user": {"displayName": name}},
        "body": {"contentType": ctype, "content": text},
        "createdDateTime": ts,
    }
    m.update(extra)
    return m


# --------------------------------------------------------------------------
# human_time (parse_ts & Co. sind schon in test_teams_export_helpers.py)
# --------------------------------------------------------------------------
def test_human_time_formats_iso_locally():
    out = te.human_time("2025-06-01T09:30:00Z")
    assert TIME_RE.fullmatch(out)   # exakte Uhrzeit ist zeitzonenabhängig
    # 7-stellige Sekundenbruchteile (Graph) werden verkraftet
    assert TIME_RE.fullmatch(te.human_time("2025-06-01T09:30:00.1234567Z"))


def test_human_time_passes_garbage_through():
    assert te.human_time("") == ""
    assert te.human_time(None) == ""
    assert te.human_time("unsinn") == "unsinn"   # unparsebar -> unverändert zurück


# --------------------------------------------------------------------------
# HOSTED_RE – Erkennung von hostedContents-URLs
# --------------------------------------------------------------------------
def test_hosted_re_matches_v1_and_beta():
    u1 = "https://graph.microsoft.com/v1.0/chats/1/messages/2/hostedContents/abc/$value"
    u2 = "https://graph.microsoft.com/beta/teams/x/channels/y/messages/z/hostedContents/q/$value"
    assert te.HOSTED_RE.search(f'<img src="{u1}">').group(0) == u1
    assert te.HOSTED_RE.search(f'<img src="{u2}">').group(0) == u2
    assert te.HOSTED_RE.search('<img src="https://example.com/bild.png">') is None


# --------------------------------------------------------------------------
# TokenClient – Retry, Backoff, TokenExpired, Paging
# --------------------------------------------------------------------------
def test_tokenclient_get_success_sends_bearer(monkeypatch):
    sess = FakeSession([FakeResponse(200, {"ok": True})])
    monkeypatch.setattr(te, "SESSION", sess)
    tc = te.TokenClient("tok123", channels_enabled=False)
    assert tc.get("https://x/y", {"$top": 5}) == {"ok": True}
    assert sess.calls == [("https://x/y", {"$top": 5})]


def test_tokenclient_get_401_raises_tokenexpired(monkeypatch):
    monkeypatch.setattr(te, "SESSION", FakeSession([FakeResponse(401)]))
    tc = te.TokenClient("tok", channels_enabled=False)
    with pytest.raises(te.TokenExpired):
        tc.get("https://x/y")


def test_tokenclient_get_retries_429_with_retry_after(monkeypatch, sleeps):
    sess = FakeSession([
        FakeResponse(429, headers={"Retry-After": "3"}),
        FakeResponse(200, {"ok": 1}),
    ])
    monkeypatch.setattr(te, "SESSION", sess)
    tc = te.TokenClient("tok", channels_enabled=False)
    assert tc.get("https://x/y") == {"ok": 1}
    assert sleeps == [3]   # Retry-After-Header wird respektiert


def test_tokenclient_get_retries_5xx_with_backoff(monkeypatch, sleeps):
    sess = FakeSession([FakeResponse(503), FakeResponse(502), FakeResponse(200, {"ok": 1})])
    monkeypatch.setattr(te, "SESSION", sess)
    tc = te.TokenClient("tok", channels_enabled=False)
    assert tc.get("https://x/y") == {"ok": 1}
    assert sleeps == [1, 2]   # exponentielles Backoff 2**attempt


def test_tokenclient_get_gives_up_after_six_attempts(monkeypatch, sleeps):
    monkeypatch.setattr(te, "SESSION", FakeSession([FakeResponse(429)] * 6))
    tc = te.TokenClient("tok", channels_enabled=False)
    with pytest.raises(RuntimeError, match="Zu viele Fehlversuche"):
        tc.get("https://x/y")


def test_tokenclient_get_raises_on_client_error(monkeypatch):
    monkeypatch.setattr(te, "SESSION", FakeSession([FakeResponse(404)]))
    tc = te.TokenClient("tok", channels_enabled=False)
    with pytest.raises(requests.HTTPError):
        tc.get("https://x/y")


def test_tokenclient_get_bytes_returns_content_and_type(monkeypatch):
    sess = FakeSession([FakeResponse(200, headers={"Content-Type": "image/png"},
                                     content=b"\x89PNG")])
    monkeypatch.setattr(te, "SESSION", sess)
    tc = te.TokenClient("tok", channels_enabled=False)
    assert tc.get_bytes("https://x/img") == (b"\x89PNG", "image/png")


def test_tokenclient_get_bytes_5xx_is_image_unavailable(monkeypatch):
    monkeypatch.setattr(te, "SESSION", FakeSession([FakeResponse(502)]))
    tc = te.TokenClient("tok", channels_enabled=False)
    with pytest.raises(te.ImageUnavailable):   # kein Retry bei Serverfehler
        tc.get_bytes("https://x/img")


def test_tokenclient_get_bytes_persistent_429_is_image_unavailable(monkeypatch, sleeps):
    monkeypatch.setattr(te, "SESSION", FakeSession([FakeResponse(429)] * 4))
    tc = te.TokenClient("tok", channels_enabled=False)
    with pytest.raises(te.ImageUnavailable):
        tc.get_bytes("https://x/img")
    assert len(sleeps) == 4


def test_tokenclient_get_bytes_401_raises_tokenexpired(monkeypatch):
    monkeypatch.setattr(te, "SESSION", FakeSession([FakeResponse(401)]))
    tc = te.TokenClient("tok", channels_enabled=False)
    with pytest.raises(te.TokenExpired):
        tc.get_bytes("https://x/img")


def test_tokenclient_paged_follows_nextlink(monkeypatch):
    sess = FakeSession([
        FakeResponse(200, {"value": [1, 2], "@odata.nextLink": "https://x/page2"}),
        FakeResponse(200, {"value": [3]}),
    ])
    monkeypatch.setattr(te, "SESSION", sess)
    tc = te.TokenClient("tok", channels_enabled=False)
    assert list(tc.paged("https://x/page1", {"$top": 2})) == [1, 2, 3]
    # nextLink wird absolut und ohne zusätzliche Parameter aufgerufen
    assert sess.calls[1] == ("https://x/page2", None)


# --------------------------------------------------------------------------
# Graph – 401 löst Token-Refresh aus, Paging läuft über get()
# --------------------------------------------------------------------------
def _bare_graph():
    """Graph-Instanz ohne interaktiven Login (kein __init__)."""
    g = object.__new__(te.Graph)
    g.token = "alt"
    g._refresh_lock = threading.Lock()
    return g


def test_graph_get_refreshes_token_on_401(monkeypatch):
    sess = FakeSession([FakeResponse(401), FakeResponse(200, {"ok": 1})])
    monkeypatch.setattr(te, "SESSION", sess)
    g = _bare_graph()
    refreshed = []

    def fake_refresh():
        refreshed.append(True)
        g.token = "neu"

    g._refresh = fake_refresh
    assert g.get("https://x/y") == {"ok": 1}
    assert refreshed == [True]


def test_graph_get_retries_and_gives_up(monkeypatch, sleeps):
    monkeypatch.setattr(te, "SESSION", FakeSession([FakeResponse(500)] * 6))
    g = _bare_graph()
    with pytest.raises(RuntimeError, match="Zu viele Fehlversuche"):
        g.get("https://x/y")
    assert len(sleeps) == 6


def test_graph_paged_follows_nextlink(monkeypatch):
    sess = FakeSession([
        FakeResponse(200, {"value": ["a"], "@odata.nextLink": "https://x/n"}),
        FakeResponse(200, {"value": ["b"]}),
    ])
    monkeypatch.setattr(te, "SESSION", sess)
    g = _bare_graph()
    assert list(g.paged("https://x/1")) == ["a", "b"]


def test_graph_get_bytes_401_refresh_then_ok(monkeypatch):
    sess = FakeSession([FakeResponse(401),
                        FakeResponse(200, headers={"Content-Type": "image/gif"}, content=b"GIF")])
    monkeypatch.setattr(te, "SESSION", sess)
    g = _bare_graph()
    g._refresh = lambda: None
    assert g.get_bytes("https://x/img") == (b"GIF", "image/gif")


# --------------------------------------------------------------------------
# Token-Modus: load_pasted_token
# --------------------------------------------------------------------------
def test_load_pasted_token_from_env(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)   # kein gx_token.txt aus dem Repo einlesen
    monkeypatch.setenv("GRAPH_TOKEN", '  "Bearer eyJ0abc"  ')
    assert te.load_pasted_token() == "eyJ0abc"   # Anführungszeichen + Präfix entfernt


def test_load_pasted_token_from_file(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    monkeypatch.delenv("GRAPH_TOKEN", raising=False)
    (tmp_path / "gx_token.txt").write_text("  eyJ0datei \n", encoding="utf-8")
    assert te.load_pasted_token() == "eyJ0datei"


def test_load_pasted_token_missing_or_empty(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    monkeypatch.delenv("GRAPH_TOKEN", raising=False)
    assert te.load_pasted_token() is None
    (tmp_path / "gx_token.txt").write_text("   \n", encoding="utf-8")
    assert te.load_pasted_token() is None


# --------------------------------------------------------------------------
# Interaktive Abfragen (ohne Terminal deterministisch)
# --------------------------------------------------------------------------
def test_read_returns_empty_on_eof(monkeypatch):
    def raise_eof(prompt):
        raise EOFError
    monkeypatch.setattr("builtins.input", raise_eof)
    assert te._read("? ") == ""


class _NoTTY:
    def isatty(self):
        return False


def test_prompt_categories_without_tty_uses_default(monkeypatch):
    monkeypatch.setattr(te.sys, "stdin", _NoTTY())
    options = [("1on1", "a"), ("group", "b"), ("meeting", "c"), ("channels", "d")]
    assert te.prompt_categories(options) == {"1on1", "group", "meeting"}


def test_select_teams_without_tty_takes_all_sorted(monkeypatch):
    monkeypatch.setattr(te.sys, "stdin", _NoTTY())
    graph = FakeGraph(pages={f"{GRAPH}/me/joinedTeams": [
        {"displayName": "Beta"}, {"displayName": "alpha"}]})
    teams = te.select_teams(graph)
    assert [t["displayName"] for t in teams] == ["alpha", "Beta"]


def test_select_teams_handles_errors():
    graph = FakeGraph(pages={f"{GRAPH}/me/joinedTeams": RuntimeError("kaputt")})
    assert te.select_teams(graph) == []
    graph = FakeGraph(pages={f"{GRAPH}/me/joinedTeams": te.TokenExpired()})
    with pytest.raises(te.TokenExpired):   # Token-Ende wird durchgereicht
        te.select_teams(graph)


# --------------------------------------------------------------------------
# member_name / chat_title
# --------------------------------------------------------------------------
def test_member_name_prefers_displayname_then_email():
    assert te.member_name({"displayName": " Alice "}) == "Alice"
    assert te.member_name({"email": "a@b.de"}) == "a@b.de"
    assert te.member_name({}) == ""


def test_chat_title_oneonone_uses_other_member():
    chat = {"chatType": "oneOnOne", "topic": "wird ignoriert", "id": "c1",
            "members": [{"userId": "me", "displayName": "Ich"},
                        {"userId": "u2", "displayName": "Alice Example"}]}
    assert te.chat_title(FakeGraph(), chat, "me") == "Alice Example"


def test_chat_title_oneonone_without_partner_is_unknown():
    chat = {"chatType": "oneOnOne", "id": "c1",
            "members": [{"userId": "me", "displayName": "Ich"}]}
    assert te.chat_title(FakeGraph(), chat, "me") == "Unbekannt"


def test_chat_title_group_prefers_topic():
    chat = {"chatType": "group", "topic": "Projekt X", "id": "c1", "members": []}
    assert te.chat_title(FakeGraph(), chat, "me") == "Projekt X"


def test_chat_title_group_joins_members_and_truncates():
    members = [{"userId": f"u{i}", "displayName": f"P{i}"} for i in range(7)]
    chat = {"chatType": "group", "topic": None, "id": "c1", "members": members}
    title = te.chat_title(FakeGraph(), chat, "me")
    assert title == "P0, P1, P2, P3, P4…"   # max. 5 Namen plus Ellipse


def test_chat_title_loads_members_when_missing():
    chat = {"chatType": "oneOnOne", "id": "c9"}
    graph = FakeGraph(pages={f"{GRAPH}/me/chats/c9/members": [
        {"userId": "me", "displayName": "Ich"},
        {"userId": "u2", "displayName": "Bob"}]})
    assert te.chat_title(graph, chat, "me") == "Bob"
    assert graph.paged_calls == [f"{GRAPH}/me/chats/c9/members"]


def test_chat_title_member_load_failure_falls_back():
    chat = {"chatType": "group", "topic": None, "id": "c9"}
    graph = FakeGraph(pages={f"{GRAPH}/me/chats/c9/members": RuntimeError("nope")})
    assert te.chat_title(graph, chat, "me") == "group"   # topic or ctype or "Chat"


# --------------------------------------------------------------------------
# Rendering: Anhänge, Reaktionen, Nachrichten, Konversation
# --------------------------------------------------------------------------
def test_render_attachments_links_and_escapes():
    out = te.render_attachments([
        {"name": "Plan <Q3>.docx", "contentUrl": "https://x/a?b=1&c=2"},
        {"contentType": "reference"},
    ])
    assert 'href="https://x/a?b=1&amp;c=2"' in out
    assert "Plan &lt;Q3&gt;.docx" in out
    assert "reference" in out          # ohne URL nur der Name/Typ
    assert te.render_attachments([]) == ""
    assert te.render_attachments(None) == ""


def test_render_reactions_counts_types():
    out = te.render_reactions([{"reactionType": "like"}, {"reactionType": "like"},
                               {"reactionType": "heart"}])
    assert "like ×2" in out and "heart ×1" in out
    assert te.render_reactions([]) == ""
    assert te.render_reactions(None) == ""


def test_render_message_text_is_escaped():
    m = _msg("Alice <X>", "<b>kein html</b>", "2025-06-01T09:30:00Z")
    out = te.render_message(m)
    assert 'class="msg text"' in out           # Textnachricht -> pre-wrap-Klasse
    assert "Alice &lt;X&gt;" in out            # Name escaped
    assert "&lt;b&gt;kein html&lt;/b&gt;" in out
    assert TIME_RE.search(out)                 # lokale Zeit gerendert


def test_render_message_html_is_cleaned():
    m = _msg("Bob", '<div onclick="evil()">Hi</div><script>x()</script>',
             "2025-06-01T09:30:00Z", ctype="html")
    out = te.render_message(m)
    assert "onclick" not in out and "script" not in out
    assert ">Hi</div>" in out                  # HTML bleibt ansonsten erhalten


def test_render_message_deleted_and_subject_and_reply():
    m = _msg("Bob", "weg", "2025-06-01T09:30:00Z",
             deletedDateTime="2025-06-02T00:00:00Z", subject="Thema <1>")
    out = te.render_message(m, is_reply=True)
    assert "[gelöscht]" in out
    assert "weg" not in out
    assert "<strong>Thema &lt;1&gt;</strong>" in out
    assert 'class="msg reply"' in out


def test_render_message_from_application_and_unknown():
    m = _msg("x", "hi", "2025-06-01T09:30:00Z")
    m["from"] = {"application": {"displayName": "Ein Bot"}}
    assert '<span class="name">Ein Bot</span>' in te.render_message(m)
    m["from"] = None
    assert "Unbekannt" in te.render_message(m)


def test_render_message_system_event_variants():
    sys_msg = {"messageType": "systemEventMessage",
               "createdDateTime": "2025-06-01T09:30:00Z",
               "body": {"content": "<p>Alice wurde hinzugefügt</p>"}}
    out = te.render_message(sys_msg)
    assert 'class="sys"' in out
    assert "Alice wurde hinzugefügt" in out

    sys_msg["body"] = {"content": ""}
    sys_msg["eventDetail"] = {"@odata.type": "#microsoft.graph.membersAddedEventMessageDetail"}
    assert "membersAddedEventMessageDetail" in te.render_message(sys_msg)

    sys_msg["eventDetail"] = {}
    assert "Systemnachricht" in te.render_message(sys_msg)


def test_render_conversation_escapes_and_marks_failed_images():
    html = te.render_conversation("Titel <X>", "1:1-Chat", "3 Nachrichten & mehr",
                                  ["<div>a</div>", f'<img src="{te.IMG_PLACEHOLDER}">',
                                   f'<img src="{te.IMG_PLACEHOLDER}">'])
    assert "<title>Titel &lt;X&gt;</title>" in html
    assert "3 Nachrichten &amp; mehr" in html
    assert "2 Bild(er) konnten nicht geladen werden" in html


def test_render_conversation_empty_blocks():
    html = te.render_conversation("T", "S", "M", [])
    assert "Keine Nachrichten." in html
    assert "konnten nicht geladen" not in html


def test_render_blocks_counts_images():
    msgs = [_msg("A", "eins", "2025-06-01T09:00:00Z"),
            _msg("B", "zwei", "2025-06-01T09:05:00Z")]
    blocks, nimg = te.render_blocks(msgs, "lbl")
    assert len(blocks) == 2 and nimg == 0
    assert "eins" in blocks[0] and "zwei" in blocks[1]


# --------------------------------------------------------------------------
# embed_hosted_images
# --------------------------------------------------------------------------
HOSTED_URL = "https://graph.microsoft.com/v1.0/chats/1/messages/2/hostedContents/abc/$value"


class FakeImgClient:
    def __init__(self, error=None):
        self.error = error
        self.calls = 0

    def get_bytes(self, url):
        self.calls += 1
        if self.error:
            raise self.error
        return b"BILD", "image/png"


def test_embed_hosted_images_noop_without_client(monkeypatch):
    monkeypatch.setattr(te, "_client", None)
    html = f'<img src="{HOSTED_URL}">'
    assert te.embed_hosted_images(html) == html


def test_embed_hosted_images_inlines_as_data_uri(monkeypatch):
    client = FakeImgClient()
    monkeypatch.setattr(te, "_client", client)
    monkeypatch.setattr(te, "IMGCACHE_DIR", None)
    counter = [0]
    out = te.embed_hosted_images(f'<img src="{HOSTED_URL}"> und <a href="https://example.com">x</a>',
                                 counter)
    assert "data:image/png;base64,QklMRA==" in out   # base64("BILD")
    assert "hostedContents" not in out
    assert "https://example.com" in out              # fremde URLs bleiben stehen
    assert counter == [1] and client.calls == 1


def test_embed_hosted_images_failure_yields_placeholder(monkeypatch):
    monkeypatch.setattr(te, "_client", FakeImgClient(error=te.ImageUnavailable("502")))
    monkeypatch.setattr(te, "IMGCACHE_DIR", None)
    counter = [0]
    out = te.embed_hosted_images(f'<img src="{HOSTED_URL}">', counter)
    assert te.IMG_PLACEHOLDER in out
    assert counter == [0]   # fehlgeschlagene Bilder zählen nicht


def test_embed_hosted_images_token_expired_propagates(monkeypatch):
    monkeypatch.setattr(te, "_client", FakeImgClient(error=te.TokenExpired()))
    monkeypatch.setattr(te, "IMGCACHE_DIR", None)
    with pytest.raises(te.TokenExpired):
        te.embed_hosted_images(f'<img src="{HOSTED_URL}">')


def test_embed_hosted_images_uses_cache(monkeypatch, tmp_path):
    client = FakeImgClient()
    monkeypatch.setattr(te, "_client", client)
    monkeypatch.setattr(te, "IMGCACHE_DIR", tmp_path)
    html = f'<img src="{HOSTED_URL}">'
    out1 = te.embed_hosted_images(html)
    out2 = te.embed_hosted_images(html)   # zweiter Lauf: Cache-Treffer, kein Download
    assert out1 == out2
    assert client.calls == 1
    assert len(list(tmp_path.iterdir())) == 1


# --------------------------------------------------------------------------
# Fortschritt: load_state / save_state / already_done / get_record / cleanup_old
# --------------------------------------------------------------------------
def test_load_state_defaults_and_roundtrip(tmp_path):
    state = te.load_state(tmp_path)
    assert state == {"version": 1, "conversations": {}}
    state["conversations"]["k"] = {"done": True, "rel": "1on1/a.html"}
    te.save_state(tmp_path, state)
    assert te.load_state(tmp_path) == state
    assert not (tmp_path / (te.STATE_FILE + ".tmp")).exists()   # atomarer Austausch


def test_load_state_ignores_corrupt_file(tmp_path):
    (tmp_path / te.STATE_FILE).write_text("{kaputt", encoding="utf-8")
    assert te.load_state(tmp_path) == {"version": 1, "conversations": {}}
    (tmp_path / te.STATE_FILE).write_text('["falsche form"]', encoding="utf-8")
    assert te.load_state(tmp_path) == {"version": 1, "conversations": {}}


def test_already_done_requires_record_and_file(tmp_path):
    state = {"version": 1, "conversations": {
        "a": {"done": True, "rel": "1on1/a.html"},
        "b": {"done": False, "rel": "1on1/b.html"},
    }}
    assert not te.already_done(tmp_path, state, "fehlt")
    assert not te.already_done(tmp_path, state, "b")       # nicht fertig
    assert not te.already_done(tmp_path, state, "a")       # Datei fehlt noch
    (tmp_path / "1on1").mkdir()
    (tmp_path / "1on1" / "a.html").write_text("x", encoding="utf-8")
    assert te.already_done(tmp_path, state, "a")


def test_get_record_variants(tmp_path):
    state = {"version": 1, "conversations": {
        "done": {"done": True, "rel": "1on1/a.html", "last_activity": "2025-06-01T00:00:00Z"},
        "empty": {"done": True, "rel": None, "empty": True},
        "open": {"done": False, "rel": "1on1/x.html"},
    }}
    assert te.get_record(tmp_path, state, "fehlt") is None
    assert te.get_record(tmp_path, state, "open") is None
    assert te.get_record(tmp_path, state, "done") is None    # Datei fehlt
    (tmp_path / "1on1").mkdir()
    (tmp_path / "1on1" / "a.html").write_text("x", encoding="utf-8")
    assert te.get_record(tmp_path, state, "done")["last_activity"] == "2025-06-01T00:00:00Z"
    # leere Chats gelten ohne Datei als gültiger Status
    assert te.get_record(tmp_path, state, "empty")["empty"] is True


def test_cleanup_old_removes_renamed_file_only(tmp_path):
    (tmp_path / "1on1").mkdir()
    old = tmp_path / "1on1" / "Unbekannt__1234.html"
    old.write_text("alt", encoding="utf-8")
    te.cleanup_old(tmp_path, {"rel": "1on1/Unbekannt__1234.html"}, "1on1/Unbekannt__1234.html")
    assert old.exists()   # gleicher Name -> nichts löschen
    te.cleanup_old(tmp_path, {"rel": "1on1/Unbekannt__1234.html"}, "1on1/Alice__1234.html")
    assert not old.exists()
    te.cleanup_old(tmp_path, {"rel": "1on1/weg.html"}, "1on1/neu.html")   # fehlend -> kein Fehler
    te.cleanup_old(tmp_path, None, "1on1/neu.html")


def test_record_done_persists_record(tmp_path):
    state = te.load_state(tmp_path)
    te.record_done(tmp_path, state, "k1", "1on1", "Alice", "1on1/a.html", 7,
                   last_activity="2025-06-01T09:30:00Z")
    rec = json.loads((tmp_path / te.STATE_FILE).read_text(encoding="utf-8"))["conversations"]["k1"]
    assert rec["category"] == "1on1" and rec["title"] == "Alice"
    assert rec["rel"] == "1on1/a.html" and rec["count"] == 7
    assert rec["done"] is True and rec["empty"] is False
    assert rec["last_activity"] == "2025-06-01T09:30:00Z"


def test_write_index_groups_and_skips_empty(tmp_path):
    state = {"version": 1, "conversations": {
        "a": {"done": True, "empty": False, "category": "1on1",
              "title": "Alice & Bob", "rel": "1on1/a.html", "count": 3},
        "b": {"done": True, "empty": False, "category": "channels",
              "title": "Team / Allgemein", "rel": "channels/T/a.html", "count": 12},
        "c": {"done": True, "empty": True, "category": "group",
              "title": "Leer", "rel": None, "count": 2},
        "d": {"done": False, "category": "group", "title": "Offen",
              "rel": "group/x.html", "count": 1},
    }}
    te.write_index(tmp_path, state)
    html = (tmp_path / "index.html").read_text(encoding="utf-8")
    assert "1:1-Chats" in html and "(1)" in html
    assert "Alice &amp; Bob" in html                 # Titel wird escaped
    assert 'href="1on1/a.html"' in html
    assert "3 Nachrichten" in html
    assert "Team-Kanäle" in html and "12 Nachrichten" in html
    assert "Leer" not in html and "Offen" not in html
    assert "Gruppenchats" not in html                # keine sichtbaren Einträge -> Gruppe fehlt


# --------------------------------------------------------------------------
# Export einer Konversation (Chat und Kanal) mit gefaktem Graph
# --------------------------------------------------------------------------
def _chat_fixture(chat_id="c1"):
    chat = {"id": chat_id, "chatType": "oneOnOne",
            "members": [{"userId": "me", "displayName": "Ich"},
                        {"userId": "u2", "displayName": "Alice Example"}]}
    msgs = [_msg("Alice Example", "zweite", "2025-06-02T08:00:00Z"),
            _msg("Ich", "erste", "2025-06-01T09:30:00Z")]
    graph = FakeGraph(pages={f"{GRAPH}/me/chats/{chat_id}/messages": msgs})
    return chat, graph


def test_export_one_chat_writes_html_and_state(tmp_path):
    chat, graph = _chat_fixture()
    state = te.load_state(tmp_path)
    status, folder, title, count, secs = te.export_one_chat(graph, tmp_path, state, "me", chat)
    assert (status, folder, title, count) == ("new", "1on1", "Alice Example", 2)

    fname = f"Alice Example__{te.short_id('c1')}.html"
    html = (tmp_path / "1on1" / fname).read_text(encoding="utf-8")
    idx_erste, idx_zweite = html.index("erste"), html.index("zweite")
    assert idx_erste < idx_zweite            # chronologisch sortiert
    assert "Alice Example" in html and "Chat-ID c1" in html

    rec = state["conversations"]["c1"]
    assert rec["rel"] == f"1on1/{fname}"
    assert rec["last_activity"] == "2025-06-02T08:00:00Z"
    assert rec["empty"] is False


def test_export_one_chat_second_run_reports_updated_and_renames(tmp_path):
    chat, graph = _chat_fixture()
    state = te.load_state(tmp_path)
    te.export_one_chat(graph, tmp_path, state, "me", chat)
    # Name ändert sich (z. B. 'Unbekannt' -> echter Name simuliert durch Member-Wechsel)
    old_rel = state["conversations"]["c1"]["rel"]
    chat["members"][1]["displayName"] = "Alice Umbenannt"
    status, _, title, _, _ = te.export_one_chat(graph, tmp_path, state, "me", chat)
    assert status == "updated" and title == "Alice Umbenannt"
    assert not (tmp_path / old_rel).exists()   # Altdatei wurde aufgeräumt
    assert (tmp_path / state["conversations"]["c1"]["rel"]).exists()


def test_export_one_chat_only_system_messages_is_empty(tmp_path):
    chat = {"id": "c2", "chatType": "meeting", "topic": "Standup"}
    sysmsg = {"messageType": "systemEventMessage",
              "createdDateTime": "2025-06-03T07:00:00Z",
              "body": {"content": "Anruf beendet"}}
    graph = FakeGraph(pages={f"{GRAPH}/me/chats/c2/messages": [sysmsg]})
    state = te.load_state(tmp_path)
    status, folder, title, count, _ = te.export_one_chat(graph, tmp_path, state, "me", chat)
    assert (status, folder, title, count) == ("empty", "meeting", "Standup", 1)
    rec = state["conversations"]["c2"]
    assert rec["empty"] is True and rec["rel"] is None
    assert rec["last_activity"] == "2025-06-03T07:00:00Z"
    assert not (tmp_path / "meeting").exists()   # keine Datei geschrieben


def _channel_fixture(reply_ts="2025-06-05T10:00:00Z"):
    team = {"id": "t1", "displayName": "Team Rakete"}
    ch = {"id": "k1", "displayName": "Allgemein", "membershipType": "standard"}
    root = _msg("Alice", "Wurzelpost", "2025-06-04T09:00:00Z")
    root["replies"] = [_msg("Bob", "Antwort", reply_ts)]
    graph = FakeGraph(pages={f"{GRAPH}/teams/t1/channels/k1/messages": [root]})
    return team, ch, graph


def test_export_one_channel_writes_nested_replies(tmp_path):
    team, ch, graph = _channel_fixture()
    state = te.load_state(tmp_path)
    status, cat, title, count, _ = te.export_one_channel(graph, tmp_path, state, team, ch)
    assert (status, cat, title, count) == ("new", "channels", "Team Rakete / Allgemein", 2)

    fname = f"Allgemein__{te.short_id('k1')}.html"
    html = (tmp_path / "channels" / "Team Rakete" / fname).read_text(encoding="utf-8")
    assert "Wurzelpost" in html and "Antwort" in html
    assert 'class="msg reply' in html                    # Antwort eingerückt
    assert "2 Nachrichten (inkl. Antworten)" in html
    assert state["conversations"]["ch:k1"]["last_activity"] == "2025-06-05T10:00:00Z"


def test_export_one_channel_unchanged_on_second_run(tmp_path):
    team, ch, graph = _channel_fixture()
    state = te.load_state(tmp_path)
    te.export_one_channel(graph, tmp_path, state, team, ch)
    status, _, _, count, _ = te.export_one_channel(graph, tmp_path, state, team, ch)
    assert status == "unchanged" and count == 2


def test_export_one_channel_updates_on_new_reply(tmp_path):
    team, ch, graph = _channel_fixture()
    state = te.load_state(tmp_path)
    te.export_one_channel(graph, tmp_path, state, team, ch)
    team, ch, graph = _channel_fixture(reply_ts="2025-06-06T12:00:00Z")   # neuere Antwort
    status, _, _, _, _ = te.export_one_channel(graph, tmp_path, state, team, ch)
    assert status == "updated"
    assert state["conversations"]["ch:k1"]["last_activity"] == "2025-06-06T12:00:00Z"


# --------------------------------------------------------------------------
# Job-Aufbau: build_chat_jobs / build_channel_jobs
# --------------------------------------------------------------------------
def test_build_chat_jobs_new_updated_and_skipped(tmp_path):
    chats = [
        {"id": "neu", "chatType": "oneOnOne",
         "lastMessagePreview": {"createdDateTime": "2025-06-01T00:00:00Z"}},
        {"id": "upd", "chatType": "group",
         "lastMessagePreview": {"createdDateTime": "2025-06-05T00:00:00Z"}},
        {"id": "alt", "chatType": "meeting",
         "lastMessagePreview": {"createdDateTime": "2025-06-01T00:00:00Z"}},
        {"id": "fremd", "chatType": "unknownType",
         "lastMessagePreview": {"createdDateTime": "2025-06-05T00:00:00Z"}},
    ]
    graph = FakeGraph(pages={f"{GRAPH}/me/chats": chats})
    state = {"version": 1, "conversations": {
        "upd": {"done": True, "rel": "group/upd.html", "count": 1, "empty": False,
                "last_activity": "2025-06-02T00:00:00Z"},
        "alt": {"done": True, "rel": "meeting/alt.html", "count": 1, "empty": False,
                "last_activity": "2025-06-02T00:00:00Z"},
    }}
    for rel in ("group/upd.html", "meeting/alt.html"):
        (tmp_path / rel).parent.mkdir(parents=True, exist_ok=True)
        (tmp_path / rel).write_text("x", encoding="utf-8")

    stats = {"new": 0, "updated": 0, "skipped": 0, "empty": 0}
    jobs = te.build_chat_jobs(graph, tmp_path, state, stats, "me",
                              {"1on1", "group", "meeting"})
    assert [(k, c["id"]) for k, c, _ in jobs] == [("chat", "neu"), ("chat", "upd")]
    assert stats["skipped"] == 1   # 'alt' unverändert; 'fremd' fällt aus den Kategorien


def test_build_channel_jobs_refresh_mode_and_error_team(monkeypatch, tmp_path):
    monkeypatch.setattr(te, "REFRESH_CHANNELS", True)
    graph = FakeGraph(pages={
        f"{GRAPH}/teams/t1/channels": [{"id": "k1", "displayName": "A"},
                                       {"id": "k2", "displayName": "B"}],
        f"{GRAPH}/teams/t2/channels": RuntimeError("keine Rechte"),
    })
    stats = {"new": 0, "updated": 0, "skipped": 0, "empty": 0}
    jobs = te.build_channel_jobs(graph, tmp_path, {"version": 1, "conversations": {}}, stats,
                                 [{"id": "t1", "displayName": "T1"},
                                  {"id": "t2", "displayName": "T2"}])
    # Fehler-Team wird übersprungen, Refresh-Modus prüft alle Kanäle erneut
    assert [(k, ch["id"]) for k, _t, ch in jobs] == [("channel", "k1"), ("channel", "k2")]


def test_build_channel_jobs_without_refresh_skips_done(monkeypatch, tmp_path):
    monkeypatch.setattr(te, "REFRESH_CHANNELS", False)
    graph = FakeGraph(pages={
        f"{GRAPH}/teams/t1/channels": [{"id": "k1", "displayName": "A"},
                                       {"id": "k2", "displayName": "B"}]})
    state = {"version": 1, "conversations": {
        "ch:k1": {"done": True, "rel": "channels/T1/a.html"}}}
    (tmp_path / "channels" / "T1").mkdir(parents=True)
    (tmp_path / "channels" / "T1" / "a.html").write_text("x", encoding="utf-8")
    stats = {"new": 0, "updated": 0, "skipped": 0, "empty": 0}
    jobs = te.build_channel_jobs(graph, tmp_path, state, stats,
                                 [{"id": "t1", "displayName": "T1"}])
    assert [ch["id"] for _k, _t, ch in jobs] == ["k2"]
    assert stats["skipped"] == 1


def test_build_channel_jobs_token_expired_propagates(tmp_path):
    graph = FakeGraph(pages={f"{GRAPH}/teams/t1/channels": te.TokenExpired()})
    stats = {"new": 0, "updated": 0, "skipped": 0, "empty": 0}
    with pytest.raises(te.TokenExpired):
        te.build_channel_jobs(graph, tmp_path, {"version": 1, "conversations": {}}, stats,
                              [{"id": "t1", "displayName": "T1"}])


# --------------------------------------------------------------------------
# make_runner / run_parallel
# --------------------------------------------------------------------------
def test_make_runner_dispatches_and_maps_errors(monkeypatch, tmp_path):
    sentinel = ("new", "1on1", "Alice", 2, 0.1)
    monkeypatch.setattr(te, "export_one_chat",
                        lambda graph, out, state, my_id, chat: sentinel)
    run = te.make_runner("g", tmp_path, {}, "me", "chat", {"id": "c"}, None)
    assert run() == sentinel

    def boom(*a, **kw):
        raise ValueError("kaputt")
    monkeypatch.setattr(te, "export_one_channel", boom)
    run = te.make_runner("g", tmp_path, {}, "me", "channel", {"id": "t"}, {"id": "k"})
    assert run() == ("error", None, "kaputt", 0, 0.0)
    assert not te.STOP.is_set()


def test_make_runner_token_expired_sets_stop(monkeypatch, tmp_path):
    def expired(*a, **kw):
        raise te.TokenExpired()
    monkeypatch.setattr(te, "export_one_chat", expired)
    run = te.make_runner("g", tmp_path, {}, "me", "chat", {"id": "c"}, None)
    assert run() == ("expired", None, None, 0, 0.0)
    assert te.STOP.is_set()
    # nach gesetztem STOP starten weitere Runner gar nicht mehr
    run2 = te.make_runner("g", tmp_path, {}, "me", "chat", {"id": "c2"}, None)
    assert run2() == ("stopped", None, None, 0, 0.0)


def test_run_parallel_counts_statuses():
    stats = {"new": 0, "updated": 0, "skipped": 0, "empty": 0}
    runners = [
        lambda: ("new", "1on1", "A", 1, 0.5),
        lambda: ("updated", "group", "B", 2, 1.5),
        lambda: ("unchanged", "channels", "C", 3, 0.1),
        lambda: ("empty", "meeting", "D", 0, 0.1),
        lambda: ("error", None, "kaputt", 0, 0.0),
        lambda: ("stopped", None, None, 0, 0.0),
    ]
    assert te.run_parallel(runners, stats, workers=2) == "done"
    assert stats == {"new": 1, "updated": 1, "skipped": 1, "empty": 1}


def test_run_parallel_reports_expired():
    stats = {"new": 0, "updated": 0, "skipped": 0, "empty": 0}
    runners = [lambda: ("new", "1on1", "A", 1, 0.5),
               lambda: ("expired", None, None, 0, 0.0)]
    assert te.run_parallel(runners, stats, workers=1) == "expired"
    assert stats["new"] == 1


def test_run_parallel_catches_raising_runner():
    def boom():
        raise RuntimeError("crash im Worker")
    stats = {"new": 0, "updated": 0, "skipped": 0, "empty": 0}
    assert te.run_parallel([boom], stats, workers=1) == "done"
    assert stats == {"new": 0, "updated": 0, "skipped": 0, "empty": 0}


def test_run_parallel_empty_list_is_done():
    assert te.run_parallel([], {}, workers=4) == "done"
