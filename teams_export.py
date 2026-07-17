#!/usr/bin/env python3
"""
Teams/Chat-Export über Microsoft Graph (delegiert, kein Admin nötig).

Beim Start fragt das Skript interaktiv ab:
  1) was exportiert werden soll (1:1 / Gruppe / Meeting / Kanäle, Mehrfachauswahl)
  2) falls Kanäle gewählt: welche Teams (jeweils alle Kanäle darin)

Exportiert eine HTML pro Chat bzw. pro Kanal, unterteilt in:
    1on1/  group/  meeting/  channels/<Team>/   + index.html

PARALLEL: mehrere Chats/Kanäle gleichzeitig (Standard 4, per Env EXPORT_WORKERS).
  Teams-Throttling: ~1 Anfrage/s je einzelnem Chat oder Kanal, 4/s je Team, und
  Chats liegen im Postfach (4 gleichzeitige Anfragen). Deshalb wird ÜBER
  Konversationen hinweg parallelisiert, nicht innerhalb einer. Kanäle holen ihre
  Antworten via $expand=replies inline (bis 1000, dann replies@odata.nextLink),
  das spart pro Kanal sehr viele Aufrufe. Drosselung (429) wird per Retry-After
  abgefangen.

Setup:   pip install msal requests
Start:   python3 teams_export.py [ausgabe-ordner] [-default]
         -default überspringt die Abfrage und nutzt die Vorgabe (1, 2, 3 =
         1:1-, Gruppen- und Meeting-Chats, keine Kanäle).

Token-Modus (wenn der Tenant für neue Apps "Approval required" verlangt):
    Access Token im Graph Explorer holen (Chat.Read bzw. ChannelMessage.Read.All
    zugestimmt), in gx_token.txt neben dieses Skript legen ODER
    export GRAPH_TOKEN="eyJ0…"

Resume / inkrementell: export_state.json im Ausgabeordner merkt sich pro
    Konversation die letzte Aktivität. Bei erneutem Lauf (z. B. per Scheduler)
    werden Chats mit NEUEN Nachrichten automatisch neu exportiert, unveränderte
    übersprungen. Kanäle werden auf Aktualität geprüft und nur bei Änderung neu
    geschrieben (REFRESH_CHANNELS=0 schaltet das ab). Token tot -> frischen Token
    setzen, neu starten. Kompletter Neu-Export: Datei (oder Ordner) löschen.

Schalter unten: WORKERS, USE_DEVICE_CODE, EMBED_IMAGES, REFRESH_CHANNELS (env).
"""

import os
import sys
import re
import json
import time
import base64
import hashlib
import threading
import html as html_lib
from datetime import datetime
from pathlib import Path
from collections import defaultdict, Counter
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    import msal
    import requests
except ImportError:
    print("Fehlende Pakete. Bitte installieren:  pip install msal requests")
    raise SystemExit(1)

# Auf Windows nutzt die Konsole standardmäßig eine Legacy-Codepage (z. B. cp1252),
# und bei Umleitung in eine Datei (python … > log.txt) die Locale-Kodierung. Beides
# lässt print() an Unicode-Zeichen wie →, ✓, · oder Emoji mit UnicodeEncodeError
# scheitern und bricht den Export ab. UTF-8 erzwingen (auf macOS/Linux ein No-op).
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

# ---------------------------------------------------------------------------
# Konfiguration
# ---------------------------------------------------------------------------
CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"  # Microsoft Graph Command Line Tools
TENANT = "organizations"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT}"

GRAPH = "https://graph.microsoft.com/v1.0"
RES = "https://graph.microsoft.com/"
SCOPES_CHAT = [RES + "Chat.Read", RES + "User.Read"]
SCOPES_FULL = SCOPES_CHAT + [
    RES + "ChannelMessage.Read.All",
    RES + "Team.ReadBasic.All",
    RES + "Channel.ReadBasic.All",
]

USE_DEVICE_CODE = False     # True = Code-Login statt Browser
EMBED_IMAGES = True         # Inline-Bilder als base64 einbetten
WORKERS = 4                 # parallele Konversationen (sinnvoll: 4; per Env EXPORT_WORKERS)
PAGE = 50                   # $top (Graph-Maximum für Nachrichten)
OUT_ROOT = "teams_export"    # fester Ordner -> Resume über mehrere Läufe hinweg
STATE_FILE = "export_state.json"

# Inkrementelle Läufe (z. B. per Scheduler): Chats werden bei neuen Nachrichten
# automatisch neu exportiert (günstig über lastMessagePreview erkannt). Kanäle
# bieten keinen günstigen Änderungs-Indikator (neue Antworten in alten Threads),
# daher werden gewählte Kanäle pro Lauf neu geholt und nur bei Änderung neu
# geschrieben. Mit REFRESH_CHANNELS=0 abschaltbar (Kanäle dann nur einmalig).
REFRESH_CHANNELS = os.environ.get("REFRESH_CHANNELS", "1") not in ("0", "false", "False")

# Heruntergeladene Inline-Bilder zwischenspeichern (Ordner .imgcache). Bei erneutem
# Export eines Chats werden so nur NEUE Bilder geladen statt aller. Kostet zusätzlichen
# Plattenplatz (Bilder liegen dann doppelt: im Cache und eingebettet im HTML).
# Mit CACHE_IMAGES=0 abschaltbar.
CACHE_IMAGES = os.environ.get("CACHE_IMAGES", "1") not in ("0", "false", "False")

# Chats, die NUR System-/Event-Nachrichten enthalten (Beitritte, Anrufe, Mitglieder-
# Änderungen, …) und keine echte Nachricht, werden standardmäßig NICHT exportiert
# und nicht in den Index aufgenommen. Mit SKIP_EMPTY_CHATS=0 doch exportieren.
SKIP_EMPTY_CHATS = os.environ.get("SKIP_EMPTY_CHATS", "1") not in ("0", "false", "False")

TYPEMAP = {"oneOnOne": "1on1", "group": "group", "meeting": "meeting"}
SUBNAME = {"1on1": "1:1-Chat", "group": "Gruppenchat",
           "meeting": "Meeting-Chat", "other": "Chat"}

SESSION = requests.Session()                 # Keep-Alive/Connection-Pooling
GATE = threading.BoundedSemaphore(WORKERS)   # gleichzeitige Postfach-/Graph-Calls <= WORKERS
STOP = threading.Event()                      # Signal: Token tot -> nichts Neues mehr starten
STATE_LOCK = threading.Lock()                 # serialisiert Schreiben der Fortschrittsdatei
PRINT_LOCK = threading.Lock()                 # saubere, nicht ineinander laufende Fortschrittszeilen

_client = None       # wird in main() gesetzt (für die Bild-Einbettung)
IMGCACHE_DIR = None  # wird in main() gesetzt, wenn CACHE_IMAGES aktiv


def log(msg):
    """Threadsichere Fortschrittsausgabe (sofort sichtbar)."""
    with PRINT_LOCK:
        print(msg, flush=True)


class TokenExpired(RuntimeError):
    """Signalisiert einen 401 im Token-Modus (kein Refresh möglich)."""


class ImageUnavailable(RuntimeError):
    """Ein Inline-Bild (hostedContent) ist nicht herunterladbar (z. B. 502).
    Wird NICHT erneut versucht – der Export läuft mit Platzhalter weiter."""


# ---------------------------------------------------------------------------
# Graph-Client: Auth, Retry, Paging
# ---------------------------------------------------------------------------
class Graph:
    def __init__(self, want_channels):
        self.app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
        self.want_channels = want_channels
        self.account = None
        self.scopes = None
        self.channels_enabled = False
        self.token = None
        self._refresh_lock = threading.Lock()
        self._login()

    def _acquire(self, scopes):
        if USE_DEVICE_CODE:
            flow = self.app.initiate_device_flow(scopes=scopes)
            if "user_code" not in flow:
                raise RuntimeError(f"Device-Flow fehlgeschlagen: {flow.get('error_description')}")
            print("\n" + flow["message"] + "\n")
            return self.app.acquire_token_by_device_flow(flow)
        return self.app.acquire_token_interactive(scopes=scopes, prompt="select_account")

    def _login(self):
        if self.want_channels:
            print("Anmeldung – fordere Zugriff auf Chats + Kanäle an…")
            res = self._acquire(SCOPES_FULL)
            if res and "access_token" in res:
                self.scopes, self.channels_enabled = SCOPES_FULL, True
            else:
                err = (res or {}).get("error_description", "") or ""
                print("Kanal-Zugriff nicht gewährt – ChannelMessage.Read.All "
                      "erfordert evtl. Admin-Consent.")
                if err:
                    print(f"  Details: {err.splitlines()[0]}")
                print("Melde mit reinem Chat-Zugriff an…")
                res = self._acquire(SCOPES_CHAT)
                if not res or "access_token" not in res:
                    raise SystemExit("Anmeldung fehlgeschlagen: "
                                     + (res or {}).get("error_description", "unbekannt"))
                self.scopes, self.channels_enabled = SCOPES_CHAT, False
        else:
            print("Anmeldung – fordere Chat-Zugriff an…")
            res = self._acquire(SCOPES_CHAT)
            if not res or "access_token" not in res:
                raise SystemExit("Anmeldung fehlgeschlagen: "
                                 + (res or {}).get("error_description", "unbekannt"))
            self.scopes, self.channels_enabled = SCOPES_CHAT, False

        accs = self.app.get_accounts()
        self.account = accs[0] if accs else None
        self.token = res["access_token"]

    def _refresh(self):
        with self._refresh_lock:   # nur ein Thread erneuert gleichzeitig
            res = self.app.acquire_token_silent(self.scopes, account=self.account) if self.account else None
            if not res or "access_token" not in res:
                res = self._acquire(self.scopes)
            if not res or "access_token" not in res:
                raise SystemExit("Token-Erneuerung fehlgeschlagen.")
            self.token = res["access_token"]

    def _headers(self):
        return {"Authorization": f"Bearer {self.token}"}

    def get(self, url, params=None):
        for attempt in range(6):
            with GATE:   # nur das eigentliche Request zählt gegen das Limit
                r = SESSION.get(url, headers=self._headers(), params=params, timeout=60)
            if r.status_code == 401:
                self._refresh()
                continue
            if r.status_code == 429 or 500 <= r.status_code < 600:
                ra = r.headers.get("Retry-After")
                w = min(int(ra) if ra and ra.isdigit() else 2 ** attempt, 60)
                print(f"    … HTTP {r.status_code}, warte {w}s (Drosselung/Server)")
                time.sleep(w)   # Pause OHNE belegten Slot
                continue
            r.raise_for_status()
            return r.json()
        raise RuntimeError(f"Zu viele Fehlversuche: {url}")

    def get_bytes(self, url):
        for attempt in range(4):
            with GATE:
                r = SESSION.get(url, headers=self._headers(), timeout=60)
            if r.status_code == 401:
                self._refresh()
                continue
            if r.status_code == 429:          # echte Drosselung -> abwarten
                ra = r.headers.get("Retry-After")
                w = min(int(ra) if ra and ra.isdigit() else 2 ** attempt, 30)
                print(f"    … HTTP 429 (Bild), warte {w}s")
                time.sleep(w)
                continue
            if 500 <= r.status_code < 600:    # Serverfehler -> Bild ist nicht ladbar
                raise ImageUnavailable(r.status_code)
            r.raise_for_status()
            return r.content, r.headers.get("Content-Type", "")
        raise ImageUnavailable("429")

    def paged(self, url, params=None):
        data = self.get(url, params)
        while True:
            for item in data.get("value", []):
                yield item
            nxt = data.get("@odata.nextLink")
            if not nxt:
                break
            data = self.get(nxt)   # nextLink ist absolut & enthält Parameter


# ---------------------------------------------------------------------------
# Token-Modus: vorhandenen Access Token nutzen (z. B. aus Graph Explorer)
# ---------------------------------------------------------------------------
def load_pasted_token():
    val = os.environ.get("GRAPH_TOKEN")
    if not val:
        p = Path("gx_token.txt")
        if p.exists():
            val = p.read_text(encoding="utf-8")
    if not val:
        return None
    val = val.strip().strip('"').strip("'").strip()
    if val.lower().startswith("bearer "):
        val = val[7:].strip()
    return val or None


class TokenClient:
    """Nutzt einen fertigen Bearer-Token; keine Anmeldung, kein Refresh."""

    def __init__(self, token, channels_enabled):
        self.token = token
        self.channels_enabled = channels_enabled
        self.account = None

    def _headers(self):
        return {"Authorization": f"Bearer {self.token}"}

    def _expired(self):
        raise TokenExpired()

    def get(self, url, params=None):
        for attempt in range(6):
            with GATE:
                r = SESSION.get(url, headers=self._headers(), params=params, timeout=60)
            if r.status_code == 401:
                self._expired()
            if r.status_code == 429 or 500 <= r.status_code < 600:
                ra = r.headers.get("Retry-After")
                w = min(int(ra) if ra and ra.isdigit() else 2 ** attempt, 60)
                print(f"    … HTTP {r.status_code}, warte {w}s (Drosselung/Server)")
                time.sleep(w)
                continue
            r.raise_for_status()
            return r.json()
        raise RuntimeError(f"Zu viele Fehlversuche: {url}")

    def get_bytes(self, url):
        for attempt in range(4):
            with GATE:
                r = SESSION.get(url, headers=self._headers(), timeout=60)
            if r.status_code == 401:
                self._expired()
            if r.status_code == 429:          # echte Drosselung -> abwarten
                ra = r.headers.get("Retry-After")
                w = min(int(ra) if ra and ra.isdigit() else 2 ** attempt, 30)
                print(f"    … HTTP 429 (Bild), warte {w}s")
                time.sleep(w)
                continue
            if 500 <= r.status_code < 600:    # Serverfehler -> Bild ist nicht ladbar
                raise ImageUnavailable(r.status_code)
            r.raise_for_status()
            return r.content, r.headers.get("Content-Type", "")
        raise ImageUnavailable("429")

    def paged(self, url, params=None):
        data = self.get(url, params)
        while True:
            for item in data.get("value", []):
                yield item
            nxt = data.get("@odata.nextLink")
            if not nxt:
                break
            data = self.get(nxt)


# ---------------------------------------------------------------------------
# Interaktive Abfragen
# ---------------------------------------------------------------------------
def _read(prompt):
    try:
        return input(prompt)
    except EOFError:
        return ""


def parse_indices(raw, n):
    out = []
    for tok in re.split(r"[\s,]+", raw.strip()):
        if tok.isdigit():
            v = int(tok)
            if 1 <= v <= n and v not in out:
                out.append(v)
    return out


def default_categories(options):
    """Standardauswahl: die ersten drei Kategorien (1:1-, Gruppen-, Meeting-Chats)."""
    return {k for k, _ in options[:3]}


def prompt_categories(options):
    if not sys.stdin.isatty():
        print("Kein interaktives Terminal – nutze die Standardauswahl (1, 2, 3).")
        return default_categories(options)
    print("Was möchtest du exportieren?")
    for i, (k, label) in enumerate(options, 1):
        print(f"  {i}) {label}")
    raw = _read("Auswahl (Zahlen kommagetrennt, Enter = 1,2,3): ").strip()
    if not raw:
        return default_categories(options)
    idxs = parse_indices(raw, len(options))
    if not idxs:
        print("Keine gültige Auswahl – nehme die Standardauswahl (1, 2, 3).")
        return default_categories(options)
    return {options[i - 1][0] for i in idxs}


def select_teams(graph):
    try:
        teams = list(graph.paged(f"{GRAPH}/me/joinedTeams", {"$top": PAGE}))
    except TokenExpired:
        raise
    except Exception as e:
        print(f"Teams konnten nicht geladen werden: {e}")
        return []
    teams.sort(key=lambda t: (t.get("displayName") or "").lower())
    if not teams:
        print("Keine Teams gefunden.")
        return []
    if not sys.stdin.isatty():
        print(f"Kein interaktives Terminal – exportiere alle {len(teams)} Teams.")
        return teams
    print(f"\n{len(teams)} Teams gefunden. Welche exportieren? (jeweils alle Kanäle)")
    for i, t in enumerate(teams, 1):
        print(f"  {i}) {t.get('displayName', 'Team')}")
    raw = _read("Auswahl (Zahlen kommagetrennt, Enter = alle): ").strip()
    if not raw:
        return teams
    idxs = parse_indices(raw, len(teams))
    if not idxs:
        print("Keine gültige Auswahl – nehme alle.")
        return teams
    return [teams[i - 1] for i in idxs]


# ---------------------------------------------------------------------------
# Helfer
# ---------------------------------------------------------------------------
def safe(name, maxlen=80):
    name = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", name or "").strip().strip(".")
    name = re.sub(r"\s+", " ", name)
    return name[:maxlen] or "unbenannt"


def short_id(s):
    return hashlib.sha1((s or "").encode()).hexdigest()[:8]


def human_time(iso):
    if not iso:
        return ""
    try:
        s = iso.replace("Z", "+00:00")
        s = re.sub(r"(\.\d{6})\d+", r"\1", s)   # Sekundenbruchteile kürzen
        return datetime.fromisoformat(s).astimezone().strftime("%Y-%m-%d %H:%M")
    except Exception:
        return iso


def parse_ts(iso):
    """ISO-8601 (Graph, UTC) -> vergleichbares datetime oder None."""
    if not iso:
        return None
    try:
        s = iso.replace("Z", "+00:00")
        s = re.sub(r"(\.\d{6})\d+", r"\1", s)
        return datetime.fromisoformat(s)
    except Exception:
        return None


def newest_iso(strings):
    """Liefert den ISO-String mit dem spätesten Zeitpunkt (robust geparst)."""
    best_iso, best_dt = None, None
    for iso in strings:
        dt = parse_ts(iso)
        if dt and (best_dt is None or dt > best_dt):
            best_dt, best_iso = dt, iso
    return best_iso


def strip_tags(s):
    return re.sub(r"<[^>]+>", "", s or "").strip()


def clean_html(s):
    s = s or ""
    s = re.sub(r"<script\b[^>]*>.*?</script>", "", s, flags=re.I | re.S)
    s = re.sub(r'\son\w+\s*=\s*"[^"]*"', "", s, flags=re.I)
    s = re.sub(r"\son\w+\s*=\s*'[^']*'", "", s, flags=re.I)
    return s


HOSTED_RE = re.compile(
    r'https://graph\.microsoft\.com/(?:v1\.0|beta)/[^\s"\'<>]*?hostedContents/[^\s"\'<>]+?/\$value'
)


_PLACEHOLDER_SVG = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="240" height="90">'
    '<rect width="240" height="90" rx="6" fill="#f3f4f6" stroke="#d6d9dd"/>'
    '<text x="120" y="40" font-family="sans-serif" font-size="13" fill="#8a8f98"'
    ' text-anchor="middle">Bild nicht verfügbar</text>'
    '<text x="120" y="60" font-family="sans-serif" font-size="11" fill="#aab0b8"'
    ' text-anchor="middle">Download fehlgeschlagen</text></svg>'
)
IMG_PLACEHOLDER = ("data:image/svg+xml;base64,"
                   + base64.b64encode(_PLACEHOLDER_SVG.encode("utf-8")).decode())


def embed_hosted_images(html_content, counter=None):
    if not _client:
        return html_content

    def repl(m):
        url = m.group(0)
        cf = None
        if IMGCACHE_DIR is not None:
            cf = IMGCACHE_DIR / hashlib.sha1(url.encode()).hexdigest()
            if cf.exists():
                try:
                    return cf.read_text(encoding="utf-8")   # Cache-Treffer: kein Download
                except OSError:
                    pass
        try:
            content, ctype = _client.get_bytes(url)
            data_uri = f"data:{ctype or 'image/png'};base64," + base64.b64encode(content).decode()
            if counter is not None:
                counter[0] += 1
            if cf is not None:
                try:
                    cf.write_text(data_uri, encoding="utf-8")
                except OSError:
                    pass
            return data_uri
        except TokenExpired:
            raise   # Token tot -> Konversation nicht halb schreiben
        except Exception:
            return IMG_PLACEHOLDER   # 502 o. Ä. -> sichtbarer Platzhalter, weitermachen

    return HOSTED_RE.sub(repl, html_content)


def render_attachments(atts):
    items = []
    for a in atts or []:
        nm = a.get("name") or a.get("contentType") or "Anhang"
        url = a.get("contentUrl")
        if url:
            items.append(f'📎 <a href="{html_lib.escape(url)}">{html_lib.escape(nm)}</a>')
        else:
            items.append(f"📎 {html_lib.escape(nm)}")
    return f'<div class="att">{"<br>".join(items)}</div>' if items else ""


def render_reactions(rs):
    if not rs:
        return ""
    c = Counter(r.get("reactionType", "?") for r in rs)
    return '<div class="react">' + html_lib.escape(
        " ".join(f"{k} ×{v}" for k, v in c.items())) + "</div>"


def render_message(msg, is_reply=False, img_counter=None):
    when = human_time(msg.get("createdDateTime"))
    if msg.get("messageType", "message") != "message":   # System-Event
        ed = msg.get("eventDetail") or {}
        label = strip_tags(html_lib.unescape((msg.get("body") or {}).get("content", ""))) \
            or ed.get("@odata.type", "").split(".")[-1] or "Systemnachricht"
        return f'<div class="sys">{html_lib.escape(label)} · {when}</div>'

    frm = msg.get("from") or {}
    user = frm.get("user") or {}
    app = frm.get("application") or {}
    name = user.get("displayName") or app.get("displayName") or "Unbekannt"

    cls = "msg reply" if is_reply else "msg"
    body = msg.get("body") or {}
    if msg.get("deletedDateTime"):
        body_html = "<em>[gelöscht]</em>"
    elif (body.get("contentType") or "text") == "html":
        body_html = clean_html(body.get("content", ""))
        if EMBED_IMAGES:
            body_html = embed_hosted_images(body_html, img_counter)
    else:
        cls += " text"
        body_html = html_lib.escape(body.get("content", ""))

    subj = msg.get("subject")
    subj_html = f'<div class="subj"><strong>{html_lib.escape(subj)}</strong></div>' if subj else ""

    return (f'<div class="{cls}"><div class="head">'
            f'<span class="name">{html_lib.escape(name)}</span>'
            f'<span class="time">{when}</span></div>{subj_html}'
            f'<div class="body">{body_html}</div>'
            f'{render_attachments(msg.get("attachments"))}'
            f'{render_reactions(msg.get("reactions"))}</div>')


def member_name(m):
    return (m.get("displayName") or m.get("email") or "").strip()


def chat_title(graph, chat, my_id):
    ctype = chat.get("chatType")
    topic = chat.get("topic")
    if ctype != "oneOnOne" and topic:
        return topic
    members = chat.get("members")            # i. d. R. schon per $expand vorhanden
    if not members:
        try:   # Fallback: Mitglieder einzeln laden
            members = list(graph.paged(f"{GRAPH}/me/chats/{chat['id']}/members", {"$top": PAGE}))
        except TokenExpired:
            raise
        except Exception:
            members = []
    others = [member_name(m) for m in members
              if m.get("userId") != my_id and member_name(m)]
    if ctype == "oneOnOne":
        return others[0] if others else "Unbekannt"
    if others:
        return ", ".join(others[:5]) + ("…" if len(others) > 5 else "")
    return topic or ctype or "Chat"


CSS = """
:root{color-scheme:light}
*{box-sizing:border-box}
body{margin:0;font:15px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;color:#1b1b1f;background:#f6f7f9}
header{position:sticky;top:0;background:#fff;border-bottom:1px solid #e3e5e8;padding:16px 24px;z-index:1}
header h1{margin:0 0 2px;font-size:18px}
header .sub{margin:0;color:#5b5f66;font-size:13px}
header .meta{margin:6px 0 0;color:#8a8f98;font-size:12px}
header .warn{margin:6px 0 0;color:#b5651d;font-size:12px}
main{max-width:860px;margin:0 auto;padding:20px 16px 60px}
.msg{padding:10px 14px;margin:2px 0;background:#fff;border:1px solid #ececef;border-radius:10px}
.msg .head{display:flex;gap:8px;align-items:baseline;margin-bottom:3px}
.msg .name{font-weight:600}
.msg .time{color:#9aa0a6;font-size:12px}
.msg .subj{margin-bottom:4px}
.msg .body{word-wrap:break-word;overflow-wrap:anywhere}
.msg .body img{max-width:100%;height:auto;border-radius:6px}
.msg.text .body{white-space:pre-wrap}
.reply{margin-left:28px;border-left:3px solid #d7dadf;border-radius:0 10px 10px 0}
.sys{background:transparent;border:none;text-align:center;color:#9aa0a6;font-size:12px;padding:6px}
.att{margin-top:6px;font-size:13px}
.att a{color:#2b6cb0;text-decoration:none}
.react{margin-top:4px;color:#8a8f98;font-size:12px}
.empty{color:#9aa0a6}
.index-group{max-width:860px;margin:0 auto;padding:8px 16px}
.index-group h2{font-size:15px;margin:22px 0 8px;color:#3b3f46}
.index-group ul{list-style:none;padding:0;margin:0}
.index-group li{padding:6px 0;border-bottom:1px solid #ececef}
.index-group a{color:#2b6cb0;text-decoration:none}
.index-group .c{color:#9aa0a6;font-size:12px;margin-left:6px}
"""


def render_conversation(title, subtitle, meta, blocks):
    body = "".join(blocks) or '<p class="empty">Keine Nachrichten.</p>'
    n_fail = body.count(IMG_PLACEHOLDER)
    warn = (f'<p class="warn">&#9888; {n_fail} Bild(er) konnten nicht geladen werden '
            f'und sind als Platzhalter markiert.</p>') if n_fail else ""
    return (f'<!DOCTYPE html><html lang="de"><head><meta charset="utf-8">'
            f'<meta name="viewport" content="width=device-width, initial-scale=1">'
            f'<title>{html_lib.escape(title)}</title><style>{CSS}</style></head><body>'
            f'<header><h1>{html_lib.escape(title)}</h1>'
            f'<p class="sub">{html_lib.escape(subtitle)}</p>'
            f'<p class="meta">{html_lib.escape(meta)}</p>{warn}</header>'
            f'<main>{body}</main></body></html>')


# ---------------------------------------------------------------------------
# Fortschritt (thread-sicher)
# ---------------------------------------------------------------------------
def load_state(out):
    p = out / STATE_FILE
    if p.exists():
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
            if isinstance(data, dict) and "conversations" in data:
                return data
        except Exception:
            print("Warnung: Fortschrittsdatei unlesbar – starte ohne Resume.")
    return {"version": 1, "conversations": {}}


def save_state(out, state):
    tmp = out / (STATE_FILE + ".tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, out / STATE_FILE)   # atomarer Austausch


def already_done(out, state, key):
    rec = state["conversations"].get(key)
    if not rec or not rec.get("done"):
        return False
    return (out / rec["rel"]).exists()   # nur überspringen, wenn Datei noch da ist


def get_record(out, state, key):
    """Vorhandenen, abgeschlossenen Datensatz lesen (samt last_activity) – oder None,
    falls nicht exportiert oder die Datei fehlt. Thread-sicher."""
    with STATE_LOCK:
        rec = state["conversations"].get(key)
    if not rec or not rec.get("done"):
        return None
    if rec.get("empty"):
        return rec   # als leer markiert: keine Datei, aber gültiger Status (für inkrementelle Prüfung)
    if not (out / rec["rel"]).exists():
        return None
    return rec


def cleanup_old(out, prior, new_rel):
    """Beim Umbenennen (z. B. 'Unbekannt' -> echter Name) die verwaiste Altdatei
    entfernen, damit kein Duplikat zurückbleibt."""
    if prior and prior.get("rel") and prior["rel"] != new_rel:
        try:
            (out / prior["rel"]).unlink()
        except OSError:
            pass


def record_done(out, state, key, category, title, rel, count, last_activity=None, empty=False):
    with STATE_LOCK:   # mehrere Worker schreiben -> serialisieren
        state["conversations"][key] = {
            "category": category, "title": title, "rel": rel,
            "count": count, "done": True, "empty": empty,
            "ts": datetime.now().isoformat(timespec="seconds"),
            "last_activity": last_activity,   # neueste Nachricht -> Basis für inkrementelle Läufe
        }
        save_state(out, state)


def write_index(out, state):
    groups = [("1:1-Chats", "1on1"), ("Gruppenchats", "group"),
              ("Meeting-Chats", "meeting"), ("Team-Kanäle", "channels")]
    by_cat = defaultdict(list)
    for rec in state["conversations"].values():
        if rec.get("done") and not rec.get("empty") and rec.get("rel"):
            by_cat[rec["category"]].append((rec["title"], rec["rel"], rec["count"]))
    parts = [f'<header><h1>Teams-Export</h1>'
             f'<p class="sub">Stand {datetime.now():%Y-%m-%d %H:%M}</p></header>']
    for label, key in groups:
        entries = by_cat.get(key, [])
        if not entries:
            continue
        parts.append(f'<div class="index-group"><h2>{html_lib.escape(label)} '
                     f'({len(entries)})</h2><ul>')
        for title, rel, count in sorted(entries, key=lambda x: x[0].lower()):
            parts.append(f'<li><a href="{html_lib.escape(rel)}">{html_lib.escape(title)}</a>'
                         f'<span class="c">{count} Nachrichten</span></li>')
        parts.append("</ul></div>")
    html = (f'<!DOCTYPE html><html lang="de"><head><meta charset="utf-8">'
            f'<title>Teams-Export</title><style>{CSS}</style></head>'
            f'<body>{"".join(parts)}</body></html>')
    (out / "index.html").write_text(html, encoding="utf-8")


# ---------------------------------------------------------------------------
# Export EINER Konversation (läuft in einem Worker-Thread)
# ---------------------------------------------------------------------------
def render_blocks(msgs, label):
    """Rendert Nachrichten-Blöcke mit Fortschritt; zeigt geladene Bilder an."""
    img, blocks, total = [0], [], len(msgs)
    for i, m in enumerate(msgs, 1):
        blocks.append(render_message(m, img_counter=img))
        if i % 1000 == 0:
            extra = f", {img[0]} Bilder geladen" if img[0] else ""
            log(f"   {label}: {i}/{total} gerendert{extra}…")
    return blocks, img[0]


def export_one_chat(graph, out, state, my_id, chat):
    t0 = time.monotonic()
    key = chat["id"]
    folder = TYPEMAP.get(chat.get("chatType"), "other")
    title = chat_title(graph, chat, my_id)
    lbl = title[:32]
    prior = get_record(out, state, key)
    log(f"→ {lbl} [{SUBNAME.get(folder, 'Chat')}] – lade Nachrichten…")
    msgs = []
    for m in graph.paged(f"{GRAPH}/me/chats/{key}/messages", {"$top": PAGE}):
        msgs.append(m)
        if len(msgs) % 2000 == 0:
            log(f"   {lbl}: {len(msgs)} Nachrichten geladen…")
    msgs.sort(key=lambda m: m.get("createdDateTime") or "")

    # Nur System-/Event-Nachrichten und keine echte Nachricht? -> standardmäßig nicht exportieren
    real = sum(1 for m in msgs if m.get("messageType", "message") == "message")
    if SKIP_EMPTY_CHATS and real == 0:
        cleanup_old(out, prior, None)   # evtl. früher geschriebene Datei entfernen
        last_act = newest_iso(m.get("createdDateTime") for m in msgs)
        record_done(out, state, key, folder, title, None, len(msgs),
                    last_activity=last_act, empty=True)
        return ("empty", folder, title, len(msgs), time.monotonic() - t0)

    if len(msgs) >= 2000:
        log(f"   {lbl}: {len(msgs)} Nachrichten geladen – rendere und lade Bilder…")
    blocks, nimg = render_blocks(msgs, lbl)
    meta = f"{len(msgs)} Nachrichten · Chat-ID {key}"
    fname = f"{safe(title)}__{short_id(key)}.html"
    new_rel = f"{folder}/{fname}"
    (out / folder).mkdir(parents=True, exist_ok=True)
    (out / folder / fname).write_text(
        render_conversation(title, SUBNAME.get(folder, "Chat"), meta, blocks),
        encoding="utf-8")
    cleanup_old(out, prior, new_rel)   # alte 'Unbekannt__…'-Datei entfernen, falls umbenannt
    last_act = newest_iso(m.get("createdDateTime") for m in msgs)
    record_done(out, state, key, folder, title, new_rel, len(msgs),
                last_activity=last_act)
    return ("updated" if prior else "new", folder, title, len(msgs), time.monotonic() - t0)


def export_one_channel(graph, out, state, team, ch):
    t0 = time.monotonic()
    tname = team.get("displayName", "Team")
    cname = ch.get("displayName", "Kanal")
    lbl = f"{tname}/{cname}"[:32]
    key = f"ch:{ch['id']}"
    prior = get_record(out, state, key)
    log(f"→ {lbl} [Kanal] – prüfe/lade…")
    base = f"{GRAPH}/teams/{team['id']}/channels/{ch['id']}/messages"
    # Wurzel-Posts MIT eingebetteten Antworten (bis 1000 inline) holen
    roots = list(graph.paged(base, {"$top": PAGE, "$expand": "replies"}))
    roots.sort(key=lambda m: m.get("createdDateTime") or "")
    img = [0]
    blocks, count, times = [], 0, []
    for root in roots:
        blocks.append(render_message(root, img_counter=img))
        count += 1
        times.append(root.get("createdDateTime"))
        times.append(root.get("lastModifiedDateTime"))
        replies = list(root.get("replies") or [])
        nxt = root.get("replies@odata.nextLink")   # nur bei > 1000 Antworten
        while nxt:
            data = graph.get(nxt)
            replies.extend(data.get("value", []))
            nxt = data.get("@odata.nextLink")
        replies.sort(key=lambda m: m.get("createdDateTime") or "")
        for rep in replies:
            blocks.append(render_message(rep, is_reply=True, img_counter=img))
            count += 1
            times.append(rep.get("createdDateTime"))
            times.append(rep.get("lastModifiedDateTime"))
        if count % 1000 == 0:
            extra = f", {img[0]} Bilder geladen" if img[0] else ""
            log(f"   {lbl}: {count} Nachrichten verarbeitet{extra}…")
    fp = newest_iso(times)   # neueste Aktivität (inkl. Antworten/Bearbeitungen)

    # Unverändert seit letztem Lauf? -> nicht neu schreiben
    if prior:
        ps, cs = parse_ts(prior.get("last_activity")), parse_ts(fp)
        if cs is None or (ps is not None and cs <= ps):
            return ("unchanged", "channels", f"{tname} / {cname}", count, time.monotonic() - t0)

    title = f"{tname} / {cname}"
    meta = f"{count} Nachrichten (inkl. Antworten) · {ch.get('membershipType', 'standard')}"
    tdir = out / "channels" / safe(tname)
    tdir.mkdir(parents=True, exist_ok=True)
    fname = f"{safe(cname)}__{short_id(ch['id'])}.html"
    new_rel = f"channels/{safe(tname)}/{fname}"
    (tdir / fname).write_text(
        render_conversation(title, "Team-Kanal", meta, blocks), encoding="utf-8")
    cleanup_old(out, prior, new_rel)
    record_done(out, state, key, "channels", title, new_rel, count, last_activity=fp)
    return ("updated" if prior else "new", "channels", title, count, time.monotonic() - t0)


def make_runner(graph, out, state, my_id, kind, a, b):
    def run():
        if STOP.is_set():
            return ("stopped", None, None, 0, 0.0)
        try:
            if kind == "chat":
                return export_one_chat(graph, out, state, my_id, a)
            return export_one_channel(graph, out, state, a, b)
        except TokenExpired:
            STOP.set()
            return ("expired", None, None, 0, 0.0)
        except Exception as e:
            return ("error", None, f"{e}", 0, 0.0)
    return run


# ---------------------------------------------------------------------------
# Job-Aufbau (im Hauptthread) + paralleler Treiber
# ---------------------------------------------------------------------------
def build_chat_jobs(graph, out, state, stats, my_id, chat_cats):
    print("\nLade Chat-Liste… (mit letzter Aktivität je Chat)")
    chats = []
    # members + lastMessagePreview inline -> richtige 1:1-Namen ohne Extra-Aufruf,
    # und der Aktivitäts-Zeitstempel je Chat für inkrementelle Läufe
    for c in graph.paged(f"{GRAPH}/me/chats",
                         {"$top": PAGE, "$expand": "members,lastMessagePreview"}):
        chats.append(c)
        if len(chats) % 50 == 0:
            print(f"  … {len(chats)} Chats geladen")
    wanted = [c for c in chats if TYPEMAP.get(c.get("chatType"), "other") in chat_cats]
    jobs, new, upd = [], 0, 0
    for chat in wanted:
        cur = (chat.get("lastMessagePreview") or {}).get("createdDateTime")
        rec = get_record(out, state, chat["id"])
        if rec is None:                       # noch nie exportiert (oder Datei fehlt)
            jobs.append(("chat", chat, None))
            new += 1
            continue
        ps, cs = parse_ts(rec.get("last_activity")), parse_ts(cur)
        if cs is not None and (ps is None or cs > ps):
            jobs.append(("chat", chat, None))   # neue Nachrichten -> erneut exportieren
            upd += 1
        else:
            stats["skipped"] += 1               # unverändert
    print(f"{len(wanted)} passende Chats: {new} neu, {upd} mit neuen Nachrichten, "
          f"{len(wanted) - len(jobs)} unverändert.")
    return jobs


def build_channel_jobs(graph, out, state, stats, selected_teams):
    jobs = []
    for team in selected_teams:
        tname = team.get("displayName", "Team")
        try:
            channels = list(graph.paged(f"{GRAPH}/teams/{team['id']}/channels", {"$top": PAGE}))
        except TokenExpired:
            raise
        except Exception as e:
            print(f"  Kanäle von '{tname}' nicht ladbar ({e})")
            continue
        for ch in channels:
            if REFRESH_CHANNELS:
                # kein günstiger Indikator für neue Antworten -> erneut holen,
                # der Worker schreibt nur bei tatsächlicher Änderung neu
                jobs.append(("channel", team, ch))
            elif already_done(out, state, f"ch:{ch['id']}"):
                stats["skipped"] += 1
            else:
                jobs.append(("channel", team, ch))
    if REFRESH_CHANNELS:
        print(f"{len(jobs)} Kanäle werden auf Aktualität geprüft.")
    else:
        print(f"{len(jobs)} Kanäle zu exportieren.")
    return jobs


def run_parallel(runners, stats, workers):
    if not runners:
        return "done"
    expired = False
    total = len(runners)
    done_count = 0
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = [ex.submit(r) for r in runners]
        for fut in as_completed(futs):
            done_count += 1
            try:
                res = fut.result()
            except Exception as e:
                res = ("error", None, str(e), 0, 0.0)
            status, cat, label, count = res[0], res[1], res[2], res[3]
            secs = res[4] if len(res) > 4 else 0.0
            dur = f"{secs:.0f}s" if secs >= 1 else f"{secs * 1000:.0f}ms"
            if status in ("new", "ok"):
                stats["new"] += 1
                log(f"✓ [{done_count}/{total}] neu · {cat}: {label} — {count} Nachrichten, {dur}")
            elif status == "updated":
                stats["updated"] += 1
                log(f"✓ [{done_count}/{total}] aktualisiert · {cat}: {label} — {count} Nachrichten, {dur}")
            elif status == "unchanged":
                stats["skipped"] += 1   # geprüft, aber keine Änderung (v. a. Kanäle)
                log(f"· [{done_count}/{total}] unverändert · {cat}: {label}")
            elif status == "empty":
                stats["empty"] += 1
                log(f"· [{done_count}/{total}] leer – nur System-Nachrichten, übersprungen · {cat}: {label}")
            elif status == "expired":
                expired = True
            elif status == "error":
                log(f"✗ [{done_count}/{total}] Fehler: {label}")
            # "stopped" -> ignorieren
    return "expired" if expired else "done"


# ---------------------------------------------------------------------------
# Hauptablauf
# ---------------------------------------------------------------------------
def main():
    global _client, OUT_ROOT, GATE, IMGCACHE_DIR
    argv = sys.argv[1:]
    use_default = any(a in ("-default", "--default") for a in argv)
    argv = [a for a in argv if a not in ("-default", "--default")]
    if argv:
        OUT_ROOT = argv[0]

    workers = WORKERS
    env = os.environ.get("EXPORT_WORKERS")
    if env:
        try:
            workers = max(1, int(env))
        except ValueError:
            pass
    GATE = threading.BoundedSemaphore(workers)
    SESSION.mount("https://", requests.adapters.HTTPAdapter(
        pool_connections=max(workers, 4), pool_maxsize=max(workers, 4)))

    # 1) Kategorien abfragen (vor dem Login, damit der Kanal-Scope nur bei Bedarf kommt)
    cat_options = [("1on1", "1:1-Chats"), ("group", "Gruppenchats"),
                   ("meeting", "Meeting-Chats"), ("channels", "Team-Kanäle")]
    if use_default:
        categories = default_categories(cat_options)
        print("Standardauswahl (-default) aktiv – keine Abfrage.")
    else:
        categories = prompt_categories(cat_options)
    labels = {k: v for k, v in cat_options}
    print("Gewählt:", ", ".join(labels[k] for k in
                                 ["1on1", "group", "meeting", "channels"] if k in categories))
    want_channels = "channels" in categories

    # 2) Login bzw. Token-Modus
    pasted = load_pasted_token()
    if pasted:
        print("Token-Modus aktiv – nutze Access Token aus Graph Explorer (kein Login).")
        graph = TokenClient(pasted, channels_enabled=want_channels)
    else:
        graph = Graph(want_channels=want_channels)
    _client = graph
    if want_channels and not graph.channels_enabled:
        print("Hinweis: Kanal-Zugriff nicht verfügbar – Kanäle werden übersprungen.")
        categories.discard("channels")
        want_channels = False

    out = Path(OUT_ROOT)
    out.mkdir(parents=True, exist_ok=True)
    if EMBED_IMAGES and CACHE_IMAGES:
        IMGCACHE_DIR = out / ".imgcache"
        IMGCACHE_DIR.mkdir(parents=True, exist_ok=True)
    state = load_state(out)
    stats = {"new": 0, "updated": 0, "skipped": 0, "empty": 0}
    result = "done"

    try:
        me = graph.get(f"{GRAPH}/me")
        my_id = me.get("id")
        print(f"Angemeldet als {me.get('displayName')} ({me.get('userPrincipalName')})")
        print(f"Parallele Konversationen: {workers}  "
              f"(Teams-Limit: ~1 Anfrage/s je Chat/Kanal, 4/s je Team)")

        selected_teams = select_teams(graph) if want_channels else []

        chat_cats = categories & {"1on1", "group", "meeting"}
        chat_jobs = build_chat_jobs(graph, out, state, stats, my_id, chat_cats) if chat_cats else []
        if not chat_cats:
            print("Keine Chat-Kategorie gewählt – überspringe Chats.")
        channel_jobs = (build_channel_jobs(graph, out, state, stats, selected_teams)
                        if (want_channels and selected_teams) else [])

        runners = [make_runner(graph, out, state, my_id, k, a, b)
                   for (k, a, b) in (chat_jobs + channel_jobs)]
        if runners:
            print(f"\nExportiere {len(runners)} Konversationen mit {workers} parallel…")
        result = run_parallel(runners, stats, workers)
    except TokenExpired:
        result = "expired"
    finally:
        write_index(out, state)

    done_total = sum(1 for r in state["conversations"].values()
                     if r.get("done") and not r.get("empty"))
    if result == "expired":
        print("\nAbgebrochen: Token abgelaufen. Frischen Access Token in gx_token.txt "
              "setzen und erneut starten – bereits exportierte Konversationen bleiben "
              "erhalten.")
        print(f"Bisher im Archiv: {done_total}.")
        sys.exit(1)

    print(f"\nFertig. Neu: {stats['new']}, aktualisiert: {stats['updated']}, "
          f"unverändert: {stats['skipped']}, leer übersprungen: {stats['empty']}, "
          f"gesamt im Archiv: {done_total}.")
    print(f"Ordner: {out.resolve()}")
    print(f"Im Browser öffnen: {out.resolve() / 'index.html'}")


if __name__ == "__main__":
    main()