#!/usr/bin/env python3
"""
Outlook/Exchange-Export als .eml über Microsoft Graph (delegiert, kein Admin nötig).

- Jede Mail als .eml (volle MIME über /messages/{id}/$value, inkl. Anhänge und
  Inline-Bildern). Direkt in jedem Mailprogramm importierbar.
- Optional zusätzlich wählbar: Kalender als .ics (Termine, Zeiten in UTC) und
  Kontakte als .vcf – in eigenen Unterordnern kalender/ und kontakte/.
- Ordnerstruktur des Postfachs wird unter E-Mail/ als Verzeichnisbaum gespiegelt
  (rekursiv) – parallel zu kalender/ und kontakte/.
- PARALLEL: bis zu 4 Downloads gleichzeitig. Exchange Online erlaubt pro Postfach
  nur 4 gleichzeitige Anfragen (MailboxConcurrency, festes Limit) – mehr erzeugt
  nur 429er. Ein globaler Semaphor hält Listing + Downloads zusammen unter dieser
  Grenze; bei 429 wird mit Retry-After zurückgenommen.

Setup:   pip install msal requests
Start:   python3 outlook_export.py [ausgabe-ordner]

Token-Modus (wenn der Tenant für neue Apps "Approval required" verlangt):
    Access Token im Graph Explorer holen (Mail.Read muss zugestimmt sein; für
    Kalender/Kontakte zusätzlich Calendars.Read und Contacts.Read),
    in gx_token.txt neben dieses Skript legen ODER  export GRAPH_TOKEN="eyJ0…"

Resume: exported.tsv im Ausgabeordner (eine Zeile pro fertige Mail). Bereits
    exportierte Mails werden übersprungen. Token tot -> frischen Token setzen,
    neu starten, es geht weiter. Kompletter Neu-Export: exported.tsv löschen.

Schalter unten: WORKERS (Parallelität, sinnvoll max 4 / per Env EXPORT_WORKERS),
    USE_DEVICE_CODE, INCLUDE_HIDDEN (versteckte Systemordner).
"""

import os
import sys
import re
import time
import html
import hashlib
import threading
from datetime import datetime, timezone
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, wait, FIRST_COMPLETED

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
SCOPES = [RES + "Mail.Read", RES + "Calendars.Read", RES + "Contacts.Read", RES + "User.Read"]

USE_DEVICE_CODE = False     # True = Code-Login statt Browser
INCLUDE_HIDDEN = False      # versteckte Ordner (Conversation History, Sync Issues …)
WORKERS = 4                 # parallele Downloads; Exchange-Limit pro Postfach = 4
PAGE = 50                   # $top für Listenabfragen
OUT_ROOT = "outlook_export"  # fester Ordner -> Resume über mehrere Läufe hinweg
DONE_FILE = "exported.tsv"
MAIL_DIR = "E-Mail"          # Postfach-Ordnerbaum liegt darunter (parallel zu kalender/kontakte)

# Diese Postfach-Ordner sind bei "alle" (Enter) standardmäßig NICHT dabei – nur per
# expliziter Auswahl. Vergleich case-insensitive über den Anzeigenamen (DE + EN).
DEFAULT_SKIP_FOLDERS = {
    "archive", "archiv",
    "entwürfe", "drafts",
    "erneut erinnern aktiviert",
    "gelöschte elemente", "deleted items",
    "junk-e-mail", "junk email", "junk-email",
    "postausgang", "outbox",
}

# Geteilte HTTP-Session (Keep-Alive/Connection-Pooling) und Drossel-Gate.
SESSION = requests.Session()
GATE = threading.BoundedSemaphore(WORKERS)   # hält gleichzeitige Postfach-Calls <= WORKERS
STOP = threading.Event()                     # Signal: Token tot -> nichts Neues mehr starten


class TokenExpired(RuntimeError):
    """Signalisiert einen 401 im Token-Modus (kein Refresh möglich)."""


# ---------------------------------------------------------------------------
# Graph-Client (Anmeldung) – Auth, Retry, Paging
# ---------------------------------------------------------------------------
class Graph:
    def __init__(self):
        self.app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
        self.account = None
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
        print("Anmeldung – fordere Lesezugriff auf Postfach, Kalender und Kontakte an…")
        res = self._acquire(SCOPES)
        if not res or "access_token" not in res:
            raise SystemExit("Anmeldung fehlgeschlagen: "
                             + (res or {}).get("error_description", "unbekannt"))
        accs = self.app.get_accounts()
        self.account = accs[0] if accs else None
        self.token = res["access_token"]

    def _refresh(self):
        with self._refresh_lock:   # nur ein Thread erneuert gleichzeitig
            res = self.app.acquire_token_silent(SCOPES, account=self.account) if self.account else None
            if not res or "access_token" not in res:
                res = self._acquire(SCOPES)
            if not res or "access_token" not in res:
                raise SystemExit("Token-Erneuerung fehlgeschlagen.")
            self.token = res["access_token"]

    def _headers(self):
        return {"Authorization": f"Bearer {self.token}"}

    def get(self, url, params=None, extra_headers=None):
        headers = self._headers()
        if extra_headers:
            headers = {**headers, **extra_headers}
        for attempt in range(6):
            with GATE:   # nur das eigentliche Request zählt gegen das Limit
                r = SESSION.get(url, headers=headers, params=params, timeout=60)
            if r.status_code == 401:
                self._refresh()
                headers = {**self._headers(), **(extra_headers or {})}
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
        for attempt in range(6):
            with GATE:
                r = SESSION.get(url, headers=self._headers(), timeout=120)
            if r.status_code == 401:
                self._refresh()
                continue
            if r.status_code == 429 or 500 <= r.status_code < 600:
                ra = r.headers.get("Retry-After")
                w = min(int(ra) if ra and ra.isdigit() else 2 ** attempt, 60)
                print(f"    … HTTP {r.status_code} (MIME), warte {w}s")
                time.sleep(w)
                continue
            r.raise_for_status()
            return r.content, r.headers.get("Content-Type", "")
        raise RuntimeError(f"Zu viele Fehlversuche: {url}")

    def paged(self, url, params=None, extra_headers=None):
        data = self.get(url, params, extra_headers)
        while True:
            for item in data.get("value", []):
                yield item
            nxt = data.get("@odata.nextLink")
            if not nxt:
                break
            data = self.get(nxt, extra_headers=extra_headers)


# ---------------------------------------------------------------------------
# Token-Modus: vorhandenen fertigen Bearer-Token nutzen (z. B. aus Graph Explorer)
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

    def __init__(self, token):
        self.token = token
        self.account = None

    def _headers(self):
        return {"Authorization": f"Bearer {self.token}"}

    def get(self, url, params=None, extra_headers=None):
        headers = self._headers()
        if extra_headers:
            headers = {**headers, **extra_headers}
        for attempt in range(6):
            with GATE:
                r = SESSION.get(url, headers=headers, params=params, timeout=60)
            if r.status_code == 401:
                raise TokenExpired()
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
        for attempt in range(6):
            with GATE:
                r = SESSION.get(url, headers=self._headers(), timeout=120)
            if r.status_code == 401:
                raise TokenExpired()
            if r.status_code == 429 or 500 <= r.status_code < 600:
                ra = r.headers.get("Retry-After")
                w = min(int(ra) if ra and ra.isdigit() else 2 ** attempt, 60)
                print(f"    … HTTP {r.status_code} (MIME), warte {w}s")
                time.sleep(w)
                continue
            r.raise_for_status()
            return r.content, r.headers.get("Content-Type", "")
        raise RuntimeError(f"Zu viele Fehlversuche: {url}")

    def paged(self, url, params=None, extra_headers=None):
        data = self.get(url, params, extra_headers)
        while True:
            for item in data.get("value", []):
                yield item
            nxt = data.get("@odata.nextLink")
            if not nxt:
                break
            data = self.get(nxt, extra_headers=extra_headers)


# ---------------------------------------------------------------------------
# Fortschritt: append-only Log, thread-sicher (skaliert auf zehntausende Mails)
# ---------------------------------------------------------------------------
class DoneLog:
    def __init__(self, path):
        self.path = path
        self.done = {}
        self._lock = threading.Lock()
        if path.exists():
            for line in path.read_text(encoding="utf-8").splitlines():
                if "\t" in line:
                    mid, rel = line.split("\t", 1)
                    self.done[mid] = rel
        self._fh = open(path, "a", encoding="utf-8")

    def is_done(self, out, mid):
        rel = self.done.get(mid)
        return bool(rel) and (out / rel).exists()

    def mark(self, mid, rel):
        with self._lock:
            self.done[mid] = rel
            self._fh.write(f"{mid}\t{rel}\n")
            self._fh.flush()

    def remap(self, fn):
        """Wendet fn(rel)->rel auf alle Einträge an und schreibt die Datei atomar neu.
        Für einmalige Pfad-Migrationen (Resume bleibt erhalten)."""
        with self._lock:
            self.done = {mid: fn(rel) for mid, rel in self.done.items()}
            try:
                self._fh.close()
            except Exception:
                pass
            tmp = self.path.with_name(self.path.name + ".tmp")
            with open(tmp, "w", encoding="utf-8") as f:
                for mid, rel in self.done.items():
                    f.write(f"{mid}\t{rel}\n")
            os.replace(tmp, self.path)
            self._fh = open(self.path, "a", encoding="utf-8")

    def close(self):
        try:
            self._fh.close()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Interaktive Ordnerauswahl
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


def list_calendars(graph):
    """Liest die Kalenderliste für die gezielte Auswahl. Leere Liste bei fehlender
    Berechtigung – dann erscheinen keine Kalender-Einträge im Menü."""
    try:
        cals = list(graph.paged(f"{GRAPH}/me/calendars", {"$top": PAGE}))
    except TokenExpired:
        raise
    except Exception as e:
        print(f"  Kalender nicht lesbar – fehlt die Berechtigung Calendars.Read? ({e})")
        return []
    cals.sort(key=lambda c: (not c.get("isDefaultCalendar"), (c.get("name") or "").lower()))
    return cals


def prompt_categories():
    """Schritt 1: Was exportieren? Mehrfachauswahl (z. B. 1,2).
    Liefert ein Set aus {"mail", "calendar", "contacts"}."""
    options = [("mail", "E-Mail (Postfach-Ordner)"),
               ("calendar", "Kalender"),
               ("contacts", "Kontakte")]
    if not sys.stdin.isatty():
        print("Kein interaktives Terminal – exportiere E-Mail, Standardkalender und Kontakte.")
        return {k for k, _ in options}
    print("\nWas möchtest du exportieren? (Mehrfachauswahl möglich, z. B. 1,2)")
    for i, (_, label) in enumerate(options, 1):
        print(f"  {i}) {label}")
    raw = _read("Auswahl (Enter = alles): ").strip()
    if not raw:
        return {k for k, _ in options}
    idxs = parse_indices(raw, len(options))
    if not idxs:
        print("Keine gültige Auswahl – nehme alles.")
        return {k for k, _ in options}
    return {options[i - 1][0] for i in idxs}


def _is_default_skip(top):
    name = (top["folder"].get("displayName") or "").strip().lower()
    return name in DEFAULT_SKIP_FOLDERS


def select_mail_folders(tops):
    """Schritt 2a: welche Postfach-Ordner (jeweils inkl. Unterordner).
    Enter = alle AUSSER den Standard-Ausschlüssen (Archiv, Entwürfe, Gelöschte
    Elemente, Junk-E-Mail, Postausgang, „Erneut erinnern aktiviert"). Diese werden
    angezeigt, aber nur auf explizite Auswahl exportiert. Mehrfachauswahl möglich."""
    tops.sort(key=lambda t: (t["folder"].get("displayName") or "").lower())
    default = [t for t in tops if not _is_default_skip(t)]
    if not tops or not sys.stdin.isatty():
        return default
    n = len(tops)
    print("\nWelche Postfach-Ordner? (Mehrfachauswahl; Enter = alle ohne die mit (aus); inkl. Unterordner)")
    for i, t in enumerate(tops, 1):
        name = t["folder"].get("displayName", "Ordner")
        subs = t["nfolders"] - 1
        extra = f", {subs} Unterordner" if subs > 0 else ""
        flag = "  (aus)" if _is_default_skip(t) else ""
        print(f"  {i}) {name}  ({t['items']} Elemente{extra}){flag}")
    raw = _read("Auswahl (Enter = alle ohne die mit (aus)): ").strip()
    if not raw:
        return default
    idxs = parse_indices(raw, n)
    if not idxs:
        print("Keine gültige Auswahl – nehme alle ohne die Standard-Ausschlüsse.")
        return default
    return [tops[i - 1] for i in idxs]


def select_calendars(cals):
    """Schritt 2b: welche Kalender. Enter = nur Standardkalender; Mehrfachauswahl möglich."""
    if not cals:
        return []
    default = [c for c in cals if c.get("isDefaultCalendar")] or cals[:1]
    if not sys.stdin.isatty():
        return default
    n = len(cals)
    print("\nWelche Kalender? (Mehrfachauswahl; Enter = nur Standardkalender)")
    for i, cal in enumerate(cals, 1):
        mark = " (Standard)" if cal.get("isDefaultCalendar") else ""
        print(f"  {i}) {cal.get('name', 'Kalender')}{mark} (.ics)")
    raw = _read("Auswahl (Enter = nur Standardkalender): ").strip()
    if not raw:
        return default
    idxs = parse_indices(raw, n)
    if not idxs:
        print("Keine gültige Auswahl – nehme den Standardkalender.")
        return default
    return [cals[i - 1] for i in idxs]


# ---------------------------------------------------------------------------
# Helfer
# ---------------------------------------------------------------------------
def safe(name, maxlen=80):
    name = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", name or "").strip().strip(".")
    name = re.sub(r"\s+", " ", name)
    return name[:maxlen] or "unbenannt"


def short_id(s):
    return hashlib.sha1((s or "").encode()).hexdigest()[:8]


def mail_filename(msg):
    dt = msg.get("receivedDateTime") or msg.get("sentDateTime") or ""
    stamp = ""
    if dt:
        try:
            s = dt.replace("Z", "+00:00")
            s = re.sub(r"(\.\d{6})\d+", r"\1", s)
            stamp = datetime.fromisoformat(s).astimezone().strftime("%Y-%m-%d_%H%M")
        except Exception:
            stamp = dt[:10]
    subj = (msg.get("subject") or "").strip() or "(kein Betreff)"
    prefix = (stamp + "__") if stamp else ""
    return f"{prefix}{safe(subj, 90)}__{short_id(msg['id'])}.eml"


def folder_params():
    p = {"$top": 100}
    if INCLUDE_HIDDEN:
        p["includeHiddenFolders"] = "true"
    return p


def list_children(graph, folder):
    """Listet die direkten Unterordner – unabhängig von childFolderCount."""
    try:
        return list(graph.paged(f"{GRAPH}/me/mailFolders/{folder['id']}/childFolders",
                                folder_params()))
    except TokenExpired:
        raise
    except Exception as e:
        print(f"  Warnung: Unterordner von '{folder.get('displayName')}' nicht lesbar: {e}")
        return []


def _subtree(graph, folder, rel_path, acc):
    """Hängt (folder, rel_path) für den Ordner und ALLE Nachkommen an acc an.
    Verlässt sich nicht auf childFolderCount, sondern listet immer die Kinder –
    so werden auch tief verschachtelte Unterordner zuverlässig erfasst."""
    acc.append((folder, rel_path))
    for child in list_children(graph, folder):
        cname = safe(child.get("displayName") or "Ordner")
        _subtree(graph, child, f"{rel_path}/{cname}", acc)


def build_tree(graph):
    """Liest die komplette Ordnerstruktur EINMAL und liefert pro oberstem Ordner
    den Teilbaum samt rekursiver Elementzahl. Das Ergebnis wird für die Auswahl UND
    den Export genutzt (kein erneutes Ordner-Listing im parallelen Download)."""
    tops = []
    roots = list(graph.paged(f"{GRAPH}/me/mailFolders", folder_params()))
    count = 0
    for tf in roots:
        rel = f"{MAIL_DIR}/{safe(tf.get('displayName') or 'Ordner')}"
        sub = []
        _subtree(graph, tf, rel, sub)
        items = sum((f.get("totalItemCount") or 0) for f, _ in sub)
        tops.append({"folder": tf, "rel": rel, "subtree": sub,
                     "items": items, "nfolders": len(sub)})
        count += len(sub)
        print(f"  … {count} Ordner erfasst", end="\r", flush=True)
    print(f"  {count} Ordner erfasst.            ")
    return tops


def iter_messages_to_export(graph, out, done, stats, selected):
    """Spiegelt die Ordner aufs Dateisystem und liefert (mid, rel) für jede
    noch nicht exportierte Mail. Listing läuft im Hauptthread (lazy)."""
    select = "id,subject,receivedDateTime,sentDateTime,from,hasAttachments"
    for top in selected:
        for folder, rel_path in top["subtree"]:
            (out / rel_path).mkdir(parents=True, exist_ok=True)
            total = folder.get("totalItemCount")
            print(f"\nOrdner: {rel_path}" + (f"  ({total} Elemente)" if total is not None else ""))
            seen = 0
            for msg in graph.paged(f"{GRAPH}/me/mailFolders/{folder['id']}/messages",
                                   {"$top": PAGE, "$select": select}):
                seen += 1
                mid = msg["id"]
                if done.is_done(out, mid):
                    stats["skipped"] += 1
                    continue
                yield mid, f"{rel_path}/{mail_filename(msg)}"
            if seen:
                print(f"  {seen} Mails gesichtet.")


# ---------------------------------------------------------------------------
# Worker + paralleler Treiber
# ---------------------------------------------------------------------------
def download_one(graph, out, done, mid, rel):
    if STOP.is_set():
        return ("stopped", mid)
    try:
        content, _ = graph.get_bytes(f"{GRAPH}/me/messages/{mid}/$value")
    except TokenExpired:
        return ("expired", mid)
    except Exception as e:
        return ("error", f"{mid[:16]}…: {e}")
    try:
        (out / rel).write_bytes(content)
    except Exception as e:
        return ("error", f"{rel}: {e}")
    done.mark(mid, rel)
    return ("ok", rel)


def run_export(graph, out, done, stats, selected, workers):
    gen = iter_messages_to_export(graph, out, done, stats, selected)
    cap = max(workers * 8, workers)      # so viele Tasks gleichzeitig in der Pipeline
    pending = set()
    expired = False

    with ThreadPoolExecutor(max_workers=workers) as ex:
        def fill():
            nonlocal expired
            while len(pending) < cap:
                try:
                    mid, rel = next(gen)
                except StopIteration:
                    return
                except TokenExpired:        # Token kann schon beim Listing sterben
                    expired = True
                    STOP.set()
                    return
                pending.add(ex.submit(download_one, graph, out, done, mid, rel))

        fill()
        while pending:
            finished, rest = wait(pending, return_when=FIRST_COMPLETED)
            pending = set(rest)
            for fut in finished:
                try:
                    status, info = fut.result()
                except Exception as e:
                    status, info = "error", str(e)
                if status == "ok":
                    stats["new"] += 1
                    if stats["new"] % 50 == 0:
                        print(f"  … {stats['new']} Mails neu exportiert")
                elif status == "expired":
                    expired = True
                    STOP.set()
                elif status == "error":
                    print(f"    Mail übersprungen ({info})")
                # "stopped" -> ignorieren
            if not expired:
                fill()

    return "expired" if expired else "done"


# ---------------------------------------------------------------------------
# Kalender (.ics) und Kontakte (.vcf)
# ---------------------------------------------------------------------------
_WD = {"monday": "MO", "tuesday": "TU", "wednesday": "WE", "thursday": "TH",
       "friday": "FR", "saturday": "SA", "sunday": "SU"}
_IDX = {"first": 1, "second": 2, "third": 3, "fourth": 4, "last": -1}


def _plain_text(body):
    body = body or {}
    c = body.get("content", "") or ""
    if (body.get("contentType") or "").lower() == "html":
        c = re.sub(r"(?is)<(script|style)\b.*?</\1>", " ", c)
        c = re.sub(r"(?is)<br\s*/?>|</p>|</div>|</li>|</tr>", " ", c)
        c = re.sub(r"<[^>]+>", " ", c)
        c = html.unescape(c)
    return " ".join(c.split())


def _esc(s):
    s = (s or "").replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,")
    return s.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\\n")


def _cn(name):
    return '"' + " ".join((name or "").split()).replace('"', "'") + '"'


def _fold(line):
    """iCal/vCard-Zeilen auf <=75 Oktette falten (CRLF + Leerzeichen)."""
    out, cur = "", 0
    for ch in line:
        w = len(ch.encode("utf-8"))
        if cur + w > 73:
            out += "\r\n "
            cur = 1
        out += ch
        cur += w
    return out


def _graph_dt(s):
    if not s:
        return None
    s = s.strip().replace("Z", "")
    s = re.sub(r"(\.\d{6})\d+", r"\1", s)
    try:
        return datetime.fromisoformat(s)
    except Exception:
        try:
            return datetime.strptime(s[:19], "%Y-%m-%dT%H:%M:%S")
        except Exception:
            return None


def _ics_dt(node, all_day):
    dt = _graph_dt((node or {}).get("dateTime") or "")
    if dt is None:
        return None
    return dt.strftime("%Y%m%d") if all_day else dt.strftime("%Y%m%dT%H%M%SZ")


def _stamp(node, all_day):
    dt = _graph_dt((node or {}).get("dateTime") or "")
    if dt is None:
        return ""
    return dt.strftime("%Y-%m-%d") if all_day else dt.strftime("%Y-%m-%d_%H%M")


def build_rrule(recurrence, all_day):
    if not recurrence:
        return None
    try:
        pat = recurrence.get("pattern") or {}
        rng = recurrence.get("range") or {}
        ptype = pat.get("type", "")
        interval = int(pat.get("interval", 1) or 1)
        days = [_WD[d.lower()] for d in (pat.get("daysOfWeek") or []) if d.lower() in _WD]
        idx = _IDX.get(pat.get("index", "first"), 1)
        parts = []
        if ptype == "daily":
            parts.append("FREQ=DAILY")
        elif ptype == "weekly":
            parts.append("FREQ=WEEKLY")
            if days:
                parts.append("BYDAY=" + ",".join(days))
        elif ptype == "absoluteMonthly":
            parts.append("FREQ=MONTHLY")
            if pat.get("dayOfMonth"):
                parts.append(f"BYMONTHDAY={pat['dayOfMonth']}")
        elif ptype == "relativeMonthly":
            parts.append("FREQ=MONTHLY")
            if days:
                parts.append("BYDAY=" + ",".join(f"{idx}{d}" for d in days))
        elif ptype == "absoluteYearly":
            parts.append("FREQ=YEARLY")
            if pat.get("month"):
                parts.append(f"BYMONTH={pat['month']}")
            if pat.get("dayOfMonth"):
                parts.append(f"BYMONTHDAY={pat['dayOfMonth']}")
        elif ptype == "relativeYearly":
            parts.append("FREQ=YEARLY")
            if pat.get("month"):
                parts.append(f"BYMONTH={pat['month']}")
            if days:
                parts.append("BYDAY=" + ",".join(f"{idx}{d}" for d in days))
        else:
            return None
        if interval != 1:
            parts.append(f"INTERVAL={interval}")
        rtype = rng.get("type", "")
        if rtype == "endDate" and rng.get("endDate"):
            d = rng["endDate"].replace("-", "")
            parts.append("UNTIL=" + (d if all_day else d + "T235959Z"))
        elif rtype == "numbered" and rng.get("numberOfOccurrences"):
            parts.append(f"COUNT={int(rng['numberOfOccurrences'])}")
        return ";".join(parts)
    except Exception:
        return None


def event_filename(ev):
    all_day = bool(ev.get("isAllDay"))
    stamp = _stamp(ev.get("start"), all_day)
    subj = (ev.get("subject") or "").strip() or "(kein Betreff)"
    prefix = (stamp + "__") if stamp else ""
    return f"{prefix}{safe(subj, 90)}__{short_id(ev.get('id') or ev.get('iCalUId') or subj)}.ics"


def build_ics(ev):
    all_day = bool(ev.get("isAllDay"))
    uid = ev.get("iCalUId") or ev.get("id") or short_id(ev.get("subject") or "")
    stamp = _graph_dt(ev.get("lastModifiedDateTime") or ev.get("createdDateTime") or "")
    dtstamp = (stamp or datetime.now(timezone.utc).replace(tzinfo=None)).strftime("%Y%m%dT%H%M%SZ")
    L = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//outlook_export//Graph//DE",
         "CALSCALE:GREGORIAN", "METHOD:PUBLISH", "BEGIN:VEVENT",
         f"UID:{_esc(uid)}", f"DTSTAMP:{dtstamp}"]
    start, end = _ics_dt(ev.get("start"), all_day), _ics_dt(ev.get("end"), all_day)
    if start:
        L.append(("DTSTART;VALUE=DATE:" if all_day else "DTSTART:") + start)
    if end:
        L.append(("DTEND;VALUE=DATE:" if all_day else "DTEND:") + end)
    L.append("SUMMARY:" + _esc(ev.get("subject") or "(kein Betreff)"))
    loc = (ev.get("location") or {}).get("displayName")
    if loc:
        L.append("LOCATION:" + _esc(loc))
    desc = _plain_text(ev.get("body"))
    if desc:
        L.append("DESCRIPTION:" + _esc(desc))
    org = (ev.get("organizer") or {}).get("emailAddress") or {}
    if org.get("address"):
        L.append(f'ORGANIZER;CN={_cn(org.get("name") or org["address"])}:mailto:{org["address"]}')
    for a in ev.get("attendees") or []:
        em = a.get("emailAddress") or {}
        if em.get("address"):
            L.append(f'ATTENDEE;CN={_cn(em.get("name") or em["address"])}:mailto:{em["address"]}')
    show = ev.get("showAs", "")
    if ev.get("isCancelled"):
        L.append("STATUS:CANCELLED")
    elif show == "tentative":
        L.append("STATUS:TENTATIVE")
    else:
        L.append("STATUS:CONFIRMED")
    L.append("TRANSP:" + ("TRANSPARENT" if show == "free" else "OPAQUE"))
    rr = build_rrule(ev.get("recurrence"), all_day)
    if rr:
        L.append("RRULE:" + rr)
    L += ["END:VEVENT", "END:VCALENDAR"]
    return "\r\n".join(_fold(x) for x in L) + "\r\n"


def export_calendar(graph, out, done, stats, cals):
    if not cals:
        return
    print("\nKalender…")
    pref = {"Prefer": 'outlook.timezone="UTC"'}      # Zeiten in UTC -> korrekte .ics
    select = ("id,iCalUId,subject,start,end,isAllDay,location,organizer,attendees,"
              "body,showAs,isCancelled,recurrence,seriesMasterId,type,"
              "createdDateTime,lastModifiedDateTime")
    for cal in cals:
        cname = safe(cal.get("name") or "Kalender")
        url = (f"{GRAPH}/me/calendars/{cal['id']}/events" if cal.get("id")
               else f"{GRAPH}/me/events")
        print(f"\nKalender: {cname}")
        seen = 0
        try:
            for ev in graph.paged(url, {"$top": PAGE, "$select": select}, extra_headers=pref):
                seen += 1
                eid = ev.get("id")
                if eid and done.is_done(out, eid):
                    stats["skipped"] += 1
                    continue
                rel = f"kalender/{cname}/{event_filename(ev)}"
                (out / "kalender" / cname).mkdir(parents=True, exist_ok=True)
                try:
                    (out / rel).write_text(build_ics(ev), encoding="utf-8")
                except Exception as e:
                    print(f"    Termin übersprungen ({e})")
                    continue
                done.mark(eid, rel)
                stats["new"] += 1
                if stats["new"] % 100 == 0:
                    print(f"  … {stats['new']} neu exportiert")
        except TokenExpired:
            raise
        except Exception as e:
            print(f"  Kalender '{cname}' abgebrochen: {e}")
            continue
        if seen:
            print(f"  {seen} Termine gesichtet.")


def contact_filename(c):
    nm = (c.get("displayName")
          or " ".join(x for x in [c.get("givenName"), c.get("surname")] if x)).strip() or "Kontakt"
    return f"{safe(nm, 90)}__{short_id(c.get('id') or nm)}.vcf"


def build_vcf(c):
    given, sur, mid = c.get("givenName") or "", c.get("surname") or "", c.get("middleName") or ""
    fn = c.get("displayName") or " ".join(x for x in [given, sur] if x).strip() or "(ohne Namen)"
    L = ["BEGIN:VCARD", "VERSION:3.0",
         f"N:{_esc(sur)};{_esc(given)};{_esc(mid)};;", "FN:" + _esc(fn)]
    org, dept = c.get("companyName") or "", c.get("department") or ""
    if org or dept:
        L.append("ORG:" + _esc(org) + (";" + _esc(dept) if dept else ""))
    if c.get("jobTitle"):
        L.append("TITLE:" + _esc(c["jobTitle"]))
    for e in c.get("emailAddresses") or []:
        if e.get("address"):
            L.append("EMAIL;TYPE=INTERNET:" + _esc(e["address"]))
    for p in c.get("businessPhones") or []:
        if p:
            L.append("TEL;TYPE=WORK,VOICE:" + _esc(p))
    for p in c.get("homePhones") or []:
        if p:
            L.append("TEL;TYPE=HOME,VOICE:" + _esc(p))
    if c.get("mobilePhone"):
        L.append("TEL;TYPE=CELL,VOICE:" + _esc(c["mobilePhone"]))
    if c.get("personalNotes"):
        L.append("NOTE:" + _esc(c["personalNotes"]))
    if c.get("id"):
        L.append("UID:" + _esc(c["id"]))
    L.append("END:VCARD")
    return "\r\n".join(_fold(x) for x in L) + "\r\n"


def export_contacts(graph, out, done, stats):
    print("\nKontakte…")
    sources = [("", f"{GRAPH}/me/contacts")]          # Standardkontakte (kein Ordner)
    try:
        folders = list(graph.paged(f"{GRAPH}/me/contactFolders", {"$top": PAGE}))
    except TokenExpired:
        raise
    except Exception as e:
        print(f"  Kontaktordner nicht lesbar – fehlt Contacts.Read? ({e})")
        folders = []
    for f in folders:
        sources.append((safe(f.get("displayName") or "Ordner"),
                        f"{GRAPH}/me/contactFolders/{f['id']}/contacts"))
    select = ("id,displayName,givenName,surname,middleName,companyName,department,"
              "jobTitle,emailAddresses,businessPhones,homePhones,mobilePhone,personalNotes")
    for sub, url in sources:
        rel_dir = "kontakte" + (f"/{sub}" if sub else "")
        seen = 0
        try:
            for c in graph.paged(url, {"$top": PAGE, "$select": select}):
                seen += 1
                cid = c.get("id")
                if cid and done.is_done(out, cid):
                    stats["skipped"] += 1
                    continue
                rel = f"{rel_dir}/{contact_filename(c)}"
                (out / rel_dir).mkdir(parents=True, exist_ok=True)
                try:
                    (out / rel).write_text(build_vcf(c), encoding="utf-8")
                except Exception as e:
                    print(f"    Kontakt übersprungen ({e})")
                    continue
                done.mark(cid, rel)
                stats["new"] += 1
        except TokenExpired:
            raise
        except Exception as e:
            print(f"  '{rel_dir}' abgebrochen: {e}")
            continue
        if seen:
            print(f"  {rel_dir}: {seen} Kontakte gesichtet.")


# ---------------------------------------------------------------------------
# Hauptablauf
# ---------------------------------------------------------------------------
def migrate_to_email_subdir(out, done):
    """Einmalige, idempotente Migration: bereits exportierte Mail-Ordner (oberste
    Ebene) nach E-Mail/ verschieben und die Resume-Pfade entsprechend umschreiben,
    damit nichts neu heruntergeladen wird. kalender/ und kontakte/ bleiben unberührt.
    No-op bei neuer/leerer Struktur."""
    reserved = {MAIL_DIR, "kalender", "kontakte"}
    try:
        children = [c for c in out.iterdir() if c.is_dir() and c.name not in reserved]
    except FileNotFoundError:
        return
    if not children:
        return
    target = out / MAIL_DIR
    target.mkdir(parents=True, exist_ok=True)
    moved = 0
    for d in children:
        dest = target / d.name
        if dest.exists():
            continue   # Teilmigration/Namenskollision -> sicherheitshalber überspringen
        try:
            d.rename(dest)
            moved += 1
        except OSError as e:
            print(f"  Migration: '{d.name}' nicht verschoben ({e})")
    if not moved:
        return

    def fix(rel):
        return rel if rel.split("/", 1)[0] in reserved else f"{MAIL_DIR}/{rel}"
    done.remap(fix)
    print(f"Struktur aktualisiert: {moved} Mail-Ordner nach '{MAIL_DIR}/' verschoben "
          f"(Resume-Liste angepasst, kein erneuter Download).")


def main():
    global OUT_ROOT, GATE
    if len(sys.argv) > 1:
        OUT_ROOT = sys.argv[1]

    workers = WORKERS
    env = os.environ.get("EXPORT_WORKERS")
    if env:
        try:
            workers = max(1, int(env))
        except ValueError:
            pass
    if workers > 4:
        print("Hinweis: Exchange Online erlaubt nur 4 gleichzeitige Anfragen pro "
              f"Postfach – {workers} Worker erzeugen v. a. Drosselung. 4 ist das "
              "sinnvolle Maximum.")
    GATE = threading.BoundedSemaphore(workers)
    SESSION.mount("https://", requests.adapters.HTTPAdapter(
        pool_connections=max(workers, 4), pool_maxsize=max(workers, 4)))

    pasted = load_pasted_token()
    if pasted:
        print("Token-Modus aktiv – nutze Access Token aus Graph Explorer (kein Login).")
        graph = TokenClient(pasted)
    else:
        graph = Graph()

    out = Path(OUT_ROOT)
    out.mkdir(parents=True, exist_ok=True)
    done = DoneLog(out / DONE_FILE)
    migrate_to_email_subdir(out, done)   # einmalig: Alt-Struktur -> E-Mail/
    stats = {"new": 0, "skipped": 0}
    result = "done"

    try:
        me = graph.get(f"{GRAPH}/me")
        print(f"Angemeldet als {me.get('displayName')} ({me.get('userPrincipalName')})")
        print(f"Parallele Downloads: {workers} (Exchange-Limit pro Postfach)")

        categories = prompt_categories()
        selected_mail, sel_cals, want_con = [], [], False

        if "mail" in categories:
            print("\nLade Ordnerstruktur inkl. aller Unterordner…")
            tops = build_tree(graph)
            selected_mail = select_mail_folders(tops)
        if "calendar" in categories:
            print("\nLade Kalenderliste…")
            sel_cals = select_calendars(list_calendars(graph))
        want_con = "contacts" in categories

        if selected_mail:
            result = run_export(graph, out, done, stats, selected_mail, workers)
        if result != "expired" and sel_cals:
            try:
                export_calendar(graph, out, done, stats, sel_cals)
            except TokenExpired:
                result = "expired"
        if result != "expired" and want_con:
            try:
                export_contacts(graph, out, done, stats)
            except TokenExpired:
                result = "expired"
    except TokenExpired:
        result = "expired"
    finally:
        done.close()

    if result == "expired":
        print("\nAbgebrochen: Token abgelaufen. Frischen Access Token in gx_token.txt "
              "setzen und erneut starten – bereits Exportiertes bleibt erhalten.")
        sys.exit(1)

    def _count(suffix):
        return sum(1 for rel in done.done.values()
                   if rel.endswith(suffix) and (out / rel).exists())
    print(f"\nFertig. Neu exportiert: {stats['new']}, übersprungen: {stats['skipped']}.")
    print(f"Im Archiv: {_count('.eml')} Mails, {_count('.ics')} Termine, "
          f"{_count('.vcf')} Kontakte.")
    print(f"Ordner: {out.resolve()}")


if __name__ == "__main__":
    main()