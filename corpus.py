#!/usr/bin/env python3
"""
corpus.py – gemeinsame Datengrundlage für die lokale RAG-Suche.

Liest Teams-Export (HTML) und Outlook-Export (.eml) in einheitliche Datensätze
und zerlegt lange Texte in überlappende Chunks. Wird von rag_index.py (Embeddings)
und rag_server.py (Retrieval/Antwort) genutzt. Nur Standardbibliothek.
"""

import re
import email
import html as html_lib
import hashlib
from email import policy
from email.parser import BytesParser
from email.utils import getaddresses
from datetime import datetime, timezone
from pathlib import Path
from html.parser import HTMLParser

CATS = {"1on1", "group", "meeting", "channels"}
CAT_LABEL = {"1on1": "1:1-Chat", "group": "Gruppenchat",
             "meeting": "Besprechung", "channels": "Kanal"}
_BLOCK = {"br", "p", "div", "li", "tr"}
SAFETY_CAP = 500_000   # absurd lange Einzeltexte begrenzen (vor dem Chunking)


# --------------------------------------------------------------------------
# Teams: exportierte Konversations-HTML parsen
# --------------------------------------------------------------------------
class ConvParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.title = ""
        self.msgs = []
        self._depth = 0
        self._cur = None
        self._msg_depth = None
        self._in_body = False
        self._body_depth = None
        self._capture = None
        self._in_h1 = False
        self._nb, self._tb, self._bb, self._h1 = [], [], [], []

    @staticmethod
    def _classes(attrs):
        for k, v in attrs:
            if k == "class":
                return (v or "").split()
        return []

    def handle_starttag(self, tag, attrs):
        cls = self._classes(attrs)
        if tag == "div":
            self._depth += 1
            if self._cur is None and "msg" in cls:
                self._cur = True
                self._msg_depth = self._depth
                self._nb, self._tb, self._bb = [], [], []
            if self._cur is not None and "body" in cls and not self._in_body:
                self._in_body = True
                self._body_depth = self._depth
            elif self._in_body:
                self._bb.append(" ")
        elif tag == "span":
            if self._cur is not None and not self._in_body:
                if "name" in cls:
                    self._capture = "name"
                elif "time" in cls:
                    self._capture = "time"
        elif tag in _BLOCK and self._in_body:
            self._bb.append(" ")
        elif tag == "h1":
            self._in_h1 = True

    def handle_endtag(self, tag):
        if tag == "span":
            self._capture = None
        elif tag == "h1":
            self._in_h1 = False
        elif tag == "div":
            if self._in_body and self._depth == self._body_depth:
                self._in_body = False
            if self._cur is not None and self._depth == self._msg_depth:
                text = " ".join("".join(self._bb).split())
                name = "".join(self._nb).strip()
                time = "".join(self._tb).strip()
                if text:
                    self.msgs.append({"n": name, "t": time, "x": text})
                self._cur = None
            self._depth -= 1

    def handle_data(self, data):
        if self._in_h1:
            self._h1.append(data)
        elif self._capture == "name":
            self._nb.append(data)
        elif self._capture == "time":
            self._tb.append(data)
        elif self._in_body:
            self._bb.append(data)

    def finish(self):
        self.title = "".join(self._h1).strip()


def parse_local(s):
    try:
        return datetime.strptime((s or "").strip(), "%Y-%m-%d %H:%M").timestamp()
    except Exception:
        return None


def load_teams(root_dir):
    recs = []
    root = Path(root_dir)
    files = [p for p in sorted(root.rglob("*.html"))
             if p.name not in ("index.html", "search.html")]
    for p in files:
        raw = p.read_text(encoding="utf-8", errors="replace")
        pr = ConvParser()
        try:
            pr.feed(raw)
            pr.finish()
            title, msgs = pr.title, pr.msgs
        except Exception:
            title, msgs = p.stem.rsplit("__", 1)[0], []
        rel = p.relative_to(root).as_posix()
        top = rel.split("/")[0]
        cat = top if top in CATS else "other"
        ctx = f"Kanal: {title}" if cat == "channels" else CAT_LABEL.get(cat, "Teams")
        for i, m in enumerate(msgs):
            recs.append({
                "uid": f"teams:{rel}:{i}", "src": "teams", "root": "teams", "rel": rel,
                "who": m["n"] or "(unbekannt)", "ppl": (m["n"] + " " + title).lower(),
                "ts": parse_local(m["t"]), "date": m["t"], "title": title, "ctx": ctx,
                "text": (m["x"] or "")[:SAFETY_CAP],
            })
    return recs


# --------------------------------------------------------------------------
# Outlook: .eml parsen
# --------------------------------------------------------------------------
def strip_html(s):
    s = re.sub(r"(?is)<(script|style)\b.*?</\1>", " ", s)
    s = re.sub(r"(?is)<br\s*/?>|</p>|</div>|</li>|</tr>", " ", s)
    s = re.sub(r"<[^>]+>", " ", s)
    return html_lib.unescape(s)


def collapse(s, cap=SAFETY_CAP):
    return " ".join((s or "").split())[:cap]


def hdr(msg, name):
    v = msg[name]
    return str(v).strip() if v is not None else ""


def decode_part(part):
    try:
        c = part.get_content()
        if isinstance(c, str):
            return c
    except Exception:
        pass
    try:
        b = part.get_payload(decode=True) or b""
        return b.decode(part.get_content_charset() or "utf-8", errors="replace")
    except Exception:
        return ""


def extract_body(msg):
    part = None
    try:
        part = msg.get_body(preferencelist=("plain", "html"))
    except Exception:
        part = None
    if part is None:
        for p in msg.walk():
            if (p.get_content_maintype() == "text"
                    and p.get_content_disposition() != "attachment"):
                part = p
                break
    if part is None:
        return ""
    text = decode_part(part)
    if part.get_content_type() == "text/html":
        text = strip_html(text)
    return collapse(text)


def addr_people(msg, *headers):
    raw = []
    for h in headers:
        vals = msg.get_all(h)
        if vals:
            raw += [str(v) for v in vals]
    names, emails = [], []
    for name, addr in getaddresses(raw):
        if name.strip():
            names.append(name.strip())
        if addr.strip():
            emails.append(addr.strip())
    return names, emails


def load_outlook(root_dir):
    recs = []
    root = Path(root_dir)
    for p in sorted(root.rglob("*.eml")):
        try:
            with open(p, "rb") as f:
                msg = BytesParser(policy=policy.default).parse(f)
        except Exception:
            continue
        fn, fe = addr_people(msg, "from")
        tn, te = addr_people(msg, "to", "cc")
        who = (fn[0] if fn else (fe[0] if fe else "")) or "(unbekannt)"
        raw_date = hdr(msg, "date")
        ts, disp = None, raw_date
        try:
            dt = email.utils.parsedate_to_datetime(raw_date)
            if dt is not None:
                ts = dt.timestamp()
                disp = dt.astimezone().strftime("%Y-%m-%d %H:%M")
        except Exception:
            pass
        rel = p.relative_to(root).as_posix()
        folder = rel.rsplit("/", 1)[0] if "/" in rel else "(Stamm)"
        recs.append({
            "uid": f"outlook:{rel}:0", "src": "outlook", "root": "outlook", "rel": rel,
            "who": who, "ppl": " ".join(fn + fe + tn + te).lower(),
            "ts": ts, "date": disp, "title": hdr(msg, "subject") or "(kein Betreff)",
            "ctx": folder, "text": extract_body(msg),
        })
    return recs


# --------------------------------------------------------------------------
# Kalender (.ics) und Kontakte (.vcf) – liegen im Outlook-Export
# --------------------------------------------------------------------------
def _unfold(text):
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    out = []
    for line in text.split("\n"):
        if line[:1] in (" ", "\t") and out:
            out[-1] += line[1:]
        else:
            out.append(line)
    return out


def _unescape(v):
    res, i = [], 0
    while i < len(v):
        ch = v[i]
        if ch == "\\" and i + 1 < len(v):
            res.append({"n": "\n", "N": "\n", ",": ",", ";": ";", "\\": "\\"}.get(v[i + 1], v[i + 1]))
            i += 2
        else:
            res.append(ch)
            i += 1
    return "".join(res)


def _prop(line):
    in_q = False
    for i, ch in enumerate(line):
        if ch == '"':
            in_q = not in_q
        elif ch == ":" and not in_q:
            name = line[:i].split(";", 1)[0].upper()
            return name, line[:i][len(name):], line[i + 1:]
    return None, None, None


def _pval(params, key):
    m = re.search(rf';{key}=("([^"]*)"|([^;:]*))', params or "", re.I)
    if not m:
        return ""
    return m.group(2) if m.group(2) is not None else (m.group(3) or "")


def _demail(v):
    return re.sub(r"(?i)^mailto:", "", (v or "").strip())


def _ics_when(val, dateonly):
    if not val:
        return None, ""
    try:
        if dateonly or (len(val) == 8 and val.isdigit()):
            dt = datetime.strptime(val[:8], "%Y%m%d")
            return dt.timestamp(), dt.strftime("%Y-%m-%d")
        utc = val.endswith("Z")
        dt = datetime.strptime(val.rstrip("Z")[:15], "%Y%m%dT%H%M%S")
        if utc:
            dt = dt.replace(tzinfo=timezone.utc)
            return dt.timestamp(), dt.astimezone().strftime("%Y-%m-%d %H:%M")
        return dt.timestamp(), dt.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return None, val


def load_calendar(root_dir):
    recs = []
    root = Path(root_dir)
    for p in sorted(root.rglob("*.ics")):
        summary = location = description = org_cn = org_mail = dtstart = ""
        dateonly = False
        att_names, att_mails = [], []
        for line in _unfold(p.read_text(encoding="utf-8", errors="replace")):
            name, params, value = _prop(line)
            if not name:
                continue
            if name == "SUMMARY":
                summary = _unescape(value)
            elif name == "LOCATION":
                location = _unescape(value)
            elif name == "DESCRIPTION":
                description = _unescape(value)
            elif name == "DTSTART":
                dtstart = value.strip()
                dateonly = "VALUE=DATE" in (params or "").upper()
            elif name == "ORGANIZER":
                org_cn, org_mail = _pval(params, "CN"), _demail(value)
            elif name == "ATTENDEE":
                cn, mail = _pval(params, "CN"), _demail(value)
                if cn:
                    att_names.append(cn)
                if mail:
                    att_mails.append(mail)
        ts, disp = _ics_when(dtstart, dateonly)
        rel = p.relative_to(root).as_posix()
        segs = rel.split("/")
        cal = segs[1] if len(segs) >= 3 and segs[0] == "kalender" else "Kalender"
        ppl = " ".join(x for x in ([org_cn, org_mail] + att_names + att_mails) if x).lower()
        text = ((f"Ort: {location}. " if location else "") + description).strip()
        recs.append({
            "uid": f"kalender:{rel}:0", "src": "kalender", "root": "outlook", "rel": rel,
            "who": org_cn or org_mail or "(unbekannt)", "ppl": ppl,
            "ts": ts, "date": disp, "title": summary or "(kein Betreff)",
            "ctx": f"Kalender: {cal}", "text": text[:SAFETY_CAP],
        })
    return recs


def load_contacts(root_dir):
    recs = []
    root = Path(root_dir)
    for p in sorted(root.rglob("*.vcf")):
        fn = org = title = note = given = family = ""
        emails, tels = [], []
        for line in _unfold(p.read_text(encoding="utf-8", errors="replace")):
            name, params, value = _prop(line)
            if not name:
                continue
            if name == "FN":
                fn = _unescape(value)
            elif name == "N":
                parts = [_unescape(x) for x in value.split(";")]
                family = parts[0] if len(parts) > 0 else ""
                given = parts[1] if len(parts) > 1 else ""
            elif name == "ORG":
                org = " · ".join(x for x in _unescape(value).split(";") if x)
            elif name == "TITLE":
                title = _unescape(value)
            elif name == "EMAIL":
                emails.append(value.strip())
            elif name == "TEL":
                tels.append(value.strip())
            elif name == "NOTE":
                note = _unescape(value)
        if not fn:
            fn = (given + " " + family).strip() or "(ohne Namen)"
        rel = p.relative_to(root).as_posix()
        segs = rel.split("/")
        folder = segs[1] if len(segs) >= 3 and segs[0] == "kontakte" else ""
        text = " · ".join(x for x in ([org, title] + emails + tels + ([note] if note else [])) if x)
        recs.append({
            "uid": f"kontakte:{rel}:0", "src": "kontakte", "root": "outlook", "rel": rel,
            "who": org or title or "Kontakt", "ppl": " ".join([fn] + emails).lower(),
            "ts": None, "date": "", "title": fn,
            "ctx": f"Kontakte: {folder}" if folder else "Kontakte",
            "text": text[:SAFETY_CAP],
        })
    return recs


# --------------------------------------------------------------------------
# Zusammenführen + Chunking
# --------------------------------------------------------------------------
def load_records(teams_dir, outlook_dir):
    recs = []
    if teams_dir and Path(teams_dir).is_dir():
        recs += load_teams(teams_dir)
    if outlook_dir and Path(outlook_dir).is_dir():
        recs += load_outlook(outlook_dir)     # .eml
        recs += load_calendar(outlook_dir)    # .ics
        recs += load_contacts(outlook_dir)    # .vcf
    return recs


def _split(text, size, overlap):
    text = text or ""
    if len(text) <= size:
        return [text.strip()] if text.strip() else []
    out, i, n = [], 0, len(text)
    while i < n:
        end = min(i + size, n)
        if end < n:
            sp = text.rfind(" ", i + int(size * 0.6), end)
            if sp != -1:
                end = sp
        piece = text[i:end].strip()
        if piece:
            out.append(piece)
        if end >= n:
            break
        i = max(end - overlap, i + 1)
    return out


def chunk_records(records, size=1500, overlap=200):
    """Eine Nachricht/Mail = Basis-Einheit; lange Texte in überlappende Stücke."""
    chunks = []
    for r in records:
        parts = _split(r["text"], size, overlap)
        for j, part in enumerate(parts):
            c = dict(r)
            c.pop("text", None)
            c["text"] = part
            c["cid"] = f'{r["uid"]}#{j}'
            chunks.append(c)
    return chunks


def embed_text(chunk):
    """Was tatsächlich eingebettet wird: Titel als Kontext + Chunk-Text."""
    return f'{chunk.get("title", "")}\n{chunk["text"]}'.strip()


def chunk_hash(chunk):
    return hashlib.sha1(embed_text(chunk).encode("utf-8")).hexdigest()
