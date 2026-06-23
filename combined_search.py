#!/usr/bin/env python3
"""
Kombinierte Volltextsuche über Teams- UND Outlook-Export in EINER search.html.

Liest beide Export-Ordner einmal ein und erzeugt eine eigenständige Suchseite mit
eingebettetem Index. Filter-Reihenfolge wie gewünscht: zuerst Person und Datum,
danach Inhalt und Komponente (Teams/Outlook). Verlinkt direkt auf die jeweilige
Quelldatei (Teams-HTML bzw. Outlook-.eml) – relativ zum Speicherort der Suchseite.

Nur Standardbibliothek – keine Installation nötig.

    python3 combined_search.py [teams-ordner] [outlook-ordner] [-o ausgabe.html]

Standard: teams_export, outlook_export. Die Ausgabe wird per Default in den
gemeinsamen übergeordneten Ordner beider Exporte geschrieben (combined_search.html),
damit die relativen Links funktionieren. Die Datei danach nicht relativ zu den
Export-Ordnern verschieben, sonst brechen die Links.
"""

import os
import sys
import re
import json
import email
import html as html_lib
from email import policy
from email.parser import BytesParser
from email.utils import getaddresses
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote
from html.parser import HTMLParser

# Auf Windows nutzt die Konsole standardmäßig eine Legacy-Codepage (z. B. cp1252),
# und bei Umleitung in eine Datei die Locale-Kodierung. Beides lässt print() an
# Unicode-Zeichen wie →, · oder … mit UnicodeEncodeError scheitern. UTF-8 erzwingen
# (auf macOS/Linux ein No-op).
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

BODY_CAP = 4000
CATS = {"1on1", "group", "meeting", "channels"}
CAT_LABEL = {"1on1": "1:1-Chat", "group": "Gruppenchat",
             "meeting": "Besprechung", "channels": "Kanal"}
MAIL_DIR = "E-Mail"   # Outlook-Export legt den Postfachbaum hierunter ab
_BLOCK = {"br", "p", "div", "li", "tr"}


# ===========================================================================
# Teams: exportierte Konversations-HTML parsen
# ===========================================================================
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
    """'YYYY-MM-DD HH:MM' (lokale Zeit aus dem Teams-Export) -> Epoch oder None."""
    try:
        return datetime.strptime(s.strip(), "%Y-%m-%d %H:%M").timestamp()
    except Exception:
        return None


def read_teams(root, out_dir, people):
    recs = []
    # nur echte Konversations-HTML – Index/Suche und versteckte Ordner (.imgcache,
    # .deltastate) überspringen; deren Inhalte sind keine Konversationen
    files = [p for p in sorted(root.rglob("*.html"))
             if p.name not in ("index.html", "search.html")
             and not any(seg.startswith(".") for seg in p.relative_to(root).parts)]
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
        ctx = (f"Kanal: {title}" if cat == "channels"
               else CAT_LABEL.get(cat, "Teams"))
        href = link(p, out_dir)
        for m in msgs:
            who = m["n"] or "(unbekannt)"
            if m["n"]:
                people.add(m["n"])
            recs.append({
                "src": "teams",
                "who": who,
                "ppl": (m["n"] + " " + title).lower(),
                "ts": parse_local(m["t"]),
                "d": m["t"],
                "title": title,
                "ctx": ctx,
                "x": m["x"][:BODY_CAP],
                "p": href,
            })
    return recs


# ===========================================================================
# Outlook: .eml parsen
# ===========================================================================
def strip_html(s):
    s = re.sub(r"(?is)<(script|style)\b.*?</\1>", " ", s)
    s = re.sub(r"(?is)<br\s*/?>|</p>|</div>|</li>|</tr>", " ", s)
    s = re.sub(r"<[^>]+>", " ", s)
    return html_lib.unescape(s)


def collapse(s, cap=BODY_CAP):
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


def read_outlook(root, out_dir, people):
    recs = []
    for p in sorted(root.rglob("*.eml")):
        try:
            with open(p, "rb") as f:
                msg = BytesParser(policy=policy.default).parse(f)
        except Exception:
            continue
        fn, fe = addr_people(msg, "from")
        tn, te = addr_people(msg, "to", "cc")
        who = (fn[0] if fn else (fe[0] if fe else "")) or "(unbekannt)"
        for x in fn + fe + tn + te:
            people.add(x)
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
        parts = rel.split("/")
        if parts and parts[0] == MAIL_DIR:
            parts = parts[1:]            # "E-Mail/" nur für die Anzeige entfernen
        folder = "/".join(parts[:-1]) if len(parts) > 1 else "(Stamm)"
        recs.append({
            "src": "outlook",
            "who": who,
            "ppl": " ".join(fn + fe + tn + te).lower(),
            "ts": ts,
            "d": disp,
            "title": hdr(msg, "subject") or "(kein Betreff)",
            "ctx": folder,
            "x": extract_body(msg),
            "p": link(p, out_dir),
        })
    return recs


# ===========================================================================
# Kalender (.ics) und Kontakte (.vcf) – liegen im Outlook-Export
# ===========================================================================
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


def read_calendar(root, out_dir, people):
    recs = []
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
        segs = p.relative_to(root).as_posix().split("/")
        cal = segs[1] if len(segs) >= 3 and segs[0] == "kalender" else "Kalender"
        for x in [org_cn] + att_names:
            if x:
                people.add(x)
        for x in [org_mail] + att_mails:
            if x:
                people.add(x)
        text = ((f"Ort: {location}. " if location else "") + description).strip()
        recs.append({
            "src": "kalender",
            "who": org_cn or org_mail or "(unbekannt)",
            "ppl": " ".join(x for x in ([org_cn, org_mail] + att_names + att_mails) if x).lower(),
            "ts": ts, "d": disp,
            "title": summary or "(kein Betreff)",
            "ctx": f"Kalender: {cal}",
            "x": text[:BODY_CAP],
            "p": link(p, out_dir),
        })
    return recs


def read_contacts(root, out_dir, people):
    recs = []
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
        segs = p.relative_to(root).as_posix().split("/")
        folder = segs[1] if len(segs) >= 3 and segs[0] == "kontakte" else ""
        people.add(fn)
        for e in emails:
            people.add(e)
        text = " · ".join(x for x in ([org, title] + emails + tels + ([note] if note else [])) if x)
        recs.append({
            "src": "kontakte",
            "who": org or title or "Kontakt",
            "ppl": " ".join([fn] + emails).lower(),
            "ts": None, "d": "",
            "title": fn,
            "ctx": f"Kontakte: {folder}" if folder else "Kontakte",
            "x": text[:BODY_CAP],
            "p": link(p, out_dir),
        })
    return recs


# ===========================================================================
# Gemeinsam
# ===========================================================================
def link(path, out_dir):
    try:
        rel = os.path.relpath(path, start=out_dir).replace(os.sep, "/")
        return "/".join(quote(seg) for seg in rel.split("/"))
    except ValueError:
        return Path(path).as_uri()


def build(teams_dir, outlook_dir, output):
    out_dir = output.parent.resolve()
    people = set()
    recs = []
    counts = {"teams": 0, "outlook": 0, "kalender": 0, "kontakte": 0}
    tp, op = Path(teams_dir), Path(outlook_dir)
    if tp.is_dir():
        r = read_teams(tp.resolve(), out_dir, people)
        recs += r
        counts["teams"] = len(r)
        print(f"  Teams:    {len(r)} Nachrichten aus {teams_dir}")
    else:
        print(f"  Teams-Ordner übersprungen (nicht gefunden): {teams_dir}")
    if op.is_dir():
        opr = op.resolve()
        m = read_outlook(opr, out_dir, people)
        c = read_calendar(opr, out_dir, people)
        k = read_contacts(opr, out_dir, people)
        recs += m + c + k
        counts.update(outlook=len(m), kalender=len(c), kontakte=len(k))
        print(f"  Mail:     {len(m)} Mails aus {outlook_dir}")
        print(f"  Kalender: {len(c)} Termine")
        print(f"  Kontakte: {len(k)} Personen")
    else:
        print(f"  Outlook-Ordner übersprungen (nicht gefunden): {outlook_dir}")

    recs.sort(key=lambda r: (r["ts"] is None, -(r["ts"] or 0)))   # neueste zuerst, undatierte zuletzt

    index = {
        "generated": datetime.now().isoformat(timespec="seconds"),
        "people": sorted(people, key=str.lower),
        "recs": recs,
    }
    payload = json.dumps(index, ensure_ascii=False).replace("</", "<\\/")
    html = TEMPLATE.replace("/*__INDEX__*/", payload)
    output.write_text(html, encoding="utf-8")
    return output, counts


TEMPLATE = r"""<!DOCTYPE html>
<html lang="de"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Teams + Outlook · Suche</title>
<style>
:root{color-scheme:light}
*{box-sizing:border-box}
body{margin:0;font:15px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;color:#1b1b1f;background:#f6f7f9}
header{position:sticky;top:0;background:#fff;border-bottom:1px solid #e3e5e8;padding:14px 20px;z-index:2}
h1{margin:0 0 10px;font-size:17px}
.primary{display:flex;flex-wrap:wrap;gap:10px;align-items:flex-end}
.field{display:flex;flex-direction:column;gap:3px}
.field label{font-size:11px;color:#8a8f98;text-transform:uppercase;letter-spacing:.03em}
input,select{padding:9px 11px;font-size:14px;border:1px solid #cfd3d8;border-radius:8px;outline:none;background:#fff}
input:focus,select:focus{border-color:#2b6cb0;box-shadow:0 0 0 3px rgba(43,108,176,.12)}
#person{min-width:200px}
#q{width:100%;margin-top:10px;padding:11px 13px;font-size:15px}
.chips{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-top:10px}
.chip{padding:5px 12px;border:1px solid #d4d8dd;border-radius:999px;background:#fff;font-size:13px;cursor:pointer;color:#3b3f46}
.chip.on{background:#2b6cb0;border-color:#2b6cb0;color:#fff}
#stats{color:#9aa0a6;font-size:12px;margin-left:auto}
#summary{color:#5b5f66;font-size:13px;margin:0 0 4px}
main{max-width:900px;margin:0 auto;padding:16px}
.hint{color:#9aa0a6;padding:24px 4px}
.rec{background:#fff;border:1px solid #ececef;border-radius:11px;margin:10px 0;padding:12px 14px}
.rtop{display:flex;align-items:center;gap:9px;flex-wrap:wrap;margin-bottom:4px}
.badge{font-size:11px;padding:2px 9px;border-radius:6px;font-weight:600}
.badge.teams{background:#efe7fb;color:#6b3fa0}
.badge.outlook{background:#e6f0fb;color:#1f5fa6}
.badge.kalender{background:#e3f3ec;color:#1f7a4d}
.badge.kontakte{background:#fdeee3;color:#b5651d}
.who{font-weight:600}
.when{color:#9aa0a6;font-size:12.5px;margin-left:auto}
.title{font-weight:600;margin:2px 0;word-wrap:break-word;overflow-wrap:anywhere}
.title a{color:#1b1b1f;text-decoration:none}
.title a:hover{color:#2b6cb0;text-decoration:underline}
.ctx{font-size:12px;color:#8a8f98;margin-bottom:4px;word-break:break-word}
.snip{color:#26282c;word-wrap:break-word;overflow-wrap:anywhere}
.snip mark,.title mark{background:#ffe9a8;padding:0 1px;border-radius:2px}
.more{padding:10px 4px;color:#9aa0a6;font-size:12.5px}
</style></head>
<body>
<header>
  <h1>Teams + Outlook · Suche</h1>
  <div class="primary">
    <div class="field"><label>Person</label>
      <input id="person" type="text" list="ppl" placeholder="Name oder E-Mail…" autocomplete="off">
      <datalist id="ppl"></datalist></div>
    <div class="field"><label>Von</label><input id="from" type="date"></div>
    <div class="field"><label>Bis</label><input id="to" type="date"></div>
  </div>
  <input id="q" type="search" placeholder="Inhalt durchsuchen… (mehrere Wörter = alle müssen vorkommen)" autocomplete="off">
  <div class="chips">
    <span class="chip on" data-src="all">Alle Quellen</span>
    <span class="chip" data-src="teams">Teams</span>
    <span class="chip" data-src="outlook">Mail</span>
    <span class="chip" data-src="kalender">Kalender</span>
    <span class="chip" data-src="kontakte">Kontakte</span>
    <span id="stats"></span>
  </div>
</header>
<main>
  <p id="summary"></p>
  <div id="results"><p class="hint">Person, Datum, Inhalt oder Quelle wählen…</p></div>
</main>
<script type="application/json" id="idx">/*__INDEX__*/</script>
<script>
const DATA = JSON.parse(document.getElementById('idx').textContent);
const recs = DATA.recs || [];
let src = 'all';
const qEl = document.getElementById('q');
const personEl = document.getElementById('person');
const fromEl = document.getElementById('from');
const toEl = document.getElementById('to');
const out = document.getElementById('results');

(DATA.people || []).forEach(p=>{ const o=document.createElement('option'); o.value=p; document.getElementById('ppl').appendChild(o); });
const LABEL = {teams:'Teams', outlook:'Mail', kalender:'Kalender', kontakte:'Kontakte'};
const cnt = {teams:0, outlook:0, kalender:0, kontakte:0};
for(const r of recs){ if(cnt[r.src]!=null) cnt[r.src]++; }
document.getElementById('stats').textContent =
  recs.length+' Einträge · Teams '+cnt.teams+' · Mail '+cnt.outlook+' · Kalender '+cnt.kalender+' · Kontakte '+cnt.kontakte;

function esc(s){return (s||'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}
function toks(q){return q.toLowerCase().split(/\s+/).filter(Boolean);}
function allIn(hay,t){hay=(hay||'').toLowerCase();return t.every(x=>hay.includes(x));}
function reEsc(t){return t.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');}
function hi(s,t){let h=esc(s||'');for(const x of t){h=h.replace(new RegExp('('+reEsc(x)+')','ig'),'<mark>$1</mark>');}return h;}
function snippet(text,t){
  text=text||'';
  if(!t.length) return esc(text.slice(0,240))+(text.length>240?' …':'');
  const low=text.toLowerCase(); let idx=-1;
  for(const x of t){const i=low.indexOf(x); if(i>=0&&(idx<0||i<idx))idx=i;}
  let start=0,pre='',suf='';
  if(idx>90){start=idx-70;pre='… ';}
  let end=Math.min(text.length,start+260); if(end<text.length)suf=' …';
  return pre+hi(text.slice(start,end),t)+suf;
}
function dayStart(s){return s?new Date(s+'T00:00:00').getTime()/1000:null;}
function dayEnd(s){return s?new Date(s+'T23:59:59').getTime()/1000:null;}

function render(){
  const q=qEl.value.trim();
  const pq=personEl.value.trim().toLowerCase();
  const fromTs=dayStart(fromEl.value), toTs=dayEnd(toEl.value);
  const active = q||pq||fromEl.value||toEl.value||src!=='all';
  if(!active){ out.innerHTML='<p class="hint">Person, Datum, Inhalt oder Quelle wählen…</p>'; document.getElementById('summary').textContent=''; return; }
  const ct=toks(q);
  const LIMIT=500;
  let shown=0,total=0,frag='';
  for(const r of recs){
    if(src!=='all' && r.src!==src) continue;
    if(pq && !(r.ppl||'').includes(pq)) continue;
    if(fromTs!==null){ if(r.ts===null || r.ts<fromTs) continue; }
    if(toTs!==null){ if(r.ts===null || r.ts>toTs) continue; }
    if(ct.length && !allIn((r.title||'')+' '+(r.x||''),ct)) continue;
    total++;
    if(shown<LIMIT){
      frag += '<div class="rec"><div class="rtop">'
            + '<span class="badge '+r.src+'">'+(LABEL[r.src]||r.src)+'</span>'
            + '<span class="who">'+esc(r.who)+'</span>'
            + '<span class="when">'+esc(r.d)+'</span></div>'
            + '<div class="title"><a href="'+r.p+'" target="_blank" rel="noopener">'+hi(r.title,ct)+'</a></div>'
            + '<div class="ctx">'+esc(r.ctx)+'</div>'
            + '<div class="snip">'+snippet(r.x,ct)+'</div></div>';
      shown++;
    }
  }
  document.getElementById('summary').textContent = total+' Treffer'+(total>shown?(' (zeige '+shown+')'):'');
  let extra = total>shown ? '<div class="more">… '+(total-shown)+' weitere – Filter verfeinern</div>' : '';
  out.innerHTML = frag ? frag+extra : '<p class="hint">Keine Treffer.</p>';
}

let timer;
function debounced(){clearTimeout(timer);timer=setTimeout(render,120);}
qEl.addEventListener('input',debounced);
personEl.addEventListener('input',debounced);
fromEl.addEventListener('change',render);
toEl.addEventListener('change',render);
document.querySelectorAll('.chip').forEach(ch=>ch.addEventListener('click',()=>{
  document.querySelectorAll('.chip').forEach(x=>x.classList.remove('on'));
  ch.classList.add('on'); src=ch.dataset.src; render();
}));
personEl.focus();
</script>
</body></html>
"""


def main():
    args = sys.argv[1:]
    output = None
    pos = []
    i = 0
    while i < len(args):
        if args[i] in ("-o", "--out") and i + 1 < len(args):
            output = args[i + 1]
            i += 2
        else:
            pos.append(args[i])
            i += 1
    teams_dir = pos[0] if len(pos) > 0 else "teams_export"
    outlook_dir = pos[1] if len(pos) > 1 else "outlook_export"

    tp, op = Path(teams_dir), Path(outlook_dir)
    if not tp.is_dir() and not op.is_dir():
        raise SystemExit(f"Weder '{teams_dir}' noch '{outlook_dir}' gefunden – nichts zu tun.")

    if output:
        outp = Path(output)
    else:
        existing = [p.resolve() for p in (tp, op) if p.is_dir()]
        if len(existing) == 2:
            try:
                base = Path(os.path.commonpath(existing))
            except ValueError:
                base = Path.cwd()
            if not base.is_dir():
                base = base.parent
        else:
            base = existing[0].parent
        outp = base / "combined_search.html"

    print(f"Erzeuge kombinierte Suche → {outp}")
    out, counts = build(teams_dir, outlook_dir, outp)
    total = sum(counts.values())
    print(f"\nFertig. {total} Einträge gesamt (Teams {counts['teams']}, Mail {counts['outlook']}, "
          f"Kalender {counts['kalender']}, Kontakte {counts['kontakte']}).")
    print(f"Suche öffnen: {out.resolve()}")


if __name__ == "__main__":
    main()