#!/usr/bin/env python3
"""
Erzeugt eine eigenständige Volltextsuche (search.html) für einen Outlook-Export.

Wie beim Teams-Generator gilt: Eine lokal geöffnete HTML-Seite (file://) darf keine
anderen lokalen Dateien per JavaScript nachladen. Darum liest dieses Skript die
.eml-Dateien EINMAL ein, extrahiert Betreff, Absender/Empfänger, Datum und den
Textkörper und schreibt search.html mit eingebettetem Index. Anhänge und Inline-
Bilder werden NICHT indexiert (der Index bleibt schlank; die vollständige Mail steckt
weiterhin in der .eml). Die Suche verlinkt direkt auf die jeweilige .eml.

Nur Standardbibliothek – keine Installation nötig.

    python3 outlook_search.py [export-ordner]     # Standard: outlook_export

Ergebnis: <export-ordner>/search.html  – im Browser öffnen und tippen.
Ein Klick auf eine Mail öffnet die .eml (i. d. R. im Mailprogramm).
"""

import sys
import re
import json
import email
import html as html_lib
from email import policy
from email.parser import BytesParser
from datetime import datetime
from pathlib import Path

# Auf Windows nutzt die Konsole standardmäßig eine Legacy-Codepage (z. B. cp1252),
# und bei Umleitung in eine Datei die Locale-Kodierung. Beides lässt print() an
# Unicode-Zeichen (… usw.) mit UnicodeEncodeError scheitern. UTF-8 erzwingen
# (auf macOS/Linux ein No-op).
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

BODY_CAP = 4000   # max. Zeichen Textkörper pro Mail im Index (hält die Datei klein)


def strip_html(s):
    s = re.sub(r"(?is)<(script|style)\b.*?</\1>", " ", s)   # Style/Script-Blöcke raus
    s = re.sub(r"(?is)<br\s*/?>|</p>|</div>|</li>|</tr>", " ", s)
    s = re.sub(r"<[^>]+>", " ", s)
    return html_lib.unescape(s)


def collapse(s, cap=BODY_CAP):
    s = " ".join((s or "").split())
    return s[:cap]


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
        cs = part.get_content_charset() or "utf-8"
        return b.decode(cs, errors="replace")
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


def parse_eml(path):
    try:
        with open(path, "rb") as f:
            msg = BytesParser(policy=policy.default).parse(f)
    except Exception:
        return None
    subject = hdr(msg, "subject") or "(kein Betreff)"
    frm = hdr(msg, "from")
    to = collapse(hdr(msg, "to") + (("; " + hdr(msg, "cc")) if msg["cc"] else ""), 300)
    raw_date = hdr(msg, "date")
    ts = 0.0
    disp = raw_date
    try:
        dt = email.utils.parsedate_to_datetime(raw_date)
        if dt is not None:
            ts = dt.timestamp()
            disp = dt.astimezone().strftime("%Y-%m-%d %H:%M")
    except Exception:
        pass
    body = extract_body(msg)
    return {"s": subject, "f": frm, "o": to, "d": disp, "ts": ts, "x": body}


def build(export_dir):
    root = Path(export_dir)
    if not root.is_dir():
        raise SystemExit(f"Ordner nicht gefunden: {root}")
    files = sorted(root.rglob("*.eml"))
    msgs = []
    folders = set()
    for p in files:
        rec = parse_eml(p)
        if rec is None:
            print(f"  übersprungen (nicht lesbar): {p.name}")
            continue
        rel = p.relative_to(root).as_posix()
        folder = rel.rsplit("/", 1)[0] if "/" in rel else "."
        rec["p"] = rel
        rec["dir"] = folder
        folders.add(folder)
        msgs.append(rec)

    msgs.sort(key=lambda m: m["ts"], reverse=True)   # neueste zuerst
    for m in msgs:
        m.pop("ts", None)

    index = {
        "generated": datetime.now().isoformat(timespec="seconds"),
        "folders": sorted(folders, key=str.lower),
        "msgs": msgs,
    }
    payload = json.dumps(index, ensure_ascii=False).replace("</", "<\\/")
    html = TEMPLATE.replace("/*__INDEX__*/", payload)
    out = root / "search.html"
    out.write_text(html, encoding="utf-8")
    return out, len(msgs), len(folders)


TEMPLATE = r"""<!DOCTYPE html>
<html lang="de"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Outlook-Export · Suche</title>
<style>
:root{color-scheme:light}
*{box-sizing:border-box}
body{margin:0;font:15px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;color:#1b1b1f;background:#f6f7f9}
header{position:sticky;top:0;background:#fff;border-bottom:1px solid #e3e5e8;padding:14px 20px;z-index:2}
h1{margin:0 0 10px;font-size:17px}
#q{width:100%;padding:11px 13px;font-size:15px;border:1px solid #cfd3d8;border-radius:9px;outline:none}
#q:focus{border-color:#2b6cb0;box-shadow:0 0 0 3px rgba(43,108,176,.12)}
.row{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin-top:10px}
select{padding:6px 9px;border:1px solid #cfd3d8;border-radius:8px;background:#fff;font-size:13px;max-width:60vw}
label{font-size:13px;color:#5b5f66}
#stats{color:#9aa0a6;font-size:12px;margin-left:auto}
#summary{color:#5b5f66;font-size:13px;margin:0 0 4px}
main{max-width:900px;margin:0 auto;padding:16px}
.hint{color:#9aa0a6;padding:24px 4px}
.mail{background:#fff;border:1px solid #ececef;border-radius:11px;margin:10px 0;padding:12px 14px}
.subj{font-weight:600;font-size:15px;margin:0 0 3px;word-wrap:break-word;overflow-wrap:anywhere}
.subj a{color:#1b1b1f;text-decoration:none}
.subj a:hover{color:#2b6cb0;text-decoration:underline}
.meta{font-size:12.5px;color:#5b5f66;margin-bottom:5px;word-wrap:break-word;overflow-wrap:anywhere}
.meta .when{color:#9aa0a6}
.badge{display:inline-block;font-size:11px;padding:2px 8px;border-radius:6px;background:#eef2f7;color:#3b5b80;margin-top:4px;word-break:break-all}
.snip{color:#26282c;word-wrap:break-word;overflow-wrap:anywhere}
.snip mark,.subj mark{background:#ffe9a8;padding:0 1px;border-radius:2px}
.more{padding:10px 4px;color:#9aa0a6;font-size:12.5px}
</style></head>
<body>
<header>
  <h1>Outlook-Export · Suche</h1>
  <input id="q" type="search" placeholder="Suchbegriff… (mehrere Wörter = alle müssen vorkommen)" autocomplete="off">
  <div class="row">
    <label>Ordner: <select id="folder"><option value="__all__">Alle</option></select></label>
    <span id="stats"></span>
  </div>
</header>
<main>
  <p id="summary"></p>
  <div id="results"><p class="hint">Suchbegriff eingeben…</p></div>
</main>
<script type="application/json" id="idx">/*__INDEX__*/</script>
<script>
const DATA = JSON.parse(document.getElementById('idx').textContent);
const msgs = DATA.msgs || [];
let curFolder = '__all__';
const input = document.getElementById('q');
const out = document.getElementById('results');
const sel = document.getElementById('folder');

(DATA.folders || []).forEach(f=>{
  const o=document.createElement('option');
  o.value=f; o.textContent=(f==='.'?'(Stammordner)':f);
  sel.appendChild(o);
});
document.getElementById('stats').textContent = msgs.length+' Mails · '+(DATA.folders||[]).length+' Ordner';

function esc(s){return (s||'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}
function tokens(q){return q.toLowerCase().split(/\s+/).filter(Boolean);}
function allIn(hay,toks){hay=(hay||'').toLowerCase();return toks.every(t=>hay.includes(t));}
function hrefFor(p){return p.split('/').map(encodeURIComponent).join('/');}
function reEsc(t){return t.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');}
function hi(s,toks){
  let html=esc(s||'');
  for(const t of toks){ html=html.replace(new RegExp('('+reEsc(t)+')','ig'),'<mark>$1</mark>'); }
  return html;
}
function snippet(text,toks){
  text=text||'';
  const low=text.toLowerCase();
  let idx=-1;
  for(const t of toks){const i=low.indexOf(t); if(i>=0&&(idx<0||i<idx))idx=i;}
  let start=0,pre='',suf='';
  if(idx>90){start=idx-70; pre='… ';}
  let end=Math.min(text.length,start+260);
  if(end<text.length) suf=' …';
  return pre+hi(text.slice(start,end),toks)+suf;
}
function inFolder(m){
  if(curFolder==='__all__') return true;
  return m.dir===curFolder || (m.dir||'').startsWith(curFolder+'/');
}
function render(){
  const q=input.value.trim();
  if(!q){ out.innerHTML='<p class="hint">Suchbegriff eingeben…</p>'; document.getElementById('summary').textContent=''; return; }
  const toks=tokens(q);
  const LIMIT=500;
  let shown=0,total=0,frag='';
  for(const m of msgs){
    if(!inFolder(m)) continue;
    const hay=(m.s||'')+' '+(m.f||'')+' '+(m.o||'')+' '+(m.x||'');
    if(!allIn(hay,toks)) continue;
    total++;
    if(shown<LIMIT){
      const folder=(m.dir==='.'?'(Stammordner)':m.dir);
      frag += '<div class="mail">'
            + '<p class="subj"><a href="'+hrefFor(m.p)+'" target="_blank" rel="noopener">'+hi(m.s,toks)+'</a></p>'
            + '<div class="meta">'+hi(m.f,toks)+(m.o?' → '+hi(m.o,toks):'')+' · <span class="when">'+esc(m.d)+'</span></div>'
            + '<div class="snip">'+snippet(m.x,toks)+'</div>'
            + '<span class="badge">'+esc(folder)+'</span>'
            + '</div>';
      shown++;
    }
  }
  document.getElementById('summary').textContent = total+' Treffer'+(total>shown?(' (zeige '+shown+')'):'');
  let extra = total>shown ? '<div class="more">… '+(total-shown)+' weitere Treffer – Suche verfeinern</div>' : '';
  out.innerHTML = frag ? frag+extra : '<p class="hint">Keine Treffer.</p>';
}
let timer;
input.addEventListener('input',()=>{clearTimeout(timer);timer=setTimeout(render,120);});
sel.addEventListener('change',()=>{curFolder=sel.value;render();});
input.focus();
</script>
</body></html>
"""


def main():
    export_dir = sys.argv[1] if len(sys.argv) > 1 else "outlook_export"
    print(f"Lese Export aus: {export_dir}")
    out, n_msgs, n_folders = build(export_dir)
    print(f"\nFertig. {n_msgs} Mails aus {n_folders} Ordnern indexiert.")
    print(f"Suche öffnen: {out.resolve()}")


if __name__ == "__main__":
    main()
