#!/usr/bin/env python3
"""
Erzeugt eine eigenständige Volltextsuche (search.html) für einen Teams-Export.

Hintergrund: Eine lokal per Doppelklick geöffnete HTML-Seite (file://) darf aus
Sicherheitsgründen keine anderen lokalen Dateien per JavaScript nachladen. Darum
liest dieses Skript die Export-HTMLs EINMAL ein, extrahiert die Nachrichtentexte
und schreibt search.html mit eingebettetem Index. Diese Datei ist offline lauffähig
und verlinkt direkt auf die jeweiligen Konversations-HTMLs.

Nur Standardbibliothek – keine Installation nötig.

    python3 teams_search.py [export-ordner]      # Standard: teams_export

Ergebnis: <export-ordner>/search.html  – im Browser öffnen und tippen.
"""

import sys
import json
import html as html_lib
from datetime import datetime
from pathlib import Path
from html.parser import HTMLParser

# Auf Windows nutzt die Konsole standardmäßig eine Legacy-Codepage (z. B. cp1252),
# und bei Umleitung in eine Datei die Locale-Kodierung. Beides lässt print() an
# Unicode-Zeichen (… usw.) mit UnicodeEncodeError scheitern. UTF-8 erzwingen
# (auf macOS/Linux ein No-op).
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

CATS = {"1on1", "group", "meeting", "channels"}
_BLOCK = {"br", "p", "div", "li", "tr"}


class ConvParser(HTMLParser):
    """Liest Titel, Untertitel und einzelne Nachrichten aus einer Export-HTML."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.title = ""
        self.subtitle = ""
        self.msgs = []           # [{"n": sender, "t": time, "x": text}]
        self._depth = 0
        self._cur = None
        self._msg_depth = None
        self._in_body = False
        self._body_depth = None
        self._capture = None     # "name" | "time" | None
        self._in_h1 = False
        self._in_sub = False
        self._nb = []            # name buffer
        self._tb = []            # time buffer
        self._bb = []            # body buffer
        self._h1 = []
        self._sub = []

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
                self._bb.append(" ")   # Block-Trenner im Body
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
        elif tag == "p" and "sub" in cls:
            self._in_sub = True

    def handle_endtag(self, tag):
        if tag == "span":
            self._capture = None
        elif tag == "h1":
            self._in_h1 = False
        elif tag == "p":
            self._in_sub = False
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
        elif self._in_sub:
            self._sub.append(data)
        elif self._capture == "name":
            self._nb.append(data)
        elif self._capture == "time":
            self._tb.append(data)
        elif self._in_body:
            self._bb.append(data)

    def finish(self):
        self.title = "".join(self._h1).strip()
        self.subtitle = "".join(self._sub).strip()


def parse_file(path):
    raw = path.read_text(encoding="utf-8", errors="replace")
    p = ConvParser()
    try:
        p.feed(raw)
        p.finish()
        title, msgs = p.title, p.msgs
    except Exception:
        title, msgs = "", []
    if not title:
        title = path.stem.rsplit("__", 1)[0]
    if not msgs:
        # Fallback: alle Tags entfernen und als ein durchsuchbarer Block ablegen
        import re
        text = re.sub(r"<[^>]+>", " ", raw)
        text = " ".join(html_lib.unescape(text).split())
        if text:
            msgs = [{"n": "", "t": "", "x": text[:20000]}]
    return title, msgs


def build(export_dir):
    root = Path(export_dir)
    if not root.is_dir():
        raise SystemExit(f"Ordner nicht gefunden: {root}")
    convs = []
    n_msgs = 0
    files = [p for p in sorted(root.rglob("*.html"))
             if p.name not in ("index.html", "search.html")]
    for p in files:
        rel = p.relative_to(root).as_posix()
        top = rel.split("/")[0]
        cat = top if top in CATS else "other"
        title, msgs = parse_file(p)
        n_msgs += len(msgs)
        convs.append({"title": title, "cat": cat, "path": rel, "msgs": msgs})
        print(f"  {rel}: {len(msgs)} Nachrichten")

    index = {"generated": datetime.now().isoformat(timespec="seconds"), "convs": convs}
    payload = json.dumps(index, ensure_ascii=False).replace("</", "<\\/")
    html = TEMPLATE.replace("/*__INDEX__*/", payload)
    out = root / "search.html"
    out.write_text(html, encoding="utf-8")
    return out, len(convs), n_msgs


TEMPLATE = r"""<!DOCTYPE html>
<html lang="de"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Teams-Export · Suche</title>
<style>
:root{color-scheme:light}
*{box-sizing:border-box}
body{margin:0;font:15px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;color:#1b1b1f;background:#f6f7f9}
header{position:sticky;top:0;background:#fff;border-bottom:1px solid #e3e5e8;padding:14px 20px;z-index:2}
h1{margin:0 0 10px;font-size:17px}
#q{width:100%;padding:11px 13px;font-size:15px;border:1px solid #cfd3d8;border-radius:9px;outline:none}
#q:focus{border-color:#2b6cb0;box-shadow:0 0 0 3px rgba(43,108,176,.12)}
.row{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-top:10px}
.chip{padding:5px 11px;border:1px solid #d4d8dd;border-radius:999px;background:#fff;font-size:13px;cursor:pointer;color:#3b3f46}
.chip.on{background:#2b6cb0;border-color:#2b6cb0;color:#fff}
#stats{color:#9aa0a6;font-size:12px;margin-left:auto}
#summary{color:#5b5f66;font-size:13px;margin:0 0 4px}
main{max-width:880px;margin:0 auto;padding:16px}
.hint{color:#9aa0a6;padding:24px 4px}
.conv{background:#fff;border:1px solid #ececef;border-radius:11px;margin:12px 0;overflow:hidden}
.conv-head{display:flex;align-items:center;gap:9px;padding:11px 14px;border-bottom:1px solid #f0f0f2;background:#fafbfc}
.conv-head a{color:#1b1b1f;text-decoration:none;font-weight:600;flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.conv-head a:hover{color:#2b6cb0;text-decoration:underline}
.cnt{color:#9aa0a6;font-size:12px}
.badge{font-size:11px;padding:2px 8px;border-radius:6px;font-weight:600;white-space:nowrap}
.badge.\31 on1{background:#e6f0fb;color:#1f5fa6}
.badge.group{background:#eaf6ec;color:#2e7d4f}
.badge.meeting{background:#f3ecfa;color:#7a4ec0}
.badge.channels{background:#fdeee3;color:#b5651d}
.badge.other{background:#eef0f2;color:#5b5f66}
.hit{padding:9px 14px;border-top:1px solid #f4f4f6}
.hit:first-of-type{border-top:none}
.who{font-weight:600;margin-right:8px}
.when{color:#9aa0a6;font-size:12px}
.snip{margin-top:3px;color:#26282c;word-wrap:break-word;overflow-wrap:anywhere}
.snip mark{background:#ffe9a8;padding:0 1px;border-radius:2px}
.more{padding:8px 14px;color:#9aa0a6;font-size:12.5px}
</style></head>
<body>
<header>
  <h1>Teams-Export · Suche</h1>
  <input id="q" type="search" placeholder="Suchbegriff… (mehrere Wörter = alle müssen vorkommen)" autocomplete="off">
  <div class="row">
    <span class="chip on" data-cat="all">Alle</span>
    <span class="chip" data-cat="1on1">1:1</span>
    <span class="chip" data-cat="group">Gruppe</span>
    <span class="chip" data-cat="meeting">Meeting</span>
    <span class="chip" data-cat="channels">Kanäle</span>
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
const convs = DATA.convs || [];
let curCat = 'all';
const input = document.getElementById('q');
const out = document.getElementById('results');
const totalMsgs = convs.reduce((a,c)=>a+(c.msgs?c.msgs.length:0),0);
document.getElementById('stats').textContent = convs.length+' Konversationen · '+totalMsgs+' Nachrichten';

function esc(s){return (s||'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}
function tokens(q){return q.toLowerCase().split(/\s+/).filter(Boolean);}
function allIn(hay,toks){hay=(hay||'').toLowerCase();return toks.every(t=>hay.includes(t));}
function hrefFor(p){return p.split('/').map(encodeURIComponent).join('/');}
function reEsc(t){return t.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');}

function snippet(text,toks){
  text = text||'';
  const low = text.toLowerCase();
  let idx = -1;
  for(const t of toks){const i=low.indexOf(t); if(i>=0&&(idx<0||i<idx))idx=i;}
  let start=0, pre='', suf='';
  if(idx>90){start=idx-70; pre='… ';}
  let end=Math.min(text.length, start+240);
  if(end<text.length) suf=' …';
  let html = esc(text.slice(start,end));
  for(const t of toks){
    html = html.replace(new RegExp('('+reEsc(t)+')','ig'),'<mark>$1</mark>');
  }
  return pre+html+suf;
}

function render(){
  const q = input.value.trim();
  if(!q){ out.innerHTML='<p class="hint">Suchbegriff eingeben…</p>'; document.getElementById('summary').textContent=''; return; }
  const toks = tokens(q);
  const LIMIT = 400;
  let rendered=0, hitMsgs=0, hitConvs=0, frag='';
  for(const c of convs){
    if(curCat!=='all' && c.cat!==curCat) continue;
    const msgs = c.msgs||[];
    const matches = msgs.filter(m=>allIn((m.x||'')+' '+(m.n||''),toks));
    const titleHit = allIn(c.title||'',toks);
    if(!matches.length && !titleHit) continue;
    hitConvs++;
    let body='';
    const show = matches.slice(0, Math.max(0, LIMIT-rendered));
    for(const m of show){
      body += '<div class="hit"><span class="who">'+esc(m.n)+'</span><span class="when">'+esc(m.t)+'</span><div class="snip">'+snippet(m.x,toks)+'</div></div>';
      rendered++; hitMsgs++;
    }
    if(matches.length>show.length) body += '<div class="more">… '+(matches.length-show.length)+' weitere Treffer hier</div>';
    if(!matches.length && titleHit) body = '<div class="more">Titel-Treffer – Konversation öffnen</div>';
    const cat = c.cat||'other';
    const label = ({'1on1':'1:1','group':'Gruppe','meeting':'Meeting','channels':'Kanal'}[cat]||cat);
    frag += '<div class="conv"><div class="conv-head"><span class="badge '+cat+'">'+esc(label)+'</span>'
          + '<a href="'+hrefFor(c.path)+'" target="_blank" rel="noopener">'+esc(c.title||c.path)+'</a>'
          + '<span class="cnt">'+matches.length+'</span></div>'+body+'</div>';
    if(rendered>=LIMIT) break;
  }
  document.getElementById('summary').textContent = hitMsgs+' Treffer in '+hitConvs+' Konversationen'+(rendered>=LIMIT?' (Anzeige gekürzt)':'');
  out.innerHTML = frag || '<p class="hint">Keine Treffer.</p>';
}

let timer;
input.addEventListener('input',()=>{clearTimeout(timer);timer=setTimeout(render,120);});
document.querySelectorAll('.chip').forEach(ch=>ch.addEventListener('click',()=>{
  document.querySelectorAll('.chip').forEach(x=>x.classList.remove('on'));
  ch.classList.add('on'); curCat=ch.dataset.cat; render();
}));
input.focus();
</script>
</body></html>
"""


def main():
    export_dir = sys.argv[1] if len(sys.argv) > 1 else "teams_export"
    print(f"Lese Export aus: {export_dir}")
    out, n_convs, n_msgs = build(export_dir)
    print(f"\nFertig. {n_convs} Konversationen, {n_msgs} Nachrichten indexiert.")
    print(f"Suche öffnen: {out.resolve()}")


if __name__ == "__main__":
    main()
