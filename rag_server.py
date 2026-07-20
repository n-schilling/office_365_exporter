#!/usr/bin/env python3
"""
rag_server.py – lokaler Dienst für die RAG-Suche (Teams + Outlook).

Lädt den von rag_index.py erzeugten Index, liefert eine UI unter
http://localhost:8000 und spricht Ollama lokal an. Zwei Modi:
  • Semantisch suchen  – nur Retrieval (sofort, keine Halluzination)
  • Frage beantworten  – Retrieval + lokale Generierung mit Quellenangaben

Der Dienst liefert auch die Quelldateien aus (Teams-HTML rendert im Browser,
Outlook-.eml öffnet im Mailprogramm), damit die Links über http funktionieren.

    ollama serve
    ollama pull bge-m3 && ollama pull qwen2.5:14b-instruct
    pip3 install numpy requests
    python3 rag_server.py --store rag_store \
        --teams teams_export --outlook outlook_export

Optionen: --embed-model bge-m3  --chat-model qwen2.5:14b-instruct
          --ollama http://localhost:11434  --port 8000
"""

import sys
import json
import argparse
from pathlib import Path
from urllib.parse import urlsplit, unquote, quote
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import numpy as np
import requests

import corpus  # noqa: F401  (gleiche Umgebung wie der Index)

# Auf Windows nutzt die Konsole standardmäßig eine Legacy-Codepage (z. B. cp1252),
# und bei Umleitung in eine Datei die Locale-Kodierung. Beides lässt print() an
# Unicode-Zeichen (·, →, … usw.) mit UnicodeEncodeError scheitern. UTF-8 erzwingen
# (auf macOS/Linux ein No-op).
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

STATE = {}

SYSTEM_PROMPT = (
    "Du beantwortest Fragen ausschließlich anhand des bereitgestellten Kontexts aus "
    "E-Mails und Teams-Nachrichten. Antworte auf Deutsch, knapp und präzise. Belege "
    "Aussagen mit Quellennummern in eckigen Klammern, z. B. [2]. Wenn der Kontext die "
    "Frage nicht beantwortet, sage das ausdrücklich und rate nicht."
)


# --------------------------------------------------------------------------
# Reine Retrieval-Logik (ohne Netz/Disk – testbar)
# --------------------------------------------------------------------------
def build_mask(meta, person, dfrom, dto, src):
    n = len(meta)
    mask = np.ones(n, dtype=bool)
    if src and src != "all":
        mask &= np.fromiter((c["src"] == src for c in meta), dtype=bool, count=n)
    if person:
        pl = person.lower()
        mask &= np.fromiter((pl in (c.get("ppl") or "") for c in meta), dtype=bool, count=n)
    if dfrom is not None or dto is not None:
        def ok(c):
            ts = c.get("ts")
            if ts is None:
                return False
            if dfrom is not None and ts < dfrom:
                return False
            if dto is not None and ts > dto:
                return False
            return True
        mask &= np.fromiter((ok(c) for c in meta), dtype=bool, count=n)
    return mask


def rank(V, qvec, mask, k):
    sims = V @ qvec
    sims = np.where(mask, sims, -1.0)
    order = np.argsort(sims)[::-1][:k]
    return [(int(i), float(sims[i])) for i in order if sims[i] > -1.0]


def browse(meta, mask, k):
    idxs = [i for i in range(len(meta)) if mask[i]]
    idxs.sort(key=lambda i: (meta[i].get("ts") is None, -(meta[i].get("ts") or 0)))
    return [(i, None) for i in idxs[:k]]


def src_link(c):
    rel = "/".join(quote(seg) for seg in c["rel"].split("/"))
    return f'/src/{c["root"]}/{rel}'


def hit_dict(c, score):
    return {"who": c["who"], "date": c["date"], "title": c["title"], "ctx": c["ctx"],
            "src": c["src"], "link": src_link(c), "preview": (c["text"] or "")[:600],
            "score": score}


# --------------------------------------------------------------------------
# Ollama
# --------------------------------------------------------------------------
def embed_query(text):
    url, model = STATE["ollama"], STATE["embed_model"]
    r = requests.post(f"{url}/api/embed", json={"model": model, "input": [text]}, timeout=120)
    r.raise_for_status()
    data = r.json()
    vec = (data.get("embeddings") or [data.get("embedding")])[0]
    v = np.asarray(vec, dtype="float32")
    nrm = np.linalg.norm(v)
    return v / nrm if nrm else v


def chat(query, hits_meta):
    url, model = STATE["ollama"], STATE["chat_model"]
    ctx = "\n\n".join(
        f'[{n}] ({c["date"]}, {c["who"]}, {c["src"]} – {c["ctx"]}) {c["text"]}'
        for n, c in enumerate(hits_meta, 1))
    user = f"Kontext:\n{ctx}\n\nFrage: {query}"
    r = requests.post(f"{url}/api/chat", json={
        "model": model, "stream": False, "options": {"temperature": 0.2},
        "messages": [{"role": "system", "content": SYSTEM_PROMPT},
                     {"role": "user", "content": user}],
    }, timeout=600)
    r.raise_for_status()
    return r.json()["message"]["content"]


# --------------------------------------------------------------------------
# HTTP-Handler
# --------------------------------------------------------------------------
class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_body(self):
        n = int(self.headers.get("Content-Length") or 0)
        return json.loads(self.rfile.read(n) or b"{}")

    # ---- GET: UI + Quelldateien ----
    def do_GET(self):
        path = unquote(urlsplit(self.path).path)
        if path in ("/", "/index.html"):
            body = UI_HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if path.startswith("/src/"):
            return self._serve_source(path)
        self.send_error(404)

    def _serve_source(self, path):
        parts = path[len("/src/"):].split("/", 1)
        if len(parts) != 2:
            return self.send_error(404)
        root, rel = parts
        base = {"teams": STATE.get("teams_dir"), "outlook": STATE.get("outlook_dir")}.get(root)
        if not base:
            return self.send_error(404)
        base = Path(base).resolve()
        target = (base / rel).resolve()
        if base not in target.parents and target != base:
            return self.send_error(403)          # Pfad-Ausbruch verhindern
        if not target.is_file():
            return self.send_error(404)
        ctype = ("text/html; charset=utf-8" if target.suffix == ".html"
                 else "message/rfc822" if target.suffix == ".eml"
                 else "text/calendar; charset=utf-8" if target.suffix == ".ics"
                 else "text/vcard; charset=utf-8" if target.suffix == ".vcf"
                 else "application/octet-stream")
        data = target.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    # ---- POST: Suche / Antwort ----
    def do_POST(self):
        path = urlsplit(self.path).path
        try:
            body = self._read_body()
        except Exception:
            return self._json({"error": "Ungültige Anfrage."}, 400)
        if path == "/api/search":
            return self._api_search(body)
        if path == "/api/answer":
            return self._api_answer(body)
        self.send_error(404)

    def _filters(self, body):
        return (body.get("query", "").strip(),
                body.get("person", "").strip(),
                body.get("fromTs"), body.get("toTs"),
                body.get("src", "all"))

    def _api_search(self, body):
        query, person, dfrom, dto, src = self._filters(body)
        k = int(body.get("k", 30))
        meta, V = STATE["meta"], STATE["V"]
        mask = build_mask(meta, person, dfrom, dto, src)
        try:
            pairs = rank(V, embed_query(query), mask, k) if query else browse(meta, mask, k)
        except requests.exceptions.RequestException:
            return self._json({"error": f"Ollama nicht erreichbar ({STATE['ollama']}). "
                                        f"Läuft 'ollama serve'?"}, 502)
        return self._json({"hits": [hit_dict(meta[i], s) for i, s in pairs],
                           "semantic": bool(query)})

    def _api_answer(self, body):
        query, person, dfrom, dto, src = self._filters(body)
        if not query:
            return self._json({"error": "Bitte eine Frage eingeben."}, 400)
        k = int(body.get("k", 8))
        meta, V = STATE["meta"], STATE["V"]
        mask = build_mask(meta, person, dfrom, dto, src)
        try:
            pairs = rank(V, embed_query(query), mask, k)
            if not pairs:
                return self._json({"answer": "Keine passenden Quellen im gewählten Filter.",
                                   "sources": []})
            picked = [meta[i] for i, _ in pairs]
            answer = chat(query, picked)
        except requests.exceptions.RequestException:
            return self._json({"error": f"Ollama nicht erreichbar ({STATE['ollama']}). "
                                        f"Läuft 'ollama serve' und sind die Modelle geladen?"}, 502)
        sources = [hit_dict(c, s) for c, (_, s) in zip(picked, pairs, strict=True)]
        return self._json({"answer": answer, "sources": sources})


UI_HTML = r"""<!DOCTYPE html>
<html lang="de"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>RAG-Suche · Teams + Outlook</title>
<style>
:root{color-scheme:light}*{box-sizing:border-box}
body{margin:0;font:15px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;color:#1b1b1f;background:#f6f7f9}
header{position:sticky;top:0;background:#fff;border-bottom:1px solid #e3e5e8;padding:14px 20px;z-index:2}
h1{margin:0 0 10px;font-size:17px}
.primary{display:flex;flex-wrap:wrap;gap:10px;align-items:flex-end}
.field{display:flex;flex-direction:column;gap:3px}
.field label{font-size:11px;color:#8a8f98;text-transform:uppercase;letter-spacing:.03em}
input,select{padding:9px 11px;font-size:14px;border:1px solid #cfd3d8;border-radius:8px;outline:none;background:#fff}
input:focus,select:focus{border-color:#2b6cb0;box-shadow:0 0 0 3px rgba(43,108,176,.12)}
#person{min-width:190px}
#q{flex:1;min-width:240px}
.qrow{display:flex;gap:8px;align-items:center;margin-top:10px}
button{padding:10px 14px;font-size:14px;border:1px solid #2b6cb0;border-radius:8px;background:#2b6cb0;color:#fff;cursor:pointer;white-space:nowrap}
button.secondary{background:#fff;color:#2b6cb0}
button:disabled{opacity:.5;cursor:default}
.chips{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-top:10px}
.chip{padding:5px 12px;border:1px solid #d4d8dd;border-radius:999px;background:#fff;font-size:13px;cursor:pointer;color:#3b3f46}
.chip.on{background:#2b6cb0;border-color:#2b6cb0;color:#fff}
#stats{color:#9aa0a6;font-size:12px;margin-left:auto}
main{max-width:900px;margin:0 auto;padding:16px}
.hint{color:#9aa0a6;padding:24px 4px}
.answer{background:#fff;border:1px solid #d9e2ec;border-left:4px solid #2b6cb0;border-radius:11px;padding:14px 16px;margin:8px 0 14px;white-space:pre-wrap;word-wrap:break-word}
.answer a{color:#2b6cb0;text-decoration:none;font-weight:600}
.srclabel{font-size:12px;color:#8a8f98;text-transform:uppercase;letter-spacing:.03em;margin:10px 0 4px}
.rec{background:#fff;border:1px solid #ececef;border-radius:11px;margin:8px 0;padding:12px 14px}
.rtop{display:flex;align-items:center;gap:9px;flex-wrap:wrap;margin-bottom:4px}
.num{font-weight:700;color:#2b6cb0;margin-right:2px}
.badge{font-size:11px;padding:2px 9px;border-radius:6px;font-weight:600}
.badge.teams{background:#efe7fb;color:#6b3fa0}.badge.outlook{background:#e6f0fb;color:#1f5fa6}
.badge.kalender{background:#e3f3ec;color:#1f7a4d}.badge.kontakte{background:#fdeee3;color:#b5651d}
.who{font-weight:600}.when{color:#9aa0a6;font-size:12.5px;margin-left:auto}
.score{color:#aab0b8;font-size:11px}
.title{font-weight:600;margin:2px 0;word-wrap:break-word}
.title a{color:#1b1b1f;text-decoration:none}.title a:hover{color:#2b6cb0;text-decoration:underline}
.ctx{font-size:12px;color:#8a8f98;margin-bottom:4px}
.snip{color:#26282c;word-wrap:break-word;overflow-wrap:anywhere}
.snip mark,.title mark{background:#ffe9a8;padding:0 1px;border-radius:2px}
</style></head>
<body>
<header>
  <h1>RAG-Suche · Teams + Outlook</h1>
  <div class="primary">
    <div class="field"><label>Person</label><input id="person" type="text" placeholder="Name oder E-Mail…" autocomplete="off"></div>
    <div class="field"><label>Von</label><input id="from" type="date"></div>
    <div class="field"><label>Bis</label><input id="to" type="date"></div>
  </div>
  <div class="qrow">
    <input id="q" type="search" placeholder="Frage stellen oder Begriffe suchen…" autocomplete="off">
    <button id="ask">Frage beantworten</button>
    <button id="find" class="secondary">Semantisch suchen</button>
  </div>
  <div class="chips">
    <span class="chip on" data-src="all">Alle Quellen</span>
    <span class="chip" data-src="teams">Teams</span>
    <span class="chip" data-src="outlook">Mail</span>
    <span class="chip" data-src="kalender">Kalender</span>
    <span class="chip" data-src="kontakte">Kontakte</span>
    <span id="stats"></span>
  </div>
</header>
<main><div id="out"><p class="hint">Frage stellen oder Begriffe suchen. Optional vorab nach Person und Datum filtern.</p></div></main>
<script>
let src='all';
const qEl=document.getElementById('q'), personEl=document.getElementById('person');
const fromEl=document.getElementById('from'), toEl=document.getElementById('to');
const out=document.getElementById('out'), askBtn=document.getElementById('ask'), findBtn=document.getElementById('find');
function esc(s){return (s||'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}
function toks(q){return q.toLowerCase().split(/\s+/).filter(Boolean);}
function reEsc(t){return t.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');}
function hi(s,t){let h=esc(s||'');for(const x of t)h=h.replace(new RegExp('('+reEsc(x)+')','ig'),'<mark>$1</mark>');return h;}
function dayStart(s){return s?new Date(s+'T00:00:00').getTime()/1000:null;}
function dayEnd(s){return s?new Date(s+'T23:59:59').getTime()/1000:null;}
function filters(){return {person:personEl.value.trim(),fromTs:dayStart(fromEl.value),toTs:dayEnd(toEl.value),src};}

function card(h,t,n){
  return '<div class="rec"><div class="rtop">'
    +(n?'<span class="num">['+n+']</span>':'')
    +'<span class="badge '+h.src+'">'+({teams:'Teams',outlook:'Mail',kalender:'Kalender',kontakte:'Kontakte'}[h.src]||h.src)+'</span>'
    +'<span class="who">'+esc(h.who)+'</span>'
    +(h.score!=null?'<span class="score">'+h.score.toFixed(3)+'</span>':'')
    +'<span class="when">'+esc(h.date)+'</span></div>'
    +'<div class="title"><a href="'+h.link+'" target="_blank" rel="noopener">'+hi(h.title,t)+'</a></div>'
    +'<div class="ctx">'+esc(h.ctx)+'</div>'
    +'<div class="snip">'+hi((h.preview||'').slice(0,300),t)+(h.preview&&h.preview.length>300?' …':'')+'</div></div>';
}
async function post(url,payload){
  const r=await fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
  return r.json();
}
function busy(b){askBtn.disabled=b;findBtn.disabled=b;}

async function doSearch(){
  const q=qEl.value.trim(); busy(true);
  out.innerHTML='<p class="hint">Suche…</p>';
  try{
    const d=await post('/api/search',{query:q,k:30,...filters()});
    if(d.error){out.innerHTML='<p class="hint">'+esc(d.error)+'</p>';return;}
    const t=toks(q);
    out.innerHTML = d.hits.length ? ('<p class="srclabel">'+d.hits.length+' Treffer'
      +(d.semantic?' (semantisch)':' (gefiltert)')+'</p>'+d.hits.map(h=>card(h,t,null)).join(''))
      : '<p class="hint">Keine Treffer.</p>';
  }catch(e){out.innerHTML='<p class="hint">Fehler: '+esc(e.message)+'</p>';}
  busy(false);
}
async function doAsk(){
  const q=qEl.value.trim();
  if(!q){qEl.focus();return;}
  busy(true);
  out.innerHTML='<p class="hint">Antwort wird erzeugt – lokales Modell, kann etwas dauern…</p>';
  try{
    const d=await post('/api/answer',{query:q,k:8,...filters()});
    if(d.error){out.innerHTML='<p class="hint">'+esc(d.error)+'</p>';return;}
    const sources=d.sources||[];
    let ans=esc(d.answer||'');
    ans=ans.replace(/\[(\d+)\]/g,(m,n)=>{const i=+n-1;return sources[i]?'<a href="'+sources[i].link+'" target="_blank" rel="noopener">['+n+']</a>':m;});
    let html='<div class="answer">'+ans+'</div>';
    if(sources.length) html+='<p class="srclabel">Quellen</p>'+sources.map((h,i)=>card(h,[],i+1)).join('');
    out.innerHTML=html;
  }catch(e){out.innerHTML='<p class="hint">Fehler: '+esc(e.message)+'</p>';}
  busy(false);
}
askBtn.addEventListener('click',doAsk);
findBtn.addEventListener('click',doSearch);
qEl.addEventListener('keydown',e=>{if(e.key==='Enter')doAsk();});
document.querySelectorAll('.chip').forEach(ch=>ch.addEventListener('click',()=>{
  document.querySelectorAll('.chip').forEach(x=>x.classList.remove('on'));ch.classList.add('on');src=ch.dataset.src;
}));
qEl.focus();
</script>
</body></html>
"""


def load_store(store):
    sp = Path(store)
    import sqlite3
    con = sqlite3.connect(f"file:{sp / 'corpus.db'}?mode=ro", uri=True)
    con.row_factory = sqlite3.Row
    meta = [dict(r) for r in con.execute("SELECT * FROM chunks ORDER BY id")]
    con.close()
    # Vektoren liegen als float16 – fürs Skalarprodukt nach float32
    V = np.load(sp / "vectors.npy").astype("float32", copy=False)
    return meta, V


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--store", default="rag_store")
    ap.add_argument("--teams", default="teams_export")
    ap.add_argument("--outlook", default="outlook_export")
    ap.add_argument("--embed-model", default="bge-m3")
    ap.add_argument("--chat-model", default="qwen2.5:14b-instruct")
    ap.add_argument("--ollama", default="http://localhost:11434")
    ap.add_argument("--port", type=int, default=8000)
    a = ap.parse_args()

    if not (Path(a.store) / "vectors.npy").exists():
        raise SystemExit(f"Kein Index in '{a.store}'. Zuerst: python3 rag_index.py")
    meta, V = load_store(a.store)
    STATE.update(meta=meta, V=V, teams_dir=a.teams, outlook_dir=a.outlook,
                 embed_model=a.embed_model, chat_model=a.chat_model, ollama=a.ollama)
    print(f"Index geladen: {len(meta)} Chunks, {V.shape[1]} Dimensionen.")
    print(f"Embedding: {a.embed_model} · Antwort: {a.chat_model} · Ollama: {a.ollama}")
    print(f"Öffne http://localhost:{a.port}")
    ThreadingHTTPServer(("127.0.0.1", a.port), Handler).serve_forever()


if __name__ == "__main__":
    main()
