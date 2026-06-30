#!/usr/bin/env python3
"""
rag_index.py – baut den Embedding-Index für die lokale RAG-Suche.

Liest beide Exporte (über corpus.py), bettet jeden Chunk per Ollama ein und legt
Vektoren + Metadaten in einem Store-Ordner ab. Inkrementell: bei erneutem Lauf
werden nur neue/geänderte Chunks neu berechnet (Abgleich über Inhalts-Hash).

    ollama serve                 # Ollama muss laufen
    ollama pull bge-m3           # mehrsprachiges Embedding-Modell (DE/EN)
    pip3 install numpy requests
    python3 rag_index.py [teams_export] [outlook_export] [--store rag_store]

Optionen: --model bge-m3  --ollama http://localhost:11434  --batch 64
"""

import sys
import json
import argparse
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

import numpy as np
import requests

import corpus

# Auf Windows nutzt die Konsole standardmäßig eine Legacy-Codepage (z. B. cp1252),
# und bei Umleitung in eine Datei die Locale-Kodierung. Beides lässt print() an
# Unicode-Zeichen wie → oder … mit UnicodeEncodeError scheitern. UTF-8 erzwingen
# (auf macOS/Linux ein No-op).
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

DEFAULT_MODEL = "bge-m3"
DEFAULT_OLLAMA = "http://localhost:11434"


def embed(texts, model, url, timeout=600):
    try:
        r = requests.post(f"{url}/api/embed",
                          json={"model": model, "input": texts}, timeout=timeout)
    except requests.exceptions.ConnectionError:
        raise SystemExit(f"Keine Verbindung zu Ollama unter {url}. "
                         f"Läuft 'ollama serve'?")
    if r.status_code == 404:
        raise SystemExit(f"Modell '{model}' nicht gefunden. Vorher: ollama pull {model}")
    r.raise_for_status()
    data = r.json()
    embs = data.get("embeddings")
    if embs is None and "embedding" in data:      # ältere Single-Form
        embs = [data["embedding"]]
    if not embs:
        raise SystemExit(f"Unerwartete Embedding-Antwort: {str(data)[:200]}")
    return embs


def load_old_vectors(store):
    mp, vp = Path(store) / "meta.json", Path(store) / "vectors.npy"
    if mp.exists() and vp.exists():
        try:
            meta = json.loads(mp.read_text(encoding="utf-8"))
            V = np.load(vp)
            return {c["hash"]: V[i] for i, c in enumerate(meta) if i < len(V)}
        except Exception:
            print("  Alter Index unlesbar – baue komplett neu.")
    return {}


def build_index(teams_dir, outlook_dir, store, model, url, batch=64):
    recs = corpus.load_records(teams_dir, outlook_dir)
    chunks = corpus.chunk_records(recs)
    if not chunks:
        raise SystemExit("Keine Inhalte gefunden – stimmen die Export-Ordner?")
    for c in chunks:
        c["hash"] = corpus.chunk_hash(c)

    old = load_old_vectors(store)
    vectors = [None] * len(chunks)
    # Pro eindeutigem Inhalts-Hash nur EINMAL einbetten und das Ergebnis auf alle
    # gleichen Chunks verteilen (identische Signaturen/Disclaimer kommen oft vor).
    uniq = {}                       # hash -> Liste der Chunk-Indizes mit diesem Hash
    for i, c in enumerate(chunks):
        v = old.get(c["hash"])
        if v is not None:
            vectors[i] = np.asarray(v, dtype="float32")
        else:
            uniq.setdefault(c["hash"], []).append(i)

    todo_groups = list(uniq.values())          # je eindeutiger Text: alle Zielindizes
    todo_texts = [corpus.embed_text(chunks[idxs[0]]) for idxs in todo_groups]
    new_total = sum(len(g) for g in todo_groups)
    print(f"{len(chunks)} Chunks: {len(chunks) - new_total} wiederverwendet, "
          f"{new_total} neu ({len(todo_texts)} eindeutig einzubetten).")

    if todo_texts:
        done = 0
        # Embedding ist GPU-gebunden und serialisiert auf einem Slot; mit zwei
        # Requests „in flight“ liegt immer schon einer in der Server-Queue, sodass
        # die GPU zwischen den Batches nicht leerläuft (kein Idle-Bubble).
        def run(b):
            texts = todo_texts[b:b + batch]
            return b, embed(texts, model, url)
        starts = range(0, len(todo_texts), batch)
        with ThreadPoolExecutor(max_workers=2) as ex:
            futs = {ex.submit(run, b): b for b in starts}
            for fut in as_completed(futs):
                b, vecs = fut.result()
                for k, vec in enumerate(vecs):
                    arr = np.asarray(vec, dtype="float32")
                    for i in todo_groups[b + k]:
                        vectors[i] = arr
                done += len(vecs)
                print(f"  … {done}/{len(todo_texts)} eingebettet", end="\r", flush=True)
        print()

    V = np.vstack(vectors).astype("float32")
    norms = np.linalg.norm(V, axis=1, keepdims=True)
    norms[norms == 0] = 1.0
    V = V / norms

    sp = Path(store)
    sp.mkdir(parents=True, exist_ok=True)
    np.save(sp / "vectors.npy", V)
    (sp / "meta.json").write_text(json.dumps(chunks, ensure_ascii=False), encoding="utf-8")
    (sp / "info.json").write_text(json.dumps({
        "model": model, "dim": int(V.shape[1]), "chunks": len(chunks),
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    return len(chunks), new_total, int(V.shape[1])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("teams", nargs="?", default="teams_export")
    ap.add_argument("outlook", nargs="?", default="outlook_export")
    ap.add_argument("--store", default="rag_store")
    ap.add_argument("--model", default=DEFAULT_MODEL)
    ap.add_argument("--ollama", default=DEFAULT_OLLAMA)
    ap.add_argument("--batch", type=int, default=64)
    a = ap.parse_args()

    print(f"Index → {a.store}  (Modell {a.model})")
    n, new, dim = build_index(a.teams, a.outlook, a.store, a.model, a.ollama, a.batch)
    print(f"\nFertig. {n} Chunks im Index ({dim} Dimensionen), davon {new} neu berechnet.")
    print(f"Jetzt: python3 rag_server.py --store {a.store}")


if __name__ == "__main__":
    main()
