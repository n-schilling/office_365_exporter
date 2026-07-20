#!/usr/bin/env python3
"""
mcp_server.py – expose the Teams + Outlook exports to Claude as an MCP server.

Instead of generating answers with a local LLM (as rag_server.py does with
qwen2.5), this server hands the *retrieval* to Claude as MCP tools and lets
Claude be the reasoning/answer layer. It reads the store built by rag_index.py:

    corpus.db     SQLite with all chunks + an FTS5 (BM25) full-text index and a
                  precomputed people table. Queried on demand – the server keeps
                  (almost) nothing in RAM and starts instantly.
    vectors.npy   float16 embedding matrix, memory-mapped – the OS pages in only
                  what a query touches.

Ranking backends, per query:
  • hybrid   – default when embeddings are available: FTS5/BM25 and semantic
               cosine ranking run side by side and are merged with Reciprocal
               Rank Fusion. Exact tokens (invoice numbers, names) and
               paraphrases both hit.
  • semantic – cosine only (needs numpy + Ollama for the query embedding).
  • lexical  – FTS5/BM25 only, standard library, no Ollama needed. Automatic
               fallback when Ollama is down.

Tools: search_messages, browse_messages, get_document, list_people,
read_source_file, corpus_stats. Every hit carries an o365:// resource URI;
the corresponding MCP resource returns the raw source file.

Install (SDK required; numpy/requests only for semantic/hybrid ranking):
    pip install mcp
    pip install numpy requests        # optional

Run (HTTP, default – one shared server for all Claude sessions):
    python3 mcp_server.py --store rag_store \
        --teams teams_export --outlook outlook_export
    # → MCP endpoint at http://127.0.0.1:8365/mcp

    Register in Claude Code (.mcp.json):
        {"mcpServers": {"office365-export":
            {"type": "http", "url": "http://127.0.0.1:8365/mcp"}}}

    The server binds to 127.0.0.1 and has no authentication – do NOT expose
    it on the network (--host 0.0.0.0) unless you know what you are doing:
    it serves your complete mail and chat history.

Run (stdio – auto-launched per client, the classic setup):
    python3 mcp_server.py --transport stdio [--store …]
"""

import re
import sys
import sqlite3
import argparse
from pathlib import Path
from datetime import datetime
from urllib.parse import quote, unquote

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

# Windows consoles default to a legacy code page; force UTF-8 so logging the
# Unicode in messages never raises (no-op on macOS/Linux).
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

STATE = {}          # populated in main(): db path, V (mmap), np, dirs, flags
mcp = FastMCP("office365-export")

_READONLY = ToolAnnotations(readOnlyHint=True, idempotentHint=True,
                            openWorldHint=False)
_WORD = re.compile(r"\w+", re.UNICODE)
_SOURCE_LABEL = {"teams": "Teams", "outlook": "Mail",
                 "kalender": "Kalender", "kontakte": "Kontakte"}
_RRF_K = 60                     # standard reciprocal-rank-fusion constant
_POOL_MIN, _POOL_MAX = 100, 1000  # candidate pool per backend before merging


def _db():
    """Fresh read-only connection per call – safe across FastMCP worker threads."""
    con = sqlite3.connect(f'file:{STATE["db"]}?mode=ro', uri=True)
    con.row_factory = sqlite3.Row
    return con


# --------------------------------------------------------------------------
# Filters (SQL WHERE fragments shared by all query tools)
# --------------------------------------------------------------------------
def _where(person, dfrom, dto, src):
    conds, params = [], []
    if src and src != "all":
        conds.append("src = ?")
        params.append(src)
    if person:
        conds.append("ppl LIKE ?")
        params.append(f"%{person.lower()}%")
    if dfrom is not None:
        conds.append("ts >= ?")                # also excludes NULL timestamps
        params.append(dfrom)
    if dto is not None:
        conds.append("ts <= ?")
        params.append(dto)
    return (" AND ".join(conds) or "1=1"), params


def _to_ts(s, end):
    s = (s or "").strip()
    if not s:
        return None
    try:
        dt = datetime.strptime(s, "%Y-%m-%d")
        if end:
            dt = dt.replace(hour=23, minute=59, second=59)
        return dt.timestamp()
    except ValueError:
        return None


# --------------------------------------------------------------------------
# Lexical backend: FTS5 / BM25
# --------------------------------------------------------------------------
def _fts_match(query):
    """Sanitize free text into an FTS5 OR-query of quoted tokens."""
    toks = _WORD.findall(query.lower())
    return " OR ".join(f'"{t}"' for t in toks)


def _lexical_rank(con, query, where, params, limit):
    match = _fts_match(query)
    if not match:
        return []
    sql = (f"SELECT c.id, bm25(chunks_fts) AS r FROM chunks_fts "
           f"JOIN chunks c ON c.id = chunks_fts.rowid "
           f"WHERE chunks_fts MATCH ? AND {where} ORDER BY r LIMIT ?")
    # bm25(): smaller = better; negate so every backend reports higher = better
    return [(row[0], -row[1]) for row in con.execute(sql, [match, *params, limit])]


# --------------------------------------------------------------------------
# Semantic backend: mmap'd float16 matrix, block-wise cosine scoring
# --------------------------------------------------------------------------
def _embed_query(text):
    import requests
    np = STATE["np"]
    r = requests.post(f"{STATE['ollama']}/api/embed",
                      json={"model": STATE["embed_model"], "input": [text]},
                      timeout=120)
    r.raise_for_status()
    data = r.json()
    vec = (data.get("embeddings") or [data.get("embedding")])[0]
    v = np.asarray(vec, dtype="float32")
    nrm = np.linalg.norm(v)
    return v / nrm if nrm else v


def _semantic_rank(con, query, where, params, limit):
    np, V = STATE["np"], STATE["V"]
    ids = np.fromiter((r[0] for r in
                       con.execute(f"SELECT id FROM chunks WHERE {where}", params)),
                      dtype=np.int64)
    if ids.size == 0:
        return []
    qvec = _embed_query(query)                       # may raise (Ollama down)
    sims = np.empty(ids.size, dtype=np.float32)
    B = 32768                                        # ~64 MB float16 per block
    for s in range(0, ids.size, B):
        block = ids[s:s + B] - 1                     # chunks.id → vector row
        sims[s:s + B] = V[block].astype(np.float32) @ qvec
    take = min(limit, ids.size)
    order = np.argpartition(-sims, take - 1)[:take]
    order = order[np.argsort(-sims[order])]
    return [(int(ids[o]), float(sims[o])) for o in order]


# --------------------------------------------------------------------------
# Fusion, dedupe, result shaping
# --------------------------------------------------------------------------
def _rrf_merge(*ranked_lists):
    """Reciprocal Rank Fusion: score = Σ 1/(K + rank). Ignores raw scales."""
    scores = {}
    for lst in ranked_lists:
        for rank, (cid, _) in enumerate(lst):
            scores[cid] = scores.get(cid, 0.0) + 1.0 / (_RRF_K + rank + 1)
    return sorted(scores.items(), key=lambda x: -x[1])


def _rank(con, query, where, params, k, offset, mode):
    """Ranked (chunk_id, score) list + the backend actually used."""
    pool = min(_POOL_MAX, max(_POOL_MIN, (offset + k) * 5))
    lex = sem = None
    if mode in ("auto", "hybrid", "semantic") and STATE.get("semantic"):
        try:
            sem = _semantic_rank(con, query, where, params, pool)
        except Exception as e:                       # Ollama down, timeout, …
            STATE["last_semantic_error"] = str(e)
            if mode == "semantic":
                raise
    if mode in ("auto", "hybrid", "lexical") or sem is None:
        lex = _lexical_rank(con, query, where, params, pool)
    if sem is not None and lex is not None:
        return _rrf_merge(sem, lex), "hybrid"
    if sem is not None:
        return sem, "semantic"
    return lex or [], "lexical"


def _dedupe_page(con, pairs, k, offset):
    """Collapse chunk hits to messages (best chunk wins), then page."""
    if not pairs:
        return []
    ids = [cid for cid, _ in pairs]
    uid_of = {}
    CHUNK = 500                                      # SQLite variable limit safety
    for s in range(0, len(ids), CHUNK):
        part = ids[s:s + CHUNK]
        q = ",".join("?" * len(part))
        uid_of.update((r[0], r[1]) for r in con.execute(
            f"SELECT id, uid FROM chunks WHERE id IN ({q})", part))
    seen, page = set(), []
    for cid, score in pairs:
        uid = uid_of.get(cid)
        if uid is None or uid in seen:
            continue
        seen.add(uid)
        if len(seen) > offset:
            page.append((cid, score))
            if len(page) >= k:
                break
    return page


def _source_uri(root, rel):
    """MCP resource URI for a source file (rel path percent-encoded)."""
    return f"o365://{root}/{quote(rel, safe='')}"


def _hit(row, score, preview_chars):
    h = {
        "uid": row["uid"],
        "source": row["src"],
        "source_label": _SOURCE_LABEL.get(row["src"], row["src"]),
        "who": row["who"],
        "date": row["date"],
        "title": row["title"],
        "context": row["ctx"],
        "path": row["rel"],
        "uri": _source_uri(row["root"], row["rel"]),
        "score": round(score, 4) if score is not None else None,
    }
    if preview_chars > 0:
        h["preview"] = (row["text"] or "")[:preview_chars]
    return h


def _rows_for(con, pairs, preview_chars):
    hits = []
    for cid, score in pairs:
        row = con.execute("SELECT * FROM chunks WHERE id = ?", (cid,)).fetchone()
        if row is not None:
            hits.append(_hit(row, score, preview_chars))
    return hits


def _join_chunks(rows):
    """Reassemble a message's full text from its overlapping chunks (by seq)."""
    text = ""
    for row in rows:
        piece = row["text"] or ""
        if not text:
            text = piece
            continue
        cut = 0
        for L in range(min(len(text), len(piece), 300), 0, -1):  # drop overlap
            if text[-L:] == piece[:L]:
                cut = L
                break
        text += piece[cut:]
    return text


def _message_text(con, uid):
    rows = con.execute("SELECT * FROM chunks WHERE uid = ? ORDER BY seq",
                       (uid,)).fetchall()
    return (rows[0] if rows else None), _join_chunks(rows)


def _resolve_source(source_root, rel):
    """Sandboxed path resolution for an export file. Returns (Path, error_str)."""
    base = {"teams": STATE.get("teams_dir"),
            "outlook": STATE.get("outlook_dir")}.get(source_root)
    if not base:
        return None, "source_root must be 'teams' or 'outlook'."
    base = Path(base).resolve()
    target = (base / rel).resolve()
    if base != target and base not in target.parents:      # prevent path escape
        return None, "Path outside the export directory."
    if not target.is_file():
        return None, f"File not found: {rel}"
    return target, None


def _read_window(target, offset, max_chars):
    """Read a byte window of a file without ever loading the whole file.

    Some exported Teams conversations exceed 100 MB; decoding them entirely
    would pin gigabytes in a long-lived server process.
    """
    total = target.stat().st_size
    start = max(0, offset)
    n = max(1, min(max_chars, 500000))
    with open(target, "rb") as f:
        f.seek(start)
        data = f.read(n)
    # A window may split a multi-byte UTF-8 sequence at either edge;
    # errors="replace" turns the clipped bytes into a replacement char.
    return data.decode("utf-8", errors="replace"), total, start, start + len(data) < total


# --------------------------------------------------------------------------
# MCP tools
# --------------------------------------------------------------------------
@mcp.tool(annotations=_READONLY)
def search_messages(query: str, person: str = "", date_from: str = "",
                    date_to: str = "", source: str = "all",
                    k: int = 12, offset: int = 0, mode: str = "auto",
                    preview_chars: int = 200) -> dict:
    """Search the exported Teams messages and Outlook mail/calendar/contacts.

    Hybrid ranking (BM25 + semantic embeddings, fused) when available. Results
    are deduped to one hit per message; use get_document(uid) for the full text
    and pass offset to page through more results.

    Args:
        query: Natural-language query or keywords (German or English).
        person: Optional. Filter to messages involving this name or email.
        date_from: Optional. Inclusive lower bound, "YYYY-MM-DD".
        date_to: Optional. Inclusive upper bound, "YYYY-MM-DD".
        source: One of "all", "teams", "outlook", "kalender", "kontakte".
        k: Number of results per page (default 12).
        offset: Results to skip, for pagination (default 0).
        mode: "auto" (hybrid if embeddings available, else lexical),
              "hybrid", "semantic", or "lexical".
        preview_chars: Preview length per hit (default 200; 0 disables previews).
    """
    con = _db()
    try:
        where, params = _where(person.strip(), _to_ts(date_from, False),
                               _to_ts(date_to, True), source)
        try:
            pairs, used = _rank(con, query.strip(), where, params,
                                max(1, k), max(0, offset), mode)
        except Exception as e:
            return {"error": f"Semantic ranking failed: {e}. "
                             f"Is Ollama running? Try mode='lexical'."}
        page = _dedupe_page(con, pairs, max(1, k), max(0, offset))
        return {"backend": used, "count": len(page), "offset": max(0, offset),
                "results": _rows_for(con, page, max(0, min(preview_chars, 2000)))}
    finally:
        con.close()


@mcp.tool(annotations=_READONLY)
def browse_messages(person: str = "", date_from: str = "", date_to: str = "",
                    source: str = "all", k: int = 30, offset: int = 0,
                    preview_chars: int = 200) -> dict:
    """List messages by filter, newest first, without a search query.

    Useful for "everything from <person> in <month>" or scanning a source.
    Pass offset to page through more results.

    Args:
        person: Optional name or email to filter by.
        date_from: Optional inclusive "YYYY-MM-DD" lower bound.
        date_to: Optional inclusive "YYYY-MM-DD" upper bound.
        source: One of "all", "teams", "outlook", "kalender", "kontakte".
        k: Max results per page (default 30).
        offset: Results to skip, for pagination (default 0).
        preview_chars: Preview length per hit (default 200; 0 disables previews).
    """
    con = _db()
    try:
        where, params = _where(person.strip(), _to_ts(date_from, False),
                               _to_ts(date_to, True), source)
        rows = con.execute(
            f"SELECT * FROM chunks WHERE seq = 0 AND {where} "
            f"ORDER BY (ts IS NULL), ts DESC LIMIT ? OFFSET ?",
            [*params, max(1, k), max(0, offset)]).fetchall()
        pc = max(0, min(preview_chars, 2000))
        return {"count": len(rows), "offset": max(0, offset),
                "results": [_hit(r, None, pc) for r in rows]}
    finally:
        con.close()


@mcp.tool(annotations=_READONLY)
def get_document(uid: str, context_before: int = 0, context_after: int = 0) -> dict:
    """Return the full text and metadata of a single message/mail by its uid.

    The uid comes from a search_messages / browse_messages result. For chat
    messages, context_before/context_after also return the neighboring messages
    of the same conversation – usually much cheaper than reading the whole
    conversation file.

    Args:
        uid: Message id from a search/browse hit.
        context_before: Neighboring messages before this one (default 0, max 20).
        context_after: Neighboring messages after this one (default 0, max 20).
    """
    con = _db()
    try:
        row, text = _message_text(con, uid)
        if row is None:
            return {"error": f"No message with uid {uid!r}."}
        out = {
            "uid": row["uid"],
            "source": row["src"],
            "source_label": _SOURCE_LABEL.get(row["src"], row["src"]),
            "who": row["who"],
            "date": row["date"],
            "title": row["title"],
            "context": row["ctx"],
            "path": row["rel"],
            "uri": _source_uri(row["root"], row["rel"]),
            "text": text,
        }
        before = max(0, min(context_before, 20))
        after = max(0, min(context_after, 20))
        if before or after:
            idx = row["msg_idx"]
            nb = con.execute(
                "SELECT DISTINCT uid FROM chunks WHERE root = ? AND rel = ? "
                "AND msg_idx BETWEEN ? AND ? AND uid != ? ORDER BY msg_idx",
                (row["root"], row["rel"], idx - before, idx + after, uid)).fetchall()
            ctx_b, ctx_a = [], []
            for (n_uid,) in nb:
                n_row, n_text = _message_text(con, n_uid)
                if n_row is None:
                    continue
                entry = {"uid": n_uid, "who": n_row["who"], "date": n_row["date"],
                         "text": n_text[:800]}
                (ctx_b if n_row["msg_idx"] < idx else ctx_a).append(entry)
            out["context_before"] = ctx_b
            out["context_after"] = ctx_a
        return out
    finally:
        con.close()


@mcp.tool(annotations=_READONLY)
def list_people(source: str = "all", contains: str = "", limit: int = 100) -> dict:
    """List the people in the corpus (senders / chat authors) with message counts.

    Use this to discover valid values for the `person` filter of search_messages
    and browse_messages.

    Args:
        source: One of "all", "teams", "outlook", "kalender", "kontakte".
        contains: Optional. Only people whose name or email contains this text.
        limit: Max number of people to return, most frequent first (default 100).
    """
    con = _db()
    try:
        conds = ["who != '' AND who != '(unbekannt)'"]
        params = []
        if source != "all":
            conds.append("src = ?")
            params.append(source)
        if contains.strip():
            # SQLite's LIKE/lower() are ASCII-only; register Python lower() so
            # umlaut-cased input ("MÜLLER") still matches. ppl is stored
            # pre-lowercased by the indexer, so only `who` needs folding.
            con.create_function("py_lower", 1,
                                lambda s: s.lower() if isinstance(s, str) else s,
                                deterministic=True)
            conds.append("(py_lower(who) LIKE ? OR ppl LIKE ?)")
            pat = f"%{contains.strip().lower()}%"
            params += [pat, pat]
        where = " AND ".join(conds)
        rows = con.execute(
            f"SELECT who, SUM(messages) AS m FROM people WHERE {where} "
            f"GROUP BY who ORDER BY m DESC, who LIMIT ?",
            [*params, max(1, limit)]).fetchall()
        total = con.execute(
            f"SELECT COUNT(DISTINCT who) FROM people WHERE {where}",
            params).fetchone()[0]
        return {"count": len(rows), "total_distinct": total,
                "people": [{"name": r[0], "messages": r[1]} for r in rows]}
    finally:
        con.close()


@mcp.tool(annotations=_READONLY)
def read_source_file(source_root: str, path: str, max_chars: int = 100000,
                     offset: int = 0) -> dict:
    """Read a raw exported source file (e.g. a whole .eml or Teams conversation).

    Large files (some Teams conversations exceed 100 MB) are returned in
    windows: the reply contains total_chars and truncated – pass offset to read
    the next window. Prefer get_document with context_before/context_after for
    chat conversations; it is far cheaper.

    Args:
        source_root: "teams" or "outlook" (the export the file belongs to).
        path: Relative path within that export, as returned in a hit's "path".
        max_chars: Max bytes to return (default 100000, cap 500000).
        offset: Byte position to start reading from (default 0).
    """
    target, err = _resolve_source(source_root, path)
    if err:
        return {"error": err}
    content, total, start, truncated = _read_window(target, offset, max_chars)
    return {"source_root": source_root, "path": path, "suffix": target.suffix,
            "total_bytes": total, "offset": start, "truncated": truncated,
            "content": content}


@mcp.tool(annotations=_READONLY)
def corpus_stats() -> dict:
    """Report corpus size, per-source counts, and which ranking backend is active."""
    con = _db()
    try:
        by_src = {r[0]: {"chunks": r[1], "messages": r[2]} for r in con.execute(
            "SELECT src, COUNT(*), COUNT(DISTINCT uid) FROM chunks GROUP BY src")}
        n = sum(v["chunks"] for v in by_src.values())
        return {
            "chunks": n,
            "by_source": by_src,
            "default_backend": "hybrid" if STATE.get("semantic") else "lexical",
            "semantic_available": bool(STATE.get("semantic")),
            "embed_model": STATE.get("embed_model") if STATE.get("semantic") else None,
            "vector_dtype": STATE.get("vector_dtype"),
            "last_semantic_error": STATE.get("last_semantic_error"),
            "teams_dir": STATE.get("teams_dir"),
            "outlook_dir": STATE.get("outlook_dir"),
        }
    finally:
        con.close()


# --------------------------------------------------------------------------
# MCP resources – fetch a source file by its URI (as advertised in each hit)
# --------------------------------------------------------------------------
@mcp.resource("o365://{root}/{path}")
def source_resource(root: str, path: str) -> str:
    """Return a raw exported source file by URI.

    URI form: o365://{root}/{path}, where {root} is "teams" or "outlook" and
    {path} is the export-relative file path, percent-encoded (slashes as %2F).
    This is the "uri" field returned with every search/browse hit. Files larger
    than 500k characters are truncated – use the read_source_file tool with
    offset to page through the rest.
    """
    target, err = _resolve_source(root, unquote(path))
    if err:
        raise ValueError(err)
    content, total, _, truncated = _read_window(target, 0, 500000)
    if truncated:
        content += (f"\n\n[truncated: {total} bytes total – use the "
                    f"read_source_file tool with offset to read more]")
    return content


# --------------------------------------------------------------------------
# Startup
# --------------------------------------------------------------------------
def _open_vectors(store, n_chunks):
    """Memory-map vectors.npy if numpy + the file are present."""
    vp = Path(store) / "vectors.npy"
    if not vp.exists():
        return None, None
    try:
        import numpy as np
    except ImportError:
        print("numpy not installed – semantic/hybrid ranking disabled.",
              file=sys.stderr)
        return None, None
    V = np.load(vp, mmap_mode="r")
    if V.shape[0] != n_chunks:
        print(f"Index/DB mismatch ({V.shape[0]} vectors vs {n_chunks} chunks) – "
              f"rebuild with rag_index.py. Lexical ranking only.", file=sys.stderr)
        return None, None
    return np, V


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--store", default="rag_store")
    ap.add_argument("--teams", default="teams_export")
    ap.add_argument("--outlook", default="outlook_export")
    ap.add_argument("--embed-model", default="bge-m3")
    ap.add_argument("--ollama", default="http://localhost:11434")
    ap.add_argument("--transport", choices=["http", "stdio"], default="http",
                    help="http: one shared server, register its URL in Claude "
                         "(default). stdio: launched per client via command.")
    ap.add_argument("--host", default="127.0.0.1",
                    help="HTTP bind address. Keep 127.0.0.1 – the server has no "
                         "auth and serves your mail/chat history.")
    ap.add_argument("--port", type=int, default=8365)
    a = ap.parse_args()

    dbp = Path(a.store) / "corpus.db"
    if not dbp.exists():
        raise SystemExit(f"No store at '{dbp}'. Build it first:\n"
                         f"  python3 rag_index.py {a.teams} {a.outlook}")

    con = sqlite3.connect(f"file:{dbp}?mode=ro", uri=True)
    n_chunks = con.execute("SELECT COUNT(*) FROM chunks").fetchone()[0]
    con.close()

    np, V = _open_vectors(a.store, n_chunks)
    STATE.update(db=str(dbp), V=V, np=np, semantic=(np is not None),
                 vector_dtype=str(V.dtype) if V is not None else None,
                 teams_dir=a.teams, outlook_dir=a.outlook,
                 embed_model=a.embed_model, ollama=a.ollama)

    backend = ("hybrid (BM25 + semantic, RRF)" if np is not None
               else "lexical (FTS5/BM25) only")
    print(f"office365-export MCP: {n_chunks} chunks · {backend}", file=sys.stderr)
    if a.transport == "http":
        mcp.settings.host = a.host
        mcp.settings.port = a.port
        print(f"MCP endpoint: http://{a.host}:{a.port}{mcp.settings.streamable_http_path}",
              file=sys.stderr)
        mcp.run(transport="streamable-http")
    else:
        mcp.run()


if __name__ == "__main__":
    main()
