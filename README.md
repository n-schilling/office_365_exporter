# Office 365 Export

Export your Microsoft 365 data (Teams chats/channels, Outlook mail, calendar,
contacts) via Microsoft Graph — delegated access, no admin consent required —
and search the exports offline: as a static HTML page, through **Claude** (MCP),
or with a local RAG web UI.

```
teams_export.py  ─┐                        ┌─ *_search.py      → static search.html
                  ├─ local export folders ─┤
outlook_export.py ┘                        └─ rag_index.py → rag_store/
                                                ├─ mcp_server.py → Claude (MCP tools)
                                                └─ rag_server.py → RAG web UI (Ollama)
```

| Script | Purpose |
|---|---|
| `teams_export.py` | Teams 1:1/group/meeting chats and channels → HTML |
| `outlook_export.py` | Mail (`.eml`), calendar (`.ics`), contacts (`.vcf`) |
| `teams_search.py` / `outlook_search.py` / `combined_search.py` | Self-contained offline `search.html` |
| `rag_index.py` | Builds the search index (`rag_store/`: SQLite + FTS5 + embeddings) |
| `mcp_server.py` | MCP server — Claude searches and reads the exports itself |
| `rag_server.py` | Local RAG web UI with AI answers (fully offline via Ollama) |
| `corpus.py` | Shared export parsing (used internally) |

Everything runs on macOS, Windows and Linux with Python 3.7+ (MCP server: 3.10+).
Commands below use `python3`; on Windows type `python` instead.

---

## 1. Setup

Create a virtual environment and install what you need:

```bash
python3 -m venv .venv
source .venv/bin/activate              # Windows PowerShell: .\.venv\Scripts\Activate.ps1
python3 -m pip install msal requests   # export tools
python3 -m pip install mcp numpy       # only for the MCP server / AI search
```

The static search tools need no packages at all.

> PowerShell blocks the activation script? Run once:
> `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`

<details>
<summary>Windows without install rights: standalone (embeddable) Python</summary>

1. Download the **"Windows embeddable package (64-bit)"** from
   [python.org/downloads/windows](https://www.python.org/downloads/windows/) and
   unzip it, e.g. to `C:\python-standalone`.
2. In that folder, open `python3XX._pth` in a text editor and remove the `#`
   before `import site`.
3. Bootstrap pip and install the packages (PowerShell, in that folder):

   ```powershell
   Invoke-WebRequest https://bootstrap.pypa.io/get-pip.py -OutFile get-pip.py
   .\python.exe get-pip.py
   .\python.exe -m pip install msal requests
   ```

Then run every command with the full path, e.g.
`C:\python-standalone\python.exe teams_export.py`. No virtual environment
needed — packages install into the standalone folder itself.
</details>

---

## 2. Authentication

The export scripts sign you in interactively (a browser window opens). If your
tenant requires admin approval for new apps, paste a token instead:

1. Log in at the [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
   and copy the token from the **"Access token"** tab.
2. Save it as `gx_token.txt` next to the scripts, or set it as the `GRAPH_TOKEN`
   environment variable.

Token mode needs the right scopes consented in Graph Explorer:
`Chat.Read` / `ChannelMessage.Read.All` (Teams), `Mail.Read`, plus
`Calendars.Read` / `Contacts.Read` for calendar/contacts.

---

## 3. Export

```bash
python3 teams_export.py                # → teams_export/
python3 outlook_export.py              # → outlook_export/
```

Both tools are interactive (they ask what to export) and **resumable** —
re-running only fetches new/changed items. Deleting the folder forces a full
re-export.

**Common variations:**

```bash
python3 teams_export.py my_archive     # custom output folder
python3 outlook_export.py -default     # no questions, default selection
                                       #   (ideal for cron/scheduled runs)
```

`-default` exports: Teams — 1:1, group and meeting chats (no channels);
Outlook — mail (except Archive, Drafts, Deleted Items, Junk, Outbox), the
default calendar and all contacts. Same as pressing Enter at every prompt.

**Options** (environment variables, e.g. `EXPORT_WORKERS=2 python3 teams_export.py`):

| Variable | Applies to | Default | What it does |
|---|---|---|---|
| `EXPORT_WORKERS` | both | `4` | Parallel downloads. `4` is the sensible max (throttling); use `1` on flaky connections. |
| `GRAPH_TOKEN` | both | — | Pasted Graph token instead of browser login (section 2). |
| `REFRESH_CHANNELS` | Teams | `1` | `0` = don't re-check exported channels for new replies. |
| `CACHE_IMAGES` | Teams | `1` | `0` = don't cache inline images (saves disk, slower re-export). |
| `SKIP_EMPTY_CHATS` | Teams | `1` | `0` = also export chats with only system messages. |

<details>
<summary>Switches at the top of each script (edit the file)</summary>

- `USE_DEVICE_CODE = True` — device-code login instead of a browser (headless machines).
- Teams `EMBED_IMAGES = False` — don't embed images as base64 (smaller HTML).
- Outlook `INCLUDE_HIDDEN = True` — also export hidden system folders.
- Outlook `DEFAULT_SKIP_FOLDERS` — folders skipped by the default selection.
</details>

---

## 4. Static search pages (offline, no install)

Each tool reads an export folder and writes a self-contained `search.html`:

```bash
python3 teams_search.py                # → teams_export/search.html
python3 outlook_search.py              # → outlook_export/search.html
python3 combined_search.py             # → both in one page
```

`combined_search.py` accepts custom folders (`[teams] [outlook] [-o out.html]`)
and writes the page to the common parent folder of both exports — don't move it
afterwards, the links are relative.

---

## 5. Search index (needed for MCP and RAG UI)

The index lives in `rag_store/`: `corpus.db` (SQLite with an FTS5 full-text
index) and `vectors.npy` (float16 embeddings, built with
[Ollama](https://ollama.com)).

```bash
ollama pull bge-m3                     # embedding model, multilingual (DE/EN)
python3 rag_index.py teams_export outlook_export
```

The build is **incremental** — re-run it after each export; only new/changed
content is re-embedded.

---

## 6. MCP server — search with Claude

`mcp_server.py` exposes the exports to Claude (Claude Code / Claude Desktop) as
[MCP](https://modelcontextprotocol.io) tools — Claude searches, reads sources
and answers with citations; no local answer model needed.

**Ranking** is hybrid by default: FTS5/BM25 and semantic cosine search merged
with Reciprocal Rank Fusion — exact tokens (invoice numbers, names) and
paraphrases both hit. If Ollama is down (or `numpy` is missing) it falls back
to pure BM25 automatically.

**Run it** (leave it running; one instance serves all Claude sessions):

```bash
python3 mcp_server.py                  # endpoint: http://127.0.0.1:8365/mcp
```

> ⚠️ The server has **no authentication** and serves your complete mail and
> chat history. Keep it on `127.0.0.1` (the default) — never bind it to the
> network with `--host`.

**Register in Claude Code** — this repo's `.mcp.json` already does it:

```json
{"mcpServers": {"office365-export": {"type": "http", "url": "http://127.0.0.1:8365/mcp"}}}
```

**Register in Claude Desktop** — `claude_desktop_config.json` only accepts
`command` entries, so bridge the HTTP endpoint with
[mcp-proxy](https://github.com/sparfenyuk/mcp-proxy):

```json
{"mcpServers": {"office365-export": {
  "command": "uvx",
  "args": ["mcp-proxy", "--transport", "streamablehttp", "http://127.0.0.1:8365/mcp"]
}}}
```

Prefer the classic auto-launched setup instead of a shared server? Register a
`command`-based entry running `python3 mcp_server.py --transport stdio`.

**Tools:** `search_messages` (hybrid search; person/date/source filters,
pagination), `browse_messages` (filtered listing, newest first), `get_document`
(full message, optionally with neighboring chat messages), `list_people`
(who is in the corpus — valid `person` filter values), `read_source_file`
(raw `.eml`/conversation, windowed for large files), `corpus_stats`. Every hit
carries an `o365://` resource URI through which Claude can fetch the source
file. All tools are read-only.

Then just ask Claude: *"Search my Teams chats with Anna about the Q3 budget."*

---

## 7. RAG web UI — fully offline AI answers

The self-contained alternative to the MCP server: retrieval plus a local answer
model, no Claude involved.

```bash
ollama pull qwen2.5:14b-instruct       # answer model, fits well in 24 GB
python3 rag_server.py --teams teams_export --outlook outlook_export
```

Then open <http://localhost:8000> — semantic search with filters, or full
question answering with source citations.
