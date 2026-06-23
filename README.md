# Office 365 Export

Tools to export Microsoft 365 data (Teams chats/channels, Outlook mail, calendar
and contacts) via Microsoft Graph — delegated access, no admin consent required —
and to search the exports offline.

All scripts run on **macOS, Windows and Linux** with Python 3.7 or newer.

---

## 1. Prerequisites

- **Python 3.7+** — check with `python3 --version` (macOS/Linux) or `python --version` (Windows).
  - Windows: install from [python.org](https://www.python.org/downloads/) and tick
    *"Add python.exe to PATH"* during setup.
- The **export** tools need two packages: `msal` and `requests`.
- The **search** tools (`teams_search.py`, `outlook_search.py`, `combined_search.py`)
  use only the Python standard library — no installation needed.

---

## 2. Install the requirements

Use a virtual environment so nothing is installed system-wide.

### macOS / Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install msal requests
```

### Windows (PowerShell)

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install msal requests
```

> If PowerShell blocks the activation script, run once:
> `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`

### Windows (Command Prompt / cmd.exe)

```bat
python -m venv .venv
.venv\Scripts\activate.bat
python -m pip install msal requests
```

In every following command, use `python3` on macOS/Linux and `python` on Windows.
The example commands below show macOS/Linux first, then the Windows equivalent.

---

## 3. Authentication

The export scripts sign you in interactively (a browser window opens). If your
tenant requires admin approval for new apps, use a pasted token instead:

1. Open the [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer)
2. Log in
3. Open the **"Access token"** tab and copy the token
4. Save it as `gx_token.txt` next to the scripts, **or** set an environment variable:

   - macOS/Linux: `export GRAPH_TOKEN="eyJ0…"`
   - Windows (PowerShell): `$env:GRAPH_TOKEN = "eyJ0…"`
   - Windows (cmd): `set GRAPH_TOKEN=eyJ0…`

The token mode needs the right scopes consented in Graph Explorer
(`Chat.Read`/`ChannelMessage.Read.All` for Teams; `Mail.Read`, and additionally
`Calendars.Read` / `Contacts.Read` for Outlook calendar/contacts).

---

## 4. Export tools

Each export tool is interactive (it asks what to export) and resumable — re-running
it only fetches new/changed items. An optional argument sets the output folder.

### Teams — chats and channels → HTML

macOS/Linux:
```bash
python3 teams_export.py            # default folder: teams_export
```
Windows:
```powershell
python teams_export.py
```

### Outlook — mail (.eml), calendar (.ics), contacts (.vcf)

macOS/Linux:
```bash
python3 outlook_export.py          # default folder: outlook_export
```
Windows:
```powershell
python outlook_export.py
```

### Export options

The exports are interactive (you pick what to export) and resumable. The most
useful options are environment variables — set them before running, no code edits
needed:

| Variable | Applies to | Default | What it does |
|---|---|---|---|
| `EXPORT_WORKERS` | both | `4` | Parallel conversations/downloads. `4` is the sensible maximum — Exchange/Teams throttle harder above it. Lower it (e.g. `1`) on flaky connections. |
| `GRAPH_TOKEN` | both | — | Use a pasted Graph token instead of the browser login (see *Authentication*). |
| `REFRESH_CHANNELS` | Teams | `1` | `0` exports Teams channels only once; otherwise channels are re-checked for new replies on every run. |
| `CACHE_IMAGES` | Teams | `1` | `0` disables caching of inline images (saves disk, re-downloads on every re-export). |
| `SKIP_EMPTY_CHATS` | Teams | `1` | `0` also exports chats that contain only system messages (joins, calls, …). |

Set a variable like this (example: 2 workers):

- macOS/Linux: `EXPORT_WORKERS=2 python3 teams_export.py`
- Windows (PowerShell): `$env:EXPORT_WORKERS=2; python teams_export.py`
- Windows (cmd): `set EXPORT_WORKERS=2 && python teams_export.py`

**Output folder** — pass it as the first argument (re-running into the same folder
resumes; deleting the folder forces a full re-export):

```
python3 teams_export.py   my_teams_archive
python3 outlook_export.py my_outlook_archive
```

A few options are switches near the top of each script (edit the file to change them):

- `USE_DEVICE_CODE = True` — log in with a device code instead of opening a browser
  (useful on headless machines).
- Teams `EMBED_IMAGES = False` — don't embed inline images as base64 (smaller HTML).
- Outlook `INCLUDE_HIDDEN = True` — also export hidden system folders
  (Conversation History, Sync Issues, …).
- Outlook `DEFAULT_SKIP_FOLDERS` — folders excluded when you just press Enter
  (Archive, Drafts, Deleted Items, Junk, Outbox); you can still pick them explicitly.

---

## 5. Search tools (offline, no install)

These read an export folder once and write a self-contained `search.html` you open
in a browser. They use only the standard library.

### Teams search

macOS/Linux:
```bash
python3 teams_search.py            # default: teams_export → teams_export/search.html
```
Windows:
```powershell
python teams_search.py
```

### Outlook search

macOS/Linux:
```bash
python3 outlook_search.py          # default: outlook_export → outlook_export/search.html
```
Windows:
```powershell
python outlook_search.py
```

### Combined search (Teams + Outlook in one page)

macOS/Linux:
```bash
python3 combined_search.py [teams-folder] [outlook-folder] [-o output.html]
```
Windows:
```powershell
python combined_search.py [teams-folder] [outlook-folder] [-o output.html]
```

Defaults are `teams_export` and `outlook_export`. The page is written to the common
parent folder of both exports so the relative links work — don't move it relative to
the export folders afterwards.

---

## 6. RAG Search (optional, AI answers)

Requires [Ollama](https://ollama.com) running locally.

```bash
ollama pull bge-m3                    # embeddings, multilingual (DE/EN)
ollama pull qwen2.5:14b-instruct      # answer model, fits well in 24 GB
```

Install the extra dependencies (in your virtual environment):

macOS/Linux:
```bash
python3 -m pip install numpy requests
```
Windows:
```powershell
python -m pip install numpy requests
```

Then build the index and start the service:

macOS/Linux:
```bash
python3 rag_index.py teams_export outlook_export          # 1) build index (incremental)
python3 rag_server.py --teams teams_export --outlook outlook_export   # 2) serve
```
Windows:
```powershell
python rag_index.py teams_export outlook_export
python rag_server.py --teams teams_export --outlook outlook_export
```

Then open <http://localhost:8000>.
