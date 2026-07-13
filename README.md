# AI Assistant

A local, non-admin AI Document Assistant: an Office Add-in task pane for
Microsoft Word that chats with a real [OpenCode](https://opencode.ai) agent
and lets it read/edit the currently open Word document through
[word-mcp-live](https://github.com/ykarapazar/word-mcp-live) (a fixed,
auditable set of Windows COM automation tools, not arbitrary code execution).

See [PRD.md](PRD.md) for the original requirements and the architecture
decisions that diverge from it (real OpenCode REST API instead of a custom
`/v1/chat` proxy; `word-mcp-live` as a local MCP server instead of a
hand-rolled "run arbitrary script" tool).

## Architecture

```
Office App (Word / PowerPoint / Excel / Outlook)
        │
        ├──► Task Pane / Mail App UI (Office.js, https://localhost:3000)
        │       │ fetch()
        │       ▼
        └──► opencode serve --port 4098 --cors https://localhost:3000
                │ stdio (local MCP servers, declared in opencode.json)
                ├──► word-mcp-live (uv) ── win32com ──► Word.Application (open document)
                ├──► ppt-mcp (uvx) ─────── win32com ──► PowerPoint.Application (open presentation)
                └──► excel (uv run) ────── JXA/COM ───► Excel.Application (open workbook)
```

The task pane talks directly to OpenCode's real REST API
(`POST /session`, `POST /session/{id}/message` — see `/doc` on the running
server for the full OpenAPI spec), not a custom backend.

## Prerequisites

| Tool | Why | Check |
|---|---|---|
| Node.js 18+ | builds/serves the task pane | `node --version` |
| Python 3.11+ | required by automation tooling | `python --version` |
| [`uv`/`uvx`](https://docs.astral.sh/uv/) | runs the MCP servers in isolated envs | `uvx --version` |
| [`opencode` CLI](https://opencode.ai) | the agent runtime (`opencode-ai` on npm) | `opencode --version` |
| Microsoft Word (desktop) | the Word automation target | — |
| Microsoft PowerPoint (desktop) | the PowerPoint automation target | — |
| Microsoft Excel (desktop) | the Excel automation target | — |
| Microsoft Outlook (desktop or web) | the Outlook email compose/read target | — |

You also need an OpenCode LLM provider configured (`opencode auth login`, or
edit `~/.config/opencode/opencode.jsonc`) — `opencode providers` should list
at least one usable model. `src/taskpane/taskpane.js` hardcodes which
provider/model the task pane sends (`OPENCODE_MODEL` near the top of the
file); update it to match whatever provider you've configured.

## One-time setup

```powershell
npm install
npx office-addin-dev-certs install
```

`office-addin-dev-certs install` triggers a one-time Windows "Security
Warning" dialog asking whether to trust the local dev HTTPS certificate.
**This click cannot be automated** — it must be approved manually, once, on
this machine. After that, `https://localhost:3000` is trusted for all future
runs.

## Autostart at Windows logon (recommended)

Sideloading the add-in (`npm start`) registers it **permanently** in
`HKCU:\Software\Microsoft\Office\16.0\WEF\Developer` — Word will try to
reconnect to `https://localhost:3000` on every future launch, not just the
session where you ran `npm start`. If the two background services aren't
already running when Word starts, this reconnect attempt fails with an
"ADD-IN ERROR" dialog.

To fix this once and for all, install a silent, per-user, admin-free autostart
entry that brings the services up automatically at every Windows logon:

```powershell
.\scripts\install-autostart.ps1
```

`install-autostart.ps1` first runs `npm run build` to produce the taskpane
bundle in `dist/`, then writes a `.vbs` wrapper to your Startup folder
(`%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup`) that silently runs
`scripts\start-background-services.ps1` at logon. That script is idempotent —
it checks whether ports 3000/4098 are already listening before starting
anything, so it's safe even if you also start things manually. Logs go to
`%LOCALAPPDATA%\AI Assistant\logs\`.

On port 3000 the autostart serves the **pre-built** bundle from `dist/` with
a tiny static HTTPS server (`scripts/serve-static.js`, `npm run serve-static`)
rather than running `webpack serve`. Compiling the bundle on every cold boot
used to leave port 3000 unserved for a minute or two after logon — long
enough that Word's reconnect failed with "ADD-IN ERROR" — so the bundle is
built ahead of time (at install) and just served as static files, which is
ready almost immediately. The static server reuses the very same trusted
`office-addin-dev-certs` certificate and serves the same files at the same
URLs as the webpack dev server, so nothing about how Word loads the add-in
changes. **If you edit anything under `src/`, re-run `npm run build` (or
`.\scripts\install-autostart.ps1`) so the autostart picks up the change** —
`npm run dev-server` (webpack, live recompile) remains the tool to use while
actively developing.

The script also self-heals the `word` and `excel` MCP connections: a logon-time cold start
can make opencode's handshake with the MCP servers time out, leaving them
in a permanent `failed` state (the agent then has no Word or Excel tools). On
every run, the script polls the live `/mcp` status and, if either `word` or `excel`
shows `failed`, first tries a runtime reconnect (`POST /mcp/<service>/connect`)
and then, as a last resort, restarts `opencode serve` once.

To start the services immediately without logging out again, run the same
script directly:

```powershell
.\scripts\start-background-services.ps1
```

To remove the autostart entry: `.\scripts\uninstall-autostart.ps1`.

This is a one-time, manual step by design — it changes your Windows user
profile's logon behavior, so it's left for you to run explicitly rather than
happening silently as a side effect of any other command.

Once the add-in is sideloaded (`npm start`, at least once) and the taskpane
has been opened once in a given document, that document is tagged to
auto-show the taskpane on reopen (via `Office.AutoShowTaskpaneWithDocument`) —
you won't need to click the ribbon button again for that specific file. This
does **not** apply to brand-new blank documents; the first taskpane open in a
new document still requires clicking the ribbon button once.

## Daily run sequence

If you haven't installed autostart above, you'll need to start both services
by hand. Two terminals, both from the project root:

```powershell
# Terminal 1 - OpenCode server. --cors is required: the task pane runs on
# https://localhost:3000 and OpenCode serves plain http://127.0.0.1:4098 -
# without --cors matching the task pane's origin, the browser blocks the
# fetch with a CORS policy error (this is NOT a mixed-content issue -
# Chromium's loopback exception already allows https -> http on localhost).
opencode serve --port 4098 --hostname 127.0.0.1 --cors https://localhost:3000

# Terminal 2 - sideloads the add-in and opens Word
npm start
```

Before asking the assistant to edit a document or presentation, make sure the target application (Word or PowerPoint) is actually open — the MCP servers attach to active instances via COM `GetActiveObject`; they do not launch the applications themselves.

To stop: `npm stop` (unloads the add-in), then Ctrl+C the `opencode serve`
terminal.

## Word, PowerPoint & Excel Automation (`opencode.json`)

All three automation targets are configured as local MCP servers inside `opencode.json`:

```json
{
  "mcp": {
    "word": {
      "type": "local",
      "command": ["uv", "run", "--with", "word-mcp-live", "scripts/word_mcp_launcher.py"],
      "enabled": true,
      "timeout": 90000
    },
    "powerpoint": {
      "type": "local",
      "command": ["uvx", "ppt-mcp"],
      "enabled": true,
      "timeout": 90000
    },
    "excel": {
      "type": "local",
      "command": ["uv", "run", "scripts/excel_mcp_server.py"],
      "enabled": true,
      "timeout": 90000
    }
  }
}
```

- **Word**: Uses `word-mcp-live` via a custom launcher script (`scripts/word_mcp_launcher.py`) that resolves the correct active document and applies robust bindings to avoid multi-document mismatch risks.
- **PowerPoint**: Uses `ppt-mcp` spawned via `uvx ppt-mcp`. This server exposes tools to manipulate slides, add shapes, resize and format elements, and extract text.
- **Excel**: Uses the `excel` MCP server written using FastMCP and pandas in [scripts/excel_mcp_server.py](file:///Users/kwongyiu/Development/opencode-office-plugin/scripts/excel_mcp_server.py). It exposes:
  - `read_active_sheet`: Reads range values, formulas, and worksheet info. Detects and highlights formula errors (e.g. `#REF!`, `#DIV/0!`).
  - `write_cell`: Writes values or Excel formulas (starting with `=`) to coordinate cells (e.g. `C4`).
  - `analyze_excel_file`: Uses pandas to inspect local CSV or Excel files, returning shape information and column types.
  - `run_pandas_query`: Evaluates pandas queries/aggregations on `df` for local CSV or Excel files, formatting output as GFM Markdown tables.

Verify that the tools are wired up correctly and showing as connected in the MCP registry:

```powershell
opencode mcp list
```

## Harness mode (auto-detected Solution Architect vault)

If a Solution Architect **harness vault** — any folder containing
`_agentic/os/AGENTS.md` (alongside `SOUL.md`, `USER.md`, and a `memory/`
folder) — sits next to or above this plugin's folder, the static server
auto-detects it (`scripts/serve-static.js` scans the plugin's ancestors and
immediate siblings) and the task pane switches into **harness mode**:

- The Settings tab **hides the manual System Instruction / Persona fields**
  and shows a banner naming the vault. Identity, user context, and conventions
  come from the vault instead: each new conversation's first message carries a
  scoped instruction telling the agent to read the vault's `SOUL.md` /
  `USER.md` / `AGENTS.md` (and today's/yesterday's memory + `MEMORY.md` if
  present) and follow them. The agent is explicitly told **not** to run the
  harness's maintenance scripts, memory-graph, or graphify — this is a
  lightweight "read the identity" integration, not full harness participation
  (opencode still runs from the plugin folder, not the vault).
- A **💾 Save to memory** button appears in the toolbar next to **💡 Prompt
  ideas**. Click it to have the agent append this conversation's key points to
  `<vault>\_agentic\os\memory\<today>.md` and promote durable items to
  `MEMORY.md`, following the vault's own memory rules (it skips if the
  memory-graph lock is held, and never overwrites). Note this writes into the
  vault, so its git working tree will show the change.

Detection is best-effort and only happens when the task pane is served by
`serve-static` (the installed/autostart path). Served any other way (webpack
dev-server, nginx), or with no vault nearby, the task pane stays in the
generic manual mode below. Because `opencode.json` grants blanket
`external_directory` access, the agent can read/write the whole vault — see
the note under the gear-icon section and `tech-design-spec.md` §4.

## Generic mode: System Instruction, Persona & Prompt library (gear icon)

The task pane's header has a gear icon (⚙) that opens a panel with two
tabs. The **Settings** tab has two fields:

- **System Instruction** — standing rules the assistant must follow in
  every conversation (e.g. "Always reply in formal English. Focus on
  tender/proposal quality.").
- **Persona** — who the assistant should be (e.g. "You are a senior bid
  manager at NCS with 15 years of government-tender experience.").

Click **Save** and the settings **take effect live**: from your next
message onward, each new conversation's first message carries a hidden
instruction with these two fields, so the agent behaves accordingly from
the start. No server restart is needed; changing and re-saving applies
from the very next message. **Clear** wipes both fields — new
conversations go back to default behavior. The fields are stored in the
browser's `localStorage` (per machine/user, not per document).

The **Prompt library** tab maintains the prompt templates that the 💡
"Prompt ideas" button picks from: the built-in templates are seeded on
first load and are fully editable — change any template's text, delete
ones you never use, or **+ Add** your own (placeholders like `[Topic]` are
fine; the assistant fills them in with document specifics before showing a
suggestion). Each row also has a ⤵ **Use it** button (below the ✕ remove
button) that stages that prompt straight into the chat input — edit it if
needed, then send. **Save library** applies from the next "Prompt ideas"
click — no restart, no new conversation needed. **Reset to defaults**
restores the built-in set. Note the two "planning mode" templates (winning theme /
business value) carry their plan-mode behavior by template *identity*, so
editing their wording keeps that behavior, while deleting them removes it.

Advanced: because this project's `opencode.json` sets
`permission.external_directory: "allow"`, a System Instruction can also
tell the agent to read local folders — e.g. "Read the agent-harness
directory at C:\\path\\to\\vault (AGENTS.md, SOUL.md, memory/) and follow
it." (If you had configured a harness root directory in the panel's
earlier design, it is migrated into exactly such a System Instruction
automatically, once.) Note that `external_directory: "allow"` grants the
agent's `read`/`write`/`edit`/`bash` tools **full read/write access to any
local path** — an acceptable posture for a local, single-user tool where
the agent only acts on your own prompts, but worth knowing. (A per-path
grant can't be applied at runtime because opencode's `PATCH /config`
doesn't persist anything — see `tech-design-spec.md` §5.)

## Sideloading, Targeting & Debugging Office Hosts

### Dual-Manifest Design
The project uses a dual-manifest system:
- `manifest.xml`: Configures the TaskPaneApp for **Word** (Document), **PowerPoint** (Presentation), and **Excel** (Workbook).
- `manifest-outlook.xml`: Configures the MailApp for **Outlook** (Mailbox).

### Targeting Word, PowerPoint, or Excel (manifest.xml)
To choose which Office host application to launch and debug with the main manifest:
1. Open `package.json` and locate the `config` block.
2. Edit `app_to_debug` to `"word"`, `"powerpoint"`, or `"excel"`:
   ```json
     "config": {
       "app_to_debug": "excel", // "word" | "powerpoint" | "excel"
       "app_type_to_debug": "desktop",
       "dev_server_port": 3000
     }
   ```
3. Run `npm start` to register the manifest and start the selected host application on Windows.

#### Sideloading on macOS
For macOS Word, Excel, or PowerPoint desktop clients, sideloading is done by copying the manifest into the host's Web Extension Framework (wef) folder:
- **Excel**: Copy `manifest.xml` to `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`
- **Word**: Copy `manifest.xml` to `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`
- **PowerPoint**: Copy `manifest.xml` to `~/Library/Containers/com.microsoft.PowerPoint/Data/Documents/wef/`
*(Create the `wef` directory if it does not exist. Reload/Insert My Add-ins in the host ribbon).*

### Sideloading Outlook (manifest-outlook.xml)
Sideloading Outlook (on Windows, macOS, or Outlook Web at outlook.office.com) is done manually via the client UI:
1. In Outlook, click **All Apps** or **More Apps** on the left navigation bar or ribbon, then click **Add apps** or **Get Add-ins**.
2. Select **Manage your add-ins** or **My Add-ins** -> **Custom Add-ins**.
3. Select **Add a custom add-in** -> **Add from file...**.
4. Browse to and upload `manifest-outlook.xml` (or `dist/manifest-outlook.xml`).
5. The "AI Assistant" ribbon button will appear in message Read surfaces and Compose windows.

### Right-Click "Analyze with AI Assistant" Context Menu
The Word, Excel, and PowerPoint manifests configure a contextual right-click menu command:
- **Access**: Highlight text in Word or PowerPoint, or right-click any cell in Excel, and choose **Analyze with AI Assistant**.
- **Mechanism**: The menu command executes `analyzeSelectionAction` in [src/commands/commands.js](file:///Users/kwongyiu/Development/opencode-office-plugin/src/commands/commands.js), which retrieves the selected content via Office.js `getSelectedDataAsync()`, serializes it as `contextMenuTrigger` in `localStorage`, and opens the taskpane. The taskpane in [src/taskpane/taskpane.js](file:///Users/kwongyiu/Development/opencode-office-plugin/src/taskpane/taskpane.js) intercepts this trigger, pre-fills the input with `Analyze this selection: "<text>"`, and automatically submits the prompt.

### Outlook Compose Mode Insertion
When composing emails in Outlook, the chat response bubbles in the taskpane show a **Insert into email** (✍) action button. Clicking this button calls `Office.context.mailbox.item.body.setSelectedDataAsync` to write the generated text directly into the email body at the cursor position.

To verify or test automation capabilities:
1. Ensure the active document, sheet, or presentation is open.
2. Start the background services using `.\scripts\start-background-services.ps1` (or `scripts/start-background-services.sh` on macOS).
3. Interact with the chat interface. Confirm edits take effect on the active document or sheet.

## Testing

Each phase of this project has its own regression test under `tests/`.
Run everything (fails fast on the first regression):

```powershell
.\tests\run-all.ps1
```

Or run a single phase, e.g. `.\tests\phase3-opencode-server.ps1`.

| Script | Verifies |
|---|---|
| `phase1-scaffold.ps1` | manifest.xml / package.json / project structure |
| `phase2-devserver.ps1` | manifest validates, dev cert installed, `https://localhost:3000/taskpane.html` serves |
| `phase3-opencode-server.ps1` | `opencode serve` responds on `/global/health` and `/doc` |
| `phase4-word-mcp.ps1` | `opencode.json` config is correct and `opencode mcp list` shows `word` connected |
| `phase5-chat-ui.ps1` (+ `phase5-chat-ui.mjs`) | full chat round trip through the real task pane UI, driven headlessly via `playwright-core` against the system's installed Edge |
| `phase6-word-mutation.ps1` | `opencode run` actually mutates a live Word document (adds a heading + table) through `word-mcp-live`, verified via COM afterward |
| `phase7-readme.ps1` | README.md documents the operational details from earlier phases |
| `phase8-autostart.ps1` | manifest/taskpane autoopen wiring, plus a real functional round trip of `start-background-services.ps1` and `install-autostart.ps1`/`uninstall-autostart.ps1` |
| `phase9-settings-panel.ps1` (+ `phase9-settings-panel.mjs`) | the gear-icon panel: Settings tab (saves/re-fills System Instruction & Persona across reloads, legacy harness-root migration, Clear) and Prompt library tab (default seeding, edit/delete/add, persistence, reset to defaults) — driven headlessly via `playwright-core`, no `opencode serve` needed since the panel is pure client-side |
| `phase10-harness.ps1` (+ `phase10-harness.mjs`) | harness auto-detection (`serve-static` `/harness-info` against a temp `_agentic/os/AGENTS.md` layout) and harness-mode UI (hidden manual Settings fields, banner, "Save to memory" pill in the toolbar posting a memory-write turn) — headless via `playwright-core`, opencode stubbed so no real vault or server is needed |

Phase 5, 6, 8, and 9 spin up real background processes (dev server, `opencode
serve`) and, for Phase 6, a real visible Word window — they are integration
tests against the live stack, not mocks.

## Troubleshooting

**Certificate warning on first run.** Expected once; see "One-time setup"
above. If `npm start` or `npm run dev-server` reports a certificate error,
re-run `npx office-addin-dev-certs install`.

**Chat shows "Could not reach OpenCode..."** Usually means `opencode serve`
isn't running, is on the wrong port, or was started without `--cors
https://localhost:3000`. Open the task pane's dev tools (right-click the
task pane → Inspect) and check the Console/Network tab for the exact error.

**`opencode mcp list` doesn't show `word` as connected.** Confirm `uv`/`uvx`
is on PATH and `opencode.json` is present in the directory `opencode serve`
/ `opencode run` was launched from — OpenCode discovers project config
relative to its working directory.

**A Word or PowerPoint edit request errors or times out.** Confirm the Word document or PowerPoint presentation is actually open (`word-mcp-live` or `ppt-mcp` needs `GetActiveObject` to find the active instance). Also avoid running `opencode run ...` against a **locked Windows session** — a locked/inactive desktop can make Office COM automation hang indefinitely with no error, since there's no interactive desktop to attach to; this was observed firsthand while building Phase 6's test.

**`opencode run` in a script never returns.** If you've wrapped a Word
automation call in a non-interactive script, give it a hard timeout (see
`tests/phase6-word-mutation.ps1` for the `Start-Job` + `Wait-Job -Timeout`
pattern) rather than assuming it will always return promptly.

**"ADD-IN ERROR" when opening Word, even without running `npm start` first.**
This means the add-in is still sideloaded from a previous `npm start` — check
with `Get-ItemProperty HKCU:\Software\Microsoft\Office\16.0\WEF\Developer`
(the manifest GUID will be listed if so; `npm stop` does not reliably clear
this). Either run `.\scripts\install-autostart.ps1` so the background
services are always up before Word needs them (see "Autostart at Windows
logon" above), or check
`%LOCALAPPDATA%\AI Assistant\logs\opencode-serve.err.log` and
`dev-server.err.log` for why they failed to start.
