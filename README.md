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
Word task pane (Office.js, https://localhost:3000)
        │ fetch()
        ▼
opencode serve --port 4098 --cors https://localhost:3000
        │ stdio (local MCP server, declared in opencode.json)
        ▼
word-mcp-live (uvx) ── win32com ──► Word.Application (already-open document)
```

The task pane talks directly to OpenCode's real REST API
(`POST /session`, `POST /session/{id}/message` — see `/doc` on the running
server for the full OpenAPI spec), not a custom backend.

## Prerequisites

| Tool | Why | Check |
|---|---|---|
| Node.js 18+ | builds/serves the task pane | `node --version` |
| Python 3.11+ | required by `word-mcp-live` | `python --version` |
| [`uv`/`uvx`](https://docs.astral.sh/uv/) | runs `word-mcp-live` in an isolated env | `uvx --version` |
| [`opencode` CLI](https://opencode.ai) | the agent runtime (`opencode-ai` on npm) | `opencode --version` |
| Microsoft Word (desktop, Windows) | the automation target | — |

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

The script also self-heals the `word` MCP connection: a logon-time cold start
can make opencode's handshake with `word-mcp-live` time out, leaving the
connection in a permanent `failed` state (the agent then has no Word tools
and says it "can't edit the document"). On every run the script polls the
live `/mcp` status and, if `word` shows `failed`, first tries a runtime
reconnect (`POST /mcp/word/connect`) and then, as a last resort, restarts
`opencode serve` once.

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

Before asking the assistant to edit a document, make sure a Word document is
actually open — `word-mcp-live` attaches to the active Word instance via
`GetActiveObject`; it does not launch Word itself.

To stop: `npm stop` (unloads the add-in), then Ctrl+C the `opencode serve`
terminal.

## Word automation (`opencode.json`)

```json
{
  "mcp": {
    "word": {
      "type": "local",
      "command": ["uvx", "--from", "word-mcp-live", "word_mcp_server"],
      "enabled": true
    }
  }
}
```

Note the command: the PyPI package is `word-mcp-live`, but the executable it
installs is `word_mcp_server` — `uvx word-mcp-live` (without `--from`) fails
because no executable of that exact name exists. Verify the tool is wired up
correctly with:

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

**A Word edit request errors or times out.** Confirm a Word document is
actually open (`word-mcp-live` needs `GetActiveObject` to find one). Also
avoid running `opencode run ...` against a **locked Windows session** — a
locked/inactive desktop can make Word's COM automation hang indefinitely
with no error, since there's no interactive desktop to attach to; this was
observed firsthand while building Phase 6's test.

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
