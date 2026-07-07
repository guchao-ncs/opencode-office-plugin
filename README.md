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

This writes a `.vbs` wrapper to your Startup folder
(`%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup`) that silently runs
`scripts\start-background-services.ps1` at logon. That script is idempotent —
it checks whether ports 3000/4098 are already listening before starting
anything, so it's safe even if you also start things manually. Logs go to
`%LOCALAPPDATA%\AI Assistant\logs\`.

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

## Optional: connecting an external "harness" folder (gear icon)

The task pane's header has a gear icon (⚙) that opens a small settings
panel where you can type the path to any other local folder — for example
an Obsidian vault or agent-harness project that has its own `AGENTS.md` —
and click **Generate config**. This produces a JSON snippet like:

```json
{
  "instructions": ["C:\\path\\to\\your\\agent-harness-folder\\AGENTS.md"],
  "permission": { "external_directory": { "C:\\path\\to\\your\\agent-harness-folder\\**": "allow" } }
}
```

Copy this into the top level of this project's `opencode.json` (next to
`"mcp"`), then restart `opencode serve` for it to take effect. The path you
type is remembered in the browser's `localStorage` so it re-fills the input
next time you open the panel, but **that's a UI convenience only** — it does
not get applied anywhere by itself. There is no live/automatic way to apply
this from the task pane; opencode's `PATCH /config` endpoint looks like it
should do this at runtime, but was found not to persist anything against a
real running server (see `tech-design-spec.md` §5), so this feature only
generates the snippet for you to paste in by hand.

Note that `permission.external_directory: "allow"` grants the agent's
`read`/`write`/`edit`/`bash` tools full access to that folder, not just
read access — only point this at a folder you're comfortable letting the
agent modify. Click **Clear** to wipe the saved path from `localStorage`
(you'll still need to manually remove the fields from `opencode.json` and
restart `opencode serve` to actually revoke the access).

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
| `phase9-harness-config.ps1` (+ `phase9-harness-config.mjs`) | the gear-icon settings panel: opens, generates a correct `opencode.json` snippet, persists/re-fills the typed path via `localStorage` across reloads, and Clear wipes it — driven headlessly via `playwright-core`, no `opencode serve` needed since this is a pure client-side feature |

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
