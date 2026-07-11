# AI Assistant — Technical Design Spec

This document describes the as-built architecture and technology choices for
AI Assistant. For product requirements see [PRD.md](PRD.md); for setup and
day-to-day usage see [README.md](README.md). This doc focuses on *how it's
built and why*.

## 1. Overview

AI Assistant is a Word task pane add-in that lets a user chat with an AI
agent (OpenCode) and have that agent read and edit the currently open Word
document. It is entirely local and non-admin: no data leaves the machine
except to whatever LLM provider OpenCode is configured to call, and nothing
requires elevated privileges to install or run.

Three independent processes cooperate at runtime:

```
┌─────────────────────────────┐       ┌──────────────────────────┐
│ Word (WebView2 host)         │       │ opencode serve            │
│  └─ task pane (Office.js)    │──────▶│  :4098, --cors for the    │
│     https://localhost:3000   │ fetch │  task pane's origin       │
└─────────────────────────────┘       └────────────┬─────────────┘
                                                     │ stdio (local MCP,
                                                     │ declared in
                                                     │ opencode.json)
                                                     ▼
                                        ┌──────────────────────────┐
                                        │ word-mcp-live (uvx)       │
                                        │  win32com ── GetActiveObject
                                        └────────────┬─────────────┘
                                                     ▼
                                        ┌──────────────────────────┐
                                        │ Word.Application (COM)    │
                                        │  the same open document   │
                                        └──────────────────────────┘
```

No custom backend/proxy server exists. The task pane talks directly to
OpenCode's own REST API, and OpenCode talks directly to word-mcp-live over
stdio as a normal MCP tool server — both are off-the-shelf, not hand-rolled.

## 2. Technology stack

| Layer | Choice | Notes |
|---|---|---|
| Add-in scaffold | `@microsoft/generator-office` (Yeoman) | XML manifest track (`manifest.xml`), not the unified JSON manifest — required for classic Word sideloading |
| Task pane UI | Vanilla JS + Office.js, Fluent UI CSS | no framework; MVP scope doesn't need React/state management |
| Build | Webpack 5 + Babel, `webpack-dev-server` | HTTPS dev server on `:3000` via `office-addin-dev-certs` |
| Markdown rendering | [`marked`](https://www.npmjs.com/package/marked) | renders assistant replies; see §4 for the XSS mitigation |
| Agent runtime | [`opencode`](https://opencode.ai) CLI (`opencode-ai` npm package), run as `opencode serve` | real REST API, not a custom wrapper |
| Word automation | [`word-mcp-live`](https://github.com/ykarapazar/word-mcp-live), invoked via `uvx --from word-mcp-live word_mcp_server` | Python, `pywin32`-based, 40+ fixed COM automation tools; runs isolated via `uv` |
| Test automation | Windows PowerShell 5.1 scripts (`tests/phase*.ps1`) + `playwright-core` for browser-driven UI tests | reuses the system-installed Edge instead of downloading a bundled Chromium |

## 3. Components

### 3.1 Task pane (`src/taskpane/`)

- `taskpane.html` — chat log (`#chat-log`) + input form (`#chat-form`,
  `#chat-input`, `#chat-send`), gated behind `Office.onReady`.
- `taskpane.js` — owns the OpenCode session lifecycle and chat rendering:
  - `ensureSession()` — lazily `POST /session` once per task pane load,
    caches the returned `session.id` in memory (`sessionId` module var).
  - `sendToOpenCode(text)` — `POST /session/{id}/message` with
    `{ model: { providerID, modelID }, parts: [{ type: "text", text }] }`,
    concatenates all returned `text`-type parts as the reply. The model
    is hardcoded via `OPENCODE_MODEL` near the top of the file — there is
    no in-UI model picker (MVP scope).
  - `appendMessage(role, text)` — escapes the text through a `textContent`
    → `innerHTML` round trip *before* `marked.parse()`, so any literal
    HTML/script tags coming from the model or the document are rendered as
    visible text, not executed (see §4).
- `taskpane.css` — flex-column chat layout (log grows, input row pinned to
  the bottom); user/assistant/error message bubbles are visually distinct.

The task pane has **no knowledge of Word document content** beyond what the
user types — it does not read the document itself. All document access goes
through the agent → MCP tool path (§3.3), not through Office.js APIs called
from the task pane.

`manifest.xml`'s ribbon button uses the `Office.AutoShowTaskpaneWithDocument`
sentinel `<TaskpaneId>`, and `Office.onReady` tags each document (via
`Office.context.document.settings`) once its task pane has been opened, so
the pane auto-shows the next time that specific document is reopened —
without a manual ribbon click. This does not apply to brand-new blank
documents (see README). Separately, `scripts/start-background-services.ps1`
(optionally installed to run at Windows logon via
`scripts/install-autostart.ps1`) starts `opencode serve` and a static HTTPS
server for the pre-built taskpane bundle idempotently, so both are already up
by the time Word tries to reconnect to the sideloaded add-in (see §5's WEF
sideload finding). The static server (`scripts/serve-static.js`) serves the
`dist/` bundle produced ahead of time by `npm run build`, reusing the same
`office-addin-dev-certs` certificate and URLs as the webpack dev server;
this replaced launching `webpack serve` at logon, whose per-boot recompile
left port 3000 unserved long enough to trigger Word's "ADD-IN ERROR" (§5).
The webpack dev server (`npm run dev-server`) remains the live-recompile tool
for active development.

The script also self-heals the `word` MCP connection: on every run it polls
the live `/mcp` status and, if `word` came up `failed` (a logon-time cold
start can make opencode's handshake with `word-mcp-live` time out), it tries
a runtime reconnect (`POST /mcp/word/connect`) and then one `opencode serve`
restart as a last resort.

### 3.2 OpenCode server

Run out-of-process via `opencode serve --port 4098 --hostname 127.0.0.1
--cors https://localhost:3000`. This project does not modify OpenCode
itself; it only supplies `opencode.json` (§3.3) as project-level
configuration that OpenCode picks up from its working directory.

The `--cors` flag is load-bearing: the task pane's origin
(`https://localhost:3000`) differs from OpenCode's
(`http://127.0.0.1:4098`), so cross-origin `fetch()` calls are blocked by
the browser without a matching `Access-Control-Allow-Origin` response
header. This was discovered empirically (browser console showed a CORS
error, not a mixed-content error — see §5) and is why the daily-run
instructions in the README always include `--cors`.

### 3.3 Word automation MCP server (`opencode.json`)

```json
{
  "mcp": {
    "word": {
      "type": "local",
      "command": ["uvx", "--from", "word-mcp-live", "word_mcp_server"],
      "enabled": true,
      "environment": {
        "MCP_AUTHOR": "Gu Chao",
        "MCP_AUTHOR_INITIALS": "GC"
      }
    }
  }
}
```

OpenCode spawns this as a local subprocess over stdio (standard MCP
transport) whenever an agent needs a Word tool — no HTTP hop, no custom
protocol glue. `word-mcp-live` itself calls `win32com.client.GetActiveObject
("Word.Application")` to attach to whichever Word instance is already
running; it does not launch Word. This is why a document must already be
open before asking the agent to edit it.

This design was chosen over the PRD's original
`run_word_com_automation(script: str)` tool (arbitrary Python execution
against the COM object) specifically because `word-mcp-live` exposes a
**fixed, auditable set of ~40 tools** (set heading style, insert table,
apply shading, replace selection, etc.) with tracked-changes support,
instead of an open code-execution surface.

### 3.4 Settings panel (gear icon): System Instruction, Persona & Prompt library

The task pane's settings panel (`#settings-toggle`/`#settings-panel` in
`taskpane.html`) has two tabs. The **Settings** tab maintains two per-user
fields — **System Instruction** (standing rules) and **Persona**
(identity/tone) — stored in
`localStorage` (keys `openCodeSystemInstruction`/`openCodePersona`). When
either is set, it is injected as a hidden instruction on each new
session's **first message** (`customizationHiddenBlock` in `taskpane.js`,
same first-message-only gating as the whole-document context), so the
agent behaves accordingly from the very start of every conversation.
Save/Clear reset the cached `sessionId`, so changes apply from the very
next message — no server restart.

The client-side injection design (rather than applying settings to the
server config at runtime) was forced by the `PATCH /config` finding in §5.
A System Instruction can also direct the agent at local folders (e.g.
"read this agent-harness directory and follow its AGENTS.md"), which works
because of `opencode.json`'s blanket
`permission.external_directory: "allow"` (§4) — without it the agent's
first read of an external path would raise an approval prompt the chat UI
cannot answer, hanging the request.

This panel replaced an earlier design that maintained a single "harness
root directory" path (injected via an equivalent hidden block). A one-time
migration (`migrateLegacyHarnessSetting`) rewrites a saved legacy path
into the equivalent read-and-follow System Instruction, then deletes the
legacy `openCodeHarnessRoot` key.

**Harness mode (auto-detected).** `scripts/serve-static.js` (Node, so it has
filesystem access the browser task pane lacks) scans the plugin folder's
ancestors and immediate siblings for a Solution Architect vault — any dir
containing `_agentic/os/AGENTS.md` — and exposes the result at
`GET /harness-info` as `{mode, root}` (signal only, never file contents). On
load the task pane calls `/harness-info` (`applyHarnessMode`); if a vault is
found it hides the manual System Instruction/Persona fields (and their
Save/Clear), shows a banner, and reveals a "Save to memory" pill in the
toolbar beside "Prompt ideas". In this mode the first-message
injection (`getSavedCustomization` → `harnessHiddenBlock`) tells the agent to
read the vault's `SOUL.md`/`USER.md`/`AGENTS.md` (+ recent memory) and adopt
them, with a hard boundary against running the harness's maintenance
scripts/memory-graph/graphify or editing anything unprompted — a deliberate
"read the identity" integration rather than full harness participation
(opencode still runs from the plugin folder, not the vault, avoiding
two-agent/maintenance/config-merge/git-nesting problems). "Save to memory"
(`onSaveToMemory`) fires a silent turn on the current session asking the agent
to append the conversation's key points to the vault's daily note and curate
`MEMORY.md` per the vault's own rules, honoring the memory-graph lock. Absent
`/harness-info` (webpack dev-server, nginx) or with no vault nearby, the task
pane stays in the generic manual mode.

The **Prompt library** tab makes the "Prompt ideas" template library
user-editable. The live library lives in `localStorage` (key
`openCodePromptLibrary`, seeded from `DEFAULT_PROMPT_LIBRARY` in
`taskpane.js` on first load) and is re-read by `deriveLibraryPromptIdeas`
on every 💡 click, so edits apply from the next click with no session
reset. Rows keep their template `id` through edits; the two plan-mode
hidden instructions are deliberately **not** part of the editable data —
they stay in code (`HIDDEN_INSTRUCTIONS`, keyed by template id), so
neither a model turn nor edited library content can inject or alter a
hidden behavior: editing a built-in template's wording keeps its plan-mode
flow, deleting the template drops it. An explicitly saved empty library
stays empty (the 💡 button then explains itself instead of calling the
LLM); "Reset to defaults" restores the built-ins. Each row also has a ⤵
"Use it" button that stages the row's current text into the chat input
(same fill-then-edit behavior as suggestion chips, hidden instruction
included by id) and closes the panel.

## 4. Security posture

- **No arbitrary code execution.** The only Word-mutating surface is
  word-mcp-live's fixed tool list; the agent cannot run arbitrary Python or
  VBA through this add-in.
- **XSS mitigation in the chat log.** Assistant replies (and echoed user
  input) are markdown-rendered. Because the reply text could itself
  originate from document content injected via a prompt, `appendMessage()`
  escapes HTML entities before handing the string to `marked.parse()` —
  literal `<script>`/`<img onerror=...>` etc. show up as visible text
  rather than executing. Markdown syntax (`**bold**`, lists, code spans)
  still works since it doesn't rely on angle brackets. Entity-escaping alone
  does **not** neutralize markdown link/image syntax (`[text](url)`),
  though — a `javascript:` URI in that position would render as a live,
  clickable `<a href="javascript:...">`. `taskpane.js` additionally
  registers a custom `marked` renderer for `link`/`image` that allowlists
  `http:`/`https:`/`mailto:` URL schemes, falling back to plain text for
  anything else.
- **Local-only network surface.** OpenCode listens on `127.0.0.1` only;
  `--cors` is scoped to the single expected origin
  (`https://localhost:3000`), not a wildcard.
- **No admin rights required anywhere** in setup or runtime — the one
  interactive step (trusting the dev HTTPS cert) is a per-user certificate
  store operation, not a system-wide install.
- **`opencode.json` grants blanket external-directory access.**
  `permission.external_directory: "allow"` (a global allow, not scoped to
  one folder) lets the agent's `read`/`write`/`edit`/`bash` tools read and
  modify **any local path**, not just the project folder or a folder a
  System Instruction points at. This is a deliberate posture, not an
  oversight: `opencode`'s
  tools are not sandboxed to the project's working directory, an approval
  prompt can't be answered from the chat UI (the request would silently
  hang), and a per-path grant can't be applied at runtime because
  `PATCH /config` doesn't persist (§5) — while the gear panel (§3.4) must
  support System Instructions that point the agent at arbitrary local
  folders without a server restart. The mitigating context is that this is
  a local, single-user tool whose agent only acts on the user's own
  prompts. There is also no validation of what the user can put in the
  settings fields — intentional, for the same reason. In harness mode (§3.4)
  this same grant is what lets the agent read the detected vault's identity
  files and write its memory notes; it also means the agent can read/write the
  entire vault (e.g. `my-solutions/` client data), so the injected harness
  instruction deliberately forbids unprompted edits and the only write path is
  the explicit "Save to memory" button.

## 5. Notable findings from implementation

- **CORS, not mixed content.** The PRD flagged "HTTPS task pane fetching
  HTTP localhost" as the top risk, assuming a mixed-content block.
  Empirical browser testing showed Chromium's loopback exception already
  permits that; the real blocker was the missing CORS header, fixed with
  `opencode serve --cors <origin>`.
- **`Office.onReady()` never resolves outside a real Office host.** Any
  code that wires up DOM listeners inside it will silently never run under
  a plain browser — this only surfaces during headless/automated testing
  (see `tests/phase5-chat-ui.mjs`, which stubs `office.js` to simulate
  `host: "Word"`), not during normal manual use inside Word.
- **Word COM automation can hang indefinitely on a locked Windows
  session** with no error — there's no interactive desktop for Word to
  attach a window to. Encountered while first running the Phase 6 test;
  documented in the README's troubleshooting section and mitigated in test
  scripts with an explicit `Wait-Job -Timeout`.
- **PyPI package name ≠ installed executable name.** `word-mcp-live` (the
  package) installs an executable called `word_mcp_server`, not
  `word-mcp-live` — `uvx word-mcp-live` fails; the correct invocation is
  `uvx --from word-mcp-live word_mcp_server`.
- **The WEF developer sideload registration is persistent, not one-shot.**
  `npm start` (`office-addin-debugging start`) writes an entry to
  `HKCU:\Software\Microsoft\Office\16.0\WEF\Developer` mapping the add-in's
  manifest GUID to its file path. Word reads this on *every* subsequent
  launch and tries to reconnect to `SourceLocation`
  (`https://localhost:3000/taskpane.html`), indefinitely — not just for the
  session where `npm start` was run. `npm stop` does not reliably clear it.
  Discovered by direct registry inspection while diagnosing a user-reported
  "ADD-IN ERROR" dialog on a plain Word launch with no dev server running.
  Addressed by making the background services (`opencode serve` + dev
  server) start automatically at Windows logon (`scripts/`, §3.1) instead of
  trying to undo the persistent registration.

- **`PATCH /config` does not persist anything, despite what the docs
  imply.** While building the gear-icon settings feature (§3.4, at the
  time a harness-root config), a
  live `opencode serve` instance (v1.16.2) was tested directly with `curl`:
  `PATCH /config` returns `HTTP 200` and echoes the request body back (or
  `HTTP 500 UnknownError` for payloads containing Windows-style backslash
  paths), but a subsequent `GET /config` showed the change was never
  actually applied — confirmed even for a trivial control field
  (`username`), with and without `?directory=`/`?workspace=` query
  params. Because of this, the settings feature does **not** attempt a
  runtime patch — it sidesteps server config entirely with a hidden
  first-message instruction (§3.4), enabled where file access is involved
  by a blanket `external_directory: "allow"` in the static config (§4).

## 6. Testing strategy

Each phase of the build has a corresponding `tests/phaseN-*.ps1` script
(see README for the full table); `tests/run-all.ps1` runs all of them in
order and fails fast on the first regression. Nothing is mocked at the
integration level — Phase 5 drives the real task pane in a real (headless)
browser against a real running `opencode serve`, and Phase 6 drives a real
Word document via a real `opencode run` + word-mcp-live call, then verifies
the mutation by reading the document back over COM. This was a deliberate
choice: static/build checks alone would have missed both the CORS issue and
the `Office.onReady` stubbing requirement.

## 7. Directory structure

```
opencode+office/
├── PRD.md                    product requirements (original + revised architecture)
├── README.md                 setup, daily run, troubleshooting
├── tech-design-spec.md        this document
├── opencode.json              word-mcp-live MCP server registration
├── manifest.xml                Office Add-in manifest (XML track)
├── package.json / webpack.config.js / babel.config.json
├── assets/                    add-in icons
├── src/
│   ├── taskpane/               chat UI (html/js/css)
│   └── commands/                generator-office boilerplate ribbon command stub
├── scripts/
│   ├── start-background-services.ps1   idempotent opencode serve + static server launcher (+ MCP self-heal)
│   ├── serve-static.js                  tiny HTTPS static server for the pre-built dist/ bundle
│   ├── install-autostart.ps1            builds dist/, writes Startup-folder .vbs autostart wrapper
│   └── uninstall-autostart.ps1          removes it
└── tests/
    ├── run-all.ps1              regression runner
    └── phase1..10-*.ps1/.mjs     per-phase tests
```
