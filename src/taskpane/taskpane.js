/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { marked } from "marked";

// OpenCode's real REST API (opencode serve), not the PRD's originally-assumed /v1/chat.
// See PRD.md FR-2 for the mixed-content (https taskpane -> http localhost) risk this depends on.
const OPENCODE_BASE_URL = "http://127.0.0.1:4098";
const OPENCODE_MODEL = { providerID: "sunshine-coder", modelID: "ssc-chat-latest" };

// escapeHtml() (below) only neutralizes literal HTML tags, not markdown link/image
// syntax - marked would otherwise happily turn `[text](javascript:...)` into a
// clickable, executable `<a href="javascript:...">`, since escaping `[`, `]`, `(`, `)`
// isn't needed to keep them literal text. Restrict rendered link/image URLs to safe
// schemes so a malicious link (typed by the user, or echoed back from document
// content by the model) can't execute script when clicked.
function isSafeUrl(href) {
  if (!href) {
    return false;
  }
  try {
    const url = new URL(href, "https://opencode-taskpane.invalid/");
    return ["http:", "https:", "mailto:"].includes(url.protocol);
  } catch {
    return false;
  }
}

marked.use({
  renderer: {
    link({ href, title, tokens }) {
      const text = this.parser.parseInline(tokens);
      if (!isSafeUrl(href)) {
        return text;
      }
      const titleAttr = title ? ` title="${title}"` : "";
      return `<a href="${href}"${titleAttr} target="_blank" rel="noopener noreferrer">${text}</a>`;
    },
    image({ href, title, text }) {
      if (!isSafeUrl(href)) {
        return text;
      }
      const titleAttr = title ? ` title="${title}"` : "";
      return `<img src="${href}" alt="${text}"${titleAttr}>`;
    },
  },
});

// Storage keys for the gear-icon Settings panel. The two fields are injected
// as a hidden instruction on each new session's first message (see
// customizationHiddenBlock). LEGACY_HARNESS_STORAGE_KEY is the pre-rework
// "harness root directory" path - migrated into a System Instruction once,
// then removed (see migrateLegacyHarnessSetting).
const SYSTEM_INSTRUCTION_STORAGE_KEY = "openCodeSystemInstruction";
const PERSONA_STORAGE_KEY = "openCodePersona";
const LEGACY_HARNESS_STORAGE_KEY = "openCodeHarnessRoot";

// Set once at load from the static server's /harness-info signal (see
// applyHarnessMode). Non-null = a Solution Architect harness vault was
// detected on this machine, and its absolute root path is here. In that
// "harness mode" the manual System Instruction/Persona fields are disabled
// and behavior instead comes from a scoped instruction pointing the agent at
// the vault's identity/memory files (see harnessHiddenBlock). Stays null in
// generic mode (no vault detected, or served without /harness-info).
let harnessRoot = null;

// Playful rotating status words shown while waiting on the model, in the
// style of Claude Code's CLI status line ("Pontificating...", "Sussing...").
const THINKING_WORDS = [
  "Doing",
  "Forming",
  "Accomplishing",
  "Whirring",
  "Pontificating",
  "Germinating",
  "Percolating",
  "Schlepping",
  "Herding",
  "Effecting",
  "Channelling",
  "Cogitating",
  "Sussing",
];

function pickThinkingWord(exclude) {
  let word;
  do {
    word = THINKING_WORDS[Math.floor(Math.random() * THINKING_WORDS.length)];
  } while (THINKING_WORDS.length > 1 && word === exclude);
  return word;
}

let sessionId = null;
let lastDocumentUrl = null;
let lastDocumentHash = null;
let lastDocumentText = null;
let documentSessionNonce = null;

// Set by a doc-suggestion chip (see renderDocSuggestions) that needs to carry
// extra hidden guidance alongside its literal chip text - e.g. the "Revise
// the doc based on winning theme" chip attaches instructions to ask a
// clarifying question and enter plan-mode-only first. Chips stage their text
// into the input for the user to edit before sending (they don't auto-send),
// so this stays pending until the next onSubmit consumes it - or until the
// user wipes the staged text entirely, which drops it (see
// onChatInputChanged), since the instruction belongs to the chip's prompt,
// not to whatever unrelated message gets typed next.
let pendingHiddenInstruction = null;

// True from onSubmit accepting a message until its reply fully settles. While
// set, the send button is in "stop" mode (clicking it aborts the run instead
// of sending) and Enter is ignored - see onSubmit/onInputKeyDown.
let isSending = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("chat-form").addEventListener("submit", onSubmit);
    document.getElementById("chat-input").addEventListener("keydown", onInputKeyDown);
    document.getElementById("chat-input").addEventListener("input", onChatInputChanged);
    document.getElementById("settings-toggle").addEventListener("click", onSettingsToggle);
    document.getElementById("settings-save").addEventListener("click", onSettingsSave);
    document.getElementById("settings-clear").addEventListener("click", onSettingsClear);
    document.getElementById("tab-btn-settings").addEventListener("click", () => onTabSwitch("settings"));
    document.getElementById("tab-btn-library").addEventListener("click", () => onTabSwitch("library"));
    document.getElementById("library-add").addEventListener("click", onLibraryAdd);
    document.getElementById("library-save").addEventListener("click", onLibrarySave);
    document.getElementById("library-reset").addEventListener("click", onLibraryReset);
    document.getElementById("memory-save").addEventListener("click", onSaveToMemory);
    document.getElementById("prompt-ideas-toggle").addEventListener("click", togglePromptIdeas);

    // Tags this document so Word auto-shows the taskpane the next time this same
    // file is reopened, instead of requiring a manual ribbon-button click every time.
    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
    Office.context.document.settings.saveAsync();

    // One-time migration of the pre-rework harness-root setting, then re-fill
    // the Settings fields from what the user saved last time. The fields only
    // live in localStorage - they take effect via the hidden first-message
    // instruction (customizationHiddenBlock), not any server-side config.
    migrateLegacyHarnessSetting();
    document.getElementById("system-instruction-input").value =
      localStorage.getItem(SYSTEM_INSTRUCTION_STORAGE_KEY) || "";
    document.getElementById("persona-input").value = localStorage.getItem(PERSONA_STORAGE_KEY) || "";

    // Ask the static server whether a harness vault was detected on this
    // machine, and switch the Settings tab to harness mode if so. Best-effort:
    // on webpack dev-server / nginx (no /harness-info) or any error, this
    // silently leaves generic mode in place.
    applyHarnessMode();

    // Open the SSE stream up front (not lazily on first send) so it has time to
    // reach OPEN before any message is sent - see connectEventStream()'s comment
    // for why sending before that race is settled would drop the first events.
    connectEventStream();

    // Best-effort, fire-and-forget: populate the on-load "explore your
    // document" chip panel. Not awaited - it must not delay the taskpane
    // becoming interactive.
    initializeDocumentSuggestions();

    // Best-effort: close our SSE connection when this taskpane instance goes
    // away (document closed, add-in reloaded, etc). opencode's own /event
    // broadcaster was observed accumulating listeners for connections that
    // were never cleanly closed (a MaxListenersExceededWarning in its server
    // log after many dev-session reloads) - this doesn't fix that server-side
    // behavior, but it does stop this client from being one more connection
    // left dangling past its actual lifetime.
    window.addEventListener("unload", () => {
      if (eventSource) {
        eventSource.close();
      }
    });
  }
});

function onSettingsToggle() {
  const panel = document.getElementById("settings-panel");
  panel.style.display = panel.style.display === "none" ? "flex" : "none";
}

// Queries the static server's harness signal and, if a vault was detected,
// switches the Settings tab into harness mode: the manual System
// Instruction/Persona fields are disabled (identity comes from the vault via
// harnessHiddenBlock), Save/Clear are replaced by "Save to memory", and a
// banner shows which vault is in effect. Best-effort - on any error, or when
// served without /harness-info (webpack dev-server / nginx), it leaves the
// editable generic mode untouched.
async function applyHarnessMode() {
  let info;
  try {
    const res = await fetch("/harness-info", { cache: "no-store" });
    if (!res.ok) {
      return;
    }
    info = await res.json();
  } catch {
    return;
  }
  if (!info || info.mode !== "harness" || !info.root) {
    return;
  }
  harnessRoot = info.root;

  // Identity comes from the vault, so the manual System Instruction/Persona
  // fields don't apply - hide them entirely and explain via the banner.
  document.getElementById("manual-settings-fields").style.display = "none";
  const banner = document.getElementById("harness-banner");
  banner.textContent =
    `Harness mode — identity, user context and conventions come from ${harnessRoot}\\_agentic ` +
    `(AGENTS.md / SOUL.md / USER.md).`;
  banner.style.display = "block";

  // Reveal the "Save to memory" pill in the toolbar beside "Prompt ideas".
  document.getElementById("memory-save").style.display = "flex";
  document.getElementById("prompt-ideas-toolbar").style.display = "flex";
}

// "Save to memory" (harness mode only): fires a silent turn asking the agent
// to append this conversation's key points to the vault's daily note and
// promote durable items to MEMORY.md, per the harness's own memory rules -
// honoring the memory-graph lock and never running maintenance scripts. Uses
// the current session (not a fresh one) so the agent has the conversation to
// summarize.
// The "Save to memory" pill lives in the toolbar (not the settings panel), so
// feedback is shown on the button label itself - "Saving..." → "Saved ✓" /
// "Nothing to save" / "Save failed" - reverting after a moment.
async function onSaveToMemory() {
  if (!harnessRoot) {
    return;
  }
  const btn = document.getElementById("memory-save");
  const label = document.getElementById("memory-save-label");
  const restore = (text) => {
    setTimeout(() => {
      label.textContent = "Save to memory";
      btn.disabled = false;
    }, 2500);
    label.textContent = text;
  };
  if (isSending) {
    // A reply is still streaming (or a tool call is stuck) on this session -
    // posting the memory-write now would just queue behind it and hang on
    // "Saving...". Make the user wait for the current turn to finish first.
    btn.disabled = true;
    restore("Finish the reply first");
    return;
  }
  if (!sessionId) {
    btn.disabled = true;
    restore("Nothing to save yet");
    return;
  }
  btn.disabled = true;
  label.textContent = "Saving...";
  try {
    await sendToOpenCodeBlocking(sessionId, [{ type: "text", text: buildSaveToMemoryInstruction() }], 150000);
    restore("Saved ✓");
  } catch {
    restore("Save failed");
  }
}

function buildSaveToMemoryInstruction() {
  const os = `${harnessRoot}\\_agentic\\os`;
  return (
    "Save the key points of THIS conversation into the Solution Architect harness memory, following the memory " +
    `rules in ${os}\\AGENTS.md. Specifically:\n` +
    `1. If ${os}\\memory-graph\\.memory-graph.lock exists, do NOT write - reply exactly "memory is locked, try ` +
    'again later" and stop.\n' +
    `2. Otherwise append a concise, structured summary of what we discussed/decided/still-to-follow-up to ` +
    `${os}\\memory\\<today's date YYYY-MM-DD>.md (create the file if it doesn't exist; NEVER overwrite existing ` +
    "content - append only).\n" +
    `3. Promote any durable feedback, decisions, or boundaries worth remembering long-term into ${os}\\MEMORY.md ` +
    "(create it if missing; curate, don't just dump).\n" +
    "Do NOT run any maintenance scripts, memory-graph, or graphify. When done, reply with a one-line confirmation " +
    "of what you wrote and to which file(s)."
  );
}

function onSettingsSave() {
  const systemInstruction = document.getElementById("system-instruction-input").value.trim();
  const persona = document.getElementById("persona-input").value.trim();
  if (systemInstruction) {
    localStorage.setItem(SYSTEM_INSTRUCTION_STORAGE_KEY, systemInstruction);
  } else {
    localStorage.removeItem(SYSTEM_INSTRUCTION_STORAGE_KEY);
  }
  if (persona) {
    localStorage.setItem(PERSONA_STORAGE_KEY, persona);
  } else {
    localStorage.removeItem(PERSONA_STORAGE_KEY);
  }
  // Drop the current session so the very next message starts a fresh one
  // whose first message carries the updated settings (see
  // customizationHiddenBlock) - no server restart needed.
  sessionId = null;
  setSettingsStatus(
    systemInstruction || persona
      ? "Saved - applies from your next message onward."
      : "Saved (both fields empty) - new conversations will use default behavior.",
    false
  );
}

function onSettingsClear() {
  document.getElementById("system-instruction-input").value = "";
  document.getElementById("persona-input").value = "";
  localStorage.removeItem(SYSTEM_INSTRUCTION_STORAGE_KEY);
  localStorage.removeItem(PERSONA_STORAGE_KEY);
  // Same session reset as Save: the next message starts a fresh session with
  // no custom settings attached.
  sessionId = null;
  setSettingsStatus("Cleared - new conversations will use default behavior.", false);
}

function setSettingsStatus(text, isError) {
  const el = document.getElementById("settings-status");
  el.textContent = text;
  el.style.color = isError ? "#a80000" : "#605e5c";
}

function setLibraryStatus(text, isError) {
  const el = document.getElementById("library-status");
  el.textContent = text;
  el.style.color = isError ? "#a80000" : "#605e5c";
}

// Switches between the panel's "Settings" and "Prompt library" tabs. The
// library list is (re-)rendered from localStorage on every switch INTO the
// library tab, discarding unsaved row edits - saved state is the only truth.
function onTabSwitch(tab) {
  const isSettings = tab === "settings";
  if (!isSettings) {
    renderPromptLibraryList();
  }
  document.getElementById("tab-settings").style.display = isSettings ? "flex" : "none";
  document.getElementById("tab-library").style.display = isSettings ? "none" : "flex";
  document.getElementById("tab-btn-settings").classList.toggle("settings-tab--active", isSettings);
  document.getElementById("tab-btn-library").classList.toggle("settings-tab--active", !isSettings);
}

// Renders the editable prompt-library rows from the saved library. Each row
// keeps its template's id (and category) in data attributes so edits preserve
// identity - which is what keeps a HIDDEN_INSTRUCTIONS behavior attached to
// an edited built-in template, and lets isPromptAlreadyUsed/dedup work
// consistently.
function renderPromptLibraryList() {
  const list = document.getElementById("prompt-library-list");
  list.innerHTML = "";
  for (const prompt of getPromptLibrary()) {
    list.appendChild(createPromptRow(prompt));
  }
}

function createPromptRow(prompt) {
  const row = document.createElement("div");
  row.className = "prompt-library-row";
  row.dataset.promptId = prompt.id;
  row.dataset.promptCategory = prompt.category || "My prompts";
  const textarea = document.createElement("textarea");
  textarea.className = "settings-textarea prompt-library-text";
  textarea.rows = 2;
  textarea.value = prompt.template || "";

  const removeBtn = document.createElement("button");
  removeBtn.type = "button";
  removeBtn.className = "prompt-library-remove";
  removeBtn.title = "Remove this prompt";
  removeBtn.setAttribute("aria-label", "Remove this prompt");
  removeBtn.textContent = "✕";
  removeBtn.addEventListener("click", () => row.remove());

  const useBtn = document.createElement("button");
  useBtn.type = "button";
  useBtn.className = "prompt-library-use";
  useBtn.title = "Use this prompt";
  useBtn.setAttribute("aria-label", "Use this prompt");
  useBtn.textContent = "⤵";
  useBtn.addEventListener("click", () => {
    // Stage this prompt (the row's CURRENT text, including unsaved edits)
    // into the chat input for review/editing - same fill-then-edit behavior
    // as suggestion chips, including the template's hidden instruction (by
    // id, from HIDDEN_INSTRUCTIONS - never from editable data). Close the
    // panel so the staged text is immediately visible.
    const text = textarea.value.trim();
    if (!text) {
      return;
    }
    pendingHiddenInstruction = HIDDEN_INSTRUCTIONS[prompt.id] || null;
    const input = document.getElementById("chat-input");
    input.value = text;
    onChatInputChanged();
    document.getElementById("settings-panel").style.display = "none";
    input.focus();
  });

  // Vertical action stack next to the textarea: remove on top, use below it.
  const actions = document.createElement("div");
  actions.className = "prompt-library-row-actions";
  actions.append(removeBtn, useBtn);
  row.append(textarea, actions);
  return row;
}

function onLibraryAdd() {
  const list = document.getElementById("prompt-library-list");
  const row = createPromptRow({ id: `user-${Date.now()}`, category: "My prompts", template: "" });
  list.appendChild(row);
  row.querySelector("textarea").focus();
}

function onLibrarySave() {
  const rows = document.querySelectorAll("#prompt-library-list .prompt-library-row");
  const library = [];
  for (const row of rows) {
    const template = row.querySelector("textarea").value.trim();
    if (!template) {
      continue;
    }
    library.push({ id: row.dataset.promptId, category: row.dataset.promptCategory, template });
  }
  localStorage.setItem(PROMPT_LIBRARY_STORAGE_KEY, JSON.stringify(library));
  renderPromptLibraryList();
  setLibraryStatus(
    `Saved ${library.length} prompt${library.length === 1 ? "" : "s"} - "Prompt ideas" picks from them on its next click.`,
    false
  );
}

function onLibraryReset() {
  localStorage.setItem(PROMPT_LIBRARY_STORAGE_KEY, JSON.stringify(DEFAULT_PROMPT_LIBRARY));
  renderPromptLibraryList();
  setLibraryStatus("Restored the built-in default prompts.", false);
}

// One-time migration from the pre-rework settings panel, which maintained a
// single "harness root directory" path instead of the System Instruction /
// Persona fields. Rather than silently dropping that configured behavior, it
// is rewritten as an equivalent System Instruction (the injection mechanism
// is the same hidden first-message block either way), so an existing user's
// harness workflow keeps working across the UI change without re-entering
// anything. The legacy key is then removed so this runs at most once.
function migrateLegacyHarnessSetting() {
  const legacyRoot = (localStorage.getItem(LEGACY_HARNESS_STORAGE_KEY) || "").trim();
  if (!legacyRoot) {
    return;
  }
  if (!localStorage.getItem(SYSTEM_INSTRUCTION_STORAGE_KEY)) {
    localStorage.setItem(
      SYSTEM_INSTRUCTION_STORAGE_KEY,
      "Read the agent-harness directory at " +
        legacyRoot +
        " with your file tools - start with AGENTS.md in its root (if present), then the companion files it " +
        "references (e.g. SOUL.md, USER.md, memory/, skills/) - and adopt and follow the instructions, persona, " +
        "and conventions defined there for the entire conversation."
    );
  }
  localStorage.removeItem(LEGACY_HARNESS_STORAGE_KEY);
}

function onInputKeyDown(event) {
  if (event.key === "Enter" && !event.shiftKey) {
    event.preventDefault();
    // While a reply is in flight the send button is in stop mode - Enter must
    // neither trigger an accidental abort nor start a second overlapping turn
    // on the same session (which corrupts both turns' message-ID tracking -
    // see activeReply's comment), so it is simply ignored until the run
    // settles.
    if (isSending) {
      return;
    }
    document.getElementById("chat-form").requestSubmit();
  }
}

// Keeps the send button's enabled state in sync with whether there is
// anything to send (Claude.ai-style: greyed out on an empty input). Also
// drops a chip-staged hidden instruction if the user wipes the staged text -
// the instruction belongs to that chip's prompt, not to whatever they type
// next. Programmatic input.value writes don't fire "input", so chip handlers
// call this directly after staging.
function onChatInputChanged() {
  const input = document.getElementById("chat-input");
  if (!isSending) {
    document.getElementById("chat-send").disabled = !input.value.trim();
  }
  if (!input.value.trim()) {
    pendingHiddenInstruction = null;
  }
}

// Swaps the send button between its two modes: a send arrow (disabled when
// the input is empty) and, while a reply is in flight, an always-enabled stop
// square that aborts the run.
function setSendButtonMode(stopping) {
  const btn = document.getElementById("chat-send");
  btn.classList.toggle("chat-send--stop", stopping);
  const label = stopping ? "Stop generating" : "Send";
  btn.title = label;
  btn.setAttribute("aria-label", label);
  btn.disabled = stopping ? false : !document.getElementById("chat-input").value.trim();
}

async function onSubmit(event) {
  event.preventDefault();
  if (isSending) {
    // The button is in stop mode while a reply is in flight, so this submit
    // is the user clicking Stop, not sending a new message.
    stopActiveRun();
    return;
  }
  const input = document.getElementById("chat-input");
  const text = input.value.trim();
  if (!text) {
    return;
  }
  input.value = "";
  appendMessage("user", text);
  sentPromptTexts.push(normalizePromptForComparison(text));
  // The on-load suggestion chips are a pre-conversation affordance (Adobe's
  // "explore your document" panel) - once the user has actually sent
  // anything, real chat history takes over and the panel would just be dead
  // space below it. The "Prompt ideas" toggle itself stays put so it can
  // still be used to re-open a freshly-filtered view later.
  document.getElementById("doc-suggestions-panel").style.display = "none";

  isSending = true;
  setSendButtonMode(true);

  // Read-and-clear: a doc-suggestion chip may have stashed extra hidden
  // guidance for this specific message (see pendingHiddenInstruction above).
  const hiddenInstruction = pendingHiddenInstruction;
  pendingHiddenInstruction = null;

  try {
    const { id: sid, isNew } = await ensureSession();
    // Only fetched for a brand-new session: once sent, the document context
    // becomes part of that session's message history and stays visible to the
    // model on every later turn, so there is no need to re-fetch/re-send it
    // on every message of the same conversation.
    const documentContext = isNew ? lastDocumentText : null;
    // Same first-message-only gating as the document context: once the
    // settings instruction is in the session's history, it stays visible to
    // the model on every later turn.
    const customization = isNew ? getSavedCustomization() : null;
    // Fetched fresh on every message, unlike the whole-document context above:
    // a selection is a right-now, in-the-moment thing (the user highlighting a
    // paragraph and then saying "rewrite this"), not something that stays
    // valid for the rest of the conversation the way the document's content
    // (mostly) does.
    const selectedText = await getSelectedText();
    await streamAssistantReply(sid, text, documentContext, selectedText, hiddenInstruction, customization);
  } catch (err) {
    appendMessage(
      "error",
      `Could not reach OpenCode at ${OPENCODE_BASE_URL}. Make sure "opencode serve --port 4098" is running. (${err.message})`
    );
  } finally {
    isSending = false;
    setSendButtonMode(false);
  }
}

async function ensureSession() {
  const docText = await getActiveDocumentText() || "";
  const docUrl = (typeof Office !== "undefined" && Office.context && Office.context.document) ? Office.context.document.url : null;
  const docHash = computeTextHash(docText);

  if (sessionId && (docUrl !== lastDocumentUrl || docHash !== lastDocumentHash)) {
    sessionId = null;
    sentPromptTexts = [];
  }
  lastDocumentUrl = docUrl;
  lastDocumentHash = docHash;
  lastDocumentText = docText;

  if (sessionId) {
    return { id: sessionId, isNew: false };
  }
  const res = await fetch(`${OPENCODE_BASE_URL}/session`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({}),
  });
  if (!res.ok) {
    throw new Error(`session create failed: HTTP ${res.status}`);
  }
  const session = await res.json();
  sessionId = session.id;
  return { id: sessionId, isNew: true };
}

function computeTextHash(text) {
  if (!text) return "empty";
  let hash = 0;
  for (let i = 0; i < text.length; i++) {
    const chr = text.charCodeAt(i);
    hash = ((hash << 5) - hash) + chr;
    hash |= 0;
  }
  return "h-" + hash.toString(36);
}

function getOrGenerateSessionNonce(isUnsaved) {
  if (!isUnsaved) {
    documentSessionNonce = null;
    return null;
  }
  if (!documentSessionNonce) {
    documentSessionNonce = "nonce-" + Math.random().toString(36).substring(2, 10);
  }
  return documentSessionNonce;
}

function getDocumentIdentity() {
  const url = (typeof Office !== "undefined" && Office.context && Office.context.document) ? Office.context.document.url : null;
  const isUnsaved = !url;
  const nonce = getOrGenerateSessionNonce(isUnsaved);
  let displayName = "Unsaved Document";
  if (url) {
    const lastSlash = Math.max(url.lastIndexOf("/"), url.lastIndexOf("\\"));
    if (lastSlash !== -1) {
      displayName = url.substring(lastSlash + 1);
    } else {
      displayName = url;
    }
  }
  return {
    url: url || null,
    isUnsaved,
    nonce,
    displayName
  };
}

function buildDocumentIdentityBlock(docText) {
  const docId = getDocumentIdentity();
  const textHash = computeTextHash(docText);
  const textLength = docText ? docText.length : 0;
  const textPreview = docText ? docText.substring(0, 150).replace(/\r?\n/g, " ") + "..." : "None";

  let block = "=== HOST DOCUMENT IDENTITY AND BINDING CONTRACT ===\n";
  if (docId.isUnsaved) {
    block += `Document Type: Unsaved Document\n`;
    block += `Fallback Display Name: ${docId.displayName}\n`;
    block += `Add-in Session Nonce: ${docId.nonce}\n`;
    block += `WARNING: This is a newly created, unsaved document. Duplicate 'Document1'-style files cannot be safely disambiguated by display name alone.\n`;
  } else {
    block += `Document Type: Saved Document\n`;
    block += `Document URL/Path: ${docId.url}\n`;
    block += `Document Name: ${docId.displayName}\n`;
  }
  block += `Fresh Text Hash: ${textHash}\n`;
  block += `Fresh Text Length: ${textLength} characters\n`;
  block += `Fresh Text Preview: "${textPreview}"\n`;
  block += `==================================================\n\n`;

  block += "CRITICAL BINDING CONTRACT & SAFETY PROTOCOL FOR MUTATIONS:\n";
  block += "Before calling any live-editing / mutation tool (such as format_text, add_heading, add_table, search_and_replace, word_live_replace_text, create_custom_style, set_paragraph_spacing, set_table_cell_shading, format_table, etc.):\n\n";

  block += "1. Read-only actions (summarizing, reading context) are permitted without restrictions.\n";
  block += "2. For saved documents (Document Type: Saved Document):\n";
  block += "   - You MUST call `word_live_list_open` first.\n";
  block += "   - Match the 'Document URL/Path' or 'Document Name' above to exactly one open document returned by `word_live_list_open`.\n";
  block += "   - You MUST pass the matched path or name as the `filename` parameter to all `word_live_*` editing tools. Never pass null or leave `filename` omitted.\n";
  block += "   - If no exact match or multiple matches exist, you MUST stop immediately, do not mutate, and ask the user to save or close duplicate files.\n";
  block += "3. For unsaved documents (Document Type: Unsaved Document):\n";
  block += "   - Destructive live edits are blocked by default.\n";
  block += "   - You MUST first call `word_live_list_open` to inspect all open documents.\n";
  block += "   - If `word_live_list_open` shows exactly one open unsaved document:\n";
  block += "     * You may proceed with the edit only after explicitly telling the user: 'I see only one unsaved document open. Please confirm if you want me to edit this document or if you prefer to save it first.'\n";
  block += "   - If `word_live_list_open` shows two or more unsaved documents (or you cannot be sure which one is hosting the taskpane):\n";
  block += "     * You MUST fail closed and stop immediately. Politely ask the user to save the target document first or close all other unsaved documents so the target can be uniquely identified.\n";
  block += "4. Mismatch Prevention:\n";
  block += "   - Do not perform any mutation if the actual document's content/structure on disk/Word does not correspond to the 'Fresh Text Hash' and preview details above.\n";

  return block;
}

// Reads the full text of the currently-open Word document via the Word JS API
// (Word.run), so the agent has it as context without needing to be told a
// file path or think to call a word-mcp-live tool on its own. `Word` only
// exists inside a real Word host - the office.js stub used by headless
// browser tests doesn't define it - so this resolves null there and the
// caller just sends the message without document context, same as before.
function getActiveDocumentText() {
  return new Promise((resolve) => {
    if (typeof Word === "undefined" || !Word.run) {
      resolve(null);
      return;
    }
    Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      resolve(body.text || "");
    }).catch(() => resolve(null));
  });
}

// Reads whatever text is currently highlighted/selected in the document, via
// Word.run's context.document.getSelection(). Unlike getActiveDocumentText()
// (the whole body, fetched only once at session start), this is fetched on
// EVERY message, because a selection is a per-turn, in-the-moment thing the
// user is actively pointing at right now - re-sending the whole document on
// every turn would be wasteful, but a selection is small and changes turn to
// turn. Resolves null (not just "") when there is no active selection (an
// empty/collapsed range, or non-Word host), so buildPromptParts can tell
// "nothing selected" apart from "selected text happens to be empty".
function getSelectedText() {
  return new Promise((resolve) => {
    if (typeof Word === "undefined" || !Word.run) {
      resolve(null);
      return;
    }
    Word.run(async (context) => {
      const range = context.document.getSelection();
      range.load("text");
      await context.sync();
      resolve(range.text && range.text.trim() ? range.text : null);
    }).catch(() => resolve(null));
  });
}

// Marker the model is asked to append after every answer, carrying its 2
// follow-up-suggestion strings as JSON. Generating suggestions this way -
// piggybacked onto the one reply already being waited for - replaces an
// earlier design that spun up a second, throwaway `POST /session` purely to
// ask "suggest 2 follow-ups": that session paid the full cost of the
// project's global `instructions` (opencode.json's harness AGENTS.md) being
// reloaded from scratch with zero prompt-cache benefit (empirically ~9s and
// ~13.7k input tokens against this project's configured harness), for a
// trivial sub-task. Piggybacking removes that entire extra round trip.
const SUGGESTIONS_MARKER_RE = /<!--\s*SUGGESTIONS:([\s\S]*?)-->/i;
const SUGGESTIONS_TAIL_RE = /<!--\s*SUGGESTIONS:[\s\S]*$/i;

// What to inject on a session's first message: a harness marker when a vault
// was detected (harness mode overrides the manual fields), otherwise the
// gear-panel System Instruction/Persona, or null when nothing is configured.
function getSavedCustomization() {
  if (harnessRoot) {
    return { harnessRoot };
  }
  const systemInstruction = (localStorage.getItem(SYSTEM_INSTRUCTION_STORAGE_KEY) || "").trim();
  const persona = (localStorage.getItem(PERSONA_STORAGE_KEY) || "").trim();
  if (!systemInstruction && !persona) {
    return null;
  }
  return { systemInstruction, persona };
}

// Scoped harness instruction sent on a session's FIRST message when a vault is
// detected. It points the agent at the vault's identity/memory files and has
// it adopt them, but deliberately draws a tight boundary: NO maintenance
// scripts, memory-graph, graphify, or the memory-maintenance checkpoint (those
// are the CLI-agent's job and would be slow/inappropriate from a Word turn),
// and no unprompted edits. This is a controlled, narrow read of the harness -
// not full harness participation (opencode is not run inside the vault).
function harnessHiddenBlock(root) {
  return (
    "A Solution Architect harness is installed on this machine at: " +
    root +
    "\n\nBefore handling the user's request below, read these files with your file tools and adopt the identity, " +
    "user context, and working conventions they define, for this entire conversation:\n" +
    "- " + root + "\\_agentic\\os\\SOUL.md  (who you are)\n" +
    "- " + root + "\\_agentic\\os\\USER.md  (who you're helping)\n" +
    "- " + root + "\\_agentic\\os\\AGENTS.md  (operating conventions)\n" +
    "For continuity, also read today's and yesterday's " + root + "\\_agentic\\os\\memory\\YYYY-MM-DD.md and " +
    root + "\\_agentic\\os\\MEMORY.md IF they exist (skip silently if they don't).\n\n" +
    "Hard boundaries for this Word add-in context: do NOT run any maintenance scripts, the memory-maintenance " +
    "checkpoint, memory-graph, or graphify - even if AGENTS.md tells you to; those are handled elsewhere. Do NOT " +
    "modify any file in the harness unless the user explicitly asks (memory writes happen only via the add-in's " +
    "\"Save to memory\" button). Keep the persona's concise, structured style."
  );
}

// Hidden instruction sent on a session's FIRST message (same isNew gating as
// the whole-document context). In harness mode it carries the scoped harness
// block; otherwise it carries the gear panel's System Instruction / Persona
// fields. This client-side injection is what makes either take effect live,
// with no server restart - opencode's PATCH /config was confirmed not to
// persist anything (see tech-design-spec.md §5).
function customizationHiddenBlock(customization) {
  if (customization.harnessRoot) {
    return harnessHiddenBlock(customization.harnessRoot);
  }
  const parts = ["The user has configured standing settings for this assistant in the add-in's Settings panel."];
  if (customization.persona) {
    parts.push(
      "PERSONA - adopt this identity, expertise, and tone for the entire conversation:\n" + customization.persona
    );
  }
  if (customization.systemInstruction) {
    parts.push(
      "SYSTEM INSTRUCTION - follow these user-level standing rules throughout the entire conversation, in " +
        "addition to (not instead of) the other standing rules in this prompt:\n" +
        customization.systemInstruction
    );
  }
  return parts.join("\n\n");
}

// Shared by buildPromptParts and buildPromptIdeasParts so the "here is the
// whole document" framing (and its file-agnostic "the document" phrasing) is
// worded identically regardless of which hidden message it rides along with.
function documentContextHiddenBlock(documentContext) {
  return (
    'Context: the user currently has a Word document open in this session. Its full text is included below - ' +
    'use it automatically whenever the user refers to "the document"/"this document" without naming a file, ' +
    "instead of asking which document they mean.\n\n" +
    "--- BEGIN CURRENT DOCUMENT CONTENT ---\n" +
    documentContext +
    "\n--- END CURRENT DOCUMENT CONTENT ---"
  );
}

// Standing rule enforcing the "always switch on Revision mode before editing"
// requirement: sent on every turn (not just ones that end up editing) since
// there is no reliable way from here to know in advance whether this turn
// will result in an edit.
const TRACK_CHANGES_INSTRUCTION =
  'Standing rule: before making any actual edit to the Word document (inserting, replacing, deleting, or ' +
  'reformatting text), first ensure Word\'s Track Changes ("Revision") mode is turned on - call the live ' +
  "toggle-track-changes tool first if it is not already enabled - so edits stay visible and reviewable rather " +
  "than being silently applied. This does not apply to read-only actions like summarizing or reading content.";

// Standing rule preventing a failure mode observed in real use: asked to "add
// the generated content to the doc", the model dumped its own markdown answer
// (pipe-delimited table rows, ** bold markers) into the document as literal
// plain text, in a font/size unrelated to the surrounding content.
const FORMATTING_INSTRUCTION =
  "Standing rule: when inserting or editing content in the Word document, produce native Word formatting - " +
  "never write markdown syntax into the document as literal text (no pipe-delimited '|' table rows, no #/##" +
  " heading markers, no **bold** asterisks, no backticks, no '-' bullet characters). Tabular data must be " +
  "inserted as a real Word table (use the add-table tool), headings via real Word heading styles, and lists " +
  "via Word's own list formatting. Before inserting, read the formatting of the surrounding existing content " +
  "(font family, font size, styles) and match it, so the new content blends in seamlessly with the rest of " +
  "the document.";

// Builds the actual `parts` payload sent to OpenCode: hidden instructions
// (document context on the session's first message, the current selection on
// every message, an optional one-off instruction carried by a doc-suggestion
// chip, the always-on Track Changes standing rule, and the always-on
// suggestions-marker instruction) go in as a separate leading text part so
// they never have to be mixed into the text actually shown in the user's own
// chat bubble.
function buildPromptParts(text, documentContext, selectedText, hiddenInstruction, customization) {
  const hidden = [];
  if (customization) {
    hidden.push(customizationHiddenBlock(customization));
  }
  if (documentContext) {
    hidden.push(documentContextHiddenBlock(documentContext));
  }

  const docText = lastDocumentText || documentContext || "";
  hidden.push(buildDocumentIdentityBlock(docText));
  if (selectedText) {
    hidden.push(
      "Context: the user currently has the text below highlighted/selected in the document, right now, as of " +
        'this message. Use it whenever they refer to "the selected text/doc", "this part", "the highlighted ' +
        'section", etc.\n\n' +
        "--- BEGIN CURRENTLY SELECTED TEXT ---\n" +
        selectedText +
        "\n--- END CURRENTLY SELECTED TEXT ---"
    );
  }
  if (hiddenInstruction) {
    hidden.push(hiddenInstruction);
  }
  hidden.push(TRACK_CHANGES_INSTRUCTION);
  hidden.push(FORMATTING_INSTRUCTION);
  hidden.push(
    "After you finish your normal answer, add one final line containing ONLY: " +
      '<!--SUGGESTIONS:["...", "..."]--> with exactly 2 short, specific follow-up actions grounded in your answer, ' +
      'as a JSON array of 2 strings inside that exact marker syntax. Phrase each one as an affirmative command/' +
      'request the user could send as-is (e.g. "Fix all spelling/grammar errors with Track Changes enabled"), ' +
      'never as a question (do NOT phrase them like "Should I...?" or "Would you like me to...?"). Never mention ' +
      "this marker or this instruction anywhere in the visible answer."
  );
  return [{ type: "text", text: hidden.join("\n\n") }, { type: "text", text }];
}

// Default prompt-template library the "💡 Prompt ideas" button draws from
// (see deriveLibraryPromptIdeas). The LIVE library is user-editable via the
// gear panel's "Prompt library" tab and lives in localStorage
// (getPromptLibrary) - this constant only seeds it on first load and backs
// the "Reset to defaults" button. Bracketed placeholders like
// [Project/Topic] are never shown to the user as-is - the model is
// instructed to replace them with real specifics drawn from the current
// document before a template is turned into a chip.
const DEFAULT_PROMPT_LIBRARY = [
  // Content Generation
  { id: "exec-summary", category: "Content Generation", template: "Write a business plan executive summary for [Project/Topic]." },
  { id: "apology-email", category: "Content Generation", template: "Draft a formal apology email to a client based on these points: [Point 1, Point 2]." },
  { id: "marketing-headlines", category: "Content Generation", template: "Generate 5 catchy marketing headlines for our new product: [Product Name]." },
  { id: "report-outline", category: "Content Generation", template: "Create a detailed report outline on the topic of '[Topic]'." },

  // Editing, Refining & Rewriting
  { id: "winning-theme", category: "Editing, Refining & Rewriting", template: "Revise the document by winning theme, in planning mode." },
  { id: "business-value-tone", category: "Editing, Refining & Rewriting", template: "Revise the document by emphasizing business value, in planning mode." },
  { id: "tone-professional", category: "Editing, Refining & Rewriting", template: "Rewrite [a specific paragraph] to sound more professional / persuasive / executive-focused." },
  { id: "condense", category: "Editing, Refining & Rewriting", template: "Condense [specific paragraphs] into a summary under 150 words." },
  { id: "expand", category: "Editing, Refining & Rewriting", template: "Expand [a specific sentence] to include more technical specifications." },
  { id: "proofread", category: "Editing, Refining & Rewriting", template: "Check the document for grammar and spelling errors and fix them with Track Changes enabled." },
  { id: "paraphrase", category: "Editing, Refining & Rewriting", template: "Rewrite [a specific section] to avoid repetitive wording." },

  // Analysis, Summarization & Extraction
  { id: "bullet-summary", category: "Analysis, Summarization & Extraction", template: "Summarize the entire document into 5 key bullet points." },
  { id: "extract-clauses", category: "Analysis, Summarization & Extraction", template: "Extract all clauses related to '[Topic, e.g. Liability, Payment Terms]' from this contract." },
  { id: "action-items", category: "Analysis, Summarization & Extraction", template: "Review the meeting minutes and list all action items along with their assigned owners." },
];

// Per-template hidden instructions, looked up locally by template id once the
// model picks that template (see deriveLibraryPromptIdeas) - NEVER stored in
// the editable library data or generated by the model itself, so neither a
// classification-style model turn nor edited library content can inject or
// alter a plan-mode behavior. Editing a template's text in the library keeps
// its id and therefore keeps this behavior; deleting the template drops it.
const HIDDEN_INSTRUCTIONS = {
  "winning-theme":
    'The user wants to revise this document based on a "winning theme" but has not told you what it is yet. ' +
    "First, ask them a clarifying question: what is the winning theme (or key differentiators) they want the " +
    "document to emphasize. Do NOT propose or make any edits in this reply. Once they answer in a follow-up " +
    "message, enter plan mode: read the document's actual sections and propose WHICH sections you would revise " +
    "and WHY, as a numbered list - but do not call any document-editing tool and do not actually change the " +
    "document until the user explicitly confirms the plan.",
  "business-value-tone":
    "The user wants to revise this document to emphasize business value (cost savings, ROI, risk reduction, " +
    "strategic impact, etc.), in planning mode. Do NOT make any edits yet - first read the document's actual " +
    "sections and propose WHICH sections you would revise and WHY, as a numbered list. Do not call any " +
    "document-editing tool and do not actually change the document until the user explicitly confirms the plan.",
};

const PROMPT_LIBRARY_STORAGE_KEY = "openCodePromptLibrary";

// The live, user-editable prompt library. Seeds localStorage with the
// defaults the first time (key absent); an explicitly saved empty library
// (user deleted every row) stays empty - it is NOT re-seeded, only the
// "Reset to defaults" button brings the built-ins back.
function getPromptLibrary() {
  const raw = localStorage.getItem(PROMPT_LIBRARY_STORAGE_KEY);
  if (raw !== null) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) {
        return parsed.filter(
          (p) => p && typeof p.id === "string" && typeof p.template === "string" && p.template.trim()
        );
      }
    } catch {
      // Corrupted JSON - fall through and re-seed the defaults.
    }
  }
  localStorage.setItem(PROMPT_LIBRARY_STORAGE_KEY, JSON.stringify(DEFAULT_PROMPT_LIBRARY));
  return DEFAULT_PROMPT_LIBRARY.slice();
}

// Marker the model is asked to reply with, and ONLY with, for the silent
// "pick from the prompt library" request triggered by the "Prompt ideas"
// button (see deriveLibraryPromptIdeas). Deliberately separate from
// SUGGESTIONS_MARKER_RE - this is a one-off request, not a normal answered
// chat turn, so it must NOT also ask for follow-up suggestions.
const PROMPT_IDEAS_MARKER_RE = /<!--\s*PROMPT_IDEAS:([\s\S]*?)-->/i;

// Builds a silent, non-conversational request asking the model to pick up to
// 4 templates from PROMPT_LIBRARY that are actually relevant to this
// document (and, since this rides on the real session, implicitly to the
// conversation so far too), deriving a concrete filled-in sentence for each.
// `documentContext` is only passed on the session's first message (isNew) -
// same gating as the whole-document injection elsewhere - since every later
// turn on this session already has the document in its own history.
function buildPromptIdeasParts(documentContext, customization, library) {
  const hidden = [];
  if (customization) {
    hidden.push(customizationHiddenBlock(customization));
  }
  if (documentContext) {
    hidden.push(documentContextHiddenBlock(documentContext));
  }
  hidden.push(
    "This is a silent, internal request to pick suggested next prompts from a fixed library - it is not shown to " +
      "the user and does not need a conversational answer. Below is a library of prompt templates, each with an " +
      "id and a template (some templates contain bracketed placeholders like [Project/Topic] or [Text] that must " +
      "be replaced, never left in literally). Library:\n" +
      JSON.stringify(library.map(({ id, category, template }) => ({ id, category, template }))) +
      "\n\nPick up to 4 of the templates most relevant to this specific document and, if applicable, the " +
      "conversation so far in this session - skip any that clearly don't fit (e.g. contract-clause extraction on " +
      "a document that isn't a contract). For each one you pick, derive a concrete, ready-to-send version of its " +
      "template by replacing any bracketed placeholders with real specifics drawn from the document (e.g. an " +
      "actual project or product name found in the text) - never leave a literal placeholder in the output. " +
      "Phrase every derived prompt as an affirmative statement/request, never as a question. If fewer than 2 " +
      "templates are clearly specific to this document, fill the remaining slots with the most generally useful " +
      "remaining LIBRARY templates (e.g. summarizing, proofreading) so at least 2 are returned whenever the " +
      "library has that many - never invent templates or ids that are not in the library. Reply with ONLY one " +
      "line, containing " +
      'exactly: <!--PROMPT_IDEAS:[{"id":"...","text":"..."}, ...]--> as a JSON array referencing each chosen ' +
      'template\'s exact "id" from the library above, with "text" being the fully derived sentence. No other text.'
  );
  return [{ type: "text", text: hidden.join("\n\n") }];
}

// Strips the (possibly still-incomplete, mid-stream) suggestions marker and
// everything after it, so it's never visible to the user even while the
// model is still typing it out token by token.
function stripSuggestionsMarker(text) {
  return text.replace(SUGGESTIONS_TAIL_RE, "");
}

// Pulls the parsed suggestions array out of a finished reply's raw text
// (marker included) plus the display-ready text with the marker removed.
function extractSuggestions(rawText) {
  const displayText = stripSuggestionsMarker(rawText).trim();
  const match = rawText.match(SUGGESTIONS_MARKER_RE);
  if (!match) {
    return { displayText, suggestions: [] };
  }
  try {
    const parsed = JSON.parse(match[1]);
    const suggestions = Array.isArray(parsed)
      ? parsed.filter((s) => typeof s === "string" && s.trim()).slice(0, 2)
      : [];
    return { displayText, suggestions };
  } catch {
    return { displayText, suggestions: [] };
  }
}

// Every user prompt actually sent this session, normalized for loose
// comparison - see isPromptAlreadyUsed(). Used to hide a suggestion chip
// once its equivalent has already been asked, so re-opening "Prompt ideas"
// later doesn't keep re-offering something already covered.
let sentPromptTexts = [];

function normalizePromptForComparison(text) {
  return text.toLowerCase().replace(/[^a-z0-9\s]/g, "").trim();
}

// "Similar" is intentionally loose (normalized substring match in either
// direction) rather than exact-string or fuzzy-distance matching - good
// enough to catch "Summarize the document" vs. "Summarize the document for
// me please" without pulling in a similarity-scoring library for what's
// fundamentally a small, fixed set of candidate chips.
function isPromptAlreadyUsed(promptText) {
  const normalized = normalizePromptForComparison(promptText);
  return sentPromptTexts.some((sent) => sent === normalized || sent.includes(normalized) || normalized.includes(sent));
}

// Decides whether the "💡 Prompt ideas" toggle is shown at all: a
// brand-new/empty document has nothing worth suggesting prompts about, so
// neither the toggle nor the panel it opens ever appears. This is a cheap,
// synchronous-feeling check (no LLM call) - the actual library lookup only
// happens later, on demand, when the button is clicked (see
// deriveLibraryPromptIdeas/togglePromptIdeas below).
async function initializeDocumentSuggestions() {
  const documentContext = await getActiveDocumentText();
  if (!documentContext || !documentContext.trim()) {
    return;
  }
  document.getElementById("prompt-ideas-toggle").style.display = "flex";
  document.getElementById("prompt-ideas-toolbar").style.display = "flex";
}

// Called fresh every time the "Prompt ideas" panel is opened (see
// togglePromptIdeas) so its picks reflect the document's current content and
// the conversation so far, not a stale snapshot from taskpane load. Sends a
// silent request (see buildPromptIdeasParts) asking the model to choose up
// to 4 templates from the live, user-editable prompt library
// (getPromptLibrary - re-read on every click, so library edits apply from
// the next click with no session reset) and derive a concrete, filled-in
// sentence for each; each returned template id is looked up locally so a
// chip's hiddenInstruction (e.g. the "winning theme" plan-mode flow) always
// comes from our own fixed library, never from the model's own output.
async function deriveLibraryPromptIdeas() {
  const library = getPromptLibrary();
  if (library.length === 0) {
    // Nothing to pick from - skip the LLM round trip entirely.
    // togglePromptIdeas pre-checks this case to show a specific message.
    return [];
  }
  const { id: sid, isNew } = await ensureSession();
  if (!lastDocumentText || !lastDocumentText.trim()) {
    return [];
  }

  // Network/timeout failures (e.g. opencode unreachable, or the model taking
  // longer than the timeout on a large document) are intentionally left to
  // propagate to togglePromptIdeas, which shows them as a visible error
  // instead of the panel silently going blank with no explanation.
  const raw = await sendToOpenCodeBlocking(
    sid,
    buildPromptIdeasParts(isNew ? lastDocumentText : null, isNew ? getSavedCustomization() : null, library),
    150000
  );
  const match = raw.match(PROMPT_IDEAS_MARKER_RE);
  if (!match) {
    return [];
  }
  // Parse leniently: models occasionally wrap the requested JSON array in
  // an extra stray brace/quote despite the instruction, so if match[1]
  // isn't valid JSON on its own, fall back to extracting just the
  // outermost [...] substring rather than discarding an otherwise-good
  // response over one stray character.
  let parsed;
  try {
    parsed = JSON.parse(match[1]);
  } catch {
    const arrayMatch = match[1].match(/\[[\s\S]*\]/);
    if (!arrayMatch) {
      return [];
    }
    parsed = JSON.parse(arrayMatch[0]);
  }
  if (!Array.isArray(parsed)) {
    return [];
  }
  const chips = [];
  for (const item of parsed) {
    if (!item || typeof item.id !== "string" || typeof item.text !== "string" || !item.text.trim()) {
      continue;
    }
    const template = library.find((t) => t.id === item.id);
    if (!template) {
      continue;
    }
    chips.push({ label: item.text.trim(), text: item.text.trim(), hiddenInstruction: HIDDEN_INSTRUCTIONS[item.id] || null });
  }
  return chips.slice(0, 4);
}

// Renders a set of prompt-idea chips into #doc-suggestions, reusing the same
// visual chip style as the post-reply follow-up suggestions
// (.chat-suggestion-chip). Each chip carries its own literal message text
// plus an optional hidden instruction (see pendingHiddenInstruction). Chips
// whose prompt has effectively already been sent (isPromptAlreadyUsed) are
// left out - if that empties the list entirely, the panel is hidden rather
// than shown blank.
function renderDocSuggestions(chips) {
  const panel = document.getElementById("doc-suggestions-panel");
  const container = document.getElementById("doc-suggestions");

  const remaining = chips.filter((chip) => !isPromptAlreadyUsed(chip.text));
  if (remaining.length === 0) {
    // Never collapse silently - a click with no visible response reads as
    // "the button is broken", not "there was nothing new to offer".
    container.innerHTML = "";
    const note = document.createElement("p");
    note.className = "doc-suggestions-loading";
    note.textContent =
      chips.length === 0
        ? 'No prompt ideas came back for this document. Click "Prompt ideas" to try again.'
        : "Every idea for this document has already been asked in this conversation - continue the chat and try again later.";
    container.appendChild(note);
    panel.style.display = "block";
    container.style.display = "flex";
    return;
  }

  container.innerHTML = "";
  for (const chip of remaining) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "chat-suggestion-chip";
    btn.textContent = chip.label;
    btn.addEventListener("click", () => {
      // Stage the chip's prompt into the input for the user to review/edit
      // (e.g. tweak specifics) and send themselves, rather than firing it off
      // immediately. Its hidden instruction rides along with the staged text
      // and is dropped if the user wipes it (see onChatInputChanged).
      pendingHiddenInstruction = chip.hiddenInstruction || null;
      const input = document.getElementById("chat-input");
      input.value = chip.text;
      onChatInputChanged();
      input.focus();
    });
    container.appendChild(btn);
  }
  panel.style.display = "block";
  container.style.display = "flex";
}

// Guards against a second overlapping fetch if the toggle is double-clicked
// while the previous library lookup is still in flight.
let promptIdeasLoading = false;

// Wired to the "💡 Prompt ideas" toggle button. Collapses the panel if it's
// currently open; if it's closed, shows a brief loading state and fetches a
// fresh, context-aware set of ideas from the prompt library (rather than
// re-showing whatever was rendered last time), then renders it.
async function togglePromptIdeas() {
  const panel = document.getElementById("doc-suggestions-panel");
  if (panel.style.display !== "none") {
    panel.style.display = "none";
    return;
  }
  if (promptIdeasLoading) {
    return;
  }

  const toggleBtn = document.getElementById("prompt-ideas-toggle");
  const container = document.getElementById("doc-suggestions");

  // Pre-check: with an emptied-out prompt library there is nothing the model
  // could pick from, so say that specifically instead of a generic "no ideas".
  if (getPromptLibrary().length === 0) {
    container.innerHTML = "";
    const emptyEl = document.createElement("p");
    emptyEl.className = "doc-suggestions-loading";
    emptyEl.textContent = 'Prompt library is empty - add prompts via the gear icon\'s "Prompt library" tab.';
    container.appendChild(emptyEl);
    container.style.display = "flex";
    panel.style.display = "block";
    return;
  }

  promptIdeasLoading = true;
  toggleBtn.disabled = true;
  container.innerHTML = "";
  const loadingEl = document.createElement("p");
  loadingEl.className = "doc-suggestions-loading";
  loadingEl.textContent = "Thinking of ideas... (can take up to a couple minutes on longer documents)";
  container.appendChild(loadingEl);
  panel.style.display = "block";

  try {
    const chips = await deriveLibraryPromptIdeas();
    renderDocSuggestions(chips);
  } catch (err) {
    container.innerHTML = "";
    const errorEl = document.createElement("p");
    errorEl.className = "doc-suggestions-loading";
    errorEl.textContent = `Couldn't get prompt ideas (${err.message}). Click "Prompt ideas" to try again.`;
    container.appendChild(errorEl);
  } finally {
    promptIdeasLoading = false;
    toggleBtn.disabled = false;
  }
}

// --- Streaming (SSE) reply rendering ------------------------------------
//
// opencode's REST API has two ways to get an assistant reply: the blocking
// POST /session/{id}/message used previously (waits for full generation,
// then returns everything at once - the "displays after a long wait" problem
// this was written to fix), and POST /session/{id}/prompt_async (returns 204
// immediately) paired with GET /event, a single global Server-Sent-Events
// stream carrying every incremental step of the reply: reasoning-token
// deltas, answer-token deltas, tool-call state changes, then a final
// message.updated with `finish`/`error` set. This lets the bubble render
// progressively (typewriter-style) and show a live "thinking"/tool-activity
// trace, instead of a blank "..." bubble sitting still for however long the
// model takes to finish.

let eventSource = null;
let eventStreamOpenPromise = null;
// The one reply in flight at a time (the send button is disabled while
// pending, so there is never more than one) - holds everything needed to
// turn opencode's part/delta events into progressive DOM updates.
let activeReply = null;
// Every assistant message ID this client has already attributed to some
// reply (its own past replies, plus whatever existed on the session before
// this reply started - see seedSeenAssistantMessageIDs). opencode
// re-broadcasts stale message.updated events around session-idle/busy
// transitions, AND a session can have a turn still generating in the
// background from before a page reload (prompt_async is fire-and-forget, so
// the server keeps working even if the client that sent it goes away). If a
// fresh reply's first-sighted "assistant" message.updated turns out to
// belong to one of these leftover messages, it would be mistaken for the new
// reply's own message - and since it may already carry `finish`, the reply
// would be finalized instantly with zero collected text ("OpenCode finished
// without returning any text."). Checking IDs against this set lets a fresh
// reply ignore any message it didn't itself create and keep waiting for its
// actual one.
let seenAssistantMessageIDs = new Set();

// Populates seenAssistantMessageIDs with every assistant message ID that
// already exists in this session BEFORE a new prompt is sent, so a reload or
// reconnect that races with a still-in-flight previous turn can't mistake
// that turn's tail end for this reply's own message. Best-effort: if this
// fails, the reply still works for the common (non-racing) case.
async function seedSeenAssistantMessageIDs(sid) {
  try {
    const res = await fetch(`${OPENCODE_BASE_URL}/session/${sid}/message`);
    if (!res.ok) {
      return;
    }
    const messages = await res.json();
    for (const m of messages) {
      if (m.info && m.info.role === "assistant" && m.info.id) {
        seenAssistantMessageIDs.add(m.info.id);
      }
    }
  } catch {
    // Best-effort - see comment above.
  }
}

// Opens the SSE connection once and reuses it for the taskpane's lifetime.
// Sending prompt_async before this connection reaches OPEN would risk losing
// the first events of the reply (the connection wouldn't exist yet to
// receive them), so callers must await this before posting a message.
function connectEventStream() {
  if (eventStreamOpenPromise) {
    return eventStreamOpenPromise;
  }
  eventStreamOpenPromise = new Promise((resolve) => {
    let settled = false;
    const settle = (ok) => {
      if (!settled) {
        settled = true;
        resolve(ok);
      }
    };
    try {
      eventSource = new EventSource(`${OPENCODE_BASE_URL}/event`);
    } catch {
      settle(false);
      return;
    }
    eventSource.onopen = () => settle(true);
    // A later error (e.g. a transient reconnect) after the stream already
    // opened must NOT flip an already-resolved promise back to false - only
    // the first open/error decides whether streaming is usable this session.
    eventSource.onerror = () => settle(false);
    eventSource.onmessage = onServerEvent;
    setTimeout(() => settle(eventSource.readyState === EventSource.OPEN), 3000);
  });
  return eventStreamOpenPromise;
}

function onServerEvent(evt) {
  if (!activeReply) {
    return;
  }
  let event;
  try {
    event = JSON.parse(evt.data);
  } catch {
    return;
  }
  const props = event.properties || {};
  if (props.sessionID && props.sessionID !== activeReply.sessionID) {
    return;
  }

  switch (event.type) {
    case "message.updated": {
      const info = props.info;
      if (!info || info.role !== "assistant") {
        return;
      }
      if (!activeReply.assistantMessageID) {
        if (seenAssistantMessageIDs.has(info.id)) {
          // Leftover/stale message this client didn't create for this reply
          // (see seenAssistantMessageIDs) - ignore it and keep waiting for
          // the genuinely new one.
          return;
        }
        // First sight of the assistant message this prompt created - lock onto
        // its ID so later part/delta events (which don't carry a role) can be
        // told apart from the user message's own echoed-back text part.
        activeReply.assistantMessageID = info.id;
        seenAssistantMessageIDs.add(info.id);
      } else if (info.id !== activeReply.assistantMessageID) {
        if (activeReply.finished) {
          return;
        }
        // opencode gives each STEP of a multi-step tool-calling turn its own
        // message ID (reasoning -> tool call -> more reasoning/text all show up
        // as separate "assistant" messages for one user prompt). The previous
        // message's `finish: "tool-calls"` was just a step boundary, not the
        // end of the reply - follow along to the new step so its parts (e.g.
        // the actual final answer text) aren't silently dropped.
        activeReply.assistantMessageID = info.id;
        seenAssistantMessageIDs.add(info.id);
      }
      if (info.error) {
        activeReply.finished = true;
        finishActiveReply(info.error);
        return;
      }
      // "tool-calls" means this step ended because the model called a tool,
      // not that the turn is over - more messages/parts (the real answer)
      // are still coming, so don't finalize the reply yet.
      if (info.finish && info.finish !== "tool-calls") {
        activeReply.finished = true;
        finishActiveReply(null);
      }
      break;
    }
    case "message.part.updated": {
      const part = props.part;
      if (!part || !activeReply.assistantMessageID || part.messageID !== activeReply.assistantMessageID) {
        return;
      }
      onPartUpdated(part);
      break;
    }
    case "message.part.delta": {
      if (
        !activeReply.assistantMessageID ||
        props.messageID !== activeReply.assistantMessageID ||
        props.field !== "text"
      ) {
        return;
      }
      onPartDelta(props.partID, props.delta);
      break;
    }
    default:
      break;
  }
}

function onPartUpdated(part) {
  if (part.type === "reasoning") {
    // Accumulate EVERY reasoning segment (not just the first) so intermediate
    // reasoning between tool calls stays visible - see renderThought.
    if (!activeReply.reasoningParts.has(part.id)) {
      activeReply.reasoningOrder.push(part.id);
    }
    activeReply.reasoningParts.set(part.id, part.text || "");
    renderThought(activeReply);
    if (part.time && part.time.end && !activeReply.dom.thoughtEl.classList.contains("chat-thought--done")) {
      finishThought(activeReply, part.time.end - part.time.start);
    }
  } else if (part.type === "text") {
    if (activeReply.reasoningParts.size === 0 && !activeReply.dom.thoughtEl.classList.contains("chat-thought--done")) {
      // Some models skip the reasoning phase entirely and go straight to the
      // answer - collapse the "Thinking..." row using wall-clock elapsed time
      // instead of leaving it spinning forever with nothing to finish it.
      finishThought(activeReply, Date.now() - activeReply.startedAt);
    }
    if (!activeReply.textOrder.includes(part.id)) {
      activeReply.textOrder.push(part.id);
    }
    activeReply.textParts.set(part.id, part.text || "");
    renderAnswerText(activeReply);
    setActivity(activeReply, null);
  } else if (part.type === "tool") {
    onToolPartUpdated(activeReply, part);
  }
}

function onPartDelta(partID, delta) {
  if (activeReply.reasoningParts.has(partID)) {
    activeReply.reasoningParts.set(partID, activeReply.reasoningParts.get(partID) + delta);
    renderThought(activeReply);
  } else if (activeReply.textParts.has(partID)) {
    if (activeReply.reasoningParts.size === 0 && !activeReply.dom.thoughtEl.classList.contains("chat-thought--done")) {
      finishThought(activeReply, Date.now() - activeReply.startedAt);
    }
    activeReply.textParts.set(partID, activeReply.textParts.get(partID) + delta);
    renderAnswerText(activeReply);
    setActivity(activeReply, null);
  }
}

// Prettifies a raw tool id for display, e.g.
// "word_word_live_toggle_track_changes" -> "toggle track changes",
// "read" -> "Read". Strips the word-mcp prefixes and turns underscores into
// spaces so the step list reads like opencode's.
function friendlyToolName(tool) {
  if (!tool) {
    return "tool";
  }
  let name = String(tool).replace(/^word_word_live_/, "").replace(/^word_/, "");
  name = name.replace(/_/g, " ").trim();
  // Capitalize the common built-in tools (read/glob/bash/write/edit/list/grep).
  if (/^(read|glob|bash|write|edit|list|grep|webfetch|task)$/i.test(name)) {
    name = name.charAt(0).toUpperCase() + name.slice(1);
  }
  return name || String(tool);
}

// A short one-line summary of a tool call's arguments. Drops `filename` first -
// every word-mcp tool carries the open document's name, so it's the same noise
// on every step and hides the actually-useful args. Then prefers a single
// salient field (the file/command/pattern/query or the content being written),
// else a compact JSON of what remains. Returns "" when nothing useful is left.
function toolArgsSummary(input) {
  if (!input || typeof input !== "object") {
    return "";
  }
  const filtered = {};
  for (const [k, v] of Object.entries(input)) {
    if (k === "filename") {
      continue;
    }
    filtered[k] = v;
  }
  const salient =
    filtered.filePath ||
    filtered.path ||
    filtered.file ||
    filtered.command ||
    filtered.pattern ||
    filtered.query ||
    filtered.url ||
    filtered.heading ||
    filtered.text ||
    filtered.new_text ||
    filtered.search_text ||
    filtered.content ||
    filtered.description;
  let text;
  if (salient != null && String(salient) !== "") {
    text = String(salient);
  } else if (Object.keys(filtered).length === 0) {
    text = "";
  } else {
    text = JSON.stringify(filtered);
  }
  if (text.length > 100) {
    text = text.slice(0, 99) + "…";
  }
  return text;
}

// A single, transient "working" line for the tool currently running - it
// updates as each tool starts and clears once the tool finishes, so the
// intermediate tool steps aren't listed out persistently (per user
// preference). Reasoning is still shown above; the answer follows below.
function onToolPartUpdated(reply, part) {
  const status = (part.state && part.state.status) || "running";
  if (status === "running" || status === "pending") {
    const name = friendlyToolName(part.tool);
    const args = toolArgsSummary(part.state && part.state.input);
    setActivity(reply, args ? `${name}: ${args}` : `${name}...`);
  } else {
    setActivity(reply, null);
  }
}

// Sends one message via the streaming path, falling back to the old blocking
// round trip if the SSE stream never came up (e.g. CORS misconfigured for
// /event specifically, or an opencode version predating prompt_async/event).
async function streamAssistantReply(sid, text, documentContext, selectedText, hiddenInstruction, customization) {
  const streamOk = await connectEventStream();
  if (!streamOk) {
    const replyText = await sendToOpenCodeBlocking(
      sid,
      buildPromptParts(text, documentContext, selectedText, hiddenInstruction, customization)
    );
    appendMessage("assistant", stripSuggestionsMarker(replyText).trim());
    return;
  }

  // Kicked off now (in parallel with building the pending bubble below) but
  // awaited just before the prompt is actually sent, so the set is guaranteed
  // populated before the server can broadcast any event this reply might see
  // - without delaying the pending bubble's instant appearance.
  const seedPromise = seedSeenAssistantMessageIDs(sid);

  const dom = appendStreamingAssistantMessage();
  const reply = {
    sessionID: sid,
    assistantMessageID: null,
    reasoningParts: new Map(),
    reasoningOrder: [],
    textOrder: [],
    textParts: new Map(),
    dom,
    startedAt: Date.now(),
    tickInterval: null,
    finished: false,
  };
  const done = new Promise((resolve) => {
    reply.settle = resolve;
  });
  activeReply = reply;
  showThought(reply);
  reply.tickInterval = setInterval(() => tickThought(reply), 1000);

  await seedPromise;

  let res;
  try {
    res = await fetch(`${OPENCODE_BASE_URL}/session/${sid}/prompt_async`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: OPENCODE_MODEL,
        parts: buildPromptParts(text, documentContext, selectedText, hiddenInstruction, customization),
      }),
    });
  } catch (err) {
    activeReply = null;
    dom.el.remove();
    throw err;
  }
  if (!res.ok) {
    activeReply = null;
    dom.el.remove();
    throw new Error(`message send failed: HTTP ${res.status}`);
  }

  await done;
}

// Invoked by the send button's stop mode (see onSubmit): best-effort asks the
// server to abort the in-flight generation, then finalizes the reply locally
// right away with whatever text has streamed so far, instead of waiting for
// the server to (maybe) emit a final event. Any late events for the aborted
// message are ignored - activeReply is already null by then, and its message
// IDs are in seenAssistantMessageIDs, so the next turn can't mistake them
// for its own.
function stopActiveRun() {
  const reply = activeReply;
  if (!reply) {
    // Nothing to stop client-side (e.g. the rare blocking-fallback path,
    // which has no incremental state to finalize) - the run ends on its own.
    return;
  }
  fetch(`${OPENCODE_BASE_URL}/session/${reply.sessionID}/abort`, { method: "POST" }).catch(() => {});
  reply.stopped = true;
  reply.finished = true;
  finishActiveReply(null);
}

function finishActiveReply(error) {
  const reply = activeReply;
  if (!reply) {
    return;
  }
  activeReply = null;
  clearInterval(reply.tickInterval);

  reply.dom.el.classList.remove("chat-message--pending");
  setActivity(reply, null);
  if (reply.dom.thoughtEl.style.display !== "none" && !reply.dom.thoughtEl.classList.contains("chat-thought--done")) {
    // Reasoning never got an explicit "ended" part update (observed on
    // aborted/errored replies) - collapse it anyway so the bubble doesn't
    // look permanently stuck mid-thought.
    reply.dom.thoughtLabelEl.textContent = "Thought";
    reply.dom.thoughtEl.classList.add("chat-thought--done");
  }

  const combined = reply.textOrder
    .map((id) => reply.textParts.get(id))
    .join("")
    .trim();
  const { displayText, suggestions } = extractSuggestions(combined);
  if (error) {
    if (!displayText) {
      reply.dom.el.remove();
    }
    appendMessage("error", describeOpenCodeError(error));
  } else if (!displayText) {
    reply.dom.el.remove();
    // A user-initiated stop before any text arrived needs no error bubble -
    // the user knows why the reply vanished; only an unexpected empty finish
    // deserves one.
    if (!reply.stopped) {
      appendMessage("error", "OpenCode finished without returning any text.");
    }
  } else {
    wireMessageActions(reply.dom, displayText);
    renderSuggestions(reply.dom, suggestions);
  }
  reply.settle();
}

function describeOpenCodeError(error) {
  const message = error && error.data && error.data.message;
  return message ? `${error.name}: ${message}` : error.name || "OpenCode returned an error.";
}

async function sendToOpenCodeBlocking(sid, parts, timeoutMs = 120000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  let res;
  try {
    res = await fetch(`${OPENCODE_BASE_URL}/session/${sid}/message`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: OPENCODE_MODEL,
        parts,
      }),
      signal: controller.signal,
    });
  } catch (err) {
    if (err.name === "AbortError") {
      throw new Error(`message send timed out after ${Math.round(timeoutMs / 1000)}s`);
    }
    throw err;
  } finally {
    clearTimeout(timer);
  }
  if (!res.ok) {
    throw new Error(`message send failed: HTTP ${res.status}`);
  }
  const body = await res.json();
  return body.parts
    .filter((part) => part.type === "text")
    .map((part) => part.text)
    .join("")
    .trim();
}

// --- DOM rendering --------------------------------------------------------

function appendMessage(role, text, { pending = false } = {}) {
  const log = document.getElementById("chat-log");
  const el = document.createElement("div");
  el.className = `chat-message chat-message--${role}${pending ? " chat-message--pending" : ""}`;
  el.innerHTML = renderMarkdown(text);
  log.appendChild(el);
  scrollChatLog();
  return el;
}

// Builds an assistant bubble with three progressively-filled regions: a
// collapsible reasoning ("chain of thought") trace, a one-line activity
// indicator for whatever's currently running (thinking / a tool call), and
// the answer text itself. Returns element references so the SSE handlers
// above can update them live as events arrive.
function appendStreamingAssistantMessage() {
  const log = document.getElementById("chat-log");
  const el = document.createElement("div");
  el.className = "chat-message chat-message--assistant chat-message--pending";

  const thoughtEl = document.createElement("div");
  thoughtEl.className = "chat-thought";
  thoughtEl.style.display = "none";
  const thoughtSummaryEl = document.createElement("button");
  thoughtSummaryEl.type = "button";
  thoughtSummaryEl.className = "chat-thought-summary";
  thoughtSummaryEl.innerHTML =
    '<span class="chat-thought-dot">◆</span><span class="chat-thought-label">Thinking...</span>';
  const thoughtDetailEl = document.createElement("div");
  thoughtDetailEl.className = "chat-thought-detail";
  thoughtSummaryEl.addEventListener("click", () => {
    thoughtDetailEl.classList.toggle("chat-thought-detail--open");
  });
  thoughtEl.append(thoughtSummaryEl, thoughtDetailEl);

  const activityEl = document.createElement("div");
  activityEl.className = "chat-activity";
  activityEl.style.display = "none";
  activityEl.innerHTML = '<span class="chat-activity-icon">◆</span><span class="chat-activity-label"></span>';

  const textEl = document.createElement("div");
  textEl.className = "chat-text";

  const actionsEl = document.createElement("div");
  actionsEl.className = "chat-actions";
  actionsEl.style.display = "none";
  actionsEl.innerHTML =
    '<button type="button" class="chat-action-btn chat-action-copy" title="Copy">⧉</button>' +
    '<button type="button" class="chat-action-btn chat-action-up" title="Good response" aria-pressed="false">👍</button>' +
    '<button type="button" class="chat-action-btn chat-action-down" title="Bad response" aria-pressed="false">👎</button>';

  const suggestionsEl = document.createElement("div");
  suggestionsEl.className = "chat-suggestions";
  suggestionsEl.style.display = "none";

  el.append(thoughtEl, activityEl, textEl, actionsEl, suggestionsEl);
  log.appendChild(el);
  scrollChatLog();

  return {
    el,
    thoughtEl,
    thoughtLabelEl: thoughtSummaryEl.querySelector(".chat-thought-label"),
    thoughtDetailEl,
    activityEl,
    activityLabelEl: activityEl.querySelector(".chat-activity-label"),
    textEl,
    actionsEl,
    suggestionsEl,
  };
}

// Reveals the Copy/👍/👎 row under a finished reply and wires it up. Feedback
// is local-only (no telemetry backend exists for this single-user tool) - the
// thumbs buttons just reflect the user's own click back as a toggle.
function wireMessageActions(dom, text) {
  dom.actionsEl.style.display = "flex";
  const copyBtn = dom.actionsEl.querySelector(".chat-action-copy");
  const upBtn = dom.actionsEl.querySelector(".chat-action-up");
  const downBtn = dom.actionsEl.querySelector(".chat-action-down");

  copyBtn.addEventListener("click", async () => {
    try {
      await navigator.clipboard.writeText(text);
      copyBtn.textContent = "✓";
      setTimeout(() => (copyBtn.textContent = "⧉"), 1200);
    } catch {
      // Clipboard permission can be denied by the host - nothing to fall back to.
    }
  });

  const togglePressed = (btn, other) => {
    const pressed = btn.getAttribute("aria-pressed") === "true";
    btn.setAttribute("aria-pressed", String(!pressed));
    other.setAttribute("aria-pressed", "false");
  };
  upBtn.addEventListener("click", () => togglePressed(upBtn, downBtn));
  downBtn.addEventListener("click", () => togglePressed(downBtn, upBtn));
}

function renderSuggestions(dom, suggestions) {
  if (suggestions.length === 0) {
    return;
  }
  dom.suggestionsEl.style.display = "flex";
  dom.suggestionsEl.innerHTML = "";
  for (const suggestion of suggestions) {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "chat-suggestion-chip";
    chip.textContent = suggestion;
    chip.addEventListener("click", () => {
      // Same fill-then-edit behavior as the prompt-ideas chips: stage the
      // suggestion in the input for review rather than sending immediately.
      const input = document.getElementById("chat-input");
      input.value = suggestion;
      onChatInputChanged();
      input.focus();
    });
    dom.suggestionsEl.appendChild(chip);
  }
  scrollChatLog();
}

function showThought(reply) {
  reply.dom.thoughtEl.style.display = "block";
  reply.currentWord = pickThinkingWord();
  reply.dom.thoughtLabelEl.textContent = `${reply.currentWord}...`;
}

function tickThought(reply) {
  if (reply.dom.thoughtEl.classList.contains("chat-thought--done")) {
    clearInterval(reply.tickInterval);
    return;
  }
  const seconds = Math.max(1, Math.round((Date.now() - reply.startedAt) / 1000));
  // Rotate to a new status word every 3s so it reads as "still working", not stuck.
  if (seconds % 3 === 0 && seconds !== reply.lastWordChangeSecond) {
    reply.lastWordChangeSecond = seconds;
    reply.currentWord = pickThinkingWord(reply.currentWord);
  }
  reply.dom.thoughtLabelEl.textContent = `${reply.currentWord}... ${seconds}s`;
}

function renderThought(reply) {
  reply.dom.thoughtDetailEl.classList.add("chat-thought-detail--open");
  // Concatenate all reasoning segments (in arrival order) so reasoning that
  // happens between tool calls is shown too, not just the first burst.
  const combined = reply.reasoningOrder.map((id) => reply.reasoningParts.get(id) || "").join("\n\n");
  reply.dom.thoughtDetailEl.innerHTML = renderMarkdown(combined);
  scrollChatLog();
}

function finishThought(reply, elapsedMs) {
  clearInterval(reply.tickInterval);
  const seconds = Math.max(1, Math.round(elapsedMs / 1000));
  reply.dom.thoughtLabelEl.textContent = `Thought for ${seconds}s`;
  reply.dom.thoughtEl.classList.add("chat-thought--done");
  // Leave the reasoning detail expanded so the intermediate reasoning stays
  // visible after the turn finishes (the user can still click "Thought for Ns"
  // to collapse it). If the turn had no reasoning, --open was never added, so
  // this stays collapsed automatically.
}

function setActivity(reply, label) {
  if (!label) {
    reply.dom.activityEl.style.display = "none";
    return;
  }
  reply.dom.activityEl.style.display = "flex";
  reply.dom.activityLabelEl.textContent = label;
  scrollChatLog();
}

function renderAnswerText(reply) {
  const combined = reply.textOrder.map((id) => reply.textParts.get(id)).join("");
  reply.dom.textEl.innerHTML = renderMarkdown(stripSuggestionsMarker(combined));
  scrollChatLog();
}

function scrollChatLog() {
  const log = document.getElementById("chat-log");
  log.scrollTop = log.scrollHeight;
}

// Escapes HTML entities first so any literal markup in the model's (or user's)
// text is displayed as text rather than parsed as DOM elements, then applies
// markdown formatting on top of the escaped text - keeps **bold**, lists, code
// spans, etc. working safely. The custom link/image renderer registered above
// additionally blocks dangerous URL schemes (e.g. javascript:), which entity
// escaping alone doesn't cover.
function renderMarkdown(text) {
  return marked.parse(escapeHtml(text));
}

function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}
