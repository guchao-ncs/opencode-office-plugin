// Phase 9 test: drives the real taskpane chat UI's settings panel (gear icon) in
// headless Edge, verifying the "generate harness config" flow end-to-end. This does
// NOT round-trip through opencode serve - the gear icon is a pure client-side helper
// (see the note in taskpane.js's generateHarnessSnippet for why: a live PATCH /config
// call was tried and found not to persist anything, confirmed by hand against a real
// opencode serve instance, so the feature was redesigned to generate a copy/paste
// snippet for opencode.json instead of trying to apply it at runtime).
//
// Assumes the caller (phase9-harness-config.ps1) already has the webpack dev server
// (https://localhost:3000) running.
import { chromium } from "playwright-core";

const EDGE_PATH = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";
const HARNESS_PATH = "C:\\AI-Workspace\\projects\\NCS-SMART-SOLUTION";

async function main() {
  const browser = await chromium.launch({
    executablePath: EDGE_PATH,
    args: ["--ignore-certificate-errors"],
  });
  const page = await browser.newPage();
  const consoleErrors = [];
  page.on("console", (msg) => {
    if (msg.type() === "error") consoleErrors.push(msg.text());
  });
  page.on("pageerror", (err) => consoleErrors.push(`pageerror: ${err.message}`));

  // Must also stub Office.context.document.settings - taskpane.js calls
  // Office.context.document.settings.set(...)/.saveAsync() (Phase 8's autoopen tagging)
  // synchronously at the top of the Office.onReady callback, before the localStorage
  // refill code below it. Without this, that line throws ("Cannot read properties of
  // undefined (reading 'document')"), which aborts the rest of the callback and made
  // the localStorage refill assertion below fail for an unrelated reason.
  await page.route("https://appsforoffice.microsoft.com/lib/1/hosted/office.js", (route) => {
    route.fulfill({
      contentType: "application/javascript",
      body: `window.Office = {
        onReady: (cb) => Promise.resolve().then(() => cb({ host: 'Word' })),
        HostType: { Word: 'Word' },
        context: { document: { settings: { set: () => {}, saveAsync: () => {} } } },
      };`,
    });
  });

  // "networkidle" would never fire: taskpane.js now opens a persistent SSE
  // connection to /event on Office.onReady for streaming replies (and, with no
  // opencode serve running in this test, keeps auto-retrying it), so the page
  // never goes network-idle. Wait for a concrete DOM readiness signal instead.
  await page.goto("https://localhost:3000/taskpane.html", { waitUntil: "load" });
  await page.waitForSelector("#settings-toggle", { state: "visible" });

  const failures = [];

  // Settings panel starts hidden, opens on gear click.
  const panelDisplayBefore = await page.$eval("#settings-panel", (el) => getComputedStyle(el).display);
  if (panelDisplayBefore !== "none") {
    failures.push(`settings panel should start hidden, but computed display was "${panelDisplayBefore}"`);
  }
  await page.click("#settings-toggle");
  const panelDisplayAfter = await page.$eval("#settings-panel", (el) => getComputedStyle(el).display);
  if (panelDisplayAfter === "none") {
    failures.push("settings panel did not become visible after clicking #settings-toggle");
  }

  // Type a harness path and generate the config snippet.
  await page.fill("#harness-root-input", HARNESS_PATH);
  await page.click("#settings-save");

  const snippetVisible = await page.$eval("#settings-snippet", (el) => getComputedStyle(el).display !== "none");
  if (!snippetVisible) {
    failures.push("#settings-snippet did not become visible after clicking Generate config");
  }
  const snippetText = await page.$eval("#settings-snippet", (el) => el.value);
  let snippetJson = null;
  try {
    snippetJson = JSON.parse(snippetText);
  } catch (err) {
    failures.push(`generated snippet is not valid JSON: ${err.message} (raw: ${snippetText})`);
  }
  if (snippetJson) {
    const expectedInstruction = `${HARNESS_PATH}\\AGENTS.md`;
    if (!Array.isArray(snippetJson.instructions) || !snippetJson.instructions.includes(expectedInstruction)) {
      failures.push(`snippet.instructions missing "${expectedInstruction}": ${JSON.stringify(snippetJson.instructions)}`);
    }
    const expectedPermissionKey = `${HARNESS_PATH}\\**`;
    const permissionValue = snippetJson.permission?.external_directory?.[expectedPermissionKey];
    if (permissionValue !== "allow") {
      failures.push(
        `snippet.permission.external_directory["${expectedPermissionKey}"] is "${permissionValue}", expected "allow"`
      );
    }
  }

  // Reload the page to verify the path was persisted to localStorage and re-filled.
  await page.reload({ waitUntil: "load" });
  await page.waitForSelector("#settings-toggle", { state: "visible" });
  await page.click("#settings-toggle");
  const refilledValue = await page.$eval("#harness-root-input", (el) => el.value);
  if (refilledValue !== HARNESS_PATH) {
    failures.push(`harness-root-input was not re-filled from localStorage after reload: got "${refilledValue}"`);
  }

  // Clear should wipe the input, hide the snippet, and clear localStorage.
  await page.click("#settings-clear");
  const clearedInputValue = await page.$eval("#harness-root-input", (el) => el.value);
  if (clearedInputValue !== "") {
    failures.push(`harness-root-input was not cleared: got "${clearedInputValue}"`);
  }
  const snippetVisibleAfterClear = await page.$eval("#settings-snippet", (el) => getComputedStyle(el).display !== "none");
  if (snippetVisibleAfterClear) {
    failures.push("#settings-snippet is still visible after clicking Clear");
  }
  const storedAfterClear = await page.evaluate(() => localStorage.getItem("openCodeHarnessRoot"));
  if (storedAfterClear !== null) {
    failures.push(`localStorage still has openCodeHarnessRoot after Clear: "${storedAfterClear}"`);
  }

  if (consoleErrors.length > 0) {
    failures.push(`browser console errors: ${JSON.stringify(consoleErrors)}`);
  }

  await browser.close();

  if (failures.length > 0) {
    console.error("FAIL:\n" + failures.map((f) => `  - ${f}`).join("\n"));
    process.exit(1);
  }
  console.log("PASS: harness config settings panel (gear icon, generate snippet, persist, clear) works");
}

main().catch((err) => {
  console.error("FAIL: script error:", err.message);
  process.exit(1);
});
