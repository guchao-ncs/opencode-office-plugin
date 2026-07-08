// Phase 9 test: drives the real taskpane chat UI's Settings panel (gear icon) in
// headless Edge. The panel maintains two localStorage-backed fields - System
// Instruction and Persona - that taskpane.js injects as a hidden instruction on
// each new session's first message (customizationHiddenBlock). This test covers
// the client-side surface: open/close, save/persist/re-fill, clear, and the
// one-time migration of the legacy "harness root directory" setting.
//
// No opencode serve needed - the panel is pure client-side; the injection path
// itself rides the same message-sending flow phase5 already covers.
//
// Assumes the caller (phase9-settings-panel.ps1) already has the webpack dev
// server (https://localhost:3000) running.
import { chromium } from "playwright-core";

const EDGE_PATH = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";
const LEGACY_HARNESS_PATH = "C:\\fake\\legacy-harness";
const SYSTEM_INSTRUCTION_TEXT = "Always reply in formal English and cite the document section you used.";
const PERSONA_TEXT = "You are a senior bid manager with 15 years of government-tender experience.";

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
  // synchronously at the top of the Office.onReady callback, before the settings-field
  // refill code below it. Without this, that line throws and aborts the rest of the
  // callback, failing the refill assertions for an unrelated reason.
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

  // "networkidle" would never fire: taskpane.js opens a persistent SSE connection
  // to /event on Office.onReady (and, with no opencode serve running in this test,
  // keeps auto-retrying it), so the page never goes network-idle. Wait for a
  // concrete DOM readiness signal instead.
  await page.goto("https://localhost:3000/taskpane.html", { waitUntil: "load" });
  await page.waitForSelector("#settings-toggle", { state: "visible" });

  const failures = [];

  // --- Legacy migration: a saved harness root becomes a System Instruction ---
  await page.evaluate((legacyPath) => {
    localStorage.clear();
    localStorage.setItem("openCodeHarnessRoot", legacyPath);
  }, LEGACY_HARNESS_PATH);
  await page.reload({ waitUntil: "load" });
  await page.waitForSelector("#settings-toggle", { state: "visible" });

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

  const migratedInstruction = await page.$eval("#system-instruction-input", (el) => el.value);
  if (!migratedInstruction.includes(LEGACY_HARNESS_PATH)) {
    failures.push(
      `legacy harness root was not migrated into System Instruction: got "${migratedInstruction.slice(0, 120)}"`
    );
  }
  const legacyKeyAfterMigration = await page.evaluate(() => localStorage.getItem("openCodeHarnessRoot"));
  if (legacyKeyAfterMigration !== null) {
    failures.push(`legacy openCodeHarnessRoot key was not removed after migration: "${legacyKeyAfterMigration}"`);
  }

  // --- Save: both fields land in localStorage ---
  await page.fill("#system-instruction-input", SYSTEM_INSTRUCTION_TEXT);
  await page.fill("#persona-input", PERSONA_TEXT);
  await page.click("#settings-save");
  const storedInstruction = await page.evaluate(() => localStorage.getItem("openCodeSystemInstruction"));
  const storedPersona = await page.evaluate(() => localStorage.getItem("openCodePersona"));
  if (storedInstruction !== SYSTEM_INSTRUCTION_TEXT) {
    failures.push(`openCodeSystemInstruction not saved: got "${storedInstruction}"`);
  }
  if (storedPersona !== PERSONA_TEXT) {
    failures.push(`openCodePersona not saved: got "${storedPersona}"`);
  }
  const statusText = await page.$eval("#settings-status", (el) => el.textContent);
  if (!statusText.trim()) {
    failures.push("settings status line is empty after Save - the user gets no confirmation");
  }

  // --- Reload: fields re-fill from localStorage ---
  await page.reload({ waitUntil: "load" });
  await page.waitForSelector("#settings-toggle", { state: "visible" });
  await page.click("#settings-toggle");
  const refilledInstruction = await page.$eval("#system-instruction-input", (el) => el.value);
  const refilledPersona = await page.$eval("#persona-input", (el) => el.value);
  if (refilledInstruction !== SYSTEM_INSTRUCTION_TEXT) {
    failures.push(`System Instruction was not re-filled after reload: got "${refilledInstruction.slice(0, 120)}"`);
  }
  if (refilledPersona !== PERSONA_TEXT) {
    failures.push(`Persona was not re-filled after reload: got "${refilledPersona.slice(0, 120)}"`);
  }

  // --- Clear: wipes the fields and localStorage ---
  await page.click("#settings-clear");
  const clearedInstruction = await page.$eval("#system-instruction-input", (el) => el.value);
  const clearedPersona = await page.$eval("#persona-input", (el) => el.value);
  if (clearedInstruction !== "" || clearedPersona !== "") {
    failures.push(`fields were not cleared: instruction="${clearedInstruction}", persona="${clearedPersona}"`);
  }
  const storedAfterClear = await page.evaluate(() => ({
    instruction: localStorage.getItem("openCodeSystemInstruction"),
    persona: localStorage.getItem("openCodePersona"),
  }));
  if (storedAfterClear.instruction !== null || storedAfterClear.persona !== null) {
    failures.push(`localStorage still has settings after Clear: ${JSON.stringify(storedAfterClear)}`);
  }

  if (consoleErrors.length > 0) {
    failures.push(`browser console errors: ${JSON.stringify(consoleErrors)}`);
  }

  await browser.close();

  if (failures.length > 0) {
    console.error("FAIL:\n" + failures.map((f) => `  - ${f}`).join("\n"));
    process.exit(1);
  }
  console.log("PASS: Settings panel (gear icon: System Instruction + Persona, migration, persist, clear) works");
}

main().catch((err) => {
  console.error(`FAIL: ${err.message}`);
  process.exit(1);
});
