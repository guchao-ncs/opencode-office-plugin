// Phase 10 test: the harness auto-detection + harness-mode UI. Two parts:
//   1. serve-static's /harness-info returns the right signal for a controlled
//      layout (driven from phase10-harness.ps1 via HARNESS_DETECT_BASE, so it
//      does NOT depend on the real vault existing on this machine).
//   2. In the task pane, harness mode disables the manual Settings fields,
//      shows the banner + "Save to memory", and hides Save/Clear; and clicking
//      "Save to memory" posts a memory-write turn to opencode.
//
// The .ps1 harness starts serve-static with HARNESS_DETECT_BASE pointed at a
// temp dir containing `_agentic/os/AGENTS.md`, so /harness-info reports harness
// mode with that temp root. Assumes serve-static is already running on :3000.
import { chromium } from "playwright-core";

const EDGE_PATH = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";
const EXPECTED_ROOT = process.env.PHASE10_EXPECTED_ROOT || "";

async function main() {
  const browser = await chromium.launch({
    executablePath: EDGE_PATH,
    args: ["--ignore-certificate-errors"],
  });
  const page = await browser.newPage();
  // Only real JS exceptions and genuine 404s count as failures. Generic
  // "Failed to load resource" console lines are network noise - in particular
  // the /event SSE stream is deliberately aborted below (net::ERR_FAILED), and
  // that must not be treated as a defect.
  const consoleErrors = [];
  page.on("console", (msg) => {
    if (msg.type() === "error" && !/Failed to load resource/i.test(msg.text())) {
      consoleErrors.push(msg.text());
    }
  });
  page.on("pageerror", (err) => consoleErrors.push(`pageerror: ${err.message}`));
  page.on("response", (resp) => {
    if (resp.status() === 404) consoleErrors.push(`404: ${resp.url()}`);
  });

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

  const failures = [];

  // 1) /harness-info reports harness mode with the temp root.
  await page.goto("https://localhost:3000/harness-info", { waitUntil: "load" });
  const infoText = await page.evaluate(() => document.body.innerText);
  let info;
  try {
    info = JSON.parse(infoText);
  } catch {
    info = null;
  }
  if (!info || info.mode !== "harness") {
    failures.push(`/harness-info did not report harness mode: ${infoText}`);
  } else if (EXPECTED_ROOT && info.root !== EXPECTED_ROOT) {
    failures.push(`/harness-info root was "${info.root}", expected "${EXPECTED_ROOT}"`);
  }

  // Stub opencode's endpoints BEFORE loading the task pane, since the task
  // pane opens its SSE stream on Office.onReady. Aborting /event forces the
  // deterministic blocking /message path regardless of whether a real opencode
  // is running on :4098 on this machine.
  let memoryRequestBody = null;
  await page.route("**/event", (route) => route.abort());
  await page.route("**/session/*/prompt_async", (route) =>
    route.fulfill({ status: 200, contentType: "application/json", body: "{}" })
  );
  await page.route("**/session", (route) => {
    if (route.request().method() === "POST") {
      route.fulfill({ contentType: "application/json", body: JSON.stringify({ id: "ses_test10" }) });
    } else {
      route.continue();
    }
  });
  await page.route("**/session/*/message", (route) => {
    const body = route.request().postData() || "";
    if (/Save the key points/i.test(body)) {
      memoryRequestBody = body;
    }
    route.fulfill({
      contentType: "application/json",
      body: JSON.stringify({ parts: [{ type: "text", text: "Wrote to memory/today.md" }] }),
    });
  });

  // 2) Task pane switches to harness mode.
  await page.goto("https://localhost:3000/taskpane.html", { waitUntil: "load" });
  await page.waitForSelector("#settings-toggle", { state: "visible" });
  await page.click("#settings-toggle");

  // Give applyHarnessMode()'s fetch a moment to resolve and update the DOM.
  await page.waitForFunction(
    () => getComputedStyle(document.getElementById("harness-banner")).display !== "none",
    undefined,
    { timeout: 5000 }
  ).catch(() => {});

  const bannerVisible = await page.$eval("#harness-banner", (el) => getComputedStyle(el).display !== "none");
  if (!bannerVisible) {
    failures.push("harness banner not shown in harness mode");
  }
  // The manual System Instruction/Persona fields (and their Save/Clear) are
  // hidden as a unit in harness mode.
  const fieldsHidden = await page.$eval("#manual-settings-fields", (el) => getComputedStyle(el).display === "none");
  if (!fieldsHidden) {
    failures.push("manual System Instruction/Persona fields should be hidden in harness mode");
  }
  // "Save to memory" lives in the toolbar next to "Prompt ideas".
  const memorySaveVisible = await page.$eval("#memory-save", (el) => getComputedStyle(el).display !== "none");
  if (!memorySaveVisible) {
    failures.push('"Save to memory" button should be visible in the toolbar in harness mode');
  }

  // 3) "Save to memory" posts a memory-write turn (opencode stubbed above).
  // Let the initial SSE connection settle, then seed a session with one normal
  // chat message via the blocking path so "Save to memory" has a current
  // session to reuse.
  await page.waitForTimeout(3500);
  await page.fill("#chat-input", "hello");
  await page.click("#chat-send");
  await page.waitForTimeout(2500);

  // The settings panel is still open from step 2; memory-save should be
  // visible in harness mode.
  await page.waitForSelector("#memory-save", { state: "visible", timeout: 5000 });
  await page.click("#memory-save");
  await page.waitForTimeout(2000);
  if (!memoryRequestBody || !/Save the key points/i.test(memoryRequestBody)) {
    failures.push(`"Save to memory" did not post a memory-write instruction (body: ${String(memoryRequestBody).slice(0, 120)})`);
  }

  if (consoleErrors.length > 0) {
    failures.push(`browser console errors: ${JSON.stringify(consoleErrors)}`);
  }

  await browser.close();

  if (failures.length > 0) {
    console.error("FAIL:\n" + failures.map((f) => `  - ${f}`).join("\n"));
    process.exit(1);
  }
  console.log("PASS: harness auto-detection + harness-mode UI (disabled fields, banner, Save to memory) works");
}

main().catch((err) => {
  console.error(`FAIL: ${err.message}`);
  process.exit(1);
});
