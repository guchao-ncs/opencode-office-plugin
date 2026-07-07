// Phase 5 test: drives the real taskpane.html chat UI in headless Edge (via
// playwright-core, reusing the system-installed browser instead of downloading
// one) and verifies a full round trip through the actual OpenCode REST API.
//
// Office.onReady() never resolves outside a real Office host, so taskpane.js's
// listener-attachment code would never run under a plain browser. This test
// stubs out Office.js (the same way Word's WebView2 host injects it) to
// simulate host === Word, which is the only way to drive the UI headlessly.
//
// Assumes the caller (phase5-chat-ui.ps1) already has both the webpack dev
// server (https://localhost:3000) and `opencode serve` (with --cors for that
// origin) running.
import { chromium } from "playwright-core";

const EDGE_PATH = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";

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
  // synchronously at the top of the Office.onReady callback. Without this, that line
  // throws ("Cannot read properties of undefined (reading 'document')"), which would
  // show up as a pageerror and fail the consoleErrors check below.
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
  // connection to /event on Office.onReady for streaming replies, so the page
  // never goes network-idle. Wait for the concrete DOM readiness signal
  // instead (the chat input becoming visible, set up at the end of onReady).
  await page.goto("https://localhost:3000/taskpane.html", { waitUntil: "load" });
  await page.waitForSelector("#chat-input", { state: "visible" });
  await page.fill("#chat-input", "Say the single word: pong");
  await page.click("#chat-send");

  // 60s, not 30s: this waits on a real LLM round trip over a corporate network that has
  // shown occasional latency spikes elsewhere in this project's testing (e.g. npm installs
  // stalling for minutes); 30s intermittently timed out here even on otherwise-healthy runs.
  //
  // page.waitForFunction(pageFunction, arg, options) takes THREE positional params -
  // arg (passed through to pageFunction) is second, options is third. Passing
  // { timeout: 60000 } as the second argument silently became `arg` (ignored, since
  // the predicate below takes no parameters) rather than `options`, so this was
  // actually running under Playwright's default 30000ms the whole time. The explicit
  // `undefined` below is required to reach the options position.
  await page.waitForFunction(
    () => document.querySelectorAll(".chat-message--assistant:not(.chat-message--pending)").length > 0,
    undefined,
    { timeout: 60000 }
  );

  const userText = await page.$eval(".chat-message--user", (el) => el.textContent.trim());
  const assistantText = await page.$eval(
    ".chat-message--assistant:not(.chat-message--pending)",
    (el) => el.textContent.trim()
  );
  const errorBubbles = await page.$$(".chat-message--error");

  await browser.close();

  const failures = [];
  if (userText !== "Say the single word: pong") {
    failures.push(`user bubble text mismatch: "${userText}"`);
  }
  if (!assistantText) {
    failures.push("assistant bubble is empty");
  }
  if (errorBubbles.length > 0) {
    failures.push(`chat log shows ${errorBubbles.length} error bubble(s)`);
  }
  if (consoleErrors.length > 0) {
    failures.push(`browser console errors: ${JSON.stringify(consoleErrors)}`);
  }

  if (failures.length > 0) {
    console.error("FAIL:\n" + failures.map((f) => `  - ${f}`).join("\n"));
    process.exit(1);
  }
  console.log(`PASS: chat round trip works (assistant replied: "${assistantText}")`);
}

main().catch((err) => {
  console.error("FAIL: script error:", err.message);
  process.exit(1);
});
