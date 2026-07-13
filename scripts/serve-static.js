/* eslint-disable no-undef */
//
// Minimal HTTPS static file server for the pre-built taskpane bundle (dist/).
//
// Why this exists: the autostart path used to run `webpack serve`, which
// recompiles the whole bundle on every launch - after a cold boot that added
// a 1-2 minute window where port 3000 was not yet serving, so Word's attempt
// to load the sideloaded add-in failed with "ADD-IN ERROR". Serving a
// pre-built dist/ (produced once by `npm run build`) instead makes the server
// ready almost immediately, closing most of that window.
//
// It deliberately reuses the SAME trusted dev certificate as webpack
// (office-addin-dev-certs' getHttpsServerOptions) and serves the SAME files
// at the SAME URLs over HTTPS on the SAME port, plus the same permissive CORS
// header webpack-dev-server sent - so from Word's / the taskpane's point of
// view nothing about the served content, origin, or trust changes. `npm run
// dev-server` (webpack serve) is untouched and remains the tool for active
// development (live recompile); this server is only for the always-on
// background/autostart path where the code isn't changing between reboots.

const https = require("https");
const fs = require("fs");
const path = require("path");
const devCerts = require("office-addin-dev-certs");

const distDir = path.join(__dirname, "..", "dist");
const pluginRoot = path.join(__dirname, "..");
const port = process.env.npm_package_config_dev_server_port || 3000;

// Detects a Solution Architect harness vault on this machine so the task pane
// can switch to "harness mode" (see taskpane.js). A vault is any directory
// containing `_agentic/os/AGENTS.md`. We look in two places, since the plugin
// is installed alongside the vault, not inside it:
//   1. every ancestor of the plugin directory (in case it IS nested one day)
//   2. the plugin's immediate siblings (the common layout: the vault and this
//      plugin are two folders under the same projects directory)
// First match wins; the resolved absolute path is shown in the UI so the user
// can confirm which vault was picked. Detection is best-effort and never
// throws - any FS error just means "no harness" (generic mode).
function detectHarnessRoot() {
  // HARNESS_DETECT_BASE overrides where detection starts (defaults to the
  // plugin root). Used by the phase10 test to point detection at a controlled
  // temp layout; also lets an advanced user force a specific search base.
  const base = process.env.HARNESS_DETECT_BASE || pluginRoot;
  const marker = path.join("_agentic", "os", "AGENTS.md");
  const hasMarker = (dir) => {
    try {
      return fs.statSync(path.join(dir, marker)).isFile();
    } catch {
      return false;
    }
  };

  // 1) Ancestors of the base directory (including the base itself).
  let dir = base;
  while (true) {
    if (hasMarker(dir)) {
      return dir;
    }
    const parent = path.dirname(dir);
    if (parent === dir) {
      break;
    }
    dir = parent;
  }

  // 2) Immediate siblings (one level only, to avoid a broad/slow scan).
  try {
    const parent = path.dirname(base);
    for (const name of fs.readdirSync(parent)) {
      const sibling = path.join(parent, name);
      if (sibling === base) {
        continue;
      }
      let isDir = false;
      try {
        isDir = fs.statSync(sibling).isDirectory();
      } catch {
        isDir = false;
      }
      if (isDir && hasMarker(sibling)) {
        return sibling;
      }
    }
  } catch {
    // parent unreadable - fall through to "no harness".
  }
  return null;
}

const harnessRoot = detectHarnessRoot();

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".map": "application/json; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".xml": "application/xml; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".ico": "image/x-icon",
  ".svg": "image/svg+xml",
};

function handleRequest(req, res) {
  // Same permissive CORS header webpack-dev-server set, to keep behavior
  // identical (source maps / sub-resources fetched cross-origin by devtools,
  // etc.). The load-bearing cross-origin grant for talking to opencode is on
  // opencode's own --cors flag, not here.
  res.setHeader("Access-Control-Allow-Origin", "*");

  // Strip query/hash, default "/" to the taskpane entry point.
  let urlPath = decodeURIComponent((req.url || "/").split("?")[0].split("#")[0]);

  // Harness detection signal for the task pane (mode + root only, never file
  // contents). Absent on webpack dev-server / nginx, where the task pane just
  // falls back to generic mode.
  if (urlPath === "/harness-info") {
    res.setHeader("Content-Type", "application/json; charset=utf-8");
    res.end(JSON.stringify(harnessRoot ? { mode: "harness", root: harnessRoot } : { mode: "generic" }));
    return;
  }

  // Browsers auto-request /favicon.ico; we don't ship one. Answer 204 (rather
  // than 404) so it doesn't surface as a console error.
  if (urlPath === "/favicon.ico") {
    res.writeHead(204);
    res.end();
    return;
  }

  if (urlPath === "/") {
    urlPath = "/taskpane.html";
  }

  // Resolve within distDir and refuse anything that escapes it (path traversal).
  const resolved = path.join(distDir, path.normalize(urlPath));
  if (resolved !== distDir && !resolved.startsWith(distDir + path.sep)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  fs.stat(resolved, (err, stats) => {
    if (err || !stats.isFile()) {
      res.writeHead(404);
      res.end("Not found");
      return;
    }
    res.setHeader("Content-Type", MIME_TYPES[path.extname(resolved).toLowerCase()] || "application/octet-stream");
    fs.createReadStream(resolved)
      .on("error", () => {
        res.writeHead(500);
        res.end("Read error");
      })
      .pipe(res);
  });
}

async function main() {
  if (!fs.existsSync(path.join(distDir, "taskpane.html"))) {
    console.error(`dist/ is not built (no taskpane.html at ${distDir}). Run "npm run build" first.`);
    process.exit(1);
  }
  const httpsOptions = await devCerts.getHttpsServerOptions();
  const server = https.createServer(
    { key: httpsOptions.key, cert: httpsOptions.cert, ca: httpsOptions.ca },
    handleRequest
  );
  server.on("error", (err) => {
    // The idempotent PowerShell launcher checks the port before starting us,
    // but guard anyway so a race just exits cleanly instead of crashing loud.
    if (err.code === "EADDRINUSE") {
      console.error(`Port ${port} already in use - another server is already serving. Exiting.`);
      process.exit(0);
    }
    console.error(`Server error: ${err.message}`);
    process.exit(1);
  });
  server.listen(port, () => {
    console.log(`Serving dist/ over HTTPS on https://localhost:${port}`);
    console.log(harnessRoot ? `Harness detected: ${harnessRoot}` : "No harness detected (generic mode)");
  });
}

main().catch((err) => {
  console.error(`Failed to start static server: ${err.message}`);
  process.exit(1);
});
