# Phase 7 test: README.md exists and documents the operational details that
# were hard-won during earlier phases (the --cors requirement, the correct
# word-mcp-live invocation, the one-time cert-trust step, and the locked-
# session Word automation hang) so a future reader doesn't have to
# rediscover them.

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$failures = @()

$readmePath = Join-Path $root "README.md"
if (-not (Test-Path $readmePath)) {
    $failures += "README.md not found at $readmePath"
} else {
    $content = Get-Content $readmePath -Raw
    $requiredSnippets = @(
        "npm install",
        "office-addin-dev-certs install",
        "opencode serve",
        "--cors",
        "npm start",
        "word_mcp_server",
        "opencode mcp list",
        "run-all.ps1"
    )
    foreach ($snippet in $requiredSnippets) {
        if ($content -notmatch [regex]::Escape($snippet)) {
            $failures += "README.md is missing expected content: '$snippet'"
        }
    }
}

if ($failures.Count -eq 0) {
    Write-Host "PASS: Phase 7 README checks (setup, run, MCP wiring, testing sections present)" -ForegroundColor Green
    exit 0
} else {
    Write-Host "FAIL: Phase 7 README checks" -ForegroundColor Red
    $failures | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    exit 1
}
