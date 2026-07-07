# Phase 5 test: full round trip through the real taskpane chat UI, driven by a
# headless Edge browser (playwright-core) against the real webpack dev server
# and a real `opencode serve` instance (with --cors, since the taskpane origin
# https://localhost:3000 is different from OpenCode's http://127.0.0.1:4098 -
# see PRD.md for the CORS-vs-mixed-content finding this depends on).

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$failures = @()
$opencodePort = 4098

Push-Location $root
try {
    if (-not (Test-Path (Join-Path $root "node_modules"))) {
        $failures += "node_modules missing - run 'npm install' first"
    } else {
        $devServerJob = Start-Job -ScriptBlock {
            param($root)
            Set-Location $root
            npm run dev-server
        } -ArgumentList $root

        $opencodeJob = Start-Job -ScriptBlock {
            param($root, $port)
            Set-Location $root
            opencode serve --port $port --hostname 127.0.0.1 --cors https://localhost:3000
        } -ArgumentList $root, $opencodePort

        try {
            $devServerReady = $false
            for ($i = 0; $i -lt 30; $i++) {
                Start-Sleep -Seconds 2
                try {
                    $resp = Invoke-WebRequest -Uri "https://localhost:3000/taskpane.html" -TimeoutSec 3 -UseBasicParsing
                    if ($resp.StatusCode -eq 200) { $devServerReady = $true; break }
                } catch { }
            }
            if (-not $devServerReady) {
                $failures += "https://localhost:3000/taskpane.html did not respond with 200 within 60s"
            }

            $opencodeReady = $false
            for ($i = 0; $i -lt 20; $i++) {
                Start-Sleep -Seconds 1
                try {
                    $resp = Invoke-WebRequest -Uri "http://127.0.0.1:$opencodePort/global/health" -TimeoutSec 3 -UseBasicParsing
                    if ($resp.StatusCode -eq 200) {
                        $body = $resp.Content | ConvertFrom-Json
                        if ($body.healthy -eq $true) { $opencodeReady = $true; break }
                    }
                } catch { }
            }
            if (-not $opencodeReady) {
                $failures += "http://127.0.0.1:$opencodePort/global/health did not report healthy within 20s"
            }

            if ($devServerReady -and $opencodeReady) {
                # No 2>&1 here: on Windows PowerShell 5.1, redirecting a native exe's stderr
                # wraps each line as a terminating NativeCommandError under -Stop, aborting
                # this script even when node exits 0 (same issue fixed in phase4-word-mcp.ps1).
                $nodeOutput = node (Join-Path $PSScriptRoot "phase5-chat-ui.mjs") | Out-String
                Write-Host $nodeOutput
                if ($LASTEXITCODE -ne 0) {
                    $failures += "Browser round-trip test (phase5-chat-ui.mjs) failed:`n$nodeOutput"
                }
            }
        } finally {
            Stop-Job $devServerJob, $opencodeJob -ErrorAction SilentlyContinue | Out-Null
            Remove-Job $devServerJob, $opencodeJob -Force -ErrorAction SilentlyContinue | Out-Null
            Get-NetTCPConnection -LocalPort 3000 -ErrorAction SilentlyContinue |
                ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }
            Get-NetTCPConnection -LocalPort $opencodePort -ErrorAction SilentlyContinue |
                ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }
        }
    }
} finally {
    Pop-Location
}

if ($failures.Count -eq 0) {
    Write-Host "PASS: Phase 5 chat UI round trip (dev server + opencode serve + browser)" -ForegroundColor Green
    exit 0
} else {
    Write-Host "FAIL: Phase 5 chat UI round trip" -ForegroundColor Red
    $failures | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    exit 1
}
