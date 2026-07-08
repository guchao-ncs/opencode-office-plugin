# Phase 9 test: verifies the taskpane's Settings panel (gear icon) - the System
# Instruction / Persona fields that get injected as a hidden instruction on each
# new session's first message. Static checks confirm the HTML/JS wiring is
# present; the functional check drives the real UI in headless Edge via
# playwright-core against the real dev server.
#
# Unlike phase5, this does NOT need a running `opencode serve` - the panel is
# pure client-side (localStorage); the injection itself rides the same
# message-sending flow phase5 already covers.

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$failures = @()
$devServerPort = 3000

Push-Location $root
try {
    # --- Static checks: taskpane.html / taskpane.js wiring ---
    $taskpaneHtmlPath = Join-Path $root "src\taskpane\taskpane.html"
    if (-not (Test-Path $taskpaneHtmlPath)) {
        $failures += "src\taskpane\taskpane.html not found at $taskpaneHtmlPath"
    } else {
        $htmlContent = Get-Content $taskpaneHtmlPath -Raw
        foreach ($id in @("settings-toggle", "system-instruction-input", "persona-input", "settings-save", "settings-clear", "tab-btn-settings", "tab-btn-library", "prompt-library-list", "library-add", "library-save", "library-reset")) {
            if ($htmlContent -notmatch [regex]::Escape("id=`"$id`"")) {
                $failures += "taskpane.html is missing an element with id=`"$id`""
            }
        }
    }

    $taskpaneJsPath = Join-Path $root "src\taskpane\taskpane.js"
    if (-not (Test-Path $taskpaneJsPath)) {
        $failures += "src\taskpane\taskpane.js not found at $taskpaneJsPath"
    } else {
        $jsContent = Get-Content $taskpaneJsPath -Raw
        if ($jsContent -notmatch [regex]::Escape("customizationHiddenBlock")) {
            $failures += "taskpane.js does not define customizationHiddenBlock"
        }
        if ($jsContent -notmatch [regex]::Escape("getPromptLibrary")) {
            $failures += "taskpane.js does not define getPromptLibrary"
        }
        if ($jsContent -match [regex]::Escape("method: `"PATCH`"")) {
            $failures += "taskpane.js still issues a PATCH request - this was found to silently not persist against a real opencode serve instance and should not be relied on"
        }
    }

    # --- Functional check: real UI, real dev server, no mocking ---
    if (-not (Test-Path (Join-Path $root "node_modules"))) {
        $failures += "node_modules missing - run 'npm install' first"
    } else {
        $devServerJob = Start-Job -ScriptBlock {
            param($root)
            Set-Location $root
            npm run dev-server
        } -ArgumentList $root

        try {
            $devServerReady = $false
            for ($i = 0; $i -lt 30; $i++) {
                Start-Sleep -Seconds 2
                try {
                    $resp = Invoke-WebRequest -Uri "https://localhost:$devServerPort/taskpane.html" -TimeoutSec 3 -UseBasicParsing
                    if ($resp.StatusCode -eq 200) { $devServerReady = $true; break }
                } catch { }
            }
            if (-not $devServerReady) {
                $failures += "https://localhost:$devServerPort/taskpane.html did not respond with 200 within 60s"
            } else {
                # No 2>&1 here: on Windows PowerShell 5.1, redirecting a native exe's stderr
                # wraps each line as a terminating NativeCommandError under -Stop, aborting
                # this script even when node exits 0 (same issue fixed in phase4-word-mcp.ps1).
                $nodeOutput = node (Join-Path $PSScriptRoot "phase9-settings-panel.mjs") | Out-String
                Write-Host $nodeOutput
                if ($LASTEXITCODE -ne 0) {
                    $failures += "Settings panel round trip (phase9-settings-panel.mjs) failed:`n$nodeOutput"
                }
            }
        } finally {
            Stop-Job $devServerJob -ErrorAction SilentlyContinue | Out-Null
            Remove-Job $devServerJob -Force -ErrorAction SilentlyContinue | Out-Null
            Get-NetTCPConnection -LocalPort $devServerPort -ErrorAction SilentlyContinue |
                ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }
        }
    }
} finally {
    Pop-Location
}

if ($failures.Count -eq 0) {
    Write-Host "PASS: Phase 9 settings panel checks (System Instruction + Persona + Prompt library)" -ForegroundColor Green
    exit 0
} else {
    Write-Host "FAIL: Phase 9 settings panel checks" -ForegroundColor Red
    $failures | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    exit 1
}
