# Phase 10 test: harness auto-detection + harness-mode task-pane UI. Static
# checks confirm the wiring is present; the functional check builds the bundle,
# stands up serve-static pointed (via HARNESS_DETECT_BASE) at a temp directory
# containing `_agentic/os/AGENTS.md`, and drives the real task pane in headless
# Edge to confirm it flips into harness mode. No real vault and no running
# opencode are required (the .mjs stubs opencode's endpoints).

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$failures = @()
$devServerPort = 3000

Push-Location $root
try {
    # --- Static checks ---
    $html = Get-Content (Join-Path $root "src\taskpane\taskpane.html") -Raw
    foreach ($id in @("harness-banner", "memory-save", "manual-settings-fields", "settings-actions")) {
        if ($html -notmatch [regex]::Escape("id=`"$id`"")) {
            $failures += "taskpane.html is missing id=`"$id`""
        }
    }
    $js = Get-Content (Join-Path $root "src\taskpane\taskpane.js") -Raw
    foreach ($fn in @("applyHarnessMode", "harnessHiddenBlock", "onSaveToMemory", "buildSaveToMemoryInstruction")) {
        if ($js -notmatch [regex]::Escape($fn)) {
            $failures += "taskpane.js does not define $fn"
        }
    }
    $ss = Get-Content (Join-Path $root "scripts\serve-static.js") -Raw
    if ($ss -notmatch [regex]::Escape("/harness-info")) {
        $failures += "serve-static.js does not serve /harness-info"
    }
    if ($ss -notmatch [regex]::Escape("detectHarnessRoot")) {
        $failures += "serve-static.js does not define detectHarnessRoot"
    }

    if (-not (Test-Path (Join-Path $root "node_modules"))) {
        $failures += "node_modules missing - run 'npm install' first"
    } else {
        # --- Functional check ---
        $buildOut = npm run build 2>&1 | Out-String
        if ($LASTEXITCODE -ne 0) {
            $failures += "npm run build failed:`n$buildOut"
        } else {
            # Temp harness layout: <temp>\_agentic\os\AGENTS.md
            $tempHarness = Join-Path ([System.IO.Path]::GetTempPath()) ("phase10-harness-" + [guid]::NewGuid().ToString("N"))
            $osDir = Join-Path $tempHarness "_agentic\os"
            New-Item -ItemType Directory -Path $osDir -Force | Out-Null
            Set-Content -Path (Join-Path $osDir "AGENTS.md") -Value "# test harness" -Encoding ASCII

            # Free port 3000 so our serve-static (not a stray one) owns it.
            Get-NetTCPConnection -LocalPort $devServerPort -ErrorAction SilentlyContinue |
                ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }
            Start-Sleep -Seconds 1

            $job = Start-Job -ScriptBlock {
                param($root, $base)
                $env:HARNESS_DETECT_BASE = $base
                Set-Location $root
                node scripts/serve-static.js
            } -ArgumentList $root, $tempHarness

            try {
                $ready = $false
                for ($i = 0; $i -lt 15; $i++) {
                    Start-Sleep -Seconds 1
                    try {
                        $resp = Invoke-WebRequest -Uri "https://localhost:$devServerPort/harness-info" -TimeoutSec 3 -UseBasicParsing
                        if ($resp.StatusCode -eq 200) { $ready = $true; break }
                    } catch { }
                }
                if (-not $ready) {
                    $failures += "serve-static did not answer /harness-info on :$devServerPort within 15s"
                } else {
                    $env:PHASE10_EXPECTED_ROOT = $tempHarness
                    # No 2>&1 (Windows PowerShell 5.1 wraps native stderr as errors under -Stop).
                    $nodeOut = node (Join-Path $PSScriptRoot "phase10-harness.mjs") | Out-String
                    Write-Host $nodeOut
                    if ($LASTEXITCODE -ne 0) {
                        $failures += "phase10-harness.mjs failed:`n$nodeOut"
                    }
                    Remove-Item Env:\PHASE10_EXPECTED_ROOT -ErrorAction SilentlyContinue
                }
            } finally {
                Stop-Job $job -ErrorAction SilentlyContinue | Out-Null
                Remove-Job $job -Force -ErrorAction SilentlyContinue | Out-Null
                Get-NetTCPConnection -LocalPort $devServerPort -ErrorAction SilentlyContinue |
                    ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }
                Remove-Item -Path $tempHarness -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
} finally {
    Pop-Location
}

if ($failures.Count -eq 0) {
    Write-Host "PASS: Phase 10 harness checks (auto-detect + harness-mode UI + Save to memory)" -ForegroundColor Green
    exit 0
} else {
    Write-Host "FAIL: Phase 10 harness checks" -ForegroundColor Red
    $failures | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    exit 1
}
