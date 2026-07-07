# Phase 8 test: verifies the autoopen manifest/taskpane wiring is in place, and
# that the autostart scripts actually do what they claim - start-background-
# services.ps1 really brings up both services (and is idempotent), and
# install-autostart.ps1 / uninstall-autostart.ps1 really round-trip the Startup
# folder entry. Cleans up any ports/files it touches so it doesn't leave a
# stray autostart entry or running process behind on the dev machine.

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$failures = @()
$devServerPort = 3000
$opencodePort = 4098

function Test-PortListening($port) {
    return [bool](Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue)
}

function Stop-PortOwner($port) {
    Get-NetTCPConnection -LocalPort $port -ErrorAction SilentlyContinue |
        ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }
}

Push-Location $root
try {
    # --- Static checks: manifest.xml + taskpane.js wiring ---
    $manifestPath = Join-Path $root "manifest.xml"
    if (-not (Test-Path $manifestPath)) {
        $failures += "manifest.xml not found at $manifestPath"
    } elseif ((Get-Content $manifestPath -Raw) -notmatch [regex]::Escape("Office.AutoShowTaskpaneWithDocument")) {
        $failures += "manifest.xml does not reference Office.AutoShowTaskpaneWithDocument"
    }

    $taskpaneJsPath = Join-Path $root "src\taskpane\taskpane.js"
    if (-not (Test-Path $taskpaneJsPath)) {
        $failures += "src\taskpane\taskpane.js not found at $taskpaneJsPath"
    } elseif ((Get-Content $taskpaneJsPath -Raw) -notmatch [regex]::Escape('Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument"')) {
        $failures += "taskpane.js does not tag the document with Office.AutoShowTaskpaneWithDocument"
    }

    # --- Functional check: start-background-services.ps1 actually starts both services ---
    if (Test-PortListening $devServerPort) {
        $failures += "port $devServerPort was already in use before the test started - can't verify start-background-services.ps1 cleanly"
    } elseif (Test-PortListening $opencodePort) {
        $failures += "port $opencodePort was already in use before the test started - can't verify start-background-services.ps1 cleanly"
    } else {
        try {
            & (Join-Path $root "scripts\start-background-services.ps1")

            $devServerReady = $false
            for ($i = 0; $i -lt 30; $i++) {
                Start-Sleep -Seconds 2
                try {
                    $resp = Invoke-WebRequest -Uri "https://localhost:$devServerPort/taskpane.html" -TimeoutSec 3 -UseBasicParsing
                    if ($resp.StatusCode -eq 200) { $devServerReady = $true; break }
                } catch { }
            }
            if (-not $devServerReady) {
                $failures += "start-background-services.ps1 did not bring up the dev server (https://localhost:$devServerPort/taskpane.html) within 60s"
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
                $failures += "start-background-services.ps1 did not bring up opencode serve (http://127.0.0.1:$opencodePort/global/health) within 20s"
            }

            # Idempotency: running it again with both ports already up should not error
            # and should not report starting new processes. Write-Host output lands on
            # the Information stream (6), not the success stream, so it must be
            # explicitly redirected (6>&1) to be captured here.
            $rerunOutput = & (Join-Path $root "scripts\start-background-services.ps1") 6>&1 | Out-String
            if ($rerunOutput -notmatch "already listening") {
                $failures += "start-background-services.ps1 did not skip already-running services on a second run:`n$rerunOutput"
            }
        } finally {
            Stop-PortOwner $devServerPort
            Stop-PortOwner $opencodePort
        }
    }

    # --- Functional check: install/uninstall-autostart.ps1 round trip ---
    $startupDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Startup"
    $vbsPath = Join-Path $startupDir "AI Assistant-autostart.vbs"
    $preExisting = Test-Path $vbsPath
    try {
        & (Join-Path $root "scripts\install-autostart.ps1") | Out-Null
        if (-not (Test-Path $vbsPath)) {
            $failures += "install-autostart.ps1 did not create $vbsPath"
        } else {
            $vbsContent = Get-Content $vbsPath -Raw
            if ($vbsContent -notmatch [regex]::Escape("start-background-services.ps1")) {
                $failures += "$vbsPath does not reference start-background-services.ps1"
            }
        }

        & (Join-Path $root "scripts\uninstall-autostart.ps1") | Out-Null
        if (Test-Path $vbsPath) {
            $failures += "uninstall-autostart.ps1 did not remove $vbsPath"
        }
    } finally {
        # Leave the machine in the state it was found: only remove the file if this
        # test run was the one that created it (i.e. it didn't pre-exist).
        if (-not $preExisting -and (Test-Path $vbsPath)) {
            Remove-Item -Path $vbsPath -Force -ErrorAction SilentlyContinue
        }
    }
} finally {
    Pop-Location
}

if ($failures.Count -eq 0) {
    Write-Host "PASS: Phase 8 autostart checks (manifest/taskpane wiring, background services, Startup folder round trip)" -ForegroundColor Green
    exit 0
} else {
    Write-Host "FAIL: Phase 8 autostart checks" -ForegroundColor Red
    $failures | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    exit 1
}
