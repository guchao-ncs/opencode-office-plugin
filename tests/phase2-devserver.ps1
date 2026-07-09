# Phase 2 test: verify manifest validates and the HTTPS dev server actually serves
# the taskpane, without requiring a full `npm start` (which launches Word's GUI and
# needs a one-time interactive cert-trust click - that part stays a manual step,
# documented in README.md).

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$failures = @()

Push-Location $root
try {
    if (-not (Test-Path (Join-Path $root "node_modules"))) {
        $failures += "node_modules missing - run 'npm install' first"
    } else {
        # 1) manifest.xml passes Microsoft's own schema/content validator
        $validateOutput = npm run validate 2>&1 | Out-String
        if ($LASTEXITCODE -ne 0) {
            $failures += "npm run validate (office-addin-manifest validate) failed:`n$validateOutput"
        }

        # 2) the dev cert exists (installed by office-addin-dev-certs on first `npm start`/`npm run dev-server`)
        $certDir = Join-Path $env:USERPROFILE ".office-addin-dev-certs"
        if (-not (Test-Path $certDir)) {
            $failures += "Dev cert directory not found at $certDir - run 'npm start' or 'npx office-addin-dev-certs install' at least once first"
        }

        # 3) webpack-dev-server actually serves the taskpane over HTTPS on :3000
        if ($failures.Count -eq 0 -or (Test-Path $certDir)) {
            $job = Start-Job -ScriptBlock {
                param($root)
                Set-Location $root
                npm run dev-server
            } -ArgumentList $root

            try {
                $ok = $false
                for ($i = 0; $i -lt 30; $i++) {
                    Start-Sleep -Seconds 2
                    try {
                        # -SkipCertificateCheck is PowerShell 7+ only; this repo targets Windows PowerShell 5.1,
                        # where the dev cert installed by office-addin-dev-certs is already trusted, so a plain
                        # request works without needing to bypass certificate validation.
                        $resp = Invoke-WebRequest -Uri "https://localhost:3000/taskpane.html" -TimeoutSec 3 -UseBasicParsing
                        if ($resp.StatusCode -eq 200) { $ok = $true; break }
                    } catch { }
                }
                if (-not $ok) {
                    $failures += "https://localhost:3000/taskpane.html did not respond with 200 within 60s"
                }
            } finally {
                Stop-Job $job -ErrorAction SilentlyContinue | Out-Null
                Remove-Job $job -Force -ErrorAction SilentlyContinue | Out-Null
                # webpack-dev-server's child process can outlive the Job; make sure port 3000 is freed.
                Get-NetTCPConnection -LocalPort 3000 -ErrorAction SilentlyContinue |
                    ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }
            }

            # 4) the static server (the autostart serving path) serves the pre-built
            #    dist/ over HTTPS on :3000 - built by `npm run build`, served by
            #    `npm run serve-static`, reusing the same trusted dev cert as webpack.
            $buildOutput = npm run build 2>&1 | Out-String
            if ($LASTEXITCODE -ne 0) {
                $failures += "npm run build failed:`n$buildOutput"
            } elseif (-not (Test-Path (Join-Path $root "dist\taskpane.html"))) {
                $failures += "npm run build did not produce dist\taskpane.html"
            } else {
                $staticJob = Start-Job -ScriptBlock {
                    param($root)
                    Set-Location $root
                    npm run serve-static
                } -ArgumentList $root
                try {
                    $staticOk = $false
                    for ($i = 0; $i -lt 15; $i++) {
                        Start-Sleep -Seconds 1
                        try {
                            $resp = Invoke-WebRequest -Uri "https://localhost:3000/taskpane.html" -TimeoutSec 3 -UseBasicParsing
                            if ($resp.StatusCode -eq 200) { $staticOk = $true; break }
                        } catch { }
                    }
                    if (-not $staticOk) {
                        $failures += "npm run serve-static did not serve https://localhost:3000/taskpane.html (200) within 15s"
                    }
                } finally {
                    Stop-Job $staticJob -ErrorAction SilentlyContinue | Out-Null
                    Remove-Job $staticJob -Force -ErrorAction SilentlyContinue | Out-Null
                    Get-NetTCPConnection -LocalPort 3000 -ErrorAction SilentlyContinue |
                        ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }
                }
            }
        }
    }
} finally {
    Pop-Location
}

if ($failures.Count -eq 0) {
    Write-Host "PASS: Phase 2 dev-server checks (manifest validate, dev cert, HTTPS taskpane.html via webpack + static server)" -ForegroundColor Green
    exit 0
} else {
    Write-Host "FAIL: Phase 2 dev-server checks" -ForegroundColor Red
    $failures | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    exit 1
}
