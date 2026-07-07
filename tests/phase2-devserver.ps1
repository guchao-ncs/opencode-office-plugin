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
        }
    }
} finally {
    Pop-Location
}

if ($failures.Count -eq 0) {
    Write-Host "PASS: Phase 2 dev-server checks (manifest validate, dev cert, HTTPS taskpane.html)" -ForegroundColor Green
    exit 0
} else {
    Write-Host "FAIL: Phase 2 dev-server checks" -ForegroundColor Red
    $failures | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    exit 1
}
