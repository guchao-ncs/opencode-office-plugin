# Phase 3 test: verify the real OpenCode server (opencode serve) is reachable and
# exposes the REST surface this project depends on - /global/health and the
# OpenAPI spec at /doc. Starts its own server instance on a dedicated test port
# so it never collides with a server the developer may already be running.

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$failures = @()
$testPort = 4097

Push-Location $root
try {
    $opencodeCmd = Get-Command opencode -ErrorAction SilentlyContinue
    if (-not $opencodeCmd) {
        $failures += "'opencode' CLI not found on PATH"
    } else {
        $job = Start-Job -ScriptBlock {
            param($root, $port)
            Set-Location $root
            opencode serve --port $port --hostname 127.0.0.1
        } -ArgumentList $root, $testPort

        try {
            $healthy = $false
            for ($i = 0; $i -lt 20; $i++) {
                Start-Sleep -Seconds 1
                try {
                    $resp = Invoke-WebRequest -Uri "http://127.0.0.1:$testPort/global/health" -TimeoutSec 3 -UseBasicParsing
                    if ($resp.StatusCode -eq 200) {
                        $body = $resp.Content | ConvertFrom-Json
                        if ($body.healthy -eq $true) { $healthy = $true; break }
                    }
                } catch { }
            }
            if (-not $healthy) {
                $failures += "http://127.0.0.1:$testPort/global/health did not report healthy within 20s"
            } else {
                # /doc must return a valid OpenAPI 3.1 spec - this is the contract the
                # taskpane chat UI (Phase 5) will be coded against, so verify its shape.
                try {
                    $docResp = Invoke-WebRequest -Uri "http://127.0.0.1:$testPort/doc" -TimeoutSec 5 -UseBasicParsing
                    $doc = $docResp.Content | ConvertFrom-Json
                    if (-not $doc.openapi) {
                        $failures += "/doc response is missing 'openapi' field - not a valid OpenAPI spec"
                    }
                    if (-not $doc.paths.'/session') {
                        $failures += "/doc does not declare a '/session' path - OpenCode API surface may have changed"
                    }
                } catch {
                    $failures += "GET /doc failed: $($_.Exception.Message)"
                }
            }
        } finally {
            Stop-Job $job -ErrorAction SilentlyContinue | Out-Null
            Remove-Job $job -Force -ErrorAction SilentlyContinue | Out-Null
            Get-NetTCPConnection -LocalPort $testPort -ErrorAction SilentlyContinue |
                ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }
        }
    }
} finally {
    Pop-Location
}

if ($failures.Count -eq 0) {
    Write-Host "PASS: Phase 3 OpenCode server checks (CLI present, /global/health, /doc OpenAPI spec)" -ForegroundColor Green
    exit 0
} else {
    Write-Host "FAIL: Phase 3 OpenCode server checks" -ForegroundColor Red
    $failures | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    exit 1
}
