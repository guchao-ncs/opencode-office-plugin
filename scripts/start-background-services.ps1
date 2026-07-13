# Idempotently starts the two background services AI Assistant needs (opencode
# serve + a static HTTPS server for the pre-built taskpane in dist/) so Word can
# connect to the sideloaded add-in the moment it launches, instead of failing
# with "ADD-IN ERROR" because nothing was listening yet. Safe to run repeatedly
# (e.g. once per Windows logon via install-autostart.ps1): each service is
# skipped if its port is already listening.
#
# After serve is up, this also verifies the `word` MCP connection actually came
# up. opencode marks an MCP server that failed its startup handshake as "failed"
# for the life of the process (observed after a logon-time cold start, where uvx
# lost the race against every other startup process and timed out) - the agent
# then has no Word tools at all and can only apologize. When that happens, this
# script tries a runtime reconnect (POST /mcp/word/connect), and if that doesn't
# recover it, restarts opencode serve once - a fresh startup handshake has been
# reliable when the machine isn't mid-logon.

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$devServerPort = 3000
$opencodePort = 4098

$logDir = Join-Path $env:LOCALAPPDATA "AI Assistant\logs"
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

function Test-PortListening($port) {
    return [bool](Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue)
}

function Start-OpenCodeServe {
    Write-Host "Starting opencode serve on port $opencodePort..."
    Start-Process -FilePath "cmd.exe" `
        -ArgumentList "/c", "cd /d `"$root`" && opencode serve --port $opencodePort --hostname 127.0.0.1 --cors https://localhost:$devServerPort" `
        -WindowStyle Hidden `
        -RedirectStandardOutput (Join-Path $logDir "opencode-serve.log") `
        -RedirectStandardError (Join-Path $logDir "opencode-serve.err.log")
}

# Reads the live MCP connection status off the running server ("connected",
# "failed", etc.) for a given server name, or $null if serve isn't answering.
function Get-McpStatus($name) {
    try {
        $resp = Invoke-RestMethod -Uri "http://127.0.0.1:$opencodePort/mcp" -TimeoutSec 5 -ErrorAction Stop
        if ($resp.$name -and $resp.$name.status) { return [string]$resp.$name.status }
        return $null
    } catch {
        return $null
    }
}

# Polls until the MCP status reaches a settled state (connected/failed) or
# the timeout elapses. Returns the last status seen.
function Wait-McpSettled($name, [int]$timeoutSeconds = 120) {
    $deadline = (Get-Date).AddSeconds($timeoutSeconds)
    do {
        $status = Get-McpStatus $name
        if ($status -eq "connected" -or $status -eq "failed") { return $status }
        Start-Sleep -Seconds 5
    } while ((Get-Date) -lt $deadline)
    return Get-McpStatus $name
}

if (Test-PortListening $devServerPort) {
    Write-Host "static server already listening on port $devServerPort - skipping"
} else {
    # Serve the pre-built bundle (dist/) rather than running `webpack serve`,
    # which recompiles on every launch and left port 3000 unserved for 1-2
    # minutes after a cold boot (Word then failed with "ADD-IN ERROR"). The
    # build is normally produced ahead of time by install-autostart.ps1 (and
    # whenever the source changes); only if dist/ is somehow missing do we
    # build here as a self-healing fallback, accepting the one-time delay.
    if (-not (Test-Path (Join-Path $root "dist\taskpane.html"))) {
        Write-Host "dist/ not built yet - building once (npm run build)..."
        Push-Location $root
        try {
            & npm run build *> (Join-Path $logDir "build.log")
        } finally {
            Pop-Location
        }
    }
    Write-Host "Starting static server (serving dist/) on port $devServerPort..."
    Start-Process -FilePath "cmd.exe" `
        -ArgumentList "/c", "cd /d `"$root`" && npm run serve-static" `
        -WindowStyle Hidden `
        -RedirectStandardOutput (Join-Path $logDir "static-server.log") `
        -RedirectStandardError (Join-Path $logDir "static-server.err.log")
}

if (Test-PortListening $opencodePort) {
    Write-Host "opencode serve already listening on port $opencodePort - skipping"
} else {
    Start-OpenCodeServe
}

# --- Self-heal: make sure the MCP connections actually came up ---
# Runs on every invocation (not just fresh starts), so an already-running serve
# whose MCP connection failed at logon gets repaired too.
$services = @("word", "excel")
$failedServices = @()

foreach ($srv in $services) {
    $status = Wait-McpSettled $srv
    if ($status -eq "failed") {
        Write-Host "$srv MCP connection is in 'failed' state - trying a runtime reconnect..."
        try {
            Invoke-RestMethod -Method Post -Uri "http://127.0.0.1:$opencodePort/mcp/$srv/connect" -TimeoutSec 120 -ErrorAction Stop | Out-Null
        } catch {
            # Best-effort
        }
        $status = Wait-McpSettled $srv -timeoutSeconds 30
    }
    if ($status -ne "connected") {
        $failedServices += $srv
    } else {
        Write-Host "$srv MCP connection verified: connected"
    }
}

if ($failedServices.Count -gt 0) {
    Write-Host "Some MCP services ($failedServices) are still failed - restarting opencode serve..."
    Get-NetTCPConnection -LocalPort $opencodePort -State Listen -ErrorAction SilentlyContinue |
        ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Seconds 2
    Start-OpenCodeServe
    
    foreach ($srv in $services) {
        $status = Wait-McpSettled $srv
        if ($status -eq "connected") {
            Write-Host "$srv MCP connection verified: connected"
        } else {
            Write-Host "warning: $srv MCP connection is still '$status' - check $logDir\opencode-serve.err.log"
        }
    }
}
