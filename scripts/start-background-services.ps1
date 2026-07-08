# Idempotently starts the two background services AI Assistant needs (opencode
# serve + the webpack dev server) so Word can connect to the sideloaded add-in the
# moment it launches, instead of failing with "ADD-IN ERROR" because nothing was
# listening yet. Safe to run repeatedly (e.g. once per Windows logon via
# install-autostart.ps1): each service is skipped if its port is already listening.
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

# Reads the live word-MCP connection status off the running server ("connected",
# "failed", etc.), or $null if serve isn't answering (yet).
function Get-WordMcpStatus {
    try {
        $resp = Invoke-RestMethod -Uri "http://127.0.0.1:$opencodePort/mcp" -TimeoutSec 5 -ErrorAction Stop
        if ($resp.word -and $resp.word.status) { return [string]$resp.word.status }
        return $null
    } catch {
        return $null
    }
}

# Polls until the word MCP status reaches a settled state (connected/failed) or
# the timeout elapses - covers both serve still booting (no answer yet) and the
# MCP handshake still in flight. Returns the last status seen ($null if serve
# never answered at all).
function Wait-WordMcpSettled([int]$timeoutSeconds = 120) {
    $deadline = (Get-Date).AddSeconds($timeoutSeconds)
    do {
        $status = Get-WordMcpStatus
        if ($status -eq "connected" -or $status -eq "failed") { return $status }
        Start-Sleep -Seconds 5
    } while ((Get-Date) -lt $deadline)
    return Get-WordMcpStatus
}

if (Test-PortListening $devServerPort) {
    Write-Host "webpack dev server already listening on port $devServerPort - skipping"
} else {
    Write-Host "Starting webpack dev server on port $devServerPort..."
    Start-Process -FilePath "cmd.exe" `
        -ArgumentList "/c", "cd /d `"$root`" && npm run dev-server" `
        -WindowStyle Hidden `
        -RedirectStandardOutput (Join-Path $logDir "dev-server.log") `
        -RedirectStandardError (Join-Path $logDir "dev-server.err.log")
}

if (Test-PortListening $opencodePort) {
    Write-Host "opencode serve already listening on port $opencodePort - skipping"
} else {
    Start-OpenCodeServe
}

# --- Self-heal: make sure the word MCP connection actually came up ---
# Runs on every invocation (not just fresh starts), so an already-running serve
# whose MCP connection failed at logon gets repaired too.
$status = Wait-WordMcpSettled
if ($status -eq "failed") {
    Write-Host "word MCP connection is in 'failed' state - trying a runtime reconnect..."
    try {
        Invoke-RestMethod -Method Post -Uri "http://127.0.0.1:$opencodePort/mcp/word/connect" -TimeoutSec 120 -ErrorAction Stop | Out-Null
    } catch {
        # Best-effort - the settled re-check below decides what happens next.
    }
    $status = Wait-WordMcpSettled -timeoutSeconds 30
}
if ($status -eq "failed") {
    Write-Host "Reconnect did not recover it - restarting opencode serve..."
    Get-NetTCPConnection -LocalPort $opencodePort -State Listen -ErrorAction SilentlyContinue |
        ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Seconds 2
    Start-OpenCodeServe
    $status = Wait-WordMcpSettled
}

if ($status -eq "connected") {
    Write-Host "word MCP connection verified: connected"
} elseif ($null -eq $status) {
    Write-Host "warning: could not read MCP status from opencode serve on port $opencodePort - check $logDir\opencode-serve.err.log"
} else {
    Write-Host "warning: word MCP connection is still '$status' after reconnect and restart - Word editing tools will be unavailable. Check $logDir\opencode-serve.err.log"
}
