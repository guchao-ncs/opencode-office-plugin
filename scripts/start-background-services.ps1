# Idempotently starts the two background services AI Assistant needs (opencode
# serve + the webpack dev server) so Word can connect to the sideloaded add-in the
# moment it launches, instead of failing with "ADD-IN ERROR" because nothing was
# listening yet. Safe to run repeatedly (e.g. once per Windows logon via
# install-autostart.ps1): each service is skipped if its port is already listening.

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
    Write-Host "Starting opencode serve on port $opencodePort..."
    Start-Process -FilePath "cmd.exe" `
        -ArgumentList "/c", "cd /d `"$root`" && opencode serve --port $opencodePort --hostname 127.0.0.1 --cors https://localhost:$devServerPort" `
        -WindowStyle Hidden `
        -RedirectStandardOutput (Join-Path $logDir "opencode-serve.log") `
        -RedirectStandardError (Join-Path $logDir "opencode-serve.err.log")
}
