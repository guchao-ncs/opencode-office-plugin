# Removes the autostart entry installed by install-autostart.ps1. Safe to run
# even if it was never installed (no-op in that case).

$ErrorActionPreference = "Stop"
$startupDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Startup"
$vbsPath = Join-Path $startupDir "AI Assistant-autostart.vbs"

if (Test-Path $vbsPath) {
    Remove-Item -Path $vbsPath -Force
    Write-Host "Removed autostart entry: $vbsPath"
} else {
    Write-Host "No autostart entry found at $vbsPath - nothing to do"
}
