# Installs a silent, per-user, admin-free autostart entry that runs
# start-background-services.ps1 at every Windows logon, so the OpenCode
# backend + dev server are already up by the time Word tries to connect to
# the sideloaded add-in. Writes only to the current user's Startup folder -
# no system-wide changes, no elevation required. Re-run safely at any time;
# it just overwrites the existing wrapper.

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$startupDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Startup"
$vbsPath = Join-Path $startupDir "AI Assistant-autostart.vbs"
$scriptPath = Join-Path $root "scripts\start-background-services.ps1"

# The 0 window-style argument to WScript.Shell.Run suppresses any visible
# window (console flash or otherwise) - this is the standard way to run a
# script silently at logon without needing a scheduled task or admin rights.
$vbsContent = @"
Set shell = CreateObject("WScript.Shell")
shell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File ""$scriptPath""", 0, False
"@

Set-Content -Path $vbsPath -Value $vbsContent -Encoding ASCII
Write-Host "Installed autostart: $vbsPath"
Write-Host "Background services will start automatically on next Windows logon."
Write-Host "To start them right now without logging out, run: scripts\start-background-services.ps1"
