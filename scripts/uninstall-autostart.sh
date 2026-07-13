#!/bin/bash
# Removes the launchd autostart plist file and unloads it.

PLIST_PATH="$HOME/Library/LaunchAgents/com.aiassistant.autostart.plist"

if [ -f "$PLIST_PATH" ]; then
  # Unload it first
  launchctl unload "$PLIST_PATH" 2>/dev/null || launchctl bootout gui/"$(id -u)" "$PLIST_PATH" 2>/dev/null
  rm -f "$PLIST_PATH"
  echo "Removed launchd autostart plist: $PLIST_PATH"
else
  echo "No autostart plist found at $PLIST_PATH - nothing to do"
fi
