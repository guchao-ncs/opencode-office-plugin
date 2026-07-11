#!/bin/bash
# Installs a launchd Launch Agent plist to start background services on macOS logon.

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(dirname "$SCRIPT_DIR")"
PLIST_PATH="$HOME/Library/LaunchAgents/com.aiassistant.autostart.plist"
LOG_DIR="$HOME/Library/Logs/AIAssistant"

mkdir -p "$LOG_DIR"
mkdir -p "$(dirname "$PLIST_PATH")"

# Generate plist content dynamically
cat <<EOF > "$PLIST_PATH"
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.aiassistant.autostart</string>
    <key>ProgramArguments</key>
    <array>
        <string>/bin/bash</string>
        <string>$SCRIPT_DIR/start-background-services.sh</string>
    </array>
    <key>RunAtLoad</key>
    <true/>
    <key>StandardOutPath</key>
    <string>$LOG_DIR/launchd-autostart.log</string>
    <key>StandardErrorPath</key>
    <string>$LOG_DIR/launchd-autostart.err.log</string>
</dict>
</plist>
EOF

# Set permissions
chmod 644 "$PLIST_PATH"

# Load the Launch Agent immediately
launchctl load "$PLIST_PATH" 2>/dev/null || launchctl bootstrap gui/"$(id -u)" "$PLIST_PATH" 2>/dev/null

echo "Installed launchd autostart plist at: $PLIST_PATH"
echo "Background services will start automatically on next macOS logon."
echo "To start them right now manually, run: scripts/start-background-services.sh"
