#!/bin/bash
# Install script for standalone executable (macOS) - no Python required

set -e

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
EXECUTABLE="$SCRIPT_DIR/outlook-attach-server"
PLIST_FILE="$HOME/Library/LaunchAgents/com.outlookautoattach.server.plist"

echo "Installing Outlook Auto Attach server (standalone)..."
echo ""

# Check if executable exists
if [ ! -f "$EXECUTABLE" ]; then
    echo "❌ Error: outlook-attach-server executable not found!"
    echo "Please build it first or use the Python version."
    exit 1
fi

# Make executable
chmod +x "$EXECUTABLE"

# Get absolute path to executable
EXECUTABLE_ABS=$(cd "$SCRIPT_DIR" && pwd)/outlook-attach-server

# Create LaunchAgent plist
echo "Creating LaunchAgent..."
cat > "$PLIST_FILE" << EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.outlookautoattach.server</string>
    <key>ProgramArguments</key>
    <array>
        <string>$EXECUTABLE_ABS</string>
    </array>
    <key>RunAtLoad</key>
    <true/>
    <key>KeepAlive</key>
    <true/>
    <key>StandardOutPath</key>
    <string>$HOME/Library/Logs/outlook-auto-attach-server.log</string>
    <key>StandardErrorPath</key>
    <string>$HOME/Library/Logs/outlook-auto-attach-server-error.log</string>
</dict>
</plist>
EOF

# Load the LaunchAgent
echo "Loading LaunchAgent..."
launchctl load "$PLIST_FILE" 2>/dev/null || launchctl bootstrap gui/$(id -u) "$PLIST_FILE" 2>/dev/null || true

# Start it now
echo "Starting server now..."
launchctl start com.outlookautoattach.server 2>/dev/null || true

echo ""
echo "✅ Auto-start installed successfully!"
echo ""
echo "The server will now start automatically when you log in."
echo "No Python installation required!"

