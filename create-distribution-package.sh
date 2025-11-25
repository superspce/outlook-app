#!/bin/bash
# Create distribution package with extension and server launcher

set -e

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

PACKAGE_NAME="outlook-auto-attach-package"
VERSION="1.0.1"
DIST_DIR="dist-package"
TIMESTAMP=$(date +%Y%m%d-%H%M%S)

echo "📦 Creating distribution package..."
echo ""

# Clean previous package
rm -rf "$DIST_DIR"
rm -f "${PACKAGE_NAME}-*.zip"

# Create package structure
mkdir -p "$DIST_DIR/Chrome Extension"
mkdir -p "$DIST_DIR/Server/Mac"
mkdir -p "$DIST_DIR/Server/Windows"
mkdir -p "$DIST_DIR/Documentation"

# Copy Chrome extension
echo "📁 Copying Chrome extension..."
if [ -d "dist/extension" ]; then
    cp -r dist/extension/* "$DIST_DIR/Chrome Extension/"
else
    # Fallback to main directory
    cp manifest.json background.js popup.html popup.js confirm.html "$DIST_DIR/Chrome Extension/" 2>/dev/null || true
    cp -r icons "$DIST_DIR/Chrome Extension/" 2>/dev/null || true
fi

# Copy server launcher (Mac) - use .app bundle
echo "🍎 Copying Mac server launcher..."
if [ -d "server/dist/Outlook Auto Attach Server.app" ]; then
    cp -R "server/dist/Outlook Auto Attach Server.app" "$DIST_DIR/Server/Mac/"
    chmod -R +x "$DIST_DIR/Server/Mac/Outlook Auto Attach Server.app"
    echo "   ✅ Copied .app bundle"
elif [ -f "server/dist/Outlook Auto Attach Server" ]; then
    cp "server/dist/Outlook Auto Attach Server" "$DIST_DIR/Server/Mac/"
    chmod +x "$DIST_DIR/Server/Mac/Outlook Auto Attach Server"
elif [ -f "dist/server/outlook-attach-server" ]; then
    cp "dist/server/outlook-attach-server" "$DIST_DIR/Server/Mac/outlook-attach-server"
    chmod +x "$DIST_DIR/Server/Mac/outlook-attach-server"
fi

# Copy server launcher (Windows)
echo "🪟 Copying Windows server launcher..."
# Check for Windows files in dist/server first (user added files here)
if [ -f "dist/server/Outlook Auto Attach Server.exe" ] && [ -d "dist/server/_internal" ]; then
    # Create folder structure and copy Windows files
    mkdir -p "$DIST_DIR/Server/Windows/Outlook Auto Attach Server"
    cp "dist/server/Outlook Auto Attach Server.exe" "$DIST_DIR/Server/Windows/Outlook Auto Attach Server/"
    cp -R "dist/server/_internal" "$DIST_DIR/Server/Windows/Outlook Auto Attach Server/"
    echo "   ✅ Copied Windows .exe and _internal folder"
# Check if we have the full Windows folder structure (from GitHub Actions artifact)
elif [ -d "dist/server/Outlook Auto Attach Server" ]; then
    cp -R "dist/server/Outlook Auto Attach Server" "$DIST_DIR/Server/Windows/"
    echo "   ✅ Copied Windows folder with _internal"
elif [ -d "server/dist/Outlook Auto Attach Server" ]; then
    # Check if this is Windows (has .exe) or Mac (has .app-like structure)
    if [ -f "server/dist/Outlook Auto Attach Server/Outlook Auto Attach Server.exe" ]; then
        cp -R "server/dist/Outlook Auto Attach Server" "$DIST_DIR/Server/Windows/"
        echo "   ✅ Copied Windows folder with _internal"
    fi
elif [ -f "dist/server/Outlook Auto Attach Server.exe" ]; then
    # .exe file found - create folder structure
    mkdir -p "$DIST_DIR/Server/Windows/Outlook Auto Attach Server"
    cp "dist/server/Outlook Auto Attach Server.exe" "$DIST_DIR/Server/Windows/Outlook Auto Attach Server/"
    echo "   ✅ Copied .exe file"
    # Check for _internal folder in same directory
    if [ -d "dist/server/_internal" ]; then
        cp -R "dist/server/_internal" "$DIST_DIR/Server/Windows/Outlook Auto Attach Server/"
        echo "   ✅ Copied _internal folder"
    elif [ -d "dist/server/Outlook Auto Attach Server/_internal" ]; then
        cp -R "dist/server/Outlook Auto Attach Server/_internal" "$DIST_DIR/Server/Windows/Outlook Auto Attach Server/"
        echo "   ✅ Copied _internal folder"
    else
        echo "   ⚠️  WARNING: Missing _internal folder!"
        echo "   💡 Make sure _internal folder is in dist/server/ or dist/server/Outlook Auto Attach Server/"
    fi
elif [ -f "server/dist/Outlook Auto Attach Server.exe" ]; then
    mkdir -p "$DIST_DIR/Server/Windows/Outlook Auto Attach Server"
    cp "server/dist/Outlook Auto Attach Server.exe" "$DIST_DIR/Server/Windows/Outlook Auto Attach Server/"
    echo "   ⚠️  Copied .exe only - Windows users need _internal folder!"
elif [ -f "dist/server/outlook-attach-server.exe" ]; then
    cp "dist/server/outlook-attach-server.exe" "$DIST_DIR/Server/Windows/"
fi

# Create installation instructions
echo "📝 Creating documentation..."

cat > "$DIST_DIR/Documentation/INSTALLATION.md" << 'EOF'
# Outlook Auto Attach - Installation Guide

## Overview
This package contains:
1. **Chrome Extension** - Detects downloads and sends them to the server
2. **Server Application** - Opens Outlook and attaches files

## Installation Steps

### Step 1: Install Chrome Extension

1. Open Chrome and go to `chrome://extensions/`
2. Enable **Developer mode** (toggle in top right)
3. Click **Load unpacked**
4. Select the `Chrome Extension` folder from this package
5. The extension should now appear in your extensions list

### Step 2: Install Server Application

#### For Mac users:
1. Open the `Server/Mac` folder
2. Double-click `Outlook Auto Attach Server.app` to start it
   - **First time**: Mac may warn about unsigned app - right-click and select "Open", then click "Open" again
3. The server starts automatically - you can minimize or close the window
4. The server will keep running in the background

**Optional - Start on Login:**
1. Drag `Outlook Auto Attach Server.app` to your **Applications** folder (recommended)
2. Go to **System Settings → Users & Groups → Login Items** (or **System Preferences → Users & Groups → Login Items**)
3. Click **+** button
4. Select "Outlook Auto Attach Server" from Applications
5. ✅ Server will start automatically when you log in

#### For Windows users:
1. Open the `Server/Windows` folder
2. Double-click `Outlook Auto Attach Server.exe` to start it
3. The server window will open - click **Start Server**
4. You can minimize the window - the server will keep running

**Optional - Start on Login:**
1. Right-click `Outlook Auto Attach Server.exe`
2. Select "Create shortcut"
3. Press `Win + R`, type `shell:startup`, press Enter
4. Copy the shortcut to the Startup folder

## Usage

1. **Make sure the server is running**
   - Double-click `Outlook Auto Attach Server.app` (Mac) or `Outlook Auto Attach Server.exe` (Windows)
   - The server starts automatically - no need to click any buttons
   - You can verify it's running by checking http://localhost:8765/status in your browser (should show "Outlook Auto Attach Server is running")

2. **Download a file** that contains "Orderbekräftelse", "Inköp", or "1000322" in the filename

3. **The extension will detect it** and show a notification with a badge on the extension icon

4. **Click the extension icon** in Chrome to open the confirmation popup

5. **Click "Open Outlook"** - Outlook will open with the file attached!

**Note**: The server runs automatically when you start the app - you don't need to manually start it!

## Troubleshooting

### Server won't start
- Check if port 8765 is already in use
- Try stopping and restarting the server
- Make sure Microsoft Outlook is installed

### Extension not working
- Make sure the server is running (Status: Running)
- Check that you're downloading files with "Orderbekräftelse", "Inköp", or "1000322" in the name
- Check Chrome's extension console for errors (chrome://extensions/ → Details → Inspect views → Service worker)

### Outlook doesn't open
- Make sure Microsoft Outlook is installed
- Check the server log window for error messages
- Try restarting both the server and Chrome

## Support
For issues or questions, contact your IT department.
EOF

# Create README
cat > "$DIST_DIR/README.txt" << EOF
Outlook Auto Attach - Distribution Package
Version: ${VERSION}
Created: ${TIMESTAMP}

QUICK START:
1. Install Chrome Extension (see Documentation/INSTALLATION.md)
2. Run the Server Application for your OS (Mac or Windows)
3. Click "Start Server" in the server window
4. Download files with "Orderbekräftelse", "Inköp", or "1000322" in the name

For detailed instructions, see: Documentation/INSTALLATION.md
EOF

# Create ZIP package
echo ""
echo "📦 Creating ZIP package..."
cd "$DIST_DIR"
zip -r "../${PACKAGE_NAME}-${VERSION}-${TIMESTAMP}.zip" . -x "*.DS_Store" "*.git*"
cd ..

# Cleanup
rm -rf "$DIST_DIR"

echo ""
echo "✅ Package created successfully!"
echo "📦 Package: ${PACKAGE_NAME}-${VERSION}-${TIMESTAMP}.zip"
echo ""
echo "📋 Package contents:"
echo "   - Chrome Extension folder"
echo "   - Server/Mac folder (with executable)"
echo "   - Server/Windows folder (with executable)"
echo "   - Documentation/INSTALLATION.md"
echo ""

