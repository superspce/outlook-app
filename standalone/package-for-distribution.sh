#!/bin/bash
# Package both Mac and Windows versions for online distribution

echo "Packaging CT Food Outlook for distribution..."
echo ""

# Check if apps are built
MAC_APP="dist/CT Food Outlook.app"
WIN_ZIP="CT-Food-Outlook-Windows.zip"

if [ ! -d "$MAC_APP" ]; then
    echo "⚠️  macOS app not found. Building..."
    ./build-standalone-mac.sh
fi

if [ ! -f "$WIN_ZIP" ]; then
    echo "⚠️  Windows package not found."
    echo "   Please build Windows version on a Windows machine:"
    echo "   build-standalone.bat"
    echo "   create-windows-installer.bat"
fi

# Create DMG for Mac
if [ -d "$MAC_APP" ]; then
    echo "Creating macOS DMG..."
    ./create-dmg.sh
    MAC_DMG="dist/CT Food Outlook.dmg"
    if [ -f "$MAC_DMG" ]; then
        echo "✅ macOS package ready: $MAC_DMG"
    fi
fi

echo ""
echo "📦 Distribution packages ready!"
echo ""
echo "For macOS:"
echo "  - Upload: dist/CT Food Outlook.dmg"
echo ""
echo "For Windows:"
echo "  - Upload: CT-Food-Outlook-Windows.zip"
echo ""
echo "Upload options:"
echo "  1. Google Drive - Upload files, right-click → Get link → Anyone with link"
echo "  2. Dropbox - Upload files, right-click → Share → Copy link"
echo "  3. OneDrive - Upload files, right-click → Share → Copy link"
echo "  4. Your website - Upload to web server"
echo ""
echo "Then share the download links with coworkers!"
