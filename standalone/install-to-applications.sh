#!/bin/bash
# Simple script to install CT Food Outlook to Applications folder

APP_NAME="CT Food Outlook.app"
SOURCE="dist/$APP_NAME"
DEST="/Applications/$APP_NAME"

if [ ! -d "$SOURCE" ]; then
    echo "❌ Error: $APP_NAME not found in dist folder"
    echo "Please build the app first: ./build-standalone-mac.sh"
    exit 1
fi

echo "Installing $APP_NAME to Applications..."
cp -R "$SOURCE" "$DEST"

if [ -d "$DEST" ]; then
    echo "✅ Successfully installed to /Applications/"
    echo ""
    echo "You can now:"
    echo "  - Launch from Applications folder"
    echo "  - Launch from Spotlight (Cmd+Space, type 'CT Food Outlook')"
    echo "  - Launch from Launchpad"
else
    echo "❌ Installation failed!"
    exit 1
fi
