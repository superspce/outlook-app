#!/bin/bash
# Create a distributable .dmg file for macOS

APP_NAME="CT Food Outlook"
APP_PATH="dist/${APP_NAME}.app"
DMG_NAME="${APP_NAME}.dmg"
DMG_PATH="dist/${DMG_NAME}"
VOLUME_NAME="${APP_NAME}"

if [ ! -d "$APP_PATH" ]; then
    echo "Error: $APP_PATH not found"
    echo "Please build the app first: ./build-standalone-mac.sh"
    exit 1
fi

echo "Creating DMG installer..."

# Fix icon first
./fix-app-icon.sh

# Create a temporary directory for DMG contents
TEMP_DIR=$(mktemp -d)
DMG_CONTENTS="$TEMP_DIR/${VOLUME_NAME}"

# Copy app to temp directory
cp -R "$APP_PATH" "$DMG_CONTENTS/"

# Create Applications symlink
ln -s /Applications "$DMG_CONTENTS/Applications"

# Create a README
cat > "$DMG_CONTENTS/README.txt" <<EOF
CT Food Outlook - Installation

1. Drag "CT Food Outlook.app" to the Applications folder
2. Open Applications and launch "CT Food Outlook"
3. The app will appear in your menu bar (top-right)

That's it! The app will automatically monitor your Downloads folder.

For more information, visit the app menu bar icon.
EOF

# Remove old DMG if it exists
rm -f "$DMG_PATH"

# Create DMG
hdiutil create -volname "$VOLUME_NAME" \
    -srcfolder "$DMG_CONTENTS" \
    -ov -format UDZO \
    "$DMG_PATH" \
    -fs HFS+ \
    -fsargs "-c c=64,a=16,e=16" \
    -imagekey zlib-level=9

# Clean up
rm -rf "$TEMP_DIR"

if [ -f "$DMG_PATH" ]; then
    echo ""
    echo "✅ DMG created successfully!"
    echo "   File: $DMG_PATH"
    echo ""
    echo "You can now:"
    echo "  - Upload this .dmg to a file sharing service"
    echo "  - Share the download link with coworkers"
    echo "  - Users download, open, and drag to Applications"
else
    echo "❌ DMG creation failed!"
    exit 1
fi
