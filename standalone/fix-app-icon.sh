#!/bin/bash
# Fix app icon after building

APP_PATH="dist/CT Food Outlook.app"
ICON_FILE="CTFood.icns"

if [ ! -d "$APP_PATH" ]; then
    echo "Error: $APP_PATH not found"
    exit 1
fi

if [ ! -f "$ICON_FILE" ]; then
    echo "Error: $ICON_FILE not found"
    echo "Creating icon from logo..."
    ./create-icon.sh
fi

echo "Fixing app icon..."

# Copy icon to app bundle
cp "$ICON_FILE" "$APP_PATH/Contents/Resources/"

# Update Info.plist
INFO_PLIST="$APP_PATH/Contents/Info.plist"

# Remove existing icon entry if it exists
/usr/libexec/PlistBuddy -c "Delete :CFBundleIconFile" "$INFO_PLIST" 2>/dev/null

# Add icon entry
/usr/libexec/PlistBuddy -c "Add :CFBundleIconFile string CTFood.icns" "$INFO_PLIST" 2>/dev/null

# Also set CFBundleIconName (required for some macOS versions)
/usr/libexec/PlistBuddy -c "Delete :CFBundleIconName" "$INFO_PLIST" 2>/dev/null
/usr/libexec/PlistBuddy -c "Add :CFBundleIconName string CTFood" "$INFO_PLIST" 2>/dev/null

echo "✅ Icon fixed!"
echo ""
echo "To see the new icon:"
echo "1. Quit the app if it's running"
echo "2. Clear icon cache: sudo killall Finder"
echo "3. Or restart your Mac"
