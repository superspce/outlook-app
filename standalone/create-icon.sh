#!/bin/bash
# Create .icns file from PNG for macOS app icon

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

LOGO="ct_food_app_logo.png"
ICONSET="CTFood.iconset"

if [ ! -f "$LOGO" ]; then
    echo "Error: $LOGO not found"
    exit 1
fi

echo "Creating macOS icon set from $LOGO..."

# Create iconset directory
rm -rf "$ICONSET"
mkdir "$ICONSET"

# Create all required icon sizes
sips -z 16 16     "$LOGO" --out "$ICONSET/icon_16x16.png" > /dev/null 2>&1
sips -z 32 32     "$LOGO" --out "$ICONSET/icon_16x16@2x.png" > /dev/null 2>&1
sips -z 32 32     "$LOGO" --out "$ICONSET/icon_32x32.png" > /dev/null 2>&1
sips -z 64 64     "$LOGO" --out "$ICONSET/icon_32x32@2x.png" > /dev/null 2>&1
sips -z 128 128   "$LOGO" --out "$ICONSET/icon_128x128.png" > /dev/null 2>&1
sips -z 256 256   "$LOGO" --out "$ICONSET/icon_128x128@2x.png" > /dev/null 2>&1
sips -z 256 256   "$LOGO" --out "$ICONSET/icon_256x256.png" > /dev/null 2>&1
sips -z 512 512   "$LOGO" --out "$ICONSET/icon_256x256@2x.png" > /dev/null 2>&1
sips -z 512 512   "$LOGO" --out "$ICONSET/icon_512x512.png" > /dev/null 2>&1
sips -z 1024 1024 "$LOGO" --out "$ICONSET/icon_512x512@2x.png" > /dev/null 2>&1

# Create .icns file
iconutil -c icns "$ICONSET" -o "CTFood.icns"

# Clean up
rm -rf "$ICONSET"

if [ -f "CTFood.icns" ]; then
    echo "✅ Created CTFood.icns"
else
    echo "❌ Failed to create icon"
    exit 1
fi
