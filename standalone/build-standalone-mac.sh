#!/bin/bash
# Build standalone macOS executable for Outlook Auto Attach

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "Building CT Food Outlook standalone executable for macOS..."
echo ""

# Check if PyInstaller is installed
if ! command -v pyinstaller &> /dev/null; then
    echo "PyInstaller not found. Installing..."
    python3 -m pip install pyinstaller
fi

# Install dependencies
echo "Installing dependencies..."
python3 -m pip install --break-system-packages -r requirements.txt 2>/dev/null || python3 -m pip install --user -r requirements.txt

# Create icon if it doesn't exist
if [ ! -f "CTFood.icns" ] && [ -f "ct_food_app_logo.png" ]; then
    echo "Creating app icon..."
    ./create-icon.sh
fi

# Build the executable as a macOS .app bundle
echo ""
echo "Building macOS application bundle..."
ICON_FLAG=""
if [ -f "CTFood.icns" ]; then
    ICON_FLAG="--icon=CTFood.icns"
    echo "Using custom icon: CTFood.icns"
fi

pyinstaller --windowed \
    --name "CT Food Outlook" \
    --clean \
    --noconfirm \
    --add-data "requirements.txt:." \
    --add-data "ct_food_app_logo.png:." \
    --hidden-import pystray._darwin \
    $ICON_FLAG \
    outlook-auto-attach-standalone.py

if [ -d "dist/CT Food Outlook.app" ]; then
    echo ""
    echo "✅ Build successful!"
    echo ""
    
    # Ask if user wants to install to Applications
    read -p "Install to Applications folder? (y/n): " -n 1 -r
    echo
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        echo "Installing to Applications..."
        cp -R "dist/CT Food Outlook.app" "/Applications/CT Food Outlook.app"
        echo "✅ Installed to /Applications/CT Food Outlook.app"
        echo ""
        echo "You can now launch it from Applications or Spotlight!"
    else
        echo "Application built at: dist/CT Food Outlook.app"
        echo "To install manually, copy it to /Applications/"
    fi
    
    echo ""
    echo "To create a distributable DMG installer:"
    echo "  ./create-dmg.sh"
else
    echo "❌ Build failed!"
    exit 1
fi
