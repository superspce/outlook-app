#!/bin/bash
# Build standalone executable for macOS using PyInstaller

echo "Building standalone executable for macOS..."

# Check if PyInstaller is installed
if ! command -v pyinstaller &> /dev/null; then
    echo "Installing PyInstaller..."
    pip3 install pyinstaller
fi

# Build the executable
pyinstaller --onefile \
    --name outlook-attach-server \
    --clean \
    --noconsole \
    outlook-attach-server.py

echo ""
echo "✅ Executable built: dist/outlook-attach-server"
echo ""
echo "You can now distribute this file - it doesn't require Python!"

