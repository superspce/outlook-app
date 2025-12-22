#!/bin/bash
# Build standalone executable for macOS Native Messaging Host

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "Building macOS native host executable..."

# Check if PyInstaller is installed
if ! command -v pyinstaller &> /dev/null; then
    echo "PyInstaller not found. Installing..."
    python3 -m pip install pyinstaller --break-system-packages || python3 -m pip install --user pyinstaller
fi

# Build the executable
pyinstaller --onefile \
    --name outlook-attach-native-host \
    --clean \
    --noconfirm \
    outlook-attach-native-host.py

if [ -f "dist/outlook-attach-native-host" ]; then
    echo ""
    echo "✅ Build successful!"
    echo "   Executable: dist/outlook-attach-native-host"
    echo ""
    echo "The install script will automatically use this executable."
else
    echo "❌ Build failed!"
    exit 1
fi

