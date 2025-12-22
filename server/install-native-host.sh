#!/bin/bash
# Install script for macOS Native Messaging Host

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
EXTENSION_ID="${1:-}"

if [ -z "$EXTENSION_ID" ]; then
    echo "Usage: $0 <extension-id>"
    echo ""
    echo "To find your extension ID:"
    echo "1. Open Chrome and go to chrome://extensions/"
    echo "2. Enable 'Developer mode'"
    echo "3. Find 'Outlook Auto Attach' extension"
    echo "4. Copy the ID shown under the extension name"
    exit 1
fi

# Determine the manifest location based on Chrome/Chromium
if [ -d "$HOME/Library/Application Support/Google/Chrome" ]; then
    CHROME_MANIFEST_DIR="$HOME/Library/Application Support/Google/Chrome/NativeMessagingHosts"
elif [ -d "$HOME/Library/Application Support/Chromium" ]; then
    CHROME_MANIFEST_DIR="$HOME/Library/Application Support/Chromium/NativeMessagingHosts"
else
    echo "Error: Could not find Chrome or Chromium installation"
    echo "Trying to create directory anyway..."
    CHROME_MANIFEST_DIR="$HOME/Library/Application Support/Google/Chrome/NativeMessagingHosts"
fi

# Create directory if it doesn't exist
echo "Creating directory: $CHROME_MANIFEST_DIR"
mkdir -p "$CHROME_MANIFEST_DIR"

if [ ! -d "$CHROME_MANIFEST_DIR" ]; then
    echo "Error: Failed to create directory: $CHROME_MANIFEST_DIR"
    exit 1
fi

# Create manifest file
MANIFEST_FILE="$CHROME_MANIFEST_DIR/com.outlookautoattach.host.json"

# Determine the path to the native host executable
# Priority: 1) Standalone executable, 2) Python script with wrapper (fallback)

# Check for standalone executable (no Python needed)
if [ -f "$SCRIPT_DIR/dist/outlook-attach-native-host" ]; then
    NATIVE_HOST_PATH="$SCRIPT_DIR/dist/outlook-attach-native-host"
    echo "Found standalone executable (no Python required)"
elif [ -f "$SCRIPT_DIR/outlook-attach-native-host" ]; then
    NATIVE_HOST_PATH="$SCRIPT_DIR/outlook-attach-native-host"
    echo "Found standalone executable in script directory"
# Fall back to Python script (requires Python) - for development only
elif [ -f "$SCRIPT_DIR/outlook-attach-native-host.py" ]; then
    echo "WARNING: Using Python script (Python 3 required)"
    echo "For production, build the executable first: ./build-native-host-mac.sh"
    
    # Check if Python 3 is installed
    if ! command -v python3 &> /dev/null; then
        echo "Error: Python 3 is not installed"
        echo "Please install Python 3 or build the standalone executable"
        exit 1
    fi
    
    # Create a wrapper script for macOS (native messaging doesn't support args)
    WRAPPER_SCRIPT="$SCRIPT_DIR/outlook-attach-native-host-wrapper.sh"
    cat > "$WRAPPER_SCRIPT" <<'WRAPPER_EOF'
#!/bin/bash
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
exec python3 "$SCRIPT_DIR/outlook-attach-native-host.py"
WRAPPER_EOF
    chmod +x "$WRAPPER_SCRIPT"
    NATIVE_HOST_PATH="$WRAPPER_SCRIPT"
else
    echo "Error: Could not find native host executable or Python script"
    echo "Please build the executable first: ./build-native-host-mac.sh"
    exit 1
fi

# Remove existing file/symlink if it exists
if [ -e "$MANIFEST_FILE" ] || [ -L "$MANIFEST_FILE" ]; then
    echo "Removing existing manifest file/symlink..."
    rm -f "$MANIFEST_FILE"
fi

# Create manifest using Python to properly escape JSON
echo "Creating manifest file: $MANIFEST_FILE"
NATIVE_HOST_PATH="$NATIVE_HOST_PATH" EXTENSION_ID="$EXTENSION_ID" MANIFEST_FILE="$MANIFEST_FILE" python3 <<'PYTHON_EOF'
import json
import os

native_host_path = os.environ.get('NATIVE_HOST_PATH', '')
extension_id = os.environ.get('EXTENSION_ID', '')
manifest_file = os.environ.get('MANIFEST_FILE', '')

manifest = {
    "name": "com.outlookautoattach.host",
    "description": "Outlook Auto Attach Native Messaging Host",
    "path": native_host_path,
    "type": "stdio",
    "allowed_origins": [
        f"chrome-extension://{extension_id}/"
    ]
}

with open(manifest_file, "w") as f:
    json.dump(manifest, f, indent=2)
PYTHON_EOF

if [ ! -f "$MANIFEST_FILE" ]; then
    echo "Error: Failed to create manifest file: $MANIFEST_FILE"
    exit 1
fi

echo ""
echo "✅ Native Messaging Host installed successfully!"
echo "   Manifest: $MANIFEST_FILE"
echo "   Extension ID: $EXTENSION_ID"
echo "   Native Host: $NATIVE_HOST_PATH"

