#!/bin/bash
# Start the standalone server executable (macOS)

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
EXECUTABLE="$SCRIPT_DIR/outlook-attach-server"

if [ ! -f "$EXECUTABLE" ]; then
    echo "Error: outlook-attach-server executable not found!"
    echo "Please build it first using: ./build-standalone.sh"
    exit 1
fi

echo "Starting Outlook Auto Attach server..."
"$EXECUTABLE"

