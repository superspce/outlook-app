#!/bin/bash
# Start the Outlook Auto Attach server

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
PYTHON_SCRIPT="$SCRIPT_DIR/outlook-attach-server.py"

echo "Starting Outlook Auto Attach server..."
python3 "$PYTHON_SCRIPT"

