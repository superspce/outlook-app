@echo off
REM Start the Outlook Auto Attach server on Windows

cd /d "%~dp0"
echo Starting Outlook Auto Attach server...
python outlook-attach-server.py

