@echo off
REM Start the standalone server executable (Windows)

cd /d "%~dp0"

if not exist "outlook-attach-server.exe" (
    echo Error: outlook-attach-server.exe not found!
    echo Please build it first using: build-standalone.bat
    pause
    exit /b 1
)

echo Starting Outlook Auto Attach server...
outlook-attach-server.exe

