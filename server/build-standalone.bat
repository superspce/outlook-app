@echo off
REM Build standalone executable for Windows using PyInstaller

echo Building standalone executable for Windows...

REM Check if PyInstaller is installed
pyinstaller --version >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
)

REM Build the executable
pyinstaller --onefile ^
    --name outlook-attach-server ^
    --clean ^
    --noconsole ^
    --add-data "outlook-attach-server.py;." ^
    outlook-attach-server.py

echo.
echo SUCCESS: Executable built: dist\outlook-attach-server.exe
echo.
echo You can now distribute this file - it doesn't require Python!
pause

