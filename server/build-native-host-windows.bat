@echo off
REM Build standalone executable for Windows Native Messaging Host

setlocal

set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

echo Building Windows native host executable...

REM Check if PyInstaller is installed
where pyinstaller >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo PyInstaller not found. Installing...
    python -m pip install pyinstaller
)

REM Build the executable
pyinstaller --onefile ^
    --name outlook-attach-native-host ^
    --clean ^
    --noconfirm ^
    outlook-attach-native-host.py

if exist "dist\outlook-attach-native-host.exe" (
    echo.
    echo ✅ Build successful!
    echo    Executable: dist\outlook-attach-native-host.exe
    echo.
    echo The install script will automatically use this executable.
) else (
    echo ❌ Build failed!
    exit /b 1
)

endlocal

