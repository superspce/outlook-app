@echo off
REM Install script for Windows Native Messaging Host

setlocal enabledelayedexpansion

set SCRIPT_DIR=%~dp0
set EXTENSION_ID=%1

if "%EXTENSION_ID%"=="" (
    echo Usage: %~nx0 ^<extension-id^>
    echo.
    echo To find your extension ID:
    echo 1. Open Chrome and go to chrome://extensions/
    echo 2. Enable 'Developer mode'
    echo 3. Find 'Outlook Auto Attach' extension
    echo 4. Copy the ID shown under the extension name
    exit /b 1
)

REM Determine the manifest location based on Chrome
set CHROME_MANIFEST_DIR=%LOCALAPPDATA%\Google\Chrome\User Data\NativeMessagingHosts

REM Create directory if it doesn't exist
if not exist "%CHROME_MANIFEST_DIR%" mkdir "%CHROME_MANIFEST_DIR%"

REM Create manifest file
set MANIFEST_FILE=%CHROME_MANIFEST_DIR%\com.outlookautoattach.host.json

REM Determine the path to the native host executable
REM Priority: 1) Standalone executable, 2) Python script with wrapper (fallback)

REM Check for standalone executable in dist folder (no Python needed)
if exist "%SCRIPT_DIR%dist\outlook-attach-native-host.exe" (
    set NATIVE_HOST_PATH=%SCRIPT_DIR%dist\outlook-attach-native-host.exe
    echo ✅ Found standalone executable (no Python required)
REM Check for standalone executable in script directory
) else if exist "%SCRIPT_DIR%outlook-attach-native-host.exe" (
    set NATIVE_HOST_PATH=%SCRIPT_DIR%outlook-attach-native-host.exe
    echo ✅ Found standalone executable in script directory
REM Fall back to Python script (requires Python) - for development only
) else if exist "%SCRIPT_DIR%outlook-attach-native-host.bat" (
    echo ⚠️  WARNING: Using Python script (Python required)
    echo    For production, build the executable first: build-native-host-windows.bat
    
    REM Check if Python is installed
    where python >nul 2>&1
    if %ERRORLEVEL% NEQ 0 (
        echo Error: Python is not installed
        echo Please install Python or build the standalone executable
        exit /b 1
    )
    
    REM Use the wrapper batch file
    set NATIVE_HOST_PATH=%SCRIPT_DIR%outlook-attach-native-host.bat
) else (
    echo Error: Could not find native host executable or Python script
    echo Please build the executable first: build-native-host-windows.bat
    exit /b 1
)

REM Escape backslashes for JSON (Windows paths need double backslashes)
set NATIVE_HOST_PATH_JSON=%NATIVE_HOST_PATH%
set NATIVE_HOST_PATH_JSON=!NATIVE_HOST_PATH_JSON:\=\\!

REM Create manifest
(
echo {
echo   "name": "com.outlookautoattach.host",
echo   "description": "Outlook Auto Attach Native Messaging Host",
echo   "path": "!NATIVE_HOST_PATH_JSON!",
echo   "type": "stdio",
echo   "allowed_origins": [
echo     "chrome-extension://%EXTENSION_ID%/"
echo   ]
echo }
) > "%MANIFEST_FILE%"

echo ✅ Native Messaging Host installed successfully!
echo    Manifest: %MANIFEST_FILE%
echo    Extension ID: %EXTENSION_ID%

endlocal

