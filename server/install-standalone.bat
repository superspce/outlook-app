@echo off
REM Install script for standalone executable (Windows) - no Python required

setlocal enabledelayedexpansion

set SCRIPT_DIR=%~dp0
set EXECUTABLE=%SCRIPT_DIR%outlook-attach-server.exe
set STARTUP_FOLDER=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
set SHORTCUT_NAME=Outlook Auto Attach Server.lnk

echo Installing Outlook Auto Attach server (standalone)...
echo.

REM Check if executable exists
if not exist "%EXECUTABLE%" (
    echo ERROR: outlook-attach-server.exe not found!
    echo Please build it first or use the Python version.
    pause
    exit /b 1
)

REM Create a VBScript to create the shortcut
set VBS_FILE=%TEMP%\create_shortcut.vbs
echo Set oWS = WScript.CreateObject("WScript.Shell") > "%VBS_FILE%"
echo sLinkFile = "%STARTUP_FOLDER%\%SHORTCUT_NAME%" >> "%VBS_FILE%"
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> "%VBS_FILE%"
echo oLink.TargetPath = "%EXECUTABLE%" >> "%VBS_FILE%"
echo oLink.WorkingDirectory = "%SCRIPT_DIR%" >> "%VBS_FILE%"
echo oLink.Description = "Outlook Auto Attach Server" >> "%VBS_FILE%"
echo oLink.Save >> "%VBS_FILE%"

REM Create the shortcut
cscript //nologo "%VBS_FILE%"
del "%VBS_FILE%"

REM Start the server now
echo Starting server now...
start "Outlook Auto Attach Server" "%EXECUTABLE%"

echo.
echo SUCCESS: Auto-start installed successfully!
echo.
echo The server will now start automatically when you log in.
echo No Python installation required!
echo.
pause

