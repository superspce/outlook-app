@echo off
REM Simple installer for CT Food Outlook
REM Right-click and "Run as Administrator"

setlocal

set APP_NAME=CT Food Outlook
set APP_EXE=outlook-auto-attach.exe
set INSTALL_DIR=%ProgramFiles%\%APP_NAME%

echo ========================================
echo %APP_NAME% - Installation
echo ========================================
echo.

REM Check admin
net session >nul 2>&1
if errorlevel 1 (
    echo ERROR: Must run as Administrator!
    echo.
    echo Right-click this file and select "Run as Administrator"
    echo.
    pause
    exit /b 1
)

REM Check if exe exists
if not exist "%APP_EXE%" (
    echo ERROR: %APP_EXE% not found!
    echo Make sure you extracted the ZIP file.
    echo.
    pause
    exit /b 1
)

echo Installing to: %INSTALL_DIR%
echo.

REM Create directory
if not exist "%INSTALL_DIR%" (
    echo Creating directory...
    mkdir "%INSTALL_DIR%"
    if errorlevel 1 (
        echo ERROR: Could not create directory
        pause
        exit /b 1
    )
)

REM Copy file
echo Copying application...
copy "%APP_EXE%" "%INSTALL_DIR%\" /Y
if errorlevel 1 (
    echo ERROR: Could not copy file
    pause
    exit /b 1
)

REM Create shortcut
echo Creating Start Menu shortcut...
set SCRIPT=%TEMP%\shortcut.vbs
(
echo Set oWS = WScript.CreateObject("WScript.Shell"^)
echo sLinkFile = "%APPDATA%\Microsoft\Windows\Start Menu\Programs\%APP_NAME%.lnk"
echo Set oLink = oWS.CreateShortcut(sLinkFile^)
echo oLink.TargetPath = "%INSTALL_DIR%\%APP_EXE%"
echo oLink.WorkingDirectory = "%INSTALL_DIR%"
echo oLink.Description = "%APP_NAME%"
echo oLink.Save
) > "%SCRIPT%"

cscript /nologo "%SCRIPT%"
del "%SCRIPT%"

echo.
echo ✅ Installation successful!
echo.
echo The app is now in Start Menu: %APP_NAME%
echo.
echo Launching application...
start "" "%INSTALL_DIR%\%APP_EXE%"

echo.
echo Press any key to close...
pause >nul

endlocal
