@echo off
REM Test version of install script that shows errors

setlocal enabledelayedexpansion

set APP_NAME=CT Food Outlook
set APP_EXE=outlook-auto-attach.exe
set INSTALL_DIR=%ProgramFiles%\%APP_NAME%
set START_MENU=%APPDATA%\Microsoft\Windows\Start Menu\Programs

echo Installing %APP_NAME%...
echo.
echo Current directory: %CD%
echo Executable: %APP_EXE%
echo.

REM Check if executable exists
if not exist "%APP_EXE%" (
    echo ERROR: %APP_EXE% not found in current directory!
    echo.
    echo Make sure you extracted the ZIP file first.
    pause
    exit /b 1
)

echo Creating installation directory...
if not exist "%INSTALL_DIR%" (
    mkdir "%INSTALL_DIR%"
    if errorlevel 1 (
        echo ERROR: Could not create directory. Run as Administrator!
        pause
        exit /b 1
    )
)

echo Copying files...
copy "%APP_EXE%" "%INSTALL_DIR%\" /Y
if errorlevel 1 (
    echo ERROR: Could not copy file. Run as Administrator!
    pause
    exit /b 1
)

echo Creating shortcuts...
set SCRIPT=%TEMP%\create_shortcut.vbs
(
echo Set oWS = WScript.CreateObject("WScript.Shell"^)
echo sLinkFile = "%START_MENU%\%APP_NAME%.lnk"
echo Set oLink = oWS.CreateShortcut(sLinkFile^)
echo oLink.TargetPath = "%INSTALL_DIR%\%APP_EXE%"
echo oLink.WorkingDirectory = "%INSTALL_DIR%"
echo oLink.Description = "%APP_NAME% - Auto attach Orderbekräftelse files to Outlook"
echo oLink.Save
) > "%SCRIPT%"

cscript /nologo "%SCRIPT%"
del "%SCRIPT%"

echo.
echo ✅ %APP_NAME% installed successfully!
echo.
echo The app is now available in Start Menu.
echo.
echo Launching application...
start "" "%INSTALL_DIR%\%APP_EXE%"

echo.
echo Press any key to close...
pause >nul

endlocal
