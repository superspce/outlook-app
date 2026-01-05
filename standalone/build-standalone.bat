@echo off
REM Build standalone Windows executable for Outlook Auto Attach

setlocal

set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

echo Building Outlook Auto Attach standalone executable...
echo.

REM Check if PyInstaller is installed
where pyinstaller >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo PyInstaller not found. Installing...
    python -m pip install pyinstaller
)

REM Install dependencies
echo Installing dependencies...
python -m pip install -r requirements.txt
python -m pip install pywin32

REM Build the executable
echo.
echo Building executable...
REM Check if logo exists and create icon
if exist "ct_food_app_logo.png" (
    echo Creating Windows icon from logo...
    REM Note: For Windows, you'd need to convert PNG to ICO
    REM For now, we'll just include the PNG and use it in the app
)

pyinstaller --onefile ^
    --name outlook-auto-attach ^
    --clean ^
    --noconfirm ^
    --add-data "requirements.txt;." ^
    --add-data "ct_food_app_logo.png;." ^
    --hidden-import win32timezone ^
    --hidden-import pystray._win32 ^
    --hidden-import platform ^
    outlook-auto-attach-standalone.py

if exist "dist\outlook-auto-attach.exe" (
    echo.
    echo ✅ Build successful!
    echo    Executable: dist\outlook-auto-attach.exe
    echo.
    echo The application is ready to distribute.
    echo Users can double-click the .exe file to run it.
    echo It will appear in the system tray (bottom-right corner).
) else (
    echo ❌ Build failed!
    exit /b 1
)

endlocal
