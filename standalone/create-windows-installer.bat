@echo off
REM Create a Windows installer package

setlocal enabledelayedexpansion

set APP_NAME=CT Food Outlook
set APP_EXE=outlook-auto-attach.exe
set INSTALLER_DIR=dist-installer

if not exist "dist\%APP_EXE%" (
    echo Error: %APP_EXE% not found in dist folder
    echo Please build the executable first using build-standalone.bat
    exit /b 1
)

echo Creating Windows installer package...
echo.

REM Create installer directory
if exist "%INSTALLER_DIR%" rmdir /s /q "%INSTALLER_DIR%"
mkdir "%INSTALLER_DIR%"

REM Copy executable
copy "dist\%APP_EXE%" "%INSTALLER_DIR%\%APP_EXE%"

REM Copy the pre-made install script (simpler and tested)
if exist "Install.bat" (
    copy "Install.bat" "%INSTALLER_DIR%\Install.bat"
) else (
REM Create a simple installer script
(
echo @echo off
echo REM %APP_NAME% - Simple Installer
echo REM IMPORTANT: Right-click this file and select "Run as Administrator"
echo setlocal enabledelayedexpansion
echo.
echo echo Installing %APP_NAME%...
echo echo.
echo.
echo REM Check if running as Administrator
echo net session ^>nul 2^>^&1
echo if %%errorlevel%% neq 0 ^(
echo     echo ERROR: This script must be run as Administrator!
echo     echo.
echo     echo Right-click this file and select "Run as Administrator"
echo     pause
echo     exit /b 1
echo ^)
echo.
echo REM Create installation directory
echo set INSTALL_DIR=%%ProgramFiles%%\%APP_NAME%
echo if not exist "%%INSTALL_DIR%%" ^(
echo     echo Creating installation directory...
echo     mkdir "%%INSTALL_DIR%%"
echo     if errorlevel 1 ^(
echo         echo ERROR: Could not create directory. Make sure you ran as Administrator!
echo         pause
echo         exit /b 1
echo     ^)
echo ^)
echo.
echo REM Check if executable exists
echo if not exist "%APP_EXE%" ^(
echo     echo ERROR: %APP_EXE% not found!
echo     echo Make sure you extracted the ZIP file first.
echo     pause
echo     exit /b 1
echo ^)
echo.
echo REM Copy files
echo echo Copying files...
echo copy "%APP_EXE%" "%%INSTALL_DIR%%\" /Y
echo if errorlevel 1 ^(
echo     echo ERROR: Could not copy file. Make sure you ran as Administrator!
echo     pause
echo     exit /b 1
echo ^)
echo.
echo REM Create Start Menu shortcut
echo set SCRIPT=%%TEMP%%\create_shortcut.vbs
echo ^(echo Set oWS = WScript.CreateObject^("WScript.Shell"^)^) ^> "%%SCRIPT%%"
echo ^(echo sLinkFile = "%%APPDATA%%\Microsoft\Windows\Start Menu\Programs\%APP_NAME%.lnk"^) ^>> "%%SCRIPT%%"
echo ^(echo Set oLink = oWS.CreateShortcut^(sLinkFile^)^) ^>> "%%SCRIPT%%"
echo ^(echo oLink.TargetPath = "%%INSTALL_DIR%%\%APP_EXE%%"^) ^>> "%%SCRIPT%%"
echo ^(echo oLink.WorkingDirectory = "%%INSTALL_DIR%%"^) ^>> "%%SCRIPT%%"
echo ^(echo oLink.Description = "%APP_NAME% - Auto attach Orderbekräftelse files to Outlook"^) ^>> "%%SCRIPT%%"
echo ^(echo oLink.Save^) ^>> "%%SCRIPT%%"
echo cscript /nologo "%%SCRIPT%%" ^>nul
echo del "%%SCRIPT%%" ^>nul
echo.
echo REM Create Desktop shortcut
echo ^(echo Set oWS = WScript.CreateObject^("WScript.Shell"^)^) ^> "%%SCRIPT%%"
echo ^(echo sLinkFile = "%%USERPROFILE%%\Desktop\%APP_NAME%.lnk"^) ^>> "%%SCRIPT%%"
echo ^(echo Set oLink = oWS.CreateShortcut^(sLinkFile^)^) ^>> "%%SCRIPT%%"
echo ^(echo oLink.TargetPath = "%%INSTALL_DIR%%\%APP_EXE%%"^) ^>> "%%SCRIPT%%"
echo ^(echo oLink.WorkingDirectory = "%%INSTALL_DIR%%"^) ^>> "%%SCRIPT%%"
echo ^(echo oLink.Save^) ^>> "%%SCRIPT%%"
echo cscript /nologo "%%SCRIPT%%" ^>nul
echo del "%%SCRIPT%%" ^>nul
echo.
echo echo.
echo echo ✅ %APP_NAME% installed successfully!
echo echo.
echo echo The app is now available in:
echo echo   - Start Menu: %APP_NAME%
echo echo   - Desktop shortcut
echo echo.
echo REM Launch the application
echo echo.
echo echo Launching application...
echo start "" "%%INSTALL_DIR%%\%APP_EXE%%"
echo.
echo echo.
echo echo ✅ Installation complete!
echo echo.
echo echo The app is now available in:
echo echo   - Start Menu: %APP_NAME%
echo echo   - Desktop shortcut
echo echo.
echo echo Press any key to close...
echo pause ^>nul
echo endlocal
) > "%INSTALLER_DIR%\Install.bat"
)

REM Create README
(
echo %APP_NAME% - Installation Package
echo ======================================
echo.
echo INSTALLATION:
echo ------------
echo 1. Right-click "Install.bat" and select "Run as Administrator"
echo 2. The application will be installed and launched automatically
echo 3. A shortcut will be created on your Desktop and in Start Menu
echo.
echo The app runs in the system tray (bottom-right corner).
echo Right-click the icon to view log or quit.
) > "%INSTALLER_DIR%\README.txt"

echo ✅ Installer package created in: %INSTALLER_DIR%
echo.

REM Create zip file for distribution
echo Creating distributable ZIP file...
set ZIP_NAME=CT-Food-Outlook-Windows.zip
if exist "%ZIP_NAME%" del "%ZIP_NAME%"

REM Use PowerShell to create zip (available on Windows 10+)
echo Compressing files...
powershell -Command "Compress-Archive -Path '%INSTALLER_DIR%\*' -DestinationPath '%ZIP_NAME%' -Force"

if exist "%ZIP_NAME%" (
    echo.
    echo ✅ Created: %ZIP_NAME%
    echo.
    echo Package ready for distribution!
    echo.
    echo To distribute online:
    echo   1. Upload "%ZIP_NAME%" to:
    echo      - Google Drive (share link)
    echo      - Dropbox (share link)
    echo      - OneDrive (share link)
    echo      - Your website/server
    echo   2. Share the download link with coworkers
    echo   3. Users click the link, download, extract, and run Install.bat
    echo.
) else (
    echo.
    echo ⚠️  Could not create ZIP automatically
    echo    Please manually zip the "%INSTALLER_DIR%" folder
    echo    Name it: CT-Food-Outlook-Windows.zip
    echo.
    echo    Or run this command:
    echo    powershell -Command "Compress-Archive -Path 'dist-installer\*' -DestinationPath 'CT-Food-Outlook-Windows.zip' -Force"
    echo.
)

echo Package contents:
echo   - %APP_EXE% (the application)
echo   - Install.bat (installation script)
echo   - README.txt (instructions)
echo.

endlocal
