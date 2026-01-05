@echo off
REM Create installer for CT Food Outlook application
REM This creates a proper Windows installation

setlocal enabledelayedexpansion

set APP_NAME=CT Food Outlook
set APP_EXE=outlook-auto-attach.exe
set INSTALL_DIR=%ProgramFiles%\%APP_NAME%
set START_MENU=%APPDATA%\Microsoft\Windows\Start Menu\Programs
set STARTUP_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup

echo ================================================
echo CT Food Outlook - Installer Creator
echo ================================================
echo.

REM Check if executable exists
if not exist "dist\%APP_EXE%" (
    echo Error: %APP_EXE% not found in dist folder
    echo Please build the executable first using build-standalone.bat
    exit /b 1
)

echo Creating installer package...
echo.

REM Create installer directory
set INSTALLER_DIR=installer
if exist "%INSTALLER_DIR%" rmdir /s /q "%INSTALLER_DIR%"
mkdir "%INSTALLER_DIR%"

REM Copy executable
copy "dist\%APP_EXE%" "%INSTALLER_DIR%\%APP_EXE%"

REM Create install script
(
echo @echo off
echo REM CT Food Outlook - Installation Script
echo setlocal
echo.
echo set APP_NAME=%APP_NAME%
echo set APP_EXE=%APP_EXE%
echo set INSTALL_DIR=%%ProgramFiles%%\%%APP_NAME%%
echo set START_MENU=%%APPDATA%%\Microsoft\Windows\Start Menu\Programs
echo.
echo echo Installing %%APP_NAME%%...
echo echo.
echo.
echo REM Create installation directory
echo if not exist "%%INSTALL_DIR%%" mkdir "%%INSTALL_DIR%%"
echo.
echo REM Copy files
echo copy "%%APP_EXE%%" "%%INSTALL_DIR%%\" /Y ^>nul
echo.
echo REM Create Start Menu shortcut
echo set SCRIPT=%%TEMP%%\create_shortcut.vbs
echo ^(echo Set oWS = WScript.CreateObject^("WScript.Shell"^)^) ^> "%%SCRIPT%%"
echo ^(echo sLinkFile = "%%START_MENU%%\%%APP_NAME%%.lnk"^) ^>> "%%SCRIPT%%"
echo ^(echo Set oLink = oWS.CreateShortcut^(sLinkFile^)^) ^>> "%%SCRIPT%%"
echo ^(echo oLink.TargetPath = "%%INSTALL_DIR%%\%%APP_EXE%%"^) ^>> "%%SCRIPT%%"
echo ^(echo oLink.WorkingDirectory = "%%INSTALL_DIR%%"^) ^>> "%%SCRIPT%%"
echo ^(echo oLink.Description = "CT Food Outlook - Auto attach Orderbekräftelse files to Outlook"^) ^>> "%%SCRIPT%%"
echo ^(echo oLink.Save^) ^>> "%%SCRIPT%%"
echo cscript /nologo "%%SCRIPT%%" ^>nul
echo del "%%SCRIPT%%" ^>nul
echo.
echo REM Create Startup shortcut (optional - uncomment to enable)
echo REM set STARTUP_DIR=%%APPDATA%%\Microsoft\Windows\Start Menu\Programs\Startup
echo REM ^(echo Set oWS = WScript.CreateObject^("WScript.Shell"^)^) ^> "%%SCRIPT%%"
echo REM ^(echo sLinkFile = "%%STARTUP_DIR%%\%%APP_NAME%%.lnk"^) ^>> "%%SCRIPT%%"
echo REM ^(echo Set oLink = oWS.CreateShortcut^(sLinkFile^)^) ^>> "%%SCRIPT%%"
echo REM ^(echo oLink.TargetPath = "%%INSTALL_DIR%%\%%APP_EXE%%"^) ^>> "%%SCRIPT%%"
echo REM ^(echo oLink.WorkingDirectory = "%%INSTALL_DIR%%"^) ^>> "%%SCRIPT%%"
echo REM ^(echo oLink.Save^) ^>> "%%SCRIPT%%"
echo REM cscript /nologo "%%SCRIPT%%" ^>nul
echo REM del "%%SCRIPT%%" ^>nul
echo.
echo echo.
echo echo ✅ %%APP_NAME%% installed successfully!
echo echo.
echo echo The application is now available in Start Menu as "%%APP_NAME%%"
echo echo You can also find it in: %%INSTALL_DIR%%
echo echo.
echo REM Launch the application
echo start "" "%%INSTALL_DIR%%\%%APP_EXE%%"
echo.
echo endlocal
) > "%INSTALLER_DIR%\install.bat"

REM Create uninstall script
(
echo @echo off
echo REM CT Food Outlook - Uninstallation Script
echo setlocal
echo.
echo set APP_NAME=%APP_NAME%
echo set INSTALL_DIR=%%ProgramFiles%%\%%APP_NAME%%
echo set START_MENU=%%APPDATA%%\Microsoft\Windows\Start Menu\Programs
echo set STARTUP_DIR=%%APPDATA%%\Microsoft\Windows\Start Menu\Programs\Startup
echo.
echo echo Uninstalling %%APP_NAME%%...
echo echo.
echo.
echo REM Stop running instance
echo taskkill /F /IM outlook-auto-attach.exe /T 2^>nul
echo.
echo REM Remove Start Menu shortcut
echo del "%%START_MENU%%\%%APP_NAME%%.lnk" 2^>nul
echo.
echo REM Remove Startup shortcut
echo del "%%STARTUP_DIR%%\%%APP_NAME%%.lnk" 2^>nul
echo.
echo REM Remove installation directory
echo if exist "%%INSTALL_DIR%%" rmdir /s /q "%%INSTALL_DIR%%"
echo.
echo echo.
echo echo ✅ %%APP_NAME%% uninstalled successfully!
echo echo.
echo endlocal
) > "%INSTALLER_DIR%\uninstall.bat"

REM Create README
(
echo CT Food Outlook - Installation Package
echo ======================================
echo.
echo INSTALLATION:
echo ------------
echo 1. Right-click "install.bat" and select "Run as Administrator"
echo 2. The application will be installed to Program Files
echo 3. A shortcut will be created in Start Menu
echo 4. The application will launch automatically after installation
echo.
echo UNINSTALLATION:
echo --------------
echo 1. Right-click "uninstall.bat" and select "Run as Administrator"
echo 2. Or manually delete from: %ProgramFiles%\CT Food Outlook
echo.
echo STARTUP:
echo --------
echo To make the app start automatically with Windows:
echo 1. Open Start Menu
echo 2. Find "CT Food Outlook"
echo 3. Right-click → More → Open file location
echo 4. Copy the shortcut
echo 5. Press Win+R, type: shell:startup
echo 6. Paste the shortcut there
echo.
echo The application runs in the system tray (bottom-right corner).
echo Right-click the icon to view log or quit.
) > "%INSTALLER_DIR%\README.txt"

echo ✅ Installer package created in: %INSTALLER_DIR%
echo.
echo Package contents:
echo   - %APP_EXE% (the application)
echo   - install.bat (installation script)
echo   - uninstall.bat (uninstallation script)
echo   - README.txt (instructions)
echo.
echo To create a distributable package:
echo   1. Zip the entire "%INSTALLER_DIR%" folder
echo   2. Send the zip file to users
echo   3. Users extract and run install.bat as Administrator
echo.

endlocal
