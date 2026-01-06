@echo off
REM Create a simple, robust Windows installer

setlocal enabledelayedexpansion

set APP_NAME=CT Food Outlook
set APP_EXE=outlook-auto-attach.exe
set INSTALLER_DIR=dist-installer

if not exist "dist\%APP_EXE%" (
    echo Error: %APP_EXE% not found in dist folder
    exit /b 1
)

echo Creating installer package...
if exist "%INSTALLER_DIR%" rmdir /s /q "%INSTALLER_DIR%"
mkdir "%INSTALLER_DIR%"

copy "dist\%APP_EXE%" "%INSTALLER_DIR%\%APP_EXE%"

REM Create a simple install script
echo @echo off > "%INSTALLER_DIR%\Install.bat"
echo setlocal enabledelayedexpansion >> "%INSTALLER_DIR%\Install.bat"
echo. >> "%INSTALLER_DIR%\Install.bat"
echo echo ======================================== >> "%INSTALLER_DIR%\Install.bat"
echo echo CT Food Outlook - Installation >> "%INSTALLER_DIR%\Install.bat"
echo echo ======================================== >> "%INSTALLER_DIR%\Install.bat"
echo echo. >> "%INSTALLER_DIR%\Install.bat"
echo. >> "%INSTALLER_DIR%\Install.bat"
echo REM Check admin >> "%INSTALLER_DIR%\Install.bat"
echo net session ^>nul 2^>^&1 >> "%INSTALLER_DIR%\Install.bat"
echo if errorlevel 1 ^( >> "%INSTALLER_DIR%\Install.bat"
echo     echo ERROR: Must run as Administrator! >> "%INSTALLER_DIR%\Install.bat"
echo     echo Right-click this file and select "Run as Administrator" >> "%INSTALLER_DIR%\Install.bat"
echo     pause >> "%INSTALLER_DIR%\Install.bat"
echo     exit /b 1 >> "%INSTALLER_DIR%\Install.bat"
echo ^) >> "%INSTALLER_DIR%\Install.bat"
echo. >> "%INSTALLER_DIR%\Install.bat"
echo set APP_NAME=%APP_NAME% >> "%INSTALLER_DIR%\Install.bat"
echo set APP_EXE=%APP_EXE% >> "%INSTALLER_DIR%\Install.bat"
echo set INSTALL_DIR=%%ProgramFiles%%\%APP_NAME% >> "%INSTALLER_DIR%\Install.bat"
echo. >> "%INSTALLER_DIR%\Install.bat"
echo echo Installing to: %%INSTALL_DIR%% >> "%INSTALLER_DIR%\Install.bat"
echo. >> "%INSTALLER_DIR%\Install.bat"
echo if not exist "%%INSTALL_DIR%%" mkdir "%%INSTALL_DIR%%" >> "%INSTALLER_DIR%\Install.bat"
echo if errorlevel 1 ^( >> "%INSTALLER_DIR%\Install.bat"
echo     echo ERROR: Could not create directory >> "%INSTALLER_DIR%\Install.bat"
echo     pause >> "%INSTALLER_DIR%\Install.bat"
echo     exit /b 1 >> "%INSTALLER_DIR%\Install.bat"
echo ^) >> "%INSTALLER_DIR%\Install.bat"
echo. >> "%INSTALLER_DIR%\Install.bat"
echo echo Copying files... >> "%INSTALLER_DIR%\Install.bat"
echo copy "%APP_EXE%" "%%INSTALL_DIR%%\" /Y >> "%INSTALLER_DIR%\Install.bat"
echo if errorlevel 1 ^( >> "%INSTALLER_DIR%\Install.bat"
echo     echo ERROR: Could not copy file >> "%INSTALLER_DIR%\Install.bat"
echo     pause >> "%INSTALLER_DIR%\Install.bat"
echo     exit /b 1 >> "%INSTALLER_DIR%\Install.bat"
echo ^) >> "%INSTALLER_DIR%\Install.bat"
echo. >> "%INSTALLER_DIR%\Install.bat"
echo echo Creating shortcuts... >> "%INSTALLER_DIR%\Install.bat"
echo set SCRIPT=%%TEMP%%\shortcut.vbs >> "%INSTALLER_DIR%\Install.bat"
echo Set oWS = WScript.CreateObject^("WScript.Shell"^) > %%SCRIPT%% >> "%INSTALLER_DIR%\Install.bat"
echo sLinkFile = "%%APPDATA%%\Microsoft\Windows\Start Menu\Programs\%APP_NAME%.lnk" >> %%SCRIPT%% >> "%INSTALLER_DIR%\Install.bat"
echo Set oLink = oWS.CreateShortcut^(sLinkFile^) >> %%SCRIPT%% >> "%INSTALLER_DIR%\Install.bat"
echo oLink.TargetPath = "%%INSTALL_DIR%%\%APP_EXE%" >> %%SCRIPT%% >> "%INSTALLER_DIR%\Install.bat"
echo oLink.WorkingDirectory = "%%INSTALL_DIR%%" >> %%SCRIPT%% >> "%INSTALLER_DIR%\Install.bat"
echo oLink.Description = "%APP_NAME%" >> %%SCRIPT%% >> "%INSTALLER_DIR%\Install.bat"
echo oLink.Save >> %%SCRIPT%% >> "%INSTALLER_DIR%\Install.bat"
echo cscript /nologo %%SCRIPT%% >> "%INSTALLER_DIR%\Install.bat"
echo del %%SCRIPT%% >> "%INSTALLER_DIR%\Install.bat"
echo. >> "%INSTALLER_DIR%\Install.bat"
echo echo. >> "%INSTALLER_DIR%\Install.bat"
echo echo ✅ Installation successful! >> "%INSTALLER_DIR%\Install.bat"
echo echo. >> "%INSTALLER_DIR%\Install.bat"
echo echo Launching application... >> "%INSTALLER_DIR%\Install.bat"
echo start "" "%%INSTALL_DIR%%\%APP_EXE%" >> "%INSTALLER_DIR%\Install.bat"
echo. >> "%INSTALLER_DIR%\Install.bat"
echo echo Press any key to close... >> "%INSTALLER_DIR%\Install.bat"
echo pause ^>nul >> "%INSTALLER_DIR%\Install.bat"
echo endlocal >> "%INSTALLER_DIR%\Install.bat"

REM Create README
(
echo CT Food Outlook - Installation
echo ===============================
echo.
echo 1. Right-click "Install.bat"
echo 2. Select "Run as Administrator"
echo 3. Follow the prompts
echo.
echo The app will be installed to Program Files and available in Start Menu.
) > "%INSTALLER_DIR%\README.txt"

echo ✅ Installer created in: %INSTALLER_DIR%
echo.
echo Creating ZIP file...
powershell -Command "Compress-Archive -Path '%INSTALLER_DIR%\*' -DestinationPath 'CT-Food-Outlook-Windows.zip' -Force"

if exist "CT-Food-Outlook-Windows.zip" (
    echo ✅ ZIP created: CT-Food-Outlook-Windows.zip
) else (
    echo ⚠️  ZIP creation failed, but installer folder is ready
)

endlocal
