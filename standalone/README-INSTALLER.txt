CT Food Outlook - Installation Guide
=====================================

BUILDING THE INSTALLER:
-----------------------

1. Build the executable:
   - Run: build-standalone.bat
   - This creates: dist\outlook-auto-attach.exe

2. Create the installer package:
   - Run: create-installer.bat
   - This creates: installer\ folder with all files

3. Distribute:
   - Zip the entire "installer" folder
   - Send to users

FOR USERS:
----------

1. Extract the zip file
2. Right-click "install.bat" → "Run as Administrator"
3. The app will be installed and appear in Start Menu as "CT Food Outlook"
4. Launch it from Start Menu like any other program

The app will:
- Appear in Start Menu as "CT Food Outlook"
- Run in system tray (background)
- Automatically monitor Downloads folder
- Open Outlook when Orderbekräftelse files are downloaded

OPTIONAL - Auto-start with Windows:
------------------------------------

1. Open Start Menu
2. Find "CT Food Outlook"
3. Right-click → More → Open file location
4. Copy the shortcut
5. Press Win+R, type: shell:startup
6. Paste the shortcut there

Now it will start automatically when Windows boots.
