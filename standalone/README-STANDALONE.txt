Outlook Auto Attach - Standalone Windows Application
====================================================

INSTALLATION:
-------------

1. Download "outlook-auto-attach.exe"
2. Double-click to run (or create a shortcut)
3. The application will run in the system tray (look for the icon in the bottom-right)
4. It will automatically start monitoring your Downloads folder

AUTOMATIC STARTUP (Optional):
-----------------------------

To make the app start automatically when Windows boots:

Method 1: Startup Folder
1. Press Win + R
2. Type: shell:startup
3. Press Enter
4. Copy or create a shortcut to "outlook-auto-attach.exe" in this folder

Method 2: Task Scheduler (More reliable)
1. Open Task Scheduler (search in Start menu)
2. Click "Create Basic Task"
3. Name: "Outlook Auto Attach"
4. Trigger: "When I log on"
5. Action: "Start a program"
6. Browse to "outlook-auto-attach.exe"
7. Click Finish

HOW IT WORKS:
-------------

- The app runs in the background (system tray icon)
- It monitors your Downloads folder
- When a file containing "Orderbekräftelse" or "Orderbekr" is downloaded
- It automatically:
  1. Creates a copy with a clean name in Desktop\businessnxtdocs\
  2. Opens Microsoft Outlook with the file attached

SYSTEM TRAY ICON:
-----------------

Right-click the system tray icon to:
- View Log: Opens the log file
- Quit: Stop the application

LOG FILE:
---------

Logs are saved to:
%USERPROFILE%\AppData\Local\OutlookAutoAttach\outlook-auto-attach.log

REQUIREMENTS:
-------------

- Windows 10 or later
- Microsoft Outlook installed
- No Python or other dependencies needed (all included in .exe)

TROUBLESHOOTING:
---------------

- If Outlook doesn't open: Check the log file for errors
- If files aren't being detected: Check that files match the criteria
- To stop the app: Right-click system tray icon → Quit
- To restart: Just run the .exe again
