Outlook Auto Attach - Installation Guide
=========================================

REQUIREMENTS:
- Chrome browser
- Microsoft Outlook installed
- NO Python required (standalone executables included)

INSTALLATION:
-------------

Step 1: Install Chrome Extension
1. Open Chrome → chrome://extensions/
2. Enable "Developer mode" (toggle top right)
3. Click "Load unpacked"
4. Select the folder containing manifest.json, background.js, popup.html, etc.
5. Note your Extension ID (shown under extension name)

Step 2: Install Native Host

Mac:
- Open Terminal
- Navigate to the "server" folder
- Run: ./install-native-host.sh YOUR_EXTENSION_ID
- Replace YOUR_EXTENSION_ID with the ID from Step 1

Windows:
- Open Command Prompt (as Administrator)
- Navigate to the "server" folder  
- Run: install-native-host.bat YOUR_EXTENSION_ID
- Replace YOUR_EXTENSION_ID with the ID from Step 1

Step 3: Test
- Download a file with "Orderbekräftelse", "Inköp", or "1000322" in filename
- Click extension icon when it appears
- Click "Open Outlook"
- Outlook opens with file attached!

TROUBLESHOOTING:
---------------
- If Outlook doesn't open: Check Chrome console (chrome://extensions/ → service worker)
- On Mac: May need to grant Chrome "Full Disk Access" in System Preferences
- Make sure native host was installed correctly (check manifest file exists)

