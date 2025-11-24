# Bygga Windows Executable på Mac

## Alternativ 1: Virtualisering (Rekommenderat)

**Steg:**

1. **Installera VirtualBox** (gratis):
   - Ladda ner från: https://www.virtualbox.org/
   - Eller använd Parallels/VMware Fusion (betalning men snabbare)

2. **Skapa Windows VM:**
   - Installera Windows 10/11 i VM
   - Dela `server`-mappen från Mac till VM

3. **I Windows VM:**
   ```cmd
   cd server
   pip install pyinstaller
   build-standalone.bat
   ```

4. **Kopiera `.exe` tillbaka till Mac**

## Alternativ 2: GitHub Actions (Gratis, molnbaserat)

Skapa en GitHub Action som bygger på Windows automatiskt.

## Alternativ 3: Wine (Svårare)

Kör Windows-program på Mac, men kan vara problematiskt.

## Alternativ 4: Be IT/kollega

Enklast - be någon med Windows att bygga den.

