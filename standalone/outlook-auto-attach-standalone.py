#!/usr/bin/env python3
"""
Outlook Auto Attach - Standalone Application
Monitors Downloads folder and automatically opens Outlook with matching files attached.
Runs in system tray/menu bar as a background service.
Supports both Windows and macOS.
"""

import os
import sys
import shutil
import time
import logging
import platform
import subprocess
from datetime import datetime
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import pystray
from PIL import Image, ImageDraw
import threading

# Platform-specific imports
SYSTEM = platform.system()
if SYSTEM == 'Windows':
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        print("Error: pywin32 not installed. Please install it: pip install pywin32")
        sys.exit(1)
elif SYSTEM == 'Darwin':  # macOS
    pass  # Use subprocess for AppleScript

# Configure logging - platform-specific paths
if SYSTEM == 'Windows':
    LOG_DIR = os.path.join(os.path.expanduser("~"), "AppData", "Local", "OutlookAutoAttach")
else:  # macOS
    LOG_DIR = os.path.join(os.path.expanduser("~"), "Library", "Logs", "OutlookAutoAttach")

os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, "outlook-auto-attach.log")

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)


def should_process_file(filename):
    """Check if filename matches criteria for processing."""
    if not filename:
        return False
    
    filename_lower = filename.lower()
    
    # Only check for Orderbekräftelse
    has_orderbekraeftelse = 'orderbekräftelse' in filename_lower or 'orderbekr' in filename_lower
    
    return has_orderbekraeftelse


def create_unique_file_copy(original_path):
    """
    Create a unique copy of the file with a clean name format.
    Returns the path to the unique copy or original path if copy fails.
    """
    if not os.path.exists(original_path):
        logger.error(f"File not found: {original_path}")
        return None
    
    try:
        original_name = os.path.basename(original_path)
        name_parts = os.path.splitext(original_name)
        file_extension = name_parts[1]
        
        home_dir = os.path.expanduser("~")
        desktop_dir = os.path.join(home_dir, "Desktop")
        businessnxtdocs_dir = os.path.join(desktop_dir, "businessnxtdocs")
        
        os.makedirs(businessnxtdocs_dir, exist_ok=True)
        
        now = datetime.now()
        timestamp = now.strftime("%Y%m%d-%H%M%S")
        microseconds = now.strftime("%f")
        
        # All files are Orderbekräftelse (since we only process those)
        unique_name = f"Orderbekräftelse-{timestamp}-{microseconds}{file_extension}"
        unique_path = os.path.join(businessnxtdocs_dir, unique_name)
        
        try:
            shutil.copy2(original_path, unique_path)
            logger.info(f"Created copy: {unique_path}")
            return unique_path
        except (PermissionError, OSError) as e:
            logger.warning(f"Could not copy file, using original: {e}")
            return original_path
        
    except Exception as e:
        logger.error(f"Error creating file copy: {e}")
        return original_path


def open_outlook_windows(file_path):
    """Open Outlook on Windows using COM automation and attach the file."""
    if not os.path.exists(file_path):
        logger.error(f"File not found: {file_path}")
        return False, f"File not found: {file_path}"
    
    # Initialize COM for this thread (required when called from background threads like watchdog)
    # CoInitialize() uses apartment-threaded model (required for Outlook automation)
    try:
        pythoncom.CoInitialize()
    except Exception:
        # COM already initialized in this thread, that's fine - continue
        pass
    
    try:
        file_path = os.path.abspath(file_path)
        file_path = file_path.replace('/', '\\')
        
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail_item = outlook.CreateItem(0)  # 0 = olMailItem
        mail_item.Attachments.Add(file_path)
        mail_item.Display()
        
        logger.info(f"Successfully opened Outlook with file: {file_path}")
        return True, "Outlook opened successfully"
        
    except Exception as e:
        logger.error(f"Error opening Outlook: {e}")
        return False, f"Error: {str(e)}"


def open_outlook_mac(file_path):
    """Open Outlook on macOS using AppleScript and attach the file."""
    if not os.path.exists(file_path):
        logger.error(f"File not found: {file_path}")
        return False, f"File not found: {file_path}"
    
    try:
        file_path = os.path.abspath(file_path)
        
        script = f'''
        tell application "Microsoft Outlook"
            activate
            set newMessage to make new outgoing message
            tell newMessage
                make new attachment with properties {{file:POSIX file "{file_path}"}}
            end tell
            open newMessage
        end tell
        '''
        
        result = subprocess.run(
            ['osascript', '-e', script],
            capture_output=True,
            text=True,
            timeout=10
        )
        
        if result.returncode == 0:
            logger.info(f"Successfully opened Outlook with file: {file_path}")
            return True, "Outlook opened successfully"
        else:
            error_msg = result.stderr.strip() if result.stderr else "Unknown error"
            logger.error(f"AppleScript error: {error_msg}")
            return False, f"AppleScript error: {error_msg}"
            
    except subprocess.TimeoutExpired:
        logger.error("Timeout opening Outlook")
        return False, "Timeout opening Outlook"
    except Exception as e:
        logger.error(f"Error opening Outlook: {e}")
        return False, f"Error: {str(e)}"


def open_outlook(file_path):
    """Open Outlook with file attached - platform-specific."""
    if SYSTEM == 'Windows':
        return open_outlook_windows(file_path)
    elif SYSTEM == 'Darwin':
        return open_outlook_mac(file_path)
    else:
        return False, f"Unsupported platform: {SYSTEM}"


class DownloadsHandler(FileSystemEventHandler):
    """Handle file system events in Downloads folder."""
    
    def __init__(self):
        super().__init__()
        self.processed_files = set()
        self.processing_lock = threading.Lock()
    
    def on_created(self, event):
        """Called when a file is created."""
        if event.is_directory:
            return
        
        # Wait a moment for file to be fully written
        time.sleep(0.5)
        
        file_path = event.src_path
        
        # Skip if already processed
        with self.processing_lock:
            if file_path in self.processed_files:
                return
            self.processed_files.add(file_path)
        
        self.process_file(file_path)
    
    def on_modified(self, event):
        """Called when a file is modified (sometimes triggered on download completion)."""
        if event.is_directory:
            return
        
        file_path = event.src_path
        
        # Only process if file is complete (not still being written)
        try:
            # Check if file is accessible (not locked by another process)
            with open(file_path, 'rb'):
                pass
        except (PermissionError, IOError):
            # File is still being written, skip
            return
        
        # Skip if already processed
        with self.processing_lock:
            if file_path in self.processed_files:
                return
            self.processed_files.add(file_path)
        
        self.process_file(file_path)
    
    def process_file(self, file_path):
        """Process a downloaded file."""
        try:
            filename = os.path.basename(file_path)
            logger.info(f"Checking file: {filename}")
            
            if not should_process_file(filename):
                logger.debug(f"File does not match criteria: {filename}")
                return
            
            logger.info(f"File matches criteria, processing: {filename}")
            
            # Create unique copy
            unique_file_path = create_unique_file_copy(file_path)
            if not unique_file_path:
                logger.error(f"Failed to create copy of: {file_path}")
                return
            
            # Delete the original file from Downloads folder (keep Downloads clean)
            # Only delete if the copy was successful and it's different from the original
            if unique_file_path != file_path and os.path.exists(unique_file_path):
                try:
                    # Wait a moment to ensure file is not locked
                    time.sleep(0.5)
                    if os.path.exists(file_path):
                        os.remove(file_path)
                        logger.info(f"Deleted original file from Downloads: {filename}")
                except (PermissionError, OSError) as e:
                    logger.warning(f"Could not delete original file {filename}: {e}")
            
            # Open Outlook with file attached
            success, message = open_outlook(unique_file_path)
            if success:
                logger.info(f"Successfully processed: {filename}")
            else:
                logger.error(f"Failed to open Outlook: {message}")
                
        except Exception as e:
            logger.error(f"Error processing file {file_path}: {e}")


def create_tray_icon():
    """Create system tray icon."""
    # Try to load the logo file
    script_dir = os.path.dirname(os.path.abspath(__file__))
    logo_path = os.path.join(script_dir, "ct_food_app_logo.png")
    
    # If running from PyInstaller bundle, check in the bundle
    if hasattr(sys, '_MEIPASS'):  # Running from PyInstaller bundle
        logo_path = os.path.join(sys._MEIPASS, "ct_food_app_logo.png")
    elif SYSTEM == 'Darwin':
        # On macOS, if running from .app bundle, check Resources
        app_bundle = os.path.dirname(os.path.dirname(os.path.dirname(script_dir)))
        if app_bundle.endswith('.app'):
            resources_path = os.path.join(app_bundle, "Contents", "Resources", "ct_food_app_logo.png")
            if os.path.exists(resources_path):
                logo_path = resources_path
    
    # Try to load the logo
    if os.path.exists(logo_path):
        try:
            img = Image.open(logo_path)
            # Resize to 64x64 for system tray
            img = img.resize((64, 64), Image.Resampling.LANCZOS)
            return img
        except Exception as e:
            logger.warning(f"Could not load logo, using default: {e}")
    
    # Fallback: Create a simple icon
    image = Image.new('RGBA', (64, 64), color=(255, 255, 255, 0))
    draw = ImageDraw.Draw(image)
    # Draw a blue circle (Outlook color)
    draw.ellipse([8, 8, 56, 56], fill=(0, 120, 212), outline=(0, 90, 180), width=2)
    # Draw 'O' for Outlook (simple text without font dependency)
    draw.text((22, 18), 'O', fill='white')
    return image


def setup_tray_icon(observer):
    """Setup system tray/menu bar icon with menu."""
    def on_quit(icon, item):
        logger.info("Shutting down...")
        observer.stop()
        icon.stop()
    
    def show_log(icon, item):
        if SYSTEM == 'Windows':
            os.startfile(LOG_FILE)
        else:  # macOS
            subprocess.run(['open', LOG_FILE])
    
    menu = pystray.Menu(
        pystray.MenuItem("View Log", show_log),
        pystray.MenuItem("Quit", on_quit)
    )
    
    icon = pystray.Icon("CT Food Outlook", create_tray_icon(), "CT Food Outlook - Auto attach Orderbekräftelse files to Outlook", menu)
    return icon


def get_downloads_folder():
    """Get the user's Downloads folder path - platform-specific. Gets actual logged-in user's folder, not Administrator."""
    if SYSTEM == 'Windows':
        # First, try to get Downloads folder using standard method
        downloads = None
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            downloads = shell.SpecialFolders("Downloads")
            if downloads and os.path.exists(downloads):
                # Check if we're running as Administrator and found Administrator's folder
                # If so, try to find the actual logged-in user's folder
                if "Administrator" in downloads:
                    logger.debug("Running as Administrator, looking for actual user's Downloads folder...")
                    downloads = None  # Continue to find actual user's folder
                else:
                    logger.info(f"Found Downloads folder: {downloads}")
                    return downloads
        except:
            pass
        
        # If we didn't find a good folder (or we're Administrator), scan C:\Users for actual user folders
        if downloads is None or "Administrator" in str(downloads):
            try:
                system_drive = os.environ.get('SystemDrive', 'C:')
                users_dir = os.path.join(system_drive, 'Users')
                if os.path.exists(users_dir):
                    # System folders to exclude
                    exclude = {'Administrator', 'Default', 'Default User', 'Public', 'All Users', '.NET v4.5', '.NET v4.5 Classic'}
                    
                    # Try to find user folders with Downloads
                    user_folders = []
                    for folder in os.listdir(users_dir):
                        if folder in exclude:
                            continue
                        user_path = os.path.join(users_dir, folder)
                        if not os.path.isdir(user_path):
                            continue
                        
                        # Try English "Downloads"
                        user_downloads = os.path.join(user_path, "Downloads")
                        if os.path.exists(user_downloads):
                            user_folders.append(user_downloads)
                            continue
                        
                        # Try Swedish "Hämtade filer"
                        user_downloads_sv = os.path.join(user_path, "Hämtade filer")
                        if os.path.exists(user_downloads_sv):
                            user_folders.append(user_downloads_sv)
                    
                    # If we found user folders, prefer the one that's not Administrator
                    for user_dl in user_folders:
                        if "Administrator" not in user_dl:
                            logger.info(f"Found Downloads folder (actual user): {user_dl}")
                            return user_dl
                    
                    # If only Administrator found, use it
                    if user_folders:
                        logger.info(f"Found Downloads folder: {user_folders[0]}")
                        return user_folders[0]
            except Exception as e:
                logger.debug(f"Could not scan user profiles: {e}")
        
        # Fallback: Use current process user's Downloads
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        if not os.path.exists(downloads):
            downloads = os.path.join(os.path.expanduser("~"), "Hämtade filer")  # Swedish Windows
    else:  # macOS
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    
    # Ensure it exists
    if not os.path.exists(downloads):
        logger.warning(f"Downloads folder not found at {downloads}, creating it")
        os.makedirs(downloads, exist_ok=True)
    
    return downloads


def main():
    """Main application entry point."""
    logger.info("Starting CT Food Outlook Application")
    
    # Get Downloads folder
    downloads_folder = get_downloads_folder()
    logger.info(f"Monitoring folder: {downloads_folder}")
    
    # Setup file system watcher
    event_handler = DownloadsHandler()
    observer = Observer()
    observer.schedule(event_handler, downloads_folder, recursive=False)
    observer.start()
    
    logger.info("File system watcher started")
    
    # Setup system tray icon
    icon = setup_tray_icon(observer)
    
    # Run icon in a separate thread
    icon_thread = threading.Thread(target=icon.run, daemon=True)
    icon_thread.start()
    
    try:
        # Keep the main thread alive
        while observer.is_alive():
            time.sleep(1)
    except KeyboardInterrupt:
        logger.info("Interrupted by user")
    finally:
        observer.stop()
        observer.join()
        icon.stop()
        logger.info("Application stopped")


if __name__ == '__main__':
    main()
