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
        
        # Try to get existing Outlook instance first, or create new one
        outlook = None
        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
            logger.info("Using existing Outlook instance")
        except:
            pass
        
        if outlook is None:
            outlook = win32com.client.Dispatch("Outlook.Application")
            logger.info("Created new Outlook instance")
        
        # Create email with attachment (retry up to 3 times)
        max_retries = 3
        for attempt in range(max_retries):
            try:
                mail_item = outlook.CreateItem(0)  # 0 = olMailItem
                mail_item.Attachments.Add(file_path)
                mail_item.Display()
                logger.info(f"Successfully opened Outlook with file: {file_path}")
                return True, "Outlook opened successfully"
            except Exception as e:
                if attempt < max_retries - 1:
                    logger.warning(f"Outlook opening attempt {attempt + 1} failed, retrying: {e}")
                    time.sleep(0.5)
                    continue
                else:
                    raise
        
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
    
    def __init__(self, downloads_folder):
        super().__init__()
        self.processed_files = {}  # file_path -> (mtime, processed_time)
        self.pending_files = {}  # file_path -> last_seen_time
        self.processing_lock = threading.Lock()
        self.pending_lock = threading.Lock()
        self.downloads_folder = downloads_folder
        self.last_scan_time = time.time()
    
    def _get_file_signature(self, file_path):
        """Get a unique signature for a file (path + modification time)."""
        try:
            if os.path.exists(file_path):
                mtime = os.path.getmtime(file_path)
                return (file_path, mtime)
        except:
            pass
        return (file_path, None)
    
    def _should_process(self, file_path):
        """Check if file should be processed (not already processed or pending). Uses file signature (path + mtime)."""
        signature = self._get_file_signature(file_path)
        with self.processing_lock:
            # Check if this exact file (same path + same modification time) was already processed or is being processed
            if signature in self.processed_files:
                processed_time, _ = self.processed_files[signature]
                # If processed very recently (within last 10 seconds), it's likely being processed by another thread
                if time.time() - processed_time < 10.0:
                    return False
            
            # Clean up old entries for files that no longer exist or have been modified
            keys_to_remove = []
            for key in list(self.processed_files.keys()):
                stored_path, stored_mtime = key
                if not os.path.exists(stored_path):
                    # File was deleted, remove from processed set
                    keys_to_remove.append(key)
                elif stored_path == file_path and stored_mtime != signature[1]:
                    # Same path but different mtime means it's a new file with same name
                    keys_to_remove.append(key)
            
            for key in keys_to_remove:
                del self.processed_files[key]
            
            # Mark this file as being processed (with current timestamp)
            self.processed_files[signature] = (time.time(), signature[1])
        return True
    
    def _mark_as_not_processed(self, file_path):
        """Remove file from processed set (in case processing failed)."""
        signature = self._get_file_signature(file_path)
        with self.processing_lock:
            # Remove this signature and any old signatures for this path
            keys_to_remove = [key for key in self.processed_files.keys() if key[0] == file_path]
            for key in keys_to_remove:
                del self.processed_files[key]
    
    def on_created(self, event):
        """Called when a file is created."""
        try:
            if event.is_directory:
                return
            
            file_path = event.src_path
            filename = os.path.basename(file_path)
            
            # Skip temporary files (browsers often create .tmp files first)
            if filename.startswith('.') or filename.endswith('.tmp') or filename.endswith('.crdownload'):
                logger.debug(f"Skipping temporary file: {filename}")
                return
            
            # Skip if already processed
            if not self._should_process(file_path):
                return
            
            # Process in a separate thread to avoid blocking
            logger.info(f"File created event: {filename}")
            threading.Thread(target=self._process_file_delayed, args=(file_path, 2.0), daemon=True).start()
        except Exception as e:
            logger.error(f"Error in on_created: {e}", exc_info=True)
    
    def on_moved(self, event):
        """Called when a file is moved/renamed (browsers often rename temp files)."""
        try:
            if event.is_directory:
                return
            
            dest_path = event.dest_path
            filename = os.path.basename(dest_path)
            
            # Skip temporary files
            if filename.startswith('.') or filename.endswith('.tmp') or filename.endswith('.crdownload'):
                return
            
            # Skip if already processed
            if not self._should_process(dest_path):
                return
            
            # Process the renamed file
            logger.info(f"File moved/renamed event: {filename}")
            threading.Thread(target=self._process_file_delayed, args=(dest_path, 2.0), daemon=True).start()
        except Exception as e:
            logger.error(f"Error in on_moved: {e}", exc_info=True)
    
    def on_modified(self, event):
        """Called when a file is modified (sometimes triggered on download completion)."""
        try:
            if event.is_directory:
                return
            
            file_path = event.src_path
            filename = os.path.basename(file_path)
            
            # Skip temporary files
            if filename.startswith('.') or filename.endswith('.tmp') or filename.endswith('.crdownload'):
                return
            
            # Skip if already processed (check BEFORE tracking as pending)
            if not self._should_process(file_path):
                return
            
            # Track pending files (might be still downloading)
            current_time = time.time()
            with self.pending_lock:
                if file_path in self.pending_files:
                    last_seen = self.pending_files[file_path]
                    # Only process if file hasn't been modified recently (likely complete)
                    if current_time - last_seen < 2.0:  # File modified less than 2 seconds ago
                        self.pending_files[file_path] = current_time  # Update timestamp
                        # Remove from processed set since we're not processing it yet
                        self._mark_as_not_processed(file_path)
                        return  # Still downloading, wait
                self.pending_files[file_path] = current_time
            
            # Process with delay to ensure file is complete
            logger.debug(f"File modified event: {filename}")
            threading.Thread(target=self._process_file_delayed, args=(file_path, 2.0), daemon=True).start()
        except Exception as e:
            logger.error(f"Error in on_modified: {e}", exc_info=True)
    
    def _process_file_delayed(self, file_path, initial_delay=2.0):
        """Process file after a delay to ensure it's complete."""
        # Wait for file to be complete (browsers need time to finish writing)
        time.sleep(initial_delay)
        
        # Check if file still exists and can be opened
        if not os.path.exists(file_path):
            logger.debug(f"File no longer exists: {os.path.basename(file_path)}")
            self._mark_as_not_processed(file_path)
            return
        
        # Double-check this file hasn't been processed by another thread while we were waiting
        signature = self._get_file_signature(file_path)
        with self.processing_lock:
            # If file signature is already in processed_files, it's being processed or was already processed
            # Check if it was recently marked (within last 5 seconds) - likely being processed by another thread
            if signature in self.processed_files:
                processed_time, _ = self.processed_files[signature]
                if time.time() - processed_time < 5.0:
                    logger.debug(f"File already being processed by another thread: {os.path.basename(file_path)}")
                    return
                # Old entry, might be a different file - remove it
                del self.processed_files[signature]
            
            # Re-mark as being processed (in case multiple threads reached here)
            self.processed_files[signature] = (time.time(), signature[1])
        
        # Try to open file to ensure it's not locked
        try:
            with open(file_path, 'rb'):
                pass
        except (PermissionError, IOError) as e:
            logger.debug(f"File still locked, skipping: {os.path.basename(file_path)}")
            self._mark_as_not_processed(file_path)
            return
        
        # Process the file
        try:
            self.process_file(file_path)
        except Exception as e:
            logger.error(f"Error in process_file for {os.path.basename(file_path)}: {e}", exc_info=True)
            # Remove from processed set so it can be retried
            self._mark_as_not_processed(file_path)
    
    def process_file(self, file_path):
        """Process a downloaded file."""
        filename = os.path.basename(file_path)
        logger.info(f"Processing file: {filename}")
        
        try:
            # Double-check file still exists
            if not os.path.exists(file_path):
                logger.warning(f"File no longer exists: {filename}")
                return
            
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
                    # Small delay to ensure copy is complete
                    time.sleep(0.3)
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
            logger.error(f"Error processing file {file_path}: {e}", exc_info=True)
            raise  # Re-raise so caller can handle it
    
    def scan_for_new_files(self):
        """Periodic scan for files that might have been missed by file system events."""
        try:
            if not os.path.exists(self.downloads_folder):
                logger.warning(f"Downloads folder does not exist: {self.downloads_folder}")
                return
            
            current_time = time.time()
            # Scan every 2 seconds (very aggressive to catch missed files)
            if current_time - self.last_scan_time < 2.0:
                return
            
            self.last_scan_time = current_time
            
            try:
                files = os.listdir(self.downloads_folder)
                logger.debug(f"Periodic scan checking {len(files)} files in Downloads folder")
                
                for filename in files:
                    file_path = os.path.join(self.downloads_folder, filename)
                    
                    # Skip directories and temporary files
                    if os.path.isdir(file_path):
                        continue
                    if filename.startswith('.') or filename.endswith('.tmp') or filename.endswith('.crdownload'):
                        continue
                    
                    # Check if it's a file we should process
                    if not should_process_file(filename):
                        continue
                    
                    # Check if already processed (using file signature)
                    signature = self._get_file_signature(file_path)
                    with self.processing_lock:
                        if signature in self.processed_files:
                            logger.debug(f"Periodic scan: file already in processed set: {filename}")
                            continue
                        # Also check if there's an old entry for this path (different mtime = new file)
                        old_entries = [key for key in self.processed_files.keys() if key[0] == file_path]
                        for old_key in old_entries:
                            # If file exists and has different mtime, it's a new file with same name
                            if os.path.exists(file_path):
                                try:
                                    current_mtime = os.path.getmtime(file_path)
                                    if old_key[1] != current_mtime:
                                        # Different file with same name - remove old entry
                                        del self.processed_files[old_key]
                                        logger.debug(f"Periodic scan: found new file with same name, removed old entry: {filename}")
                                except:
                                    pass
                    
                    # Check file modification time (only process files older than 2 seconds - likely complete)
                    try:
                        file_mtime = os.path.getmtime(file_path)
                        if current_time - file_mtime < 2.0:
                            logger.debug(f"Periodic scan: file too recent (downloading?): {filename}")
                            continue  # File too recent, might still be downloading
                    except Exception as e:
                        logger.debug(f"Periodic scan: error getting mtime for {filename}: {e}")
                        continue
                    
                    # Check if file is stable (not locked)
                    try:
                        with open(file_path, 'rb'):
                            pass
                    except (PermissionError, IOError) as e:
                        logger.debug(f"Periodic scan: file locked: {filename}")
                        continue  # File is locked, skip
                    
                    # Process this file
                    logger.info(f"Periodic scan found unprocessed file: {filename}")
                    if self._should_process(file_path):
                        logger.info(f"Periodic scan: starting processing thread for {filename}")
                        threading.Thread(target=self._process_file_delayed, args=(file_path, 0.5), daemon=True).start()
                    else:
                        logger.debug(f"Periodic scan: file marked as processed by _should_process: {filename}")
            except Exception as e:
                logger.error(f"Error during periodic scan (listing files): {e}", exc_info=True)
        except Exception as e:
            logger.error(f"Error in scan_for_new_files: {e}", exc_info=True)


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
        """Open the log file in default text editor."""
        try:
            if SYSTEM == 'Windows':
                os.startfile(LOG_FILE)
            else:  # macOS
                subprocess.run(['open', LOG_FILE])
        except Exception as e:
            logger.error(f"Error opening log file: {e}")
            # Fallback: open the folder
            open_log_folder(icon, item)
    
    def open_log_folder(icon, item):
        """Open the log file folder in File Explorer/Finder."""
        try:
            if SYSTEM == 'Windows':
                # Open folder and select the log file
                subprocess.run(['explorer', '/select,', LOG_FILE], check=False)
            else:  # macOS
                subprocess.run(['open', '-R', LOG_FILE], check=False)
        except Exception as e:
            logger.error(f"Error opening log folder: {e}")
    
    menu = pystray.Menu(
        pystray.MenuItem("View Log File", show_log),
        pystray.MenuItem("Open Log Folder", open_log_folder),
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
    event_handler = DownloadsHandler(downloads_folder)
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
        # Keep the main thread alive and periodically scan for missed files
        last_observer_check = time.time()
        while observer.is_alive():
            time.sleep(1)
            
            # Periodic scan as fallback in case file system events are missed
            try:
                event_handler.scan_for_new_files()
            except Exception as e:
                logger.error(f"Error in periodic scan: {e}", exc_info=True)
            
            # Log observer status periodically (for debugging)
            current_time = time.time()
            if current_time - last_observer_check > 30.0:
                if observer.is_alive():
                    # Clean up old processed file entries for files that no longer exist
                    with event_handler.processing_lock:
                        keys_to_remove = [key for key in event_handler.processed_files.keys() 
                                        if not os.path.exists(key[0])]
                        for key in keys_to_remove:
                            del event_handler.processed_files[key]
                    logger.debug(f"Observer is running. Processed files count: {len(event_handler.processed_files)}")
                else:
                    logger.error("Observer stopped unexpectedly!")
                last_observer_check = current_time
    except KeyboardInterrupt:
        logger.info("Interrupted by user")
    finally:
        observer.stop()
        observer.join()
        icon.stop()
        logger.info("Application stopped")


if __name__ == '__main__':
    main()
