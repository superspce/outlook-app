#!/usr/bin/env python3
"""
Outlook Auto Attach Native Messaging Host
Communicates with Chrome extension via Native Messaging API (stdin/stdout).
Opens Outlook with file attached.
Supports both macOS (AppleScript) and Windows (COM automation).
"""

import json
import sys
import os
import subprocess
import platform
import shutil
import re
from datetime import datetime


def create_unique_file_copy(original_path):
    """
    Create a unique copy of the file with a clean name format based on file type:
    - Files with 7-digit numbers → "Faktura-datum-tid.pdf"
    - Files with "Inköp" → "Order-datum-tid.pdf"
    - Files with "Orderbekräftelse" → "Orderbekräftelse-datum-tid.pdf"
    Includes microseconds to ensure uniqueness and avoid system-appended numbers.
    Returns the path to the unique copy.
    If copying fails due to permissions, returns the original path.
    """
    if not os.path.exists(original_path):
        return None, f"File not found: {original_path}"
    
    # Check if we can read the original file
    if not os.access(original_path, os.R_OK):
        # If we can't read it, try to use the original path directly
        import sys
        sys.stderr.write(f"DEBUG: Cannot read original file, will try to use it directly: {original_path}\n")
        sys.stderr.flush()
        return original_path, None
    
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
        
        original_lower = original_name.lower()
        
        has_7digits = bool(re.search(r'\d{7}', original_name))
        has_inkop = 'inköp' in original_lower or 'inkop' in original_lower
        has_orderbekraeftelse = 'orderbekräftelse' in original_lower or 'orderbekr' in original_lower
                
        if has_7digits:
            base_name = "Faktura"
        elif has_inkop:
            base_name = "Order"
        else:
            base_name = "Orderbekräftelse"
        
        unique_name = f"{base_name}-{timestamp}-{microseconds}{file_extension}"
        unique_path = os.path.join(businessnxtdocs_dir, unique_name)
        
        # Try to copy, but if it fails due to permissions, use original
        try:
            shutil.copy2(original_path, unique_path)
            return unique_path, None
        except (PermissionError, OSError) as e:
            # If we can't copy due to permissions, use the original file
            import sys
            sys.stderr.write(f"DEBUG: Could not copy file (permission denied), using original: {original_path}\n")
            sys.stderr.flush()
            return original_path, None
        
    except Exception as e:
        # If anything else fails, try to use the original path
        import sys
        sys.stderr.write(f"DEBUG: Error in create_unique_file_copy, using original: {str(e)}\n")
        sys.stderr.flush()
        return original_path, None


def open_outlook_mac(file_path):
    """Open Outlook on macOS using AppleScript and attach the file."""
    if not os.path.exists(file_path):
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
            return True, "Outlook opened successfully"
        else:
            error_msg = result.stderr.strip() if result.stderr else "Unknown error"
            return False, f"AppleScript error: {error_msg}"
            
    except subprocess.TimeoutExpired:
        return False, "Timeout opening Outlook"
    except Exception as e:
        return False, f"Error: {str(e)}"


def open_outlook_windows(file_path):
    """Open Outlook on Windows using COM automation and attach the file."""
    if not os.path.exists(file_path):
        return False, f"File not found: {file_path}"
    
    try:
        import win32com.client
        
        file_path = os.path.abspath(file_path)
        file_path = file_path.replace('/', '\\')
        
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        mail_item = outlook.CreateItem(0)
        
        mail_item.Attachments.Add(file_path)
        
        mail_item.Display()
        
        return True, "Outlook opened successfully"
        
    except ImportError:
        return False, "pywin32 not installed"
    except Exception as e:
        return False, f"Error: {str(e)}"


def send_message(message):
    """Send a message to the extension via stdout."""
    # Native Messaging protocol: 4-byte length header (little-endian) + JSON message
    message_json = json.dumps(message)
    message_bytes = message_json.encode('utf-8')
    length = len(message_bytes)
    
    # Write length as 4-byte little-endian integer
    sys.stdout.buffer.write(length.to_bytes(4, byteorder='little'))
    # Write message
    sys.stdout.buffer.write(message_bytes)
    sys.stdout.buffer.flush()


def read_message():
    """Read a message from the extension via stdin."""
    # Native Messaging protocol: 4-byte length header (little-endian) + JSON message
    length_bytes = sys.stdin.buffer.read(4)
    if len(length_bytes) < 4:
        return None
    
    length = int.from_bytes(length_bytes, byteorder='little')
    if length == 0:
        return None
    
    message_bytes = sys.stdin.buffer.read(length)
    if len(message_bytes) < length:
        return None
    
    message_json = message_bytes.decode('utf-8')
    return json.loads(message_json)


def handle_attach_request(data):
    """Handle an attach request from the extension."""
    file_path = data.get('filePath')
    
    # Log to stderr (Chrome captures this for debugging)
    import sys
    sys.stderr.write(f"DEBUG: Received attach request for: {file_path}\n")
    sys.stderr.flush()
    
    if not file_path:
        return {
            'success': False,
            'message': 'Missing filePath in request'
        }
    
    # Create unique file copy
    unique_file_path, copy_error = create_unique_file_copy(file_path)
    if not unique_file_path:
        sys.stderr.write(f"DEBUG: Failed to create copy: {copy_error}\n")
        sys.stderr.flush()
        return {
            'success': False,
            'message': copy_error or 'Failed to create unique file copy'
        }
    
    sys.stderr.write(f"DEBUG: Created unique copy: {unique_file_path}\n")
    sys.stderr.flush()
    
    # Open Outlook with file attached
    system = platform.system()
    sys.stderr.write(f"DEBUG: Platform: {system}\n")
    sys.stderr.flush()
    
    if system == 'Darwin':
        success, message = open_outlook_mac(unique_file_path)
    elif system == 'Windows':
        success, message = open_outlook_windows(unique_file_path)
    else:
        return {
            'success': False,
            'message': f'Unsupported platform: {system}'
        }
    
    sys.stderr.write(f"DEBUG: Outlook open result: success={success}, message={message}\n")
    sys.stderr.flush()
    
    return {
        'success': success,
        'message': message
    }


def main():
    """Main loop: read messages from stdin and respond via stdout."""
    # Native Messaging hosts should not buffer stdout
    sys.stdout.reconfigure(line_buffering=False)
    
    # Log startup to stderr
    sys.stderr.write("DEBUG: Native host started\n")
    sys.stderr.flush()
    
    try:
        while True:
            message = read_message()
            if message is None:
                sys.stderr.write("DEBUG: Received None message, exiting\n")
                sys.stderr.flush()
                break
            
            sys.stderr.write(f"DEBUG: Received message: {message}\n")
            sys.stderr.flush()
            
            # Handle different message types
            if message.get('action') == 'attach':
                response = handle_attach_request(message)
                send_message(response)
            elif message.get('action') == 'ping':
                # Health check
                send_message({'success': True, 'message': 'pong'})
            else:
                send_message({
                    'success': False,
                    'message': f'Unknown action: {message.get("action")}'
                })
                
    except KeyboardInterrupt:
        sys.stderr.write("DEBUG: Interrupted by user\n")
        sys.stderr.flush()
        pass
    except Exception as e:
        import traceback
        error_msg = f'Internal error: {str(e)}\n{traceback.format_exc()}'
        sys.stderr.write(f"DEBUG: {error_msg}\n")
        sys.stderr.flush()
        # Send error response if possible
        try:
            send_message({
                'success': False,
                'message': f'Internal error: {str(e)}'
            })
        except:
            pass
        sys.exit(1)


if __name__ == '__main__':
    main()

