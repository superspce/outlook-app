#!/usr/bin/env python3
"""
Outlook Auto Attach Server
Receives file paths from Chrome extension and opens Outlook with file attached.
Supports both macOS (AppleScript) and Windows (COM automation).
"""

import http.server
import json
import sys
import os
import subprocess
import platform
import shutil
import tempfile
from datetime import datetime

# Server configuration
PORT = 8765


def create_unique_file_copy(original_path):
    """
    Create a unique copy of the file with a clean name format: Orderbekräftelse-datum-tid.pdf
    Includes microseconds to ensure uniqueness and avoid system-appended numbers.
    Returns the path to the unique copy.
    """
    if not os.path.exists(original_path):
        return None, f"File not found: {original_path}"
    
    try:
        # Get original file directory and extension
        original_dir = os.path.dirname(original_path)
        original_name = os.path.basename(original_path)
        name_parts = os.path.splitext(original_name)
        file_extension = name_parts[1]  # Keep original extension (.pdf, etc.)
        
        # Create clean filename based on file type: Orderbekräftelse-datum-tid.pdf or Inköp-datum-tid.pdf
        now = datetime.now()
        timestamp = now.strftime("%Y%m%d-%H%M%S")
        microseconds = now.strftime("%f")  # Include microseconds to prevent duplicates
        
        # Detect file type from original filename
        original_lower = original_name.lower()
        if 'inköp' in original_lower or 'inkop' in original_lower:
            base_name = "Inköp"
        else:
            base_name = "Orderbekräftelse"
        
        unique_name = f"{base_name}-{timestamp}-{microseconds}{file_extension}"
        unique_path = os.path.join(original_dir, unique_name)
        
        # Copy file to unique name
        shutil.copy2(original_path, unique_path)
        
        return unique_path, None
        
    except Exception as e:
        return None, f"Error creating unique copy: {str(e)}"


def open_outlook_mac(file_path):
    """Open Outlook on macOS using AppleScript and attach the file."""
    if not os.path.exists(file_path):
        return False, f"File not found: {file_path}"
    
    try:
        # Normalize path for AppleScript
        file_path = os.path.abspath(file_path)
        
        # AppleScript to create a new email in Outlook with attachment
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
        
        # Run AppleScript
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
        
        # Normalize path for Windows
        file_path = os.path.abspath(file_path)
        file_path = file_path.replace('/', '\\')
        
        # Create Outlook COM object
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        # Create new mail item
        mail_item = outlook.CreateItem(0)  # 0 = olMailItem
        
        # Attach the file
        mail_item.Attachments.Add(file_path)
        
        # Display the email (opens Outlook window)
        mail_item.Display()
        
        return True, "Outlook opened successfully"
        
    except ImportError:
        return False, "pywin32 not installed"
    except Exception as e:
        return False, f"Error: {str(e)}"


class AttachHandler(http.server.BaseHTTPRequestHandler):
    """HTTP request handler for /attach endpoint."""
    
    def log_message(self, format, *args):
        """Override to use custom log format with timestamp."""
        timestamp = datetime.now().strftime("[%d/%b/%Y %H:%M:%S]")
        sys.stderr.write(f"{timestamp} {format % args}\n")
    
    def do_OPTIONS(self):
        """Handle CORS preflight requests."""
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
    
    def do_POST(self):
        """Handle POST requests to /attach endpoint."""
        if self.path != '/attach':
            self.send_response(404)
            self.end_headers()
            return
        
        # Read request body
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length)
        
        try:
            data = json.loads(body.decode('utf-8'))
            file_path = data.get('filePath')
            
            if not file_path:
                self.send_error_response(400, "Missing filePath in request")
                return
            
            # Create a unique copy of the file with timestamp to avoid conflicts
            # This ensures we attach the exact file that was confirmed
            unique_file_path, copy_error = create_unique_file_copy(file_path)
            if not unique_file_path:
                self.send_error_response(500, copy_error or "Failed to create unique file copy")
                return
            
            # Use the unique file path for attachment
            file_to_attach = unique_file_path
            
            # Open Outlook based on platform
            system = platform.system()
            if system == 'Darwin':  # macOS
                success, message = open_outlook_mac(file_to_attach)
            elif system == 'Windows':
                success, message = open_outlook_windows(file_to_attach)
            else:
                # Clean up unique file if platform unsupported
                try:
                    os.remove(unique_file_path)
                except:
                    pass
                self.send_error_response(400, f"Unsupported platform: {system}")
                return
            
            # Send response
            response_data = {
                'success': success,
                'message': message
            }
            
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps(response_data).encode('utf-8'))
            
            # Log the result
            status = "Success" if success else "Failed"
            original_filename = os.path.basename(file_path)
            unique_filename = os.path.basename(file_to_attach)
            self.log_message(f"Attached file: {original_filename} (unique: {unique_filename}) - {status}: {success}")
            
            # Note: We keep the unique file copy (user can delete it later if needed)
            # This ensures the attachment always points to the correct file
            
        except json.JSONDecodeError:
            self.send_error_response(400, "Invalid JSON in request body")
        except Exception as e:
            self.send_error_response(500, f"Internal server error: {str(e)}")
    
    def send_error_response(self, status_code, message):
        """Send an error response with JSON body."""
        response_data = {
            'success': False,
            'message': message
        }
        
        self.send_response(status_code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(json.dumps(response_data).encode('utf-8'))
    
    def do_GET(self):
        """Handle GET requests - return simple status."""
        if self.path == '/' or self.path == '/status':
            self.send_response(200)
            self.send_header('Content-Type', 'text/plain')
            self.end_headers()
            self.wfile.write(b'Outlook Auto Attach Server is running')
        else:
            self.send_response(404)
            self.end_headers()


def main():
    """Start the HTTP server."""
    server_address = ('', PORT)
    httpd = http.server.HTTPServer(server_address, AttachHandler)
    
    print(f"Outlook Auto Attach server started on http://localhost:{PORT}")
    print("Press Ctrl+C to stop the server")
    
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nServer stopped")
        httpd.shutdown()


if __name__ == '__main__':
    main()

