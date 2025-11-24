#!/usr/bin/env python3
"""
Local web server for Outlook Auto Attach Chrome extension
This server receives file paths from the Chrome extension and opens Outlook with the file attached
"""

import http.server
import json
import socketserver
import platform
import os
import sys
from urllib.parse import urlparse, parse_qs

# Platform-specific imports
if platform.system() == 'Darwin':  # Mac
    import subprocess
elif platform.system() == 'Windows':  # Windows
    try:
        import win32com.client
    except ImportError:
        print("Warning: pywin32 not installed. Install it with: pip install pywin32", file=sys.stderr)

PORT = 8765

def open_outlook_mac(file_path):
    """Open Outlook on Mac using AppleScript."""
    try:
        # Ensure file path is absolute and exists
        if not os.path.isabs(file_path):
            file_path = os.path.abspath(file_path)
        
        if not os.path.isfile(file_path):
            return {"success": False, "message": f"File not found: {file_path}"}
        
        # Escape file path for AppleScript
        escaped_path = file_path.replace("\\", "\\\\").replace('"', '\\"')
        
        # AppleScript to open Outlook and attach file
        # Email message text
        message_text = "Hej,\n\nHär kommer din orderbekräftelse"
        # Escape message text for AppleScript
        escaped_message = message_text.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")
        
        # AppleScript to open Outlook and attach file
        applescript = f'''
        tell application "Microsoft Outlook"
            activate
            set newMessage to make new outgoing message
            tell newMessage
                set content to "{escaped_message}"
                make new attachment with properties {{file: POSIX file "{escaped_path}"}}
            end tell
            open newMessage
        end tell
        '''
        
        result = subprocess.run(
            ["osascript", "-e", applescript],
            capture_output=True,
            text=True,
            timeout=10
        )
        
        if result.returncode == 0:
            return {"success": True, "message": "Outlook opened with file attached successfully"}
        else:
            return {"success": False, "message": f"AppleScript error: {result.stderr}"}
    
    except subprocess.TimeoutExpired:
        return {"success": False, "message": "Timeout while trying to open Outlook"}
    except Exception as e:
        return {"success": False, "message": f"Error: {str(e)}"}

def open_outlook_windows(file_path):
    """Open Outlook on Windows using COM automation."""
    try:
        # Ensure file path is absolute and exists
        if not os.path.isabs(file_path):
            file_path = os.path.abspath(file_path)
        
        # Normalize Windows paths (handle forward slashes)
        file_path = file_path.replace('/', '\\')
        file_path = os.path.normpath(file_path)
        
        if not os.path.isfile(file_path):
            return {"success": False, "message": f"File not found: {file_path}"}
        
        # Create Outlook application object
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        # Create a new mail item (0 = olMailItem)
        mail = outlook.CreateItem(0)
        
        # Set email body with message
        mail.Body = "Hej,\n\nHär kommer din orderbekräftelse"
        
        # Add attachment
        mail.Attachments.Add(file_path)
        
        # Display the email window
        mail.Display()
        
        return {"success": True, "message": "Outlook opened with file attached successfully"}
    
    except ImportError:
        return {"success": False, "message": "pywin32 not installed. Install with: pip install pywin32"}
    except Exception as e:
        return {"success": False, "message": f"Error: {str(e)}"}

def open_outlook_with_attachment(file_path):
    """Open Outlook with attachment, platform-specific."""
    system = platform.system()
    
    if system == 'Darwin':
        return open_outlook_mac(file_path)
    elif system == 'Windows':
        return open_outlook_windows(file_path)
    else:
        return {"success": False, "message": f"Unsupported operating system: {system}"}

class OutlookAttachHandler(http.server.SimpleHTTPRequestHandler):
    """HTTP request handler for Outlook Auto Attach"""
    
    def do_GET(self):
        """Handle GET requests"""
        parsed_path = urlparse(self.path)
        
        if parsed_path.path == '/attach':
            # Get file path from query parameter
            query_params = parse_qs(parsed_path.query)
            file_path = query_params.get('file', [None])[0]
            
            if file_path:
                # Decode URL encoding
                file_path = file_path.replace('%20', ' ')
                
                # Open Outlook with attachment
                result = open_outlook_with_attachment(file_path)
                
                # Send JSON response
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                
                response_json = json.dumps(result).encode('utf-8')
                self.wfile.write(response_json)
                
                # Log to console
                print(f"[{self.log_date_time_string()}] Attached file: {file_path} - Success: {result.get('success', False)}")
            else:
                self.send_error(400, "Missing 'file' parameter")
        else:
            self.send_error(404, "Not Found")
    
    def do_POST(self):
        """Handle POST requests"""
        if self.path == '/attach':
            try:
                # Read request body
                content_length = int(self.headers.get('Content-Length', 0))
                body = self.rfile.read(content_length).decode('utf-8')
                data = json.loads(body)
                
                file_path = data.get('filePath')
                
                if file_path:
                    # Open Outlook with attachment
                    result = open_outlook_with_attachment(file_path)
                    
                    # Send JSON response
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/json')
                    self.send_header('Access-Control-Allow-Origin', '*')
                    self.end_headers()
                    
                    response_json = json.dumps(result).encode('utf-8')
                    self.wfile.write(response_json)
                    
                    # Log to console
                    print(f"[{self.log_date_time_string()}] Attached file: {file_path} - Success: {result.get('success', False)}")
                else:
                    self.send_error(400, "Missing 'filePath' in request body")
            
            except json.JSONDecodeError:
                self.send_error(400, "Invalid JSON in request body")
            except Exception as e:
                self.send_error(500, f"Internal server error: {str(e)}")
        else:
            self.send_error(404, "Not Found")
    
    def log_message(self, format, *args):
        """Override to prevent default logging"""
        # Only log important messages
        if '200' in format or 'ERROR' in format:
            print(f"[{self.log_date_time_string()}] {format % args}")

def main():
    """Start the web server"""
    with socketserver.TCPServer(("localhost", PORT), OutlookAttachHandler) as httpd:
        print(f"Outlook Auto Attach server started on http://localhost:{PORT}")
        print("Press Ctrl+C to stop the server")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServer stopped")

if __name__ == '__main__':
    main()

