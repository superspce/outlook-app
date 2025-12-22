#!/usr/bin/env python3
"""
Test script for Native Messaging Host
Simulates Chrome's communication with the native host
"""

import json
import subprocess
import sys
import os

def send_native_message(process, message):
    """Send a message using Native Messaging protocol."""
    message_json = json.dumps(message)
    message_bytes = message_json.encode('utf-8')
    length = len(message_bytes)
    
    # Write length as 4-byte little-endian integer
    process.stdin.buffer.write(length.to_bytes(4, byteorder='little'))
    # Write message
    process.stdin.buffer.write(message_bytes)
    process.stdin.flush()

def read_native_message(process):
    """Read a message using Native Messaging protocol."""
    # Read 4-byte length header
    length_bytes = process.stdout.buffer.read(4)
    if len(length_bytes) < 4:
        return None
    
    length = int.from_bytes(length_bytes, byteorder='little')
    if length == 0:
        return None
    
    # Read message
    message_bytes = process.stdout.buffer.read(length)
    if len(message_bytes) < length:
        return None
    
    message_json = message_bytes.decode('utf-8')
    return json.loads(message_json)

def main():
    """Test the native messaging host."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    host_script = os.path.join(script_dir, 'outlook-attach-native-host.py')
    
    if not os.path.exists(host_script):
        print(f"Error: Could not find {host_script}")
        sys.exit(1)
    
    print("Starting native host...")
    print("=" * 50)
    
    # Start the native host process
    process = subprocess.Popen(
        [sys.executable, host_script],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=False
    )
    
    try:
        # Test 1: Ping
        print("\n[Test 1] Sending ping...")
        send_native_message(process, {'action': 'ping'})
        response = read_native_message(process)
        if response and response.get('success'):
            print(f"✅ Ping successful: {response}")
        else:
            print(f"❌ Ping failed: {response}")
            return False
        
        # Test 2: Attach with test file (optional - requires a real file)
        if len(sys.argv) > 1:
            test_file = sys.argv[1]
            if os.path.exists(test_file):
                print(f"\n[Test 2] Testing attach with file: {test_file}")
                send_native_message(process, {
                    'action': 'attach',
                    'filePath': test_file
                })
                response = read_native_message(process)
                if response:
                    if response.get('success'):
                        print(f"✅ Attach successful: {response.get('message')}")
                    else:
                        print(f"⚠️  Attach returned error: {response.get('message')}")
                else:
                    print("❌ No response received")
            else:
                print(f"⚠️  Test file not found: {test_file}")
        else:
            print("\n[Test 2] Skipped (no test file provided)")
            print("   To test attach, run: python test-native-host.py /path/to/test-file.pdf")
        
        print("\n" + "=" * 50)
        print("✅ All tests completed!")
        return True
        
    except KeyboardInterrupt:
        print("\n\nTest interrupted by user")
        return False
    except Exception as e:
        print(f"\n❌ Error during testing: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        # Clean up
        process.terminate()
        try:
            process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            process.kill()

if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)

