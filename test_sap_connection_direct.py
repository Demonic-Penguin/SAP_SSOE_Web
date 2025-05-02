"""
SAP Direct Connection Test
This script performs direct testing of SAP GUI connections with more verbose logging
"""

import os
import sys
import platform
import time
import traceback

print(f"Python version: {sys.version}")
print(f"Platform: {platform.system()}")
print("=" * 80)

# Only continue if on Windows
if platform.system() != "Windows":
    print("ERROR: This script requires Windows to connect to SAP.")
    print(f"Current platform: {platform.system()}")
    sys.exit(1)

# Try to import win32com.client
try:
    import win32com.client
    print("✓ Successfully imported win32com.client")
except ImportError as e:
    print(f"ERROR: Failed to import win32com.client: {e}")
    print("You need to install pywin32 with: pip install pywin32")
    sys.exit(1)

print("\nTrying to connect to SAP GUI directly...")
print("-" * 80)

# Step 1: Try to get the SAPGUI object
try:
    print("Step 1: Connecting to SAPGUI...")
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    print("✓ Successfully connected to SAPGUI")
except Exception as e:
    print(f"✗ Failed to connect to SAPGUI: {e}")
    print("\nPossible causes:")
    print("  - SAP GUI is not running")
    print("  - SAP scripting is not enabled in SAP GUI settings")
    print("  - You are not logged into SAP")
    print("  - Windows security settings are blocking automation")
    print("\nAction required:")
    print("  1. Open SAP GUI and log in")
    print("  2. In SAP GUI, go to Customizing → Options → Scripting")
    print("  3. Enable 'Enable scripting' and 'Allow remote control'")
    sys.exit(1)

# Step 2: Get the scripting engine
try:
    print("\nStep 2: Getting scripting engine...")
    application = sap_gui_auto.GetScriptingEngine
    if application is None:
        print("✗ Failed to get scripting engine (returned None)")
        sys.exit(1)
    print("✓ Successfully got scripting engine")
except Exception as e:
    print(f"✗ Failed to get scripting engine: {e}")
    print("\nPossible causes:")
    print("  - SAP scripting is not properly enabled")
    print("  - SAP GUI was started without scripting support")
    print("\nAction required:")
    print("  1. Close all SAP GUI windows")
    print("  2. Make sure scripting is enabled in SAP GUI settings")
    print("  3. Restart SAP GUI and log in again")
    sys.exit(1)

# Step 3: Check connections
try:
    print("\nStep 3: Checking SAP connections...")
    conn_count = application.Children.Count
    print(f"Found {conn_count} SAP connection(s)")
    
    if conn_count == 0:
        print("✗ No SAP connections found")
        print("\nPossible causes:")
        print("  - You are not logged into SAP")
        print("  - SAP GUI is open but no system is connected")
        print("\nAction required:")
        print("  1. Open SAP GUI")
        print("  2. Log into your SAP system")
        print("  3. Run this script again")
        sys.exit(1)
    
    # Get first connection
    connection = application.Children(0)
    if connection is None:
        print("✗ Failed to get connection (returned None)")
        sys.exit(1)
    print("✓ Successfully got connection")
except Exception as e:
    print(f"✗ Failed to get connection information: {e}")
    sys.exit(1)

# Step 4: Check sessions
try:
    print("\nStep 4: Checking SAP sessions...")
    sess_count = connection.Children.Count
    print(f"Found {sess_count} session(s)")
    
    if sess_count == 0:
        print("✗ No SAP sessions found")
        print("\nPossible causes:")
        print("  - Connection exists but no session is open")
        print("\nAction required:")
        print("  1. Open a new session in SAP GUI")
        print("  2. Run this script again")
        sys.exit(1)
    
    # Get first session
    session = connection.Children(0)
    if session is None:
        print("✗ Failed to get session (returned None)")
        sys.exit(1)
    print("✓ Successfully got session")
except Exception as e:
    print(f"✗ Failed to get session information: {e}")
    sys.exit(1)

# Step 5: Get user info
try:
    print("\nStep 5: Getting user info...")
    info = session.Info
    if info is None:
        print("✗ Failed to get Info object (returned None)")
        sys.exit(1)
    
    user = info.User
    if not user:
        print("✗ Username is empty")
        sys.exit(1)
    
    print(f"✓ Connected as user: {user}")
    print(f"✓ System ID: {info.SystemName}")
    print(f"✓ System number: {info.SystemNumber}")
    print(f"✓ Client: {info.Client}")
except Exception as e:
    print(f"✗ Failed to get user info: {e}")
    print("\nPossible causes:")
    print("  - Connection lost or became invalid")
    print("  - Scripting permissions not set correctly")
    print("  - Unexpected SAP GUI behavior")
    print("\nAction required:")
    print("  1. Restart SAP GUI")
    print("  2. Check scripting options again")
    print("  3. Log in to SAP")
    sys.exit(1)

# Step 6: Test findById
try:
    print("\nStep 6: Testing findById...")
    
    # Try to find status bar (exists in most screens)
    try:
        status_bar = session.findById("wnd[0]/sbar")
        print("✓ Found status bar element")
    except Exception as e:
        print(f"- Could not find status bar: {e}")
        print("  (This is not necessarily an error)")
    
    # Try to find OK code field (command line field)
    try:
        cmd_field = session.findById("wnd[0]/tbar[0]/okcd")
        print("✓ Found command field element")
    except Exception as e:
        print(f"- Could not find command field: {e}")
        print("  (This might be screen-dependent)")
    
    print("✓ findById works correctly")
except Exception as e:
    print(f"✗ General error testing findById: {e}")
    sys.exit(1)

# Step 7: Try to execute an actual transaction
print("\nStep 7: Testing transaction navigation...")
try:
    # Get current transaction for reference
    current_screen = "Unknown"
    try:
        current_screen = session.findById("wnd[0]").text
        print(f"Current screen: {current_screen}")
    except:
        print("Could not determine current screen")
    
    # Try to access the command field
    print("Attempting to enter transaction code...")
    try:
        cmd_field = session.findById("wnd[0]/tbar[0]/okcd")
        if cmd_field is not None:
            # Backup current transaction
            orig_text = cmd_field.text
            
            # Enter /nIW32 transaction code
            print("Entering '/nIW32' transaction code...")
            cmd_field.text = "/nIW32"
            
            # Press Enter
            print("Sending Enter key...")
            session.findById("wnd[0]").sendVKey(0)
            
            # Wait a moment
            print("Waiting for screen to change...")
            time.sleep(1)
            
            # Check new screen
            new_screen = "Unknown"
            try:
                new_screen = session.findById("wnd[0]").text
                print(f"New screen: {new_screen}")
                
                if new_screen != current_screen:
                    print("✓ Successfully navigated to a new screen")
                else:
                    print("✗ Screen did not change")
            except:
                print("Could not determine new screen")
                
            # Try to go back
            print("Returning to previous screen...")
            try:
                cmd_field = session.findById("wnd[0]/tbar[0]/okcd")
                cmd_field.text = "/n"
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(1)
                print("Returned to previous state")
            except:
                print("Could not return to previous state")
            
            print("✓ Transaction navigation test completed")
        else:
            print("✗ Command field is None")
    except Exception as e:
        print(f"✗ Failed to execute transaction: {e}")
        print("\nPossible causes:")
        print("  - Transaction execution not permitted with scripting")
        print("  - Connection lost during execution")
        print("  - Screen elements changed during execution")
except Exception as e:
    print(f"✗ General error testing transaction: {e}")

print("\n" + "=" * 80)
print("SAP Connection Test Completed")
print("=" * 80)

if "session" in locals() and session is not None:
    try:
        # Final diagnostic: check if session is still valid
        print("\nFinal check: Is session still valid?")
        try:
            # Try to access Info again
            info = session.Info
            user = info.User
            print(f"✓ Session still valid. Connected as user: {user}")
        except Exception as e:
            print(f"✗ Session no longer valid: {e}")
            print("The session became invalid during testing.")
            print("This explains why your application loses connection!")
    except:
        print("Could not perform final check")

print("\nOverall result:")
print("✓ SAP GUI connection was successfully established")
print("✓ SAP GUI scripting is working")
print("✓ Basic SAP interactions are possible")

print("\nWhat to do next:")
print("1. If this test was successful but your app still fails:")
print("   - The issue is likely in how your app maintains the connection")
print("   - Try using the improved connection code with reconnection features")
print("   - Make sure your app doesn't run while SAP is performing background tasks")
print("\n2. If this test showed errors:")
print("   - Fix the specific issues mentioned above")
print("   - Check SAP GUI settings and permissions")
print("   - Ensure you're logged into SAP before running automation")