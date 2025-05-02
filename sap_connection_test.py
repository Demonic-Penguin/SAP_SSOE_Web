"""
SAP Connection Test Script
Run this script to test your SAP connection and diagnose any issues
"""

import os
import sys
import traceback
import platform
import time
import json

# Check if we're on Windows
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

# Function to test SAP connection
def test_sap_connection():
    print("\n=== SAP CONNECTION TEST ===\n")
    print(f"Platform: {platform.system()}")
    print(f"Python Version: {sys.version}")
    
    results = {
        "platform": platform.system(),
        "python_version": sys.version,
        "steps": [],
        "success": False
    }
    
    # Step 1: Try to get the SAPGUI Automation object
    print("\nStep 1: Getting SAPGUI Automation object...")
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        print("✓ Successfully connected to SAPGUI")
        results["steps"].append({"step": "Get SAPGUI", "status": "success"})
    except Exception as e:
        error_msg = str(e)
        print(f"✗ Failed to connect to SAPGUI: {error_msg}")
        results["steps"].append({"step": "Get SAPGUI", "status": "failed", "error": error_msg})
        print("\nPossible causes:")
        print("  - SAP GUI is not running")
        print("  - SAP Scripting is not enabled")
        print("  - No active SAP session")
        
        print("\nRecommended actions:")
        print("  1. Make sure SAP GUI is running and you're logged in")
        print("  2. Enable scripting in SAP (see SAP Note 480149)")
        print("  3. Check Windows security settings")
        
        results["error"] = error_msg
        return results
    
    # Step 2: Get the Scripting Engine
    print("\nStep 2: Getting Scripting Engine...")
    try:
        application = sap_gui_auto.GetScriptingEngine
        if application is None:
            print("✗ Failed to get Scripting Engine (returned None)")
            results["steps"].append({"step": "Get Scripting Engine", "status": "failed", "error": "Returned None"})
            results["error"] = "Scripting Engine is None"
            return results
        print("✓ Successfully got Scripting Engine")
        results["steps"].append({"step": "Get Scripting Engine", "status": "success"})
    except Exception as e:
        error_msg = str(e)
        print(f"✗ Failed to get Scripting Engine: {error_msg}")
        results["steps"].append({"step": "Get Scripting Engine", "status": "failed", "error": error_msg})
        results["error"] = error_msg
        return results
    
    # Step 3: Get Connection
    print("\nStep 3: Getting Connection...")
    try:
        # Check how many connections are available
        conn_count = application.Children.Count
        print(f"Found {conn_count} SAP connection(s)")
        results["connection_count"] = conn_count
        
        if conn_count == 0:
            print("✗ No SAP connections found")
            results["steps"].append({"step": "Get Connection", "status": "failed", "error": "No connections"})
            results["error"] = "No SAP connections found"
            return results
        
        connection = application.Children(0)
        if connection is None:
            print("✗ Failed to get connection (returned None)")
            results["steps"].append({"step": "Get Connection", "status": "failed", "error": "Returned None"})
            results["error"] = "Connection is None"
            return results
        print("✓ Successfully got connection")
        results["steps"].append({"step": "Get Connection", "status": "success"})
    except Exception as e:
        error_msg = str(e)
        print(f"✗ Failed to get connection: {error_msg}")
        results["steps"].append({"step": "Get Connection", "status": "failed", "error": error_msg})
        results["error"] = error_msg
        return results
    
    # Step 4: Get Session
    print("\nStep 4: Getting Session...")
    try:
        # Check how many sessions are available
        sess_count = connection.Children.Count
        print(f"Found {sess_count} SAP session(s)")
        results["session_count"] = sess_count
        
        if sess_count == 0:
            print("✗ No SAP sessions found")
            results["steps"].append({"step": "Get Session", "status": "failed", "error": "No sessions"})
            results["error"] = "No SAP sessions found"
            return results
        
        session = connection.Children(0)
        if session is None:
            print("✗ Failed to get session (returned None)")
            results["steps"].append({"step": "Get Session", "status": "failed", "error": "Returned None"})
            results["error"] = "Session is None"
            return results
        print("✓ Successfully got session")
        results["steps"].append({"step": "Get Session", "status": "success"})
    except Exception as e:
        error_msg = str(e)
        print(f"✗ Failed to get session: {error_msg}")
        results["steps"].append({"step": "Get Session", "status": "failed", "error": error_msg})
        results["error"] = error_msg
        return results
    
    # Step 5: Get User Info
    print("\nStep 5: Getting User Info...")
    try:
        info = session.Info
        if info is None:
            print("✗ Failed to access session.Info (returned None)")
            results["steps"].append({"step": "Get Info", "status": "failed", "error": "Returned None"})
            results["error"] = "Info is None"
            return results
        
        user = info.User
        if not user:
            print("✗ Failed to get username (empty string)")
            results["steps"].append({"step": "Get Username", "status": "failed", "error": "Empty username"})
            results["error"] = "Empty username"
            return results
        
        print(f"✓ Successfully got user info: {user}")
        results["steps"].append({"step": "Get User Info", "status": "success"})
        results["username"] = user
    except Exception as e:
        error_msg = str(e)
        print(f"✗ Failed to get user info: {error_msg}")
        results["steps"].append({"step": "Get User Info", "status": "failed", "error": error_msg})
        results["error"] = error_msg
        return results
    
    # Step 6: Test FindById with status bar (should exist in any SAP screen)
    print("\nStep 6: Testing FindById with status bar...")
    try:
        status_bar = session.findById("wnd[0]/sbar")
        print(f"✓ Successfully found status bar")
        results["steps"].append({"step": "Test FindById", "status": "success"})
    except Exception as e:
        error_msg = str(e)
        print(f"✗ Failed to find status bar: {error_msg}")
        results["steps"].append({"step": "Test FindById", "status": "failed", "error": error_msg})
        results["warning"] = error_msg
        print("\nNOTE: This is not necessarily an error if your SAP screen doesn't have a status bar")
        print("     The test will continue, but be cautious about FindById working correctly")
    
    # Step 7: Get current transaction code
    print("\nStep 7: Getting current transaction code...")
    try:
        # The transaction code is visible in the system status field or title bar
        # We'll try both approaches
        transaction = "Unknown"
        
        try:
            # Try to get it from the command field
            cmd_field = session.findById("wnd[0]/tbar[0]/okcd")
            if cmd_field:
                transaction = cmd_field.text
                if not transaction:
                    transaction = "Empty (command field exists but is empty)"
        except Exception:
            pass
            
        if transaction == "Unknown":
            try:
                # Try to get it from the title bar
                title = session.findById("wnd[0]").text
                if title:
                    parts = title.split(" - ")
                    if len(parts) > 1:
                        transaction = parts[0].strip()
            except Exception:
                pass
        
        print(f"Current transaction: {transaction}")
        results["transaction"] = transaction
        results["steps"].append({"step": "Get Transaction", "status": "success"})
    except Exception as e:
        error_msg = str(e)
        print(f"✗ Failed to get transaction: {error_msg}")
        results["steps"].append({"step": "Get Transaction", "status": "warning", "error": error_msg})
    
    # All steps completed successfully
    print("\n=== SAP CONNECTION TEST SUCCESSFUL ===")
    print(f"Connected to SAP as user: {user}")
    results["success"] = True
    
    return results

# Run the test
try:
    results = test_sap_connection()
    
    # Save results to a file
    try:
        with open("sap_connection_test_results.json", "w") as f:
            json.dump(results, f, indent=2)
        print(f"\nTest results saved to sap_connection_test_results.json")
    except Exception as e:
        print(f"Failed to save test results: {e}")
    
    # Exit with appropriate code
    if results.get("success", False):
        print("\nYou're good to go! The SAP connection is working properly.")
        sys.exit(0)
    else:
        print("\nThe SAP connection test failed. Review the errors above and fix them before trying again.")
        sys.exit(1)
        
except Exception as e:
    print(f"\nUnexpected error during testing: {e}")
    print(traceback.format_exc())
    sys.exit(1)