"""
SAP GUI Screen Recorder
This script helps identify correct SAP screen element IDs by recording them as you click
Run this on a Windows machine with SAP GUI open and logged in
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

print("\nSAP Screen Element Recorder")
print("=" * 80)
print("This tool will help you identify the correct element IDs for your SAP screens.")
print("It creates a file with all element IDs it can find as you navigate in SAP.")
print("\nInstructions:")
print("1. Start SAP GUI and log in before running this script")
print("2. Run this script")
print("3. Navigate to the screens you want to record in SAP")
print("4. Press Ctrl+C in this console window when you're done")
print("\nNOTE: This will not interfere with your SAP session, it just observes.")
print("=" * 80)

try:
    # Connect to SAP
    print("\nConnecting to SAP GUI...")
    sap_gui_auto = win32com.client.GetObject("SAPGUI")
    application = sap_gui_auto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    
    # Create output file
    output_file = "sap_element_ids.txt"
    print(f"\nRecording SAP elements to {output_file}")
    print("Press Ctrl+C to stop recording")
    
    # Define which SAP GUI elements to try to find
    common_elements = [
        "wnd[0]/tbar[0]/okcd",  # Command field
        "wnd[0]/sbar",  # Status bar
        "wnd[0]/usr/ctxtAUFNR",  # Common service order field
        "wnd[0]/usr/ctxtRIWO00-AUFNR",  # Alternative service order field
        "wnd[0]/usr/ctxtVORG",  # Another alternative service order field
        "wnd[0]/usr/ctxtIW32-AUFNR",  # IW32 service order field
        "wnd[0]/usr/ctxtCAUFVD-AUFNR"  # Another service order field
    ]
    
    # Function to find all potentially useful elements
    def record_screen_elements():
        screen_elements = []
        current_screen = "Unknown"
        
        # Try to get screen title
        try:
            current_screen = session.findById("wnd[0]").text
            screen_elements.append(f"Screen Title: {current_screen}")
        except:
            screen_elements.append("Could not determine screen title")
        
        # Check common elements
        for element_id in common_elements:
            try:
                element = session.findById(element_id)
                if element is not None:
                    try:
                        element_text = element.text
                        screen_elements.append(f"Found: {element_id}, Text: {element_text}")
                    except:
                        screen_elements.append(f"Found: {element_id}, Text: <unavailable>")
            except:
                continue
        
        # Check ZIWBN-specific paths from VBS script
        ziwbn_paths = [
            "wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_KUNUM",
            "wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell"
        ]
        
        for element_id in ziwbn_paths:
            try:
                element = session.findById(element_id)
                if element is not None:
                    screen_elements.append(f"Found ZIWBN element: {element_id}")
            except:
                continue
        
        # Look for tab strips
        try:
            tabstrip = session.findById("wnd[0]/usr/tabsTABSTRIP")
            if tabstrip is not None:
                screen_elements.append(f"Found tab strip: wnd[0]/usr/tabsTABSTRIP")
                # Look for tab pages
                for i in range(1, 10):  # Try common tab numbers
                    try:
                        tab_id = f"wnd[0]/usr/tabsTABSTRIP/tabpT\\0{i}"
                        tab = session.findById(tab_id)
                        if tab is not None:
                            screen_elements.append(f"Found tab page: {tab_id}")
                    except:
                        continue
        except:
            pass
            
        # Check for grids/tables
        for i in range(1, 5):  # Try a few different grid patterns
            try:
                grid_id = f"wnd[0]/usr/cntlGRID{i}/shellcont/shell"
                grid = session.findById(grid_id)
                if grid is not None:
                    screen_elements.append(f"Found grid: {grid_id}")
                    # Try to get row count
                    try:
                        row_count = grid.RowCount
                        screen_elements.append(f"  Grid rows: {row_count}")
                    except:
                        screen_elements.append(f"  Could not get row count")
            except:
                continue
        
        return screen_elements
    
    # Main recording loop
    with open(output_file, "w") as f:
        f.write(f"SAP Screen Element Recording\n")
        f.write(f"Date: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"User: {session.Info.User}\n")
        f.write(f"System: {session.Info.SystemName}\n")
        f.write("=" * 80 + "\n\n")
        
        last_screen = ""
        try:
            while True:
                # Get current screen title
                try:
                    current_screen = session.findById("wnd[0]").text
                except:
                    current_screen = "Unknown"
                
                # Only record if screen changed
                if current_screen != last_screen:
                    print(f"\nNew screen detected: {current_screen}")
                    
                    # Record elements
                    elements = record_screen_elements()
                    
                    # Write to file
                    f.write(f"\n\n{'=' * 40}\n")
                    f.write(f"Screen: {current_screen}\n")
                    f.write(f"Time: {time.strftime('%H:%M:%S')}\n")
                    f.write(f"{'-' * 40}\n")
                    
                    for element in elements:
                        f.write(f"{element}\n")
                        print(f"  Recorded: {element}")
                    
                    f.flush()  # Make sure it's written immediately
                    last_screen = current_screen
                
                # Wait before checking again
                time.sleep(1)
                
        except KeyboardInterrupt:
            print("\n\nRecording stopped by user")
            f.write(f"\n\nRecording stopped by user at {time.strftime('%H:%M:%S')}")
            print(f"\nSAP element IDs have been saved to {output_file}")
            print("Use these IDs to update your SAP connection code")

except Exception as e:
    print(f"\nError: {e}")
    print(traceback.format_exc())
    print("\nPlease make sure SAP GUI is open and you are logged in")
    sys.exit(1)