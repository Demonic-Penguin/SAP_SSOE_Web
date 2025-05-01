"""
Local SAP Connection Script
This script will run the SAP Service Order Automation using a real SAP connection
instead of the simulation mode. This requires:
1. Running on Windows
2. Having SAP GUI installed
3. Proper SAP system access
"""

import sys
import os
import tkinter as tk
from tkinter import messagebox
from sap_service_order_automation import SapServiceAutomation

def main():
    """Main function to run SAP automation with real SAP connection"""
    # Verify we're on Windows
    if sys.platform != "win32":
        print("Error: This script requires Windows to connect to a real SAP system.")
        return

    # Try to import win32com
    try:
        import win32com.client
    except ImportError:
        print("Error: win32com.client module is not available.")
        print("Please install it with: pip install pywin32")
        return
    
    # Create root Tkinter window for dialogs
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    # Display confirmation dialog
    response = messagebox.askyesno(
        "SAP Connection", 
        "This will attempt to connect to a real SAP system.\n\n"
        "Make sure:\n"
        "1. SAP GUI is installed\n"
        "2. You're logged into SAP\n"
        "3. You have proper access rights\n\n"
        "Continue?"
    )
    
    if not response:
        print("Operation canceled by user.")
        return
    
    # Create SAP automation instance with mock mode disabled
    app = SapServiceAutomation(use_mock=False)
    
    # Attempt to connect to SAP
    if app.connect_to_sap():
        print(f"Successfully connected to SAP as user: {app.username}")
        
        # Now run the main process
        app.main()
        
        # Run Tkinter main loop to ensure all dialogs are properly handled
        root.mainloop()
    else:
        messagebox.showerror(
            "SAP Connection Failed", 
            "Could not connect to SAP system.\n\n"
            "Please make sure:\n"
            "1. SAP GUI is running\n"
            "2. You're logged into the SAP system\n"
            "3. SAP scripting is enabled"
        )

if __name__ == "__main__":
    main()