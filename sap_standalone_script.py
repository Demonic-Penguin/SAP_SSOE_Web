"""
SAP Standalone Script for Service Order Automation
This is a standalone version of the service order automation that doesn't require the web interface
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

class SapServiceOrderAutomation:
    def __init__(self):
        """Initialize SAP connection"""
        self.sap_gui_auto = None
        self.application = None
        self.connection = None
        self.session = None
        self.username = "UNKNOWN"
        self.is_connected = False
        
    def connect_to_sap(self):
        """Connect to SAP GUI"""
        print("\nConnecting to SAP GUI...")
        try:
            # Get SAP GUI scripting object
            self.sap_gui_auto = win32com.client.GetObject("SAPGUI")
            print("✓ Connected to SAPGUI")
            
            # Get the scripting engine
            self.application = self.sap_gui_auto.GetScriptingEngine
            if self.application is None:
                print("✗ Failed to get scripting engine (returned None)")
                return False
            print("✓ Got scripting engine")
            
            # Check connections
            conn_count = self.application.Children.Count
            print(f"Found {conn_count} SAP connection(s)")
            if conn_count == 0:
                print("✗ No SAP connections found")
                return False
            
            # Get first connection
            self.connection = self.application.Children(0)
            if self.connection is None:
                print("✗ Failed to get connection (returned None)")
                return False
            print("✓ Got connection")
            
            # Check sessions
            sess_count = self.connection.Children.Count
            print(f"Found {sess_count} session(s)")
            if sess_count == 0:
                print("✗ No SAP sessions found")
                return False
            
            # Get first session
            self.session = self.connection.Children(0)
            if self.session is None:
                print("✗ Failed to get session (returned None)")
                return False
            print("✓ Got session")
            
            # Get user info
            info = self.session.Info
            self.username = info.User
            print(f"✓ Connected as user: {self.username}")
            
            self.is_connected = True
            return True
            
        except Exception as e:
            print(f"✗ Failed to connect to SAP: {e}")
            print(traceback.format_exc())
            return False
    
    def check_connection(self):
        """Check if the SAP connection is still valid"""
        if not self.is_connected or self.session is None:
            print("Connection not established. Connecting...")
            return self.connect_to_sap()
            
        try:
            # Try to access session info to check connection
            info = self.session.Info
            username = info.User
            return True
        except Exception as e:
            print(f"Connection lost: {e}")
            print("Reconnecting...")
            return self.connect_to_sap()
    
    def safe_find_by_id(self, element_id, max_retries=3):
        """Safely find SAP element with error handling and retry logic"""
        if not self.check_connection():
            return None
            
        for attempt in range(max_retries):
            try:
                element = self.session.findById(element_id)
                return element
            except Exception as e:
                print(f"Error finding element {element_id} (attempt {attempt+1}): {e}")
                if attempt < max_retries - 1:
                    print(f"Retrying...")
                    self.check_connection()
                    time.sleep(0.5)
                else:
                    print(f"Failed to find element after {max_retries} attempts")
                    return None
    
    def navigate_to_transaction(self, transaction_code):
        """Navigate to a SAP transaction"""
        print(f"\nNavigating to transaction {transaction_code}...")
        
        if not self.check_connection():
            return False
            
        try:
            # Get command field
            cmd_field = self.safe_find_by_id("wnd[0]/tbar[0]/okcd")
            if cmd_field is None:
                print("✗ Could not find command field")
                return False
            
            # Enter transaction code with /n prefix to ensure new transaction
            cmd_field.text = f"/n{transaction_code}"
            
            # Press Enter
            window = self.safe_find_by_id("wnd[0]")
            if window is None:
                print("✗ Could not find main window")
                return False
                
            window.sendVKey(0)
            
            # Wait for screen to change
            time.sleep(1)
            
            print(f"✓ Navigated to {transaction_code}")
            return True
            
        except Exception as e:
            print(f"✗ Failed to navigate to transaction: {e}")
            print(traceback.format_exc())
            return False
    
    def enter_service_order(self, service_order):
        """Enter service order number"""
        print(f"\nEntering service order {service_order}...")
        
        if not self.check_connection():
            return False
            
        try:
            # Try different possible service order field IDs
            order_fields = [
                "wnd[0]/usr/ctxtAUFNR",
                "wnd[0]/usr/ctxtRIWO00-AUFNR",
                "wnd[0]/usr/ctxtVORG",
                "wnd[0]/usr/ctxtIW32-AUFNR",
                "wnd[0]/usr/ctxtCAUFVD-AUFNR"
            ]
            
            field_found = False
            for field_id in order_fields:
                order_field = self.safe_find_by_id(field_id)
                if order_field is not None:
                    # Found a valid field
                    print(f"Found order field: {field_id}")
                    order_field.text = service_order
                    field_found = True
                    break
            
            if not field_found:
                print("✗ Could not find service order input field")
                return False
            
            # Press Enter
            window = self.safe_find_by_id("wnd[0]")
            if window is None:
                print("✗ Could not find main window")
                return False
                
            window.sendVKey(0)
            
            # Wait for data to load
            time.sleep(1)
            
            print(f"✓ Entered service order {service_order}")
            return True
            
        except Exception as e:
            print(f"✗ Failed to enter service order: {e}")
            print(traceback.format_exc())
            return False
    
    def get_service_order_data(self, service_order):
        """Get data for a service order"""
        print(f"\nGetting data for service order {service_order}...")
        
        # First navigate to IW32
        if not self.navigate_to_transaction("IW32"):
            return None
        
        # Enter service order
        if not self.enter_service_order(service_order):
            return None
        
        # Data dictionary to collect information
        data = {
            'service_order': service_order,
            'part_number': None,
            'serial_number': None,
            'customer': None,
            'op_comments': None,
            'mod_status': None
        }
        
        # Try to extract data
        try:
            # Look for part number
            part_fields = [
                "wnd[0]/usr/tabsTABSTRIP/tabpT\\01/ssubSUB_DATA:SAPLIQS0:7235/subGENERAL:SAPLIQS0:7212/txtLTAP-MATNR",
                "wnd[0]/usr/tabsTABSTRIP/tabpDESC/ssubDETAIL:SAPLITO0:0115/txtITOBJ-MATXT",
                "wnd[0]/usr/ctxtRIWO00-MATNR"
            ]
            
            for field_id in part_fields:
                field = self.safe_find_by_id(field_id)
                if field is not None:
                    try:
                        data['part_number'] = field.text
                        print(f"Found part number: {data['part_number']}")
                        break
                    except:
                        pass
            
            # Try to find customer information
            customer_fields = [
                "wnd[0]/usr/tabsTABSTRIP/tabpT\\01/ssubSUB_DATA:SAPLIQS0:7235/subCUSTOMER:SAPLIQS0:7280/txtKUAGV-NAME1",
                "wnd[0]/usr/ctxtRIWO00-KUNUM"
            ]
            
            for field_id in customer_fields:
                field = self.safe_find_by_id(field_id)
                if field is not None:
                    try:
                        data['customer'] = field.text
                        print(f"Found customer: {data['customer']}")
                        break
                    except:
                        pass
            
            # Try to switch to Equipment tab to get serial number
            try:
                tab_ids = [
                    "wnd[0]/usr/tabsTABSTRIP/tabpT\\02",
                    "wnd[0]/usr/tabsTABSTRIP/tabpEQUIPMENT"
                ]
                
                tab_selected = False
                for tab_id in tab_ids:
                    tab = self.safe_find_by_id(tab_id)
                    if tab is not None:
                        tab.select()
                        print(f"Selected tab: {tab_id}")
                        tab_selected = True
                        time.sleep(0.5)
                        break
                
                if tab_selected:
                    # Try to get serial number
                    serial_fields = [
                        "wnd[0]/usr/tabsTABSTRIP/tabpT\\02/ssubSUB_DATA:SAPLIQS0:7236/subOBJ:SAPLIQS0:7322/txtVIQMEL-SERGE",
                        "wnd[0]/usr/tabsTABSTRIP/tabpEQUIPMENT/ssubDETAIL:SAPLITO0:0115/txtITOB-SERGE"
                    ]
                    
                    for field_id in serial_fields:
                        field = self.safe_find_by_id(field_id)
                        if field is not None:
                            try:
                                data['serial_number'] = field.text
                                print(f"Found serial number: {data['serial_number']}")
                                break
                            except:
                                pass
            except Exception as e:
                print(f"Error getting equipment data: {e}")
            
            print(f"\nData collected for service order {service_order}:")
            for key, value in data.items():
                print(f"  - {key}: {value}")
            
            return data
            
        except Exception as e:
            print(f"✗ Failed to get service order data: {e}")
            print(traceback.format_exc())
            return None

    def navigate_to_ziwbn(self, service_order):
        """Navigate to ZIWBN transaction with service order"""
        print(f"\nNavigating to ZIWBN for service order {service_order}...")
        
        # First navigate to ZIWBN
        if not self.navigate_to_transaction("ZIWBN"):
            return False
        
        try:
            # Try to find the service order input field
            input_field = self.safe_find_by_id("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB1:SAPLYAFF_ZIWBNGUI:0100/ssubSUB2:SAPLYAFF_ZIWBNGUI:0102/ctxtW_INP_DATA")
            
            if input_field is None:
                print("✗ Could not find ZIWBN service order input field")
                return False
            
            # Enter service order
            input_field.text = service_order
            
            # Press Enter
            window = self.safe_find_by_id("wnd[0]")
            if window is None:
                print("✗ Could not find main window")
                return False
                
            window.sendVKey(0)
            
            # Wait for data to load
            time.sleep(1)
            
            print(f"✓ Navigated to ZIWBN for service order {service_order}")
            return True
            
        except Exception as e:
            print(f"✗ Failed to navigate to ZIWBN: {e}")
            print(traceback.format_exc())
            return False

    def get_ziwbn_data(self, service_order):
        """Get data from ZIWBN transaction"""
        print(f"\nGetting ZIWBN data for service order {service_order}...")
        
        # Navigate to ZIWBN
        if not self.navigate_to_ziwbn(service_order):
            return None
        
        # Data dictionary to collect information
        data = {
            'service_order': service_order,
            'part_number': None,
            'serial_number': None,
            'customer': None,
            'customer_number': None
        }
        
        try:
            # Get customer number from service order header
            customer_field = self.safe_find_by_id("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_KUNUM")
            if customer_field is not None:
                data['customer_number'] = customer_field.text
                print(f"Found customer number: {data['customer_number']}")
            
            # Switch to Equipment tab
            equip_tab = self.safe_find_by_id("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H")
            if equip_tab is not None:
                equip_tab.select()
                time.sleep(0.5)
                print("Selected Equipment tab")
                
                # Try different versions of the grid
                try:
                    # First try with standard shell
                    grid = self.safe_find_by_id("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell")
                    if grid is not None:
                        print("Found equipment grid")
                        data['part_number'] = grid.getCellValue(0, "MATNR")
                        data['serial_number'] = grid.getCellValue(0, "SERNR")
                        print(f"Found part number: {data['part_number']}")
                        print(f"Found serial number: {data['serial_number']}")
                except Exception as e1:
                    print(f"Error with standard grid: {e1}")
                    
                    # Try version 10 method
                    try:
                        grid = self.safe_find_by_id("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont[0]/shell")
                        if grid is not None:
                            print("Found version 10 equipment grid")
                            data['part_number'] = grid.getCellValue(0, "MATNR")
                            data['serial_number'] = grid.getCellValue(0, "SERNR")
                            print(f"Found part number: {data['part_number']}")
                            print(f"Found serial number: {data['serial_number']}")
                    except Exception as e2:
                        print(f"Error with version 10 grid: {e2}")
            
            print(f"\nZIWBN data collected for service order {service_order}:")
            for key, value in data.items():
                print(f"  - {key}: {value}")
            
            return data
            
        except Exception as e:
            print(f"✗ Failed to get ZIWBN data: {e}")
            print(traceback.format_exc())
            return None
    
    def main(self, service_order):
        """Main function to execute service order automation"""
        # Connect to SAP
        if not self.connect_to_sap():
            print("Failed to connect to SAP")
            return False
        
        # Try to get data from ZIWBN first (as the original script does)
        ziwbn_data = self.get_ziwbn_data(service_order)
        
        # If ZIWBN didn't work, try IW32
        if ziwbn_data is None or (ziwbn_data['part_number'] is None and ziwbn_data['serial_number'] is None):
            print("\nZIWBN data incomplete, trying IW32...")
            iw32_data = self.get_service_order_data(service_order)
            
            if iw32_data is None:
                print("Failed to get service order data from both ZIWBN and IW32")
                return False
                
            print("\nService order data retrieved successfully")
            return iw32_data
        else:
            print("\nService order data retrieved successfully from ZIWBN")
            return ziwbn_data

# Main execution
if __name__ == "__main__":
    print("SAP Service Order Automation - Standalone Version")
    print("=" * 80)
    
    # Create automation instance
    sap = SapServiceOrderAutomation()
    
    # Get service order from user
    service_order = input("\nEnter service order number: ")
    if not service_order:
        print("No service order entered. Exiting.")
        sys.exit(1)
    
    # Run automation
    result = sap.main(service_order)
    
    if result:
        print("\nAutomation completed successfully!")
        print("\nService Order Data:")
        for key, value in result.items():
            print(f"  {key}: {value}")
    else:
        print("\nAutomation failed.")
        
    print("\nDone.")