"""
SAP Connection Debug and Fix
This module provides improved SAP connection handling with detailed error reporting
"""

import sys
import os
import platform
import time
import traceback

# Flag to indicate if we're running in Replit environment
IN_REPLIT = 'REPL_ID' in os.environ if 'os' in sys.modules else False

class SapConnection:
    """
    A more robust SAP connection handler with detailed error reporting and debugging
    """
    def __init__(self):
        self.sap_gui_auto = None
        self.application = None
        self.connection = None
        self.session = None
        self.username = "UNKNOWN"
        self.is_connected = False
        self.connection_error = None
        self.debug_info = {}
        
    def connect(self):
        """Connect to SAP with enhanced error handling"""
        # Only attempt connection on Windows
        if platform.system() != "Windows":
            self.connection_error = "SAP connection requires Windows"
            return False
            
        # Try to import win32com
        try:
            import win32com.client
            self.debug_info["win32com_import"] = "Success"
        except ImportError as e:
            self.connection_error = f"Failed to import win32com.client: {str(e)}"
            self.debug_info["win32com_import"] = f"Failed: {str(e)}"
            return False
            
        # Step 1: Check if SAP GUI is running by attempting to get the SAPGUI object
        try:
            self.sap_gui_auto = win32com.client.GetObject("SAPGUI")
            self.debug_info["sap_gui_auto"] = "Success"
        except Exception as e:
            self.connection_error = f"Failed to connect to SAPGUI: {str(e)}"
            self.debug_info["sap_gui_auto"] = f"Failed: {str(e)}"
            return False
            
        # Step 2: Get the scripting engine
        try:
            if self.sap_gui_auto is None:
                self.connection_error = "SAP GUI Automation object is None"
                self.debug_info["application"] = "Failed: SapGuiAuto is None"
                return False
                
            self.application = self.sap_gui_auto.GetScriptingEngine
            if self.application is None:
                self.connection_error = "Failed to get SAP scripting engine"
                self.debug_info["application"] = "Failed: GetScriptingEngine returned None"
                return False
                
            self.debug_info["application"] = "Success"
        except Exception as e:
            self.connection_error = f"Failed to get SAP scripting engine: {str(e)}"
            self.debug_info["application"] = f"Failed: {str(e)}"
            return False
            
        # Step 3: Get the SAP connection
        try:
            if self.application is None:
                self.connection_error = "SAP application object is None"
                self.debug_info["connection"] = "Failed: Application is None"
                return False
                
            # Check how many connections exist
            conn_count = 0
            try:
                conn_count = self.application.Children.Count
                self.debug_info["connection_count"] = conn_count
            except Exception as conn_e:
                self.debug_info["connection_count_error"] = str(conn_e)
            
            if conn_count == 0:
                self.connection_error = "No active SAP connections found. Please log into SAP first."
                self.debug_info["connection"] = "Failed: No connections"
                return False
                
            self.connection = self.application.Children(0)
            if self.connection is None:
                self.connection_error = "Failed to get SAP connection"
                self.debug_info["connection"] = "Failed: Children(0) returned None"
                return False
                
            self.debug_info["connection"] = "Success"
        except Exception as e:
            self.connection_error = f"Failed to get SAP connection: {str(e)}"
            self.debug_info["connection"] = f"Failed: {str(e)}"
            return False
            
        # Step 4: Get the SAP session
        try:
            if self.connection is None:
                self.connection_error = "SAP connection object is None"
                self.debug_info["session"] = "Failed: Connection is None"
                return False
                
            # Check how many sessions exist
            sess_count = 0
            try:
                sess_count = self.connection.Children.Count
                self.debug_info["session_count"] = sess_count
            except Exception as sess_e:
                self.debug_info["session_count_error"] = str(sess_e)
            
            if sess_count == 0:
                self.connection_error = "No active SAP sessions found"
                self.debug_info["session"] = "Failed: No sessions"
                return False
                
            self.session = self.connection.Children(0)
            if self.session is None:
                self.connection_error = "Failed to get SAP session"
                self.debug_info["session"] = "Failed: Children(0) returned None"
                return False
                
            self.debug_info["session"] = "Success"
        except Exception as e:
            self.connection_error = f"Failed to get SAP session: {str(e)}"
            self.debug_info["session"] = f"Failed: {str(e)}"
            return False
            
        # Step 5: Get the user name to verify connection
        try:
            if self.session is None:
                self.connection_error = "SAP session object is None"
                self.debug_info["username"] = "Failed: Session is None"
                return False
                
            # Test if Info exists and is accessible
            try:
                info = self.session.Info
                if info is None:
                    self.connection_error = "Cannot access session.Info"
                    self.debug_info["info"] = "Failed: Info is None"
                    return False
                    
                self.debug_info["info"] = "Success"
            except Exception as info_e:
                self.connection_error = f"Failed to access session.Info: {str(info_e)}"
                self.debug_info["info"] = f"Failed: {str(info_e)}"
                return False
                
            # Test if we can access the username
            try:
                self.username = self.session.Info.User
                if not self.username:
                    self.connection_error = "Failed to get username from SAP"
                    self.debug_info["username"] = "Failed: Username is empty"
                    return False
                    
                self.debug_info["username"] = f"Success: {self.username}"
            except Exception as usr_e:
                self.connection_error = f"Failed to access username: {str(usr_e)}"
                self.debug_info["username"] = f"Failed: {str(usr_e)}"
                return False
                
        except Exception as e:
            self.connection_error = f"Failed to get username: {str(e)}"
            self.debug_info["username"] = f"Failed: {str(e)}"
            return False
            
        # If we got here, the connection is successful
        self.is_connected = True
        return True
        
    def test_find_by_id(self):
        """
        Test the findById method to make sure it's working
        This tests both a valid and invalid ID to see what happens
        """
        if not self.is_connected or self.session is None:
            return {
                "status": "error", 
                "message": "Not connected to SAP", 
                "error": self.connection_error
            }
            
        try:
            # Test findById with a common valid element (status bar)
            try:
                status_bar = self.session.findById("wnd[0]/sbar")
                result = {
                    "status": "success",
                    "valid_id_test": "Success",
                    "status_bar_found": status_bar is not None
                }
            except Exception as valid_e:
                result = {
                    "status": "warning",
                    "valid_id_test": f"Failed: {str(valid_e)}",
                    "trace": traceback.format_exc()
                }
                
            # Test with an intentionally invalid ID to see error behavior
            try:
                invalid = self.session.findById("not/a/real/sap/id")
                result["invalid_id_test"] = "Didn't raise expected error"
            except Exception as invalid_e:
                result["invalid_id_test"] = f"Expected error: {str(invalid_e)}"
                
            return result
        except Exception as e:
            return {
                "status": "error",
                "message": f"Error testing findById: {str(e)}",
                "trace": traceback.format_exc()
            }
            
    def get_connection_details(self):
        """
        Return detailed information about the connection state
        """
        return {
            "platform": platform.system(),
            "is_connected": self.is_connected,
            "connection_error": self.connection_error,
            "username": self.username,
            "debug_info": self.debug_info,
            "sap_gui_auto": self.sap_gui_auto is not None,
            "application": self.application is not None,
            "connection": self.connection is not None,
            "session": self.session is not None
        }
        
    def check_and_reconnect(self):
        """
        Check if SAP connection is still valid and try to reconnect if not
        Returns True if connected (or reconnected successfully), False otherwise
        """
        try:
            # Quick test to see if session is still valid
            if self.session is None:
                print("Session is None, attempting to reconnect")
                return self.connect()
                
            # Try to access something simple in SAP
            try:
                # Just check if we can get the Info property
                info = self.session.Info
                if info is None:
                    print("Session Info is None, attempting to reconnect")
                    return self.connect()
                    
                # Check if we can still get the username
                username = info.User
                if not username:
                    print("Username is empty, attempting to reconnect")
                    return self.connect()
                    
                # Connection still looks good
                return True
            except Exception as e:
                print(f"Error checking SAP connection: {e}")
                print("Attempting to reconnect...")
                return self.connect()
                
        except Exception as e:
            print(f"Error in check_and_reconnect: {e}")
            return False
            
    def safe_find_by_id(self, element_id, max_retries=3):
        """
        Safely find an element by ID with retry logic and connection check
        """
        if not self.is_connected or self.session is None:
            if not self.check_and_reconnect():
                return None
                
        errors = []
        
        for attempt in range(max_retries):
            try:
                element = self.session.findById(element_id)
                return element
            except Exception as e:
                errors.append(f"Attempt {attempt+1}: {str(e)}")
                
                # Try to reconnect before next attempt
                if not self.check_and_reconnect():
                    break
                    
                # Wait a bit before retry
                time.sleep(0.5)
                
        # All attempts failed
        print(f"Failed to find element {element_id} after {max_retries} attempts")
        print(f"Errors: {errors}")
        return None
            
    def get_service_order_data(self, service_order):
        """
        Get service order data with detailed error reporting
        """
        if not self.is_connected or self.session is None:
            if not self.check_and_reconnect():
                return {
                    "status": "error", 
                    "message": "Not connected to SAP", 
                    "error": self.connection_error
                }
            
        try:
            # Create a record of each step for debugging
            steps = []
            
            # Step 1: Navigate to the IW32 transaction
            try:
                # Get the command field
                steps.append("Attempting to get command field")
                cmd_field = self.safe_find_by_id("wnd[0]/tbar[0]/okcd").text = "/nZIWBN"
                
                if cmd_field is None:
                    raise Exception("Could not find command field")
                    
                steps.append("Got command field")
                
                # Enter the IW32 transaction code
                steps.append("Entering transaction code")
                cmd_field.text = "/nIW32"
                steps.append("Transaction code entered")
                
                # Press Enter
                steps.append("Sending VKey(0)")
                window = self.safe_find_by_id("wnd[0]")
                
                if window is None:
                    raise Exception("Could not find main window")
                    
                window.sendVKey(0)
                steps.append("VKey sent successfully")
                
                # Wait a moment for the screen to load
                steps.append("Waiting for screen to load")
                time.sleep(1)
                steps.append("Wait complete")
            except Exception as e:
                return {
                    "status": "error",
                    "message": f"Failed to navigate to IW32: {str(e)}",
                    "steps": steps,
                    "trace": traceback.format_exc()
                }
            
            # Step 2: Enter the service order number
            try:
                # Find the service order input field
                steps.append("Attempting to find service order input field")
                
                # Different SAP systems might use different field IDs
                # Try a few common ones
                so_field = None
                
                # Try different possible field IDs
                service_order_field_ids = [
                    "wnd[0]/usr/ctxtAUFNR",
                    "wnd[0]/usr/ctxtRIWO00-AUFNR",
                    "wnd[0]/usr/ctxtVORG",
                    "wnd[0]/usr/ctxtIW32-AUFNR",
                    "wnd[0]/usr/ctxtCAUFVD-AUFNR"
                ]
                
                for field_id in service_order_field_ids:
                    try:
                        field = self.safe_find_by_id(field_id)
                        if field:
                            so_field = field
                            steps.append(f"Found service order field: {field_id}")
                            break
                    except Exception:
                        continue
                
                if so_field is None:
                    steps.append("Could not find any service order input field using known IDs")
                    raise Exception("Could not find service order input field")
                
                # Enter the service order number
                steps.append(f"Entering service order: {service_order}")
                so_field.text = service_order
                steps.append("Service order entered")
                
                # Press Enter
                steps.append("Sending VKey(0)")
                window = self.safe_find_by_id("wnd[0]")
                if window is None:
                    raise Exception("Could not find main window")
                window.sendVKey(0)
                steps.append("VKey sent successfully")
                
                # Wait a moment for the data to load
                steps.append("Waiting for data to load")
                time.sleep(1)
                steps.append("Wait complete")
            except Exception as e:
                return {
                    "status": "error",
                    "message": f"Failed to enter service order: {str(e)}",
                    "steps": steps,
                    "trace": traceback.format_exc()
                }
            
            # Step 3: Extract data from the first tab
            service_order_data = {
                "service_order": service_order,
                "part_number": None,
                "serial_number": None,
                "customer": None,
                "op_comments": None,
                "mod_status": None,
                "auth_documents": [],
                "notifications": [],
                "test_sheets": []
            }
            
            try:
                steps.append("Attempting to extract data from first tab")
                
                # Try to get part number using different possible field IDs
                try:
                    part_field_ids = [
                        "wnd[0]/usr/tabsTABSTRIP/tabpT\\01/ssubSUB_DATA:SAPLIQS0:7235/subGENERAL:SAPLIQS0:7212/txtLTAP-MATNR",
                        "wnd[0]/usr/tabsTABSTRIP/tabpDESC/ssubDETAIL:SAPLITO0:0115/txtITOBJ-MATXT",
                        "wnd[0]/usr/tabsTABSTRIP/tabpDESC/ssubDETAIL:SAPLITO0:0115/txtITOBJ-MATXT",
                        "wnd[0]/usr/ctxtRIWO00-MATNR",
                        # ZIWBN-specific paths from original script
                        "wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell"
                    ]
                    
                    # First try to get it from standard fields
                    for field_id in part_field_ids:
                        try:
                            field = self.safe_find_by_id(field_id)
                            if field:
                                # Check if it's a direct text field
                                try:
                                    part_number = field.text
                                    if part_number:
                                        service_order_data["part_number"] = part_number
                                        steps.append(f"Found part number in field {field_id}: {part_number}")
                                        break
                                except Exception:
                                    # It might be a grid - try to get cell value for MATNR column
                                    try:
                                        part_number = field.getCellValue(0, "MATNR")
                                        if part_number:
                                            service_order_data["part_number"] = part_number
                                            steps.append(f"Found part number in grid {field_id}: {part_number}")
                                            break
                                    except Exception:
                                        continue
                        except Exception:
                            continue
                                
                    if not service_order_data["part_number"]:
                        steps.append("Could not find part number in any expected fields")
                except Exception as part_e:
                    steps.append(f"Error getting part number: {str(part_e)}")
                
                # Try to get customer information
                try:
                    customer_field_ids = [
                        "wnd[0]/usr/tabsTABSTRIP/tabpT\\01/ssubSUB_DATA:SAPLIQS0:7235/subCUSTOMER:SAPLIQS0:7280/txtKUAGV-NAME1",
                        "wnd[0]/usr/ctxtRIWO00-KUNUM",
                        # ZIWBN-specific path from original script
                        "wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_KUNUM"
                    ]
                    
                    for field_id in customer_field_ids:
                        try:
                            field = self.safe_find_by_id(field_id)
                            if field:
                                service_order_data["customer"] = field.text
                                steps.append(f"Found customer in field {field_id}: {field.text}")
                                break
                        except Exception:
                            continue
                                
                    if not service_order_data["customer"]:
                        steps.append("Could not find customer in any expected fields")
                except Exception as cust_e:
                    steps.append(f"Error getting customer: {str(cust_e)}")
                
            except Exception as e:
                steps.append(f"Error extracting data from first tab: {str(e)}")
                # Continue to try the other tabs even if this one fails
            
            # Step 4: Go to the Equipment tab and get serial number
            try:
                steps.append("Attempting to switch to Equipment tab")
                
                # Try different tab IDs
                tab_ids = [
                    "wnd[0]/usr/tabsTABSTRIP/tabpT\\02",
                    "wnd[0]/usr/tabsTABSTRIP/tabpEQUIPMENT",
                    "wnd[0]/usr/tabsTABSTRIP/tabpDESC"
                ]
                
                tab_selected = False
                for tab_id in tab_ids:
                    try:
                        tab = self.session.findById(tab_id)
                        if tab:
                            tab.select()
                            steps.append(f"Selected tab {tab_id}")
                            tab_selected = True
                            break
                    except Exception:
                        continue
                        
                if not tab_selected:
                    steps.append("Could not select any equipment tab")
                else:
                    # Wait for tab to load
                    time.sleep(0.5)
                    
                    # Try to get serial number
                    serial_field_ids = [
                        "wnd[0]/usr/tabsTABSTRIP/tabpT\\02/ssubSUB_DATA:SAPLIQS0:7236/subOBJ:SAPLIQS0:7322/txtVIQMEL-SERGE",
                        "wnd[0]/usr/tabsTABSTRIP/tabpDESC/ssubDETAIL:SAPLITO0:0115/txtITOB-SERGE"
                    ]
                    
                    for field_id in serial_field_ids:
                        try:
                            field = self.session.findById(field_id)
                            if field:
                                service_order_data["serial_number"] = field.text
                                steps.append(f"Found serial number in field {field_id}: {field.text}")
                                break
                        except Exception:
                            continue
                            
                    if not service_order_data["serial_number"]:
                        steps.append("Could not find serial number in any expected fields")
            except Exception as e:
                steps.append(f"Error handling equipment tab: {str(e)}")
                # Continue to try the other tabs
            
            # Return the data we found, even if incomplete
            return {
                "status": "success",
                "message": "Data retrieved with some success",
                "data": service_order_data,
                "steps": steps
            }
            
        except Exception as e:
            return {
                "status": "error",
                "message": f"Failed to get service order data: {str(e)}",
                "trace": traceback.format_exc()
            }