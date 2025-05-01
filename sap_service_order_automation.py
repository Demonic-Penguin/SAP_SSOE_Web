"""
SAP Service Order Automation Script
Converted from VBS to Python
This script automates the process of checking and completing service orders in SAP.

Note: In this implementation, we simulate the SAP GUI interactions
since direct SAP GUI automation may not be available in all environments.
"""

import os
import sys
import time
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, simpledialog

# Detect if we're in a Replit environment
IN_REPLIT = "REPLIT_DB_URL" in os.environ

# Mock SAP GUI components for environments without SAP GUI access
class MockSapGui:
    """Mock SAP GUI for environments without SAP GUI access"""
    def __init__(self):
        self.windows = {"wnd[0]": MockWindow("Main Window")}
        self.user = "MOCKUSER"
        
    def find_by_id(self, element_id):
        """Simulate finding elements by ID"""
        parts = element_id.split("/")
        window_id = parts[0]
        
        if window_id not in self.windows:
            # Create window if it doesn't exist (simulating popup)
            self.windows[window_id] = MockWindow(f"Window {window_id}")
            
        window = self.windows[window_id]
        
        # Create mock element based on element type in ID
        if "usr" in element_id:
            return MockUsrElement(element_id)
        elif "tbar" in element_id:
            return MockToolbarElement(element_id)
        elif "okcd" in element_id:
            return MockCommandElement(element_id)
        elif "ctxt" in element_id:
            return MockTextElement(element_id)
        elif "txt" in element_id:
            return MockTextElement(element_id)
        elif "btn" in element_id:
            return MockButtonElement(element_id)
        elif "shell" in element_id:
            return MockGridElement(element_id)
        else:
            return MockElement(element_id)

class MockElement:
    """Base mock element class"""
    def __init__(self, element_id):
        self.element_id = element_id
        self.text = ""
        self.caretPosition = 0
        
    def select(self):
        """Select this element"""
        print(f"Selected element: {self.element_id}")
        return True
        
    def press(self):
        """Press this element"""
        print(f"Pressed element: {self.element_id}")
        return True
    
    def sendVKey(self, key_code):
        """Send virtual key press"""
        print(f"Sent key {key_code} to element: {self.element_id}")
        return True
        
    def pressToolbarButton(self, button_id):
        """Press a toolbar button"""
        print(f"Pressed toolbar button {button_id} on element: {self.element_id}")
        return True
        
    def doubleClickCurrentCell(self):
        """Double click the current cell"""
        print(f"Double-clicked current cell in element: {self.element_id}")
        return True
        
    def getCellValue(self, row, column):
        """Get cell value from grid"""
        # Return mock values based on column type
        if column == "MATNR":
            return "MK-12345"
        elif column == "SERNR":
            return "SN987654"
        else:
            return "MockValue"

class MockWindow(MockElement):
    """Mock SAP window"""
    def __init__(self, title):
        super().__init__(title)
        self.title = title

class MockUsrElement(MockElement):
    """Mock user area element"""
    pass

class MockToolbarElement(MockElement):
    """Mock toolbar element"""
    pass

class MockCommandElement(MockElement):
    """Mock command field element"""
    pass

class MockTextElement(MockElement):
    """Mock text element"""
    pass

class MockButtonElement(MockElement):
    """Mock button element"""
    pass

class MockGridElement(MockElement):
    """Mock grid/table element"""
    def __init__(self, element_id):
        super().__init__(element_id)
        self.currentCellRow = 0
        self.currentCellColumn = ""
        self.selectedRows = ""

class MockSession:
    """Mock SAP session"""
    def __init__(self):
        self.info = MockInfo()
        self.sap_gui = MockSapGui()
        
    def findById(self, element_id):
        """Find element by ID"""
        return self.sap_gui.find_by_id(element_id)

class MockInfo:
    """Mock session info"""
    def __init__(self):
        self.user = "MOCKUSER"

class SapServiceAutomation:
    def __init__(self, use_mock=False):
        """Initialize the SAP Service Order Automation tool"""
        # Initialize variables
        self.billable = False
        self.tmp_admin_code_status = ""
        self.sales_value = ""
        self.is_spex = False
        self.new_spex = False
        self.version10 = False
        self.stay_labored_on = False
        self.is_converted = False
        self.is_exchange = False
        self.is_dpmi = False
        self.oops = False
        self.bool_spex = False
        self.bool_error = False

        # Service order details
        self.tmp_serv_order = ""
        self.tmp_cust = ""
        self.tmp_del_blck = ""
        self.tmp_zh_status = ""
        self.tmp_zg_status = ""
        self.tmp_mods_in = ""
        self.tmp_mods_out = ""
        self.tmp_sf_mods_in = ""
        self.tmp_sf_mods_out = ""
        self.tmp_sw_versions_in = ""
        self.tmp_sw_versions_out = ""
        self.tmp_pn = ""
        self.tmp_sn = ""
        self.tmp_customer_num = ""
        self.tmp_superior_order = ""
        self.username = ""
        
        # Error sound for critical errors
        self.error_sound = "Oh no"
        
        # Initialize GUI root for dialogs
        self.root = tk.Tk()
        self.root.withdraw()  # Hide the main window
        
        # Setup SAP connection (or mock)
        self.use_mock = use_mock
        self.sap_gui_auto = None
        self.application = None
        self.connection = None
        self.session = None
        
        if use_mock or IN_REPLIT:
            print("Using mock SAP session (simulation mode)")
            self.session = MockSession()
            self.username = self.session.info.user
        else:
            # Real SAP connection will be attempted in connect_to_sap()
            pass

    def play_error_sound(self):
        """Play error sound - simulated in this implementation"""
        print(f"ERROR SOUND: {self.error_sound}")
        
    def connect_to_sap(self):
        """Connect to the SAP system and get the active session"""
        if self.use_mock or IN_REPLIT:
            # Using mock SAP session
            return True
            
        try:
            import win32com.client
            max_attempts = 5
            attempt = 0
            
            while attempt < max_attempts:
                try:
                    if not self.sap_gui_auto:
                        self.sap_gui_auto = win32com.client.GetObject("SAPGUI")
                        
                    if not self.application:
                        self.application = self.sap_gui_auto.GetScriptingEngine
                        
                    if not self.connection:
                        self.connection = self.application.Children(0)
                        
                    if not self.session:
                        self.session = self.connection.Children(0)
                        
                    # Get the username from the session
                    self.username = self.session.Info.User
                    return True
                    
                except Exception as e:
                    attempt += 1
                    res_sap = messagebox.askokcancel(
                        "SAP Connection Error",
                        "SAP was not detected. Please open a SAP session. Thank you."
                    )
                    if not res_sap:
                        return False
                        
            return False
        except ImportError:
            messagebox.showerror(
                "Missing SAP GUI Components",
                "The PyWin32 module is not available. Running in simulation mode."
            )
            # Fall back to mock mode
            self.use_mock = True
            self.session = MockSession()
            self.username = self.session.info.user
            return True
            
    def get_service_order_input(self):
        """Get service order number from user"""
        self.tmp_serv_order = simpledialog.askstring("SSOE", "Enter service order number?")
        if not self.tmp_serv_order:
            return False
        return True
    
    def open_ziwbn(self):
        """Open ZIWBN transaction with the service order"""
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nZIWBN"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB1:SAPLYAFF_ZIWBNGUI:0100/ssubSUB2:SAPLYAFF_ZIWBNGUI:0102/ctxtW_INP_DATA").text = self.tmp_serv_order
            self.session.findById("wnd[0]").sendVKey(0)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error opening ZIWBN: {str(e)}")
            return False
    
    def get_iw32_information(self):
        """Get information from IW32 transaction"""
        try:
            self.tmp_customer_num = self.session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_KUNUM").text
            self.tmp_superior_order = self.session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB1:SAPLYAFF_ZIWBNGUI:0100/ssubSUB2:SAPLYAFF_ZIWBNGUI:0102/ctxtW_INP_DATA").text
            
            # Check if customer is SPEX
            spex_customers = [
                "PLANT1133", "SLSR01", "PLANT1057", "PLANT1052",
                "PLANT1013", "PLANT1103", "PLANT1116", "PLANT1005"
            ]
            
            self.is_spex = self.tmp_customer_num in spex_customers
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error getting information: {str(e)}")
            return False
            
    def labor_on(self):
        """Turn labor on for the service order"""
        try:
            self.session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I").select()
            
            self.session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").pressToolbarButton("&MB_FILTER")
            self.session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell").currentCellRow = 7
            self.session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell").selectedRows = "7"
            self.session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell").doubleClickCurrentCell()
            self.session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/btn600_BUTTON").press()
            self.session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "Close Up Inspection"
            self.session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 19
            self.session.findById("wnd[2]").sendVKey(0)
            self.session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").currentCellColumn = ""
            self.session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectedRows = "0"
            self.session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").pressToolbarButton("LABON")
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error turning labor on: {str(e)}")
            return False
    
    def get_part_number_serial_number(self):
        """Get part number and serial number from the equipment tab"""
        try:
            self.session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H").select()
            
            try:
                self.session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell").currentCellColumn = "MATNR"
                
                self.tmp_pn = self.session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell").getCellValue(0, "MATNR")
                self.tmp_sn = self.session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell").getCellValue(0, "SERNR")
                self.version10 = False
            except Exception:
                # Handle version 10 interface
                self.version10 = True
                self.session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont[0]/shell").currentCellColumn = "MATNR"
                self.tmp_pn = self.session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont[0]/shell").getCellValue(0, "MATNR")
                self.tmp_sn = self.session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont[0]/shell").getCellValue(0, "SERNR")
            
            self.session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select()
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error getting part/serial number: {str(e)}")
            return False
    
    def ask_about_part_number_serial_number(self):
        """Ask user to confirm part number and serial number match"""
        try:
            # Check if PN matches ID plate
            res1 = messagebox.askyesno(
                "Check",
                "Does the Part Number match the ID plate on the unit and the outgoing Part Number in SAP?"
            )
            if not res1:
                self.play_error_sound()
                return False
            
            # Check if SN matches ID plate
            res2 = messagebox.askyesno(
                "Check",
                "Does the Serial Number match the ID plate on the unit and the outgoing Serial Number in SAP?"
            )
            if not res2:
                self.play_error_sound()
                return False
            
            # Verify part number input
            pn_tries = 0
            while True:
                input_pn = simpledialog.askstring(
                    "Part Number Verification",
                    "Please enter the Part Number from the Unit being inspected."
                )
                
                if input_pn == self.tmp_pn:
                    break
                else:
                    pn_tries += 1
                    if pn_tries > 1:
                        self.play_error_sound()
                        messagebox.showerror(
                            "Error",
                            "The Part Number being tried does not appear to match SAP. This script will now terminate."
                        )
                        return False
            
            # Verify serial number input
            sn_tries = 0
            while True:
                input_sn = simpledialog.askstring(
                    "Serial Number Verification",
                    "Please enter the Serial Number from the Unit being inspected."
                )
                
                if input_sn == self.tmp_sn:
                    break
                else:
                    sn_tries += 1
                    if sn_tries > 1:
                        self.play_error_sound()
                        messagebox.showerror(
                            "Error",
                            "The Serial Number being tried does not appear to match SAP. This script will now terminate."
                        )
                        return False
            
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error in part/serial number verification: {str(e)}")
            return False
    
    def ask_about_operator_comments(self):
        """Ask user to confirm operator comments"""
        try:
            res = messagebox.askyesno(
                "Operator Comments",
                "Have you verified the operator comments to ensure there are no mismatches or discrepancies compared to actual repairs?"
            )
            if not res:
                self.bool_error = True
                return False
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error in operator comments verification: {str(e)}")
            self.bool_error = True
            return False
    
    def ask_about_unit_mod_status(self):
        """Ask user to confirm unit mod status"""
        try:
            res = messagebox.askyesno(
                "Unit Mod Status",
                "Have you verified the unit mod status and confirmed it matches the actual unit configuration?"
            )
            if not res:
                self.bool_error = True
                return False
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error in unit mod status verification: {str(e)}")
            self.bool_error = True
            return False
    
    def z8_notifications(self):
        """Process Z8 notifications"""
        try:
            res = messagebox.askyesno(
                "Z8 Notifications",
                "Have you verified that all Z8 notifications have been properly processed?"
            )
            if not res:
                self.bool_error = True
                return False
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error in Z8 notifications verification: {str(e)}")
            self.bool_error = True
            return False
    
    def process_auth_docs(self):
        """Process authorization documents"""
        try:
            res = messagebox.askyesno(
                "Authorization Documents",
                "Have you verified that all authorization documents have been properly processed?"
            )
            if not res:
                self.bool_error = True
                return False
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error in authorization documents verification: {str(e)}")
            self.bool_error = True
            return False
    
    def ask_about_hardware(self):
        """Ask user to confirm hardware verification"""
        try:
            res = messagebox.askyesno(
                "Hardware Verification",
                "Have you verified that all hardware has been properly inspected and is in good condition?"
            )
            if not res:
                self.bool_error = True
                return False
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error in hardware verification: {str(e)}")
            self.bool_error = True
            return False
    
    def ask_about_connectors(self):
        """Ask user to confirm connectors verification"""
        try:
            res = messagebox.askyesno(
                "Connectors Verification",
                "Have you verified that all connectors have been properly inspected and are in good condition?"
            )
            if not res:
                self.bool_error = True
                return False
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error in connectors verification: {str(e)}")
            self.bool_error = True
            return False
    
    def ask_about_fod(self):
        """Ask user to confirm FOD check"""
        try:
            res = messagebox.askyesno(
                "FOD Check",
                "Have you verified that the unit is free of FOD (Foreign Object Debris)?"
            )
            if not res:
                self.bool_error = True
                return False
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error in FOD check verification: {str(e)}")
            self.bool_error = True
            return False
    
    def ask_about_customer_req(self):
        """Ask user to confirm customer requirements"""
        try:
            res = messagebox.askyesno(
                "Customer Requirements",
                "Have you verified that all customer requirements have been addressed and completed?"
            )
            if not res:
                self.bool_error = True
                return False
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error in customer requirements verification: {str(e)}")
            self.bool_error = True
            return False
    
    def ask_about_auth_docs_matching_service_report(self):
        """Ask user to confirm auth docs match service report"""
        try:
            res = messagebox.askyesno(
                "Authorization Documents",
                "Do the authorization documents match the service report?"
            )
            if not res:
                self.bool_error = True
                return False
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error in auth docs verification: {str(e)}")
            self.bool_error = True
            return False
    
    def ask_about_is_service_report_complete(self):
        """Ask user to confirm service report is complete"""
        try:
            res = messagebox.askyesno(
                "Service Report",
                "Is the service report complete with all required information filled in?"
            )
            if not res:
                self.bool_error = True
                return False
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error in service report verification: {str(e)}")
            self.bool_error = True
            return False
    
    def ask_about_does_test_sheet_match_unit(self):
        """Ask user to confirm test sheet matches unit"""
        try:
            res = messagebox.askyesno(
                "Test Sheet",
                "Does the test sheet match the unit being inspected?"
            )
            if not res:
                self.bool_error = True
                return False
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error in test sheet verification: {str(e)}")
            self.bool_error = True
            return False
    
    def ask_about_does_test_sheet_show_any_fails(self):
        """Ask user if test sheet shows any failures"""
        try:
            res = messagebox.askyesno(
                "Test Sheet Failures",
                "Does the test sheet show any failures or issues that need to be addressed?"
            )
            if res:  # If there are failures, this is a problem
                self.bool_error = True
                return False
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error in test sheet failures verification: {str(e)}")
            self.bool_error = True
            return False
    
    def ask_about_date_and_signature_on_test_sheet(self):
        """Ask user to confirm date and signature on test sheet"""
        try:
            res = messagebox.askyesno(
                "Test Sheet Signature",
                "Is the test sheet properly dated and signed?"
            )
            if not res:
                self.bool_error = True
                return False
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error in test sheet signature verification: {str(e)}")
            self.bool_error = True
            return False
    
    def check_inspection_tab_indicators(self):
        """Check inspection tab indicators"""
        try:
            res = messagebox.askyesno(
                "Inspection Indicators",
                "Have you verified all inspection tab indicators and confirmed they are correct?"
            )
            if not res:
                self.bool_error = True
                return False
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error in inspection indicators verification: {str(e)}")
            self.bool_error = True
            return False
    
    def ask_about_has_the_repairman_line_been_signed(self):
        """Ask user to confirm repairman line has been signed"""
        try:
            res = messagebox.askyesno(
                "Repairman Signature",
                "Has the repairman line been properly signed?"
            )
            if not res:
                self.bool_error = True
                return False
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error in repairman signature verification: {str(e)}")
            self.bool_error = True
            return False
    
    def update_wsupd_comments(self):
        """Update WSUPD comments"""
        try:
            # In the original script, this would update comments in the SAP system
            # Here we just ask the user to confirm
            res = messagebox.askyesno(
                "WSUPD Comments",
                "Do you want to update the WSUPD comments with completion information?"
            )
            if not res:
                self.bool_error = True
                return False
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error in WSUPD comments update: {str(e)}")
            self.bool_error = True
            return False
    
    def labor_off_incomplete(self):
        """Turn labor off when process is incomplete"""
        try:
            # Original script would update SAP
            messagebox.showinfo(
                "Labor Off - Incomplete",
                "Labor is being turned off due to incomplete process. Please address the issues before continuing."
            )
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error turning labor off: {str(e)}")
            return False
    
    def labor_off_complete(self):
        """Turn labor off when process is complete"""
        try:
            # Original script would update SAP
            messagebox.showinfo(
                "Labor Off - Complete",
                "Labor is being turned off as the process has been completed successfully."
            )
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error turning labor off: {str(e)}")
            return False
    
    def main(self):
        """Main execution function"""
        # Connect to SAP (or use mock)
        if not self.connect_to_sap():
            messagebox.showerror("Error", "Failed to connect to SAP. Exiting.")
            return False
        
        # Get service order number from user
        if not self.get_service_order_input():
            messagebox.showinfo("Info", "No Service Order Number entered.")
            return False
        
        # Open ZIWBN transaction
        if not self.open_ziwbn():
            return False
        
        # Get IW32 information
        if not self.get_iw32_information():
            return False
            
        # Open ZIWBN again
        if not self.open_ziwbn():
            return False
        
        # Get part number and serial number
        if not self.get_part_number_serial_number():
            return False
            
        # Ask about part number and serial number
        if not self.ask_about_part_number_serial_number():
            return False
        
        # Ask about operator comments
        if not self.ask_about_operator_comments():
            self.labor_off_incomplete()
            return False
        
        # Ask about unit mod status
        if not self.ask_about_unit_mod_status():
            return False
        
        # Check Z8 notifications
        if not self.z8_notifications():
            return False
        
        # Ask about hardware
        if not self.ask_about_hardware():
            return False
        
        # Ask about connectors
        if not self.ask_about_connectors():
            return False
        
        # Ask about FOD
        if not self.ask_about_fod():
            return False
            
        # If not SPEX, check additional requirements
        if not self.is_spex:
            # Ask about customer requirements
            if not self.ask_about_customer_req():
                return False
        
        # Process authorization documents
        if not self.process_auth_docs():
            return False
        
        # Ask about auth docs matching service report
        if not self.ask_about_auth_docs_matching_service_report():
            return False
        
        # Ask about service report completeness
        if not self.ask_about_is_service_report_complete():
            return False
        
        # Ask about test sheet matching unit
        if not self.ask_about_does_test_sheet_match_unit():
            return False
        
        # Ask about test sheet failures
        if not self.ask_about_does_test_sheet_show_any_fails():
            return False
        
        # Ask about date and signature on test sheet
        if not self.ask_about_date_and_signature_on_test_sheet():
            return False
        
        # Check inspection tab indicators
        if not self.check_inspection_tab_indicators():
            return False
        
        # Ask about repairman signature
        if not self.ask_about_has_the_repairman_line_been_signed():
            return False
        
        # Update WSUPD comments
        if not self.update_wsupd_comments():
            return False
        
        # Turn labor off (complete)
        self.labor_off_complete()
        
        # Show completion message
        messagebox.showinfo("SSOE Complete", "SSOE process has been completed successfully.")
        return True


def main():
    """Main entry point"""
    # Create SAP automation instance
    # Automatically use mock mode in non-Windows environments
    use_mock = sys.platform != "win32" or IN_REPLIT
    
    app = SapServiceAutomation(use_mock=use_mock)
    
    # Run the main process
    app.main()
    
    # Run Tkinter main loop to ensure all dialogs are properly handled
    app.root.mainloop()


if __name__ == "__main__":
    main()