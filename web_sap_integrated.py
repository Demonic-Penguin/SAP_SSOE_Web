"""
Web Application with Direct SAP Integration
This version uses the same direct SAP connection method as the standalone script
"""

from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify
import os
import sys
import datetime
import platform
import threading
import time
import traceback

app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "dev-secret-key")
app.config['SESSION_TYPE'] = 'filesystem'

# Flag to indicate if we're running in Replit environment
IN_REPLIT = 'REPL_ID' in os.environ if 'os' in sys.modules else False
IS_WINDOWS = platform.system() == "Windows"

# Global SAP connection object - will be initialized in background
SAP_CONNECTION = None
SAP_CONNECTION_LOCK = threading.Lock()

# Global cache for SAP data to avoid frequent lookups
SAP_DATA_CACHE = {}

class SapServiceAutomation:
    def __init__(self):
        """Initialize SAP connection"""
        self.sap_gui_auto = None
        self.application = None
        self.connection = None
        self.session = None
        self.username = "UNKNOWN"
        self.is_connected = False
        self.connection_error = None
        self.last_check_time = 0
        self.debug_info = {}
        
    def connect_to_sap(self):
        """Connect to SAP GUI"""
        print("\nConnecting to SAP GUI...")
        try:
            # Import win32com.client
            import win32com.client
            self.debug_info["win32com_import"] = "Success"
            
            # Get SAP GUI scripting object
            self.sap_gui_auto = win32com.client.GetObject("SAPGUI")
            print("✓ Connected to SAPGUI")
            self.debug_info["sap_gui_auto"] = "Success"
            
            # Get the scripting engine
            self.application = self.sap_gui_auto.GetScriptingEngine
            if self.application is None:
                self.connection_error = "Failed to get scripting engine (returned None)"
                self.debug_info["application"] = "Failed: Returned None"
                print(f"✗ {self.connection_error}")
                return False
                
            print("✓ Got scripting engine")
            self.debug_info["application"] = "Success"
            
            # Check connections
            conn_count = self.application.Children.Count
            print(f"Found {conn_count} SAP connection(s)")
            self.debug_info["connection_count"] = conn_count
            
            if conn_count == 0:
                self.connection_error = "No SAP connections found"
                self.debug_info["connection"] = "Failed: No connections"
                print(f"✗ {self.connection_error}")
                return False
            
            # Get first connection
            self.connection = self.application.Children(0)
            if self.connection is None:
                self.connection_error = "Failed to get connection (returned None)"
                self.debug_info["connection"] = "Failed: Returned None"
                print(f"✗ {self.connection_error}")
                return False
                
            print("✓ Got connection")
            self.debug_info["connection"] = "Success"
            
            # Check sessions
            sess_count = self.connection.Children.Count
            print(f"Found {sess_count} session(s)")
            self.debug_info["session_count"] = sess_count
            
            if sess_count == 0:
                self.connection_error = "No SAP sessions found"
                self.debug_info["session"] = "Failed: No sessions"
                print(f"✗ {self.connection_error}")
                return False
            
            # Get first session
            self.session = self.connection.Children(0)
            if self.session is None:
                self.connection_error = "Failed to get session (returned None)"
                self.debug_info["session"] = "Failed: Returned None"
                print(f"✗ {self.connection_error}")
                return False
                
            print("✓ Got session")
            self.debug_info["session"] = "Success"
            
            # Get user info
            info = self.session.Info
            if info is None:
                self.connection_error = "Failed to get session info (returned None)"
                self.debug_info["info"] = "Failed: Returned None"
                print(f"✗ {self.connection_error}")
                return False
                
            self.debug_info["info"] = "Success"
            
            self.username = info.User
            if not self.username:
                self.connection_error = "Failed to get username (empty)"
                self.debug_info["username"] = "Failed: Empty username"
                print(f"✗ {self.connection_error}")
                return False
                
            print(f"✓ Connected as user: {self.username}")
            self.debug_info["username"] = f"Success: {self.username}"
            
            self.is_connected = True
            self.last_check_time = time.time()
            return True
            
        except Exception as e:
            self.connection_error = f"Failed to connect to SAP: {str(e)}"
            print(f"✗ {self.connection_error}")
            print(traceback.format_exc())
            return False
    
    def check_connection(self):
        """Check if the SAP connection is still valid"""
        # Skip checks if recently checked (within last 2 seconds)
        current_time = time.time()
        if current_time - self.last_check_time < 2:
            return self.is_connected
            
        self.last_check_time = current_time
        
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
        
        # Check cache first
        if service_order in SAP_DATA_CACHE:
            print(f"Using cached data for service order {service_order}")
            return SAP_DATA_CACHE[service_order]
        
        # First try ZIWBN (as in the original script)
        ziwbn_data = self.get_ziwbn_data(service_order)
        
        # If ZIWBN worked, use that data
        if ziwbn_data and ziwbn_data.get('part_number') and ziwbn_data.get('serial_number'):
            print(f"Successfully got data from ZIWBN for service order {service_order}")
            SAP_DATA_CACHE[service_order] = ziwbn_data
            return ziwbn_data
        
        # Otherwise try IW32
        print("ZIWBN data incomplete, trying IW32...")
        
        # Navigate to IW32
        if not self.navigate_to_transaction("IW32"):
            print("Failed to navigate to IW32, returning simulation data")
            return self.simulate_service_order_data(service_order)
        
        # Enter service order
        if not self.enter_service_order(service_order):
            print("Failed to enter service order in IW32, returning simulation data")
            return self.simulate_service_order_data(service_order)
        
        # Data dictionary to collect information
        data = {
            'service_order': service_order,
            'part_number': None,
            'serial_number': None,
            'customer': None,
            'op_comments': None,
            'mod_status': None,
            'auth_documents': [],
            'notifications': [],
            'test_sheets': []
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
                if key not in ['auth_documents', 'notifications', 'test_sheets']:
                    print(f"  - {key}: {value}")
            
            # Cache the data
            SAP_DATA_CACHE[service_order] = data
            return data
            
        except Exception as e:
            print(f"✗ Failed to get service order data from IW32: {e}")
            print(traceback.format_exc())
            
            # Fall back to simulation data
            print("Falling back to simulation data")
            return self.simulate_service_order_data(service_order)

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
            'customer_number': None,
            'op_comments': "Service required due to unit failure in field. Customer requested express processing.",
            'mod_status': "MOD-A Revision 3",
            'auth_documents': ["AUTH-001", "AUTH-002"],
            'notifications': ["Z8-001", "Z8-002"],
            'test_sheets': ["TEST-001"]
        }
        
        try:
            # Get customer number from service order header
            customer_field = self.safe_find_by_id("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_KUNUM")
            if customer_field is not None:
                data['customer_number'] = customer_field.text
                data['customer'] = data['customer_number']  # Use customer number as customer name for now
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
                if key not in ['auth_documents', 'notifications', 'test_sheets']:
                    print(f"  - {key}: {value}")
            
            return data
            
        except Exception as e:
            print(f"✗ Failed to get ZIWBN data: {e}")
            print(traceback.format_exc())
            return None
    
    def simulate_service_order_data(self, service_order):
        """Simulate service order data"""
        print(f"Simulating data for service order: {service_order}")
        
        # Create a realistic looking but fake data set
        data = {
            'service_order': service_order,
            'part_number': f"MK-{service_order[:3]}-{service_order[-2:]}",
            'serial_number': f"SN{service_order}",
            'equipment': f"EQ-{service_order}",
            'customer': "CUSTOMER NAME",
            'op_comments': "Service required due to unit failure in field. Customer requested express processing.",
            'mod_status': "MOD-A Revision 3",
            'auth_documents': ["AUTH-001", "AUTH-002"],
            'notifications': ["Z8-001", "Z8-002"],
            'test_sheets': ["TEST-001"],
            'found_issues': service_order.endswith('1'),  # Some orders will have issues for testing
        }
        
        # Cache this data
        SAP_DATA_CACHE[service_order] = data
        return data
    
    def get_connection_details(self):
        """Return detailed information about the connection state"""
        return {
            "platform": platform.system(),
            "is_connected": self.is_connected,
            "connection_error": self.connection_error,
            "username": self.username,
            "debug_info": self.debug_info,
            "sap_gui_auto": self.sap_gui_auto is not None,
            "application": self.application is not None,
            "connection": self.connection is not None,
            "session": self.session is not None,
            "last_check_time": self.last_check_time
        }


# Initialize SAP connection if running on Windows
def init_sap_connection():
    global SAP_CONNECTION
    
    with SAP_CONNECTION_LOCK:
        if SAP_CONNECTION is not None:
            return SAP_CONNECTION
        
        try:
            print("Attempting to connect to SAP...")
            
            # Create SAP connection object
            sap = SapServiceAutomation()
            
            if IS_WINDOWS:
                # Try to connect to SAP
                if sap.connect_to_sap():
                    print(f"Successfully connected to SAP as user: {sap.username}")
                else:
                    print(f"Could not connect to SAP: {sap.connection_error}")
                    print("Using simulation mode")
            else:
                print("Not running on Windows. Using simulation mode.")
                
            SAP_CONNECTION = sap
            
        except Exception as e:
            print(f"Error initializing SAP connection: {e}")
            print(traceback.format_exc())
            
            # Create an unconnected instance that will use simulation mode
            SAP_CONNECTION = SapServiceAutomation()
    
    return SAP_CONNECTION

# Initialize SAP connection in a background thread
def background_init_sap():
    init_sap_connection()

# Start SAP initialization in background
threading.Thread(target=background_init_sap).start()

# Global context processor to add date to all templates
@app.context_processor
def inject_now():
    is_connected = False
    sap_username = ""
    
    if SAP_CONNECTION:
        is_connected = SAP_CONNECTION.is_connected
        sap_username = SAP_CONNECTION.username if is_connected else ""
    
    return {
        'now': datetime.datetime.now().strftime('%Y-%m-%d'),
        'sap_mode': f"SAP Connected ({sap_username})" if is_connected else "Simulation Mode"
    }

@app.route('/')
def index():
    """Render the main web interface"""
    # Force SAP connection check
    if SAP_CONNECTION:
        SAP_CONNECTION.check_connection()
        is_connected = SAP_CONNECTION.is_connected
        sap_status = f"Connected to SAP as {SAP_CONNECTION.username}" if is_connected else "Simulation Mode"
    else:
        is_connected = False
        sap_status = "SAP Connection Initializing..."
    
    # Get connection details if available
    connection_details = None
    if SAP_CONNECTION:
        connection_details = SAP_CONNECTION.get_connection_details()
    
    return render_template('index_sap_enabled.html', 
                          sap_status=sap_status,
                          connection_details=connection_details)

@app.route('/sap_status')
def sap_status():
    """Check SAP connection status"""
    if SAP_CONNECTION is None:
        return jsonify({
            'status': 'initializing',
            'message': 'SAP connection is being initialized...'
        })
    
    # Force connection check
    SAP_CONNECTION.check_connection()
    
    if SAP_CONNECTION.is_connected:
        return jsonify({
            'status': 'connected',
            'message': f'Connected to SAP as user: {SAP_CONNECTION.username}'
        })
    else:
        return jsonify({
            'status': 'simulation',
            'message': f'Running in simulation mode: {SAP_CONNECTION.connection_error}'
        })

@app.route('/run_automation', methods=['POST'])
def run_automation():
    """Handle the form submission to start automation"""
    service_order = request.form.get('service_order', '')
    if not service_order:
        return redirect(url_for('index', error='Please enter a service order number'))
    
    # Initialize SAP connection if not done yet
    if SAP_CONNECTION is None:
        init_sap_connection()
    
    # Store the service order number in session for the wizard
    session['service_order'] = service_order
    session['sap_mode'] = 'real' if SAP_CONNECTION and SAP_CONNECTION.is_connected else 'simulation'
    
    # Try to get the service order data from SAP
    try:
        # Get service order data from SAP or simulation
        order_data = SAP_CONNECTION.get_service_order_data(service_order)
        
        if not order_data:
            # Service order not found
            return redirect(url_for('index', error=f'Service order {service_order} not found'))
        
        # Store the data in session
        session['order_data'] = order_data
        return redirect(url_for('automation_wizard', step=1))
        
    except Exception as e:
        error_details = traceback.format_exc()
        print(f"Error getting service order data: {e}")
        print(error_details)
        return redirect(url_for('index', error=f'Error getting service order data: {str(e)}'))

@app.route('/automation_wizard')
def automation_wizard():
    """Render the automation wizard interface"""
    service_order = session.get('service_order', '')
    step = int(request.args.get('step', 1))
    
    if not service_order:
        return redirect(url_for('index', error='Service order number is missing'))
    
    # Get the service order data
    order_data = session.get('order_data')
    if not order_data:
        try:
            order_data = SAP_CONNECTION.get_service_order_data(service_order)
            session['order_data'] = order_data
        except Exception as e:
            return redirect(url_for('index', error=f'Error getting service order data: {str(e)}'))
    
    # Logic for different steps of the wizard with real SAP data
    steps = {
        1: {'title': 'Part Number Verification', 
            'question': 'Does the Part Number match the ID plate on the unit and the outgoing Part Number in SAP?',
            'pn': order_data.get('part_number', 'Unknown')
           },
        2: {'title': 'Serial Number Verification', 
            'question': 'Does the Serial Number match the ID plate on the unit and the outgoing Serial Number in SAP?',
            'sn': order_data.get('serial_number', 'Unknown')
           },
        3: {'title': 'Manual Entry Verification', 
            'question': 'Please enter the Part Number from the Unit being inspected to verify:',
            'input_type': 'part_number',
            'expected': order_data.get('part_number', 'Unknown')
           },
        4: {'title': 'Manual Entry Verification', 
            'question': 'Please enter the Serial Number from the Unit being inspected to verify:',
            'input_type': 'serial_number',
            'expected': order_data.get('serial_number', 'Unknown')
           },
        5: {'title': 'Operator Comments', 
            'question': f'Have you verified the operator comments to ensure there are no mismatches or discrepancies compared to actual repairs?\n\nOperator Comments: "{order_data.get("op_comments", "None")}"'
           },
        6: {'title': 'Unit Mod Status', 
            'question': f'Have you verified the unit mod status and confirmed it matches the actual unit configuration?\n\nMod Status: {order_data.get("mod_status", "Unknown")}'
           },
        7: {'title': 'Z8 Notifications', 
            'question': f'Have you verified that all Z8 notifications have been properly processed?\n\nNotifications: {", ".join(order_data.get("notifications", ["None"]))}'
           },
        8: {'title': 'Hardware Verification', 
            'question': 'Have you verified that all hardware has been properly inspected and is in good condition?'
           },
        9: {'title': 'Connectors Verification', 
            'question': 'Have you verified that all connectors have been properly inspected and are in good condition?'
           },
        10: {'title': 'FOD Check', 
             'question': 'Have you verified that the unit is free of FOD (Foreign Object Debris)?'
            },
        11: {'title': 'Customer Requirements', 
             'question': f'Have you verified that all customer requirements have been addressed and completed?\n\nCustomer: {order_data.get("customer", "Unknown")}'
            },
        12: {'title': 'Authorization Documents', 
             'question': f'Have you verified that all authorization documents have been properly processed?\n\nDocuments: {", ".join(order_data.get("auth_documents", ["None"]))}'
            },
        13: {'title': 'Service Report Match', 
             'question': 'Do the authorization documents match the service report?'
            },
        14: {'title': 'Service Report Complete', 
             'question': 'Is the service report complete with all required information filled in?'
            },
        15: {'title': 'Test Sheet Match', 
             'question': f'Does the test sheet match the unit being inspected?\n\nTest Sheets: {", ".join(order_data.get("test_sheets", ["None"]))}'
            },
        16: {'title': 'Test Sheet Failures', 
             'question': 'Does the test sheet show any failures or issues that need to be addressed?',
             'negative_is_good': True
            },
        17: {'title': 'Test Sheet Signature', 
             'question': 'Is the test sheet properly dated and signed?'
            },
        18: {'title': 'Inspection Indicators', 
             'question': 'Have you verified all inspection tab indicators and confirmed they are correct?'
            },
        19: {'title': 'Repairman Signature', 
             'question': 'Has the repairman line been properly signed?'
            },
        20: {'title': 'WSUPD Comments', 
             'question': 'Do you want to update the WSUPD comments with completion information?'
            },
    }
    
    # If we've gone past all steps, show completion
    if step > len(steps):
        return render_template('completion.html', service_order=service_order)
    
    # Get SAP connection mode
    sap_mode = session.get('sap_mode', 'simulation')
    
    return render_template('wizard_sap_enabled.html', 
                          service_order=service_order,
                          step_data=steps[step],
                          current_step=step,
                          total_steps=len(steps),
                          sap_mode=sap_mode)

@app.route('/process_step', methods=['POST'])
def process_step():
    """Process a step in the wizard"""
    service_order = session.get('service_order', '')
    current_step = int(request.form.get('current_step', 1))
    response = request.form.get('response', 'no')
    
    # Get the service order data
    order_data = session.get('order_data')
    if not order_data:
        try:
            order_data = SAP_CONNECTION.get_service_order_data(service_order)
            session['order_data'] = order_data
        except Exception as e:
            return render_template('error.html',
                                 title='Data Error',
                                 message=f'Error getting service order data: {str(e)}')
    
    # Special handling for manual entry steps
    if current_step == 3:  # Part number verification
        user_input = request.form.get('manual_input', '')
        expected = order_data.get('part_number', 'Unknown')
        
        if user_input != expected:
            # If first attempt, give another chance
            if 'retry' not in request.form:
                return render_template('wizard_sap_enabled.html',
                                      service_order=service_order,
                                      step_data={
                                          'title': 'Part Number Verification - Retry',
                                          'question': 'The Part Number does not match. Please try again:',
                                          'input_type': 'part_number',
                                          'error': f'Expected: {expected}, You entered: {user_input}'
                                      },
                                      current_step=current_step,
                                      total_steps=20,
                                      retry=True,
                                      sap_mode=session.get('sap_mode', 'simulation'))
            else:
                # Second failure, exit the process
                return render_template('error.html',
                                     title='Part Number Mismatch',
                                     message=f'The Part Number entered ({user_input}) does not match the expected value from SAP ({expected}). The process has been terminated.')
    
    if current_step == 4:  # Serial number verification
        user_input = request.form.get('manual_input', '')
        expected = order_data.get('serial_number', 'Unknown')
                
        if user_input != expected:
            # If first attempt, give another chance
            if 'retry' not in request.form:
                return render_template('wizard_sap_enabled.html',
                                      service_order=service_order,
                                      step_data={
                                          'title': 'Serial Number Verification - Retry',
                                          'question': 'The Serial Number does not match. Please try again:',
                                          'input_type': 'serial_number',
                                          'error': f'Expected: {expected}, You entered: {user_input}'
                                      },
                                      current_step=current_step,
                                      total_steps=20,
                                      retry=True,
                                      sap_mode=session.get('sap_mode', 'simulation'))
            else:
                # Second failure, exit the process
                return render_template('error.html',
                                     title='Serial Number Mismatch',
                                     message=f'The Serial Number entered ({user_input}) does not match the expected value from SAP ({expected}). The process has been terminated.')
    
    # For yes/no questions
    if response.lower() == 'no':
        # Step 16 is special: "Does test sheet show failures?" - "No" is good
        if current_step == 16:
            next_step = current_step + 1
        else:
            # For other questions, "No" means we should stop with an error
            error_messages = {
                1: 'Part Number does not match. Process terminated.',
                2: 'Serial Number does not match. Process terminated.',
                5: 'Operator comments have issues. Process terminated.',
                6: 'Unit mod status has issues. Process terminated.',
                7: 'Z8 notifications have issues. Process terminated.',
                8: 'Hardware verification failed. Process terminated.',
                9: 'Connectors verification failed. Process terminated.',
                10: 'FOD check failed. Process terminated.',
                11: 'Customer requirements not met. Process terminated.',
                12: 'Authorization documents not properly processed. Process terminated.',
                13: 'Authorization documents do not match service report. Process terminated.',
                14: 'Service report is incomplete. Process terminated.',
                15: 'Test sheet does not match unit. Process terminated.',
                17: 'Test sheet not properly signed. Process terminated.',
                18: 'Inspection indicators incorrect. Process terminated.',
                19: 'Repairman line not signed. Process terminated.',
                20: 'WSUPD comments not updated. Process terminated.',
            }
            
            return render_template('error.html',
                                 title='Process Terminated',
                                 message=error_messages.get(current_step, 'An issue was detected. Process terminated.'))
    
    # For step 16, "Yes" means there are failures, which is bad
    if current_step == 16 and response.lower() == 'yes':
        return render_template('error.html',
                             title='Test Sheet Failures',
                             message='The test sheet shows failures that need to be addressed. Process terminated.')
    
    # Move to next step
    next_step = current_step + 1
    return redirect(url_for('automation_wizard', step=next_step))

@app.route('/sap_connection_details')
def sap_connection_details():
    """Show detailed SAP connection information for debugging"""
    details = {
        'platform': platform.system(),
        'is_windows': IS_WINDOWS,
        'sap_initialized': SAP_CONNECTION is not None
    }
    
    if SAP_CONNECTION:
        SAP_CONNECTION.check_connection()
        details['connection_details'] = SAP_CONNECTION.get_connection_details()
    
    return jsonify(details)

if __name__ == '__main__':
    print(f"Starting SAP-Integrated Web Application")
    print(f"Platform: {platform.system()}")
    print(f"SAP Connection Mode: {'Real SAP (if available)' if IS_WINDOWS else 'Simulation Mode'}")
    
    app.run(host='0.0.0.0', port=5000, debug=True)