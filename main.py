"""
SAP Service Order Automation with Enhanced Error Handling
This is an improved version of main_sap_integrated.py that fixes the findById error
"""

import os
import sys
import datetime
import platform
import threading
import time
import json
import traceback
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify

app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "dev-secret-key")
app.config['SESSION_TYPE'] = 'filesystem'

# Flag to indicate if we're running in Replit environment
IN_REPLIT = 'REPL_ID' in os.environ

# Import the robust SAP connection class
import sap_debug_fix

# Global SAP connection - will be initialized if possible
SAP_CONNECTION = None
SAP_CONNECTION_LOCK = threading.Lock()
IS_WINDOWS = platform.system() == "Windows"

# Global cache for SAP data to avoid frequent lookups
SAP_DATA_CACHE = {}

# Initialize SAP connection if running on Windows
def init_sap_connection():
    global SAP_CONNECTION
    
    with SAP_CONNECTION_LOCK:
        if SAP_CONNECTION is not None:
            return SAP_CONNECTION
        
        try:
            print("Attempting to connect to SAP...")
            
            # Create and initialize the robust SAP connection handler
            sap_conn = sap_debug_fix.SapConnection()
            
            if sap_conn.connect():
                print(f"Successfully connected to SAP as user: {sap_conn.username}")
                SAP_CONNECTION = sap_conn
                
                # Test findById immediately to make sure it's working
                test_result = sap_conn.test_find_by_id()
                print(f"FindById test result: {json.dumps(test_result, indent=2)}")
                
                if test_result.get("status") != "success":
                    print(f"WARNING: FindById test did not fully succeed. This may cause issues later.")
            else:
                print(f"Could not connect to SAP: {sap_conn.connection_error}")
                print(f"Debug info: {json.dumps(sap_conn.debug_info, indent=2)}")
                
                # Create a mock connection for simulation mode
                from sap_service_order_automation import SapServiceAutomation
                SAP_CONNECTION = SapServiceAutomation(use_mock=True)
                
        except Exception as e:
            print(f"Error initializing SAP connection: {e}")
            print(traceback.format_exc())
            
            # Create a mock connection for simulation mode
            from sap_service_order_automation import SapServiceAutomation
            SAP_CONNECTION = SapServiceAutomation(use_mock=True)
    
    return SAP_CONNECTION

# Initialize SAP connection in a background thread to avoid blocking the Flask app startup
def background_init_sap():
    init_sap_connection()

# Start SAP initialization in background
threading.Thread(target=background_init_sap).start()

# Global context processor to add date to all templates
@app.context_processor
def inject_now():
    is_real_sap = False
    
    if SAP_CONNECTION:
        # Check if we have a real SAP connection (not a mock)
        if hasattr(SAP_CONNECTION, 'is_connected'):
            is_real_sap = SAP_CONNECTION.is_connected
        elif hasattr(SAP_CONNECTION, 'use_mock'):
            is_real_sap = not SAP_CONNECTION.use_mock
    
    return {
        'now': datetime.datetime.now().strftime('%Y-%m-%d'),
        'sap_mode': 'SAP Connected' if is_real_sap else 'Simulation Mode'
    }

# Function to get real data from SAP for a service order
def get_sap_service_order_data(service_order):
    """
    Pull real data from SAP for the given service order
    Returns a dictionary with service order information or None if not found
    """
    # Check if the data is in the cache
    if service_order in SAP_DATA_CACHE:
        return SAP_DATA_CACHE[service_order]
    
    # Check if we have a real SAP connection
    is_real_sap = False
    if SAP_CONNECTION:
        if hasattr(SAP_CONNECTION, 'is_connected'):
            is_real_sap = SAP_CONNECTION.is_connected
            print("YAY CONNECTED!")
        elif hasattr(SAP_CONNECTION, 'use_mock'):
            is_real_sap = not SAP_CONNECTION.use_mock
            print("BOO!! NOT CONNECTED...")
    
    if not is_real_sap:
        # Simulation mode - use mock data
        return simulate_service_order_data(service_order)
    
    try:
        print(f"Pulling data from SAP for service order: {service_order}")
        
        # Use the improved SAP connection to get service order data
        if hasattr(SAP_CONNECTION, 'get_service_order_data'):
            # Use the new method if available
            result = SAP_CONNECTION.get_service_order_data(service_order)
            print("SAP result is:", result)
            if result.get("status") == "success":
                data = result.get("data", {})
                # Save to cache
                SAP_DATA_CACHE[service_order] = data
                return data
            else:
                print(f"Error getting service order data: {result.get('message')}")
                return simulate_service_order_data(service_order)
        else:
            # Fallback to simulation
            return simulate_service_order_data(service_order)
        
    except Exception as e:
        print(f"Error getting data from SAP: {e}")
        print(traceback.format_exc())
        # Fall back to simulation mode if there's an error
        return simulate_service_order_data(service_order)

def simulate_service_order_data(service_order):
    """Simulate service order data for development/testing"""
    # In a real implementation, this would be replaced with actual SAP data
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

# Configure routes
@app.route('/')
def index():
    """Render the main web interface"""
    # Check if we have a real SAP connection
    is_real_sap = False
    if SAP_CONNECTION:
        if hasattr(SAP_CONNECTION, 'is_connected'):
            is_real_sap = SAP_CONNECTION.is_connected
        elif hasattr(SAP_CONNECTION, 'use_mock'):
            is_real_sap = not SAP_CONNECTION.use_mock
    
    sap_status = "Connected to SAP" if is_real_sap else "Simulation Mode"
    
    # Get connection details if available
    connection_details = None
    if hasattr(SAP_CONNECTION, 'get_connection_details'):
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
    
    # Check if we have a real SAP connection
    is_real_sap = False
    if SAP_CONNECTION:
        if hasattr(SAP_CONNECTION, 'is_connected'):
            is_real_sap = SAP_CONNECTION.is_connected
            username = getattr(SAP_CONNECTION, 'username', 'Unknown')
        elif hasattr(SAP_CONNECTION, 'use_mock'):
            is_real_sap = not SAP_CONNECTION.use_mock
            username = getattr(SAP_CONNECTION, 'username', 'Unknown')
    
    if is_real_sap:
        return jsonify({
            'status': 'connected',
            'message': f'Connected to SAP as user: {username}'
        })
    else:
        return jsonify({
            'status': 'simulation',
            'message': 'Running in simulation mode'
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
    
    # Check if we have a real SAP connection
    is_real_sap = False
    if SAP_CONNECTION:
        if hasattr(SAP_CONNECTION, 'is_connected'):
            is_real_sap = SAP_CONNECTION.is_connected
        elif hasattr(SAP_CONNECTION, 'use_mock'):
            is_real_sap = not SAP_CONNECTION.use_mock
    
    # Store the service order number in session for the wizard
    session['service_order'] = service_order
    session['sap_mode'] = 'real' if is_real_sap else 'simulation'
    
    # Try to get the service order data from SAP
    try:
        # Get service order data from SAP
        order_data = get_sap_service_order_data(service_order)
        
        if not order_data and is_real_sap:
            # Service order not found in real SAP
            return redirect(url_for('index', error=f'Service order {service_order} not found in SAP'))
        
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
            order_data = get_sap_service_order_data(service_order)
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
            order_data = get_sap_service_order_data(service_order)
            session['order_data'] = order_data
        except Exception as e:
            return render_template('error.html',
                                  title='Data Error',
                                  message=f'Error getting service order data: {str(e)}')
    
    # Special handling for manual entry steps
    if current_step == 3:  # Part number verification
        user_input = request.form.get('manual_input', '')
        expected = order_data.get('part_number', 'Unknown')
        
        # Check if we have a real SAP connection
        is_real_sap = False
        if SAP_CONNECTION:
            if hasattr(SAP_CONNECTION, 'is_connected'):
                is_real_sap = SAP_CONNECTION.is_connected
            elif hasattr(SAP_CONNECTION, 'use_mock'):
                is_real_sap = not SAP_CONNECTION.use_mock
        
        if is_real_sap:
            try:
                # This would be actual SAP verification logic in a real implementation
                print(f"Verifying part number in SAP: {user_input}")
                # In real implementation, would call appropriate SAP verification
            except Exception as e:
                flash(f"Error communicating with SAP: {str(e)}", "error")
        
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
        
        # Check if we have a real SAP connection
        is_real_sap = False
        if SAP_CONNECTION:
            if hasattr(SAP_CONNECTION, 'is_connected'):
                is_real_sap = SAP_CONNECTION.is_connected
            elif hasattr(SAP_CONNECTION, 'use_mock'):
                is_real_sap = not SAP_CONNECTION.use_mock
        
        if is_real_sap:
            try:
                # This would be actual SAP verification logic in a real implementation
                print(f"Verifying serial number in SAP: {user_input}")
                # In real implementation, would call appropriate SAP verification
            except Exception as e:
                flash(f"Error communicating with SAP: {str(e)}", "error")
                
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
            
            # Check if we should update real SAP
            is_real_sap = False
            if SAP_CONNECTION:
                if hasattr(SAP_CONNECTION, 'is_connected'):
                    is_real_sap = SAP_CONNECTION.is_connected
                elif hasattr(SAP_CONNECTION, 'use_mock'):
                    is_real_sap = not SAP_CONNECTION.use_mock
            
            # If connected to real SAP, update SAP about the failure
            if is_real_sap and current_step > 5:
                try:
                    step_name = {
                        5: 'Operator Comments',
                        6: 'Unit Mod Status',
                        7: 'Z8 Notifications',
                        8: 'Hardware Verification',
                        9: 'Connectors Verification',
                        10: 'FOD Check',
                        11: 'Customer Requirements',
                        12: 'Authorization Documents',
                        13: 'Service Report Match',
                        14: 'Service Report Complete',
                        15: 'Test Sheet Match',
                        17: 'Test Sheet Signature',
                        18: 'Inspection Indicators',
                        19: 'Repairman Signature',
                        20: 'WSUPD Comments',
                    }.get(current_step, f'Step {current_step}')
                    
                    # This would be real SAP update logic in a real implementation
                    print(f"Updating SAP for service order {service_order}: Failed at {step_name}")
                    # Would do actual SAP updates here
                except Exception as e:
                    print(f"Error updating SAP: {e}")
            
            return render_template('error.html',
                                 title='Process Terminated',
                                 message=error_messages.get(current_step, 'An issue was detected. Process terminated.'))
    
    # For step 16, "Yes" means there are failures, which is bad
    if current_step == 16 and response.lower() == 'yes':
        # Check if we should update real SAP
        is_real_sap = False
        if SAP_CONNECTION:
            if hasattr(SAP_CONNECTION, 'is_connected'):
                is_real_sap = SAP_CONNECTION.is_connected
            elif hasattr(SAP_CONNECTION, 'use_mock'):
                is_real_sap = not SAP_CONNECTION.use_mock
                
        # If connected to real SAP, update SAP about the test sheet failures
        if is_real_sap:
            try:
                # This would be real SAP update logic in a real implementation
                print(f"Updating SAP for service order {service_order}: Test sheet has failures")
                # Would do actual SAP updates here
            except Exception as e:
                print(f"Error updating SAP: {e}")
                
        return render_template('error.html',
                             title='Test Sheet Failures',
                             message='The test sheet shows failures that need to be addressed. Process terminated.')
    
    # If on the last step and user said yes, update SAP with completion info
    if current_step == 20 and response.lower() == 'yes':
        # Check if we should update real SAP
        is_real_sap = False
        if SAP_CONNECTION:
            if hasattr(SAP_CONNECTION, 'is_connected'):
                is_real_sap = SAP_CONNECTION.is_connected
            elif hasattr(SAP_CONNECTION, 'use_mock'):
                is_real_sap = not SAP_CONNECTION.use_mock
                
        if is_real_sap:
            try:
                # Update WSUPD in the real SAP system
                print(f"Updating WSUPD comments in SAP for service order {service_order}")
                # This would be real SAP code to update WSUPD
                # SAP_CONNECTION.session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW32"
                # SAP_CONNECTION.session.findById("wnd[0]").sendVKey(0)
                # And so on...
            except Exception as e:
                print(f"Error updating SAP WSUPD: {e}")
    
    # Move to next step
    next_step = current_step + 1
    return redirect(url_for('automation_wizard', step=next_step))

@app.route('/sap_connection_details')
def sap_connection_details():
    """Show detailed SAP connection information for debugging"""
    details = {
        'platform': platform.system(),
        'is_windows': IS_WINDOWS,
        'sap_initialized': SAP_CONNECTION is not None,
        'connection_details': {}
    }
    
    if SAP_CONNECTION:
        if hasattr(SAP_CONNECTION, 'get_connection_details'):
            details['connection_details'] = SAP_CONNECTION.get_connection_details()
        elif hasattr(SAP_CONNECTION, 'use_mock'):
            details['connection_details'] = {
                'is_mock': SAP_CONNECTION.use_mock,
                'username': getattr(SAP_CONNECTION, 'username', 'Unknown')
            }
    
    return jsonify(details)

if __name__ == '__main__':
    print(f"Starting SAP Service Order Automation Web Interface")
    print(f"Platform: {platform.system()}")
    
    # Check if we're on Windows
    if IS_WINDOWS:
        print("Running on Windows - Will attempt to connect to SAP")
    else:
        print("Not running on Windows - Will use simulation mode")
    
    # Show very detailed connection info at startup
    if IS_WINDOWS:
        try:
            import win32com.client
            print("win32com.client module is available")
            
            try:
                print("\nTesting SAP connection with win32com...")
                sap_gui = win32com.client.GetObject("SAPGUI")
                print("  ✓ Successfully connected to SAPGUI")
                
                script_engine = sap_gui.GetScriptingEngine
                print("  ✓ Successfully got scripting engine")
                
                print(f"  Found {script_engine.Children.Count} connection(s)")
                
                if script_engine.Children.Count > 0:
                    connection = script_engine.Children(0)
                    print("  ✓ Successfully got first connection")
                    
                    print(f"  Found {connection.Children.Count} session(s)")
                    
                    if connection.Children.Count > 0:
                        session = connection.Children(0)
                        print("  ✓ Successfully got first session")
                        
                        try:
                            info = session.Info
                            user = info.User
                            print(f"  ✓ Connected as user: {user}")
                        except Exception as e:
                            print(f"  ✗ Failed to get user info: {e}")
                    else:
                        print("  ✗ No SAP sessions found")
                else:
                    print("  ✗ No SAP connections found")
                    
            except Exception as e:
                print(f"Error testing SAP connection: {e}")
            
        except ImportError:
            print("win32com.client module is not available")
            print("Install pywin32 for SAP integration: pip install pywin32")
    
    # Run the Flask app
    app.run(host='0.0.0.0', port=5000, debug=True)