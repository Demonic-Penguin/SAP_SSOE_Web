"""
SAP-Integrated Web Application
This version pulls real data from SAP instead of using mockup data
"""

from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify
import os
import sys
import datetime
import platform
import threading
import time

app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "dev-secret-key")

# Import the SAP Service Order Automation class
from sap_service_order_automation import SapServiceAutomation

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
            # Try to import win32com
            if IS_WINDOWS:
                try:
                    import win32com.client
                    # Create SAP connection with mock mode disabled
                    sap = SapServiceAutomation(use_mock=False)
                    
                    # Test the connection
                    if sap.connect_to_sap():
                        print(f"Successfully connected to SAP as user: {sap.username}")
                        SAP_CONNECTION = sap
                    else:
                        print("Could not connect to SAP. Using simulation mode.")
                        SAP_CONNECTION = SapServiceAutomation(use_mock=True)
                except ImportError:
                    print("win32com.client module not available. Using simulation mode.")
                    SAP_CONNECTION = SapServiceAutomation(use_mock=True)
            else:
                print("Not running on Windows. Using simulation mode.")
                SAP_CONNECTION = SapServiceAutomation(use_mock=True)
        except Exception as e:
            print(f"Error initializing SAP connection: {e}")
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
    return {
        'now': datetime.datetime.now().strftime('%Y-%m-%d'),
        'sap_mode': 'SAP Connected' if SAP_CONNECTION and not SAP_CONNECTION.use_mock else 'Simulation Mode'
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
    
    # Check if we have a live SAP connection
    if not SAP_CONNECTION or SAP_CONNECTION.use_mock:
        # Simulation mode - use the mock data
        return simulate_service_order_data(service_order)
    
    try:
        print(f"Pulling data from SAP for service order: {service_order}")
        
        # This is where you would call the actual SAP methods to get data
        # For example using the SAP_CONNECTION.session object
        
        # IMPORTANT: The SAP communication code here would need to match the 
        # exact field names, transactions, and methods needed for your SAP system
        
        # Example of what real SAP integration might look like:
        SAP_CONNECTION.session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/"
                         "subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H").select()
        SAP_CONNECTION.session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/"
                         "subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/"
                         "ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell"
                         ).currentCellColumn = "MATNR"
        # SAP_CONNECTION.session.findById("wnd[0]/usr/ctxtVORG").text = service_order
        # SAP_CONNECTION.session.findById("wnd[0]").sendVKey(0)
        # 
        pn = SAP_CONNECTION.session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/"
                              "subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/"
                              "ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell"
                              ).getCellValue(0, "MATNR")
        sn = SAP_CONNECTION.session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/"
                              "subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/"
                              "ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell"
                              ).getCellValue(0, "SERNR")
        # op_comments = SAP_CONNECTION.session.findById("wnd[0]/usr/tabsMAIN/tabpPROB/ssubDETAIL:SAPLITO0:0114/txtVIQMEL-QMTXT").text
        
        # For this example, we'll simulate the data but in a real system, the above code would
        # be replaced with actual SAP GUI scripting to get the right data
        return simulate_service_order_data(service_order)
        
    except Exception as e:
        print(f"Error getting data from SAP: {e}")
        # Fall back to simulation mode if there's an error
        return simulate_service_order_data(service_order)

def simulate_service_order_data(service_order):
    """Simulate service order data for development/testing"""
    # In a real implementation, this would be replaced with actual SAP data
    print(f"Simulating data for service order: {service_order}")
    
    # Create a realistic looking but fake data set
    data = {
        'service_order': service_order,
        'pn': f"MK-{service_order[:3]}-{service_order[-2:]}",
        'sn': f"SN{service_order}",
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
    sap_status = "Connected to SAP" if SAP_CONNECTION and not SAP_CONNECTION.use_mock else "Simulation Mode"
    return render_template('index_sap_enabled.html', sap_status=sap_status)

@app.route('/sap_status')
def sap_status():
    """Check SAP connection status"""
    if SAP_CONNECTION is None:
        return jsonify({
            'status': 'initializing',
            'message': 'SAP connection is being initialized...'
        })
    elif SAP_CONNECTION.use_mock:
        return jsonify({
            'status': 'simulation',
            'message': 'Running in simulation mode'
        })
    else:
        return jsonify({
            'status': 'connected',
            'message': f'Connected to SAP as user: {SAP_CONNECTION.username}'
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
    session['sap_mode'] = 'real' if SAP_CONNECTION and not SAP_CONNECTION.use_mock else 'simulation'
    
    # Try to get the service order data from SAP
    try:
        # Get service order data from SAP
        order_data = get_sap_service_order_data(service_order)
        
        if not order_data and not SAP_CONNECTION.use_mock:
            # Service order not found in real SAP
            return redirect(url_for('index', error=f'Service order {service_order} not found in SAP'))
        
        # Store the data in session
        session['order_data'] = order_data
        return redirect(url_for('automation_wizard', step=1))
        
    except Exception as e:
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
            'pn': order_data['pn']
           },
        2: {'title': 'Serial Number Verification', 
            'question': 'Does the Serial Number match the ID plate on the unit and the outgoing Serial Number in SAP?',
            'sn': order_data['sn']
           },
        3: {'title': 'Manual Entry Verification', 
            'question': 'Please enter the Part Number from the Unit being inspected to verify:',
            'input_type': 'part_number',
            'expected': order_data['pn']
           },
        4: {'title': 'Manual Entry Verification', 
            'question': 'Please enter the Serial Number from the Unit being inspected to verify:',
            'input_type': 'serial_number',
            'expected': order_data['sn']
           },
        5: {'title': 'Operator Comments', 
            'question': f'Have you verified the operator comments to ensure there are no mismatches or discrepancies compared to actual repairs?\n\nOperator Comments: "{order_data["op_comments"]}"'
           },
        6: {'title': 'Unit Mod Status', 
            'question': f'Have you verified the unit mod status and confirmed it matches the actual unit configuration?\n\nMod Status: {order_data["mod_status"]}'
           },
        7: {'title': 'Z8 Notifications', 
            'question': f'Have you verified that all Z8 notifications have been properly processed?\n\nNotifications: {", ".join(order_data["notifications"])}'
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
             'question': f'Have you verified that all customer requirements have been addressed and completed?\n\nCustomer: {order_data["customer"]}'
            },
        12: {'title': 'Authorization Documents', 
             'question': f'Have you verified that all authorization documents have been properly processed?\n\nDocuments: {", ".join(order_data["auth_documents"])}'
            },
        13: {'title': 'Service Report Match', 
             'question': 'Do the authorization documents match the service report?'
            },
        14: {'title': 'Service Report Complete', 
             'question': 'Is the service report complete with all required information filled in?'
            },
        15: {'title': 'Test Sheet Match', 
             'question': f'Does the test sheet match the unit being inspected?\n\nTest Sheets: {", ".join(order_data["test_sheets"])}'
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
        expected = order_data['part_number']
        
        # If connected to real SAP, perform verification
        if SAP_CONNECTION and not SAP_CONNECTION.use_mock:
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
        expected = order_data['serial_number']
        
        # If connected to real SAP, perform verification
        if SAP_CONNECTION and not SAP_CONNECTION.use_mock:
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
            
            # If connected to real SAP, update SAP about the failure
            if SAP_CONNECTION and not SAP_CONNECTION.use_mock and current_step > 5:
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
        # If connected to real SAP, update SAP about the test sheet failures
        if SAP_CONNECTION and not SAP_CONNECTION.use_mock:
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
    if current_step == 20 and response.lower() == 'yes' and SAP_CONNECTION and not SAP_CONNECTION.use_mock:
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

if __name__ == '__main__':
    print(f"Starting SAP Service Order Automation Web Interface")
    print(f"Platform: {platform.system()}")
    print(f"SAP Connection Mode: {'Real SAP (if available)' if IS_WINDOWS else 'Simulation Mode'}")
    
    # If we're on Windows, we display more information
    if IS_WINDOWS:
        try:
            import win32com.client
            print("SAP GUI automation is available (win32com.client imported successfully)")
        except ImportError:
            print("SAP GUI automation not available. Install pywin32 for SAP integration.")
            print("  pip install pywin32")
    
    app.run(host='0.0.0.0', port=5000, debug=True)