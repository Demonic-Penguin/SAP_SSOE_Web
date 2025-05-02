"""
Web Application with File-Based SAP Data
This version reads SAP data from a file instead of connecting directly
"""

from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify
import os
import sys
import datetime
import json
import time

app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "dev-secret-key")
app.config['SESSION_TYPE'] = 'filesystem'

# Global cache for SAP data to avoid frequent lookups
SAP_DATA_CACHE = {}

# Path to the SAP data file
SAP_DATA_FILE = "sap_data.json"

def get_service_order_data(service_order):
    """
    Get service order data from the file or cache
    """
    # Check cache first
    if service_order in SAP_DATA_CACHE:
        return SAP_DATA_CACHE[service_order]
    
    # Try to read from file
    try:
        if os.path.exists(SAP_DATA_FILE):
            with open(SAP_DATA_FILE, 'r') as f:
                data = json.load(f)
                
            # Only use the data if it matches the requested service order
            if data.get('service_order') == service_order:
                # Cache the data
                SAP_DATA_CACHE[service_order] = data
                return data
    except Exception as e:
        print(f"Error reading SAP data file: {e}")
    
    # If no data found or an error occurred, use simulation data
    return simulate_service_order_data(service_order)

def simulate_service_order_data(service_order):
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

# Global context processor to add date to all templates
@app.context_processor
def inject_now():
    # Check if we have SAP data file
    sap_mode = "File Data Mode"
    if not os.path.exists(SAP_DATA_FILE):
        sap_mode = "Simulation Mode"
    
    return {
        'now': datetime.datetime.now().strftime('%Y-%m-%d'),
        'sap_mode': sap_mode
    }

@app.route('/')
def index():
    """Render the main web interface"""
    # Check if we have SAP data file
    sap_status = "Using SAP data from file"
    if not os.path.exists(SAP_DATA_FILE):
        sap_status = "No SAP data file found. Will use simulation data."
    
    # Check data file timestamp
    file_details = {}
    if os.path.exists(SAP_DATA_FILE):
        try:
            file_stat = os.stat(SAP_DATA_FILE)
            file_details['size'] = file_stat.st_size
            file_details['modified'] = datetime.datetime.fromtimestamp(file_stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
            
            # Read the contents to get service order
            with open(SAP_DATA_FILE, 'r') as f:
                data = json.load(f)
                file_details['service_order'] = data.get('service_order', 'Unknown')
                file_details['part_number'] = data.get('part_number', 'Unknown')
                file_details['serial_number'] = data.get('serial_number', 'Unknown')
        except Exception as e:
            file_details['error'] = str(e)
    
    return render_template('index_sap_enabled.html', 
                          sap_status=sap_status,
                          file_details=file_details)

@app.route('/sap_status')
def sap_status():
    """Check SAP data status"""
    if os.path.exists(SAP_DATA_FILE):
        try:
            file_stat = os.stat(SAP_DATA_FILE)
            modified = datetime.datetime.fromtimestamp(file_stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
            
            with open(SAP_DATA_FILE, 'r') as f:
                data = json.load(f)
                
            return jsonify({
                'status': 'file',
                'message': f'Using SAP data from file (last updated: {modified})',
                'service_order': data.get('service_order', 'Unknown')
            })
        except Exception as e:
            return jsonify({
                'status': 'error',
                'message': f'Error reading SAP data file: {e}'
            })
    else:
        return jsonify({
            'status': 'simulation',
            'message': 'No SAP data file found. Will use simulation data.'
        })

@app.route('/run_automation', methods=['POST'])
def run_automation():
    """Handle the form submission to start automation"""
    service_order = request.form.get('service_order', '')
    if not service_order:
        return redirect(url_for('index', error='Please enter a service order number'))
    
    # Store the service order number in session for the wizard
    session['service_order'] = service_order
    
    # Check if there's a data file with this service order
    file_mode = False
    if os.path.exists(SAP_DATA_FILE):
        try:
            with open(SAP_DATA_FILE, 'r') as f:
                data = json.load(f)
                if data.get('service_order') == service_order:
                    file_mode = True
        except:
            pass
    
    session['sap_mode'] = 'file' if file_mode else 'simulation'
    
    # Try to get the service order data
    try:
        # Get service order data from file or simulation
        order_data = get_service_order_data(service_order)
        
        if not order_data:
            # Service order not found
            return redirect(url_for('index', error=f'Service order {service_order} not found'))
        
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
            order_data = get_service_order_data(service_order)
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
            order_data = get_service_order_data(service_order)
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

if __name__ == '__main__':
    print(f"Starting SAP-Enabled Web Application (File Mode)")
    
    # Check if SAP data file exists
    if os.path.exists(SAP_DATA_FILE):
        try:
            file_stat = os.stat(SAP_DATA_FILE)
            modified = datetime.datetime.fromtimestamp(file_stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
            
            with open(SAP_DATA_FILE, 'r') as f:
                data = json.load(f)
                
            print(f"Found SAP data file: {SAP_DATA_FILE}")
            print(f"  Last modified: {modified}")
            print(f"  Service order: {data.get('service_order', 'Unknown')}")
            print(f"  Part number: {data.get('part_number', 'Unknown')}")
            print(f"  Serial number: {data.get('serial_number', 'Unknown')}")
        except Exception as e:
            print(f"Error reading SAP data file: {e}")
            print(f"Will use simulation data")
    else:
        print(f"No SAP data file found. Will use simulation data.")
        print(f"To use real SAP data, run sap_data_extractor.py first.")
    
    app.run(host='0.0.0.0', port=5000, debug=True)