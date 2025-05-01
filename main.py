from flask import Flask, render_template, request, redirect, url_for
import os
import sys
import datetime

app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "dev-secret-key")

# Import the SAP Service Order Automation class
from sap_service_order_automation import SapServiceAutomation

# Global context processor to add date to all templates
@app.context_processor
def inject_now():
    return {'now': datetime.datetime.now().strftime('%Y-%m-%d')}

# Configure routes
@app.route('/')
def index():
    """Render the main web interface"""
    return render_template('index.html')

@app.route('/run_automation', methods=['POST'])
def run_automation():
    """Handle the form submission to start automation"""
    service_order = request.form.get('service_order', '')
    if not service_order:
        return redirect(url_for('index', error='Please enter a service order number'))
    
    # Store the service order number in session for the wizard
    # (We'll use a query parameter for now as we're not setting up a database)
    return redirect(url_for('automation_wizard', service_order=service_order, step=1))

@app.route('/automation_wizard')
def automation_wizard():
    """Render the automation wizard interface"""
    service_order = request.args.get('service_order', '')
    step = int(request.args.get('step', 1))
    
    if not service_order:
        return redirect(url_for('index', error='Service order number is missing'))
    
    # Logic for different steps of the wizard
    steps = {
        1: {'title': 'Part Number Verification', 
            'question': 'Does the Part Number match the ID plate on the unit and the outgoing Part Number in SAP?',
            'pn': 'MK-12345' # In a real app, this would come from SAP
           },
        2: {'title': 'Serial Number Verification', 
            'question': 'Does the Serial Number match the ID plate on the unit and the outgoing Serial Number in SAP?',
            'sn': 'SN987654' # In a real app, this would come from SAP
           },
        3: {'title': 'Manual Entry Verification', 
            'question': 'Please enter the Part Number from the Unit being inspected to verify:',
            'input_type': 'part_number'
           },
        4: {'title': 'Manual Entry Verification', 
            'question': 'Please enter the Serial Number from the Unit being inspected to verify:',
            'input_type': 'serial_number'
           },
        5: {'title': 'Operator Comments', 
            'question': 'Have you verified the operator comments to ensure there are no mismatches or discrepancies compared to actual repairs?'
           },
        6: {'title': 'Unit Mod Status', 
            'question': 'Have you verified the unit mod status and confirmed it matches the actual unit configuration?'
           },
        7: {'title': 'Z8 Notifications', 
            'question': 'Have you verified that all Z8 notifications have been properly processed?'
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
             'question': 'Have you verified that all customer requirements have been addressed and completed?'
            },
        12: {'title': 'Authorization Documents', 
             'question': 'Have you verified that all authorization documents have been properly processed?'
            },
        13: {'title': 'Service Report Match', 
             'question': 'Do the authorization documents match the service report?'
            },
        14: {'title': 'Service Report Complete', 
             'question': 'Is the service report complete with all required information filled in?'
            },
        15: {'title': 'Test Sheet Match', 
             'question': 'Does the test sheet match the unit being inspected?'
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
    
    return render_template('wizard.html', 
                          service_order=service_order,
                          step_data=steps[step],
                          current_step=step,
                          total_steps=len(steps))

@app.route('/process_step', methods=['POST'])
def process_step():
    """Process a step in the wizard"""
    service_order = request.form.get('service_order', '')
    current_step = int(request.form.get('current_step', 1))
    response = request.form.get('response', 'no')
    
    # Special handling for manual entry steps
    if current_step == 3:  # Part number verification
        user_input = request.form.get('manual_input', '')
        expected = 'MK-12345'  # In a real app, this would be from SAP
        if user_input != expected:
            # If first attempt, give another chance
            if 'retry' not in request.form:
                return render_template('wizard.html',
                                      service_order=service_order,
                                      step_data={
                                          'title': 'Part Number Verification - Retry',
                                          'question': 'The Part Number does not match. Please try again:',
                                          'input_type': 'part_number',
                                          'error': f'Expected: {expected}, You entered: {user_input}'
                                      },
                                      current_step=current_step,
                                      total_steps=20,
                                      retry=True)
            else:
                # Second failure, exit the process
                return render_template('error.html',
                                     title='Part Number Mismatch',
                                     message='The Part Number being tried does not appear to match SAP. The process has been terminated.')
    
    if current_step == 4:  # Serial number verification
        user_input = request.form.get('manual_input', '')
        expected = 'SN987654'  # In a real app, this would be from SAP
        if user_input != expected:
            # If first attempt, give another chance
            if 'retry' not in request.form:
                return render_template('wizard.html',
                                      service_order=service_order,
                                      step_data={
                                          'title': 'Serial Number Verification - Retry',
                                          'question': 'The Serial Number does not match. Please try again:',
                                          'input_type': 'serial_number',
                                          'error': f'Expected: {expected}, You entered: {user_input}'
                                      },
                                      current_step=current_step,
                                      total_steps=20,
                                      retry=True)
            else:
                # Second failure, exit the process
                return render_template('error.html',
                                     title='Serial Number Mismatch',
                                     message='The Serial Number being tried does not appear to match SAP. The process has been terminated.')
    
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
    return redirect(url_for('automation_wizard', service_order=service_order, step=next_step))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)