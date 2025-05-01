# SAP Data Integration Guide

This guide explains how the SAP Service Order Automation application can be configured to pull real data from SAP instead of using simulation data.

## Overview

The application has been enhanced to:
1. Attempt to connect to a real SAP system when running on Windows
2. Pull actual service order data including part numbers, serial numbers, and other details
3. Update SAP in real-time as the wizard progresses
4. Maintain fallback simulation when SAP is unavailable

## How SAP Data Integration Works

The enhanced application (`main_sap_integrated.py`) implements several key functions:

### 1. SAP Connection

The application attempts to establish a connection to SAP using:
```python
import win32com.client
sap_gui_auto = win32com.client.GetObject("SAPGUI")
application = sap_gui_auto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)
```

### 2. Data Retrieval

When a service order number is entered, the application queries SAP for details:
```python
def get_sap_service_order_data(service_order):
    # In real implementation, would do something like:
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW32"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtVORG").text = service_order
    session.findById("wnd[0]").sendVKey(0)
    
    # And then extract the data from various fields
    part_number = session.findById("...").text
    serial_number = session.findById("...").text
    # etc.
```

### 3. Real-time Updates

As the user progresses through the wizard, the application updates SAP:
```python
# Example of updating SAP when a step fails
session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW32"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/ctxtVORG").text = service_order
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/tabsMAIN/tabpDOC/...").text = "Failed at step X: Details..."
session.findById("wnd[0]").sendVKey(0)
```

## Customizing for Your SAP System

To adapt this application to your specific SAP system:

### 1. Identify Correct SAP GUI Paths

You need to determine the correct object paths in your SAP GUI:
- Use the SAP GUI Scripting Tracker (part of SAP GUI tools)
- Enable scripting in SAP GUI and record the actions you want to automate
- Note the ID paths shown in the recorder

### 2. Update the SAP Integration Functions

Modify these functions in `main_sap_integrated.py`:
- `get_sap_service_order_data`: Update to use your SAP system's specific paths
- Process steps that update SAP (in the `process_step` function)

### 3. Add Error Handling for SAP-Specific Issues

Different SAP systems may have different error conditions:
- Add checks for specific SAP errors in your environment
- Create appropriate fallbacks for each error condition

## Example: Customizing for IW32 Transaction

If your service orders use the IW32 transaction, you might need code like this:

```python
def get_sap_service_order_data(service_order):
    try:
        # Navigate to IW32
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW32"
        session.findById("wnd[0]").sendVKey(0)
        
        # Enter service order number
        session.findById("wnd[0]/usr/ctxtAUFNR").text = service_order
        session.findById("wnd[0]").sendVKey(0)
        
        # Extract data from various tabs
        # Note: These paths are examples and must be adjusted for your system
        
        # General data tab
        part_number = session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\\01/ssubSUB_DATA:SAPLIQS0:7235/subGENERAL:SAPLIQS0:7212/txtLTAP-MATNR").text
        
        # Equipment tab
        session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\\02").select()
        serial_number = session.findById("wnd[0]/usr/tabsTABSTRIP/tabpT\\02/ssubSUB_DATA:SAPLIQS0:7236/subOBJ:SAPLIQS0:7322/txtVIQMEL-SERGE").text
        
        # Return the collected data
        return {
            'service_order': service_order,
            'part_number': part_number,
            'serial_number': serial_number,
            # Add other fields as needed
        }
    except Exception as e:
        print(f"Error retrieving data from SAP: {e}")
        return None
```

## Security Considerations

When integrating with real SAP systems:
1. Ensure the application runs with appropriate user credentials
2. SAP scripting must be enabled in your SAP system
3. Consider the impacts of automation on your SAP system's performance
4. Never store SAP credentials in the code

## Testing SAP Integration

Before deploying to production:
1. Test with a few known service orders in a development/test SAP system
2. Verify that the data is correctly extracted from all fields
3. Check that updates to SAP are performed correctly
4. Verify that error handling works as expected

## Further Resources

- SAP GUI Scripting API Documentation: [SAP Help Portal](https://help.sap.com/viewer/b47d018c3b9b45e897faf66a6c0885a8/LATEST/en-US)
- PyWin32 Documentation: [PyWin32 GitHub](https://github.com/mhammond/pywin32)