# SAP Integration Guide

This guide explains how to use the SAP Service Order Automation tool with a real SAP connection instead of the simulation mode.

## Prerequisites

1. **Windows Operating System**: SAP GUI automation is only supported on Windows.

2. **SAP GUI for Windows**: You must have SAP GUI for Windows installed.
   - Recommended version: SAP GUI 7.50 or newer
   - SAP Scripting must be enabled (usually done by your SAP administrator)

3. **Python Libraries**:
   - `pywin32`: Required for COM automation to control SAP GUI
     ```
     pip install pywin32
     ```
   - Other required packages:
     ```
     pip install flask flask-sqlalchemy tkinter pyttsx3
     ```

4. **Active SAP Session**: 
   - You must have an active SAP GUI session open
   - You must be logged into the SAP system before running the script

## Running with Real SAP Connection

### Option 1: Using the local_sap_run.py Script

The simplest way to run with a real SAP connection is to use the included script:

```bash
python local_sap_run.py
```

This script will:
1. Verify you're on Windows
2. Check if win32com.client is available
3. Attempt to connect to an active SAP session
4. Run the SAP Service Order Automation with real SAP controls

### Option 2: Modifying the Code Directly

If you need to customize the behavior, you can modify the relevant code:

1. In `sap_service_order_automation.py`, find this code (around line 800):
   ```python
   use_mock = sys.platform != "win32" or IN_REPLIT
   app = SapServiceAutomation(use_mock=use_mock)
   ```

2. You can manually override this to:
   ```python
   app = SapServiceAutomation(use_mock=False)
   ```

## Troubleshooting SAP Connections

### Common Issues

1. **Cannot Find SAP GUI Automation**: 
   - Make sure SAP GUI is running
   - Make sure you have at least one SAP system connection open
   - Verify SAP Scripting is enabled (ask your SAP administrator)

2. **Permission Errors**:
   - SAP scripting requires specific permissions
   - Your SAP user may not have sufficient rights
   - Check with your SAP administrator about scripting permissions

3. **Import Error for win32com**:
   - Install the pywin32 package: `pip install pywin32`
   - If installation fails, download from [GitHub releases](https://github.com/mhammond/pywin32/releases)

4. **SAP GUI Version Issues**:
   - Different SAP GUI versions may have slightly different object models
   - If you encounter errors about missing properties or methods, check your SAP GUI version

## Notes for Developers

If you're extending the application to work with SAP:

1. Use the SAP GUI Scripting Tracker tool (comes with SAP GUI) to identify object IDs
2. Always handle exceptions, as SAP GUI automation can be fragile
3. Test your scripts with minimal user impact first

## Security Considerations

SAP GUI automation has security implications:
- Scripts can perform any action the user can, including sensitive transactions
- SAP administrators often restrict scripting for security reasons
- Never store SAP credentials in scripts

## Further Resources

- [SAP GUI Scripting API Documentation](https://help.sap.com/viewer/b47d018c3b9b45e897faf66a6c0885a8/LATEST/en-US)
- [SAP GUI Scripting Security](https://help.sap.com/viewer/8ecea00c1f854fd0a433c4aef5da1ea2/LATEST/en-US/6b189d8d3a6a4541a3c29527695d0880.html)