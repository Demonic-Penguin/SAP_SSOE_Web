# SAP Service Order Automation

Python script for automating SAP service order processing, converted from VBS.

## Overview

This script automates a Service Sheet of Excellence (SSOE) process for service orders in SAP. It guides technicians through a checklist to ensure all required steps are completed when processing a service order.

## Requirements

- Python 3.x
- PyWin32 (`pip install pywin32`)
- pyttsx3 (`pip install pyttsx3`)
- tkinter (included in standard Python installation)
- Active SAP GUI with scripting enabled

## Features

- Connects to SAP GUI using COM automation
- Guides users through a comprehensive service order checklist
- Validates part numbers and serial numbers
- Verifies operator comments, unit modifications, and hardware conditions
- Processes authorization documents and service reports
- Provides voice feedback for critical errors

## Usage

1. Ensure SAP GUI is open and you're logged in
2. Run the script: `python sap_service_order_automation.py`
3. Enter the service order number when prompted
4. Follow the on-screen prompts to complete the SSOE process

## Conversion Notes

This script was converted from a VBS script to Python. The original workflow and functionality have been preserved, with adaptations for Python's syntax and libraries:

- COM automation is handled via PyWin32 instead of native VBS
- Dialog boxes are created with tkinter instead of VBScript's MsgBox
- Text-to-speech is implemented with pyttsx3 instead of SAPI.SpVoice
- Error handling follows Python's exception model while maintaining the same logic
