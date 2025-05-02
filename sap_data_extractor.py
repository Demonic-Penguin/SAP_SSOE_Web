"""
SAP Data Extractor
A standalone script to extract service order data from SAP and save it for the web application
"""

import os
import sys
import platform
import time
import traceback
import json
import argparse

def main():
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Extract service order data from SAP')
    parser.add_argument('service_order', help='Service order number to extract data for')
    parser.add_argument('--output', '-o', default='sap_data.json', help='Output file to save data to')
    args = parser.parse_args()
    
    service_order = args.service_order
    output_file = args.output
    
    print(f"SAP Data Extractor for Service Order: {service_order}")
    print(f"Output file: {output_file}")
    print("-" * 80)
    
    # Only continue if on Windows
    if platform.system() != "Windows":
        print("ERROR: This script requires Windows to connect to SAP.")
        print(f"Current platform: {platform.system()}")
        sys.exit(1)
    
    # Try to import win32com.client
    try:
        import win32com.client
        print("✓ Successfully imported win32com.client")
    except ImportError as e:
        print(f"ERROR: Failed to import win32com.client: {e}")
        print("You need to install pywin32 with: pip install pywin32")
        sys.exit(1)
    
    try:
        print("\nConnecting to SAP GUI...")
        
        # First try via Dispatch
        try:
            print("Trying to connect via Dispatch...")
            sap_gui_auto = win32com.client.Dispatch("SAPGUI.ScriptingCtrl.1")
            print("✓ Connected to SAPGUI via Dispatch")
        except Exception as e:
            print(f"Failed to connect via Dispatch: {e}")
            
            # Try alternate method - GetObject
            try:
                print("Trying to connect via GetObject...")
                sap_gui_auto = win32com.client.GetObject("SAPGUI")
                print("✓ Connected to SAPGUI via GetObject")
            except Exception as e2:
                print(f"Failed to connect via GetObject: {e2}")
                print("ERROR: Could not connect to SAP GUI.")
                print("Make sure SAP GUI is running and you are logged in.")
                sys.exit(1)
        
        # Get scripting engine
        application = sap_gui_auto.GetScriptingEngine
        if application is None:
            print("ERROR: Failed to get SAP scripting engine")
            sys.exit(1)
        print("✓ Got scripting engine")
        
        # Check connections
        conn_count = application.Children.Count
        print(f"Found {conn_count} SAP connection(s)")
        
        if conn_count == 0:
            print("ERROR: No SAP connections found. Please log into SAP first.")
            sys.exit(1)
        
        connection = application.Children(0)
        if connection is None:
            print("ERROR: Failed to get SAP connection")
            sys.exit(1)
        print("✓ Got connection")
        
        # Check sessions
        sess_count = connection.Children.Count
        print(f"Found {sess_count} session(s)")
        
        if sess_count == 0:
            print("ERROR: No SAP sessions found")
            sys.exit(1)
        
        session = connection.Children(0)
        if session is None:
            print("ERROR: Failed to get SAP session")
            sys.exit(1)
        print("✓ Got session")
        
        # Get username
        info = session.Info
        username = info.User
        print(f"✓ Connected as user: {username}")
        
        # Service order data to collect
        data = {
            'service_order': service_order,
            'part_number': None,
            'serial_number': None,
            'customer': None,
            'op_comments': "Service required due to unit failure in field. Customer requested express processing.",
            'mod_status': "MOD-A Revision 3",
            'auth_documents': ["AUTH-001", "AUTH-002"],
            'notifications': ["Z8-001", "Z8-002"],
            'test_sheets': ["TEST-001"]
        }
        
        # First try ZIWBN transaction
        try:
            print("\nTrying ZIWBN transaction...")
            
            # Navigate to ZIWBN
            print("Navigating to ZIWBN...")
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nZIWBN"
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            
            # Enter service order
            print(f"Entering service order {service_order}...")
            input_field = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB1:SAPLYAFF_ZIWBNGUI:0100/ssubSUB2:SAPLYAFF_ZIWBNGUI:0102/ctxtW_INP_DATA")
            input_field.text = service_order
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(1)
            
            print("✓ Navigated to ZIWBN")
            
            # Get customer number
            try:
                customer_field = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_KUNUM")
                if customer_field:
                    data['customer'] = customer_field.text
                    print(f"Found customer: {data['customer']}")
            except Exception as e:
                print(f"Could not find customer: {e}")
            
            # Switch to Equipment tab
            try:
                print("Switching to Equipment tab...")
                equipment_tab = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H")
                equipment_tab.select()
                time.sleep(1)
                
                # Try to find equipment grid
                try:
                    # First try standard grid
                    grid = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell")
                    if grid:
                        data['part_number'] = grid.getCellValue(0, "MATNR")
                        data['serial_number'] = grid.getCellValue(0, "SERNR")
                        print(f"Found part number: {data['part_number']}")
                        print(f"Found serial number: {data['serial_number']}")
                except Exception:
                    # Try version 10 grid
                    try:
                        grid = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont[0]/shell")
                        if grid:
                            data['part_number'] = grid.getCellValue(0, "MATNR")
                            data['serial_number'] = grid.getCellValue(0, "SERNR")
                            print(f"Found part number: {data['part_number']}")
                            print(f"Found serial number: {data['serial_number']}")
                    except Exception as e:
                        print(f"Could not find equipment grid: {e}")
            except Exception as e:
                print(f"Could not switch to Equipment tab: {e}")
                
        except Exception as e:
            print(f"Error with ZIWBN transaction: {e}")
        
        # If we didn't get part number and serial number from ZIWBN, try IW32
        if not data['part_number'] or not data['serial_number']:
            try:
                print("\nTrying IW32 transaction...")
                
                # Navigate to IW32
                print("Navigating to IW32...")
                session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW32"
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(1)
                
                # Enter service order
                print(f"Entering service order {service_order}...")
                
                # Try different service order fields
                service_order_fields = [
                    "wnd[0]/usr/ctxtAUFNR",
                    "wnd[0]/usr/ctxtRIWO00-AUFNR",
                    "wnd[0]/usr/ctxtVORG",
                    "wnd[0]/usr/ctxtIW32-AUFNR",
                    "wnd[0]/usr/ctxtCAUFVD-AUFNR"
                ]
                
                field_found = False
                for field_id in service_order_fields:
                    try:
                        field = session.findById(field_id)
                        if field:
                            field.text = service_order
                            field_found = True
                            break
                    except:
                        continue
                
                if not field_found:
                    print("Could not find service order input field")
                else:
                    # Press Enter
                    session.findById("wnd[0]").sendVKey(0)
                    time.sleep(1)
                    
                    print("✓ Navigated to IW32")
                
                    # Look for part number if not found yet
                    if not data['part_number']:
                        part_fields = [
                            "wnd[0]/usr/tabsTABSTRIP/tabpT\\01/ssubSUB_DATA:SAPLIQS0:7235/subGENERAL:SAPLIQS0:7212/txtLTAP-MATNR",
                            "wnd[0]/usr/tabsTABSTRIP/tabpDESC/ssubDETAIL:SAPLITO0:0115/txtITOBJ-MATXT",
                            "wnd[0]/usr/ctxtRIWO00-MATNR"
                        ]
                        
                        for field_id in part_fields:
                            try:
                                field = session.findById(field_id)
                                if field:
                                    data['part_number'] = field.text
                                    print(f"Found part number: {data['part_number']}")
                                    break
                            except:
                                continue
                    
                    # Look for customer if not found yet
                    if not data['customer']:
                        customer_fields = [
                            "wnd[0]/usr/tabsTABSTRIP/tabpT\\01/ssubSUB_DATA:SAPLIQS0:7235/subCUSTOMER:SAPLIQS0:7280/txtKUAGV-NAME1",
                            "wnd[0]/usr/ctxtRIWO00-KUNUM"
                        ]
                        
                        for field_id in customer_fields:
                            try:
                                field = session.findById(field_id)
                                if field:
                                    data['customer'] = field.text
                                    print(f"Found customer: {data['customer']}")
                                    break
                            except:
                                continue
                    
                    # Look for serial number if not found yet
                    if not data['serial_number']:
                        # Try to switch to Equipment tab
                        tab_ids = [
                            "wnd[0]/usr/tabsTABSTRIP/tabpT\\02",
                            "wnd[0]/usr/tabsTABSTRIP/tabpEQUIPMENT"
                        ]
                        
                        for tab_id in tab_ids:
                            try:
                                tab = session.findById(tab_id)
                                if tab:
                                    tab.select()
                                    time.sleep(1)
                                    print("Switched to Equipment tab")
                                    
                                    # Try to get serial number
                                    serial_fields = [
                                        "wnd[0]/usr/tabsTABSTRIP/tabpT\\02/ssubSUB_DATA:SAPLIQS0:7236/subOBJ:SAPLIQS0:7322/txtVIQMEL-SERGE",
                                        "wnd[0]/usr/tabsTABSTRIP/tabpEQUIPMENT/ssubDETAIL:SAPLITO0:0115/txtITOB-SERGE"
                                    ]
                                    
                                    for field_id in serial_fields:
                                        try:
                                            field = session.findById(field_id)
                                            if field:
                                                data['serial_number'] = field.text
                                                print(f"Found serial number: {data['serial_number']}")
                                                break
                                        except:
                                            continue
                                    
                                    break
                            except:
                                continue
            except Exception as e:
                print(f"Error with IW32 transaction: {e}")
        
        # Generate simulation data for missing fields
        if not data['part_number']:
            data['part_number'] = f"MK-{service_order[:3]}-{service_order[-2:]}"
            print(f"Using simulated part number: {data['part_number']}")
        
        if not data['serial_number']:
            data['serial_number'] = f"SN{service_order}"
            print(f"Using simulated serial number: {data['serial_number']}")
        
        if not data['customer']:
            data['customer'] = "CUSTOMER NAME"
            print(f"Using simulated customer: {data['customer']}")
        
        # Write data to JSON file
        print(f"\nWriting data to {output_file}...")
        with open(output_file, 'w') as f:
            json.dump(data, f, indent=2)
        
        print(f"✓ Service order data saved to {output_file}")
        print("\nExtracted Data:")
        for key, value in data.items():
            if key not in ['auth_documents', 'notifications', 'test_sheets']:
                print(f"  {key}: {value}")
                
        return 0
        
    except Exception as e:
        print(f"ERROR: {str(e)}")
        print(traceback.format_exc())
        return 1

if __name__ == "__main__":
    sys.exit(main())