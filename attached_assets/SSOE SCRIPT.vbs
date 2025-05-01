on ERROR RESUME NEXT
do
    if Not IsObject(application) Then
        Set SapGuiAuto = GetObject("SAPGUI")
        Set application = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(connection) Then
        Set connection = application.Children(0)
    End If
    If Not IsObject(session) Then
        Set session = connection.Children(0)
    End If
    If IsObject(WScript) Then
        WScript.ConnectObject session, "on"
        WScript.ConnectObject application, "on"
    End If

    if err.number <> 0 then
        resSAP = msgbox("SAP was not detected please open a SAP session.  Thank you.", vbOkCancel)
        err.clear
        if resSAP = vbCancel then
            wscript.quit
        end if
    else
        exit do
    end if
loop
on error goto 0

Billable = False
dim tmpAdminCodeStatus
dim salesvalue
isSPEX = False
newSPEX = False
version10 = False
STAYLABOREDON = False
isConverted = False
isExchange = False
isDPMI = False
OOPS = False
boolSPEX = False

dim tmpServOrder
dim tmpCust
'dim tmpCustomerNUM
'dim salesvalue
dim tmpDELBLCK
dim tmpZHStatus
dim tmpZGStatus
dim tmpMODSIN
dim tmpMODSOUT
dim tmpSFMODSIN
dim tmpSFMODSOUT
dim tmpSWVERSIONSIN
dim tmpSWVERSIONSOUT
dim tmpPN

UserName = session.Info.User

strText = "Oh no"
set objVoice = CreateObject("SAPI.SpVoice")

    dim msgBoxResult

    dim tmpSN

    isSPEX = False
    newSPEX = False
    version10 = False
    STAYLABOREDON = False

boolError = False
UserName = session.Info.User

Call Main

Sub Main

	dim tmpServOrd
	dim tmpCust
	isSPEX = False

	dim tmpPN
	dim tmpSN

	dim msgBoxResult

	Call GetServiceOrderInput(tmpServOrd)

   'call recordSO(tmpServOrd,UserName)

	If tmpServOrd = "" then 
		MsgBox("No Service Order Number entered.")
		Exit Sub
	End If



	Call OpenZiwbn(tmpServOrd)
	
	Call GetIW32Information(tmpServOrd, isSPEX)
	
	Call OpenZiwbn(tmpServOrd)

	On Error Resume Next
	'Call LaborOn
	Call GetPartNumberSerialNumber(tmpPn, tmpSn)
	Call AskAboutPartNumberSerialNumber(tmpPn,tmpSn)
	
	Call AskAboutOperatorComments

	if boolError = True Then 
		Call LaborOFFIncomplete
		Exit Sub
	End if 

	Call AskAboutUnitModStatus

	If boolError = True Then 
		'Call LaborOFFIncomplete
		Exit Sub
	End If


	Call Z8Notifications
	If boolError = True Then 
		'Call LaborOFFIncomplete
		Exit Sub
	End If

	'Call ProcessAuthDocs
	'if boolError = True Then 
		'Call LaborOFFIncomplete
	'	Exit Sub
	'End if 

	Call AskAboutHardware
	if boolError = True Then 
		'Call LaborOFFIncomplete
		Exit Sub
	End if 

	Call AskAboutConnectors
	if boolError = True Then 
		'Call LaborOFFIncomplete
		Exit Sub
	End if 

	Call AskAboutFOD
	if boolError = True Then 
		'Call LaborOFFIncomplete
		Exit Sub
	End if 

	'Call AskAboutCompletedCustomerRequest
	if boolError = True Then 
		'Call LaborOFFIncomplete
		Exit Sub
	End if 
	
	if isSPEX = False then
        'Call checkQT

        Call AskAboutCustomerReq

        'Call Z8Notifications

        'Call checkWARRANTY

        'Call checkZTASKS

    End if
	
	Call ProcessAuthDocs
	if boolError = True Then 
		'Call LaborOFFIncomplete
		Exit Sub
	End if 

	Call AskAboutAuthDocsMatchingServiceReport
	if boolError = True Then 
		'Call LaborOFFIncomplete
		Exit Sub
	End if 

	Call AskAboutIsServiceReportComplete
	if boolError = True Then 
		'Call LaborOFFIncomplete
		Exit Sub
	End if 

	Call AskAboutDoesTestSheetMatchUnit
	if boolError = True Then 
		'Call LaborOFFIncomplete
		Exit Sub
	End if 

	Call AskAboutDoesTestSheetShowAnyFails
	if boolError = True Then 
		'Call LaborOFFIncomplete
		Exit Sub
	End if 

	Call AskAboutDateAndSignatureOnTestSheet
	if boolError = True Then 
		'Call LaborOFFIncomplete
		Exit Sub
	End if 

	'msgbox(isSPEX)

	'If isSPEX = False then 
	'	Call CheckSalesOrder
	'End If

	if boolError = True Then 
		'Call LaborOFFIncomplete
		Exit Sub
	End If

	'Call UpdateWandingStatus
	if boolError = True Then 
		'Call LaborOFFIncomplete
		Exit Sub
	End if 

	

	'Call LaborOFFComplete

	Call CheckInspectionTabIndicators
	if boolError = True Then 
	'	Call LaborOFFIncomplete
		Exit Sub
	End if 

	Call AskAboutHasTheRepairmanLineBeenSigned
	if boolError = True Then 
	'	Call LaborOFFIncomplete
		Exit Sub
	End if
	
	Call UpdateWSUPDComments
	if boolError = True Then 
		'Call LaborOFFIncomplete
		Exit Sub
	End if 



	MsgBox("SSOE Complete")
	'session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
	'session.findById("wnd[0]").sendVKey 0
	'session.findById("wnd[0]").sendVKey 0

End Sub


Sub GetServiceOrderInput(servOrd)
	servOrd = InputBox("Enter service order number?", "SSOE")
	    If servOrd = "" then
        wscript.quit
		End If
End Sub


Sub GetIW32Information(servOrd, boolSPEX)

    tmpCustomerNUM = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_KUNUM").text
    tmpSuperiorOrder = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB1:SAPLYAFF_ZIWBNGUI:0100/ssubSUB2:SAPLYAFF_ZIWBNGUI:0102/ctxtW_INP_DATA").text

	' session.findById("wnd[0]/tbar[0]/okcd").text = "/nIW33"
	' session.findById("wnd[0]").sendVKey 0
	' session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = servOrd
	' session.findById("wnd[0]").sendVKey 0
	' session.findById("wnd[0]").sendVKey 0
	' session.findById("wnd[0]").sendVKey 0

	'tmpCust = session.findById("wnd[0]/usr/subSUB_ALL:SAPLCOIH:3001/ssubSUB_LEVEL:SAPLCOIH:1100/tabsTS_1100/tabpIHKZ/ssubSUB_AUFTRAG:SAPLCOIH:1120/subSUB_ADRESSE:SAPLIPAR:0704/tabsTSTRIP_700/tabpKUND/ssubTSTRIP_SCREEN:SAPLIPAR:0130/subADRESSE:SAPLIPAR:0150/txtDIADR-NAME1").text

	if tmpCustomerNUM = "PLANT1133" then
        boolSPEX = TRUE
    end if

    if tmpCustomerNUM = "SLSR01" then
        boolSPEX = TRUE
    end if

    if tmpCustomerNUM = "PLANT1057" then
        boolSPEX = TRUE
    end if

    if tmpCustomerNUM = "PLANT1052" then
        boolSPEX = TRUE
    end if

    if tmpCustomerNUM = "PLANT1013" then
        boolSPEX = TRUE
    end if

    if tmpCustomerNUM = "PLANT1103" then
        boolSPEX = TRUE
    end if

    if tmpCustomerNUM = "PLANT1116" then
        boolSPEX = TRUE
    end if

    if tmpCustomerNUM = "PLANT1005" then
        boolSPEX = TRUE
    end if


End Sub 

Sub OpenZIWBN(tmpServOrd)
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nZIWBN"
	session.findById("wnd[0]").sendVKey 0
	session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB1:SAPLYAFF_ZIWBNGUI:0100/ssubSUB2:SAPLYAFF_ZIWBNGUI:0102/ctxtW_INP_DATA").text = tmpServOrd
	session.findById("wnd[0]").sendVKey 0
End Sub

Sub LaborOn
	session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I").select

	session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").pressToolbarButton "&MB_FILTER"
	session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell").currentCellRow = 7
	session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell").selectedRows = "7"
	session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell").doubleClickCurrentCell
	session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/btn600_BUTTON").press
	session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "Close Up Inspection"
	session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 19
	session.findById("wnd[2]").sendVKey 0
	session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").currentCellColumn = ""
	session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectedRows = "0"
	session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").pressToolbarButton "LABON"
End Sub


Sub GetPartNumberSerialNumber(pn,sn)
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H").select
    on error resume next
    err.clear
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell").currentCellColumn = "MATNR"
    if err.number <> 0 then
        version10 = True
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont[0]/shell").currentCellColumn = "MATNR"
        pn = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont[0]/shell").getCellValue(0, "MATNR")

        sn = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont[0]/shell").getCellValue(0, "SERNR")

        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
    else
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell").currentCellColumn = "MATNR"
        pn = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell").getCellValue(0, "MATNR")

        sn = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpEQUIPMENT_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0233/cntlG_CNTR_HDR_EQUIPMENT/shellcont/shell").getCellValue(0, "SERNR")
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select

    end if
    on error goto 0
    err.clear
End Sub

Sub AskAboutPartNumberSerialNumber(pn, sn)
   snTries = 0
    pnTries = 0
    inputPNRes = ""
    inputSNRes = ""
    res = ""

    res1 = msgbox("Does the Part Number match the ID plate on the unit and the outgoing Part Number in SAP?", VBYesNo, "Check")
    if res1 = vbNo then
        objVoice.Speak strText
        wscript.quit
    end if

    res2 = msgbox("Does the Serial Number match the ID plate on the unit and the outgoing Serial Number in SAP?", VBYesNo, "Check")
    if res1 = vbNo then
        objVoice.Speak strText
        wscript.quit
    end if

    Do While pn <> inputPNRes
        inputPNRes = InputBox("Please enter the Part Number from the Unit being inspected.")
        if pn = inputPNRes then
            Exit Do
        else

            if pnTries > 1 then
                objVoice.Speak strText
                MsgBox("The Part Number being tried does not appear to match SAP. This script will now terminate.")
                wscript.quit
            end if
            objVoice.Speak strText
            MsgBox("The Part Number entered does not match the Part Number of this record in SAP. Please double check your entry.")
            pnTries = pnTries + 1
        end if
    Loop

    dim result
    dim tmpSTATUS
    dim SearchChar

    SearchChar = "DPMI"

    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
    tmpSTATUS = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_SUBORD_STAT").TEXT
    result = InStr(tmpSTATUS, SearchChar)
    If result = 0 Then
        Do While UCASE(sn) <> UCASE(inputSNRes)
            inputSNRes = InputBox("Please enter the Serial Number from the unit being inspected.")
            If UCASE(sn) = UCASE(inputSNRes) then
                Exit Do
            else
                if snTries > 1 Then
                    objVoice.Speak strText
                    Msgbox("The Serial Number being tried does not appear to match SAP. This script will now terminate.")
                    wscript.quit
                end if
                objVoice.Speak strText
                MsgBox("The Serial Number entered does not match the Serial Number of this record in SAP. Please double check your entry.")
                snTries = snTries + 1
            End If
        Loop
    end if
    If result <> 0 Then
        isDPMI = True
    End If
End Sub



Sub AskAboutUnitModStatus

	res = MsgBox("Does the Mod Plate Information match the Service Report, including Mods, PN, SN?", vbYesNo, "Modification Status")
	If res = vbNo then 
		boolError = True
		MsgBox("Mod Status Failure. Script terminated.")
	End If

End Sub


Sub Z8Notifications
	session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
	session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/btnG_H_SERORD_BT_ALERT").press

	MsgBox("Open each Z8 to find special instructions about this order.")

	res = MsgBox("Have all applicable Z8 Notifications been complied with?", vbYesNo, "Z8 Compliance")

	If res = vbNo then 
		MsgBox("Z8 Notifications have not been fully complied with. Script terminated.")
		boolError = True
	End If 

	session.findById("wnd[1]").close
End Sub


Sub ProcessAuthDocs
	session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpAUTHDOCS_I").select
	'session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpAUTHDOCS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0313/cntlG_CNTR_AUTH_DOC/shellcont/shellcont/shell/shellcont[0]/shell").setCurrentCell -1,"USTAT"
	'session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpAUTHDOCS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0313/cntlG_CNTR_AUTH_DOC/shellcont/shellcont/shell/shellcont[0]/shell").selectColumn "USTAT"
	'session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpAUTHDOCS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0313/cntlG_CNTR_AUTH_DOC/shellcont/shellcont/shell/shellcont[0]/shell").contextMenu
	'session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpAUTHDOCS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0313/cntlG_CNTR_AUTH_DOC/shellcont/shellcont/shell/shellcont[0]/shell").selectContextMenuItem "&SORT_ASC"
End Sub



Sub AskAboutHardware
	res = MsgBox("Has all hardware been checked?" & vbcrlf & "(No loose/missing screws, knobs turn freely, handles tight, etc...)", vbYesNo, "Hardware Status")
	If res = vbNo then 
		boolError = True
		MsgBox("Hardware Failure. Script terminated.")

	End If
End Sub



Sub AskAboutConnectors
	res = MsgBox("Has connectors been checked?" & vbcrlf & "(No bent pins, FOD, etc...)", vbYesNo, "Connector Status")
	If res = vbNo then 
		boolError = True
		MsgBox("Connector Check Failure. Script terminated.")

	End If
End Sub



Sub AskAboutFOD
	res = MsgBox("Does the unit pass FOD check?", vbYesNo, "FOD Status")
	If res = vbNo then 
		boolError = True
		MsgBox("FOD Check Failure. Script terminated.")

	End If
End Sub



Sub AskAboutCompletedCustomerRequest
	res = MsgBox("Does the work performed match what customer requested?", vbYesNo, "Customer Workscope Status")
	If res = vbNo then 
		boolError = True
		MsgBox("Customer Workscope Failure. Script terminated.")

	End If
End Sub



Sub AskAboutAuthDocsMatchingServiceReport
	res = MsgBox("Do the documents used in AuthDocs match whats on the Service Report?", vbYesNo, "AuthDocs Status")
	If res = vbNo then 
		boolError = True
		MsgBox("AuthDocs Failure. Script terminated.")

	End If
End Sub



Sub AskAboutIsServiceReportComplete
	res = MsgBox("Has the Service Report been reviewed for incorrect or missing data?", vbYesNo, "Service Report Status")
	If res = vbNo then 
		boolError = True
		MsgBox("Service Report Failure. Script terminated.")

	End If
End Sub



Sub AskAboutDoesTestSheetMatchUnit
	res = MsgBox("Is the test sheet header information correct? P/N, S/N, Cal Dates, etc...", vbYesNo, "Test Sheet Header Status")
	If res = vbNo then 
		boolError = True
		MsgBox("Test Sheet Header Failure. Script terminated.")

	End If
End Sub



Sub AskAboutDoesTestSheetShowAnyFails
	res = MsgBox("Did all tests pass or have been corrected on test sheet?", vbYesNo, "Test Sheet Data Status")
	If res = vbNo then 
		boolError = True
		MsgBox("Test Sheet Data Failure. Script terminated.")

	End If
End Sub



Sub AskAboutDateAndSignatureOnTestSheet
	res = MsgBox("Was the test sheet signed and dated?", vbYesNo, "Test Sheet Signature Status")
	If res = vbNo then 
		boolError = True
		MsgBox("Test Sheet Signature Failure. Script terminated.")

	End If
End Sub





'***********************************			THIS ASKS IF THE UNIT IS OR IS NOT SPEX. IF/THEN to Process SALES ORDER
'********************************************************************************************************************************
Sub CheckSalesOrder

	'msgbox("Running CheckSalesOrder")

	session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
	session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_SALES_ORD").setFocus

	session.findById("wnd[0]").sendVKey 2
	

	res = MsgBox("Does the dollar amount in the Sales Order match what was noted on the traveler?", vbYesNo, "Funded Sales Order")

	if res = vbNo then 
		boolError = True
		MsgBox("The sales order is eather not funded, or does not match the notes on the traveler. Script terminated.")
	End If

	'Below button press to kick out of Sales Order
	session.findById("wnd[0]/tbar[0]/btn[3]").press
End Sub



Sub UpdateWandingStatus
	ON ERROR RESUME NEXT
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
    IF ERR.NUMBER <> 0 THEN
        msgbox("Oops can you fix that")
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
        err.clear
        on error goto 0
    END IF

    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/ctxtYAFS_ZIWBN_HEADER-SRV_LOCATION").text = "FININSP"
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/ctxtYAFS_ZIWBN_HEADER-SRV_LOCATION").caretPosition = 7
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/btnG_H_SRV_BT_SAVE").press
End Sub



Sub UpdateWSUPDComments

 dim tmpDateTime
    tmpDateTime = Now
	
		ON ERROR RESUME NEXT
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
    IF ERR.NUMBER <> 0 THEN
        msgbox("Oops can you fix that")
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
        err.clear
        on error goto 0
    END IF

    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/ctxtYAFS_ZIWBN_HEADER-SRV_LOCATION").text = "FININSP"
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/ctxtYAFS_ZIWBN_HEADER-SRV_LOCATION").caretPosition = 7
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/btnG_H_SRV_BT_SAVE").press

    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/btnG_H_SERORD_BT_LTXT").press
    session.findById("wnd[1]/usr/cntlW_TEXT_LTXT/shellcont/shell").text = "SSOE, Rev 1, " + CStr(tmpDateTime) + "  " + UserName + vbCr+session.findById("wnd[1]/usr/cntlW_TEXT_LTXT/shellcont/shell").text
    session.findById("wnd[1]/tbar[0]/btn[11]").press
End Sub




Sub LaborOFFComplete
	session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I").select
	session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").pressToolbarButton "&MB_FILTER"
	session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell").currentCellRow = 7
	session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell").selectedRows = "7"
	session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell").doubleClickCurrentCell
	session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/btn600_BUTTON").press
	session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "Close Up Inspection"
	session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 19
	session.findById("wnd[2]").sendVKey 0
	session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").currentCellColumn = ""
	session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").selectedRows = "0"
	session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpOPERATIONS_I/ssubG_IWB_ITEMS:SAPLYAFF_ZIWBNGUI:0314/cntlG_CNTR_ITM_OPERATION/shellcont/shellcont/shell/shellcont[0]/shell").pressToolbarButton "LABOFFC"
End Sub



Sub AskAboutHasTheRepairmanLineBeenSigned
	res = MsgBox("Has the SSOE checklist been filled out properly and followed?", vbYesNo, "SSOE Sign Status")
	If res = vbNo then 
		boolError = True
		MsgBox("SSOE Failure. Script terminated.")

	End If

	'ress = MsgBox("Did the Repairman sign the Repairman line on the Unit Status Report?", vbYesNo, "Repairman Sign Status")
	'If ress = vbNo then 
		'boolError = True
		'MsgBox("Repairman Sign Failure. Script terminated.")

	'End If
End Sub



Sub CheckInspectionTabIndicators

    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpFINALINSP_I").select
    if boolSPEX = False then
        session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/ssubSUB3:SAPLYAFF_ZIWBNGUI:0300/subSUB2:SAPLYAFF_ZIWBNGUI:0302/tabsG_ITEMS_TBSTR_CTRL/tabpFINALINSP_I/ssubG_IWB_ITEMS:SAPLYAFF_GDBE:0110/btnW_SIM").press
    end if
    MsgBox("Check red indicators for errors.")
    'if newSPEX = True then
     '   exit sub
    'end if

    'if STAYLABOREDON = True then
     '   msgbox("You will have to labor off complete of Final before sending to QA.")
    'end if

    'if newSPEX = True then
     '   exit sub
    'end if

End Sub




Sub LaborOFFIncomplete
	
		MsgBox("")
End Sub

sub recordSO(tmpServOrder,tmpUSER_RESPONSIBLE)
		 Const fsoForAppend = 8

		 Dim objFSO
		 Set objFSO = CreateObject("Scripting.FileSystemObject")
		 d = now

'		 dim tmpServOrder 
'		 tmpServOrder = "501111"
'		 dim tmpUSER_RESPONSIBLE
'		 tmpUSER_RESPONSIBLE = "e848528"


		 'Open the text file
		 Dim objTextStream
		 Set objTextStream = objFSO.OpenTextFile("S:\CSC DataBases\RepairmanScript\Reference\autorun.TXT", fsoForAppend)

		 'Display the contents of the text file
		 objTextStream.WriteLine(tmpServOrder & ", " & tmpUSER_RESPONSIBLE & ", "& formatdatetime(d,2) & ", " & formatdatetime(d,4) &", " & "Repairman")
		 'objTextStream.Write Now

		 'Close the file and clean up
		 objTextStream.Close
		 Set objTextStream = Nothing
		 Set objFSO = Nothing

end sub

sub AskAboutOperatorComments
    USRres = msgbox("Are the Operation comments filled out on each labor line?", vbYesNo, "OP comments")
    If USRres = vbNo then
        objVoice.Speak strText
        MsgBox("!!ALERT!!" & vbcrlf & "Operation comments must be filled out. TERMINATING SCRIPT.")
        wscript.quit
    End If
end sub

sub AskAboutCustomerReq

        err.clear
		
		'if boolSPEX = False then
		dim result
        dim tmpAdminCodeStatus
        dim SearchChar
		 SearchChar = tmpAdminCodeStatus
		 
		session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpNOTIFICATION_H").select
		session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpNOTIFICATION_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0232/cntlG_CNTR_HDR_NOTIFICATION/shellcont/shell").doubleClickCurrentCell
		session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02").select
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]").sendVKey 0		
		'session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0102/tabsNOTIF_TAB/tabpADMN/ssubADMN:SAPLXQQM:9007/ctxtQMEL-YYWRPC").setFocus
		session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0102/tabsNOTIF_TAB/tabpADMN").select
		tmpAdminCodeStatus = session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0102/tabsNOTIF_TAB/tabpADMN/ssubADMN:SAPLXQQM:9007/ctxtQMEL-YYWRPC").text
		session.findById("wnd[0]/tbar[0]/btn[3]").press
		session.findById("wnd[0]").sendVKey 0
				session.findById("wnd[0]").sendVKey 0
						session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
		
        'session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpNOTIFICATION_H").select
		'tmpAdminCodeStatus = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpNOTIFICATION_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0232/cntlG_CNTR_HDR_NOTIFICATION/shellcont/shell").text
		'session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
 
		
				'msgbox(tmpAdminCodeStatus)
				'else
		
		    CRDres = msgbox("Have you reviewed the PO for special instructions?", vbYesNo, "Customer Requirements")
    if CRDres = vbNo then
        objVoice.Speak strText
        msgbox("!!ALERT" & vbcrlf & "You must review these for customers, TERMINATING SCRIPT.")
        wscript.quit
    end if
	'end if
	
'if boolSPEX = False then
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
    session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0211/txtYAFS_ZIWBN_HEADER-SRV_SALES_ORD").setFocus
    session.findById("wnd[0]").sendVKey 2
    session.findById("wnd[0]").sendVKey 0

    salesvalue = session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBAK-NETWR").text

        'SearchChar = tmpAdminCodeStatus
		
		'tmpAdminCodeStatus = session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpNOTIFICATION_H/ssubG_IWB_HEADER:SAPLYAFF_ZIWBNGUI:0232/cntlG_CNTR_HDR_NOTIFICATION/shellcont/shell").text
		
		'session.findById("wnd[0]/usr/subSUB1:SAPLYAFF_ZIWBNGUI:0011/subSUB2:SAPLYAFF_ZIWBNGUI:0200/subSUB2:SAPLYAFF_ZIWBNGUI:0202/tabsG_HEADER_TBSTRP_CTRL/tabpSERORDER_H").select
		

		
		if tmpAdminCodeStatus = "CNA" or tmpAdminCodeStatus = "BTC" then
		Billable = True
		else 
		end if
		'exit sub
		
		if Billable = True then
			if salesvalue <> 0 then
            RES = msgbox("Does Net value match traveler? This unit is " & tmpAdminCodeStatus & " and requires a value in the sales order", vbYesNo, "Net Value")
				if RES = vbNo then
                session.findById("wnd[0]/tbar[0]/btn[3]").press
                wscript.quit
				end if
			else
				msgbox("Script ending, needs money in sales order before continuing.")
				session.findById("wnd[0]/tbar[0]/btn[3]").press
				wscript.quit
			end if
			session.findById("wnd[0]/tbar[0]/btn[3]").press
		else
			res = MsgBox("Does the dollar amount in the Sales Order match the funded quote or funded PO?" & vbcrlf & "(Units that are warranty or MSA should be a 0$ value)", vbYesNo, "Funded Sales Order")
			if res = vbNo then
				objVoice.Speak strText
				MsgBox("!!ALERT!!" & vbcrlf & "The sales order is either not funded, or does not match the notes on the traveler. TERMINATING SCRIPT.")
				session.findById("wnd[0]/tbar[0]/btn[3]").press
				wscript.quit
			End If

			tmpDELBLCK = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4400/ssubHEADER_FRAME:SAPMV45A:4440/cmbVBAK-LIFSK").key
			if tmpDELBLCK = "" then

			else
			msgbox("There is a Delivery Block, before going to QA this Delivery Block must be removed, see your workflow/supervisor for help.")
			session.findById("wnd[0]/tbar[0]/btn[3]").press
			wscript.quit
			end if

			session.findById("wnd[0]/tbar[0]/btn[3]").press
		End if
		'end if

end sub