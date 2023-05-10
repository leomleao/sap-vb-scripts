Option Explicit
Dim WSHShell, SAPGUIPath, SID, InstanceNo, application, SapGuiAuto, connection, session, connected, d, e, i, j, gridView, rowCount, productionOrders, reqQty, withdrawlQty, qtyItems, attempts, test

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

Function addItem(arr, val)
    ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = val
    addItem = arr
End Function

' remove duplicates from array
Function removeDuplicates(arr)
   Set d = CreateObject("Scripting.Dictionary")
   For Each e In arr
   d(e) = e
   Next
   removeDuplicates = d.Items
End Function

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nCOOIS"
    session.findById("wnd[0]").sendVKey 0     

    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").key = "PPIOM000"
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = "/UK_COMP"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_WERKS-LOW").text = "0008"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SELID").text = "ZCOFA01"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").text = "TECO"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").selected = true
    session.findById("wnd[0]/tbar[1]/btn[8]").press



    productionOrders = Array()
    gridView = "wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell"
    rowCount = session.findById(gridView).rowCount
    


    For l = 0 To rowCount - 1 : Do   
        reqQty = session.findById(gridView).getCellValue(l,"MENGE")
        withdrawlQty = session.findById(gridView).getCellValue(l,"DENMNG")
        If reqQty <> withdrawlQty Then
            productionOrders = addItem(productionOrders, session.findById(gridView).getCellValue(l,"AUFNR"))
        End If   
    Loop While False: Next

    productionOrders = removeDuplicates(productionOrders)
   
    ' strData = "4146907,4146908,4223908,4234058,4234059,4249790,4272584,4278434,4278439,4284044,4284096,4284143,4286058,4286059,4286060,4286061,4286065,4286068,4288088,4294166,4296011"

    ' productionOrders = Split(strData,",")

    For i = 0 To UBound(productionOrders) : Do

        ' For each production order remvove the final issue of anything not deleted and that is missing picking.
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nco02"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = productionOrders(i)
        session.findById("wnd[0]/tbar[1]/btn[6]").press

        visibleRowCount = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120").visiblerowcount
        rowCount = session.findById("wnd[0]/usr/txtRC27X-ENTRIES").text

        iterations = rowCount \ visibleRowCount
        for j = 0 to iterations
            For y = 0 To (visibleRowCount - 1)         
            currentItem = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/ctxtRESBD-MATNR[1," & y & "]").text
            if currentItem <> "" Then        
                If session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/chkRESBD-XLOEK[17," & y & "]").selected = 0 Then           
                    reqQty2 = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3," & y & "]").text
                    withdrawlQty2 = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-DENMNG[4," & y & "]").text
                    If withdrawlQty2 < reqQty2 Then
                    session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/chkRESBD-KZEAR[23," & y & "]").selected = false
                    End If
                End If         
            End if
            Next
            session.findById("wnd[0]/tbar[0]/btn[82]").press
        Next
        session.findById("wnd[0]").sendVKey 11 
        
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nmigo"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0011/cmbGODYNPRO-ACTION").key = "A07"
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0011/cmbGODYNPRO-REFDOC").key = "R08"
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0011/ctxtGODEFAULT_TV-BWART").text = "261"
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2070/ctxtGODYNPRO-ORDER_NUMBER").text = "4146907"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMLIST:SAPLMIGO:0200/subSUB_BUTTONS:SAPLMIGO:0210/btnOK_DETERMINE_ALL").press
        session.findById("wnd[0]/tbar[0]/btn[11]").press
                
    Loop While False: Next

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nCOOIS"
    session.findById("wnd[0]").sendVKey 0     

    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").key = "PPIOM000"
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = "/UK_COMP"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_WERKS-LOW").text = "0008"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SELID").text = "ZCOFA01"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").text = "TECO"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").selected = true
    session.findById("wnd[0]/tbar[1]/btn[8]").press



    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton "&NAVIGATION_PROFILE_TOOLBAR_EXPAND"
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem "&XXL"
    session.findById("wnd[1]").sendVKey 4
    session.findById("wnd[2]/usr/ctxtDY_PATH").text = "C:\Users\u081715\OneDrive - WAGO\Desktop\TEMP\Automations\Production order missing picking\"
    session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = "Orders-missing-picking" & "_" & Day(Now) & "-" & Month(Now) & "-" & Year(Now) & ".XLSX"
    session.findById("wnd[2]/tbar[0]/btn[11]").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    Dim oShell : Set oShell = CreateObject("WScript.Shell")
    oShell.Run "taskkill /f /im excel.exe"

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
    session.findById("wnd[0]").sendVKey 0