'-Begin-----------------------------------------------------------------

'-Directives----------------------------------------------------------
Option Explicit
On Error Resume Next

'-Variables-----------------------------------------------------------
Dim WSHShell, SAPGUIPath, SID, InstanceNo, application, SapGuiAuto, connection, session, connected, d, e, i, j, y, q, u, gridView, reqQty, withdrawlQty, rowCount, visibleRowCount, iterations, productionOrders, reqQty, withdrawlQty, qtyItems, attempts, test

'-Help functions -----------------------------------------------------
Dim svc, sQuery, cProc, iniProc

'-Connection----------------------------------------------------------------
Set WSHShell = WScript.CreateObject("WScript.Shell")
If IsObject(WSHShell) Then
   '-Set the path to the SAP GUI directory---------------------------
   SAPGUIPath = "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\"

   '-Set the SAP system ID-------------------------------------------
   SID = "WGP"

   '-Starts the SAP GUI----------------------------------------------
   WSHShell.Exec SAPGUIPath & "sapshcut.exe start -system=" & SID & " -maxgui"   
   Set WSHShell = Nothing
End If

connected = false

Do While connected = false
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
   If session.children.count > 1 Then
      session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").select
      session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").setFocus
      session.findById("wnd[1]/tbar[0]/btn[0]").press
   End If    
   If Err.Number = 0 Then
      connected = true
   End If
   Err.Clear 
   WScript.sleep 1000
Loop

Action(session)

'-End Connection------------------------------------------------------------


'' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' ''
'Supporting procedures and functions
'' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' ''
' add item to array
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

Sub Action(session)
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nCOOIS"
    session.findById("wnd[0]").sendVKey 0     

    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").key = "PPIOM000"
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = "/UK_COMP"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_WERKS-LOW").text = "0008"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").text = "CNF"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").text = "TECO"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").selected = true
    session.findById("wnd[0]/tbar[1]/btn[8]").press

    If Session.ActiveWindow.Name = "wnd[1]" Then
        If session.findById("wnd[1]/usr/txtMESSTXT1").text = "There is no data for the selection" Then
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
            session.findById("wnd[0]").sendVKey 0
            Exit Sub
        End If
    End If

    productionOrders = Array()

    gridView = "wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell"
    rowCount = session.findById(gridView).rowCount

    If rowCount = 0 Then
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
        session.findById("wnd[0]").sendVKey 0
        Exit Sub
    End If

    For i = 0 To rowCount - 1 : Do
    reqQty = session.findById(gridView).getCellValue(i,"MENGE")
    withdrawlQty = session.findById(gridView).getCellValue(i,"DENMNG")
    If reqQty <> withdrawlQty Then
        productionOrders = addItem(productionOrders, session.findById(gridView).getCellValue(i,"AUFNR"))
    End If   
    Loop While False: Next

    productionOrders = removeDuplicates(productionOrders)


    ' Section to go on the POs, remove TECO and remove the Final issue Flag for each unpicked component line
    For i = 0 To UBound(productionOrders) : Do

      session.findById("wnd[0]/tbar[0]/okcd").text = "/nco02"
      session.findById("wnd[0]").sendVKey 0
      session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = productionOrders(i)
      session.findById("wnd[0]").sendVKey 0

      If Session.ActiveWindow.Name = "wnd[1]" Then
         If InStr(session.findById("wnd[1]/usr/txtMESSTXT1").text,"""Change"" is not allowed") > 0 Then
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            session.findById("wnd[0]/mbar/menu[1]/menu[7]/menu[4]").select
         End if
      End if
      session.findById("wnd[0]/tbar[1]/btn[6]").press
      visibleRowCount = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120").visiblerowcount
      'Total of rows in the table
      rowCount = session.findById("wnd[0]/usr/txtRC27X-ENTRIES").text
      'How many iterations needs to be done to process every single line
      iterations = rowCount \ visibleRowCount

      For y = 0 To iterations
         For q = 0 To visibleRowCount - 1
            qtyReq = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3," & q &"]").text
            qtyWithdrawn = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-DENMNG[4," & q &"]").text
            If qtyWithdrawn < qtyReq Then
               If Not session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/chkRESBD-XLOEK[17," & q &"]").selected Then
                  session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/chkRESBD-KZEAR[23," & q &"]").selected = false
               End If
            End If
         Next 
         'Press page down button (Not tcode specific)
         session.findById("wnd[0]/tbar[0]/btn[82]").press
      Next     

      'Save
      session.findById("wnd[0]/tbar[0]/btn[11]").press   
    Loop While False: Next
    ' END Section to go on the POs, remove TECO and remove the Final issue Flag for each unpicked component line

    ' Picking in CO27
    For u = 0 To UBound(productionOrders) : Do
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nco27"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/usr/ctxtS_AUFNR-LOW").text = productionOrders(i)
        session.findById("wnd[0]/tbar[1]/btn[8]").press
        session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").currentCellColumn = ""
        session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectedRows = "0"
        session.findById("wnd[0]/tbar[1]/btn[8]").press
        session.findById("wnd[0]/usr/subPUSHBUTTON:SAPLCOWB:0400/btnMALL").press
        session.findById("wnd[0]/usr/subPUSHBUTTON:SAPLCOWB:0400/btnCHFI").press
        session.findById("wnd[0]/tbar[0]/btn[11]").press
    Loop While False: Next
    ' END Picking in CO27

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nCOOIS"
    session.findById("wnd[0]").sendVKey 0     

    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").key = "PPIOM000"
    session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = "/UK_COMP"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_WERKS-LOW").text = "0008"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").text = "CNF"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").text = "TECO"
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").selected = true
    session.findById("wnd[0]/tbar[1]/btn[8]").press

    If Session.ActiveWindow.Name = "wnd[1]" Then
        If session.findById("wnd[1]/usr/txtMESSTXT1").text = "There is no data for the selection" Then
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
            session.findById("wnd[0]").sendVKey 0
            Exit Sub
        End If
    End If

    If session.findById(gridView).rowCount = 0 Then
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
        session.findById("wnd[0]").sendVKey 0
        Exit Sub
    End If

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
End Sub