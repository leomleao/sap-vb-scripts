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

session.findById("wnd[0]/tbar[0]/okcd").text = "/nCOOIS"
session.findById("wnd[0]").sendVKey 0     

session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").key = "PPIOM000"
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_WERKS-LOW").text = "0008"
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST1").text = "DLV"
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SYST2").text = "TECO"
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/chkP_KZ_E2").selected = true
session.findById("wnd[0]/tbar[1]/btn[8]").press


' If session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").ToolbarButtonCount < 10 Then
'    session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton "&NAVIGATION_PROFILE_TOOLBAR_EXPAND"
' End If
' session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton "&MB_FILTER"
' session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem "&FILTER"

' gridView = "wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell"
' rowCount = session.findById(gridView).rowCount
' For i = 0 To rowCount - 3 : Do
'    currentRow = session.findById(gridView).getCellValue(i,"SELTEXT")
'    If currentRow = "Requirement quantity (EINHEIT)" or currentRow = "Quantity withdrawn (EINHEIT)" Then
'       Call session.findById(gridView).DoubleClick(i,"SELTEXT")
'    End If
' Loop While False: Next

' session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/btn600_BUTTON").press
' session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/txt%%DYN002-LOW").setFocus
' session.findById("wnd[2]").sendVKey 2
' session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 0,"TEXT"
' session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "0"
' session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
' session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/txt%%DYN003-LOW").setFocus 
' session.findById("wnd[2]").sendVKey 2
' session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 3,"TEXT"
' session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "3"
' session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
' session.findById("wnd[2]/tbar[0]/btn[0]").press


productionOrders = Array()

gridView = "wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell"
rowCount = session.findById(gridView).rowCount

For i = 0 To rowCount - 1 : Do
   reqQty = session.findById(gridView).getCellValue(i,"MENGE")
   withdrawlQty = session.findById(gridView).getCellValue(i,"DENMNG")
   If reqQty <> withdrawlQty Then
      productionOrders = addItem(productionOrders, session.findById(gridView).getCellValue(i,"AUFNR"))
   End If   
Loop While False: Next

productionOrders = removeDuplicates(productionOrders)

For i = 0 To UBound(productionOrders) : Do
   Wscript.Echo productionOrders(i)
Loop While False: Next


session.findById("wnd[0]/tbar[0]/okcd").text = "/nMB1A"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").text = "261"
session.findById("wnd[0]/usr/ctxtRM07M-WERKS").text = "0008"
session.findById("wnd[0]/usr/ctxtRM07M-LGORT").text = "0001"

' strData = "3847118,3847139,3847151,3847153,3847154,3847158,3847159,3849575,3849692,3849700,3849702,3849703,3849704,3849705,3849707,3849708,3849710,3849711,3849714,3849717,3849719,3849720,3849721,3849723,3849727"

' productionOrders = Split(strData,",")

For i = 0 To UBound(productionOrders) : Do
    session.findById("wnd[0]/tbar[1]/btn[13]").press
    session.findById("wnd[1]/usr/sub:SAPMM07M:1405/ctxtRM07M-AUFNR[0,0]").text = productionOrders(i)
    session.findById("wnd[1]").sendVKey 0 

    If Session.ActiveWindow.Name = "wnd[2]" Then
        session.findById("wnd[2]/tbar[0]/btn[0]").press
        session.findById("wnd[1]/tbar[0]/btn[12]").press
        Exit Do 
    End if
    
    qtyItems = session.findById("wnd[0]/usr/txtRM07M-POSNM").text

    If qtyItems > 31 Then
        For j = 1 To 31
            session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[" & j - 1 &",53]").text = "0000999908"
        Next 
        session.findById("wnd[0]/tbar[0]/btn[82]").press
        For j = 1 To qtyItems - 31
            session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[" & j - 1 &",53]").text = "0000999908"
        Next 
    Else
        For j = 1 To qtyItems
            session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[" & j - 1 &",53]").text = "0000999908"
        Next
    End If 

    If session.findById("wnd[0]/sbar/pane[0]").text <> "" Then
       session.findById("wnd[0]").sendVKey 0 
    End If      

    session.findById("wnd[0]/tbar[1]/btn[9]").press

    If session.findById("wnd[0]/sbar/pane[0]").text <> "" Then
       session.findById("wnd[0]").sendVKey 0 
    End If 
    session.findById("wnd[0]/tbar[0]/btn[11]").press

    test = true
    attempts = 0

    Do
        If session.findById("wnd[0]/sbar/pane[0]").text <> "" Then
            WScript.Sleep 500
            If InStr(session.findById("wnd[0]/sbar/pane[0]").text,"Deficit of BA Unrestricted-use") > 0 Then
                session.findById("wnd[0]/usr/txtMSEG-ERFMG").text = 0            
            End If
            session.findById("wnd[0]").sendVKey 0 
            attempts = attempts + 1
            If attempts > qtyItems Then
                test = false
            End If
        Else
            test = false
        End If 
    Loop While test = true
    
Loop While False: Next