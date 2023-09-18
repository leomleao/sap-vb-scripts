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

session.findById("wnd[0]/tbar[0]/okcd").text = "/nsq01"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/mbar/menu[5]/menu[0]").select
session.findById("wnd[1]/usr/radRAD1").select
session.findById("wnd[1]/tbar[0]/btn[2]").press
session.findById("wnd[0]/mbar/menu[1]/menu[7]").select

For i = 0 To session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").rowcount - 1
    If session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").getCellValue(i,"DBGBNUM") = "Z_MM_UK" Then
        session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = i
    End If
Next

session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell").currentCellRow = 8
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "PROD_ORD_DATE"
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").setFocus
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").caretPosition = 13
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem "&XXL"

session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\u081715\OneDrive - WAGO\08_Rugby All\Production\Production_Schedule\Data"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "PO 1 conf dates.XLSX"
session.findById("wnd[1]/tbar[0]/btn[11]").press