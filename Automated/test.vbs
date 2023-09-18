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

For i = 0 To session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").rowcount - 1
    If session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").getCellValue(i,"DBGBNUM") = "Z_MM_UK" Then
        session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = i
    End If
Next