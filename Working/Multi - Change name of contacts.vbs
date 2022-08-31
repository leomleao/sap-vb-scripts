Set objArgs = WScript.Arguments
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
'    Set session    = connection.Children(0)
   Set session    = connection.Children(CInt(objArgs(0)))
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If


contactsStr = objArgs(1)
contacts = Split(contactsStr,",")

session.findById("wnd[0]/tbar[0]/okcd"). Text = "/nVAP2"
session.findById("wnd[0]").sendVKey 0


For i = 0 To UBound(contacts)
    session.findById("wnd[0]/usr/ctxtRF02D-PARNR").text = contacts(i)
    session.findById("wnd[0]").sendVKey 0
    If InStr(session.findById("wnd[0]/sbar").text,"marked for deletion") > 0 Then
        session.findById("wnd[0]").sendVKey 0
    End if
    test = true
    Do
    If InStr(session.findById("wnd[0]/sbar/pane[0]").text,"currently blocked by user U081715") > 0 Then
        WScript.sleep 1000
        session.findById("wnd[0]").sendVKey 0
    Else
        test = false
    End If 
    Loop While test = true

    session.findById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtADDR3_DATA-NAME_LAST").text = "DELETE"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 11
Next