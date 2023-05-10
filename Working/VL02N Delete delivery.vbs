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

session.findById("wnd[0]/tbar[0]/okcd"). Text = "/nVL02n"
session.findById("wnd[0]").sendVKey 0

strData =  "820769781,820769782,820769783,820769784,820769785,820769786,820769787,820769788,820769789,820769790,820769791,820769792,820769793,820769794,820769795,820769796,820769797,820769798,820769799,820769800,820769801,820769802,820769803,820769804,820769805,820769806,820769807,820769808,820769809,820769810,820769811,820769812,820769813,820769814"


arr = Split(strData,",")

For i = 0 To UBound(arr)  
    session.findById("wnd[0]/usr/ctxtLIKP-VBELN").text = arr(i)
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[14]").press
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
Next




