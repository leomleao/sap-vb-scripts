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


poStr =  "4511595,4510632,4487442,4504689,4503056,4506025,4507411,4510515,4510630,4633291,4633292,4633293,4633294,4512183,4487441,4494894,4494890,4494891,4417267,4487380,4487375,4474909,4512085,4554041,4633202,4506000,4507405,4507364,4496343,4609989,4406006,4406005,4460608,4504608,4639391,4633200,4608726,4448511,4456301,4502946,4492874"

pos = Split(poStr,",")


For i = 0 To UBound(pos) : Do
session.findById("wnd[0]/tbar[0]/okcd").text = "/nco02"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = pos(i)
session.findById("wnd[0]").sendVKey 0
   If session.ActiveWindow.Name = "wnd[1]" Then
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[0]/mbar/menu[1]/menu[7]/menu[10]").select
   Else 
      session.findById("wnd[0]/mbar/menu[1]/menu[9]/menu[9]").select
   End If

   If session.ActiveWindow.Name = "wnd[1]" Then
session.findById("wnd[1]").sendVKey 0
 End If
session.findById("wnd[0]").sendVKey 11

Loop While False: Next



