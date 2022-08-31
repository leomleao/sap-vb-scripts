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
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

session.findById("wnd[0]/tbar[0]/okcd"). Text = "/nVD02"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = "800753"
session.findById("wnd[1]/usr/ctxtRF02D-VKORG").text = "0008"
session.findById("wnd[1]/usr/ctxtRF02D-VTWEG").text = "00"
session.findById("wnd[1]/usr/ctxtRF02D-SPART").text = "00"
session.findById("wnd[1]").sendVKey 0

   test = true
   Do
   If session.ActiveWindow.Name = "wnd[2]" Then
      WScript.sleep 1000
      session.findById("wnd[2]/tbar[0]/btn[0]").press
      session.findById("wnd[1]").sendVKey 0         
   Else
      test = false
   End If 
   Loop While test = true