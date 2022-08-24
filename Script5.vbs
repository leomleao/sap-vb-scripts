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

strData = "800513,7311804,800216,800217,800247,801681,802160,802228,802403,806984,807375,808793,809861,810256,810288,810479,812279,7307559,7313210,7323130"
arr = Split(strData,",")

For i = 0 To UBound(arr) : Do
   session.findById("wnd[0]/tbar[0]/okcd").text = "/nvd03"
   session.findById("wnd[0]").sendVKey 0
   session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = arr(i)
   session.findById("wnd[1]/usr/ctxtRF02D-VKORG").text = "0008"
   session.findById("wnd[1]/usr/ctxtRF02D-VTWEG").text = "00"
   session.findById("wnd[1]/usr/ctxtRF02D-VTWEG").setFocus
   session.findById("wnd[1]/usr/ctxtRF02D-VTWEG").caretPosition = 2
   session.findById("wnd[1]").sendVKey 0
   session.findById("wnd[0]/tbar[0]/btn[3]").press   
Loop While False: Next

   


