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

session.findById("wnd[0]/tbar[0]/okcd"). Text = "/nVDH1N"
session.findById("wnd[0]").sendVKey 0

topLevel = "92836"
strData =  "7305687,811420,7374722,7313255,7418877,7421893"

currentDate = Day(Now) & "." & Month(Now) & "." & Year(Now)
startDate = "01.01.2020"


session.findById("wnd[0]/usr/ctxtS_DATE").text = currentDate
session.findById("wnd[0]/usr/ctxtS_HITYP").text = "A"
session.findById("wnd[0]/usr/ctxtS_KUNNR-LOW").text = topLevel
session.findById("wnd[0]/usr/ctxtS_DATE").setFocus
session.findById("wnd[0]/usr/ctxtS_DATE").caretPosition = 2
session.findById("wnd[0]/tbar[1]/btn[8]").press


arr = Split(strData,",")

For i = 0 To UBound(arr)
   session.findById("wnd[0]/tbar[1]/btn[27]").press
   session.findById("wnd[0]/usr/ctxtKNVH-HKUNNR").text = topLevel
   session.findById("wnd[0]/usr/ctxtKNVH-HVKORG").text = "0008"
   session.findById("wnd[0]/usr/ctxtKNVH-HVTWEG").text = "00"
   session.findById("wnd[0]/usr/ctxtKNVH-HSPART").text = "00"
   session.findById("wnd[0]/usr/ctxtKNVH-VKORG").text = "0008"
   session.findById("wnd[0]/usr/ctxtKNVH-VTWEG").text = "00"
   session.findById("wnd[0]/usr/ctxtKNVH-SPART").text = "00"
   session.findById("wnd[0]/usr/ctxtKNVH-KUNNR").text = arr(i)
   session.findById("wnd[0]/usr/ctxtKNVH-DATAB").text = startDate
   session.findById("wnd[0]/usr/btnDVF_GET_DETAIL_BUTTON").press
Next




