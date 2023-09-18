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

matFromStr =  "60152510,60272370,60374741"
matToStr =  "60272369,60374740,99701489"


matFrom = Split(matFromStr,",")
matTo = Split(matToStr,",")


For i = 0 To UBound(matFrom)

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nVK12"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRV13A-KSCHL").text = "ZPR0"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[0,0]").select
    session.findById("wnd[1]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtF001").text = "46"
    session.findById("wnd[0]/usr/ctxtF002").text = "202301"
    session.findById("wnd[0]/usr/ctxtF003-LOW").text = matFrom(i)
    session.findById("wnd[0]/usr/ctxtF003-HIGH").text = matTo(i)
    session.findById("wnd[0]/tbar[1]/btn[8]").press

    session.findById("wnd[0]/usr/btnFCODE_MARL").press
    session.findById("wnd[0]/usr/btnFCODE_TICH").press
    session.findById("wnd[1]/usr/ctxtRV13A-DATBI").text = "19.06.2023"
    session.findById("wnd[1]").sendVKey 0
    session.findById("wnd[0]").sendVKey 11

Next 