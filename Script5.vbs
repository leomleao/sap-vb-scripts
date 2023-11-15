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
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectColumn "AUFNR"
session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectColumn "MATNR"
session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectColumn "POSNR"
session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectColumn "BDTER"
session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectColumn "MENGE"
session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectColumn "DENMNG"
session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectColumn "EINHEIT"
session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectColumn "MATXT"
session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectColumn "CHARG"
session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectColumn "PLNFL"
session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectColumn "VORNR"
session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectColumn "RSPOS"
session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectColumn "WERKS"
session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectColumn "LGORT"
session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectColumn "ICON_MSGTY"
session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectColumn "STTXT"
session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectAll