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

prodOrdersStr = "4157986,4158199,4158201,4158204,4158235,4158236,4158238"
prodOrders = Split(prodOrdersStr,",")

session.findById("wnd[0]/tbar[0]/okcd"). Text = "/nCO02"
session.findById("wnd[0]").sendVKey 0

For i = 0 To UBound(prodOrders)
    session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = prodOrders(i)
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[5]").press
    If InStr(session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-ARBPL[4,1]").text,"QA") > 0 Then      
        session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-ARBPL[4,1]").text = "QA"
    End If 
    If InStr(session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-ARBPL[4,2]").text,"QA") > 0 Then      
        session.findById("wnd[0]/usr/tblSAPLCOVGTCTRL_0100/ctxtAFVGD-ARBPL[4,2]").text = "QA"
    End If 
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 11
Next



