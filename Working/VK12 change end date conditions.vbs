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

customerStr =  "3128938,3128938,3128938,3128938,3128938,3128938,3128938,3128938,3128938,3128938,3128938,807138,807138,807138,3124034,3124034,3124034,3124034,3124034,810537,807423,7326012"
materialStr =  "2790626,2790682,2800650,2800653,2820682,2830672,7590302,7810601,7810604,7810651,7820601,60027404,60220970,60220980,50033602,51169644,60047884,60194879,60194914,60027090,60051942,60362957"


customers = Split(customerStr,",")
materials = Split(materialStr,",")

session.findById("wnd[0]/tbar[0]/okcd").text = "/nVK12"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRV13A-KSCHL").text = "ZRA1"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[2,0]").select
session.findById("wnd[1]").sendVKey 0

For i = 0 To UBound(customers)

    session.findById("wnd[0]/usr/ctxtSEL_DATE").text = "15.08.2022"
    session.findById("wnd[0]/usr/ctxtF001").text = "1901"
    session.findById("wnd[0]/usr/ctxtF002").text = "00"
    session.findById("wnd[0]/usr/ctxtF003").text = customers(i)
    session.findById("wnd[0]/usr/ctxtF004-LOW").text = materials(i)
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    If InStr(session.findById("wnd[0]/sbar/pane[0]").text,"No condition records exist for this selection") > 0 Then
        session.findById("wnd[0]/tbar[0]/btn[3]").press  
        session.findById("wnd[0]").sendVKey 0
    Else
        session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/ctxtRV13A-DATBI[11,0]").text = "31.08.2022"
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 11 
        If InStr(session.findById("wnd[0]/sbar/pane[0]").text,"Saving not necessary. No changes were made") > 0 Then
            session.findById("wnd[0]/tbar[0]/btn[3]").press  
            session.findById("wnd[0]").sendVKey 0
        End If
    End If
Next 