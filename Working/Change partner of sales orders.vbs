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

ordersStr = "15967654,15957291,15957284,15957273,15957263,15957257,15957253,15957238,15957197,15953612,15947796,15947795,15947710,15947791,15941093,15941090,15941082,15941078,15941075,15941069,15941065,15941063,15941055,15937452,15937426,15937423,15937419,15937414,15937402,15937315,15937301,15937257,15937224,15931983,15931977,15931960,15931945,15929162,15929148,15929155,15908088,15867604,15854530,15854520,15850708,15850698,15850318,15843229,15843221,15843213,15843208,15843199,15843188,15843145,15843138,15818637,15818634,15818626,15816149,15799598,15799597,15799589,15799581,15799577,15799568,15799495,15799483,15789035,15789034,15773988,15773956,15773926,15773920,15773590,15773581,15755323,15749230,15746146,15744355,15744345,15744115,15744103,15734676,15734675,15734673,15733672,15733670,15733662,15733644,15731040,15731038,15731033,15731029,15730863,15730856,15673650,15654503,15654506,15613798,15613756,15613760"
partnerType = "SP"
newPartner = "22002077" 

partnerTable = "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW"

orders = Split(ordersStr,",")
For i = 0 To UBound(orders)
    session.findById("wnd[0]/tbar[0]/okcd"). Text = "/nVA02"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = orders(i)
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press

    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").select
    For k = 0 To session.findById(partnerTable).visibleRowCount    
        partner = session.findById(partnerTable & "/cmbGVS_TC_DATA-REC-PARVW[0," & k & "]").text
        If Left(partner,2) = partnerType Then
            session.findById(partnerTable & "/ctxtGVS_TC_DATA-REC-PARTNER[1," & k & "]").text = newPartner  
        Exit For
        End If
    Next
    session.findById("wnd[0]").sendVKey 11
Next