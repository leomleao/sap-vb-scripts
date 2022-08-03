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

pFNumber = "wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtRF02D-KTONR[2,0]"

customerStr =  "807803,811641,7380943,7322387,7381142,7321642"
partnerTypeStr =  "YA,YA,YA,YA,YA,YA"
newPartnerStr =  "1001583,1001583,1001583,1001583,1001583,1001583"

customers = Split(customerStr,",")
partnerType = Split(partnerTypeStr,",")
newPartner = Split(newPartnerStr,",")

For i = 0 To UBound(customers)
    session.findById("wnd[0]/tbar[0]/okcd"). Text = "/nVD02"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = customers(i)
    session.findById("wnd[1]/usr/ctxtRF02D-VKORG").text = "0008"
    session.findById("wnd[1]/usr/ctxtRF02D-VTWEG").text = "00"
    session.findById("wnd[1]/usr/ctxtRF02D-SPART").text = "00"
    session.findById("wnd[1]").sendVKey 0

    If InStr(session.findById("wnd[0]").text,"Sales Area Data") = 0 Then
    session.findById("wnd[0]/tbar[1]/btn[27]").press  
    End If
    session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05").select 
    session.findById("wnd[0]").sendVKey 0

    If session.ActiveWindow.Name = "wnd[1]" Then
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End if

    If findByFunction(partnerType(i)) Then 
        session.findById(pFNumber).text = newPartner(i)
        session.findById("wnd[0]").sendVKey 11        
    End If
Next

Function findByFunction (partnerFunction)
   session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/ctxt*KNVP-PARVW").text = partnerFunction
   session.findById("wnd[0]").sendVKey 0
   If session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,0]").text = partnerFunction Then
      findByFunction = 1
   Else
      findByFunction = 0
   End If
End Function




