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

ordersStr = "16147512,16257499,16331587,16339654,16340895,16366974,16371558,16375126,16377090,16390917,16397635,16401733,16408980,16413380,16439836,16463254,16467315,16481349,16484765,16488314,16498793,16498845,16499473,16502287,16502428,16502481,16502486,16502632,16503481,16503506,16503644,16505772,16505784"
partnerType = "RG"
newPartner = "800403" 

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
            session.findById(partnerTable & "/ctxtGVS_TC_DATA-REC-PARTNER_EXT[1," & k & "]").text = newPartner              
            session.findById("wnd[0]").sendVKey 11
            If session.ActiveWindow.Name = "wnd[1]" Then
               session.findById("wnd[1]/tbar[0]/btn[0]").press
            End if
            Exit for
        End If
   Next
Next