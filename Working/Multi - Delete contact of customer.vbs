Set objArgs = WScript.Arguments
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(CInt(objArgs(0)))
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

pFNumber = "wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtRF02D-KTONR[2,0]"

customerStr = objArgs(1)

customers = Split(customerStr,",")

For i = 0 To UBound(customers)
   If IsNumeric(customers(i)) = True Then
      session.findById("wnd[0]/tbar[0]/okcd"). Text = "/nVD02"
      session.findById("wnd[0]").sendVKey 0
      session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = customers(i)
      session.findById("wnd[1]/usr/ctxtRF02D-VKORG").text = "0008"
      session.findById("wnd[1]/usr/ctxtRF02D-VTWEG").text = "00"
      session.findById("wnd[1]/usr/ctxtRF02D-SPART").text = "00"
      session.findById("wnd[1]").sendVKey 0

      If session.ActiveWindow.Name = "wnd[2]" Then
         session.findById("wnd[2]/tbar[0]/btn[0]").press
      End if
      
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
      

      If InStr(session.findById("wnd[0]").text,"General Data") = 0 Then
      session.findById("wnd[0]/tbar[1]/btn[25]").press  
      End If
      session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB07").select 

      If session.ActiveWindow.Name = "wnd[1]" Then
         session.findById("wnd[1]/tbar[0]/btn[0]").press
      End if

      If InStr(session.findById("wnd[0]/sbar").text,"marked for deletion") > 0 Then
         session.findById("wnd[0]").sendVKey 0
      End if

      If findByFunction("DELETE") Then 
         session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB07/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7360/btnDELETE_ROW").press
         session.findById("wnd[0]").sendVKey 11        
      End If
    End If 
Next

Function findByFunction (name)
   session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB07/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7360/txt*KNVK-NAME1").text = name
   session.findById("wnd[0]").sendVKey 0
   If session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB07/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7360/tblSAPMF02DTCTRL_ANSPRECHPARTNER/txtKNVK-NAMEV[2,0]").text = name Then
      findByFunction = 1
   Else
      findByFunction = 0
   End If
End Function




