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

customerStr =  "801934,802067,802562,802602,803136,803810,803977,805033,806641,806686,806708,807161,807911,807947,807984,808030,808133,808303,808332,808886,809023,810037,810453,810454,813183,7311234,7332113,800438"

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
         If session.ActiveWindow.Name = "wnd[1]" Then
         session.findById("wnd[1]/tbar[0]/btn[0]").press
         End if
         session.findById("wnd[0]").sendVKey 0     
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




