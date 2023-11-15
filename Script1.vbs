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
  
   
   strData = "4473252,4473275,4473279,4473304,4474727,4480708,4482343,4485912,4487465,4487466,4489308,4491707,4504656,4505979,4505980,4505983,4506077,4507406,4509074,4509164,4525181,4525248,4527110,4527112,4544952,4546653,4551233,4551235,4551236,4552373,4553177,4570959,4572950,4574472,4574479,4577857,4583804,4583865,4585213,4588032,4590027,4590028,4591532,4591587,4591588,4594069,4594081,4595642,4598578,4598622,4600646,4600647,4605585,4608727,4609989,4610168,4614400,4616514,4616516,4616517,4616549,4616551,4616552,4616655,4618381,4618382,4618383,4620255,4621786,4623038,4623061,4623082,4623937,4627798,4627858,4627887,4627918,4627919,4630349,4630350,4630356,4631573,4633127,4633129,4633130,4633191,4634534,4636046,4636048,4636050,4636102,4636170,4647131,4652820,4652821,4652823,4657565"
   productionOrders = Split(strData,",")


' ' Section to go on the POs, remove TECO and remove the Final issue Flag for each unpicked component line
'     For i = 0 To UBound(productionOrders) : Do

'       session.findById("wnd[0]/tbar[0]/okcd").text = "/nco02"
'       session.findById("wnd[0]").sendVKey 0
'       session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").text = productionOrders(i)
'       session.findById("wnd[0]").sendVKey 0

'       If Session.ActiveWindow.Name = "wnd[1]" Then
'          If InStr(session.findById("wnd[1]/usr/txtMESSTXT1").text,"""Change"" is not allowed") > 0 Then
'             session.findById("wnd[1]/tbar[0]/btn[0]").press
'             session.findById("wnd[0]/mbar/menu[1]/menu[7]/menu[4]").select
'          End if
'       End if
'       session.findById("wnd[0]/tbar[1]/btn[6]").press
'       visibleRowCount = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120").visiblerowcount
'       'Total of rows in the table
'       rowCount = session.findById("wnd[0]/usr/txtRC27X-ENTRIES").text
'       'How many iterations needs to be done to process every single line
'       iterations = rowCount \ visibleRowCount

'       For y = 0 To iterations
'          For q = 0 To visibleRowCount - 1
'             qtyReq = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-MENGE[3," & q &"]").text
'             qtyWithdrawn = session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/txtRESBD-DENMNG[4," & q &"]").text
'             If qtyWithdrawn < qtyReq Then
'                If Not session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/chkRESBD-XLOEK[17," & q &"]").selected Then
'                   session.findById("wnd[0]/usr/tblSAPLCOMKTCTRL_0120/chkRESBD-KZEAR[23," & q &"]").selected = false
'                End If
'             End If
'          Next 
'          'Press page down button (Not tcode specific)
'          session.findById("wnd[0]/tbar[0]/btn[82]").press
'       Next     

'       'Save
'       session.findById("wnd[0]/tbar[0]/btn[11]").press   
'     Loop While False: Next
'     ' END Section to go on the POs, remove TECO and remove the Final issue Flag for each unpicked component line

    ' Picking in CO27
    For u = 0 To UBound(productionOrders) : Do
         session.findById("wnd[0]/tbar[0]/okcd").text = "/nco27"
         session.findById("wnd[0]").sendVKey 0
         session.findById("wnd[0]/usr/ctxtS_AUFNR-LOW").text = productionOrders(u)
         session.findById("wnd[0]/tbar[1]/btn[8]").press
         If Session.ActiveWindow.Name = "wnd[1]" Then
            session.findById("wnd[1]/tbar[0]/btn[0]").press
         Else
            If InStr(session.findById("wnd[0]/sbar").text,"The pick list chosen contains no entries") > 0 Then
               session.findById("wnd[0]/tbar[0]/btn[3]").press  
               session.findById("wnd[0]").sendVKey 0
            Else 
               session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectAll            
               session.findById("wnd[0]/tbar[1]/btn[8]").press
               session.findById("wnd[0]/usr/subPUSHBUTTON:SAPLCOWB:0400/btnMALL").press
               session.findById("wnd[0]/usr/subPUSHBUTTON:SAPLCOWB:0400/btnCHFI").press
               session.findById("wnd[0]/tbar[0]/btn[11]").press
            End if
         End if
    Loop While False: Next
    ' END Picking in CO27