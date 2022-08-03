'-Begin-----------------------------------------------------------------
'-
'- Scripting to fill up the delivery picking tab with the Dlv. Qty
'- and fill the batch number as well.
'-
'- Author: Leonardo Leao
'-
'-----------------------------------------------------------------------

'-Directives------------------------------------------------------------
  Option Explicit

'-Sub Action------------------------------------------------------------
'-
'- Get the selected session and do the action inside it
'-
'-----------------------------------------------------------------------
Sub Action(session)
   session.findById("wnd[0]/tbar[0]/okcd").text = "/nCOOIS"
   session.findById("wnd[0]").sendVKey 0

   session.findById("wnd[0]/tbar[1]/btn[17]").press 'Get variant
   session.findById("wnd[1]/usr/txtV-LOW").text = "/UK_NO_BOOK" 'Get variant
   session.findById("wnd[1]/usr/txtENAME-LOW").text = "" 'Erase user name if there
   session.findById("wnd[1]/tbar[0]/btn[8]").press 'Execute to find variant

   session.findById("wnd[0]/tbar[1]/btn[8]").press 'Execute report

   If Session.ActiveWindow.Name = "wnd[1]" Then
      MsgBox "Everything is clear, goob job!"
   Else

      Dim count, i, j
      count = session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").RowCount
      Dim arr()
      Redim arr(count - 1)
      For i = 0 To (count - 1) : Do
         arr(i) = session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").GetCellValue(i, "AUFNR")
         
         
      Loop While False: Next

      session.findById("wnd[0]/tbar[0]/okcd").text = "/nMIGO"
      session.findById("wnd[0]").sendVKey 0

      session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0011/cmbGODYNPRO-ACTION").key = "A01"
      session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0011/cmbGODYNPRO-REFDOC").key = "R08"


      For j = 0 To UBound(arr)
      'Add all prod orders.
      session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2070/ctxtGODYNPRO-ORDER_NUMBER").text = arr(j)
      session.findById("wnd[0]").sendVKey 0
      If (j > 0) Then
      'Press Item OK on all the item except 0 (button not changeable)
      session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[3," & j & "]").selected = true
      End If
      Next
      'Press Item OK on the item 0
      session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/subSUB_DETAIL_TAKE:SAPLMIGO:0304/chkGODYNPRO-DETAIL_TAKE").selected = true
      MsgBox "If everything is ok, please save the document now!"

   End if   

End Sub

'-Sub Main--------------------------------------------------------------
'-
'- Main procedure to select the session
'-
'-----------------------------------------------------------------------
Sub Main()

  '-Local variables-----------------------------------------------------
  Dim SapAppl, SapGuiAuto, CollCon, i, oCon, CollSes, j, oSes, unExe
  Dim oSesInf, SID, Trans

  unExe = True

  Set SapGuiAuto = GetObject("SAPGUI")
  If Not IsObject(SapGuiAuto) Then
    Exit Sub
  End If

  Set SapAppl = SapGuiAuto.GetScriptingEngine
  If Not IsObject(SapAppl) Then
    Exit Sub
  End If

  Set CollCon = SapAppl.Connections()
  If Not IsObject(CollCon) Then
    Exit Sub
  End If

  '-Loop over connections-----------------------------------------------
  For i = 0 To CollCon.Count() - 1

    Set oCon = SapAppl.Children(CLng(i))
    If Not IsObject(oCon) Then
      Exit Sub
    End If

    Set CollSes = oCon.Sessions()
    If Not IsObject(CollSes) Then
      Exit Sub
    End If

    '-Loop over sessions------------------------------------------------
    For j = 0 To CollSes.Count() - 1

      Set oSes = oCon.Children(CLng(j))
      If Not IsObject(oSes) Then
        Exit Sub
      End If 

      If oSes.Busy() = vbFalse Then

        Set oSesInf = oSes.Info()
        If IsObject(oSesInf) Then

          '-------------------------------------------------------------
          '-
          '- The action should only be executed in the active session in VL02N
          '-
          '-------------------------------------------------------------
            SID = oSesInf.SystemName()
            Trans = oSesInf.Transaction()             

            If SID = "WGP" And unExe Then
                Action oSes  
                unExe = false
            End If

        End If

      End If

    Next

  Next

End Sub

'-Main------------------------------------------------------------------
Main()

'-End-------------------------------------------------------------------





