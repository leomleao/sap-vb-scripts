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

    Dim visibleRowCount, rowCount, iterations, currentItem, qty, y, k

    'Check if button delete is changeable (it means we're inside the delivery)
    If session.findById("wnd[0]/tbar[1]/btn[17]").changeable Then
        'Open tab picking
        session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").select
        'How many visible rows are appearing on screen
        visibleRowCount = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK").visiblerowcount
        'Total of rows in the table
        rowCount = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK").rowCount
        'How many iterations needs to be done to process every single line
        iterations = rowCount \ visibleRowCount

        For y = 0 To iterations
            For k = 0 To (visibleRowCount - 1)
            'Check if the item line is filled
            currentItem = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPS-POSNR[0," & k & "]").text
            if currentItem <> "" Then
                qty = session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-G_LFIMG[4," & k & "]").text
                If session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-PIKMG[6," & k & "]").changeable Then
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/txtLIPSD-PIKMG[6," & k & "]").text = qty
                End If
                If session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/ctxtLIPS-CHARG[8," & k & "]").changeable Then
                    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1104/tblSAPMV50ATC_LIPS_PICK/ctxtLIPS-CHARG[8," & k & "]").text = "0000999908"
                End If
            End if
            Next
            'Press page down button
            session.findById("wnd[0]/tbar[0]/btn[82]").press
        Next
        'Save it
        session.findById("wnd[0]/tbar[0]/btn[11]").press

    End If

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

            If SID = "WGQ" And Trans = "VL02N" And unExe Then
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


