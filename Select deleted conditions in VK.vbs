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


strData = "7388472,7390228,7393127,7421908,7430550"

arr = Split(strData,",")

For i = 0 To UBound(arr) : Do

   session.findById("wnd[0]/usr/ctxtF003").text = arr(i)
   session.findById("wnd[0]/tbar[1]/btn[8]").press

   'How many visible rows are appearing on screen
   visibleRowCount = session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY").visiblerowcount
   'Total of rows in the table
   rowCount = session.findById("wnd[0]/usr/txtRV130-TABCA").text
   'How many iterations needs to be done to process every single line
   iterations = rowCount \ visibleRowCount

   For y = 0 To iterations
      For j = 0 To visibleRowCount - 1
         If session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/chkKONP-LOEVM_KO[12," & j &"]").selected Then
            session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY").getAbsoluteRow(j + y * (visibleRowCount - 1)).selected = true
         End If
      Next 
      'Press page down button
      session.findById("wnd[0]/tbar[0]/btn[82]").press
   Next

   session.findById("wnd[0]/usr/btnFCODE_TICH").press
   session.findById("wnd[1]/usr/ctxtRV13A-DATBI").text = "31.12.2021"
   'Press ok button to pop up
   session.findById("wnd[1]/tbar[0]/btn[0]").press
   'Press save button
   session.findById("wnd[0]/tbar[0]/btn[11]").press

Loop While False: Next






