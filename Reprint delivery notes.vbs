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

'Create an object of type GuiMainWindow
Set Wnd0 = session.findById ("wnd[0]")

'Create an object of type GuiMenubar
Set Menubar = Wnd0.findById ("mbar")

'Create an object of type GuiUserArea
Set UserArea = Wnd0.findById ("usr")

'Create an object of type GuiStatusbar
Set Statusbar = Wnd0.findById ("sbar")

'Define the user's login
UserName = session.Info.User

'' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' ''
'Supporting procedures and functions
'' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' ''
' Pressing the "Enter"
Sub pressEnter ()
Wnd0.sendVKey 0
End Sub
Dim sInput
sInput = InputBox("Which printer?")


' Pressing the "Enter"
Sub printDelivery ()    

    'Into delivery now
    session.findById("wnd[0]/mbar/menu[3]/menu[1]/menu[0]").select 'Header Delivery output

    ' 'Get first row
    session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3").getAbsoluteRow(0).selected = -1
    session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/lblDV70A-STATUSICON[0,0]").setFocus
    session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/lblDV70A-STATUSICON[0,0]").caretPosition = 0    
    session.findById("wnd[0]/tbar[1]/btn[6]").press 'Repeat output

    session.findById("wnd[0]/tbar[1]/btn[2]").press ' Communication method
    session.findById("wnd[0]/usr/ctxtNAST-LDEST").text = sInput
    session.findById("wnd[0]").sendVKey 0 'Enter

    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Go Back
    ' session.findById("wnd[0]/tbar[0]/btn[3]").press 'Go Back
    session.findById("wnd[0]/tbar[0]/btn[11]").press
End Sub

' session.findById("wnd[0]/tbar[1]/btn[5]").press 'Select All
    session.findById("wnd[0]/tbar[1]/btn[13]").press 'Change deliveries
    printDelivery ()

test = true
Do            
    If Session.ActiveWindow.Name = "wnd[1]" Then
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        printDelivery ()
    Else
    ' Dim sContinue
    ' sContinue = msgbox("Everything OK?" & Session.ActiveWindow.Name, vbYesNo,"Check")
    '     If sContinue = 7 Then
    '         Exit Do
    '     Else
    '         Exit Do
    '     End If
        test = false
    End if
Loop While test = true





