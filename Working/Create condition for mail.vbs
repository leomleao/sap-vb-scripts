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



'Pressing the F3 button
Sub pressF3 ()
Wnd0.sendVKey 3
End Sub

Wnd0.findById ("tbar[0]/okcd"). Text = "/nVV31"
session.findById("wnd[0]").sendVKey 0


Dim sInput
sInput = InputBox("Enter bill-to-party no:")

if Len(sInput) > 10 Then
    arr = Split(strData,",")

    For i = 0 To UBound(arr)

        ' if i = 0 Then
        '     session.findById("wnd[0]/usr/ctxtRV13B-KSCHL").text = "ZIGB"
        '     session.findById("wnd[0]").sendVKey 0
        '     session.findById("wnd[1]/tbar[0]/btn[0]").press
        '     session.findById("wnd[0]/usr/ctxtKOMB-VKORG").text = "0008"
        '     session.findById("wnd[0]/usr/ctxtKOMB-FKART").text = "Z210"
        ' End if

        ' session.findById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtKOMB-KUNRE[0,0]").text = sInput
        ' session.findById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtKOMB-KUNRE[0,0]").setFocus
        ' session.findById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtKOMB-KUNRE[0,0]").caretPosition = 6
        ' session.findById("wnd[0]").sendVKey 0
        ' session.findById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtNACH-SPRAS[6,0]").text = "EN"
        ' session.findById("wnd[0]/tbar[1]/btn[2]").press
        ' session.findById("wnd[0]/usr/chkNACH-DIMME").selected = true
        ' session.findById("wnd[0]/usr/ctxtNACH-LDEST").text = "8037"
        ' session.findById("wnd[0]/usr/txtNACH-ANZAL").text = "0"
        ' session.findById("wnd[0]/usr/txtNACH-DSNAM").text = "ZIGB"
        ' session.findById("wnd[0]/usr/txtNACH-DSUF1").text = "UK"
        ' session.findById("wnd[0]/usr/txtNACH-DSUF2").text = "VKORG0008"
        ' session.findById("wnd[0]/usr/txtNACH-DSUF2").setFocus
        ' session.findById("wnd[0]/usr/txtNACH-DSUF2").caretPosition = 9
        ' session.findById("wnd[0]/tbar[0]/btn[3]").press

        ' Dim sContinue
        ' sContinue = msgbox("Everything OK?", vbYesNo,"Check")
        ' If sContinue = 7 Then
        '     Wnd0.findById ("tbar[0]/okcd"). Text = "/n"
        '     session.findById("wnd[0]").sendVKey 0
        ' Else
        '     session.findById("wnd[0]/tbar[0]/btn[11]").press
        ' End If

    Next

Else
    session.findById("wnd[0]/usr/ctxtRV13B-KSCHL").text = "ZIGB"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/usr/ctxtKOMB-VKORG").text = "0008"
    session.findById("wnd[0]/usr/ctxtKOMB-FKART").text = "Z210"
    session.findById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtKOMB-KUNRE[0,0]").text = sInput
    session.findById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtKOMB-KUNRE[0,0]").setFocus
    session.findById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtKOMB-KUNRE[0,0]").caretPosition = 6
    session.findById("wnd[0]").sendVKey 0

    If session.findById("wnd[0]/sbar/pane[0]").text = "The condition record entered already exists" Then

    Dim sError
    sError = msgbox("Sorry, check the existing condition in VV32.", vbYesNo,"Check")
        If sError = 7 Then
            Wnd0.findById ("tbar[0]/okcd"). Text = "/n"
            session.findById("wnd[0]").sendVKey 0
        Else
            Wnd0.findById ("tbar[0]/okcd"). Text = "/nVV32"
            session.findById("wnd[0]").sendVKey 0
        End If    

    Else 
    session.findById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtNACH-SPRAS[6,0]").text = "EN"
    session.findById("wnd[0]/tbar[1]/btn[2]").press
    session.findById("wnd[0]/usr/chkNACH-DIMME").selected = true
    session.findById("wnd[0]/usr/ctxtNACH-LDEST").text = "8038"
    session.findById("wnd[0]/usr/txtNACH-ANZAL").text = "0"
    session.findById("wnd[0]/usr/txtNACH-DSNAM").text = "ZIGB"
    session.findById("wnd[0]/usr/txtNACH-DSUF1").text = "UK"
    session.findById("wnd[0]/usr/txtNACH-DSUF2").text = "VKORG0008"
    session.findById("wnd[0]/usr/txtNACH-DSUF2").setFocus
    session.findById("wnd[0]/usr/txtNACH-DSUF2").caretPosition = 9
    session.findById("wnd[0]/tbar[0]/btn[3]").press

    Dim sContinue
    sContinue = msgbox("Everything OK?", vbYesNo,"Check")
        If sContinue = 7 Then
            Wnd0.findById ("tbar[0]/okcd"). Text = "/n"
            session.findById("wnd[0]").sendVKey 0
        Else
            session.findById("wnd[0]/tbar[0]/btn[11]").press
        End If

    End If

End If






