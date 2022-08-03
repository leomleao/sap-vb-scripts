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

'Pressing the F3 button
Sub pressF3 ()
Wnd0.sendVKey 3
End Sub

Wnd0.findById ("tbar[0]/okcd"). Text = "/nVF02"
pressEnter ()

Dim arr(21, 1)

arr(0,0) = 6080001510
arr(0,1) = false
arr(1,0) = 6080001516
arr(1,1) = false
arr(2,0) = 6080001884
arr(2,1) = false
arr(3,0) = 6080002411
arr(3,1) = false
arr(4,0) = 6080002668
arr(4,1) = false
arr(5,0) = 6080002917
arr(5,1) = true
arr(6,0) = 6080003857
arr(6,1) = true
arr(7,0) = 6080003859
arr(7,1) = true
arr(8,0) = 6080004135
arr(8,1) = true
arr(9,0) = 6080004789
arr(9,1) = true
arr(10,0) = 6080001061
arr(10,1) = true
arr(11,0) = 6080001716
arr(11,1) = true
arr(12,0) = 6080002127
arr(12,1) = true
arr(13,0) = 6080002322
arr(13,1) = true
arr(14,0) = 6080005188
arr(14,1) = true
arr(15,0) = 6080005452
arr(15,1) = true
arr(16,0) = 6080001189
arr(16,1) = true
arr(17,0) = 6080002932
arr(17,1) = true
arr(18,0) = 6080004798
arr(18,1) = true
arr(19,0) = 6080003147
arr(19,1) = true
arr(20,0) = 6080003148
arr(20,1) = true
arr(21,0) = 6080003754
arr(21,1) = true

For i = 0 To UBound(arr)

if arr(i, 1) then

    session.findById("wnd[0]/usr/ctxtVBRK-VBELN").text = arr(i, 0)
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/btnTC_OUTPUT").press
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[5]").press
    session.findById("wnd[0]/usr/cmbNAST-VSZTP").key = "4"
    session.findById("wnd[0]/tbar[0]/btn[3]").press    
    session.findById("wnd[0]/tbar[0]/btn[11]").press
End if

Next




