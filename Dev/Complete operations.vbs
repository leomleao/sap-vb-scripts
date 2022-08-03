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

Wnd0.findById ("tbar[0]/okcd"). Text = "/nCO11N"
pressEnter ()

Dim arr(19,2)

arr(0,0) = 2647569
arr(0,1) = 3

arr(1,0) = 2647570
arr(1,1) = 3

arr(2,0) = 2647574
arr(2,1) = 3

arr(3,0) = 2647571
arr(3,1) = 3

arr(4,0) = 2647572
arr(4,1) = 3

arr(5,0) = 2647575
arr(5,1) = 3

arr(6,0) = 2647576
arr(6,1) = 3

arr(7,0) = 2647577
arr(7,1) = 3

arr(8,0) = 2647578
arr(8,1) = 3

arr(9,0) = 2647579
arr(9,1) = 3

arr(10,0) = 2647580
arr(10,1) = 3

arr(11,0) = 2647581
arr(11,1) = 3

arr(12,0) = 2647582
arr(12,1) = 3

arr(13,0) = 2647583
arr(13,1) = 3

arr(14,0) = 2647584
arr(14,1) = 3

arr(15,0) = 2647585
arr(15,1) = 3

arr(16,0) = 2647586
arr(16,1) = 3

arr(17,0) = 2647591
arr(17,1) = 3

arr(18,0) = 2647592
arr(18,1) = 2

arr(19,0) = 2647593
arr(19,1) = 2



For i = 2 To UBound(arr)
' For i = 0 To 1
' session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").text = "261"
' session.findById("wnd[0]/usr/ctxtRM07M-WERKS").text = "0008"
' session.findById("wnd[0]/usr/ctxtRM07M-LGORT").text = "0001"

If arr(i,1) = 3 Then

session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:0113/ctxtAFRUD-AUFNR").text = arr(i,0)
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:0113/ctxtAFRUD-VORNR").text = "0005"
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_DET1:SAPLCORU_S:0420/ctxtAFRUD-PERNR").text = "01001109"
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_DET1:SAPLCORU_S:0420/ctxtAFRUD-PERNR").setFocus
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_DET1:SAPLCORU_S:0420/ctxtAFRUD-PERNR").caretPosition = 8
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 11


session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:0113/ctxtAFRUD-AUFNR").text = arr(i,0)
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:0113/ctxtAFRUD-VORNR").text = "0010"
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_DET1:SAPLCORU_S:0420/ctxtAFRUD-PERNR").text = "01000115"
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_DET1:SAPLCORU_S:0420/ctxtAFRUD-PERNR").setFocus
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_DET1:SAPLCORU_S:0420/ctxtAFRUD-PERNR").caretPosition = 8
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 11

session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:0113/ctxtAFRUD-AUFNR").text = arr(i,0)
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:0113/ctxtAFRUD-VORNR").text = "0015"
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_DET1:SAPLCORU_S:0420/ctxtAFRUD-PERNR").text = "01000275"
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_DET1:SAPLCORU_S:0420/ctxtAFRUD-PERNR").setFocus
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_DET1:SAPLCORU_S:0420/ctxtAFRUD-PERNR").caretPosition = 8
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 11

Else

session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:0113/ctxtAFRUD-AUFNR").text = arr(i,0)
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:0113/ctxtAFRUD-VORNR").text = "0010"
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_DET1:SAPLCORU_S:0420/ctxtAFRUD-PERNR").text = "01000115"
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_DET1:SAPLCORU_S:0420/ctxtAFRUD-PERNR").setFocus
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_DET1:SAPLCORU_S:0420/ctxtAFRUD-PERNR").caretPosition = 8
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 11

session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:0113/ctxtAFRUD-AUFNR").text = arr(i,0)
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:0113/ctxtAFRUD-VORNR").text = "0015"
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_DET1:SAPLCORU_S:0420/ctxtAFRUD-PERNR").text = "01000275"
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_DET1:SAPLCORU_S:0420/ctxtAFRUD-PERNR").setFocus
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_DET1:SAPLCORU_S:0420/ctxtAFRUD-PERNR").caretPosition = 8
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 11

End If 
        

Next




