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

strData = "6080015947,6080015949,6080017796,6080018259,6080018260,6080019083,6080020480,6080020482,6080020484,6080020486,6080020489,6080020496,6080020497,6080021421,6080021586,6080021661,6080021663,6080022131,6080022132,6080022429,6080022535,6080022537,6080022785,6080023298,6080023299,6080023367,6080023368,6080023738,6080023740"
' strData = "6015503517"

arr = Split(strData,",")

For i = 0 To UBound(arr)

session.findById("wnd[0]/usr/ctxtVBRK-VBELN").text = arr(i)
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/btnTC_OUTPUT").press
' session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3").columns.elementAt(1).width = 6
' session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-KSCHL[1,10]").text = "ZIDE"
' session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/cmbNAST-NACHA[3,10]").key = "5"
' session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-PARVW[4,10]").text = "YE"
' session.findById("wnd[0]").sendVKey 0
' session.findById("wnd[0]/tbar[1]/btn[5]").press
' session.findById("wnd[0]/usr/cmbNAST-VSZTP").key = "4"
' session.findById("wnd[0]/tbar[0]/btn[3]").press
' session.findById("wnd[0]/tbar[1]/btn[2]").press
' session.findById("wnd[0]/usr/ctxtNAST-TCODE").text = "CS01"
' session.findById("wnd[0]").sendVKey 0
' session.findById("wnd[0]/usr/chkNAST-DIMME").selected = true
' session.findById("wnd[0]/usr/ctxtNAST-LDEST").text = "LOCL"
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[11]").press

Next




