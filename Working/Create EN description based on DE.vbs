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




strData = "60443237,60443240,60446202,60449745,60479521,60479524,60479525,60479526,60479527,60489734,60489735,60489737,60489738,60489753,60489754,60489756,60491903,60491904,60492047,60492049,60492241,60493796"

arr = Split(strData,",")


For i = 0 To UBound(arr)

Wnd0.findById ("tbar[0]/okcd"). Text = "/nMM17"
pressEnter ()

   session.findById("wnd[0]/usr/tabsTAB/tabpTAB1/ssubSUBTAB:SAPMMSDL:1000/tblSAPMMSDLTC_TAB").getAbsoluteRow(1).selected = true  
   session.findById("wnd[0]/tbar[1]/btn[8]").press
   session.findById("wnd[0]/usr/tabsTAB/tabpNEW").select
   session.findById("wnd[0]/usr/tabsTAB/tabpNEW/ssubSUB_ALL:SAPLMASS_SEL_DIALOG:0400/subSUB_SEL:SAPLMASSFREESELECTIONS:1000/sub:SAPLMASSFREESELECTIONS:1000/ctxtMASSFREESEL-LOW[0,24]").text = arr(i)
   session.findById("wnd[0]/usr/tabsTAB/tabpNEW/ssubSUB_ALL:SAPLMASS_SEL_DIALOG:0400/subSUB_SEL:SAPLMASSFREESELECTIONS:1000/sub:SAPLMASSFREESELECTIONS:1000/ctxtMASSFREESEL-LOW[1,24]").text = "EN"
   session.findById("wnd[0]/usr/tabsTAB/tabpNEW/ssubSUB_ALL:SAPLMASS_SEL_DIALOG:0400/subSUB_PARA:SAPLMASSFREESELECTIONS:2000/sub:SAPLMASSFREESELECTIONS:2000/ctxtMASSFREESEL_P-LOW[0,22]").text = arr(i)
   session.findById("wnd[0]/usr/tabsTAB/tabpNEW/ssubSUB_ALL:SAPLMASS_SEL_DIALOG:0400/subSUB_PARA:SAPLMASSFREESELECTIONS:2000/sub:SAPLMASSFREESELECTIONS:2000/ctxtMASSFREESEL_P-LOW[1,22]").text = "DE"
   session.findById("wnd[0]/tbar[1]/btn[8]").press
   session.findById("wnd[0]/tbar[0]/btn[11]").press
   session.findById("wnd[1]/tbar[0]/btn[0]").press

Next




