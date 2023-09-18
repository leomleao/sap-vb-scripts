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


strData = "6080173096,6080173087,6080173090,6080173099,6080173089,6080173095,6080173088,6080173086,6080173301,6080173093,6080173097,6080173091,6080173101,6080173302,6080173092,6080173094,6080173100,6080173102,6080173098"

arr = Split(strData,",")


For i = 0 To UBound(arr)

    session.findById("wnd[0]/usr/ctxtVBRK-VBELN").text = arr(i)
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/btnTC_OUTPUT").press
    session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-KSCHL[1,9]").text = "ZIID"
    session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/cmbNAST-NACHA[3,9]").key = "6"
    session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-PARVW[4,9]").text = "BP"
    session.findById("wnd[0]").sendVKey 0

    firstlineStatus = session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/lblDV70A-STATUSICON[0,0]").iconname
    firstlinePartner = session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-PARVW[4,0]").text

    if ((firstlineStatus = "S_TL_Y") And (firstlinePartner = "YE")) then
        session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-KSCHL[1,1]").setFocus
    End if
    
    session.findById("wnd[0]/tbar[1]/btn[5]").press
    session.findById("wnd[0]/usr/cmbNAST-VSZTP").key = "4" 
    session.findById("wnd[0]/tbar[0]/btn[3]").press

    if ((firstlineStatus = "S_TL_Y") And (firstlinePartner = "YE")) then
        session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-KSCHL[1,1]").setFocus
    End if

   '  session.findById("wnd[0]/tbar[1]/btn[2]").press 
   '  session.findById("wnd[0]/usr/chkNAST-DIMME").selected = true
   '  session.findById("wnd[0]/usr/ctxtNAST-LDEST").text = "LOCL"
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/tbar[0]/btn[11]").press 

Next




