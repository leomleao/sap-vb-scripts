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


strData = "6080000623,6080001414,6080001415,6080001416,6080001417,6080001418,6080001419,6080001420,6080001421,6080001422,6080001423,6080001424,6080001425,6080001701,6080001702,6080001703,6080001704,6080003585,6080004052,6080004701,6080005079,6080005230,6080006230,6080006403,6080006576,6080006577,6080007343,6080007436,6080007548,6080008440,6080008441,6080008589,6080008812,6080009704,6080009705,6080009706,6080010151,6080011575,6080012133,6080012134,6080013579,6080013787,6080013788,6080013961,6080013962,6080014149,6080015135,6080015136,6080015287,6080015692"

arr = Split(strData,",")


For i = 0 To UBound(arr)

    session.findById("wnd[0]/usr/ctxtVBRK-VBELN").text = arr(i)
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/btnTC_OUTPUT").press
    session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-KSCHL[1,9]").text = "ZIGB"
    session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/cmbNAST-NACHA[3,9]").key = "1"
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

    session.findById("wnd[0]/tbar[1]/btn[2]").press 
    session.findById("wnd[0]/usr/chkNAST-DIMME").selected = true
    session.findById("wnd[0]/usr/ctxtNAST-LDEST").text = "8037"
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/tbar[0]/btn[11]").press 

Next




