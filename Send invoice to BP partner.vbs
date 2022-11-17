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


strData = "6080130273,6080130274,6080130275,6080130276,6080130277,6080130278,6080130279,6080130280,6080130281,6080130282,6080130283,6080130284,6080130285,6080130286,6080130287,6080130288,6080130289,6080130290,6080130291,6080130292,6080130293,6080130294,6080130295,6080130296,6080131165,6080131166,6080131167,6080131168,6080131169,6080131170,6080131171,6080131172,6080131173,6080131175,6080131176,6080131177,6080131178,6080131179,6080131180,6080131181,6080131182,6080131183,6080131184,6080131185,6080131186,6080131187,6080131188,6080131189,6080131190,6080131191,6080131192,6080131193,6080131194,6080131195,6080131196,6080131197,6080131198,6080131199,6080131200,6080131201,6080131202,6080131203,6080131204,6080131677,6080132020,6080132021,6080132022,6080132023,6080132024,6080132025,6080132026,6080132027,6080132028,6080132029,6080132030,6080132031,6080132032,6080132033,6080132034,6080132035,6080132036,6080132037,6080132038,6080132039,6080132040,6080132041,6080132042,6080132043,6080132044,6080132045,6080132046,6080132047,6080132048,6080132049,6080132050,6080132051,6080132052,6080132053,6080132054,6080132055,6080132056,6080132057,6080132058,6080132059,6080132060,6080132241"
' strData = "6080130273"

arr = Split(strData,",")


For i = 0 To UBound(arr)

    

    ' session.findById("wnd[0]/usr/ctxtVBRK-VBELN").text = arr(i)
    ' session.findById("wnd[0]").sendVKey 0
    ' session.findById("wnd[0]/usr/btnTC_OUTPUT").press
    ' session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-KSCHL[1,9]").text = "ZIGB"
    ' session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/cmbNAST-NACHA[3,9]").key = "5"
    ' session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-PARVW[4,9]").text = "BP"
    ' session.findById("wnd[0]").sendVKey 0
    ' session.findById("wnd[0]/tbar[1]/btn[5]").press
    ' session.findById("wnd[0]/usr/cmbNAST-VSZTP").key = "4" 
    ' session.findById("wnd[0]/tbar[0]/btn[3]").press
    ' session.findById("wnd[0]/tbar[1]/btn[2]").press    
    ' session.findById("wnd[0]/usr/ctxtNAST-TCODE").text = "CS01"
    ' session.findById("wnd[0]").sendVKey 0
    ' session.findById("wnd[0]/usr/chkNAST-DIMME").selected = true
    ' session.findById("wnd[0]/usr/ctxtNAST-LDEST").text = "LOCL"
    ' session.findById("wnd[0]/tbar[0]/btn[3]").press
    ' session.findById("wnd[0]/tbar[0]/btn[11]").press

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

   '  if ((firstlineStatus = "S_TL_Y") And (firstlinePartner = "YE")) then
   '      session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-KSCHL[1,1]").setFocus
   '  End if

   '  session.findById("wnd[0]/tbar[1]/btn[2]").press    
   '  session.findById("wnd[0]/usr/ctxtNAST-TCODE").text = "CS01"
   '  session.findById("wnd[0]").sendVKey 0
   '  session.findById("wnd[0]/usr/chkNAST-DIMME").selected = true
   '  session.findById("wnd[0]/usr/ctxtNAST-LDEST").text = "LOCL"
   '  session.findById("wnd[0]/tbar[0]/btn[3]").press


    session.findById("wnd[0]/tbar[0]/btn[11]").press 

Next




