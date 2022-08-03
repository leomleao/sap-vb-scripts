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


strData = "6080073087,6080073081,6080073084,6080073069,6080073059,6080073078,6080073067,6080073071,6080073062,6080073063,6080073060,6080073089,6080073086,6080073076,6080073077,6080073094,6080073079,6080073092,6080073055,6080073085,6080073091,6080073054,6080073080,6080073072,6080073068,6080073090,6080073074,6080073061,6080073058,6080073083,6080073082,6080073064,6080073070,6080073073,6080073056,6080073057,6080073066,6080073075"

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




