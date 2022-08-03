
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

session.findById("wnd[0]/tbar[0]/okcd"). Text = "/nVA42"
pressEnter ()

strData = "13500021,13532007,13605495,13636570,13690733,13691048,13703910,13985599,14189015,14189326,14189332,14189333,14189397,14189398,14189400,14196971,14233790,14297276,14297285,14357227,14389370,14720859,14767190,14777846,14778094,14972557,15097739,15152489,15294207,15331574,15365296,15384711,15451014,15499501,15517726,15550742,15550774,15566323,15570519,15581499,15591740,15591889,15591954,15612078,15612970,15634655,15642443,15652350,15660710,15665077,15700660,15708029,15726915,15726949,15730041,15730046,15737963,15738007,15750838,15754522,15770383,15803018,15807258,15818051,15832044,15842397,15852710,15888074"
pricingDate = "01.07.2022"
arr = Split(strData,",")

For i = 0 To UBound(arr) : Do   
    session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = arr(i)
    session.findById("wnd[0]").sendVKey 0
    If session.children.count > 1 Then
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If 
    session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4022/btnBT_KKAU").press
    session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBKD-PRSDT").text = pricingDate
    session.findById("wnd[0]").sendVKey 0    
    session.findById("wnd[0]").sendVKey 0
    If session.children.count > 1 Then
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End If
    session.findById("wnd[0]").sendVKey 11  
Loop While False: Next