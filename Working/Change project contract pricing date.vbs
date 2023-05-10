
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


'' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' ''
'Supporting procedures and functions
'' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' ''

session.findById("wnd[0]/tbar[0]/okcd"). Text = "/nVA42"
session.findById("wnd[0]").sendVKey 0

strData = "13487681,13487682,13487683,13487685,13487686,13487687,13487688,13487689,13487690,13487691,13487692,13487693,13487694,13487696,13487697,13487698,13487699,13487700,13487701,13487702,13487703,13487704,13487705,13487706,13487707,13487708,13487709,13487710,13487711,13487712,13487713,13487714,13487715,13487716,13487717,13487718,13487719,13487720,13487721,13487723,13487724,13487725,13487726,13487727,13487728,13487729,13487730,13487731,13487732,13487733,13487734,13487735,13487736,13487737,13487738,13487739,13487740,13487741,13487742,13487743,13487744,13487745,13487746,13487747,13487748,13487749,13487751,13487752,13487753,13487754,13487755,13487756,13487757,13487758,13487759,13487760,13500021,13532007,13605495,13636570,13691048,13703910,13985599,14189015,14189326,14189328,14189332,14189333,14189397,14189398,14189400,14196971,14233790,14297276,14297285,14357227,14386097,14389370,14612932,14720859,14767190,14777846,14778094,14782228,14972557,15097739,15152489,15294207,15331574,15365296,15384711,15451014,15499501,15517726,15550742,15550774,15566323,15570519,15581499,15591740,15591889,15591954,15612078,15612970,15634655,15642443,15652350,15660710,15665077,15700660,15708029,15726915,15726949,15730046,15737963,15738007,15750838,15754522,15770383,15803018,15807258,15818051,15832044,15842397,15852710,15888074,15922512,15953312,15992736,16009483,16042482,16068373,16136330,16153391,16261287,16292497,16361039,16361439,16362049,16366881,16399439,16405799,16424152,16447051"
pricingDate = "02.03.2023"
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