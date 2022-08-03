
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

' Writing log
Sub WriteLog(strMessage)
	Const FOR_APPENDING = 8
 
	strFileName = "C:\Users\u081715\AppData\Roaming\SAP\SAP GUI\Scripts\log.txt"
 
	Set objFS = CreateObject("Scripting.FileSystemObject")
 
	If objFS.FileExists(strFileName) Then
		Set oFile = objFS.OpenTextFile(strFileName, FOR_APPENDING)
	Else
		Set oFile = objFS.CreateTextFile(strFileName)
	End If
 
	oFile.WriteLine strMessage
End Sub

Wnd0.findById("tbar[0]/okcd"). Text = "/nVA02"
pressEnter ()

strData = "15299303,14867920,14959783,15276277,14650283,15287673,15362238,15266985,15268810,15273542,15242914,14869969,15354301,14846950,14754772,15234538,15297682,15219640,15358093,15273812,15318590,15266172,14679474,14662474,15269139,15273554,14630554,15266300,15269191,15275604,15282598,15354354,15247746,15273030,14638647,15313391,15286380,15266742,15269408,15318750,14836070,15269305,15159598,15354334,15354310,15269438,15266359,15282553,15293715,15276020,15246408,14987952,15265915,15356194,15302709,15265858,15354300,15266782,15265889,15356982,15354149,15275937,15270046,15204336,15281862,15161822,14695297,15266243,14800428,15241958,15064006,15350255,15262298,15345456,15364976,15356999,15341489,15287657,15209126,15193635,14789008,15357918,15365965,15266345,14866662,15265893,15279383,15282283,15268849,15165912,15361293,15293526,14885204,14903679,15266555,15266765,15251814,15243822,15268699,15283830,15338193,15274017,15354246,15261990,15268985,15266294,15235172,15251850,15282681,15289981,15270657,15231516,15272937,14932798,15360465,15257465,15046997,14799246,15269761,14634508,15303256,15269769,14720890,15357595,15148531,15266364,15342150,15272904,15361499,15279334,15265921,15265057,15263263,15266967,15269415,15251743,15265881,14886805,15207740,15265807,15354296,15270674,15303575,15065865,15270148,15266711,15177040,15269015,15293552,15302184,15266279,15354313,15276338,15290297,14905941,15272379,15337452,14621306,15354123,15286697,15269355,14645916,15243386,14583005,15268827,15268917,15223102,15276081,15266148,15272984,15273312,15305737,15268886,15003379,15291141,15269474,15266952,14944262,15275941,14766366,15250762,15273192,15266787,15003210,15266821,15354250,15266313,15235062,15279381,15268652,14895642,15265902,15313440,15303755,15266870,15274133,15273836,14813698,15334301,15297834,15208148,14911971,15156595,15078646,15269126,15276144,14874325,15282727,15275952,15234071,15286127,15356862,15276076,15238796,15265854,15303462,15250795,15227252,15215957,14719864,15270688,14725267,15356851,15276077,15266108,15263299,15265984,15275852,15126821,15346166,14655943,15365097,15287383,14947207,14791073,15232078,15189721,15293966,15364985,15287624,15360417,15295321,15270690,15272867,15321482,15236169,15239278,15020401,14832074,15282570,15270099,15302285,15276280,15172453,15169612,15346560,14592277,15126906,15242617,14699568,14672841,15364099,15361021,15275626,15281905,14721084,15337991,15342901,14716881,15290973,15289878,14842442,15356340,15023495,14948420,15261880,15341690,15293604,15261704,15353295,15295267,15318703,15190839,15247745,15309626,15278894,15230963,15279720,15266049,15247819,15364561,15338793,15295317,15101556,15007287,15356627,15266738,15272607,15071265,15365523,15277770,15185043,15283141,14987689,15360374,15129329,15274086,15356678,15337196,15313987,15246790,15273851,14885963,14734008,15364431,14791304,15272914,15279396,15275872,14836139,15356430,15032581,14611635,14667826,14771489,15145837,15283831,15364808,15365575,15278408,15314638,15360993,15362149,15173008,15298222,15216386,15331586,14672259,15364352,14782754,15364867,15274060,15326646,15189491,15305976,15286777,15364528,15246696,15325883,15306254,14796919,15357544,15242465,15228721,15096895,15216381,15327033,15273193,15212433,15364881,15321952,14625999,15155737,15342297,15279173,15326532,14637575,15278365,14790069,15287619,15349185,15309680,15306113,15293983,15060461,15281781,15365369,15356732,15365271,14992540,14676170,15278334,15270678,15290950,15250607,15250586,15275892,15239562,15278639,15294014,15334235,15278620,15270453,15008833,15279034,15278835,15357524,14750224,15278346,15238639,15253145,15321405,15270281,15222881,15356542,15287525,15303698"

arr = Split(strData,",")

For i = 0 To UBound(arr) : Do
    Wnd0.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = arr(i)
    Wnd0.findById("wnd[0]").sendVKey 0

    If Session.ActiveWindow.Name = "wnd[1]" Then
        session.findById("wnd[1]/tbar[0]/btn[0]").press
    End if

    Wnd0.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_OUTPUT").press

    firstOutputType = Wnd0.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-KSCHL[1,0]").text
    processed = Wnd0.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/lblDV70A-STATUSICON[0,0]").tooltip

    If (firstOutputType = "ZCGB" AND processed = "Not processed") Then
        Wnd0.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3").getAbsoluteRow(0).selected = true
        Wnd0.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/lblDV70A-STATUSICON[0,0]").setFocus
        Wnd0.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/lblDV70A-STATUSICON[0,0]").caretPosition = 0
        Wnd0.findById("wnd[0]/tbar[1]/btn[5]").press
        Wnd0.findById("wnd[0]/usr/cmbNAST-VSZTP").key = "4"
        Wnd0.findById("wnd[0]/tbar[0]/btn[3]").press
        Wnd0.findById("wnd[0]/tbar[0]/btn[3]").press
    Else
        Wnd0.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-KSCHL[1,11]").text = "ZCGB"
        Wnd0.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/cmbNAST-NACHA[3,11]").key = "5"
        Wnd0.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-PARVW[4,11]").text = "YI"
        Wnd0.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-PARVW[4,11]").setFocus
        Wnd0.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-PARVW[4,11]").caretPosition = 2
        Wnd0.findById("wnd[0]").sendVKey 0
        Wnd0.findById("wnd[0]/tbar[1]/btn[5]").press
        Wnd0.findById("wnd[0]/usr/cmbNAST-VSZTP").key = "4"
        Wnd0.findById("wnd[0]/tbar[0]/btn[3]").press
        Wnd0.findById("wnd[0]/tbar[1]/btn[2]").press
        Wnd0.findById("wnd[0]/usr/ctxtNAST-TCODE").text = "CS01"
        Wnd0.findById("wnd[0]/usr/ctxtNAST-TCODE").caretPosition = 4
        Wnd0.findById("wnd[0]").sendVKey 0
        Wnd0.findById("wnd[0]/tbar[0]/btn[3]").press
        Wnd0.findById("wnd[0]/tbar[0]/btn[3]").press
    End If

        Wnd0.findById("wnd[0]/tbar[0]/btn[11]").press

    
Loop While False: Next



