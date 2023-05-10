
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


strMaterial = "0209-0173|2090173,0210-0112|2100112,0210-0118|2100118,0221-0412|2210412,0221-0413|2210413,0221-0413|2210413,0221-0413|2210413,0221-0413|2210413,0221-0413|2210413,0221-0413|2210413,0222-0415|2220415,0280-0324|2800324,0282-0301|2820301,0885-0108|8850108,0890-0122|8900122,2002-3291|20023291,2273-0208|22730208,0231-0602|51110641,0862-2504|51131515,0862-2504|51131515,0262-0138|51139681,0262-0138|51139681,0262-0138|51139681,0210-0133|51170633,0753-0629|51221211,0280-0623|51243829,2006-0401|51301259,0769-0607|60283638,0721-0468|60352265,0210-0509|60354589,2002-0400|60356162,0281-0611|60359536,0281-0611|60359536,0281-0611|60359536,2102-1201|60388238,2102-1201|60388238,2102-1201|60388238,2102-1201|60388238,2102-1201|60388238,2102-1201|60388238,2102-1201|60388238,2102-1201|60388238,2102-1301|60388241,2000-2231|60402180,0750-0451|60404071,0280-0901|60412441,0280-0687|60412573,2773-0406|60416310,2052-1207|60467048,2773-0405|60473950,0207-1431|60489010,0207-1432|60489011,2050-1204|60499718,0859-0410|60507163"
strBatches = "0004026529,2000544906,2000495663,0004200106,0004160453,0004183488,0004187243,0004187245,0004187250,0004193607,0004189674,0003999125,0003849393,0003861458,0004095596,0004008004,0004176923,0003804798,0003917681,0004212120,0003829987,0003829999,0004120435,2000535209,0003909199,0003884412,0004040271,0004166411,0001945101,2000549090,0004188060,0004012828,0004024769,0004027987,0004140699,0004140710,0004140715,0004140717,0004140722,0004140744,0004140746,0004140753,0004140789,0003807837,0004167605,0004152251,0004118990,0004130889,0004162290,0003860048,0004168782,0004220190,0000999911,0004004549"
strQty = "100,200,9,176000,88000,88000,88000,88000,88000,176000,51200,1200,200,25,50,100,88000,100,2000,2400,150,100,350,35,29,200,100,25,500,10,150,250,300,450,100,1000,900,1050,1050,350,1000,1050,50,50,5,400,100,100,200,15,4,3,200,500"

arrMaterial = Split(strMaterial,",")
arrBatches = Split(strBatches,",")
arrQty = Split(strQty,",")


For i = 0 To UBound(arrMaterial) : Do
   session.findById ("wnd[0]/tbar[0]/okcd"). Text = "/nMB11"
   pressEnter ()

   session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").text = "309"
   session.findById("wnd[0]/usr/ctxtRM07M-WERKS").text = "0008"
   session.findById("wnd[0]/usr/ctxtRM07M-LGORT").text = "0001"
   session.findById("wnd[0]/usr/txtMKPF-BKTXT").text = "Fix dummy batch"
   session.findById("wnd[0]").sendVKey 0
   session.findById("wnd[0]/usr/ctxtMSEGK-UMWRK").text = "0008"
   session.findById("wnd[0]/usr/ctxtMSEGK-UMLGO").text = "0001"
   session.findById("wnd[0]/usr/ctxtMSEGK-UMMAT").text = arrMaterial(i)
   session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[0,7]").text = arrMaterial(i)
   session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[0,26]").text = arrQty(i)
   session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-LGORT[0,48]").text = "0001"
   session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[0,53]").text = arrBatches(i)
   session.findById("wnd[0]").sendVKey 11
   session.findById("wnd[0]").sendVKey 0
   session.findById("wnd[0]").sendVKey 0
   session.findById("wnd[0]/usr/ctxtMSEG-UMCHA").text = "0000999908"
   session.findById("wnd[0]").sendVKey 11
   session.findById("wnd[0]").sendVKey 0
   session.findById("wnd[0]/tbar[0]/btn[3]").press
Loop While False: Next