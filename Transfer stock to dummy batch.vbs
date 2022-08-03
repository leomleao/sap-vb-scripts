
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


strMaterial = "0770-0515|7700515,0770-0515|7700515,0221-0412|60273342,0221-0412|60273342,2773-0402|60408237,0280-0901|60412441,0280-0833|60412719,0280-0833|60412719,2773-0404|60414567,0282-0681|60428093,2787-2134|60452125,2787-2135|60452134,2787-2154|60452136,2787-2154|60452136,2787-2157|60452137,2687-2142|60461071,0221-2505|60485995"
strBatches = "0003623484,0003755040,0003864454,0003870486,0003782505,0003797603,0003812236,0003812798,0003855178,0003769253,2000427277,2000427274,2000426015,2000430900,2000444918,2000466392,0003776069"
strQty = "75,250,5000,2500,240,100,400,100,480,46,10,10,4,6,2,2,5"

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