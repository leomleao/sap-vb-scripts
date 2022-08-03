
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

Wnd0.findById ("tbar[0]/okcd"). Text = "/nMB11"
pressEnter ()

session.findById("wnd[0]/usr/ctxtRM07M-BWARTWA").text = "309"
session.findById("wnd[0]/usr/ctxtRM07M-WERKS").text = "0008"
session.findById("wnd[0]/usr/ctxtRM07M-LGORT").text = "0001"
session.findById("wnd[0]/usr/txtMKPF-BKTXT").text = "Transfer new replcmt IDs"

' strMaterial = "0248-0501|60434629"
strMaterial = "51008291"
' strNewMaterial = "0248-0501|60479507"
strNewMaterial = "60502384"
' strQty = "1"
strQty = "36110"

arrMaterial = Split(strMaterial,",")
arrNewMaterial = Split(strNewMaterial,",")
arrQty = Split(strQty,",")


For i = 0 To UBound(arrMaterial) : Do
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtMSEGK-UMWRK").text = "0008"
    session.findById("wnd[0]/usr/ctxtMSEGK-UMLGO").text = "0001"
    session.findById("wnd[0]/usr/ctxtMSEGK-UMMAT").text = arrNewMaterial(i)
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-MATNR[0,7]").text = arrMaterial(i)
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/txtMSEG-ERFMG[0,26]").text = arrQty(i)
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-LGORT[0,48]").text = "0001"
    session.findById("wnd[0]/usr/sub:SAPMM07M:0421/ctxtMSEG-CHARG[0,53]").text = "0000999908"
    session.findById("wnd[0]").sendVKey 11
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[0]/btn[3]").press
Loop While False: Next