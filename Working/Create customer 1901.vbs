
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

Wnd0.findById ("tbar[0]/okcd"). Text = "/nVD01"
session.findById("wnd[0]").sendVKey 0

strData = "801343,802304,806873,808226,810578,7323182,7323513,7324604,7326011,807238,7431074,7310500,507197,7323594"
' strData = "7449730"

arr = Split(strData,",")

For i = 0 To UBound(arr) : Do
    session.findById("wnd[1]/usr/cmbRF02D-KTOKD").key = "ZDEB"
    session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = arr(i)
    session.findById("wnd[1]/usr/ctxtRF02D-VKORG").text = "1901"
    session.findById("wnd[1]/usr/ctxtRF02D-REF_KUNNR").text = arr(i)
    session.findById("wnd[1]/usr/ctxtRF02D-REF_VKORG").text = "0008"
    session.findById("wnd[1]/usr/ctxtRF02D-REF_VTWEG").text = "00"
    session.findById("wnd[1]/usr/ctxtRF02D-REF_SPART").text = "00"
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:SAPLATAB:0200/subAREA1:SAPMF02D:7310/ctxtKNVV-VKBUR").text = "1902"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 11
Loop While False: Next




