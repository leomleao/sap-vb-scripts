
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

session.findById("wnd[0]/tbar[0]/okcd"). Text = "/nMI02"
pressEnter ()

' strData = "3364045,3428446,3428447,3455069,3466577,3474559,3476700,3476701,3476706,3476707,3476708,3497948,3497952,3509582,3509603,3520388,3531038,3531334,3534285,3534307,3534313,3534319,3534335,3534345,3534359,3538586,3538587,3538588,3538591,3541225,3545610,3547254,3552267,3554934,3557102,3559791,3562118,3564393,3566882,3566893,3566916,3566935,3566936,3574334,3574385,3576574,3576579,3576692,3579376,3579378,3579380,3581449,3585978,3585984,3585992,3586265,3586266,3588328,3588350,3588399,3588420,3588432,3588517,3591152,3593279,3593282,3593285,3593286,3593287,3593292,3595397,3597355,3599613,3602072,3602098,3602124,3604101,3617956,3618016,3618051,3618072,3618079,3622211,3624124,3624250,3624265,3626374,3628313,3628335,3628391,3630510,3634987,3635073,3635075,3645279,3646849,3646946,3646947,3653884,3660534,3660535,3662713"
strData = "101122037,101122038,101122039,101122080,101122081,101122082,101122083,101122084,101122085,101122086,101122087,101122088,101122089,101122090"
arr = Split(strData,",")

For i = 0 To UBound(arr) : Do
    session.findById("wnd[0]/usr/ctxtRM07I-IBLNR").text = arr(i)
    session.findById("wnd[0]/tbar[1]/btn[14]").press
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press    
Loop While False: Next