
'Create a WScript.Shell object
Set WshShell = WScript.CreateObject ("WScript.Shell")

'' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' ''
'Connecting to SAP
'' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' ''

'Create an object
Set SapGuiAuto = GetObject ("SAPGUI")

'Create an object of type GuiApplication (COM-interface)
Set application = SapGuiAuto.GetScriptingEngine ()

'Creating an object of type GuiSession is a session that corresponds to the active SAP window
'Ie. when you start WSF, the script itself will run in the same SAP window from which it is running
Set session = application.ActiveSession ()

WScript.ConnectObject session, "on"
WScript.ConnectObject application, "on"

'Create an object of type GuiMainWindow
Set Wnd0 = session.findById ("wnd [0]")

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

'Starting the transaction
Sub startTransaction (transaction_name)
Wnd0.findById ("tbar [0] / okcd"). Text = transaction_name
pressEnter ()
End Sub

'Pressing the "Enter"
Sub pressEnter ()
Wnd0.sendVKey 0
End Sub

'Pressing F2
Sub pressF2 ()
Wnd0.sendVKey 2
End Sub

'Pressing the F3 button
Sub pressF3 ()
Wnd0.sendVKey 3
End Sub

'Pressing the F5
Sub pressF5 ()
Wnd0.sendVKey 5
End Sub

'Pressing the F8 button
Sub pressF8 ()
Wnd0.sendVKey 8
End Sub

'Dialog for selecting a file, creating read streams from a file and writing to a file
Function selectFile (createOuputFile)
Set objDialog = CreateObject ("UserAccounts.CommonDialog")
'Fill in the properties and open the dialog
With objDialog
.InitialDir = WshShell.SpecialFolders ("Desktop") 'The start folder is the desktop
.Filter = "Text files (* .csv; *. Txt) | * .csv; *. Txt"
result = .ShowOpen
End With
'If the file is not selected - we exit
If (result = 0) Then WScript.Quit
inputFile = objDialog.FileName 'Full path to the selected file
Set fso = CreateObject ("Scripting.FileSystemObject")
Set inputStream = fso.OpenTextFile (inputFile)
'Create output file?
If (createOuputFile) Then
outputFile = Left (inputFile, Len (inputFile) - 3) & "out" & Right (inputFile, 4)
Set outputStream = fso.CreateTextFile (outputFile, True)
'Return an array from the read stream from the file and the write stream to the file
selectFile = Array (inputStream, outputStream)
Else
'Return the read stream from file
selectFile = inputStream
End If
End Function

'Fill one row in the table (for ME51N)
Sub fill_row (row, material, kolvo, zavod, zatreboval)
Set grid = session.findById (UserArea.findByName ("GRIDCONTROL", "GuiCustomControl"). Id & "/ shellcont / shell")
grid.modifyCell row, "KNTTP", "K" 'Account assignment type
grid.modifyCell row, "MATNR", material 'Material
grid.modifyCell row, "MENGE", number 'Number
grid.modifyCell row, "NAME1", factory 'Factory
grid.modifyCell row, "LGOBE", "0001 '' Warehouse
grid.modifyCell row, "AFNAM", requested 'Requested
End Sub
