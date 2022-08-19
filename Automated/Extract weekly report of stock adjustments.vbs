'-Begin-----------------------------------------------------------------

'-Directives----------------------------------------------------------
Option Explicit
On Error Resume Next

'-Variables-----------------------------------------------------------
Dim WSHShell, SAPGUIPath, SID, InstanceNo, application, SapGuiAuto, connection, session, currentDate, connected
Dim strFileName, objFS, oFile, FOR_APPENDING, strMessage

'-Help functions -----------------------------------------------------
Dim shl, mType

currentDate = Day(Now) & "-" & Month(Now) & "-" & Year(Now)

Sub WriteLog(mType, strMessage)
	Set shl = CreateObject("WScript.Shell")
   Call shl.LogEvent(mType, WScript.ScriptName & " " & strMessage)
   Set shl = Nothing
   ' 0    SUCCESS
   ' 1    ERROR
   ' 2    WARNING
   ' 4    INFORMATION
   ' 8    AUDIT_SUCCESS
   ' 16   AUDIT_FAILURE
End Sub

' -Connection----------------------------------------------------------
Set WSHShell = WScript.CreateObject("WScript.Shell")
If IsObject(WSHShell) Then
   '-Set the path to the SAP GUI directory---------------------------
   SAPGUIPath = "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\"

   '-Set the SAP system ID-------------------------------------------
   SID = "WGP"

   '-Starts the SAP GUI----------------------------------------------
   WSHShell.Exec SAPGUIPath & "sapshcut.exe start -system=" & SID & " -maxgui"   
   Set WSHShell = Nothing
End If

connected = false

Do While connected = false
   WriteLog 4, "Trying to get Scripting Engine."   
   If Not IsObject(application) Then
      Set SapGuiAuto  = GetObject("SAPGUI")
      Set application = SapGuiAuto.GetScriptingEngine
   End If
   WriteLog 4, "Trying to stablish connection." 
   If Not IsObject(connection) Then
      Set connection = application.Children(0)
   End If
   WriteLog 4, "Trying to get session." 
   If Not IsObject(session) Then
      Set session    = connection.Children(0)
   End If
   WriteLog 4, "Connecting to session and app."    
   If IsObject(WScript) Then
      WScript.ConnectObject session,     "on"
      WScript.ConnectObject application, "on"
   End If
   WriteLog 4, "Checking for another login."  
   If session.children.count > 1 Then
      WriteLog 4, "There is another login."
      session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").select
      session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").setFocus
      session.findById("wnd[1]/tbar[0]/btn[0]").press
   End If       
   If Err.Number = 0 Then
      WriteLog 4, "We're connected."  
      Action(session)
      connected = true
   End If
   Err.Clear 
   WScript.sleep 1000
Loop
'-End Connection------------------------------------------------------------


Sub Action(session)
   WriteLog 4, "Performing script." 
   Dim oShell : Set oShell = CreateObject("WScript.Shell")  

   session.findById("wnd[0]/tbar[0]/okcd").text = "/nmb51"
   session.findById("wnd[0]").sendVKey 0
   session.findById("wnd[0]/tbar[1]/btn[17]").press
   session.findById("wnd[1]/usr/txtV-LOW").text = "UK_STOCK_ADJS"
   session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
   session.findById("wnd[1]/tbar[0]/btn[8]").press
   session.findById("wnd[0]/tbar[1]/btn[8]").press  

   If session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").RowCount = 0 Then
      Const FOR_APPENDING = 8
      strFileName = "C:\Users\u081715\OneDrive - WAGO\Desktop\TEMP\Automations\Stock adjustments\Message.txt"
      Set objFS = CreateObject("Scripting.FileSystemObject")
      If objFS.FileExists(strFileName) Then
         Set oFile = objFS.OpenTextFile(strFileName, FOR_APPENDING)
      Else
         Set oFile = objFS.CreateTextFile(strFileName)
      End If
      oFile.WriteLine "No stock adjustments performed this week."
   Else
      session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select
      session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\u081715\OneDrive - WAGO\Desktop\TEMP\Automations\Stock adjustments"
      session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Weekly stock adjustments" & "_" & currentDate & ".XLSX"
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      oShell.Run "taskkill /f /im excel.exe" 
   End If   
   
   session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
   session.findById("wnd[0]").sendVKey 0
   
End Sub

'-Sub Main--------------------------------------------------------------