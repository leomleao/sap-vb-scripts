'-Begin-----------------------------------------------------------------

'-Directives----------------------------------------------------------
Option Explicit
On Error Resume Next

'-Variables-----------------------------------------------------------
Dim WSHShell, SAPGUIPath, SID, InstanceNo, WinTitle, currentDate, currentDate2, totalPicking, totalPacking, jsonExport, strFileName, objFS, oFile, application, SapGuiAuto, connection, session, connected

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

'-Connection----------------------------------------------------------------
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
   Dim oShell : Set oShell = CreateObject("WScript.Shell")
   
   currentDate2 = Day(Now) & "." & Month(Now) & "." & Year(Now)
   currentDate = Day(Now) & "-" & Month(Now) & "-" & Year(Now)

   session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl06f"
   session.findById("wnd[0]").sendVKey 0
   session.findById("wnd[0]/tbar[1]/btn[17]").press
   session.findById("wnd[1]/usr/txtV-LOW").text = "UK-NO-PICKING"
   session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
   session.findById("wnd[1]/tbar[0]/btn[8]").press
   session.findById("wnd[0]/tbar[1]/btn[8]").press
   If session.findById("wnd[0]/sbar/pane[0]").text = "No deliveries selected" Then
      totalPicking = 0
   Else
      session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,"WADAT"
      session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "WADAT"
      session.findById("wnd[0]/tbar[1]/btn[29]").press
      session.findById("wnd[1]").sendVKey 2
      session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 2,"TEXT"
      session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
      session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = currentDate2
      session.findById("wnd[1]").sendVKey 0
      totalPicking = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").rowCount
      session.findById("wnd[0]/tbar[1]/btn[43]").press
      session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\u081715\OneDrive - WAGO\Desktop\TEMP\Automations\Extract count of deliveries\NEW"
      session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "UK-NO-PICKING" & "_" & currentDate & ".XLSX"
      session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      oShell.Run "taskkill /f /im excel.exe"  
   End If  
  
   session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl06f"
   session.findById("wnd[0]").sendVKey 0
   session.findById("wnd[0]/tbar[1]/btn[17]").press
   session.findById("wnd[1]/usr/txtV-LOW").text = "UK-NO-PACKING"
   session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
   session.findById("wnd[1]/tbar[0]/btn[8]").press
   session.findById("wnd[0]/tbar[1]/btn[8]").press
   
   If session.findById("wnd[0]/sbar/pane[0]").text = "No deliveries selected" Then
      totalPacking = 0
   Else
      session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,"WADAT"
      session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "WADAT"
      session.findById("wnd[0]/tbar[1]/btn[29]").press
      session.findById("wnd[1]").sendVKey 2
      session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 2,"TEXT"
      session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
      session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = currentDate2
      session.findById("wnd[1]").sendVKey 0
      totalPacking = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").rowCount
      session.findById("wnd[0]/tbar[1]/btn[43]").press
      session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\u081715\OneDrive - WAGO\Desktop\TEMP\Automations\Extract count of deliveries\NEW"
      session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "UK-NO-PACKING" & "_" & currentDate & ".XLSX"
      session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      oShell.Run "taskkill /f /im excel.exe"
   End If  
  
   session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
   session.findById("wnd[0]").sendVKey 0

  
   oShell.Run "taskkill /f /im excel.exe"

   jsonExport = "{""UK-NO-PICKING"": " & totalPicking & ", ""UK-NO-PACKING:"": " & totalPacking & "}"

   Const FOR_APPENDING = 8
   strFileName = "C:\Users\u081715\OneDrive - WAGO\Desktop\TEMP\Automations\Extract count of deliveries\NEW\count.json"

   Set objFS = CreateObject("Scripting.FileSystemObject")

   If objFS.FileExists(strFileName) Then
      Set oFile = objFS.OpenTextFile(strFileName, FOR_APPENDING)
   Else
      Set oFile = objFS.CreateTextFile(strFileName)
   End If

   oFile.WriteLine jsonExport
End Sub

'-Sub Main--------------------------------------------------------------