'-Begin-----------------------------------------------------------------

'-Directives----------------------------------------------------------
Option Explicit
On Error Resume Next

'-Variables-----------------------------------------------------------
Dim WSHShell, SAPGUIPath, SID, InstanceNo, application, SapGuiAuto, connection, session, currentDate, connected
Dim strFileName, objFS, oFile, FOR_APPENDING, strMessage, i

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

    WriteLog 4, "Opening SQ01."  
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nsq01"

    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/mbar/menu[5]/menu[0]").select
    session.findById("wnd[1]/usr/radRAD1").select
    session.findById("wnd[1]/tbar[0]/btn[2]").press
    session.findById("wnd[0]/mbar/menu[1]/menu[2]").select

    WriteLog 4, "Selecting the user group."   
    For i = 0 To session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").rowcount - 1
        WriteLog 4, Err.Description   
        If session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").getCellValue(i,"DBGBNUM") = "Z_MM_UK" Then
            WriteLog 4, "Found the key."   
            session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = i
        End If
    Next

    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell").currentCellRow = 8
    session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "PROD_ORD_DATE"
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press

    WriteLog 4, "Extracting report in spreasheet."
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
    session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem "&XXL"

    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\u081715\OneDrive - WAGO\08_Rugby All\Production\Production_Schedule\Data"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "PO 1 conf dates.XLSX"
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    oShell.Run "taskkill /f /im excel.exe" 


    WriteLog 0, "All done, closing down."  
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
    session.findById("wnd[0]").sendVKey 0
   oShell.Run "taskkill /f /im saplogon.exe"    

   
End Sub

'-Sub Main--------------------------------------------------------------