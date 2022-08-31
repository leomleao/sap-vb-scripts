
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

str = "101029076,101032285,101029105,101032311,101029132,101032336,101032366,101029166,101032529,101029343,101029349,101032534,101029357,101032541,101029526,101032702,101029595,101032767,101029615,101032785,101029645,101032814,101032850,101029683,101032901,101029738,101032917,101029755,101029776,101032937,101032956,101029796,101029907,101033063"
splits = 4

For i = 1 To splits - 1
    session.createSession
Next

WScript.sleep 10000

str1 = ""
str2 = ""
str3 = ""
str4 = ""
str5 = ""
last = 1
current = ""

object = Split(str,",")
For each x in object    
    Execute ("str" & last & " = str" & last & " & "","" & x")
    If last = splits Then
        last = 1
    Else
        last = last + 1
    End If 
Next

For i = 1 To splits
    Execute ("str" & i & " = Right(str" & i & ",Len(str" & i & ") - 1)")  
    Execute ("current = str" & i )

    Set shell = CreateObject("WScript.Shell")
    shell.CurrentDirectory = "C:\Users\u081715\OneDrive - WAGO\Desktop\TEMP\Scripts\"
    shell.Run "C:\Windows\System32\cscript.exe ""C:\Users\u081715\OneDrive - WAGO\Desktop\TEMP\Scripts\Working\Change name of contacts.vbs"" " & """" & i - 1 & """" & " " & """" & current & """" 
Next



