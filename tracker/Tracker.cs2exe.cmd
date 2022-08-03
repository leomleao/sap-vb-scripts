
@rem -Begin-------------------------------------------------------------

@rem -x64---------------------------------------------------------------
c:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe /target:winexe /platform:x64 /out:Tracker_RunScript.cs.exe /reference:"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\Microsoft.VisualBasic.dll" Tracker_RunScript.cs

@rem -x86---------------------------------------------------------------
@rem c:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe /target:winexe /platform:x86 /out:Tracker_RunScript.cs.exe /reference:"C:\Windows\Microsoft.NET\Framework\v4.0.30319\Microsoft.VisualBasic.dll" Tracker_RunScript.cs

start /min Tracker_RunScript.cs.exe

@rem -End---------------------------------------------------------------
