
@rem -Begin-------------------------------------------------------------

@rem -x64---------------------------------------------------------------
c:\Windows\Microsoft.NET\Framework64\v4.0.30319\vbc.exe /target:winexe /platform:x64 /out:Tracker_RunScript.vb.exe Tracker_RunScript.vb

@rem -x86---------------------------------------------------------------
@rem c:\Windows\Microsoft.NET\Framework\v4.0.30319\vbc.exe /target:winexe /platform:x86 /out:Tracker_RunScript.vb.exe Tracker_RunScript.vb

start /min Tracker_RunScript.vb.exe

@rem -End---------------------------------------------------------------
