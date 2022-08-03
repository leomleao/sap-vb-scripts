
@rem -Begin-------------------------------------------------------------
@rem
@rem  Install dotNET 5 or 6 SDK and set DOTNET_ROOT environment
@rem  variable. Depending on which release is used, the comments in
@rem  the code below are to be moved. Standard is dotNET 6.
@rem
@rem -------------------------------------------------------------------

mkdir Tracker_RunScript

copy Tracker_RunScript.cs Tracker_RunScript\Tracker_RunScript.cs
@rem copy Tracker.RunScript.net5.csproj Tracker_RunScript\Tracker_RunScript.csproj
copy Tracker.RunScript.net6.csproj Tracker_RunScript\Tracker_RunScript.csproj

@cd Tracker_RunScript

dotnet build

@rem start /w /min bin\Debug\net5.0\Tracker_RunScript.exe
start /w /min bin\Debug\net6.0\Tracker_RunScript.exe

@cd..

rmdir /s /q Tracker_RunScript

@rem -End---------------------------------------------------------------
