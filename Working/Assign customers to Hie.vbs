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

session.findById("wnd[0]/tbar[0]/okcd"). Text = "/nVDH1N"
session.findById("wnd[0]").sendVKey 0

topLevel = "92836"
strData =  "92167,90498,800023,801161,804199,804910,805488,806388,806633,806688,806769,806880,807112,807741,807794,808025,808032,808078,808172,808173,808227,808314,808374,808482,808608,808676,808842,809067,809068,809132,809133,809282,809282,809387,809479,809512,809604,809630,809706,810269,810392,810393,810423,810462,810464,810514,810583,810678,810685,811483,811778,812011,812134,812137,812138,812517,812518,812541,812564,812574,812840,813055,813266,7305819,7306831,7306871,7307062,7307671,7308647,7309105,7310380,7312122,7312556,7312633,7313887,7314014,7314014,7314961,7316053,7316666,7317525,7317725,7317749,7318428,7321189,7321646,7321898,7322535,7323075,7323518,7323593,7324005,7324055,7324345,7324539,7324604,7324926,7325056,7325266,7325266,7325995,7326148,7327768,7329930,7330110,7330213,7330214,7330383,7333257,7333279,7335359,7335403,7335405,7335533,7335683,7335935,7373092,7384289,7384289,7386652,7388399,7389764,7402094,7408563,7411349,7413607,7423250,7425250,7428162,7428162,7430801,7431000,7435764,7440982,7448962,7450327,7450791,7462807,7463120,7307183"

currentDate = Day(Now) & "." & Month(Now) & "." & Year(Now)
startDate = "01.01.2020"


session.findById("wnd[0]/usr/ctxtS_DATE").text = currentDate
session.findById("wnd[0]/usr/ctxtS_KUNNR-LOW").text = topLevel
session.findById("wnd[0]/usr/ctxtS_DATE").setFocus
session.findById("wnd[0]/usr/ctxtS_DATE").caretPosition = 2
session.findById("wnd[0]/tbar[1]/btn[8]").press


arr = Split(strData,",")

For i = 0 To UBound(arr)
   session.findById("wnd[0]/tbar[1]/btn[27]").press
   session.findById("wnd[0]/usr/ctxtKNVH-HKUNNR").text = topLevel
   session.findById("wnd[0]/usr/ctxtKNVH-HVKORG").text = "0008"
   session.findById("wnd[0]/usr/ctxtKNVH-HVTWEG").text = "00"
   session.findById("wnd[0]/usr/ctxtKNVH-HSPART").text = "00"
   session.findById("wnd[0]/usr/ctxtKNVH-VKORG").text = "0008"
   session.findById("wnd[0]/usr/ctxtKNVH-VTWEG").text = "00"
   session.findById("wnd[0]/usr/ctxtKNVH-SPART").text = "00"
   session.findById("wnd[0]/usr/ctxtKNVH-KUNNR").text = arr(i)
   session.findById("wnd[0]/usr/ctxtKNVH-HKUNNR").caretPosition = 5
   session.findById("wnd[0]/usr/ctxtKNVH-DATAB").text = startDate
   session.findById("wnd[0]/usr/btnDVF_GET_DETAIL_BUTTON").press
Next




