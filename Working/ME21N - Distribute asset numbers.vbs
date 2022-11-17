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


'Open acct assignment tab:
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13").select


table = "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1000/tblSAPLMEACCTVIDYN_1000TC"

qty = CInt(session.findById(table & "/txtMEACCT1000-MENGE[2,0]").text)

If qty = 0 Then 
    qty = InputBox("Please type the quantity:")
End If

If session.findById(table & "/ctxtMEACCT1000-ANLN1[6,0]").text = "" Then 
    firstAsset = InputBox("Please type the first asset:")
    session.findById(table & "/ctxtMEACCT1000-ANLN1[6,0]").text = firstAsset
End If

firstAsset = CDbl(session.findById(table & "/ctxtMEACCT1000-ANLN1[6,0]").text)

session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB1:SAPLMEACCTVI:1200/cmbMEACCT1200-VRTKZ").key = ""
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB1:SAPLMEACCTVI:1200/cmbMEACCT1200-VRTKZ").key = "1"


For i = 1 To qty - 1
    session.findById(table).verticalScrollbar.position = i
    session.findById(table & "/txtMEACCT1000-MENGE[2,1]").text = "1"
    session.findById(table & "/ctxtMEACCT1000-ANLN1[6,1]").text = firstAsset + i    
Next

session.findById(table).verticalScrollbar.position = 0
session.findById(table & "/txtMEACCT1000-MENGE[2,0]").text = "1"
session.findById("wnd[0]").sendVKey 0






