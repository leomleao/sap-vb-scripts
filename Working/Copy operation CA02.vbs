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



groupStr = "81670,81998,82389,82579,82580,87358"
groups = Split(groupStr,",")

For i = 0 To UBound(groups)

''Get the next group
session.findById("wnd[0]/tbar[0]/okcd"). Text = "/nCA02"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRC27M-MATNR").text = ""
session.findById("wnd[0]/usr/ctxtRC271-PLNNR").text = groups(i)
session.findById("wnd[0]/usr/ctxtRC271-AENNR").text = ""
session.findById("wnd[0]").sendVKey 0

    ''Find the line with validity up to 31.12.2999
    If session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-DATUB[39,2]").text = "30.12.2999" Then

        session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VORNR[0,3]").text = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VORNR[0,2]").text
        session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-ARBPL[2,3]").text = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-ARBPL[2,2]").text
        session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-STEUS[4,3]").text = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-STEUS[4,2]").text
        session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-KTSCH[5,3]").text = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-KTSCH[5,2]").text
        session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,3]").text = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,2]").text
        session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-BMSCH[14,3]").text = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-BMSCH[14,2]").text
        session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[16,3]").text = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[16,2]").text
        session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-VGE01[17,3]").text = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-VGE01[17,2]").text
        session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[19,3]").text = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[19,2]").text
        session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-VGE02[20,3]").text = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-VGE02[20,2]").text
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400").getAbsoluteRow(2).selected = true
        session.findById("wnd[0]/tbar[1]/btn[14]").press
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        session.findById("wnd[0]").sendVKey 11 
    Else 
            If session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-DATUB[39,1]").text = "30.12.2999" Then
            session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VORNR[0,2]").text = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VORNR[0,1]").text
            session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-ARBPL[2,2]").text = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-ARBPL[2,1]").text
            session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-STEUS[4,2]").text = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-STEUS[4,1]").text
            session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-KTSCH[5,2]").text = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-KTSCH[5,1]").text
            session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,2]").text = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,1]").text
            session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-BMSCH[14,2]").text = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-BMSCH[14,1]").text
            session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[16,2]").text = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[16,1]").text
            session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-VGE01[17,2]").text = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-VGE01[17,1]").text
            session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[19,2]").text = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[19,1]").text
            session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-VGE02[20,2]").text = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-VGE02[20,1]").text
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400").getAbsoluteRow(1).selected = true
            session.findById("wnd[0]/tbar[1]/btn[14]").press
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            session.findById("wnd[0]").sendVKey 11 
        End If
    End If

session.findById("wnd[0]/tbar[0]/btn[15]").press

Next

