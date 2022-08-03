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

session.findById("wnd[0]").maximize
'Open tcode
session.findById("wnd[0]/tbar[0]/okcd"). Text = "/n/MCL/TC01"
session.findById("wnd[0]").sendVKey 0

'Fill basic company data
session.findById("wnd[0]/usr/ctxt/MCL/TC_DY_ALL_1000-VSTEL").text = "0008"
session.findById("wnd[0]/usr/ctxt/MCL/TC_DY_ALL_1000-BUKRS").text = "0008"
session.findById("wnd[0]/usr/ctxt/MCL/TC_DY_ALL_1000-VTWEG").text = "01"
session.findById("wnd[0]/usr/ctxt/MCL/TC_DY_ALL_1000-SPART").text = "00"
session.findById("wnd[0]/usr/ctxt/MCL/TC_DY_ALL_1000-VSTEL").text = "0008"
session.findById("wnd[0]/usr/ctxt/MCL/TC_DY_ALL_1000-UPS_SHIPACCT").text = "6A4235"
 

'Fill address data
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/tabsTAB_01/tabpFK_BASIS/ssubSUB00:/MCL/TC_UPS:1020/ctxt/MCL/TC_S_SENDUNG-CHARGE_TYPE").text = "PRE"
session.findById("wnd[0]/usr/tabsTAB_01/tabpFK_BASIS/ssubSUB00:/MCL/TC_UPS:1020/ctxt/MCL/TC_S_SENDUNG-PICKUP_DATE").text = "06.01.2022"

session.findById("wnd[0]/usr/tabsTAB_01/tabpFK_BASIS/ssubSUB00:/MCL/TC_UPS:1020/txt/MCL/TC_S_RECEIVER-ADDRESS2").text = "Wago Limited"
session.findById("wnd[0]/usr/tabsTAB_01/tabpFK_BASIS/ssubSUB00:/MCL/TC_UPS:1020/txt/MCL/TC_S_RECEIVER-ADDRESS1").text = "Attention"
session.findById("wnd[0]/usr/tabsTAB_01/tabpFK_BASIS/ssubSUB00:/MCL/TC_UPS:1020/txt/MCL/TC_S_RECEIVER-ADDRESS3").text = "Str and house number99"
session.findById("wnd[0]/usr/tabsTAB_01/tabpFK_BASIS/ssubSUB00:/MCL/TC_UPS:1020/txt/MCL/TC_S_RECEIVER-ADDRESS4").text = "Adress line 2"
session.findById("wnd[0]/usr/tabsTAB_01/tabpFK_BASIS/ssubSUB00:/MCL/TC_UPS:1020/txt/MCL/TC_S_RECEIVER-ADDRESS5").text = "Addressline 3"
session.findById("wnd[0]/usr/tabsTAB_01/tabpFK_BASIS/ssubSUB00:/MCL/TC_UPS:1020/txt/MCL/TC_S_RECEIVER-CITY").text = "City"
session.findById("wnd[0]/usr/tabsTAB_01/tabpFK_BASIS/ssubSUB00:/MCL/TC_UPS:1020/ctxt/MCL/TC_S_RECEIVER-STATE").text = "WA"
session.findById("wnd[0]/usr/tabsTAB_01/tabpFK_BASIS/ssubSUB00:/MCL/TC_UPS:1020/txt/MCL/TC_S_RECEIVER-PLZ").text = "CV21 4DQ"
session.findById("wnd[0]/usr/tabsTAB_01/tabpFK_BASIS/ssubSUB00:/MCL/TC_UPS:1020/ctxt/MCL/TC_S_RECEIVER-COUNTRY").text = "GB"
session.findById("wnd[0]/usr/tabsTAB_01/tabpFK_BASIS/ssubSUB00:/MCL/TC_UPS:1020/txt/MCL/TC_S_RECEIVER-PHONE").text = "07775596971"

'Add box
session.findById("wnd[0]/tbar[1]/btn[5]").press
session.findById("wnd[0]/usr/cntlCONTAINER1110/shellcont/shell").pressToolbarButton "ADD"
session.findById("wnd[0]/usr/tabsTAB_02/tabpFK_BASIS/ssubSUB00:/MCL/TC_UPS:1120/ctxt/MCL/TC_S_PACKET-PACKTYP_DESCR").text = "CUSTOMER SUPPLIED BOX"
session.findById("wnd[0]/usr/tabsTAB_02/tabpFK_BASIS/ssubSUB00:/MCL/TC_UPS:1120/txt/MCL/TC_S_PACKET-PACK_WEIGHT").text = "2"

'Release document
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[1]/usr/btnBUTTON_1").press
session.findById("wnd[1]/usr/btnBUTTON_2").press

' strData = "3364045,3428446,3428447,3455069,3466577,3474559,3476700,3476701,3476706,3476707,3476708,3497948,3497952,3509582,3509603,3520388,3531038,3531334,3534285,3534307,3534313,3534319,3534335,3534345,3534359,3538586,3538587,3538588,3538591,3541225,3545610,3547254,3552267,3554934,3557102,3559791,3562118,3564393,3566882,3566893,3566916,3566935,3566936,3574334,3574385,3576574,3576579,3576692,3579376,3579378,3579380,3581449,3585978,3585984,3585992,3586265,3586266,3588328,3588350,3588399,3588420,3588432,3588517,3591152,3593279,3593282,3593285,3593286,3593287,3593292,3595397,3597355,3599613,3602072,3602098,3602124,3604101,3617956,3618016,3618051,3618072,3618079,3622211,3624124,3624250,3624265,3626374,3628313,3628335,3628391,3630510,3634987,3635073,3635075,3645279,3646849,3646946,3646947,3653884,3660534,3660535,3662713"
' strData = "101122037,101122038,101122039,101122080,101122081,101122082,101122083,101122084,101122085,101122086,101122087,101122088,101122089,101122090"
' arr = Split(strData,",")

' For i = 0 To UBound(arr) : Do
'     session.findById("wnd[0]/usr/ctxtRM07I-IBLNR").text = arr(i)
'     session.findById("wnd[0]/tbar[1]/btn[14]").press
'     session.findById("wnd[1]/usr/btnSPOP-OPTION1").press    
' Loop While False: Next


