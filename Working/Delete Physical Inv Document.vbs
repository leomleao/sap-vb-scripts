
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

'Create an object of type GuiMainWindow
Set Wnd0 = session.findById ("wnd[0]")

'Create an object of type GuiMenubar
Set Menubar = Wnd0.findById ("mbar")

'Create an object of type GuiUserArea
Set UserArea = Wnd0.findById ("usr")

'Create an object of type GuiStatusbar
Set Statusbar = Wnd0.findById ("sbar")

'Define the user's login
UserName = session.Info.User

'' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' ''
'Supporting procedures and functions
'' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' '' ''
' Pressing the "Enter"
Sub pressEnter ()
Wnd0.sendVKey 0
End Sub

' Writing log
Sub WriteLog(strMessage)
	Const FOR_APPENDING = 8
 
	strFileName = "C:\Users\u081715\AppData\Roaming\SAP\SAP GUI\Scripts\log.txt"
 
	Set objFS = CreateObject("Scripting.FileSystemObject")
 
	If objFS.FileExists(strFileName) Then
		Set oFile = objFS.OpenTextFile(strFileName, FOR_APPENDING)
	Else
		Set oFile = objFS.CreateTextFile(strFileName)
	End If
 
	oFile.WriteLine strMessage
End Sub

session.findById("wnd[0]/tbar[0]/okcd"). Text = "/nMI02"
pressEnter ()

' strData = "3364045,3428446,3428447,3455069,3466577,3474559,3476700,3476701,3476706,3476707,3476708,3497948,3497952,3509582,3509603,3520388,3531038,3531334,3534285,3534307,3534313,3534319,3534335,3534345,3534359,3538586,3538587,3538588,3538591,3541225,3545610,3547254,3552267,3554934,3557102,3559791,3562118,3564393,3566882,3566893,3566916,3566935,3566936,3574334,3574385,3576574,3576579,3576692,3579376,3579378,3579380,3581449,3585978,3585984,3585992,3586265,3586266,3588328,3588350,3588399,3588420,3588432,3588517,3591152,3593279,3593282,3593285,3593286,3593287,3593292,3595397,3597355,3599613,3602072,3602098,3602124,3604101,3617956,3618016,3618051,3618072,3618079,3622211,3624124,3624250,3624265,3626374,3628313,3628335,3628391,3630510,3634987,3635073,3635075,3645279,3646849,3646946,3646947,3653884,3660534,3660535,3662713"
strData = "101255922,101255922,101255922,101255922,101255922,101255922,101255922,101255922,101255922,101255922,101255922,101255922,101255922,101255922,101255922,101255923,101255923,101255923,101255923,101255923,101255923,101255923,101255923,101255923,101255923,101255923,101255923,101255924,101255925,101255926,101255927,101255928,101255929,101255940,101255941,101255942,101255943,101255944,101255945,101255946,101255947,101255948,101255949,101255950,101255951,101255952,101255953,101255954,101255955,101255956,101255957,101255958,101255959,101255960,101255961,101255962,101255963,101255964,101255965,101255966,101255967,101255968,101255969,101255970,101255971,101255972,101255973,101255974,101255975,101255976,101255977,101255978,101255979,101255980,101255981,101255982,101255983,101255984,101255985,101255986,101255987,101255988,101255989,101255990,101255991,101255992,101255993,101255994,101255995,101255996,101255997,101255998,101255999,101256000,101256001,101256002,101256003,101256004,101256005,101256006,101256007,101256008,101256009,101256010,101256011,101256012,101256013,101256014,101256015,101256016,101256017,101256018,101256019,101256020,101256021,101256022,101256023,101256024,101256025,101256026,101256027,101256028,101256029,101256030,101256031,101256032,101256033,101256034,101256035,101256036,101256037,101256038,101256039,101256040,101256041,101256042,101256043,101256044,101256045,101256046,101256047,101256048,101256049,101256050,101256051,101256052,101256053,101256054,101256055,101256056,101256057,101256058,101256059,101256060,101256061,101256062,101256063,101256064,101256065,101256066,101256067,101256068,101256069,101256070,101256071,101256072,101256073,101256074,101256074,101256075,101256076,101256077,101256078,101256079,101256080,101256081,101256082,101256083,101256084,101256085,101256086,101256087,101256088,101256089,101256090,101256091,101256092,101256093,101256094,101256095,101256096,101256097,101256098,101256099,101256100,101256101,101256102,101256103,101256104,101256105,101256106,101256107,101256108,101256109,101256110,101256111,101256112,101256113,101256114,101256115,101256116,101256117,101256118,101256119,101256120,101256121,101256122,101256123,101256124,101256125,101256126,101256127,101256128,101256129,101256130,101256131,101256132,101256133,101256134,101256135,101256136,101256137,101256138,101256139,101256140,101256141,101256142,101256143,101256144,101256145,101256146,101256147,101256148,101256149,101256150,101256151,101256152,101256153,101256154,101256155,101256156,101256157,101256158,101256159,101256160,101256161,101256162,101256163,101256164,101256165,101256166,101256167,101256168,101256169,101256170,101256171,101256172,101256173,101256174,101256175,101256176,101256177,101256178,101256179,101256180,101256181,101256182,101256183,101256184,101256185,101256186,101256187,101256188,101256189,101256190,101256191,101256192,101256193,101256194,101256195,101256196,101256197,101256198,101256199,101256200,101256201,101256202,101256203,101256204,101256205,101256206,101256207,101256208,101256209,101256210,101256211,101256212,101256213,101256214,101256215,101256216,101256217,101256218,101256219,101256220,101256221,101256222,101256223,101256224,101256225,101256226,101256227,101256228,101256229,101256230,101256231,101256232,101256233,101256234,101256235,101256236,101256237,101256238,101256239,101256240,101256241,101256242,101256243,101256244,101256245,101256246,101256247,101256248,101256249,101256250,101256251,101256252,101256253,101256254,101256255,101256256,101256257,101256258,101256259,101256260,101256261,101256262,101256263,101256264,101256265,101256266,101256267,101256268,101256269,101256270,101256271,101256272,101256273,101256274,101256275,101256276,101256277,101256278,101256279,101256280,101256281,101256282,101256283,101256284,101256285,101256286,101256287,101256288,101256289,101256290,101256291,101256292,101256293,101256294,101256295,101256296,101256297,101256298,101256299,101256300,101256301,101256302,101256303,101256304,101256305,101256306,101256307,101256308,101256309,101256310,101256311,101256312,101256313,101256314,101256315,101256316,101256317,101256318,101256319,101256320,101256321,101256322,101256323,101256324,101256325,101256326,101256327,101256328,101256329,101256330,101256331,101256332,101256333,101256334,101256335,101256336,101256337,101256338,101256339,101256340,101256341,101256342,101256343,101256344,101256345,101256346,101256347,101256348,101256349,101256350,101256351,101256352,101256353,101256354,101256355,101256356,101256357,101256358,101256359,101256360,101256361,101256362,101256363,101256364,101256365,101256366,101256367,101256368,101256369,101256370,101256371,101256372,101256373,101256374,101256375,101256376,101256377,101256378,101256379,101256380,101256381,101256382,101256383,101256384,101256385,101256386,101256387,101256388,101256389,101256390,101256391,101256392,101256393,101256394,101256395,101256396,101256397,101256398,101256399,101256400,101256401,101256402,101256403,101256404,101256405,101256406,101256407,101256408,101256409,101256410,101256411,101256412,101256413,101256414,101256415,101256416,101256417,101256418,101256419,101256420,101256421,101256422,101256423,101256424,101256425,101256426,101256427,101256428,101256429,101256430,101256431,101256432,101256433,101256434,101256435,101256436,101256437,101256438,101256439,101256440,101256441,101256442,101256443,101256444,101256445,101256446,101256447,101256448,101256449,101256450,101256451,101256452,101256453,101256454,101256455,101256456,101256457,101256457,101256457,101256459,101256464,101256466,101256467,101256468,101256469,101256470,101256472,101256473,101256475,101256476,101256477,101256478,101256479,101256480,101256481,101256482,101256483,101256484,101256485,101256486,101256488,101256494,101256495,101256496,101256499,101256500,101256502,101256505,101256506,101256507,101256507,101256508,101256508,101256508,101256509,101256510,101256511,101256511,101256511,101256511,101256511,101256512,101256512,101256513,101256514,101256515,101256516,101256517,101256518,101256519,101256520,101256521,101256522,101256523,101256524,101256525,101256526,101256527,101256528,101256529,101256530,101256531,101256532,101256533,101256534,101256535,101256536,101256537,101256538,101256539,101256540,101256541,101256542,101256543,101256544,101256545,101256546,101256547,101256548,101256549,101256550,101256551,101256552,101256553,101256554,101256555,101256556,101256557,101256558,101256559,101256560,101256561,101256562,101256563,101256564,101256565,101256566,101256567,101256568,101256569,101256570,101256571,101256572,101256573,101256574,101256575,101256576,101256577,101256578,101256579,101256580,101256581,101256582,101256583,101256584,101256585,101256586,101256587,101256588,101256589,101256590,101256591,101256592,101256593,101256594,101256595,101256596,101256597,101256598,101256599,101256600,101256601,101256602,101256603,101256604,101256605,101256606,101256607,101256608,101256609,101256611,101256612,101256613,101256614,101256615,101256616,101256617,101256618,101256619,101256620,101256621,101256622,101256623,101256624,101256625,101256626,101256627,101256628,101256629,101256630,101256631,101256632,101256633,101256634,101256635,101256636,101256637,101256638,101256639,101256640,101256641,101256642,101256643,101256644,101256645,101256646,101256647,101256648,101256649,101256650,101256651,101256652,101256653,101256654,101256655,101256656,101256657,101256658,101256659,101256660,101256661,101256662,101256663,101256664,101256665,101256666,101256667,101256668,101256669,101256670,101256671,101256672,101256673,101256674,101256675,101256676,101256677,101256678,101256679,101256680,101256681,101256682,101256683,101256684,101256685,101256686,101256687,101256688,101256689,101256690,101256691,101256692,101256693,101256694,101256695,101256696,101256697,101256698,101256699,101256700,101256701,101256702,101256703,101256704,101256706,101256707,101256708,101256709,101256710,101256711,101256712,101256713,101256714,101256715,101256716,101256717,101256718,101256719,101256720,101256721,101256722,101256723,101256724,101256725,101256726,101256727,101256728,101256729,101256730,101256731,101256732,101256733,101256734,101256735,101256736,101256737,101256738,101256739,101256740,101256741,101256742,101256743,101256744,101256745,101256746,101256747,101256748,101256749,101256750,101256751,101256752,101256753,101256754,101256755,101256756,101256757,101256758,101256759,101256760,101256761,101256762,101256763,101256764,101256765,101256765,101256766,101256767,101256768,101256769,101256770,101256771,101256772,101256773,101256774,101256775,101256776,101256777,101256778,101256779,101256780,101256781,101256782,101256784,101256785,101256786,101256787,101256789,101256790,101256791,101256792,101256793,101256794,101256795,101256796,101256797,101256798,101256799,101256800,101256801,101256802,101256803,101256804,101256805,101256806,101256807,101256808,101256809"

arr = Split(strData,",")

For i = 0 To UBound(arr) : Do
    session.findById("wnd[0]/usr/ctxtRM07I-IBLNR").text = arr(i)
    session.findById("wnd[0]/tbar[1]/btn[14]").press
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press    
Loop While False: Next