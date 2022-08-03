If session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").ToolbarButtonCount < 10 Then
   session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton "&NAVIGATION_PROFILE_TOOLBAR_EXPAND"
End If
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton "&MB_FILTER"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem "&FILTER"

gridView = "wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell"
rowCount = session.findById(gridView).rowCount
For i = 0 To rowCount - 3 : Do
   currentRow = session.findById(gridView).getCellValue(i,"SELTEXT")
   If currentRow = "Requirement quantity (EINHEIT)" or currentRow = "Quantity withdrawn (EINHEIT)" Then
      Call session.findById(gridView).DoubleClick(i,"SELTEXT")
   End If
Loop While False: Next

session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/btn600_BUTTON").press
session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/txt%%DYN002-LOW").setFocus
session.findById("wnd[2]").sendVKey 2
session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 0,"TEXT"
session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "0"
session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/txt%%DYN003-LOW").setFocus 
session.findById("wnd[2]").sendVKey 2
session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 3,"TEXT"
session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "3"
session.findById("wnd[3]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[2]/tbar[0]/btn[0]").press