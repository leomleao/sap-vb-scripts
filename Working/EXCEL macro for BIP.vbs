Sub transposeSelectionIntoString()
'
' transposeSelectionIntoString Macro
' Used to create scripts for SAP
'
' Keyboard Shortcut: Ctrl+q
'
Dim strIn As String
Dim cl As Range
Dim rng As Range
strIn = ""

If TypeName(Selection) = "Range" Then
        If Selection.Areas.Count = 1 Then
            Debug.Print Selection.Address
      
        Else
            MsgBox "Please select only one area.", vbInformation
        End If
    End If


Set rng = Range(Selection.Address)

For Each cl In rng
        If cl.EntireRow.Hidden = False Then
            Debug.Print cl
            strIn = strIn & "," & cl.Value
        End If
    Next
    
strIn = Right(strIn, Len(strIn) - 1)
strIn = """" & strIn & """"

'Place string into clipboard
Clipboard strIn

End Sub
Sub transposeSelectionIntoString2()
'
' transposeSelectionIntoString Macro
' Used to create scripts for SAP
'
' Keyboard Shortcut: Ctrl+q
'
Dim strIn As String
Dim cl As Range
Dim rng As Range
strIn = ""

If TypeName(Selection) = "Range" Then
        If Selection.Areas.Count = 1 Then
            Debug.Print Selection.Address
      
        Else
            MsgBox "Please select only one area.", vbInformation
        End If
    End If


Set rng = Range(Selection.Address)

For Each cl In rng
        If cl.EntireRow.Hidden = False Then
            Debug.Print cl
            strIn = strIn & ";" & cl.Value
        End If
    Next
    
strIn = Right(strIn, Len(strIn) - 1)

'Place string into clipboard
Clipboard strIn


End Sub

Function Clipboard(Optional StoreText As String) As String
'PURPOSE: Read/Write to Clipboard
'Source: ExcelHero.com (Daniel Ferry)

Dim x As Variant

'Store as variant for 64-bit VBA support
  x = StoreText

'Create HTMLFile Object
  With CreateObject("htmlfile")
    With .parentWindow.clipboardData
      Select Case True
        Case Len(StoreText)
          'Write to the clipboard
            .setData "text", x
        Case Else
          'Read from the clipboard (no variable passed through)
            Clipboard = .GetData("text")
      End Select
    End With
  End With

End Function

Public Function RemoveCharsBeg(RemChars As String, qty As Long) As String
Dim TempString As String
TempString = RemChars
If Left(RemChars, 1) = "1" Then
    If Len(RemChars) > 1 Then
        TempString = Right(RemChars, Len(RemChars) - 1)
    End If
End If
RemoveCharsBeg = TempString
End Function

