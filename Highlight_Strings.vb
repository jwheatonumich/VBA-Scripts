Sub HighlightStrings()
    Dim xHStr As String, xStrTmp As String
    Dim xHStrLen As Long, xCount As Long, I As Long
    Dim xCell As Range
    Dim xArr
    On Error Resume Next
    xHStr = Application.InputBox("What is the string to highlight:", "KuTools For Excel", , , , , , 2)
    If TypeName(xHStr) <> "String" Then Exit Sub
    Application.ScreenUpdating = False
        xHStrLen = Len(xHStr)
        For Each xCell In Selection
            xArr = Split(xCell.Value, xHStr)
            xCount = UBound(xArr)
            If xCount > 0 Then
                xStrTmp = ""
                For I = 0 To xCount - 1
                    xStrTmp = xStrTmp & xArr(I)
                    xCell.Characters(Len(xStrTmp) + 1, xHStrLen).Font.ColorIndex = 3
                    xStrTmp = xStrTmp & xHStr
                Next
            End If
        Next
    Application.ScreenUpdating = True
End Sub
