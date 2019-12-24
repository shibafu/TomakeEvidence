Attribute VB_Name = "TextProofreading"
Public Sub ‘¶İ‚µ‚Ä‚¢‚é‚à‚Ì‚ğ”‚¦‚é()

    'Œ»İ‚Ìˆóü”ÍˆÍ‚ğƒZƒ‹‚É‚µ‚Ä‚¢
    Dim Worksheet As Worksheet
    Set ws = ActiveWorkSheet

    Dim PrintArea As String
    PrintArea = ws.PageSetup.PrintArea
    
    'Œ»İ‚Ì‘I‘ğ”ÍˆÍ‚Ìs‚Æ—ñ‚ğæ“¾
    ws.Range(PrintArea).Activate
    
    
    Dim RowsNumber As Integer
    RowsNumber = Selection.Rows.Count
    
    Dim ColumnNumber As Integer
    ColumnNumber = Selection.Columns.Count
    
    For i = 1 To RowsNumber
        For j = 1 To ColumnNumber
            Something (Range(Cells(i, j)))
    
        Next j
    Next i

End Sub

Function Something(Range As Range)
    If Range.Value Like "*" Then

    End If
End Function

