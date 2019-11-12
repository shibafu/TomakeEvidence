Public Sub 存在しているものを数える()

    '現在の印刷範囲をセルにしてい
    Dim Worksheet As Worksheet
    Set ws = ActiveWorksheet

    Dim PrintArea As String
    PrintArea = ws.PageSetup.PrintArea
    
    '現在の選択範囲の行と列を取得
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

