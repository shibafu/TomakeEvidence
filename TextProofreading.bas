Attribute VB_Name = "TextProofreading"
Public Sub ���݂��Ă�����̂𐔂���()

    '���݂̈���͈͂��Z���ɂ��Ă�
    Dim Worksheet As Worksheet
    Set ws = ActiveWorkSheet

    Dim PrintArea As String
    PrintArea = ws.PageSetup.PrintArea
    
    '���݂̑I��͈͂̍s�Ɨ���擾
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

