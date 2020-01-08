' テーブルを図にします’

Public Sub セルを縦の図にする()
'
' テーブルを図にします
'
    '開始位置
    RowAddress = Selection.Row
    ColumnAddress = Selection.Column
    '選択範囲
    ColumnCount = Selection.Columns.Count - 1
    RowCount = Selection.Rows.Count


    ''線を消す
    Selection.Borders.LineStyle = xlNone

For i = 1 To RowCount
'縦一行を選択
    '選択範囲を結合
    'Range(Cells(RowAddress + i - 1, ColumnAddress), Cells(RowAddress + i - 1, ColumnAddress + ColumnCount)).Merge
    ''線を引く
    With Range(Cells(RowAddress + i - 1, ColumnAddress), Cells(RowAddress + i - 1, ColumnAddress + ColumnCount)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = vbBlack
    End With
    With Range(Cells(RowAddress + i - 1, ColumnAddress), Cells(RowAddress + i - 1, ColumnAddress + ColumnCount)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = vbBlack
    End With
    With Range(Cells(RowAddress + i - 1, ColumnAddress), Cells(RowAddress + i - 1, ColumnAddress + ColumnCount)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = vbBlack
    End With
    With Range(Cells(RowAddress + i - 1, ColumnAddress), Cells(RowAddress + i - 1, ColumnAddress + ColumnCount)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = vbBlack
    End With
Next i

    
End Sub
