Attribute VB_Name = "TableMerge"
Public Sub CellTableMake()
Attribute CellTableMake.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CellTableMake Macro
'
' テーブルを図にします（未完成）
'

    '開始位置
    RowAddress = Selection.Row
    ColumnAddress = Selection.Column
    '選択範囲
    ColumnCount = Selection.Rows.Count
    RowCount = Selection.Rows.Count


    ''線を消す
    Selection.Borders.LineStyle = xlNone

For i = 1 To RowCount
'縦一行を選択
    Range(Cells(RowAddress + i - 1, ColumnAddress), Cells(RowAddress + i - 1, ColumnAddress + ColumnCount)).Select

    '選択範囲を結合
    'Selection.Merge
    ''線を引く
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = vbBlack
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = vbBlack
    End With
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = vbBlack
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = vbBlack
    End With
Next i

'選択範囲をアンマージ
Selection.UnMerge
    
End Sub
