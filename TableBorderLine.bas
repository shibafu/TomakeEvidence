Attribute VB_Name = "TableBorderLine"
Public Sub 表作成整列()
Attribute 表作成整列.VB_ProcData.VB_Invoke_Func = "T\n14"
'
' Macro1 Macro
'
' Keyboard Shortcut: Ctrl+Shift+T
' 選択範囲のテーブルを2列ずつの表に線を引く変更
'

    Dim ColumnNum As Integer
    ColumnNum = Selection.Columns.Count / 2
    '選択範囲の数表を出力する
    For i = 1 To Selection.Rows.Count
    
        ActiveCell.Cells.Select
        ActiveCell.Range(Cell(1, 1), Cell(2, ColumnNum)).Select
        
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        Selection.Borders(xlInsideVertical).LineStyle = xlNone
        Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
        
        '文字列セルに変更
        Selection.NumberFormatLocal = "@"
        
        'セルをマージ
                'ActiveCell.Range(Cell(1, 1), Cell(2, ColumnNum)).Select
                'Selection.HorizontalAlignment = xlLeft
        
        '次セルを選択
        
        ActiveCell.Range(Cell(1, 1), Cell(2, ColumnNum)).Select
        ActiveCell.Offset(２, 0).Activate
    
    Next i

End Sub
