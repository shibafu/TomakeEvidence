Attribute VB_Name = "CellTableMake"
' �e�[�u����}�ɂ��܂��f

Public Sub �Z�����c�̐}�ɂ���()
Attribute �Z�����c�̐}�ɂ���.VB_ProcData.VB_Invoke_Func = "w\n14"
'
' �e�[�u����}�ɂ��܂�
'
    '�J�n�ʒu
    RowAddress = Selection.Row
    ColumnAddress = Selection.Column
    '�I��͈�
    ColumnCount = Selection.Columns.Count - 1
    RowCount = Selection.Rows.Count


    ''��������
    Selection.Borders.LineStyle = xlNone

For i = 1 To RowCount
'�c��s��I��
    Range(Cells(RowAddress + i - 1, ColumnAddress), Cells(RowAddress + i - 1, ColumnAddress + ColumnCount)).Select

    '�I��͈͂�����
    'Selection.Merge
    ''��������
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

'�I��͈͂��A���}�[�W
Selection.UnMerge
    
End Sub
