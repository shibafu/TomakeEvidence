' �e�[�u����}�ɂ��܂��f

Public Sub �Z�����c�̐}�ɂ���()
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
    '�I��͈͂�����
    'Range(Cells(RowAddress + i - 1, ColumnAddress), Cells(RowAddress + i - 1, ColumnAddress + ColumnCount)).Merge
    ''��������
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