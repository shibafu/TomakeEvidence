Attribute VB_Name = "Today"
Public Sub �����̓��t������()
Attribute �����̓��t������.VB_ProcData.VB_Invoke_Func = "T\n14"
'
' EraseLined Macro
' �ԐF�ɂ���
'
' Keyboard Shortcut: Ctrl+Shift+R



With ActiveCell
    .NumberFormatLocal = "@"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlCenter
End With
With ActiveCell.Font
    .Size = 10
    .Name = "���C���I"
    '.Name = "���S�V�b�N"
End With
ActiveCell.Value = ActiveCell.Value + " " + Format(Date, "yyyy/mm/dd") _
+ " ���V�@�C���i��������No." + CStr(�����������擾����()) + "�j"

End Sub

Private Function �����������擾����() As Integer

Dim i
Dim resultNm As Integer

Do While Worksheets("��������").Cells(7 + i, 4).Value <> ""
    
    '�ǂݍ��݃Z�����`����
    resultNm = Worksheets("��������").Cells(7 + i, 2).Value
    
    i = i + 1

 Loop

�����������擾���� = resultNm


End Function
