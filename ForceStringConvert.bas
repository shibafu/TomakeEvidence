Attribute VB_Name = "ForceStringConvert"

'�R�s�[�����Z���������I�ɕ�����`���ɕϊ�����B
'���̃V�[�g�̃}�N���ɃR�s�[���Ďg���I

Private Sub Worksheet_Change(ByVal Target As Range)
 If Application.CutCopyMode = False Then Exit Sub
 Selection.NumberFormatLocal = "@"
End Sub



