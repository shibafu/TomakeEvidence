Attribute VB_Name = "ReviewDesigin_CursorSetting"
Public Sub �V�[�g������ɐݒ肷��()
Attribute �V�[�g������ɐݒ肷��.VB_ProcData.VB_Invoke_Func = "U\n14"

Dim sheet As Worksheet

'�e�V�[�g�̃V�[�g�̉E��������ɂ���
For Each sheet In ActiveWorkbook.Sheets

    sheet.Activate
    sheet.Cells(1, 1).Activate

Next sheet

'���������̈�ԉ���I��
Dim i
Dim RevisionName As String
RevisionName = "��������"

Dim ws As Worksheet
Dim Revisionflag As Boolean
For Each ws In ActiveWorkbook.Sheets
    If ws.Name = RevisionName Then
        Revisionflag = True
    End If
Next ws

If Revisionflag Then
    Do While Worksheets(RevisionName).Cells(7 + i, 4).Value <> ""
        '�ǂݍ��݃Z�����`����
        i = i + 1
    Loop
    Worksheets(RevisionName).Activate
    ActiveWorkbook.Sheets(RevisionName).Cells(7 + i - 1, 4).Activate
End If

MsgBox "�J�[�\���ݒ肪�������܂����I"
End Sub
