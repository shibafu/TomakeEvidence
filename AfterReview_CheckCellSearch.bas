Attribute VB_Name = "AfterReview_CheckCellSearch"
'���r���[�Ń`�F�b�N���ꂽ�Z������������
'@Author ���V
'
Public Sub ���r���[�Ń`�F�b�N���ꂽ�Z�����T�[�`����()
'Workbook(ActiveWorkbook.Name).Active

Dim TargetWorkSheet As Worksheet
' �S�V�[�g�ɑ΂��čs��
For Each TargetWorkSheet In ActiveWorkbook.Sheets

    '�`�F�b�N���ꂽ�Z����I��������b�Z�[�W�E�B���h�E�ɕ\��
    Call �`�F�b�N���ꂽ�Z����I��������b�Z�[�W�E�B���h�E�ɕ\��(TargetWorkSheet)

Next TargetWorkSheet

MsgBox "�`�F�b�N���I�����܂����I"
End Sub

Private Sub �`�F�b�N���ꂽ�Z����I��������b�Z�[�W�E�B���h�E�ɕ\��(ws As Worksheet)

Dim PrArea As String '����͈�
Dim AreaColumn As Integer '����͈�
Dim AreaRow As Integer '����͈�

ws.Activate
'����͈͂��擾����
PrArea = ws.PageSetup.PrintArea
If PrArea = "" Then
    Exit Sub
End If
'����͈͂�I��
ws.Range(Replace(PrArea, "$", "")).Select

'�����G���A�񐔂��擾����
AreaColumn = Selection.Columns.Count
'�����G���A�s�����擾����
AreaRow = Selection.Rows.Count

'��ƍs�������[�v
For RowPointNm = 1 To AreaRow
    For ColumnPointNm = 1 To AreaColumn
    
    '�Z�����s���N�̎��A�Z�����擾
    If ws.Cells(RowPointNm, ColumnPointNm).Interior.Color = ColorConstants.vbMagenta Or _
    ws.Cells(RowPointNm, ColumnPointNm).Interior.Color = 16751103 Then
        '���b�Z�[�W�{�b�N�X�Ƀ��b�Z�[�W��\������
        MsgBox "�C���Z����������܂����I�F" & ws.Name & "�A�h���X�F" & ws.Cells(RowPointNm, ColumnPointNm).Address & vbCrLf _
        & "���e�F" & ws.Cells(RowPointNm, ColumnPointNm).Value
    End If
    Next ColumnPointNm
Next RowPointNm


End Sub
