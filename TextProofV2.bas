Attribute VB_Name = "TextProof"
Public Sub �v���Z�X�ڍׂ𕶏͂��Z������()

    '���݂̈���͈͂�ݒ�
    '���݂͑I��͈͂�ݒ肷��
    Dim ws As Worksheet
    Set ws = ActiveSheet
    '�I��͈�
    Dim ActiveArea As String
    ActiveArea = ws.PageSetup.PrintArea
    'ActiveArea = Selection.Address
    ws.Range(ActiveArea).Activate
    
    '�ΏۃG���A���Z�b�g
    Dim TargetArea As Range
    'Set TargetArea = ws.Range(ActiveArea)
    Set TargetArea = Selection

    '�@�����@�\���C���X�^���X��
    Dim tp_Indent As TextProof_SerialNumberCreate
    Set tp_Indent = New TextProof_SerialNumberCreate

    
    '�C���f���g�A�Ԃ��쐬����
    Call tp_Indent.Init
    Call tp_Indent.�C���f���g�ɉ������A�Ԃ����o�͂���(TargetArea)
    

    MsgBox "���͍Z���̃��r���[���I���܂����I�:.߁R(�L��`�)��.:� �K"

End Sub


