Attribute VB_Name = "ColumnTextOutput"
'���[�N�V�[�g�̍��ڗ���e�L�X�g�Ɏ����o�͂���'

Public Sub DataOutputText()

'�I�𒆂̃��[�N�V�[�g��I������B
Dim ws As Worksheet
Set ws = ActiveSheet

'�t�H���_�_�C�A���O���J���A�t�H���_�p�X���擾����
Dim Path As String
Path = get_folder()

'�t�@�C���쐬�I�u�W�F�N�g���쐬����
Dim ObjFso As Object
Set ObjFso = CreateObject("Scripting.FileSystemObject")

'�������ރt�@�C�����쐬����
Dim OutputFile As String

ObjFso.CreateTextFile (Path & "\" & ws.Name & ".txt")
Output = Path & "\" & ws.Name & ".txt"

Open OutputFile For Output As #1

'�V�[�g���������o��
    Print #1, ws.Name
    Print #1, '\r\n'

'6�Ԗڂ̗���e�L�X�g�ɏ����o��
Dim i As Long
i = 1

Do While ws.Cells(i + 10, 6).Value <> ""
    Print #1 ws.Cells(i + 10, 6).Value
    i = i + 1
Loop

'�e�L�X�g�����
Close #1
'�������b�Z�[�W��\��
MsgBox "�G���e�B�e�B�̈ꗗ�o�͂��I���܂���"

End Sub

Function get_folder() As String
    Dim i As Integer
    Dim Path As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "�o�̓t�H���_�[��I��"
    If Show = -1 Then
        '�I�������t�H���_�p�X�𕶎���ɑ}������
        Path = .SelectedItem(1)
    Else
        Exit Function
    End If
    
    End With
    
    get_folder = Path

End Function

