Attribute VB_Name = "ColumnTextInput"
Public Sub DataInput()
    '�o�͂����e�L�X�g�t�@�C���𕡐��I��
    Inputs = get_files()
    
    Dim ColumnTextNum As Integer
    Dim Buffer As String
    
    '�J�n����w��,�f�t�H���g��B�񂩂�
    ColumnNum = 2
    
    '�ǂݍ��݃t�@�C�����X�g���J�����[�v
    For Each Input_data In Inputs
        If Input_data <> "" Then
            Open Input_data For Input As #1
            
            Dim RawNum As Integer
            RawNum = 1
            
            '�ǂݍ��݃e�L�X�g�擾���[�v
            '�s���ɃZ����ł�����
            
            Do Until EOF(1)
                Line Input #1, Buffer
                ActiveSheet.Cells(RawNum, ColumnNum).Vaue = Buffer
                RawNum = RawNum + 1
            Loop
            
            Close #1
            '�������炷
        ColumnNum = ColumnNum + 1
        End If
    Next Input_data
End Sub

Function get_files()
    Dim i As Integer
    Dim InputTextFiles(99)
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "�e�L�X�g�t�@�C����I��"
        .AllowMultiSelect = True
        If .Show - 1 Then
            '�����I���������t�@�C���ɔz���}������
            For i = 1 To .SelectItems.Count
                InputTextFiles(i) = SelectItems(i)
            Next i
        Else
            Exit Function
        End If
    End With
    
    get_files = InputTextFiles

End Function

