Dim BufferAll() As String

Public Sub ������Java�t�@�C���𐳋K�\���ŕҏW����()

    '�o�͂����e�L�X�g�t�@�C���𕡐��I��
    Inputs = get_files()
    
    '�t�@�C���쐬�I�u�W�F�N�g���쐬����
    Dim ObjFso As Object
    Set ObjFso = CreateObject("Scripting.FileSystemObject")

    '�ǂݍ��݃t�@�C�����X�g���J�����[�v
    
    For Each Input_data In Inputs
    If Input_data <> "" Then
    
        '�ǂݍ��݃t�@�C�����J��
        Open Input_data For Input As #1
    
        '�������݃t�@�C�����J��
        ObjFso.CreateTextFile (Input_data & ".new")
        OutputFile = Input_date & ".new"

        '�ҏW�������Ɋg���q��Java�ɖ߂��R�[�h
        'ObjFso.CreateTextFile (Mid(Input_data, 1, Len(Input_data) - 4))
        'OutputFile = Mid(Input_data, 1, Len(Input_data) - 4)

         Open OutputFile For Output As #2
         
         '�ǂݍ��݃e�L�X�g�擾���[�v
         '�s���ɃZ����ǂݍ���
         
         '�ǂݍ��݃��[�v
         Dim i As Integer
         Do Until EOF(1)
            ReadJavaFile (i)
            i = i + 1
         Loop
         
         i = 0
         
         '�������݃��[�v
         Dim j As Integer
         For Each Buf In BufferAl
          a = EditJavaFile(j, CStr(Buf))
          j = j + 1
         Next Buf
         
         j = 0
         
         '�t�@�C�������
         Close #1
         Close #2
    End If
    Next Input_data
         MsgBox "������Java�t�@�C���̕ҏW���I���܂���"
         
End Sub

'�t�@�C���I���E�B���h�E���J��
Private Function get_files()

    Dim i As Integer
    Dim InputTextFiles(200)
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Java�t�@�C����I�����Ă�������(�����I����)"
        .AllowMultiSelect = True
        
        If .Show = -1 Then
            '�����I�������t�@�C���ɔz���}������
            For i = 1 To .SelectedItems.Count
                InputTextFiles(i) = .SelectedItems(i)
            Next i
        Else
            Exit Function
        End If
        
        get_files = InputTextFiles

End Function

'�ǂݍ���Ńo�b�t�@�[�ɕۑ�
Private Function ReadJavaFile(i As Integer)

    '�ۑ��o�b�t�@�[���Ē�`
    '�����̕ϐ������̂܂܂�i�̐������z����`���Ȃ���
    ReDim Preserve BufferAll(i)
    '��s���Ƃɓǂݍ���
    Line Input #1, Buffer
    BufferAll(i) = CStr(Buffer)

End Function

'JavaFile��ҏW����
Private Function EditJavaFile(j As Integer, Buffer As String)

'���L�֐��œǂݍ��񂾃e�L�X�g��ҏW�@�g��������
'Buffer = HeaderEdit(Buffer)
'Buffer = ClassJavaDocComment(Buffer)

    'Open�����t�@�C���ɓǂݍ��񂾃o�b�t�@�[��ҏW���ď�������
    Print #2, Buffer
    
End Function
End Function
'�p�b�P�[�W�̏�Ƀw�b�_�[��t���ĕԋp

Private Function HeaderEdit(RawText As String)

    Dim StrAppend(4) As String
    StrAppend(0) = "/**" & vbCrLf
    StrAppend(1) = " *" & vbCrLf
    StrAppend(2) = " * �R�����g�ł�" & vbCrLf
    StrAppend(3) = " *XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & vbCrLf
    StrAppend(4) = " */" & vbCrLf

    'LIKE���g���������܂������BVBA��LIKE�����ł��u���K�\���v�ɋ߂��������o����
    '�p�b�P�[�W�̒P�ꂪ���������炻�̍s��ҏW
    If RawText Like "package *" Then
        RawText = StrAppend(0) & vcrlf & StrAppend(1) & vcrlf & StrAppend(2) & vcrlf & StrAppend(3) & vcrlf & _
        StrAppend(4) & vcrlf & RawText
    End If
    
    HeaderEdit = RawText
    
End Function
'�N���X�̏��JavaDoc�R�����g��t���ĕԋp

Private Function HeaderEdit(RawText As String) As String


Dim StrAppend(4) As String
StrAppend(0) = "/**" & vbCrLf
StrAppend(1) = " *"
StrAppend(2) = " * @author"
StrAppend(3) = " *" & vbCrLf
StrAppend(4) = " */" & vbCrLf


    '���K�\�����g���ăN���X����T�m
    Dim IndentFirstPattern As Object
    Set IndentFirstPattern = New VBScriptRegExp_55.RegExp
    '���K�\�����g���Ĉ�ԏ��߂̃N���X����
    '�uTestCase<4���̐���>�v�Ƃ����p�^�[��������������T�m����
    With IndentFirstPattern
        .Pattern = " (TestCase[0-9] {4})"
        .Global = True
    End With
    
        '���K�\�����������s
        Dim FirstMatches As VBScript_RegExp_55.MatchCollection
        Set FirstMatches = IndentFirstPattern.Execute(RawText)
        
    '�p�^�[�������������珈�����s��
    If FirstMatches.Count > 0 Then
        '������������������ăR�����g�����
        StrAppend(1) = StrAppend(1) & "�e�X�g�N���X�@" & FirstMatches.Item(0) & vbCrLf
        '�N���X�R�����g��ҏW
        RawText = StrAppend(0) & StrAppend(1) & StrAppend(2) & StrAppend(3) & StrAppend(4) & RawText
    End If

'�R�����g��ԋp
ClassJavaDocComment = RawText

End Function

