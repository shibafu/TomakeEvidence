Declare Function OpenClipboard Lib "user32" (Optional ByVal hwnd As Long = 0) As Long
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function EmptyClipboard Lib "user32" () As Long

'�J��Ԃ��t���O
Public isLogging As Boolean
'�ς���s�����g���ĂȂ�
Public LocalOffset As Load

'�X�N���[���V���b�g���N���b�v�{�[�h�ɃR�s�[������
'�G�N�Z���Ɏ����Œ���t����

'�L���v�`���[���[�h���N������
Public Sub Capture()

    MsgBox "�G�r�f���X�̎����\��t�����J�n���܂�"
    
    'ESC�L�[�Œ�~
    Application.OnKey "{ESC}", "StopCapture"
    
    '�N���b�v�{�[�h����ɂ���
    a = OpenClipboard()
    a = EmptyClipboard()
    a = CloseClipboard()
    
    
    '�J��Ԃ��t���O���I���ɂ���
    isLogging = True
    
    '�摜���N���b�v�{�[�h�ɕۑ����ꂽ��N��
    If Application.ClipboardFormats(1) = xlClipboardFormatBitmap Then
    
        '�\��t��
        ActiveSheet.Paste
        '���݃Z������10�����Ɉړ�
        ActiveSheet.Offset(10, 0).Select
        
            '�N���b�v�{�[�h����ɂ���
            a = OpenClipboard()
            a = EmptyClipboard()
            a = CloseClipboard()
    
    End If
    
    '1�b�Ԋu Capture�֐����N������
    Application.OnTime Now + TimeValue("00:00:01"), "Capture", , isLogging
    
    Exit Sub

End Sub

'�L���v�`���[���[�h���~����
Public Sub StopCapture()

    'ESC�ւ̃L�[�o�^����������
    Application.OnKey "{ESC}", ""
    '�J��Ԃ��t���O���I�t��
    isLogging = False
    
    MsgBox "�L���v�`���̎擾���~���܂���"
    
End Sub
