Attribute VB_Name = "DataSizeAnalyze"
Public Sub DigiAnalySisExe()
Attribute DigiAnalySisExe.VB_ProcData.VB_Invoke_Func = "I\n14"
'�Z���̍ŏI�s���擾
    Dim n As Integer
    n = Cells(Rows.Count, "A").End(xlUp).Row
     
     
  Dim DigiResult As String
    '���������烋�[�v�̃X�^�[�g
    For i = 1 To n
    '�󔒂łȂ��Ƃ��̂ݎ��s

          '�t�H�[�}�b�g�𕶎���ɕύX
         Cells(i, 1).NumberFormatLocal = "@"
           '�Z���̒lA�񂩂當�����擾
         Dim CharSize As Long
         Dim RawChar As String
         RawChar = Cells(i, 1).Value
    
         '�o�C�g���擾
         CharSize = LenB(StrConv(RawChar, vbFromUnicode))
    
         Cells(i, 2).Value = CharSize
         
         Dim HankakuSize As Long
         Dim ZenkakuSize As Long

         For k = 1 To Len(RawChar)
                '�ꕶ���𔼊p���ǂ�������
                 If Mid(RawChar, k, 1) Like "[- �-� 0-9a-zA-Z]" Then
                 HankakuSize = HankakuSize + 1
                 Else
                 '����ȊO�͑S�p
                 ZenkakuSize = ZenkakuSize + 1
                 End If
         Next k


        '�S�p�̃T�C�Y�����L�ɏo�͂����
        Cells(i, 3).Value = ZenkakuSize
        Cells(i, 4).Value = HankakuSize
         

    
        '������
         RawChar = ""
         CharSize = 0
         HankakuSize = 0
         ZenkakuSize = 0
     

    
    Next i
    '�������܂ł����[�v

MsgBox "�����̕��͂��I���܂���"

        Application.OnKey "^I", "DigiAnalySisExe"

End Sub
