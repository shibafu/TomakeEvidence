Attribute VB_Name = "DummyData"
Public Sub DummyDataGenerate()
Attribute DummyDataGenerate.VB_ProcData.VB_Invoke_Func = "O\n14"

'�Z���̍ŏI�s���擾
    Dim n As Integer
    n = Cells(Rows.Count, "A").End(xlUp).Row
     
     
  Dim DummyDataNum As String
  Dim DummyDataChr As String
    '���������烋�[�v�̃X�^�[�g
    For i = 1 To n
    
    

    '�Z���̒lA�񂩂���擾
    Dim Ketasu As Integer
    Ketasu = Cells(i, 1).Value
    
        '�����̃_�~�[�f�[�^����
       '�_�~�[�f�[�^�𐶐�
        For j = 1 To Ketasu
         Dim NextNum As String
         NextNum = j Mod 10
         DummyDataNum = DummyDataNum + NextNum
        Next j
        
    '�l��B��ɑ��
    Cells(i, 2).Value = DummyDataNum
    
    '�S�p�f�[�^����
    Dim ZenkakuResult As String
    ZenkakuResult = LeftB(StrConv(DummyDataNum, vbWide), Len(DummyDataNum))
    Cells(i, 3).Value = ZenkakuResult
        
        '���������f�[�^����
        For j = 1 To Ketasu
         Dim DigiNum As String
         If j = 1 Then
            DigiNum = "1"
         Else
            DigiNum = DigiNum + "0"
         End If
        Next j
    Cells(i, 4).Value = DigiNum
        
    '������
    DummyDataNum = ""
    ZenkakuResult = ""
    DigiNum = ""
    
        '�����̃_�~�[�f�[�^����
       '�_�~�[�f�[�^�𐶐�
        For j = 1 To Ketasu
         Dim NextChr As String
         NextChr = "N"
         DummyDataChr = DummyDataChr + NextChr
        Next j
        
    '�l��J��ɑ��
    Cells(i, 5).Value = DummyDataChr
    DummyDataChr = ""
    Next i
    '�������܂ł����[�v

MsgBox "�_�~�[�f�[�^�̐������I���܂���"

        Application.OnKey "^O", "DummyDataGenerate"

End Sub



