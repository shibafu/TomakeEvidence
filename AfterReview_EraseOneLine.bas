Attribute VB_Name = "AfterReview_EraseOneLine"
'�C���O����
Dim ImprooveSentence As String

Public Sub �����ӏ��ɑŏ�����������()
Attribute �����ӏ��ɑŏ�����������.VB_ProcData.VB_Invoke_Func = "Y\n14"

    ' �C���O�̕��͂��`
    ThroughSentence = "�Ώۍ��ڂ̊���/�񊈐�"
    ' �ŏ����t���O
    EraseFlg = True

    ' �ŏ������J�n
    Dim StrikethroughStart As Integer
    ' �ŏ������I��
    Dim StrikethroughEnd As Integer
    
    ' �I��͈͂̕��A���[�v������
    For i = 1 To Selection.Rows.Count
        For j = 1 To Selection.Columns.Count
        
            ' �Z���^�[�Q�b�g�̃A�h���X�i��j
            Dim TargetPointRowAddress
            TargetPointRowAddress = Selection.Row + i - 1
            
            ' �Z���^�[�Q�b�g�̃A�h���X�i�s�j
            Dim TargetPointColumnAddress
            TargetPointColumnAddress = Selection.Column + j - 1
        
            If InStr(Cells(TargetPointRowAddress, TargetPointColumnAddress).Value, ThroughSentence) > 0 Then
                ' �ŏ������J�n������
                StrikethroughStart = InStr(Cells(TargetPointRowAddress, TargetPointColumnAddress).Value, ThroughSentence)
                ' �ŏ������I��������
                StrikethroughEnd = InStr(Cells(TargetPointRowAddress, TargetPointColumnAddress).Value, ThroughSentence) + _
                    Len(ThroughSentence) - 1
                ' �J�n�������ƁA��������̏��ɂ̂݁A�ŏ�����������I������
                Cells(TargetPointRowAddress, TargetPointColumnAddress). _
                Characters(StrikethroughStart, StrikethroughEnd).Font.Strikethrough = EraseFlg
            End If
        Next j
    Next i

End Sub
