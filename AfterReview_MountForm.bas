Attribute VB_Name = "AfterReview_MountForm"
' ���z��

Public Sub ���z�̌�����ύX()
    ' �C���O�̕��͂��`
    ThroughSentence = "9"
    ' �C����̕��͂��`
    Dim fixedSentence As String
    fixedSentene = "15"

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
        
            If InStr(Cells(TargetPointRowAddress, 5).Value, "���z") > 0 Then
                ' �ŏ������J�n������
                StrikethroughStart = InStr(Cells(TargetPointRowAddress, TargetPointColumnAddress).Value, ThroughSentence)
                ' �ŏ������I��������
                StrikethroughEnd = InStr(Cells(TargetPointRowAddress, TargetPointColumnAddress).Value, ThroughSentence) + _
                    Len(ThroughSentence)
                ' �J�n�������ƁA��������̏��ɂ̂݁A�ŏ�����������I������
                Cells(TargetPointRowAddress, TargetPointColumnAddress). _
                Characters(StrikethroughStart, StrikethroughEnd).Font.Strikethrough = True
                
                ' �ŏ������ƁA�ǉ�������B
                Cells(TargetPointRowAddress, TargetPointColumnAddress).Value = CStr(Cells(TargetPointRowAddress, TargetPointColumnAddress).Value) + " " _
                 + fixedSentene
            End If
        Next j
    Next i
End Sub
