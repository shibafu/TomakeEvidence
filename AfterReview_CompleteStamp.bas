Attribute VB_Name = "AfterReview_CompleteStamp"
Public Sub ��_����_�̕�����������()
Attribute ��_����_�̕�����������.VB_ProcData.VB_Invoke_Func = "C\n14"

    '���K�\�����g���ăN���X����T�m
    Dim RegOb As Object
    Set RegOb = CreateObject("VBScript.RegExp")
    '���K�\�����g���ĒT�m����
    With RegOb
        .Pattern = "(.*)�y��_([0-9\/])+�z"
        .Global = True
    End With
    
    '���K�\�����������s
    Dim Matches
    Set Matches = RegOb.Execute(ActiveCell.Value)

    '���ɍσ}�[�N������ꍇ
    If Matches.Count > 0 Then
        '�σ}�[�N�������āA�l�����Ȃ���
        ActiveCell.Value = Matches.Item(0).SubMatches(0)
    Else
        '���݂̒l�̌��ɍσ}�[�N������
        ActiveCell.Value = ActiveCell.Value & "�y��_" & Date & "�z"
    End If

End Sub


