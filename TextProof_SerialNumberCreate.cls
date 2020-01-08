VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextProof_SerialNumberCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'�C���f���g�J�E���^�[�P
Dim CounterNm1st As Integer
'�C���f���g�J�E���^�[�Q
Dim CounterNm2nd As Integer
'�C���f���g�J�E���^�[�R
Dim CounterNm3rd As Integer
'�C���f���g�J�E���^�[�S
Dim CounterNm4th As Integer
'�C���f���g�J�E���^�[�T
Dim CounterNm5th As Integer
'�C���f���g�J�E���^�[�U
Dim CounterNm6th As Integer
'�C���f���g��A��l
Dim FIRST_INDENT_ROW As Integer

'����������
Public Sub Init()

CounterNm1st = 0
CounterNm2nd = 0
CounterNm3rd = 0
CounterNm4th = 0
CounterNm5th = 0
CounterNm6th = 0
FIRST_INDENT_ROW = 3

End Sub


Public Sub �C���f���g�ɉ������A�Ԃ����o�͂���(ByRef SelectionArea As Range)

    For i = 0 To SelectionArea.Rows.Count - 1
        For j = 0 To SelectionArea.Columns.Count - 1
            '���K�\�����g���ăN���X����T�m
            Dim IndentPattern As Object
            Set IndentPattern = CreateObject("VBScript.RegExp")
            '���̃p�^�[���̐��K�\���Ăяo��������
            'Set RE = CreateObject("VBScript.RegExp")
            'X-X-X�������L�q���Ƃ���������������
            With IndentPattern
                .Pattern = "^([1-9\-]+\.)(.*)"
                .Global = True
            End With
    
            '���K�\�����������s
            Set Matches = IndentPattern.Execute(Cells(SelectionArea.Row + i, SelectionArea.Column + j).Value)
        
            Dim MatchedIndent As String
            MatchedIndent = ""
            Dim OtherSentence As String
            OtherSentence = ""
            '�p�^�[�������������珈�����s��
            If Matches.Count > 0 Then
                '������������������ăR�����g�����
                MatchedIndent = Matches.Item(0).SubMatches.Item(0)
                OtherSentence = Matches.Item(0).SubMatches.Item(1)
            End If


            '�}�b�`�����Ƃ��Ɏ��s
            If MatchedIndent <> "" Then
                '�}�b�`�����Ƃ��Z���̒��g����ɂ���
                Cells(SelectionArea.Row + i, SelectionArea.Column + j).Value = ""
               'X.�̎��A�ŏ��̏��������͂��đ��J�E���^�[��������
                If Len(MatchedIndent) = 2 Then
                  
                    CounterNm1st = CounterNm1st + 1
                    '�Z���ɏo��
                    Cells(SelectionArea.Row + i, FIRST_INDENT_ROW).Value = CStr(CounterNm1st) & "." & OtherSentence
            
                    CounterNm2nd = 0
                    CounterNm3rd = 0
                    CounterNm4th = 0
                    CounterNm5th = 0
                    CounterNm6th = 0
                    'X-X.�̎��A
                ElseIf Len(MatchedIndent) = 4 Then
        
                    CounterNm2nd = CounterNm2nd + 1
                    '�Z���ɏo��
                    Cells(SelectionArea.Row + i, FIRST_INDENT_ROW + 1).Value = CStr(CounterNm1st) & "-" & CStr(CounterNm2nd) & "." & OtherSentence
            
                    CounterNm3rd = 0
                    CounterNm4th = 0
                    CounterNm5th = 0
                    CounterNm6th = 0
                    'X-X-X.�̎��A�ŏ��̏��������͂��đ��J�E���^�[��������
                ElseIf Len(MatchedIndent) = 6 Then
        
                    CounterNm3rd = CounterNm3rd + 1
                    '�Z���ɏo��
                    Cells(SelectionArea.Row + i, FIRST_INDENT_ROW + 2).Value = CStr(CounterNm1st) & "-" & CStr(CounterNm2nd) _
                    & "-" & CStr(CounterNm3rd) & "." & OtherSentence
            
                    CounterNm4th = 0
                    CounterNm5th = 0
                    CounterNm6th = 0
                    'X-X-X.�̎��A�ŏ��̏��������͂��đ��J�E���^�[��������
                ElseIf Len(MatchedIndent) = 8 Then
        
                    CounterNm4th = CounterNm4th + 1
                    '�Z���ɏo��
                    Cells(SelectionArea.Row + i, FIRST_INDENT_ROW + 3).Value = CStr(CounterNm1st) & "-" & CStr(CounterNm2nd) _
                    & "-" & CStr(CounterNm3rd) & "-" & CStr(CounterNm4th) & "." & OtherSentence
            
                    CounterNm5th = 0
                    CounterNm6th = 0
                'X-X-X-X.�̎��A�ŏ��̏��������͂��đ��J�E���^�[��������
                ElseIf Len(MatchedIndent) = 10 Then
        
                    CounterNm5th = CounterNm5th + 1
                    '�Z���ɏo��
                    Cells(SelectionArea.Row + i, FIRST_INDENT_ROW + 4).Value = CStr(CounterNm1st) & "-" & CStr(CounterNm2nd) _
                    & "-" & CStr(CounterNm3rd) & "-" & CStr(CounterNm4th) & "-" _
                    & CStr(CounterNm5th) & "." & OtherSentence
            
                    CounterNm6th = 0
                    'X-X-X-X.�̎��A�ŏ��̏��������͂��đ��J�E���^�[��������
                ElseIf Len(MatchedIndent) = 12 Then
        
                    CounterNm6th = CounterNm6th + 1
                    '�Z���ɏo��
                    Cells(SelectionArea.Row + i, FIRST_INDENT_ROW + 5).Value = CStr(CounterNm1st) & "-" & CStr(CounterNm2nd) _
                    & "-" & CStr(CounterNm3rd) & "-" & CStr(CounterNm4th) & "-" _
                    & CStr(CounterNm5th) & "-" & CStr(CounterNm6th) & "." & OtherSentence
            
                End If
             End If
        Next j
    Next i
    

End Sub