Attribute VB_Name = "ReportDesignEdi_DispInsert"
Dim TARGET_CELL_JP_NAME_COLUMN
Dim TARGET_CELL_REMARK_COLUMN
Dim TARGET_CELL_REVISION_COLUMN
'���[�ׂ̗ɏC����}������
Public Sub DispInsert()

'�萔��������
TARGET_CELL_JP_NAME_COLUMN = 2
TARGET_CELL_REMARK_COLUMN = 12
TARGET_CELL_REVISION_COLUMN = 13

Dim CellRowIndex As Integer
CellRowIndex = 1 + 6

Do While ActiveSheet.Cells(CellRowIndex, TARGET_CELL_JP_NAME_COLUMN).Value <> ""

    '���K�\�����g���ăN���X����T�m
    Dim RegOb As Object
    Set RegOb = CreateObject("VBScript.RegExp")
    '���K�\�����g���ĒT�m����
    With RegOb
        .Pattern = "^[^%]+([�O-�X0-9]+)"
        .Global = True
    End With
    
        '���K�\�����������s
        Dim Matches
        Set Matches = RegOb.Execute(ActiveSheet.Cells(CellRowIndex, TARGET_CELL_JP_NAME_COLUMN).Value)
        
    '�p�^�[�������������珈�����s��
    If Matches.Count > 0 Then
        '������������������ăR�����g�����
        ActiveSheet.Cells(CellRowIndex, TARGET_CELL_REMARK_COLUMN).Value = Cells(CellRowIndex, TARGET_CELL_REMARK_COLUMN).Value & vbCrLf _
        & "Z,ZZ9"
        '�N���X�R�����g��ҏW
        ActiveSheet.Cells(CellRowIndex, TARGET_CELL_REVISION_COLUMN).Value = ActiveSheet.Cells(CellRowIndex, TARGET_CELL_REVISION_COLUMN).Value & _
        "2019/12/13�@���V�@�C���@���������iNo.8�j"
        ActiveSheet.Cells(CellRowIndex, TARGET_CELL_REVISION_COLUMN).VerticalAlignment = xlCenter
    End If

    CellRowIndex = CellRowIndex + 1
Loop
End Sub
