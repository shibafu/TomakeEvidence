Attribute VB_Name = "SheetRename"
Public Sub �V�[�g���A�ԐU�蕪��()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    '�������
    Dim RenameResult As Boolean
    RenameResult = SheetRenameProc(wb)

    MsgBox "�V�[�g���ϊ���Ƃ��I���܂����:.߁R(�L��`�)��.:� �K"

End Sub

'�V�[�g�����֐�
Function SheetRenameProc(wbObject As Workbook) As Boolean

    Dim SheetCounter As Integer
    SheetCounter = 1
       '�V�[�g���������J�n
   For Each sheet In wbObject.Sheets
           '���K�\����錾����
        Dim RegExp As VBScript_RegExp_55.RegExp
        Set RegExp = New VBScript_RegExp_55.RegExp
        With RegExp
            .Pattern = "[0-9]+\.(.+)"  '���ɃC���f�b�N�X���t�����Ă��邩�`�F�b�N
            .Global = True
        End With

        '���K�\�����s
        Dim Results As MatchCollection
        Set Results = RegExp.Execute(sheet.Name)
    
        '�q�b�g�����ꍇ�A�f�[�^�𐶐�����
        If Results.Count > 0 Then

            '���ʂ��Z���ɑ��
            Dim SheetNameWithourIndex As String
            SheetNameWithourIndex = Results.Item(0).SubMatches.Item(0)

            '���݂̃V�[�g���̓��ɐ�����t����
            sheet.Name = CStr(SheetCounter) + "." + SheetNameWithourIndex
        Else

            '���݂̃V�[�g���̓��ɐ�����t����
            sheet.Name = CStr(SheetCounter) + "." + sheet.Name
        End If
       
       SheetCounter = SheetCounter + 1

   Next sheet
End Function

