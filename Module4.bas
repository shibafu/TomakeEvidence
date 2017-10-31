Attribute VB_Name = "Module4"
Sub set_all()
     fns = get_filenames()
     '�z��̕����������s
     For Each fn In fns
         If fn <> "" Then
             Workbooks.Open Filename:=fn
             '�^��"&[�y�[�W�ԍ�]/&[���y�[�W��]"��ݒ�
               Call set_header_footer("", "", "", "", "&P" & "/" & "&N", "")
                ActiveWorkbook.Save
                MsgBox ("(^-^)b")
                ActiveWindow.Close
            End If
         Next fn
    End Sub

Function get_filenames()
'�t�@�C�������擾����
    Dim i             As Integer
    Dim ars(99)
    
     With Application.FileDialog(msoFileDialogFilePicker)
         .Title = "�t�H���_��I��"
         .AllowMultiSelect = True
         If .Show = -1 Then
         
            '�����I�������t�@�C����z��ɑ}������
            For i = 1 To .SelectedItems.Count
             ars(i) = .SelectedItems(i)
            Next i
         
         Else
             Exit Function
         End If
     End With

     get_filenames = ars
End Function
    
Function set_header_footer(lh, ch, rh, lf, cf, rf)
'�w�b�_�[�t�b�^�[���Z�b�g����
'�S�Ẵ��[�N�V�[�g�ɓ��w�b�_�[�t�b�^�[���Z�b�g����
     For Each ws In Worksheets
         With ws.PageSetup
             .LeftHeader = lh
             .CenterHeader = ch
             .RightHeader = rh
             .LeftFooter = lf
             .CenterFooter = cf
             .RightFooter = rf
         End With
     Next
End Function

