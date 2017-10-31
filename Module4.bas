Attribute VB_Name = "Module4"
Sub set_all()
     fns = get_filenames()
     '配列の文字数分実行
     For Each fn In fns
         If fn <> "" Then
             Workbooks.Open Filename:=fn
             '真ん中"&[ページ番号]/&[総ページ数]"を設定
               Call set_header_footer("", "", "", "", "&P" & "/" & "&N", "")
                ActiveWorkbook.Save
                MsgBox ("(^-^)b")
                ActiveWindow.Close
            End If
         Next fn
    End Sub

Function get_filenames()
'ファイル名を取得する
    Dim i             As Integer
    Dim ars(99)
    
     With Application.FileDialog(msoFileDialogFilePicker)
         .Title = "フォルダを選択"
         .AllowMultiSelect = True
         If .Show = -1 Then
         
            '複数選択したファイルを配列に挿入する
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
'ヘッダーフッターをセットする
'全てのワークシートに同ヘッダーフッターをセットする
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

