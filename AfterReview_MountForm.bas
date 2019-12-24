Attribute VB_Name = "AfterReview_MountForm"
' 金額を

Public Sub 金額の桁数を変更()
    ' 修正前の文章を定義
    ThroughSentence = "9"
    ' 修正後の文章を定義
    Dim fixedSentence As String
    fixedSentene = "15"

    ' 打消し線開始
    Dim StrikethroughStart As Integer
    ' 打消し線終了
    Dim StrikethroughEnd As Integer
    
    ' 選択範囲の分、ループさせる
    For i = 1 To Selection.Rows.Count
        For j = 1 To Selection.Columns.Count
        
            ' セルターゲットのアドレス（列）
            Dim TargetPointRowAddress
            TargetPointRowAddress = Selection.Row + i - 1
            
            ' セルターゲットのアドレス（行）
            Dim TargetPointColumnAddress
            TargetPointColumnAddress = Selection.Column + j - 1
        
            If InStr(Cells(TargetPointRowAddress, 5).Value, "金額") > 0 Then
                ' 打消し線開始文字数
                StrikethroughStart = InStr(Cells(TargetPointRowAddress, TargetPointColumnAddress).Value, ThroughSentence)
                ' 打消し線終了文字数
                StrikethroughEnd = InStr(Cells(TargetPointRowAddress, TargetPointColumnAddress).Value, ThroughSentence) + _
                    Len(ThroughSentence)
                ' 開始文字数と、狩猟文字数の所にのみ、打消し文字数を選択する
                Cells(TargetPointRowAddress, TargetPointColumnAddress). _
                Characters(StrikethroughStart, StrikethroughEnd).Font.Strikethrough = True
                
                ' 打消し線と、追加をする。
                Cells(TargetPointRowAddress, TargetPointColumnAddress).Value = CStr(Cells(TargetPointRowAddress, TargetPointColumnAddress).Value) + " " _
                 + fixedSentene
            End If
        Next j
    Next i
End Sub
