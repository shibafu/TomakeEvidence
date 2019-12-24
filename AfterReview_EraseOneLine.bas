Attribute VB_Name = "AfterReview_EraseOneLine"
'修正前文章
Dim ImprooveSentence As String

Public Sub 改訂箇所に打消し線を引く()
Attribute 改訂箇所に打消し線を引く.VB_ProcData.VB_Invoke_Func = "Y\n14"

    ' 修正前の文章を定義
    ThroughSentence = "対象項目の活性/非活性"
    ' 打消しフラグ
    EraseFlg = True

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
        
            If InStr(Cells(TargetPointRowAddress, TargetPointColumnAddress).Value, ThroughSentence) > 0 Then
                ' 打消し線開始文字数
                StrikethroughStart = InStr(Cells(TargetPointRowAddress, TargetPointColumnAddress).Value, ThroughSentence)
                ' 打消し線終了文字数
                StrikethroughEnd = InStr(Cells(TargetPointRowAddress, TargetPointColumnAddress).Value, ThroughSentence) + _
                    Len(ThroughSentence) - 1
                ' 開始文字数と、狩猟文字数の所にのみ、打消し文字数を選択する
                Cells(TargetPointRowAddress, TargetPointColumnAddress). _
                Characters(StrikethroughStart, StrikethroughEnd).Font.Strikethrough = EraseFlg
            End If
        Next j
    Next i

End Sub
