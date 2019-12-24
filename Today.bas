Attribute VB_Name = "Today"
Public Sub 今日の日付を入れる()
Attribute 今日の日付を入れる.VB_ProcData.VB_Invoke_Func = "T\n14"
'
' EraseLined Macro
' 赤色にする
'
' Keyboard Shortcut: Ctrl+Shift+R



With ActiveCell
    .NumberFormatLocal = "@"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlCenter
End With
With ActiveCell.Font
    .Size = 10
    .Name = "メイリオ"
    '.Name = "游ゴシック"
End With
ActiveCell.Value = ActiveCell.Value + " " + Format(Date, "yyyy/mm/dd") _
+ " 野澤　修正（改訂履歴No." + CStr(改訂履歴を取得する()) + "）"

End Sub

Private Function 改訂履歴を取得する() As Integer

Dim i
Dim resultNm As Integer

Do While Worksheets("改訂履歴").Cells(7 + i, 4).Value <> ""
    
    '読み込みセルを定義する
    resultNm = Worksheets("改訂履歴").Cells(7 + i, 2).Value
    
    i = i + 1

 Loop

改訂履歴を取得する = resultNm


End Function
