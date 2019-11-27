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
End With
ActiveCell.Value = Format(Date, "yyyy/mm/dd")

End Sub
