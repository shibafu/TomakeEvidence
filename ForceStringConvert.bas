Attribute VB_Name = "ForceStringConvert"

'コピーしたセルを強制的に文字列形式に変換する。
'そのシートのマクロにコピーして使う！

Private Sub Worksheet_Change(ByVal Target As Range)
 If Application.CutCopyMode = False Then Exit Sub
 Selection.NumberFormatLocal = "@"
End Sub



