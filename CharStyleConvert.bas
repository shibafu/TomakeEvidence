Attribute VB_Name = "CharStyleConvert"
Public Sub カンマ区切り()
Attribute カンマ区切り.VB_ProcData.VB_Invoke_Func = "K\n14"
'
' カンマ区切り Macro
'　文字列セルに対してのみ使用
'
' Keyboard Shortcut: Ctrl+Shift+K

    '文字列セルに変更
    Selection.NumberFormatLocal = "@"
    '現在の選択セル
    Dim rg As Range
    Set rg = Selection
    '代入する文字列
    Dim Result As String
    '
    Result = Format(rg.Value, "#,#")

    rg.Value = Result


End Sub
