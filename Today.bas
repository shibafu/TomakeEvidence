Attribute VB_Name = "Today"
Public Sub �����̓��t������()
Attribute �����̓��t������.VB_ProcData.VB_Invoke_Func = "T\n14"
'
' EraseLined Macro
' �ԐF�ɂ���
'
' Keyboard Shortcut: Ctrl+Shift+R

With ActiveCell
    .NumberFormatLocal = "@"
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlCenter
End With
With ActiveCell.Font
    .Size = 10
    .Name = "���C���I"
End With
ActiveCell.Value = Format(Date, "yyyy/mm/dd")

End Sub
