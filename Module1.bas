Attribute VB_Name = "Module1"
Sub 横業()
Attribute 横業.VB_Description = "列をはみ出します"
Attribute 横業.VB_ProcData.VB_Invoke_Func = "L\n14"
'
' 横業 Macro
' 列をはみ出します
'
' Keyboard Shortcut: Ctrl+Shift+L
'
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
