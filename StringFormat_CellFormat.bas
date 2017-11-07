Attribute VB_Name = "StringFormat_CellFormat"
Public Sub YokoOver()
Attribute YokoOver.VB_Description = "列をはみ出します"
Attribute YokoOver.VB_ProcData.VB_Invoke_Func = "L\n14"
'
' 横業 Macro
' 列をはみ出します
'
' Keyboard Shortcut: Ctrl+Shift+L
'
    With Selection
        .WrapText = False
        .AddIndent = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Application.OnKey "^L", "YokoOver"
End Sub

Public Sub フォントサイズ修正()
Attribute フォントサイズ修正.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' フォントサイズ修正 Macro
'
' Keyboard Shortcut: Ctrl+Shift+F
'

    With Selection.Font
        .Name = "メイリオ"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    With Selection
            .Borders.LineStyle = False
    End With
End Sub

'コピーした値を文字列に
Public Sub FormatChangeString()
Attribute FormatChangeString.VB_ProcData.VB_Invoke_Func = "Q\n14"

    With Selection
        .NumberFormatLocal = "@"
        .Borders.LineStyle = False
    End With
    
    
    Application.OnKey "^Q", "FormatChangeString"
End Sub
Public Sub BorderLineOn()
Attribute BorderLineOn.VB_ProcData.VB_Invoke_Func = "W\n14"
'
' 横枠修正 Macro
'
' Keyboard Shortcut: Ctrl+Shift+W

    With Selection
        .Borders.LineStyle = True
    End With
    
        Application.OnKey "^W", "BorderLineOn"
End Sub

