Attribute VB_Name = "Module3"
Sub 画像拡縮_大()
Attribute 画像拡縮_大.VB_ProcData.VB_Invoke_Func = "M\n14"
'
' 画像拡縮 Macro
'
' Keyboard Shortcut: Ctrl+Shift+M
'
    Selection.ShapeRange.ScaleWidth 0.835, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 0.835, msoFalse, msoScaleFromTopLeft

'   元に戻す倍率 1/ 0.835
'    Selection.ShapeRange.ScaleWidth 1.19760479042, msoFalse, msoScaleFromTopLeft
'    Selection.ShapeRange.ScaleHeight 1.19760479042, msoFalse, msoScaleFromTopLeft
    
    Selection.ShapeRange.ZOrder msoSendToBack
End Sub
Sub フォントサイズ修正()
Attribute フォントサイズ修正.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' フォントサイズ修正 Macro
'
' Keyboard Shortcut: Ctrl+Shift+F
'

    With Selection.Font
        .Name = "游ゴシック"
        .Size = 10
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
    With Selection.Font
        .Name = "游ゴシック"
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
End Sub
Sub 横幅左修正()
Attribute 横幅左修正.VB_ProcData.VB_Invoke_Func = "r\n14"
'
' Macro8 Macro
'
' Keyboard Shortcut: Ctrl+r
'
        Selection.Left = 55
End Sub
Sub 画像拡縮_小()
Attribute 画像拡縮_小.VB_ProcData.VB_Invoke_Func = "H\n14"
'
' Macro9 Macro
'
' Keyboard Shortcut: Ctrl+Shift+H
' よく使うサイズ0.673　0.635
' 今回のサイズ 0.625
'中文字サイズ0.66
'小文字サイズ 0.69

    Selection.ShapeRange.ScaleWidth 0.64, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 0.64, msoFalse, msoScaleFromTopLeft
    
    Selection.ShapeRange.ZOrder msoSendToBack
End Sub
