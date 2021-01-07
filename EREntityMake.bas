Attribute VB_Name = "EREntityMake"
Public Sub EREntityMake()
Attribute EREntityMake.VB_ProcData.VB_Invoke_Func = "p\n14"
'
' EREntityMake Macro
'
' Keyboard Shortcut: Ctrl+p
'
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.Left, Selection.Top, 160, 90).Select
    ' 枠線を赤色に
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
        .Transparency = 0
    End With
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    ' フォントに色を付ける
    With Selection.ShapeRange(1).TextFrame.Characters.Font
        .Color = msoThemeColorText1
    End With
    'PKを追加
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "PK"
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    ActiveSheet.Shapes.AddShape(msoShapeOval, 284.5, 32.5, 15.5, 18).Select
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    Range("G5").Select
End Sub
