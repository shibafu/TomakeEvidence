Attribute VB_Name = "MakeCallOut"
Public Sub 吹き出しを付ける()
Attribute 吹き出しを付ける.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' 吹き出しを付ける Macro
'
' Keyboard Shortcut: Ctrl+Shift+F
'

    Dim RectWidth
    RectWidth = 190
    Dim RectHeight
    RectHeight = 70

    Dim RectangleName As String
    Dim LineName As String
    '四角形を付ける
    Set Rectangle = ActiveSheet.Shapes.AddShape(msoShapeRectangle, Selection.Left, Selection.Top, RectWidth, RectHeight)
    With Rectangle.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent6
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0.8000000119
        .Transparency = 0
        .Solid
    End With
    
    '四角の中の文字を黒色にする
    With Rectangle.TextFrame2.TextRange.Characters.Font.Fill.ForeColor
        .RGB = vbBlack
    End With

    'フリーフォームを付ける
    Dim LineObj
    Set LineObj = ActiveSheet.Shapes.BuildFreeform(msoEditingCorner, Selection.Left - 10, Selection.Top + RectHeight / 2)
    With LineObj
        .AddNodes msoSegmentLine, msoEditingAuto, Selection.Left - 60, Selection.Top + RectHeight / 2
        .AddNodes msoSegmentLine, msoEditingAuto, Selection.Left - 160, Selection.Top + RectHeight / 2 + 50
    End With
    
    Dim LineShape
    Set LineShape = LineObj.ConvertToShape
    
    LineShape.Line.Weight = 2
    '線と資格をグループ化
    LineShape.Select
    Rectangle.Select False
    
    Selection.ShapeRange.Group.Name = Rectangle.Name & "_" & LineShape.Name
End Sub

