Attribute VB_Name = "EROneToNoneMany"
Public Sub OneToMany()
'
' OneToOn Macro
'
'
    ' 選択せるの座標を記録
    Dim SelectionTopXPosition As Integer
    SelectionTopXPosition = Selection.Left
    Dim SelectionTopYPosition As Integer
    SelectionTopYPosition = Selection.Top
    ' 線を追加
    Dim Select1 As Shape
    Set Select1 = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, _
    SelectionTopXPosition, SelectionTopYPosition, SelectionTopXPosition + 264, SelectionTopYPosition)
    
    Select1.Line.ForeColor.RGB = RGB(0, 0, 0)
    Dim Select2 As Shape
    Set Select2 = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, _
    SelectionTopXPosition + 10, SelectionTopYPosition - 10, SelectionTopXPosition + 10, SelectionTopYPosition + 10)
    Select2.Line.ForeColor.RGB = RGB(0, 0, 0)
    
    '足を書く
    Dim Select3_1 As Shape
    Set Select3_1 = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, _
    SelectionTopXPosition + 241, SelectionTopYPosition, SelectionTopXPosition + 256, SelectionTopYPosition + 15)
    Select3_1.Line.ForeColor.RGB = RGB(0, 0, 0)
    
    Dim Select3_2 As Shape
    Set Select3_2 = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, _
    SelectionTopXPosition + 241, SelectionTopYPosition, SelectionTopXPosition + 256, SelectionTopYPosition - 15)
    Select3_2.Line.ForeColor.RGB = RGB(0, 0, 0)
       
    Dim Select3 As Shape
    Set Select3 = ActiveSheet.Shapes.AddShape(msoShapeOval, SelectionTopXPosition + 241, SelectionTopYPosition - 7, _
         15, 15)
    With Select3.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
        .Solid
    End With
    With Select3.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
        
    ' グループ化
    ActiveSheet.Shapes.Range(Array(Select1.Name, _
        Select2.Name, Select3.Name, Select3_1.Name, Select3_2.Name)).Select
    Selection.ShapeRange.Group.Select
End Sub
