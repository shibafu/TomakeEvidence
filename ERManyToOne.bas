Attribute VB_Name = "ERManyToOne"
Public Sub ManyToOne()
'
' OneToOn Macro
'
'

    ' 選択せるの座標を記録
    Dim SelectionTopXPosition As Integer
    SelectionTopXPosition = Selection.Top
    Dim SelectionTopYPosition As Integer
    SelectionTopYPosition = Selection.Left
    ' 線を追加
    Dim Select1 As Shape
    Set Select1 = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, _
    SelectionTopXPosition, SelectionTopYPosition, SelectionTopXPosition + 264, SelectionTopYPosition)
    
    Select1.Line.ForeColor.RGB = RGB(0, 0, 0)
    Dim Select2 As Shape
    Set Select2 = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, _
    SelectionTopXPosition + 10, SelectionTopYPosition - 10, SelectionTopXPosition + 10, SelectionTopYPosition + 10)
    Select2.Line.ForeColor.RGB = RGB(0, 0, 0)
        
    Dim Select3 As Shape
    Set Select3 = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, _
    SelectionTopXPosition + 246, SelectionTopYPosition - 10, SelectionTopXPosition + 246, SelectionTopYPosition + 10)
    Select3.Line.ForeColor.RGB = RGB(0, 0, 0)
    
    '足を書く
    Dim Select3_1 As Shape
    Set Select3_1 = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, _
    SelectionTopXPosition + 246, SelectionTopYPosition, SelectionTopXPosition + 261, SelectionTopYPosition + 15)
    Select3_1.Line.ForeColor.RGB = RGB(0, 0, 0)
    
    Dim Select3_2 As Shape
    Set Select3_2 = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, _
    SelectionTopXPosition + 246, SelectionTopYPosition, SelectionTopXPosition + 261, SelectionTopYPosition - 15)
    Select3_2.Line.ForeColor.RGB = RGB(0, 0, 0)

        
    ' グループ化
    ActiveSheet.Shapes.Range(Array(Select1.Name, _
        Select2.Name, Select3.Name, Select3_1.Name, Select3_2.Name _
        )).Select
    Selection.ShapeRange.Group.Select
End Sub






