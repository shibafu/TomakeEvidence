Attribute VB_Name = "EROneToOne"
Public Sub OneToOn()
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
        
    ' グループ化
    ActiveSheet.Shapes.Range(Array(Select1.Name, _
        Select2.Name, Select3.Name)).Select
    Selection.ShapeRange.Group.Select
End Sub


