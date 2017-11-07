Attribute VB_Name = "ImazeZORder"
Private Sub All_Rectangles_Front()
Attribute All_Rectangles_Front.VB_ProcData.VB_Invoke_Func = "w\n14"

        For Each myShape In ActiveSheet.Rectangles
        
            '青色以外の四角形を最前面に
            If myShape.ShapeRange.Fill.ForeColor.SchemeColor <> 49 Then
                myShape.ShapeRange.ZOrder msoBringToFront
            End If
            
        Next myShape
 
 End Sub

Public Sub AddRedRectangle()
Attribute AddRedRectangle.VB_ProcData.VB_Invoke_Func = "m\n14"
'四角形を追加
      'フォームアイテム用サイズ
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 70, 15).Select
    
      'テーブル用サイズ
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 70, 13).Select
    
       'DB用サイズ
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 70, 15).Select

'枠線を赤枠に設定
    With Selection.ShapeRange
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(255, 0, 0)
        .Fill.Visible = False
    End With

    '位置修正 現在セルに
    Selection.Top = ActiveCell.Top
    Selection.Left = ActiveCell.Left
    
    
End Sub

Sub AddArrow()
Attribute AddArrow.VB_ProcData.VB_Invoke_Func = "l\n14"
'矢印を作成

'現在セルを取得
      
    ActiveSheet.Shapes.AddShape(msoShapeDownArrow, 0, 0, 50, 55).Select


    '位置修正 現在セルに
    Selection.Top = ActiveCell.Top
    Selection.Left = ActiveCell.Left + 30
    
    
End Sub

'場所修正
Public Sub RectanglePosition()
Attribute RectanglePosition.VB_ProcData.VB_Invoke_Func = "b\n14"

    Selection.ShapeRange.Left = Selection.ShapeRange.Left + 1
    '上位置変更
    'Selection.ShapeRange.Top = Selection.ShapeRange.Top + 1
    
    
End Sub

'幅修正
Public Sub RectangleSize()
Attribute RectangleSize.VB_ProcData.VB_Invoke_Func = "n\n14"
    Selection.Width = Selection.Width + 2
    
End Sub

'全画像選択
Private Sub AllShapesSelected()

  Dim shp As Shape

  For Each shp In ActiveSheet.Shapes
    shp.Select Replace:=False
  Next shp

End Sub

'全画像に青シートを敷く
Private Sub AllShapesBuleBased()

  '現時点でシートに存在する画像全て
  Set ws = ActiveSheet
  

  For Each pct In ws.Pictures
        '画像のみ選択
          If TypeName(pct) <> "OLEObject" Then
          
          ActiveSheet.Shapes.AddShape(msoShapeRectangle, pct.Left - 10, pct.Top - 10, pct.Width + 20, pct.Height + 20).Select
         '背面に移動
          Selection.ShapeRange.ZOrder msoSendToBack
          
          End If
  
     Next pct

End Sub


