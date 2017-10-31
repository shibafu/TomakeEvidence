Attribute VB_Name = "Module2"
Sub 三枠作成()
Attribute 三枠作成.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' 枠横幅訂正 Macro
'
' Keyboard Shortcut: Ctrl+Shift+R
'
'現在セルを取得
      
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 56, 2010, 500).Select

'枠線を点線に設定
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .DashStyle = msoLineSysDash
    End With


    '最背面に移動
    Selection.ShapeRange.ZOrder msoSendToBack
    Selection.ShapeRange.Fill.Visible = msoFalse
    '位置修正 現在セルに
    Selection.Top = ActiveCell.Top
    Selection.Left = 56
    
End Sub

Sub 四枠作成()
Attribute 四枠作成.VB_ProcData.VB_Invoke_Func = "E\n14"
'
' 枠横幅訂正 Macro
'
' Keyboard Shortcut: Ctrl+Shift+R
'
'現在セルを取得
      
'   通常のエビデンスに使うサイズ
'    ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 35, 1350, 950).Select

'   DB画面のエビデンスに使うサイズ
'    ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 35, 1350, 950).Select

  ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 35, 1900, 1120).Select
    
'枠線を点線に設定
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .DashStyle = msoLineSysDash
    End With

    '最背面に移動
    Selection.ShapeRange.ZOrder msoSendToBack
    Selection.ShapeRange.Fill.Visible = msoFalse
    '位置修正 現在セルに
    Selection.Top = ActiveCell.Top
    Selection.Left = 30
    
End Sub
