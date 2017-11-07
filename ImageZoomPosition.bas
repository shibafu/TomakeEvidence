Attribute VB_Name = "ImageZoomPosition"
Public Sub ‰æ‘œŠgk_‘å()
Attribute ‰æ‘œŠgk_‘å.VB_ProcData.VB_Invoke_Func = "M\n14"
'
' ‰æ‘œŠgk Macro
'
' Keyboard Shortcut: Ctrl+Shift+M
'
    Selection.ShapeRange.ScaleWidth 0.835, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 0.835, msoFalse, msoScaleFromTopLeft

'   Œ³‚É–ß‚·”{—¦ 1/ 0.835
'    Selection.ShapeRange.ScaleWidth 1.19760479042, msoFalse, msoScaleFromTopLeft
'    Selection.ShapeRange.ScaleHeight 1.19760479042, msoFalse, msoScaleFromTopLeft
    
    Selection.Top = Selection.Top + 5
    Selection.Left = Selection.Left + 15
    
    Selection.ShapeRange.ZOrder msoSendToBack
End Sub

Public Sub ‰¡•¶C³()
Attribute ‰¡•¶C³.VB_ProcData.VB_Invoke_Func = "r\n14"
'
' Macro8 Macro
'
' Keyboard Shortcut: Ctrl+r
'
'
        Selection.Top = Selection.Top + 5
        Selection.Left = Selection.Left + 15
        
End Sub
Public Sub ‰æ‘œŠgk_¬()
Attribute ‰æ‘œŠgk_¬.VB_ProcData.VB_Invoke_Func = "H\n14"
'
' Macro9 Macro
'
' Keyboard Shortcut: Ctrl+Shift+H
' ‚æ‚­g‚¤ƒTƒCƒY0.673@0.635
' ¡‰ñ‚ÌƒTƒCƒY 0.625
'’†•¶šƒTƒCƒY0.66
'¬•¶šƒTƒCƒY 0.69

    Selection.ShapeRange.ScaleWidth 0.94, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 0.94, msoFalse, msoScaleFromTopLeft
    
    Selection.Top = Selection.Top + 5
    Selection.Left = Selection.Left + 15
    
End Sub
