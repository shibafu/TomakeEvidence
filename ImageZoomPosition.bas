Attribute VB_Name = "ImageZoomPosition"
Public Sub �摜�g�k_��()
Attribute �摜�g�k_��.VB_ProcData.VB_Invoke_Func = "M\n14"
'
' �摜�g�k Macro
'
' Keyboard Shortcut: Ctrl+Shift+M
'
    Selection.ShapeRange.ScaleWidth 0.835, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 0.835, msoFalse, msoScaleFromTopLeft

'   ���ɖ߂��{�� 1/ 0.835
'    Selection.ShapeRange.ScaleWidth 1.19760479042, msoFalse, msoScaleFromTopLeft
'    Selection.ShapeRange.ScaleHeight 1.19760479042, msoFalse, msoScaleFromTopLeft
    
    Selection.Top = Selection.Top + 5
    Selection.Left = Selection.Left + 15
    
    Selection.ShapeRange.ZOrder msoSendToBack
End Sub

Public Sub �������C��()
Attribute �������C��.VB_ProcData.VB_Invoke_Func = "r\n14"
'
' Macro8 Macro
'
' Keyboard Shortcut: Ctrl+r
'
'
        Selection.Top = Selection.Top + 5
        Selection.Left = Selection.Left + 15
        
End Sub
Public Sub �摜�g�k_��()
Attribute �摜�g�k_��.VB_ProcData.VB_Invoke_Func = "H\n14"
'
' Macro9 Macro
'
' Keyboard Shortcut: Ctrl+Shift+H
' �悭�g���T�C�Y0.673�@0.635
' ����̃T�C�Y 0.625
'�������T�C�Y0.66
'�������T�C�Y 0.69

    Selection.ShapeRange.ScaleWidth 0.94, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 0.94, msoFalse, msoScaleFromTopLeft
    
    Selection.Top = Selection.Top + 5
    Selection.Left = Selection.Left + 15
    
End Sub
