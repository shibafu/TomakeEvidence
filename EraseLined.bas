Attribute VB_Name = "EraseLined"
Public Sub 打消し線を引く()
Attribute 打消し線を引く.VB_ProcData.VB_Invoke_Func = "E\n14"
'
' EraseLined Macro
' 打消し線を引く
'
' Keyboard Shortcut: Ctrl+Shift+E
'
    If Selection.Font.Strikethrough Then
        With Selection.Font
            .Strikethrough = False
        End With
    Else
        With Selection.Font
            .Strikethrough = True
        End With
    End If
    
End Sub

Public Sub 赤色にする()
Attribute 赤色にする.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' EraseLined Macro
' 赤色にする
'
' Keyboard Shortcut: Ctrl+Shift+R
'
    If Selection.Font.Color <> vbRed Then
        With Selection.Font
           .Color = vbRed
        End With
    Else
        With Selection.Font
            .Color = Black
        End With
    End If
End Sub
