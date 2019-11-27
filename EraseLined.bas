Attribute VB_Name = "EraseLined"
Public Sub ‘ÅÁ‚µü‚ğˆø‚­()
Attribute ‘ÅÁ‚µü‚ğˆø‚­.VB_ProcData.VB_Invoke_Func = "E\n14"
'
' EraseLined Macro
' ‘ÅÁ‚µü‚ğˆø‚­
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

Public Sub ÔF‚É‚·‚é()
Attribute ÔF‚É‚·‚é.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' EraseLined Macro
' ÔF‚É‚·‚é
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
