Attribute VB_Name = "EraseLined"
Public Sub �ŏ�����������()
Attribute �ŏ�����������.VB_ProcData.VB_Invoke_Func = "E\n14"
'
' EraseLined Macro
' �ŏ�����������
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

Public Sub �ԐF�ɂ���()
Attribute �ԐF�ɂ���.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' EraseLined Macro
' �ԐF�ɂ���
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
