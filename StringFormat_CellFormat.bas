Attribute VB_Name = "StringFormat_CellFormat"
Public Sub YokoOver()
Attribute YokoOver.VB_Description = "����͂ݏo���܂�"
Attribute YokoOver.VB_ProcData.VB_Invoke_Func = "L\n14"
'
' ���� Macro
' ����͂ݏo���܂�
'
' Keyboard Shortcut: Ctrl+Shift+L
'
    With Selection
        .WrapText = False
        .AddIndent = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Application.OnKey "^L", "YokoOver"
End Sub

Public Sub �t�H���g�T�C�Y�C��()
Attribute �t�H���g�T�C�Y�C��.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' �t�H���g�T�C�Y�C�� Macro
'
' Keyboard Shortcut: Ctrl+Shift+F
'

    With Selection.Font
        .Name = "���C���I"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    With Selection
            .Borders.LineStyle = False
    End With
End Sub

'�R�s�[�����l�𕶎����
Public Sub FormatChangeString()
Attribute FormatChangeString.VB_ProcData.VB_Invoke_Func = "Q\n14"

    With Selection
        .NumberFormatLocal = "@"
        .Borders.LineStyle = False
    End With
    
    
    Application.OnKey "^Q", "FormatChangeString"
End Sub
Public Sub BorderLineOn()
Attribute BorderLineOn.VB_ProcData.VB_Invoke_Func = "W\n14"
'
' ���g�C�� Macro
'
' Keyboard Shortcut: Ctrl+Shift+W

    With Selection
        .Borders.LineStyle = True
    End With
    
        Application.OnKey "^W", "BorderLineOn"
End Sub

