Attribute VB_Name = "CharStyleConvert"
Public Sub �J���}��؂�()
Attribute �J���}��؂�.VB_ProcData.VB_Invoke_Func = "K\n14"
'
' �J���}��؂� Macro
'�@������Z���ɑ΂��Ă̂ݎg�p
'
' Keyboard Shortcut: Ctrl+Shift+K

    '������Z���ɕύX
    Selection.NumberFormatLocal = "@"
    '���݂̑I���Z��
    Dim rg As Range
    Set rg = Selection
    '������镶����
    Dim Result As String
    '
    Result = Format(rg.Value, "#,#")

    rg.Value = Result


End Sub
