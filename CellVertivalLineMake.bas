Attribute VB_Name = "CellVertivalLineMake"
Public Sub cü‚ğÁ‚µ‚Ä‰¡ü‚ğ•`‚­()
Attribute cü‚ğÁ‚µ‚Ä‰¡ü‚ğ•`‚­.VB_ProcData.VB_Invoke_Func = "q\n14"

    '‘I‘ğ”ÍˆÍ‚Ìü‚ğÁ‚·
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    
    '‰¡ü‚ğ•`‚­
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With




End Sub



