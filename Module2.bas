Attribute VB_Name = "Module2"
Sub �O�g�쐬()
Attribute �O�g�쐬.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' �g�������� Macro
'
' Keyboard Shortcut: Ctrl+Shift+R
'
'���݃Z�����擾
      
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 56, 2010, 500).Select

'�g����_���ɐݒ�
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .DashStyle = msoLineSysDash
    End With


    '�Ŕw�ʂɈړ�
    Selection.ShapeRange.ZOrder msoSendToBack
    Selection.ShapeRange.Fill.Visible = msoFalse
    '�ʒu�C�� ���݃Z����
    Selection.Top = ActiveCell.Top
    Selection.Left = 56
    
End Sub

Sub �l�g�쐬()
Attribute �l�g�쐬.VB_ProcData.VB_Invoke_Func = "E\n14"
'
' �g�������� Macro
'
' Keyboard Shortcut: Ctrl+Shift+R
'
'���݃Z�����擾
      
'   �ʏ�̃G�r�f���X�Ɏg���T�C�Y
'    ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 35, 1350, 950).Select

'   DB��ʂ̃G�r�f���X�Ɏg���T�C�Y
'    ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 35, 1350, 950).Select

  ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 35, 1900, 1120).Select
    
'�g����_���ɐݒ�
    With Selection.ShapeRange.Line
        .Visible = msoTrue
        .DashStyle = msoLineSysDash
    End With

    '�Ŕw�ʂɈړ�
    Selection.ShapeRange.ZOrder msoSendToBack
    Selection.ShapeRange.Fill.Visible = msoFalse
    '�ʒu�C�� ���݃Z����
    Selection.Top = ActiveCell.Top
    Selection.Left = 30
    
End Sub
