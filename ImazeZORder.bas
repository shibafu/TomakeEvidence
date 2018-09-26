Attribute VB_Name = "ImazeZORder"
Private Sub All_Rectangles_Front()
Attribute All_Rectangles_Front.VB_ProcData.VB_Invoke_Func = "w\n14"

        For Each myShape In ActiveSheet.Rectangles
        
            '�F�ȊO�̎l�p�`���őO�ʂ�
            If myShape.ShapeRange.Fill.ForeColor.SchemeColor <> 49 Then
                myShape.ShapeRange.ZOrder msoBringToFront
            End If
            
        Next myShape
 
 End Sub

Public Sub AddRedRectangle()
Attribute AddRedRectangle.VB_ProcData.VB_Invoke_Func = "m\n14"
'�l�p�`��ǉ�
      '�t�H�[���A�C�e���p�T�C�Y
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 70, 15).Select
    
      '�e�[�u���p�T�C�Y
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 70, 13).Select
    
       'DB�p�T�C�Y
    'ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 70, 15).Select

'�g����Ԙg�ɐݒ�
    With Selection.ShapeRange
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(255, 0, 0)
        .Fill.Visible = False
    End With

    '�ʒu�C�� ���݃Z����
    Selection.Top = ActiveCell.Top
    Selection.Left = ActiveCell.Left
    
    
End Sub

Sub AddArrow()
Attribute AddArrow.VB_ProcData.VB_Invoke_Func = "l\n14"
'�����쐬

'���݃Z�����擾
      
    ActiveSheet.Shapes.AddShape(msoShapeDownArrow, 0, 0, 50, 55).Select


    '�ʒu�C�� ���݃Z����
    Selection.Top = ActiveCell.Top
    Selection.Left = ActiveCell.Left + 30
    
    
End Sub

'�ꏊ�C��
Public Sub RectanglePosition()
Attribute RectanglePosition.VB_ProcData.VB_Invoke_Func = "b\n14"

    Selection.ShapeRange.Left = Selection.ShapeRange.Left + 1
    '��ʒu�ύX
    'Selection.ShapeRange.Top = Selection.ShapeRange.Top + 1
    
    
End Sub

'���C��
Public Sub RectangleSize()
Attribute RectangleSize.VB_ProcData.VB_Invoke_Func = "n\n14"
    Selection.Width = Selection.Width + 2
    
End Sub

'�S�摜�I��
Private Sub AllShapesSelected()

  Dim shp As Shape

  For Each shp In ActiveSheet.Shapes
    shp.Select Replace:=False
  Next shp

End Sub

'�S�摜�ɐV�[�g��~��
Private Sub AllShapesBuleBased()

  '�����_�ŃV�[�g�ɑ��݂���摜�S��
  Set ws = ActiveSheet
  

  For Each pct In ws.Pictures
        '�摜�̂ݑI��
          If TypeName(pct) <> "OLEObject" Then
          
          ActiveSheet.Shapes.AddShape(msoShapeRectangle, pct.Left - 10, pct.Top - 10, pct.Width + 20, pct.Height + 20).Select
         '�w�ʂɈړ�
          Selection.ShapeRange.ZOrder msoSendToBack
          
          End If
  
     Next pct

End Sub


