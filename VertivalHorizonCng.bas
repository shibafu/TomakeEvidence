Attribute VB_Name = "VertivalHorizonCng"
Public Sub �c����ϊ�����()

Dim i As Integer
i = 1


Vertival = Vertival + 1
Horizontal = Horizontal + 1

'�ꎞ�ϐ�
Dim Temp As String

Do While ActiveSheet.Cells(i, 1).Value <> ""


    Temp = ActiveSheet.Cells(i, 1).Value

    '���Z���̒l������
    ActiveSheet.Cells(i, 1).Value = ""
    '���Z���̒l�������~
    ActiveSheet.Cells(1, i).Value = Temp

    i = i + 1
Loop



End Sub
