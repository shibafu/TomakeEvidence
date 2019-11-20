Attribute VB_Name = "VertivalHorizonCng"
Public Sub 縦横を変換する()

Dim i As Integer
i = 1


Vertival = Vertival + 1
Horizontal = Horizontal + 1

'一時変数
Dim Temp As String

Do While ActiveSheet.Cells(i, 1).Value <> ""


    Temp = ActiveSheet.Cells(i, 1).Value

    '元セルの値を消す
    ActiveSheet.Cells(i, 1).Value = ""
    '元セルの値をだ遠敷
    ActiveSheet.Cells(1, i).Value = Temp

    i = i + 1
Loop



End Sub
