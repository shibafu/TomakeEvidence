Attribute VB_Name = "DummyData"
Public Sub DummyDataGenerate()
Attribute DummyDataGenerate.VB_ProcData.VB_Invoke_Func = "O\n14"

'セルの最終行を取得
    Dim n As Integer
    n = Cells(Rows.Count, "A").End(xlUp).Row
     
     
  Dim DummyDataNum As String
  Dim DummyDataChr As String
    '↓ここからループのスタート
    For i = 1 To n
    
    

    'セルの値A列からを取得
    Dim Ketasu As Integer
    Ketasu = Cells(i, 1).Value
    
        '数字のダミーデータ生成
       'ダミーデータを生成
        For j = 1 To Ketasu
         Dim NextNum As String
         NextNum = j Mod 10
         DummyDataNum = DummyDataNum + NextNum
        Next j
        
    '値をB列に代入
    Cells(i, 2).Value = DummyDataNum
    
    '全角データを代入
    Dim ZenkakuResult As String
    ZenkakuResult = LeftB(StrConv(DummyDataNum, vbWide), Len(DummyDataNum))
    Cells(i, 3).Value = ZenkakuResult
        
        '桁数数字データを代入
        For j = 1 To Ketasu
         Dim DigiNum As String
         If j = 1 Then
            DigiNum = "1"
         Else
            DigiNum = DigiNum + "0"
         End If
        Next j
    Cells(i, 4).Value = DigiNum
        
    '初期化
    DummyDataNum = ""
    ZenkakuResult = ""
    DigiNum = ""
    
        '文字のダミーデータ生成
       'ダミーデータを生成
        For j = 1 To Ketasu
         Dim NextChr As String
         NextChr = "N"
         DummyDataChr = DummyDataChr + NextChr
        Next j
        
    '値をJ列に代入
    Cells(i, 5).Value = DummyDataChr
    DummyDataChr = ""
    Next i
    '↑ここまでがループ

MsgBox "ダミーデータの生成が終わりました"

        Application.OnKey "^O", "DummyDataGenerate"

End Sub



