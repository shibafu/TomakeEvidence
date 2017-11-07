Attribute VB_Name = "DataSizeAnalyze"
Public Sub DigiAnalySisExe()
Attribute DigiAnalySisExe.VB_ProcData.VB_Invoke_Func = "I\n14"
'セルの最終行を取得
    Dim n As Integer
    n = Cells(Rows.Count, "A").End(xlUp).Row
     
     
  Dim DigiResult As String
    '↓ここからループのスタート
    For i = 1 To n
    '空白でないときのみ実行

          'フォーマットを文字列に変更
         Cells(i, 1).NumberFormatLocal = "@"
           'セルの値A列から文字を取得
         Dim CharSize As Long
         Dim RawChar As String
         RawChar = Cells(i, 1).Value
    
         'バイト数取得
         CharSize = LenB(StrConv(RawChar, vbFromUnicode))
    
         Cells(i, 2).Value = CharSize
         
         Dim HankakuSize As Long
         Dim ZenkakuSize As Long

         For k = 1 To Len(RawChar)
                '一文字を半角かどうか判定
                 If Mid(RawChar, k, 1) Like "[- ｡-ﾟ 0-9a-zA-Z]" Then
                 HankakuSize = HankakuSize + 1
                 Else
                 'それ以外は全角
                 ZenkakuSize = ZenkakuSize + 1
                 End If
         Next k


        '全角のサイズが左記に出力される
        Cells(i, 3).Value = ZenkakuSize
        Cells(i, 4).Value = HankakuSize
         

    
        '初期化
         RawChar = ""
         CharSize = 0
         HankakuSize = 0
         ZenkakuSize = 0
     

    
    Next i
    '↑ここまでがループ

MsgBox "桁数の分析が終わりました"

        Application.OnKey "^I", "DigiAnalySisExe"

End Sub
