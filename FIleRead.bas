Public Sub ファイルの操作の実行()

  'ファイルを開く
  Dim FileAddress As String
  FileAddress = ファイルダイアログを開く()
  
  '読み込みバッファを用意
  Dim buf As String

  ' データ格納コレクションを用意
  Dim inputData As Collection
  Set inputData = New Collection
  
  'ファイルを開く
  Open FileAddress For Input As #1
  
  'ファイル入力
  Do Until EOF(1)
        '1行ずつバッファに入れる
        Line Input #1, buf
        ' データ格納コレクションに入れる
        inputData.Add (CStr(buf))
  Loop
  'ファイルを閉じる
  Close #1    ''1番のファイルを閉じます

End Sub
    
    
Private Function ファイルダイアログを開く() As String

    Dim result As Variant

    result = Application.GetOpenFilename( _
            Title:="テキストを選択してください", _
            MultiSelect:=True)

ファイルダイアログを開く = result(1)

         
End Function