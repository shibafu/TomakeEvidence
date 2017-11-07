ColumnTextInput.bas

Attribute VB_name = "ColumnTextInput"
Public Sub DataInput()
	'出力したテキストファイルを複数選択
	Inputs = get_files()
	
	Dim ColumnTextNum As Integer
	Dim Buffer As String
	
	'開始列を指定,デフォルトはB列から
	ColumnNum = 2
	
	'読み込みファイルリストを開くループ
	For Each Input_data In Inputs
		If Input_data <> "" Then
			Open Input_data For Input As #1
			
			Dim RawNum As Integer
			RawNum = 1
			
			'読み込みテキスト取得ループ
			'行頭にセルを打ち込む
			
			Do Until EOF(1)
				Line Input #1 , Buffer
				ActiveSheet.Cells(RawNum, ColumnNum).Vaue = Buffer
				RawNum = RawNum + 1
			Loop
			
			Close #1
			'列を一つずらす
		ColumnNum = ColumnNum + 1
		End If			
	Next Input_data
End Sub

Function get_files()
	Dim i As Integer
	Dim InputTextFiles(99)
	
	With Application.FileDialog(msoFileDialogFilePicker)
		.Title = "テキストファイルを選択"
		.AllowMultiSelect = True
		If .Show -1 Then
			'複数選択をしたファイルに配列を挿入する
			For i = 1 To .SelectItems.Count
				InputTextFiles(i) = SelectItems(i)
			Next i
		Else
			Exit Function
		End If
	End With
	
	get_files = InputTextFiles

End Function
