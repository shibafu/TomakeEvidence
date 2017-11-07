TextProfroof. bas
Attribute VB_Name = "TextProfroof" Public Sub I
Public Sub 処理記述を構成する () Attribute
Attribute 処理記述シートの文章を構成する VB_ProcData.VB_Invoke_Func = "L\n14"

'現在の印刷範囲を設定
	Dim ws As Worksheet Set ws = ActiveSheet Dim ActiveArea As String
	ActiveArea = ws. PageSetup. PrintArea
	ws. Range (ActiveArea). Activate

	'選択範囲のデータを取得
	Dim Rows As Integer
	Rows = Selection. Rows. Count
	Dim Columns As Integer 4 Columns = Selection. Columns. Count
	Dim CellPointer As Range
	'発見したセルを全精査し、処理をかける
	For i = 1 To Rows
		For j = 1 To Columns
			Set CellPointer = ActiveSheet.Cells(i, j)
			Dim Result As Integer
			Dim GrammerCheck As Boolean GrammerCheck = False
			Result = ErrorIDAndCodeChange (CellPointer)
			Result = KutenCummaChange (CellPointer)
			Result = IndentImproove (CellPointer)
			Result = TableWriteFormat (CellPointer)
			Result = GettingPoint Format (CellPointer)
			Result = EditTableBeanIsCheck (CellPointer)

			If Result = 1 Then
				MsgBox "▲【記述エラー！】取得項目が集合体になっていません！▲"
			GoTo GrammerCheckIsFailure End If

			
			If Result = 2 Then
				MsgBox "▲【記述エラー！】集合体から入力されていないテーブル編集があります！▲"
				GoTo GrammerCheckIsFailure
				End If 
		Next j
	Next i

	MsgBox "文章校正のレビューが終わりました！｡:.ﾟヽ(´∀`｡)ﾉﾟ.:｡ ゜"
	
GrammerCheckIsFailure:

End Sub
'ェラーIDをエラーコードに置換する 

Function ErrorIDAndCodeChange (Input Cell As Range)

	If InputCell. Value Like "*" & "エラーID" & "*" Then
		Dim RawText As String RawText = InputCell. Value
		Dim ResultText As String ResultText = Replace (RawText, "エラーID", "エラーコード").
		InputCell. Value = ResultText End If
End Function
	
'日本語句点をカンマに変換する
 Function KutenCummaChange (InputCell As Range)

 If InputCell. Value Like "*" & "、" & "*" Then
	Dim RawText As String RawText = InputCell. Value
	Dim ResultText As String ResultText = Replace (RawText, "、","，")
	InputCell. Value = ResultText End If

 End Function

'インデントを修正するをカンマに変換する
Function IndentImproove(InputCell As Range)

		Dim RawText As String
		Dim Row As long
'最初のインデントを18列目に
If InputCell.Value LIKE "[０-９]." & "*" or InputCell.Value LIKE "[０-９]．" & "*" Then 
	
	RawText = InputCel1.Value
	Row = InputCel1.Row

	InputCell. ClearContents
	ActiveSheet.Cells (Row, 18).Value = RawText
	Exit Function
End If

'次のインデントを19列目に
If InputCell.Value LIKE "（[０-９]）." & "*" or InputCell.Value LIKE "（[０-９]）" & "*" Then 
	
	RawText = InputCel1.Value
	Row = InputCel1.Row

	InputCell. ClearContents
	ActiveSheet.Cells (Row, 19).Value = RawText
	Exit Function
End If

'次のインデントを21列目に
If InputCell.Value LIKE "[ａ-ｚ]." & "*" or InputCell.Value LIKE "[ａ-ｚ]．" & "*" Then 
	
	RawText = InputCel1.Value
	Row = InputCel1.Row

	InputCell. ClearContents
	ActiveSheet.Cells (Row, 21).Value = RawText
	Exit Function
End If

'次のインデントを23列目に
If InputCell.Value LIKE "（[ａ-ｚ]）." & "*" or InputCell.Value LIKE "（[ａ-ｚ]）．" & "*" Then 
	
	RawText = InputCel1.Value
	Row = InputCel1.Row

	InputCell. ClearContents
	ActiveSheet.Cells (Row, 23).Value = RawText
	Exit Function
End If

'次のインデントを24列目に
If InputCell.Value LIKE "[ⅰ-ⅸ]." & "*" or InputCell.Value LIKE "[ⅰ-ⅸ]．" & "*" Then 
	
	RawText = InputCel1.Value
	Row = InputCel1.Row

	InputCell. ClearContents
	ActiveSheet.Cells (Row, 24).Value = RawText
	Exit Function
End If

'次のインデントを25列目に

If InputCell.Value LIKE "（[ⅰ-ⅸ]）." & "*" or InputCell.Value LIKE "（[ⅰ-ⅸ]）．" & "*" Then 
	
	RawText = InputCel1.Value
	Row = InputCel1.Row

	InputCell. ClearContents
	ActiveSheet.Cells (Row, 25).Value = RawText
	Exit Function
End If
	End Function
	
	
'セル内の文字列を、文字配列に変換:
 Function StrSplit (RawText As String) As String()
'分割して文字配列に分ける
		Dim Buffer() As String ReDim Buffer (Len (RawText) - 1)
		For i = 0 To (Len(RawText) - 1)
			Buffer(i) = Mid (RawText, i + 1, 1)
		Next i
		StrSplit = Buffer
End Function
'セル内の文字列を、文字配列に変換
Function StrCombine (RawText() As String) As String
	'結合して一つの文字に
	Dim Result As String Result = Join (RawText, "")
	StrCombine = Result
End Function
'対象テーブル名からかぎDかっこを外す
 Function TableWriteFormat (InputCell As Range) As Integer
'対象テーブル名という列があるなら起動
If InputCel1.Value Like "対象テーブル名" Then
	'ターゲットセルの21セル横を編集
	Dim TargetCell As Range
	Set TargetCel1 = ActiveSheet.Cells(InputCel1.Row, InputCel1.Column + 20)
	'鍵かっこ削除
	TargetCel1. Value = Replace(TargetCel1.Value, "「", "")
	TargetCell.Value = Replace (TargetCel1.Value, "」", "")
	
	TableWriteFormat = 1
Else
	TableWriteFormat = 0
End If
 End Function

'取得項目が集合体名になっているかチェック
Function GettingPointFormat (InputCell As Range) As Integer
'取得項目名という文字列をチェック
If InputCell.Value Like "取得項目" Then
	'ターゲットセルの21セル横を編集
	Dim TargetCell As Range Set TargetCel1 = ActiveSheet. Cells(InputCel1.Row, InputCel1.Column + 20)
	'鍵かっこ削除 Dim BeanWriteIs As Boolean
If Not TargetCel1.Value Like "※*"Then
	BeanWritels = Not TargetCel1.Value Like "<*>” 
End If
 If BeanWriteIs Then
 'エラーコード "1"
 GettingPointFormat = 1
Else
 GettingPointFormat = 0
End If
End Function

'テーブル編集が集合体からされるかチェック
Function EditTableBeanIsCheck (InputCell As Range) As Integer

Dim CreOrRef As Range 
'上のセルが登録と書いているか取得とか言いているかチェック
If InputCell.Row <> 1 Then
	Set CreOrRef = ActiveSheet.Cells(Input Cell. Row - 1, InputCell. Column)
Else
	Set CreOrRef = ActiveSheet.Cells(Input Cell. Row, Input Cell. Column)
End If
' SQLIDという文字をチェック

If Input Cell.Value Like "SQLID” And (CreOrRef. Value Like "*登録*" Or CreOrRef. Value Like "*取得*") Then
'ターゲットセルの21セル横を編集

'") Then
	Dim TargetCell As Range
	Set TargetCell = ActiveSheet.Cells (Input Cell. Row, 1)

	'鍵かっこ削除
	Dim BeanWriteIs As Boolean
		BeanWriteIs = Not TargetCell. Value Like " <*>"
		
	If BeanWriteIs Then 
		EditTableBeanIsCheck = 2
	End If 
	Else
EditTableBeanIsCheck = 0
 End If
 End Function
