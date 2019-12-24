Attribute VB_Name = "AfterReview_CheckCellSearch"
'レビューでチェックされたセルを検索する
'@Author 野澤
'
Public Sub レビューでチェックされたセルをサーチする()
'Workbook(ActiveWorkbook.Name).Active

Dim TargetWorkSheet As Worksheet
' 全シートに対して行う
For Each TargetWorkSheet In ActiveWorkbook.Sheets

    'チェックされたセルを選択し､メッセージウィンドウに表示
    Call チェックされたセルを選択し､メッセージウィンドウに表示(TargetWorkSheet)

Next TargetWorkSheet

MsgBox "チェックが終了しました！"
End Sub

Private Sub チェックされたセルを選択し､メッセージウィンドウに表示(ws As Worksheet)

Dim PrArea As String '印刷範囲
Dim AreaColumn As Integer '印刷範囲
Dim AreaRow As Integer '印刷範囲

ws.Activate
'印刷範囲を取得する
PrArea = ws.PageSetup.PrintArea
If PrArea = "" Then
    Exit Sub
End If
'印刷範囲を選択
ws.Range(Replace(PrArea, "$", "")).Select

'検索エリア列数を取得する
AreaColumn = Selection.Columns.Count
'検索エリア行数を取得する
AreaRow = Selection.Rows.Count

'列と行数分ループ
For RowPointNm = 1 To AreaRow
    For ColumnPointNm = 1 To AreaColumn
    
    'セルがピンクの時、セルを取得
    If ws.Cells(RowPointNm, ColumnPointNm).Interior.Color = ColorConstants.vbMagenta Or _
    ws.Cells(RowPointNm, ColumnPointNm).Interior.Color = 16751103 Then
        'メッセージボックスにメッセージを表示する
        MsgBox "修正セルが見つかりました！：" & ws.Name & "アドレス：" & ws.Cells(RowPointNm, ColumnPointNm).Address & vbCrLf _
        & "内容：" & ws.Cells(RowPointNm, ColumnPointNm).Value
    End If
    Next ColumnPointNm
Next RowPointNm


End Sub
