Attribute VB_Name = "ReviewDesigin_CursorSetting"
Public Sub シートを左上に設定する()
Attribute シートを左上に設定する.VB_ProcData.VB_Invoke_Func = "U\n14"

Dim sheet As Worksheet

'各シートのシートの右上を活性にする
For Each sheet In ActiveWorkbook.Sheets

    sheet.Activate
    sheet.Cells(1, 1).Activate

Next sheet

'改訂履歴の一番下を選択
Dim i
Dim RevisionName As String
RevisionName = "改訂履歴"

Dim ws As Worksheet
Dim Revisionflag As Boolean
For Each ws In ActiveWorkbook.Sheets
    If ws.Name = RevisionName Then
        Revisionflag = True
    End If
Next ws

If Revisionflag Then
    Do While Worksheets(RevisionName).Cells(7 + i, 4).Value <> ""
        '読み込みセルを定義する
        i = i + 1
    Loop
    Worksheets(RevisionName).Activate
    ActiveWorkbook.Sheets(RevisionName).Cells(7 + i - 1, 4).Activate
End If

MsgBox "カーソル設定が完了しました！"
End Sub
