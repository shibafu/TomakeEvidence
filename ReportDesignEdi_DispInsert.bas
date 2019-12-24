Attribute VB_Name = "ReportDesignEdi_DispInsert"
Dim TARGET_CELL_JP_NAME_COLUMN
Dim TARGET_CELL_REMARK_COLUMN
Dim TARGET_CELL_REVISION_COLUMN
'帳票の隣に修正を挿入する
Public Sub DispInsert()

'定数を初期化
TARGET_CELL_JP_NAME_COLUMN = 2
TARGET_CELL_REMARK_COLUMN = 12
TARGET_CELL_REVISION_COLUMN = 13

Dim CellRowIndex As Integer
CellRowIndex = 1 + 6

Do While ActiveSheet.Cells(CellRowIndex, TARGET_CELL_JP_NAME_COLUMN).Value <> ""

    '正規表現を使ってクラス名を探知
    Dim RegOb As Object
    Set RegOb = CreateObject("VBScript.RegExp")
    '正規表現を使って探知する
    With RegOb
        .Pattern = "^[^%]+([０-９0-9]+)"
        .Global = True
    End With
    
        '正規表現検索を実行
        Dim Matches
        Set Matches = RegOb.Execute(ActiveSheet.Cells(CellRowIndex, TARGET_CELL_JP_NAME_COLUMN).Value)
        
    'パターンが見つかったら処理を行う
    If Matches.Count > 0 Then
        '見つかった文字列を入れてコメントを作る
        ActiveSheet.Cells(CellRowIndex, TARGET_CELL_REMARK_COLUMN).Value = Cells(CellRowIndex, TARGET_CELL_REMARK_COLUMN).Value & vbCrLf _
        & "Z,ZZ9"
        'クラスコメントを編集
        ActiveSheet.Cells(CellRowIndex, TARGET_CELL_REVISION_COLUMN).Value = ActiveSheet.Cells(CellRowIndex, TARGET_CELL_REVISION_COLUMN).Value & _
        "2019/12/13　野澤　修正　改訂履歴（No.8）"
        ActiveSheet.Cells(CellRowIndex, TARGET_CELL_REVISION_COLUMN).VerticalAlignment = xlCenter
    End If

    CellRowIndex = CellRowIndex + 1
Loop
End Sub
