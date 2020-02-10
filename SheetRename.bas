Attribute VB_Name = "SheetRename"
Public Sub シート名連番振り分け()
    Dim wb As Workbook
    Set wb = ActiveWorkbook

    '改名作業
    Dim RenameResult As Boolean
    RenameResult = SheetRenameProc(wb)

    MsgBox "シート名変換作業が終わりました｡:.ﾟヽ(´∀`｡)ﾉﾟ.:｡ ゜"

End Sub

'シート改名関数
Function SheetRenameProc(wbObject As Workbook) As Boolean

    Dim SheetCounter As Integer
    SheetCounter = 1
       'シート改名処理開始
   For Each sheet In wbObject.Sheets
           '正規表現を宣言する
        Dim RegExp As VBScript_RegExp_55.RegExp
        Set RegExp = New VBScript_RegExp_55.RegExp
        With RegExp
            .Pattern = "[0-9]+\.(.+)"  '既にインデックスが付けられているかチェック
            .Global = True
        End With

        '正規表現実行
        Dim Results As MatchCollection
        Set Results = RegExp.Execute(sheet.Name)
    
        'ヒットした場合、データを生成する
        If Results.Count > 0 Then

            '結果をセルに代入
            Dim SheetNameWithourIndex As String
            SheetNameWithourIndex = Results.Item(0).SubMatches.Item(0)

            '現在のシート名の頭に数字を付ける
            sheet.Name = CStr(SheetCounter) + "." + SheetNameWithourIndex
        Else

            '現在のシート名の頭に数字を付ける
            sheet.Name = CStr(SheetCounter) + "." + sheet.Name
        End If
       
       SheetCounter = SheetCounter + 1

   Next sheet
End Function

