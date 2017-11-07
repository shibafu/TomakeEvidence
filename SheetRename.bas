SheetRename.bas

Public Sub SheetRename() 
    Dim wb As Workbook
	Set wb = ActviveWorkbook

    Dim ActiveSheetStart As Integer 'アクティブなシートが何枚目か
	ActiveSheetStart = wb.ActiveSheet.Index
    Dim SerialSheetStart As Integer '処理を開始するシートが何枚目か
    Dim SerialSheetMount As Integer  '処理するシート枚数

    '開始枚数と処理枚数を計算
    Dim CountResult() As Integer
    CountResult=StartSheetCount (wb, ActiveSheetStart)

    SerialSheetStart=CountResult (1) 
    SerialSheetMount=CountResult (2)


    '改名作業
    Dim RenameResult As Boolean
    RenameResult=SheetRenameProc (wb, SerialSheetStart, SerialSheetAmount)

    '結果判定
    If RenameResult=True Then
        MsgBox "シート名変換作業が終わりました｡:.ﾟヽ(´∀`｡)ﾉﾟ.:｡ ゜"
    Else
        MsgBox "対応形式のシートが見つかりませんでした(´・ω・｀)"
    Esd If

End Sub
'シート改名開始計測関数
Function StartSheetCount(ActiveworkbookObject As Workbook, ActiveSheetStart As Integer) As Integer()
    '現在のワークブックのシートを精査

    Dim SerialSheetStart As Integer '処理を開始するシートが何枚目か
    SerialSheetStart=1
    Dim SerialSheetMount As Integer  '処理するシート枚数
    SerialSheetCount=0
    Dim SheetThroughedFlag As Boolean 
    SheetThroughedFlag=False

    For Each pointerSheet In ActiveworkbookObject.Sheets

        '／／アクティブなシートからルーフを開始するために'最初はループを空回しにする。
        If ActiveSheetStart > 1 Then
           ActiveSheetStart = ActiveSheetStart - 1 
           GoTo CONTINUE
        End If


    'シート名がく文字列）+'('＋数字＋つ'のパターン時にシート処理枚数を加算
        If pointerSheet. Name Like "*" & "(" &"[0-9]*" &")" Then 
           SerialSheetMount = SerialSheetMount + 1
           '既定の処理シートが見つかったので'処理開始枚数は加算しない 
           SheetThroughedFlag = True
        End If

        'パターンの一連の読み込みが終了したらループを脱出
        If SerialSheetCount = 0 And SheetThroughedFlag = False Then 
              SerialSheetStart = SerialSheetStart + 1
        ElseIf SerialSheetMount = 0 Then
           Exit For
        End If
        'ループの終端。最初は空回しする 
CONTINUE:
    Next pointerSheet


    '開始位置を最終決定。
    SerialSheetStart = SerialSheetStart + ActiveworkbookObject.ActiveSheet.Index - 1

    '結果を返却する
    Dim ReturnResult(2) As Integer 
    ReturnResult(1) = SerialSheetStart 
    ReturnResult(2) = SerialSheetCount

    Start SheetCount = ReturnResult

End Function
'シート改名関数
Function SheetRenameProc(ActjveworkbookObject As Workbook, SerialSheetStart As Integer, SerialSheetMount As Integer) As Boolean

    '処理枚数が0の時は関数を終了。
	'失敗フラグを返す
   If SerialSheetAmount  0 Then 
       SheetRenameProc = False
       Exit Function 
   End If

       'シート改名処理開始
   For i = 1 To SerialSheetMount

       '現在のシート名を取得
       Dim SheetName As String
       SheetName = ActiveworkbookObject.Sheets(SerialSheetStart + i - 1).Name
       '"(”が見っかったところまでをヘッダーとして保存
	Dim SheetNameHeader As String 
       SheetNameHeader=Mid(SheetName, 1, InStr(SheetName, "("))

       'シート名を改名
       SheetName = SheetNameHeader& i + 1 & ")"

       Activeworkbook. Sheets (SerialSheetStart + i - 1).Name = SheetName
   Next i

   '成功フラグ
	SheetRenameProc = True 
End Function
