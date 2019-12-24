Attribute VB_Name = "ColumnTextOutput"
'ワークシートの項目列をテキストに自動出力する'

Public Sub DataOutputText()

'選択中のワークシートを選択する。
Dim ws As Worksheet
Set ws = ActiveSheet

'フォルダダイアログを開き、フォルダパスを取得する
Dim Path As String
Path = get_folder()

'ファイル作成オブジェクトを作成する
Dim ObjFso As Object
Set ObjFso = CreateObject("Scripting.FileSystemObject")

'書き込むファイルを作成する
Dim OutputFile As String

ObjFso.CreateTextFile (Path & "\" & ws.Name & ".txt")
Output = Path & "\" & ws.Name & ".txt"

Open OutputFile For Output As #1

'シート名を書き出し
    Print #1, ws.Name
    Print #1, '\r\n'

'6番目の列をテキストに書き出し
Dim i As Long
i = 1

Do While ws.Cells(i + 10, 6).Value <> ""
    Print #1 ws.Cells(i + 10, 6).Value
    i = i + 1
Loop

'テキストを閉じる
Close #1
'完了メッセージを表示
MsgBox "エンティティの一覧出力が終わりました"

End Sub

Function get_folder() As String
    Dim i As Integer
    Dim Path As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "出力フォルダーを選択"
    If Show = -1 Then
        '選択したフォルダパスを文字列に挿入する
        Path = .SelectedItem(1)
    Else
        Exit Function
    End If
    
    End With
    
    get_folder = Path

End Function

