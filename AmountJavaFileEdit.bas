Dim BufferAll() As String

Public Sub 複数のJavaファイルを正規表現で編集する()

    '出力したテキストファイルを複数選択
    Inputs = get_files()
    
    'ファイル作成オブジェクトを作成する
    Dim ObjFso As Object
    Set ObjFso = CreateObject("Scripting.FileSystemObject")

    '読み込みファイルリストを開くループ
    
    For Each Input_data In Inputs
    If Input_data <> "" Then
    
        '読み込みファイルを開く
        Open Input_data For Input As #1
    
        '書き込みファイルを開く
        ObjFso.CreateTextFile (Input_data & ".new")
        OutputFile = Input_date & ".new"

        '編集完了時に拡張子をJavaに戻すコード
        'ObjFso.CreateTextFile (Mid(Input_data, 1, Len(Input_data) - 4))
        'OutputFile = Mid(Input_data, 1, Len(Input_data) - 4)

         Open OutputFile For Output As #2
         
         '読み込みテキスト取得ループ
         '行順にセルを読み込む
         
         '読み込みループ
         Dim i As Integer
         Do Until EOF(1)
            ReadJavaFile (i)
            i = i + 1
         Loop
         
         i = 0
         
         '書き込みループ
         Dim j As Integer
         For Each Buf In BufferAl
          a = EditJavaFile(j, CStr(Buf))
          j = j + 1
         Next Buf
         
         j = 0
         
         'ファイルを閉じる
         Close #1
         Close #2
    End If
    Next Input_data
         MsgBox "複数のJavaファイルの編集が終わりました"
         
End Sub

'ファイル選択ウィンドウを開く
Private Function get_files()

    Dim i As Integer
    Dim InputTextFiles(200)
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Javaファイルを選択してください(複数選択可)"
        .AllowMultiSelect = True
        
        If .Show = -1 Then
            '複数選択したファイルに配列を挿入する
            For i = 1 To .SelectedItems.Count
                InputTextFiles(i) = .SelectedItems(i)
            Next i
        Else
            Exit Function
        End If
        
        get_files = InputTextFiles

End Function

'読み込んでバッファーに保存
Private Function ReadJavaFile(i As Integer)

    '保存バッファーを再定義
    '内部の変数をそのままにiの数だけ配列を定義しなおす
    ReDim Preserve BufferAll(i)
    '一行ごとに読み込み
    Line Input #1, Buffer
    BufferAll(i) = CStr(Buffer)

End Function

'JavaFileを編集する
Private Function EditJavaFile(j As Integer, Buffer As String)

'下記関数で読み込んだテキストを編集　使う時だけ
'Buffer = HeaderEdit(Buffer)
'Buffer = ClassJavaDocComment(Buffer)

    'Openしたファイルに読み込んだバッファーを編集して書き込み
    Print #2, Buffer
    
End Function
End Function
'パッケージの上にヘッダーを付けて返却

Private Function HeaderEdit(RawText As String)

    Dim StrAppend(4) As String
    StrAppend(0) = "/**" & vbCrLf
    StrAppend(1) = " *" & vbCrLf
    StrAppend(2) = " * コメントです" & vbCrLf
    StrAppend(3) = " *XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & vbCrLf
    StrAppend(4) = " */" & vbCrLf

    'LIKEを使ったあいまい検索。VBAはLIKEだけでも「正規表現」に近い検索が出来る
    'パッケージの単語が見つかったらその行を編集
    If RawText Like "package *" Then
        RawText = StrAppend(0) & vcrlf & StrAppend(1) & vcrlf & StrAppend(2) & vcrlf & StrAppend(3) & vcrlf & _
        StrAppend(4) & vcrlf & RawText
    End If
    
    HeaderEdit = RawText
    
End Function
'クラスの上にJavaDocコメントを付けて返却

Private Function HeaderEdit(RawText As String) As String


Dim StrAppend(4) As String
StrAppend(0) = "/**" & vbCrLf
StrAppend(1) = " *"
StrAppend(2) = " * @author"
StrAppend(3) = " *" & vbCrLf
StrAppend(4) = " */" & vbCrLf


    '正規表現を使ってクラス名を探知
    Dim IndentFirstPattern As Object
    Set IndentFirstPattern = New VBScriptRegExp_55.RegExp
    '正規表現を使って一番初めのクラス名を
    '「TestCase<4桁の数字>」というパターンをを見つけたら探知する
    With IndentFirstPattern
        .Pattern = " (TestCase[0-9] {4})"
        .Global = True
    End With
    
        '正規表現検索を実行
        Dim FirstMatches As VBScript_RegExp_55.MatchCollection
        Set FirstMatches = IndentFirstPattern.Execute(RawText)
        
    'パターンが見つかったら処理を行う
    If FirstMatches.Count > 0 Then
        '見つかった文字列を入れてコメントを作る
        StrAppend(1) = StrAppend(1) & "テストクラス　" & FirstMatches.Item(0) & vbCrLf
        'クラスコメントを編集
        RawText = StrAppend(0) & StrAppend(1) & StrAppend(2) & StrAppend(3) & StrAppend(4) & RawText
    End If

'コメントを返却
ClassJavaDocComment = RawText

End Function

