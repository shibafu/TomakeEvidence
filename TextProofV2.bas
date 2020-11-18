Attribute VB_Name = "TextProof"
Public Sub プロセス詳細を文章を校正する()

    '現在の印刷範囲を設定
    '現在は選択範囲を設定する
    Dim ws As Worksheet
    Set ws = ActiveSheet
    '選択範囲
    Dim ActiveArea As String
    ActiveArea = ws.PageSetup.PrintArea
    'ActiveArea = Selection.Address
    ws.Range(ActiveArea).Activate
    
    '対象エリアをセット
    Dim TargetArea As Range
    'Set TargetArea = ws.Range(ActiveArea)
    Set TargetArea = Selection

    '　処理機能をインスタンス化
    Dim tp_Indent As TextProof_SerialNumberCreate
    Set tp_Indent = New TextProof_SerialNumberCreate

    
    'インデント連番を作成する
    Call tp_Indent.Init
    Call tp_Indent.インデントに応じた連番を作り出力する(TargetArea)
    

    MsgBox "文章校正のレビューが終わりました！｡:.ﾟヽ(´∀`｡)ﾉﾟ.:｡ ゜"

End Sub


