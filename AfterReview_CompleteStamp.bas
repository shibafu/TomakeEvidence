Attribute VB_Name = "AfterReview_CompleteStamp"
Public Sub 済_今日_の文字を加える()
Attribute 済_今日_の文字を加える.VB_ProcData.VB_Invoke_Func = "C\n14"

    '正規表現を使ってクラス名を探知
    Dim RegOb As Object
    Set RegOb = CreateObject("VBScript.RegExp")
    '正規表現を使って探知する
    With RegOb
        .Pattern = "(.*)【済_([0-9\/])+】"
        .Global = True
    End With
    
    '正規表現検索を実行
    Dim Matches
    Set Matches = RegOb.Execute(ActiveCell.Value)

    '既に済マークがある場合
    If Matches.Count > 0 Then
        '済マークを消して、値を入れなおす
        ActiveCell.Value = Matches.Item(0).SubMatches(0)
    Else
        '現在の値の後ろに済マークを入れる
        ActiveCell.Value = ActiveCell.Value & "【済_" & Date & "】"
    End If

End Sub


