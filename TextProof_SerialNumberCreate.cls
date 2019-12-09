VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextProof_SerialNumberCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'インデントカウンター１
Dim CounterNm1st As Integer
'インデントカウンター２
Dim CounterNm2nd As Integer
'インデントカウンター３
Dim CounterNm3rd As Integer
'インデントカウンター４
Dim CounterNm4th As Integer
'インデントカウンター５
Dim CounterNm5th As Integer
'インデントカウンター６
Dim CounterNm6th As Integer
'インデント列、基準値
Dim FIRST_INDENT_ROW As Integer

'初期化処理
Public Sub Init()

CounterNm1st = 0
CounterNm2nd = 0
CounterNm3rd = 0
CounterNm4th = 0
CounterNm5th = 0
CounterNm6th = 0
FIRST_INDENT_ROW = 3

End Sub


Public Sub インデントに応じた連番を作り出力する(ByRef SelectionArea As Range)

    For i = 0 To SelectionArea.Rows.Count - 1
        For j = 0 To SelectionArea.Columns.Count - 1
            '正規表現を使ってクラス名を探知
            Dim IndentPattern As Object
            Set IndentPattern = CreateObject("VBScript.RegExp")
            'このパターンの正規表現呼び出しもあり
            'Set RE = CreateObject("VBScript.RegExp")
            'X-X-X＜処理記述＞という文を検索する
            With IndentPattern
                .Pattern = "^([1-9\-]+\.)(.*)"
                .Global = True
            End With
    
            '正規表現検索を実行
            Set Matches = IndentPattern.Execute(Cells(SelectionArea.Row + i, SelectionArea.Column + j).Value)
        
            Dim MatchedIndent As String
            MatchedIndent = ""
            Dim OtherSentence As String
            OtherSentence = ""
            'パターンが見つかったら処理を行う
            If Matches.Count > 0 Then
                '見つかった文字列を入れてコメントを作る
                MatchedIndent = Matches.Item(0).SubMatches.Item(0)
                OtherSentence = Matches.Item(0).SubMatches.Item(1)
            End If


            'マッチしたときに実行
            If MatchedIndent <> "" Then
                'マッチしたときセルの中身を空にする
                Cells(SelectionArea.Row + i, SelectionArea.Column + j).Value = ""
               'X.の時、最初の所だけ入力して他カウンターを初期化
                If Len(MatchedIndent) = 2 Then
                  
                    CounterNm1st = CounterNm1st + 1
                    'セルに出力
                    Cells(SelectionArea.Row + i, FIRST_INDENT_ROW).Value = CStr(CounterNm1st) & "." & OtherSentence
            
                    CounterNm2nd = 0
                    CounterNm3rd = 0
                    CounterNm4th = 0
                    CounterNm5th = 0
                    CounterNm6th = 0
                    'X-X.の時、
                ElseIf Len(MatchedIndent) = 4 Then
        
                    CounterNm2nd = CounterNm2nd + 1
                    'セルに出力
                    Cells(SelectionArea.Row + i, FIRST_INDENT_ROW + 1).Value = CStr(CounterNm1st) & "-" & CStr(CounterNm2nd) & "." & OtherSentence
            
                    CounterNm3rd = 0
                    CounterNm4th = 0
                    CounterNm5th = 0
                    CounterNm6th = 0
                    'X-X-X.の時、最初の所だけ入力して他カウンターを初期化
                ElseIf Len(MatchedIndent) = 6 Then
        
                    CounterNm3rd = CounterNm3rd + 1
                    'セルに出力
                    Cells(SelectionArea.Row + i, FIRST_INDENT_ROW + 2).Value = CStr(CounterNm1st) & "-" & CStr(CounterNm2nd) _
                    & "-" & CStr(CounterNm3rd) & "." & OtherSentence
            
                    CounterNm4th = 0
                    CounterNm5th = 0
                    CounterNm6th = 0
                    'X-X-X.の時、最初の所だけ入力して他カウンターを初期化
                ElseIf Len(MatchedIndent) = 8 Then
        
                    CounterNm4th = CounterNm4th + 1
                    'セルに出力
                    Cells(SelectionArea.Row + i, FIRST_INDENT_ROW + 3).Value = CStr(CounterNm1st) & "-" & CStr(CounterNm2nd) _
                    & "-" & CStr(CounterNm3rd) & "-" & CStr(CounterNm4th) & "." & OtherSentence
            
                    CounterNm5th = 0
                    CounterNm6th = 0
                'X-X-X-X.の時、最初の所だけ入力して他カウンターを初期化
                ElseIf Len(MatchedIndent) = 10 Then
        
                    CounterNm5th = CounterNm5th + 1
                    'セルに出力
                    Cells(SelectionArea.Row + i, FIRST_INDENT_ROW + 4).Value = CStr(CounterNm1st) & "-" & CStr(CounterNm2nd) _
                    & "-" & CStr(CounterNm3rd) & "-" & CStr(CounterNm4th) & "-" _
                    & CStr(CounterNm5th) & "." & OtherSentence
            
                    CounterNm6th = 0
                    'X-X-X-X.の時、最初の所だけ入力して他カウンターを初期化
                ElseIf Len(MatchedIndent) = 12 Then
        
                    CounterNm6th = CounterNm6th + 1
                    'セルに出力
                    Cells(SelectionArea.Row + i, FIRST_INDENT_ROW + 5).Value = CStr(CounterNm1st) & "-" & CStr(CounterNm2nd) _
                    & "-" & CStr(CounterNm3rd) & "-" & CStr(CounterNm4th) & "-" _
                    & CStr(CounterNm5th) & "-" & CStr(CounterNm6th) & "." & OtherSentence
            
                End If
             End If
        Next j
    Next i
    

End Sub
