Declare Function OpenClipboard Lib "user32" (Optional ByVal hwnd As Long = 0) As Long
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function EmptyClipboard Lib "user32" () As Long

'繰り返しフラグ
Public isLogging As Boolean
'変える行数※使ってない
Public LocalOffset As Long

'スクリーンショットがクリップボードにコピーされると
'エクセルに自動で張り付ける

'キャプチャーモードを起動する
Public Sub Capture()

    MsgBox "エビデンスの自動貼り付けを開始します"
    
    'ESCキーで停止
    Application.OnKey "{ESC}", "StopCapture"
    
    'クリップボードを空にする
    a = OpenClipboard()
    a = EmptyClipboard()
    a = CloseClipboard()
    
    
    '繰り返しフラグをオンにする
    isLogging = True
    
    '画像がクリップボードに保存されたら起動
    If Application.ClipboardFormats(1) = xlClipboardFormatBitmap Then
    
        '貼り付け
        ActiveSheet.Paste
        '現在セルから10したに移動
        ActiveSheet.Offset(10, 0).Select
        
            'クリップボードを空にする
            a = OpenClipboard()
            a = EmptyClipboard()
            a = CloseClipboard()
    
    End If
    
    '1秒間隔 Capture関数を起動する
    Application.OnTime Now + TimeValue("00:00:01"), "Capture", , isLogging
    
    Exit Sub

End Sub

'キャプチャーモードを停止する
Public Sub StopCapture()

    'ESCへのキー登録を解除する
    Application.OnKey "{ESC}", ""
    '繰り返しフラグをオフに
    isLogging = False
    
    MsgBox "キャプチャの取得を停止しました"
    
End Sub
