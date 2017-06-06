' ----------------------------------------------------------- 
' TCP/IPで通信するサンプル(VBS/VBA)
' 通信オブジェクト生成
' [regsvr32.exe NONCOMSCK.OCX]が必要
Set Winsock1 = CreateObject("NonComSck.Winsock")

' TCP/IP接続
Winsock1.Connect "127.0.0.1", 8888

' TCP/IP接続待ち
Do While Winsock1.State = 6
    WScript.Sleep(500)
Loop

' データ送信(文字列をByte配列に変換して送信)
Winsock1.SendData _
    Winsock1.StrToByteArray(Chr(2) & "1234567890" & Chr(3))

' データ受信
Winsock1.Start_EventForScript()
Do
    WScript.Sleep(500)
    Evt = Winsock1.GetEventParameters()
    If Ubound(Evt) >= 0 Then
        ' Evt(0) : イベント名
        If Evt(0) = "DataArrival" Then
            ' Evt(9) : 受信データのByte配列
            ' Byte配列を文字列に変換
            MsgBox Winsock1.ByteArrayToStr(Evt(9))
            Exit Do
        End If
    End If
Loop
Winsock1.End_EventForScript()

' TCP/IP切断
Winsock1.Close2

MsgBox "終了"
' ----------------------------------------------------------- 
