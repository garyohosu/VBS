'http://homepage2.nifty.com/nonnon/Chinamini/20100001/20100305.html
'★VBSでRS232Cのサンプルソース
' COMポート設定
Set MSComm1 = CreateObject("MSCommLib.MSComm")
MSComm1.CommPort = 20
MSComm1.Settings = "9600,N,8,1"

' COMポートオープン
MSComm1.PortOpen = True

' データ送信
MSComm1.Output = "12345あいうえお"

' ACKまたはNAKを受信するまでまつ
Do
    WScript.Sleep(1000)
    RcvDat = RcvDat & MSComm1.Input
Loop While InStr(RcvDat, Chr(6)) <= 0 And _
           InStr(RcvDat, Chr(21)) <= 0

' 受信データ表示
MsgBox RcvDat

' COMポートクローズ
MSComm1.PortOpen = False