'WScript.Shellオブジェクト名.Run "実行するコマンド",ウィンドウサイズ指定,同期モード指定
'
'ウィンドウサイズの指定は、コマンド実行時のウィンドウサイズを数値で指定する。
'指定できる値	実行時のウィンドウサイズ(状態)
'0	非表示
'1	通常ウィンドウ
'2	最小化
'3	最大化
'
'同期モード指定
'false:非同期
'True:同期

'Set objShell = CreateObject("WScript.Shell")
'objShell.Run "cmd /c ipconfig /all > c:\ip.txt",0,false
sub Shell(cmd)

	dim objShell

	Set objShell = CreateObject("WScript.Shell")
	objShell.Run cmd,1,True

end sub
