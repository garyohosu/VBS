'http://wsh.style-mods.net/topic11.htm

'イベントログに書き込むには

'WSHでイベントログに書き込むには、WshShell オブジェクトの LogEvent メソッドを使用します。

'イベントログは「アプリケーション」「セキュリティ」「システム」といったカテゴリに分かれており、
'書き込めるのは「アプリケーション」です。

'また、「成功」「エラー」「警告」「情報」「成功の監査」「失敗の監査」といった種類も
'指定することができます。

sub logPrintln(s)
	logPrint(s & vbcrlf)
end sub

sub logPrint(s)
	dim WshShell
	Set WshShell = WScript.CreateObject("WScript.Shell")
	WshShell.LogEvent 0, "[" & WScript.ScriptName & "]" & s
	Set WshShell = Nothing
end sub

sub logPrintlnError(s)
	logPrint(s & vbcrlf)
end sub

sub logPrintError(s)
	dim WshShell
	Set WshShell = WScript.CreateObject("WScript.Shell")
	WshShell.LogEvent 1,  "[" & WScript.ScriptName & "]" & s
	Set WshShell = Nothing
end sub

'Set WshShell = WScript.CreateObject("WScript.Shell")
'WshShell.LogEvent 0, "「成功」のメッセージ"
'WshShell.LogEvent 1, "「エラー」のメッセージ"
'WshShell.LogEvent 2, "「警告」のメッセージ"
'WshShell.LogEvent 4, "「情報」のメッセージ"
'WshShell.LogEvent 8, "「成功の監査」のメッセージ"
'WshShell.LogEvent 16, "「失敗の監査」のメッセージ"

