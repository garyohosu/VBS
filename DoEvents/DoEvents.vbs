'http://blog.livedoor.jp/shikaku_li/archives/7515681.html
'VBA�ł͎g����DoEvents��VBS�ł͎g���Ȃ��̂�Sleep�ő�p����


Option Explicit

Call Main()

Sub Main()

	dim Alarm
	dim msg
	dim strDate

	Alarm="15:10"

	msg="�ʒk15:40�@3�K���ڎ�"

	do while 1
		strDate = FormatDateTime(Now, 4)

		if strDate >= Alarm then
			msgbox msg, vbSystemModal
			exit do
		end if
		WScript.Sleep 100
	loop

end sub
