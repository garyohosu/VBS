'http://blog.livedoor.jp/shikaku_li/archives/7515681.html
'VBAでは使えるDoEventsがVBSでは使えないのでSleepで代用する


Option Explicit

Call Main()

Sub Main()

	dim Alarm
	dim msg
	dim strDate

	Alarm="15:10"

	msg="面談15:40　3階応接室"

	do while 1
		strDate = FormatDateTime(Now, 4)

		if strDate >= Alarm then
			msgbox msg, vbSystemModal
			exit do
		end if
		WScript.Sleep 100
	loop

end sub
