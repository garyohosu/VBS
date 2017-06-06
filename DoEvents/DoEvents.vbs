'http://blog.livedoor.jp/shikaku_li/archives/7515681.html
'VBA‚Å‚ÍŽg‚¦‚éDoEvents‚ªVBS‚Å‚ÍŽg‚¦‚È‚¢‚Ì‚ÅSleep‚Å‘ã—p‚·‚é


Option Explicit

Call Main()

Sub Main()

	dim Alarm
	dim msg
	dim strDate

	Alarm="15:10"

	msg="–Ê’k15:40@3ŠK‰žÚŽº"

	do while 1
		strDate = FormatDateTime(Now, 4)

		if strDate >= Alarm then
			msgbox msg, vbSystemModal
			exit do
		end if
		WScript.Sleep 100
	loop

end sub
