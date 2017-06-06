Option Explicit

Call Main()

Sub Main()

	dim Alarm
	dim msg
	dim strDate

	Alarm="11:09"

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
