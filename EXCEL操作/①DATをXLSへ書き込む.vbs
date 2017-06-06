' Polaris ブザー鳴動検査データコンバーター(DAT->XLS変換)
'
'DATE      VER  NAME     COMMENT
'2016/08/19 1.00 HANTANI  新規

dim rdLine
dim rdLineList
dim x,y,i
dim item
dim x_start
dim x_end
dim y_start

logPrintln("[START]" & date & " " & time)

' WScript.CreateObjectを使ってブックを開く。
Set excel = WScript.CreateObject("Excel.Application")
Set book = excel.WorkBooks.Open(apppath & "\CSVFileRead_V135_SII.xls")
'Set sheet = book.WorkSheets.Item(1)
Set sheet = book.WorkSheets("CSV_Data")

'前のデータがあればクリア
y_start = 27
x_start = 5
x_end = 23

y = y_start

do while 1
	if sheet.Cells(y, x_start) <> "" then
		for i = x_start to x_end
			sheet.Cells(y, i) = ""
		next
	else
		exit do
	end if
	y = y + 1
loop


dim x_serial
dim x_kisyu
dim x_date1
dim x_data1
dim x_data2
dim C2

x_serial=0
x_kisyu=1
x_date1=2
x_data1=4
x_data2=5

Set objFso = CreateObject("Scripting.FileSystemObject")
Set objFile = objFso.OpenTextFile("DATA101\OK.CSV", 1, False)

If Err.Number > 0 Then
    WScript.Echo "Open Error"
Else
	y = y_start
    Do Until objFile.AtEndOfStream
        rdLine = objFile.ReadLine
		rdLineList = split(rdLine,",")
		logPrintln(rdLine)
		logPrintln("y=" & y & "objFile.AtEndOfStream=" & objFile.AtEndOfStream)
		x = 0
		sheet.Cells(y, x_start + x) = "C:\分布図作成ﾏｸﾛ\RS-335\Type1\量産\基板検査\.\DATA101\OK101.dat"
		x = x + 1
		'sheet.Cells(y, x_start + x) = rdLineList(x_kisyu)'"RS-335"
		sheet.Cells(y, x_start + x) = "RS-335"
		x = x + 1
		sheet.Cells(y, x_start + x) = "0001-101"
		x = x + 1
		sheet.Cells(y, x_start + x) = rdLineList(x_serial)'"S316C143601"
		x = x + 1
		sheet.Cells(y, x_start + x) = "10X"
		x = x + 1
		sheet.Cells(y, x_start + x) = "102"
		x = x + 1
		sheet.Cells(y, x_start + x) = ""
		x = x + 1
		sheet.Cells(y, x_start + x) = "OK"
		x = x + 1
		sheet.Cells(y, x_start + x) = "1.83"
		x = x + 1
		sheet.Cells(y, x_start + x) = "utou"
		x = x + 1

		C2 = rdLineList(x_date1)

		sheet.Cells(y, x_start + x) = LEFT(C2,4) & "/" & MID(C2,5,2) & "/" & MID(C2,7,2)'"2016/03/09"
		x = x + 1
		sheet.Cells(y, x_start + x) = MID(C2,9,2) & ":" & MID(C2,11,2) & ":" & MID(C2,13,2)'"16:44:03"
		x = x + 1
		sheet.Cells(y, x_start + x) = "0 "
		x = x + 1
		sheet.Cells(y, x_start + x) = rdLineList(x_date1)'"20160309164403"
		x = x + 1
		sheet.Cells(y, x_start + x) = ""
		x = x + 1
		sheet.Cells(y, x_start + x) = "1.8"
		x = x + 1
		sheet.Cells(y, x_start + x) = "1.8"
		x = x + 1
		sheet.Cells(y, x_start + x) = rdLineList(x_data1)'"103.24"
		x = x + 1
		sheet.Cells(y, x_start + x) = rdLineList(x_data2)'"13.71"
		x = x + 1

		y = y + 1
    Loop
End If

logPrintln("do exit")

objFile.Close
Set objFile = Nothing
Set objFso = Nothing

book.Save
logPrintln("book.Save")

excel.Quit()
logPrintln("excel.Quit")

WScript.Echo "終了"


Set sheet = Nothing
Set book = Nothing
Set excel = Nothing

logPrintln("[END]" & date & " " & time)


function apppath
    dim fso
    set fso = createObject("Scripting.FileSystemObject")
    apppath = fso.getParentFolderName(WScript.ScriptFullName)
end function

sub logPrintln(s)
	logPrint(s & vbcrlf)
end sub

sub logPrint(s)
	dim objFsoWR
	dim objFileWR
	dim LogFile
	dim SerialFieldNo


	LogFile = apppath & "\log.log"

	Set objFsoWR = CreateObject("Scripting.FileSystemObject")
	Set objFileWR = objFsoWR.OpenTextFile(LogFile, 8, True)

	If Err.Number > 0 Then
	    WScript.Echo "Open Error"
	Else
		objFileWR.WriteLine s
	End If

	objFileWR.Close
	Set objFileWR = Nothing
	Set objFsoWR = Nothing

end sub
