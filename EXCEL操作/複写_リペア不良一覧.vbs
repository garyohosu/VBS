' WScript.GetObject���g���ău�b�N���J���ăf�[�^��ǂ�

dim rdLine
dim rdLineList
dim x,y,i
dim item

'Set book = WScript.GetObject("���y�A�s�ǈꗗ.xls")
'Set sheet = book.WorkSheets.Item(1)
''WScript.Echo sheet.Cells(2, 1).Text

' WScript.CreateObject���g���ău�b�N���J���B
Set excel = WScript.CreateObject("Excel.Application")
Set book = excel.WorkBooks.Open(apppath & "\���y�A�s�ǈꗗ.xls")
Set sheet = book.WorkSheets.Item(1)

'�O�̃f�[�^������΃N���A
y = 13
do while 1
	if sheet.Cells(y, 1) <> "" then
		for i = 1 to 7
			sheet.Cells(y, i) = ""
		next
	else
		exit do
	end if
loop

Set objFso = CreateObject("Scripting.FileSystemObject")
Set objFile = objFso.OpenTextFile("���y�A�s�ǈꗗ.csv", 1, False)

If Err.Number > 0 Then
    WScript.Echo "Open Error"
Else
	y = 0
    Do Until objFile.AtEndOfStream
		y = y + 1
        rdLine = objFile.ReadLine
		rdLineList = split(rdLine,",")
		x = 0
		for each item in rdLineList
			x = x + 1
			sheet.Cells(y, x) = item
		next
    Loop
End If

objFile.Close
Set objFile = Nothing
Set objFso = Nothing

'book.SaveAs("c:\temp\Book10.xls")
excel.Quit()
WScript.Echo "�I��"

Set sheet = Nothing
Set book = Nothing
Set excel = Nothing

function apppath
    dim fso
    set fso = createObject("Scripting.FileSystemObject")
    apppath = fso.getParentFolderName(WScript.ScriptFullName)
end function


