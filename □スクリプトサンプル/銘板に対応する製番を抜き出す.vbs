'============================================================
'[内容]銘板に対応する製番
'[使い方]　銘板ファイルを読み、そこに含まれる製造番号を含む
'         製番読込を出力する
' DATE       VER  NAME    COMMENT
'2017/03/09  1.00 HANTANI NEW
'============================================================
Option Explicit

dim meibanFileName
dim seibanFileName
dim OutputSeibanFileName

meibanFileName = "銘OK.csv"
seibanFileName = "1702_5A.csv"
OutputSeibanFileName = "1703_A"

call main

function csvRead(str,n)
	dim rdline
	dim ret

	rdline = split(str,",")
	ret = rdline(n)
	csvRead = ret
end function

sub main

	logPrint("[START]" & date & " " & time)

	dim meiban
	dim seiban

	set meiban = new clsText
	set seiban = new clsText

	meiban.FileName = meibanFileName
	meiban.ReadFile
	'msgbox(meiban.count)

	seiban.FileName = seibanFileName
	seiban.ReadFile
	'msgbox(seiban.count)

	dim dic
	set dic = CreateObject("Scripting.dictionary")

	dim i
	for i = 0 to meiban.Count - 1
		logPrint(meiban.Items(i))
		logPrint(csvRead(meiban.Items(i),6))
		if len(csvRead(meiban.Items(i),6)) = 11 then
			dic.add csvRead(meiban.Items(i),6),1
		end if
	next

	dim serial

	for i = 0 to seiban.Count - 1
		serial = csvRead(seiban.Items(i),0)
		if dic.exists(serial) = False then
			seiban.Change i,""
			logPrint("SER NOT EXIST:" & serial)
		else
			logPrint("SER EXIST:" & serial)
		end if
	next
	seiban.filename = OutputSeibanFileName
	seiban.WriteFile

	msgbox("END")
	logPrint("[END]" & date & " " & time)
end sub

class clsText

	dim LineData
	dim RDobjFile
	dim RDobjFso
	dim WRobjFso
	dim WRobjFile
	dim m_FileName

    Public Property Get FileName
        FileName = m_FileName
    End Property

    Public Property Let FileName(vData)
        m_FileName = vData
    End Property


    Private Sub Class_Initialize()
        set LineData = new ArrayList
		Set RDobjFso = CreateObject("Scripting.FileSystemObject")
		Set WRobjFso = CreateObject("Scripting.FileSystemObject")
    End Sub

    Private Sub Class_Terminate()
		Set RDobjFso = Nothing
		Set WRobjFso = Nothing
    End Sub

	Public Sub Clear
		LineData.Clear
	end sub

	public Function Count
		Count = LineData.Count
	end Function

	public Function Items(n)
		Items = LineData.Items(n)
	end Function

	public Sub Change(i,x)
		LineData.Change i,x
	end sub

	Public Sub ReadFile
		Set RDobjFile = RDobjFso.OpenTextFile(m_FileName, 1, False)
		If Err.Number > 0 Then
		    WScript.Echo "Open Error"
		Else
		    Do Until RDobjFile.AtEndOfStream
		        LineData.add RDobjFile.ReadLine
		    Loop
		End If
		RDobjFile.Close
		Set RDobjFile = Nothing
	end sub

	Public Sub WriteFile
		Set WRobjFile = WRobjFso.OpenTextFile(m_FileName, 2, True)
		If Err.Number > 0 Then
		    WScript.Echo "Open Error"
		Else
			dim item
			for each item in LineData.item
				if len(item)>0 then
	    				WRobjFile.WriteLine item
				end if
			next
		End If
		WRobjFile.Close
		Set WRobjFile = Nothing
	end sub

	Public Sub AppendFile
		Set WRobjFile = WRobjFso.OpenTextFile(m_FileName, 8, True)
		If Err.Number > 0 Then
		    WScript.Echo "Open Error"
		Else
			for each item in LineData.item
	    		WRobjFile.WriteLine item
			next
		End If
		WRobjFile.Close
		Set WRobjFile = Nothing
	end sub

'Scripting.FileSystemObjectはファイル操作をするオブジェクトです。
'OpenTextFileでファイルを開きます。
'第1パラメータ→ 必ず指定します。
'第2パラメータ→ 1:読み取り専用、2:書き込み専用、8:ファイルの最後に書き込み
'第3パラメータ→ True(規定値):新しいファイルを作成する、False:新しいファイルを作成しない
'第4パラメータ→ 0(規定値):ASCII ファイルとして開く、-1:Unicode ファイルとして開く、-2:システムの既定値で開く
'ReadLineでテキストファイルを読み込みます。
'Closeでファイルをクローズします。

	sub OpenFileDialog(title)

	    Dim obj, filename
	    Set obj = CreateObject("Excel.Application")
	    filename = obj.GetOpenFilename("ALL File,*.*",1,title)
	    obj.Quit
	    Set obj = Nothing
	    If filename <> False Then
	          m_FileName = filename
	    End If

	end sub

end Class



'動的配列版ArrayList
class ArrayList

	private m_Item()
	private m_count

	public sub Add(x)
		ReDim Preserve m_item(m_count)
		If IsObject(x) Then
			set m_item(m_count) = x
		else
			m_item(m_count) = x
		end if
		m_count = m_count + 1
	end sub

	public sub Change(i,x)
		If IsObject(x) Then
			set m_item(i) = x
		else
			m_item(i) = x
		end if
	end sub

	public function Count
		Count = m_count
	end function

	public function Clear
		m_count=0
		Erase m_item
	end function

	public function Item
		Item = m_Item
	end function

	public function Items(n)
		If IsObject(m_Item(n)) Then
			set Items = m_Item(n)
		else
			Items = m_Item(n)
		end if
	end function

end class
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
