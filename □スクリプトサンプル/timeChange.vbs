'============================================================
'[内容]検査データの日付を変更する
'[使い方]　変更したいファイルのあるフォルダで
'         「dir /S/B >fileList.txt」を実行しファイルのリストを作成
'          同じフォルダでこのスクリプトを実行する
'[条件]CSVファイルで２項目目にYYYYMMDDの形式で日付が保存されていること
' DATE       VER  NAME    COMMENT
'2017/03/09  1.00 HANTANI NEW
'============================================================
Option Explicit

call main

function dateChange(str)
	dim dic
	dim ret
	
	ret = str

	set dic = CreateObject("Scripting.dictionary")
	'        変更前　　　変更後
	dic.add "20170207","20170306"
	dic.add "20170208","20170307"
	dic.add "20170209","20170307"
	dic.add "20170213","20170307"
	dic.add "20170214","20170308"
	dic.add "20170216","20170309"
	dic.add "20170217","20170310"
	dic.add "20170218","20170311"
	dic.add "20170220","20170313"
	dic.add "20170221","20170314"
	dic.add "20170222","20170315"
	dic.add "20170223","20170316"
	dic.add "20170224","20170317"
	dic.add "20170301","20170317"
	dic.add "20170306","20170317"

	dim rdline
	dim oldDate
	dim newDate

	rdline = split(str,",")
	oldDate = rdline(1)
	if len(oldDate)=8 then
		if dic.exists(oldDate) = True then
			newDate = dic(oldDate)
			ret = replace(str,oldDate,newDate)
			logPrint("[" & oldDate & ">" & newDate & "]" & ret)
		else
			logPrint("[SKIP2]" & str)
		end if
	else
		logPrint("[SKIP1]" & str)
	end if
	dateChange = ret
end function


sub main

	logPrint("[START]" & date & " " & time)

	dim fileList
	dim outFile

	set fileList = new clsText
	set outFile = new clsText

	fileList.FileName = "fileList.txt"
	fileList.ReadFile
	

	dim i
	dim fileName
	dim n
	
	n = fileList.Count

	for i = 0 to n -1
		outfile.FileName = fileList.Items(i)
		logPrint(outfile.FileName)
		outfile.Clear
		outfile.ReadFile
		dim j
		dim m
		m = outfile.Count
		for j = 0 to m -1
			outfile.Change j,dateChange(outfile.Items(j))
		next
		outfile.WriteFile
	next
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
	    			WRobjFile.WriteLine item
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
