Option Explicit


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
		Set RDWobjFso = CreateObject("Scripting.FileSystemObject")
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
		LineData.Change(i,x)
	end Function

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
