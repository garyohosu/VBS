Option Explicit

Main()


sub Main()

	dim tf
	dim s

	set tf = new ClsTxtFile

	'１行書込み
	tf.FileOpen "C:\Temp\test.txt","W"
	tf.Print "line1"
	tf.Print "line2"
	tf.FileClose

	'一行読み出し
	tf.FileOpen "C:\Temp\test.txt","R"
	do while not tf.eof
		msgbox(tf.Input)
	loop
	tf.FileClose

	msgbox("ReadAll")

	'追加書込み
	tf.FileOpen "C:\Temp\test.txt","A"
	tf.Print "line3"
	tf.Print "line4"
	tf.FileClose

	'一括読み出し
	tf.FileOpen "C:\Temp\test.txt","R"
	tf.ReadAll
	for each s in tf.Lines.Item
		msgbox(s)
	next
	tf.FileClose

end sub


class ClsTxtFile

	dim objFso
	dim objFile
	public Lines


	'Scripting.FileSystemObjectはファイル操作をするオブジェクトです。
	'OpenTextFileでファイルを開きます。
	'第1パラメータ→ 必ず指定します。
	'第2パラメータ→ 1:読み取り専用、2:書き込み専用、8:ファイルの最後に書き込み
	'第3パラメータ→ True(規定値):新しいファイルを作成する、False:新しいファイルを作成しない
	'第4パラメータ→ 0(規定値):ASCII ファイルとして開く、-1:Unicode ファイルとして開く、-2:システムの既定値で開く
	'ReadLineでテキストファイルを読み込みます。
	'Closeでファイルをクローズします。
	public function FileOpen(filename,mode)
		Set objFso = CreateObject("Scripting.FileSystemObject")
		Set Lines = new ArrayList

		select case ucase(mode)
			case "W"
				Set objFile = objFso.OpenTextFile(filename, 2, True)
			case "A"
				Set objFile = objFso.OpenTextFile(filename, 8, True)
			case "R"
				Set objFile = objFso.OpenTextFile(filename, 1, False)
		end select
	
	end function

	public sub Print(s)

		If Err.Number > 0 Then
		    WScript.Echo "Open Error"
		Else
		    objFile.WriteLine s
		End If

	end sub

	public function EOF
		EOF = objFile.AtEndOfStream
	end function

	public function Input
		If Err.Number > 0 Then
		    WScript.Echo "Open Error"
		Else
	        Input =  objFile.ReadLine
		End If
	end function

	public sub ReadAll

		If Err.Number > 0 Then
		    WScript.Echo "Open Error"
		Else
			Lines.Clear
		    Do Until objFile.AtEndOfStream
		        Lines.add objFile.ReadLine
		    Loop
		End If

	end sub

	public function FileClose
		objFile.Close
		Set objFile = Nothing
		Set objFso = Nothing
		Set Lines = Nothing
	end function

end class

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

