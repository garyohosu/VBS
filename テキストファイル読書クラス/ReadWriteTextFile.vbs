Option Explicit

Main()


sub Main()

	dim tf
	dim s

	set tf = new ClsTxtFile

	'�P�s������
	tf.FileOpen "C:\Temp\test.txt","W"
	tf.Print "line1"
	tf.Print "line2"
	tf.FileClose

	'��s�ǂݏo��
	tf.FileOpen "C:\Temp\test.txt","R"
	do while not tf.eof
		msgbox(tf.Input)
	loop
	tf.FileClose

	msgbox("ReadAll")

	'�ǉ�������
	tf.FileOpen "C:\Temp\test.txt","A"
	tf.Print "line3"
	tf.Print "line4"
	tf.FileClose

	'�ꊇ�ǂݏo��
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


	'Scripting.FileSystemObject�̓t�@�C�����������I�u�W�F�N�g�ł��B
	'OpenTextFile�Ńt�@�C�����J���܂��B
	'��1�p�����[�^�� �K���w�肵�܂��B
	'��2�p�����[�^�� 1:�ǂݎ���p�A2:�������ݐ�p�A8:�t�@�C���̍Ō�ɏ�������
	'��3�p�����[�^�� True(�K��l):�V�����t�@�C�����쐬����AFalse:�V�����t�@�C�����쐬���Ȃ�
	'��4�p�����[�^�� 0(�K��l):ASCII �t�@�C���Ƃ��ĊJ���A-1:Unicode �t�@�C���Ƃ��ĊJ���A-2:�V�X�e���̊���l�ŊJ��
	'ReadLine�Ńe�L�X�g�t�@�C����ǂݍ��݂܂��B
	'Close�Ńt�@�C�����N���[�Y���܂��B
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

'���I�z���ArrayList
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

