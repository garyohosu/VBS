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

'Scripting.FileSystemObject�̓t�@�C�����������I�u�W�F�N�g�ł��B
'OpenTextFile�Ńt�@�C�����J���܂��B
'��1�p�����[�^�� �K���w�肵�܂��B
'��2�p�����[�^�� 1:�ǂݎ���p�A2:�������ݐ�p�A8:�t�@�C���̍Ō�ɏ�������
'��3�p�����[�^�� True(�K��l):�V�����t�@�C�����쐬����AFalse:�V�����t�@�C�����쐬���Ȃ�
'��4�p�����[�^�� 0(�K��l):ASCII �t�@�C���Ƃ��ĊJ���A-1:Unicode �t�@�C���Ƃ��ĊJ���A-2:�V�X�e���̊���l�ŊJ��
'ReadLine�Ńe�L�X�g�t�@�C����ǂݍ��݂܂��B
'Close�Ńt�@�C�����N���[�Y���܂��B

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
