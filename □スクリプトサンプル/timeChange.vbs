'============================================================
'[���e]�����f�[�^�̓��t��ύX����
'[�g����]�@�ύX�������t�@�C���̂���t�H���_��
'         �udir /S/B >fileList.txt�v�����s���t�@�C���̃��X�g���쐬
'          �����t�H���_�ł��̃X�N���v�g�����s����
'[����]CSV�t�@�C���łQ���ږڂ�YYYYMMDD�̌`���œ��t���ۑ�����Ă��邱��
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
	'        �ύX�O�@�@�@�ύX��
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
