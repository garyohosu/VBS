'DAT CSV�ϊ��c�[��
'
'[�@�\]�@OK.DAT����w�肵�����t�ȍ~�̃f�[�^�̂ݐ؂�o���B
'        ���̃X�N���v�g������t�H���_�ƃT�u�t�H���_�ɂ���OK.DAT��S�ď�������B
'
'      �A�t�@�C������[�t�H���_��].CSV�ɕύX����B
'        �� \401\OK.DAT �� 401.CSV
'
'[����]OK.DAT�� 2��ڂɌ�������YYYYMMDD�̃t�H�[�}�b�g�ő��݂��Ă��邱�ƁB
'
'[�g�p���@]�@�ϊ�����@��t�H���_(�� \RS-387-0001)�Ŗ{�X�N���v�g���_�u���N���b�N����B
'          �A�O��̏o�ד�����͂���(�� 2016/06/08 �� 20160608�����)
'          �B�e�T�u�t�H���_����OK.DAT���t�H���_��.CSV�ɕϊ�����A���̃f�[�^�͇A�̓��t�ȍ~�̃f�[�^�݂̂ɂȂ�B
'
'DATE       VER  NAME    COMMENT
'2016/06/08 0.00 HANTANI �V�K�쐬
'2016/06/08 0.01 HANTANI ���������u2016/06/03�v�ł������悤�C��

Option Explicit

Call Main()

sub Main

	dim ff
	dim oldShipDate

	set ff = new ClsFiles

	ff.filter = "OK.DAT"
	ff.getFile(apppath)
	dim item

	oldShipDate=inputbox("�O��o�ד��Hyyyymmdd")

	if len(oldShipDate) = 8 then
		for each item in ff.FileList.item
			dim folders
			dim dataPath
			dim filename
			dim newFilename

			dataPath = item
			folders = split(dataPath,"\")

		    filename ="OK.DAT"
			newFilename = folders(UBound(folders) - 1) &".CSV"
			'Wscript.echo dataPath & ":" & filename & ":" & newFilename & ":" & oldShipDate

			csv2dat dataPath,filename,newFilename,oldShipDate
		next
	end if
	msgbox("�I��")
end sub

'�g���q��CSV����DAT�Ɋ����Â��f�[�^�͍폜
sub csv2dat(path,filename,newFilename,shipDate)
	dim objFsoRD
	dim objFileRD
	dim objFsoWR
	dim objFileWR

	dim rdLine
	dim dataItem
	dim InspDate

	Set objFsoRD = CreateObject("Scripting.FileSystemObject")

	Set objFileRD = objFsoRD.OpenTextFile(path , 1, False)

	Set objFsoWR = CreateObject("Scripting.FileSystemObject")

	path = ucase(path)

	Set objFileWR = objFsoWR.OpenTextFile(replace(path,"OK.DAT","") & "\" & newFilename, 2, True)

	If Err.Number > 0 Then
	    WScript.Echo "Open Error"
	Else
	    Do Until objFileRD.AtEndOfStream
	        rdLine = objFileRD.ReadLine
			dataItem = split(rdLine,",")
			'WScript.Echo rdLine
			if UBound(dataItem) > 2 then
				InspDate=replace(dataItem(1),"/","")
				if InspDate > shipDate and dataItem(1) <> "�N����" then
					objFileWR.WriteLine rdLine
				end if
			end if
	    Loop
	End If

	objFileRD.Close
	Set objFileRD = Nothing
	Set objFsoRD = Nothing

	objFileWR.Close
	Set objFileWR = Nothing
	Set objFsoWR = Nothing

end sub

'���݂̃p�X��Ԃ�
function apppath
    dim fso
    set fso = createObject("Scripting.FileSystemObject")
    apppath = fso.getParentFolderName(WScript.ScriptFullName)
end function

'�t�@�C���I���_�C�A���O���o��
function OpenFileDialog(title)

    Dim obj, filename
    Set obj = CreateObject("Excel.Application")
    filename = obj.GetOpenFilename("ALL File,*.*",1,title)
    obj.Quit
    Set obj = Nothing
    If filename <> False Then
          OpenFileDialog = filename
    End If

end function

class ClsFiles

	public FileList
	Dim objFSO          ' FileSystemObject
	public Filter

	Private Sub Class_Initialize()
		'Set FileList = CreateObject("System.Collections.ArrayList")
		Set FileList = new ArrayList
		Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	End Sub

	'�t�@�C�����ύX
	public sub rename(orgFilePath,newFileName)

		dim sfo
		dim gfile

		set sfo = createobject("scripting.filesystemobject")
		set gfile = sfo.getfile(orgFilePath)
		gfile.name = newFileName
		set sfo = nothing
		set gfile = nothing

	end sub

	'�t�@�C���폜
	public sub del(strDelFile)
		objFSO.DeleteFile strDelFile, True
	end sub

	'�t�H���_���݃`�F�b�N
	public function FolderExists(FolderPath)
		FolderExists = objFSO.FolderExists(FolderPath) 
	end function

	'�t�@�C�����݃`�F�b�N
	public function FileExists(FolderPath)
		FolderExists = objFSO.FileExists(FolderPath) 
	end function

	'�t�@�C�����𓾂�
	public function getFileName(FolderPath)
		getFileName = objFSO.getFileName(FolderPath) 
	end function

	'�t�H���_���𓾂�
	public function getFolderName(FolderPath)
		dim FileName

		if right(FolderPath,1)="\" then
			FolderPath = left(FolderPath,len(FolderPath) - 1)
		end if
		FileName = objFSO.getFileName(FolderPath)
		if filename <> "" then
			FolderPath = Replace(FolderPath,filename,"")
		end if
		getFolderName = FolderPath
	end function

	public sub getFile(path)
		FindFolder objFSO.getFolder(path)
	end sub

	' �t�H���_�ċA�I�����֐��i���ʏo�͂� FileList)
	private Sub FindFolder(ByVal objParentFolder)

		Dim objFile
		Dim resultLine
		For Each objFile In objParentFolder.Files
			'if instr(1,ucase(objFile.Name),ucase(Filter)) > 0 then
			if ucase(objFile.Name)=ucase(Filter) then
				FileList.add objFile.ParentFolder & "\" & objFile.Name
			end if
		    'FIND_RESULT_FILE_OBJ.Write(objFile.ParentFolder & "\" & objFile.Name & ",")
		    'FIND_RESULT_FILE_OBJ.Write(objFile.Size & ",") 'byte
		    'FIND_RESULT_FILE_OBJ.Write(objFile.DateLastModified & ",")
		    'FIND_RESULT_FILE_OBJ.Write(Fix(Date() - objFile.DateLastModified) & ",")
		    'FIND_RESULT_FILE_OBJ.Write(objFile.DateLastAccessed & ",")
		    'FIND_RESULT_FILE_OBJ.Write(Fix(Date() - objFile.DateLastAccessed))
		    'FIND_RESULT_FILE_OBJ.WriteLine("")
		Next

		Dim objSubFolder    ' �T�u�t�H���_
		For Each objSubFolder In objParentFolder.SubFolders
		    FindFolder objSubFolder
		Next

	End Sub

	' �t�H���_���ċA�I�ɍ쐬����
	Sub CreateFolderEx(ByVal strPath)
		WScript.Echo  "CreateFolderEx:[" & strPath & "]"

	    Dim strParent   ' �e�t�H���_
	    strParent = objFSO.GetParentFolderName(strPath)
	    If objFSO.FolderExists(strParent) = True Then
	        If objFSO.FolderExists(strPath) <> True Then
	            objFSO.CreateFolder strPath
	        End If
	    Else
			if strParent <> "" then
		        CreateFolderEx strParent
			end if
	        objFSO.CreateFolder strPath
	    End If
	End Sub

    Private Sub Class_Terminate()
		set FileList = nothing
		Set objFSO = nothing
    End Sub

end class

'���݂̃p�X��Ԃ�
function apppath
    dim fso
    set fso = createObject("Scripting.FileSystemObject")
    apppath = fso.getParentFolderName(WScript.ScriptFullName)
end function



'���I�z���
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



