Option Explicit

call Main

sub Main

	dim ff
	dim I

	set ff = new ClsFiles

	ff.filter = "para.dat"
	ff.getFile(apppath)
	dim item
	for each item in ff.FileList.Item
	'for I = 0 to ff.FileList.size - 1
		'WScript.echo item
		'ChangePara(item)
		'ff.del(item)
		'ff.rename item & "1","para.dat"

		ChangePara(ff.FileList.item(I))
		ff.del(ff.FileList.item(I))
		ff.rename ff.FileList.item(I) & "1","para.dat"

	next
end sub


sub ChangePara(path)
	dim RDobjFso
	dim RDobjFile
	dim WRobjFso
	dim WRobjFile
	dim rdData
	
	Dim FS
	Dim filename
	Dim FolderPath
	
	dim OutPath
	dim rdLine
	dim wrLine
	dim ff

	set ff = new ClsFiles


	OutPath = path+"1"

	Set RDobjFso = CreateObject("Scripting.FileSystemObject")
	Set RDobjFile = RDobjFso.OpenTextFile(path, 1, False)

	Set WRobjFso = CreateObject("Scripting.FileSystemObject")
	Set WRobjFile = WRobjFso.OpenTextFile(OutPath, 2, True)

	Set FS = CreateObject("Scripting.FileSystemObject")

	WScript.Echo "[FileName:" & path & "]"

	If Err.Number > 0 Then
	    WScript.Echo "Open Error"
	Else
		if not RDobjFile.AtEndOfStream  then
			rdLine = RDobjFile.ReadLine
	        WScript.Echo "[Skip]" & rdLine & vbCrLf	'1�s�ڂ�ǂݔ�΂�
			WRobjFile.WriteLine rdLine
		end if
	    Do Until RDobjFile.AtEndOfStream
			rdLine = RDobjFile.ReadLine
			if instr(1,rdLine,"RS-387-0") > 0 then
				wrLine = Replace(rdLine,"RS-387-0","RS-387-9")
		        WScript.Echo "[" & rdLine & "]->[" & wrLine & "]" & vbCrLf
				WRobjFile.WriteLine wrLine
				rdData = split(wrLine,",")
				'�t�H���_�����݂��邩�`�F�b�N���Ȃ���΍쐬
				filename = FS.getFileName(rdData(0))
				FolderPath = rdData(0)

				if instr(1,FolderPath,"deneb_eeprom_data") = 0 then
					if right(FolderPath,1) <> "\" then
						if filename <> "" then
							FolderPath = Replace(FolderPath,filename,"")
						end if
					end if
					WScript.Echo FolderPath
					if FS.FolderExists(FolderPath) = False Then
						WScript.Echo "[�V�K�쐬]"
						ff.CreateFolderEx(FolderPath)
						'FS.CreateFolder(FolderPath)
				        If Err.Number = 0 Then
				            WScript.Echo "�t�H���_ " & FolderPath & " ���쐬���܂����B"
				        Else
				            WScript.Echo strMessage = "�G���[: " & Err.Description
				        End If					
					end if
				end if
			else
				WRobjFile.WriteLine rdLine
			end if
	    Loop
	End If

	RDobjFile.Close
	Set RDobjFile = Nothing
	Set RDobjFso = Nothing

	WRobjFile.Close
	Set WRobjFile = Nothing
	Set WRobjFso = Nothing

end sub


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
