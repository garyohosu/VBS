'再帰的にフォルダ、サブフォルダのファイル一覧を得る

Option Explicit

call Main

sub Main

	FileList "msgbox(filename)"

end sub

'ファイル一覧に「exec」に書いてあることを実行する。
'ファイル名は「exec」内では「fileName」と書く
sub FileList(exec)

	dim ff
	set ff = new ClsFiles

	'ff.filter = "para.dat"
	ff.getFile(appPath)
	dim fileName
	for each fileName in ff.FileList.Item
		Execute exec	
	next
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

	'ファイル名変更
	public sub rename(orgFilePath,newFileName)

		dim sfo
		dim gfile

		set sfo = createobject("scripting.filesystemobject")
		set gfile = sfo.getfile(orgFilePath)
		gfile.name = newFileName
		set sfo = nothing
		set gfile = nothing

	end sub

	'ファイル削除
	public sub del(strDelFile)
		objFSO.DeleteFile strDelFile, True
	end sub

	'フォルダ存在チェック
	public function FolderExists(FolderPath)
		FolderExists = objFSO.FolderExists(FolderPath) 
	end function

	'ファイル存在チェック
	public function FileExists(FolderPath)
		FileExists= objFSO.FileExists(FolderPath) 
	end function

	'ファイル名を得る
	public function getFileName(FolderPath)
		getFileName = objFSO.getFileName(FolderPath) 
	end function

	'フォルダ名を得る
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

	' フォルダ再帰的検索関数（結果出力は FileList)
	private Sub FindFolder(ByVal objParentFolder)

		Dim objFile
		Dim resultLine
		For Each objFile In objParentFolder.Files
			if Filter<>"" then
				'if instr(1,ucase(objFile.Name),ucase(Filter)) > 0 then
				if ucase(objFile.Name)=ucase(Filter) then
					FileList.add objFile.ParentFolder & "\" & objFile.Name
				end if
			else
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

		Dim objSubFolder    ' サブフォルダ
		For Each objSubFolder In objParentFolder.SubFolders
		    FindFolder objSubFolder
		Next

	End Sub

	' フォルダを再帰的に作成する
	Sub CreateFolderEx(ByVal strPath)
		WScript.Echo  "CreateFolderEx:[" & strPath & "]"

	    Dim strParent   ' 親フォルダ
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

'現在のパスを返す
function appPath
    dim fso
    set fso = createObject("Scripting.FileSystemObject")
    appPath = fso.getParentFolderName(WScript.ScriptFullName)
end function


'動的配列版 ArrayList
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
