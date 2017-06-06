'DAT CSV変換ツール
'
'[機能]①OK.DATから指定した日付以降のデータのみ切り出す。
'        このスクリプトがあるフォルダとサブフォルダにあるOK.DATを全て処理する。
'
'      ②ファイル名を[フォルダ名].CSVに変更する。
'        例 \401\OK.DAT → 401.CSV
'
'[条件]OK.DATの 2列目に検査日がYYYYMMDDのフォーマットで存在していること。
'
'[使用方法]①変換する機種フォルダ(例 \RS-387-0001)で本スクリプトをダブルクリックする。
'          ②前回の出荷日を入力する(例 2016/06/08 → 20160608を入力)
'          ③各サブフォルダ内のOK.DATがフォルダ名.CSVに変換され、中のデータは②の日付以降のデータのみになる。
'
'DATE       VER  NAME    COMMENT
'2016/06/08 0.00 HANTANI 新規作成
'2016/06/08 0.01 HANTANI 検査日が「2016/06/03」でも動くよう修正

Option Explicit

Call Main()

sub Main

	dim ff
	dim oldShipDate

	set ff = new ClsFiles

	ff.filter = "OK.DAT"
	ff.getFile(apppath)
	dim item

	oldShipDate=inputbox("前回出荷日？yyyymmdd")

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
	msgbox("終了")
end sub

'拡張子をCSVからDATに換え古いデータは削除
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
				if InspDate > shipDate and dataItem(1) <> "年月日" then
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

'現在のパスを返す
function apppath
    dim fso
    set fso = createObject("Scripting.FileSystemObject")
    apppath = fso.getParentFolderName(WScript.ScriptFullName)
end function

'ファイル選択ダイアログを出す
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
		FolderExists = objFSO.FileExists(FolderPath) 
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
function apppath
    dim fso
    set fso = createObject("Scripting.FileSystemObject")
    apppath = fso.getParentFolderName(WScript.ScriptFullName)
end function



'動的配列版
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



