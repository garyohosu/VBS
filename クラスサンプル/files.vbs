Option Explicit


call main

sub main
	dim objfiles

	set objfiles = new Files

	objfiles.Filter = "\.TXT$"

	logprintln "[START]" & date & " " & time

	logprintln "===curfile==="

	objfiles.getCurFile
	dim rdFile
	for each rdFile in objfiles.FileList.Item
		logprintln rdFile.path
	next

	logprintln "===curfilesub==="

	objfiles.getCurFileWithSub
	for each rdFile in objfiles.FileList.Item
		logprintln rdFile.path
	next

	logprintln("===file===")
	objfiles.Path="C:\DATA"
	objfiles.getFile
	for each rdFile in objfiles.FileList.Item
		'logprintln rdFile.path
		logprintln rdFile.filename
	next

	logprintln("===filesub===")
	objfiles.Path="C:\DATA"
	objfiles.getFileWithSub
	for each rdFile in objfiles.FileList.Item
		logprintln rdFile.path
		rdFile.Read
		logPrintln rdFile.Lines.count
	next

	msgbox("終了")
end sub
	


class File
	dim m_fullPath
	dim m_objFso
	dim m_objLines
	dim m_file

    Public Property Get Lines
		set Lines = m_objLines
    End Property

    Public Property Get Path
        Path = m_fullPath
		
    End Property

    Public Property Let Path(vPath)
        m_fullPath = vPath
    End Property


    Public Property Get filename
		set m_file = m_objFso.getfile(m_fullPath)
        filename = m_file.name
		set m_file = nothing
    End Property

    Public Property Let filename(vfn)
		set m_file = m_objFso.getfile(m_fullPath)
        m_file.name = vfn
		set m_file = nothing
    End Property

    Private Sub Class_Initialize()
		Set m_objFso = CreateObject("Scripting.FileSystemObject")
		set m_objLines = new ArrayList
    End Sub

	sub Read
		dim RDobjFile

		Set RDobjFile = m_objFso.OpenTextFile(m_fullPath, 1, False)

		If Err.Number > 0 Then
		    WScript.Echo "Open Error"
		Else
			m_objLines.Clear
		    Do Until RDobjFile.AtEndOfStream
		        m_objLines.add  RDobjFile.ReadLine
		    Loop
		End If

		RDobjFile.Close
		Set RDobjFile = Nothing
	end sub

	sub Save
		dim WRobjFile

		Set WRobjFile = m_objFso.OpenTextFile(m_fullPath, 2, True)

		If Err.Number > 0 Then
    		WScript.Echo "Open Error"
		Else
			dim wrLine
			for each wrLine in m_objLines.FileList
	    		WRobjFile.WriteLine wrLine
			next
		End If

		WRobjFile.Close
		Set WRobjFile = Nothing
		'Scripting.FileSystemObjectはファイル操作をするオブジェクトです。
		'OpenTextFileでファイルを開きます。
		'第1パラメータ→ 必ず指定します。
		'第2パラメータ→ 1:読み取り専用、2:書き込み専用、8:ファイルの最後に書き込み
		'第3パラメータ→ True(規定値):新しいファイルを作成する、False:新しいファイルを作成しない
		'第4パラメータ→ 0(規定値):ASCII ファイルとして開く、-1:Unicode ファイルとして開く、-2:システムの既定値で開く
		'ReadLineでテキストファイルを読み込みます。
		'Closeでファイルをクローズします。
	end sub
end class

class Files
	dim m_Path
	Dim m_FileList
	Dim m_RegExp	'正規表現
	Dim m_filter	'正規表現文字列
	'https://msdn.microsoft.com/ja-jp/library/ms974570.aspx
	'^	文字列の先頭にのみマッチします。
	'$	文字列の末尾にのみマッチします。
	'\b	任意の単語境界にマッチします。
	'\B	任意の単語境界以外の位置にマッチします。
	
    Public Property Get FileList
        set FileList = m_FileList
    End Property

    Public Property Let FileList(vFileList)
        set m_FileList = vFileList
    End Property

    Public Property Get Path
        Path = m_Path
    End Property

    Public Property Let Path(vPath)
        m_Path = vPath
    End Property

    Public Property Get Filter
        Filter = m_Filter
    End Property

    Public Property Let Filter(vFilter)
        m_Filter = vFilter
    End Property

    Private Sub Class_Initialize()
		''連想配列の作成
		'Set Item = WScript.CreateObject("Scripting.Dictionary")
		Set m_FileList = new ArrayList

		Set m_RegExp = new RegExp
		m_RegExp.IgnoreCase = True	'大文字と小文字の区別をしない

		m_Filter = ""
    End Sub

	sub getFile	'フォルダのファイル一覧を取得
		m_FileList.Clear
		findFile m_Path,false
	end sub

	sub getFileWithSub	'サブフォルダこみでファイル一覧
		m_FileList.Clear
		findFile m_Path,true
	end sub

	sub getCurFile	'現在のフォルダのファイル一覧
		m_Path = apppath
		getFile
	end sub

	sub getCurFileWithSub	'現在のフォルダからサブフォルダ込でファイル一覧
		m_Path = apppath
		getFileWithSub
	end sub

	function apppath
	    dim fso
	    set fso = createObject("Scripting.FileSystemObject")
	    apppath = fso.getParentFolderName(WScript.ScriptFullName)
	end function

	sub findFile(path,f_subfolder)
		dim fso
		set fso = createObject("Scripting.FileSystemObject")

		dim folder
		set folder = fso.getFolder(path)

		' ファイル一覧
		dim rdfile
		for each rdfile in folder.files
			if m_filter<>"" then
				m_RegExp.pattern = m_filter
				if m_RegExp.Test(rdfile) = true then
					'msgbox(m_RegExp.Test(m_filter) & rdfile)
					dim newfile
					set newfile = new file
					newfile.path = rdfile
				    m_FileList.add newfile
				end if
			else
				dim newfile1
				set newfile1 = new file
				newfile1.path = rdfile
			    m_FileList.add newfile1
			end if
		next 

		if f_subfolder = true then
			' サブフォルダ一覧
			dim subfolder
			for each subfolder in folder.subfolders
			    findFile subfolder,f_subfolder
			next
		end if
	end sub

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
