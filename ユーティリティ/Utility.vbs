'VBS Utility Class
'DATE       VER  NAME    COMMENT
'2016/05/10 0.00 HANTANI 新規作成


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

'Include "Utility.vbs"  ' 外部ファイル「Library.vbs」を取込み
'
'Sub Include(ByVal strFile)
'  Dim objFSO , objStream , strDir
'
'  Set objFSO = WScript.CreateObject("Scripting.FileSystemObject") 
'  strDir = objFSO.GetFile(WScript.ScriptFullName).ParentFolder 
'
'  Set objStream = objFSO.OpenTextFile(strDir & "\" & strFile, 1)
'
'  ExecuteGlobal objStream.ReadAll() 
'  objStream.Close 
'
'  Set objStream = Nothing 
'  Set objFSO = Nothing
'End Sub


sub mkdir(path)

	Dim ObjFso

	Set ObjFso=WScript.CreateObject("Scripting.FileSystemObject")

	If ObjFso.FolderExists(path) = False Then
		ObjFso.Createfolder(path)
	End If
	set ObjFso = Nothing

end sub


sub copy(strCopyFile,strCopyFolder)

	On Error Resume Next

	Dim objFSO          ' FileSystemObject
'	Dim strCopyFile     ' コピー対象ファイル
'	Dim strCopyFolder   ' コピー先フォルダ名

'	strCopyFile = "C:\A*"
'	strCopyFolder = "C:\Work"
	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
	If Err.Number = 0 Then
	    ' コピー先フォルダが存在しないときは作成する
	    If objFSO.FolderExists(strCopyFolder) <> True Then
	        objFSO.CreateFolder(strCopyFolder)
	    End If

	    ' ファイルコピー
	    objFSO.CopyFile strCopyFile, strCopyFolder, True
	    If Err.Number = 0 Then
	        WScript.Echo strCopyFile & " を " & _
	            strCopyFolder & " にコピーしました。"
	    Else
	        WScript.Echo "エラー: " & Err.Description
	    End If
	Else
	    WScript.Echo "エラー: " & Err.Description
	End If

	Set objFSO = Nothing
end sub



sub shell(x)
'WScript.Shellオブジェクト名.Run "実行するコマンド",ウィンドウサイズ指定,同期モード指定
'
'ウィンドウサイズの指定は、コマンド実行時のウィンドウサイズを数値で指定する。
'指定できる値	実行時のウィンドウサイズ(状態)
'0	非表示
'1	通常ウィンドウ
'2	最小化
'3	最大化
'
'同期モード指定
'false:非同期
'True:同期

	dim objShell
	Set objShell = CreateObject("WScript.Shell")
	objShell.Run x,0,false
	Set objShell = Nothing

end sub
