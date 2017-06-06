

Option Explicit
On Error Resume Next

Dim objFSO          ' FileSystemObject
Dim strCopyFile     ' コピー対象ファイル
Dim strCopyFolder   ' コピー先フォルダ名

strCopyFile = "C:\A*"
strCopyFolder = "C:\Work"
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
