Option Explicit

dim WRobjFso
dim WRobjFile

Set WRobjFso = CreateObject("Scripting.FileSystemObject")
Set WRobjFile = WRobjFso.OpenTextFile("c:\b.txt", 2, True)

If Err.Number > 0 Then
    WScript.Echo "Open Error"
Else
    WRobjFile.WriteLine "書き込む文字列です。"
End If

WRobjFile.Close
Set WRobjFile = Nothing
Set WRobjFso = Nothing

'Scripting.FileSystemObjectはファイル操作をするオブジェクトです。
'OpenTextFileでファイルを開きます。
'第1パラメータ→ 必ず指定します。
'第2パラメータ→ 1:読み取り専用、2:書き込み専用、8:ファイルの最後に書き込み
'第3パラメータ→ True(規定値):新しいファイルを作成する、False:新しいファイルを作成しない
'第4パラメータ→ 0(規定値):ASCII ファイルとして開く、-1:Unicode ファイルとして開く、-2:システムの既定値で開く
'ReadLineでテキストファイルを読み込みます。
'Closeでファイルをクローズします。

