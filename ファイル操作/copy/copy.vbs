

Option Explicit
On Error Resume Next

Dim objFSO          ' FileSystemObject
Dim strCopyFile     ' �R�s�[�Ώۃt�@�C��
Dim strCopyFolder   ' �R�s�[��t�H���_��

strCopyFile = "C:\A*"
strCopyFolder = "C:\Work"
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
If Err.Number = 0 Then
    ' �R�s�[��t�H���_�����݂��Ȃ��Ƃ��͍쐬����
    If objFSO.FolderExists(strCopyFolder) <> True Then
        objFSO.CreateFolder(strCopyFolder)
    End If

    ' �t�@�C���R�s�[
    objFSO.CopyFile strCopyFile, strCopyFolder, True
    If Err.Number = 0 Then
        WScript.Echo strCopyFile & " �� " & _
            strCopyFolder & " �ɃR�s�[���܂����B"
    Else
        WScript.Echo "�G���[: " & Err.Description
    End If
Else
    WScript.Echo "�G���[: " & Err.Description
End If

Set objFSO = Nothing
