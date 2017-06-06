Option Explicit

Dim FIND_START_FOLDER
FIND_START_FOLDER = "c:\temp"                  '�T���J�nfolder
Dim FIND_RESULT_FILE_NAME
FIND_RESULT_FILE_NAME = "c:\temp\FIND_RESULT.TXT" '�T�����ʈꗗ
Dim FIND_RESULT_FILE_OBJ

Sub Main()

Dim objFSO          ' FileSystemObject

Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
'refer to http://msdn.microsoft.com/ja-jp/library/ie/cc428044.aspx
'2=�����p�Ƃ���open  , True=file�V�K�쐬 , -1=unicode�ŏ���
Set FIND_RESULT_FILE_OBJ = objFSO.OpenTextFile(FIND_RESULT_FILE_NAME,2,True,-1)


FIND_RESULT_FILE_OBJ.Write("#PATH,SIZE(byte),MODIFY DATE,MODIFY DATE AGE,")
FIND_RESULT_FILE_OBJ.Write("ACCESS DATE,ACCESS DATE AGE")
FIND_RESULT_FILE_OBJ.WriteLine("")

FindFolder objFSO.getFolder(FIND_START_FOLDER)

FIND_RESULT_FILE_OBJ.Close

msgbox("END")

End Sub


' �t�H���_�����֐�
Sub FindFolder(ByVal objParentFolder)

Dim objFile
Dim resultLine
For Each objFile In objParentFolder.Files
    FIND_RESULT_FILE_OBJ.Write(objFile.ParentFolder & "\" & objFile.Name & ",")
    FIND_RESULT_FILE_OBJ.Write(objFile.Size & ",") 'byte
    FIND_RESULT_FILE_OBJ.Write(objFile.DateLastModified & ",")
    FIND_RESULT_FILE_OBJ.Write(Fix(Date() - objFile.DateLastModified) & ",")
    FIND_RESULT_FILE_OBJ.Write(objFile.DateLastAccessed & ",")
    FIND_RESULT_FILE_OBJ.Write(Fix(Date() - objFile.DateLastAccessed))
    FIND_RESULT_FILE_OBJ.WriteLine("")

Next

Dim objSubFolder    ' �T�u�t�H���_
For Each objSubFolder In objParentFolder.SubFolders
    FindFolder objSubFolder
Next

End Sub

Main
