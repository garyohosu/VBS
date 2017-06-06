'VBS Utility Class
'DATE       VER  NAME    COMMENT
'2016/05/10 0.00 HANTANI �V�K�쐬


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

'Include "Utility.vbs"  ' �O���t�@�C���uLibrary.vbs�v���捞��
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
'	Dim strCopyFile     ' �R�s�[�Ώۃt�@�C��
'	Dim strCopyFolder   ' �R�s�[��t�H���_��

'	strCopyFile = "C:\A*"
'	strCopyFolder = "C:\Work"
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
end sub



sub shell(x)
'WScript.Shell�I�u�W�F�N�g��.Run "���s����R�}���h",�E�B���h�E�T�C�Y�w��,�������[�h�w��
'
'�E�B���h�E�T�C�Y�̎w��́A�R�}���h���s���̃E�B���h�E�T�C�Y�𐔒l�Ŏw�肷��B
'�w��ł���l	���s���̃E�B���h�E�T�C�Y(���)
'0	��\��
'1	�ʏ�E�B���h�E
'2	�ŏ���
'3	�ő剻
'
'�������[�h�w��
'false:�񓯊�
'True:����

	dim objShell
	Set objShell = CreateObject("WScript.Shell")
	objShell.Run x,0,false
	Set objShell = Nothing

end sub
