Option Explicit

dim WRobjFso
dim WRobjFile

Set WRobjFso = CreateObject("Scripting.FileSystemObject")
Set WRobjFile = WRobjFso.OpenTextFile("c:\b.txt", 2, True)

If Err.Number > 0 Then
    WScript.Echo "Open Error"
Else
    WRobjFile.WriteLine "�������ޕ�����ł��B"
End If

WRobjFile.Close
Set WRobjFile = Nothing
Set WRobjFso = Nothing

'Scripting.FileSystemObject�̓t�@�C�����������I�u�W�F�N�g�ł��B
'OpenTextFile�Ńt�@�C�����J���܂��B
'��1�p�����[�^�� �K���w�肵�܂��B
'��2�p�����[�^�� 1:�ǂݎ���p�A2:�������ݐ�p�A8:�t�@�C���̍Ō�ɏ�������
'��3�p�����[�^�� True(�K��l):�V�����t�@�C�����쐬����AFalse:�V�����t�@�C�����쐬���Ȃ�
'��4�p�����[�^�� 0(�K��l):ASCII �t�@�C���Ƃ��ĊJ���A-1:Unicode �t�@�C���Ƃ��ĊJ���A-2:�V�X�e���̊���l�ŊJ��
'ReadLine�Ńe�L�X�g�t�@�C����ǂݍ��݂܂��B
'Close�Ńt�@�C�����N���[�Y���܂��B

