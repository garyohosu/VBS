Option Explicit

dim RDobjFile
dim RDobjFso

Set RDobjFso = CreateObject("Scripting.FileSystemObject")
Set RDobjFile = RDobjFso.OpenTextFile("c:\a.txt", 1, False)

If Err.Number > 0 Then
    WScript.Echo "Open Error"
Else
    Do Until RDobjFile.AtEndOfStream
        WScript.Echo RDobjFile.ReadLine & vbCrLf
    Loop
End If

RDobjFile.Close
Set RDobjFile = Nothing
Set RDobjFso = Nothing


'Scripting.FileSystemObject�̓t�@�C�����������I�u�W�F�N�g�ł��B
'OpenTextFile�Ńt�@�C�����J���܂��B
'��1�p�����[�^�� �K���w�肵�܂��B
'��2�p�����[�^�� 1:�ǂݎ���p�A2:�������ݐ�p�A8:�t�@�C���̍Ō�ɏ�������
'��3�p�����[�^�� True(�K��l):�V�����t�@�C�����쐬����AFalse:�V�����t�@�C�����쐬���Ȃ�
'��4�p�����[�^�� 0(�K��l):ASCII �t�@�C���Ƃ��ĊJ���A-1:Unicode �t�@�C���Ƃ��ĊJ���A-2:�V�X�e���̊���l�ŊJ��
'ReadLine�Ńe�L�X�g�t�@�C����ǂݍ��݂܂��B
'Close�Ńt�@�C�����N���[�Y���܂��B

