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

'Set objShell = CreateObject("WScript.Shell")
'objShell.Run "cmd /c ipconfig /all > c:\ip.txt",0,false
sub Shell(cmd)

	dim objShell

	Set objShell = CreateObject("WScript.Shell")
	objShell.Run cmd,1,True

end sub
