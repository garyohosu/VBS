'5. Dictionary �I�u�W�F�N�g
'VBScript �ɂ̓n�b�V���ϐ��͂���܂��񂪁AScripting.Dictionary �I�u�W�F�N�g��p���邱�Ƃő�ւ����\�ł��B
'Dictionary �I�u�W�F�N�g�̐錾�Ƒ��

Dim hash

Set hash = CreateObject("Scripting.Dictionary")
hash.Add "a", 1
hash.Add "b", 2

'Dictionary �̗v�f�̎Q�ƂƑ��
' �v�f�̎Q��
WScript.Echo hash("a") ' 1
WScript.Echo hash("b") ' 2

' �v�f�̑��
hash("a") = 5
hash("b") = 7

'Dictionary �̃v���p�e�B�ƃ��\�b�h
' �L�[�̎擾
keys = hash.Keys

' �l�̎擾
values = hash.Items

' �L�[�̑��݊m�F
hash.Exists("a")

' �L�[�̍폜
hash.Remove "a"
