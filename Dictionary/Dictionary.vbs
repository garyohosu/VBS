
Option Explicit
Dim objDictionary
'�A�z�z��̍쐬
Set objDictionary = WScript.CreateObject("Scripting.Dictionary")
'�L�[�ƒl�̒ǉ�
objDictionary.Add "kona","Konata Izumi"
objDictionary.Add "kaga","Kagami Hiiragi"
objDictionary.Add "tuka","Tsukasa Hiiragi"
objDictionary.Add "yuki","Miyuki Takara"

'�l�̎Q��
MsgBox "�L�[�ukaga�v�̒l�́A�u" & objDictionary("kaga") & "�v�ł��B"

'�L�[�̑��݊m�F
If objDictionary.Exists("yuki") Then
	MsgBox "�L�[�uyuki�v�́A���݂��܂��B"
End If

'�v�f�̐���\��
MsgBox "Dictinary�Ɋ܂܂��v�f���F" & objDictionary.Count

Dim strKey, strMessage
'�����̗�
For Each strKey In objDictionary.Keys
        strMessage = strMessage & objDictionary(strKey) & vbCrLf
Next
MsgBox strMessage
objDictionary.RemoveAll() '�����̑S�v�f�폜
objDictionary = Nothing
