'���K�\�����g���ăp�^�[���Ɉ�v���邩���ׂ�

'New���g����RegExp�I�u�W�F�N�g���擾����B�p�^�[����RegExp�I�u�W�F�N�g��Pattern�v���p�e�B�ɑ������B

Dim re
Set re = New RegExp
re.Pattern = "^A.*$"
If re.Test("ABC") Then
    WScript.Echo "matched"
End If
'�啶���������𖳎�����ꍇ��IgnoreCase��True��������B

re.IgnoreCase = True
'�ŏ��̈�v�݂̂ł͂Ȃ��A������S�̂������Ώۂɂ���ꍇ�ɂ�Global��True��������B

re.Global = True
'���K�\�����g���Ĉ�v���镶������擾����

'RegExp�I�u�W�F�N�g��Execute���g���B

Dim re, matches
Set re = New RegExp
re.Pattern = "^A.*$"
Set matches = re.Execute("ABC")
If matches.Count > 0 Then
    WScript.Echo matches(0)
End If
'�o�͌��ʂ́uABC�v�ƂȂ�B

'���K�\�����g���ăL���v�`��������������擾����

'Item�v���p�e�B�𗘗p����SubMatches���擾����B

Dim re, matches
Set re = New RegExp
re.Pattern = "^([A-F]+)_([A-F]+).+"
Set matches = re.Execute("ABC_DEF_GHI")
If matches.Count > 0 Then
    WScript.Echo matches.Item(0).SubMatches.Item(0)
    WScript.Echo matches.Item(0).SubMatches.Item(1)
End If
'�o�͌��ʂ́uABC�v�uDEF�v�ƂȂ�B

'���K�\�����g���ĕ������u������

'RegExp�I�u�W�F�N�g��Replace���g���B

Dim re, matches
Set re = New RegExp
re.Pattern = "^ABC_"
WScript.Echo re.Replace("ABC_DEF_GHI", "XXX_")
'�o�͌��ʂ́uXXX_DEF_GHI�v�ƂȂ�B

'���K�\�����g���ĕ������u������i����Q�Ɓj

'RegExp�I�u�W�F�N�g��Replace���g���Ēu�������񒆂�$�����̌`���ŎQ�Ƃ���B

Dim re, matches
Set re = New RegExp
re.Pattern = "^([A-Z]{2})(.*)"
WScript.Echo re.Replace("ABCDE", "XX$2")
'�o�͌��ʂ́uXXCDE�v�ƂȂ�B

