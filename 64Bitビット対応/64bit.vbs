
	RunOn32bit
	msgbox("end")


'64bit��VBScript�̏ꍇ��32bit�łŋN�����Ȃ���
Public Sub RunOn32bit

	dim p_objWshShell
	dim p_admWscriptCscript
	dim p_admArrayArguments
	dim p_admArg
	dim p_admCommand


	Set p_objWshShell = CreateObject("Wscript.Shell")
	
	'Environment("Process").Item("PROCESSOR_ARCHITECTURE")��x86�̏ꍇ��32bit�A�����łȂ����64bit��VBScript�������Ă���
	'32bit�������ꍇ�͉������Ȃ��A64bit�������ꍇ�͈ȉ�32bit�ŋN�����Ȃ���
	If p_objWshShell.Environment("Process").Item("PROCESSOR_ARCHITECTURE") <> "x86" Then
		
		'�R�}���h���C���������w�肳��Ă���ꍇ�͂�����ė��p���邽�ߎ擾����
		If Not WScript.Arguments.Count = 0 Then
			For Each p_admArg In Wscript.Arguments
			  p_admArrayArguments = p_admArrayArguments & " """ & p_admArg & """"
			Next
		End If
		
		'WScript.FullName�ŋN�����Ă���v���Z�X�̖��O���킩��̂ŁA�������̂ŋN�����Ȃ�������WScript��CScript�����m�F����
		If InStr(LCase(WScript.FullName), "wscript") > 0 Then
			p_admWscriptCscript = "WScript.exe"
		Else
			p_admWscriptCscript = "CScript.exe"
		End If
		
		'WScript.ScriptFullName�ŃX�N���v�g�̃p�X���擾���A����܂łɎ擾�������Ƃ��킹�ăR�}���h���쐬�AWscript.Shell��Run���\�b�h�Ŏ��s����
		p_admCommand = """" &  p_objWshShell.Environment("Process").Item("windir") & "\SysWOW64\" & p_admWscriptCscript & """ """ & WScript.ScriptFullName & """" & p_admArrayArguments
		p_objWshShell.Run p_admCommand

		'���݂�(64bit��)�v���Z�X���I������
		WScript.Quit
	End If
End Sub
