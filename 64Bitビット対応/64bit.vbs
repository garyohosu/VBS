
	RunOn32bit
	msgbox("end")


'64bit版VBScriptの場合は32bit版で起動しなおす
Public Sub RunOn32bit

	dim p_objWshShell
	dim p_admWscriptCscript
	dim p_admArrayArguments
	dim p_admArg
	dim p_admCommand


	Set p_objWshShell = CreateObject("Wscript.Shell")
	
	'Environment("Process").Item("PROCESSOR_ARCHITECTURE")がx86の場合は32bit、そうでなければ64bitでVBScriptが動いている
	'32bitだった場合は何もしない、64bitだった場合は以下32bitで起動しなおす
	If p_objWshShell.Environment("Process").Item("PROCESSOR_ARCHITECTURE") <> "x86" Then
		
		'コマンドライン引数が指定されている場合はそれを再利用するため取得する
		If Not WScript.Arguments.Count = 0 Then
			For Each p_admArg In Wscript.Arguments
			  p_admArrayArguments = p_admArrayArguments & " """ & p_admArg & """"
			Next
		End If
		
		'WScript.FullNameで起動しているプロセスの名前がわかるので、同じもので起動しなおすためWScriptかCScriptかを確認する
		If InStr(LCase(WScript.FullName), "wscript") > 0 Then
			p_admWscriptCscript = "WScript.exe"
		Else
			p_admWscriptCscript = "CScript.exe"
		End If
		
		'WScript.ScriptFullNameでスクリプトのパスを取得し、これまでに取得した情報とあわせてコマンドを作成、Wscript.ShellのRunメソッドで実行する
		p_admCommand = """" &  p_objWshShell.Environment("Process").Item("windir") & "\SysWOW64\" & p_admWscriptCscript & """ """ & WScript.ScriptFullName & """" & p_admArrayArguments
		p_objWshShell.Run p_admCommand

		'現在の(64bitの)プロセスを終了する
		WScript.Quit
	End If
End Sub
