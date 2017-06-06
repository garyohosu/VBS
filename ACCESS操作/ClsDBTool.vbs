'DATE       Name    Ver  Comment
'2011/11/11 Hantani 1.00 �V�K�쐬
'2017/05/12 Hantani 1.00 VBS�p �V�K�쐬

Class ClsDBTool

	dim adOpenDynamic
	dim adLockReadOnly
	dim adOpenKeyset
	dim adLockOptimistic

	Private DbtCn

    Private Sub Class_Initialize()
		Set DbtCn = CreateObject("ADODB.Connection")'�ް��ް��ڑ��p�ȸ��ݵ�޼ު��
		adOpenDynamic = 2
		adLockReadOnly = 1
		adOpenKeyset = 1
		adLockOptimistic = 3
    End Sub

	Private Sub ErrMsg
					MsgBox "�G���[�ԍ�:" & Err.Number & vbCrLf & "����:" & Err.Description & vbCrLf & "�\�[�X:" & Err.Source & vbCrLf
	end sub

	'==============================================================================
	'[�@�@�\]�@�f�[�^�x�[�X�ڑ�
	'[�֐���]�@Function setDB(DBPath As String) As Boolean
	'[���@��]�@DBPath:�f�[�^�x�[�X�t�@�C����
	'[�o�@��]�@True:�ڑ������@False:�ڑ����s
	'[���@�L]�@�g�p��
	'
	'		if dbt.setDB(DB_Path) = False Then MsgBox("DB�ڑ��G���[")
	'				'DB_Path�ɂ̓f�[�^�x�[�X�t�@�C���̃p�X������
	'==============================================================================
	Public function setDB(DBPath)
	    
		On Error Resume Next

	    'DB�A����޼ު�Ă̐ݒ�
	    If DbtCn.State = 1 Then
	        DbtCn.Close
	    End If

		DbtCn.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & DBPath & ";"
		
		if Err.Number = 0 Then
			setDB = True
		else
			setDB = False
			ErrMsg
		end if
	End function


	'==============================================================================
	'[�@�@�\]�@SQL���s
	'[�֐���]�@Function exeSQL(ByVal cmd As String) As Boolean
	'[���@��]�@cmd:SQL��
	'[�o�@��]�@True:�����@False:���s
	'[���@�L]  �����̃f�[�^�̈ꊇ�ǉ���X�V�Ȃǂ̃N�G���[�����s
	'
	'�g�p��
	'        if dbt.exeSQL("INSERT INTO Data (�H���Ǘ��ԍ�,���i�V���A���ԍ�) VALUES ('A00001','1AB00001')") = False Then MsgBox("�f�[�^�x�[�X�G���[")
	'
	'        if dbt.exeSQL("UPDATE Data SET �H���Ǘ��ԍ� = 'A00001',CPU��������� = 'TRUE' WHERE ���i�V���A���ԍ� ='1AB00001'") = False Then MsgBox("�f�[�^�x�[�X�G���[")
	'
	'==============================================================================
	Function exeSQL(cmd) 'As Boolean

	    Dim Ret 'As String
	    Dim Rs 'As Object
	    Dim mySql 'As String
	    Dim DbCmd 'As ADODB.Command

		On Error Resume Next

	    DbtCn.BeginTrans

	    mySql = cmd

	'    Set DbCmd = New ADODB.Command
	    Set DbCmd = CreateObject("ADODB.Command")

	    DbCmd.ActiveConnection = DbtCn
	    DbCmd.CommandText = mySql
	    DbCmd.Execute
	    Set DbCmd = Nothing

	    DbtCn.CommitTrans

		if Err.Number = 0 Then
	    	exeSQL = True
		else
			exeSQL = False
			ErrMsg
		end if

	End Function

	'==============================================================================
	'�L�[�����ɕʂ̃��R�[�h�ɒl��ݒ肷��B
	'[�@�@�\]�@�L�[�����ɕʂ̃��R�[�h�ɒl��ݒ肷��B
	'[�֐���]�@Public Function setValue(TableName As String, RecordName As String, Key As String, DataRecordName As String, rData As String) As Boolean
	'[���@��]�@TableName:�e�[�u����
	'�@�@�@�@�@RecordName�F��������t�B�[���h��
	'�@�@�@�@�@Key�F��������l
	'�@�@�@�@�@DataRecordName�F�l���擾����t�B�[���h��
	'�@�@�@�@�@rData�F�ݒ肷��l
	'[�o�@��]�@True:���������@False:�������s
	'[���@�L]
	'
	'�g�p��
	'
	'        if dbt.setValue("Data", "�H���Ǘ��ԍ�", "A00001", "���i�V���A���ԍ�","1AB00001") = False Then MsgBox("�f�[�^�x�[�X�G���[")
	'==============================================================================
	Public Function setValue(TableName, RecordName, Key, DataRecordName, rData) 'As Boolean
	    Dim Ret 'As String
	    Dim Rs 'As Object
	    Dim mySql 'As String
	    
		On Error Resume Next
	    
	    mySql = "SELECT * FROM " & TableName
	    mySql = mySql & " WHERE " & RecordName & "='" & Key & "'"
	    
	    Set Rs = CreateObject("ADODB.Recordset")
	    
	    'Rs.Open mySql, DbtCn, adOpenDynamic, adLockReadOnly
	    Rs.Open mySql, DbtCn, adOpenKeyset, adLockOptimistic
'	    Rs.Open mySql, DbtCn
            Rs.MoveFirst
	        
	    If Rs.EOF = True Or Rs.BOF = True Then
	        setValue = False
	    Else
	        Rs(DataRecordName) = rData
	        Rs.Update
	    
	        setValue = True
	    End If

	    Rs.Close
	    Set Rs = Nothing

	    If setValue = True and  Err.Number <> 0  Then
	        setValue = False
			ErrMsg
		end if

	End Function
	'==============================================================================
	'[�@�@�\]�@�L�[�����ɕʂ̃��R�[�h�̒l���擾����
	'[�֐���]�@Public Function getValue(TableName As String, RecordName As String, Key As String, DataRecordName As String) As String
	'[���@��]�@TableName:�e�[�u����
	'�@�@�@�@�@RecordName�F��������t�B�[���h��
	'�@�@�@�@�@Key�F��������l
	'�@�@�@�@�@DataRecordName�F�l���擾����t�B�[���h��
	'[�o�@��]�@"":�f�[�^�����@"error":�f�[�^�x�[�X�G���[�@���̑��F�擾�����l
	'[���@�L]�@RecordName = DataRecordName���B���̏ꍇ�A�t�B�[���h�Ɍ�������l�����邩
	'�@�@�@�@�@�i���ɓo�^�ς݂��H�j���m�F���邱�ƂɎg����B
	'�g�p��
	'
	'		��@�H���Ǘ��ԍ���^���āA���i�V���A���ԍ��𓾂�
	'
	'        DBRet = dbt.getValue("Data", "�H���Ǘ��ԍ�", frmMain.txtSerialNo.Text, "���i�V���A���ԍ�")
	'
	'�@�@�@�@�@�@�H���Ǘ��ԍ���^���āA�H���Ǘ��ԍ��𓾂�i�o�^�ς݂��m�F�j
	'
	'        DBRet = dbt.getValue("Data", "�H���Ǘ��ԍ�", frmMain.txtSerialNo.Text, "�H���Ǘ��ԍ�")
	'
	'==============================================================================
	Public Function getValue(TableName, RecordName, Key, DataRecordName)' As String
	    Dim Ret 'As String
	    Dim Rs 'As Object
	    Dim mySql 'As String
	    
		On Error Resume Next
	    
	    mySql = "SELECT * FROM " & TableName
	    mySql = mySql & " WHERE " & RecordName & "='" & Key & "'"
	    
	    Set Rs = CreateObject("ADODB.Recordset")
	    
'	    Rs.Open mySql, DbtCn, adOpenDynamic, adLockReadOnly
	    Rs.Open mySql, DbtCn
            Rs.MoveFirst
	    
	        
	    If Rs.EOF = True Or Rs.BOF = True Then
	        getValue = ""
	    Else
	        getValue = Rs(DataRecordName)
	    End If
	        
	    Rs.Close
	    Set Rs = Nothing
	    
	    If Err.Number <> 0  Then
	        getValue = "error"
			ErrMsg
		end if	

	End Function
	'==============================================================================
	'[�@�@�\]�@1�t�B�[���h�̃��R�[�h�̒l���擾����
	'[�֐���]�@Public Function getFieldValue(TableName As String, DataRecordName As String) As String
	'[���@��]�@TableName:�e�[�u����
	'�@�@�@�@�@DataRecordName�F�l���擾����t�B�[���h��
	'[�o�@��]�@"":�f�[�^�����@"error":�f�[�^�x�[�X�G���[�@���̑��F�擾�����l(�J���}��؂�j
	'[���@�L]�@
	'
	'==============================================================================
	Public Function getFieldValue(TableName, DataRecordName)' As String
	    Dim Ret 'As String
	    Dim Rs 'As Object
	    Dim mySql 'As String
	    
	    On Error Resume Next
	    
	    mySql = "SELECT * FROM " & TableName


	    Set Rs = CreateObject("ADODB.Recordset")

	    
'	    Rs.Open mySql, DbtCn, adOpenDynamic, adLockReadOnly
	    Rs.Open mySql, DbtCn

            Rs.MoveFirst
	    
	    ret = ""
	    Do While Rs.EOF = False and Rs.BOF = False

	    	If Rs.EOF = True Or Rs.BOF = True Then
	        	exit do
	    	Else
			if ret = "" then
				ret = Rs(DataRecordName)
			else
				ret = ret & "," & Rs(DataRecordName)
			end if
	    	End If
		Rs.MoveNext
	    loop

	    Rs.Close
	    Set Rs = Nothing

	    getFieldValue = ret

	    If Err.Number <> 0  Then
	        getFieldValue = "error"
		ErrMsg
	    end if

	End Function

	'==============================================================================
	'[�@�@�\]�@�L�[�����ɕʂ�1���R�[�h�̒l���擾����
	'[�֐���]�@Public Function get1RecodeValue(TableName As String, RecordName As String, Key As String) As String
	'[���@��]�@TableName:�e�[�u����
	'�@�@�@�@�@RecordName�F��������t�B�[���h��
	'�@�@�@�@�@Key�F��������l
	'[�o�@��]�@"":�f�[�^�����@"error":�f�[�^�x�[�X�G���[�@���̑��F�擾�����l
	'[���@�L]
	'�g�p��
	'
	'
	'==============================================================================
	Public Function get1Recode(TableName, RecordName, Key)' As String
	    Dim Ret 'As String
	    Dim Rs 'As Object
	    Dim mySql 'As String
	    
		'On Error Resume Next
	    
	    mySql = "SELECT * FROM " & TableName
	    mySql = mySql & " WHERE " & RecordName & "='" & Key & "'"
	    
	    Set Rs = CreateObject("ADODB.Recordset")
	    
'	    Rs.Open mySql, DbtCn, adOpenDynamic, adLockReadOnly
	    Rs.Open mySql, DbtCn
        Rs.MoveFirst
	    If Rs.EOF = True Or Rs.BOF = True Then
	        get1Recode = ""
	    Else
	        get1Recode = ""
			dim item
			for each item in Rs.Fields
				msgbox(item.name)
				if get1Recode = "" then
					get1Recode = item.name & ":" & Rs(item.name)
				else
					get1Recode = get1Recode & "," & item.name & ":" & Rs(item.name) 
				end if
			next
	        'getValue = Rs(DataRecordName)
	    End If
	        
	    Rs.Close
	    Set Rs = Nothing
	    
	    If Err.Number <> 0  Then
	        getValue = "error"
			ErrMsg
		end if	

	End Function

	'==============================================================================
	'[�@�@�\]�@�P���R�[�h�݂̂̒ǉ��B
	'[�֐���]�@Public Function addValue(TableName As String, RecordName As String, rData As String) As Boolean
	'[���@��]�@TableName:�e�[�u����
	'�@�@�@�@�@RecordName�F�ݒ肷��t�B�[���h��
	'�@�@�@�@�@rData�F�ݒ肷��l
	'[�o�@��]�@True:�ǉ������@False:�ǉ����s
	'[���@�L]
	'
	'�g�p��
	'
	'        if dbt.addValue("Data", "�H���Ǘ��ԍ�", "A00001") = False Then MsgBox("�f�[�^�x�[�X�G���[")
	'
	'==============================================================================
	Public Function addValue(TableName, RecordName, rData) 'As Boolean
	    
	    Dim Ret 'As String
	    Dim Rs 'As Object
	    Dim mySql 'As String
	    Dim DbCmd 'As ADODB.Command

	    On Error Resume Next

	    mySql = "INSERT INTO " & TableName & " (" & RecordName & ") VALUES ('" & rData & "');"

		if exeSQL(mySql) = True Then
		    addValue = True
	    else
		    addValue = false
		end if

	    If Err.Number <> 0  Then
		    addValue = False
			ErrMsg
		end if

	End Function


	'==============================================================================
	'[�@�@�\]�@�f�[�^�x�[�X�ؒf
	'[�֐���]�@Public Function CloseDB() As Boolean
	'[���@��]�@����
	'[�o�@��]�@True:�ؒf�����@False:�ؒf���s
	'[���@�L]�@setDB()�Ƒ΂ɂȂ��Ă���B�f�[�^�x�[�X�ڑ����K�v�ȏꍇ
	'�@�@�@�@�@setDB()�ˏ�����CloseDB()���s���B�ȉ��̃N���X�J�����s���܂ŁA���x�ł��J��Ԃ��\�B
	'
	'�@�@�@�@�@�g�p��
	'�@�@�@�@�@dbt.CloseDB
	'==============================================================================
	Public Function CloseDB() 'As Boolean
	    If DbtCn.State = 1 Then DbtCn.Close
	End Function

    Private Sub Class_Terminate()
		Set DbtCn = Nothing
    End Sub

End Class
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

