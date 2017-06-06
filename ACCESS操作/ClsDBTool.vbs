'DATE       Name    Ver  Comment
'2011/11/11 Hantani 1.00 新規作成
'2017/05/12 Hantani 1.00 VBS用 新規作成

Class ClsDBTool

	dim adOpenDynamic
	dim adLockReadOnly
	dim adOpenKeyset
	dim adLockOptimistic

	Private DbtCn

    Private Sub Class_Initialize()
		Set DbtCn = CreateObject("ADODB.Connection")'ﾃﾞｰﾀﾍﾞｰｽ接続用ｺﾈｸｼｮﾝｵﾌﾞｼﾞｪｸﾄ
		adOpenDynamic = 2
		adLockReadOnly = 1
		adOpenKeyset = 1
		adLockOptimistic = 3
    End Sub

	Private Sub ErrMsg
					MsgBox "エラー番号:" & Err.Number & vbCrLf & "説明:" & Err.Description & vbCrLf & "ソース:" & Err.Source & vbCrLf
	end sub

	'==============================================================================
	'[機　能]　データベース接続
	'[関数名]　Function setDB(DBPath As String) As Boolean
	'[入　力]　DBPath:データベースファイル名
	'[出　力]　True:接続成功　False:接続失敗
	'[注　記]　使用例
	'
	'		if dbt.setDB(DB_Path) = False Then MsgBox("DB接続エラー")
	'				'DB_Pathにはデータベースファイルのパスを入れる
	'==============================================================================
	Public function setDB(DBPath)
	    
		On Error Resume Next

	    'DB連結ｵﾌﾞｼﾞｪｸﾄの設定
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
	'[機　能]　SQL実行
	'[関数名]　Function exeSQL(ByVal cmd As String) As Boolean
	'[入　力]　cmd:SQL文
	'[出　力]　True:成功　False:失敗
	'[注　記]  複数のデータの一括追加や更新などのクエリーを実行
	'
	'使用例
	'        if dbt.exeSQL("INSERT INTO Data (工程管理番号,製品シリアル番号) VALUES ('A00001','1AB00001')") = False Then MsgBox("データベースエラー")
	'
	'        if dbt.exeSQL("UPDATE Data SET 工程管理番号 = 'A00001',CPU基板検査完了 = 'TRUE' WHERE 製品シリアル番号 ='1AB00001'") = False Then MsgBox("データベースエラー")
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
	'キーを元に別のレコードに値を設定する。
	'[機　能]　キーを元に別のレコードに値を設定する。
	'[関数名]　Public Function setValue(TableName As String, RecordName As String, Key As String, DataRecordName As String, rData As String) As Boolean
	'[入　力]　TableName:テーブル名
	'　　　　　RecordName：検索するフィールド名
	'　　　　　Key：検索する値
	'　　　　　DataRecordName：値を取得するフィールド名
	'　　　　　rData：設定する値
	'[出　力]　True:書換成功　False:書換失敗
	'[注　記]
	'
	'使用例
	'
	'        if dbt.setValue("Data", "工程管理番号", "A00001", "製品シリアル番号","1AB00001") = False Then MsgBox("データベースエラー")
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
	'[機　能]　キーを元に別のレコードの値を取得する
	'[関数名]　Public Function getValue(TableName As String, RecordName As String, Key As String, DataRecordName As String) As String
	'[入　力]　TableName:テーブル名
	'　　　　　RecordName：検索するフィールド名
	'　　　　　Key：検索する値
	'　　　　　DataRecordName：値を取得するフィールド名
	'[出　力]　"":データ無し　"error":データベースエラー　その他：取得した値
	'[注　記]　RecordName = DataRecordNameも可。その場合、フィールドに検索する値があるか
	'　　　　　（既に登録済みか？）を確認することに使える。
	'使用例
	'
	'		例　工程管理番号を与えて、製品シリアル番号を得る
	'
	'        DBRet = dbt.getValue("Data", "工程管理番号", frmMain.txtSerialNo.Text, "製品シリアル番号")
	'
	'　　　　　　工程管理番号を与えて、工程管理番号を得る（登録済みか確認）
	'
	'        DBRet = dbt.getValue("Data", "工程管理番号", frmMain.txtSerialNo.Text, "工程管理番号")
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
	'[機　能]　1フィールドのレコードの値を取得する
	'[関数名]　Public Function getFieldValue(TableName As String, DataRecordName As String) As String
	'[入　力]　TableName:テーブル名
	'　　　　　DataRecordName：値を取得するフィールド名
	'[出　力]　"":データ無し　"error":データベースエラー　その他：取得した値(カンマ区切り）
	'[注　記]　
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
	'[機　能]　キーを元に別の1レコードの値を取得する
	'[関数名]　Public Function get1RecodeValue(TableName As String, RecordName As String, Key As String) As String
	'[入　力]　TableName:テーブル名
	'　　　　　RecordName：検索するフィールド名
	'　　　　　Key：検索する値
	'[出　力]　"":データ無し　"error":データベースエラー　その他：取得した値
	'[注　記]
	'使用例
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
	'[機　能]　１レコードのみの追加。
	'[関数名]　Public Function addValue(TableName As String, RecordName As String, rData As String) As Boolean
	'[入　力]　TableName:テーブル名
	'　　　　　RecordName：設定するフィールド名
	'　　　　　rData：設定する値
	'[出　力]　True:追加成功　False:追加失敗
	'[注　記]
	'
	'使用例
	'
	'        if dbt.addValue("Data", "工程管理番号", "A00001") = False Then MsgBox("データベースエラー")
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
	'[機　能]　データベース切断
	'[関数名]　Public Function CloseDB() As Boolean
	'[入　力]　無し
	'[出　力]　True:切断成功　False:切断失敗
	'[注　記]　setDB()と対になっている。データベース接続が必要な場合
	'　　　　　setDB()⇒処理⇒CloseDB()を行う。以下のクラス開放を行うまで、何度でも繰り返し可能。
	'
	'　　　　　使用例
	'　　　　　dbt.CloseDB
	'==============================================================================
	Public Function CloseDB() 'As Boolean
	    If DbtCn.State = 1 Then DbtCn.Close
	End Function

    Private Sub Class_Terminate()
		Set DbtCn = Nothing
    End Sub

End Class
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

