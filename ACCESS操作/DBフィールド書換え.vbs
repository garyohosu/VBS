Option Explicit

'DATE       Name    Ver  Comment
'2011/11/11 Hantani 1.00 新規作成
'2017/05/12 Hantani 1.00 VBS用 新規作成

Call Main()

Sub Main()

	dim DB_Path
	dim DB_Table
	dim DB_Key_Field
	dim DB_Field
	dim DB_Value
	dim KEY_File
	dim objFile
	dim objFso
	dim DB_Key
	dim Ret
	dim Count
	dim NGCount
	dim msg

	DB_Path="Z:\RS-387-9001\DB\RS-387.mdb"
	DB_Table="Data"
	DB_Key_Field = "工程管理番号"
	DB_Field="機種名"
	DB_Value="RS-387-0101"
	KEY_File = "Serial.txt"
	Count = 0
	NGCount = 0

	logPrintln("[START]" & Date & " " & time)

	Dim m_dbt
	Set m_dbt = new ClsDBTool

	call m_dbt.setDB(DB_Path)
	
	Set objFso = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFso.OpenTextFile(KEY_File, 1, False)

	If Err.Number > 0 Then
	    WScript.Echo "Open Error"
	Else
	    Do Until objFile.AtEndOfStream
	        DB_Key = objFile.ReadLine
			Ret = m_dbt.setValue(DB_Table,DB_Key_Field,DB_Key,DB_Field,DB_Value)
			if ret = True then
				count = count + 1
			else
				ngcount = ngcount + 1
			end if
			logPrintln(Ret & ":" & DB_Table & ":" & DB_Key_Field & ":" & DB_Key & ":" & DB_Field & ":" & DB_Value)
	    Loop
	End If

	objFile.Close
	Set objFile = Nothing
	Set objFso = Nothing

	m_dbt.CloseDB

	msg = count & "件書き換えました。" & ngcount & "件エラーです"
	logPrintln(msg)
	logPrintln("[END]" & Date & " " & time)
	msgbox(msg)

end sub

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

function apppath
    dim fso
    set fso = createObject("Scripting.FileSystemObject")
    apppath = fso.getParentFolderName(WScript.ScriptFullName)
end function

sub logPrintln(s)
	logPrint(s & vbcrlf)
end sub

sub logPrint(s)
	dim objFsoWR
	dim objFileWR
	dim LogFile

	LogFile = apppath & "\log.log"

	Set objFsoWR = CreateObject("Scripting.FileSystemObject")
	Set objFileWR = objFsoWR.OpenTextFile(LogFile, 8, True)

	If Err.Number > 0 Then
	    WScript.Echo "Open Error"
	Else
		objFileWR.WriteLine s
	End If

	objFileWR.Close
	Set objFileWR = Nothing
	Set objFsoWR = Nothing

end sub
