VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsDBTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'使用には以下への参照が必要
'Microsoft DAO 3.51 Object Library
'Microsoft ActiveX Data Objects 2.8 Library

'

'DATE       Name    Ver  Comment
'2011/11/11 Hantani 1.00 新規作成


Private DbtCn As New ADODB.Connection         'ﾃﾞｰﾀﾍﾞｰｽ接続用ｺﾈｸｼｮﾝｵﾌﾞｼﾞｪｸﾄ

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
Function exeSQL(ByVal cmd As String) As Boolean
    Dim Ret As String
    Dim Rs As Object
    Dim mySql As String
    Dim DbCmd As ADODB.Command

    On Error GoTo ERRHANDLER

    DbtCn.BeginTrans

    mySql = cmd

    Set DbCmd = New ADODB.Command
    DbCmd.ActiveConnection = DbtCn
    DbCmd.CommandText = mySql
    DbCmd.Execute
    Set DbCmd = Nothing

    DbtCn.CommitTrans

    exeSQL = True
    Exit Function

ERRHANDLER:

    exeSQL = False
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
Public Function setValue(TableName As String, RecordName As String, Key As String, DataRecordName As String, rData As String) As Boolean
    Dim Ret As String
    Dim Rs As Object
    Dim mySql As String
    
    On Error GoTo ERRHANDLER
    
    mySql = "SELECT * FROM " & TableName
    mySql = mySql & " WHERE " & RecordName & "='" & Key & "'"
    
    Set Rs = New ADODB.Recordset
    
    Rs.Open mySql, DbtCn, adOpenKeyset, adLockOptimistic
    
        
    If Rs.EOF = True Or Rs.BOF = True Then
        setValue = False
    Else
        Rs(DataRecordName) = rData
        Rs.Update
    
        setValue = True
    End If
        
    Rs.Close
    Set Rs = Nothing
    
    Exit Function
    
    
ERRHANDLER:
    
    setValue = False

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
Public Function getValue(TableName As String, RecordName As String, Key As String, DataRecordName As String) As String
    Dim Ret As String
    Dim Rs As Object
    Dim mySql As String
    
    On Error GoTo ERRHANDLER
    
    mySql = "SELECT * FROM " & TableName
    mySql = mySql & " WHERE " & RecordName & "='" & Key & "'"
    
    Set Rs = New ADODB.Recordset
    
    Rs.Open mySql, DbtCn, adOpenDynamic, adLockReadOnly
    
        
    If Rs.EOF = True Or Rs.BOF = True Then
        getValue = ""
    Else
        getValue = Rs(DataRecordName)
    End If
        
    Rs.Close
    Set Rs = Nothing
    
    Exit Function
    
    
ERRHANDLER:
    
    getValue = "error"


End Function

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
Public Function setDB(DBPath As String) As Boolean
    On Error GoTo ERRHANDLER
    
    'DB連結ｵﾌﾞｼﾞｪｸﾄの設定
    If DbtCn.State = 1 Then
        DbtCn.Close
    End If
        
    DbtCn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & DBPath
    DbtCn.Open
    
    setDB = True
    Exit Function
    
ERRHANDLER:
    setDB = False
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
Public Function addValue(TableName As String, RecordName As String, rData As String) As Boolean
    
    Dim Ret As String
    Dim Rs As Object
    Dim mySql As String
    Dim DbCmd As ADODB.Command

    On Error GoTo ERRHANDLER

    mySql = "INSERT INTO " & TableName & " (" & RecordName & ") VALUES ('" & rData & "');"

	if exeSQL(mySql) = True Then
	    addValue = True
    else
	    addValue = false
	endif

    Exit Function


ERRHANDLER:
    
    addValue = False

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
Public Function CloseDB() As Boolean
    If DbtCn.State = 1 Then DbtCn.Close
End Function

