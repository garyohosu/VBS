Include "ClsDBTool.vbs"  ' 外部ファイルを取込み

Dim m_dbt
Dim mySql

Set m_dbt = new ClsDBTool

call m_dbt.setDB("Z:\RS-387-9001\DB\RS-387_1.mdb")
msgbox(m_dbt.getValue("Data","工程管理番号","AM1000604","製品シリアル番号"))
msgbox(m_dbt.setValue("Data","工程管理番号","AM1000604","製品シリアル番号","A1277777777"))
msgbox(m_dbt.getValue("Data","工程管理番号","AM1000604","製品シリアル番号"))
msgbox(m_dbt.setValue("Data","工程管理番号","AM1000604","製品シリアル番号","A12K0000378"))

mySql = "INSERT INTO Data (工程管理番号) VALUES ('AM2000001');"
msgbox(m_dbt.exeSQL(mySql))
msgbox(m_dbt.addValue("Data","工程管理番号","AM3000001"))

m_dbt.CloseDB

Sub Include(ByVal strFile)
  Dim objFSO , objStream , strDir
  Set objFSO = WScript.CreateObject("Scripting.FileSystemObject") 
  strDir = objFSO.GetFile(WScript.ScriptFullName).ParentFolder 
  Set objStream = objFSO.OpenTextFile(strDir & "\" & strFile, 1)
  ExecuteGlobal objStream.ReadAll() :  objStream.Close 
  Set objStream = Nothing : Set objFSO = Nothing
End Sub