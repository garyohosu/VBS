Include "ClsDBTool.vbs"  ' �O���t�@�C�����捞��

Dim m_dbt
Dim mySql

Set m_dbt = new ClsDBTool

call m_dbt.setDB("Z:\RS-387-9001\DB\RS-387_1.mdb")
msgbox(m_dbt.getValue("Data","�H���Ǘ��ԍ�","AM1000604","���i�V���A���ԍ�"))
msgbox(m_dbt.setValue("Data","�H���Ǘ��ԍ�","AM1000604","���i�V���A���ԍ�","A1277777777"))
msgbox(m_dbt.getValue("Data","�H���Ǘ��ԍ�","AM1000604","���i�V���A���ԍ�"))
msgbox(m_dbt.setValue("Data","�H���Ǘ��ԍ�","AM1000604","���i�V���A���ԍ�","A12K0000378"))

mySql = "INSERT INTO Data (�H���Ǘ��ԍ�) VALUES ('AM2000001');"
msgbox(m_dbt.exeSQL(mySql))
msgbox(m_dbt.addValue("Data","�H���Ǘ��ԍ�","AM3000001"))

m_dbt.CloseDB

Sub Include(ByVal strFile)
  Dim objFSO , objStream , strDir
  Set objFSO = WScript.CreateObject("Scripting.FileSystemObject") 
  strDir = objFSO.GetFile(WScript.ScriptFullName).ParentFolder 
  Set objStream = objFSO.OpenTextFile(strDir & "\" & strFile, 1)
  ExecuteGlobal objStream.ReadAll() :  objStream.Close 
  Set objStream = Nothing : Set objFSO = Nothing
End Sub