Include "Library.vbs"  ' 外部ファイル「Library.vbs」を取込み

Sub Include(ByVal strFile)
  Dim objFSO , objStream , strDir

  Set objFSO = WScript.CreateObject("Scripting.FileSystemObject") 
  strDir = objFSO.GetFile(WScript.ScriptFullName).ParentFolder 

  Set objStream = objFSO.OpenTextFile(strDir & "\" & strFile, 1)

  ExecuteGlobal objStream.ReadAll() 
  objStream.Close 

  Set objStream = Nothing 
  Set objFSO = Nothing
End Sub