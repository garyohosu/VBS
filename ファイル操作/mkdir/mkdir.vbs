Option Explicit

Const MKDIR = "C:\temp"

Dim ObjFso

Set ObjFso=WScript.CreateObject("Scripting.FileSystemObject")

If ObjFso.FolderExists(MKDIR) = False Then
ObjFso.Createfolder(MKDIR)
End If

set ObjFso = Nothing