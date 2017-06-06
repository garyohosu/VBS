Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("C:\temp")

For Each objFile In objFolder.Files
    WScript.Echo objFile.Name
Next
