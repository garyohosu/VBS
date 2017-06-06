option explicit



sub chdir(path)
	Dim WshShell

	Set WshShell = WScript.CreateObject("WScript.Shell")
	WshShell.CurrentDirectory = path

	set WshShell = Nothing
end sub
