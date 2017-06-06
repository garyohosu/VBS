Option Explicit

Call Main()

Sub Main()
	dim dt
	dim newfilename
	logPrintln("[START]" & date & " " & time)
	dt = inputbox("出荷日 yyyymmdd")
	if len(dt)=8 then
		dim fso
		set fso = createObject("Scripting.FileSystemObject")

		dim folder
		set folder = fso.getFolder(apppath)

		' ファイル一覧
		dim file
		for each file in folder.files
			if left(file.name,4)="data" then
				newfilename = "data" & dt & right(file.name,len(file.name)-12)
				logPrintln("[" & file.name & "]->[" & newfilename & "]")
				rename file,newfilename
			end if
		next 
	end if
	logPrintln("[END]" & date & " " & time)
	msgbox("終了")
end sub

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
	dim SerialFieldNo


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

sub rename(filepath,filename)
	dim sfo, gfile

	logPrintln("[" & filepath & "]->[" & filename & "]")
	
	set sfo = createobject("scripting.filesystemobject")
	set gfile = sfo.getfile(filepath)

	gfile.name = filename

	set sfo = nothing
	set gfile = nothing
end sub
