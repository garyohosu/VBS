'�P�̌����̐����ԍ������ɒ�R�l�t�@�C���𒊏o����
'�g�p���@�@��R�l��OK.dat(S01�t�H���_)�����̃t�@�C���փh���b�v����

Option Explicit


call main

	dim rd
	Dim args

	logprintln("[START]")

	Set args = WScript.Arguments

	set rd = new RegistData

	rd.ReadSerial "C:\���z�}�쐬ϸ�\RS-335(Polaris)\Type1\�ʎY\�����\DATA101\101.DAT"
	rd.data_selection args(0),"C:\���z�}�쐬ϸ�\RS-335(Polaris)\Type1\��R�l\��R�l�m�F\DATA101\S01.DAT"
	msgbox("�I��")

sub main



end sub


class RegistData
	
	Dim objDictionary
	Dim duplication_list

	sub ReadSerial(okFile)

		dim objFile
		dim objFso
		dim rdLine
		dim serial
		dim SerialFieldNo

		SerialFieldNo = 6


		'�A�z�z��̍쐬
		Set objDictionary = WScript.CreateObject("Scripting.Dictionary")
		Set duplication_list = WScript.CreateObject("Scripting.Dictionary")

		Set objFso = CreateObject("Scripting.FileSystemObject")
		Set objFile = objFso.OpenTextFile(okFile, 1, False)

		If Err.Number > 0 Then
		    WScript.Echo "Open Error"
		Else
		    Do Until objFile.AtEndOfStream
		        rdLine = objFile.ReadLine
				serial = split(rdLine,",")
				logprintln(rdLine)
				objDictionary.add serial(SerialFieldNo),""
				logprintln(serial(SerialFieldNo))
		    Loop
		End If

		objFile.Close
		Set objFile = Nothing
		Set objFso = Nothing
	end sub

	function isShipping(serial)
		'�L�[�̑��݊m�F
		isShipping  = False
		if objDictionary.Exists(serial) then
			isShipping  = True
		end if
	end function

	function duplication_check(serial)
		'�d���`�F�b�N
		duplication_check=false
		if duplication_list.Exists(serial) = False then
			duplication_check = True
			duplication_list.add serial,""
		end if

	end function

	sub data_selection(inFile,outFile)

		dim robjFile
		dim robjFso

		dim wobjFile
		dim wobjFso

		dim serial
		dim SerialFieldNo
		dim rdLine

		SerialFieldNo = 6
		logprintln(inFile & "," & outFile)

		Set robjFso = CreateObject("Scripting.FileSystemObject")
		Set robjFile = robjFso.OpenTextFile(inFile, 1, False)


		Set wobjFso = CreateObject("Scripting.FileSystemObject")
		Set wobjFile = wobjFso.OpenTextFile(outFile, 2, True)

		If Err.Number > 0 Then
		    WScript.Echo "Open Error"
		Else
		    Do Until robjFile.AtEndOfStream
		        rdLine = robjFile.ReadLine
				serial = split(rdLine,",")
				logprintln(rdLine)
				logprintln(serial(SerialFieldNo))
				if isShipping(serial(SerialFieldNo)) = True then
					if duplication_check(serial(SerialFieldNo)) = True then
						wobjFile.WriteLine rdLine
						logprintln("out:" & rdLine)
					else
						logprintln("[�d��]:" & serial(SerialFieldNo))
					end if
				end if
		    Loop
		End If

		robjFile.Close
		Set robjFile = Nothing
		Set robjFso = Nothing

		wobjFile.Close
		Set wobjFile = Nothing
		Set wobjFso = Nothing

	end sub

end class

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
