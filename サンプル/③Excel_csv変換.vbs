Option Explicit

Call Main()

Sub Main()

	Dim objExcel
	Dim targetSheetNum
	Dim fullPathFileName

	'Dim FullPath="C:\���z�}�쐬ϸ�\RS-387(Minerva)\-0004(Type4)\-0004(Type4)_UM03\�ʎY"
	Dim FullPath
    FullPath = apppath

    Dim Stage
    Stage = "MAIN�����\DATA101\101_OK DAT�p.xlsx"
    Stage = Stage & ",IP�����\DATA201\201_OK DAT�p.xlsx"
    Stage = Stage & ",�����i����1\DATA301\301_OK DAT�p.xlsx"
    Stage = Stage & ",RF����DOCOMO\DATA401\401_OK DAT�p.xlsx"
    Stage = Stage & ",RF����KDDI\DATA402\402_OK DAT�p.xlsx"

    Dim TargetCSV
    TargetCSV = "103.DAT,202.DAT,303.DAT,403.DAT,404.DAT"

	Dim StageList
	StageList = split(Stage,",")
	Dim TargetCSVList
	TargetCSVList = split(TargetCSV,",")

	Set objExcel  = CreateObject("Excel.Application")

	fullPathFileName = "C:\���z�}�쐬ϸ�\RS-387(Minerva)\-0004(Type4)\-0004(Type4)_UM03\�ʎY\IP�����\DATA201\201_OK DAT�p.xlsx"
	targetSheetNum = 1

	dim I

	for I = 0 to UBound(StageList)
		fullPathFileName = FullPath & "\" & StageList(I)
		ExcelToCsv objExcel, fullPathFileName, TargetCSVList(I),targetSheetNum
	next
	msgbox("�I��")

end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ExcelToCSV
'  - Excel�̋@�\���g����Excel��CSV�ɕϊ�
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub ExcelToCsv(objExcel, fullPathFileName,targetName, targetSheetNum)
    Dim xlSheet
    Dim csvFileName
    Dim FSO, PathName, FileName

    Set FSO = CreateObject("Scripting.FileSystemObject")
    FileName = FSO.GetFileName(fullPathFileName)
    PathName = FSO.GetParentFolderName(fullPathFileName)
    Set FSO = Nothing

    objExcel.DisplayAlerts = False
    objExcel.Workbooks.Open(fullPathFileName)
    Set xlSheet = objExcel.Worksheets(targetSheetNum)
    ' 1�s��(���o���s)�͍폜
    xlSheet.Rows(1).Delete
    'csvFileName = Left(fullPathFileName, Len(fullPathFileName) - 4) + ".csv"
    csvFileName = PathName & "\" &  targetName


    ' CSV�ŕۑ�����ɂ� SaveAs ��2�ڂ̈����� 6 ���w��
    xlSheet.SaveAs csvFileName, 6
    objExcel.Workbooks.Close
End Sub

function apppath
    dim fso
    set fso = createObject("Scripting.FileSystemObject")
    apppath = fso.getParentFolderName(WScript.ScriptFullName)
end function
