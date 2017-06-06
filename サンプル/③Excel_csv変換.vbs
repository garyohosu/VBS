Option Explicit

Call Main()

Sub Main()

	Dim objExcel
	Dim targetSheetNum
	Dim fullPathFileName

	'Dim FullPath="C:\分布図作成ﾏｸﾛ\RS-387(Minerva)\-0004(Type4)\-0004(Type4)_UM03\量産"
	Dim FullPath
    FullPath = apppath

    Dim Stage
    Stage = "MAIN基板検査\DATA101\101_OK DAT用.xlsx"
    Stage = Stage & ",IP基板検査\DATA201\201_OK DAT用.xlsx"
    Stage = Stage & ",完成品検査1\DATA301\301_OK DAT用.xlsx"
    Stage = Stage & ",RF検査DOCOMO\DATA401\401_OK DAT用.xlsx"
    Stage = Stage & ",RF検査KDDI\DATA402\402_OK DAT用.xlsx"

    Dim TargetCSV
    TargetCSV = "103.DAT,202.DAT,303.DAT,403.DAT,404.DAT"

	Dim StageList
	StageList = split(Stage,",")
	Dim TargetCSVList
	TargetCSVList = split(TargetCSV,",")

	Set objExcel  = CreateObject("Excel.Application")

	fullPathFileName = "C:\分布図作成ﾏｸﾛ\RS-387(Minerva)\-0004(Type4)\-0004(Type4)_UM03\量産\IP基板検査\DATA201\201_OK DAT用.xlsx"
	targetSheetNum = 1

	dim I

	for I = 0 to UBound(StageList)
		fullPathFileName = FullPath & "\" & StageList(I)
		ExcelToCsv objExcel, fullPathFileName, TargetCSVList(I),targetSheetNum
	next
	msgbox("終了")

end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ExcelToCSV
'  - Excelの機能を使ってExcelをCSVに変換
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
    ' 1行目(見出し行)は削除
    xlSheet.Rows(1).Delete
    'csvFileName = Left(fullPathFileName, Len(fullPathFileName) - 4) + ".csv"
    csvFileName = PathName & "\" &  targetName


    ' CSVで保存するには SaveAs の2つ目の引数で 6 を指定
    xlSheet.SaveAs csvFileName, 6
    objExcel.Workbooks.Close
End Sub

function apppath
    dim fso
    set fso = createObject("Scripting.FileSystemObject")
    apppath = fso.getParentFolderName(WScript.ScriptFullName)
end function
