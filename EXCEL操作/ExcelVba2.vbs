Dim ExcelApp
Dim ExcelBook
Dim FilePath

FilePath = "T:\TEST_PRG\îƒóp\VBS\EXCELëÄçÏ\ExcelVba.xlsm"

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelBook = ExcelApp.Workbooks.Open(FilePath)
ExcelApp.ActiveWorkbook.Worksheets("Sheet1").select

ExcelApp.Visible = True

ExcelApp.Run "Macro1()"

ExcelBook.Close True
ExcelApp.quit

Set ExcelApp = Nothing
Set ExcelBook = Nothing 
