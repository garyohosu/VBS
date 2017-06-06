Dim ExcelApp
Dim ExcelBook
Dim FilePath

FilePath = "C:\project\VBS\EXCELëÄçÏ\ExcelVba.xlsm"

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelBook = ExcelApp.Workbooks.Open(FilePath)
ExcelApp.ActiveWorkbook.Worksheets("Sheet1").select

ExcelApp.Visible = True

ExcelApp.Run "Sheet1.Macro1()"

ExcelBook.Close True
ExcelApp.quit

Set ExcelApp = Nothing
Set ExcelBook = Nothing 