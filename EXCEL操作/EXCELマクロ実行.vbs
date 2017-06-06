'http://d.hatena.ne.jp/necoyama3/20081030/1225379056

Dim ExcelApp
Dim ExcelBook
Dim FilePath

FilePath = "ExcelVba.xls"

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelBook = ExcelApp.Workbooks.Open(FilePath)
ExcelApp.ActiveWorkbook.Worksheets("マクロ呼び出しのボタンのあるシート名").select

ExcelApp.Visible = True

ExcelApp.Run "シート名.ボタン_Click()"

ExcelBook.Close True
ExcelApp.quit

Set ExcelApp = Nothing
Set ExcelBook = Nothing 