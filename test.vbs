Dim objExcelApp ,objExcelBook
Dim macro_path

macro_path = "C:\cygwin\home\TSUNEMICHI\repo\vba-study\sample\test.xlsm"

Set objExcelApp = CreateObject("Excel.Application")
Set objExcelBook = objExcelApp.Workbooks.Open(macro_path, , True)

objExcelApp.Run "'" + macro_path + "'!ThisWorkbook.reloadModule"
objExcelApp.Run "'" + macro_path + "'!ThisWorkbook.runAllTests"

objExcelBook.Saved = True
objExcelBook.Close False
Set objExcelBook = Nothing
Set objExcelApp = Nothing