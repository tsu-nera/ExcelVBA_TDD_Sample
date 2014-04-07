' Input parameter
Dim oParam
Set oParam = WScript.Arguments

' Param Check
Dim file_name, macro_name, excel_path
If oParam.Count <> 0 Then
	file_name = oParam(0)
	excel_path = file_name
End If

' Convert relative path to absolute path
' Dim objFileSys
' Set objFileSys = 
excel_path = _
     CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(excel_path)

' Open Excel
Dim objExcelApp ,objExcelBook
Set objExcelApp = CreateObject("Excel.Application")
Set objExcelBook = objExcelApp.Workbooks.Open(excel_path, , True)

' Execute Macro
Dim reload, runtest

reload = "'" + excel_path + "'!ThisWorkbook.reloadModule"
runtest = "'" + excel_path + "'!ThisWorkbook.runAllTests"

objExcelApp.Run reload
objExcelApp.Run runtest

' TearDown
objExcelBook.Saved = True
objExcelBook.Close False
Set objExcelBook = Nothing
Set objExcelApp = Nothing
