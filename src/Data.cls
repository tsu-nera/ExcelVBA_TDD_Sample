VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Title As String
Public Number As Integer

Sub getData()
	Title = ThisWorkbook.WorkSheets("Sheet2").Cells(1, 1)
	Number = ThisWorkbook.WorkSheets("Sheet2").Cells(1, 2)
End Sub
