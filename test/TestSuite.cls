VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------
' Name     : TestSuite.cls
'-------------------------------------------------------------------
Option Explicit

Implements ITestCase

Private Function ITestCase_TestCaseInstance() As ITestCase
	Dim objResult As New TestSuite
	' Test for cut/paste error - programmer forgot to change type of objResult.
	Debug.Assert TypeName(objResult) = TypeName(Me)
	Set ITestCase_TestCaseInstance = objResult
End Function

Private Property Get ITestCase_Suite() As Variant
	ITestCase_Suite = Array( _
				New Test_Data, _
				New Test_Data2 _
				)
End Property

Private Sub ITestCase_RunTest( _
			      ByVal TestNum As Long, _
			      TestName As String, _
			      ExpectErrNum As Long, _
			      EndOfTests As Boolean _
			      )
			      EndOfTests = True
End Sub