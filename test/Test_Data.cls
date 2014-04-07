VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Test_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------
' Name     : Tst_Data.cls
' Function :
'-------------------------------------------------------------------
Option Explicit
Implements ITestCase

Private Function ITestCase_TestCaseInstance() As ITestCase
	Dim objResult As New Test_Data
	' Test for cut/paste error - programmer forgot to change type of objResult.
	'Debug.Assert TypeName(objResult) = TypeName(Me)
	Set ITestCase_TestCaseInstance = objResult
End Function

Private Property Get ITestCase_Suite() As Variant
	' Do nothing here
	' Only run tests within RunTest.
End Property

Private Sub ITestCase_RunTest( _
			      ByVal TestNum As Long, _
			      TestName As String, _
			      ExpectErrNum As Long, _
			      EndOfTests As Boolean _
			      )

	Dim objTestNum As New TestNumGenerator
	Dim target As New DataSet
	
	Select Case TestNum
		' Case tstcTestActionSetup
		' Case tstcTestActionTeardown

		Case objTestNum.NextNumber
			target.getData
			AssertEqual "ƒeƒXƒg", target.Title
			AssertEqual 11, target.Number
		Case Else
		EndOfTests = True
	End Select
End Sub