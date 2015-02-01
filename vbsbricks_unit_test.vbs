'! Includes unit tests for vbsbricks_unit.vbs.
'! These are unit tests to test the unit test module.
'! This set of unit tests are special:
'! Some tests will fail, but it is ok since it is expected to fail.
'! It is a design choice to obtain code coverage so that all the branches
'! in "vbsbricks_unit.vbs" are covered.
'! The developer needs to follow the results manually and decide whether the tests actually pass or fail.
'! Part of VBSbricks.


Option Explicit


'! Returns the directory of the script file.
'! @return the directory of the script file.
Private Function GetScriptDirectory()
    Dim strScriptFullName
    Dim objFso
    Dim objFile
    Dim strDirectoryName

    strScriptFullName = Wscript.ScriptFullName
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFso.GetFile(strScriptFullName)
    strDirectoryName = objFso.GetParentFolderName(objFile)

    GetScriptDirectory = strDirectoryName
End Function


'! Combines two paths by considering the path separator
'! and returns a new combined path.
'! @param strPath1 First path
'! @param strPath2 Second path
'! @return Two paths joined, string
Private Function PathCombine(strPath1, strPath2)
    Dim strResult
    Dim objFso
    Set objFso = CreateObject("Scripting.FileSystemObject")
    strResult = objFso.BuildPath(strPath1, strPath2)
    PathCombine = strResult
End Function


'! Includes a file so that its functions and subroutines can be accessed.
'!
'! @param strFileName file name to be included, string
Private Sub IncludeFile(strFileName)
    Dim objFso
    Set objFso = CreateObject("Scripting.FileSystemObject")
    ExecuteGlobal objFso.OpenTextFile(strFileName).ReadAll()
End Sub


'! Tests the AssertAreEqual() subroutine.
'! The test is expected to pass.
Public Sub TestAssertAreEqualPass()
    Dim routineName
    routineName = "TestAssertAreEqualPass()"
    Call AssertAreEqual(routineName, "one", "one")
End Sub


'! Tests the AssertAreEqual() subroutine.
'! The test is expected to fail.
Public Sub TestAssertAreEqualFail()
    Dim routineName
    routineName = "TestAssertAreEqualFail()"
    Call AssertAreEqual(routineName, "one", "two")
End Sub


'! Tests the AssertAreEqual() subroutine.
'! The test is expected to pass.
Public Sub TestAssertAreNotEqualPass()
    Dim routineName
    routineName = "TestAssertAreNotEqualPass()"
    Call AssertAreNotEqual(routineName, "one", "two")
End Sub


'! Tests the AssertAreEqual() subroutine.
'! The test is expected to fail.
Public Sub TestAssertAreNotEqualFail()
    Dim routineName
    routineName = "TestAssertAreNotEqualFail()"
    Call AssertAreNotEqual(routineName, "one", "one")
End Sub


'! Tests the AssertAreArraysEqual() subroutine.
'! In this test, all the elements are equal, and it is expected to pass.
Public Sub TestAssertAreArraysEqualWithSameArraysPass()
    Dim routineName
    routineName = "TestAssertAreArraysEqualWithSameArraysPass()"

    Dim arrA1(2)
    arrA1(0) = "zero"
    arrA1(1) = "one"
    arrA1(2) = "two"

    Dim arrA2(2)
    arrA2(0) = "zero"
    arrA2(1) = "one"
    arrA2(2) = "two"
    Call AssertAreArraysEqual(routineName, arrA1, arrA2)
End Sub


'! Tests the AssertAreArraysEqual() subroutine.
'! In this test, there are equal number of elements,
'! but one of them is different, so it is expected to fail.
Public Sub TestAssertAreArraysEqualWithDifferentValuesFail()
    Dim routineName
    routineName = "TestAssertAreArraysEqualWithDifferentValuesFail()"

    Dim arrA1(2)
    arrA1(0) = "zero"
    arrA1(1) = "one"
    arrA1(2) = "two"

    Dim arrA2(2)
    arrA2(0) = "0000" ' different
    arrA2(1) = "one"
    arrA2(2) = "two"
    Call AssertAreArraysEqual(routineName, arrA1, arrA2)

    ' repeat the test, but with the other array
    Dim arrB1(2)
    arrB1(0) = "zero"
    arrB1(1) = "one"
    arrB1(2) = "222" ' different

    Dim arrB2(2)
    arrB2(0) = "zero"
    arrB2(1) = "one"
    arrB2(2) = "two"
    Call AssertAreArraysEqual(routineName, arrB1, arrB2)
End Sub


'! Tests the AssertAreArraysEqual() subroutine.
'! In this test, the number of elements does not match.
'! It is expected to fail.
Public Sub TestAssertAreArraysEqualWithDifferentSizeFail()
    Dim routineName
    routineName = "TestAssertAreArraysEqualWithDifferentSizeFail()"

    Dim arrA1(2)
    arrA1(0) = "zero"
    arrA1(1) = "one"
    arrA1(2) = "two"

    Dim arrA2(1) ' different
    arrA2(0) = "zero"
    arrA2(1) = "one"
    Call AssertAreArraysEqual(routineName, arrA1, arrA2)

    ' repeat the test, but with the other array
    Dim arrB1(1) ' different
    arrB1(0) = "zero"
    arrB1(1) = "one"

    Dim arrB2(2)
    arrB2(0) = "zero"
    arrB2(1) = "one"
    arrB2(2) = "two"
    Call AssertAreArraysEqual(routineName, arrB1, arrB2)
End Sub


'! Starting point of unit tests.
Private Sub Main()
    Call IncludeFile(PathCombine(GetScriptDirectory(), "vbsbricks_unit.vbs"))
    Call TestAssertAreEqualPass()
    Call TestAssertAreEqualFail()
    Call TestAssertAreNotEqualPass()
    Call TestAssertAreNotEqualFail()
    Call TestAssertAreArraysEqualWithSameArraysPass()
    Call TestAssertAreArraysEqualWithDifferentValuesFail()
    Call TestAssertAreArraysEqualWithDifferentSizeFail()
End Sub


Call Main
