'! Testing module for: vbsbricks_arrays.vbs
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


'! Tests vbsbricks_arrays.vbs : AddToArray()
Private Sub TestAddToArray1()
    Dim routineName
    routineName = "TestAddToArray1()"

    ReDim arrAny(2)
    arrAny(0) = "zero"
    arrAny(1) = "one"
    arrAny(2) = "two"
    Call AddToArray(arrAny, "three")
    Call AssertAreEqual(routineName, 3, UBound(arrAny)) ' pass
    Call AssertAreEqual(routineName, "zero", arrAny(0)) ' pass
    Call AssertAreEqual(routineName, "three", arrAny(3)) ' pass
End Sub


'! Tests vbsbricks_arrays.vbs : ArrayToString()
Private Sub TestArrayToString1()
    Dim routineName
    routineName = "TestArrayToString1()"

    ReDim arrAny(2)
    arrAny(0) = "zero"
    arrAny(1) = "one"
    arrAny(2) = "two"

    Dim strExpected
    strExpected = "zero,one,two"

    Dim strActual
    strActual = ArrayToString(arrAny, ",")

    Call AssertAreEqual(routineName, strExpected, strActual) ' pass
End Sub


'! Tests vbsbricks_arrays.vbs : PrintArray()
'! Actually, this test contains no equality or unequality check, but it contributes to code coverage.
Private Sub TestPrintArray()
    Dim routineName
    routineName = "TestPrintArray()"

    ReDim arrAny(2)
    arrAny(0) = "zero"
    arrAny(1) = "one"
    arrAny(2) = "two"
    Call PrintArray(arrAny)
End Sub


'! Starting point of unit tests.
Private Sub Main
    Call IncludeFile(PathCombine(GetScriptDirectory(), "vbsbricks_unit.vbs"))
    Call IncludeFile(PathCombine(GetScriptDirectory(), "vbsbricks_arrays.vbs"))
    Call TestAddToArray1()
    Call TestArrayToString1()
    Call TestPrintArray()
End Sub


Call Main
