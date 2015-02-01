'! Testing module for: vbsbricks_strings.vbs
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


'! Tests vbsbricks_strings.vbs : PadLeft()
Private Sub TestPadLeft1()
    Dim strNeedle
    Dim strExpected
    Dim strActual
    Dim routineName
    routineName = "TestPadLeft1()"

    ' a simple padding test
    strNeedle = "aaa"
    strExpected = "--" & strNeedle
    strActual = PadLeft("aaa", 5, "-")
    Call AssertAreEqual(routineName, strExpected, strActual)

    ' same length, so no padding
    strNeedle = "aaa"
    strExpected = strNeedle
    strActual = PadLeft("aaa", 3, "-")
    Call AssertAreEqual(routineName, strExpected, strActual)

    ' the string is already shorter than the provided maximum lenght
    strNeedle = "aaa"
    strExpected = strNeedle
    strActual = PadLeft("aaa", 2, "-")
    Call AssertAreEqual(routineName, strExpected, strActual)
End Sub


'! Tests vbsbricks_strings.vbs : PadRight()
Private Sub TestPadRight1()
    Dim strNeedle
    Dim strExpected
    Dim strActual
    Dim routineName
    routineName = "TestPadRight1()"

    ' a simple padding test
    strNeedle = "aaa"
    strExpected = strNeedle & "--"
    strActual = PadRight("aaa", 5, "-")
    Call AssertAreEqual(routineName, strExpected, strActual)

    ' same length, so no padding
    strNeedle = "aaa"
    strExpected = strNeedle
    strActual = PadRight("aaa", 3, "-")
    Call AssertAreEqual(routineName, strExpected, strActual)

    ' the string is already shorter than the provided maximum lenght
    strNeedle = "aaa"
    strExpected = strNeedle
    strActual = PadRight("aaa", 2, "-")
    Call AssertAreEqual(routineName, strExpected, strActual)
End Sub


'! Tests vbsbricks_strings.vbs : GetProperDate()
Private Sub TestGetProperDate1()
    Dim strExpected
    Dim strActual
    Dim routineName
    routineName = "TestGetProperDate1()"

    Dim dateSeparator
    Dim groupSeparator
    Dim timeSeparator

    dateSeparator = "//"
    groupSeparator = "+"
    timeSeparator = ":::"

    strActual = Len(GetProperDate(dateSeparator, groupSeparator, timeSeparator))
    strExpected = 25 ' 2014//05//19+15:::49:::22
    Call AssertAreEqual(routineName, strExpected, strActual)
End Sub


'! Tests vbsbricks_strings.vbs : GetDateTimeStamp()
Private Sub TestGetDateTimeStamp1()
    Dim strExpected
    Dim strActual
    Dim routineName
    routineName = "TestGetDateTimeStamp1()"
    strActual = Len(GetDateTimeStamp())
    strExpected = 15 ' 20140519_154922
    Call AssertAreEqual(routineName, strExpected, strActual)
End Sub


'! Starting point of unit tests.
Private Sub Main()
    Call IncludeFile(PathCombine(GetScriptDirectory(), "vbsbricks_unit.vbs"))
    Call IncludeFile(PathCombine(GetScriptDirectory(), "vbsbricks_strings.vbs"))
    Call TestPadLeft1()
    Call TestPadRight1()
    Call TestGetDateTimeStamp1()
    Call TestGetProperDate1()
End Sub


Call Main
