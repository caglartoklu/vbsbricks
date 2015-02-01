'! Testing module for: vbsbricks_files.vbs
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


'! Tests vbsbricks_files_test.vbs : GetScriptDirectory()
'! Note that this method does not test the GetScriptDirectory() in vbsbricks_files.vbs
'! since they both include a function with the same name.
Private Sub TestGetScriptDirectory()
    Dim routineName
    routineName = "TestGetScriptDirectory()"

    Dim strActual
    strActual = GetScriptDirectory()

    Call AssertAreEqual(routineName, ":\", Right(Left(strActual, 3), 2))
End Sub


'! Tests vbsbricks_files.vbs : GetPathSeparator()
Private Sub TestGetPathSeparator()
    Dim routineName
    routineName = "TestGetPathSeparator()"

    Dim strExpected
    strExpected = "\"

    Dim strActual
    strActual = GetPathSeparator()

    Call AssertAreEqual(routineName, strExpected, strActual)
End Sub


'! Tests vbsbricks_files.vbs : PathCombine()
Private Sub TestPathCombine()
    Dim routineName
    routineName = "TestPathCombine()"

    Dim strExpected
    Dim strActual

    strExpected = "x\y"
    strActual = PathCombine("x", "y")
    Call AssertAreEqual(routineName, strExpected, strActual)

    strExpected = "x\y"
    strActual = PathCombine("x\", "\y")
    Call AssertAreEqual(routineName, strExpected, strActual)
End Sub


'! Tests vbsbricks_files.vbs : PathCombine(), FileExists(), DeleteFileIfExists(),
'! WriteToFile(), ReadAllFileInAscii(), ReadAllFileInUtf8().
Private Sub TestWriteReadFile1()
    Dim routineName
    routineName = "TestWriteReadFile1()"

    Dim filePath
    filePath = PathCombine(GetScriptDirectory(), "temp.txt")

    Dim contentToWrite
    contentToWrite = "because the night" & vbCrlf & "belongs to coders."

    Call DeleteFileIfExists(filePath)
    Call AssertAreEqual(routineName, False, FileExists(filePath))

    Call WriteToFile(filePath, contentToWrite)
    Call AssertAreEqual(routineName, True, FileExists(filePath))

    Dim contentToRead
    contentToRead = ReadAllFileInAscii(filePath)
    Call AssertAreEqual(routineName, contentToWrite, contentToRead)

    contentToRead = ""
    contentToRead = ReadAllFileInUtf8(filePath)
    Call AssertAreEqual(routineName, contentToWrite, contentToRead)

    Call DeleteFileIfExists(filePath)
    Call AssertAreEqual(routineName, False, FileExists(filePath))
End Sub


'! Starting point of unit tests.
Private Sub Main
    Call IncludeFile(PathCombine(GetScriptDirectory(), "vbsbricks_unit.vbs"))
    Call IncludeFile(PathCombine(GetScriptDirectory(), "vbsbricks_files.vbs"))
    Call TestGetScriptDirectory()
    Call TestGetPathSeparator()
    Call TestPathCombine()
    Call TestWriteReadFile1()
End Sub


Call Main
