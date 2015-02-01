'! Testing module for: vbsbricks_XXXMODULENAME.vbs
'! Part of VBSbricks.
'! @todo replace all occurences of: XXXMODULENAME
'! @todo replace all occurences of: XXXMETHODNAME


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


'! Tests vbsbricks_XXXMODULENAME.vbs : XXXMETHODNAME()
Private Sub TestXXXMETHODNAME1()
    Dim routineName
    routineName = "TestXXXMETHODNAME1()"

    Dim strExpected
    Dim strActual

    strExpected = 1
    strActual = 1
    ' strActual = XXXMETHODNAME()
    Call AssertAreEqual(routineName, strExpected, strActual)
End Sub


'! Starting point of unit tests.
Private Sub Main
    Call IncludeFile(PathCombine(GetScriptDirectory(), "vbsbricks_unit.vbs"))
    ' Call IncludeFile(PathCombine(GetScriptDirectory(), "vbsbricks_XXXMODULENAME.vbs"))
    Call TestMethod1()
End Sub


Call Main
