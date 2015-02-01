'! Testing module for: vbsbricks_services.vbs
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


'! Tests vbsbricks_services.vbs : GetServiceStatus()
Private Sub TestGetServiceStatus1()
    Dim serviceName
    Dim serviceStatus
    Dim routineName
    routineName = "TestGetServiceStatus1()"

    ' COM+ Event System is among the default Windows services,
    ' so it is expected to be running.
    serviceName = "EventSystem"
    serviceStatus = GetServiceStatus(serviceName)
    Call AssertAreEqual(routineName, "running", serviceStatus)

    ' "smash it" is not a real service, so it can not be running.
    serviceName = "smash it"
    serviceStatus = GetServiceStatus(serviceName)
    Call AssertAreEqual(routineName, "stopped", serviceStatus)
End Sub


'! Tests vbsbricks_services.vbs : ServiceExists()
Private Sub TestServiceExists()
    Dim serviceName
    Dim result
    Dim routineName
    routineName = "TestServiceExists()"

    ' COM+ Event System is among the default Windows services,
    ' so it is expected to be running.
    serviceName = "EventSystem"
    result = ServiceExists(serviceName)
    Call AssertAreEqual(routineName, True, result)

    ' "smash it" is not a real service, so it does not exist.
    serviceName = "smash it"
    result = ServiceExists(serviceName)
    Call AssertAreEqual(routineName, False, result)
End Sub


'! Starting point of unit tests.
Private Sub Main()
    Call IncludeFile(PathCombine(GetScriptDirectory(), "vbsbricks_unit.vbs"))
    Call IncludeFile(PathCombine(GetScriptDirectory(), "vbsbricks_services.vbs"))
    Call TestGetServiceStatus1()
    Call TestServiceExists()
End Sub


Call Main
