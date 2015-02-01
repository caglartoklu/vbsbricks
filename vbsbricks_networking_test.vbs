'! Testing module for: vbsbricks_networking.vbs
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


'! Tests vbsbricks_networking.vbs : SendMailUsingSmtp()
'! This is not an actual test, but a test template instead.
'! The user of this subroutine needs to adapt it to his/her own SNMP server.
'! @todo Adapt the subroutine for local tests.
Private Sub TestSendMailUsingSmtp()
    Dim routineName
    routineName = "TestSendMailUsingSmtp()"

    Dim strSubject
    Dim strTo
    Dim strFrom
    Dim strCC
    Dim strBody
    Dim arrAttachmentFileNames()
    Dim strSmtpServerAddress
    Dim intSmtpServerPort

    strSubject = "TEST subject"
    strTo = "some.name@example.com"
    strFrom = "your.name@example.com"
    strCC = ""
    strBody = "because the night belongs to coders"
    strSmtpServerAddress = "some.server.address"
    intSmtpServerPort = 25

    ' Disabled to prevent accidents.
    ' Tests for this subroutine should be specific and manual.
    ' Call SendMailUsingSmtp(strSubject, strTo, strFrom, strCC, strBody, arrAttachmentFileNames, strSmtpServerAddress, intSmtpServerPort)
End Sub


'! Starting point of unit tests.
Private Sub Main()
    Call IncludeFile(PathCombine(GetScriptDirectory(), "vbsbricks_unit.vbs"))
    Call IncludeFile(PathCombine(GetScriptDirectory(), "vbsbricks_networking.vbs"))
    Call TestSendMailUsingSmtp()
End Sub


Call Main
