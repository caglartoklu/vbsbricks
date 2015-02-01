'! Functions and subroutines about files, folders and IO.
'! Part of VBSbricks.


Option Explicit

'! ForReading = 1 is a constant to provide to FileSystemObject.OpenTextFile() method.
'! @see http://msdn.microsoft.com/en-us/library/314cz14s%28v=vs.84%29.aspx
Public Const ForReading = 1

'! ForWriting = 2 is a constant to provide to FileSystemObject.OpenTextFile() method.
'! @see http://msdn.microsoft.com/en-us/library/314cz14s%28v=vs.84%29.aspx
Public Const ForWriting = 2

'! ForAppending = 8 is a constant to provide to FileSystemObject.OpenTextFile() method.
'! @see http://msdn.microsoft.com/en-us/library/314cz14s%28v=vs.84%29.aspx
Public Const ForAppending = 8


'! Returns the directory of the script file.
'! @return the directory of the script file.
'! @see https://msdn.microsoft.com/en-us/library/cc5ywscw.aspx
Public Function GetScriptDirectory()
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


'! Returns the path separator which is "\" on Microsoft Windows.
'! Even though the return value is fixed, this function is used
'! for unification of the path separators.
'! @return The path separator, string
Public Function GetPathSeparator()
    Dim strTemp
    strTemp = PathCombine("x", "y")

    Dim strSep
    strSep = Mid(strTemp, 2, 1)

    ' GetPathSeparator = "\"
    GetPathSeparator = strSep
End Function


'! Combines two paths by considering the path separator
'! and returns a new combined path.
'! @param strPath1 First path
'! @param strPath2 Second path
'! @return Two paths joined, string
'! @see https://msdn.microsoft.com/en-us/library/z0z2z1zt.aspx
Public Function PathCombine(strPath1, strPath2)
    Dim strResult
    Dim objFso
    Set objFso = CreateObject("Scripting.FileSystemObject")
    strResult = objFso.BuildPath(strPath1, strPath2)
    PathCombine = strResult
End Function


'! Writes a content to a text file.
'!
'! @param strFileName the name of the file to write, string
'! @param strText the content to be written, string
'! @see https://msdn.microsoft.com/en-us/library/5t9b5c0c.aspx
Public Sub WriteToFile(strFileName, strText)
    Dim objFso
    Dim objFile
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFso.CreateTextFile(strFileName, True)
    objFile.Write strText
    objFile.Close
End Sub


'! Reads a text file and returns its content in ASCII encoding.
'!
'! @param strFileName the name of the file to be read, string
'! @return The content of the file in ASCII encoding, string
'! @see ReadAllFileInUtf8()
'! @see https://msdn.microsoft.com/en-us/library/314cz14s.aspx
Public Function ReadAllFileInAscii(strFileName)
    Dim strFileContent

    Dim objFS
    Set objFS = CreateObject("Scripting.FileSystemObject")
    Dim objFile
    Set objFile = objFS.OpenTextFile(strFileName, ForReading)
    strFileContent = objFile.ReadAll
    objFile.Close

    ReadAllFileInAscii = strFileContent
End Function


'! Reads a text file and returns its content in UTF-8 encoding.
'!
'! @param strFileName the name of the file to be read, string
'! @return The content of the file in UTF-8 encoding, string
'! @see ReadAllFileInAscii()
'! @see http://stackoverflow.com/a/7235192
'! @see https://msdn.microsoft.com/en-us/library/windows/desktop/ms675032.aspx
Public Function ReadAllFileInUtf8(strFileName)
    Dim strFileContent

    Dim adoStream
    Set adoStream = CreateObject("Adodb.Stream")
    adoStream.Open
    adoStream.Charset = "UTF-8"
    adoStream.LoadFromFile strFileName
    strFileContent = adoStream.ReadText(-1)
    adoStream.Close
    Set adoStream = Nothing

    ReadAllFileInUtf8 = strFileContent
End Function


'! Returns True if the file exists, False otherwise.
'!
'! @param fileNameToTest name of the file to check, string
'! @return  True if the file exists, False otherwise.
'! @see https://msdn.microsoft.com/en-us/library/x23stk5t.aspx
Public Function FileExists(fileNameToTest)
    Dim result
    result = False
    Dim objFso
    Set objFso = CreateObject("Scripting.FileSystemObject")
    If objFso.FileExists(fileNameToTest) Then
        result = True
    End If
    FileExists = result
End Function


'! Deletes a file if it exists.
'!
'! @param fileName name of the file to be deleted, string
'! @see FileExists()
'! @see https://msdn.microsoft.com/en-us/library/thx0f315.aspx
Public Sub DeleteFileIfExists(fileName)
    If FileExists(fileName) Then
        Dim objFso
        Set objFso = CreateObject("Scripting.FileSystemObject")
        objFso.DeleteFile(fileName)
    End If
End Sub
