'! A VBScript module about strings.
'! Part of VBSbricks.


Option Explicit


'! Pads the string with pad characters to the left until the specified maximum length
'!
'! @param strNeedle the string to be padded, string
'! @param intMax specified maximum length, integer
'! @param strPadChar the pad character to be used, string
'! @return padded copy of the string, string
Public Function PadLeft(strNeedle, intMax, strPadChar)
    Dim i
    Dim strResult

    ' make sure that strPadChar is exactly 1 byte;
    ' nothing less:
    strPadChar = Trim(strPadChar)
    If Len(strPadChar) = 0 Then
        strPadChar = " "
    End If
    ' nothing more:
    strPadChar = Left(strPadChar, 1)

    Dim strMissing

    Dim intMissing
    intMissing = intMax - Len(strNeedle)

    strResult = strNeedle
    If intMissing > 0 Then
        For i = 1 To intMissing
            strResult = strPadChar & strResult
        Next
    End IF

    PadLeft = strResult
End Function


'! Pads the string with pad characters to the right until the specified maximum length
'!
'! @param strNeedle the string to be padded, string
'! @param intMax specified maximum length, integer
'! @param strPadChar the pad character to be used, string
'! @return padded copy of the string, string
Public Function PadRight(strNeedle, intMax, strPadChar)
    Dim i
    Dim strResult

    ' make sure that strPadChar is exactly 1 byte;
    ' nothing less:
    strPadChar = Trim(strPadChar)
    If Len(strPadChar) = 0 Then
        strPadChar = " "
    End If
    ' nothing more:
    strPadChar = Left(strPadChar, 1)

    Dim strMissing

    Dim intMissing
    intMissing = intMax - Len(strNeedle)

    strResult = strNeedle
    If intMissing > 0 Then
        For i = 1 To intMissing
            strResult = strResult & strPadChar
        Next
    End IF

    PadRight = strResult
End Function


'! Formats the current time and returns it as a string.
'!
'! @param dateSeparator the separator between year and month, month and day, such as "-", string
'! @param groupSeparator the separator between date block and time block, such as "_", string
'! @param timeSeparator the separator between hour and minute, minute and second, such as ":", string
'! @return a string like "2014-05-19_15:49:22", string
Public Function GetProperDate(dateSeparator, groupSeparator, timeSeparator)
    Dim moment
    moment = now

    Dim strResult
    strResult = ""
    strResult = strResult & PadLeft(CStr(Year(moment)), 4, "0")
    strResult = strResult & dateSeparator
    strResult = strResult & PadLeft(CStr(Month(moment)), 2, "0")
    strResult = strResult & dateSeparator
    strResult = strResult & PadLeft(CStr(Day(moment)), 2, "0")
    strResult = strResult & groupSeparator
    strResult = strResult & PadLeft(CStr(Hour(moment)), 2, "0")
    strResult = strResult & timeSeparator
    strResult = strResult & PadLeft(CStr(Minute(moment)), 2, "0")
    strResult = strResult & timeSeparator
    strResult = strResult & PadLeft(CStr(Second(moment)), 2, "0")
    GetProperDate = strResult
End Function


'! Formats the current time and returns it as a string.
'! The result of this function can be used in files and folder names.
'! For this purpose, it uses the following values by default:
'! - dateSeparator = ""
'! - groupSeparator = "_"
'! - timeSeparator = ""
'! This function uses GetProperDate() function.
'! @return a string like "20140519_154922", string
Public Function GetDateTimeStamp()
    Dim dateSeparator
    Dim groupSeparator
    Dim timeSeparator

    dateSeparator = ""
    groupSeparator = "_"
    timeSeparator = ""

    GetDateTimeStamp = GetProperDate(dateSeparator, groupSeparator, timeSeparator)
End Function
