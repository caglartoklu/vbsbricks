'! Functions and subroutines about arrays.
'! Part of VBSbricks.


Option Explicit

'! Adds an element to an array. Note that arr and varElement must be value types, not objects.
'!
'! @param arr array
'! @param varElement the element to be added to the array
Public Sub AddToArray(arr, varElement)
    ' Not for objects, only for value types
    Dim intCurrentMax
    Dim intNewMax
    intCurrentMax = UBound(arr)
    If intCurrentMax = -1 Then
        intNewMax = 0
    Else
        intNewMax = intCurrentMax + 1
    End If
    ReDim Preserve arr(intNewMax)
    arr(intNewMax) = varElement
End Sub


'! Returns the string representation of an array.
'!
'! @param arr array
'! @param strSep separator to be used between array elements
'! @return the string representation of an array.
Public Function ArrayToString(arr, strSep)
    Dim strResult
    Dim varElement
    Dim strSepCurrent
    strSepCurrent = ""
    strResult = strResult & ""
    For Each varElement In arr
        strResult = strResult & strSepCurrent & varElement
        strSepCurrent = strSep
    Next
    ArrayToString = strResult
End Function


'! Prints the array, each element in a separate line
'!
'! @param arr array
Public Sub PrintArray(arr)
    Dim i
    Dim strResult
    Dim varElement
    i = 0
    For Each varElement In arr
        strResult = strResult & "[" & CStr(i) & "]" & varElement & vbCrlf
        i = i + 1
    Next
    WScript.Echo strResult
End Sub
