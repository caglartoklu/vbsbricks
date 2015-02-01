'! A simple unit testing module for VBScript.
'! Part of VBSbricks.


Option Explicit


'! Returns the pass message,
'! so that it is the same as for all test subroutines.
'!
'! @return the pass message, string
Public Function GetPassMessage()
    GetPassMessage = "pass ... "
End Function


'! Returns the fail message,
'! so that it is the same as for all test subroutines.
'!
'! @return the fail message, string
Public Function GetFailMessage()
    GetFailMessage = "FAIL ... "
End Function


'! Tests if two values are equal.
'!
'! @param routineName name of the test subroutine, string
'! @param expected expected value of any type, except array
'! @param actual actual value of any type, except array
Public Sub AssertAreEqual(routineName, expected, actual)
    Dim msg
    Dim bolResult
    bolResult = False
    If expected = actual Then
        msg = GetPassMessage() & routineName
    Else
        msg = GetFailMessage() & routineName & " - "
        msg = msg & "  expected : " & expected & vbCrlf
        msg = msg & "  actual   : " & actual & vbCrlf
    End If
    WScript.Echo msg
End Sub


'! Tests if two values are not equal.
'!
'! @param routineName name of the test subroutine, string
'! @param expected expected value of any type, except array
'! @param actual actual value of any type, except array
Public Sub AssertAreNotEqual(routineName, expected, actual)
    Dim msg
    Dim bolResult
    bolResult = False
    If expected = actual Then
        msg = GetFailMessage() & routineName & " - "
        msg = msg & "  expected : " & expected & vbCrlf
        msg = msg & "  actual   : " & actual & vbCrlf
    Else
        msg = GetPassMessage() & routineName
    End If
    WScript.Echo msg
End Sub


'! Tests if the contents of two arrays are equal.
'!
'! @param routineName name of the test subroutine, string
'! @param arrExpected expected array of any type
'! @param arrActual actual array of any type
Public Sub AssertAreArraysEqual(routineName, arrExpected, arrActual)
    Dim bolResult
    Dim msg
    bolResult = False
    If LBound(arrExpected) = LBound(arrActual) Then
        If UBound(arrExpected) = UBound(arrActual) Then
            Dim bolAllEqual
            bolAllEqual = True
            Dim i
            For i = LBound(arrExpected) To UBound(arrExpected)
                If arrExpected(i) <> arrActual(i) Then
                    ' same number of elements but at least one of them does not match
                    msg = GetFailMessage() & routineName & " - "
                    msg = msg & "  expected(" & i & ") : " & arrExpected(i) & vbCrlf
                    msg = msg & "  actual(" & i & ")   : " & arrActual(i) & vbCrlf
                    bolAllEqual = False
                    Exit For
                End If
            Next
            If bolAllEqual = True Then
                ' same number of elements and all elements are equal
                msg = GetPassMessage() & routineName
            End If
        Else
            ' UBound of array do not match
            msg = GetFailMessage() & routineName & " - "
            msg = msg & "  expected UBound : " & UBound(arrExpected) & vbCrlf
            msg = msg & "  actual   UBound : " & UBound(arrActual) & vbCrlf
        End If
    Else
        ' LBound of array do not match
        msg = GetFailMessage() & routineName & " - "
        msg = msg & "  expected LBound : " & LBound(arrExpected) & vbCrlf
        msg = msg & "  actual   LBound : " & LBound(arrActual) & vbCrlf
    End If
    WScript.Echo msg
End Sub
