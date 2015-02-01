'! A VBScript module about Windows Services.
'! Part of VBSbricks.


Option Explicit


'! Returns the status of a provided service name.
'! @param strServiceName the name of the service, string.
'! Not to be confused with the display name, it is the service name.
'! @return the status of a provided service name, "running", "stopped" or "error", string.
'! If the service does not exist, it will return "stopped".
Public Function GetServiceStatus(strServiceName)
    Err.Clear
    Dim strComputer
    strComputer = "."

    Dim result

    Dim bolServiceRunning
    bolServiceRunning = IsServiceRunning(strComputer, strServiceName)
    If Err.Number = 0 Then
        If bolServiceRunning Then
            result = "running"
        Else
            result = "stopped"
        End If
    Else
        ' there is an error
        result = "error"
    End If

    Err.Clear
    GetServiceStatus = result
End Function



'! Returns True if the service is running, False otherwise or it does not exist.
'! @param strComputer name of the computer in the domain, "." for local, string
'! @param strServiceName the name of the service, string.
'! Not to be confused with the display name, it is the service name.
'! @return True if the service is running, False otherwise or it does not exist, boolean
'! @see http://www.wisesoft.co.uk/scripts/vbscript_check_if_service_is_running.aspx
Public Function IsServiceRunning(strComputer, strServiceName)
    Dim objWMIService
    Dim strWMIQuery
    Dim result

    strWMIQuery = "SELECT * FROM Win32_Service WHERE Name = '" & strServiceName & "' AND state='Running'"

    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

    If objWMIService.ExecQuery(strWMIQuery).Count > 0 then
        result = True
    Else
        result = False
    End If
    IsServiceRunning = result
End Function


'! Returns True if the service exists, False otherwise.
'! @param strComputer name of the computer in the domain, "." for local, string
'! @param strServiceName the name of the service, string.
'! Not to be confused with the display name, it is the service name.
'! @return True if the service exists, False otherwise.
'! @see http://www.wisesoft.co.uk/scripts/vbscript_check_if_service_is_running.aspx
Public Function ServiceExists(strComputer, strServiceName)
    Dim objWMIService
    Dim strWMIQuery
    Dim result

    strWMIQuery = "SELECT * FROM Win32_Service WHERE Name = '" & strServiceName & "'"

    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

    If objWMIService.ExecQuery(strWMIQuery).Count > 0 then
        result = True
    Else
        result = False
    End If

    ServiceExists = result
End Function
