'! Networking functions and subroutines.
'! Part of VBSbricks.


Option Explicit


'! Sends a UTF-8 e-mail using SMTP.
'!
'! @param strSubject Subject of the e-mail.
'! @param strTo recipient list, "myname@example.com" or " My Name <myname@example.com>" are valid.
'! string separated by ","
'! @param strFrom sender, string
'! @param strCC carbon copy recipient list, string separated by ","
'! @param strBody the e-mail body. To send UTF-8 content, write the contents to a file first,
'! then provide it to this parameter using ReadAllFileInUtf8() function.
'! @param arrAttachmentFileNames() The list of files to be attached to this e-mail.
'! @param strSmtpServerAddress The address of the SMTP server, string
'! @param intSmtpServerPort The address of the SMTP port, integer
'!
'! @see ReadAllFileInUtf8() in vbsbricks_files.vb
'! @see http://www.paulsadowski.com/wsh/cdo.htm
Public Sub SendMailUsingSmtp(strSubject, strTo, strFrom, strCC, strBody, arrAttachmentFileNames(), strSmtpServerAddress, intSmtpServerPort)
    Dim objEmail
    Set objEmail = CreateObject("CDO.Message")
    objEmail.Subject = strSubject
    objEmail.To = strTo

    If Len(Trim(strFrom)) > 0 Then
        objEmail.From = strFrom
    End If

    objEmail.CC = strCC
    objEmail.BodyPart.Charset = "UTF-8"
    objEmail.Textbody = strBody


    On Error Resume Next
    ' if the array is not initialized, on error resume next will prevent the raise of the error.
    Dim i
    For i = LBound(arrAttachmentFileNames) To UBound(arrAttachmentFileNames)
        objEmail.AddAttachment arrAttachmentFileNames(i)
    Next
    On Error Goto 0 ' error handling is turned off again

    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSmtpServerAddress
    objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = intSmtpServerPort
    objEmail.Configuration.Fields.Update

    objEmail.Send
End Sub
