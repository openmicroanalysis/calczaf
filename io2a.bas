Attribute VB_Name = "CodeIO2"
' (c) Copyright 1995-2021 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

' Wav file play constants
'Private Const SND_SYNC& = &H0
Private Const SND_ASYNC& = &H1
'Private Const SND_NODEFAULT& = &H2
'Private Const SND_LOOP& = &H8
'Private Const SND_NOSTOP& = &H10

Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Sub IOMsgBox(errstring As String, msgboxflags As Integer, procstring As String)
'  Handles all real time errors

ierror = False
On Error GoTo IOMsgBoxError

Dim tmsg As String

' If automation is running and e-mail notification flag is set, send e-mail error message
If EmailNotificationOfErrorsFlag And AcquisitionOnAutomate Then
Call IOSendEMail(Int(0), errstring$, procstring$)
If ierror Then Exit Sub
End If

' Write error to log file
Call IOWriteError(errstring$, procstring$)
If ierror Then Exit Sub

' Write error to log window
tmsg$ = "ERROR in " & procstring$ & " : " & errstring$
Call IOWriteLogRichText(tmsg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
If ierror Then Exit Sub

' Normal error handling
MsgBox errstring$, msgboxflags%, procstring$
'FormMSGBOXDOEVENTS3.Caption = procstring$
'FormMSGBOXDOEVENTS3.Label1.Caption = errstring$
'FormMSGBOXDOEVENTS3.Show vbModal

ierror = True
Exit Sub

' Errors
IOMsgBoxError:
MsgBox Error$, vbOKOnly + vbCritical, "IOMsgBox"
ierror = True
Exit Sub

End Sub

Sub IOMsgBox2(errstring As String, msgboxflags As Integer, procstring As String)
'  Handles all real time errors silently (no MsgBox notification and returns no error)

ierror = False
On Error GoTo IOMsgBox2Error

Dim tmsg As String

' If automation is running and e-mail notification flag is set, send e-mail error message
If EmailNotificationOfErrorsFlag And AcquisitionOnAutomate Then
Call IOSendEMail(Int(0), errstring$, procstring$)
If ierror Then Exit Sub
End If

' Write error to log file
Call IOWriteError(errstring$, procstring$)
If ierror Then Exit Sub

' Write error to log window
tmsg$ = "ERROR in " & procstring$ & " : " & errstring$
Call IOWriteLogRichText(tmsg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
If ierror Then Exit Sub

Exit Sub

' Errors
IOMsgBox2Error:
MsgBox Error$, vbOKOnly + vbCritical, "IOMsgBox2"
ierror = True
Exit Sub

End Sub

Sub IOWriteLogError(errstring As String, msgboxflags As Integer, procstring As String)
'  Just writes the error to the log window (no MsgBox)

ierror = False
On Error GoTo IOWriteLogErrorError

Dim tmsg As String

' Write error to log window
tmsg$ = "ERROR in " & procstring$ & " : " & errstring$
Call IOWriteLogRichText(tmsg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
If ierror Then Exit Sub

Exit Sub

' Errors
IOWriteLogErrorError:
MsgBox Error$, vbOKOnly + vbCritical, "IOWriteLogError"
ierror = True
Exit Sub

End Sub

Sub IOSendEMail(mode As Integer, errstring As String, procstring As String)
' Send e-mail message to specified address (in INI file)
' mode = 0 error message
' mode = 1 normal message

ierror = False
On Error GoTo IOSendEMailError

Dim itest As Long

Static tmsg As String

' Check for valid addresses
If Trim$(SMTPServerAddress$) = vbNullString Then GoTo IOSendEMailBadSMTPServerAddress
If Trim$(SMTPAddressFrom$) = vbNullString Then GoTo IOSendEMailBadSMTPAddressFrom
If Trim$(SMTPAddressTo$) = vbNullString Then GoTo IOSendEMailBadSMTPAddressTo
If Trim$(SMTPUserName$) = vbNullString Then GoTo IOSendEMailBadSMTPUserName

If InStr(SMTPAddressFrom$, "@") = 0 Then GoTo IOSendEMailBadSMTPAddressFrom
If InStr(SMTPAddressTo$, "@") = 0 Then GoTo IOSendEMailBadSMTPAddressTo
  
' If first time, ask user for password and check again
If Trim$(SMTPUserPassword$) = vbNullString Then
Call PasswordLoad(SMTPUserPassword$, "Enter SMTP Password", "Enter the SMTP password for the specified username (" & SMTPUserName$ & "). This password will be retained until the probe run is closed.")
If ierror Then Exit Sub
End If

' Check password again
If Trim$(SMTPUserPassword$) = vbNullString Then GoTo IOSendEMailBadSMTPUserPassword
  
' Load SMTP control
FormMAIN.SmtpClient1.HostName = Trim$(SMTPServerAddress$)
FormMAIN.SmtpClient1.RemotePort = smtpPortSecure    ' smtpPortStandard or smtpPortSecure
FormMAIN.SmtpClient1.Timeout = 20                   ' assume 20 seconds
FormMAIN.SmtpClient1.UserName = Trim$(SMTPUserName$)
FormMAIN.SmtpClient1.password = Trim$(SMTPUserPassword$)
FormMAIN.SmtpClient1.Options = smtpOptionNone
FormMAIN.SmtpClient1.Extended = True
    
If DebugMode Then
FormMAIN.SmtpClient1.TraceFile = ApplicationCommonAppData$ & "SmtpTest.log"
FormMAIN.SmtpClient1.TraceFlags = 4
FormMAIN.SmtpClient1.Trace = True
Else
FormMAIN.SmtpClient1.Trace = False
End If
    
' Set securce connection (explicit SSL/TLS)
FormMAIN.SmtpClient1.Secure = True
FormMAIN.SmtpClient1.Options = &H2000 ' Implicit SSL/TLS on port 465
        
' Make SMTP connection
itest& = FormMAIN.SmtpClient1.Connect()
If itest& > 0 Then GoTo IOSendEMailBadConnection
    
itest& = FormMAIN.SmtpClient1.Authenticate()
If itest& > 0 Then GoTo IOSendEMailBadAuthentication

' Build message
FormMAIN.MailMessage1.Reset
FormMAIN.MailMessage1.From = Trim$(SMTPAddressFrom$)
FormMAIN.MailMessage1.To = Trim$(SMTPAddressTo$)
FormMAIN.MailMessage1.date = FormMAIN.SmtpClient1.CurrentDate

If mode% = 0 Then
FormMAIN.MailMessage1.Subject = "Automation error from Probe for EPMA, at " & Now
Else
FormMAIN.MailMessage1.Subject = "Automation message from Probe for EPMA, at " & Now
End If

' Construct message
If mode% = 0 Then
tmsg$ = "An error occured during a Probe for EPMA automation procedure!" & vbCrLf & vbCrLf
Else
tmsg$ = "A message occured during a Probe for EPMA automation procedure!" & vbCrLf & vbCrLf
End If
tmsg$ = tmsg$ & "Time: " & Now & vbCrLf
tmsg$ = tmsg$ & "File: " & ProbeDataFile$ & vbCrLf
tmsg$ = tmsg$ & "Version: " & Str$(DataFileVersionNumber!) & vbCrLf
tmsg$ = tmsg$ & "File Type: " & MDBFileType$ & vbCrLf
tmsg$ = tmsg$ & "File Title: " & MDBFileTitle$ & vbCrLf
tmsg$ = tmsg$ & "User: " & MDBUserName$ & vbCrLf
tmsg$ = tmsg$ & "File Description : " & MDBFileDescription$ & vbCrLf & vbCrLf

' Send additional info
If mode% = 0 Then
tmsg$ = tmsg$ & "Procedure : " & procstring$ & vbCrLf
tmsg$ = tmsg$ & "Error : " & errstring$ & vbCrLf
Else
If NumberofSamples% > 0 Then
tmsg$ = tmsg$ & "Sample : " & SampleNams$(NumberofSamples%) & vbCrLf
tmsg$ = tmsg$ & "Line : " & NumberofLines& & vbCrLf
End If
tmsg$ = tmsg$ & "Message : " & errstring$ & vbCrLf
End If
FormMAIN.MailMessage1.Text = tmsg$
    
' Send message
itest& = FormMAIN.SmtpClient1.SendMessage(FormMAIN.MailMessage1.From, FormMAIN.MailMessage1.To, FormMAIN.MailMessage1.Message)
If itest& > 0 Then GoTo IOSendEMailBadSendMessage
    
FormMAIN.SmtpClient1.Disconnect
Exit Sub

' Errors
IOSendEMailError:
MsgBox Error$, vbOKOnly + vbCritical, "IOSendEMail"
ierror = True
Exit Sub

IOSendEMailBadSMTPServerAddress:
msg$ = "Not a valid SMTP Server Address. Make sure the keyword SMTPServerAddress is properly specified in the [general] section of file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "IOSendEMail"
ierror = True
Exit Sub

IOSendEMailBadSMTPAddressFrom:
msg$ = "Not a valid SMTP Address From. Make sure the keyword SMTPAddressFrom is properly specified in the [general] section of file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "IOSendEMail"
ierror = True
Exit Sub

IOSendEMailBadSMTPAddressTo:
msg$ = "Not a valid SMTP Address To. Make sure the keyword SMTPAddressTo is properly specified in the [general] section of file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "IOSendEMail"
ierror = True
Exit Sub

IOSendEMailBadSMTPUserName:
msg$ = "Not a valid user name. Make sure the keyword SMTPUserName is properly specified in the [general] section of file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "IOSendEMail"
ierror = True
Exit Sub

IOSendEMailBadSMTPUserPassword:
msg$ = "Not a valid user password"
MsgBox msg$, vbOKOnly + vbExclamation, "IOSendEMail"
ierror = True
Exit Sub

IOSendEMailBadConnection:
msg$ = "Could not establish an SMTP connection to " & SMTPServerAddress$ & ", Error: " & FormMAIN.SmtpClient1.LastErrorString
MsgBox msg$, vbOKOnly + vbExclamation, "IOSendEMail"
ierror = True
Exit Sub

IOSendEMailBadAuthentication:
msg$ = "Bad authentication for SMTP username " & SMTPUserName$ & ", Error: " & FormMAIN.SmtpClient1.LastErrorString
MsgBox msg$, vbOKOnly + vbExclamation, "IOSendEMail"
FormMAIN.SmtpClient1.Disconnect
ierror = True
Exit Sub

IOSendEMailBadSendMessage:
msg$ = "Bad message send for SMTP username " & SMTPUserName$ & ", Error: " & FormMAIN.SmtpClient1.LastErrorString
MsgBox msg$, vbOKOnly + vbExclamation, "IOSendEMail"
FormMAIN.SmtpClient1.Disconnect
ierror = True
Exit Sub
        
End Sub

Sub IOPlayWavFile(wavfile As String)
' Plays the wave file

ierror = False
On Error GoTo IOPlayWavFileError

Dim wFlags As Long, istatus As Long

' Just exit if file is blank
If Trim$(wavfile$) = vbNullString Then Exit Sub

' Play asychronously
wFlags& = SND_ASYNC

' Play wave file
istatus& = sndPlaySound(wavfile$, wFlags&)
If istatus& = 0 Then GoTo IOPlayWavFileNotPlayed

Exit Sub

' Errors
IOPlayWavFileError:
MsgBox Error$, vbOKOnly + vbCritical, "IOPlayWavFile"
ierror = True
Exit Sub

IOPlayWavFileNotPlayed:
msg$ = "The play sound function was unable to play the .WAV file " & wavfile$
MsgBox msg$, vbOKOnly + vbExclamation, "IOPlayWavFile"
ierror = True
Exit Sub

End Sub

Sub IOWriteError(errstring As String, procstring As String)
' Write to the error log

ierror = False
On Error GoTo IOWriteErrorError

Dim astring As String

' Open file and write error
Open ProbeErrorLogFile$ For Append As #ProbeErrorLogFileNumber%
astring$ = Now & ", Program: " & app.EXEName & ", Error: " & errstring$ & ", Procedure: " & procstring$
Print #ProbeErrorLogFileNumber%, astring$
Close #ProbeErrorLogFileNumber%

Exit Sub

' Errors
IOWriteErrorError:
MsgBox Error$, vbOKOnly + vbCritical, "IOWriteError"
Close #ProbeErrorLogFileNumber%
ierror = True
Exit Sub

End Sub

Sub IOWriteText(txtstring As String)
' Write to the text log

ierror = False
On Error GoTo IOWriteTextError

Dim astring As String

' Open file and write error
Open ProbeTextLogFile$ For Append As #ProbeTextLogFileNumber%
astring$ = Now & ", " & txtstring$
Print #ProbeTextLogFileNumber%, astring$
Close #ProbeTextLogFileNumber%

Exit Sub

' Errors
IOWriteTextError:
MsgBox Error$, vbOKOnly + vbCritical, "IOWriteText"
Close #ProbeTextLogFileNumber%
ierror = True
Exit Sub

End Sub

Sub IOCleanWindowPositionFile()
' Deletes the current user from the WINDOW.INI file

ierror = False
On Error GoTo IOCleanWindowPositionFileError

Dim deleted_lines As Long

' Copy the current WINDOW.INI file to temp file without any lines containing the current user
deleted_lines& = MiscDeleteLines(WindowINIFile$, ApplicationCommonAppData$ & "TEMP.INI", MDBUserName$, "[")
If ierror Then Exit Sub

' Delete the original file
Kill WindowINIFile$

' Copy temp file to new WINDOW.INI
FileCopy ApplicationCommonAppData$ & "TEMP.INI", WindowINIFile$

msg$ = "All window position references to current user (" & MDBUserName$ & ") deleted from " & WindowINIFile$
MsgBox msg$, vbOKOnly + vbInformation, "IOCleamWindowPositionFile"
Exit Sub

' Errors
IOCleanWindowPositionFileError:
MsgBox Error$, vbOKOnly + vbCritical, "IOCleanWindowPositionFile"
ierror = True
Exit Sub

End Sub

