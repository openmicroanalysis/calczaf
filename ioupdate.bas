Attribute VB_Name = "CodeIOUpdate"
' (c) Copyright 1995-2024 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Dim DownloadMode As Integer
Dim tLocalFile As String, tRemoteFile As String
Dim tBackupFile As String

Dim tLocalFileDate As Variant, tRemoteFileDate As Variant

Sub IOUpdateClose(Cancel As Integer)
' Check if download is running and close dialog

ierror = False
On Error GoTo IOUpdateCloseError

Dim response As Integer

' Check if download is running (warn user)
If DownloadMode% = 1 Or DownloadMode% = 3 Then
If FormUPDATE.FtpClient1.Connected Then
msg$ = "The download is not complete, are you sure you want to cancel the update download?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton2, "IOUpdateClose")

If response% = vbYes Then
FormUPDATE.FtpClient1.Cancel
If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
FormUPDATE.Caption = "Update CalcZAF [Please wait for termination...]"
Else
FormUPDATE.Caption = "Update Penepma [Please wait for termination...]"
End If
End If

If UCase$(app.EXEName) = UCase$("Probewin") Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
FormUPDATE.Caption = "Update Probe for EPMA [Please wait for termination...]"
Else
FormUPDATE.Caption = "Update Penepma [Please wait for termination...]"
End If
End If

Call MiscDelay(CDbl(2#), Now)
FormUPDATE.FtpClient1.Disconnect
DoEvents
Exit Sub
Else
Cancel = True
End If
End If

Else
If FormUPDATE.HttpClient1.Connected Then
msg$ = "The download is not complete, are you sure you want to cancel the update download?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton2, "IOUpdateClose")
If response% = vbYes Then
icancel = True
FormUPDATE.HttpClient1.Cancel

If UCase$(app.EXEName) = UCase$("CalcZAF") Then
FormUPDATE.Caption = "Update CalcZAF [Please wait for termination...]"
End If

If UCase$(app.EXEName) = UCase$("Probewin") Then
FormUPDATE.Caption = "Update Probe for EPMA [Please wait for termination...]"
End If

Call MiscDelay(CDbl(2#), Now)
FormUPDATE.HttpClient1.Disconnect
DoEvents
Exit Sub
Else
Cancel = True
End If
End If

End If

Exit Sub

' Errors
IOUpdateCloseError:
MsgBox Error$, vbOKOnly + vbCritical, "IOUpdateClose"
ierror = True
Exit Sub

End Sub

Sub IOUpdateGetUpdate(mode As Integer)
' Get the latest update for Probe for EPMA
'   mode = 1 use FTP (Explicit FTP over TLS)
'   mode = 2 use HTTPS
'   mode = 3 use alternative FTP

ierror = False
On Error GoTo IOUpdateGetUpdateError

' Check for valid version
If ProgramVersionNumber! < 8.11 Then GoTo IOUpdateGetUpdateOldVersion

' Save download mode
DownloadMode% = mode%

' Check for trace mode (only for PFE because CalcZAF runs in debug mode by default)
If UCase$(app.EXEName) = UCase$("Probewin") And DebugMode Then
FormUPDATE.FtpClient1.TraceFile = ApplicationCommonAppData$ & "FtpTest.log"
FormUPDATE.FtpClient1.TraceFlags = 4
FormUPDATE.FtpClient1.Trace = True
FormUPDATE.HttpClient1.TraceFile = ApplicationCommonAppData$ & "HttpTest.log"
FormUPDATE.HttpClient1.TraceFlags = 4
FormUPDATE.HttpClient1.Trace = True
Else
FormUPDATE.FtpClient1.Trace = False
FormUPDATE.HttpClient1.Trace = False
End If

' Load filenames for CalcZAF update
If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
tLocalFile$ = ApplicationCommonAppData$ & "CALCZAF.MSI"
tBackupFile$ = ApplicationCommonAppData$ & "CALCZAF_Backup.MSI"

' Download Penepma12.ZIP
Else
tLocalFile$ = ApplicationCommonAppData$ & "PENEPMA12.ZIP"
tBackupFile$ = ApplicationCommonAppData$ & "PENEPMA12_Backup.ZIP"
End If
End If

' Load filenames for Probewin update
If UCase$(app.EXEName) = UCase$("Probewin") Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
tLocalFile$ = ApplicationCommonAppData$ & "ProbeForEPMA.MSI"
tBackupFile$ = ApplicationCommonAppData$ & "ProbeForEPMA_Backup.MSI"

' Download Penepma12.ZIP
Else
tLocalFile$ = ApplicationCommonAppData$ & "PENEPMA12.ZIP"
tBackupFile$ = ApplicationCommonAppData$ & "PENEPMA12_Backup.ZIP"
End If
End If

' Check date on existing download ZIP
If Dir$(tLocalFile$) <> vbNullString Then
tLocalFileDate = FileDateTime(tLocalFile$)
Else
tLocalFileDate = "01/20/1956 8:26 AM"   ' before I was born would be a safe assumption!
End If

' Get date/time using FTP
If mode% = 1 Or mode% = 3 Then
If mode% = 3 And UCase$(app.EXEName) = UCase$("CalcZAF") Then GoTo IOUpdateGetUpdateNotAvailable    ' mode = 3 currently uses SFTP not SSH, so only use normal FTP for PFE updates
Call IOUpdateGetUpdateFTP(Int(0))
If ierror Then Exit Sub
End If

' Get date/time using HTTP
If mode% = 2 Then
Call IOUpdateGetUpdateHTTP(Int(0))
If ierror Then Exit Sub
End If

' See if update is necessary
Call IOWriteLog(vbCrLf & "IOUpdateGetUpdate: date/time of last update (" & tLocalFileDate & "), date/time of current update (" & tRemoteFileDate & ")...")
DoEvents
If CDate(tLocalFileDate) > CDate(tRemoteFileDate) Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
msg$ = "The current version of this program (" & ProgramVersionString$ & ") is already up to date. To force an update download, please use the Delete Update button first and try again."
Else
msg$ = "The current PENEPMA12.ZIP file is already up to date. To force an update download, please use the Delete Update button first and try again."
End If
MsgBox msg$, vbOKOnly + vbExclamation, "IOUpdateGetUpdate"
ierror = True
Exit Sub
End If

' Check if file exists already and if so backup and then delete
If Dir$(tLocalFile$) <> vbNullString Then
If Dir$(tBackupFile$) <> vbNullString Then
Call IOWriteLog("IOUpdateGetUpdate: Removing previous update backup (" & tBackupFile$ & ")...")
DoEvents
Kill tBackupFile$
End If
Call IOWriteLog("IOUpdateGetUpdate: Copying previous update (" & tLocalFile$ & ") to update backup (" & tBackupFile$ & ")...")
DoEvents
FileCopy tLocalFile$, tBackupFile$
Call IOWriteLog("IOUpdateGetUpdate: Removing previous update (" & tLocalFile$ & ")...")
DoEvents
Kill tLocalFile$
DoEvents
End If

' Get using FTP
If mode% = 1 Or mode% = 3 Then
Screen.MousePointer = vbHourglass
Call IOWriteLog("IOUpdateGetUpdate: Downloading update using secure FTP (" & tLocalFile$ & ")...")
DoEvents
If RealTimeMode% Then Call IOAutomationPause(Int(1))
Call IOUpdateGetUpdateFTP(Int(1))
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End If

' Get using HTTPS
If mode% = 2 Then
Screen.MousePointer = vbHourglass
Call IOWriteLog("IOUpdateGetUpdate: Downloading update using secure HTTPS (" & tLocalFile$ & ")...")
DoEvents
If RealTimeMode% Then Call IOAutomationPause(Int(1))
Call IOUpdateGetUpdateHTTP(Int(1))
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End If

Call IOWriteLog("IOUpdateGetUpdate: Download of file (" & tLocalFile$ & ") is complete")

' Notify user
If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
FormUPDATE.Caption = "Update CalcZAF [download complete]"
Else
FormUPDATE.Caption = "Update Penepma [download complete]"
End If
End If

If UCase$(app.EXEName) = UCase$("Probewin") Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
FormUPDATE.Caption = "Update Probe for EPMA [download complete]"
Else
FormUPDATE.Caption = "Update Penepma [download complete]"
End If
End If

DoEvents
If RealTimeMode% Then Call IOAutomationPause(Int(2))
Call IOUpdateComplete
If ierror Then Exit Sub

Unload FormUPDATE
Exit Sub

' Errors
IOUpdateGetUpdateError:
MsgBox Error$, vbOKOnly + vbCritical, "IOUpdateGetUpdate"
ierror = True
Exit Sub

IOUpdateGetUpdateOldVersion:
msg$ = "The current version of this program (" & ProgramVersionString$ & ") is too old to update automatically. The program will have to be re-installed. Please contact Probe Software for more information."
MsgBox msg$, vbOKOnly + vbExclamation, "IOUpdateGetUpdate"
ierror = True
Exit Sub

IOUpdateGetUpdateNotAvailable:
msg$ = "This download site is not currently available for CalcZAF, please try a different download option or download using your browser at https://www.probesoftware.com/resources/."
MsgBox msg$, vbOKOnly + vbExclamation, "IOUpdateGetUpdate"
ierror = True
Exit Sub

End Sub

Sub IOUpdateGetUpdateFTP(mode As Integer)
' Get the update via ftp
'  mode = 0 get download file date/time only
'  mode = 1 download file

ierror = False
On Error GoTo IOUpdateGetUpdateFTPError

Dim tHostname As String, tusername As String, tpassword As String
Dim nError As Long

' Download from whitewater
If DownloadMode% = 1 Then
tHostname$ = "whitewater.uoregon.edu"
tusername$ = "micro"
tpassword$ = "analysis"

' Download CalcZAF
If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
tRemoteFile$ = "Probe for EPMA\V11\CALCZAF.MSI"

' Download penepma12.zip
Else
tRemoteFile$ = "Probe for EPMA\Penepma12\PENEPMA12.ZIP"
End If
End If

' Download Probewin
If UCase$(app.EXEName) = UCase$("Probewin") Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
tRemoteFile$ = "Probe for EPMA\V11\ProbeForEPMA.MSI"

' Download penepma12.zip
Else
tRemoteFile$ = "Probe for EPMA\Penepma12\PENEPMA12.ZIP"
End If
End If

' Download from probe software ftp sub folder
ElseIf DownloadMode% = 3 Then
tHostname$ = "205.178.145.65"
tusername$ = "micro%003a750"
tpassword$ = "4rfvVGY&"

' Download CalcZAF (disabled for now from IOGetUpdate)
If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
tRemoteFile$ = "V11/CalcZAF.msi"

' Download penepma12.zip (disabled for now from IOGetUpdate)
Else
tRemoteFile$ = "PENEPMA12.ZIP"
End If
End If

' Download Probewin
If UCase$(app.EXEName) = UCase$("Probewin") Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
tRemoteFile$ = "V11/ProbeForEPMA.msi"

' Download penepma12.zip
Else
tRemoteFile$ = "PENEPMA12.ZIP"
End If
End If
End If

' Load username and password
FormUPDATE.FtpClient1.UserName = tusername$
FormUPDATE.FtpClient1.password = tpassword$

' Use secure connection (Network Solutions only supports SSH, which is not supported by v. 4.5 of Catalyst- need to update Catalyst OCX component to v. 8 for SSH support)
FormUPDATE.FtpClient1.AutoResolve = False
FormUPDATE.FtpClient1.Blocking = True
FormUPDATE.FtpClient1.Secure = True
If DownloadMode% = 3 Then FormUPDATE.FtpClient1.Secure = False       ' DownloadMode% = 3 is not secure FTP (use for PFE only)

' Set the local file path
If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
tLocalFile$ = ApplicationCommonAppData$ & "CalcZAF.msi"

' Download penepma12.zip
Else
tLocalFile$ = ApplicationCommonAppData$ & "PENEPMA12.ZIP"
End If
End If

If UCase$(app.EXEName) = UCase$("Probewin") Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
tLocalFile$ = ApplicationCommonAppData$ & "ProbeForEPMA.msi"

' Download penepma12.zip
Else
tLocalFile$ = ApplicationCommonAppData$ & "PENEPMA12.ZIP"
End If
End If

' Establish a connection to the server and display any errors to the user
nError& = FormUPDATE.FtpClient1.Connect(tHostname$)
If nError& > 0 Then GoTo IOUpdateGetUpdateFTPBadConnect

' Check date/time of remote file
nError& = FormUPDATE.FtpClient1.Localize = True
nError& = FormUPDATE.FtpClient1.GetFileTime(tRemoteFile$, tRemoteFileDate)
If nError& > 0 Then GoTo IOUpdateGetUpdateFTPBadGetFileTime

' If just getting remote file date then exit sub
If mode% = 0 Then
FormUPDATE.FtpClient1.Disconnect
Exit Sub
End If

' Download the file to the local system
nError& = FormUPDATE.FtpClient1.Localize = True
nError& = FormUPDATE.FtpClient1.GetFile(tLocalFile$, tRemoteFile$)
If nError& > 0 Then GoTo IOUpdateGetUpdateFTPBadGetFile

' Disconnect from the server
FormUPDATE.FtpClient1.Disconnect
Exit Sub

' Errors
IOUpdateGetUpdateFTPError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "IOUpdateGetUpdateFTP"
ierror = True
Exit Sub

IOUpdateGetUpdateFTPBadConnect:
Screen.MousePointer = vbDefault
msg$ = "FTP connection error: " & FormUPDATE.FtpClient1.LastErrorString
MsgBox msg$, vbOKOnly + vbExclamation, "IOUpdateGetUpdateFTP"
ierror = True
Exit Sub

IOUpdateGetUpdateFTPBadGetFileTime:
Screen.MousePointer = vbDefault
msg$ = "FTP get file time error: " & FormUPDATE.FtpClient1.LastErrorString
MsgBox msg$, vbOKOnly + vbExclamation, "IOUpdateGetUpdateFTP"
ierror = True
FormUPDATE.FtpClient1.Disconnect
Exit Sub

IOUpdateGetUpdateFTPBadGetFile:
Screen.MousePointer = vbDefault
msg$ = "FTP download error: " & FormUPDATE.FtpClient1.LastErrorString
MsgBox msg$, vbOKOnly + vbExclamation, "IOUpdateGetUpdateFTP"
ierror = True
FormUPDATE.FtpClient1.Disconnect
Exit Sub

End Sub

Sub IOUpdateGetUpdateHTTP(mode As Integer)
' Get the update via http
'  mode = 0 get download file date/time only
'  mode = 1 download file

ierror = False
On Error GoTo IOUpdateGetUpdateHTTPError
    
Dim tURL As String, tusername As String, tpassword As String
Dim nError As Long

' Download CalcZAF
If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
tURL$ = "https://epmalab.uoregon.edu/Calczaf/V11/CalcZAF.msi"

' Download penepma12.zip
Else
tURL$ = "https://epmalab.uoregon.edu/Calczaf/PENEPMA12.ZIP"
End If
End If

If UCase$(app.EXEName) = UCase$("Probewin") Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
tURL$ = "https://epmalab.uoregon.edu/updates/V11/ProbeForEPMA.msi"

' Download penepma12.zip
Else
tURL$ = "https://epmalab.uoregon.edu/Calczaf/PENEPMA12.ZIP"
End If
End If

' Set the URL resource user and password (need to use instead of tRemoteFile because of .htaccess security)
tusername$ = "micro"
tpassword$ = "analysis"

' Set the local file path for CalcZAF
If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
tLocalFile$ = ApplicationCommonAppData$ & "CALCZAF.MSI"

' Set local path for penepma12.zip
Else
tLocalFile$ = ApplicationCommonAppData$ & "PENEPMA12.ZIP"
End If
End If

' Set local path for PFE
If UCase$(app.EXEName) = UCase$("Probewin") Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
tLocalFile$ = ApplicationCommonAppData$ & "ProbeForEPMA.MSI"

' Set local path for penepma12.zip
Else
tLocalFile$ = ApplicationCommonAppData$ & "PENEPMA12.ZIP"
End If
End If

' If the user enters an invalid URL, setting the URL property will throw an exception, so that needs to be handled here
On Error Resume Next: Err.Clear
FormUPDATE.HttpClient1.URL = tURL$
If Err.number > 0 Then GoTo IOUpdateGetUpdateHTTPBadURL
On Error GoTo IOUpdateGetUpdateHTTPError

' Load username and password
FormUPDATE.HttpClient1.UserName = tusername$
FormUPDATE.HttpClient1.password = tpassword$

' Establish a connection to the server and display any errors to the user
nError& = FormUPDATE.HttpClient1.Connect()
If nError& > 0 Then GoTo IOUpdateGetUpdateHTTPBadConnection

' Check date/time of remote file
nError& = FormUPDATE.HttpClient1.Localize = True
nError& = FormUPDATE.HttpClient1.GetFileTime(FormUPDATE.HttpClient1.Resource, tRemoteFileDate)
If nError& > 0 Then GoTo IOUpdateGetUpdateHTTPBadGetFileTime

' If just getting remote file date then exit sub
If mode% = 0 Then
FormUPDATE.HttpClient1.Disconnect
Exit Sub
End If

' Download the file to the local system
nError& = FormUPDATE.HttpClient1.Localize = True
nError& = FormUPDATE.HttpClient1.GetFile(tLocalFile$, FormUPDATE.HttpClient1.Resource)
If nError& > 0 Then GoTo IOUpdateGetUpdateHTTPBadGetFile

' Disconnect from the server
FormUPDATE.HttpClient1.Disconnect
Exit Sub

' Errors
IOUpdateGetUpdateHTTPError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "IOUpdateGetUpdateHTTP"
ierror = True
Exit Sub

IOUpdateGetUpdateHTTPBadURL:
Screen.MousePointer = vbDefault
msg$ = "Invalid URL error: " & FormUPDATE.HttpClient1.LastErrorString
MsgBox msg$, vbOKOnly + vbExclamation, "IOUpdateGetUpdateHTTP"
ierror = True
Exit Sub

IOUpdateGetUpdateHTTPBadConnection:
Screen.MousePointer = vbDefault
msg$ = "HTTPS connection error: " & FormUPDATE.HttpClient1.LastErrorString
MsgBox msg$, vbOKOnly + vbExclamation, "IOUpdateGetUpdateHTTP"
ierror = True
Exit Sub

IOUpdateGetUpdateHTTPBadGetFileTime:
Screen.MousePointer = vbDefault
msg$ = "HTTPS get file time error: " & FormUPDATE.HttpClient1.LastErrorString
MsgBox msg$, vbOKOnly + vbExclamation, "IOUpdateGetUpdateHTTP"
ierror = True
FormUPDATE.HttpClient1.Disconnect
Exit Sub

IOUpdateGetUpdateHTTPBadGetFile:
Screen.MousePointer = vbDefault
msg$ = "HTTPS get file error: " & FormUPDATE.HttpClient1.LastErrorString
MsgBox msg$, vbOKOnly + vbExclamation, "IOUpdateGetUpdateHTTP"
ierror = True
FormUPDATE.HttpClient1.Disconnect
Exit Sub

End Sub

Sub IOUpdateComplete()
' Download complete. Notify user and extract files using a spawned process.

ierror = False
On Error GoTo IOUpdateCompleteError

Dim taskID As Long

Screen.MousePointer = vbDefault
msg$ = "The download is complete and the update file is " & tLocalFile$ & ". Click OK to close the current program and update the application files automatically."
FormMSGBOXDOEVENTS2.Caption = "IOUpdateComplete"
FormMSGBOXDOEVENTS2.Label1.Caption = msg$
FormMSGBOXDOEVENTS2.Show vbModal

' Change path to download folder
Call MiscChangePath(ApplicationCommonAppData$)
If ierror Then Exit Sub

' Run the batch file to extract the CalcZAF update
If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
taskID& = Shell("msiexec /i calczaf.msi", vbNormalFocus)
'taskID& = Shell("msiexec /i calczaf.msi /l*v install.log", vbNormalFocus)   ' creates installer log
'Call IORunShellExecute("open", "calczaf.msi", "/i", ApplicationCommonAppData$, SW_SHOWNORMAL&)

' Extract Penepma.ZIP to Penepma folder
Else
taskID& = Shell("cmd.exe /k pkzip25 -extract -direct -overwrite -exclude=pkzip25.exe " & VbDquote$ & ApplicationCommonAppData$ & "Penepma12.zip" & VbDquote$ & " " & VbDquote$ & PENEPMA_Root$ & VbDquote$, vbNormalFocus)
End If
End If

' Run the batch file to extract the Probewin update
If UCase$(app.EXEName) = UCase$("Probewin") Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
taskID& = Shell("msiexec /i ProbeForEPMA.msi", vbNormalFocus)
'taskID& = Shell("msiexec /i ProbeForEPMA.msi /l*v install.log", vbNormalFocus)   ' creates installer log
'Call IORunShellExecute("open", "ProbeForEPMA.msi", "/i", ApplicationCommonAppData$, SW_SHOWNORMAL&)

' Extract Penepma.ZIP
Else
taskID& = Shell("cmd.exe /k pkzip25 -extract -direct -overwrite -exclude=pkzip25.exe " & VbDquote$ & ApplicationCommonAppData$ & "Penepma12.zip" & VbDquote$ & " " & VbDquote$ & PENEPMA_Root$ & VbDquote$, vbNormalFocus)
End If
End If

' Close the current program (make sure automation is not paused)
If RealTimeMode Then
If RealTimePauseAutomation Then RealTimePauseAutomation = False
End If

Unload FormMAIN
Exit Sub

' Errors
IOUpdateCompleteError:
MsgBox Error$, vbOKOnly + vbCritical, "IOUpdateComplete"
ierror = True
Exit Sub

End Sub

Sub IOUpdateDeleteUpdate()
' Delete the previously (failed?) update file download. This will force an update.

ierror = False
On Error GoTo IOUpdateDeleteUpdateError

If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
tLocalFile$ = ApplicationCommonAppData$ & "CALCZAF.MSI"

' Delete existing Penepma12.zip
Else
tLocalFile$ = ApplicationCommonAppData$ & "Penepma12.zip"
End If
End If

If UCase$(app.EXEName) = UCase$("Probewin") Then
If FormUPDATE.CheckUpdatePenepmaOnly.value = vbUnchecked Then
tLocalFile$ = ApplicationCommonAppData$ & "ProbeForEPMA.MSI"

' Delete existing Penepma12.zip
Else
tLocalFile$ = ApplicationCommonAppData$ & "Penepma12.zip"
End If
End If

' Check if file exists
If Dir$(tLocalFile$) <> vbNullString Then
Kill tLocalFile$
msg$ = "The previously downloaded update file has been deleted. This will allow the update to proceed even if it is not really necessary."
MsgBox msg$, vbOKOnly + vbExclamation, "IOUpdateDeleteUpdate"
Else
msg$ = "No previously downloaded update file was found. You may proceed with the update in any case though your current version will not be automatically backed up."
MsgBox msg$, vbOKOnly + vbExclamation, "IOUpdateDeleteUpdate"
End If

Exit Sub

' Errors
IOUpdateDeleteUpdateError:
MsgBox Error$, vbOKOnly + vbCritical, "IOUpdateDeleteUpdate"
ierror = True
Exit Sub

End Sub


