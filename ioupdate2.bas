Attribute VB_Name = "CodeIOUpdate2"
' (c) Copyright 1995-2026 by John J. Donovan
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
Dim tLocalFile As String
Dim tBackupFile As String

Dim tLocalFileDate As Variant, tRemoteFileDate As Variant

Sub IOUpdate2Close(Cancel As Integer)
' Check if download is running and close dialog

ierror = False
On Error GoTo IOUpdate2CloseError

Dim response As Integer

If FormUPDATE2.HttpClient1.Connected Then
msg$ = "The download is not complete, are you sure you want to cancel the update download?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton2, "IOUpdate2Close")
If response% = vbYes Then
icancel = True
FormUPDATE2.HttpClient1.Cancel

If UCase$(app.EXEName) = UCase$("CalcZAF") Then
FormUPDATE2.Caption = "Update CalcZAF [Please wait for termination...]"
End If

If UCase$(app.EXEName) = UCase$("Probewin") Then
FormUPDATE2.Caption = "Update Probe for EPMA [Please wait for termination...]"
End If

Call MiscDelay(CDbl(2#), Now)
FormUPDATE2.HttpClient1.Disconnect
DoEvents
Exit Sub
Else
Cancel = True
End If
End If

Exit Sub

' Errors
IOUpdate2CloseError:
MsgBox Error$, vbOKOnly + vbCritical, "IOUpdate2Close"
ierror = True
Exit Sub

End Sub

Sub IOUpdate2GetUpdate(mode As Integer)
' Get the latest update for Probe for EPMA/CalcZAF/PENEPMA using HTTPS
'  mode = 1 download from EPMALab (U Oregon)
'  mode = 2 download from Probe Software

ierror = False
On Error GoTo IOUpdate2GetUpdateError

' Check for valid version
If ProgramVersionNumber! < 8.11 Then GoTo IOUpdate2GetUpdateOldVersion

DownloadMode% = mode%

' Check for trace mode (only for PFE because CalcZAF runs in debug mode by default)
If UCase$(app.EXEName) = UCase$("Probewin") Then
If DebugMode Then
FormUPDATE2.HttpClient1.TraceFile = ApplicationCommonAppData$ & "HttpTest.log"
FormUPDATE2.HttpClient1.TraceFlags = httpTraceHexDump&
FormUPDATE2.HttpClient1.Trace = True
End If
End If

' Load filenames for CalcZAF update
If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If FormUPDATE2.CheckUpdatePenepmaOnly.value = vbUnchecked Then
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
If FormUPDATE2.CheckUpdatePenepmaOnly.value = vbUnchecked Then
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

' Get date/time using HTTP
Call IOUpdate2GetUpdateHTTP(Int(0))
If ierror Then Exit Sub

' See if update is necessary
Call IOWriteLog(vbCrLf & "IOUpdate2GetUpdate: date/time of last update (" & tLocalFileDate & "), date/time of current update (" & tRemoteFileDate & ")...")
DoEvents
If CDate(tLocalFileDate) > CDate(tRemoteFileDate) Then
If FormUPDATE2.CheckUpdatePenepmaOnly.value = vbUnchecked Then
msg$ = "The current version of this program (" & ProgramVersionString$ & ") is already up to date. To force an update download, please use the Delete Update button first and try again."
Else
msg$ = "The current PENEPMA12.ZIP file is already up to date. To force an update download, please use the Delete Update button first and try again."
End If
MsgBox msg$, vbOKOnly + vbExclamation, "IOUpdate2GetUpdate"
ierror = True
Exit Sub
End If

' Check if file exists already and if so backup and then delete
If Dir$(tLocalFile$) <> vbNullString Then
If Dir$(tBackupFile$) <> vbNullString Then
Call IOWriteLog("IOUpdate2GetUpdate: Removing previous update backup (" & tBackupFile$ & ")...")
DoEvents
Kill tBackupFile$
End If
Call IOWriteLog("IOUpdate2GetUpdate: Copying previous update (" & tLocalFile$ & ") to update backup (" & tBackupFile$ & ")...")
DoEvents
FileCopy tLocalFile$, tBackupFile$
Call IOWriteLog("IOUpdate2GetUpdate: Removing previous update (" & tLocalFile$ & ")...")
DoEvents
Kill tLocalFile$
DoEvents
End If

' Get using HTTPS
Screen.MousePointer = vbHourglass
Call IOWriteLog("IOUpdate2GetUpdate: Downloading update using secure HTTPS (" & tLocalFile$ & ")...")
DoEvents
If RealTimeMode% Then Call IOAutomationPause(Int(1))
Call IOUpdate2GetUpdateHTTP(Int(1))
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

Call IOWriteLog("IOUpdate2GetUpdate: Download of file (" & tLocalFile$ & ") is complete")

' Notify user
If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If FormUPDATE2.CheckUpdatePenepmaOnly.value = vbUnchecked Then
FormUPDATE2.Caption = "Update CalcZAF [download complete]"
Else
FormUPDATE2.Caption = "Update Penepma [download complete]"
End If
End If

If UCase$(app.EXEName) = UCase$("Probewin") Then
If FormUPDATE2.CheckUpdatePenepmaOnly.value = vbUnchecked Then
FormUPDATE2.Caption = "Update Probe for EPMA [download complete]"
Else
FormUPDATE2.Caption = "Update Penepma [download complete]"
End If
End If

DoEvents
If RealTimeMode% Then Call IOAutomationPause(Int(2))
Call IOUpdate2Complete
If ierror Then Exit Sub

Unload FormUPDATE2
Exit Sub

' Errors
IOUpdate2GetUpdateError:
MsgBox Error$, vbOKOnly + vbCritical, "IOUpdate2GetUpdate"
ierror = True
Exit Sub

IOUpdate2GetUpdateOldVersion:
msg$ = "The current version of this program (" & ProgramVersionString$ & ") is too old to update automatically. The program will have to be re-installed. Please contact Probe Software for more information."
MsgBox msg$, vbOKOnly + vbExclamation, "IOUpdate2GetUpdate"
ierror = True
Exit Sub

End Sub

Sub IOUpdate2GetUpdateHTTP(mode As Integer)
' Get the update via http
'  mode = 0 get download file date/time only
'  mode = 1 download file

ierror = False
On Error GoTo IOUpdate2GetUpdateHTTPError
    
Dim tURL As String
Dim nError As Long

' Download CalcZAF
If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If FormUPDATE2.CheckUpdatePenepmaOnly.value = vbUnchecked Then
If DownloadMode% = 1 Then
tURL$ = "https://epmalab.uoregon.edu/download/CalcZAF.msi"
Else
tURL$ = "https://www.probesoftware.com/download/CalcZAF.msi"
End If

' Download penepma12.zip
Else
If DownloadMode% = 1 Then
tURL$ = "https://epmalab.uoregon.edu/download/CalcZAF.msi"
Else
tURL$ = "https://www.probesoftware.com/download/PENEPMA12.ZIP"
End If
End If
End If

' Download Probe for EPMA
If UCase$(app.EXEName) = UCase$("Probewin") Then
If FormUPDATE2.CheckUpdatePenepmaOnly.value = vbUnchecked Then
If DownloadMode% = 1 Then
tURL$ = "https://epmalab.uoregon.edu/download/ProbeForEPMA.msi"
Else
tURL$ = "https://www.probesoftware.com/download/ProbeForEPMA.msi"
End If

' Download penepma12.zip
Else
If DownloadMode% = 1 Then
tURL$ = "https://epmalab.uoregon.edu/download/PENEPMA12.ZIP"
Else
tURL$ = "https://www.probesoftware.com/download/PENEPMA12.ZIP"
End If
End If
End If

Call IOWriteLog("IOUpdate2GetUpdateHTTP: Download URL: " & tURL$)
DoEvents

' Set the local file path for CalcZAF
If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If FormUPDATE2.CheckUpdatePenepmaOnly.value = vbUnchecked Then
tLocalFile$ = ApplicationCommonAppData$ & "CALCZAF.MSI"

' Set local path for penepma12.zip
Else
tLocalFile$ = ApplicationCommonAppData$ & "PENEPMA12.ZIP"
End If
End If

' Set local path for PFE
If UCase$(app.EXEName) = UCase$("Probewin") Then
If FormUPDATE2.CheckUpdatePenepmaOnly.value = vbUnchecked Then
tLocalFile$ = ApplicationCommonAppData$ & "ProbeForEPMA.MSI"

' Set local path for penepma12.zip
Else
tLocalFile$ = ApplicationCommonAppData$ & "PENEPMA12.ZIP"
End If
End If

' If the user enters an invalid URL, setting the URL property will throw an exception, so that needs to be handled here
On Error Resume Next: Err.Clear
FormUPDATE2.HttpClient1.URL = tURL$
If Err.number > 0 Then GoTo IOUpdate2GetUpdateHTTPBadURL
On Error GoTo IOUpdate2GetUpdateHTTPError

' Establish a connection to the server and display any errors to the user
nError& = FormUPDATE2.HttpClient1.Connect()
If nError& > 0 Then GoTo IOUpdate2GetUpdateHTTPBadConnection

' Check date/time of remote file
nError& = FormUPDATE2.HttpClient1.Localize = True
nError& = FormUPDATE2.HttpClient1.GetFileTime(FormUPDATE2.HttpClient1.Resource, tRemoteFileDate)
If nError& > 0 Then GoTo IOUpdate2GetUpdateHTTPBadGetFileTime

' If just getting remote file date then exit sub
If mode% = 0 Then
FormUPDATE2.HttpClient1.Disconnect
Exit Sub
End If

' Download the file to the local system
nError& = FormUPDATE2.HttpClient1.Localize = False
nError& = FormUPDATE2.HttpClient1.GetFile(tLocalFile$, FormUPDATE2.HttpClient1.Resource)
If nError& > 0 Then GoTo IOUpdate2GetUpdateHTTPBadGetFile

' Disconnect from the server
FormUPDATE2.HttpClient1.Disconnect
Exit Sub

' Errors
IOUpdate2GetUpdateHTTPError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "IOUpdate2GetUpdateHTTP"
ierror = True
Exit Sub

IOUpdate2GetUpdateHTTPBadURL:
Screen.MousePointer = vbDefault
msg$ = "Invalid URL error: " & FormUPDATE2.HttpClient1.LastErrorString
MsgBox msg$, vbOKOnly + vbExclamation, "IOUpdate2GetUpdateHTTP"
ierror = True
Exit Sub

IOUpdate2GetUpdateHTTPBadConnection:
Screen.MousePointer = vbDefault
msg$ = "HTTPS connection error: " & FormUPDATE2.HttpClient1.LastErrorString
MsgBox msg$, vbOKOnly + vbExclamation, "IOUpdate2GetUpdateHTTP"
ierror = True
Exit Sub

IOUpdate2GetUpdateHTTPBadGetFileTime:
Screen.MousePointer = vbDefault
msg$ = "HTTPS get file time error: " & FormUPDATE2.HttpClient1.LastErrorString
MsgBox msg$, vbOKOnly + vbExclamation, "IOUpdate2GetUpdateHTTP"
ierror = True
FormUPDATE2.HttpClient1.Disconnect
Exit Sub

IOUpdate2GetUpdateHTTPBadGetFile:
Screen.MousePointer = vbDefault
msg$ = "HTTPS get file error: " & FormUPDATE2.HttpClient1.LastErrorString
MsgBox msg$, vbOKOnly + vbExclamation, "IOUpdate2GetUpdateHTTP"
ierror = True
FormUPDATE2.HttpClient1.Disconnect
Exit Sub

End Sub

Sub IOUpdate2Complete()
' Download complete. Notify user and extract files using a spawned process.

ierror = False
On Error GoTo IOUpdate2CompleteError

Dim taskID As Long

Screen.MousePointer = vbDefault
msg$ = "The download is complete and the update file is " & tLocalFile$ & ". Click OK to close the current program and update the application files automatically."
FormMSGBOXDOEVENTS2.Caption = "IOUpdate2Complete"
FormMSGBOXDOEVENTS2.Label1.Caption = msg$
FormMSGBOXDOEVENTS2.Show vbModal

' Change path to download folder
Call MiscChangePath(ApplicationCommonAppData$)
If ierror Then Exit Sub

' Run the batch file to extract the CalcZAF update
If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If FormUPDATE2.CheckUpdatePenepmaOnly.value = vbUnchecked Then
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
If FormUPDATE2.CheckUpdatePenepmaOnly.value = vbUnchecked Then
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
IOUpdate2CompleteError:
MsgBox Error$, vbOKOnly + vbCritical, "IOUpdate2Complete"
ierror = True
Exit Sub

End Sub

Sub IOUpdate2DeleteUpdate()
' Delete the previously (failed?) update file download. This will force an update.

ierror = False
On Error GoTo IOUpdate2DeleteUpdateError

If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If FormUPDATE2.CheckUpdatePenepmaOnly.value = vbUnchecked Then
tLocalFile$ = ApplicationCommonAppData$ & "CALCZAF.MSI"

' Delete existing Penepma12.zip
Else
tLocalFile$ = ApplicationCommonAppData$ & "Penepma12.zip"
End If
End If

If UCase$(app.EXEName) = UCase$("Probewin") Then
If FormUPDATE2.CheckUpdatePenepmaOnly.value = vbUnchecked Then
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
MsgBox msg$, vbOKOnly + vbExclamation, "IOUpdate2DeleteUpdate"
Else
msg$ = "No previously downloaded update file was found. You may proceed with the update in any case though your current version will not be automatically backed up."
MsgBox msg$, vbOKOnly + vbExclamation, "IOUpdate2DeleteUpdate"
End If

Exit Sub

' Errors
IOUpdate2DeleteUpdateError:
MsgBox Error$, vbOKOnly + vbCritical, "IOUpdate2DeleteUpdate"
ierror = True
Exit Sub

End Sub


