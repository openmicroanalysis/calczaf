Attribute VB_Name = "CodeIO2"
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

