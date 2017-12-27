Attribute VB_Name = "CodeIOStatus"
' (c) Copyright 1995-2018 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub IOAutomationCancel()
' Cancel the current automation (non realtime version)

ierror = False
On Error GoTo IOAutomationCancelError

Dim response As Integer

FormMAIN.StatusBarAuto.Panels(2).Bevel = sbrInset
Call MiscTimer(CSng(0.2))
FormMAIN.StatusBarAuto.Panels(2).Bevel = sbrRaised

If AcquisitionOnAutomate Then
msg$ = "Are you sure you want to cancel the current automation action?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton2, "IOAutomationCancel")
If response% = vbNo Then Exit Sub
End If

RealTimeInterfaceBusy = False   ' force attention
DoEvents
RealTimePauseAutomation = False   ' force attention
DoEvents

icancelauto = True
DoEvents
Exit Sub

' Errors
IOAutomationCancelError:
MsgBox Error$, vbOKOnly + vbCritical, "IOAutomationCancel"
ierror = True
Exit Sub

End Sub

Sub IOAutomationPause(mode As Integer)
' Pause the current automation
' 0 = toggle the current pause status
' 1 = force automation pause
' 2 = force automation continue

ierror = False
On Error GoTo IOAutomationPauseError

FormMAIN.StatusBarAuto.Panels(3).Bevel = sbrInset
Call MiscTimer(CSng(0.2))
FormMAIN.StatusBarAuto.Panels(3).Bevel = sbrRaised

' Pause
If (mode% = 0 And Not RealTimePauseAutomation) Or mode% = 1 Then
FormMAIN.StatusBarAuto.Panels(3).Text = "Continue"
Call IOStatusAuto("Warning: automation paused...")
End If

' Continue
If (mode% = 0 And RealTimePauseAutomation) Or mode% = 2 Then
FormMAIN.StatusBarAuto.Panels(3).Text = "Pause"
Call IOStatusAuto(vbNullString)
End If


If mode% = 0 Then RealTimePauseAutomation = Not RealTimePauseAutomation
If mode% = 1 Then RealTimePauseAutomation = True
If mode% = 2 Then RealTimePauseAutomation = False
DoEvents

Exit Sub

' Errors
IOAutomationPauseError:
MsgBox Error$, vbOKOnly + vbCritical, "IOAutomationPause"
ierror = True
Exit Sub

End Sub

Sub IOStatusAuto(astring As String)
' This routine writes the string to the proper label for automation status purposes

ierror = False
On Error GoTo IOStatusAutoError

' Update status bar
FormMAIN.StatusBarAuto.Panels(1).Text = astring$
DoEvents
Exit Sub

' Errors
IOStatusAutoError:
MsgBox Error$, vbOKOnly + vbCritical, "IOStatusAuto"
ierror = True
Exit Sub

End Sub


