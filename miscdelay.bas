Attribute VB_Name = "CodeMiscDelay"
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub MiscTimer(timeinterval As Single)
' Waits a specified number of seconds before returning (fractional second precision)
' Usage: Call MiscTimer(secondsdelay!)

ierror = False
On Error GoTo MiscTimerError

Dim startsec As Single, elapsedsec As Single

If timeinterval! <= 0# Then Exit Sub
If timeinterval! > SECPERDAY# Then GoTo MiscTimerBadInterval

Screen.MousePointer = vbHourglass
icancelauto = False
startsec! = CSng(Timer)   ' set start seconds
Do While elapsedsec! < startsec! + timeinterval!
    
    elapsedsec! = CSng(Timer)
    If elapsedsec! < startsec! Then startsec! = startsec! - SECPERDAY#     ' in case timer goes through midnight
    DoEvents   ' yield to other processes
    Sleep (timeinterval! * MSECPERSEC# / 10#) ' yield to other apps

    ' Check for cancel
    If icancelauto Then
    Screen.MousePointer = vbDefault
    ierror = True
    Exit Sub
    End If

Loop

Screen.MousePointer = vbDefault
Exit Sub

' Errors
MiscTimerError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "MiscTimer"
ierror = True
Exit Sub

MiscTimerBadInterval:
msg$ = "Timer interval exceeds number of seconds in a day"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscTimer"
ierror = True
Exit Sub

End Sub

Sub MiscTimer2(timeinterval As Single)
' Waits a specified number of seconds before returning (fractional second precision)
' Usage: Call MiscTimer2(secondsdelay!) (no hourglass)

ierror = False
On Error GoTo MiscTimer2Error

Dim startsec As Single, elapsedsec As Single

If timeinterval! <= 0# Then Exit Sub
If timeinterval! > SECPERDAY# Then GoTo MiscTimer2BadInterval

icancelauto = False
startsec! = CSng(Timer)   ' set start seconds
Do While elapsedsec! < startsec! + timeinterval!
    
    elapsedsec! = CSng(Timer)
    If elapsedsec! < startsec! Then startsec! = startsec! - SECPERDAY#     ' in case timer goes through midnight
    DoEvents   ' yield to other processes
    Sleep (timeinterval! * MSECPERSEC# / 10#) ' yield to other apps

    ' Check for cancel
    If icancelauto Then
    ierror = True
    Exit Sub
    End If

Loop

Exit Sub

' Errors
MiscTimer2Error:
MsgBox Error$, vbOKOnly + vbCritical, "MiscTimer2"
ierror = True
Exit Sub

MiscTimer2BadInterval:
msg$ = "Timer interval exceeds number of seconds in a day"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscTimer2"
ierror = True
Exit Sub

End Sub

Sub MiscDelay(timeinterval As Double, previoustime As Double)
' Waits a specified number of seconds before returning
' Usage: Call MiscDelay(CDbl(secondsdelay!), Now)

ierror = False
On Error GoTo MiscDelayError

If timeinterval# <= 0# Then Exit Sub

' Loop until current date/time exceeds specified interval in seconds
Screen.MousePointer = vbHourglass
Do
If Now > timeinterval# / SECPERDAY# + previoustime# Then Exit Do
DoEvents

' Check for cancel
If icancelauto Then
Screen.MousePointer = vbDefault
ierror = True
Exit Sub
End If

Loop

Screen.MousePointer = vbDefault
Exit Sub

' Errors
MiscDelayError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "MiscDelay"
ierror = True
Exit Sub

End Sub

Sub MiscDelay2(nloop As Long)
' Waits a specified number of loops before returning
' Usage: Call MiscDelay2(1000)

ierror = False
On Error GoTo MiscDelay2Error

Dim n As Long

' Loop
Screen.MousePointer = vbHourglass
For n = 1 To nloop&
DoEvents

' Check for cancel
If icancelauto Then
Screen.MousePointer = vbDefault
ierror = True
Exit Sub
End If

Next n&

Screen.MousePointer = vbDefault
Exit Sub

' Errors
MiscDelay2Error:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "MiscDelay2"
ierror = True
Exit Sub

End Sub

Sub MiscDelay3(tStatus As StatusBar, timeinterval As Double, previoustime As Double)
' Waits a specified number of seconds before returning (updates status text)
' Usage: Call MiscDelay3(tStatus as StatusBar, CDbl(secondsdelay!), Now)

ierror = False
On Error GoTo MiscDelay3Error

Dim atemp As Double

If timeinterval# <= 0# Then Exit Sub

' Loop until current date/time exceeds specified interval in seconds
Screen.MousePointer = vbHourglass
Do
If Now > timeinterval# / SECPERDAY# + previoustime# Then Exit Do
DoEvents

' Update caption
atemp# = (timeinterval# / SECPERDAY# + previoustime#) - Now
tStatus.Panels(1).Text = Format$(atemp# * SECPERDAY#, f81$) & " seconds remaining for next automation process..."
DoEvents    ' yield to this app
Sleep (100) ' yield to other apps

' Check for cancel
If icancelauto Then
Screen.MousePointer = vbDefault
ierror = True
Exit Sub
End If

Loop

Screen.MousePointer = vbDefault
Exit Sub

' Errors
MiscDelay3Error:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "MiscDelay3"
ierror = True
Exit Sub

End Sub

Sub MiscDelay4(timeinterval As Double, previoustime As Double)
' Waits a specified number of seconds before returning
'  (no doevents- program will "freeze" during the delay)
' Usage: Call MiscDelay4(CDbl(secondsdelay!), Now)

ierror = False
On Error GoTo MiscDelay4Error

If timeinterval# <= 0# Then Exit Sub

' Loop until current date/time exceeds specified interval in seconds
Screen.MousePointer = vbHourglass
Do
If Now > timeinterval# / SECPERDAY# + previoustime# Then Exit Do
Loop

Screen.MousePointer = vbDefault
Exit Sub

' Errors
MiscDelay4Error:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "MiscDelay4"
ierror = True
Exit Sub

End Sub

Sub MiscDelay5(timeinterval As Double, previoustime As Double)
' Waits a specified number of seconds before returning  (without hourglass cursor)
' Usage: Call MiscDelay5(CDbl(secondsdelay!), Now)

ierror = False
On Error GoTo MiscDelay5Error

If timeinterval# <= 0# Then Exit Sub

' Loop until current date/time exceeds specified interval in seconds
Do
If Now > timeinterval# / SECPERDAY# + previoustime# Then Exit Do
DoEvents

' Check for cancel
If icancelauto Then
ierror = True
Exit Sub
End If

Loop

Exit Sub

' Errors
MiscDelay5Error:
MsgBox Error$, vbOKOnly + vbCritical, "MiscDelay5"
ierror = True
Exit Sub

End Sub


