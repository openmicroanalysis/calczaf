Attribute VB_Name = "CodeCOND2"
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Dim CondTmpSample(1 To 1) As TypeSample

Sub Cond2Load(sample() As TypeSample)
' Loads the analytical condition defaults for user to change (for single element only for combined conditions)

ierror = False
On Error GoTo Cond2LoadError

Dim i As Integer

' Load sample into module level array
CondTmpSample(1) = sample(1)

' Load element list
FormCOND2.ComboElementXraySpectrometerCrystal.Clear
For i% = 1 To CondTmpSample(1).LastElm%
msg$ = CondTmpSample(1).Elsyms$(i%) & " " & CondTmpSample(1).Xrsyms$(i%) & ", Spec " & Str$(CondTmpSample(1).MotorNumbers%(i%)) & " " & CondTmpSample(1).CrystalNames$(i%)
FormCOND2.ComboElementXraySpectrometerCrystal.AddItem msg$
FormCOND2.ComboElementXraySpectrometerCrystal.ItemData(FormCOND2.ComboElementXraySpectrometerCrystal.NewIndex) = i%
Next i%

' Select first element
If FormCOND2.ComboElementXraySpectrometerCrystal.ListCount > 0 Then
FormCOND2.ComboElementXraySpectrometerCrystal.ListIndex = 0
End If

' Realtime modifications
If OperatingVoltagePresent Then
FormCOND2.TextKiloVolts.Enabled = True
Else
FormCOND2.TextKiloVolts.Enabled = False
End If

' Load element list for channel order operations
Call Cond2LoadChannels
If ierror Then Exit Sub

Exit Sub

' Errors
Cond2LoadError:
MsgBox Error$, vbOKOnly + vbCritical, "Cond2Load"
ierror = True
Exit Sub

End Sub

Sub Cond2SaveField()
' Save changed conditions to defaults (analytical condition for a single element)

ierror = False
On Error GoTo Cond2SaveFieldError

Dim i As Integer
Dim tfov As Single

If Val(FormCOND2.TextKiloVolts.Text) < MINKILOVOLTS! Or Val(FormCOND2.TextKiloVolts.Text) > MAXKILOVOLTS! Then
msg$ = FormCOND2.TextKiloVolts.Text & " for Kilovolts is out of range!"
MsgBox msg$, vbOKOnly + vbExclamation, "Cond2SaveField"
ierror = True
Exit Sub
End If

If Val(FormCOND2.TextTakeOff.Text) < 1# Or Val(FormCOND2.TextTakeOff.Text) > 90# Then
msg$ = FormCOND2.TextTakeOff.Text & " for TakeOff is out of range!"
MsgBox msg$, vbOKOnly + vbExclamation, "Cond2SaveField"
ierror = True
Exit Sub
End If

' Get selected element channel
If FormCOND2.ComboElementXraySpectrometerCrystal.ListCount < 1 Then Exit Sub
If FormCOND2.ComboElementXraySpectrometerCrystal.ListIndex < 0 Then Exit Sub
i% = FormCOND2.ComboElementXraySpectrometerCrystal.ItemData(FormCOND2.ComboElementXraySpectrometerCrystal.ListIndex)

CondTmpSample(1).TakeoffArray!(i%) = Val(FormCOND2.TextTakeOff.Text)
CondTmpSample(1).KilovoltsArray!(i%) = Val(FormCOND2.TextKiloVolts.Text)

' Check if sample is still combined
CondTmpSample(1).CombinedConditionsFlag = False
If MiscIsDifferent3(CondTmpSample(1).LastElm%, CondTmpSample(1).TakeoffArray!()) Then CondTmpSample(1).CombinedConditionsFlag = True
If MiscIsDifferent3(CondTmpSample(1).LastElm%, CondTmpSample(1).KilovoltsArray!()) Then CondTmpSample(1).CombinedConditionsFlag = True

' Force reload of afactor arrays in case conditions change
AllAnalysisUpdateNeeded = True
AllAFactorUpdateNeeded = True

Exit Sub

' Errors
Cond2SaveFieldError:
MsgBox Error$, vbOKOnly + vbCritical, "Cond2SaveField"
ierror = True
Exit Sub

End Sub

Sub Cond2Return(sample() As TypeSample)
' Return the modified sample for other code modules (AcquireChangeConditions2)

ierror = False
On Error GoTo Cond2ReturnError

sample(1) = CondTmpSample(1)
Exit Sub

' Errors
Cond2ReturnError:
MsgBox Error$, vbOKOnly + vbCritical, "Cond2Return"
ierror = True
Exit Sub

End Sub

Sub Cond2LoadField()
' Load the fields

ierror = False
On Error GoTo Cond2LoadFieldError

Dim i As Integer

If FormCOND2.ComboElementXraySpectrometerCrystal.ListCount < 1 Then Exit Sub
If FormCOND2.ComboElementXraySpectrometerCrystal.ListIndex < 0 Then Exit Sub
i% = FormCOND2.ComboElementXraySpectrometerCrystal.ItemData(FormCOND2.ComboElementXraySpectrometerCrystal.ListIndex)

' Load defaults if necessary
If CondTmpSample(1).KilovoltsArray!(i%) = 0# Then CondTmpSample(1).KilovoltsArray!(i%) = CondTmpSample(1).kilovolts!
If CondTmpSample(1).TakeoffArray!(i%) = 0# Then CondTmpSample(1).TakeoffArray!(i%) = CondTmpSample(1).takeoff!

' Load fields
FormCOND2.TextKiloVolts.Text = Format$(CondTmpSample(1).KilovoltsArray!(i%))
FormCOND2.TextTakeOff.Text = Format$(CondTmpSample(1).TakeoffArray!(i%))

Exit Sub

' Errors
Cond2LoadFieldError:
MsgBox Error$, vbOKOnly + vbCritical, "Cond2LoadField"
ierror = True
Exit Sub

End Sub

Sub Cond2Sort(upordown As Integer)
' Save new channel order (channel up or down)
' chan = channel to be re-sorted
' upordown 1 = up, 2 = down

ierror = False
On Error GoTo Cond2SortError

Dim chan As Integer, chan2 As Integer

' Get selected channel
If FormCOND2.ListElements.ListCount < 1 Then Exit Sub
If FormCOND2.ListElements.ListIndex < 0 Then Exit Sub
chan% = FormCOND2.ListElements.ListIndex + 1

' Get new channel
If upordown = 1 Then
chan2% = chan% + 1
If chan2% > CondTmpSample(1).LastElm% Then Exit Sub
End If
If upordown = 2 Then
chan2% = chan% - 1
If chan2% < 1 Then Exit Sub
End If

' Sort new sample
Call GetElmSaveSampleOnly(CondTmpSample(), chan%, chan2%)
If ierror Then Exit Sub

' Reload the form
Call Cond2Load(CondTmpSample())
If ierror Then Exit Sub

Call Cond2LoadField
If ierror Then Exit Sub

' Reselect channel
FormCOND2.ListElements.Selected(chan2% - 1) = True

Exit Sub

' Errors
Cond2SortError:
MsgBox Error$, vbOKOnly + vbCritical, "Cond2Sort"
ierror = True
Exit Sub

End Sub

Sub Cond2Apply()
' Save condition changes and update form

ierror = False
On Error GoTo Cond2ApplyError

Call Cond2SaveField
If ierror Then Exit Sub

' Update channel order list
Call Cond2LoadChannels
If ierror Then Exit Sub

Call Cond2LoadField
If ierror Then Exit Sub

Exit Sub

' Errors
Cond2ApplyError:
MsgBox Error$, vbOKOnly + vbCritical, "Cond2Apply"
ierror = True
Exit Sub

End Sub

Sub Cond2LoadChannels()
' Load the chanel order list

ierror = False
On Error GoTo Cond2LoadChannelsError

Dim i As Integer

FormCOND2.ListElements.Clear
For i% = 1 To CondTmpSample(1).LastElm%
msg$ = Format$(CondTmpSample(1).Elsyms$(i%) & " " & CondTmpSample(1).Xrsyms$(i%), a50$) & vbTab
msg$ = msg$ & " KeV" & Str$(CondTmpSample(1).KilovoltsArray!(i%))
FormCOND2.ListElements.AddItem msg$
FormCOND2.ListElements.ItemData(FormCOND2.ListElements.NewIndex) = i%
Next i%

Exit Sub

' Errors
Cond2LoadChannelsError:
MsgBox Error$, vbOKOnly + vbCritical, "Cond2LoadChannels"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub
