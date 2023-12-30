Attribute VB_Name = "CodeGETPTC"
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

Sub GetPTCLoad()
' Load PTC options

ierror = False
On Error GoTo GetPTCLoadError

Dim i As Integer

' Load PTC flag
If iptc% = 0 Then
FormGETPTC.CheckUsePTC.value = vbUnchecked
Else
FormGETPTC.CheckUsePTC.value = vbChecked
End If

' Load models
FormGETPTC.ComboPTCModel.Clear
For i% = 1 To UBound(ptcstring$)
FormGETPTC.ComboPTCModel.AddItem ptcstring$(i%)
Next i%

' Load current value
If PTCModel% = 0 Then PTCModel% = 1
FormGETPTC.ComboPTCModel.ListIndex = PTCModel% - 1

' Load numerical values
If PTCDiameter! = 0# Then PTCDiameter! = 10000#
FormGETPTC.TextPTCDiameter.Text = MiscAutoFormat$(PTCDiameter!)

If PTCDensity! = 0# Then PTCDensity! = 3#
FormGETPTC.TextPTCDensity.Text = MiscAutoFormat$(PTCDensity!)

If PTCThicknessFactor! = 0 Then PTCThicknessFactor! = 1#
FormGETPTC.TextPTCThicknessFactor.Text = MiscAutoFormat$(PTCThicknessFactor!)

If PTCNumericalIntegrationStep! = 0 Then PTCNumericalIntegrationStep! = 0.00001
FormGETPTC.TextPTCNumericalIntegrationStep.Text = MiscAutoFormat$(PTCNumericalIntegrationStep!)

If PTCDoNotNormalizeSpecifiedFlag Then
FormGETPTC.CheckPTCDoNotNormalizeSpecifiedFlag.value = vbChecked
Else
FormGETPTC.CheckPTCDoNotNormalizeSpecifiedFlag.value = vbUnchecked
End If
If ProbeDataFileVersionNumber! < 10.68 Then FormGETPTC.CheckPTCDoNotNormalizeSpecifiedFlag.Enabled = False

FormGETPTC.Show vbModal

Exit Sub

' Errors
GetPTCLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "GetPTCLoad"
ierror = True
Exit Sub

End Sub

Sub GetPTCSave()
' Save PTC options

ierror = False
On Error GoTo GetPTCSaveError

' Save PTC flag
If FormGETPTC.CheckUsePTC.value = vbUnchecked Then
iptc% = 0
Else
iptc% = 1
UseParticleCorrectionFlag = True    ' if user enables any particle then turn on global flag
End If

' Save model
If FormGETPTC.ComboPTCModel.ListIndex% >= 0 Then
PTCModel% = FormGETPTC.ComboPTCModel.ListIndex% + 1
End If

' Save numerical values (diameter saved in microns)
If Val(FormGETPTC.TextPTCDiameter.Text) < 0.001 Or Val(FormGETPTC.TextPTCDiameter.Text) > 10000# Then
msg$ = "Particle diameter is out of range (must be between 0.001 and 10000 microns)"
MsgBox msg$, vbOKOnly + vbExclamation, "GetPTCSave"
ierror = True
Exit Sub
Else
PTCDiameter! = Val(FormGETPTC.TextPTCDiameter.Text)
End If

If Val(FormGETPTC.TextPTCDensity.Text) < 0.1 Or Val(FormGETPTC.TextPTCDensity.Text) > 20# Then
msg$ = "Particle density is out of range (must be between 0.1 and 20)"
MsgBox msg$, vbOKOnly + vbExclamation, "GetPTCSave"
ierror = True
Exit Sub
Else
PTCDensity! = Val(FormGETPTC.TextPTCDensity.Text)
End If

If Val(FormGETPTC.TextPTCThicknessFactor.Text) < 0.001 Or Val(FormGETPTC.TextPTCThicknessFactor.Text) > 1000# Then
msg$ = "Particle thickness factor is out of range (must be between 0.001 and 1000)"
MsgBox msg$, vbOKOnly + vbExclamation, "GetPTCSave"
ierror = True
Exit Sub
Else
PTCThicknessFactor! = Val(FormGETPTC.TextPTCThicknessFactor.Text)
End If

If Val(FormGETPTC.TextPTCNumericalIntegrationStep.Text) < 0.0000001 Or Val(FormGETPTC.TextPTCNumericalIntegrationStep.Text) > 0.001 Then
msg$ = "Particle numerical integration step is out of range (must be between 0.0000001 and 0.001)"
MsgBox msg$, vbOKOnly + vbExclamation, "GetPTCSave"
ierror = True
Exit Sub
Else
PTCNumericalIntegrationStep! = Val(FormGETPTC.TextPTCNumericalIntegrationStep.Text)
End If

If FormGETPTC.CheckPTCDoNotNormalizeSpecifiedFlag.value = vbChecked Then
PTCDoNotNormalizeSpecifiedFlag = True
Else
PTCDoNotNormalizeSpecifiedFlag = False
End If

Call GetPTCCheck
If ierror Then Exit Sub

Exit Sub

' Errors
GetPTCSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "GetPTCSave"
ierror = True
Exit Sub

End Sub

Sub GetPTCCheck()
' Check PTC options

ierror = False
On Error GoTo GetPTCCheckError

Dim i As Integer

' Check that Phi-Rho-Z is selected (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% <> 0 Then
msg$ = "Only ZAF or Phi-rho-z corrections are supported with particle/thin film corrections."
MsgBox msg$, vbOKOnly + vbExclamation, "GetPTCCheck"
ierror = True
Exit Sub
End If

' Check for proper Phi-Rho-Z options for PTC calculations
If izaf% = 0 Then GoTo 1200     ' undocumented option for experts
If izaf% = 1 Then GoTo 1200
If izaf% = 6 Then GoTo 1200
If izaf% = 7 Then GoTo 1200
If izaf% = 8 Then GoTo 1200
msg$ = "Only the following Phi-Rho-Z options are supported for particle and thin film calculations:" & vbCrLf
For i% = 1 To UBound(zafstring$)    ' do not print Individual Selections
If i% = 0 Then msg$ = msg$ & vbCrLf & zafstring$(i%)
If i% = 1 Then msg$ = msg$ & vbCrLf & zafstring$(i%)
If i% = 6 Then msg$ = msg$ & vbCrLf & zafstring$(i%)
If i% = 7 Then msg$ = msg$ & vbCrLf & zafstring$(i%)
If i% = 8 Then msg$ = msg$ & vbCrLf & zafstring$(i%)
Next i%
MsgBox msg$, vbOKOnly + vbExclamation, "GetPTCCheck"
ierror = True
Exit Sub
1200:

' Check for proper abscor options for PTC calculations
If iabs% = 7 Then GoTo 1300
If iabs% = 8 Then GoTo 1300
If iabs% = 9 Then GoTo 1300
If iabs% = 10 Then GoTo 1300
If iabs% = 11 Then GoTo 1300
If iabs% = 14 Then GoTo 1300
If iabs% = 15 Then GoTo 1300
msg$ = "Only the following absorption correction options are supported for particle and thin film calculations:" & vbCrLf
For i% = 1 To UBound(absstring$)
If i% = 7 Then msg$ = msg$ & vbCrLf & absstring$(i%)
If i% = 8 Then msg$ = msg$ & vbCrLf & absstring$(i%)
If i% = 9 Then msg$ = msg$ & vbCrLf & absstring$(i%)
If i% = 10 Then msg$ = msg$ & vbCrLf & absstring$(i%)
If i% = 11 Then msg$ = msg$ & vbCrLf & absstring$(i%)
If i% = 14 Then msg$ = msg$ & vbCrLf & absstring$(i%)
If i% = 15 Then msg$ = msg$ & vbCrLf & absstring$(i%)
Next i%
MsgBox msg$, vbOKOnly + vbExclamation, "GetPTCCheck"
ierror = True
Exit Sub
1300:

Exit Sub

' Errors
GetPTCCheckError:
MsgBox Error$, vbOKOnly + vbCritical, "GetPTCCheck"
ierror = True
Exit Sub

End Sub
