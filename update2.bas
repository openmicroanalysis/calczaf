Attribute VB_Name = "CodeUPDATE2"
' (c) Copyright 1995-2017 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Dim UpdateStdSample(1 To 1) As TypeSample

Sub UpdateAllStdKfacs(method As Integer, analysis As TypeAnalysis, sample() As TypeSample, stdsample() As TypeSample)
' This routine calculates the standard k-factors for all standards. This
' routine is called to make sure that no anomalous conditions exist.
'   method = 0 normal call (check for different element, x-ray, spectro and crystal)
'   method = 1 MAN call (check for different element, x-ray, spectro and crystal *and* keV)

ierror = False
On Error GoTo UpdateAllStdKfacsError

Dim i As Integer

' Calculate standard k-factors for all standards
Call IOWriteLog(vbCrLf & vbCrLf & "Calculating All Standard K-factors...")
For i% = 1 To NumberofStandards%
If VerboseMode Then Call IOWriteLog(vbNullString)
If VerboseMode Then Call IOWriteLog(vbNullString)
msg$ = "Calculating k-factors for standard " & Str$(StandardNumbers%(i%)) & " " & StandardNames$(i%)
Call IOWriteLog(msg$)
DoEvents

' Get standard composition and cations from the standard database and load in Tmp sample arrays
Call UpdateCalculate(method%, i%, StandardNumbers%(i%), analysis, sample(), stdsample())
If ierror Then Exit Sub
Next i%

' Print the standard parameters (after the last standard is calculated)
If DebugMode And (CorrectionFlag% = 0 Or CorrectionFlag% = MAXCORRECTION%) Then
Call ZAFPrintStandards(analysis, sample())
If ierror Then Exit Sub
End If

' 0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters
If CorrectionFlag% = 0 Or CorrectionFlag% = 5 Or CorrectionFlag% = MAXCORRECTION% Then
msg$ = "Standard K-factors Calculated"
Call IOWriteLog(msg$)

Else
msg$ = "Standard Beta-factors Calculated at " & Str$(sample(1).takeoff!) & " takeoff and " & Str$(sample(1).kilovolts!) & " keV"
Call IOWriteLog(msg$)
End If

Exit Sub

' Errors
UpdateAllStdKfacsError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateAllStdKfacs"
ierror = True
Exit Sub

End Sub

Sub UpdateCalculate(method As Integer, row As Integer, num As Integer, analysis As TypeAnalysis, sample() As TypeSample, stdsample() As TypeSample)
' Loads a single standard stdsample array and calculates k-ratios
'   method = 0 normal call (check for different element, x-ray, spectro and crystal)
'   method = 1 MAN call (check for different element, x-ray, spectro and crystal *and* keV)
'   "row" is the standard position in the standard list
'   "num" is the standard number

ierror = False
On Error GoTo UpdateCalculateError

' Update the standard for current sample conditions
Call UpdateCalculateUpdateStandard(method%, num%, sample(), stdsample())
If ierror Then Exit Sub

' Reload the element arrays for the current sample
If sample(1).LastElm% > 0 Then
Call ElementGetData(sample())
If ierror Then Exit Sub

' Calculate the standard k factors for this standard (also calculate ZAF for calibration curve for demo mode)
If CorrectionFlag% = 0 Or CorrectionFlag% = 5 Then
Call ZAFStd(row%, analysis, sample(), stdsample())
If ierror Then Exit Sub
ElseIf CorrectionFlag% = MAXCORRECTION% Then
'Call ZAFStd3(row%, analysis, sample(), stdsample())
'If ierror Then Exit Sub
End If

' Calculate the standard beta factors for this standard (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% > 0 And CorrectionFlag% < 5 Then
Call AFactorStd(row%, analysis, sample(), stdsample())
If ierror Then Exit Sub
End If
End If

Exit Sub

' Errors
UpdateCalculateError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateCalculate"
ierror = True
Exit Sub

End Sub

Sub UpdateStdElements(analysis As TypeAnalysis, sample() As TypeSample, stdsample() As TypeSample)
' Update the not analyzed elements for the standards based on the sample

ierror = False
On Error GoTo UpdateStdElementsError

Dim elementadded As Boolean
Dim i As Integer, j As Integer
Dim ip As Integer, ipp As Integer

If NumberofSamples% = 0 Then Exit Sub
If sample(1).LastChan% = 0 Then Exit Sub

' Check for additional elements if sample is a standard and add to the analyzed element list
If VerboseMode Then
Call IOWriteLog(vbCrLf & "Checking for unanalyzed elements in standard sample...")
End If

' Get data from the standard database if the sample is a standard
elementadded = False
If sample(1).Type% = 1 Then
Call StandardGetMDBStandard(sample(1).number%, stdsample())
If ierror Then Exit Sub

' If element not found, add to the analyzed sample element list
For j% = 1 To stdsample(1).LastChan%
ip% = IPOS1(sample(1).LastChan%, stdsample(1).Elsyms$(j%), sample(1).Elsyms$())

' Add to sample as specified eleemnt if concentration is greater than MinSpecifiedValue
If ip% = 0 Then
If stdsample(1).ElmPercents!(j%) > MinSpecifiedValue! Then
If sample(1).LastChan% + 1 > MAXCHAN% Then GoTo UpdateStdElementsTooManyElements

' Add the new specified element to the sample arrays
elementadded = True
sample(1).LastChan% = sample(1).LastChan% + 1
sample(1).TakeoffArray!(sample(1).LastChan%) = sample(1).takeoff!
sample(1).KilovoltsArray!(sample(1).LastChan%) = sample(1).kilovolts!
sample(1).Elsyms$(sample(1).LastChan%) = stdsample(1).Elsyms$(j%)
sample(1).Xrsyms$(sample(1).LastChan%) = vbNullString
ip% = IPOS1(MAXELM%, sample(1).Elsyms$(sample(1).LastChan%), Symlo$())
If ip% > 0 Then
sample(1).numcat%(sample(1).LastChan%) = AllCat%(ip%)
sample(1).numoxd%(sample(1).LastChan%) = AllOxd%(ip%)
sample(1).AtomicCharges!(sample(1).LastChan%) = AllAtomicCharges!(ip%)
End If

End If
End If
Next j%
End If

' If calculating oxygen for sample, add oxygen if not already present for excess oxygen calculation
Call UpdateAddCalculatedOxygen(elementadded, sample())
If ierror Then Exit Sub

' Load the standard percents for AnalyzeTypeResults "PUBL" field, if sample is a standard
If sample(1).Type% = 1 Then
ip% = IPOS2(NumberofStandards%, sample(1).number%, StandardNumbers%())
If ip% > 0 Then
For i% = 1 To sample(1).LastChan%
analysis.StdPercents!(ip%, i%) = NOT_ANALYZED_VALUE_SINGLE!

' Add to sample (also if using duplicated element)
ipp% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample()) ' find if element is duplicated in sample
If Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ipp% = 0) Then
ipp% = IPOS1(stdsample(1).LastChan%, sample(1).Elsyms$(i%), stdsample(1).Elsyms$())
If ipp% > 0 Then analysis.StdPercents!(ip%, i%) = stdsample(1).ElmPercents!(ipp%)
End If

Next i%
End If
End If

Exit Sub

' Errors
UpdateStdElementsError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateStdElements"
ierror = True
Exit Sub

UpdateStdElementsTooManyElements:
msg$ = "Too many elements in standard number " & Str$(stdsample(1).number%)
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateStdElements"
ierror = True
Exit Sub

End Sub

Sub UpdateStdKfacs(analysis As TypeAnalysis, sample() As TypeSample, stdsample() As TypeSample)
' This routine calculates the standard k-factors for the assigned
' standards and interference standards for the analyzed sample.

ierror = False
On Error GoTo UpdateStdKfacsError

Dim i As Integer, j As Integer
Dim ip As Integer

ReDim standardcalculated(1 To MAXSTD%) As Integer

If NumberofSamples% = 0 Then Exit Sub
If sample(1).LastChan% = 0 Then Exit Sub

' ZAFStd calculation, initialize the standard already calculated flags
For i% = 1 To MAXSTD%
standardcalculated(i%) = False
Next i%

' Calculate standard k-factors for each assigned standard for this sample
If VerboseMode Then Call IOWriteLog("Calculating assigned standard K-factors...")
For i% = 1 To sample(1).LastElm%

' Make sure standards are assigned to channels
If sample(1).DisableQuantFlag(i%) = 0 Then
If sample(1).StdAssigns%(i%) = 0 Then GoTo UpdateStdKfacsNoStandard

' Find position of standard in standard list and check that it was not already calculated
ip% = IPOS2(NumberofStandards%, sample(1).StdAssigns%(i%), StandardNumbers%())
If ip% = 0 Then GoTo UpdateStdKfacsBadStd
If Not standardcalculated(ip%) Then

' Calculate the standard k ratios
If CorrectionFlag% = 0 Or CorrectionFlag% = 5 Or CorrectionFlag% = MAXCORRECTION% Then
Call AnalyzeStatusAnal("Calculating standard k-factors for standard " & Format$(sample(1).StdAssigns%(i%)) & "...")
End If
If CorrectionFlag% > 0 And CorrectionFlag% < 5 Then
Call AnalyzeStatusAnal("Calculating standard beta-factors for standard " & Format$(sample(1).StdAssigns%(i%)) & "...")
End If
Call UpdateCalculate(Int(0), ip%, sample(1).StdAssigns%(i%), analysis, sample(), stdsample())
If ierror Then Exit Sub

If VerboseMode% Then
msg$ = vbCrLf & "Standard " & Str$(sample(1).StdAssigns%(i%)) & ", (after UpdateCalculate):"
Call IOWriteLog(msg$)
For j% = 1 To sample(1).LastElm%

ip% = IPOS2(NumberofStandards%, sample(1).StdAssigns%(i%), StandardNumbers%())
If ip% > 0 Then
msg$ = Str$(j%) & " " & sample(1).Elsyms$(j%) & " " & sample(1).Xrsyms$(j%) & MiscAutoFormat$(analysis.StdPercents!(ip%, j%))
Call IOWriteLog(msg$)
End If

Next j%
End If

standardcalculated(ip%) = True
End If
End If
Next i%


' Calculate standard k-factors for each assigned interference standard for this sample
If VerboseMode Then Call IOWriteLog("Calculating assigned interference standard K-factors...")
For i% = 1 To sample(1).LastElm%
For j% = 1 To MAXINTF%

' Make sure interference standard is assigned for this position
If sample(1).StdAssignsIntfStds%(j%, i%) > 0 Then

' Find position of standard in standard list and check that it was not already calculated
ip% = IPOS2(NumberofStandards%, sample(1).StdAssignsIntfStds%(j%, i%), StandardNumbers%())
If ip% = 0 Then GoTo UpdateStdKfacsBadStdInterf
If Not standardcalculated(ip%) Then

' Calculate the standard k ratios
Call UpdateCalculate(Int(0), ip%, sample(1).StdAssignsIntfStds%(j%, i%), analysis, sample(), stdsample())
If ierror Then Exit Sub

standardcalculated(ip%) = True
End If
End If

Next j%
DoEvents
Next i%


' Calculate standard k-factors for each assigned MAN standard for this sample (in case used for any standard)
If VerboseMode Then Call IOWriteLog("Calculating assigned MAN standard K-factors...")
For i% = 1 To sample(1).LastElm%
For j% = 1 To MAXMAN%

' Make sure MAN standard is assigned for this position
If sample(1).MANStdAssigns%(j%, i%) > 0 Then

' Find position of standard in standard list and check that it was not already calculated
ip% = IPOS2(NumberofStandards%, sample(1).MANStdAssigns%(j%, i%), StandardNumbers%())
If ip% = 0 Then GoTo UpdateStdKfacsBadMANStd
If Not standardcalculated(ip%) Then

' Calculate the standard k ratios
Call UpdateCalculate(Int(0), ip%, sample(1).MANStdAssigns%(j%, i%), analysis, sample(), stdsample())
If ierror Then Exit Sub

standardcalculated(ip%) = True
End If
End If

Next j%
DoEvents
Next i%


' Check that the assigned standards contain sufficient concentration of the assigned element
For i% = 1 To sample(1).LastElm%
If sample(1).DisableQuantFlag(i%) = 0 Then

ip% = IPOS2(NumberofStandards%, sample(1).StdAssigns%(i%), StandardNumbers%())
If ip% = 0 Then GoTo UpdateStdKfacsBadStd

' Skip if disabled quant flag (add check for UseAggregateIntensitiesFlag 3/15/2006)
If Not UseAggregateIntensitiesFlag Then
If ip% > 0 Then If analysis.StdPercents!(ip%, i%) < StdMinimumValue! Then GoTo UpdateStdKfacsBadPercent

Else
j% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())
If j% = 0 Then
If analysis.StdPercents!(ip%, i%) < StdMinimumValue! Then GoTo UpdateStdKfacsBadPercent
End If
End If

End If
Next i%

' Print the standard parameters (after the last standard is calculated)
If DebugMode And (CorrectionFlag% = 0 Or CorrectionFlag% = MAXCORRECTION%) Then
Call ZAFPrintStandards(analysis, sample())
If ierror Then Exit Sub
End If

Call AnalyzeStatusAnal(vbNullString)
Exit Sub

' Errors
UpdateStdKfacsError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateStdKfacs"
ierror = True
Exit Sub

UpdateStdKfacsNoStandard:
msg$ = TypeLoadString$(sample())
msg$ = "No standard assigned for " & sample(1).Elsyms$(i%) & " in sample " & msg$ & ". " & vbCrLf & vbCrLf
msg$ = msg$ & "Make sure that all elements are assigned a standard by first clicking the Standard "
msg$ = msg$ & "Assignments button in the Analyze! window and confirming the default standard assignments." & vbCrLf & vbCrLf
msg$ = msg$ & "Note also, that because new samples are created based on the last unknown sample, the "
msg$ = msg$ & "samples selected for standard assignments should generally include the last unknown sample."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateStdKfacs"
ierror = True
Exit Sub

UpdateStdKfacsBadStd:
msg$ = "Standard number " & Str$(sample(1).StdAssigns%(i%)) & ", the assigned standard for " & sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & ", does not exist in this run. " & vbCrLf & vbCrLf
msg$ = msg$ & "Change the standard assignment or use the Add Standards to "
msg$ = msg$ & "Run menu under the Standard menu to add the standard to the run "
msg$ = msg$ & "and acquire data for it."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateStdKfacs"
ierror = True
Exit Sub

UpdateStdKfacsBadStdInterf:
msg$ = "Standard number " & Str$(sample(1).StdAssignsIntfStds%(j%, i%)) & ", the assigned interference standard for " & sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & ", does not exist in this run. " & vbCrLf & vbCrLf
msg$ = msg$ & "Change the standard interference assignment or use the Add Standards to "
msg$ = msg$ & "Run menu under the Standard menu to add the standard to the run "
msg$ = msg$ & "and acquire data for it."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateStdKfacs"
ierror = True
Exit Sub

UpdateStdKfacsBadMANStd:
msg$ = "MAN Standard number " & Str$(sample(1).MANStdAssigns%(j%, i%)) & ", an assigned MAN standard for " & sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & ", does not exist in this run. " & vbCrLf & vbCrLf
msg$ = msg$ & "Change the MAN standard assignments under the Analytical | MAN Fits menu, or use the Add Standards to "
msg$ = msg$ & "Run menu under the Standard menu to add the MAN standard to the run "
msg$ = msg$ & "and acquire data for it."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateStdKfacs"
ierror = True
Exit Sub

UpdateStdKfacsBadPercent:
msg$ = "Insufficient assigned standard percent on " & sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & " channel " & Str$(i%) & ", "
msg$ = msg$ & " in standard " & Str$(StandardNumbers(ip%)) & " " & StandardNames$(ip%) & "." & vbCrLf & vbCrLf
msg$ = msg$ & "Try assigning a different standard to sample " & Str$(sample(1).number%) & " " & sample(1).Name$ & " for the analysis "
msg$ = msg$ & "of " & sample(1).Elsyms$(i%) & " using the Standard Assigments button in the ANALYZE! window."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateStdKfacs"
ierror = True
Exit Sub

End Sub

Sub UpdateCalculateUpdateStandard(method As Integer, num As Integer, sample() As TypeSample, stdsample() As TypeSample)
' Update the standard composition for the current sample conditions
'   method = 0 normal call (check for different element, x-ray, spectro and crystal)
'   method = 1 MAN call (check for different element, x-ray, spectro and crystal *and* keV)

ierror = False
On Error GoTo UpdateCalculateUpdateStandardError

Dim i As Integer, j As Integer, ip As Integer

' Get standard composition and cations from the standard database and load in stdsample arrays
If num% > 0 Then
Call StandardGetMDBStandard(num%, stdsample())
If ierror Then Exit Sub

' Load standard as unknown composition if stdnum is zero (random composition)
Else
Call StandardGetMDBStandard(num%, sample())
If ierror Then Exit Sub
stdsample(1) = sample(1)
End If

' Save standard composition for subsequent loading of concentrations for std k-factors
UpdateStdSample(1) = stdsample(1)

If VerboseMode% Then
msg$ = vbCrLf & "Standard " & Str$(stdsample(1).number%) & ", (before standard update):"
Call IOWriteLog(msg$)
For i% = 1 To stdsample(1).LastChan%
msg$ = Str$(i%) & " " & stdsample(1).Elsyms$(i%) & " " & stdsample(1).Xrsyms$(i%) & Str$(stdsample(1).MotorNumbers%(i%)) & " " & stdsample(1).CrystalNames$(i%)
Call IOWriteLog(msg$)
Next i%
End If

' Load analyzed sample conditions for standard k factor calculation
stdsample(1).EDSSpectraFlag% = sample(1).EDSSpectraFlag%            ' for no WDS elements
For j% = 1 To stdsample(1).LastChan%
ip% = IPOS1(sample(1).LastElm%, stdsample(1).Elsyms$(j%), sample(1).Elsyms$()) ' only check element
If ip% > 0 Then
stdsample(1).TakeoffArray!(j%) = sample(1).TakeoffArray!(ip%)
stdsample(1).KilovoltsArray!(j%) = sample(1).KilovoltsArray!(ip%)
Else
stdsample(1).Xrsyms$(j%) = vbNullString   ' not used in sample
stdsample(1).MotorNumbers%(j%) = 0
stdsample(1).CrystalNames$(j%) = vbNullString
End If
Next j%

' Load x-ray lines (and motor and crystal) for elements present in analyzed sample for standard k-factor calculation
For j% = 1 To stdsample(1).LastChan%
ip% = IPOS9a(stdsample(1).Elsyms$(j%), sample())  ' changed 09/28/2012 for Pt La/Ma disable quant issue, IPOS9 checks disable flag, IPOS9a does NOT check disable flag
If ip% > 0 And ip% <= sample(1).LastElm% And stdsample(1).MotorNumbers%(j%) = 0 Then
stdsample(1).Xrsyms$(j%) = sample(1).Xrsyms$(ip%)
stdsample(1).MotorNumbers%(j%) = sample(1).MotorNumbers%(ip%)
stdsample(1).CrystalNames$(j%) = sample(1).CrystalNames$(ip%)
End If
Next j%

If VerboseMode% Then
msg$ = vbCrLf & "Standard " & Str$(stdsample(1).number%) & ", (after standard update for same element as unknown):"
Call IOWriteLog(msg$)
For i% = 1 To stdsample(1).LastChan%
msg$ = Str$(i%) & " " & stdsample(1).Elsyms$(i%) & " " & stdsample(1).Xrsyms$(i%) & Str$(stdsample(1).MotorNumbers%(i%)) & " " & stdsample(1).CrystalNames$(i%)
Call IOWriteLog(msg$)
Next i%
End If

' Now check for analyzed elements in the sample that are not in the standard, and load them (for the quantitative interference correction calculation).
For j% = 1 To sample(1).LastElm%
ip% = IPOS1(stdsample(1).LastChan%, sample(1).Elsyms$(j%), stdsample(1).Elsyms$()) ' only check element
If ip% = 0 Then
If stdsample(1).LastChan% + 1 > MAXCHAN% Then GoTo UpdateCalculateUpdateStandardTooManyElements

' Add the analyzed element to the standard element arrays
stdsample(1).LastChan% = stdsample(1).LastChan% + 1
stdsample(1).TakeoffArray!(stdsample(1).LastChan%) = sample(1).TakeoffArray!(j%)
stdsample(1).KilovoltsArray!(stdsample(1).LastChan%) = sample(1).KilovoltsArray!(j%)
stdsample(1).Elsyms$(stdsample(1).LastChan%) = sample(1).Elsyms$(j%)
stdsample(1).Xrsyms$(stdsample(1).LastChan%) = sample(1).Xrsyms$(j%)
stdsample(1).MotorNumbers%(stdsample(1).LastChan%) = sample(1).MotorNumbers%(j%)
stdsample(1).CrystalNames$(stdsample(1).LastChan%) = sample(1).CrystalNames$(j%)
stdsample(1).numcat%(stdsample(1).LastChan%) = sample(1).numcat%(j%)
stdsample(1).numoxd%(stdsample(1).LastChan%) = sample(1).numoxd%(j%)
stdsample(1).AtomicCharges!(stdsample(1).LastChan%) = sample(1).AtomicCharges!(j%)
stdsample(1).ElmPercents!(stdsample(1).LastChan%) = NOT_ANALYZED_VALUE_SINGLE!  ' use a non-zero value
stdsample(1).LastElm% = stdsample(1).LastChan%
End If
Next j%

If VerboseMode% Then
msg$ = vbCrLf & "Standard " & Str$(stdsample(1).number%) & ", (after standard update for new unknown element):"
Call IOWriteLog(msg$)
For i% = 1 To stdsample(1).LastChan%
msg$ = Str$(i%) & " " & stdsample(1).Elsyms$(i%) & " " & stdsample(1).Xrsyms$(i%) & Str$(stdsample(1).MotorNumbers%(i%)) & " " & stdsample(1).CrystalNames$(i%)
Call IOWriteLog(msg$)
Next i%
End If

' Add in duplicate elements with different x-ray, motor, crystal (composition will total over 100% unless quant disabled or aggregate mode) (add check for different keV, for call from MAN dialog, 08-17-2017)
For j% = 1 To sample(1).LastElm%
If method% = 0 Then ip% = IPOS5(Int(1), j%, sample(), stdsample())  ' find position of sample element in std
If method% = 1 Then ip% = IPOS13B(Int(1), sample(1).Elsyms$(j%), sample(1).Xrsyms$(j%), sample(1).MotorNumbers%(j%), sample(1).CrystalNames$(j%), sample(1).KilovoltsArray!(j%), stdsample())
If ip% = 0 Then
If stdsample(1).LastChan% + 1 > MAXCHAN% Then GoTo UpdateCalculateUpdateStandardTooManyElements

' Add the analyzed element to the standard element arrays
stdsample(1).LastChan% = stdsample(1).LastChan% + 1
stdsample(1).TakeoffArray!(stdsample(1).LastChan%) = sample(1).TakeoffArray!(j%)
stdsample(1).KilovoltsArray!(stdsample(1).LastChan%) = sample(1).KilovoltsArray!(j%)
stdsample(1).Elsyms$(stdsample(1).LastChan%) = sample(1).Elsyms$(j%)
stdsample(1).Xrsyms$(stdsample(1).LastChan%) = sample(1).Xrsyms$(j%)
stdsample(1).MotorNumbers%(stdsample(1).LastChan%) = sample(1).MotorNumbers%(j%)
stdsample(1).CrystalNames$(stdsample(1).LastChan%) = sample(1).CrystalNames$(j%)
stdsample(1).numcat%(stdsample(1).LastChan%) = sample(1).numcat%(j%)
stdsample(1).numoxd%(stdsample(1).LastChan%) = sample(1).numoxd%(j%)
stdsample(1).AtomicCharges!(stdsample(1).LastChan%) = sample(1).AtomicCharges!(j%)
stdsample(1).ElmPercents!(stdsample(1).LastChan%) = NOT_ANALYZED_VALUE_SINGLE!  ' use a non-zero value
stdsample(1).LastElm% = stdsample(1).LastChan%
End If
Next j%

If VerboseMode% Then
msg$ = vbCrLf & "Standard " & Str$(stdsample(1).number%) & ", (after standard update for duplicate element and concentration):"
Call IOWriteLog(msg$)
For i% = 1 To stdsample(1).LastChan%
msg$ = Str$(i%) & " " & stdsample(1).Elsyms$(i%) & " " & stdsample(1).Xrsyms$(i%) & Str$(stdsample(1).MotorNumbers%(i%)) & " " & stdsample(1).CrystalNames$(i%)
Call IOWriteLog(msg$)
Next i%
End If

' Check for disable quant flag from unknown for analyzed elements based on element and x-ray
For j% = 1 To sample(1).LastElm%
ip% = IPOS5(Int(1), j%, sample(), stdsample())
If ip% > 0 Then
stdsample(1).DisableQuantFlag%(ip%) = sample(1).DisableQuantFlag%(j%)
End If
Next j%

If VerboseMode% Then
msg$ = vbCrLf & "Standard " & Str$(stdsample(1).number%) & ", (after standard update for disable quant):"
Call IOWriteLog(msg$)
For i% = 1 To stdsample(1).LastChan%
msg$ = Str$(i%) & " " & stdsample(1).Elsyms$(i%) & " " & stdsample(1).Xrsyms$(i%) & " " & Format$(stdsample(1).MotorNumbers%(i%)) & " " & stdsample(1).CrystalNames$(i%) & " " & Format$(stdsample(1).DisableQuantFlag%(i%))
Call IOWriteLog(msg$)
Next i%
End If

' Finally add in standard concentrations for standard k-factor calculations (added 04/15/2014, see also recent changes in ZAFStd)
For j% = 1 To stdsample(1).LastChan%
ip% = IPOS1(UpdateStdSample(1).LastElm%, stdsample(1).Elsyms$(j%), UpdateStdSample(1).Elsyms$()) ' only check first occurance of element in standard
If ip% > 0 Then
stdsample(1).ElmPercents!(j%) = UpdateStdSample(1).ElmPercents!(ip%)    ' always use previously stored standard concentrations to calculate std k-factors properly (changed 08/28/2014, Buse)
End If
Next j%

If VerboseMode% Then
msg$ = vbCrLf & "Standard " & Str$(stdsample(1).number%) & ", (after standard update for standard percents for std k-factors):"
Call IOWriteLog(msg$)
For i% = 1 To stdsample(1).LastChan%
msg$ = Str$(i%) & " " & stdsample(1).Elsyms$(i%) & " " & stdsample(1).Xrsyms$(i%) & " " & Format$(stdsample(1).MotorNumbers%(i%)) & " " & stdsample(1).CrystalNames$(i%) & " " & Format$(stdsample(1).ElmPercents!(i%)) & " " & Format$(stdsample(1).DisableQuantFlag%(i%))
Call IOWriteLog(msg$)
Next i%
End If

' Re-order standard sample in case x-ray lines were changed to unanalyzed
Call GetElmSaveSampleOnly(method%, stdsample(), Int(0), Int(0))
If ierror Then Exit Sub

If VerboseMode% Then
msg$ = vbCrLf & "Standard " & Str$(stdsample(1).number%) & ", (after sort):"
Call IOWriteLog(msg$)
For i% = 1 To stdsample(1).LastChan%
msg$ = Str$(i%) & " " & stdsample(1).Elsyms$(i%) & " " & stdsample(1).Xrsyms$(i%) & Str$(stdsample(1).MotorNumbers%(i%)) & " " & stdsample(1).CrystalNames$(i%) & " " & Str$(stdsample(1).ElmPercents!(i%)) & " " & Str$(stdsample(1).DisableQuantFlag%(i%))
Call IOWriteLog(msg$)
Next i%
Call IOWriteLog(vbNullString)
End If

Exit Sub

' Errors
UpdateCalculateUpdateStandardError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateCalculateUpdateStandard"
ierror = True
Exit Sub

UpdateCalculateUpdateStandardTooManyElements:
msg$ = "Too many elements in standard number " & Str$(stdsample(1).number%)
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateCalculateUpdateStandard"
ierror = True
Exit Sub

End Sub
