Attribute VB_Name = "CodeUPDATE2"
' (c) Copyright 1995-2025 by John J. Donovan
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

Sub UpdateAllStdKfacs(analysis As TypeAnalysis, sample() As TypeSample, stdsample() As TypeSample)
' This routine calculates the standard k-factors for all standards. This
' routine is called to make sure that no anomalous conditions exist.

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
Call UpdateCalculate(i%, StandardNumbers%(i%), analysis, sample(), stdsample())
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
msg$ = "Standard Beta-factors Calculated"
Call IOWriteLog(msg$)
End If

Exit Sub

' Errors
UpdateAllStdKfacsError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateAllStdKfacs"
ierror = True
Exit Sub

End Sub

Sub UpdateCalculate(row As Integer, num As Integer, analysis As TypeAnalysis, sample() As TypeSample, stdsample() As TypeSample)
' Loads a single standard stdsample array and calculates k-ratios
'   "row" is the standard position in the standard list
'   "num" is the standard number

ierror = False
On Error GoTo UpdateCalculateError

' New std k-factor calculation  code for samples with duplicate elements (with different x-ray lines or different kilovolts!)
If MiscIsElementDuplicatedAndXrayOrKilovoltsDifferent(sample()) Then
Call UpdateCalculate2(row%, num%, analysis, sample(), stdsample)
If ierror Then Exit Sub
Exit Sub
End If

' Update the standard for current sample conditions
Call UpdateCalculateUpdateStandard2(num%, sample(), stdsample())
If ierror Then Exit Sub

' Reload the element arrays for the current sample
If sample(1).LastElm% > 0 Then
Call ElementGetData(sample())
If ierror Then Exit Sub

' Calculate the standard k factors for this standard (also calculate ZAF for calibration curve for demo mode)
If CorrectionFlag% = 0 Or CorrectionFlag% = 5 Then
Call ZAFStd2(row%, analysis, sample(), stdsample())
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

Sub UpdateCalculate2(row As Integer, num As Integer, analysis As TypeAnalysis, sample() As TypeSample, stdsample() As TypeSample)
' Loads a single standard stdsample array and calculates k-ratios (version to handle duplicate elements)
'   "row" is the standard position in the standard list
'   "num" is the standard number

ierror = False
On Error GoTo UpdateCalculate2Error

Dim chan As Integer, i As Integer, j As Integer, ip As Integer

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

' Loop on each analyzed element in original sample and load standard composition
For chan% = 1 To sample(1).LastElm%

' Skip disable quant element
If sample(1).DisableQuantFlag%(chan%) = 0 Then

' Skip duplicate element
ip% = IPOS8A(chan%, sample(1).Elsyms$(chan%), sample(1).Xrsyms$(chan%), sample(1).KilovoltsArray!(chan%), sample()) ' find if element is duplicated
If Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0) Then

' Loop on each analyzed element in original sample and calculate standard k-factors *one element at a time*
If VerboseMode% Then
msg$ = vbCrLf & vbCrLf & "Calculating standard k-factors for " & Str$(chan%) & " " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & ", " & Str$(sample(1).MotorNumbers%(chan%)) & " " & sample(1).CrystalNames$(chan%) & ", " & Format$(sample(1).TakeoffArray!(chan%)) & " " & Format$(sample(1).KilovoltsArray!(chan%))
Call IOWriteLog(msg$)
End If

' Load passed sample into temp std sample
UpdateStdSample(1) = sample(1)

' Now zero out concentrations (in case any are loaded) and specifiy as fixed concentrations
For i% = 1 To UpdateStdSample(1).LastChan%
UpdateStdSample(1).ElmPercents!(i%) = NOT_ANALYZED_VALUE_SINGLE!                                 ' use a non-zero value
Next i%

If VerboseMode% Then
msg$ = vbCrLf & "Standard " & Format$(num%) & ", (original sample):"
Call IOWriteLog(msg$)
For i% = 1 To UpdateStdSample(1).LastChan%
msg$ = Str$(i%) & " " & UpdateStdSample(1).Elsyms$(i%) & " " & UpdateStdSample(1).Xrsyms$(i%) & Str$(UpdateStdSample(1).MotorNumbers%(i%)) & " " & UpdateStdSample(1).CrystalNames$(i%) & ", " & Format$(UpdateStdSample(1).TakeoffArray!(i%)) & " " & Format$(UpdateStdSample(1).KilovoltsArray!(i%))
Call IOWriteLog(msg$)
Next i%
End If

' Now check for elements in the standard that are not in the sample, and add them if necesssary as specified elements
For j% = 1 To stdsample(1).LastChan%
'ip% = IPOS1(UpdateStdSample(1).LastChan%, stdsample(1).Elsyms$(j%), UpdateStdSample(1).Elsyms$())         ' only check element
ip% = IPOS9(stdsample(1).Elsyms$(j%), UpdateStdSample())                                         ' only check first not disabled occurance of element to load concentration
If ip% = 0 Then
If UpdateStdSample(1).LastChan% + 1 > MAXCHAN% Then GoTo UpdateCalculate2TooManyElements

' Add the standard element to the sample element arrays as a specified element
UpdateStdSample(1).LastChan% = UpdateStdSample(1).LastChan% + 1
UpdateStdSample(1).Elsyms$(UpdateStdSample(1).LastChan%) = stdsample(1).Elsyms$(j%)
UpdateStdSample(1).Xrsyms$(UpdateStdSample(1).LastChan%) = vbNullString
UpdateStdSample(1).MotorNumbers%(UpdateStdSample(1).LastChan%) = 0
UpdateStdSample(1).CrystalNames$(UpdateStdSample(1).LastChan%) = vbNullString
UpdateStdSample(1).numcat%(UpdateStdSample(1).LastChan%) = stdsample(1).numcat%(j%)
UpdateStdSample(1).numoxd%(UpdateStdSample(1).LastChan%) = stdsample(1).numoxd%(j%)

UpdateStdSample(1).AtomicCharges!(UpdateStdSample(1).LastChan%) = stdsample(1).AtomicCharges!(j%)
UpdateStdSample(1).AtomicWts!(UpdateStdSample(1).LastChan%) = stdsample(1).AtomicWts!(j%)

UpdateStdSample(1).TakeoffArray!(UpdateStdSample(1).LastChan%) = 0#
UpdateStdSample(1).KilovoltsArray!(UpdateStdSample(1).LastChan%) = 0#

' Do update atomic weights from standard database (in case standard element is enriched) (v. 13.3.2)
Else
UpdateStdSample(1).numcat%(ip%) = stdsample(1).numcat%(j%)     ' element is already present, update cations/oxygens from standard database?????
UpdateStdSample(1).numoxd%(ip%) = stdsample(1).numoxd%(j%)

UpdateStdSample(1).AtomicCharges!(ip%) = stdsample(1).AtomicCharges!(j%)     ' element is already present, update atomic charges from standard database
UpdateStdSample(1).AtomicWts!(ip%) = stdsample(1).AtomicWts!(j%)             ' element is already present, update atomic weights from standard database
End If
Next j%

' Now make all elements specified
For i% = 1 To UpdateStdSample(1).LastChan%
UpdateStdSample(1).Xrsyms$(i%) = vbNullString
UpdateStdSample(1).MotorNumbers%(i%) = 0
UpdateStdSample(1).CrystalNames$(i%) = vbNullString
UpdateStdSample(1).TakeoffArray!(i%) = 0#
UpdateStdSample(1).KilovoltsArray!(i%) = 0#
Next i%

If VerboseMode% Then
msg$ = vbCrLf & "Standard " & Format$(num%) & ", (before update for standard composition):"
Call IOWriteLog(msg$)
For i% = 1 To UpdateStdSample(1).LastChan%
msg$ = Str$(i%) & " " & UpdateStdSample(1).Elsyms$(i%) & " " & UpdateStdSample(1).Xrsyms$(i%) & Str$(UpdateStdSample(1).MotorNumbers%(i%)) & " " & UpdateStdSample(1).CrystalNames$(i%) & ", " & Format$(UpdateStdSample(1).TakeoffArray!(i%)) & " " & Format$(UpdateStdSample(1).KilovoltsArray!(i%))
Call IOWriteLog(msg$)
Next i%
End If

' Now load standard concentrations for all elements other than the current element
For j% = 1 To stdsample(1).LastChan%
If UCase$(stdsample(1).Elsyms$(j%)) <> UCase$(sample(1).Elsyms$(chan%)) Then                       ' skip loading concentration for current element (it might be duplicated)
ip% = IPOS1(UpdateStdSample(1).LastChan%, stdsample(1).Elsyms$(j%), UpdateStdSample(1).Elsyms$())  ' load first occurrance of all elements in standard
If ip% > 0 Then
UpdateStdSample(1).ElmPercents!(ip%) = stdsample(1).ElmPercents!(j%)                               ' load concentration from standard database
End If
End If
Next j%

' Now load the concentration for the current element in the current sample for ZAFStd calculations
ip% = IPOS1(stdsample(1).LastChan%, sample(1).Elsyms$(chan%), stdsample(1).Elsyms$())
If ip% > 0 Then
UpdateStdSample(1).ElmPercents!(chan%) = stdsample(1).ElmPercents!(ip%)                            ' load concentration from standard database if present in standard composition
End If
UpdateStdSample(1).Xrsyms$(chan%) = sample(1).Xrsyms$(chan%)                                       ' load original x-ray line (still in original sample order)
UpdateStdSample(1).MotorNumbers%(chan%) = sample(1).MotorNumbers%(chan%)                           ' load original sample spectrometer number (still in original sample order)
UpdateStdSample(1).CrystalNames$(chan%) = sample(1).CrystalNames$(chan%)                           ' load original sample crystal name (still in original sample order)
UpdateStdSample(1).TakeoffArray!(chan%) = sample(1).TakeoffArray!(chan%)                           ' load original sample takeoff angle (still in original sample order)
UpdateStdSample(1).KilovoltsArray!(chan%) = sample(1).KilovoltsArray!(chan%)                       ' load original sample kilovolts (still in original sample order)

If VerboseMode% Then
msg$ = vbCrLf & "Standard " & Format$(num%) & ", (after update for standard concentrations):"
Call IOWriteLog(msg$)
For i% = 1 To UpdateStdSample(1).LastChan%
msg$ = Str$(i%) & " " & UpdateStdSample(1).Elsyms$(i%) & " " & UpdateStdSample(1).Xrsyms$(i%) & Str$(UpdateStdSample(1).MotorNumbers%(i%)) & " " & UpdateStdSample(1).CrystalNames$(i%) & ", " & Format$(UpdateStdSample(1).TakeoffArray!(i%)) & " " & Format$(UpdateStdSample(1).KilovoltsArray!(i%)) & ", " & Format$(UpdateStdSample(1).ElmPercents!(i%))
Call IOWriteLog(msg$)
Next i%
End If

' Re-order standard sample because all except one x-ray line was changed to unanalyzed
Call GetElmSaveSampleOnly(Int(1), UpdateStdSample(), Int(0), Int(0))
If ierror Then Exit Sub

If VerboseMode% Then
msg$ = vbCrLf & "Standard " & Str$(num%) & ", (after sort):"
Call IOWriteLog(msg$)
For i% = 1 To UpdateStdSample(1).LastChan%
msg$ = Str$(i%) & " " & UpdateStdSample(1).Elsyms$(i%) & " " & UpdateStdSample(1).Xrsyms$(i%) & Str$(UpdateStdSample(1).MotorNumbers%(i%)) & " " & UpdateStdSample(1).CrystalNames$(i%) & ", " & Format$(UpdateStdSample(1).TakeoffArray!(i%)) & " " & Format$(UpdateStdSample(1).KilovoltsArray!(i%)) & ", " & Format$(UpdateStdSample(1).ElmPercents!(i%))
Call IOWriteLog(msg$)
Next i%
Call IOWriteLog(vbNullString)
End If

' Update some standard sample parameters for ZAFStd calculation
UpdateStdSample(1).number% = num%
UpdateStdSample(1).OxideOrElemental% = 2      ' always calculate standard k-factors as elemental

' Reload the element arrays for the current sample
If UpdateStdSample(1).LastElm% > 0 Then
Call ElementGetData(UpdateStdSample())
If ierror Then Exit Sub

' Calculate the standard k factors for this standard (also calculate ZAF for calibration curve for demo mode)
If CorrectionFlag% = 0 Or CorrectionFlag% = 5 Then
Call ZAFStd2(row%, analysis, sample(), UpdateStdSample())
If ierror Then Exit Sub
ElseIf CorrectionFlag% = MAXCORRECTION% Then
'Call ZAFStd3(row%, analysis, sample(), UpdateStdSample())
'If ierror Then Exit Sub
End If

' Calculate the standard beta factors for this standard (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% > 0 And CorrectionFlag% < 5 Then
Call AFactorStd(row%, analysis, sample(), UpdateStdSample())
If ierror Then Exit Sub
End If
End If

' Print out StdZAFCors for MAN dialog (continuum absorption correction)
If DebugMode And VerboseMode Then
Call IOWriteLog("UpdateCalculate2: Standard " & Format$(num%) & " " & StandardNames$(row%) & ", " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & ", " & Format$(sample(1).TakeoffArray!(chan%)) & " " & Format$(sample(1).KilovoltsArray!(chan%)))
For i% = 1 To sample(1).LastElm%
ip% = IPOS14(i%, sample(), UpdateStdSample())  ' check element, xray, take-off and kilovolts (only one element will qualify at a time)
If ip% > 0 Then
msg$ = UpdateStdSample(1).Elsyms$(ip%) & " " & UpdateStdSample(1).Xrsyms$(ip%) & ", " & Format$(UpdateStdSample(1).TakeoffArray!(ip%)) & " " & Format$(UpdateStdSample(1).KilovoltsArray!(ip%)) & ", " & Format$(UpdateStdSample(1).ElmPercents!(ip%)) & ", " & Format$(analysis.StdZAFCors!(1, row%, ip%)) & ", " & Format$(analysis.StdContinuumCorrections!(row%, ip%))
Call IOWriteLog(msg$)
End If
Next i%
End If

' Re-set current element to specified
UpdateStdSample(1).ElmPercents!(chan%) = NOT_ANALYZED_VALUE_SINGLE!
UpdateStdSample(1).Xrsyms$(chan%) = vbNullString
UpdateStdSample(1).TakeoffArray!(chan%) = 0#
UpdateStdSample(1).KilovoltsArray!(chan%) = 0#

' Calculate std k-factors for next original sample element
End If
End If
Next chan%

Exit Sub

' Errors
UpdateCalculate2Error:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateCalculate2"
ierror = True
Exit Sub

UpdateCalculate2TooManyElements:
msg$ = "Too many elements in standard number " & Str$(stdsample(1).number%)
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateCalculate2"
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

' Add to sample as specified element if concentration is greater than MinSpecifiedValue
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
sample(1).AtomicWts!(sample(1).LastChan%) = AllAtomicWts!(ip%)
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
Call UpdateCalculate(ip%, sample(1).StdAssigns%(i%), analysis, sample(), stdsample())
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
Call UpdateCalculate(ip%, sample(1).StdAssignsIntfStds%(j%, i%), analysis, sample(), stdsample())
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
Call UpdateCalculate(ip%, sample(1).MANStdAssigns%(j%, i%), analysis, sample(), stdsample())
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

Sub UpdateCalculateUpdateStandard2(num As Integer, sample() As TypeSample, stdsample() As TypeSample)
' Update the standard composition for the current sample conditions (simplified code, 10/27/2017)

ierror = False
On Error GoTo UpdateCalculateUpdateStandard2Error

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

' Load passed sample into temp std sample
UpdateStdSample(1) = sample(1)

' Now check for elements in the standard that are not in the sample, and add them if necesssary as specified elements
For j% = 1 To stdsample(1).LastChan%
'ip% = IPOS1(UpdateStdSample(1).LastChan%, stdsample(1).Elsyms$(j%), UpdateStdSample(1).Elsyms$())         ' only check element
ip% = IPOS9(stdsample(1).Elsyms$(j%), UpdateStdSample())                                         ' only check first not disabled occurance of element to load concentration
If ip% = 0 Then
If UpdateStdSample(1).LastChan% + 1 > MAXCHAN% Then GoTo UpdateCalculateUpdateStandard2TooManyElements

' Add the standard element to the sample element arrays as a specified element
UpdateStdSample(1).LastChan% = UpdateStdSample(1).LastChan% + 1
UpdateStdSample(1).Elsyms$(UpdateStdSample(1).LastChan%) = stdsample(1).Elsyms$(j%)
UpdateStdSample(1).Xrsyms$(UpdateStdSample(1).LastChan%) = vbNullString
UpdateStdSample(1).numcat%(UpdateStdSample(1).LastChan%) = stdsample(1).numcat%(j%)
UpdateStdSample(1).numoxd%(UpdateStdSample(1).LastChan%) = stdsample(1).numoxd%(j%)

UpdateStdSample(1).AtomicCharges!(UpdateStdSample(1).LastChan%) = stdsample(1).AtomicCharges!(j%)
UpdateStdSample(1).AtomicWts!(UpdateStdSample(1).LastChan%) = stdsample(1).AtomicWts!(j%)

' Do update atomic weights from standard database (in case standard element is enriched) (v. 13.3.2)
Else
UpdateStdSample(1).numcat%(ip%) = stdsample(1).numcat%(j%)     ' element is already present, update cations/oxygens from standard database?????
UpdateStdSample(1).numoxd%(ip%) = stdsample(1).numoxd%(j%)

UpdateStdSample(1).AtomicCharges!(ip%) = stdsample(1).AtomicCharges!(j%)     ' element is already present, update atomic charges from standard database
UpdateStdSample(1).AtomicWts!(ip%) = stdsample(1).AtomicWts!(j%)             ' element is already present, update atomic weights from standard database
End If
Next j%

If VerboseMode% Then
msg$ = vbCrLf & "Standard " & Format$(num%) & ", (before update for standard composition):"
Call IOWriteLog(msg$)
For i% = 1 To UpdateStdSample(1).LastChan%
msg$ = Str$(i%) & " " & UpdateStdSample(1).Elsyms$(i%) & " " & UpdateStdSample(1).Xrsyms$(i%) & Str$(UpdateStdSample(1).MotorNumbers%(i%)) & " " & UpdateStdSample(1).CrystalNames$(i%) & ", " & Format$(UpdateStdSample(1).TakeoffArray!(i%)) & " " & Format$(UpdateStdSample(1).KilovoltsArray!(i%))
Call IOWriteLog(msg$)
Next i%
End If

' Now zero out concentrations (in case any are loaded)
For i% = 1 To UpdateStdSample(1).LastChan%
UpdateStdSample(1).ElmPercents!(i%) = NOT_ANALYZED_VALUE_SINGLE!                                 ' use a non-zero value
Next i%

' Now load standard concentrations for normal samples (only load concentration once for each elements)
For j% = 1 To stdsample(1).LastChan%
ip% = IPOS9(stdsample(1).Elsyms$(j%), UpdateStdSample())                                         ' only check first not disabled occurance of element to load concentration
If ip% > 0 Then
UpdateStdSample(1).ElmPercents!(ip%) = stdsample(1).ElmPercents!(j%)                             ' load concentration from standard database
End If
Next j%

If VerboseMode% Then
msg$ = vbCrLf & "Standard " & Format$(num%) & ", (after update for standard concentrations):"
Call IOWriteLog(msg$)
For i% = 1 To UpdateStdSample(1).LastChan%
msg$ = Str$(i%) & " " & UpdateStdSample(1).Elsyms$(i%) & " " & UpdateStdSample(1).Xrsyms$(i%) & Str$(UpdateStdSample(1).MotorNumbers%(i%)) & " " & UpdateStdSample(1).CrystalNames$(i%) & ", " & Format$(UpdateStdSample(1).TakeoffArray!(i%)) & " " & Format$(UpdateStdSample(1).KilovoltsArray!(i%)) & ", " & Format$(UpdateStdSample(1).ElmPercents!(i%))
Call IOWriteLog(msg$)
Next i%
End If

' Reload passed standard sample from modified unknown sample
stdsample(1) = UpdateStdSample(1)

' Update some standard sample parameters for ZAFStd calculation
stdsample(1).number% = num%
stdsample(1).OxideOrElemental% = 2      ' always calculate standard k-factors as elemental

Exit Sub

' Errors
UpdateCalculateUpdateStandard2Error:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateCalculateUpdateStandard2"
ierror = True
Exit Sub

UpdateCalculateUpdateStandard2TooManyElements:
msg$ = "Too many elements in standard number " & Str$(stdsample(1).number%)
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateCalculateUpdateStandard2"
ierror = True
Exit Sub

End Sub

