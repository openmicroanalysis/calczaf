Attribute VB_Name = "CodeType2"
' (c) Copyright 1995-2019 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub TypeZAFSelections()
' This routine only types out the ZAF selections

ierror = False
On Error GoTo TypeZAFSelectionsError

' Type out correction method
msg$ = vbCrLf & vbCrLf & "Correction Method and Mass Absorption Coefficient File:"
Call IOWriteLog(msg$)

msg$ = corstring$(CorrectionFlag%)      ' 0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters
Call IOWriteLog(msg$)

msg$ = macstring$(MACTypeFlag%)
Call IOWriteLog(msg$)

' If alpha factor correction, print empirical alpha flag
If CorrectionFlag% > 0 And CorrectionFlag% < 5 Then
msg$ = empstring$(EmpiricalAlphaFlag%)
Call IOWriteLog(msg$)
End If

' Check if calibration curve or fundamental parameter corrections
If CorrectionFlag% = 5 Or CorrectionFlag% = MAXCORRECTION% Then Exit Sub
msg$ = vbCrLf & "Current ZAF or Phi-Rho-Z Selection:"
Call IOWriteLog(msg$)

msg$ = zafstring$(izaf%)
Call IOWriteLog(msg$)

' PTC modification (if not Probewin)
If Not MiscStringsAreSame(app.EXEName, "Probewin") Then
If UseParticleCorrectionFlag And iptc% = 1 Then
msg$ = vbCrLf & "Using Particle and Thin Film Corrections:"
Call IOWriteLog(msg$)
msg$ = ptcstring$(PTCModel%)
Call IOWriteLog(msg$)

msg$ = "Particle Diameter: " & MiscAutoFormat$(PTCDiameter!)
Call IOWriteLog(msg$)
msg$ = "Particle Density: " & MiscAutoFormat$(PTCDensity!)
Call IOWriteLog(msg$)
msg$ = "Thickness factor: " & MiscAutoFormat$(PTCThicknessFactor!)
Call IOWriteLog(msg$)
msg$ = "Numerical Integration Step: " & MiscAutoFormat$(PTCNumericalIntegrationStep!)
Call IOWriteLog(msg$)
End If
End If

msg$ = vbCrLf & "Correction Selections:"
Call IOWriteLog(msg$)

msg$ = absstring$(iabs%)
Call IOWriteLog(msg$)

msg$ = stpstring$(istp%)
Call IOWriteLog(msg$)

msg$ = bscstring$(ibsc%)
Call IOWriteLog(msg$)

msg$ = bksstring$(ibks%)
Call IOWriteLog(msg$)

msg$ = mipstring$(imip%)
Call IOWriteLog(msg$)

If iabs% > 6 Then
msg$ = phistring$(iphi%)
Call IOWriteLog(msg$)
End If

msg$ = flustring$(iflu%)
Call IOWriteLog(msg$)

' 0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters
If UseFluorescenceByBetaLinesFlag Or (CorrectionFlag% <> 0 And CorrectionFlag% <> 5 And UsePenepmaKratiosFlag = 2) Or iflu% = 5 Then
msg$ = "Fluorescence by Beta Lines Included"
Call IOWriteLog(msg$)
Else
msg$ = "Fluorescence by Beta Lines NOT Included"
Call IOWriteLog(msg$)
End If

' Print if using Penepma alpha factors
If CorrectionFlag% > 0 And CorrectionFlag% < 5 And UsePenepmaKratiosFlag = 2 Then
msg$ = "Using Penepma Derived Alpha Factors (if available)"
Call IOWriteLog(msg$)
msg$ = "Continuum Fluorescence Correction Included"
Call IOWriteLog(msg$)
Else
msg$ = "Continuum Fluorescence Correction Not Included"
End If

' Skip a line
Call IOWriteLog(vbNullString)

Exit Sub

' Errors
TypeZAFSelectionsError:
MsgBox Error$, vbOKOnly + vbCritical, "TypeZAFSelections"
ierror = True
Exit Sub

End Sub

Function TypeLoadString(sample() As TypeSample) As String
' Returns a sample type, number, set and name string based on the sample

ierror = False
On Error GoTo TypeLoadStringError

Dim tmsg As String

' Load string based on sample type
TypeLoadString = vbNullString
If sample(1).Set% > 0 Then
If sample(1).Type% = 1 Then tmsg$ = "St " & Format$(sample(1).number%, a40$) & " Set " & Format$(sample(1).Set%, a30$) & " " & sample(1).Name$
If sample(1).Type% = 2 Then tmsg$ = "Un " & Format$(sample(1).number%, a40$) & " " & sample(1).Name$
If sample(1).Type% = 3 Then tmsg$ = "Wa " & Format$(sample(1).number%, a40$) & " " & sample(1).Name$

' Standard database application only
Else
tmsg$ = "St " & Format$(sample(1).number%, a40$) & " " & sample(1).Name$
End If

TypeLoadString = tmsg$
Exit Function

' Errors
TypeLoadStringError:
MsgBox Error$, vbOKOnly + vbCritical, "TypeLoadString"
ierror = True
Exit Function

End Function

Sub TypeZbar(mode As Integer, analysis As TypeAnalysis)
' Print zbar, etc to Log Window
' mode = 1 print iterations (PROBEWIN)
' mode = 2 don't print iterations (STANDARD) or halogen corrected oxygen total

ierror = False
On Error GoTo TypeZbarError

msg$ = vbNullString
Call IOWriteLog(msg$)

msg$ = "Average Total Oxygen:     " & Format$(Format$(analysis.totaloxygen!, f83$), a80$)
msg$ = msg$ & "     " & "Average Total Weight%: " & Format$(Format$(analysis.TotalPercent!, f83), a80$)
Call IOWriteLog(msg$)

msg$ = "Average Calculated Oxygen:" & Format$(Format$(analysis.CalculatedOxygen!, f83$), a80$)
msg$ = msg$ & "     " & "Average Atomic Number: " & Format$(Format$(analysis.Zbar!, f83$), a80$)
Call IOWriteLog(msg$)

msg$ = "Average Excess Oxygen:    " & Format$(Format$(analysis.ExcessOxygen!, f83$), a80$)
msg$ = msg$ & "     " & "Average Atomic Weight: " & Format$(Format$(analysis.AtomicWeight!, f83$), a80$)
Call IOWriteLog(msg$)

' Halogen analysis
If analysis.OxygenFromHalogens! > 0# Then
msg$ = "Oxygen Equiv. from Halogen:" & Format$(Format$(analysis.OxygenFromHalogens!, f83$), a70$)
If mode% = 1 Then
msg$ = msg$ & "  " & "Halogen Corrected Oxygen: " & Format$(Format$(analysis.HalogenCorrectedOxygen!, f83$), a80$)
End If
Call IOWriteLog(msg$)
End If

' Charge balance analysis
If UseChargeBalanceCalculationFlag Then
msg$ = "Average Charge Balance:   " & Format$(Format$(analysis.ChargeBalance!, f83$), a80$)
If analysis.FeCharge! <> 0# Then
msg$ = msg$ & "     " & "Fe+ Atomic Charge:     " & Format$(Format$(analysis.FeCharge!, f83$), a80$)
Else
msg$ = msg$ & "     " & "Fe+ Atomic Charge:         ----"
End If
Call IOWriteLog(msg$)
End If

' Output matrix iterations (if not calibration curve)
If mode% = 1 And CorrectionFlag% <> 5 Then
If CorrectionFlag% = 0 Or CorrectionFlag% = MAXCORRECTION% Then
msg$ = "Average ZAF Iteration:    " & Format$(Format$(analysis.ZAFIter!, f82$), a80$)
ElseIf CorrectionFlag% > 0 And CorrectionFlag% < 5 Then
msg$ = "Average BET Iteration:    " & Format$(Format$(analysis.ZAFIter!, f82$), a80$)
End If
msg$ = msg$ & "     " & "Average Quant Iterate: " & Format$(Format$(analysis.MANIter!, f82$), a80$)
Call IOWriteLog(msg$)
End If

Exit Sub

' Errors
TypeZbarError:
MsgBox Error$, vbOKOnly + vbCritical, "TypeZbar"
ierror = True
Exit Sub

End Sub

Sub TypeGetRange(mode As Integer, i As Integer, ii As Integer, jj As Integer, sample() As TypeSample)
' Calculates typeout based on "ExtendedFormat" flag
' mode = 1 return type out range for analyzed elements only
' mode = 2 return type out range for all elements
' mode = 3 return type out range for specified elements only

ierror = False
On Error GoTo TypeGetRangeError

Dim elementsperline As Integer

' Type 8 elements per line, unless sample is a wavescan sample
elementsperline% = 8
If sample(1).Type% = 3 Then elementsperline% = 4

' Calculate the analyzed element array limits based on sample(1).LastElm% and sample(1).LastChan%
If mode% = 1 Then
ii% = elementsperline% * i% - (elementsperline% - 1)
jj% = elementsperline% * i%
If jj% > sample(1).LastElm% Then jj% = sample(1).LastElm%
End If

If mode% = 2 Then
ii% = elementsperline% * i% - (elementsperline% - 1)
jj% = elementsperline% * i%
If jj% > sample(1).LastChan% Then jj% = sample(1).LastChan%
End If

If mode% = 3 Then
If i% = 1 Then ii% = sample(1).LastElm% + 1
If i% > 1 Then ii% = jj% + 1
jj% = ii% + elementsperline% - 1
If jj% > sample(1).LastChan% Then jj% = sample(1).LastChan%
End If

' If user has selected "ExtendedFormat" line format output, adjust ranges
If ExtendedFormat Then
If i% > 1 Then
ii% = sample(1).LastChan% + 1
Exit Sub
End If

If mode% = 1 Then
ii% = 1
jj% = sample(1).LastElm%
End If

If mode% = 2 Then
ii% = 1
jj% = sample(1).LastChan%
End If

If mode% = 3 Then
ii% = sample(1).LastElm% + 1
jj% = sample(1).LastChan%
End If
End If

Exit Sub

' Errors
TypeGetRangeError:
MsgBox Error$, vbOKOnly + vbCritical, "TypeGetRange"
ierror = True
Exit Sub

End Sub

Function TypeWeight(mode As Integer, sample() As TypeSample) As String
' Function to return a string of elements and weight percents only (used by STANDARD.EXE)
'  mode = 0 return composition only
'  mode = 1 return name and composition only
'  mode = 2 return name, xray and composition

ierror = False
On Error GoTo TypeWeightError

Dim i As Integer
Dim tmsg As String
Dim sum As Single

' Loop through elements
tmsg$ = vbNullString
If mode% > 0 Then tmsg$ = sample(1).Name$ & vbCrLf
sum! = 0#
For i% = 1 To sample(1).LastChan%
If mode% = 0 Or mode% = 1 Then
tmsg$ = tmsg$ & Format$(sample(1).Elsyms$(i%), a20$) & " = "
Else
tmsg$ = tmsg$ & Format$(sample(1).Elsyms$(i%), a20$) & " " & Format$(sample(1).Xrsyms$(i%), a20$) & " = "
End If
tmsg$ = tmsg$ & Format$(Format$(sample(1).ElmPercents!(i%), f82$), a80$) & vbCrLf
sum! = sum! + sample(1).ElmPercents!(i%)
Next i%

If mode% = 0 Or mode% = 1 Then
tmsg$ = tmsg$ & "Sum: " & Format$(Format$(sum!, f82$), a80$) & vbCrLf
Else
tmsg$ = tmsg$ & "Sum: " & "   " & Format$(Format$(sum!, f82$), a80$) & vbCrLf
End If

' Return string
tmsg$ = tmsg$ & vbCrLf
TypeWeight$ = tmsg$

Exit Function

' Errors
TypeWeightError:
MsgBox Error$, vbOKOnly + vbCritical, "TypeWeight"
ierror = True
Exit Function

End Function

Sub TypeSampleFlags(analysis As TypeAnalysis, sample() As TypeSample)
' Routine to type sample calculation flags

ierror = False
On Error GoTo TypeSampleFlagsError

Dim astring As String

' Add a space to output
If sample(1).OxideOrElemental% = 1 Or (sample(1).Type% = 2 And sample(1).OxygenChannel% > sample(1).LastElm%) Or sample(1).DifferenceElementFlag% Or sample(1).DifferenceFormulaFlag% Or sample(1).StoichiometryElementFlag% Or sample(1).RelativeElementFlag% Then
Call IOWriteLog(vbNullString)
End If

' If oxide sample inform user
If sample(1).OxideOrElemental% = 1 Then
msg$ = "Oxygen Calculated by Cation Stoichiometry and Included in the Matrix Correction"
Call IOWriteLog(msg$)
If sample(1).OxygenChannel% > 0 And sample(1).OxygenChannel% <= sample(1).LastElm% Then
If sample(1).DisableQuantFlag%(sample(1).OxygenChannel%) = 0 Then
msg$ = "Warning: Oxygen is both measured and calculated by cation stoichiometry! You may observe high totals in your analyses!"
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
End If
End If

If analysis.OxygenFromHalogens! > 0# Then
If Not UseOxygenFromHalogensCorrectionFlag Then
msg$ = "Oxygen Equivalent from Halogens (F/Cl/Br/I), Not Subtracted in the Matrix Correction"
Else
msg$ = "Oxygen Equivalent from Halogens (F/Cl/Br/I), Subtracted in the Matrix Correction"
End If
Call IOWriteLog(msg$)
End If

' Also warn if unknown and calculated oxygen contains a specified amount of oxygen
If sample(1).Type% = 2 And sample(1).OxygenChannel% > sample(1).LastElm% Then
If sample(1).ElmPercents!(sample(1).OxygenChannel%) > 0# Then
msg$ = "Oxygen Calculated by Cation Stoichiometry Also Includes " & Format$(sample(1).ElmPercents!(sample(1).OxygenChannel%), f83$) & " Oxygen as a Specified Concentration"
Call IOWriteLog(msg$)
End If
End If
End If

' Elements by calculation
If sample(1).DifferenceElementFlag% Then
msg$ = "Element " & MiscAutoUcase$(sample(1).DifferenceElement$) & " is Calculated by Difference from 100%"
Call IOWriteLog(msg$)
End If

If sample(1).DifferenceFormulaFlag% Then
msg$ = "Formula " & sample(1).DifferenceFormula$ & " is Calculated by Difference from 100%"
Call IOWriteLog(msg$)
End If

If sample(1).StoichiometryElementFlag% Then
msg$ = "Element " & MiscAutoUcase$(sample(1).StoichiometryElement$) & " is Calculated by Stoichiometry to Oxygen"
msg$ = "Element " & MiscAutoUcase$(sample(1).StoichiometryElement$) & " is Calculated " & Str$(sample(1).StoichiometryRatio!) & " Atoms Relative To 1.0 Atom of Oxygen"
Call IOWriteLog(msg$)
End If

If sample(1).RelativeElementFlag% Then
msg$ = "Element " & MiscAutoUcase$(sample(1).RelativeElement$) & " is Calculated " & Str$(sample(1).RelativeRatio!) & " Atoms Relative To 1.0 Atom of " & MiscAutoUcase$(sample(1).RelativeToElement$)
Call IOWriteLog(msg$)
End If

If MiscIsElementDuplicated(sample()) And Not UseAggregateIntensitiesFlag Then
msg$ = "WARNING- Duplicate analyzed elements are present in the sample matrix!!" & vbCrLf
msg$ = msg$ & "Use Aggregate Intensity option or Disable Quant feature for accurate matrix correction."
Call IOWriteLog(msg$)
End If

If sample(1).EDSSpectraFlag And sample(1).EDSSpectraUseFlag And sample(1).EDSSpectraQuantMethodOrProject$ <> vbNullString Then
msg$ = "EDS Quant Method or Project: " & sample(1).EDSSpectraQuantMethodOrProject$
Call IOWriteLog(msg$)
End If

' PTC modification
If UseParticleCorrectionFlag And iptc% = 1 Then
msg$ = vbCrLf & "Using Particle and Thin Film Corrections:"
Call IOWriteLog(msg$)
msg$ = ptcstring$(PTCModel%)
Call IOWriteLog(msg$)

msg$ = "Particle Diameter: " & MiscAutoFormat$(PTCDiameter!)
Call IOWriteLog(msg$)
msg$ = "Particle Density: " & MiscAutoFormat$(PTCDensity!)
Call IOWriteLog(msg$)
msg$ = "Thickness factor: " & MiscAutoFormat$(PTCThicknessFactor!)
Call IOWriteLog(msg$)
msg$ = "Numerical Integration Step: " & MiscAutoFormat$(PTCNumericalIntegrationStep!)
Call IOWriteLog(msg$)
End If

' Conductive coating flags
msg$ = vbNullString
If (UCase$(app.EXEName) = UCase$("CalcZAF")) Or ProbeDataFileVersionNumber! > 7.42 Then
If UseConductiveCoatingCorrectionForElectronAbsorption = True Or UseConductiveCoatingCorrectionForXrayTransmission = True Then
If sample(1).CoatingFlag% = 1 Then
astring$ = "Sample Coating=" & Trim$(Symup$(sample(1).CoatingElement%))
astring$ = astring$ & ", Density=" & Format$(sample(1).CoatingDensity!) & " gm/cm3"
astring$ = astring$ & ", Thickness=" & Format$(sample(1).CoatingThickness!) & " angstroms"
astring$ = astring$ & ", Sin(Thickness)=" & Format$(sample(1).CoatingSinThickness!) & " angstroms"

If UseConductiveCoatingCorrectionForElectronAbsorption = True And Not UseConductiveCoatingCorrectionForXrayTransmission = True Then
msg$ = vbCrLf & "Using Conductive Coating Correction For Electron Absorption: " & vbCrLf & astring$
End If
If Not UseConductiveCoatingCorrectionForElectronAbsorption = True And UseConductiveCoatingCorrectionForXrayTransmission = True Then
msg$ = vbCrLf & "Using Conductive Coating Correction For X-Ray Transmission: " & vbCrLf & astring$
End If
If UseConductiveCoatingCorrectionForElectronAbsorption = True And UseConductiveCoatingCorrectionForXrayTransmission = True Then
msg$ = vbCrLf & "Using Conductive Coating Correction For Electron Absorption and X-Ray Transmission: " & vbCrLf & astring$
End If

' No sample coating
Else
msg$ = vbCrLf & "No Sample Coating and/or No Sample Coating Correction"
End If

Call IOWriteLog(msg$)
End If
End If

Exit Sub

' Errors
TypeSampleFlagsError:
MsgBox Error$, vbOKOnly + vbCritical, "TypeSampleFlags"
ierror = True
Exit Sub

End Sub

