Attribute VB_Name = "CodeZAF4"
' (c) Copyright 1995-2015 by John J. Donovan (credit to John Armstrong for original code)
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub ZAFPrintMAC2(zaf As TypeZAF)
' Print current ZAF mass absorption coefficients

ierror = False
On Error GoTo ZAFPrintMAC2Error

Dim empflag As Boolean
Dim i As Integer, i1 As Integer, im4 As Integer
Dim emtz As Integer, absz As Integer
Dim mac As Single
Dim tstring As String, tmsg As String
Dim tfactor As Single, tstandard As String

Call IOWriteLog(vbCrLf & "Current Mass Absorption Coefficients From:")
msg$ = macstring$(MACTypeFlag%)
Call IOWriteLog(msg$)
msg$ = vbCrLf & Format$("Z-LINE", a80$) & Format$("X-RAY", a80$) & Format$("Z-ABSOR", a80$) & Format$("MAC", a80$)
Call IOWriteLog(msg$)

' Print element MACs
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
For i1% = 1 To zaf.in0%
tmsg$ = Format$(Symup$(Int(zaf.Z!(i%))), a80$) & Format$(Xraylo$(zaf.il(i%)), a80$) & Format$(Symup$(Int(zaf.Z!(i1%))), a80$) & "  " & Format$(zaf.mup!(i1%, i%), e104$)

' Add asterisk if empirical
If UCase$(app.EXEName) = UCase$("Calczaf") Or UseMACFlag Then
emtz% = Int(zaf.Z!(i%))
im4% = zaf.il%(i%)
absz% = Int(zaf.Z!(i1%))
Call EmpLoadMACAPF(Int(1), emtz%, im4%, absz%, mac!, tstring$, tfactor!, tstandard$)
If ierror Then Exit Sub
If mac! > 0# Then
tmsg$ = tmsg$ & " *"
empflag = True
End If
End If

Call IOWriteLog(tmsg$)
Next i1%
End If
Next i%

If empflag Then
msg$ = " * indicates empirical MAC"
Call IOWriteLog(msg$)

' Print empirical MACs
Call IOWriteLog(vbCrLf & "Empirical Mass Absorption Coefficients From:")
msg$ = EmpMACFile$
Call IOWriteLog(msg$)
msg$ = vbCrLf & Format$("Z-LINE", a80$) & Format$("X-RAY", a80$) & Format$("Z-ABSOR", a80$) & Format$("MAC", a80$)
Call IOWriteLog(msg$)

For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
For i1% = 1 To zaf.in0%
emtz% = Int(zaf.Z!(i%))
im4% = zaf.il%(i%)
absz% = Int(zaf.Z!(i1%))
Call EmpLoadMACAPF(Int(1), emtz%, im4%, absz%, mac!, tstring$, tfactor!, tstandard$)
If ierror Then Exit Sub
If mac! > 0# Then
msg$ = Format$(Symup$(emtz%), a80$) & Format$(Xraylo$(im4%), a80$) & Format$(Symup$(absz%), a80$) & "  " & Format$(mac!, e104$) & "    " & tstring$
Call IOWriteLog(msg$)
End If
Next i1%
End If
Next i%
End If

Exit Sub

' Errors
ZAFPrintMAC2Error:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFPrintMAC2"
ierror = True
Exit Sub

End Sub

Sub ZAFPrintAPF2(zaf As TypeZAF)
' Print current ZAF area peak factors

ierror = False
On Error GoTo ZAFPrintAPF2Error

Dim i As Integer, i1 As Integer, im4 As Integer
Dim emtz As Integer, absz As Integer
Dim apf As Single
Dim tstring As String
Dim tfactor As Single, tstandard As String

' Print empirical MACs
Call IOWriteLog(vbCrLf & "Empirical Area Peak Factors (APF) From:")
msg$ = EmpAPFFile$
Call IOWriteLog(msg$)
msg$ = vbCrLf & Format$("Z-LINE", a80$) & Format$("X-RAY", a80$) & Format$("Z-ABSOR", a80$) & "  " & Format$("APF", a80$) & "  " & Format$("RE-NORM", a80$)
Call IOWriteLog(msg$)

For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
For i1% = 1 To zaf.in0%
emtz% = Int(zaf.Z!(i%))
im4% = zaf.il%(i%)
absz% = Int(zaf.Z!(i1%))
Call EmpLoadMACAPF(Int(2), emtz%, im4%, absz%, apf!, tstring$, tfactor!, tstandard$)
If ierror Then Exit Sub
If apf! <> 1# Then
msg$ = Format$(Symup$(emtz%), a80$) & Format$(Xraylo$(im4%), a80$) & Format$(Symup$(absz%), a80$) & "  " & Format$(Format$(apf!, f84$), a80$) & "  " & Format$(Format$(tfactor!, f84$), a80$) & "    " & tstring$
Call IOWriteLog(msg$)
End If
Next i1%
End If
Next i%

Exit Sub

' Errors
ZAFPrintAPF2Error:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFPrintAPF2"
ierror = True
Exit Sub

End Sub

Sub ZAFPrintSmp(zaf As TypeZAF, tzbar As Single, tdisplayoxide As Integer, texcess As Single)
' Print sample ZAF results if debug

ierror = False
On Error GoTo ZAFPrintSmpError

Dim i As Integer

' Print sample heading and parameters
msg$ = vbCrLf & "SAMPLE: " & Format$(zaf.n8&) & ", TOA: " & Format$(zaf.TOA!) & ", ITERATIONS: " & Format$(zaf.iter%) & ", Z-BAR: " & Format$(tzbar!)
Call IOWriteLog(msg$)

' Print particle parameters if specified (averaged for all elements?)
If UseParticleCorrectionFlag And iptc% = 1 Then
msg$ = vbCrLf & "E-RANGE: " & MiscAutoFormat$(zaf.erange!(1) * MICRONSPERCM&) & ", INTE-STEP: " & Str$(zaf.intnum&(1)) & vbCrLf
msg$ = msg$ & "Particle or thin film corrections utilized were " & ptcstring$(PTCModel%) & vbCrLf & vbCrLf
msg$ = msg$ & "Particle parameters were a particle diameter of " & Format$(PTCDiameter!) & " microns, "
msg$ = msg$ & "a particle density of " & Format$(PTCDensity!) & " gm/cm^3, "
msg$ = msg$ & "a thickness factor of " & Format$(PTCThicknessFactor!) & ", "
msg$ = msg$ & "and a numerical integration step size of " & Format$(PTCNumericalIntegrationStep!) & " microns."
Call IOWriteLog(msg$)
End If

' Output the element z and the correction factors
msg$ = vbCrLf & " ELEMENT  ABSCOR  FLUCOR  ZEDCOR  ZAFCOR STP-POW BKS-COR   F(x)u      Ec   Eo/Ec    MACs"
Call IOWriteLog(msg$)

For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
msg$ = Format$(Symup$(Int(zaf.Z!(i%))) & " " & Xraylo$(Int(zaf.il%(i%))), a80$)

' Absorption
If zaf.genstd!(i%) <> 0# And (1# + zaf.vv!(i%)) <> 0# And zaf.gensmp!(i%) <> 0# And zaf.il%(i%) <> 10 Then
If Not ZAFEquationMode Then
msg$ = msg$ & Format$(Format$(zaf.gensmp!(i%) / zaf.genstd!(i%), f84$), a80$)
Else
msg$ = msg$ & Format$(Format$(1# / ((zaf.gensmp!(i%) / zaf.genstd!(i%))), f84$), a80$)
End If
Else
msg$ = msg$ & Format$("*****", a80$)
End If

' Fluorescence
If zaf.genstd!(i%) <> 0# And (1# + zaf.vv!(i%)) <> 0# And zaf.gensmp!(i%) <> 0# And zaf.il%(i%) <> 10 Then
If Not ZAFEquationMode Then
msg$ = msg$ & Format$(Format$(1# / (1# + zaf.vv!(i%)), f84$), a80$)
Else
msg$ = msg$ & Format$(Format$((1# + zaf.vv!(i%)), f84$), a80$)
End If
Else
msg$ = msg$ & Format$("*****", a80$)
End If

' Atomic number
If Not ZAFEquationMode Then
msg$ = msg$ & Format$(Format$(zaf.zed!(i%), f84$), a80$)
Else
msg$ = msg$ & Format$(Format$(1# / zaf.zed!(i%), f84$), a80$)
End If

' ZAF Correction
If zaf.genstd!(i%) <> 0# And (1# + zaf.vv!(i%)) <> 0# And zaf.gensmp!(i%) <> 0# And zaf.il%(i%) <> 10 Then
If Not ZAFEquationMode Then
msg$ = msg$ & Format$(Format$((zaf.gensmp!(i%) / zaf.genstd!(i%)) * zaf.zed!(i%) / (1# + zaf.vv!(i%)), f84$), a80$)
Else
msg$ = msg$ & Format$(Format$((1# / (zaf.gensmp!(i%) / zaf.genstd!(i%)) * (1# + zaf.vv!(i%)) / zaf.zed!(i%)), f84$), a80$)
End If
Else
msg$ = msg$ & Format$("*****", a80$)
End If

' Stopping power
If zaf.stp!(i%) <> 0# Then
If Not ZAFEquationMode Then
msg$ = msg$ & Format$(Format$(zaf.stp!(i%), f84$), a80$)
Else
msg$ = msg$ & Format$(Format$(1# / zaf.stp!(i%), f84$), a80$)
End If
Else
msg$ = msg$ & Format$("*****", a80$)
End If

' Backscatter loss
msg$ = msg$ & Format$(Format$(zaf.bks!(i%), f84$), a80$)

' F(chi)
If zaf.gensmp!(i%) <> 0# Then
msg$ = msg$ & Format$(Format$(1# / zaf.gensmp!(i%), f84$), a80$)
Else
msg$ = msg$ & Format$("*****", a80$)
End If

' Edge and overvoltage
msg$ = msg$ & Format$(Format$(zaf.eC!(i%), f84$), a80$)
msg$ = msg$ & Format$(Format$(zaf.v!(i%), f84$), a80$)

' Emitter MACs
msg$ = msg$ & MiscAutoFormat$(zaf.MACs!(i%))

' Output line
Call IOWriteLog(msg$)
End If
Next i%

msg$ = vbCrLf & " ELEMENT   K-RAW K-VALUE ELEMWT% OXIDWT% ATOMIC% FORMULA KILOVOL"
If UseParticleCorrectionFlag And iptc% = 1 Then msg$ = msg$ & " NORMEL% NORMOX%"
If UseConductiveCoatingCorrectionForXrayTransmission Then
msg$ = msg$ & " COATTRN COATU/S"
Else
msg$ = msg$ & "                "
End If
If UseConductiveCoatingCorrectionForElectronAbsorption Then
msg$ = msg$ & " KILOVOL"
msg$ = msg$ & " COATABS COATU/S"
Else
msg$ = msg$ & "        "
msg$ = msg$ & "                "
End If
Call IOWriteLog(msg$)

' Output the element Z, k-ratio and concentrations
For i% = 1 To zaf.in0%

' Stoichiometric oxygen
If zaf.il%(zaf.in0%) = 0 And i% = zaf.in0 Then
If Not UseAutomaticFormatForResultsFlag Then
msg$ = Format$(Symup$(Int(zaf.Z!(i%))) & "   ", a80$) & a8x$ & a8x$ & Format$(Format$(100# * zaf.conc!(i%), f83$), a80$) & Format$(DASHED5$, a80$) & Format$(Format$(zaf.AtPercents!(i%), f83$), a80$) & Format$(Format$(zaf.Formulas!(i%), f83$), a80$)
Else
msg$ = Format$(Symup$(Int(zaf.Z!(i%))) & "   ", a80$) & a8x$ & a8x$ & MiscAutoFormat$(100# * zaf.conc!(i%)) & Format$(DASHED5$, a80$) & MiscAutoFormat$(zaf.AtPercents!(i%)) & MiscAutoFormat$(zaf.Formulas!(i%))
End If
If UseParticleCorrectionFlag And iptc% = 1 Then msg$ = msg$ & a8x$ & MiscAutoFormat$(zaf.NormElPercents!(i%))

' Specified element
ElseIf zaf.il%(i%) > MAXRAY% - 1 Then
If Int(zaf.Z!(i%)) <> 8 Then
If zaf.in1% <> zaf.in0% Or tdisplayoxide% Then     ' using stoichiometric oxygen
If Not UseAutomaticFormatForResultsFlag Then
msg$ = Format$(Symup$(Int(zaf.Z!(i%))) & "   ", a80$) & a8x$ & a8x$ & Format$(Format$(100# * zaf.conc!(i%), f83$), a80$) & Format$(Format$(zaf.OxPercents!(i%), f83$), a80$) & Format$(Format$(zaf.AtPercents!(i%), f83$), a80$) & Format$(Format$(zaf.Formulas!(i%), f83$), a80$)
Else
msg$ = Format$(Symup$(Int(zaf.Z!(i%))) & "   ", a80$) & a8x$ & a8x$ & MiscAutoFormat$(100# * zaf.conc!(i%)) & MiscAutoFormat$(zaf.OxPercents!(i%)) & MiscAutoFormat$(zaf.AtPercents!(i%)) & MiscAutoFormat$(zaf.Formulas!(i%))
End If
Else
If Not UseAutomaticFormatForResultsFlag Then
msg$ = Format$(Symup$(Int(zaf.Z!(i%))) & "   ", a80$) & a8x$ & a8x$ & Format$(Format$(100# * zaf.conc!(i%), f83$), a80$) & Format$(DASHED5$, a80$) & Format$(Format$(zaf.AtPercents!(i%), f83$), a80$) & Format$(Format$(zaf.Formulas!(i%), f83$), a80$)
Else
msg$ = Format$(Symup$(Int(zaf.Z!(i%))) & "   ", a80$) & a8x$ & a8x$ & MiscAutoFormat$(100# * zaf.conc!(i%)) & Format$(DASHED5$, a80$) & MiscAutoFormat$(zaf.AtPercents!(i%)) & MiscAutoFormat$(zaf.Formulas!(i%))
End If
End If
Else
If zaf.in1% <> zaf.in0% Or tdisplayoxide% Then     ' using stoichiometric oxygen
If Not UseAutomaticFormatForResultsFlag Then
msg$ = Format$(Symup$(Int(zaf.Z!(i%))) & "   ", a80$) & a8x$ & a8x$ & Format$(Format$(100# * zaf.conc!(i%), f83$), a80$) & Format$(Format$(texcess!, f83$), a80$) & Format$(Format$(zaf.AtPercents!(i%), f83$), a80$) & Format$(Format$(zaf.Formulas!(i%), f83$), a80$)
Else
msg$ = Format$(Symup$(Int(zaf.Z!(i%))) & "   ", a80$) & a8x$ & a8x$ & MiscAutoFormat$(100# * zaf.conc!(i%)) & MiscAutoFormat$(texcess!) & MiscAutoFormat$(zaf.AtPercents!(i%)) & MiscAutoFormat$(zaf.Formulas!(i%))
End If
Else
If Not UseAutomaticFormatForResultsFlag Then
msg$ = Format$(Symup$(Int(zaf.Z!(i%))) & "   ", a80$) & a8x$ & a8x$ & Format$(Format$(100# * zaf.conc!(i%), f83$), a80$) & Format$(DASHED5$, a80$) & Format$(Format$(zaf.AtPercents!(i%), f83$), a80$) & Format$(Format$(zaf.Formulas!(i%), f83$), a80$)
Else
msg$ = Format$(Symup$(Int(zaf.Z!(i%))) & "   ", a80$) & a8x$ & a8x$ & MiscAutoFormat$(100# * zaf.conc!(i%)) & Format$(DASHED5$, a80$) & MiscAutoFormat$(zaf.AtPercents!(i%)) & MiscAutoFormat$(zaf.Formulas!(i%))
End If
End If
End If

' Analyzed element
Else
If Int(zaf.Z!(i%)) <> 8 Then
If zaf.in1% <> zaf.in0% Or tdisplayoxide% Then      ' using stoichiometric oxygen
If Not UseAutomaticFormatForResultsFlag Then
msg$ = Format$(Symup$(Int(zaf.Z!(i%))) & " " & Xraylo$(Int(zaf.il%(i%))), a80$) & Format$(Format$(zaf.kraw!(i%), f85$), a80$) & Format$(Format$(zaf.krat!(i%), f85$), a80$) & Format$(Format$(100# * zaf.conc!(i%), f83$), a80$) & Format$(Format$(zaf.OxPercents!(i%), f83$), a80$) & Format$(Format$(zaf.AtPercents!(i%), f83$), a80$) & Format$(Format$(zaf.Formulas!(i%), f83$), a80$) & Format$(Format$(zaf.eO!(i%), f82$), a80$)
Else
msg$ = Format$(Symup$(Int(zaf.Z!(i%))) & " " & Xraylo$(Int(zaf.il%(i%))), a80$) & MiscAutoFormat$(zaf.kraw!(i%)) & MiscAutoFormat$(zaf.krat!(i%)) & MiscAutoFormat$(100# * zaf.conc!(i%)) & MiscAutoFormat$(zaf.OxPercents!(i%)) & MiscAutoFormat$(zaf.AtPercents!(i%)) & MiscAutoFormat$(zaf.Formulas!(i%)) & Format$(Format$(zaf.eO!(i%), f82$), a80$)
End If
Else
If Not UseAutomaticFormatForResultsFlag Then
msg$ = Format$(Symup$(Int(zaf.Z!(i%))) & " " & Xraylo$(Int(zaf.il%(i%))), a80$) & Format$(Format$(zaf.kraw!(i%), f85$), a80$) & Format$(Format$(zaf.krat!(i%), f85$), a80$) & Format$(Format$(100# * zaf.conc!(i%), f83$), a80$) & Format$(DASHED5$, a80$) & Format$(Format$(zaf.AtPercents!(i%), f83$), a80$) & Format$(Format$(zaf.Formulas!(i%), f83$), a80$) & Format$(Format$(zaf.eO!(i%), f82$), a80$)
Else
msg$ = Format$(Symup$(Int(zaf.Z!(i%))) & " " & Xraylo$(Int(zaf.il%(i%))), a80$) & MiscAutoFormat$(zaf.kraw!(i%)) & MiscAutoFormat$(zaf.krat!(i%)) & MiscAutoFormat$(100# * zaf.conc!(i%)) & Format$(DASHED5$, a80$) & MiscAutoFormat$(zaf.AtPercents!(i%)) & MiscAutoFormat$(zaf.Formulas!(i%)) & Format$(Format$(zaf.eO!(i%), f82$), a80$)
End If
End If
Else
If zaf.in1% <> zaf.in0% Or tdisplayoxide% Then      ' using stoichiometric oxygen
If Not UseAutomaticFormatForResultsFlag Then
msg$ = Format$(Symup$(Int(zaf.Z!(i%))) & "   ", a80$) & a8x$ & a8x$ & Format$(Format$(100# * zaf.conc!(i%), f83$), a80$) & Format$(Format$(texcess!, f83$), a80$) & Format$(Format$(zaf.AtPercents!(i%), f83$), a80$) & Format$(Format$(zaf.Formulas!(i%), f83$), a80$)
Else
msg$ = Format$(Symup$(Int(zaf.Z!(i%))) & "   ", a80$) & a8x$ & a8x$ & MiscAutoFormat$(100# * zaf.conc!(i%)) & MiscAutoFormat$(texcess!) & MiscAutoFormat$(zaf.AtPercents!(i%)) & MiscAutoFormat$(zaf.Formulas!(i%))
End If
Else
If Not UseAutomaticFormatForResultsFlag Then
msg$ = Format$(Symup$(Int(zaf.Z!(i%))) & " " & Xraylo$(Int(zaf.il%(i%))), a80$) & Format$(Format$(zaf.kraw!(i%), f85$), a80$) & Format$(Format$(zaf.krat!(i%), f85$), a80$) & Format$(Format$(100# * zaf.conc!(i%), f83$), a80$) & Format$(DASHED5$, a80$) & Format$(Format$(zaf.AtPercents!(i%), f83$), a80$) & Format$(Format$(zaf.Formulas!(i%), f83$), a80$) & Format$(Format$(zaf.eO!(i%), f82$), a80$)
Else
msg$ = Format$(Symup$(Int(zaf.Z!(i%))) & " " & Xraylo$(Int(zaf.il%(i%))), a80$) & MiscAutoFormat$(zaf.kraw!(i%)) & MiscAutoFormat$(zaf.krat!(i%)) & MiscAutoFormat$(100# * zaf.conc!(i%)) & Format$(DASHED5$, a80$) & MiscAutoFormat$(zaf.AtPercents!(i%)) & MiscAutoFormat$(zaf.Formulas!(i%)) & Format$(Format$(zaf.eO!(i%), f82$), a80$)
End If
End If
End If

If UseParticleCorrectionFlag And iptc% = 1 Then
If zaf.in1% <> zaf.in0% Then     ' using stoichiometric oxygen
msg$ = msg$ & MiscAutoFormat$(zaf.NormElPercents!(i%)) & MiscAutoFormat$(zaf.NormOxPercents!(i%))
Else
msg$ = msg$ & MiscAutoFormat$(zaf.NormElPercents!(i%)) & Format$(DASHED5$, a80$)
End If
End If

' Output coating corrections
If zaf.coating_flag% = 1 Then
If UseConductiveCoatingCorrectionForXrayTransmission Then
If zaf.coating_trans_std_assigns!(i%) <> 0# Then
msg$ = msg$ & Format$(Format$(zaf.coating_trans_smp!(i%), f85$), a80$) & Format$(Format$(zaf.coating_trans_smp!(i%) / zaf.coating_trans_std_assigns!(i%), f85$), a80$)
Else
msg$ = msg$ & Format$(DASHED5$, a80$) & Format$(DASHED5$, a80$)
End If
Else
msg$ = msg$ & Format$("     ", a80$) & Format$("     ", a80$)
End If
End If

If UseConductiveCoatingCorrectionForElectronAbsorption Then
If zaf.coating_absorbs_std_assigns!(i%) <> 0# Then
msg$ = msg$ & Format$(Format$(zaf.coating_actual_kilovolts!(i%), f82$), a80$) & Format$(Format$(zaf.coating_absorbs_smp!(i%), f85$), a80$) & Format$(Format$(zaf.coating_absorbs_smp!(i%) / zaf.coating_absorbs_std_assigns!(i%), f85$), a80$)
Else
msg$ = msg$ & Format$(DASHED5$, a80$) & Format$(DASHED5$, a80$) & Format$(DASHED5$, a80$)
End If
Else
msg$ = msg$ & Format$("     ", a80$) & Format$("     ", a80$) & Format$("     ", a80$)
End If
End If

Call IOWriteLog(msg$)
Next i%

' Print total line
If zaf.in1% <> zaf.in0% Or tdisplayoxide% Then      ' using stoichiometric oxygen
msg$ = "   TOTAL: " & a6x$ & a8x$ & Format$(Format$(100# * zaf.ksum!, f83$), a80$) & Format$(Format$(100# * zaf.ksum!, f83$), a80$) & Format$(Format$(100#, f83$), a80$) & Format$(Format$(zaf.TotalCations!, f83$), a80$)
Else
msg$ = "   TOTAL: " & a6x$ & a8x$ & Format$(Format$(100# * zaf.ksum!, f83$), a80$) & Format$(DASHED5$, a80$) & Format$(Format$(100#, f83$), a80$) & Format$(Format$(zaf.TotalCations!, f83$), a80$)
End If
If UseParticleCorrectionFlag And iptc% = 1 Then
If zaf.in1% <> zaf.in0% Then     ' using stoichiometric oxygen
msg$ = msg$ & a8x$ & Format$(Format$(100#, f83$), a80$) & Format$(Format$(100#, f83$), a80$)
Else
msg$ = msg$ & a8x$ & Format$(Format$(100#, f83$), a80$) & Format$(DASHED5$, a80$)
End If
End If
Call IOWriteLog(msg$)

Exit Sub

' Errors
ZAFPrintSmpError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFPrintSmp"
ierror = True
Exit Sub

End Sub

Sub ZAFPrintStd(zaf As TypeZAF)
' Type out section for standards (primary intensities)

ierror = False
On Error GoTo ZAFPrintStdError

Dim i As Integer

If VerboseMode Or (Not VerboseMode And UCase$(app.EXEName) = UCase$("CalcZAF") And CalcZAFMode% = 0) Then

' Type MAC header
Call ZAFPrintMAC2(zaf)
If ierror Then Exit Sub

msg$ = vbCrLf & " ELEMENT  ABSFAC  ZEDFAC  FINFAC STP-POW BKS-COR   F(x)e"
Call IOWriteLog(msg$)

' Element primary factors
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
msg$ = Format$(Symup$(Int(zaf.Z!(i%))) & " " & Xraylo$(Int(zaf.il%(i%))), a80$)

' Absorption
msg$ = msg$ & Format$(Format$(zaf.genstd!(i%), f84), a80$)

' Atomic number
msg$ = msg$ & Format$(Format$(zaf.r!(i%, i%) / zaf.s!(i%, i%), f84), a80$)

msg$ = msg$ & Format$(Format$(zaf.genstd!(i%) * zaf.r!(i%, i%) / zaf.s!(i%, i%), f84), a80$)

' Stopping power
msg$ = msg$ & Format$(Format$(zaf.s!(i%, i%), f84), a80$)

' Backscatter loss
msg$ = msg$ & Format$(Format$(zaf.r!(i%, i%), f84), a80$)

' F(chi)
msg$ = msg$ & Format$(Format$(1# / zaf.genstd!(i%), f84), a80$)

Call IOWriteLog(msg$)
End If
Next i%

End If

Exit Sub

' Errors
ZAFPrintStdError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFPrintStd"
ierror = True
Exit Sub

End Sub

Sub ZAFPrintStandards2(zaf As TypeZAF, analysis As TypeAnalysis, sample() As TypeSample)
' Print out the standard k-factors, etc. for all standards in the run

ierror = False
On Error GoTo ZAFPrintStandards2Error

Dim i As Integer

If (UCase$(app.EXEName) = UCase$("CalcZAF") And CalcZAFMode% < 3) Or (UCase$(app.EXEName) = UCase$("Probewin") And DebugMode) Then

' Print sample heading and parameters
msg$ = vbCrLf & "STANDARD PARAMETERS (TOA= " & Format$(zaf.TOA!) & "):"
Call IOWriteLog(msg$)

' Output the standard composition and k-factors, zbar, etc
msg$ = vbCrLf & " ELEMENT  STDNUM STDCONC STDKFAC   Z-BAR  ABSCOR  FLUCOR  ZEDCOR  ZAFCOR"
If UseConductiveCoatingCorrectionForXrayTransmission Then
msg$ = msg$ & " COATTRNS"
Else
msg$ = msg$ & "         "
End If
If UseConductiveCoatingCorrectionForElectronAbsorption Then
msg$ = msg$ & " COATABSR"
Else
msg$ = msg$ & "         "
End If
Call IOWriteLog(msg$)

For i% = 1 To sample(1).LastElm%
msg$ = Format$(MiscAutoUcase$(analysis.Elsyms$(i%)) & " " & MiscAutoUcase$(analysis.Xrsyms$(i%)), a80$)
If sample(1).StdAssigns%(i%) <> MAXINTEGER% And CalcZAFMode% < 3 Then
msg$ = msg$ & Format$(Format$(sample(1).StdAssigns%(i%)), a80$)
Else
msg$ = msg$ & Format$(DASHED5$, a80$)
End If
msg$ = msg$ & Format$(Format$(analysis.StdAssignsPercents!(i%), f83$), a80$)
msg$ = msg$ & Format$(Format$(analysis.StdAssignsKrats!(i%), f84$), a80$)
msg$ = msg$ & Format$(Format$(analysis.StdAssignsZbars!(i%), f84$), a80$)

' Absorption
If Not ZAFEquationMode% Then
msg$ = msg$ & Format$(Format$(analysis.StdAssignsZAFCors!(1, i%), f84$), a80$)
Else
msg$ = msg$ & Format$(Format$(1# / analysis.StdAssignsZAFCors!(1, i%), f84$), a80$)
End If

' Fluorescence
If Not ZAFEquationMode% Then
msg$ = msg$ & Format$(Format$(analysis.StdAssignsZAFCors!(2, i%), f84$), a80$)
Else
msg$ = msg$ & Format$(Format$(1# / analysis.StdAssignsZAFCors!(2, i%), f84$), a80$)
End If

' Atomic number
msg$ = msg$ & Format$(Format$(analysis.StdAssignsZAFCors!(3, i%), f84$), a80$)

' ZAF correction
msg$ = msg$ & Format$(Format$(analysis.StdAssignsZAFCors!(4, i%), f84$), a80$)

' Coating x-ray transmission for each standard
If UseConductiveCoatingCorrectionForXrayTransmission Then
If DefaultStandardCoatingFlag% = 1 Then
msg$ = msg$ & Format$(Format$(analysis.Coating_StdAssignsTrans!(i%), f85$), a80$)
Else
msg$ = msg$ & Format$(DASHED5$, a80$)
End If
End If

' Coating electron absorption for each standard
If UseConductiveCoatingCorrectionForElectronAbsorption Then
If DefaultStandardCoatingFlag% = 1 Then
msg$ = msg$ & Format$(Format$(analysis.Coating_StdAssignsAbsorbs!(i%), f85$), a80$)
Else
msg$ = msg$ & Format$(DASHED5$, a80$)
End If
End If

Call IOWriteLog(msg$)
Next i%

msg$ = vbCrLf & " ELEMENT STP-POW BKS-COR   F(x)e   F(x)s      Eo      Ec   Eo/Ec"
Call IOWriteLog(msg$)

For i% = 1 To sample(1).LastElm%
msg$ = Format$(MiscAutoUcase$(analysis.Elsyms$(i%)) & " " & MiscAutoUcase$(analysis.Xrsyms$(i%)), a80$)

' Stopping power
If Not ZAFEquationMode% Then
msg$ = msg$ & Format$(Format$(analysis.StdAssignsZAFCors!(5, i%), f84$), a80$)
Else
msg$ = msg$ & Format$(Format$(1# / analysis.StdAssignsZAFCors!(5, i%), f84$), a80$)
End If

' Backscatter loss
msg$ = msg$ & Format$(Format$(analysis.StdAssignsZAFCors!(6, i%), f84$), a80$)

' Emitted to generated ratio intensities for pure element and standard
msg$ = msg$ & Format$(Format$(analysis.StdAssignsZAFCors!(7, i%), f84$), a80$)
msg$ = msg$ & Format$(Format$(analysis.StdAssignsZAFCors!(8, i%), f84$), a80$)

' Edge and overvoltage
msg$ = msg$ & Format$(Format$(analysis.StdAssignsActualKilovolts!(i%), f82$), a80$)
msg$ = msg$ & Format$(Format$(analysis.StdAssignsEdgeEnergies!(i%), f84$), a80$)
msg$ = msg$ & Format$(Format$(analysis.StdAssignsActualOvervoltages!(i%), f84$), a80$)

Call IOWriteLog(msg$)
Next i%

End If
Exit Sub

' Errors
ZAFPrintStandards2Error:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFPrintStandards2"
ierror = True
Exit Sub

End Sub

Sub ZAFPrintCalculate(zaf As TypeZAF, analysis As TypeAnalysis, sample() As TypeSample)
' Calculate print out variables (atomic, oxide, formulas, etc)

ierror = False
On Error GoTo ZAFPrintCalculateError

Dim i As Integer, ip As Integer
Dim sum As Single

' Determine formula basis if not sum of cations
If sample(1).FormulaElementFlag And sample(1).FormulaElement$ <> vbNullString Then
ip% = IPOS1(sample(1).LastChan%, sample(1).FormulaElement$, sample(1).Elsyms$())
If ip% = sample(1).OxygenChannel% And sample(1).OxideOrElemental% = 1 Then ip% = sample(1).LastChan% + 1   ' if oxygen by stoichiometry, use that channel
If ip% = 0 Then GoTo ZAFPrintCalculateInvalidFormulaElement
End If

' Calculate the atomic percent for each concentration
analysis.TotalCations! = 0#
For i% = 1 To zaf.in0%
analysis.TotalCations! = analysis.TotalCations! + zaf.conc!(i%) / zaf.atwts!(i%)
Next i%
For i% = 1 To zaf.in0%
analysis.AtPercents!(i%) = 0#
If analysis.TotalCations! > 0# Then analysis.AtPercents!(i%) = (zaf.conc!(i%) / zaf.atwts!(i%)) * (100# / analysis.TotalCations!)
Next i%

' Calculate formula if indicated
If UCase$(app.EXEName) = UCase$("CalcZAF") And (CalcZAFMode% = 0 Or sample(1).FormulaElementFlag = False) Then
sample(1).FormulaElementFlag = True
If sample(1).FormulaRatio! = 0# Then         ' calculate formula basis of 1 cation for CalcZAFMode% = 0
ip% = 0
sample(1).FormulaRatio! = 1#
End If
End If

analysis.TotalCations! = 0#
If sample(1).FormulaElementFlag Then
If ip% <> 0 Then     ' normal formula calculation

' Normalize to formula basis
For i% = 1 To zaf.in0%
analysis.Formulas!(i%) = 0#
If analysis.AtPercents!(ip%) > 0# Then analysis.Formulas!(i%) = (sample(1).FormulaRatio! / analysis.AtPercents!(ip%)) * analysis.AtPercents!(i%)
Next i%

' Calculate sum of cations
Else
analysis.TotalCations! = 0#
For i% = 1 To zaf.in0%
If AllAtomicCharges!(CInt(zaf.Z!(i%))) > 0# Then analysis.TotalCations! = analysis.TotalCations! + analysis.AtPercents!(i%)   ' skip anions
Next i%

' Normalize to total number of cations
For i% = 1 To zaf.in0%
analysis.Formulas!(i%) = 0#
If analysis.TotalCations! > 0# Then analysis.Formulas!(i%) = analysis.AtPercents!(i%) * sample(1).FormulaRatio! / analysis.TotalCations!
Next i%
End If

' Sum formula cations
analysis.TotalCations! = 0#
For i% = 1 To zaf.in0%
analysis.TotalCations! = analysis.TotalCations! + analysis.Formulas!(i%)
Next i%
End If

' Calculate sum of elemental data (for particle normalization)
sum! = 0#
For i% = 1 To zaf.in0%  ' elemental normalization
sum! = sum! + zaf.conc!(i%)
Next i%

' Calculate oxides
For i% = 1 To zaf.in0%

' Stoichiometric oxygen
If zaf.il%(zaf.in0%) = 0 And i% = zaf.in0 Then
If sum! > 0# Then analysis.NormElPercents!(i%) = 100# * zaf.conc!(i%) / sum!
analysis.OxPercents!(i%) = 0#
analysis.NormOxPercents!(i%) = 0#

' Specified element
ElseIf zaf.il%(i%) > MAXRAY% - 1 Then
If Int(zaf.Z!(i%)) <> 8 Then
If sum! > 0# Then analysis.NormElPercents!(i%) = analysis.WtPercents!(i%) / sum!
analysis.OxPercents!(i%) = 100# * zaf.conc!(i%) * (1# + zaf.p1(i%))
analysis.NormOxPercents!(i%) = analysis.NormElPercents!(i%) * (1# + zaf.p1(i%))

Else
If sum! > 0# Then analysis.NormElPercents!(i%) = analysis.WtPercents!(i%) / sum!
analysis.OxPercents!(i%) = analysis.ExcessOxygen!
analysis.NormOxPercents!(i%) = analysis.ExcessOxygen! / sum!
End If

' Analyzed element
Else
If Int(zaf.Z!(i%)) <> 8 Then
If sum! > 0# Then analysis.NormElPercents!(i%) = analysis.WtPercents!(i%) / sum!
analysis.OxPercents!(i%) = 100# * zaf.conc!(i%) * (1# + zaf.p1(i%))
analysis.NormOxPercents!(i%) = analysis.NormElPercents!(i%) * (1# + zaf.p1(i%))

Else
If sum! > 0# Then analysis.NormElPercents!(i%) = analysis.WtPercents!(i%) / sum!
analysis.OxPercents!(i%) = analysis.ExcessOxygen!
analysis.NormOxPercents!(i%) = analysis.ExcessOxygen! / sum!
End If

End If
Next i%

' Load zaf arrays from analysis arrays (for ZAFPrintSmp procedure)
zaf.TotalCations! = analysis.TotalCations!
For i% = 1 To zaf.in0%
If i% <= MAXCHAN% Then
zaf.Formulas!(i%) = analysis.Formulas!(i%)
zaf.AtPercents!(i%) = analysis.AtPercents!(i%)
zaf.Formulas!(i%) = analysis.Formulas!(i%)
zaf.OxPercents!(i%) = analysis.OxPercents!(i%)
zaf.NormElPercents!(i%) = analysis.NormElPercents!(i%)
zaf.NormOxPercents!(i%) = analysis.NormOxPercents!(i%)
End If
Next i%

Exit Sub

' Errors
ZAFPrintCalculateError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFPrintCalculate"
ierror = True
Exit Sub

ZAFPrintCalculateInvalidFormulaElement:
msg$ = TypeLoadString$(sample())
msg$ = "Element " & sample(1).FormulaElement$ & " is an invalid formula element for sample " & msg$
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFPrintCalculate"
ierror = True
Exit Sub

End Sub

