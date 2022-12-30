Attribute VB_Name = "CodeZAF2"
' (c) Copyright 1995-2023 by John J. Donovan (credit to John Armstrong for original code)
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub ZAFCalZBar(tCalculatedOxygen As Single, analysis As TypeAnalysis, sample() As TypeSample)
' This routine calculates the total "analysis.TotalPercent" and the mean atomic number "analysis.Zbar!"
' for the sample from the "analysis.WtPercents!". Note: "tCalculatedOxygen" is the oxygen calculated by stoichiometry only.

ierror = False
On Error GoTo ZAFCalZbarError

Dim chan As Integer, ip As Integer
Dim temp As Single, sum_atoms As Single, sum_cations As Single
Dim total_atoms As Single, total_cations As Single

ReDim atoms(1 To MAXCHAN1%) As Single
ReDim basis(1 To MAXCHAN1%) As Single

If VerboseMode Then
Call IOWriteLog(vbCrLf & "Entering ZAFCalZBar...")
End If

' Init passed parameters
analysis.TotalPercent = 0#
analysis.totaloxygen! = 0#
analysis.TotalCations! = 0#
analysis.totalatoms! = 0#
analysis.ExcessOxygen! = 0#
analysis.CalculatedOxygen! = 0#
analysis.zbar! = 0#
analysis.AtomicWeight! = 0#
analysis.OxygenFromHalogens! = 0#
analysis.HalogenCorrectedOxygen! = 0#
analysis.ChargeBalance! = 0#
analysis.FeCharge! = 0#

' If calculating oxygen, and oxygen is analyzed or specified, add from AFACTORSMP or ZAFSMP calculation
If sample(1).OxideOrElemental% = 1 Then
analysis.CalculatedOxygen! = tCalculatedOxygen!
analysis.totaloxygen! = analysis.CalculatedOxygen!

If sample(1).OxygenChannel% > 0 Then
If sample(1).DisableQuantFlag%(sample(1).OxygenChannel%) = 0 Or (sample(1).DisableQuantFlag%(sample(1).OxygenChannel%) = 1 And sample(1).OxideOrElemental% = 1) Then
analysis.WtPercents!(sample(1).OxygenChannel%) = analysis.WtPercents!(sample(1).OxygenChannel%) + analysis.CalculatedOxygen!
analysis.totaloxygen! = analysis.WtPercents!(sample(1).OxygenChannel%)
End If
End If

' If measuring and displaying as oxide, display calculated, measured and excess oxygen
Else
If sample(1).OxygenChannel% > 0 And sample(1).DisplayAsOxideFlag Then
analysis.totaloxygen! = analysis.WtPercents!(sample(1).OxygenChannel%)
analysis.CalculatedOxygen! = ConvertOxygenFromCations3!(analysis, sample())
End If
End If

' Sum the total percents
For chan% = 1 To sample(1).LastChan%
If sample(1).DisableQuantFlag%(chan%) = 0 Or (sample(1).DisableQuantFlag%(chan%) = 1 And sample(1).OxideOrElemental% = 1 And sample(1).OxygenChannel% = chan%) Then
analysis.TotalPercent! = analysis.TotalPercent! + analysis.WtPercents!(chan%)
End If
Next chan%

' Sum total atoms, total cations
For chan% = 1 To sample(1).LastChan%
If sample(1).DisableQuantFlag%(chan%) = 0 Or (sample(1).DisableQuantFlag%(chan%) = 1 And sample(1).OxideOrElemental% = 1 And sample(1).OxygenChannel% = chan%) Then
If sample(1).AtomicWts!(chan%) = 0# Then GoTo ZAFCalZbarNoAtomicWeights
If sample(1).AtomicNums%(chan%) = 0 Then GoTo ZAFCalZbarNoAtomicNumbers
atoms!(chan%) = analysis.WtPercents!(chan%) / sample(1).AtomicWts!(chan%)
sum_atoms! = sum_atoms! + atoms!(chan%)
If sample(1).AtomicCharges!(chan%) > 0# Then sum_cations! = sum_cations! + atoms!(chan%)
analysis.AtomicNumbers!(chan%) = sample(1).AtomicNums%(chan%)
analysis.AtomicWeights!(chan%) = sample(1).AtomicWts!(chan%)
End If
Next chan%

' Determine formula element basis if specified
If sum_atoms! >= 0.01 And sample(1).FormulaElementFlag Then

' Element formula basis
If Trim$(sample(1).FormulaElement$) <> vbNullString And sample(1).FormulaRatio! <> 0# Then
ip% = IPOS1(sample(1).LastChan%, sample(1).FormulaElement$, sample(1).Elsyms$())
If ip% = 0 Then GoTo ZAFCalZbarInvalidFormulaElement
If atoms!(ip%) >= 0.01 Then
temp! = sample(1).FormulaRatio! / atoms!(ip%)
total_cations! = 0#
total_atoms! = 0#
For chan% = 1 To sample(1).LastChan%
basis!(chan%) = atoms!(chan%) * temp!
total_atoms! = total_atoms! + basis!(chan%)
If sample(1).AtomicCharges!(chan%) > 0# Then total_cations! = total_cations! + basis!(chan%)
Next chan%
End If

' Sum of cation basis (assume 8 cations if not specified)
Else
If sample(1).FormulaRatio! = 0# Then sample(1).FormulaRatio! = 8#
total_cations! = 0#
total_atoms! = 0#
temp! = 0#
If sum_cations! > 0# Then temp! = sample(1).FormulaRatio! / sum_cations!
For chan% = 1 To sample(1).LastChan%
basis!(chan%) = atoms!(chan%) * temp!
total_atoms! = total_atoms! + basis!(chan%)
If sample(1).AtomicNums%(chan%) <> 8 Then total_cations! = total_cations! + basis!(chan%)
Next chan%
End If
End If

' Return cations and atoms
analysis.TotalCations! = total_cations!
analysis.totalatoms! = total_atoms!

' Calculate average atomic weight and average atomic number (Zbar)
If analysis.TotalPercent! <> 0# And sum_atoms! <> 0# Then

' Calculate average atomic weight
For chan% = 1 To sample(1).LastChan%
'If sample(1).DisableQuantFlag%(chan%) = 0 Or (sample(1).DisableQuantFlag%(chan%) = 1 And sample(1).OxideOrElemental% = 1 And sample(1).OxygenChannel% = chan%) Then
analysis.AtomicWeight! = analysis.AtomicWeight! + sample(1).AtomicWts!(chan%) * atoms!(chan%) / sum_atoms!
'End If
Next chan%

' Calculate mass fraction Zbar
If Not UseZFractionZbarCalculationsFlag Then
analysis.zbar! = ConvertWeightsToZBar!(Int(0), sample(1).LastChan%, analysis.WtPercents!(), sample(1).AtomicNums%(), sample(1).AtomicWts!(), ZFractionZbarCalculationsExponent!)
If ierror Then Exit Sub

' Calculate Z fraction Zbar
Else
analysis.zbar! = ConvertWeightsToZBar!(Int(1), sample(1).LastChan%, analysis.WtPercents!(), sample(1).AtomicNums%(), sample(1).AtomicWts!(), ZFractionZbarCalculationsExponent!)
If ierror Then Exit Sub
End If

End If

' Calculate Excess Oxygen
analysis.ExcessOxygen! = analysis.totaloxygen! - analysis.CalculatedOxygen!

' Calculate oxygen equivalent of halogens
analysis.OxygenFromHalogens! = ConvertHalogensToOxygen!(sample(1).LastChan%, sample(1).Elsyms$(), sample(1).DisableQuantFlag%(), analysis.WtPercents!())

' Calculate halogen corrected oxygen (do not perform if oxygen is measured)
If Not UseOxygenFromHalogensCorrectionFlag And sample(1).OxygenChannel% > sample(1).LastElm% And sample(1).OxideOrElemental% = 1 Then
analysis.HalogenCorrectedOxygen! = analysis.totaloxygen! - analysis.OxygenFromHalogens!
Else
analysis.HalogenCorrectedOxygen! = analysis.totaloxygen!
End If

' Calculate charge balance
analysis.ChargeBalance! = ConvertChargeBalance!(sample(1).LastChan%, sample(1).AtomicWts!(), analysis.WtPercents!(), sample(1).AtomicCharges!())

' Load Fe charge
ip% = IPOS1(sample(1).LastChan%, Symlo$(ATOMIC_NUM_IRON%), sample(1).Elsyms$())   ' find Fe index
If ip% > 0 Then analysis.FeCharge! = sample(1).AtomicCharges!(ip%)

If VerboseMode Then
msg$ = "ELEMENT "
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%), a80$)
Next chan%
Call IOWriteLog(msg$)

msg$ = "UNK WT% "
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(analysis.WtPercents!(chan%), f84$), a80$)
Next chan%
Call IOWriteLog(msg$)
End If

Exit Sub

' Errors
ZAFCalZbarError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFCalZbar"
ierror = True
Exit Sub

ZAFCalZbarNoAtomicWeights:
msg$ = "The atomic weight for " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " is zero"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFCalZbar"
ierror = True
Exit Sub

ZAFCalZbarNoAtomicNumbers:
msg$ = "The atomic number for " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " is zero"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFCalZbar"
ierror = True
Exit Sub

ZAFCalZbarInvalidFormulaElement:
msg$ = "The formula basis element " & sample(1).Elsyms$(chan%) & " is not present in the sample"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFCalZbar"
ierror = True
Exit Sub

End Sub

Sub ZAFCalZbarLoadText(tForm As Form, analysis As TypeAnalysis)
' This routine loads the Total, Z-bar, etc. text boxes

ierror = False
On Error GoTo ZAFCalZbarLoadTextError

' Load the text fields
tForm.LabelTotal.Caption = Format$(Format$(analysis.TotalPercent!, f83), a80$)
tForm.LabelAtomic.Caption = Format$(Format$(analysis.AtomicWeight!, f83), a80$)

tForm.LabelZbar.Caption = Format$(Format$(analysis.zbar!, f83), a80$)
If Not UseZFractionZbarCalculationsFlag Then
tForm.LabelZBarText.Caption = "Z-Bar (mass frac.)"
Else
tForm.LabelZBarText.Caption = "Z-Bar (Z frac.)"
End If

' Special for Oxide standards
tForm.LabelCalculated.Caption = Format$(Format$(analysis.CalculatedOxygen!, f83), a80$)
tForm.LabelTotalOxygen.Caption = Format$(Format$(analysis.totaloxygen!, f83), a80$)
tForm.LabelExcess.Caption = Format$(Format$(analysis.ExcessOxygen!, f83), a80$)

' Check total and if below or above tolerance make red
tForm.LabelTotal.ForeColor = vbBlack
If analysis.TotalPercent! <> 0# Then
If analysis.TotalPercent! < AnalyticalTotalMinimum! Or analysis.TotalPercent! > AnalyticalTotalMaximum! Then
tForm.LabelTotal.ForeColor = vbRed
End If
End If

Exit Sub

' Errors
ZAFCalZbarLoadTextError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFCalZbarLoadText"
ierror = True
Exit Sub

End Sub

Sub ZAFCalculateRanges(mup() As Single, analysis As TypeAnalysis, sample() As TypeSample)
' This routine calculates electron and x-ray ranges

ierror = False
On Error GoTo ZAFCalculateRangesError

Const MAXCALCDENSITY% = 8   ' maximum density in gm/cm3

Dim i As Integer, k As Integer, n As Integer
Dim ii As Integer, jj As Integer, notransmission As Integer
Dim tRadius As Single, averageatomicweight As Single
Dim averageatomicnumber As Single
Dim astring As String

' Check for elements
If sample(1).LastElm% = 0 Then GoTo ZAFCalculateNoAnalyzedElements

' Check for analysis
If analysis.WtPercents!(1) = 0# Then GoTo ZAFCalculateNoAnalyzedConcentrations

' Calculate average atomic weight and number
For i% = 1 To sample(1).LastChan%
averageatomicweight! = averageatomicweight! + analysis.WtPercents!(i%) / 100# * sample(1).AtomicWts!(i%)
averageatomicnumber! = averageatomicnumber! + analysis.WtPercents!(i%) / 100# * sample(1).AtomicNums%(i%)
Next i%

' Electron ranges
astring$ = astring$ & vbCrLf & "Kanaya-Okayama Range (1972) at " & Str$(sample(1).kilovolts!) & " KeV, in microns:" & vbCrLf
For k% = 1 To MAXCALCDENSITY%     ' loop on densities 1 to MAXCALCDENSITY%
tRadius! = (0.0276 * averageatomicweight! * sample(1).kilovolts! ^ 1.67) / (k% * averageatomicnumber! ^ 0.89)
astring$ = astring$ & "Density =" & Str$(k%) & ", radius = " & MiscAutoFormat$(tRadius!) & vbCrLf
Next k%

Call IOWriteLog(astring)

' Xray ranges for analyzed elements
For k% = 1 To MAXCALCDENSITY%     ' loop on densities
notransmission = True

astring$ = vbCrLf & "Xray Transmission Efficiency for density = " & Str$(k%)
Call IOWriteLog(astring)

n% = 0
Do Until False
n% = n% + 1
Call TypeGetRange(Int(1), n%, ii%, jj%, sample())
If ierror Then Exit Sub
If ii% > sample(1).LastElm% Then Exit Do

' Elements
msg$ = "ELEM: "
For i% = ii% To jj%
msg$ = msg$ & Format$(sample(1).Elsyup$(i%), a80$)
Next i%
Call IOWriteLog(msg$)

' Xrays
msg$ = "XRAY: "
For i% = ii% To jj%
msg$ = msg$ & Format$(sample(1).Xrsyms$(i%), a80$)
Next i%
Call IOWriteLog(msg$)

' Transmission for different thicknesses (microns)
Call ZAFCalculateRanges2(notransmission%, Int(1), Int(8), Int(1), ii%, jj%, k%, mup!(), analysis, sample())
If ierror Then Exit Sub
Call ZAFCalculateRanges2(notransmission%, Int(10), Int(18), Int(2), ii%, jj%, k%, mup!(), analysis, sample())
If ierror Then Exit Sub
Call ZAFCalculateRanges2(notransmission%, Int(20), Int(45), Int(5), ii%, jj%, k%, mup!(), analysis, sample())
If ierror Then Exit Sub
Call ZAFCalculateRanges2(notransmission%, Int(50), Int(100), Int(10), ii%, jj%, k%, mup!(), analysis, sample())
If ierror Then Exit Sub
Call ZAFCalculateRanges2(notransmission%, Int(125), Int(200), Int(25), ii%, jj%, k%, mup!(), analysis, sample())
If ierror Then Exit Sub

Loop

If notransmission Then Exit For
Next k% ' next density

Exit Sub

' Errors
ZAFCalculateRangesError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFCalculateRanges"
ierror = True
Exit Sub

ZAFCalculateNoAnalyzedElements:
msg$ = "No analyzed elements, enter or load at least one analyzed element"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFCalculateRanges"
ierror = True
Exit Sub

ZAFCalculateNoAnalyzedConcentrations:
msg$ = "No analyzed concentrations, calculate the concentrations first and try again"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFCalculateRanges"
ierror = True
Exit Sub

End Sub

Sub ZAFCalculateRanges2(notransmission As Integer, lomicron As Integer, himicron As Integer, istep As Integer, ii As Integer, jj As Integer, k As Integer, mup() As Single, analysis As TypeAnalysis, sample() As TypeSample)
' Calculate x-ray transmission for a range of microns

ierror = False
On Error GoTo ZAFCalculateRanges2Error

Const MINTRANSMISSION! = 0.005
Const MAXTRANSMISSION! = 0.01

Dim i As Integer, j As Integer
Dim l As Integer
Dim tmsg As String
Dim averagemassabsorption As Single
Dim transmission As Single
Dim hightransmission As Single

For j% = lomicron% To himicron Step istep%
hightransmission! = 0#

tmsg$ = "THIC: "
For i% = ii% To jj%
tmsg$ = tmsg$ & Format$(j%, a80$)
Next i%

msg$ = "TRAN: "
For i% = ii% To jj%

' Calculate average mass absorption for this emitter
averagemassabsorption! = 0#
For l% = 1 To sample(1).LastChan%
averagemassabsorption! = averagemassabsorption! + analysis.WtPercents!(l%) / 100# * mup!(l%, i%)
Next l%

' Calculate x-ray transmission
transmission! = NATURALE# ^ (-1# * averagemassabsorption! * CSng(k%) * CSng(j%) * CMPERMICRON#)

' Check
If lomicron% > 5 And transmission! > MINTRANSMISSION! Then notransmission = False
If transmission! > hightransmission! Then hightransmission! = transmission!

msg$ = msg$ & MiscAutoFormat$(transmission!)
Next i%

If hightransmission! < MAXTRANSMISSION! Then Exit For
Call IOWriteLog(tmsg$)
Call IOWriteLog(msg$)

Next j% ' next transmission

Exit Sub

' Errors
ZAFCalculateRanges2Error:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFCalculateRanges2"
ierror = True
Exit Sub

End Sub

Sub ZAFGetOxygenChannel(sample() As TypeSample)
' Determines the channel that is oxygen

ierror = False
On Error GoTo ZAFGetOxygenChannelError

Dim ip As Integer

' Calculate "sample(1).OxygenChannel%" for use in calculating elemental to oxide conversions
sample(1).OxygenChannel% = 0
'ip% = IPOS1(sample(1).LastChan%, Symlo$(ATOMIC_NUM_OXYGEN%), sample(1).Elsyms$())
ip% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_OXYGEN%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())      ' to handle analyzed oxygen disabled for quant (06/06/2022)
sample(1).OxygenChannel% = ip%

Exit Sub

' Errors
ZAFGetOxygenChannelError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFGetOxygenChannel"
ierror = True
Exit Sub

End Sub

Function ZAFConvertExcessOxygenToHydrogen(conc() As Single, zaf As TypeZAF, sample() As TypeSample) As Single
' Calculate the equivalent hydrogen based on the excess oxygen concentration

ierror = False
On Error GoTo ZAFConvertExcessOxygenToHydrogenError

Dim j As Integer
Dim temp As Single, p1 As Single
Dim stoichiometricoxygen As Single
Dim alloxygen As Single, extraoxygen As Single

' Calculate stoichiometric oxygen
stoichiometricoxygen! = 0#
For j% = 1 To zaf.in1%
p1! = sample(1).numoxd%(j%) / CSng(sample(1).numcat%(j%)) * AllAtomicWts!(ATOMIC_NUM_OXYGEN%) / sample(1).AtomicWts!(j%)
If j% <> sample(1).OxygenChannel% Then
stoichiometricoxygen! = stoichiometricoxygen! + conc!(j%) * p1!
End If
Next j%

' If calculating oxygen, add oxygen from specified concentration
If zaf.il%(zaf.in0%) = 0 Then
alloxygen! = stoichiometricoxygen!

If sample(1).OxygenChannel% > sample(1).LastElm% Then
alloxygen! = alloxygen! + sample(1).ElmPercents!(sample(1).OxygenChannel%) / 100#
End If

' If measuring oxygen, load analyzed oxygen
Else
If sample(1).OxygenChannel% > 0 Then alloxygen! = conc!(sample(1).OxygenChannel%)
End If

' Calculate excess oxygen
extraoxygen! = alloxygen! - stoichiometricoxygen!

' Convert to hydrogen if there is excess oxygen (allow negative correction)
temp! = 0#
If extraoxygen! > 0# Or (extraoxygen! < 0# And Not ForceNegativeKratiosToZeroFlag) Then
temp! = extraoxygen! / AllAtomicWts!(ATOMIC_NUM_OXYGEN%) * sample(1).HydrogenStoichiometryRatio! * AllAtomicWts!(ATOMIC_NUM_HYDROGEN%)
End If

ZAFConvertExcessOxygenToHydrogen! = temp!
Exit Function

' Errors
ZAFConvertExcessOxygenToHydrogenError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFConvertExcessOxygenToHydrogen"
ierror = True
Exit Function

End Function

Sub ZAFCalculatePhiRhoZCurves(i As Integer, Alpha As Single, beta As Single, gamma0 As Single, chi As Single, phi_0_mix As Single, zaf As TypeZAF)
' Generate phi(rho*z) curves for the passed element (both generated and emitted) for plot display (Bastin, Armstrong, Proza, etc)
' Calculation code provided by Brian Joy (2018)
'   i = element array index

ierror = False
On Error GoTo ZAFCalculatePhiRhoZCurvesError

Const eps! = 0.00001          ' tolerance for Simpson's rule

Dim n_rhoz As Integer
Dim nemax As Long, n As Long, it As Long, k As Long
Dim phi As Single, ost As Single, os As Single, rhoz As Single
Dim a As Single, b As Single, phi1 As Single, phi2 As Single
Dim sum As Single, tnm As Single, st As Single, ss As Single
Dim Step As Single, r_limit As Single, del As Single, i_g_mix As Single

' Specify number of steps to calculate intensities
n_rhoz% = 200

' Dimension global arrays if first element
If i% = 1 Then

' Determine how many emitting elements there actually are
n& = 0
For k& = 1 To zaf.in1%
If zaf.il%(k&) <= 12 Then n& = n& + 1
Next k&
If n& = 0 Then GoTo ZAFCalculatePhiRhoZCurvesNoElements

ReDim PhiRhoZPlotX(1 To n&, 1 To n_rhoz% + 1) As Single       ' z depth
ReDim PhiRhoZPlotY1(1 To n&, 1 To n_rhoz% + 1) As Single      ' generated intensities
ReDim PhiRhoZPlotY2(1 To n&, 1 To n_rhoz% + 1) As Single      ' emitted intensities

PhiRhoZPlotSets& = n&
PhiRhoZPlotPoints& = n_rhoz% + 1
End If

' Calculate i_g_mix
i_g_mix = (Sqr(PI!) / (2# * Alpha)) * (gamma0 - (gamma0 - phi_0_mix) * ZAFErrorFunction(beta / (2# * Alpha)))
    
If VerboseMode Then
Call IOWriteLog(vbNullString)
Call IOWriteLog("ZAFCalculatePhiRhoZCurves: Emitter = " & Symlo$(zaf.Z(i)) & " " & Xraylo$(zaf.il(i)))
Call IOWriteLog("Armstrong/Love-Scott mixture parameters:")
Call IOWriteLog("Alpha = " & Format$(Alpha!))
Call IOWriteLog("Beta = " & Format$(beta!))
Call IOWriteLog("Gamma0 = " & Format$(gamma0!))
Call IOWriteLog("Alpha = " & Format$(Alpha!))
Call IOWriteLog("Chi = " & Format$(chi!))
Call IOWriteLog("phi_0_mix = " & Format$(phi_0_mix!))
Call IOWriteLog("i_g_mix = " & Format$(i_g_mix!))
Call IOWriteLog(vbNullString)
End If

    rhoz = 0#
    
    ' Find rho*z at which phi(rho*z) decreases to a small value and then set the rho*z limit according to this
'    Do
'    phi = gamma0 * Exp(-1# * Alpha ^ 2# * rhoz ^ 2#) * (1# - ((gamma0 - phi_0_mix) / gamma0) * Exp(-1# * beta * rhoz))
'    If (phi < 0.01) Then Exit Do
'    rhoz = rhoz + 0.00001                              ' [g/cm**2]
'    Loop

    ' Use Simpson's rule to find rhoz based on .999 of area under the curve
    nemax = 20
    
    Do
    a = 0#
    b = rhoz
    ost = -1E+30
    os = -1E+30
        For n = 1 To nemax
            If n = 1 Then
               phi1 = gamma0 * Exp(-1# * Alpha ^ 2 * a ^ 2) * (1# - ((gamma0 - phi_0_mix) / gamma0) * Exp(-1# * beta * a))
               phi2 = gamma0 * Exp(-1# * Alpha ^ 2 * b ^ 2) * (1# - ((gamma0 - phi_0_mix) / gamma0) * Exp(-1# * beta * b))
               st = 0.5 * (b - a) * (phi1 + phi2)
            Else
               it = 2 ^ (n - 2)
               tnm = it
               del = (b - a) / tnm
               rhoz = a + 0.5 * del
               sum = 0#
               For k = 1 To it
                  sum = sum + gamma0 * Exp(-1# * Alpha ^ 2 * rhoz ^ 2) * (1# - ((gamma0 - phi_0_mix) / gamma0) * Exp(-1# * beta * rhoz))
                  rhoz = rhoz + del
               Next k
               st = (st + (b - a) * sum / tnm) / 2#
            End If
            
            ss = (4# * st - ost) / 3#
            'If VerboseMode Then Call IOWriteLog("n = " & Format$(n) & ", rhoz = " & Format$(rhoz) & ", ss = " & Format$(ss))
            'If VerboseMode Then Call IOWriteLog("sum = " & Format$(sum))
            If (Abs(ss - os) < eps * Abs(os)) Then Exit For
            os = ss
            ost = st
         Next n
         
         If ss / i_g_mix > 0.999 Then Exit Do
         rhoz = rhoz + 0.00001                        ' [g/cm**2]
    Loop
      
    r_limit = rhoz
    Step = r_limit / n_rhoz                             ' n_rhoz is number of intensities to calculate
    rhoz = 0#
      
    If VerboseMode Then
    Call IOWriteLog("R_limit[mg/cm**2] = " & Format$(r_limit * 1000#))
    Call IOWriteLog(vbNullString)
    Call IOWriteLog("rho*z[mg/cm**2], phi(rho*z), phi(rho*z)*EXP(-chi*rhoz):")
    End If
      
    For k = 1 To n_rhoz + 1
        phi = gamma0 * Exp(-1# * Alpha ^ 2 * rhoz ^ 2) * (1# - ((gamma0 - phi_0_mix) / gamma0) * Exp(-1# * beta * rhoz))
        If VerboseMode Then Call IOWriteLog("rhoz = " & MiscAutoFormat$(CSng(rhoz * 1000#)) & ", phi = " & MiscAutoFormat$(CSng(phi)) & ", phi' = " & MiscAutoFormat$(CSng(phi * Exp(-1 * chi * rhoz))))
        rhoz = rhoz + Step
        
        ' Load into global arrays for plotting
        PhiRhoZPlotX(i%, k) = rhoz * 1000#
        PhiRhoZPlotY1(i%, k) = phi
        PhiRhoZPlotY2(i%, k) = phi * Exp(-1 * chi * rhoz)
        
    Next k
    If VerboseMode Then Call IOWriteLog(vbNullString)

Exit Sub

' Errors
ZAFCalculatePhiRhoZCurvesError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFCalculatePhiRhoZCurves"
ierror = True
Exit Sub

ZAFCalculatePhiRhoZCurvesNoElements:
msg$ = "No analyzed elements for this sample"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFCalculatePhiRhoZCurves"
ierror = True
Exit Sub

End Sub

Sub ZAFCalculatePhiRhoZCurvesPAP(i As Integer, rm_mix As Double, rc_mix As Double, rx_mix As Double, a1 As Double, a2 As Double, b1 As Double, chi As Single, zaf As TypeZAF)
' Generate phi(rho*z) curves for the passed element (both generated and emitted) for plot display (PAP only)
' Calculation code provided by Brian Joy (2018)
'   i = element array index

ierror = False
On Error GoTo ZAFCalculatePhiRhoZCurvesPAPError

Dim n_rhoz As Integer
Dim n As Long, k As Long
Dim phi As Double, rhoz As Double
Dim Step As Double

' Specify number of steps to calculate intensities
n_rhoz% = 200

' Dimension global arrays if first element
If i% = 1 Then

' Determine how many emitting elements there actually are
n& = 0
For k& = 1 To zaf.in1%
If zaf.il%(k&) <= 12 Then n& = n& + 1
Next k&
If n& = 0 Then GoTo ZAFCalculatePhiRhoZCurvesPAPNoElements

ReDim PhiRhoZPlotX(1 To n&, 1 To n_rhoz% + 1) As Single       ' z depth
ReDim PhiRhoZPlotY1(1 To n&, 1 To n_rhoz% + 1) As Single      ' generated intensities
ReDim PhiRhoZPlotY2(1 To n&, 1 To n_rhoz% + 1) As Single      ' emitted intensities

PhiRhoZPlotSets& = n&
PhiRhoZPlotPoints& = n_rhoz% + 1
End If
    
If VerboseMode Then
Call IOWriteLog(vbNullString)
Call IOWriteLog("ZAFCalculatePhiRhoZCurvesPAP: Emitter = " & Symlo$(zaf.Z(i)) & " " & Xraylo$(zaf.il(i)))
Call IOWriteLog("PAP mixture parameters:")
Call IOWriteLog("Rm[mg/cm**2] = " & Format$(CSng(rm_mix# * 1000#)))
Call IOWriteLog("Rc[mg/cm**2] = " & Format$(CSng(rc_mix# * 1000#)))
Call IOWriteLog("Rx[mg/cm**2] = " & Format$(CSng(rx_mix# * 1000#)))
Call IOWriteLog("A1 = " & Format$(CSng(a1#)))
Call IOWriteLog("A2 = " & Format$(CSng(a2#)))
Call IOWriteLog("phi(Rm) = " & Format$(CSng(b1#)))
Call IOWriteLog("Chi = " & Format$(chi!))
End If
        
If VerboseMode Then
    Call IOWriteLog(vbNullString)
    Call IOWriteLog("rho*z[mg/cm**2], phi(rho*z), phi(rho*z)*EXP(-chi*rhoz):")
End If
      
    ' Generate phi(rho*z) curve
    Step = rx_mix / n_rhoz
    rhoz = 0#
      
    For k = 1 To n_rhoz + 1
         If (rhoz < rc_mix) Then
            phi = a1 * (rhoz - rm_mix) ^ 2 + b1
         Else
            phi = a2 * (rhoz - rx_mix) ^ 2
         End If
        
        If VerboseMode Then Call IOWriteLog("rhoz = " & MiscAutoFormat$(CSng(rhoz * 1000#)) & ", phi = " & MiscAutoFormat$(CSng(phi)) & ", phi' = " & MiscAutoFormat$(CSng(phi * Exp(-1 * chi * rhoz))))
        rhoz = rhoz + Step
      
        ' Load into global arrays for plotting
        PhiRhoZPlotX(i%, k) = rhoz * 1000#
        PhiRhoZPlotY1(i%, k) = phi
        PhiRhoZPlotY2(i%, k) = phi * Exp(-1 * chi * rhoz)
        
    Next k
    Call IOWriteLog(vbNullString)

Exit Sub

' Errors
ZAFCalculatePhiRhoZCurvesPAPError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFCalculatePhiRhoZCurvesPAP"
ierror = True
Exit Sub

ZAFCalculatePhiRhoZCurvesPAPNoElements:
msg$ = "No analyzed elements for this sample"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFCalculatePhiRhoZCurvesPAP"
ierror = True
Exit Sub

End Sub

Sub ZAFCalculatePhiRhoZCurvesXPP(i As Integer, a As Double, b As Double, aa As Double, bb As Double, phi_0_mix As Double, FZ_mix As Double, chi As Single, zaf As TypeZAF)
' Generate phi(rho*z) curves for the passed element (both generated and emitted) for plot display (XPP only)
' Calculation code provided by Brian Joy (2018)
'   i = element array index

ierror = False
On Error GoTo ZAFCalculatePhiRhoZCurvesXPPError

Const eps! = 0.00001          ' tolerance for Simpson's rule

Dim n_rhoz As Integer
Dim n As Long, k As Long, nemax As Long, it As Long
Dim aaa As Single, bbb As Single, ost As Single, os As Single, st As Single
Dim phi As Single, phi1 As Single, phi2 As Single, rhoz As Double, tnm As Single
Dim Step As Single, del As Single, ssum As Single, ss As Single, r_limit As Single

' Specify number of steps to calculate intensities
n_rhoz% = 200

' Dimension global arrays if first element
If i% = 1 Then

' Determine how many emitting elements there actually are
n& = 0
For k& = 1 To zaf.in1%
If zaf.il%(k&) <= 12 Then n& = n& + 1
Next k&
If n& = 0 Then GoTo ZAFCalculatePhiRhoZCurvesXPPNoElements

ReDim PhiRhoZPlotX(1 To n&, 1 To n_rhoz% + 1) As Single       ' z depth
ReDim PhiRhoZPlotY1(1 To n&, 1 To n_rhoz% + 1) As Single      ' generated intensities
ReDim PhiRhoZPlotY2(1 To n&, 1 To n_rhoz% + 1) As Single      ' emitted intensities

PhiRhoZPlotSets& = n&
PhiRhoZPlotPoints& = n_rhoz% + 1
End If
    
If VerboseMode Then
Call IOWriteLog(vbNullString)
Call IOWriteLog("ZAFCalculatePhiRhoZCurvesXPP: Emitter = " & Symlo$(zaf.Z(i)) & " " & Xraylo$(zaf.il(i)))
Call IOWriteLog("XPP mixture parameters:")
Call IOWriteLog("A = " & Format$(CSng(a#)))
Call IOWriteLog("B = " & Format$(CSng(b#)))
Call IOWriteLog("AA = " & Format$(CSng(aa#)))
Call IOWriteLog("BB = " & Format$(CSng(bb#)))
Call IOWriteLog("Phi_0_mix = " & Format$(CSng(phi_0_mix#)))
Call IOWriteLog("FZ_mix = " & Format$(CSng(FZ_mix#)))
Call IOWriteLog("Chi = " & Format$(chi!))
Call IOWriteLog(vbNullString)
End If
             
      ' Find rho*z at which phi(rho*z) decreases to a small value and then set the rho*z limit according to this
'      rhoz = 0#
'      Do
'      phi = a * Exp(-1# * aa * rhoz) + (b * rhoz + phi_0_mix - a) * Exp(-1# * bb * rhoz)
'      If (phi < 0.01) Then Exit Do
'      rhoz = rhoz + 0.00001                              ' [g/cm**2]
'      Loop
             
      ' Use Simpson's rule to find rhoz based on .999 of area under the curve
      rhoz = 0#
      nemax = 20
      Do
         aaa = 0#
         bbb = rhoz
         ost = -1E+30
         os = -1E+30
         For n = 1 To nemax
            If n = 1 Then
               phi1 = a * Exp(-1# * aa * aaa) + (b * aaa + phi_0_mix - a) * Exp(-1# * bb * aaa)
               phi2 = a * Exp(-1# * aa * bbb) + (b * bbb + phi_0_mix - a) * Exp(-1# * bb * bbb)
               st = 0.5 * (bbb - aaa) * (phi1 + phi2)
            Else
               it = 2 ^ (n - 2)
               tnm = it
               del = (bbb - aaa) / tnm
               rhoz = aaa + 0.5 * del
               ssum = 0#
               For k = 1 To it
                  ssum = ssum + a * Exp(-1# * aa * rhoz) + (b * rhoz + phi_0_mix - a) * Exp(-1# * bb * rhoz)
                  rhoz = rhoz + del
               Next k&
               st = (st + (bbb - aaa) * ssum / tnm) / 2#
            End If
            
            ss = (4# * st - ost) / 3#
            'If VerboseMode Then Call IOWriteLog("n = " & Format$(n) & ", rhoz = " & Format$(rhoz) & ", ss = " & Format$(ss))
            'If VerboseMode Then Call IOWriteLog("ssum = " & Format$(ssum))
            If (Abs(ss - os) < eps * Abs(os)) Then Exit For
            os = ss
            ost = st
         Next n&
         If (ss / FZ_mix > 0.999) Then Exit Do
         rhoz = rhoz + 0.00001                        ' [g/cm**2]
      Loop
      
    r_limit = rhoz

    If VerboseMode Then
    Call IOWriteLog("R_limit[mg/cm**2] = " & Format$(r_limit * 1000#))
    Call IOWriteLog(vbNullString)
    Call IOWriteLog("rho*z[mg/cm**2], phi(rho*z), phi(rho*z)*EXP(-chi*rhoz):")
    End If

    ' Generate phi(rho*z) curve
    Step = r_limit / n_rhoz
    rhoz = 0#
      
    For k = 1 To n_rhoz + 1
        phi = a * Exp(-1# * aa * rhoz) + (b * rhoz + phi_0_mix - a) * Exp(-1# * bb * rhoz)
        If VerboseMode Then Call IOWriteLog("rhoz = " & MiscAutoFormat$(CSng(rhoz * 1000#)) & ", phi = " & MiscAutoFormat$(CSng(phi)) & ", phi' = " & MiscAutoFormat$(CSng(phi * Exp(-1 * chi * rhoz))))
        rhoz = rhoz + Step
      
        ' Load into global arrays for plotting
        PhiRhoZPlotX(i%, k) = rhoz * 1000#
        PhiRhoZPlotY1(i%, k) = phi
        PhiRhoZPlotY2(i%, k) = phi * Exp(-1 * chi * rhoz)
        
    Next k
    Call IOWriteLog(vbNullString)

Exit Sub

' Errors
ZAFCalculatePhiRhoZCurvesXPPError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFCalculatePhiRhoZCurvesXPP"
ierror = True
Exit Sub

ZAFCalculatePhiRhoZCurvesXPPNoElements:
msg$ = "No analyzed elements for this sample"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFCalculatePhiRhoZCurvesXPP"
ierror = True
Exit Sub

End Sub
