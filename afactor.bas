Attribute VB_Name = "CodeAFACTOR"
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

' Alpha factors have "lastelm" rows, corresponding to the emitting elements
' and "lastchan" columns, corresponding to the absorbing or fluorescing
' elements that are present in the sample.
'
' In the alpha arrays, the row or record is the radiation, while the
' column is the absorbing or fluorescing matrix element. The
' record and column a certain element occupies is determined by its
' atomic number.

' Sample for alpha-factor binary k-factor calculations
Dim AFactorTmpSample(1 To 1) As TypeSample
Dim FormulaTmpSample(1 To 1) As TypeSample

' Dimension here to return values for plotting
Dim AlphaNpts1 As Integer
Dim AlphaXdata1() As Single
Dim AlphaYdata1() As Single
Dim AlphaAcoeff1() As Single
Dim AlphaAstddev1 As Single

Dim AlphaNpts2 As Integer
Dim AlphaXdata2() As Single
Dim AlphaYdata2() As Single
Dim AlphaAcoeff2() As Single
Dim AlphaAstddev2 As Single

' Empirical alpha factors (elemental only)
Dim empfackev(1 To MAXEMPFAC%) As Single
Dim empfactof(1 To MAXEMPFAC%) As Single
Dim empfaco(1 To MAXEMPFAC%) As Integer             ' alpha regression mode (1, 2, 3 or 4) (constant, linear, polynomial or non-linear)

Dim empface(1 To MAXEMPFAC%) As String      ' emitter
Dim empfacx(1 To MAXEMPFAC%) As String      ' xray
Dim empfaca(1 To MAXEMPFAC%) As String      ' absorber

Dim empfac1(1 To MAXEMPFAC%) As Single
Dim empfac2(1 To MAXEMPFAC%) As Single
Dim empfac3(1 To MAXEMPFAC%) As Single
Dim empfac4(1 To MAXEMPFAC%) As Single

Dim empfacstr(1 To MAXEMPFAC%) As String ' string

' Operating voltage and takeoff for alpha-factors look-up tables (combined condition samples not supported yet!)
Dim alphakev As Single, alphatof As Single

' X-rays for alpha-factor look up tables (blank = not calculated)
Dim AlphaXray(1 To MAXELM%, 1 To MAXELM%) As String                 ' emitter x-ray

' Alpha-factor look-up tables
Dim alphal1(1 To MAXELM%, 1 To MAXELM%) As Single
Dim alphal2(1 To MAXELM%, 1 To MAXELM%) As Single
Dim alphal3(1 To MAXELM%, 1 To MAXELM%) As Single
Dim alphal4(1 To MAXELM%, 1 To MAXELM%) As Single

' Alpha-factor sample arrays used for quantitative calculations
Dim alpha1(1 To MAXCHAN%, 1 To MAXCHAN%) As Single
Dim alpha2(1 To MAXCHAN%, 1 To MAXCHAN%) As Single
Dim alpha3(1 To MAXCHAN%, 1 To MAXCHAN%) As Single
Dim alpha4(1 To MAXCHAN%, 1 To MAXCHAN%) As Single

Dim PenepmaKratiosFlag(1 To MAXCHAN%, 1 To MAXCHAN%) As Boolean

Sub AFactorBeta(analysis As TypeAnalysis, sample() As TypeSample)
' This routine accepts an array of weight percents and returns an array of beta factors

ierror = False
On Error GoTo AFactorBetaError

Dim i As Integer, ip As Integer
Dim emitter As Integer, absorber As Integer
Dim betafraction As Single
Dim astring As String

ReDim wtfractions(1 To MAXCHAN%) As Single

' Convert to weight fractions
If analysis.TotalPercent! = 0# Then GoTo AFactorBetaBadTotal
For i% = 1 To sample(1).LastChan%
wtfractions!(i%) = analysis.WtPercents!(i%) / analysis.TotalPercent!
Next i%

If VerboseMode Then
Call IOWriteLog(vbNullString)
End If

' Factors for "LastElm" emitters and "LastChan" absorbers
For emitter% = 1 To sample(1).LastElm%
analysis.UnkBetas!(emitter%) = 0#

' Skip if quant is disabled
If sample(1).DisableQuantFlag(emitter%) = 0 Then

' Determine if element is duplicated
ip% = IPOS8(emitter%, sample(1).Elsyms$(emitter%), sample(1).Xrsyms$(emitter%), sample())
If Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0) Then

' Calculate beta factors (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
For absorber% = 1 To sample(1).LastChan%

' Calculate constant, linear and polynomial beta factors
If CorrectionFlag% < 4 Then
betafraction! = (alpha1!(emitter%, absorber%) + wtfractions!(absorber%) * alpha2!(emitter%, absorber%) + wtfractions!(absorber%) ^ 2 * alpha3!(emitter%, absorber%)) * wtfractions!(absorber%)

' Calculate non-linear beta factors
Else
betafraction! = (alpha1!(emitter%, absorber%) + wtfractions!(absorber%) * alpha2!(emitter%, absorber%) + wtfractions!(absorber%) ^ 2 * alpha3!(emitter%, absorber%) + Exp(wtfractions!(absorber%)) * alpha4!(emitter%, absorber%)) * wtfractions!(absorber%)
End If

' Debug mode
If VerboseMode Then
astring$ = "AFactorBeta: " & sample(1).Elsyup$(emitter%) & " " & sample(1).Xrsyms$(emitter%) & " in " & sample(1).Elsyup$(absorber%) & ", betafraction=" & Format$(betafraction!)
Call IOWriteLog(astring$)
End If

' Sum beta factors fractions
analysis.UnkBetas!(emitter%) = analysis.UnkBetas!(emitter%) + betafraction!

Next absorber%
End If
End If

' Debug mode
If VerboseMode Then
astring$ = "AFactorBeta: " & sample(1).Elsyup$(emitter%) & " " & sample(1).Xrsyms$(emitter%) & ", Total Beta=" & Format$(analysis.UnkBetas!(emitter%))
Call IOWriteLog(astring$)
End If
Next emitter%

Exit Sub

' Errors
AFactorBetaError:
MsgBox Error$, vbOKOnly + vbCritical, "AFactorBeta"
ierror = True
Exit Sub

AFactorBetaBadTotal:
msg$ = "Total percent equals zero, cannot calculate beta factors"
MsgBox msg$, vbOKOnly + vbExclamation, "AFactorBeta"
ierror = True
Exit Sub

End Sub

Sub AFactorCalculateFactor(i As Integer, j As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' Calculate the specified or all alpha-factors for the sample
' i = emitter, j = absorber position of binary in sample()

ierror = False
On Error GoTo AFactorCalculateFactorError

Dim emitter As Integer, absorber As Integer
Dim numberofbinaries As Integer, inum As Integer

' If emitter and absorber are zero, calculate all binaries in sample
icancel = False
If i% = 0 And j% = 0 Then

' Calculate total number of binaries
numberofbinaries% = 0
For emitter% = 1 To sample(1).LastElm%
For absorber% = emitter% + 1 To sample(1).LastChan%
numberofbinaries% = numberofbinaries% + 1
Next absorber%
Next emitter%

Call IOWriteLog("Number of alpha-factor binaries to be calculated = " & Str$(numberofbinaries%))
Call AnalyzeStatusAnal(vbNullString)

' Calculate each binary in sample (all elements)
inum% = 0
For emitter% = 1 To sample(1).LastElm%
If sample(1).DisableQuantFlag(emitter%) = 0 Then
For absorber% = emitter% + 1 To sample(1).LastChan%
inum% = inum% + 1

msg$ = "Calculating alpha-factor binary " & Format$(inum%, a50$) & " of " & Format$(numberofbinaries%, a50$) & "..."
Call AnalyzeStatusAnal(msg$)
If icancel Then
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub
End If

' K-factor calculation
Screen.MousePointer = vbHourglass
Call AFactorCalculateKFactors(emitter%, absorber%, analysis, sample())
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

Next absorber%
End If
Next emitter%
Call AnalyzeStatusAnal(vbNullString)

' Save alpha conditions for next time
alphakev! = sample(1).kilovolts!
alphatof! = sample(1).takeoff!

' Calculate a single binary in sample
Else
emitter% = i%
absorber% = j%

msg$ = "Calculating alpha-factor binary " & sample(1).Elsyms$(emitter%) & " " & sample(1).Xrsyms$(emitter%) & " in " & sample(1).Elsyms$(absorber%) & "..."
Call AnalyzeStatusAnal(msg$)
If icancel Then
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub
End If

' K-factor calculation
Screen.MousePointer = vbHourglass
Call AFactorCalculateKFactors(emitter%, absorber%, analysis, sample())
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End If

Call AnalyzeStatusAnal(vbNullString)
Call IOStatusAuto(vbNullString)
Exit Sub

' Errors
AFactorCalculateFactorError:
MsgBox Error$, vbOKOnly + vbCritical, "AFactorCalculateFactor"
Call AnalyzeStatusAnal(vbNullString)
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub AFactorCalculateFitFactors(k As Integer, l As Integer, wout() As Single, kout() As Single, eout() As String, xout() As String, zout() As Integer)
' Calculate the alpha-factor from the passed concentration and intensity arrays

ierror = False
On Error GoTo AFactorCalculateFitFactorsError

Dim m As Integer, ibin As Integer, p As Integer
Dim kmax As Integer, nmax As Integer, npts As Integer
Dim temp As Single
Dim astring As String

Dim syme As String, symx As String, symm As String
Dim rec1 As Integer, rec2 As Integer
Dim stddev As Single
Dim alph1 As Single, alph2 As Single, alph3 As Single, alph4 As Single

ReDim weightfractions(1 To MAXBINARY% * 2) As Single

ReDim xdata(1 To MAXBINARY%) As Single
ReDim ydata(1 To MAXBINARY%) As Single
ReDim acoeff(1 To MAXCOEFF4%) As Single

ReDim AlphaXdata1(1 To MAXBINARY%) As Single
ReDim AlphaYdata1(1 To MAXBINARY%) As Single
ReDim AlphaAcoeff1(1 To MAXCOEFF4%) As Single

ReDim AlphaXdata2(1 To MAXBINARY%) As Single
ReDim AlphaYdata2(1 To MAXBINARY%) As Single
ReDim AlphaAcoeff2(1 To MAXCOEFF4%) As Single

' 0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters
If DebugMode Then
Call IOWriteLog(vbNullString)
If CorrectionFlag% = 1 Then astring$ = " CONSTANT Alpha Factors"
If CorrectionFlag% = 2 Then astring$ = " LINEAR Alpha Factors"
If CorrectionFlag% = 3 Then astring$ = " POLYNOMIAL Alpha Factors"
If CorrectionFlag% = 4 Then astring$ = " NON-LINEAR Alpha Factors"
astring$ = astring$ & ", Takeoff=" & Str$(DefaultTakeOff!) & ", KeV=" & Str$(DefaultKiloVolts!)
Call IOWriteLog(astring$)
End If

' Convert to weight fraction
For m% = 1 To MAXBINARY% * 2
weightfractions!(m%) = wout!(m%) / 100#
Next m%

' Calculate alpha-factors for both binary components. "ibin" is the emitting element index in the binary calculation.
For ibin% = 1 To 2

' Init labels for each binary
syme$ = vbNullString            ' emitting element
symx$ = vbNullString            ' emitting x-ray
symm$ = vbNullString            ' matrix element
rec1% = 0
rec2% = 0

' Skip hydrogen emitter and helium emitter components (check first binary only)
If LCase$(eout$(ibin%)) = Symlo$(ATOMIC_NUM_HYDROGEN%) Or LCase$(eout$(ibin%)) = Symlo$(ATOMIC_NUM_HELIUM%) Then GoTo 600

' Calculate the nominal alpha-factor at the given concentrations and perform a least squares fit based on "CorrectionFlag%".
'
' The alpha-factors use the following alpha expression. Constant alpha factors are based on a linear fit of alpha vs. concentration, and
' back calculated for a 50/50 concentration. Polynomial alpha-factors are fit to a 2nd order polynomial. Non-linear alpha-factors are fit to
' a non-linear equation.
'
'         (C/K - C)/(1 - C) = alpha    or   alpha vs. C
'
' Load weight fractions and k-ratios into arrays to fit
npts% = 0
If DebugMode Then Call IOWriteLog(vbNullString)
For p% = 1 To MAXBINARY%
xdata!(p%) = 0#
ydata!(p%) = 0#
m% = (p% - 1) * 2 + ibin%

' Check for absorber only
If Trim$(xout$(m%)) <> vbNullString Then

temp! = ((weightfractions!(m%) / kout!(m%)) - weightfractions!(m%)) / (1# - weightfractions!(m%))

' Only fit points with positive k-ratios and alpha factors
If kout!(m%) > 0# And temp! > 0# Then
npts% = npts% + 1
xdata!(npts%) = weightfractions!(m%)
ydata!(npts%) = temp!

If DebugMode% Then
msg$ = "P=" & Format$(p%) & ", C=" & Format$(weightfractions!(m%), f84$) & ", K=" & Format$(kout!(m%), f84$) & ", Alpha=" & Format$(ydata!(npts%), f84$)
Call IOWriteLog(msg$)
End If

' Load labels and alpha array index numbers if not yet loaded (this is in case using PenepmaKratioLimits)
If syme$ = vbNullString Then
syme$ = eout$(m%)
symx$ = xout$(m%)
symm$ = eout$(m% + 1)
rec1% = zout%(m%)
rec2% = zout%(m% + 1)
End If

End If
End If

Next p%

' Check for at least two data points
If npts% < 2 Then GoTo 600
nmax% = npts%

' Constant and linear alpha factors (linear fit)
If CorrectionFlag% = 1 Or CorrectionFlag% = 2 Then
kmax% = 1   ' first order
Call LeastSquares(kmax%, nmax%, xdata!(), ydata!(), acoeff!())                  ' constant or linear regression
If ierror Then Exit Sub

' Calculate a constant alpha-factor at 50/50 composition
If CorrectionFlag% = 1 Then
alph1! = acoeff!(1) + acoeff!(2) * 0.5
alph2! = 0#
alph3! = 0#

' Calculate average % deviation for constant alpha-factor
acoeff(2) = 0#      ' to force constant
Call LeastDeviation(Int(1), stddev!, nmax%, xdata!(), ydata!(), acoeff!())       ' average deviation
If ierror Then Exit Sub
End If

' Linear alpha-factor fit
If CorrectionFlag% = 2 Then
alph1! = acoeff!(1)
alph2! = acoeff!(2)
alph3! = 0#

' Calculate average % deviation of linear fit from mean for linear alpha-factor
Call LeastDeviation(Int(1), stddev!, nmax%, xdata!(), ydata!(), acoeff!())       ' average deviation
If ierror Then Exit Sub
End If
End If

' Do polynomial alpha-factor fit next
If CorrectionFlag% = 3 Then
kmax% = 2   ' second order
Call LeastSquares(kmax%, nmax%, xdata!(), ydata!(), acoeff!())                  ' polynomial regression
If ierror Then Exit Sub
alph1! = acoeff!(1)
alph2! = acoeff!(2)
alph3! = acoeff!(3)

' Calculate average % deviation of 2nd order fit
Call LeastDeviation(Int(1), stddev!, nmax%, xdata!(), ydata!(), acoeff!())       ' quadratic deviation
If ierror Then Exit Sub
End If

' Do non-linear alpha factor fit next
If CorrectionFlag% = 4 Then

' Try non-linear fit
Call LeastMathNonLinear(nmax%, xdata!(), ydata!(), acoeff!())
If ierror Then Exit Sub
alph1! = acoeff!(1)
alph2! = acoeff!(2)
alph3! = acoeff!(3)
alph4! = acoeff!(4)

Call LeastMathNonLinearDeviation(stddev!, nmax%, xdata!(), ydata!(), acoeff!())
If ierror Then Exit Sub
End If

' Display results
If DebugMode Then
If ibin% = 1 And PenepmaKratiosFlag(k%, l%) Then
Call IOWriteLog("Xray  Matrix   Alpha1  Alpha2  Alpha3  Alpha4 %AvgDev   *from Penepma 2012 Calculations")
ElseIf ibin% = 2 And PenepmaKratiosFlag(l%, k%) Then
Call IOWriteLog("Xray  Matrix   Alpha1  Alpha2  Alpha3  Alpha4 %AvgDev   *from Penepma 2012 Calculations")
Else
Call IOWriteLog("Xray  Matrix   Alpha1  Alpha2  Alpha3  Alpha4 %AvgDev")
End If
msg$ = syme$ & " " & symx$ & " in " & symm$ & "  "
msg$ = msg$ & Format$(Format$(alph1!, f84$), a80$) & Format$(Format$(alph2!, f84$), a80$) & Format$(Format$(alph3!, f84$), a80$) & Format$(Format$(alph4!, f84$), a80$) & MiscAutoFormat$(stddev!)
Call IOWriteLog(msg$)
End If

' Save calculated alpha-factors to look-up table
If rec1% = 0 Or rec2% = 0 Then GoTo AFactorCalculateFitFactorsNoRecord
alphal1!(rec1%, rec2%) = alph1!
alphal2!(rec1%, rec2%) = alph2!
alphal3!(rec1%, rec2%) = alph3!
alphal4!(rec1%, rec2%) = alph4!

' Save x-ray symbol for alpha-factor look-up table
AlphaXray$(rec1%, rec2%) = symx$

' Save to module level for plotting
If ibin% = 1 Then
AlphaNpts1% = npts%
For p% = 1 To npts%
AlphaXdata1!(p%) = xdata!(p%)
AlphaYdata1!(p%) = ydata!(p%)
Next p%

AlphaAcoeff1!(1) = acoeff!(1)
AlphaAcoeff1!(2) = acoeff!(2)
AlphaAcoeff1!(3) = acoeff!(3)
AlphaAcoeff1!(4) = acoeff!(4)

AlphaAstddev1! = stddev!

Else
AlphaNpts2% = npts%
For p% = 1 To npts%
AlphaXdata2!(p%) = xdata!(p%)
AlphaYdata2!(p%) = ydata!(p%)
Next p%

AlphaAcoeff2!(1) = acoeff!(1)
AlphaAcoeff2!(2) = acoeff!(2)
AlphaAcoeff2!(3) = acoeff!(3)
AlphaAcoeff2!(4) = acoeff!(4)

AlphaAstddev2! = stddev!
End If

600:  Next ibin%

Exit Sub

' Errors
AFactorCalculateFitFactorsError:
MsgBox Error$, vbOKOnly + vbCritical, "AFactorCalculateFitFactors"
ierror = True
Exit Sub

AFactorCalculateFitFactorsNoRecord:
msg$ = "The alpha array index numbers for the calculated binary are zero. This error should not occur. Please contact Probe Software with details."
MsgBox msg$, vbOKOnly + vbExclamation, "AFactorCalculateFitactors"
ierror = True
Exit Sub

End Sub

Sub AFactorCalculateKFactors(k As Integer, l As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' Calculate k-factors for alpha-factor calculations of the passed binary.
'
' This routine will calculate K-factors for a binary elemental system
' at MAXBIN compositions calculated using the currently selected ZAF
' selections. If both elements are loaded as absorbers only, then
' an error is returned.
'
' k = array position of 1st binary component in sample array
' l = array position of 2nd binary component in sample array

ierror = False
On Error GoTo AFactorCalculateKFactorsError

Dim notfoundA As Boolean, notfoundB As Boolean
Dim ibin As Integer, m As Integer, n As Integer
Dim astring As String

ReDim wout(1 To MAXBINARY% * 2) As Single, rout(1 To MAXBINARY% * 2) As Single
ReDim eout(1 To MAXBINARY% * 2) As String, xout(1 To MAXBINARY% * 2) As String
ReDim zout(1 To MAXBINARY% * 2) As Integer

' Check for emitter and absorber equal (already initialized to 1.000)
If sample(1).Elsyms$(k%) = sample(1).Elsyms(l%) Then Exit Sub

' Check for both elements as absorbers
If sample(1).Xrsyms$(k%) = vbNullString And sample(1).Xrsyms$(l%) = vbNullString Then Exit Sub

msg$ = vbCrLf & "Calculating alpha-factor binary " & Format$(sample(1).Elsyms$(k%), a20$) & " " & sample(1).Xrsyms$(k%) & " in " & Format$(sample(1).Elsyms$(l%), a20$)
If DebugMode Then Call IOWriteLog(msg$)
Call AnalyzeStatusAnal(msg$)

' Create a dummy sample for this binary k-factor calculation
AFactorTmpSample(1) = sample(1)

AFactorTmpSample(1).LastChan% = 2    ' binary only
AFactorTmpSample(1).OxideOrElemental% = 2   ' always elemental alpha-factors

' Load analyzed, then specified
If sample(1).Xrsyms$(k%) = vbNullString Then
AFactorTmpSample(1).LastElm% = 1    ' one absorber
AFactorTmpSample(1).Elsyms$(1) = sample(1).Elsyms$(l%)
AFactorTmpSample(1).Xrsyms$(1) = sample(1).Xrsyms$(l%)
AFactorTmpSample(1).Elsyms$(2) = sample(1).Elsyms$(k%)
AFactorTmpSample(1).Xrsyms$(2) = sample(1).Xrsyms$(k%)

ElseIf sample(1).Xrsyms$(l%) = vbNullString Then
AFactorTmpSample(1).LastElm% = 1    ' one absorber
AFactorTmpSample(1).Elsyms$(1) = sample(1).Elsyms$(k%)
AFactorTmpSample(1).Xrsyms$(1) = sample(1).Xrsyms$(k%)
AFactorTmpSample(1).Elsyms$(2) = sample(1).Elsyms$(l%)
AFactorTmpSample(1).Xrsyms$(2) = sample(1).Xrsyms$(l%)

Else
AFactorTmpSample(1).LastElm% = 2    ' no absorbers
AFactorTmpSample(1).Elsyms$(1) = sample(1).Elsyms$(k%)
AFactorTmpSample(1).Xrsyms$(1) = sample(1).Xrsyms$(k%)
AFactorTmpSample(1).Elsyms$(2) = sample(1).Elsyms$(l%)
AFactorTmpSample(1).Xrsyms$(2) = sample(1).Xrsyms$(l%)
End If

' Load element arrays
Call ElementGetData(AFactorTmpSample())
If ierror Then Exit Sub

' Initialize
For ibin% = 1 To 2
For n% = 1 To MAXBINARY%
m% = (n% - 1) * 2 + ibin%
wout!(m%) = 0#
rout!(m%) = 1#
eout$(m%) = AFactorTmpSample(1).Elsyms$(ibin%)
xout$(m%) = AFactorTmpSample(1).Xrsyms$(ibin%)
zout%(m%) = AFactorTmpSample(1).AtomicNums%(ibin%)
Next n%
Next ibin%

' If use Penepma k-ratios flag then load values for this binary (if found) (1 = do not use, 2 = use)
If UsePenepmaKratiosFlag = 2 Then
notfoundA = True
notfoundB = True
Call AFactorPenepmaReadMatrix(wout!(), rout!(), eout$(), xout$(), zout%(), notfoundA, notfoundB, AFactorTmpSample())
If ierror Then Exit Sub
If Not notfoundA And sample(1).Xrsyms$(k%) <> vbNullString Then PenepmaKratiosFlag(k%, l%) = True
If Not notfoundB And sample(1).Xrsyms$(l%) <> vbNullString Then PenepmaKratiosFlag(l%, k%) = True
End If

' If Penepma k-ratios were not found, initialize calculations for ZAF or phi-rho-z calculations (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If UsePenepmaKratiosFlag = 1 Or (UsePenepmaKratiosFlag = 2 And (notfoundA And sample(1).Xrsyms$(k%) <> vbNullString) Or (notfoundB And sample(1).Xrsyms$(l%) <> vbNullString)) Then

If CorrectionFlag% <> MAXCORRECTION% Then
Call ZAFSetZAF(AFactorTmpSample())
If ierror Then Exit Sub
Else
'Call ZAFSetZAF3(AFactorTmpSample())
'If ierror Then Exit Sub
End If

' Calculate array of intensities using ZAF or Phi-Rho-Z (zout() is the atomic number for alpha record)
Call ZAFAFactor(wout!(), rout!(), eout$(), xout$(), zout%(), analysis, AFactorTmpSample())
If ierror Then Exit Sub
End If

' Calculate alpha factors, fit and load into alpha-factor look up tables
Call AFactorCalculateFitFactors(k%, l%, wout!(), rout!(), eout$(), xout$(), zout%())
If ierror Then Exit Sub

' Output to log
If VerboseMode Then
For ibin% = 1 To 2
If AFactorTmpSample(1).Xrsyms$(ibin%) <> vbNullString Then

If ibin% = 1 Then
astring$ = "AFactorCalculateKFactors: " & AFactorTmpSample(1).Elsyms$(1) & " " & AFactorTmpSample(1).Xrsyms$(1) & " in " & AFactorTmpSample(1).Elsyms$(2)
Call IOWriteLog$(vbCrLf & astring$ & " at " & Format$(AFactorTmpSample(1).TakeoffArray!(1)) & " degrees and " & Format$(AFactorTmpSample(1).KilovoltsArray!(1)) & " keV")
Else
astring$ = "AFactorCalculateKFactors: " & AFactorTmpSample(1).Elsyms$(2) & " " & AFactorTmpSample(1).Xrsyms$(2) & " in " & AFactorTmpSample(1).Elsyms$(1)
Call IOWriteLog$(vbCrLf & astring$ & " at " & Format$(AFactorTmpSample(1).TakeoffArray!(2)) & " degrees and " & Format$(AFactorTmpSample(1).KilovoltsArray!(2)) & " keV")
End If
astring$ = Format$(vbTab & "Conc%", a08$) & vbTab & Format$("Kratios", a08$) & vbTab & Format$("Alpha", a08$)
Call IOWriteLog$(astring$)

For n% = 1 To MAXBINARY%
m% = (n% - 1) * 2 + ibin%

If ibin% = 1 Then
astring$ = vbTab & MiscAutoFormat$(wout!(m%)) & vbTab & MiscAutoFormat$(rout!(m%)) & MiscAutoFormat$(AlphaYdata1!(n%))
Else
astring$ = vbTab & MiscAutoFormat$(wout!(m%)) & vbTab & MiscAutoFormat$(rout!(m%)) & MiscAutoFormat$(AlphaYdata2!(n%))
End If
Call IOWriteLog$(astring$)

Next n%

End If
Next ibin%
End If

Exit Sub

' Errors
AFactorCalculateKFactorsError:
MsgBox Error$, vbOKOnly + vbCritical, "AFactorCalculateKFactors"
ierror = True
Exit Sub

End Sub

Sub AFactorInitFactors()
' Initialize the alpha arrays (to force a recalculation)

ierror = False
On Error GoTo AFactorInitFactorsError

Dim i As Integer, j As Integer

Call IOWriteLog(vbCrLf & "Initializing alpha-factors...")

alphakev! = 0#
alphatof! = 0#

For i% = 1 To MAXELM%
For j% = 1 To MAXELM%
AlphaXray$(i%, j%) = vbNullString
alphal1!(i%, j%) = 1#
alphal2!(i%, j%) = 0#
alphal3!(i%, j%) = 0#
alphal4!(i%, j%) = 0#
Next j%
Next i%

For i% = 1 To MAXCHAN%
For j% = 1 To MAXCHAN%
alpha1!(i%, j%) = 0#
alpha2!(i%, j%) = 0#
alpha3!(i%, j%) = 0#
alpha4!(i%, j%) = 0#
PenepmaKratiosFlag(i%, j%) = False
Next j%
Next i%

' Indicate that alpha-factors are initialized
AllAFactorUpdateNeeded = False

Exit Sub

' Errors
AFactorInitFactorsError:
MsgBox Error$, vbOKOnly + vbCritical, "AFactorInitFactors"
ierror = True
Exit Sub

End Sub

Sub AFactorLoadEmpirical(sample() As TypeSample)
' Load empirical alpha-factors from the empirical alpha factor arrays
' Note that only "elemental" alpha-factors are supported in Probe, "oxide"
' alpha factors will have to be re-calculated based on the original measured
' k-factors by first normalizing them to oxide end-member k-factors.

ierror = False
On Error GoTo AFactorLoadEmpiricalError

Dim ip As Integer, ipp As Integer
Dim n As Integer

Call IOWriteLog("Loading Empirical Alpha-Factors...")

' Loop on each empirical alpha and check if it is needed
For n% = 1 To MAXEMPFAC%

' Check conditions
If empfackev!(n%) = sample(1).kilovolts! Then
If empfactof!(n%) = sample(1).takeoff! Then

' Find each emitter/absorber position in the sample
ip% = IPOS1(sample(1).LastElm%, empface$(n%), sample(1).Elsyms$())
ipp% = IPOS1(sample(1).LastChan%, empfaca$(n%), sample(1).Elsyms$())

' Check that emitter and absorber match
If ip% <> 0 And ipp% <> 0 Then

' Check that emitter x-ray match
If MiscStringsAreSame(empfacx$(n%), sample(1).Xrsyms$(ip%)) Then

' Load pre-calculated alpha-factor in alpha arrays (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% = 1 And empfaco%(n%) = 1 Then
alpha1!(ip%, ipp%) = empfac1!(n%)
alpha2!(ip%, ipp%) = 0#
alpha3!(ip%, ipp%) = 0#
alpha4!(ip%, ipp%) = 0#
If DebugMode Then
Call IOWriteLog("Loading constant empirical alpha for " & sample(1).Elsyms$(ip%) & " " & sample(1).Xrsyms$(ip%) & " in " & sample(1).Elsyms$(ipp%) & Str$(empfac1!(n%)) & Str$(empfac2!(n%)) & Str$(empfac3!(n%)))
End If
End If

' Linear
If CorrectionFlag% = 2 And empfaco%(n%) = 2 Then
alpha1!(ip%, ipp%) = empfac1!(n%)
alpha2!(ip%, ipp%) = empfac2!(n%)
alpha3!(ip%, ipp%) = 0#
alpha4!(ip%, ipp%) = 0#
If DebugMode Then
Call IOWriteLog("Loading linear empirical alpha for " & sample(1).Elsyms$(ip%) & " " & sample(1).Xrsyms$(ip%) & " in " & sample(1).Elsyms$(ipp%) & Str$(empfac1!(n%)) & Str$(empfac2!(n%)) & Str$(empfac3!(n%)))
End If
End If

' Polynomial
If CorrectionFlag% = 3 And empfaco%(n%) = 3 Then
alpha1!(ip%, ipp%) = empfac1!(n%)
alpha2!(ip%, ipp%) = empfac2!(n%)
alpha3!(ip%, ipp%) = empfac3!(n%)
alpha4!(ip%, ipp%) = 0#
If DebugMode Then
Call IOWriteLog("Loading polynomial empirical alpha for " & sample(1).Elsyms$(ip%) & " " & sample(1).Xrsyms$(ip%) & " in " & sample(1).Elsyms$(ipp%) & Str$(empfac1!(n%)) & Str$(empfac2!(n%)) & Str$(empfac3!(n%)))
End If
End If

' Non-linear
If CorrectionFlag% = 4 And empfaco%(n%) = 4 Then
alpha1!(ip%, ipp%) = empfac1!(n%)
alpha2!(ip%, ipp%) = empfac2!(n%)
alpha3!(ip%, ipp%) = empfac3!(n%)
alpha4!(ip%, ipp%) = empfac4!(n%)
If DebugMode Then
Call IOWriteLog("Loading non-linear empirical alpha for " & sample(1).Elsyms$(ip%) & " " & sample(1).Xrsyms$(ip%) & " in " & sample(1).Elsyms$(ipp%) & Str$(empfac1!(n%)) & Str$(empfac2!(n%)) & Str$(empfac3!(n%)) & Str$(empfac4!(n%)))
End If
End If

End If
End If

End If
End If

Next n%

Exit Sub

' Errors
AFactorLoadEmpiricalError:
MsgBox Error$, vbOKOnly + vbCritical, "AFactorLoadEmpirical"
ierror = True
Exit Sub

End Sub

Sub AFactorLoadFactors(analysis As TypeAnalysis, sample() As TypeSample)
' Load the specific alpha-factors based on sample setup

ierror = False
On Error GoTo AFactorLoadFactorsError

Dim emitter As Integer, absorber As Integer
Dim ip As Integer, ipp As Integer, i As Integer

' Check that sample is not a combined conditions sample (only check takeoff and kilovolts)
If sample(1).LastElm% > 1 Then
If MiscIsDifferent3(sample(1).LastElm%, sample(1).TakeoffArray!()) Then GoTo AFactorLoadFactorsCombinedConditions
If MiscIsDifferent3(sample(1).LastElm%, sample(1).KilovoltsArray!()) Then GoTo AFactorLoadFactorsCombinedConditions
End If

' Check for update
If AllAFactorUpdateNeeded Then
Call AFactorInitFactors
If ierror Then Exit Sub

' Re-load the empirical alpha-factors
If EmpiricalAlphaFlag = 2 Then
Call AFactorReadEmpirical
If ierror Then Exit Sub
End If
End If

' Check for a change in operating conditions and re-calculate all if necessary
If alphakev! <> sample(1).kilovolts! Or alphatof! <> sample(1).takeoff! Then
Call AFactorCalculateFactor(Int(0), Int(0), analysis, sample())
If ierror Then Exit Sub

' Loop on each emitting element in sample
Else
For emitter% = 1 To sample(1).LastElm%
ip% = IPOS1(MAXELM%, sample(1).Elsyms$(emitter%), Symlo$())

' Loop on each absorbing element in sample
For absorber% = 1 To sample(1).LastChan%
ipp% = IPOS1(MAXELM%, sample(1).Elsyms$(absorber%), Symlo$())

' Check for missing emitter, xray or absorber combination
If ip% > 0 And ipp% > 0 Then
If AlphaXray$(ip%, ipp%) <> sample(1).Xrsyms$(emitter%) Then
Call AFactorCalculateFactor(emitter%, absorber%, analysis, sample())
If ierror Then Exit Sub
End If
End If

Next absorber%
Next emitter%
End If

' Loop on each emitting element in sample
For emitter% = 1 To sample(1).LastElm%
ip% = IPOS1(MAXELM%, sample(1).Elsyms$(emitter%), Symlo$())

' Loop on each absorbing element in sample
For absorber% = 1 To sample(1).LastChan%
ipp% = IPOS1(MAXELM%, sample(1).Elsyms$(absorber%), Symlo$())

' Load pre-calculated alpha-factors into sample alpha arrays
If ip% > 0 And ipp% > 0 Then
alpha1!(emitter%, absorber%) = alphal1!(ip%, ipp%)
alpha2!(emitter%, absorber%) = alphal2!(ip%, ipp%)
alpha3!(emitter%, absorber%) = alphal3!(ip%, ipp%)
alpha4!(emitter%, absorber%) = alphal4!(ip%, ipp%)
End If

Next absorber%
Next emitter%

' Next check for empirical alpha factors
If EmpiricalAlphaFlag = 2 Then
Call AFactorLoadEmpirical(sample())
If ierror Then Exit Sub
End If

Exit Sub

' Errors
AFactorLoadFactorsError:
MsgBox Error$, vbOKOnly + vbCritical, "AFactorLoadFactors"
ierror = True
Exit Sub

AFactorLoadFactorsCombinedConditions:
msg$ = "Alpha factor calculations are not supported for combined condition samples at this time"
msg$ = msg$ & vbCrLf & "Lastelm: " & Format$(sample(1).LastElm%)
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & vbCrLf & "Element: " & sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & ", TO: " & Format$(sample(1).TakeoffArray!(i%)) & ", KeV: " & Format$(sample(1).KilovoltsArray!(i%))
Next i%
MsgBox msg$, vbOKOnly + vbExclamation, "AFactorLoadFactors"
ierror = True
Exit Sub

End Sub

Sub AFactorLoadFactorsReturn(sample() As TypeSample, talpha1() As Single, talpha2() As Single, talpha3() As Single, talpha4() As Single)
' Return the module level sample specific alpha factors

ierror = False
On Error GoTo AFactorLoadFactorsReturnError

Dim emitter As Integer, absorber As Integer
Dim ip As Integer, ipp As Integer

' Loop on each emitting element in sample
For emitter% = 1 To sample(1).LastElm%
ip% = IPOS1(MAXELM%, sample(1).Elsyms$(emitter%), Symlo$())

' Loop on each absorbing element in sample
For absorber% = 1 To sample(1).LastChan%
ipp% = IPOS1(MAXELM%, sample(1).Elsyms$(absorber%), Symlo$())

' Load pre-calculated alpha-factors into sample alpha arrays
If ip% > 0 And ipp% > 0 Then
talpha1!(emitter%, absorber%) = alphal1!(ip%, ipp%)
talpha2!(emitter%, absorber%) = alphal2!(ip%, ipp%)
talpha3!(emitter%, absorber%) = alphal3!(ip%, ipp%)
talpha4!(emitter%, absorber%) = alphal4!(ip%, ipp%)
End If

Next absorber%
Next emitter%

Exit Sub

' Errors
AFactorLoadFactorsReturnError:
MsgBox Error$, vbOKOnly + vbCritical, "AFactorLoadFactorsReturn"
ierror = True
Exit Sub

End Sub

Sub AFactorReadEmpirical()
' Read empirical alpha-factors from ASCII file

ierror = False
On Error GoTo AFactorReadEmpiricalError

Dim n As Integer

' Load filename
If Dir$(EmpFACFile$) = vbNullString Then GoTo AFactorReadEmpiricalNotFound
Call IOWriteLog("Reading file " & EmpFACFile$ & "...")

' Open file
Open EmpFACFile$ For Input As #EMPFacFileNumber%

n% = 1
Do While Not EOF(EMPFacFileNumber%)

Input #EMPFacFileNumber%, empfactof!(n%), empfackev!(n%), empfaco%(n%), empface$(n%), empfacx$(n%), empfaca$(n%), empfac1!(n%), empfac2!(n%), empfac3!(n%), empfacstr$(n%)

n% = n% + 1
If n% > MAXEMPFAC% Then Exit Do

Loop

Close #EMPFacFileNumber%
Exit Sub

' Errors
AFactorReadEmpiricalError:
MsgBox Error$, vbOKOnly + vbCritical, "AFactorReadEmpirical"
Close #EMPFacFileNumber%
ierror = True
Exit Sub

AFactorReadEmpiricalNotFound:
msg$ = "Empirical alpha-factor file " & EmpFACFile$ & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "AFactorReadEmpirical"
ierror = True
Exit Sub

End Sub

Sub AFactorAFASaveFactors(analysis As TypeAnalysis, sample() As TypeSample)
' Save the caluclated conditions and alpha-factors to an ASCII file.
' Note that in addition to the alpha-factors (constant, linear, polynomial and
' non-linear) themselves, this routine stores other information required
' for quantitative calculations (e.g., standard assignments and intensities)
' for quantitative image calculations (CalcImage).

ierror = False
On Error GoTo AFactorAFASaveFactorsError

Dim i As Integer, j As Integer

' Open the alpha-factor ASCII file
Open AFactorDataFile$ For Output As #AFactorDataFileNumber%

' Write run info
Print #AFactorDataFileNumber%, VbDquote$ & ProbeDataFile$ & VbDquote$
Print #AFactorDataFileNumber%, VbDquote$ & MDBUserName$ & VbDquote$
Print #AFactorDataFileNumber%, VbDquote$ & MDBFileTitle$ & VbDquote$
Print #AFactorDataFileNumber%, VbDquote$ & Now & VbDquote$
Print #AFactorDataFileNumber%, sample(1).kilovolts!, sample(1).takeoff!
Print #AFactorDataFileNumber%, sample(1).LastElm%, sample(1).LastChan%, sample(1).OxideOrElemental%, NumberofStandards%

' Write column labels
msg$ = Space$(8) & vbTab
msg$ = msg$ & Format$(VbDquote$ & "Spec" & VbDquote$, a80$) & vbTab
msg$ = msg$ & Format$(VbDquote$ & "Xtal" & VbDquote$, a80$) & vbTab
msg$ = msg$ & Format$(VbDquote$ & "Std#" & VbDquote$, a80$) & vbTab
msg$ = msg$ & Format$(VbDquote$ & "Std%" & VbDquote$, a80$) & vbTab
msg$ = msg$ & Format$(VbDquote$ & "Beta" & VbDquote$, a80$) & vbTab
msg$ = msg$ & Format$(VbDquote$ & "Count" & VbDquote$, a80$) & vbTab
Print #AFactorDataFileNumber%, msg$

' Write standard counts, weight percents and beta factors
For j% = 1 To sample(1).LastElm%
msg$ = Format$(VbDquote$ & sample(1).Elsyms$(j%) & VbDquote$ & " " & VbDquote$ & sample(1).Xrsyms$(j%) & VbDquote$, a80$) & vbTab
msg$ = msg$ & Format$(sample(1).MotorNumbers%(j%), a80$) & vbTab
msg$ = msg$ & Format$(VbDquote$ & sample(1).CrystalNames$(j%), a80$) & VbDquote$ & vbTab
msg$ = msg$ & Format$(sample(1).StdAssigns%(j%), a80$) & vbTab
msg$ = msg$ & Format$(Format$(analysis.StdAssignsPercents!(j%), f83$), a80$) & vbTab
msg$ = msg$ & Format$(Format$(analysis.StdAssignsBetas!(j%), f84$), a80$) & vbTab
msg$ = msg$ & Format$(Format$(analysis.StdAssignsCounts!(j%), f81$), a80$) & vbTab
Print #AFactorDataFileNumber%, msg$
Next j%

' Write column labels
msg$ = Space$(8) & vbTab
msg$ = msg$ & Format$(VbDquote$ & "Oxide" & VbDquote$, a80$) & vbTab
msg$ = msg$ & Format$(VbDquote$ & "Cation" & VbDquote$, a80$) & vbTab
Print #AFactorDataFileNumber%, msg$

' Write oxide and cations
For j% = 1 To sample(1).LastChan%
msg$ = Format$(VbDquote$ & sample(1).Elsyms$(j%) & VbDquote$ & " " & VbDquote$ & sample(1).Xrsyms$(j%) & VbDquote$, a80$) & vbTab
msg$ = msg$ & Format$(sample(1).numoxd%(j%), a80$) & vbTab
msg$ = msg$ & Format$(sample(1).numcat%(j%), a80$) & vbTab
Print #AFactorDataFileNumber%, msg$
Next j%

' Write interference calibration
msg$ = Space$(8) & vbTab
For i% = 1 To MAXINTF%
msg$ = msg$ & Format$(VbDquote$ & "Interf" & Trim$(Str$(i%)) & VbDquote$, a80$) & vbTab
Next i%
Print #AFactorDataFileNumber%, msg$

' Interference elements
For j% = 1 To sample(1).LastElm%
msg$ = Format$(VbDquote$ & sample(1).Elsyms$(j%) & VbDquote$ & " " & VbDquote$ & sample(1).Xrsyms$(j%) & VbDquote$, a80$) & vbTab
For i% = 1 To MAXINTF%
msg$ = msg$ & VbDquote$ & Format$(sample(1).StdAssignsIntfElements$(i%, j%), a80$) & VbDquote$ & vbTab
Next i%
Print #AFactorDataFileNumber%, msg$
Next j%

' Interference std assignments
For j% = 1 To sample(1).LastElm%
msg$ = Format$(VbDquote$ & sample(1).Elsyms$(j%) & VbDquote$ & " " & VbDquote$ & sample(1).Xrsyms$(j%) & VbDquote$, a80$) & vbTab
For i% = 1 To MAXINTF%
msg$ = msg$ & Format$(sample(1).StdAssignsIntfStds%(i%, j%), a80$) & vbTab
Next i%
Print #AFactorDataFileNumber%, msg$
Next j%

' Interference counts
For j% = 1 To sample(1).LastElm%
msg$ = Format$(VbDquote$ & sample(1).Elsyms$(j%) & VbDquote$ & " " & VbDquote$ & sample(1).Xrsyms$(j%) & VbDquote$, a80$) & vbTab
For i% = 1 To MAXINTF%
msg$ = msg$ & MiscAutoFormat$(analysis.StdAssignsIntfCounts!(i%, j%)) & vbTab
Next i%
Print #AFactorDataFileNumber%, msg$
Next j%

' Standard percents
msg$ = Space$(8) & vbTab
For i% = 1 To NumberofStandards%
msg$ = msg$ & Format$(Format$(StandardNumbers%(i%)), a80$) & vbTab
Next i%
Print #AFactorDataFileNumber%, msg$

For j% = 1 To sample(1).LastElm%
msg$ = Format$(VbDquote$ & sample(1).Elsyms$(j%) & VbDquote$ & " " & VbDquote$ & sample(1).Xrsyms$(j%) & VbDquote$, a80$) & vbTab
For i% = 1 To NumberofStandards%
msg$ = msg$ & MiscAutoFormat$(analysis.StdPercents!(i%, j%)) & vbTab
Next i%
Print #AFactorDataFileNumber%, msg$
Next j%

' Write MAN background coeff column labels
msg$ = Space$(8) & vbTab
msg$ = msg$ & Format$(VbDquote$ & "BgdTyp" & VbDquote$, a80$) & vbTab
msg$ = msg$ & Format$(VbDquote$ & "MANAbs" & VbDquote$, a80$) & vbTab
For i% = 1 To MAXCOEFF%
msg$ = msg$ & Format$(VbDquote$ & "Coeff" & Trim$(Str$(i%)) & VbDquote$, a80$) & vbTab
Next i%
Print #AFactorDataFileNumber%, msg$

' MAN background types and fit coefficients
For j% = 1 To sample(1).LastElm%
msg$ = Format$(VbDquote$ & sample(1).Elsyms$(j%) & VbDquote$ & " " & VbDquote$ & sample(1).Xrsyms$(j%) & VbDquote$, a80$) & vbTab
msg$ = msg$ & Format$(sample(1).BackgroundTypes%(j%), a80$) & vbTab
msg$ = msg$ & Format$(sample(1).MANAbsCorFlags%(j%), a80$) & vbTab
For i% = 1 To MAXCOEFF%
msg$ = msg$ & MiscAutoFormat$(analysis.MANFitCoefficients!(i%, j%)) & vbTab
Next i%
Print #AFactorDataFileNumber%, msg$
Next j%

' Write absorbers column labels
msg$ = Space$(8) & vbTab
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(VbDquote$ & sample(1).Elsyms$(i%) & VbDquote$, a80$) & vbTab
Next i%
Print #AFactorDataFileNumber%, msg$

' Loop on emitters
For j% = 1 To sample(1).LastElm%

' Loop on absorbers
msg$ = Format$(VbDquote$ & sample(1).Elsyms$(j%) & VbDquote$ & " " & VbDquote$ & sample(1).Xrsyms$(j%) & VbDquote$, a80$) & vbTab
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(alpha1(j%, i%), f84$), a80$) & vbTab
Next i%
Print #AFactorDataFileNumber%, msg$

msg$ = Format$(VbDquote$ & sample(1).Elsyms$(j%) & VbDquote$ & " " & VbDquote$ & sample(1).Xrsyms$(j%) & VbDquote$, a80$) & vbTab
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(alpha2(j%, i%), f84$), a80$) & vbTab
Next i%
Print #AFactorDataFileNumber%, msg$

msg$ = Format$(VbDquote$ & sample(1).Elsyms$(j%) & VbDquote$ & " " & VbDquote$ & sample(1).Xrsyms$(j%) & VbDquote$, a80$) & vbTab
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(alpha3(j%, i%), f84$), a80$) & vbTab
Next i%
Print #AFactorDataFileNumber%, msg$

msg$ = Format$(VbDquote$ & sample(1).Elsyms$(j%) & VbDquote$ & " " & VbDquote$ & sample(1).Xrsyms$(j%) & VbDquote$, a80$) & vbTab
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(alpha4(j%, i%), f84$), a80$) & vbTab
Next i%
Print #AFactorDataFileNumber%, msg$

Next j%

' Close file
Close #AFactorDataFileNumber%

Exit Sub

' Errors
AFactorAFASaveFactorsError:
MsgBox Error$, vbOKOnly + vbCritical, "AFactorAFASaveFactors"
Close #AFactorDataFileNumber%
ierror = True
Exit Sub

End Sub

Sub AFactorSmp(row As Integer, excess As Single, UnkCounts() As Single, zerror As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' Calculate alpha-factor correction using a Bence-Albee iteration. Note excess oxygen is
' passed because AFactorSmp cannot keep excess and calculated oxygen separated.

ierror = False
On Error GoTo AFactorSmpError

Dim MaxALPIter As Integer
Dim ALPMinToler As Single, ALPMinTotal As Single
Dim maxdiff As Single, oxygen As Single
Dim chan As Integer, iter As Integer
Dim ip As Integer, ipp As Integer, ippp As Integer
Dim temp1 As Single, temp2 As Single
Dim diff As Single, temp As Single, sum As Single

ReDim kratios(1 To MAXCHAN%) As Single
ReDim kratios2(1 To MAXCHAN%) As Single
ReDim unkwts(1 To MAXCHAN%) As Single
ReDim fixedelement(1 To MAXCHAN%) As Integer    ' specified (not analyzed) elements

MaxALPIter% = 30
ALPMinToler! = 0.005 ' in weight percent
ALPMinTotal! = 0.001 ' in weight percent

If DebugMode Then
Call IOWriteLog(vbCrLf & "Entering AFactorSmp...")

msg$ = "ELEMENT "
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%), a80$)
Next chan%
Call IOWriteLog(msg$)

msg$ = "UNK WT% "
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(analysis.WtPercents!(chan%), f83$), a80$)
Next chan%
Call IOWriteLog(msg$)

msg$ = "UNK CNT "
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & MiscAutoFormat$(UnkCounts!(chan%))
Next chan%
Call IOWriteLog(msg$)

msg$ = "STD CNT "
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & MiscAutoFormat$(analysis.StdAssignsCounts!(chan%))
Next chan%
Call IOWriteLog(msg$)

msg$ = "STD WT% "
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.StdAssignsPercents!(chan%), f83$), a80$)
Next chan%
Call IOWriteLog(msg$)

msg$ = "STD BETA"
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.StdAssignsBetas!(chan%), f84$), a80$)
Next chan%
Call IOWriteLog(msg$)

msg$ = "EXCESS O" & Format$(Format$((excess!), f82$), a80$)
Call IOWriteLog(msg$)
Call IOWriteLog(vbNullString)
End If

' Load specified element flags
If sample(1).StoichiometryElementFlag% Then ip% = IPOS1B(sample(1).LastElm% + 1, sample(1).LastChan%, sample(1).StoichiometryElement$, sample(1).Elsyms$())
If sample(1).RelativeElementFlag% Then ipp% = IPOS1B(sample(1).LastElm% + 1, sample(1).LastChan%, sample(1).RelativeElement$, sample(1).Elsyms$())
If sample(1).DifferenceElementFlag% Then ippp% = IPOS1B(sample(1).LastElm% + 1, sample(1).LastChan%, sample(1).DifferenceElement$, sample(1).Elsyms$())
For chan% = sample(1).LastElm% + 1 To sample(1).LastChan%
If chan% <> ip% And chan% <> ipp% And chan% <> ippp% Then
fixedelement%(chan%) = True
End If
Next chan%

' Calculate elemental k-ratios from measured kratios
For chan% = 1 To sample(1).LastElm%
If sample(1).DisableAcqFlag(chan%) = 0 Then
If sample(1).DisableQuantFlag%(chan%) = 0 Then

If Not UseAggregateIntensitiesFlag Then
If analysis.StdAssignsCounts!(chan%) <= 0# Then GoTo AFactorSmpBadStdCounts
kratios!(chan%) = UnkCounts!(chan%) / analysis.StdAssignsCounts!(chan%)       ' calculate raw kratio
kratios!(chan%) = kratios!(chan%) * analysis.StdAssignsPercents!(chan%) / analysis.StdAssignsBetas!(chan%)  ' correct for standard
kratios2!(chan%) = kratios!(chan%) / 100#                                     ' calculate elemental k-ratio (alpha calculations use wt%)

Else
ip% = IPOS8(chan%, sample(1).Elsyms$(chan%), sample(1).Xrsyms$(chan%), sample())
If ip% = 0 Then
If analysis.StdAssignsCounts!(chan%) <= 0# Then GoTo AFactorSmpBadStdCounts
kratios!(chan%) = UnkCounts!(chan%) / analysis.StdAssignsCounts!(chan%)       ' calculate raw kratio
kratios!(chan%) = kratios!(chan%) * analysis.StdAssignsPercents!(chan%) / analysis.StdAssignsBetas!(chan%)  ' correct for standard
kratios2!(chan%) = kratios!(chan%) / 100#                                     ' calculate elemental k-ratio (alpha calculations use wt%)
End If
End If

End If
End If
Next chan%

' Perform secondary boundary fluorescence correction on measured k-ratios
If UseSecondaryBoundaryFluorescenceCorrectionFlag Then
Call SecondaryCorrection(row%, kratios2!(), sample())
If ierror Then Exit Sub
End If

' Bence-Albee iteration begins here
analysis.ZAFIter! = 1#
For iter% = 1 To MaxALPIter%

' Calculate weight percents for analyzed elements
For chan% = 1 To sample(1).LastElm%
If sample(1).DisableAcqFlag(chan%) = 0 Then
If sample(1).DisableQuantFlag%(chan%) = 0 Then

ip% = IPOS8(chan%, sample(1).Elsyms$(chan%), sample(1).Xrsyms$(chan%), sample()) ' find if element is duplicated
If Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0) Then

If sample(1).DisableQuantFlag(chan%) = 0 Then             ' no disabled quant flag
If analysis.UnkBetas!(chan%) <= 0# Then GoTo AFactorSmpBadUnkBetas
If analysis.StdAssignsPercents!(chan%) <= 0# Then GoTo AFactorSmpBadStdPercents
If analysis.StdAssignsBetas!(chan%) <= 0# Then GoTo AFactorSmpBadStdBetas

' Calculate the unknown weight percents (alpha calculations use wt%)
unkwts!(chan%) = 100# * kratios2!(chan%) * analysis.UnkBetas!(chan%)            ' this is the matrix correction calculation!!!

' Check for force to zero flag
If ForceNegativeKratiosToZeroFlag = True Then
If unkwts!(chan%) <= 0# Then unkwts!(chan%) = NOT_ANALYZED_VALUE_SINGLE! ' use a non-zero value
End If

End If
End If

End If
End If
Next chan%

If VerboseMode Then
msg$ = "ANAL WT%"
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(unkwts!(chan%), f83$), a80$)
Next chan%
Call IOWriteLog(msg$)
End If

' Load specified concentrations (default)
For chan% = sample(1).LastElm% + 1 To sample(1).LastChan%
unkwts!(chan%) = 0#
If fixedelement%(chan%) Then
unkwts!(chan%) = analysis.WtPercents!(chan%)
End If
Next chan%

' Load stoichiometric oxygen
oxygen! = 0#
If sample(1).OxideOrElemental% = 1 Then

' Add specified (excess) oxygen
If sample(1).OxygenChannel% > 0 Then unkwts!(sample(1).OxygenChannel%) = excess!

' Sum calculated oxygen (do not include oxygen from element by difference)
temp1! = 0#
For chan% = 1 To sample(1).LastChan%
If sample(1).DisableQuantFlag%(chan%) = 0 Then
If chan% <> ippp% Then
temp2! = ConvertElmToOxd!(analysis.WtPercents!(chan%), sample(1).Elsyms$(chan%), sample(1).numcat%(chan%), sample(1).numoxd%(chan%))
temp1! = temp1! + (temp2! - analysis.WtPercents!(chan%))  ' calculate oxygen from oxide minus elemental
End If
End If
Next chan%

' Check for analyzed oxygen disabled but calculating oxygen by stoichiometry
If sample(1).OxygenChannel% > 0 And sample(1).OxideOrElemental% = 1 Then
If sample(1).DisableQuantFlag%(sample(1).OxygenChannel%) = 1 Then
temp1! = ConvertOxygenFromCations3!(analysis, sample())
End If
End If

' Calculate equivalent oxygen from halogens and subtract from calculated oxygen if flagged
If UseOxygenFromHalogensCorrectionFlag Then temp1! = temp1! - ConvertHalogensToOxygen(sample(1).LastChan%, sample(1).Elsyms$(), sample(1).DisableQuantFlag%(), unkwts!())

' Save oxygen calculated by stoichiometry for calculations below
oxygen! = temp1!
End If

' Load element by stoichiometry to stoichiometric element (oxygen)
If sample(1).StoichiometryElementFlag% And sample(1).StoichiometryElement$ <> vbNullString Then
ip% = IPOS1B(sample(1).LastElm% + 1, sample(1).LastChan%, sample(1).StoichiometryElement$, sample(1).Elsyms$())
If ip% > 0 And sample(1).OxideOrElemental% = 1 Then
unkwts!(ip%) = oxygen! / AllAtomicWts!(ATOMIC_NUM_OXYGEN%) * sample(1).StoichiometryRatio! * sample(1).AtomicWts!(ip%)

If VerboseMode% Then
msg$ = "STOI WT%" & Format$(Format$(unkwts!(ip%), f83$), a80$)
Call IOWriteLog(msg$)
End If
End If
End If

' Load element relative to another element
If sample(1).RelativeElementFlag% And sample(1).RelativeElement$ <> vbNullString And sample(1).RelativeToElement$ <> vbNullString Then
ip% = IPOS1B(sample(1).LastElm% + 1, sample(1).LastChan%, sample(1).RelativeElement$, sample(1).Elsyms$())
ipp% = IPOS1(sample(1).LastChan%, sample(1).RelativeToElement$, sample(1).Elsyms$())
If ip% > 0 And ipp% > 0 Then
unkwts!(ip%) = (unkwts!(ipp%) / sample(1).AtomicWts!(ipp%)) * sample(1).AtomicWts!(ip%) * sample(1).RelativeRatio!

' Add stoichiometric oxygen from element by stoichiometry to another element
If sample(1).OxideOrElemental% = 1 Then
temp! = ConvertElmToOxd!(unkwts!(ip%), sample(1).Elsyms$(ip%), sample(1).numcat%(ip%), sample(1).numoxd%(ip%))
oxygen! = oxygen! + (temp! - unkwts!(ip%))
End If

If VerboseMode% Then
msg$ = "RELA WT%" & Format$(Format$(unkwts!(ip%), f83$), a80$)
Call IOWriteLog(msg$)
End If
End If
End If

' Total all weight percents for difference element calculation
analysis.TotalPercent! = 0#
For chan% = 1 To sample(1).LastChan%
If sample(1).DisableQuantFlag%(chan%) = 0 Then
analysis.TotalPercent! = analysis.TotalPercent! + unkwts!(chan%)
End If
Next chan%
analysis.TotalPercent! = analysis.TotalPercent! + oxygen!

If VerboseMode% Then
msg$ = "SUM WT% " & Format$(Format$(analysis.TotalPercent!, f83$), a80$)
Call IOWriteLog(msg$)
End If

' Calculate element by difference based on current weight percents
If sample(1).DifferenceElementFlag% And sample(1).DifferenceElement$ <> vbNullString Then
ip% = IPOS1B(sample(1).LastElm% + 1, sample(1).LastChan%, sample(1).DifferenceElement$, sample(1).Elsyms$())
If ip% > 0 And analysis.TotalPercent! < 100# Then
unkwts!(ip%) = 100# - analysis.TotalPercent!
analysis.TotalPercent! = 100#

' Add stoichiometric oxygen from element by difference (convert oxide difference to elemental first)
If sample(1).OxideOrElemental% = 1 Then
unkwts!(ip%) = ConvertOxdToElm!(unkwts!(ip%), sample(1).Elsyms$(ip%), sample(1).numcat%(ip%), sample(1).numoxd%(ip%))
temp! = ConvertElmToOxd!(unkwts!(ip%), sample(1).Elsyms$(ip%), sample(1).numcat%(ip%), sample(1).numoxd%(ip%))
oxygen! = oxygen! + (temp! - unkwts!(ip%))
End If

If VerboseMode% Then
msg$ = "DIFF WT%" & Format$(Format$(unkwts!(ip%), f83$), a80$)
Call IOWriteLog(msg$)
End If
End If
End If

' Calculate formula by difference based on current weight percents
If sample(1).DifferenceFormulaFlag% And sample(1).DifferenceFormula$ <> vbNullString Then
Call FormulaFormulaToSample(sample(1).DifferenceFormula$, FormulaTmpSample())
If ierror Then Exit Sub

' Calculate sum of composition skipping formula by difference elements
temp! = 0#
For chan% = 1 To sample(1).LastChan%
ip% = IPOS1%(FormulaTmpSample(1).LastChan%, sample(1).Elsyms$(chan%), FormulaTmpSample(1).Elsyms$())
If ip% = 0 Then
temp! = temp! + unkwts!(chan%)
End If
Next chan%

' Determine difference from 100%
temp! = 100# - temp!
If temp! < 0# Then temp! = 100#

' Add in formula by difference elements (search from 1 to LastChan in FormulaTmpSample())
For chan% = 1 To sample(1).LastChan%
ip% = IPOS1%(FormulaTmpSample(1).LastChan%, sample(1).Elsyms$(chan%), FormulaTmpSample(1).Elsyms$())
If ip% > 0 Then
ipp% = IPOS1B(Int(1), FormulaTmpSample(1).LastChan%, sample(1).Elsyms$(chan%), FormulaTmpSample(1).Elsyms$())
If ipp% > 0 Then
unkwts!(chan%) = FormulaTmpSample(1).ElmPercents!(ipp%) * temp! / 100#
End If
End If
Next chan%

analysis.TotalPercent! = 100#
End If

' Check for bad total during iteration
If analysis.TotalPercent! <= ALPMinTotal! Then GoTo AFactorSmpBadTotal

' Add stoichiometric oxygen to oxygen channel
If sample(1).OxygenChannel% > 0 Then
unkwts!(sample(1).OxygenChannel%) = unkwts!(sample(1).OxygenChannel%) + oxygen!
End If

' Compute change in weight percents since last iteration
maxdiff = 0#
For chan% = 1 To sample(1).LastChan%
If sample(1).DisableQuantFlag%(chan%) = 0 Or (sample(1).DisableQuantFlag%(chan%) = 1 And sample(1).OxideOrElemental% = 1 And sample(1).OxygenChannel% = chan%) Then
diff! = Abs(unkwts!(chan%) - analysis.WtPercents!(chan%))
If diff! > maxdiff! Then maxdiff! = diff!
analysis.UnkKrats!(chan%) = kratios2!(chan%)
analysis.WtPercents!(chan%) = unkwts!(chan%)
analysis.Elsyms$(chan%) = sample(1).Elsyms$(chan%)
End If
Next chan%

' Check for convergence
If maxdiff! <= ALPMinToler! Then GoTo 6000

If VerboseMode Then
msg$ = "UNK WT% "
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(analysis.WtPercents!(chan%), f83$), a80$)
Next chan%
Call IOWriteLog(msg$)

msg$ = "UNK BETA"
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.UnkBetas!(chan%), f84$), a80$)
Next chan%
Call IOWriteLog(msg$)
End If

' Compute sum of alpha factors for each element (beta factors)
Call AFactorBeta(analysis, sample())
If ierror Then Exit Sub

' Not converged, try again
analysis.ZAFIter! = analysis.ZAFIter! + 1
Next iter%

' If we get here, convergence failed
msg$ = "WARNING in AFactorSmp- Convergence failed on line " & Str$(sample(1).Linenumber&(row%))
Call IOWriteLog(msg$)

' Converged, calculate actual excess oxygen (if calculating oxygen by stoichiometry)
6000:

If sample(1).OxideOrElemental% = 1 And sample(1).OxygenChannel% > 0 Then
oxygen! = analysis.WtPercents!(sample(1).OxygenChannel%) - excess!
analysis.WtPercents!(sample(1).OxygenChannel%) = excess!
End If

' Calculate total, z-bar, etc
Call ZAFCalZBar(oxygen!, analysis, sample())
If ierror Then Exit Sub

' Calculate atomic percents
If sample(1).AtomicPercentFlag Then
Call ConvertWeightToAtomic(sample(1).LastChan%, analysis.AtomicWeights!(), analysis.WtPercents!(), analysis.AtPercents!())
If ierror Then Exit Sub
For chan% = 1 To sample(1).LastChan%
analysis.AtPercents!(chan%) = 100# * analysis.AtPercents!(chan%)
Next chan%
End If

' Calculate oxide percents
If sample(1).DisplayAsOxideFlag Then
Call ConvertWeightToOxide(sample(1).LastChan%, analysis.AtomicWeights!(), sample(1).numcat%(), sample(1).numoxd%(), analysis.WtPercents!(), excess!, analysis.OxPercents!())
If ierror Then Exit Sub
End If

If Not DebugMode Then Exit Sub

' Debug mode, type out
msg$ = vbCrLf & "ELEMENT "
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%), a80$)
Next chan%
msg$ = msg$ & Format$("Total", a80$)
Call IOWriteLog(msg$)

msg$ = "UNK KRAT"
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(kratios2!(chan%), f84), a80$)
Next chan%
Call IOWriteLog(msg$)

msg$ = "UNK WT% "
sum! = 0#
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(analysis.WtPercents!(chan%), f83), a80$)
sum! = sum! + analysis.WtPercents!(chan%)
Next chan%
msg$ = msg$ & Format$(Format$(sum!, f83), a80$)
Call IOWriteLog(msg$)

If sample(1).AtomicPercentFlag Then
msg$ = "UNK AT% "
sum! = 0#
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(analysis.AtPercents!(chan%), f83), a80$)
sum! = sum! + analysis.AtPercents!(chan%)
Next chan%
msg$ = msg$ & Format$(Format$(sum!, f83), a80$)
Call IOWriteLog(msg$)
End If

If sample(1).DisplayAsOxideFlag Then
msg$ = "UNK OX% "
sum! = 0#
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(analysis.OxPercents!(chan%), f83), a80$)
sum! = sum! + analysis.OxPercents!(chan%)
Next chan%
msg$ = msg$ & Format$(Format$(sum!, f83), a80$)
Call IOWriteLog(msg$)
End If

msg$ = "UNK BETA"
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.UnkBetas!(chan%), f84), a80$)
Next chan%
Call IOWriteLog(msg$)

msg$ = "ALPITER " & Format$(Format$(analysis.ZAFIter!, f84), a80$)
Call IOWriteLog(msg$ & vbCrLf)

Exit Sub

' Errors
AFactorSmpError:
MsgBox Error$, vbOKOnly + vbCritical, "AFactorSmp"
ierror = True
Exit Sub

AFactorSmpBadUnkBetas:
msg$ = "WARNING in AFactorSmp- Bad unknown beta factor (" & Format$(analysis.UnkBetas!(chan%)) & ") on " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " on channel " & Format$(chan%)
Call IOWriteLog(msg$)
zerror = True
Exit Sub

AFactorSmpBadTotal:
msg$ = "WARNING in AFactorSmp- Insufficient total on line " & Str$(sample(1).Linenumber&(row%))
Call IOWriteLog(msg$)
zerror = True
Exit Sub

AFactorSmpBadStdPercents:
msg$ = "Insufficient standard percents on " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " on channel " & Format$(chan%)
MsgBox msg$, vbOKOnly + vbExclamation, "AFactorSmp"
ierror = True
Exit Sub

AFactorSmpBadStdCounts:
msg$ = "Insufficient standard counts on " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " on channel " & Format$(chan%)
MsgBox msg$, vbOKOnly + vbExclamation, "AFactorSmp"
ierror = True
Exit Sub

AFactorSmpBadStdBetas:
msg$ = "Bad standard beta factor on " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " on channel " & Format$(chan%)
MsgBox msg$, vbOKOnly + vbExclamation, "AFactorSmp"
ierror = True
Exit Sub

End Sub

Sub AFactorStd(row As Integer, analysis As TypeAnalysis, sample() As TypeSample, stdsample() As TypeSample)
' Calculates standard beta factors for a standard

ierror = False
On Error GoTo AFactorStdError

Dim i As Integer, ip As Integer
Dim sum As Single

' Check for missing elements in alpha-factor look-up tables
Call AFactorLoadFactors(analysis, stdsample())
If ierror Then Exit Sub

' Load the element arrays
Call ElementGetData(stdsample())
If ierror Then Exit Sub

' Calculate sum of standard (skip duplicated elements)
sum! = 0#
For i% = 1 To stdsample(1).LastChan%
analysis.WtPercents!(i%) = 0#
ip% = IPOS8(i%, stdsample(1).Elsyms$(i%), stdsample(1).Xrsyms$(i%), stdsample())
If ip% = 0 Then
sum! = sum! + stdsample(1).ElmPercents!(i%)
End If
Next i%

' Check for valid total
If sum! <= 0# Then GoTo AFactorStdBadTotal
If sum! < 98# Then
msg$ = "WARNING in AFactorStd- standard " & Str$(StandardNumbers%(row%)) & " total is " & Str$(sum!)
Call IOWriteLog(msg$)
End If

' Normalize standard weight percents to 100 (modifed 3/09/2007 to avoid issue with duplicate elements)
For i% = 1 To stdsample(1).LastChan%
ip% = IPOS8(i%, stdsample(1).Elsyms$(i%), stdsample(1).Xrsyms$(i%), stdsample()) ' find if element is duplicated
If Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0) Then
analysis.WtPercents!(i%) = 100# * stdsample(1).ElmPercents!(i%) / sum!
End If
Next i%
analysis.TotalPercent! = 100#

' Calculate standard beta-factors for each element in "stdpcnts" array
Call AFactorBeta(analysis, stdsample())
If ierror Then Exit Sub

' Type if in debug mode the analyzed element betas
If VerboseMode Then
Call AFactorTypeStandard(analysis, stdsample())
If ierror Then Exit Sub
End If

' Load standard beta factors for analyzed elements only
For i% = 1 To sample(1).LastElm%

' See if this standard is the assigned standard for this element
If sample(1).StdAssigns%(i%) <> stdsample(1).number% Then GoTo 8400

ip% = IPOS5(Int(1), i%, sample(), stdsample())  ' check element, xray, motor crystal
If ip% = 0 Then GoTo 8400

' Load assigned standard beta factor and weight percent
analysis.StdAssignsBetas!(i%) = analysis.UnkBetas!(ip%)
analysis.StdAssignsPercents!(i%) = analysis.WtPercents!(ip%)
8400:
Next i%

' Un-normalize weight percents to original sum
For i% = 1 To sample(1).LastChan%
ip% = IPOS8(i%, stdsample(1).Elsyms$(i%), stdsample(1).Xrsyms$(i%), stdsample()) ' find if element is duplicated
If Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0) Then
analysis.WtPercents!(i%) = sum! * analysis.WtPercents!(i%) / 100#
End If
Next i%

' Calculate total and zbar for standard sample. Oxygen is not calculated for standards!!!!
Call ZAFCalZBar(CSng(0#), analysis, stdsample())
If ierror Then Exit Sub

' Load standard arrays for this standard into the standard list arrays
For i% = 1 To sample(1).LastChan%
ip% = IPOS5(Int(1), i%, sample(), stdsample())  ' check element, xray, motor crystal
If ip% = 0 Then GoTo 8500

' Load the standard percents and z-bar
analysis.Elsyms$(i%) = stdsample(1).Elsyms$(ip%)
analysis.StdPercents!(row%, i%) = stdsample(1).ElmPercents!(ip%)
analysis.StdZbars!(row%) = analysis.Zbar!

' Load standard arrays
If i% <= sample(1).LastElm% Then
analysis.StdBetas!(row%, i%) = analysis.UnkBetas!(ip%)
End If

8500:
Next i%

Exit Sub

' Errors
AFactorStdError:
MsgBox Error$, vbOKOnly + vbCritical, "AFactorStd"
ierror = True
Exit Sub

AFactorStdBadTotal:
msg$ = "Standard " & Str$(StandardNumbers%(row%)) & " has a zero or less total sum"
MsgBox msg$, vbOKOnly + vbExclamation, "AFactorStd"
ierror = True
Exit Sub

End Sub

Sub AFactorTypeStandard(analysis As TypeAnalysis, stdsample() As TypeSample)
' Type standard a-factor calculation

ierror = False
On Error GoTo AFactorTypeStandardError

Dim i As Integer
Dim temp As Single

' Type standard name
msg$ = StandardLoadDescription$(stdsample())
If ierror Then Exit Sub
Call IOWriteLog(vbCrLf & msg$)

msg$ = "Standard Z-bar: " & Str$(analysis.Zbar!) & vbCrLf
Call IOWriteLog(msg$)

' Type elements
msg$ = "ELEM: "
For i% = 1 To stdsample(1).LastChan%
msg$ = msg$ & Format$(stdsample(1).Elsyup$(i%), a80$)
Next i%
Call IOWriteLog(msg$)

' Type out weight percents
msg$ = "ELWT: "
For i% = 1 To stdsample(1).LastChan%
msg$ = msg$ & Format$(Format$(stdsample(1).ElmPercents!(i%), f83$), a80$)
Next i%
Call IOWriteLog(msg$)

' Type out weight percents (normalized)
msg$ = "NRWT: "
For i% = 1 To stdsample(1).LastChan%
msg$ = msg$ & Format$(Format$(analysis.WtPercents!(i%), f83$), a80$)
Next i%
Call IOWriteLog(msg$)

' Type out beta factors
msg$ = "BETA: "
For i% = 1 To stdsample(1).LastChan%
msg$ = msg$ & Format$(Format$(analysis.UnkBetas!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

' Type out derived k-factor
msg$ = "KRAT: "
For i% = 1 To stdsample(1).LastChan%
temp! = 0#
If analysis.UnkBetas!(i%) <> 0# Then temp! = analysis.WtPercents!(i%) / analysis.UnkBetas!(i%)
msg$ = msg$ & Format$(Format$(temp! / 100#, f84$), a80$)
Next i%
Call IOWriteLog(msg$)

' Skip a space
Call IOWriteLog(vbNullString)

Exit Sub

' Errors
AFactorTypeStandardError:
MsgBox Error$, vbOKOnly + vbCritical, "AFactorTypeStandard"
ierror = True
Exit Sub

End Sub

Sub AFactorAFALoadFactors(analysis As TypeAnalysis, sample() As TypeSample)
' Read the analytical conditions from the AFA file

ierror = False
On Error GoTo AFactorAFALoadFactorsError

Dim astring As String, bstring As String
Dim cstring As String, dstring As String
Dim estring As String, fstring As String
Dim i As Integer, j As Integer

' Open the alpha-factor ASCII file
Open AFactorDataFile For Input As #AFactorDataFileNumber%

' Read run info
Input #AFactorDataFileNumber%, ProbeDataFile$
Input #AFactorDataFileNumber%, MDBUserName$
Input #AFactorDataFileNumber%, MDBFileTitle$
Input #AFactorDataFileNumber%, astring$

Input #AFactorDataFileNumber%, sample(1).kilovolts!, sample(1).takeoff!
Input #AFactorDataFileNumber%, sample(1).LastElm%, sample(1).LastChan%, sample(1).OxideOrElemental%, NumberofStandards%

' Read column labels
Input #AFactorDataFileNumber%, astring$, bstring$, cstring$, dstring$, estring$, fstring$

' Read standard counts, weight percents and beta factors
For j% = 1 To sample(1).LastElm%
Input #AFactorDataFileNumber%, sample(1).Elsyms$(j%), sample(1).Xrsyms$(j%), sample(1).MotorNumbers%(j%), sample(1).CrystalNames$(j%), sample(1).StdAssigns%(j%), analysis.StdAssignsPercents!(j%), analysis.StdAssignsBetas!(j%), analysis.StdAssignsCounts!(j%)
Next j%

' Read column labels
Input #AFactorDataFileNumber%, astring$, bstring$

' Read oxide and cations
For j% = 1 To sample(1).LastChan%
Input #AFactorDataFileNumber%, sample(1).Elsyms$(j%), sample(1).Xrsyms$(j%), sample(1).numoxd%(j%), sample(1).numcat%(j%)
Next j%

' Read interference calibration
Input #AFactorDataFileNumber%, astring$, bstring$, cstring$, dstring$

' Interference elements
For j% = 1 To sample(1).LastElm%
Input #AFactorDataFileNumber%, sample(1).Elsyms$(j%), sample(1).Xrsyms$(j%)
For i% = 1 To MAXINTF%
Input #AFactorDataFileNumber%, sample(1).StdAssignsIntfElements$(i%, j%)
Next i%
Next j%

' Interference std assignments
For j% = 1 To sample(1).LastElm%
Input #AFactorDataFileNumber%, sample(1).Elsyms$(j%), sample(1).Xrsyms$(j%)
For i% = 1 To MAXINTF%
Input #AFactorDataFileNumber%, sample(1).StdAssignsIntfStds%(i%, j%)
Next i%
Next j%

' Interference counts
For j% = 1 To sample(1).LastElm%
Input #AFactorDataFileNumber%, sample(1).Elsyms$(j%), sample(1).Xrsyms$(j%)
For i% = 1 To MAXINTF%
Input #AFactorDataFileNumber%, analysis.StdAssignsIntfCounts!(i%, j%)
Next i%
Next j%

' Standard percents
For i% = 1 To NumberofStandards%
Input #AFactorDataFileNumber%, StandardNumbers%(i%)
Next i%

For j% = 1 To sample(1).LastElm%
Input #AFactorDataFileNumber%, sample(1).Elsyms$(j%), sample(1).Xrsyms$(j%)
For i% = 1 To NumberofStandards%
Input #AFactorDataFileNumber%, analysis.StdPercents!(i%, j%)
Next i%
Next j%

' Read MAN background coeff column labels
Input #AFactorDataFileNumber%, astring$, bstring$, cstring$, dstring$, estring$

' MAN fit coefficients
For j% = 1 To sample(1).LastElm%
Input #AFactorDataFileNumber%, sample(1).Elsyms$(j%), sample(1).Xrsyms$(j%), sample(1).BackgroundTypes%(j%), sample(1).MANAbsCorFlags%(j%)
For i% = 1 To MAXCOEFF%
Input #AFactorDataFileNumber%, analysis.MANFitCoefficients!(i%, j%)
Next i%
Next j%

' Read alpha factor absorbers column labels
For i% = 1 To sample(1).LastChan%
Input #AFactorDataFileNumber%, sample(1).Elsyms$(i%)
Next i%

' Loop on emitters
For j% = 1 To sample(1).LastElm%

' Loop on absorbers
Input #AFactorDataFileNumber%, sample(1).Elsyms$(j%), sample(1).Xrsyms$(j%)
For i% = 1 To sample(1).LastChan%
Input #AFactorDataFileNumber%, alpha1(j%, i%)
Next i%

Input #AFactorDataFileNumber%, sample(1).Elsyms$(j%), sample(1).Xrsyms$(j%)
For i% = 1 To sample(1).LastChan%
Input #AFactorDataFileNumber%, alpha2(j%, i%)
Next i%

Input #AFactorDataFileNumber%, sample(1).Elsyms$(j%), sample(1).Xrsyms$(j%)
For i% = 1 To sample(1).LastChan%
Input #AFactorDataFileNumber%, alpha3(j%, i%)
Next i%

Input #AFactorDataFileNumber%, sample(1).Elsyms$(j%), sample(1).Xrsyms$(j%)
For i% = 1 To sample(1).LastChan%
Input #AFactorDataFileNumber%, alpha4(j%, i%)
Next i%

Next j%

' Close file
Close #AFactorDataFileNumber%

Exit Sub

' Errors
AFactorAFALoadFactorsError:
MsgBox Error$, vbOKOnly + vbCritical, "AFactorAFALoadFactors"
Close #AFactorDataFileNumber%
ierror = True
Exit Sub

End Sub

Sub AFactorReturnAFactors(mode As Integer, npts As Integer, xdata() As Single, ydata() As Single, acoeff() As Single, stddev As Single)
' Return the calculated a-factors for plotting
' mode = 1 return first emitter
' mode = 2 return second emitter

ierror = False
On Error GoTo AFactorReturnAFactorsError

Dim i As Integer

ReDim xdata(1 To MAXBINARY%) As Single
ReDim ydata(1 To MAXBINARY%) As Single
ReDim acoeff(1 To MAXCOEFF4%) As Single

' Load return arrays
If mode% = 1 Then
npts% = AlphaNpts1%
For i% = 1 To MAXBINARY%
xdata!(i%) = AlphaXdata1!(i%)
ydata!(i%) = AlphaYdata1!(i%)
Next i%

For i% = 1 To MAXCOEFF4%
acoeff!(i%) = AlphaAcoeff1!(i%)
Next i%

stddev! = AlphaAstddev1!

Else
npts% = AlphaNpts2%
For i% = 1 To MAXBINARY%
xdata!(i%) = AlphaXdata2!(i%)
ydata!(i%) = AlphaYdata2!(i%)
Next i%

For i% = 1 To MAXCOEFF4%
acoeff!(i%) = AlphaAcoeff2!(i%)
Next i%

stddev! = AlphaAstddev2!
End If

Exit Sub

' Errors
AFactorReturnAFactorsError:
MsgBox Error$, vbOKOnly + vbCritical, "AFactorReturnAFactors"
ierror = True
Exit Sub

End Sub

Sub AFactorPenepmaReadMatrix(wout() As Single, rout() As Single, eout() As String, xout() As String, zout() As Integer, notfoundA As Boolean, notfoundB As Boolean, sample() As TypeSample)
' Read the matrix database for the specified energy, emitter, xray and matrix element of the specified binary

ierror = False
On Error GoTo AFactorPenepmaReadMatrixError

Dim n As Integer, m As Integer, ibin As Integer
Dim EmitterElement As Integer, EmitterXray As Integer, MatrixElement As Integer
Dim EmitterTakeOff As Single
Dim EmitterKilovolts As Single
Dim astring As String, tmsg As String

Dim tKratios1(1 To MAXBINARY%) As Double
Dim tKratios2(1 To MAXBINARY%) As Double

' Load the emitter/absorber arrays
For ibin% = 1 To 2

If ibin% = 1 Then
EmitterTakeOff! = sample(1).TakeoffArray!(1)
EmitterKilovolts! = sample(1).KilovoltsArray!(1)

EmitterElement% = sample(1).AtomicNums%(1)
EmitterXray% = sample(1).XrayNums%(1)
MatrixElement% = sample(1).AtomicNums%(2)
End If

If ibin% = 2 Then
EmitterTakeOff! = sample(1).TakeoffArray!(2)
EmitterKilovolts! = sample(1).KilovoltsArray!(2)

EmitterElement% = sample(1).AtomicNums%(2)
EmitterXray% = sample(1).XrayNums%(2)
MatrixElement% = sample(1).AtomicNums%(1)
End If

' Get the specified binary data
If ibin% = 1 Then
If EmitterXray% < MAXRAY% Then
Call Penepma12MatrixReadMDB2(EmitterTakeOff!, EmitterKilovolts!, EmitterElement%, EmitterXray%, MatrixElement%, tKratios1#(), notfoundA)
If ierror Then Exit Sub
End If

Else
If EmitterXray% < MAXRAY% Then
Call Penepma12MatrixReadMDB2(EmitterTakeOff!, EmitterKilovolts!, EmitterElement%, EmitterXray%, MatrixElement%, tKratios2#(), notfoundB)
If ierror Then Exit Sub
End If
End If

' Check if found
If (ibin% = 1 And Not notfoundA And EmitterXray% < MAXRAY%) Or (ibin% = 2 And Not notfoundB And EmitterXray% < MAXRAY%) Then

' Output to log
If DebugMode Then
astring$ = "AFactorPenepmaReadMatrix: " & Symup$(EmitterElement%) & " " & Xraylo$(EmitterXray%) & " in " & Symup$(MatrixElement%)
Call IOWriteLog$(vbCrLf & astring$ & " at " & Format$(EmitterTakeOff!) & " degrees and " & Format$(EmitterKilovolts!) & " keV")
astring$ = Format$(vbTab & "Conc", a08$) & vbTab & Format$("Kratios", a08$)
Call IOWriteLog$(astring$)

For n% = 1 To MAXBINARY%
If ibin% = 1 Then
astring$ = vbTab & MiscAutoFormat$(BinaryRanges!(n%)) & vbTab & MiscAutoFormatD$(tKratios1#(n%))
Else
astring$ = vbTab & MiscAutoFormat$(BinaryRanges!(n%)) & vbTab & MiscAutoFormatD$(tKratios2#(n%))
End If
Call IOWriteLog$(astring$)
Next n%
End If

' Load weight percents and k-ratios for fit
For n% = 1 To MAXBINARY%
m% = (n% - 1) * 2 + ibin%

' Check for UsePenepmaKratiosLimitFlag
If Not UsePenepmaKratiosLimitFlag Or (UsePenepmaKratiosLimitFlag And BinaryRanges!(n%) <= PenepmaKratiosLimitValue!) Then
wout!(m%) = BinaryRanges!(n%)

If ibin% = 1 Then
rout!(m%) = tKratios1#(n%) / 100#
Else
rout!(m%) = tKratios2#(n%) / 100#
End If

' Load symbols for print out
If ibin% = 1 Then
eout$(m%) = sample(1).Elsyms$(1)
Else
eout$(m%) = sample(1).Elsyms$(2)
End If

If ibin% = 1 Then
xout$(m%) = sample(1).Xrsyms$(1)
Else
xout$(m%) = sample(1).Xrsyms$(2)
End If

' Load atomic numbers for look-up tables
If ibin% = 1 Then
zout%(m%) = sample(1).AtomicNums%(1)
Else
zout%(m%) = sample(1).AtomicNums%(2)
End If

End If
Next n%

' Penepma binary was not found. Save warning to CalcZAF.ERR
Else
If DebugMode Then
If EmitterXray% < MAXRAY% Then
tmsg$ = "Binary for " & Symup$(EmitterElement%) & " " & Xraylo$(EmitterXray%) & " in " & Symup$(MatrixElement%) & " at " & Format$(EmitterKilovolts!) & " keV, " & Format$(EmitterTakeOff!) & " degrees, was not found."
Call IOWriteError(tmsg$, "AFactorPenepmaReadmatrix")
If ierror Then Exit Sub
End If
End If
End If
Next ibin%

Exit Sub

' Errors
AFactorPenepmaReadMatrixError:
MsgBox Error$, vbOKOnly + vbCritical, "AFactorPenepmaReadMatrix"
ierror = True
Exit Sub

End Sub

Sub AFactorTypeAlphas(analysis As TypeAnalysis, sample() As TypeSample)
' Print out the alpha factors

ierror = False
On Error GoTo AFactorTypeAlphasError

Dim i As Integer, j As Integer
Dim ip As Integer, ipp As Integer, n As Integer

' Check for alpha factors (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% < 1 Or CorrectionFlag% > 4 Then
MsgBox "Alpha factor calculations are not selected (see Analytical | ZAF, Phi-Rho-Z, Alpha Factor and Calibration Curve Selections menu)", vbOKOnly + vbInformation, "AFactorTypeAlphas"
Exit Sub
End If

' Reload the element arrays
Call ElementGetData(sample())
If ierror Then Exit Sub

' Load elements to arrays
Call ZAFSetZAF(sample())
If ierror Then Exit Sub

' Re-load alpha-factors based on current selections
AllAFactorUpdateNeeded = True
Call AFactorLoadFactors(analysis, sample())
If ierror Then Exit Sub

' Print empirical alphas
If EmpiricalAlphaFlag = 2 Then
Call IOWriteLog(vbCrLf & "Empirical K-Ratio Alpha Factors:")

Call IOWriteLog("Xray  Matrix   Alpha1  Alpha2  Alpha3  Alpha4")
For n% = 1 To MAXEMPFAC%
If empfaco%(n%) = CorrectionFlag% Then
ip% = IPOS1(sample(1).LastElm%, empface$(n%), sample(1).Elsyms$())
ipp% = IPOS1(sample(1).LastChan%, empfaca$(n%), sample(1).Elsyms$())

If ip% <> 0 And ipp% <> 0 Then
msg$ = sample(1).Elsyms$(ip%) & " " & sample(1).Xrsyms$(ip%) & " in " & sample(1).Elsyms$(ipp%)
If empfaco%(n%) > 0 Then msg$ = msg$ & Format$(Format$(empfac1!(n%), f84$), a80$)
If empfaco%(n%) > 1 Then msg$ = msg$ & Format$(Format$(empfac2!(n%), f84$), a80$)
If empfaco%(n%) > 2 Then msg$ = msg$ & Format$(Format$(empfac3!(n%), f84$), a80$)
If empfaco%(n%) > 3 Then msg$ = msg$ & Format$(Format$(empfac4!(n%), f84$), a80$)
Call IOWriteLog(msg$)
End If
End If
Next n%
End If

' Print Penepma alphas (1 = do not use, 2 = use)
If UsePenepmaKratiosFlag% = 2 Then
Call IOWriteLog(vbCrLf & vbCrLf & "Penepma K-Ratio Alpha Factors:")

Call IOWriteLog("Xray  Matrix   Alpha1  Alpha2  Alpha3  Alpha4")
For i% = 1 To sample(1).LastChan%   ' absorbers
For j% = 1 To sample(1).LastElm%    ' emitters
If PenepmaKratiosFlag(j%, i%) Then

msg$ = sample(1).Elsyup$(j%) & " " & sample(1).Xrsyms$(j%) & " in " & sample(1).Elsyup$(i%) & "  "
msg$ = msg$ & Format$(Format$(alpha1!(j%, i%), f84$), a80$) & Format$(Format$(alpha2!(j%, i%), f84$), a80$) & Format$(Format$(alpha3!(j%, i%), f84$), a80$) & Format$(Format$(alpha4!(j%, i%), f84$), a80$) & "    *from Penepma 2012 Calculations"
Call IOWriteLog(msg$)

End If
Next j%
Next i%
End If

' If debug mode print all alpha factors
Call IOWriteLog(vbCrLf & "All Alpha Factors:")

' Write absorbers column labels
msg$ = Space$(8) & vbTab
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(sample(1).Elsyms$(i%), a80$) & vbTab
Next i%
Call IOWriteLog(msg$)

' Loop on emitters
For j% = 1 To sample(1).LastElm%

' Loop on absorbers
msg$ = Format$(sample(1).Elsyms$(j%) & " " & sample(1).Xrsyms$(j%), a80$) & vbTab
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(alpha1(j%, i%), f84$), a80$) & vbTab
Next i%
Call IOWriteLog(msg$)

If CorrectionFlag% > 1 Then
msg$ = Format$(sample(1).Elsyms$(j%) & " " & sample(1).Xrsyms$(j%), a80$) & vbTab
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(alpha2(j%, i%), f84$), a80$) & vbTab
Next i%
Call IOWriteLog(msg$)
End If

If CorrectionFlag% > 2 Then
msg$ = Format$(sample(1).Elsyms$(j%) & " " & sample(1).Xrsyms$(j%), a80$) & vbTab
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(alpha3(j%, i%), f84$), a80$) & vbTab
Next i%
Call IOWriteLog(msg$)
End If

If CorrectionFlag% > 3 Then
msg$ = Format$(sample(1).Elsyms$(j%) & " " & sample(1).Xrsyms$(j%), a80$) & vbTab
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(alpha4(j%, i%), f84$), a80$) & vbTab
Next i%
Call IOWriteLog(msg$)
End If

Next j%

Exit Sub

' Errors
AFactorTypeAlphasError:
MsgBox Error$, vbOKOnly + vbCritical, "AFactorTypeAlphas"
ierror = True
Exit Sub

End Sub

Sub AFactorAlpha(analysis As TypeAnalysis, sample() As TypeSample)
' Calculate the alpha factors only for the passed k-ratios

ierror = False
On Error GoTo AFactorAlphaError

Dim ip As Integer
Dim i As Integer, row As Integer
Dim k As Single, c As Single

Dim afactors(1 To MAXROW%, 1 To MAXCHAN%) As Single

msg$ = vbCrLf & "Calculated alpha factors (for binary compositions only!)..."
Call IOWriteLog(msg$)

' Type out symbols for count time data lines for analyzed elements
msg$ = "ELEM: "
For i% = 1 To sample(1).LastElm%
If sample(1).DisableAcqFlag%(i%) = 1 Then
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & "-D", a80$)
Else
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%), a80$)
End If
Next i%
Call IOWriteLog(msg$)

' Get position of standard in standard array
ip% = IPOS2(NumberofStandards%, sample(1).number%, StandardNumbers%())
If ip% = 0 Then GoTo AFactorAlphaNotLoaded

For row% = 1 To sample(1).Datarows%
For i% = 1 To sample(1).LastElm%
If analysis.StdPercents!(ip%, i%) > NOT_ANALYZED_VALUE_SINGLE! Then

' Calculate alpha factor for this composition
c! = analysis.StdPercents!(ip%, i%) / 100#
k! = analysis.CalData!(row%, i%)
afactors!(row%, i%) = ((c! / k!) - c!) / (1 - c!)        ' calculate binary alpha factors

End If
Next i%

' For each row type out alpha factors (all lines)
msg$ = Format$(Format$(sample(1).Linenumber&(row%), i50$), a50$)
If sample(1).LineStatus(row%) Then msg$ = msg$ & "G"
If Not sample(1).LineStatus(row%) Then msg$ = msg$ & "B"

' Type corrected counts
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(afactors!(row%, i%), f84$), a80$)
Next i%

Call IOWriteLog(msg$)
Next row%

Exit Sub

' Errors
AFactorAlphaError:
MsgBox Error$, vbOKOnly + vbCritical, "AFactorAlpha"
ierror = True
Exit Sub

AFactorAlphaNotLoaded:
msg$ = "Standard " & Str$(sample(1).Type%) & " is not in the standard list"
MsgBox msg$, vbOKOnly + vbExclamation, "AFactorAlphaNotLoaded"
ierror = True
Exit Sub

End Sub

