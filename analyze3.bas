Attribute VB_Name = "CodeANALYZE3"
' (c) Copyright 1995-2017 by John J. Donovan
Option Explicit

' Maximum precision and detection limits
Global tPrecision(1 To MAXCHAN%) As Single, tDetection(1 To MAXCHAN%) As Single     ' for total statistics on sum
Global aPrecision(1 To MAXCHAN%) As Single, aDetection(1 To MAXCHAN%) As Single     ' for average total statistics on sum

' Unknown Row arrays (for averaging)
Dim RowUnkZAFIters(1 To MAXROW%)  As Single
Dim RowUnkMANIters(1 To MAXROW%)  As Single

Dim RowUnkTotalPercents(1 To MAXROW%) As Single
Dim RowUnkTotalOxygens(1 To MAXROW%)  As Single
Dim RowUnkTotalCations(1 To MAXROW%)  As Single
Dim RowUnkTotalAtoms(1 To MAXROW%)  As Single
Dim RowUnkCalculatedOxygens(1 To MAXROW%) As Single
Dim RowUnkExcessOxygens(1 To MAXROW%) As Single
Dim RowUnkZbars(1 To MAXROW%)  As Single
Dim RowUnkAtomicWeights(1 To MAXROW%)  As Single
Dim RowUnkOxygenFromHalogens(1 To MAXROW%) As Single
Dim RowUnkHalogenCorrectedOxygen(1 To MAXROW%) As Single
Dim RowUnkChargeBalance(1 To MAXROW%)  As Single
Dim RowUnkFeChargeBalance(1 To MAXROW%)  As Single

Dim RowUnkKRaws(1 To MAXROW%, 1 To MAXCHAN%) As Single
Dim RowUnkKrats(1 To MAXROW%, 1 To MAXCHAN%) As Single
Dim RowUnkZAFCors(1 To MAXROW%, 1 To MAXCHAN%) As Single
Dim RowUnkBetas(1 To MAXROW%, 1 To MAXCHAN%) As Single

Dim RowUnkZCors(1 To MAXROW%, 1 To MAXCHAN%) As Single
Dim RowUnkACors(1 To MAXROW%, 1 To MAXCHAN%) As Single
Dim RowUnkFCors(1 To MAXROW%, 1 To MAXCHAN%) As Single

Dim RowUnkMACs(1 To MAXROW%, 1 To MAXCHAN%) As Single   ' compound MACs

Dim RowUnkPeakToBgds(1 To MAXROW%, 1 To MAXCHAN%) As Single
Dim RowStdAssignsCounts(1 To MAXROW%, 1 To MAXCHAN%) As Single

' Interference, MAN, Volatile and APF Correction factors
Dim RowUnkIntfCors(1 To MAXROW%, 1 To MAXCHAN%) As Single
Dim RowUnkMANAbsCors(1 To MAXROW%, 1 To MAXCHAN%) As Single
Dim RowUnkAPFCors(1 To MAXROW%, 1 To MAXCHAN%) As Single
Dim RowUnkVolElCors(1 To MAXROW%, 1 To MAXCHAN%) As Single
Dim RowUnkVolElDevs(1 To MAXROW%, 1 To MAXCHAN%) As Single

' Calibration curve arrays
Dim RowCurve1Coeffs(1 To MAXROW%, 1 To MAXCHAN%) As Single
Dim RowCurve2Coeffs(1 To MAXROW%, 1 To MAXCHAN%) As Single
Dim RowCurve3Coeffs(1 To MAXROW%, 1 To MAXCHAN%) As Single
Dim RowCurveFits(1 To MAXROW%, 1 To MAXCHAN%) As Single

' MAN/Interf/VolEl/APF convergence difference
Dim convergencedifference(1 To MAXCHAN%) As Single

' Standard variances
Dim StandardPublishedValues(1 To MAXCHAN%) As Single
Dim StandardPercentVariances(1 To MAXCHAN%) As Single
Dim StandardAlgebraicDifferences(1 To MAXCHAN%) As Single

Dim stdsample(1 To 1) As TypeSample

Sub AnalyzeWeightCalculate(linerow As Integer, firsttime As Boolean, zerror As Integer, analysis As TypeAnalysis, sample() As TypeSample, stdsample() As TypeSample)
' Called by routine AnalyzeSample to calculate the weight percents for the specified sample linerow

ierror = False
On Error GoTo AnalyzeWeightCalculateError

Dim alldone As Boolean
Dim i As Integer, j As Integer, jmax As Integer
Dim chan As Integer, ip As Integer, ipp As Integer
Dim ippp As Integer, ipppp As Integer
Dim temp As Single, excess As Single
Dim MaxMANIter As Integer

ReDim uncts(1 To MAXCHAN%) As Single
ReDim oldcts(1 To MAXCHAN%) As Single

' Reset ZAF error flag (non-fatal error)
zerror = False

' Write debug info
If DebugMode Then
Call IOWriteLog(vbCrLf & "Sample Line Number: " & Str$(sample(1).Linenumber&(linerow%)))

msg$ = "Elements:" & vbCrLf
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).Elsyms$(i%), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "Element Standards:" & vbCrLf
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(sample(1).StdAssigns%(i%), i80$), a80$)
Next i%
Call IOWriteLog(msg$)

' 0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters
If CorrectionFlag% = 0 Or CorrectionFlag% = MAXCORRECTION% Then
msg$ = "Element Standard K-Factors:" & vbCrLf
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.StdAssignsKfactors!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

Else
msg$ = "Element Standard Beta-Factors:" & vbCrLf
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.StdAssignsBetas!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)
End If

msg$ = "Element Standard Counts (MAN/Interference corrected):" & vbCrLf
For i% = 1 To sample(1).LastElm%
If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(analysis.StdAssignsCounts!(i%), f81$), a80$)
Else
msg$ = msg$ & Format$(Format$(analysis.StdAssignsCounts!(i%), f82$), a80$)
End If
Next i%
Call IOWriteLog(msg$)

msg$ = "Element Standard Percents:" & vbCrLf
For i% = 1 To sample(1).LastElm%
ip% = IPOS2(NumberofStandards%, sample(1).StdAssigns%(i%), StandardNumbers%())
If ip% > 0 Then
msg$ = msg$ & Format$(Format$(analysis.StdPercents!(ip%, i%), f83$), a80$)
End If
Next i%
Call IOWriteLog(msg$)

' Calculate maximum number of interference assignments in sample
jmax% = UpdateGetMaxInterfAssign(sample())

Call IOWriteLog(vbNullString)
For j% = 1 To jmax%
msg$ = "Interfering " & Str$(j%) & " Elements:" & vbCrLf
For i% = 1 To sample(1).LastElm%
If sample(1).StdAssignsIntfElements$(j%, i%) <> vbNullString Then
msg$ = msg$ & Format$(sample(1).StdAssignsIntfElements$(j%, i%) & " (" & sample(1).StdAssignsIntfXrays$(j%, i%) & ")", a80$)
Else
msg$ = msg$ & Space$(8)
End If
Next i%
Call IOWriteLog(msg$)
Next j%

For j% = 1 To jmax%
msg$ = "Interfering Element " & Str$(j%) & " Standards:" & vbCrLf
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(sample(1).StdAssignsIntfStds%(j%, i%), i80$), a80$)
Next i%
Call IOWriteLog(msg$)
Next j%

For j% = 1 To jmax%
msg$ = "Interference Standard " & Str$(j%) & " Counts:" & vbCrLf
For i% = 1 To sample(1).LastElm%
If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(analysis.StdAssignsIntfCounts!(j%, i%), f81$), a80$)
Else
msg$ = msg$ & Format$(Format$(analysis.StdAssignsIntfCounts!(j%, i%), f82$), a80$)
End If
Next i%
Call IOWriteLog(msg$)
Next j%

For j% = 1 To jmax%
msg$ = "Interference Standard " & Str$(j%) & " Percents:" & vbCrLf
For i% = 1 To sample(1).LastElm%
ip% = IPOS2(NumberofStandards%, sample(1).StdAssignsIntfStds%(j%, i%), StandardNumbers%())
If ProbeDataFileVersionNumber! > 6.41 Then
ipp% = IPOS1A(sample(1).LastElm%, sample(1).StdAssignsIntfElements$(j%, i%), sample(1).StdAssignsIntfXrays$(j%, i%), sample(1).Elsyms$(), sample(1).Xrsyms$())
Else
ipp% = IPOS1(sample(1).LastElm%, sample(1).StdAssignsIntfElements$(j%, i%), sample(1).Elsyms$())
End If
If ip% > 0 And ipp% > 0 Then
msg$ = msg$ & Format$(Format$(analysis.StdPercents!(ip%, ipp%), f83$), a80$)
Else
msg$ = msg$ & Format$(Format$(0#, f83$), a80$)
End If
Next i%
Call IOWriteLog(msg$)
Next j%
Call IOWriteLog(vbNullString)

' 0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters
If CorrectionFlag% = 0 Or CorrectionFlag% = MAXCORRECTION% Then
For j% = 1 To jmax%
msg$ = "Interference Standard " & Str$(j%) & " Absorption Corrections:" & vbCrLf
For i% = 1 To sample(1).LastElm%
ip% = IPOS2(NumberofStandards%, sample(1).StdAssignsIntfStds%(j%, i%), StandardNumbers%())
If ip% > 0 Then
msg$ = msg$ & Format$(Format$(analysis.StdZAFCors!(1, ip%, i%), f84$), a80$)
Else
msg$ = msg$ & Format$(Format$(0#, f83$), a80$)
End If
Next i%
Call IOWriteLog(msg$)
Next j%

For j% = 1 To jmax%
msg$ = "Interference Standard " & Str$(j%) & " Atomic Number Corrections:" & vbCrLf
For i% = 1 To sample(1).LastElm%
ip% = IPOS2(NumberofStandards%, sample(1).StdAssignsIntfStds%(j%, i%), StandardNumbers%())
If ip% > 0 Then
msg$ = msg$ & Format$(Format$(analysis.StdZAFCors!(3, ip%, i%), f84$), a80$)
Else
msg$ = msg$ & Format$(Format$(0#, f83$), a80$)
End If
Next i%
Call IOWriteLog(msg$)
Next j%

' Alpha-factors
Else
For j% = 1 To jmax%
msg$ = "Interference Standard " & Str$(j%) & " Beta Corrections:" & vbCrLf
For i% = 1 To sample(1).LastElm%
ip% = IPOS2(NumberofStandards%, sample(1).StdAssignsIntfStds%(j%, i%), StandardNumbers%())
If ip% > 0 Then
msg$ = msg$ & Format$(Format$(analysis.StdBetas!(ip%, i%), f84$), a80$)
Else
msg$ = msg$ & Format$(Format$(0#, f83$), a80$)
End If
Next i%
Call IOWriteLog(msg$)
Next j%
End If

' Print MAN counts and fit coefficients
msg$ = "Elements:" & vbCrLf
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).Elsyms$(i%), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "MAN Assignments:"
Call IOWriteLog(msg$)

' Calculate actual maximum number of MAN assignments for this sample
jmax% = UpdateGetMaxMANAssign(sample())

For j% = 1 To jmax%
msg$ = vbNullString
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(sample(1).MANStdAssigns%(j%, i%), i80$), a80$)
Next i%
Call IOWriteLog(msg$)
Next j%

msg$ = "BackgroundTypes:" & vbCrLf
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(bgdtypestrings$(sample(1).BackgroundTypes%(i%)), a80$)  ' 0=off-peak, 1=MAN, 2=multipoint
Next i%
Call IOWriteLog(msg$)

msg$ = "MAN Fit Orders:" & vbCrLf
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(sample(1).MANLinearFitOrders%(i%), i80$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "MAN Absorption Correction Flags:" & vbCrLf
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(sample(1).MANAbsCorFlags%(i%), i80$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "MAN Counts:"
Call IOWriteLog(msg$)
For j% = 1 To jmax%
msg$ = vbNullString
For i% = 1 To sample(1).LastElm%
If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(analysis.MANAssignsCounts!(j%, i%), f81$), a80$)
Else
msg$ = msg$ & Format$(Format$(analysis.MANAssignsCounts!(j%, i%), f82$), a80$)
End If
Next i%
Call IOWriteLog(msg$)
Next j%

msg$ = "MAN Fit Coefficients:"
Call IOWriteLog(msg$)
For j% = 1 To MAXCOEFF%
msg$ = vbNullString
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & MiscAutoFormat$(analysis.MANFitCoefficients!(j%, i%))
Next i%
Call IOWriteLog(msg$)
Next j%
End If

' Load specified weight percents for this sample
excess! = 0#
For chan% = 1 To sample(1).LastChan%
analysis.WtPercents!(chan%) = 0#
uncts!(chan%) = 0#
oldcts!(chan%) = 0#

' Load specified elements from standard sample if element is a specified
' channel. If sample is a standard, the standard sample was loaded from the
' standard database. If sample is an unknown, the standard sample is simply
' a copy of the sample.
If chan% > sample(1).LastElm% Then

' Determine if element is in sample
ip% = IPOS11(sample(1).Elsyms$(chan%), stdsample())

' Load specified weight percent if not specified oxygen
If ip% > 0 Then
If chan% <> sample(1).OxygenChannel% Then
ippp% = IPOS1(sample(1).LastElm%, sample(1).Elsyms$(chan%), sample(1).Elsyms$())    ' is element already analysed?
If ippp% = 0 Then
analysis.WtPercents!(chan%) = stdsample(1).ElmPercents!(ip%)            ' no, so load
Else
If sample(1).DisableQuantFlag%(ippp%) = 1 Then analysis.WtPercents!(chan%) = stdsample(1).ElmPercents!(ip%)     ' if analyzed but disabled, then load anway
'If sample(1).DisableQuantFlag%(ippp%) = 1 And chan% <= sample(1).LastElm% Then analysis.WtPercents!(chan%) = stdsample(1).ElmPercents!(ip%)     ' if analyzed but disabled, then load anyway (this line handles the situation when element is present as disabled WDS, enabled EDS and also specified)
End If

' Zero if analyzing oxygen and sample is a standard (3/1/2004 code changes)
If sample(1).Type% = 1 Then
If sample(1).OxygenChannel% > 0 And sample(1).OxygenChannel% <= sample(1).LastElm% Then
If UCase$(Trim$(sample(1).Elsyms$(chan%))) = UCase$(Trim$(Symlo$(ATOMIC_NUM_OXYGEN%))) Then
analysis.WtPercents!(chan%) = 0#                                        ' zero specified value
ippp% = IPOS2(NumberofStandards%, sample(1).number%, StandardNumbers%())
analysis.StdPercents!(ippp%, chan%) = 0#                                ' fix PUBL: values
End If
End If
End If

' If element is oxygen, make sure that correct excess oxygen is specified based on the sample cations (standards only)
Else

If sample(1).Type% = 1 Then
analysis.WtPercents!(chan%) = ConvertTotalToExcessOxygen!(Int(1), sample(), stdsample())

' For unknowns, use specified oxygen weight percent
Else
analysis.WtPercents!(chan%) = stdsample(1).ElmPercents!(ip%)
End If

excess! = analysis.WtPercents!(chan%)   ' store excess oxygen for iteration
End If
End If

End If
Next chan%

' Iterate on the MAN, interference, volatile and APF corrections
MaxMANIter% = 100
analysis.MANIter! = 1#
analysis.Zbar! = 10.8    ' assume quartz z-bar for first MAN iteration
alldone = False

' Iterate on matrix and other compositionally dependent corrections
Do Until alldone

' Calculate MAN backgrounds, peak interferences, volatile element and APF corrections
Call AnalyzeWeightCorrect(linerow%, uncts!(), oldcts!(), alldone, firsttime, zerror, analysis, sample())
firsttime = False
If ierror Then Exit Sub

' Re-load excess oxygen each iteration since it is overwritten in ZAFSmp when calculated oxygen is added to already specified oxygen
If sample(1).OxygenChannel% > 0 Then
If DebugMode Then
msg$ = "Nominal excess oxygen weight percent: " & Str$(excess!)
Call IOWriteLog(msg$)
msg$ = "Current oxygen weight percent: " & Str$(analysis.WtPercents!(sample(1).OxygenChannel%))
Call IOWriteLog(msg$)
End If
analysis.WtPercents!(sample(1).OxygenChannel%) = excess!
End If

' ZAF or Phi-Rho-Z calculation for this data line (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% = 0 Then
Call ZAFSmp(linerow%, uncts!(), zerror, analysis, sample())
If ierror Then Exit Sub

' Alpha-factor calculation for this data line (note excess oxygen is passed because AFactorSmp cannot keep excess and calculated oxygen separated)
ElseIf CorrectionFlag% > 0 And CorrectionFlag% < 5 Then
Call AFactorSmp(linerow%, excess!, uncts!(), zerror, analysis, sample())
If ierror Then Exit Sub

' Fundamental parameter correction
ElseIf CorrectionFlag% = MAXCORRECTION% Then
'Call ZAFSmp3(linerow%, uncts!(), zerror, analysis, sample())
'If ierror Then Exit Sub
End If

' Check if alldone with AnalyzeWeightCorrect
If alldone Then Exit Do

' Increment iterations, check for too many
analysis.MANIter! = analysis.MANIter! + 1#
If analysis.MANIter! > MaxMANIter% Then
msg$ = vbCrLf & "Warning in AnalyzeWeightCalculate- Too many MAN/Interf/APF/Vol iterations on line " & Str$(sample(1).Linenumber&(linerow%)) & ": " & Str$(analysis.MANIter)
Call IOWriteLog(msg$)

msg$ = "Convergence Difference Counts:" & vbCrLf
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).Elsyms$(chan%), a80$)
Next chan%
Call IOWriteLog(msg$)

msg$ = vbNullString
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & MiscAutoFormat$(convergencedifference!(chan%))
Next chan%
Call IOWriteLog(msg$)
Exit Do
End If

Loop

' MAN/Interf iterations completed, load weight data and modified counts
For chan% = 1 To sample(1).LastChan%
analysis.WtsData!(linerow%, chan%) = analysis.WtPercents!(chan%)
sample(1).CorData!(linerow%, chan%) = uncts!(chan%)
Next chan%

' Calculate raw k-ratios using MAN and interference corrected unknown counts
For chan% = 1 To sample(1).LastChan%
ipppp% = IPOS8(chan%, sample(1).Elsyms$(chan%), sample(1).Xrsyms$(chan%), sample())
If Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ipppp% = 0) Then       ' check for duplicate element
If analysis.StdAssignsCounts!(chan%) <> 0# Then RowUnkKRaws!(linerow%, chan%) = sample(1).CorData!(linerow%, chan%) / analysis.StdAssignsCounts!(chan%)
RowUnkKrats!(linerow%, chan%) = analysis.UnkKrats!(chan%)
RowUnkBetas!(linerow%, chan%) = analysis.UnkBetas!(chan%)

RowUnkZAFCors!(linerow%, chan%) = analysis.UnkZAFCors!(4, chan%)
RowUnkZCors!(linerow%, chan%) = analysis.UnkZAFCors!(3, chan%)
RowUnkACors!(linerow%, chan%) = analysis.UnkZAFCors!(1, chan%)
RowUnkFCors!(linerow%, chan%) = analysis.UnkZAFCors!(2, chan%)

RowUnkMACs!(linerow%, chan%) = analysis.UnkMACs!(chan%)
End If

' Calculate and load peak to background ratios. MAN data calculated in routine AnalyzeWeightCorrect and
' loaded into "...BgdData()". Add background back in since we are now using corrected counts for both off-peak and MAN.
temp! = 0#
If sample(1).BgdData!(linerow%, chan%) > MINCPSPERNA! Then    ' to avoid issues with missing off-peak data
temp! = (sample(1).CorData!(linerow%, chan%) + sample(1).BgdData!(linerow%, chan%)) / sample(1).BgdData!(linerow%, chan%)
End If
RowUnkPeakToBgds!(linerow%, chan%) = temp!

' Save drift corrected standard counts for each data linerow
RowStdAssignsCounts!(linerow%, chan%) = analysis.StdAssignsCounts!(chan%)
Next chan%

' Load total weight percents, ZAF iterations and zbar from routine ZAFCalZbar
analysis.WtsData!(linerow%, sample(1).LastChan% + 1) = analysis.TotalPercent!
RowUnkZAFIters!(linerow%) = analysis.ZAFIter!
RowUnkMANIters!(linerow%) = analysis.MANIter!

RowUnkTotalPercents!(linerow%) = analysis.TotalPercent!
RowUnkTotalOxygens!(linerow%) = analysis.totaloxygen!
RowUnkTotalCations!(linerow%) = analysis.TotalCations!
RowUnkTotalAtoms!(linerow%) = analysis.totalatoms!
RowUnkCalculatedOxygens!(linerow%) = analysis.CalculatedOxygen!
RowUnkExcessOxygens!(linerow%) = analysis.ExcessOxygen!
RowUnkZbars!(linerow%) = analysis.Zbar!
RowUnkAtomicWeights!(linerow%) = analysis.AtomicWeight!
RowUnkOxygenFromHalogens!(linerow%) = analysis.OxygenFromHalogens!
RowUnkHalogenCorrectedOxygen!(linerow%) = analysis.HalogenCorrectedOxygen!
RowUnkChargeBalance!(linerow%) = analysis.ChargeBalance!
RowUnkFeChargeBalance!(linerow%) = analysis.FeCharge!

' Check for zero total
If analysis.TotalPercent! < MinTotalValue! Then
msg$ = vbCrLf & "Warning in AnalyzeWeightCalculate- Total is near zero on line " & Str$(sample(1).Linenumber&(linerow%)) & ". If this is a trace analysis, it may be necessary to specify some matrix elements if they are not analyzed, by using the Specified Concentrations button in the Analyze! window."
Call IOWriteLog(msg$)
End If

Exit Sub

' Errors
AnalyzeWeightCalculateError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeWeightCalculate"
ierror = True
Exit Sub

End Sub

Sub AnalyzeWeightCorrect(linerow As Integer, uncts() As Single, oldcts() As Single, alldone As Boolean, firsttime As Boolean, zerror As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' Correct counts for MAN backgrounds, peak interferences, volatile element (TDI) and area peak factor (APF) corrections

ierror = False
On Error GoTo AnalyzeWeightCorrectError

Const INTERF_TOO_NEGATIVE# = -0.01    ' in raw k-ratio units (-0.01 = -1%)

Dim chan As Integer, j As Integer, jmax As Integer, response As Integer
Dim intfchan As Integer, intfstd As Integer, assignstd As Integer, ip As Integer
Dim temp As Single, apf As Single, MANMin As Single
Dim temp1 As Single, temp2 As Single, temp3 As Single, voluncts As Single
Dim tstring As String, tmsg As String
Dim tfactor As Single, tstandard As String

ReDim bgdcts(1 To MAXCHAN%) As Single
ReDim unkintfcts(1 To MAXINTF%, 1 To MAXCHAN%) As Single

ReDim continuum_absorbtion(1 To MAXCHAN%) As Single
ReDim blankcts(1 To MAXCHAN%) As Single

' Reset ZAF error flag
zerror = False

If DebugMode Then
Call IOWriteLog(vbCrLf & "Entering AnalyzeWeightCorrect...")
End If

MANMin! = 0.001 ' minimum cps count rate for MAN/APF/Interference convergence test

' Zero the interference count arrays
For chan% = 1 To MAXCHAN%
For j% = 1 To MAXINTF%
unkintfcts(j%, chan%) = 0#
Next j%
Next chan%

' Zero counts for elements that are disabled
For chan% = 1 To sample(1).LastElm%
If sample(1).DisableQuantFlag%(chan%) = 1 Then
sample(1).CorData!(linerow%, chan%) = 0#
sample(1).BgdData!(linerow%, chan%) = 0#
End If
Next chan%

' Load fresh unknown count data and weight fractions for MAN continuum calculations
For chan% = 1 To sample(1).LastElm%
uncts!(chan%) = sample(1).CorData!(linerow%, chan%)
If CorrectionFlag% = 0 Or CorrectionFlag% = MAXCORRECTION% Then continuum_absorbtion!(chan%) = analysis.UnkZAFCors!(1, chan%)   ' use characteristic absorption for continuum absorption correction (MAN)
'If CorrectionFlag% = 0 Or CorrectionFlag% = MAXCORRECTION% Then continuum_absorbtion!(chan%) = 1# / analysis.UnkZAFCors!(8, chan%)   ' use generated sample intensity f(chi) for continuum absorption correction(MAN)
If CorrectionFlag% > 0 And CorrectionFlag% < 5 Then continuum_absorbtion!(chan%) = analysis.UnkBetas!(chan%)  ' use alpha factor correction for continuum absorption (MAN)
Next chan%

' Calculate maximum number of interference assignments in sample
jmax% = UpdateGetMaxInterfAssign(sample())

' Correct each analyzed channel for volatile elements, MAN, interferences and APF
For chan% = 1 To sample(1).LastElm%
ip% = IPOS8(chan%, sample(1).Elsyms$(chan%), sample(1).Xrsyms$(chan%), sample()) ' find if element is duplicated
If Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0) Then
If sample(1).DisableQuantFlag%(chan%) = 0 Then  ' no disabled quant flag

' First correct counts for volatile element loss if volatile correction flag is set
If UseVolElFlag Then
RowUnkVolElCors!(linerow%, chan%) = 0#
RowUnkVolElDevs!(linerow%, chan%) = 0
If sample(1).VolatileCorrectionUnks%(chan%) <> 0 Then

' Only perform TDI correction if counts are > 0
If uncts!(chan%) > 0# Then

' Calculate volatile correction using count time plus intervals for elapsed time
Call VolatileCalculateCorrection(chan%, linerow%, uncts!(chan%), voluncts!, sample())
If ierror Then Exit Sub

RowUnkVolElCors!(linerow%, chan%) = 100# * (voluncts! - uncts!(chan%)) / uncts!(chan%)
RowUnkVolElDevs!(linerow%, chan%) = sample(1).VolatileFitAvgDev!(chan%)
uncts!(chan%) = voluncts!

' If negative counts, warn user
Else
tmsg$ = "Warning- Negative TDI counts (" & Format$(uncts!(chan%)) & ") for " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " on channel " & Format$(chan%) & ", line " & Format$(sample(1).Linenumber&(linerow%)) & ", sample " & SampleGetString2$(sample()) & " (unable to perform TDI correction)."
Call IOWriteLogRichText(tmsg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
End If
End If
End If

' Do MAN background correction if MAN channel data based on unknown zbar (background type is set to MAN in DataCorrectData if UseMANForOffPeakElementsFlag is true)
If sample(1).BackgroundTypes%(chan%) = 1 Then  ' 0=off-peak, 1=MAN, 2=multipoint
If analysis.Zbar! <= 0# Then GoTo AnalyzeWeightCorrectBadZbar
bgdcts!(chan%) = analysis.MANFitCoefficients!(1, chan%) + analysis.MANFitCoefficients!(2, chan%) * analysis.Zbar! + analysis.MANFitCoefficients!(3, chan%) * analysis.Zbar! ^ 2

' Uncorrect the calculated MAN counts for absorption of the continuum if "MANAbsCorFlags()" is true
If UseMANAbsFlag And sample(1).MANAbsCorFlags(chan%) Then

' De-normalize MAN counts for absorption correction for unknown composition (see MANFitData for standard composition correction)
If continuum_absorbtion!(chan%) > 0# Then
RowUnkMANAbsCors!(linerow%, chan%) = bgdcts!(chan%) ' store uncorrected counts
bgdcts!(chan%) = bgdcts!(chan%) / continuum_absorbtion!(chan%) ' de-normalize
If RowUnkMANAbsCors!(linerow%, chan%) > 0# Then
RowUnkMANAbsCors!(linerow%, chan%) = 100# * (bgdcts!(chan%) - RowUnkMANAbsCors!(linerow%, chan%)) / RowUnkMANAbsCors!(linerow%, chan%)
End If
End If
End If

' Background correction for unknowns and load counts for P/B calculation
uncts!(chan%) = uncts!(chan%) - bgdcts!(chan%)
sample(1).BgdData!(linerow%, chan%) = bgdcts!(chan%)
End If

' Calculate unknown interference counts on interfered channel
If UseInterfFlag Then
assignstd% = IPOS2(NumberofStandards%, sample(1).StdAssigns%(chan%), StandardNumbers%())
If assignstd% = 0 Then GoTo AnalyzeWeightCorrectBadAssignStd

' Is there an assigned interference on this channel?
For j% = 1 To jmax%
If sample(1).StdAssignsIntfStds%(j%, chan%) > 0 Then

' Find the position of the interfering element in the analyzed sample arrays (skip disable quant elements)
intfchan% = IPOSDQ(sample(1).LastElm%, sample(1).StdAssignsIntfElements$(j%, chan%), sample(1).StdAssignsIntfXrays$(j%, chan%), sample(1).Elsyms$(), sample(1).Xrsyms$(), sample(1).DisableQuantFlag%())
If intfchan% = 0 Then GoTo AnalyzeWeightCorrectBadIntfElement
If sample(1).DisableQuantFlag%(intfchan%) = 0 Then  ' no disabled quant flag

' Find the position of the standard used for the interference correction in the standard list
intfstd% = IPOS2(NumberofStandards%, sample(1).StdAssignsIntfStds%(j%, chan%), StandardNumbers%())
If intfstd% = 0 Then GoTo AnalyzeWeightCorrectBadIntfStandard

' Check for valid weight percents and counts on interference standard
If analysis.StdPercents!(intfstd%, intfchan%) <= 0.01 Then GoTo AnalyzeWeightCorrectBadIntfPercents
If ForceNegativeInterferenceIntensitiesToZeroFlag Then
If analysis.StdAssignsIntfCounts!(j%, chan%) < 0# Then analysis.StdAssignsIntfCounts!(j%, chan%) = 0#
Else
If analysis.StdAssignsIntfCounts!(j%, chan%) = 0# Then GoTo AnalyzeWeightCorrectNoIntfCounts
End If

' Check the raw k-ratio of the interference intensity and see if it is too negative
If firsttime And Not ForceNegativeInterferenceIntensitiesToZeroFlag Then
If analysis.StdAssignsIntfCounts!(j%, chan%) / analysis.StdAssignsCounts!(intfchan%) < INTERF_TOO_NEGATIVE# Then
msg$ = "Interference Standard " & Str$(sample(1).StdAssignsIntfStds%(j%, chan%)) & " "
msg$ = msg$ & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " interference counts for sample " & Str$(sample(1).number%) & " "
msg$ = msg$ & sample(1).Name$ & " are very negative (" & analysis.StdAssignsIntfCounts!(j%, chan%) & ")."
msg$ = msg$ & vbCrLf & vbCrLf
msg$ = msg$ & "Although an interference of " & sample(1).StdAssignsIntfElements$(j%, chan%) & " on "
msg$ = msg$ & sample(1).Elsyms$(chan%) & " was assigned, there does not appear to be an actual interference "
msg$ = msg$ & "present or there is an interference on the off-peak positions (e.g., a background fit problem) for the interference standard. Please check for off-peak "
msg$ = msg$ & "interferences, examine the MAN fits or change the interference standard used for the interference correction, or disable the "
msg$ = msg$ & "interference correction, using the Standard Assignments button in the ANALYZE! window."
MsgBox msg$, vbOKOnly + vbExclamation, "AnalyzeWeightCorrect"
ElseIf analysis.StdAssignsIntfCounts!(j%, chan%) < 0# Then
tmsg$ = "Warning- Negative interference counts (" & Format$(analysis.StdAssignsIntfCounts!(j%, chan%)) & ") for " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " on interference standard " & Str$(sample(1).StdAssignsIntfStds%(j%, chan%))
Call IOWriteLogRichText(tmsg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
End If
End If

' Calculate interference on unknown counts (note with Gilfrich, analysis.WtPercents!(intfchan%)is calculated semi-quantitatively using intensity ratio only)
If DisableFullQuantInterferenceCorrectionFlag = 0 Then
unkintfcts!(j%, chan%) = analysis.StdAssignsIntfCounts!(j%, chan%) * analysis.WtPercents!(intfchan%) / analysis.StdPercents!(intfstd%, intfchan%)
Else
temp1! = analysis.StdAssignsIntfCounts!(j%, chan%) * sample(1).CorData!(linerow%, intfchan%) * analysis.StdAssignsPercents(intfchan%)
temp2! = analysis.StdPercents!(intfstd%, intfchan%) * analysis.StdAssignsCounts(intfchan%)
If temp2! <> 0# Then unkintfcts!(j%, chan%) = temp1! / temp2!            ' use approximation method of Gilfrich for educational purposes
End If

' Check for negative correction factors
If Not IgnoreZAFandAlphaFactorWarnings Then
If (CorrectionFlag% = 0 And analysis.UnkZAFCors!(4, chan%) < 0#) Or (CorrectionFlag% > 0 And CorrectionFlag% < 5 And analysis.UnkBetas!(chan%) < 0#) Then
msg$ = "Invalid unknown correction factors on line " & Str$(sample(1).Linenumber&(linerow%)) & " " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " channel"
response% = MsgBox(msg$, vbAbortRetryIgnore + vbExclamation, "AnalyzeWeightCorrect")

If response% = vbAbort Then
ierror = True
Exit Sub
End If

If response% = vbIgnore Then
IgnoreZAFandAlphaFactorWarnings = True
zerror = True
End If

End If
End If

' Calculate matrix correction for unknown interference counts (use Z, A and F since F for K-alpha lines for beta-alpha fluorescence will be correct)
If CorrectionFlag% = 0 Then
If analysis.UnkZAFCors!(4, chan%) <> 0# Then temp! = analysis.StdZAFCors!(4, intfstd%, chan%) / analysis.UnkZAFCors!(4, chan%)
ElseIf CorrectionFlag% > 0 And CorrectionFlag% < 5 Then
If analysis.UnkBetas!(chan%) <> 0# Then temp! = analysis.StdBetas!(intfstd%, chan%) / analysis.UnkBetas!(chan%)
ElseIf CorrectionFlag% = MAXCORRECTION% Then
End If

' Correct interfering counts on unknown for matrix correction (do only for first order interferences for now)
If DisableFullQuantInterferenceCorrectionFlag = 0 And DisableMatrixCorrectionInterferenceCorrectionFlag = 0 And sample(1).StdAssignsIntfOrders%(j%, chan%) = 1 Then
unkintfcts!(j%, chan%) = unkintfcts!(j%, chan%) * temp!
End If

' Display interference calculations
If VerboseMode Then
If DisableFullQuantInterferenceCorrectionFlag = 0 Then
Call AnalyzePrintEquationParameters(Int(0), linerow%, j%, chan%, intfchan%, intfstd%, unkintfcts!(j%, chan%), analysis, sample())
Else
Call AnalyzePrintEquationParameters(Int(1), linerow%, j%, chan%, intfchan%, intfstd%, unkintfcts!(j%, chan%), analysis, sample())
End If
End If

End If
End If
Next j%

' Correct unknown counts for interference
RowUnkIntfCors!(linerow%, chan%) = uncts!(chan%)
For j% = 1 To jmax%
If unkintfcts!(j%, chan%) > 0# Then uncts!(chan%) = uncts!(chan%) - unkintfcts!(j%, chan%)
Next j%
If RowUnkIntfCors!(linerow%, chan%) <> 0# Then  ' calculate interference correction magnitude
RowUnkIntfCors!(linerow%, chan%) = 100# * (uncts!(chan%) - RowUnkIntfCors!(linerow%, chan%)) / RowUnkIntfCors!(linerow%, chan%)
End If
End If

' Correct counts on unknown for peak shape changes using APF's (Area Peak
' Factors). Sum weight fraction of APF from each absorber. The sum of APFs
' is calculated from the element weight fractions not including the affected
' element.
If Not sample(1).IntegratedIntensitiesUseIntegratedFlags(chan%) Then
If UseAPFFlag And UseAPFOption% = 0 Then
RowUnkAPFCors!(linerow%, chan%) = 0#
temp! = 0#
For j% = 1 To sample(1).LastChan%
If chan% <> j% Then temp! = temp! + analysis.WtPercents!(j%)    ' calculate partial sum (not including affected element)
Next j%

If temp! <> 0# Then                 ' was "If temp! > 0# Then" but caused problem when temp was slightly negative
For j% = 1 To sample(1).LastChan%
If chan% <> j% Then
Call EmpLoadMACAPF(Int(2), sample(1).AtomicNums%(chan%), sample(1).XrayNums%(chan%), sample(1).AtomicNums%(j%), apf!, tstring$, tfactor!, tstandard$)
If ierror Then Exit Sub
RowUnkAPFCors!(linerow%, chan%) = RowUnkAPFCors!(linerow%, chan%) + apf! * analysis.WtPercents!(j%) / temp!     ' sum APFs based on relative abundance
End If
Next j%

' Perform APF normalization based on partial sum (added 02/22/2009 to deal with Si Ka peak shift Si -> SiO2)
If analysis.WtPercents!(chan%) <> 0# Then
temp3! = 1# / (temp! / analysis.WtPercents!(chan%))         ' calculate scaling factor based on partial sum
If Abs(temp3!) >= 1# Then RowUnkAPFCors!(linerow%, chan%) = 1# + (RowUnkAPFCors!(linerow%, chan%) - 1) / temp3!
End If

uncts!(chan%) = uncts!(chan%) * RowUnkAPFCors!(linerow%, chan%)     ' perform compound APF correction to intensities
End If
End If

' Correct using "specified" APF factor for this emitter (and not assigned as the standard)
If UseAPFFlag And UseAPFOption% = 1 Then
If sample(1).Type% = 2 Or (sample(1).Type% = 1 And sample(1).number% <> sample(1).StdAssigns%(chan%)) Then
uncts!(chan%) = uncts!(chan%) * sample(1).SpecifiedAreaPeakFactors!(chan%)
End If
End If
End If

' Do blank correction for trace elements
If UseBlankCorFlag Then
If Not MiscAllZero(sample(1).LastElm%, sample(1).BlankCorrectionUnks%()) Then
If sample(1).BlankCorrectionUnks%(chan%) > 0 Then
temp! = (AnalBlankCorrectionPercents!(chan%) - sample(1).BlankCorrectionLevels!(chan%))              ' calculate wt% blank correction
temp! = analysis.StdAssignsCounts!(chan%) * temp! / analysis.StdAssignsPercents!(chan%)              ' convert to unknown intensity (divide % by %)

If CorrectionFlag% = 0 Then
blankcts!(chan%) = temp! * analysis.StdAssignsZAFCors!(4, chan%) / analysis.UnkZAFCors!(4, chan%)    ' correct for matrix effect
ElseIf CorrectionFlag% > 0 And CorrectionFlag% < 5 Then
blankcts!(chan%) = temp! * analysis.StdAssignsBetas!(chan%) / analysis.UnkBetas!(chan%)              ' correct for matrix effect
ElseIf CorrectionFlag% = MAXCORRECTION% Then
End If

uncts!(chan%) = uncts!(chan%) - blankcts!(chan%)
End If
End If
End If

End If  ' disable quant endif
End If  ' use aggregate intensity endif
Next chan%

' Check for change in unknown counts due to all compositionally dependent corrections
alldone = True
For chan% = 1 To sample(1).LastElm%
convergencedifference!(chan%) = Abs(uncts!(chan%) - oldcts!(chan%))
If convergencedifference!(chan%) > MANMin! And convergencedifference!(chan%) > Abs(uncts!(chan%) / 1000#) Then
alldone = False
End If
oldcts!(chan%) = uncts!(chan%)
Next chan%

' Debug statements
If Not DebugMode Then Exit Sub

msg$ = "Elements:" & vbCrLf
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).Elsyms$(chan%), a80$)
Next chan%
Call IOWriteLog(msg$)

msg$ = vbCrLf & "Uncorrected Unknown Counts:" & vbCrLf
For chan% = 1 To sample(1).LastElm%
If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(sample(1).CorData!(linerow%, chan%), f81$), a80$)
Else
msg$ = msg$ & Format$(Format$(sample(1).CorData!(linerow%, chan%), f82$), a80$)
End If
Next chan%
Call IOWriteLog(msg$)

msg$ = "MAN Background Counts on Unknown (based on unknown Z-bar: " & Str$(analysis.Zbar!) & ")" & vbCrLf
For chan% = 1 To sample(1).LastElm%
If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(bgdcts!(chan%), f81$), a80$)
Else
msg$ = msg$ & Format$(Format$(bgdcts!(chan%), f82$), a80$)
End If
Next chan%
Call IOWriteLog(msg$)

Call IOWriteLog(vbNullString)
For j% = 1 To jmax%
msg$ = "Interfering " & Str$(j%) & " Counts on Unknown:" & vbCrLf
For chan% = 1 To sample(1).LastElm%
If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(unkintfcts!(j%, chan%), f81$), a80$)
Else
msg$ = msg$ & Format$(Format$(unkintfcts!(j%, chan%), f82$), a80$)
End If
Next chan%
Call IOWriteLog(msg$)
Next j%
        
msg$ = vbCrLf & "Continuum Absorption Correction Factors on Unknown:" & vbCrLf
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(continuum_absorbtion!(chan%), f84$), a80$)
Next chan%
Call IOWriteLog(msg$)
        
msg$ = vbCrLf & "Absorption Correction Factors on Unknown:" & vbCrLf
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.UnkZAFCors!(1, chan%), f84$), a80$)
Next chan%
Call IOWriteLog(msg$)
        
msg$ = "Fluorescence Correction Factors on Unknown:" & vbCrLf
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.UnkZAFCors!(2, chan%), f84$), a80$)
Next chan%
Call IOWriteLog(msg$)
        
msg$ = "Atomic Number Correction Factors on Unknown:" & vbCrLf
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.UnkZAFCors!(3, chan%), f84$), a80$)
Next chan%
Call IOWriteLog(msg$)

msg$ = "ZAF Correction Factors on Unknown:" & vbCrLf
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.UnkZAFCors!(4, chan%), f84$), a80$)
Next chan%
Call IOWriteLog(msg$)

msg$ = vbCrLf & "Disabled Quant Flag (zeroed intensity if set):" & vbCrLf
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & MiscAutoFormatI$(sample(1).DisableQuantFlag%(chan%))
Next chan%
Call IOWriteLog(msg$)

msg$ = "Corrected Unknown Counts:" & vbCrLf
For chan% = 1 To sample(1).LastElm%
If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(uncts!(chan%), f81$), a80$)
Else
msg$ = msg$ & Format$(Format$(uncts!(chan%), f82$), a80$)
End If
Next chan%
Call IOWriteLog(msg$)

msg$ = vbCrLf & "Convergence Difference Counts:" & vbCrLf
For chan% = 1 To sample(1).LastElm%
If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(convergencedifference!(chan%), f81$), a80$)
Else
msg$ = msg$ & Format$(Format$(convergencedifference!(chan%), f82$), a80$)
End If
Next chan%
Call IOWriteLog(msg$)

msg$ = vbCrLf & "Sum Area-Peak-Factors:" & vbCrLf
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(RowUnkAPFCors!(linerow%, chan%), f84$), a80$)
Next chan%
Call IOWriteLog(msg$)

msg$ = "Time Dependent Intensity (TDI) Element Correction Percents:" & vbCrLf
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(RowUnkVolElCors!(linerow%, chan%), f82$), a80$)
Next chan%
Call IOWriteLog(msg$)

msg$ = "Time Dependent Intensity (TDI) Element Correction Percent Relative Deviations:" & vbCrLf
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(RowUnkVolElDevs!(linerow%, chan%), f81$), a80$)
Next chan%
Call IOWriteLog(msg$)

msg$ = "MAN Absorption Correction Percents:" & vbCrLf
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(RowUnkMANAbsCors!(linerow%, chan%), f82$), a80$)
Next chan%
Call IOWriteLog(msg$)

Exit Sub

' Errors
AnalyzeWeightCorrectError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeWeightCorrect"
ierror = True
Exit Sub

AnalyzeWeightCorrectBadZbar:
msg$ = "MAN Z-bar (" & Format$(analysis.Zbar!) & ") is bad on line " & Str$(sample(1).Linenumber&(linerow%)) & ", probably from a very low analytical total, e.g., hitting epoxy."
If Not CalcImageQuantFlag Then
MiscMsgBoxTim FormMSGBOXTIME, "AnalyzeWeightCorrect", msg$, 20#
Call IOWriteLog(msg$)
Else
Call IOWriteLog(msg$)
End If
zerror = True
Exit Sub

AnalyzeWeightCorrectBadIntfPercents:
msg$ = "Insufficient weight percent of " & sample(1).StdAssignsIntfElements$(j%, chan%) & " in "
msg$ = msg$ & "interference standard " & Str$(sample(1).StdAssignsIntfStds%(j%, chan%)) & " for "
msg$ = msg$ & " for sample " & Str$(sample(1).number%) & " " & sample(1).Name$ & "."
msg$ = msg$ & vbCrLf & vbCrLf
msg$ = msg$ & "Although an interference of " & sample(1).StdAssignsIntfElements$(j%, chan%) & " on "
msg$ = msg$ & sample(1).Elsyms$(chan%) & " was assigned, there is an insufficient concentration of the interfering element "
msg$ = msg$ & "present. Please change the interference standard used for the interference correction, or "
msg$ = msg$ & "disable the interference correction, using the Standard Assignments button in the ANALYZE! window."
MsgBox msg$, vbOKOnly + vbExclamation, "AnalyzeWeightCorrect"
ierror = True
Exit Sub

AnalyzeWeightCorrectNoIntfCounts:
msg$ = "No interference counts on interference standard " & Str$(sample(1).StdAssignsIntfStds%(j%, chan%)) & " for "
msg$ = msg$ & sample(1).Elsyms$(chan%) & " for sample " & Str$(sample(1).number%) & " " & sample(1).Name$ & "."
MsgBox msg$, vbOKOnly + vbExclamation, "AnalyzeWeightCorrect"
ierror = True
Exit Sub

AnalyzeWeightCorrectBadAssignStd:
msg$ = "Invalid assigned standard on " & sample(1).Elsyms$(chan%) & " channel"
MsgBox msg$, vbOKOnly + vbExclamation, "AnalyzeWeightCorrect"
ierror = True
Exit Sub

AnalyzeWeightCorrectBadIntfElement:
msg$ = "Invalid interfering element on " & sample(1).Elsyms$(chan%) & " channel"
MsgBox msg$, vbOKOnly + vbExclamation, "AnalyzeWeightCorrect"
ierror = True
Exit Sub

AnalyzeWeightCorrectBadIntfStandard:
msg$ = "Standard number " & Str$(sample(1).StdAssignsIntfStds%(j%, chan%)) & " is an invalid interference standard on " & sample(1).Elsyms$(chan%) & " channel"
MsgBox msg$, vbOKOnly + vbExclamation, "AnalyzeWeightCorrect"
ierror = True
Exit Sub

End Sub

Sub AnalyzeWeightCalculateCurve(curvecoeffs() As Single, curvedevs() As Single, linerow As Integer, analysis As TypeAnalysis, sample() As TypeSample, stdsample() As TypeSample)
' Called by routine AnalyzeSampleCurve to calculate the weight percents for the specified sample linerow

ierror = False
On Error GoTo AnalyzeWeightCalculateCurveError

Dim j As Integer, ip As Integer, ipp As Integer, chan As Integer
Dim temp As Single, excess As Single, stoichoxygen As Single, sum As Single

ReDim uncts(1 To MAXCHAN%) As Single

' Write debug info
If DebugMode Then
Call IOWriteLog(vbCrLf & "Sample line number: " & Str$(sample(1).Linenumber&(linerow%)))

msg$ = "Elements:" & vbCrLf
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).Elsyms$(chan%), a80$)
Next chan%
Call IOWriteLog(msg$)

For j% = 1 To MAXCOEFF%
msg$ = "Calibration curve coefficients " & Str$(j%) & ":" & vbCrLf
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(curvecoeffs!(j%, chan%), f84$), a80$)
Next chan%
Call IOWriteLog(msg$)
Next j%
End If

' Load specified weight percents for this sample
For chan% = 1 To sample(1).LastChan%
uncts!(chan%) = 0#
analysis.WtPercents!(chan%) = 0#

' Load specified elements from standard sample if element is a specified
' channel. If sample is a standard, the standard sample was loaded from the
' standard database. If sample is an unknown, the standard sample is simply
' a copy of the sample.
If chan% > sample(1).LastElm% Then
ip% = IPOS11(sample(1).Elsyms$(chan%), stdsample())

' Load specified weight percent if not specified oxygen channel
If ip% > 0 Then
If chan% <> sample(1).OxygenChannel% Then
analysis.WtPercents!(chan%) = stdsample(1).ElmPercents!(ip%)

' If element is oxygen, make sure that correct excess oxygen is specified based on the sample cations (standards only)
Else

If sample(1).Type% = 1 Then
analysis.WtPercents!(chan%) = ConvertTotalToExcessOxygen!(Int(1), sample(), stdsample())

' For unknowns, use specified oxygen weight percent
Else
analysis.WtPercents!(chan%) = stdsample(1).ElmPercents!(ip%)
End If
excess! = analysis.WtPercents!(chan%)   ' store excess oxygen for iteration
End If
End If
End If

Next chan%

' Zero counts for elements that are disabled
For chan% = 1 To sample(1).LastElm%
If sample(1).DisableQuantFlag%(chan%) = 1 Then
sample(1).CorData!(linerow%, chan%) = 0#
sample(1).BgdData!(linerow%, chan%) = 0#
End If
Next chan%

' Load unknown counts
For chan% = 1 To sample(1).LastElm%
uncts!(chan%) = sample(1).CorData!(linerow%, chan%)
Next chan%

If DebugMode Then
msg$ = "Elements:" & vbCrLf
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(sample(1).Elsyms$(chan%), a80$)
Next chan%
Call IOWriteLog(msg$)

msg$ = "Disabled Quant Flag (zeroed intensity if set):" & vbCrLf
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & MiscAutoFormatI$(sample(1).DisableQuantFlag%(chan%))
Next chan%
Call IOWriteLog(msg$)

msg$ = "Unknown counts:" & vbCrLf
For chan% = 1 To sample(1).LastElm%
If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(uncts!(chan%), f81$), a80$)
Else
msg$ = msg$ & Format$(Format$(uncts!(chan%), f82$), a80$)
End If
Next chan%
Call IOWriteLog(msg$)

msg$ = "Specified weights:" & vbCrLf
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(analysis.WtPercents!(chan%), f83$), a80$)
Next chan%
Call IOWriteLog(msg$)
End If

' Calculate weight percents using calibration curve !!!!!!!!!!!!!!!!!!!!!!!
For chan% = 1 To sample(1).LastElm%
analysis.WtPercents!(chan%) = curvecoeffs!(1, chan%) + curvecoeffs!(2, chan%) * uncts!(chan%) + curvecoeffs!(3, chan%) * uncts!(chan%) ^ 2
Next chan%

If DebugMode Then
msg$ = "Calculated weights:" & vbCrLf
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(analysis.WtPercents!(chan%), f83$), a80$)
Next chan%
Call IOWriteLog(msg$)
End If

' If calculating oxygen by stoichiometry, do it here
stoichoxygen! = 0#
If sample(1).OxideOrElemental% = 1 Then
If sample(1).OxygenChannel% > 0 Then analysis.WtPercents!(sample(1).OxygenChannel%) = excess!
For chan% = 1 To sample(1).LastChan%
If chan% <> sample(1).OxygenChannel% Then
temp! = ConvertElmToOxd(analysis.WtPercents!(chan%), sample(1).Elsyms$(chan%), sample(1).numcat%(chan%), sample(1).numoxd%(chan%))
stoichoxygen! = stoichoxygen! + (temp! - analysis.WtPercents!(chan%))
End If
Next chan%

' Calculate equivalent oxygen from halogens and subtract from calculated oxygen if flagged
If UseOxygenFromHalogensCorrectionFlag Then stoichoxygen! = stoichoxygen! - ConvertHalogensToOxygen(sample(1).LastChan%, sample(1).Elsyms$(), sample(1).DisableQuantFlag%(), analysis.WtPercents!())
End If

' Element by stoichiometry to stoichiometric oxygen
If sample(1).StoichiometryElementFlag% Then
ip% = IPOS1(sample(1).LastChan%, sample(1).StoichiometryElement$, sample(1).Elsyms$())
If ip% > 0 And sample(1).OxideOrElemental% = 1 Then
temp! = (stoichoxygen! / AllAtomicWts!(ATOMIC_NUM_OXYGEN%)) * sample(1).StoichiometryRatio! * sample(1).AtomicWts!(ip%)
analysis.WtPercents!(ip%) = temp!

' Add stoichiometric oxygen from element by stoichiometry to stoichiometric oxygen
If sample(1).OxideOrElemental% = 1 Then
temp! = ConvertElmToOxd(analysis.WtPercents!(ip%), sample(1).Elsyms$(ip%), sample(1).numcat%(ip%), sample(1).numoxd%(ip%))
stoichoxygen! = stoichoxygen! + (temp! - analysis.WtPercents!(ip%))
End If
End If
End If

' Element relative to another element
If sample(1).RelativeElementFlag% Then
ip% = IPOS1(sample(1).LastChan%, sample(1).RelativeElement$, sample(1).Elsyms$())
ipp% = IPOS1(sample(1).LastChan%, sample(1).RelativeToElement$, sample(1).Elsyms$())
If ip% > 0 And ipp% > 0 Then
analysis.WtPercents!(ip%) = analysis.WtPercents!(ipp%) / sample(1).AtomicWts!(ipp%) * sample(1).RelativeRatio! * sample(1).AtomicWts!(ip%)

' Add stoichiometric oxygen from element relative to another element
If sample(1).OxideOrElemental% = 1 Then
temp! = ConvertElmToOxd(analysis.WtPercents!(ip%), sample(1).Elsyms$(ip%), sample(1).numcat%(ip%), sample(1).numoxd%(ip%))
stoichoxygen! = stoichoxygen! + (temp! - analysis.WtPercents!(ip%))
End If
End If
End If

' Do element by difference
If sample(1).DifferenceElementFlag% Then
ip% = IPOS1(sample(1).LastChan%, sample(1).DifferenceElement$, sample(1).Elsyms$())
If ip% > 0 Then
sum! = 0#
For chan% = 1 To sample(1).LastChan%
sum! = sum! + analysis.WtPercents!(chan%)
Next chan%
sum! = sum! + stoichoxygen!
analysis.WtPercents!(ip%) = 100# - sum!

' Add stoichiometric oxygen for element by difference
If sample(1).OxideOrElemental% = 1 Then
temp! = ConvertOxdToElm(analysis.WtPercents!(ip%), sample(1).Elsyms$(ip%), sample(1).numcat%(ip%), sample(1).numoxd%(ip%))
analysis.WtPercents!(ip%) = temp!   ' load elemental weight percent
temp! = ConvertElmToOxd(analysis.WtPercents!(ip%), sample(1).Elsyms$(ip%), sample(1).numcat%(ip%), sample(1).numoxd%(ip%))
stoichoxygen! = stoichoxygen! + (temp! - analysis.WtPercents!(ip%))
End If
End If
End If

' Do formula by difference
If sample(1).DifferenceFormulaFlag% Then




End If

' Calculate z-bar, etc.
Call ZAFCalZBar(stoichoxygen!, analysis, sample())
If ierror Then Exit Sub

If DebugMode Then
msg$ = "ZAFCal weights:" & vbCrLf
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(analysis.WtPercents!(chan%), f83$), a80$)
Next chan%
Call IOWriteLog(msg$)
End If

' Load calculated data
For chan% = 1 To sample(1).LastChan%
analysis.WtsData!(linerow%, chan%) = analysis.WtPercents!(chan%)

' Load fit deviations from calibration curve calculation
RowCurve1Coeffs!(linerow%, chan%) = curvecoeffs!(1, chan%)
RowCurve2Coeffs!(linerow%, chan%) = curvecoeffs!(2, chan%)
RowCurve3Coeffs!(linerow%, chan%) = curvecoeffs!(3, chan%)
RowCurveFits!(linerow%, chan%) = curvedevs!(chan%)

' Calculate and load peak to background ratios
temp! = 0#
If sample(1).BgdData!(linerow%, chan%) <> 0# Then
temp! = (sample(1).CorData!(linerow%, chan%) + sample(1).BgdData!(linerow%, chan%)) / sample(1).BgdData!(linerow%, chan%)
End If
RowUnkPeakToBgds!(linerow%, chan%) = temp!

' Zero these
RowStdAssignsCounts!(linerow%, chan%) = 0#
RowUnkKRaws!(linerow%, chan%) = 1#
RowUnkKrats!(linerow%, chan%) = 1#
RowUnkZAFCors!(linerow%, chan%) = 1#
RowUnkBetas!(linerow%, chan%) = 1#

Next chan%

' Load total weight percents, ZAF iterations and zbar from routine ZAFCalZbar
analysis.WtsData!(linerow%, sample(1).LastChan% + 1) = analysis.TotalPercent!
RowUnkZAFIters!(linerow%) = 0#
RowUnkMANIters!(linerow%) = 0#

RowUnkTotalPercents!(linerow%) = analysis.TotalPercent!
RowUnkTotalOxygens!(linerow%) = analysis.totaloxygen!
RowUnkTotalCations!(linerow%) = analysis.totaloxygen!
RowUnkCalculatedOxygens!(linerow%) = analysis.CalculatedOxygen!
RowUnkExcessOxygens!(linerow%) = analysis.ExcessOxygen!
RowUnkZbars!(linerow%) = analysis.Zbar!
RowUnkAtomicWeights!(linerow%) = analysis.AtomicWeight!
RowUnkOxygenFromHalogens!(linerow%) = analysis.OxygenFromHalogens!
RowUnkHalogenCorrectedOxygen!(linerow%) = analysis.HalogenCorrectedOxygen!
RowUnkChargeBalance!(linerow%) = analysis.ChargeBalance!
RowUnkFeChargeBalance!(linerow%) = analysis.FeCharge!

Exit Sub

' Errors
AnalyzeWeightCalculateCurveError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeWeightCalculateCurve"
ierror = True
Exit Sub

End Sub

Sub AnalyzeGetStdCts(datarow As Integer, stdcts() As Single)
' Return the array of standard counts for this line

ierror = False
On Error GoTo AnalyzeGetStdCtsError

Dim i As Integer

' Load passed array with standard data
For i% = 1 To MAXCHAN%
stdcts!(i%) = RowStdAssignsCounts!(datarow%, i%)
Next i%
    
Exit Sub

' Errors
AnalyzeGetStdCtsError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeGetStdCts"
ierror = True
Exit Sub

End Sub

Sub AnalyzeInitAnalysis(analysis As TypeAnalysis)
' Initialize the analysis arrays

ierror = False
On Error GoTo AnalyzeInitAnalysisError

Dim i As Integer, j As Integer

' Reset global analysis error flag
icancelanal = False

' Initialize counters
analysis.ZAFIter! = 0#
analysis.MANIter! = 0#

' Initialize sample row arrays
For j% = 1 To MAXROW%
RowUnkZAFIters!(j%) = 0#
RowUnkMANIters!(j%) = 0#
RowUnkTotalPercents!(j%) = 0#
RowUnkTotalOxygens!(j%) = 0#
RowUnkTotalCations!(j%) = 0#
RowUnkCalculatedOxygens!(j%) = 0#
RowUnkExcessOxygens!(j%) = 0#
RowUnkZbars!(j%) = 0#
RowUnkAtomicWeights!(j%) = 0#
RowUnkOxygenFromHalogens!(j%) = 0#
RowUnkHalogenCorrectedOxygen!(j%) = 0#
RowUnkChargeBalance!(j%) = 0#
RowUnkFeChargeBalance!(j%) = 0#

For i% = 1 To MAXCHAN%
RowUnkKrats!(j%, i%) = 0#
RowUnkZAFCors!(j%, i%) = 0#
RowUnkKRaws!(j%, i%) = 0#
RowUnkBetas!(j%, i%) = 0#

RowUnkPeakToBgds!(j%, i%) = 0#
RowUnkIntfCors!(j%, i%) = 0#
RowUnkMANAbsCors!(j%, i%) = 0#
RowUnkAPFCors!(j%, i%) = 0#
RowUnkVolElCors!(j%, i%) = 0#
RowUnkVolElDevs!(j%, i%) = 0#
Next i%
Next j%

' Initialize "analysis.WtsData" and "analysis.CalData" arrays
For i% = 1 To MAXCHAN1%
For j% = 1 To MAXROW%
analysis.WtsData!(j%, i%) = 0#
analysis.CalData!(j%, i%) = 0#
Next j%
Next i%
    
analysis.TotalPercent! = 0#
analysis.AtomicWeight! = 0#
analysis.Zbar! = 0#

analysis.CalculatedOxygen! = 0#
analysis.totaloxygen! = 0#
analysis.ExcessOxygen! = 0#

Exit Sub

' Errors
AnalyzeInitAnalysisError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeInitAnalysis"
ierror = True
Exit Sub

End Sub

Sub AnalyzeCalculateAverages(analysis As TypeAnalysis, sample() As TypeSample)
' Calculate row averages

ierror = False
On Error GoTo AnalyzeCalculateAveragesError

Dim average As TypeAverage

Call MathAverage(average, RowUnkTotalPercents!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
analysis.TotalPercent! = average.averags!(1)

Call MathAverage(average, RowUnkTotalOxygens!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
analysis.totaloxygen! = average.averags!(1)

Call MathAverage(average, RowUnkTotalCations!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
analysis.TotalCations! = average.averags!(1)

Call MathAverage(average, RowUnkTotalAtoms!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
analysis.totalatoms! = average.averags!(1)

Call MathAverage(average, RowUnkCalculatedOxygens!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
analysis.CalculatedOxygen! = average.averags!(1)

Call MathAverage(average, RowUnkExcessOxygens!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
analysis.ExcessOxygen! = average.averags!(1)

Call MathAverage(average, RowUnkZAFIters!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
analysis.ZAFIter! = average.averags!(1)

Call MathAverage(average, RowUnkMANIters!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
analysis.MANIter! = average.averags!(1)

Call MathAverage(average, RowUnkZbars!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
analysis.Zbar! = average.averags!(1)

Call MathAverage(average, RowUnkAtomicWeights!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
analysis.AtomicWeight! = average.averags!(1)

Call MathAverage(average, RowUnkOxygenFromHalogens!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
analysis.OxygenFromHalogens! = average.averags!(1)

Call MathAverage(average, RowUnkHalogenCorrectedOxygen!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
analysis.HalogenCorrectedOxygen! = average.averags!(1)

Call MathAverage(average, RowUnkChargeBalance!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
analysis.ChargeBalance! = average.averags!(1)

Call MathAverage(average, RowUnkFeChargeBalance!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
analysis.FeCharge! = average.averags!(1)

Exit Sub

' Errors
AnalyzeCalculateAveragesError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeCalculateAverages"
ierror = True
Exit Sub

End Sub

Sub AnalyzeTypeSpecified(mode As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' Type the specified elements

ierror = False
On Error GoTo AnalyzeTypeSpecifiedError

Dim linerow As Integer
Dim i As Integer, ii As Integer, jj As Integer, n As Integer

Dim average As TypeAverage

' Type out specified elements
n = 0
Do Until False
n% = n% + 1
Call TypeGetRange(Int(3), n%, ii%, jj%, sample())
If ierror Then Exit Sub
If ii% > sample(1).LastChan% Then Exit Do

' Specified element or oxide symbols
msg$ = vbCrLf & "SPEC: "
For i% = ii% To jj%
If mode% = 2 Or mode% = 6 Then
msg$ = msg$ & Format$(sample(1).Oxsyup$(i%), a80$)
Else
msg$ = msg$ & Format$(sample(1).Elsyup$(i%), a80$)
End If
Next i%
Call IOWriteLog(msg$)

' Type specified element types if indicated
If UseDetailedFlag Then
msg$ = "TYPE: "
For i% = ii% To jj%
If sample(1).DifferenceElementFlag% And sample(1).DifferenceElement$ = sample(1).Elsyms$(i%) Then
msg$ = msg$ & Format$("DIFF", a80$)
ElseIf sample(1).DifferenceFormulaFlag% And ConvertIsDifferenceFormulaElement(sample(1).DifferenceFormula$, sample(1).Elsyms$(i%)) Then
msg$ = msg$ & Format$("FORM", a80$)
ElseIf sample(1).StoichiometryElementFlag% And sample(1).StoichiometryElement$ = sample(1).Elsyms$(i%) Then
msg$ = msg$ & Format$("STOI", a80$)
ElseIf sample(1).RelativeElementFlag% And sample(1).RelativeElement$ = sample(1).Elsyms$(i%) Then
msg$ = msg$ & Format$("RELA", a80$)
ElseIf sample(1).OxideOrElemental% = 1 And sample(1).OxygenChannel% = i% Then
msg$ = msg$ & Format$("CALC", a80$)
Else
msg$ = msg$ & Format$("SPEC", a80$)
End If
Next i%
Call IOWriteLog(msg$)
End If

' Type out specified weight percents and sum for each data line (if debug)
If DebugMode Then
For linerow% = 1 To sample(1).Datarows
If sample(1).LineStatus(linerow%) Then
msg$ = Format$(Format$(sample(1).Linenumber&(linerow%), i50$), a60$)
For i% = ii% To jj%
If Not UseAutomaticFormatForResultsFlag Then
If mode% = 1 Then
msg$ = msg$ & Format$(Format$(analysis.WtsData!(linerow%, i%), f83$), a80$)
Else
msg$ = msg$ & Format$(Format$(analysis.CalData!(linerow%, i%), f83$), a80$)
End If
Else
If mode% = 1 Then
msg$ = msg$ & AnalyzeFormatAnalysisResult$(Int(0), Int(0), analysis.WtsData!(linerow%, i%), analysis, sample())
Else
msg$ = msg$ & AnalyzeFormatAnalysisResult$(Int(0), Int(0), analysis.CalData!(linerow%, i%), analysis, sample())
End If
End If
Next i%
Call IOWriteLog(msg$)
End If
Next linerow%
End If

' Type average specified weight percents and std deviation
If mode% = 1 Then
Call MathArrayAverage(average, analysis.WtsData!(), sample(1).Datarows%, sample(1).LastChan% + 1, sample())
If ierror Then Exit Sub
Else
Call MathArrayAverage(average, analysis.CalData!(), sample(1).Datarows%, sample(1).LastChan% + 1, sample())
If ierror Then Exit Sub
End If

If UseDetailedFlag Then
msg$ = vbCrLf & "AVER: "
Else
msg$ = "AVER: "
End If
For i% = ii% To jj%
If Not UseAutomaticFormatForResultsFlag Or i% > sample(1).LastElm% Then
msg$ = msg$ & Format$(Format$(average.averags!(i%), f83$), a80$)
Else
msg$ = msg$ & AnalyzeFormatAnalysisResult$(Int(0), Int(0), average.averags!(i%), analysis, sample())
End If
Next i%
Call IOWriteLog(msg$)
        
msg$ = "SDEV: "
For i% = ii% To jj%
If Not UseAutomaticFormatForResultsFlag Or i% > sample(1).LastElm% Then
msg$ = msg$ & Format$(Format$(average.Stddevs!(i%), f83$), a80$)
Else
msg$ = msg$ & AnalyzeFormatAnalysisResult$(Int(0), Int(0), average.Stddevs!(i%), analysis, sample())
End If
Next i%
Call IOWriteLog(msg$)
Loop

Exit Sub

' Errors
AnalyzeTypeSpecifiedError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeTypeSpecified"
ierror = True
Exit Sub

End Sub

Sub AnalyzeGetCorrectionFactors(tZAFCors() As Single, tBetas() As Single, tStdCts() As Single)
' Return the line by line correction factors

ierror = False
On Error GoTo AnalyzeGetCorrectionFactorsError

Dim i As Integer, j As Integer

For i% = 1 To MAXROW%
For j% = 1 To MAXCHAN%
tZAFCors!(i%, j%) = RowUnkZAFCors!(i%, j%)
tBetas!(i%, j%) = RowUnkBetas!(i%, j%)
tStdCts!(i%, j%) = RowStdAssignsCounts!(i%, j%)
Next j%
Next i%

Exit Sub

' Errors
AnalyzeGetCorrectionFactorsError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeGetCorrectionFactorsError"
ierror = True
Exit Sub

End Sub

Sub AnalyzeTypeResults(mode As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' mode = 0 type out sample description and warnings only
' mode = 1 type out "analysis.WtsData!" array (elemental weight percents)
' mode = 2 type out "analysis.CalData!" array (oxide weight percents)
' mode = 3 type out "analysis.CalData!" array (atomic percents)
' mode = 4 type out "analysis.CalData!" array (formula atoms)
' mode = 5 type out "analysis.CalData!" array (normalized elemental)
' mode = 6 type out "analysis.CalData!" array (normalized oxide)

ierror = False
On Error GoTo AnalyzeTypeResultsError

Dim linerow As Integer
Dim i As Integer, j As Integer, ii As Integer, jj As Integer, n As Integer
Dim ip As Integer
Dim nthpnt As Long
Dim sString As String
Dim temp3 As Single

Dim npts As Integer, nrows As Integer
Dim npixels As Long
Dim txdata() As Single, tydata() As Single, ttdata() As Single
Dim tldata() As Long
Dim tedata() As Single

Dim average As TypeAverage, average2 As TypeAverage

' Check for do not output to log window flag
If FormANALYZE.CheckDoNotOutputToLog.value = vbChecked Then Exit Sub

' Zero out the precision and detection limit arrays
For i% = 1 To MAXCHAN%
tPrecision!(i%) = 0#
tDetection!(i%) = 0#
aPrecision!(i%) = 0#
aDetection!(i%) = 0#
Next i%

' Write sample description
sString$ = TypeLoadString(sample())
If ierror Then Exit Sub
If mode% = 0 Then
Call IOWriteLogRichText(vbCrLf & sString$, vbNullString, Int(LogWindowFontSize% + 4), vbBlue, Int(FONT_ITALIC% Or FONT_UNDERLINE%), Int(0))
msg$ = TypeLoadDescription$(sample())
If ierror Then Exit Sub
Call IOWriteLog(vbCrLf & msg$)

' Print Number of lines and number of "good" lines and date and time
If UseDetailedFlag Then
msg$ = "Number of Data Lines: " & Format$(sample(1).Datarows%, a30$) & Space$(12)
msg$ = msg$ & " Number of 'Good' Data Lines: " & Format$(sample(1).GoodDataRows%, a30$)
Call IOWriteLog(msg$)
If sample(1).Datarows% > 0 Then
msg$ = "First/Last Date-Time: " & Format$(sample(1).DateTimes#(1), "mm/dd/yyyy hh:mm:ss AM/PM") & " to " & Format$(sample(1).DateTimes#(sample(1).Datarows%), "mm/dd/yyyy hh:mm:ss AM/PM")
Call IOWriteLog(msg$)
End If
End If

' Type warnings for analytical conditions
Call TypeWarnings(sample())
If ierror Then Exit Sub

' Type sample z-bar etc.
If UseDetailedFlag Then
Call TypeZbar(Int(1), analysis)
If ierror Then Exit Sub
End If

' Type Sample Flags
If UseDetailedFlag Then
Call TypeSampleFlags(analysis, sample())
If ierror Then Exit Sub
End If
Exit Sub    ' sample description only
End If

' If combined sample type arrays
If mode% = 1 And sample(1).CombinedConditionsFlag And UseDetailedFlag Then
Call TypeCombined(Int(1), sample())
If ierror Then Exit Sub
End If

' If elemental output and not using detailed output, skip output if unknown sample is display as oxide
If mode% = 1 And Not UseDetailedFlag Then
If sample(1).Type% = 2 And sample(1).DisplayAsOxideFlag Then Exit Sub
End If

' Write data type to log window
If mode% = 1 Then msg$ = sString$ & ", Results in Elemental Weight Percents"
If mode% = 2 Then msg$ = sString$ & ", Results in Oxide Weight Percents"
If mode% = 3 Then msg$ = sString$ & ", Results in Atomic Percents"
If mode% = 4 Then
If sample(1).FormulaElement$ <> vbNullString Then msg$ = sString$ & ", Results Based on " & sample(1).FormulaRatio! & " Atoms of " & sample(1).FormulaElement$
If sample(1).FormulaElement$ = vbNullString Then msg$ = sString$ & ", Results Based on Sum of " & sample(1).FormulaRatio! & " Cations"
End If
If mode% = 5 Then msg$ = sString$ & ", Results in Normalized Elemental Weight Percents (Particle Corrections)"
If mode% = 6 Then msg$ = sString$ & ", Results in Normalized Oxide Weight Percents (Particle Corrections)"
Call IOWriteLogRichText(vbCrLf & msg$, vbNullString, Int(LogWindowFontSize% + 2), vbBlue, Int(FONT_UNDERLINE%), Int(0))

' Specified elements
If Not PrintAnalyzedAndSpecifiedOnSameLineFlag Then
Call AnalyzeTypeSpecified(mode%, analysis, sample())
If ierror Then Exit Sub
End If

' Type out analyzed condition data for the sample
n% = 0
Do Until False
n% = n% + 1
If Not PrintAnalyzedAndSpecifiedOnSameLineFlag Then
Call TypeGetRange(Int(1), n%, ii%, jj%, sample())
If ierror Then Exit Sub
If ii% > sample(1).LastElm% Then Exit Do
Else
Call TypeGetRange(Int(2), n%, ii%, jj%, sample())
If ierror Then Exit Sub
If ii% > sample(1).LastChan% Then Exit Do
End If

' Type out symbols for conditions of analyzed elements (only mode=1)
If mode% = 1 Then
If UseDetailedFlag Or PrintAnalyzedAndSpecifiedOnSameLineFlag Then
msg$ = " "
Call IOWriteLog(msg$)
msg$ = "ELEM: "
For i% = ii% To jj%
msg$ = msg$ & Format$(sample(1).Elsyup$(i%), a80$)
Next i%
Call IOWriteLog(msg$)
End If

' Type specified element types if indicated
If PrintAnalyzedAndSpecifiedOnSameLineFlag Then
msg$ = "TYPE: "
For i% = ii% To jj%
If i% <= sample(1).LastElm% Then
msg$ = msg$ & Format$("ANAL", a80$)
Else
If sample(1).DifferenceElementFlag% And sample(1).DifferenceElement$ = sample(1).Elsyms$(i%) Then
msg$ = msg$ & Format$("DIFF", a80$)
ElseIf sample(1).DifferenceFormulaFlag% And ConvertIsDifferenceFormulaElement(sample(1).DifferenceFormula$, sample(1).Elsyms$(i%)) Then
msg$ = msg$ & Format$("FORM", a80$)
ElseIf sample(1).StoichiometryElementFlag% And sample(1).StoichiometryElement$ = sample(1).Elsyms$(i%) Then
msg$ = msg$ & Format$("STOI", a80$)
ElseIf sample(1).RelativeElementFlag% And sample(1).RelativeElement$ = sample(1).Elsyms$(i%) Then
msg$ = msg$ & Format$("RELA", a80$)
ElseIf sample(1).OxideOrElemental% = 1 And sample(1).OxygenChannel% = i% Then
msg$ = msg$ & Format$("CALC", a80$)
Else
msg$ = msg$ & Format$("SPEC", a80$)
End If
End If
Next i%
Call IOWriteLog(msg$)
End If

' Type out background correction types
If UseDetailedFlag Then
msg$ = "BGDS: "
For i% = ii% To jj%
If i% <= sample(1).LastElm% Then
If sample(1).IntegratedIntensitiesUseIntegratedFlags%(i%) Then
msg$ = msg$ & Format$("INT", a80$)
Else
If sample(1).BackgroundTypes%(i%) <> 1 Then  ' 0=off-peak, 1=MAN, 2=multipoint
If sample(1).CrystalNames$(i%) <> EDS_CRYSTAL$ Then
msg$ = msg$ & Format$(bgstrings$(sample(1).OffPeakCorrectionTypes%(i%)), a80$)
Else
msg$ = msg$ & Format$(EDS_CRYSTAL$, a80$)
End If
Else
msg$ = msg$ & Format$("MAN", a80$)
End If
End If
Else
msg$ = msg$ & Format$(vbNullString, a80$)
End If
Next i%
Call IOWriteLog(msg$)
End If

' Type out MAN absorption correction
If sample(1).MANBgdFlag Then
If UseDetailedFlag And UseMANParametersFlag Then

For j% = 1 To MAXCOEFF%
msg$ = "MAN" & Format$(j%) & ": "
For i% = ii% To jj%
If i% <= sample(1).LastElm% Then
msg$ = msg$ & MiscAutoFormat$(analysis.MANFitCoefficients!(j%, i%))
Else
msg$ = msg$ & Format$(vbNullString, a80$)
End If
Next i%
Call IOWriteLog(msg$)
Next j%

If UseMANAbsFlag Then
Call MathArrayAverage(average, RowUnkMANAbsCors!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
msg$ = "ABS%: "
For i% = ii% To jj%
If i% <= sample(1).LastElm% Then
msg$ = msg$ & Format$(Format$(average.averags!(i%), f82$), a80$)
Else
msg$ = msg$ & Format$(vbNullString, a80$)
End If
Next i%
Call IOWriteLog(msg$)
End If
End If
End If

' Type out the on-peak count time and beam current for each channel (average)
If UseDetailedFlag Then
Call MathArrayAverage(average, sample(1).OnTimeData!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
If Not sample(1).CombinedConditionsFlag Then      ' use OnBeamData in case of aggregate intensity calculation (use average aggregate beam)
Call MathArrayAverage(average2, sample(1).OnBeamData!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
Else                                              ' use OnBeamDataArray in case of aggregate intensity calculation (use average aggregate beam)
Call MathArrayAverage(average2, sample(1).OnBeamDataArray!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
End If

msg$ = "TIME: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

If i% <= sample(1).LastElm% Then
msg$ = msg$ & Format$(Format$(average.averags!(i%), f82$), a80$)
Else
msg$ = msg$ & Format$(vbNullString, a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)

msg$ = "BEAM: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

If i% <= sample(1).LastElm% Then
msg$ = msg$ & Format$(Format$(average2.averags!(i%), f82$), a80$)
Else
msg$ = msg$ & Format$(vbNullString, a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)

' Display number of aggregate channels
If UseDetailedFlag And UseAggregateIntensitiesFlag Then
msg$ = "AGGR: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

If i% <= sample(1).LastElm% And sample(1).AggregateNumChannels%(1, i%) > 1 Then
msg$ = msg$ & Format$(Format$(sample(1).AggregateNumChannels%(1, i%), i50$), a80$)
Else
msg$ = msg$ & Format$(a8x$, a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)
End If

End If
End If

' Type out weight percents and sum for each data line. Note, "analysis.WtsData!(linerow%,MAXCHAN+1)" contains the total weight percents.
msg$ = vbCrLf & "ELEM: "
For i% = ii% To jj%
If mode% = 2 Or mode% = 6 Then
If sample(1).DisableQuantFlag%(i%) = 1 Then
msg$ = msg$ & Format$(sample(1).Oxsyup$(i%) & "-D", a80$)
Else
msg$ = msg$ & Format$(sample(1).Oxsyup$(i%), a80$)
End If
Else
If sample(1).DisableQuantFlag%(i%) = 1 Then
msg$ = msg$ & Format$(sample(1).Elsyup$(i%) & "-D", a80$)
Else
msg$ = msg$ & Format$(sample(1).Elsyup$(i%), a80$)
End If
End If
Next i%
msg$ = msg$ & Format$("   SUM  ", a80$)
Call IOWriteLog(msg$)

If MiscIsElementDuplicated%(sample()) Then
msg$ = "XRAY: "
For i% = ii% To jj%
msg$ = msg$ & Format$("(" & sample(1).Xrsyms$(i%) & ")", a80$)
Next i%
Call IOWriteLog(msg$)
End If

' Print results for each line
For linerow% = 1 To sample(1).Datarows
If sample(1).LineStatus(linerow%) Then
msg$ = Format$(Format$(sample(1).Linenumber&(linerow%), i50$), a60$)
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 Then

If Not UseAutomaticFormatForResultsFlag Then
If mode% = 1 Then
msg$ = msg$ & Format$(Format$(analysis.WtsData!(linerow%, i%), f83$), a80$)
Else
msg$ = msg$ & Format$(Format$(analysis.CalData!(linerow%, i%), f83$), a80$)
End If
Else
If mode% = 1 Then
If i% > sample(1).LastElm% Then
msg$ = msg$ & AnalyzeFormatAnalysisResult$(Int(0), Int(0), analysis.WtsData!(linerow%, i%), analysis, sample())
Else
msg$ = msg$ & AnalyzeFormatAnalysisResult$(linerow%, i%, analysis.WtsData!(linerow%, i%), analysis, sample())
End If
Else
If i% > sample(1).LastElm% Then
msg$ = msg$ & AnalyzeFormatAnalysisResult$(Int(0), Int(0), analysis.CalData!(linerow%, i%), analysis, sample())
Else
msg$ = msg$ & AnalyzeFormatAnalysisResult$(linerow%, i%, analysis.CalData!(linerow%, i%), analysis, sample())
End If
End If
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%

' Output total sum for each line
If Not UseAutomaticFormatForResultsFlag Then
If mode% = 1 Then
msg$ = msg$ & Format$(Format$(analysis.WtsData!(linerow%, sample(1).LastChan% + 1), f83$), a80$)
Else
msg$ = msg$ & Format$(Format$(analysis.CalData!(linerow%, sample(1).LastChan% + 1), f83$), a80$)
End If
Else

If UseAutomaticFormatForResultsType% = 0 Then
If mode% = 1 Then
msg$ = msg$ & AnalyzeFormatAnalysisResult$(Int(0), Int(0), analysis.WtsData!(linerow%, sample(1).LastChan% + 1), analysis, sample())
Else
msg$ = msg$ & AnalyzeFormatAnalysisResult$(Int(0), Int(0), analysis.CalData!(linerow%, sample(1).LastChan% + 1), analysis, sample())
End If

' Calculate (weighted) precision for the analysis  (don't worry about detection limit, it should always be small relative to total)
Else
If mode% = 1 Then
Call MathAverageWeighted(temp3!, aPrecision!(), jj%, linerow%, analysis.WtsData!(), sample())    ' calculate weighted precision for this line
If ierror Then Exit Sub
Else
Call MathAverageWeighted(temp3!, aPrecision!(), jj%, linerow%, analysis.CalData!(), sample())    ' calculate weighted precision for this line
If ierror Then Exit Sub
End If

If mode% = 1 Then
msg$ = msg$ & MiscAutoFormatQ$(temp3!, NOT_ANALYZED_VALUE_SINGLE!, analysis.WtsData!(linerow%, sample(1).LastChan% + 1))
Else
msg$ = msg$ & MiscAutoFormatQ$(temp3!, NOT_ANALYZED_VALUE_SINGLE!, analysis.CalData!(linerow%, sample(1).LastChan% + 1))
End If
End If
End If

' Output results for line
Call IOWriteLog(msg$)
End If
Next linerow%

' Pre-calculate the average raw k-ratios and zero average results if the average kraw is < 0.
Call MathArrayAverage(average2, RowUnkKRaws!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub

' Type average measured weight percents and std deviation. Note that the average total sum is in "average.Averags!(sample(1).LastChan%+1)".
If mode% = 1 Then
Call MathArrayAverage(average, analysis.WtsData!(), sample(1).Datarows%, sample(1).LastChan% + 1, sample())
If ierror Then Exit Sub
Else
Call MathArrayAverage(average, analysis.CalData!(), sample(1).Datarows%, sample(1).LastChan% + 1, sample())
If ierror Then Exit Sub
End If

msg$ = vbCrLf & "AVER: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 Then

If Not UseAutomaticFormatForResultsFlag Or i% > sample(1).LastElm% Then
msg$ = msg$ & Format$(Format$(average.averags!(i%), f83$), a80$)

Else
If UseAutomaticFormatForResultsType% = 0 Then
msg$ = msg$ & AnalyzeFormatAnalysisResult$(Int(0), Int(0), average.averags!(i%), analysis, sample())
Else
msg$ = msg$ & MiscAutoFormatQ$(aPrecision!(i%), aDetection!(i%), average.averags!(i%))      ' use maximum precision values from analysis
End If
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%

' Output average total
If Not UseAutomaticFormatForResultsFlag Then
msg$ = msg$ & Format$(Format$(average.averags!(sample(1).LastChan% + 1), f83$), a80$)
Else

' Calculate average (weighted) precision for the analysis (don't worry about detection limit, it should always be small relative to total)
Call MathAverageWeighted2(temp3!, average.Reldevs!(), sample(1).LastElm%, average.averags!(), sample())
If ierror Then Exit Sub

' For 99% confidence use 3 deviations
temp3! = temp3! * 3#

' Use special formatting
If UseAutomaticFormatForResultsType% = 0 Then
msg$ = msg$ & AnalyzeFormatAnalysisResult$(Int(0), Int(0), average.averags!(sample(1).LastChan% + 1), analysis, sample())
Else
msg$ = msg$ & MiscAutoFormatQ$(temp3!, NOT_ANALYZED_VALUE_SINGLE!, average.averags!(sample(1).LastChan% + 1))
End If
End If
Call IOWriteLogRichText(msg$, vbNullString, Int(0), VbDarkBlue&, Int(0), Int(0))
        
msg$ = "SDEV: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 Then

If Not UseAutomaticFormatForResultsFlag Or i% > sample(1).LastElm% Then
msg$ = msg$ & Format$(Format$(average.Stddevs!(i%), f83$), a80$)
Else
msg$ = msg$ & AnalyzeFormatAnalysisResult$(Int(0), Int(0), average.Stddevs!(i%), analysis, sample())
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%

' Output average std devs
If Not UseAutomaticFormatForResultsFlag Then
msg$ = msg$ & Format$(Format$(average.Stddevs!(sample(1).LastChan% + 1), f83$), a80$)
Else
msg$ = msg$ & AnalyzeFormatAnalysisResult$(Int(0), Int(0), average.Stddevs!(sample(1).LastChan% + 1), analysis, sample())
End If
Call IOWriteLogRichText(msg$, vbNullString, Int(0), VbDarkBlue&, Int(0), Int(0))
    
If UseDetailedFlag Then
msg$ = "SERR: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 Then

If Not UseAutomaticFormatForResultsFlag Then
msg$ = msg$ & Format$(Format$(average.Stderrs!(i%), f83$), a80$)
Else
msg$ = msg$ & AnalyzeFormatAnalysisResult$(Int(0), Int(0), average.Stderrs!(i%), analysis, sample())
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)

msg$ = "%RSD: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 Then

ip% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())
If Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0) Then       ' check for duplicate element

If Abs(100# * average.Reldevs!(i%)) < MAXRELDEV% Then
If Not UseAutomaticFormatForResultsFlag Then
msg$ = msg$ & Format$(Format$(100# * average.Reldevs!(i%), f82$), a80$)
Else
msg$ = msg$ & AnalyzeFormatAnalysisResult$(Int(0), Int(0), 100# * average.Reldevs!(i%), analysis, sample())
End If
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If

Else
msg$ = msg$ & Format$(Format$(0#, f84$), a80$)   ' if duplicate element
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)
End If
   
' For analyses of standards, type out published composition
Call AnalyzeTypePublished(mode%, ii%, jj%, average, analysis, sample())
If ierror Then Exit Sub

' End of type out for standard samples, now type out standard assignments for each channel (unless calibration curve)
If (mode% = 1 Or mode% = 2) And CorrectionFlag% <> 5 Then
msg$ = "STDS: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

ip% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())
If Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0) Then       ' check for duplicate element
msg$ = msg$ & Format$(Format$(sample(1).StdAssigns%(i%), i50$), a80$)
Else
msg$ = msg$ & Format$(Format$(Int(0), i50$), a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLogRichText(msg$, vbNullString, Int(0), VbDarkBlue&, Int(0), Int(0))
End If

' Type standard k-factors (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If mode% = 1 Then
If UseDetailedFlag Then
If CorrectionFlag% = 0 Then
msg$ = vbCrLf & "STKF: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

ip% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())
If Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0) Then       ' check for duplicate element
msg$ = msg$ & Format$(Format$(analysis.StdAssignsKfactors!(i%), f84$), a80$)
Else
msg$ = msg$ & Format$(Format$(0#, f84$), a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)
End If
  
' Standard beta factors
If CorrectionFlag% > 0 And CorrectionFlag% < 5 Then
msg$ = vbCrLf & "STBE: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

ip% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())
If Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0) Then       ' check for duplicate element
msg$ = msg$ & Format$(Format$(analysis.StdAssignsBetas!(i%), f84$), a80$)
Else
msg$ = msg$ & Format$(Format$(0#, f84$), a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)
End If

' Type standard counts (unless calibration curve)
If CorrectionFlag% <> 5 Then
msg$ = "STCT: "
Call MathArrayAverage(average, RowStdAssignsCounts!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(average.averags!(i%), f81$), a80$)
Else
msg$ = msg$ & Format$(Format$(average.averags!(i%), f82$), a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)
End If
End If
  
' Type average sample K-ratio and ZAF correction factors and raw k-ratios
If UseDetailedFlag Then
If CorrectionFlag% = 0 Then
Call MathArrayAverage(average, RowUnkKrats!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
msg$ = vbCrLf & "UNKF: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

ip% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())
If Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0) Then       ' check for duplicate element
msg$ = msg$ & Format$(Format$(average.averags!(i%), f84$), a80$)
Else
msg$ = msg$ & Format$(Format$(0#, f84$), a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)
End If

If CorrectionFlag% > 0 And CorrectionFlag% < 5 Then
Call MathArrayAverage(average, RowUnkBetas!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
msg$ = vbCrLf & "UNBE: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

ip% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())
If Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0) Then       ' check for duplicate element
msg$ = msg$ & Format$(Format$(average.averags!(i%), f84$), a80$)
Else
msg$ = msg$ & Format$(Format$(0#, f84$), a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)
End If

' Type average sample counts (corrected for Off-Peak/MAN and interferences)
Call MathArrayAverage(average, sample(1).CorData!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
msg$ = "UNCT: "
If CorrectionFlag% = 5 Then msg$ = vbCrLf & msg$
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(average.averags!(i%), f81$), a80$)
Else
msg$ = msg$ & Format$(Format$(average.averags!(i%), f82$), a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)

' Type average sample background counts (corrected for MAN absorption)
Call MathArrayAverage(average, sample(1).BgdData!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
msg$ = "UNBG: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(average.averags!(i%), f81$), a80$)
Else
msg$ = msg$ & Format$(Format$(average.averags!(i%), f82$), a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)

' ZAF or Phi-Rho-Z corrections
If CorrectionFlag% = 0 Then
Call MathArrayAverage(average, RowUnkZAFCors!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
msg$ = vbCrLf & "ZCOR: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

ip% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())
If Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0) Then       ' check for duplicate element
msg$ = msg$ & Format$(Format$(average.averags!(i%), f84$), a80$)
Else
msg$ = msg$ & Format$(Format$(0#, f84$), a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)
End If

Call MathArrayAverage(average, RowUnkKRaws!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
msg$ = "KRAW: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

If Not UseAutomaticFormatForResultsFlag Then
msg$ = msg$ & Format$(Format$(average.averags!(i%), f84$), a80$)
Else
msg$ = msg$ & AnalyzeFormatAnalysisResult$(Int(0), Int(0), average.averags!(i%), analysis, sample())
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)
        
Call MathArrayAverage(average, RowUnkPeakToBgds!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
msg$ = "PKBG: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

If Not UseAutomaticFormatForResultsFlag Then
msg$ = msg$ & Format$(Format$(average.averags!(i%), f82$), a80$)
Else
msg$ = msg$ & AnalyzeFormatAnalysisResult$(Int(0), Int(0), average.averags!(i%), analysis, sample())
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)
End If
        
' Type interference correction percent change (unless calibration curve)
If UseInterfFlag Then
If CorrectionFlag% <> 5 Then
Call MathArrayAverage(average, RowUnkIntfCors!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
msg$ = "INT%: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

If average.averags!(i%) <> 0# Then
msg$ = msg$ & Format$(Format$(average.averags!(i%), f82$), a80$)
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)
End If
End If
        
' Type APF correction factors (unless calibration curve)
If UseAPFFlag And UseAPFOption% = 0 Then
If CorrectionFlag% <> 5 Then
Call MathArrayAverage(average, RowUnkAPFCors!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
msg$ = "APF:  "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

' Check if channel is not using integrated intensities and is using compound APF factors
If Not sample(1).IntegratedIntensitiesUseIntegratedFlags(i%) And EmpCheckAPF(sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%)) Then
msg$ = msg$ & Format$(Format$(average.averags!(i%), f83$), a80$)
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)
End If

' Type specified APF factors (unless calibration curve)
If UseAPFFlag And UseAPFOption% = 1 Then
If CorrectionFlag% <> 5 Then
msg$ = "APF*: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

If average.averags!(i%) <> 1# Then
msg$ = msg$ & Format$(Format$(sample(1).SpecifiedAreaPeakFactors!(i%), f83$), a80$)
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)
End If
End If

End If

' Type Volatile element correction percent change
If UseVolElFlag And CorrectionFlag% <> 5 Then
If (MiscIsDifferent(sample(1).LastElm%, sample(1).VolatileCorrectionUnks%()) Or Not MiscAllZero(sample(1).LastElm%, sample(1).VolatileCorrectionUnks%())) Then

' Average the volatile correction percentages
Call MathArrayAverage(average, RowUnkVolElCors!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
msg$ = vbCrLf & "TDI%: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

If sample(1).VolatileCorrectionUnks%(i%) <> 0 Then      ' -1 = self, 0 = none, >0 = assigned)
msg$ = msg$ & Format$(Format$(average.averags!(i%), f83$), a80$)
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)

Call MathArrayAverage(average, RowUnkVolElDevs!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
msg$ = "DEV%: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

If sample(1).VolatileCorrectionUnks%(i%) <> 0 Then
msg$ = msg$ & Format$(Format$(average.averags!(i%), f81$), a80$)
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)

' Only print volatile assignment if assigned (not self) (-1 = self, 0 = none, >0 = assigned)
If DebugMode Or Not MiscAllEqualToPassed(Int(-1), sample(1).LastElm%, sample(1).VolatileCorrectionUnks%()) Then
msg$ = "TDI#: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

If sample(1).VolatileCorrectionUnks%(i%) <> 0 Then
msg$ = msg$ & Format$(sample(1).VolatileCorrectionUnks%(i%), a80$)
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)
End If

msg$ = "TDIF: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

ip% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())
If sample(1).VolatileCorrectionUnks%(i%) <> 0 And (Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0)) Then       ' check for duplicate element
msg$ = msg$ & Format$(vstring$(sample(1).VolatileFitTypes%(i%)), a80$)
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)

msg$ = "TDIT: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

' Average the self volatile fit parameters (time and intercept). Check if called from Probe for EPMA (Nth sample rows) or CalcImage (Nth pixels)
ip% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())
If sample(1).VolatileCorrectionUnks%(i%) < 0 And (Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0)) Then       ' check for duplicate element
nthpnt& = 1                        ' use all sample rows or all pixels
If UCase$(Trim$(app.EXEName)) <> UCase$(Trim$("CalcImage")) Then
Call VolatileCalculateFitSelfAll(nthpnt&, txdata!(), tydata!(), ttdata!(), tedata!(), tldata&(), npts%, nrows%, i%, sample())
If ierror Then Exit Sub
Else
Call VolatileCalculateFitSelfAll_CI(nthpnt&, txdata!(), tydata!(), ttdata!(), tedata!(), tldata&(), npts%, npixels&, i%, sample())
If ierror Then Exit Sub
End If
msg$ = msg$ & Format$(Format$(sample(1).VolatileFitAvgTime!(i%), f82$), a80$)

ElseIf sample(1).VolatileCorrectionUnks%(i%) > 0 Then
Call VolatileCalculateFitAssigned(sample(1).VolatileCorrectionUnks%(i%), txdata!(), tydata!(), ttdata!(), tedata!(), npts%, i%, sample())
If ierror Then Exit Sub
msg$ = msg$ & Format$(Format$(sample(1).VolatileFitAvgTime!(i%), f82$), a80$)
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)

' Print volatile fit intercepts (log)
msg$ = "TDII: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

ip% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())
If sample(1).VolatileCorrectionUnks%(i%) <> 0 And sample(1).VolatileFitIntercepts!(i%) < MAXLOGEXPS! And (Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0)) Then        ' check for duplicate element
msg$ = msg$ & MiscAutoFormatBB$(NATURALE# ^ (sample(1).VolatileFitIntercepts!(i%)))
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)

' Print volatile fit intercepts (linear)
msg$ = "TDIL: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

ip% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())
If sample(1).VolatileCorrectionUnks%(i%) <> 0 And (Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0)) Then         ' check for duplicate element
msg$ = msg$ & MiscAutoFormatBB$(sample(1).VolatileFitIntercepts!(i%))
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)
End If
End If

' Print blank assignment
If sample(1).Type% <> 3 And UseBlankCorFlag Then
msg$ = "BLNK#:"
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

If sample(1).BlankCorrectionUnks%(i%) > 0 Then
ip% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())
If Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0) Then       ' check for duplicate element
msg$ = msg$ & Format$(sample(1).BlankCorrectionUnks%(i%), a80$)
Else
msg$ = msg$ & Format$(Format$(Int(0), i50$), a80$)
End If

Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)

msg$ = "BLNKL:"
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

If sample(1).BlankCorrectionUnks%(i%) > 0 Then
ip% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())
If Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0) Then       ' check for duplicate element
msg$ = msg$ & MiscAutoFormat$(sample(1).BlankCorrectionLevels!(i%))
Else
msg$ = msg$ & Format$(Format$(Int(0), i50$), a80$)
End If

Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)

msg$ = "BLNKV:"
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then

If sample(1).BlankCorrectionUnks%(i%) > 0 Then
ip% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())
If Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0) Then       ' check for duplicate element
msg$ = msg$ & MiscAutoFormat$(AnalBlankCorrectionPercents!(i%))
Else
msg$ = msg$ & Format$(Format$(Int(0), i50$), a80$)
End If

Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)
End If
End If

' Type calibration curve fits
If CorrectionFlag% = 5 Then
Call MathArrayAverage(average, RowCurve1Coeffs!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
msg$ = vbCrLf & "FIT1: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then
msg$ = msg$ & Format$(Format$(average.averags!(i%), f84$), a80$)

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)

Call MathArrayAverage(average, RowCurve2Coeffs!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
msg$ = "FIT2: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then
msg$ = msg$ & Format$(Format$(average.averags!(i%), f84$), a80$)

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)

Call MathArrayAverage(average, RowCurve3Coeffs!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
msg$ = "FIT3: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then
msg$ = msg$ & Format$(Format$(average.averags!(i%), f86$), a80$)

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)

Call MathArrayAverage(average, RowCurveFits!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
msg$ = "DEV:  "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).Xrsyms$(i%) <> vbNullString Then
msg$ = msg$ & Format$(Format$(average.averags!(i%), f81$), a80$)

Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if quant disabled
End If
Next i%
Call IOWriteLog(msg$)
End If
        
Loop
Exit Sub

' Errors
AnalyzeTypeResultsError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeTypeResults"
ierror = True
Exit Sub

End Sub

Sub AnalyzeReturnAnalysisFactors(mode As Integer, tarray() As Single, sample() As TypeSample)
' Return some of the analysis arrays
' mode = 1 return UnkKraws
' mode = 2 return UnkKrats
' mode = 3 return UnkZCors
' mode = 4 return UnkACors
' mode = 5 return UnkFCors
' mode = 6 return UnkZAFCors
' mode = 7 return UnkMACs
'
' mode = 8 return UnkIntfCors
' mode = 9 return UnkMANAbsCors
' mode = 10 return UnkAPFCors
' mode = 11 return UnkVolElCors
' mode = 12 return UnkVolElDevs

ierror = False
On Error GoTo AnalyzeReturnAnalysisFactorsError

Dim i As Integer, j As Integer

For j% = 1 To sample(1).Datarows%
For i% = 1 To sample(1).LastElm%
If mode% = 1 Then tarray!(j%, i%) = RowUnkKRaws!(j%, i%)
If mode% = 2 Then tarray!(j%, i%) = RowUnkKrats!(j%, i%)
If mode% = 3 Then tarray!(j%, i%) = RowUnkZCors!(j%, i%)
If mode% = 4 Then tarray!(j%, i%) = RowUnkACors!(j%, i%)
If mode% = 5 Then tarray!(j%, i%) = RowUnkFCors!(j%, i%)
If mode% = 6 Then tarray!(j%, i%) = RowUnkZAFCors!(j%, i%)
If mode% = 7 Then tarray!(j%, i%) = RowUnkMACs!(j%, i%)

If mode% = 8 Then tarray!(j%, i%) = RowUnkIntfCors!(j%, i%)
If mode% = 9 Then tarray!(j%, i%) = RowUnkMANAbsCors!(j%, i%)
If mode% = 10 Then tarray!(j%, i%) = RowUnkAPFCors!(j%, i%)
If mode% = 11 Then tarray!(j%, i%) = RowUnkVolElCors!(j%, i%)
If mode% = 12 Then tarray!(j%, i%) = RowUnkVolElDevs!(j%, i%)

Next i%
Next j%

Exit Sub

' Errors
AnalyzeReturnAnalysisFactorsError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeReturnAnalysisFactors"
ierror = True
Exit Sub

End Sub

Sub AnalyzeReturnAnalysisFactors2(mode As Integer, tvalue As Single, i As Integer, j As Integer)
' Return a single value from the analysis arrays (i = chan, j = row)
' mode = 1 return UnkKraws
' mode = 2 return UnkKrats
' mode = 3 return UnkZCors
' mode = 4 return UnkACors
' mode = 5 return UnkFCors
' mode = 6 return UnkZAFCors
' mode = 7 return UnkMACs

ierror = False
On Error GoTo AnalyzeReturnAnalysisFactors2Error

If mode% = 1 Then tvalue! = RowUnkKRaws!(j%, i%)
If mode% = 2 Then tvalue! = RowUnkKrats!(j%, i%)
If mode% = 3 Then tvalue! = RowUnkZCors!(j%, i%)
If mode% = 4 Then tvalue! = RowUnkACors!(j%, i%)
If mode% = 5 Then tvalue! = RowUnkFCors!(j%, i%)
If mode% = 6 Then tvalue! = RowUnkZAFCors!(j%, i%)
If mode% = 7 Then tvalue! = RowUnkMACs!(j%, i%)

Exit Sub

' Errors
AnalyzeReturnAnalysisFactors2Error:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeReturnAnalysisFactors2"
ierror = True
Exit Sub

End Sub

Sub AnalyzeReturnAnalysisFactors3(mode As Integer, row As Integer, tvalue As Single)
' Return a single value from the analysis arrays (row = linerow)
' mode = 1 return RowUnkTotalPercents
' mode = 2 return RowUnkTotalOxygens
' mode = 3 return RowUnkTotalCations
' mode = 4 return RowUnkCalculatedOxygens
' mode = 5 return RowUnkExcessOxygens
' mode = 6 return RowUnkZbars
' mode = 7 return RowUnkAtomicWeights
' mode = 8 return RowUnkOxygenFromHalogens
' mode = 9 return RowUnkHalogenCorrectedOxygen
' mode = 10 return RowUnkChargeBalance
' mode = 11 return RowUnkFeChargeBalance
' mode = 12 return RowUnkTotalAtoms

ierror = False
On Error GoTo AnalyzeReturnAnalysisFactors3Error

If mode% = 1 Then tvalue! = RowUnkTotalPercents!(row%)
If mode% = 2 Then tvalue! = RowUnkTotalOxygens!(row%)
If mode% = 3 Then tvalue! = RowUnkTotalCations!(row%)
If mode% = 4 Then tvalue! = RowUnkCalculatedOxygens!(row%)
If mode% = 5 Then tvalue! = RowUnkExcessOxygens!(row%)
If mode% = 6 Then tvalue! = RowUnkZbars!(row%)
If mode% = 7 Then tvalue! = RowUnkAtomicWeights!(row%)
If mode% = 8 Then tvalue! = RowUnkOxygenFromHalogens!(row%)
If mode% = 9 Then tvalue! = RowUnkHalogenCorrectedOxygen!(row%)
If mode% = 10 Then tvalue! = RowUnkChargeBalance!(row%)
If mode% = 11 Then tvalue! = RowUnkFeChargeBalance!(row%)
If mode% = 12 Then tvalue! = RowUnkTotalAtoms!(row%)

Exit Sub

' Errors
AnalyzeReturnAnalysisFactors3Error:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeReturnAnalysisFactors3"
ierror = True
Exit Sub

End Sub

Sub AnalyzeReturnStandardValues(mode As Integer, tvalue() As Single)
' Return the standard values from the analysis arrays (published, % variance, etc)
' mode = 1 return published values
' mode = 2 return percent variances
' mode = 3 return algebraic differences

ierror = False
On Error GoTo AnalyzeReturnStandardValuesError

Dim i As Integer

For i% = 1 To MAXCHAN%
If mode% = 1 Then tvalue!(i%) = StandardPublishedValues!(i%)
If mode% = 2 Then tvalue!(i%) = StandardPercentVariances!(i%)
If mode% = 3 Then tvalue!(i%) = StandardAlgebraicDifferences!(i%)
Next i%

Exit Sub

' Errors
AnalyzeReturnStandardValuesError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeReturnStandardValues"
ierror = True
Exit Sub

End Sub

Function AnalyzeFormatAnalysisResult(row As Integer, chan As Integer, avalue As Single, analysis As TypeAnalysis, sample() As TypeSample) As String
' Format the value based on UseAutomaticFormatForResultsType (0=maximum decimals, 1=significant decimals)

ierror = False
On Error GoTo AnalyzeFormatAnalysisResultError

Dim percenterror As Single, detectionlimit As Single, wtscalratio As Single

' Format for maximum number of decimals
If UseAutomaticFormatForResultsType% = 0 Then
AnalyzeFormatAnalysisResult$ = MiscAutoFormatA$(avalue!)
End If

' Format for significant number of digits, based on counting statistics
If UseAutomaticFormatForResultsType% = 1 And (row% <> 0 And chan% <> 0) Then
percenterror! = ConvertAnalyticalSensitivity2!(row%, chan%, sample())
If ierror Then Exit Function
detectionlimit! = ConvertDetectionLimits2!(row%, chan%, RowUnkZAFCors!(), RowStdAssignsCounts!(), analysis, sample())
If ierror Then Exit Function
wtscalratio! = 1#
If analysis.CalData!(row%, chan%) <> 0# Then
wtscalratio! = analysis.WtsData!(row%, chan%) / analysis.CalData!(row%, chan%)  ' for oxide/atomic scaling
If chan% <= sample(1).LastElm% And sample(1).AtomicNums%(chan%) = 8 And sample(1).HydrogenStoichiometryFlag Then wtscalratio! = 1#      ' turn off scaling for hydrogen stoichiometry on excess oxygen
detectionlimit! = detectionlimit! / wtscalratio!
End If
AnalyzeFormatAnalysisResult$ = MiscAutoFormatQ$(percenterror!, detectionlimit!, avalue!)

' Save precision and detection limits (to module level arrays) for calculation of total (sum) statistics
tPrecision!(chan%) = Abs(percenterror!)
tDetection!(chan%) = detectionlimit!

' Save maximum precision and detection limits (to module level arrays) for calculation of average statistics
If Abs(percenterror!) < aPrecision!(chan%) Or aPrecision!(chan%) = 0# Then aPrecision!(chan%) = Abs(percenterror!)
If detectionlimit! < aDetection!(chan%) Or aDetection!(chan%) = 0# Then aDetection!(chan%) = detectionlimit!
End If

' Cannot format for significant digits, just do normal format
If UseAutomaticFormatForResultsType% = 1 And (row% = 0 Or chan% = 0) Then
AnalyzeFormatAnalysisResult$ = Format$(Format$(avalue!, f83$), a80$)
End If

Exit Function

' Errors
AnalyzeFormatAnalysisResultError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeFormatAnalysisResult"
ierror = True
Exit Function

End Function

Sub AnalyzeTypePublished(mode As Integer, ii As Integer, jj As Integer, average As TypeAverage, analysis As TypeAnalysis, sample() As TypeSample)
' Type out published standard values
'   mode = 1 element
'   mode = 2 oxide

ierror = False
On Error GoTo AnalyzeTypePublishedError

Dim ip As Integer, i As Integer, ipp As Integer, ippp As Integer
Dim sum As Single, stdpercent As Single

ReDim temp(1 To MAXCHAN%) As Single

' Print out element or oxide published values
If (mode% = 1 Or mode% = 2) And sample(1).Type = 1 Then
ip% = IPOS2(NumberofStandards%, sample(1).number%, StandardNumbers%())
If ip% = 0 Then GoTo AnalyzeTypePublishedNotFound

' Get the standard composition from standard database (for sum and excess oxygen calculations)
Call StandardGetMDBStandard(sample(1).number%, stdsample())
If ierror Then Exit Sub

' Sum standard database weight percents (use elemental percent to avoid problems with missing trace oxygen from unspecified elements)
sum! = 0#
For i% = 1 To stdsample(1).LastChan%
stdpercent! = stdsample(1).ElmPercents!(i%)    ' load all elements analyzed and specified in sample
ipp% = IPOS1(sample(1).LastChan%, stdsample(1).Elsyms$(i%), sample(1).Elsyms$())
If ipp% > 0 Then
sum! = sum! + stdpercent!
End If
Next i%

msg$ = vbCrLf & "PUBL: "
For i% = ii% To jj%
stdpercent! = analysis.StdPercents!(ip%, i%)
ippp% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())   ' check if duplicate element
If mode% = 2 And analysis.StdPercents!(ip%, i%) <> NOT_ANALYZED_VALUE_SINGLE! Then
stdpercent! = ConvertElmToOxd(analysis.StdPercents!(ip%, i%), sample(1).Elsyms$(i%), sample(1).numcat%(i%), sample(1).numoxd%(i%))
If i% = sample(1).OxygenChannel% Then stdpercent! = ConvertTotalToExcessOxygen!(mode%, sample(), stdsample())
End If
If stdpercent! <> NOT_ANALYZED_VALUE_SINGLE! And sample(1).DisableQuantFlag%(i%) = 0 And (Not UseAggregateIntensitiesFlag Or UseAggregateIntensitiesFlag And ippp% = 0) Then
If Not UseAutomaticFormatForResultsFlag Then
msg$ = msg$ & Format$(Format$(stdpercent!, f83$), a80$)
StandardPublishedValues!(i%) = stdpercent!
Else
msg$ = msg$ & AnalyzeFormatAnalysisResult$(Int(0), Int(0), stdpercent!, analysis, sample())
StandardPublishedValues!(i%) = stdpercent!
End If
Else
msg$ = msg$ & Format$("    n.a.", a80$)
StandardPublishedValues!(i%) = 0#
End If
Next i%
If Not UseAutomaticFormatForResultsFlag Then
msg$ = msg$ & Format$(Format$(sum!, f83$), a80$)
Else
msg$ = msg$ & AnalyzeFormatAnalysisResult$(Int(0), Int(0), sum!, analysis, sample())
End If
Call IOWriteLogRichText(msg$, vbNullString, Int(0), VbDarkBlue&, Int(0), Int(0))
  
' Calculate percent error from published percents
For i% = ii% To jj%
stdpercent! = analysis.StdPercents!(ip%, i%)
ippp% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())   ' check if duplicate element
If mode% = 2 And analysis.StdPercents!(ip%, i%) <> NOT_ANALYZED_VALUE_SINGLE! Then
stdpercent! = ConvertElmToOxd(analysis.StdPercents!(ip%, i%), sample(1).Elsyms$(i%), sample(1).numcat%(i%), sample(1).numoxd%(i%))
If i% = sample(1).OxygenChannel% Then stdpercent! = ConvertTotalToExcessOxygen!(mode%, sample(), stdsample())
End If
temp!(i%) = 0#
If stdpercent! <> NOT_ANALYZED_VALUE_SINGLE! And stdpercent! <> 0# And (Not UseAggregateIntensitiesFlag Or UseAggregateIntensitiesFlag And ippp% = 0) Then
temp!(i%) = (average.averags!(i%) - stdpercent!) * 100# / stdpercent!
End If
Next i%

msg$ = "%VAR: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 Then
If analysis.StdPercents!(ip%, i%) <> NOT_ANALYZED_VALUE_SINGLE! Then
If sample(1).StdAssigns%(i%) <> sample(1).number% Then
msg$ = msg$ & Format$(Format$(temp!(i%), f82$), a80$)
StandardPercentVariances!(i%) = temp!(i%)
Else
msg$ = msg$ & Format$("(" & Format$(temp!(i%), f82$) & ")", a80$)
StandardPercentVariances!(i%) = temp!(i%)
End If
Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if not analyzed value
StandardPercentVariances!(i%) = 0#
End If
Else
msg$ = msg$ & Format$(DASHED3$, a80$)
StandardPercentVariances!(i%) = 0#
End If
Next i%
Call IOWriteLogRichText(msg$, vbNullString, Int(0), VbDarkBlue&, Int(0), Int(0))
  
' Calculate absolute variance (subtraction)
If UseDetailedFlag Then
For i% = ii% To jj%
stdpercent! = analysis.StdPercents!(ip%, i%)
If mode% = 2 And analysis.StdPercents!(ip%, i%) <> NOT_ANALYZED_VALUE_SINGLE! Then
stdpercent! = ConvertElmToOxd(analysis.StdPercents!(ip%, i%), sample(1).Elsyms$(i%), sample(1).numcat%(i%), sample(1).numoxd%(i%))
If i% = sample(1).OxygenChannel% Then stdpercent! = ConvertTotalToExcessOxygen!(mode%, sample(), stdsample())
End If
temp!(i%) = 0#
If stdpercent! <> NOT_ANALYZED_VALUE_SINGLE! And stdpercent! <> 0# Then
temp!(i%) = average.averags!(i%) - stdpercent!
End If
Next i%

msg$ = "DIFF: "
For i% = ii% To jj%
ippp% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())   ' check if duplicate element
If sample(1).DisableQuantFlag%(i%) = 0 And (Not UseAggregateIntensitiesFlag Or UseAggregateIntensitiesFlag And ippp% = 0) Then
If analysis.StdPercents!(ip%, i%) <> NOT_ANALYZED_VALUE_SINGLE! Then
If Not UseAutomaticFormatForResultsFlag Then
If sample(1).StdAssigns%(i%) <> sample(1).number% Then
msg$ = msg$ & Format$(Format$(temp!(i%), f83$), a80$)
StandardAlgebraicDifferences!(i%) = temp!(i%)
Else
msg$ = msg$ & Format$("(" & Format$(temp!(i%), f82$) & ")", a80$)
StandardAlgebraicDifferences!(i%) = temp!(i%)
End If
Else
If sample(1).StdAssigns%(i%) <> sample(1).number% Then
msg$ = msg$ & AnalyzeFormatAnalysisResult$(Int(0), Int(0), temp!(i%), analysis, sample())
StandardAlgebraicDifferences!(i%) = temp!(i%)
Else
msg$ = msg$ & Format$("(" & Format$(temp!(i%), f82$) & ")", a80$)
StandardAlgebraicDifferences!(i%) = temp!(i%)
End If
End If
Else
msg$ = msg$ & Format$(DASHED3$, a80$)   ' if not analyzed value
StandardAlgebraicDifferences!(i%) = 0#
End If
Else
msg$ = msg$ & Format$(DASHED3$, a80$)
StandardAlgebraicDifferences!(i%) = 0#
End If
Next i%
Call IOWriteLog(msg$)
End If
End If

Exit Sub

' Errors
AnalyzeTypePublishedError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeTypePublished"
ierror = True
Exit Sub

AnalyzeTypePublishedNotFound:
msg$ = "Standard weight percent data was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "AnalyzeTypePublished"
ierror = True
Exit Sub

End Sub

Sub AnalyzePrintEquationParameters(mode As Integer, linerow As Integer, j As Integer, chan As Integer, intfchan As Integer, intfstd As Integer, intfcts As Single, analysis As TypeAnalysis, sample() As TypeSample)
' Print out the equation parameters indicated by "mode"
'  mode = 0 print full quant expression
'  mode = 1 print Gilfrich expression

ierror = False
On Error GoTo AnalyzePrintEquationParametersError

Dim astring As String, bstring As String, cstring As String
Dim dstring As String, estring As String, fstring As String

' Print name and line number
Call IOWriteLog(vbCrLf & "Interference correction calculations for " & SampleGetString2$(sample()) & ", line " & Format$(sample(1).Linenumber&(j%)))

' Print full quant interference parameters
If mode% = 0 Then

' Load ZAF strings (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% = 0 Then
dstring$ = "(" & Trim$(MiscAutoFormat$(analysis.StdZAFCors!(4, intfstd%, chan%))) & ")"
estring$ = "(" & Trim$(MiscAutoFormat$(analysis.UnkZAFCors!(4, chan%))) & ")"
ElseIf CorrectionFlag% > 0 And CorrectionFlag% < 5 Then
dstring$ = "(" & Trim$(MiscAutoFormat$(analysis.StdBetas!(intfstd%, chan%))) & ")"
estring$ = "(" & Trim$(MiscAutoFormat$(analysis.UnkBetas!(chan%))) & ")"
End If

' Load base equation string
astring$ = "(" & Trim$(MiscAutoFormat$(analysis.WtPercents!(intfchan%))) & ")"
bstring$ = "(" & Trim$(MiscAutoFormat$(analysis.StdPercents!(intfstd%, intfchan%))) & ")"
fstring$ = "(" & Trim$(MiscAutoFormat$(analysis.StdAssignsIntfCounts!(j%, chan%))) & ")"

' Full matrix correction (also use only if first order interference)
If DisableMatrixCorrectionInterferenceCorrectionFlag = 0 And sample(1).StdAssignsIntfOrders%(j%, chan%) = 1 Then
Call IOWriteLogRTF(ProgramPath$ & "equation4a.rtf")

astring$ = Space$(29) & dstring$ & " " & astring$
cstring$ = "Cps " & Format$(sample(1).Elsyms$(chan%), a20$) & " " & sample(1).Xrsyms$(chan%) & " by " & Format$(sample(1).Elsyms$(intfchan%), a20$) & " = " & Trim$(MiscAutoFormat$(intfcts!)) & " = ---------------------"
bstring$ = Space$(29) & estring$ & " " & bstring$

Call IOWriteLog(astring$)
Call IOWriteLog(cstring$ & " " & fstring$)
Call IOWriteLog(bstring$)

' No matrix correction
Else
Call IOWriteLogRTF(ProgramPath$ & "equation4b.rtf")

astring$ = Space$(29) & astring$
cstring$ = "Cps " & Format$(sample(1).Elsyms$(chan%), a20$) & " " & sample(1).Xrsyms$(chan%) & " by " & Format$(sample(1).Elsyms$(intfchan%), a20$) & " = " & Trim$(MiscAutoFormat$(intfcts!)) & " = -----------"
bstring$ = Space$(29) & bstring$

Call IOWriteLog(astring$)
Call IOWriteLog(cstring$ & " " & fstring$)
Call IOWriteLog(bstring$)
End If

End If

' Print Gilfrich interference parameters
If mode% = 1 Then
Call IOWriteLogRTF(ProgramPath$ & "equation8.rtf")

' Load base equation string
astring$ = "(" & Trim$(MiscAutoFormat$(analysis.StdAssignsIntfCounts!(j%, chan%))) & ")"
bstring$ = "(" & Trim$(MiscAutoFormat$(analysis.StdPercents!(intfstd%, intfchan%))) & ")"
dstring$ = "(" & Trim$(MiscAutoFormat$(sample(1).CorData!(linerow%, intfchan%))) & ")"
estring$ = "(" & Trim$(MiscAutoFormat$(analysis.StdAssignsCounts(intfchan%))) & ")"

fstring$ = "(" & Trim$(MiscAutoFormat$(analysis.StdAssignsPercents(intfchan%))) & ")"

astring$ = Space$(29) & astring$
cstring$ = "Cps " & Format$(sample(1).Elsyms$(chan%), a20$) & " " & sample(1).Xrsyms$(chan%) & " by " & Format$(sample(1).Elsyms$(intfchan%), a20$) & " = " & Trim$(MiscAutoFormat$(intfcts!)) & " = ---------------------"
bstring$ = Space$(29) & bstring$

Call IOWriteLog(astring$ & " " & dstring$)
Call IOWriteLog(cstring$ & " " & fstring$)
Call IOWriteLog(bstring$ & " " & estring$)
End If

Exit Sub

' Errors
AnalyzePrintEquationParametersError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzePrintEquationParameters"
ierror = True
Exit Sub

End Sub

