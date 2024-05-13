Attribute VB_Name = "CodeGETELM2"
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

Dim GetElmOldSample(1 To 1) As TypeSample
Dim GetElmTmpSample(1 To 1) As TypeSample

Sub GetElmSaveSampleOnly(method As Integer, sample() As TypeSample, chan As Integer, chan2 As Integer)
' Loads the passed sample and saves it re-ordered
'   method = 0 normal call (check for different element, x-ray, spectro and crystal)
'   method = 1 MAN call (check for different element, x-ray, spectro and crystal *and* keV)

ierror = False
On Error GoTo GetElmSaveSampleOnlyError

' Load the passed sample
GetElmTmpSample(1) = sample(1)

' Save the sample
Call GetElmSave(method%, chan%, chan2%)

' Return the saved sample
sample(1) = GetElmTmpSample(1)

Exit Sub

' Errors
GetElmSaveSampleOnlyError:
MsgBox Error$, vbOKOnly + vbCritical, "GetElmSaveSampleOnly"
ierror = True
Exit Sub

End Sub

Sub GetElmSave(method As Integer, chan As Integer, chan2 As Integer)
' Load the GetElmOldSample array from the GetElmTmpSample and sort the elements into
' analyzed and specified in case the user has changed the analyzed and
' specified elements (added or deleted elements).
'   method = 0 normal call (check for different element, x-ray, spectro and crystal)
'   method = 1 MAN call (check for different element, x-ray, spectro and crystal *and* keV)

ierror = False
On Error GoTo GetElmSaveError

Dim sym As String
Dim i As Integer, j As Integer, n As Integer, i2 As Integer, m As Integer
Dim ip As Integer, ipp As Integer, ippp As Integer

' Pre-load the GetElmOldSample arrays (kilovolts, takeoff, beam counts, etc)
GetElmOldSample(1) = GetElmTmpSample(1)
GetElmOldSample(1).LastElm% = 0
GetElmOldSample(1).LastChan% = 0

' Load the analyzed elements first
For n% = 1 To MAXCHAN%
i% = n%

' Initialize the element channel
Call InitElement(i%, GetElmOldSample())
If ierror Then Exit Sub

' See if sorting channels
If chan% > 0 And chan2% > 0 Then

' Shift channel up
If chan% < chan2% Then
If n% = chan% Then
i% = chan2% ' reset to skipped channel
End If
If n% = chan2% Then
i% = chan%    ' reset to new channel
End If
End If

' Shift channel down
If chan% > chan2% Then
If n% = chan2% Then
i2% = n%    ' save skipped channel
i% = chan% ' reset to new channel
End If
If n% = chan% Then
i% = i2%    ' reset to skipped channel
End If
End If
End If

' Find element and xray symbol
sym$ = GetElmTmpSample(1).Elsyms$(i%)
ip% = IPOS1(MAXELM%, sym$, Symlo$())

sym$ = GetElmTmpSample(1).Xrsyms$(i%)
ipp% = IPOS1(MAXRAY%, sym, Xraylo$())   ' including unanalyzed element

' Skip if element if NOT analyzed
If ip% = 0 Or ipp% = 0 Or ipp% > MAXRAY% - 1 Then GoTo 2000

' Check for *analyzed* element already loaded (element, x-ray, motor, crystal) (also add check for different keV, for call from MAN dialog, 08-17-2017)
If method% = 0 Then ippp% = IPOS5(Int(0), i%, GetElmTmpSample(), GetElmOldSample())
If method% = 0 Then ippp% = IPOS13B(Int(0), GetElmTmpSample(1).Elsyms$(i%), GetElmTmpSample(1).Xrsyms$(i%), GetElmTmpSample(1).MotorNumbers%(i%), GetElmTmpSample(1).CrystalNames$(i%), GetElmTmpSample(1).KilovoltsArray!(i%), GetElmOldSample())
If ippp% > 0 Then
If method% = 0 Then msg$ = "Error in " & SampleGetString2$(GetElmTmpSample()) & ", " & GetElmTmpSample(1).Elsyms$(i%) & " " & GetElmTmpSample(1).Xrsyms$(i%) & ", Spectrometer " & Str$(GetElmTmpSample(1).MotorNumbers%(i%)) & " " & GetElmTmpSample(1).CrystalNames$(i%) & ", is already present as an analyzed element, it will be skipped"
If method% = 1 Then msg$ = "Error in " & SampleGetString2$(GetElmTmpSample()) & ", " & GetElmTmpSample(1).Elsyms$(i%) & " " & GetElmTmpSample(1).Xrsyms$(i%) & ", Spectrometer " & Str$(GetElmTmpSample(1).MotorNumbers%(i%)) & " " & GetElmTmpSample(1).CrystalNames$(i%) & ", " & Format$(GetElmTmpSample(1).KilovoltsArray!(i%)) & " keV, is already present as an analyzed element, it will be skipped"
MsgBox msg$, vbOKOnly + vbExclamation, "GetElmSave"
GoTo 2000
End If

' Increment number of analyzed elements
GetElmOldSample(1).LastElm% = GetElmOldSample(1).LastElm% + 1
GetElmOldSample(1).Elsyms$(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).Elsyms$(i%)
GetElmOldSample(1).Xrsyms$(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).Xrsyms$(i%)
    
' Make sure cations and oxygens are loaded
If GetElmTmpSample(1).numcat%(i%) = 0 Or (GetElmTmpSample(1).numcat%(i%) = 0 And GetElmTmpSample(1).numoxd%(i%) = 0) Then
GetElmTmpSample(1).numcat%(i%) = AllCat%(ip%)
GetElmTmpSample(1).numoxd%(i%) = AllOxd%(ip%)
End If
GetElmOldSample(1).numcat%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).numcat%(i%)
GetElmOldSample(1).numoxd%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).numoxd%(i%)
GetElmOldSample(1).ElmPercents!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).ElmPercents!(i%)

GetElmOldSample(1).AtomicCharges!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).AtomicCharges!(i%)
GetElmOldSample(1).AtomicWts!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).AtomicWts!(i%)

' Load off peak correction type, 0=linear, 1=average, 2=high only, 3=low only, 4=exponential, 5=slope hi, 6=slope lo, 7=polynomial, 8=multi-point
GetElmOldSample(1).OffPeakCorrectionTypes%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).OffPeakCorrectionTypes%(i%)

' Load slope and polynomial coefficients
GetElmOldSample(1).BackgroundExponentialBase!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).BackgroundExponentialBase!(i%)
GetElmOldSample(1).BackgroundSlopeCoefficients!(1, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).BackgroundSlopeCoefficients!(1, i%)
GetElmOldSample(1).BackgroundSlopeCoefficients!(2, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).BackgroundSlopeCoefficients!(2, i%)

GetElmOldSample(1).BackgroundPolynomialPositions!(1, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).BackgroundPolynomialPositions!(1, i%)
GetElmOldSample(1).BackgroundPolynomialPositions!(2, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).BackgroundPolynomialPositions!(2, i%)
GetElmOldSample(1).BackgroundPolynomialPositions!(3, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).BackgroundPolynomialPositions!(3, i%)

GetElmOldSample(1).BackgroundPolynomialCoefficients!(1, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).BackgroundPolynomialCoefficients!(1, i%)
GetElmOldSample(1).BackgroundPolynomialCoefficients!(2, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).BackgroundPolynomialCoefficients!(2, i%)
GetElmOldSample(1).BackgroundPolynomialCoefficients!(3, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).BackgroundPolynomialCoefficients!(3, i%)

GetElmOldSample(1).BackgroundPolynomialNominalBeam!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).BackgroundPolynomialNominalBeam!(i%)

' Load other real time element parameters
GetElmOldSample(1).BraggOrders%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).BraggOrders%(i%)
GetElmOldSample(1).MotorNumbers%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).MotorNumbers%(i%)

If GetElmOldSample(1).CrystalNames$(GetElmOldSample(1).LastElm%) <> EDS_CRYSTAL$ Then
GetElmOldSample(1).OrderNumbers%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).OrderNumbers%(i%)
Else
GetElmOldSample(1).OrderNumbers%(GetElmOldSample(1).LastElm%) = 1           ' EDS element order is always one
End If

GetElmOldSample(1).CrystalNames$(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).CrystalNames$(i%)
GetElmOldSample(1).Crystal2ds!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).Crystal2ds!(i%)
GetElmOldSample(1).CrystalKs!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).CrystalKs!(i%)

GetElmOldSample(1).OnPeaks!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).OnPeaks!(i%)
GetElmOldSample(1).HiPeaks!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).HiPeaks!(i%)
GetElmOldSample(1).LoPeaks!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).LoPeaks!(i%)

GetElmOldSample(1).StdAssigns%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).StdAssigns%(i%)
GetElmOldSample(1).BackgroundTypes%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).BackgroundTypes%(i%)  ' 0=off-peak, 1=MAN, 2=multipoint
GetElmOldSample(1).DeadTimes!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).DeadTimes!(i%)

GetElmOldSample(1).Baselines!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).Baselines!(i%)
GetElmOldSample(1).Windows!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).Windows!(i%)
GetElmOldSample(1).Gains!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).Gains!(i%)
GetElmOldSample(1).Biases!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).Biases!(i%)
GetElmOldSample(1).InteDiffModes%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).InteDiffModes%(i%)

GetElmOldSample(1).DetectorSlitSizes$(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).DetectorSlitSizes$(i%)
GetElmOldSample(1).DetectorSlitPositions$(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).DetectorSlitPositions$(i%)
GetElmOldSample(1).DetectorModes$(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).DetectorModes$(i%)

GetElmOldSample(1).IntegratedIntensitiesUseIntegratedFlags%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).IntegratedIntensitiesUseIntegratedFlags%(i%)
GetElmOldSample(1).IntegratedIntensitiesInitialStepSizes!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).IntegratedIntensitiesInitialStepSizes!(i%)
GetElmOldSample(1).IntegratedIntensitiesMinimumStepSizes!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).IntegratedIntensitiesMinimumStepSizes!(i%)
GetElmOldSample(1).IntegratedIntensitiesIntegratedTypes%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).IntegratedIntensitiesIntegratedTypes%(i%)

' Interference correction
For j% = 1 To MAXINTF%
GetElmOldSample(1).StdAssignsIntfElements$(j%, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).StdAssignsIntfElements$(j%, i%)
GetElmOldSample(1).StdAssignsIntfXrays$(j%, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).StdAssignsIntfXrays$(j%, i%)
GetElmOldSample(1).StdAssignsIntfStds%(j%, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).StdAssignsIntfStds%(j%, i%)
GetElmOldSample(1).StdAssignsIntfOrders%(j%, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).StdAssignsIntfOrders%(j%, i%)
Next j%

' Volatile element correction and specified apf
GetElmOldSample(1).VolatileCorrectionUnks%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).VolatileCorrectionUnks%(i%)
GetElmOldSample(1).SpecifiedAreaPeakFactors!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).SpecifiedAreaPeakFactors!(i%)

' Blank correction
GetElmOldSample(1).BlankCorrectionUnks%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).BlankCorrectionUnks%(i%)
GetElmOldSample(1).BlankCorrectionLevels!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).BlankCorrectionLevels!(i%)

' MAN correction
GetElmOldSample(1).MANAbsCorFlags%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).MANAbsCorFlags%(i%)
GetElmOldSample(1).MANLinearFitOrders%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).MANLinearFitOrders%(i%)
For j% = 1 To MAXMAN%
GetElmOldSample(1).MANStdAssigns%(j%, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).MANStdAssigns%(j%, i%)
Next j%

' Count data and time
For j% = 1 To MAXROW%
GetElmOldSample(1).OnPeakCounts!(j%, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).OnPeakCounts!(j%, i%)
GetElmOldSample(1).HiPeakCounts!(j%, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).HiPeakCounts!(j%, i%)
GetElmOldSample(1).LoPeakCounts!(j%, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).LoPeakCounts!(j%, i%)
GetElmOldSample(1).OnCountTimes!(j%, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).OnCountTimes!(j%, i%)
GetElmOldSample(1).HiCountTimes!(j%, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).HiCountTimes!(j%, i%)
GetElmOldSample(1).LoCountTimes!(j%, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).LoCountTimes!(j%, i%)
GetElmOldSample(1).UnknownCountFactors!(j%, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).UnknownCountFactors!(j%, i%)
GetElmOldSample(1).UnknownMaxCounts&(j%, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).UnknownMaxCounts&(j%, i%)

GetElmOldSample(1).VolCountTimesStart(j%, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).VolCountTimesStart(j%, i%)
GetElmOldSample(1).VolCountTimesStop(j%, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).VolCountTimesStop(j%, i%)
GetElmOldSample(1).VolCountTimesDelay!(j%, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).VolCountTimesDelay!(j%, i%)

GetElmOldSample(1).OnBeamCountsArray!(j%, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).OnBeamCountsArray!(j%, i%)
GetElmOldSample(1).AbBeamCountsArray!(j%, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).AbBeamCountsArray!(j%, i%)
GetElmOldSample(1).OnBeamCountsArray2!(j%, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).OnBeamCountsArray2!(j%, i%)
GetElmOldSample(1).AbBeamCountsArray2!(j%, GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).AbBeamCountsArray2!(j%, i%)
Next j%

' Save last count times
GetElmOldSample(1).LastOnCountTimes!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).LastOnCountTimes!(i%)
GetElmOldSample(1).LastHiCountTimes!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).LastHiCountTimes!(i%)
GetElmOldSample(1).LastLoCountTimes!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).LastLoCountTimes!(i%)
GetElmOldSample(1).LastWaveCountTimes!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).LastWaveCountTimes!(i%)
GetElmOldSample(1).LastPeakCountTimes!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).LastPeakCountTimes!(i%)
GetElmOldSample(1).LastQuickCountTimes!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).LastQuickCountTimes!(i%)
GetElmOldSample(1).LastCountFactors!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).LastCountFactors!(i%)
GetElmOldSample(1).LastMaxCounts&(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).LastMaxCounts&(i%)

' Load combined conditions
If GetElmTmpSample(1).CombinedConditionsFlag Then
GetElmOldSample(1).TakeoffArray!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).TakeoffArray!(i%)
GetElmOldSample(1).KilovoltsArray!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).KilovoltsArray!(i%)
GetElmOldSample(1).BeamCurrentArray!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).BeamCurrentArray!(i%)
GetElmOldSample(1).BeamSizeArray!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).BeamSizeArray!(i%)

GetElmOldSample(1).ColumnConditionMethodArray%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).ColumnConditionMethodArray%(i%)
GetElmOldSample(1).ColumnConditionStringArray$(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).ColumnConditionStringArray$(i%)

' If any combined conditions are zero or blank, use sample level parameters
If GetElmOldSample(1).TakeoffArray!(GetElmOldSample(1).LastElm%) = 0# Then GetElmOldSample(1).TakeoffArray!(GetElmOldSample(1).LastElm%) = GetElmOldSample(1).takeoff!
If GetElmOldSample(1).KilovoltsArray!(GetElmOldSample(1).LastElm%) = 0# Then GetElmOldSample(1).KilovoltsArray!(GetElmOldSample(1).LastElm%) = GetElmOldSample(1).kilovolts!
If GetElmOldSample(1).BeamCurrentArray!(GetElmOldSample(1).LastElm%) = 0# Then GetElmOldSample(1).BeamCurrentArray!(GetElmOldSample(1).LastElm%) = GetElmOldSample(1).beamcurrent!
If GetElmOldSample(1).BeamSizeArray!(GetElmOldSample(1).LastElm%) = 0# Then GetElmOldSample(1).BeamSizeArray!(GetElmOldSample(1).LastElm%) = GetElmOldSample(1).beamsize!

' Load analytical conditions
Else
GetElmOldSample(1).TakeoffArray!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).takeoff!
GetElmOldSample(1).KilovoltsArray!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).kilovolts!
GetElmOldSample(1).BeamCurrentArray!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).beamcurrent!
GetElmOldSample(1).BeamSizeArray!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).beamsize!

GetElmOldSample(1).ColumnConditionMethodArray%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).ColumnConditionMethod%
GetElmOldSample(1).ColumnConditionStringArray$(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).ColumnConditionString$
End If

' Load std assignments and element disable flags
GetElmOldSample(1).StdAssignsFlag%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).StdAssignsFlag%(i%)
GetElmOldSample(1).DisableQuantFlag%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).DisableQuantFlag%(i%)
GetElmOldSample(1).DisableAcqFlag%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).DisableAcqFlag%(i%)

GetElmOldSample(1).VolatileFitCurvatures!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).VolatileFitCurvatures!(i%)
GetElmOldSample(1).PeakingBeforeAcquisitionElementFlags%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).PeakingBeforeAcquisitionElementFlags%(i%)
GetElmOldSample(1).VolatileFitTypes%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).VolatileFitTypes%(i%)

GetElmOldSample(1).WDSWaveScanHiPeaks!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).WDSWaveScanHiPeaks!(i%)
GetElmOldSample(1).WDSWaveScanLoPeaks!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).WDSWaveScanLoPeaks!(i%)
GetElmOldSample(1).WDSWaveScanPoints%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).WDSWaveScanPoints%(i%)

GetElmOldSample(1).WDSQuickScanHiPeaks!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).WDSQuickScanHiPeaks!(i%)
GetElmOldSample(1).WDSQuickScanLoPeaks!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).WDSQuickScanLoPeaks!(i%)
GetElmOldSample(1).WDSQuickScanSpeeds!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).WDSQuickScanSpeeds!(i%)

GetElmOldSample(1).NthPointAcquisitionFlags%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).NthPointAcquisitionFlags%(i%)
GetElmOldSample(1).NthPointAcquisitionIntervals%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).NthPointAcquisitionIntervals%(i%)

' Save multi-points
GetElmOldSample(1).MultiPointNumberofPointsAcquireHi%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).MultiPointNumberofPointsAcquireHi%(i%)
GetElmOldSample(1).MultiPointNumberofPointsAcquireLo%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).MultiPointNumberofPointsAcquireLo%(i%)
GetElmOldSample(1).MultiPointNumberofPointsIterateHi%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).MultiPointNumberofPointsIterateHi%(i%)
GetElmOldSample(1).MultiPointNumberofPointsIterateLo%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).MultiPointNumberofPointsIterateLo%(i%)
GetElmOldSample(1).MultiPointBackgroundFitType%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).MultiPointBackgroundFitType%(i%)     ' 0 = linear, 1 = 2nd order polynomial, 2 = exponential

For m% = 1 To MAXMULTI%
GetElmOldSample(1).MultiPointAcquirePositionsHi!(GetElmOldSample(1).LastElm%, m%) = GetElmTmpSample(1).MultiPointAcquirePositionsHi!(i%, m%)
GetElmOldSample(1).MultiPointAcquirePositionsLo!(GetElmOldSample(1).LastElm%, m%) = GetElmTmpSample(1).MultiPointAcquirePositionsLo!(i%, m%)

GetElmOldSample(1).MultiPointProcessLastManualFlagHi%(GetElmOldSample(1).LastElm%, m%) = GetElmTmpSample(1).MultiPointProcessLastManualFlagHi%(i%, m%)  ' last manual override flag (-1 = never use, 0 = automatic, 1 = always use)
GetElmOldSample(1).MultiPointProcessLastManualFlagLo%(GetElmOldSample(1).LastElm%, m%) = GetElmTmpSample(1).MultiPointProcessLastManualFlagLo%(i%, m%)  ' last manual override flag (-1 = never use, 0 = automatic, 1 = always use)
Next m%

GetElmOldSample(1).UnknownCountTimeForInterferenceStandardChanFlag(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).UnknownCountTimeForInterferenceStandardChanFlag(i%)

GetElmOldSample(1).SecondaryFluorescenceBoundaryFlag(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).SecondaryFluorescenceBoundaryFlag(i%)
GetElmOldSample(1).SecondaryFluorescenceBoundaryKratiosDATFile$(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFile$(i%)
GetElmOldSample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine1$(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine1$(i%)
GetElmOldSample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine2$(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine2$(i%)
GetElmOldSample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine3$(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine3$(i%)

GetElmOldSample(1).SecondaryFluorescenceBoundaryMatA_String$(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).SecondaryFluorescenceBoundaryMatA_String$(i%)
GetElmOldSample(1).SecondaryFluorescenceBoundaryMatB_String$(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).SecondaryFluorescenceBoundaryMatB_String$(i%)
GetElmOldSample(1).SecondaryFluorescenceBoundaryMatBStd_String$(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).SecondaryFluorescenceBoundaryMatBStd_String$(i%)

GetElmOldSample(1).ConditionNumbers%(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).ConditionNumbers%(i%)    ' list order

GetElmOldSample(1).AtomicCharges!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).AtomicCharges!(i%)
GetElmOldSample(1).AtomicWts!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).AtomicWts!(i%)

GetElmOldSample(1).EffectiveTakeOffs!(GetElmOldSample(1).LastElm%) = GetElmTmpSample(1).EffectiveTakeOffs!(i%)
2000:  Next n%

' Update number of analyzed elements
GetElmOldSample(1).LastChan% = GetElmOldSample(1).LastElm%

' Load the specified elements next, set x-ray, etc. to blank
For n% = 1 To MAXCHAN%
i% = n%

' Get specified element
sym$ = GetElmTmpSample(1).Elsyms$(i%)
ip% = IPOS1(MAXELM%, sym$, Symlo$())

sym$ = GetElmTmpSample(1).Xrsyms$(i%)
ipp% = IPOS1(MAXRAY%, sym$, Xraylo$())  ' including unanalyzed element

' Skip if element if IS analyzed
If ip% = 0 Or ipp% = 0 Or ipp% <= MAXRAY% - 1 Then GoTo 3000

' Check for *specified* element already analyzed or specified and skip if found
ippp% = IPOS5(Int(1), i%, GetElmTmpSample(), GetElmOldSample())
If ippp% > 0 Then GoTo 3000

' Increment number of specified elements
If GetElmOldSample(1).LastChan% + 1 > MAXCHAN% Then GoTo GetElmSaveTooManyElements
GetElmOldSample(1).LastChan% = GetElmOldSample(1).LastChan% + 1

' Load specified element parameters
GetElmOldSample(1).Elsyms$(GetElmOldSample(1).LastChan%) = GetElmTmpSample(1).Elsyms$(i%)
GetElmOldSample(1).Xrsyms$(GetElmOldSample(1).LastChan%) = vbNullString
    
' Make sure cations are loaded
If GetElmTmpSample(1).numcat%(i%) = 0 Then GetElmTmpSample(1).numcat%(i%) = AllCat%(ip%)
'If GetElmTmpSample(1).numoxd%(i%) = 0 Then GetElmTmpSample(1).numoxd%(i%) = AllOxd%(ip%)       ' zero oxygens is valid
GetElmOldSample(1).numcat%(GetElmOldSample(1).LastChan%) = GetElmTmpSample(1).numcat%(i%)
GetElmOldSample(1).numoxd%(GetElmOldSample(1).LastChan%) = GetElmTmpSample(1).numoxd%(i%)
GetElmOldSample(1).ElmPercents!(GetElmOldSample(1).LastChan%) = GetElmTmpSample(1).ElmPercents!(i%)

GetElmOldSample(1).AtomicCharges!(GetElmOldSample(1).LastChan%) = GetElmTmpSample(1).AtomicCharges!(i%)
GetElmOldSample(1).AtomicWts!(GetElmOldSample(1).LastChan%) = GetElmTmpSample(1).AtomicWts!(i%)
3000:  Next n%

' Re-sort condition orders just in case
Call Cond2ConditionDefaultOrder(GetElmOldSample())
If ierror Then Exit Sub

' Check assignments if in Probe for EPMA
Call GetElmCheckAssignments(GetElmOldSample())
If ierror Then Exit Sub

' Reload the GetElmTmpSample
GetElmTmpSample(1) = GetElmOldSample(1)

Exit Sub

' Errors
GetElmSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "GetElmSave"
ierror = True
Exit Sub

GetElmSaveTooManyElements:
msg$ = "Too many elements in sample"
MsgBox msg$, vbOKOnly + vbExclamation, "GetElmSave"
ierror = True
Exit Sub

End Sub

Sub GetElmCheckAssignments(sample() As TypeSample)
' Check the interference and stoichiometry assignments

ierror = False
On Error GoTo GetElmCheckAssignmentsError

Dim i As Integer, j As Integer
Dim ip As Integer, ipp As Integer

' Make sure that interference assignments do not reference specified or deleted elements
For i% = 1 To sample(1).LastElm%
For j% = 1 To MAXINTF%
If ProbeDataFileVersionNumber! > 6.41 Then
ip% = IPOS1A(sample(1).LastElm%, sample(1).StdAssignsIntfElements$(j%, i%), sample(1).StdAssignsIntfXrays$(j%, i%), sample(1).Elsyms$(), sample(1).Xrsyms$())
Else
ip% = IPOS1(sample(1).LastElm%, sample(1).StdAssignsIntfElements$(j%, i%), sample(1).Elsyms$())
End If
If ip% = 0 Then
sample(1).StdAssignsIntfElements$(j%, i%) = vbNullString
sample(1).StdAssignsIntfXrays$(j%, i%) = vbNullString
sample(1).StdAssignsIntfStds%(j%, i%) = 0
sample(1).StdAssignsIntfOrders%(j%, i%) = 0
End If
Next j%
Next i%

' Make sure that elements by difference, stoichiometry or formula basis are ok for new sample setup
ip% = IPOS1(sample(1).LastChan%, sample(1).FormulaElement$, sample(1).Elsyms$())
If ip% = 0 And sample(1).FormulaElement$ <> vbNullString And sample(1).Type% <> 1 Then               ' allow blank element for sum of cations (new code 06-16-2017), also allow missing formula element if standard since it will get loaded automatically (new code 04-04-2018)
sample(1).FormulaElementFlag% = False
sample(1).FormulaElement$ = vbNullString
sample(1).FormulaRatio! = 0#
sample(1).MineralFlag% = 0
End If

ip% = IPOS1(sample(1).LastChan%, sample(1).DifferenceElement$, sample(1).Elsyms$())
If ip% <= sample(1).LastElm% Then
sample(1).DifferenceElementFlag% = False
sample(1).DifferenceElement$ = vbNullString
End If

ip% = IPOS1(sample(1).LastChan%, sample(1).StoichiometryElement$, sample(1).Elsyms$())
If ip% <= sample(1).LastElm% Then
sample(1).StoichiometryElementFlag% = False
sample(1).StoichiometryElement$ = vbNullString
sample(1).StoichiometryRatio! = 0#
End If

ip% = IPOS1(sample(1).LastChan%, sample(1).RelativeElement$, sample(1).Elsyms$())
ipp% = IPOS1(sample(1).LastChan%, sample(1).RelativeToElement$, sample(1).Elsyms$())
If ip% <= sample(1).LastElm% Or ipp% = 0 Then
sample(1).RelativeElementFlag% = False
sample(1).RelativeElement$ = vbNullString
sample(1).RelativeToElement$ = vbNullString
sample(1).RelativeRatio! = 0#
End If

Exit Sub

' Errors
GetElmCheckAssignmentsError:
MsgBox Error$, vbOKOnly + vbCritical, "GetElmCheckAssignments"
ierror = True
Exit Sub

End Sub
