Attribute VB_Name = "CodeANALYZE4"
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

Global TimerToggle As Integer

Sub AnalyzeUpdateList(mode As Integer, tForm As Form, tList As ListBox)
' Updates the sample list box for FormANALYZE and FormPLOT, FormLOCATE and FormIMAGEEDIT
' mode = 0 select last item in list
' mode = 1 select previously selected items in list

ierror = False
On Error GoTo AnalyzeUpdateListError

Dim samplerow As Integer
Dim i As Integer

' Save current selections if indicated
If mode% = 1 And tList.ListCount > 0 Then
ReDim ListSelected(0 To tList.ListCount - 1) As Integer
For i% = 0 To tList.ListCount - 1
ListSelected%(i%) = tList.Selected(i%)
Next i%
End If

' Loop through samples and load appropriate samples
tList.Clear
For samplerow% = 1 To NumberofSamples%

' Load standards or unknowns or all
If tForm.OptionStandard.value = True And SampleTyps%(samplerow%) <> 1 Then GoTo 1000
If tForm.OptionUnknown.value = True And SampleTyps%(samplerow%) <> 2 Then GoTo 1000
If tForm.OptionWavescan.value = True And SampleTyps%(samplerow%) <> 3 Then GoTo 1000

' Check for FormANALYZE and if so, check display only samples with data checkbox
If tForm.Name = "FormANALYZE" Or tForm.Name = "FormPLOT_WAVE" Or tForm.Name = "FormPLOT" Then
If tForm.CheckOnlyDisplaySamplesWithData.value = vbChecked And SampleDels%(samplerow%) = True Then GoTo 1000
End If

' Load number set and name
msg$ = SampleGetString(samplerow%)
tList.AddItem msg$
tList.ItemData(tList.NewIndex) = samplerow%
1000:  Next samplerow%

' Set list box to last loaded sample
If tList.ListCount > 0 Then
If mode% = 0 Then
tList.ListIndex = tList.ListCount - 1
tList.Selected(tList.ListCount - 1) = True

' Select previously selected items
Else
If tList.ListCount > 0 Then
ReDim Preserve ListSelected(0 To tList.ListCount - 1) As Integer
End If
For i% = 0 To tList.ListCount - 1
tList.Selected(i%) = ListSelected%(i%)
Next i%
End If
End If

Exit Sub

' Errors
AnalyzeUpdateListError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeUpdateList"
ierror = True
Exit Sub

End Sub

Sub AnalyzeCombineSamples(mode As Integer, tmpsample() As TypeSample, sample() As TypeSample)
' Combine the samples (add tmpsample to sample)
' mode = 1 load analyzed elements only
' mode = 3 load specified elements only

ierror = False
On Error GoTo AnalyzeCombineSamplesError

Dim i As Integer, j As Integer
Dim ielm As Integer, m As Integer

' Load each analyzed elements
If mode% = 1 Then
For i% = 1 To tmpsample(1).LastElm%
If tmpsample(1).DisableAcqFlag%(i%) = 0 Then

' Find if element is already loaded as an analyzed element, and if so warn
ielm% = IPOS5(Int(0), i%, tmpsample(), sample())
If ielm% > 0 And tmpsample(1).DisableQuantFlag%(i%) = 0 Then
If sample(1).DisableQuantFlag%(ielm%) = 0 Then  ' if not disabled in first sample, warn user
msg$ = "Warning in AnalyzeCombineSamples: element " & tmpsample(1).Elsyms$(i%) & " is already present in the combined sample and will not be loaded again"
Call IOWriteLog(msg$)
GoTo 4000
End If
End If

' Increment number of analyzed elements
'If tmpsample(1).DisableQuantFlag%(i%) = 0 Then
If sample(1).LastElm% + 1 > MAXCHAN% Then GoTo AnalyzeCombineSamplesTooManyAnalyzedElements
sample(1).LastElm% = sample(1).LastElm% + 1

' Add conditions to array (the whole point being different voltages and/or currents!!!)
sample(1).TakeoffArray!(sample(1).LastElm%) = tmpsample(1).TakeoffArray!(i%)
sample(1).KilovoltsArray!(sample(1).LastElm%) = tmpsample(1).KilovoltsArray!(i%)
sample(1).BeamCurrentArray!(sample(1).LastElm%) = tmpsample(1).BeamCurrentArray!(i%)
sample(1).BeamSizeArray!(sample(1).LastElm%) = tmpsample(1).BeamSizeArray!(i%)

sample(1).ColumnConditionMethodArray%(sample(1).LastElm%) = tmpsample(1).ColumnConditionMethodArray%(i%)
sample(1).ColumnConditionStringArray$(sample(1).LastElm%) = tmpsample(1).ColumnConditionStringArray$(i%)

' Now add in normal element parameters
sample(1).Elsyms$(sample(1).LastElm%) = tmpsample(1).Elsyms$(i%)
sample(1).Xrsyms$(sample(1).LastElm%) = tmpsample(1).Xrsyms$(i%)
    
sample(1).numcat%(sample(1).LastElm%) = tmpsample(1).numcat%(i%)
sample(1).numoxd%(sample(1).LastElm%) = tmpsample(1).numoxd%(i%)
sample(1).ElmPercents!(sample(1).LastElm%) = tmpsample(1).ElmPercents!(i%)
sample(1).AtomicCharges!(sample(1).LastElm%) = tmpsample(1).AtomicCharges!(i%)

' 0=linear, 1=average, 2=high only, 3=low only, 4=exponential, 5=slope hi, 6=slope lo, 7=polynomial, 8=multi-point
sample(1).OffPeakCorrectionTypes%(sample(1).LastElm%) = tmpsample(1).OffPeakCorrectionTypes%(i%)

sample(1).BackgroundExponentialBase!(sample(1).LastElm%) = tmpsample(1).BackgroundExponentialBase!(i%)
sample(1).BackgroundSlopeCoefficients!(1, sample(1).LastElm%) = tmpsample(1).BackgroundSlopeCoefficients!(1, i%)
sample(1).BackgroundSlopeCoefficients!(2, sample(1).LastElm%) = tmpsample(1).BackgroundSlopeCoefficients!(2, i%)

sample(1).BackgroundPolynomialPositions!(1, sample(1).LastElm%) = tmpsample(1).BackgroundPolynomialPositions!(1, i%)
sample(1).BackgroundPolynomialPositions!(2, sample(1).LastElm%) = tmpsample(1).BackgroundPolynomialPositions!(2, i%)
sample(1).BackgroundPolynomialPositions!(3, sample(1).LastElm%) = tmpsample(1).BackgroundPolynomialPositions!(3, i%)

sample(1).BackgroundPolynomialCoefficients!(1, sample(1).LastElm%) = tmpsample(1).BackgroundPolynomialCoefficients!(1, i%)
sample(1).BackgroundPolynomialCoefficients!(2, sample(1).LastElm%) = tmpsample(1).BackgroundPolynomialCoefficients!(2, i%)
sample(1).BackgroundPolynomialCoefficients!(3, sample(1).LastElm%) = tmpsample(1).BackgroundPolynomialCoefficients!(3, i%)

sample(1).BackgroundPolynomialNominalBeam!(sample(1).LastElm%) = tmpsample(1).BackgroundPolynomialNominalBeam!(i%)

' Load other real time element parameters
sample(1).BraggOrders%(sample(1).LastElm%) = tmpsample(1).BraggOrders%(i%)
sample(1).MotorNumbers%(sample(1).LastElm%) = tmpsample(1).MotorNumbers%(i%)
sample(1).OrderNumbers%(sample(1).LastElm%) = tmpsample(1).OrderNumbers%(i%)
sample(1).CrystalNames$(sample(1).LastElm%) = tmpsample(1).CrystalNames$(i%)
sample(1).Crystal2ds!(sample(1).LastElm%) = tmpsample(1).Crystal2ds!(i%)
sample(1).CrystalKs!(sample(1).LastElm%) = tmpsample(1).CrystalKs!(i%)

sample(1).OnPeaks!(sample(1).LastElm%) = tmpsample(1).OnPeaks!(i%)
sample(1).HiPeaks!(sample(1).LastElm%) = tmpsample(1).HiPeaks!(i%)
sample(1).LoPeaks!(sample(1).LastElm%) = tmpsample(1).LoPeaks!(i%)

sample(1).StdAssigns%(sample(1).LastElm%) = tmpsample(1).StdAssigns%(i%)
sample(1).BackgroundTypes%(sample(1).LastElm%) = tmpsample(1).BackgroundTypes%(i%)  ' 0=off-peak, 1=MAN, 2=multipoint
sample(1).DeadTimes!(sample(1).LastElm%) = tmpsample(1).DeadTimes!(i%)

sample(1).Baselines!(sample(1).LastElm%) = tmpsample(1).Baselines!(i%)
sample(1).Windows!(sample(1).LastElm%) = tmpsample(1).Windows!(i%)
sample(1).Gains!(sample(1).LastElm%) = tmpsample(1).Gains!(i%)
sample(1).Biases!(sample(1).LastElm%) = tmpsample(1).Biases!(i%)
sample(1).InteDiffModes%(sample(1).LastElm%) = tmpsample(1).InteDiffModes%(i%)

sample(1).DetectorSlitSizes$(sample(1).LastElm%) = tmpsample(1).DetectorSlitSizes$(i%)
sample(1).DetectorSlitPositions$(sample(1).LastElm%) = tmpsample(1).DetectorSlitPositions$(i%)
sample(1).DetectorModes$(sample(1).LastElm%) = tmpsample(1).DetectorModes$(i%)

sample(1).IntegratedIntensitiesUseIntegratedFlags%(sample(1).LastElm%) = tmpsample(1).IntegratedIntensitiesUseIntegratedFlags%(i%)
sample(1).IntegratedIntensitiesInitialStepSizes!(sample(1).LastElm%) = tmpsample(1).IntegratedIntensitiesInitialStepSizes!(i%)
sample(1).IntegratedIntensitiesMinimumStepSizes!(sample(1).LastElm%) = tmpsample(1).IntegratedIntensitiesMinimumStepSizes!(i%)
sample(1).IntegratedIntensitiesIntegratedTypes%(sample(1).LastElm%) = tmpsample(1).IntegratedIntensitiesIntegratedTypes%(i%)

' Interference correction (ok for combining)
For j% = 1 To MAXINTF%
sample(1).StdAssignsIntfElements$(j%, sample(1).LastElm%) = tmpsample(1).StdAssignsIntfElements$(j%, i%)
sample(1).StdAssignsIntfXrays$(j%, sample(1).LastElm%) = tmpsample(1).StdAssignsIntfXrays$(j%, i%)
sample(1).StdAssignsIntfStds%(j%, sample(1).LastElm%) = tmpsample(1).StdAssignsIntfStds%(j%, i%)
sample(1).StdAssignsIntfOrders%(j%, sample(1).LastElm%) = tmpsample(1).StdAssignsIntfOrders%(j%, i%)
Next j%

' Volatile element correction
sample(1).VolatileCorrectionUnks%(sample(1).LastElm%) = tmpsample(1).VolatileCorrectionUnks%(i%)
sample(1).SpecifiedAreaPeakFactors!(sample(1).LastElm%) = tmpsample(1).SpecifiedAreaPeakFactors!(i%)

' Blank correction
sample(1).BlankCorrectionUnks%(sample(1).LastElm%) = tmpsample(1).BlankCorrectionUnks%(i%)
sample(1).BlankCorrectionLevels!(sample(1).LastElm%) = tmpsample(1).BlankCorrectionLevels!(i%)

sample(1).VolatileFitCurvatures!(sample(1).LastElm%) = tmpsample(1).VolatileFitCurvatures!(i%)
sample(1).PeakingBeforeAcquisitionElementFlags%(sample(1).LastElm%) = tmpsample(1).PeakingBeforeAcquisitionElementFlags%(i%)
sample(1).VolatileFitTypes%(sample(1).LastElm%) = tmpsample(1).VolatileFitTypes%(i%)

' MAN correction
sample(1).MANAbsCorFlags%(sample(1).LastElm%) = tmpsample(1).MANAbsCorFlags%(i%)
sample(1).MANLinearFitOrders%(sample(1).LastElm%) = tmpsample(1).MANLinearFitOrders%(i%)
For j% = 1 To MAXMAN%
sample(1).MANStdAssigns%(j%, sample(1).LastElm%) = tmpsample(1).MANStdAssigns%(j%, i%)
Next j%

' Count data and time
For j% = 1 To MAXROW%
sample(1).OnPeakCounts!(j%, sample(1).LastElm%) = tmpsample(1).OnPeakCounts!(j%, i%)
sample(1).HiPeakCounts!(j%, sample(1).LastElm%) = tmpsample(1).HiPeakCounts!(j%, i%)
sample(1).LoPeakCounts!(j%, sample(1).LastElm%) = tmpsample(1).LoPeakCounts!(j%, i%)
sample(1).OnCountTimes!(j%, sample(1).LastElm%) = tmpsample(1).OnCountTimes!(j%, i%)
sample(1).HiCountTimes!(j%, sample(1).LastElm%) = tmpsample(1).HiCountTimes!(j%, i%)
sample(1).LoCountTimes!(j%, sample(1).LastElm%) = tmpsample(1).LoCountTimes!(j%, i%)
sample(1).UnknownCountFactors!(j%, sample(1).LastElm%) = tmpsample(1).UnknownCountFactors!(j%, i%)
sample(1).UnknownMaxCounts&(j%, sample(1).LastElm%) = tmpsample(1).UnknownMaxCounts&(j%, i%)

sample(1).VolCountTimesStart(j%, sample(1).LastElm%) = tmpsample(1).VolCountTimesStart(j%, i%)
sample(1).VolCountTimesStop(j%, sample(1).LastElm%) = tmpsample(1).VolCountTimesStop(j%, i%)
sample(1).VolCountTimesDelay(j%, sample(1).LastElm%) = tmpsample(1).VolCountTimesDelay(j%, i%)

' Set line status to deleted if additional sample line is deleted
If Not tmpsample(1).LineStatus%(j%) Then sample(1).LineStatus%(j%) = False
Next j%

' Load different beam currents for each line for each line/element (combined analysis only)
For j% = 1 To MAXROW%
If Not tmpsample(1).CombinedConditionsFlag Then
sample(1).OnBeamCountsArray!(j%, sample(1).LastElm%) = tmpsample(1).OnBeamCounts!(j%)
sample(1).AbBeamCountsArray!(j%, sample(1).LastElm%) = tmpsample(1).AbBeamCounts!(j%)
sample(1).OnBeamCountsArray2!(j%, sample(1).LastElm%) = tmpsample(1).OnBeamCounts2!(j%)
sample(1).AbBeamCountsArray2!(j%, sample(1).LastElm%) = tmpsample(1).AbBeamCounts2!(j%)
Else
sample(1).OnBeamCountsArray!(j%, sample(1).LastElm%) = tmpsample(1).OnBeamCountsArray!(j%, i%)
sample(1).AbBeamCountsArray!(j%, sample(1).LastElm%) = tmpsample(1).AbBeamCountsArray!(j%, i%)
sample(1).OnBeamCountsArray2!(j%, sample(1).LastElm%) = tmpsample(1).OnBeamCountsArray2!(j%, i%)
sample(1).AbBeamCountsArray2!(j%, sample(1).LastElm%) = tmpsample(1).AbBeamCountsArray2!(j%, i%)
End If
Next j%

' Save last count times
sample(1).LastOnCountTimes!(sample(1).LastElm%) = tmpsample(1).LastOnCountTimes!(i%)
sample(1).LastHiCountTimes!(sample(1).LastElm%) = tmpsample(1).LastHiCountTimes!(i%)
sample(1).LastLoCountTimes!(sample(1).LastElm%) = tmpsample(1).LastLoCountTimes!(i%)
sample(1).LastWaveCountTimes!(sample(1).LastElm%) = tmpsample(1).LastWaveCountTimes!(i%)
sample(1).LastPeakCountTimes!(sample(1).LastElm%) = tmpsample(1).LastPeakCountTimes!(i%)
sample(1).LastQuickCountTimes!(sample(1).LastElm%) = tmpsample(1).LastQuickCountTimes!(i%)
sample(1).LastCountFactors!(sample(1).LastElm%) = tmpsample(1).LastCountFactors!(i%)
sample(1).LastMaxCounts&(sample(1).LastElm%) = tmpsample(1).LastMaxCounts&(i%)

sample(1).StdAssignsFlag%(sample(1).LastElm%) = tmpsample(1).StdAssignsFlag%(i%)
sample(1).DisableQuantFlag%(sample(1).LastElm%) = tmpsample(1).DisableQuantFlag%(i%)
sample(1).DisableAcqFlag%(sample(1).LastElm%) = tmpsample(1).DisableAcqFlag%(i%)

sample(1).WDSWaveScanHiPeaks!(sample(1).LastElm%) = tmpsample(1).WDSWaveScanHiPeaks!(i%)
sample(1).WDSWaveScanLoPeaks!(sample(1).LastElm%) = tmpsample(1).WDSWaveScanLoPeaks!(i%)
sample(1).WDSWaveScanPoints%(sample(1).LastElm%) = tmpsample(1).WDSWaveScanPoints%(i%)

sample(1).WDSQuickScanHiPeaks!(sample(1).LastElm%) = tmpsample(1).WDSQuickScanHiPeaks!(i%)
sample(1).WDSQuickScanLoPeaks!(sample(1).LastElm%) = tmpsample(1).WDSQuickScanLoPeaks!(i%)
sample(1).WDSQuickScanSpeeds!(sample(1).LastElm%) = tmpsample(1).WDSQuickScanSpeeds!(i%)

sample(1).NthPointAcquisitionFlags%(sample(1).LastElm%) = tmpsample(1).NthPointAcquisitionFlags%(i%)
sample(1).NthPointAcquisitionIntervals%(sample(1).LastElm%) = tmpsample(1).NthPointAcquisitionIntervals%(i%)

' Do not load integrated intensities here (first sample loaded must have the only integrated intensities when combining samples!)
If tmpsample(1).IntegratedIntensitiesUseIntegratedFlags(i%) Then
End If

' Load MPB parameters
sample(1).MultiPointNumberofPointsAcquireHi%(sample(1).LastElm%) = tmpsample(1).MultiPointNumberofPointsAcquireHi%(i%)
sample(1).MultiPointNumberofPointsAcquireLo%(sample(1).LastElm%) = tmpsample(1).MultiPointNumberofPointsAcquireLo%(i%)
sample(1).MultiPointNumberofPointsIterateHi%(sample(1).LastElm%) = tmpsample(1).MultiPointNumberofPointsIterateHi%(i%)
sample(1).MultiPointNumberofPointsIterateLo%(sample(1).LastElm%) = tmpsample(1).MultiPointNumberofPointsIterateLo%(i%)
sample(1).MultiPointBackgroundFitType%(sample(1).LastElm%) = tmpsample(1).MultiPointBackgroundFitType%(i%)      ' 0 = linear, 1 = 2nd order polynomial, 2 = exponential

' Update multi-point acquisition positions and count times only
For m% = 1 To MAXMULTI%
sample(1).MultiPointAcquirePositionsHi!(sample(1).LastElm%, m%) = tmpsample(1).MultiPointAcquirePositionsHi!(i%, m%)
sample(1).MultiPointAcquirePositionsLo!(sample(1).LastElm%, m%) = tmpsample(1).MultiPointAcquirePositionsLo!(i%, m%)
sample(1).MultiPointAcquireLastCountTimesHi(sample(1).LastElm%, m%) = tmpsample(1).MultiPointAcquireLastCountTimesHi(i%, m%)
sample(1).MultiPointAcquireLastCountTimesLo(sample(1).LastElm%, m%) = tmpsample(1).MultiPointAcquireLastCountTimesLo(i%, m%)
Next m%

sample(1).UnknownCountTimeForInterferenceStandardChanFlag(sample(1).LastElm%) = tmpsample(1).UnknownCountTimeForInterferenceStandardChanFlag(i%)

sample(1).SecondaryFluorescenceBoundaryFlag(sample(1).LastElm%) = tmpsample(1).SecondaryFluorescenceBoundaryFlag(i%)
sample(1).SecondaryFluorescenceBoundaryCorrectionMethod%(sample(1).LastElm%) = tmpsample(1).SecondaryFluorescenceBoundaryCorrectionMethod%(i%)

sample(1).SecondaryFluorescenceBoundaryKratiosDATFile$(sample(1).LastElm%) = tmpsample(1).SecondaryFluorescenceBoundaryKratiosDATFile$(i%)
sample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine1$(sample(1).LastElm%) = tmpsample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine1$(i%)
sample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine2$(sample(1).LastElm%) = tmpsample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine2$(i%)
sample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine3$(sample(1).LastElm%) = tmpsample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine3$(i%)

sample(1).SecondaryFluorescenceBoundaryMatA_String$(sample(1).LastElm%) = tmpsample(1).SecondaryFluorescenceBoundaryMatA_String$(i%)
sample(1).SecondaryFluorescenceBoundaryMatB_String$(sample(1).LastElm%) = tmpsample(1).SecondaryFluorescenceBoundaryMatB_String$(i%)
sample(1).SecondaryFluorescenceBoundaryMatBStd_String$(sample(1).LastElm%) = tmpsample(1).SecondaryFluorescenceBoundaryMatBStd_String$(i%)

'sample(1).SecondaryFluorescenceBoundaryMaterialB_Elsyms$(sample(1).LastElm%) = tmpsample(1).SecondaryFluorescenceBoundaryMaterialB_Elsyms$(i%)
'sample(1).SecondaryFluorescenceBoundaryMaterialB_WtPercents!(sample(1).LastElm%) = tmpsample(1).SecondaryFluorescenceBoundaryMaterialB_WtPercents!(i%)

sample(1).ConditionNumbers%(sample(1).LastElm%) = tmpsample(1).ConditionNumbers%(i%)    ' list order
'End If
End If
4000: Next i%

' Update number of analyzed elements
sample(1).LastChan% = sample(1).LastElm%
End If

' Load the specified elements last, set x-ray, etc. to blank
If mode% = 3 Then
For i% = tmpsample(1).LastElm% + 1 To tmpsample(1).LastChan%

' Find element, add if not already added
ielm% = IPOS5(Int(1), i%, tmpsample(), sample())
If ielm% > 0 Then GoTo 6000

' Increment number of specified elements
If sample(1).LastChan% + 1 > MAXCHAN% Then GoTo AnalyzeCombineSamplesTooManyElements
sample(1).LastChan% = sample(1).LastChan% + 1

' Load specified element parameters
sample(1).Elsyms$(sample(1).LastChan%) = tmpsample(1).Elsyms$(i%)
sample(1).Xrsyms$(sample(1).LastChan%) = vbNullString
    
' Make sure cations are loaded
sample(1).numcat%(sample(1).LastChan%) = tmpsample(1).numcat%(i%)
sample(1).numoxd%(sample(1).LastChan%) = tmpsample(1).numoxd%(i%)
sample(1).AtomicCharges!(sample(1).LastChan%) = tmpsample(1).AtomicCharges!(i%)
sample(1).ElmPercents!(sample(1).LastChan%) = tmpsample(1).ElmPercents!(i%)
6000: Next i%

' Since loading specified elements, load stoichiometry assinments
sample(1).FormulaElementFlag% = tmpsample(1).FormulaElementFlag%
sample(1).DifferenceElementFlag% = tmpsample(1).DifferenceElementFlag%
sample(1).DifferenceFormulaFlag% = tmpsample(1).DifferenceFormulaFlag%
sample(1).StoichiometryElementFlag% = tmpsample(1).StoichiometryElementFlag%
sample(1).RelativeElementFlag% = tmpsample(1).RelativeElementFlag%

sample(1).FormulaElement$ = tmpsample(1).FormulaElement$
sample(1).FormulaRatio! = tmpsample(1).FormulaRatio!
sample(1).MineralFlag% = tmpsample(1).MineralFlag%

sample(1).DifferenceElement$ = tmpsample(1).DifferenceElement$
sample(1).DifferenceFormula$ = tmpsample(1).DifferenceFormula$
sample(1).StoichiometryElement$ = tmpsample(1).StoichiometryElement$
sample(1).StoichiometryRatio! = tmpsample(1).StoichiometryRatio!
sample(1).StoichiometryRatio! = tmpsample(1).StoichiometryRatio!
sample(1).RelativeToElement$ = tmpsample(1).RelativeToElement$
sample(1).RelativeRatio! = tmpsample(1).RelativeRatio!
End If

' Check stoichiometry assignments (only call when mode%=3 to avoid problems with formula element assignments)
If mode% = 3 Then Call GetElmCheckAssignments(sample())
If ierror Then Exit Sub

' Set background type flags (AllMANBgdFlag = true if all MAN, MANBgdFlag = true if any MAN)
sample(1).AllMANBgdFlag = True
sample(1).MANBgdFlag = False
For i% = 1 To sample(1).LastElm%
If sample(1).BackgroundTypes%(i%) = 1 Then  ' 0=off-peak, 1=MAN, 2=multipoint
sample(1).MANBgdFlag = True
Else
sample(1).AllMANBgdFlag = False
End If
Next i%

Exit Sub

' Errors
AnalyzeCombineSamplesError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeCombineSamples"
ierror = True
Exit Sub

AnalyzeCombineSamplesTooManyAnalyzedElements:
msg$ = "Too many analyzed elements to combine into a single sample"
MsgBox msg$, vbOKOnly + vbExclamation, "AnalyzeCombineSamples"
ierror = True
Exit Sub

AnalyzeCombineSamplesTooManyElements:
msg$ = "Too many elements (analyzed + specified) to combine into a single sample"
MsgBox msg$, vbOKOnly + vbExclamation, "AnalyzeCombineSamples"
ierror = True
Exit Sub

End Sub

Sub AnalyzeCancel(tForm As Form)
' Cancel the analysis

ierror = False
On Error GoTo AnalyzeCancelError

tForm.StatusBarAnal.Panels(2).Bevel = sbrInset
FormPLOT.StatusBarPlot.Panels(2).Bevel = sbrInset
Call MiscTimer(CSng(0.2))
tForm.StatusBarAnal.Panels(2).Bevel = sbrRaised
FormPLOT.StatusBarPlot.Panels(2).Bevel = sbrRaised

icancelanal = True
DoEvents

Exit Sub

' Errors
AnalyzeCancelError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeCancel"
ierror = True
Exit Sub

End Sub

Sub AnalyzeNext(mode As Integer)
' Analyze the next sample
'  mode = 0 disable the button
'  mode = 1 enable the button
'  mode = 2 set the next flag

ierror = False
On Error GoTo AnalyzeNextError

FormANALYZE.StatusBarAnal.Panels(3).Bevel = sbrInset
Call MiscTimer(CSng(0.2))
FormANALYZE.StatusBarAnal.Panels(3).Bevel = sbrRaised

' Check mode
If mode% = 0 Then
FormANALYZE.StatusBarAnal.Panels(3).Enabled = False
FormANALYZE.StatusBarAnal.Panels(3).Picture = Nothing
ElseIf mode% = 1 Then
FormANALYZE.StatusBarAnal.Panels(3).Enabled = True
Else
NextSample = True
End If

DoEvents
Exit Sub

' Errors
AnalyzeNextError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeNext"
ierror = True
Exit Sub

End Sub

Sub AnalyzeStatusAnal(astring As String)
' This routine writes the string to the status panel for analysis status purposes

ierror = False
On Error GoTo AnalyzeStatusAnalError

' Update status bar
FormANALYZE.StatusBarAnal.Panels(1).Text = astring$
DoEvents
FormPLOT.StatusBarPlot.Panels(1).Text = astring$
DoEvents

If Not RealTimeMode% Then
FormMAIN.StatusBarAuto.Panels(1).Text = astring$
DoEvents
End If

Exit Sub

' Errors
AnalyzeStatusAnalError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeStatusAnal"
ierror = True
Exit Sub

End Sub

Sub AnalyzeNextToggle()
' Toggle the Next button for attention

ierror = False
On Error GoTo AnalyzeNextToggleError

If TimerToggle Then
FormANALYZE.StatusBarAnal.Panels(3).Picture = FormANALYZE.CommandNext.Picture
Else
FormANALYZE.StatusBarAnal.Panels(3).Picture = Nothing
End If
TimerToggle = Not TimerToggle
DoEvents

Exit Sub

' Errors
AnalyzeNextToggleError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeNextToggle"
ierror = True
Exit Sub

End Sub

Sub AnalyzeSelectSamples(tList As ListBox)
' Select the selected samples in the list (include "continued" samples)

ierror = False
On Error GoTo AnalyzeSelectSamplesError

Dim i As Integer, n As Integer

' Save original selected sample
n% = tList.ListIndex

'tList.Enabled = False      ' commented out to avoid problem in FormPLOT_WAVE loading X and Y lists
'DoEvents

' Select all sample after selected sample
For i% = n% + 1 To tList.ListCount - 1
If InStr(tList.List(i%), CONTINUED$) Then
tList.Selected(i%) = True
Else
Exit For
End If
Next i%

' Select all samples backwards
If InStr(tList.List(n%), CONTINUED$) Then
For i% = n% - 1 To 0 Step -1
If InStr(tList.List(i%), CONTINUED$) Then
tList.Selected(i%) = True
Else
Exit For
End If
Next i%
tList.Selected(i%) = True
End If

'tList.Enabled = True      ' commented out to avoid problem in FormPLOT_WAVE loading X and Y lists
'DoEvents

Exit Sub

' Errors
AnalyzeSelectSamplesError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeSelectSamples"
tList.Enabled = True
ierror = True
Exit Sub

End Sub
