Attribute VB_Name = "CodeINIT2"
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

Sub InitElement(chan As Integer, sample() As TypeSample)
' Routine to initialize a single element chan

ierror = False
On Error GoTo InitElementError

Dim j As Integer, m As Integer

' Check for valid chan
If chan% < 1 Or chan% > MAXCHAN% Then GoTo InitElementBadchan

sample(1).TakeoffArray!(chan%) = 0#
sample(1).KilovoltsArray!(chan%) = 0#
sample(1).BeamCurrentArray!(chan%) = 0#
sample(1).BeamSizeArray!(chan%) = 0#
sample(1).ColumnConditionMethodArray%(chan%) = 0
sample(1).ColumnConditionStringArray$(chan%) = vbNullString

' Initialize arrays for a single element
sample(1).Elsyms$(chan%) = vbNullString
sample(1).Xrsyms$(chan%) = vbNullString
sample(1).BraggOrders%(chan%) = 1           ' assume first order as default
sample(1).numcat%(chan%) = 0
sample(1).numoxd%(chan%) = 0
sample(1).AtomicCharges!(chan%) = 0#
sample(1).ElmPercents!(chan%) = 0#

' 0=linear, 1=average, 2=high only, 3=low only, 4=exponential, 5=slope hi, 6=slope lo, 7=polynomial, 8=multi-point
sample(1).OffPeakCorrectionTypes%(chan%) = 0

sample(1).MotorNumbers%(chan%) = 0
sample(1).OrderNumbers%(chan%) = 0
sample(1).CrystalNames$(chan%) = vbNullString
sample(1).Crystal2ds!(chan%) = 0#
sample(1).CrystalKs!(chan%) = 0#

sample(1).OnPeaks!(chan%) = 0#
sample(1).HiPeaks!(chan%) = 0#
sample(1).LoPeaks!(chan%) = 0#

sample(1).StdAssigns%(chan%) = 0
sample(1).BackgroundTypes%(chan%) = 0  ' 0=off-peak, 1=MAN, 2=multipoint
sample(1).DeadTimes!(chan%) = 0#

sample(1).Baselines!(chan%) = 0#
sample(1).Windows!(chan%) = 0#
sample(1).Gains!(chan%) = 0#
sample(1).Biases!(chan%) = 0#
sample(1).InteDiffModes%(chan%) = False ' False = integral / True = differential

sample(1).AtomicCharges!(chan%) = 0#

sample(1).DetectorSlitSizes$(chan%) = vbNullString
sample(1).DetectorSlitPositions$(chan%) = vbNullString
sample(1).DetectorModes$(chan%) = vbNullString

sample(1).IntegratedIntensitiesUseIntegratedFlags%(chan%) = False
sample(1).IntegratedIntensitiesInitialStepSizes!(chan%) = 0#
sample(1).IntegratedIntensitiesMinimumStepSizes!(chan%) = 0#
sample(1).IntegratedIntensitiesIntegratedTypes%(chan%) = 0

' Interference correction
For j% = 1 To MAXINTF%
sample(1).StdAssignsIntfElements$(j%, chan%) = vbNullString
sample(1).StdAssignsIntfXrays$(j%, chan%) = vbNullString
sample(1).StdAssignsIntfStds%(j%, chan%) = 0
sample(1).StdAssignsIntfOrders%(j%, chan%) = 0
Next j%

' Volatile element correction and specified area peak factors
sample(1).VolatileCorrectionUnks%(chan%) = 0        ' no assignment by default (-1 = self, 0 = none, >0 = assigned)
sample(1).SpecifiedAreaPeakFactors!(chan%) = 1#     ' no correction by default

' Blank correction
sample(1).BlankCorrectionUnks%(chan%) = 0        ' no assignment by default
sample(1).BlankCorrectionLevels!(chan%) = 0#

' Peaking before acquisition flags
sample(1).PeakingBeforeAcquisitionElementFlags%(chan%) = False

' MAN correction
sample(1).MANAbsCorFlags%(chan%) = 0
sample(1).MANLinearFitOrders%(chan%) = 0
For j% = 1 To MAXMAN%
sample(1).MANStdAssigns%(j%, chan%) = 0
Next j%

' Slope and polynomial coefficients
sample(1).BackgroundExponentialBase!(chan%) = 1#

sample(1).BackgroundSlopeCoefficients!(1, chan%) = 1#
sample(1).BackgroundSlopeCoefficients!(2, chan%) = 1#

sample(1).BackgroundPolynomialPositions!(1, chan%) = 0#
sample(1).BackgroundPolynomialPositions!(2, chan%) = 0#
sample(1).BackgroundPolynomialPositions!(3, chan%) = 0#

sample(1).BackgroundPolynomialCoefficients!(1, chan%) = 0#
sample(1).BackgroundPolynomialCoefficients!(2, chan%) = 0#
sample(1).BackgroundPolynomialCoefficients!(3, chan%) = 0#

sample(1).BackgroundPolynomialNominalBeam!(chan%) = 0#

sample(1).LastOnCountTimes!(chan%) = DefaultOnCountTime!
sample(1).LastHiCountTimes!(chan%) = DefaultOffCountTime!
sample(1).LastLoCountTimes!(chan%) = DefaultOffCountTime!
sample(1).LastWaveCountTimes!(chan%) = DefaultWavescanCountTime!
sample(1).LastPeakCountTimes!(chan%) = DefaultPeakingCountTime!
sample(1).LastQuickCountTimes!(chan%) = DefaultQuickscanCountTime!
sample(1).LastCountFactors!(chan%) = 1#
sample(1).LastMaxCounts&(chan%) = DefaultUnknownMaxCounts&

' Other
sample(1).Offsets!(chan%) = 0#
sample(1).LineEnergy!(chan%) = 0#   ' in eV
sample(1).LineEdge!(chan%) = 0#     ' in eV
sample(1).AtomicWts!(chan%) = 0#
sample(1).AtomicNums%(chan%) = 0
sample(1).XrayNums%(chan%) = 0
sample(1).Elsyup$(chan%) = vbNullString
sample(1).Oxsyup$(chan%) = vbNullString

' New
sample(1).StdAssignsFlag%(chan%) = 0
sample(1).DisableQuantFlag%(chan%) = 0
sample(1).DisableAcqFlag%(chan%) = 0

' 05/23/2009
sample(1).WDSWaveScanHiPeaks!(chan%) = 0#
sample(1).WDSWaveScanLoPeaks!(chan%) = 0#
sample(1).WDSWaveScanPoints%(chan%) = 0

sample(1).WDSQuickScanHiPeaks!(chan%) = 0#
sample(1).WDSQuickScanLoPeaks!(chan%) = 0#
sample(1).WDSQuickScanSpeeds!(chan%) = 0#

' 09/05/2009
sample(1).NthPointAcquisitionFlags%(chan%) = 0
sample(1).NthPointAcquisitionIntervals%(chan%) = DefaultNthPointAcquisitionInterval%

' 11/01/2009
sample(1).MultiPointNumberofPointsAcquireHi%(chan%) = DefaultMultiPointNumberofPointsAcquireHi%
sample(1).MultiPointNumberofPointsAcquireLo%(chan%) = DefaultMultiPointNumberofPointsAcquireLo%
sample(1).MultiPointNumberofPointsIterateHi%(chan%) = DefaultMultiPointNumberofPointsIterateHi%
sample(1).MultiPointNumberofPointsIterateLo%(chan%) = DefaultMultiPointNumberofPointsIterateLo%
sample(1).MultiPointBackgroundFitType%(chan%) = 0    ' 0 = linear, 1 = 2nd order polynomial, 2 = exponential

' 11/28/2009
For m% = 1 To MAXMULTI%
sample(1).MultiPointAcquirePositionsHi!(chan%, m%) = 0#
sample(1).MultiPointAcquirePositionsLo!(chan%, m%) = 0#
sample(1).MultiPointAcquireLastCountTimesHi!(chan%, m%) = sample(1).LastHiCountTimes!(chan%) / 2#   ' divide off-peak count time by two for multi-point
sample(1).MultiPointAcquireLastCountTimesLo!(chan%, m%) = sample(1).LastLoCountTimes!(chan%) / 2#   ' divide off-peak count time by two for multi-point
Next m%

' 12/17/2011
sample(1).UnknownCountTimeForInterferenceStandardChanFlag(chan%) = False

' 11/3/2012
sample(1).SecondaryFluorescenceBoundaryCorrectionMethod%(chan%) = 0

sample(1).SecondaryFluorescenceBoundaryFlag(chan%) = False
sample(1).SecondaryFluorescenceBoundaryKratiosDATFile$(chan%) = vbNullString
sample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine1$(chan%) = vbNullString
sample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine2$(chan%) = vbNullString
sample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine3$(chan%) = vbNullString

sample(1).SecondaryFluorescenceBoundaryMatA_String$(chan%) = vbNullString
sample(1).SecondaryFluorescenceBoundaryMatB_String$(chan%) = vbNullString
sample(1).SecondaryFluorescenceBoundaryMatBStd_String$(chan%) = vbNullString

'sample(1).SecondaryFluorescenceBoundaryMaterialB_Elsyms$(chan%) = vbNullString
'sample(1).SecondaryFluorescenceBoundaryMaterialB_WtPercents!(chan%) = 0#

sample(1).ConditionNumbers%(chan%) = 1

' Count data and time
For j% = 1 To MAXROW%
sample(1).OnPeakCounts!(j%, chan%) = 0#
sample(1).HiPeakCounts!(j%, chan%) = 0#
sample(1).LoPeakCounts!(j%, chan%) = 0#
sample(1).OnCountTimes!(j%, chan%) = DefaultOnCountTime!
sample(1).HiCountTimes!(j%, chan%) = DefaultOffCountTime!
sample(1).LoCountTimes!(j%, chan%) = DefaultOffCountTime!
sample(1).UnknownCountFactors!(j%, chan%) = 1#
sample(1).UnknownMaxCounts&(j%, chan%) = DefaultUnknownMaxCounts
sample(1).VolCountTimesStart(j%, chan%) = 0#    ' variants must be initialized
sample(1).VolCountTimesStop(j%, chan%) = 0#    ' variants must be initialized

sample(1).OnBeamCountsArray!(j%, chan%) = 0#
sample(1).AbBeamCountsArray!(j%, chan%) = 0#
sample(1).OnBeamCountsArray2!(j%, chan%) = 0#
sample(1).AbBeamCountsArray2!(j%, chan%) = 0#
Next j%

Exit Sub

' Errors
InitElementError:
MsgBox Error$, vbOKOnly + vbCritical, "InitElement"
ierror = True
Exit Sub

InitElementBadchan:
msg$ = "Invalid channel number"
MsgBox msg$, vbOKOnly + vbExclamation, "InitElement"
ierror = True
Exit Sub

End Sub

Sub InitLine(analysis As TypeAnalysis)
' Initializes arrays for a single data line calculation (see InitStandards for basic initialization)

ierror = False
On Error GoTo InitLineError

Dim i As Integer

analysis.TotalPercent! = 0#
analysis.totaloxygen! = 0#
analysis.TotalCations! = 0#
analysis.totalatoms! = 0#
analysis.CalculatedOxygen! = 0#
analysis.ExcessOxygen! = 0#
analysis.zbar! = 0#
analysis.AtomicWeight! = 0#
analysis.OxygenFromHalogens! = 0#
analysis.HalogenCorrectedOxygen! = 0#
analysis.ChargeBalance! = 0#
analysis.FeCharge! = 0#

For i% = 1 To MAXCHAN%
analysis.UnkZAFCors!(1, i%) = 1#
analysis.UnkZAFCors!(2, i%) = 1#
analysis.UnkZAFCors!(3, i%) = 1#
analysis.UnkZAFCors!(4, i%) = 1#
analysis.UnkZAFCors!(5, i%) = 1#
analysis.UnkZAFCors!(6, i%) = 1#
analysis.UnkZAFCors!(7, i%) = 1#
analysis.UnkZAFCors!(8, i%) = 1#

analysis.UnkKrats(i%) = 0#
analysis.UnkBetas(i%) = 1#   ' alpha factor calculations only
analysis.UnkMACs(i%) = 0#

analysis.WtPercents!(i%) = 0#
analysis.OxPercents!(i%) = 0#
analysis.AtPercents!(i%) = 0#
analysis.OxMolPercents!(i%) = 0#
analysis.ElPercents!(i%) = 0#
analysis.Formulas!(i%) = 0#
analysis.NormElPercents!(i%) = 0#
analysis.NormOxPercents!(i%) = 0#
Next i%

Exit Sub

' Errors
InitLineError:
MsgBox Error$, vbOKOnly + vbCritical, "InitLine"
ierror = True
Exit Sub

End Sub

Sub InitSample(sample() As TypeSample)
' Routine to initialize a sample array

ierror = False
On Error GoTo InitSampleError

Dim i As Integer, j As Integer, k As Integer
Dim amsg As String

amsg$ = "Dimensioning raw data sample arrays..."
ReDim sample(1).OnPeakCounts(1 To MAXROW%, 1 To MAXCHAN%) As Single ' data arrays
ReDim sample(1).HiPeakCounts(1 To MAXROW%, 1 To MAXCHAN%) As Single
ReDim sample(1).LoPeakCounts(1 To MAXROW%, 1 To MAXCHAN%) As Single

ReDim sample(1).OnCountTimes(1 To MAXROW%, 1 To MAXCHAN%) As Single ' data arrays
ReDim sample(1).HiCountTimes(1 To MAXROW%, 1 To MAXCHAN%) As Single
ReDim sample(1).LoCountTimes(1 To MAXROW%, 1 To MAXCHAN%) As Single

' Dimension corrected data arrays (allocated in DataGetMDBSample to get around 64K limit for user defined type)
'amsg$ = "Dimensioning corrected data sample arrays..."
'ReDim sample(1).CorData(1 To MAXROW%, 1 To MAXCHAN1%) As Single
'ReDim sample(1).BgdData(1 To MAXROW%, 1 To MAXCHAN1%) As Single
'ReDim sample(1).ErrData(1 To MAXROW%, 1 To MAXCHAN1%) As Single
    
'ReDim sample(1).OnTimeData(1 To MAXROW%, 1 To MAXCHAN1%) As Single      ' for aggregate intensity calculations
'ReDim sample(1).HiTimeData(1 To MAXROW%, 1 To MAXCHAN1%) As Single      ' for aggregate intensity calculations
'ReDim sample(1).LoTimeData(1 To MAXROW%, 1 To MAXCHAN1%) As Single      ' for aggregate intensity calculations

'ReDim sample(1).OnBeamData(1 To MAXROW%, 1 To MAXCHAN1%) As Single      ' for aggregate intensity calculations (average aggregate beam)
'ReDim sample(1).OnBeamDataArray(1 To MAXROW%, 1 To MAXCHAN1%) As Single      ' for aggregate intensity calculations (average aggregate beam)
'ReDim sample(1).AggregateNumChannels(1 To MAXROW%, 1 To MAXCHAN1%) As Integer      ' for aggregate intensity calculations (number of aggregate channels)

amsg$ = "Dimensioning TDI sample arrays..."
ReDim sample(1).VolCountTimesStart(1 To MAXROW%, 1 To MAXCHAN%) As Variant  ' volatile time arrays
ReDim sample(1).VolCountTimesStop(1 To MAXROW%, 1 To MAXCHAN%) As Variant
ReDim sample(1).VolCountTimesDelay(1 To MAXROW%, 1 To MAXCHAN%) As Single

amsg$ = "Dimensioning beam current sample arrays..."
ReDim sample(1).OnBeamCountsArray(1 To MAXROW%, 1 To MAXCHAN%) As Single    ' multiple conditions
ReDim sample(1).AbBeamCountsArray(1 To MAXROW%, 1 To MAXCHAN%) As Single    ' multiple conditions
ReDim sample(1).OnBeamCountsArray2(1 To MAXROW%, 1 To MAXCHAN%) As Single    ' multiple conditions
ReDim sample(1).AbBeamCountsArray2(1 To MAXROW%, 1 To MAXCHAN%) As Single    ' multiple conditions

ReDim sample(1).UnknownCountFactors(1 To MAXROW%, 1 To MAXCHAN%) As Single
ReDim sample(1).UnknownMaxCounts(1 To MAXROW%, 1 To MAXCHAN%) As Long

amsg$ = "Dimensioning MAN sample arrays..."
ReDim sample(1).MANStdAssigns(1 To MAXMAN%, 1 To MAXCHAN%) As Integer
ReDim sample(1).MANLinearFitOrders(1 To MAXCHAN%) As Integer
ReDim sample(1).MANAbsCorFlags(1 To MAXCHAN%) As Integer                               ' MAN matrix correction flag

amsg$ = "Dimensioning interference sample arrays..."
ReDim sample(1).StdAssignsIntfElements(1 To MAXINTF%, 1 To MAXCHAN%) As String         ' interfering element
ReDim sample(1).StdAssignsIntfXrays(1 To MAXINTF%, 1 To MAXCHAN%) As String            ' interfering x-ray (for channel ID only)
ReDim sample(1).StdAssignsIntfStds(1 To MAXINTF%, 1 To MAXCHAN%) As Integer            ' interference standard
ReDim sample(1).StdAssignsIntfOrders(1 To MAXINTF%, 1 To MAXCHAN%) As Integer          ' order of interfering line (for matrix correction adjustment)

ReDim sample(1).StagePositions(1 To MAXROW%, 1 To MAXAXES%) As Single

' Integrated intensity arrays (allocated in DataGetMDBSample to avoid 64K limit)
'amsg$ = "Dimensioning integrated intensity sample arrays..."
'ReDim sample(1).IntegratedPoints(1 To MAXROW%, 1 To MAXCHAN%) As Integer
'ReDim sample(1).IntegratedPeakIntensities(1 To MAXROW%, 1 To MAXCHAN%) As Single
'ReDim sample(1).IntegratedPositions(1 To MAXROW%, 1 To MAXCHAN%, 1 To n%) As Single
'ReDim sample(1).IntegratedIntensities(1 To MAXROW%, 1 To MAXCHAN%, 1 To n%) As Single
'ReDim sample(1).IntegratedCountTimes(1 To MAXROW%, 1 To MAXCHAN%, 1 To n%) As Single

amsg$ = "Dimensioning background fit sample arrays..."
ReDim sample(1).BackgroundExponentialBase(1 To MAXCHAN%) As Single
ReDim sample(1).BackgroundSlopeCoefficients(1 To 2, 1 To MAXCHAN%) As Single    ' 1 = High, 2 = Low
ReDim sample(1).BackgroundPolynomialPositions(1 To MAXCOEFF%, 1 To MAXCHAN%) As Single
ReDim sample(1).BackgroundPolynomialCoefficients(1 To MAXCOEFF%, 1 To MAXCHAN%) As Single
ReDim sample(1).BackgroundPolynomialNominalBeam(1 To MAXCHAN%) As Single

' Multi-point background positions (same for all lines in sample)
amsg$ = "Dimensioning MPB sample arrays..."
ReDim sample(1).MultiPointAcquirePositionsHi(1 To MAXCHAN%, 1 To MAXMULTI%) As Single
ReDim sample(1).MultiPointAcquirePositionsLo(1 To MAXCHAN%, 1 To MAXMULTI%) As Single

ReDim sample(1).MultiPointAcquireLastCountTimesHi(1 To MAXCHAN%, 1 To MAXMULTI%) As Single
ReDim sample(1).MultiPointAcquireLastCountTimesLo(1 To MAXCHAN%, 1 To MAXMULTI%) As Single

' Multi-point acquisition data
ReDim sample(1).MultiPointAcquireCountTimesHi(1 To MAXROW%, 1 To MAXCHAN%, 1 To MAXMULTI%) As Single  ' high off-peak time
ReDim sample(1).MultiPointAcquireCountTimesLo(1 To MAXROW%, 1 To MAXCHAN%, 1 To MAXMULTI%) As Single  ' low off-peak time
    
ReDim sample(1).MultiPointAcquireCountsHi(1 To MAXROW%, 1 To MAXCHAN%, 1 To MAXMULTI%) As Single  ' x-ray counts (cps raw counts)
ReDim sample(1).MultiPointAcquireCountsLo(1 To MAXROW%, 1 To MAXCHAN%, 1 To MAXMULTI%) As Single  ' x-ray counts (cps raw counts)
    
ReDim sample(1).MultiPointProcessManualFlagHi(1 To MAXROW%, 1 To MAXCHAN%, 1 To MAXMULTI%) As Integer  ' manual override flag (-1 = never use, 0 = automatic, 1 = always use)
ReDim sample(1).MultiPointProcessManualFlagLo(1 To MAXROW%, 1 To MAXCHAN%, 1 To MAXMULTI%) As Integer  ' manual override flag (-1 = never use, 0 = automatic, 1 = always use)

ReDim sample(1).MultiPointProcessLastManualFlagHi(1 To MAXCHAN%, 1 To MAXMULTI%) As Integer  ' last manual override flag (-1 = never use, 0 = automatic, 1 = always use)
ReDim sample(1).MultiPointProcessLastManualFlagLo(1 To MAXCHAN%, 1 To MAXMULTI%) As Integer  ' last manual override flag (-1 = never use, 0 = automatic, 1 = always use)

' Init boundary secondary fluorescence arrays
amsg$ = "Dimensioning secondary fluorescence sample arrays..."
ReDim sample(1).SecondaryFluorescenceBoundaryDistance(1 To MAXROW%) As Single

ReDim sample(1).SecondaryFluorescenceBoundaryKratiosDATFile(1 To MAXCHAN%) As String
ReDim sample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine1(1 To MAXCHAN%) As String
ReDim sample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine2(1 To MAXCHAN%) As String
ReDim sample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine3(1 To MAXCHAN%) As String

ReDim sample(1).SecondaryFluorescenceBoundaryKratios(1 To MAXROW%, 1 To MAXCHAN%) As Single

' EDS spectra elements
amsg$ = "Dimensioning EDS sample arrays..."
sample(1).EDSSpectraFlag% = False
sample(1).EDSSpectraUseFlag% = False

ReDim sample(1).EDSUnknownCountFactors(1 To MAXROW%) As Single
  
' Allocated in DataEDSSpectraGetData to avoid 64K limit
'ReDim sample(1).EDSSpectraIntensities(1 To MAXROW%, 1 To MAXSPECTRA%) As Long
'ReDim sample(1).EDSSpectraStrobes(1 To MAXROW%, 1 To MAXSTROBE%) As Long    ' only for Oxford EDS

ReDim sample(1).EDSSpectraSampleTime(1 To MAXROW%) As Single     ' sample counting time (estimated)
ReDim sample(1).EDSSpectraElapsedTime(1 To MAXROW%) As Single    ' real time
ReDim sample(1).EDSSpectraDeadTime(1 To MAXROW%) As Single       ' dead time (in percentage)
ReDim sample(1).EDSSpectraLiveTime(1 To MAXROW%) As Single       ' live time (actual count integration time)

ReDim sample(1).EDSSpectraNumberofChannels(1 To MAXROW%) As Integer
ReDim sample(1).EDSSpectraNumberofStrobes(1 To MAXROW%) As Integer  ' (new 11/06/04)

ReDim sample(1).EDSSpectraEVPerChannel(1 To MAXROW%) As Single
ReDim sample(1).EDSSpectraTakeOff(1 To MAXROW%) As Single
ReDim sample(1).EDSSpectraAcceleratingVoltage(1 To MAXROW%) As Single

ReDim sample(1).EDSSpectraMaxCounts(1 To MAXROW%) As Long
ReDim sample(1).EDSSpectraStartEnergy(1 To MAXROW%) As Single
ReDim sample(1).EDSSpectraEndEnergy(1 To MAXROW%) As Single
ReDim sample(1).EDSSpectraADCTimeConstant(1 To MAXROW%) As Single

ReDim sample(1).EDSSpectraKLineBCoefficient(1 To MAXROW%) As Single         ' used by Bruker only
ReDim sample(1).EDSSpectraKLineCCoefficient(1 To MAXROW%) As Single         ' used by Bruker only

ReDim sample(1).EDSSpectraEDSFileName(1 To MAXROW%) As String               ' used by JEOL OEM EDS only

amsg$ = "Dimensioning CL sample arrays..."

' Allocated in DataCLSpectraGetData to avoid 64K limit
'ReDim sample(1).CLSpectraIntensities(1 To MAXROW%, 1 To MAXSPECTRA_CL%) As Long
'ReDim sample(1).CLSpectraDarkIntensities(1 To MAXROW%, 1 To MAXSPECTRA_CL%) As Long
'ReDim sample(1).CLSpectraNanometers(1 To MAXROW%, 1 To MAXSPECTRA_CL%) As Long

ReDim sample(1).CLSpectraNumberofChannels(1 To MAXROW%) As Integer
ReDim sample(1).CLSpectraStartEnergy(1 To MAXROW%) As Single
ReDim sample(1).CLSpectraEndEnergy(1 To MAXROW%) As Single
ReDim sample(1).CLSpectraKilovolts(1 To MAXROW%) As Single

ReDim sample(1).CLAcquisitionCountTime(1 To MAXROW%) As Single
ReDim sample(1).CLUnknownCountFactors(1 To MAXROW%) As Single
ReDim sample(1).CLDarkSpectraCountTimeFraction(1 To MAXROW%) As Single

' Calculated, not stored in data file
amsg$ = "Initializing sample parameters..."
sample(1).LastElm% = 0
sample(1).LastChan% = 0
sample(1).AllMANBgdFlag = False
sample(1).MANBgdFlag = False

sample(1).Datarows% = 0
sample(1).GoodDataRows% = 0

' Initialize sample row array variables
sample(1).number% = 0
sample(1).Set% = 0
sample(1).Type% = 0
sample(1).Name$ = vbNullString

sample(1).CombinedConditionsFlag = False
sample(1).MultipleSetupNumber% = 0

' Initialize sample arrays
sample(1).SampleSetupNumber% = 0  ' sample setup number
sample(1).FileSetupName$ = vbNullString  ' file setup name
sample(1).FileSetupNumber% = 0  ' file setup number

sample(1).VolatileAcquisitionType% = 0  ' 0 = none, 1 = self, 2 = assigned
sample(1).WavescanAcquisitionType% = 0  ' 1 = normal, 2 = quick, 3 = normal ROM, 4 = quick ROM

sample(1).kilovolts! = 0#
sample(1).takeoff! = 0#
sample(1).beamcurrent! = 0#
sample(1).beamsize! = NOT_ANALYZED_VALUE_SINGLE!     ' since it can be zero
sample(1).ColumnConditionMethod% = 0
sample(1).ColumnConditionString$ = vbNullString
sample(1).Magnification! = 0#
sample(1).magnificationanalytical! = 0#
sample(1).magnificationimaging! = 0#
sample(1).ImageShiftX! = 0#
sample(1).ImageShiftY! = 0#
sample(1).beammode% = DefaultBeamMode%      ' 0 = spot, 1 = scan, 2 = digital spot

sample(1).OxideOrElemental% = 0#
sample(1).Description$ = vbNullString
sample(1).DisplayAsOxideFlag = False
sample(1).SampleDensity! = DEFAULTDENSITY!

sample(1).AtomicPercentFlag = False
sample(1).FormulaElementFlag% = 0
sample(1).DifferenceElementFlag% = 0
sample(1).DifferenceFormulaFlag% = 0
sample(1).StoichiometryElementFlag% = 0
sample(1).RelativeElementFlag% = 0

sample(1).FormulaElement$ = vbNullString
sample(1).FormulaRatio! = 0#
sample(1).DifferenceElement$ = vbNullString
sample(1).DifferenceFormula$ = vbNullString
sample(1).StoichiometryElement$ = vbNullString
sample(1).StoichiometryRatio! = 0#
sample(1).RelativeElement$ = vbNullString
sample(1).RelativeToElement$ = vbNullString
sample(1).RelativeRatio! = 0#
sample(1).MineralFlag% = 0
sample(1).DetectionLimitsFlag = False
sample(1).DetectionLimitsProjectedFlag% = False
sample(1).HomogeneityFlag% = False
sample(1).HomogeneityAlternateFlag% = False
sample(1).CorrelationFlag% = False

sample(1).DisplayAmphiboleCalculationFlag% = False
sample(1).DisplayBiotiteCalculationFlag% = False

sample(1).HydrogenStoichiometryFlag% = False
sample(1).HydrogenStoichiometryRatio! = 0#

sample(1).FerrousFerricCalculationFlag = False
sample(1).FerrousFerricTotalCations! = 0#
sample(1).FerrousFerricTotalOxygens! = 0#

sample(1).CoatingFlag = DefaultSampleCoatingFlag%  ' 0 = uncoated, 1 = coated
sample(1).CoatingElement% = DefaultSampleCoatingElement%
sample(1).CoatingDensity! = DefaultSampleCoatingDensity!
sample(1).CoatingThickness! = DefaultSampleCoatingThickness!  ' in angstroms
sample(1).CoatingSinThickness! = DefaultSampleCoatingThickness! / Sin(DefaultTakeOff! * PI! / 180#)

' EDS sample flags
sample(1).EDSSpectraFlag = False
sample(1).EDSSpectraUseFlag = False

sample(1).CLSpectraFlag = False

sample(1).IntegratedIntensitiesFlag% = False
sample(1).IntegratedIntensitiesUseFlag% = False

sample(1).FiducialSetNumber% = 0

sample(1).ImageShiftX! = 0#
sample(1).ImageShiftY! = 0#

For j% = 1 To MAXROW%
sample(1).Linenumber&(j%) = 0
sample(1).LineStatus(j%) = False
sample(1).DateTimes(j%) = 0

sample(1).OnBeamCounts!(j%) = 0#
sample(1).AbBeamCounts!(j%) = 0#
sample(1).OnBeamCounts2!(j%) = 0#
sample(1).AbBeamCounts2!(j%) = 0#
sample(1).EDSUnknownCountFactors!(j%) = 1
sample(1).CLUnknownCountFactors!(j%) = 1

For k% = 1 To MAXAXES%
sample(1).StagePositions!(j%, k%) = 0#
Next k%

Next j%

' Initialize the element arrays
amsg$ = "Initializing sample element parameters..."
For i% = 1 To MAXCHAN%
Call InitElement(i%, sample())
If ierror Then Exit Sub
Next i%

' Reset combined condition acquisition orders too
amsg$ = "Initializing misc sample parameters..."
For k% = 1 To MAXCOND%
sample(1).ConditionOrders%(k%) = 1
Next k%

' Other
sample(1).OxygenChannel% = 0

' PTC options
sample(1).iptc% = 0
sample(1).PTCModel% = 0
sample(1).PTCDiameter! = 0#
sample(1).PTCDensity! = 0#
sample(1).PTCThicknessFactor! = 0#
sample(1).PTCNumericalIntegrationStep! = 0#

sample(1).AlternatingOnAndOffPeakAcquisitionFlag% = 0

sample(1).SpecifyMatrixByAnalysisUnknownNumber% = 0
sample(1).UnknownCountTimeForInterferenceStandardFlag = False

sample(1).OnPeakTimeFractionFlag = False
sample(1).OnPeakTimeFractionValue! = 1#
sample(1).ChemicalAgeCalculationFlag = False

sample(1).PTCDoNotNormalizeSpecifiedFlag = False
sample(1).EDSSpectraQuantMethodOrProject$ = vbNullString

sample(1).FerrousFerricCalculationFlag = False
sample(1).FerrousFerricTotalCations! = 0#
sample(1).FerrousFerricTotalOxygens! = 0#

sample(1).LastEDSSpecifiedCountTime! = EDSSpecifiedCountTime!
sample(1).LastEDSUnknownCountFactor! = EDSUnknownCountFactor!

sample(1).LastCLSpecifiedCountTime! = CLSpecifiedCountTime!
sample(1).LastCLUnknownCountFactor! = CLUnknownCountFactor!
sample(1).LastCLDarkSpectraCountTimeFraction! = CLDarkSpectraCountTimeFraction!

sample(1).MaterialType$ = vbNullString

Exit Sub

InitSampleError:
MsgBox Error$ & ", " & amsg$, vbOKOnly + vbCritical, "InitSample"
ierror = True
Exit Sub

End Sub

Sub InitStandard()
' Initialize standard array

ierror = False
On Error GoTo InitStandardError

Dim i As Integer

NumberofStandards% = 0
For i% = 1 To MAXSTD%
StandardNames$(i%) = vbNullString
StandardNumbers%(i%) = 0
StandardDescriptions$(i%) = vbNullString
StandardDensities!(i%) = 0#

StandardCoatingFlag%(i%) = 0
StandardCoatingElement%(i%) = 0
StandardCoatingDensity!(i%) = 0#
StandardCoatingThickness!(i%) = 0#
Next i%

Exit Sub

InitStandardError:
MsgBox Error$, vbOKOnly + vbCritical, "InitStandard"
ierror = True
Exit Sub

End Sub

Sub InitStandardIndex()
' Initialize the standard index

ierror = False
On Error GoTo InitStandardIndexError

Dim i As Integer

' Zero program arrays
NumberOfAvailableStandards% = 0
For i% = 1 To MAXINDEX%
StandardIndexNumbers%(i%) = 0
StandardIndexNames$(i%) = vbNullString
StandardIndexDescriptions$(i%) = vbNullString
StandardIndexDensities!(i%) = 0#
StandardIndexMaterialTypes$(i%) = vbNullString
Next i%

Exit Sub

' Errors
InitStandardIndexError:
MsgBox Error$, vbOKOnly + vbCritical, "InitStandardIndex"
ierror = True
Exit Sub

End Sub

Sub InitStandards(analysis As TypeAnalysis)
' Initialize analysis arrays for standard factors

ierror = False
On Error GoTo InitStandardsError

Dim i As Integer, j As Integer, k As Integer

' Dimension dynamic arrays
ReDim analysis.WtsData(1 To MAXROW%, 1 To MAXCHAN1%) As Single
ReDim analysis.CalData(1 To MAXROW%, 1 To MAXCHAN1%) As Single

ReDim analysis.StdZAFCors(1 To MAXZAFCOR%, 1 To MAXSTD%, 1 To MAXCHAN%) As Single

ReDim analysis.StdBetas(1 To MAXSTD%, 1 To MAXCHAN%) As Single
ReDim analysis.StdContinuumCorrections(1 To MAXSTD%, 1 To MAXCHAN%) As Single
ReDim analysis.StdMACs(1 To MAXSTD%, 1 To MAXCHAN%) As Single
ReDim analysis.StdPercents(1 To MAXSTD%, 1 To MAXCHAN%) As Single

ReDim analysis.UnkZAFCors(1 To MAXZAFCOR%, 1 To MAXCHAN%) As Single
ReDim analysis.StdAssignsZAFCors(1 To MAXZAFCOR%, 1 To MAXCHAN%) As Single

ReDim analysis.MANFitCoefficients(1 To MAXCOEFF%, 1 To MAXCHAN%) As Single
ReDim analysis.MANAssignsCounts(1 To MAXMAN%, 1 To MAXCHAN%) As Single
ReDim analysis.MANAssignsRows(1 To MAXMAN%, 1 To MAXCHAN%) As Integer

ReDim analysis.MANAssignsCountTimes(1 To MAXMAN%, 1 To MAXCHAN%) As Single
ReDim analysis.MANAssignsBeamCurrents(1 To MAXMAN%, 1 To MAXCHAN%) As Single

ReDim analysis.StdAssignsIntfCounts(1 To MAXINTF%, 1 To MAXCHAN%) As Single
ReDim analysis.StdAssignsIntfRows(1 To MAXINTF%, 1 To MAXCHAN%) As Integer

' Initialize Standard arrays
For i% = 1 To MAXCHAN%
analysis.StdAssignsPercents!(i%) = 0#
analysis.StdAssignsKfactors!(i%) = 0#
analysis.StdAssignsBetas!(i%) = 0#
analysis.StdAssignsCounts(i%) = 0#
analysis.StdAssignsActualKilovolts(i%) = 0#
analysis.StdAssignsEdgeEnergies(i%) = 0#
analysis.StdAssignsActualOvervoltages(i%) = 0#

analysis.UnkContinuumCorrections!(i%) = 0#

For j% = 1 To MAXSTD%
analysis.StdPercents!(j%, i%) = 0#
analysis.StdContinuumCorrections!(j%, i%) = 0#
analysis.StdMACs!(j%, i%) = 0#

For k% = 1 To MAXZAFCOR%
analysis.StdZAFCors!(k%, j%, i%) = 1#
Next k%
analysis.StdBetas!(j%, i%) = 1#
Next j%

For j% = 1 To MAXINTF%
analysis.StdAssignsIntfCounts(j%, i%) = 0#
Next j%
Next i%

For j% = 1 To MAXSTD%
analysis.StdZbars!(j%) = 0#
Next j%

Exit Sub

' Errors
InitStandardsError:
MsgBox Error$, vbOKOnly + vbCritical, "InitStandards"
ierror = True
Exit Sub

End Sub

Sub InitImage(tImage As TypeImage)
' Initializes arrays for an image

ierror = False
On Error GoTo InitImageError

tImage.ImageAnalogAverages% = 0
tImage.ImageChannelName$ = vbNullString
tImage.ImageIx% = 0
tImage.ImageIy% = 0
tImage.ImageNumber% = 0
tImage.ImageSampleNumber% = 0
tImage.ImageXmin! = 0
tImage.ImageXmax! = 0
tImage.ImageYmin! = 0
tImage.ImageYmax! = 0
tImage.ImageZmin& = 0
tImage.ImageZmax& = 0

tImage.ImageMag! = 0#

tImage.ImageZ1! = 0#
tImage.ImageZ2! = 0#
tImage.ImageZ3! = 0#
tImage.ImageZ4! = 0#

'For i% = 1 To MAXIMAGEIX%
'For j% = 1 To MAXIMAGEIY%
'tImage.ImageData&(i%, j%) = 0   ' dimensioned dynamically
'Next j%
'Next i%

Exit Sub

' Errors
InitImageError:
MsgBox Error$, vbOKOnly + vbCritical, "InitImage"
ierror = True
Exit Sub

End Sub

Sub InitScan(tScan As TypeScan)
' Initializes arrays for a scan object

ierror = False
On Error GoTo InitScanError

tScan.ScanToRow% = 0
tScan.ScanNumber& = 0
tScan.ScanChannel% = 0
tScan.ScanType% = 0       ' 1 = PHA, 2 = Bias, 3 = Gain, 4 = interval peaking, 5 = parabolic peaking, 6 = ROM peaking, 7 = prescan, 8 = postscan
tScan.ScanMotor% = 0
tScan.ScanElsyms$ = vbNullString
tScan.ScanXrsyms$ = vbNullString
tScan.ScanCrystal$ = vbNullString
tScan.ScanCountTime! = 0#
tScan.ScanFitCentroid! = 0#
tScan.ScanFitThreshold! = 0#
tScan.ScanFitPtoB! = 0#
tScan.ScanFitCoeff1! = 0#
tScan.ScanFitCoeff2! = 0#
tScan.ScanFitCoeff3! = 0#
tScan.ScanFitDeviation! = 0#
tScan.ScanPoints% = 0

'For i% = 1 To MAXROMSCAN%
'tScan.ScanXdata!(i%) = 0#   ' dimensioned dynamically
'tScan.ScanYdata!(i%) = 0#
'Next i%

tScan.ScanCurrentTakeOff! = 0#
tScan.ScanCurrentKilovolts! = 0#
tScan.ScanCurrentBeamCurrent! = 0#
tScan.ScanCurrentBeamSize! = 0#
tScan.ScanCurrentColumnConditionMethod% = 0    ' 0 = TKCS, 1 = condition string
tScan.ScanCurrentColumnConditionString$ = vbNullString
tScan.ScanCurrentMagnification! = 0#
tScan.ScanCurrentBeamMode% = 0
tScan.ScanCurrentBaseline! = 0#
tScan.ScanCurrentWindow! = 0#
tScan.ScanCurrentGain! = 0#
tScan.ScanCurrentBias! = 0#
tScan.ScanCurrentInteDiffMode% = 0
tScan.ScanCurrentDeadTime! = 0#
tScan.ScanCurrentPeakingStartSize! = 0#
tScan.ScanCurrentPeakingStopSize! = 0#
tScan.ScanCurrentPositionSample$ = vbNullString
tScan.ScanCurrentStageX! = 0#
tScan.ScanCurrentStageY! = 0#
tScan.ScanCurrentStageZ! = 0#
tScan.ScanROMPeakingType% = 0
tScan.ScanDateTime = Now
tScan.ScanROMPeakingSet% = 0    ' 0 = final, 1 = coarse
tScan.ScanPHAHardwareType% = PHAHardwareType%   ' 0 = traditional PHA, 1 = MCA PHA

Exit Sub

' Errors
InitScanError:
MsgBox Error$, vbOKOnly + vbCritical, "InitScan"
ierror = True
Exit Sub

End Sub

Sub InitGetZAFSetZAF2(itemp As Integer)
' Load pre-defined ZAF correction (used by CalcZAF for binary k-ratio calculations and Matrix)

ierror = False
On Error GoTo InitGetZAFSetZAF2Error

' Select Individual Corrections (to maintain backward compatibility with Probe for EPMA)
If itemp% = 0 Then

' Select the default (Armstrong/Love-Scott)ZAF matrix correction algorithms
ElseIf itemp% = 1 Then
ibsc% = 2
imip% = 1
iphi% = 2
iabs% = 9
istp% = 4
ibks% = 4   ' note ibks is indexed from zero

' Conventional Philibert/Duncumb Reed
ElseIf itemp% = 2 Then
ibsc% = 1
imip% = 2
iphi% = 5
iabs% = 1
istp% = 1
ibks% = 1   ' note ibks is indexed from zero

' Heinrich/Duncumb-Reed
ElseIf itemp% = 3 Then
ibsc% = 1
imip% = 1
iphi% = 5
iabs% = 1
istp% = 1
ibks% = 2   ' note ibks is indexed from zero

' Love-Scott I
ElseIf itemp% = 4 Then
ibsc% = 2
imip% = 1
iphi% = 2
iabs% = 4
istp% = 4
ibks% = 4   ' note ibks is indexed from zero

' Love-Scott II
ElseIf itemp% = 5 Then
ibsc% = 2
imip% = 1
iphi% = 2
iabs% = 6
istp% = 4
ibks% = 4   ' note ibks is indexed from zero

' Packwood Phi(PZ) (EPQ-91)
ElseIf itemp% = 6 Then
ibsc% = 1
imip% = 5
iphi% = 7
iabs% = 14
istp% = 6
ibks% = 0   ' note ibks is indexed from zero

' Bastin original Phi(PZ)
ElseIf itemp% = 7 Then
ibsc% = 2
imip% = 3
iphi% = 2
iabs% = 10
istp% = 6
ibks% = 0

' Bastin PROZA Phi(PZ) (EPQ-91)
ElseIf itemp% = 8 Then
ibsc% = 3
imip% = 3
iphi% = 4
iabs% = 15
istp% = 5
ibks% = 7   ' note ibks is indexed from zero

' Pouchout & Pichoir - Full
ElseIf itemp% = 9 Then
ibsc% = 3
imip% = 3
iphi% = 4
iabs% = 12
istp% = 5
ibks% = 7   ' note ibks is indexed from zero

' Pouchout & Pichoir - Simplified
ElseIf itemp% = 10 Then
ibsc% = 3
imip% = 3
iphi% = 4
iabs% = 13
istp% = 5
ibks% = 7   ' note ibks is indexed from zero
End If

Exit Sub

' Errors
InitGetZAFSetZAF2Error:
MsgBox Error$, vbOKOnly + vbCritical, "InitGetZAFSetZAF2"
ierror = True
Exit Sub

End Sub
