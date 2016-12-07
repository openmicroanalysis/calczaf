Attribute VB_Name = "CodeUPDATE"
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

Private Const NOTLOADEDVALUE! = -0.0000000009   ' flag for array member not loaded

Dim UpdateAnalyzeReloadCorrection As Boolean
Dim UpdateAnalyzeReloadDrift As Boolean

' Drift calculation arrays
Dim StdAssignsDriftCounts(1 To MAXSET%, 1 To MAXCHAN%) As Single
Dim StdAssignsDriftTimes(1 To MAXSET%, 1 To MAXCHAN%) As Single
Dim StdAssignsDriftBeams(1 To MAXSET%, 1 To MAXCHAN%) As Single
Dim StdAssignsDateTimes(1 To MAXSET%, 1 To MAXCHAN%) As Double
Dim StdAssignsBackgroundTypes(1 To MAXSET%, 1 To MAXCHAN%) As Integer
Dim StdAssignsSampleRows(1 To MAXSET%, 1 To MAXCHAN%) As Integer

Dim StdAssignsDriftBgdCounts(1 To MAXSET%, 1 To MAXCHAN%) As Single

Dim StdAssignsIntfDriftCounts(1 To MAXSET%, 1 To MAXINTF%, 1 To MAXCHAN%) As Single
Dim StdAssignsIntfDateTimes(1 To MAXSET%, 1 To MAXINTF%, 1 To MAXCHAN%) As Double
Dim StdAssignsIntfBackgroundTypes(1 To MAXSET%, 1 To MAXINTF%, 1 To MAXCHAN%) As Integer
Dim StdAssignsIntfSampleRows(1 To MAXSET%, 1 To MAXINTF%, 1 To MAXCHAN%) As Integer

Dim StdAssignsSets(1 To MAXCHAN%) As Integer
Dim StdAssignsIntfSets(1 To MAXINTF%, 1 To MAXCHAN%) As Integer

' Missing MAN assignments array
Dim MissingMANAssignmentsStringsNumberOf As Integer
Dim MissingMANAssignmentsStrings() As String

Dim UpdateTmpSample(1 To 1) As TypeSample
Dim UpdateOldSample(1 To 1) As TypeSample   ' use only for function UpdateChangedSample!
Dim UpdateAnalysis As TypeAnalysis

Sub UpdateAnalyze(analysis As TypeAnalysis, sample() As TypeSample, stdsample() As TypeSample)
' Load the analysis arrays for the specified sample to be analyzed

ierror = False
On Error GoTo UpdateAnalyzeError

Dim changedsample As Boolean

icancelanal = False

' Load MAN assignments in case sample or assigned standards or interference standard have MAN corrected elements (always do this) (added 7/20/2011, v. 8.50)
Call DataMAN(Int(1), sample())
If ierror Then Exit Sub

' Init the analysis structure
Call InitStandards(analysis)
If ierror Then Exit Sub

' Update not analyzed elements for standards and oxygen by stoichiometry (always do this) (added 7/16/2011, v. 8.50)
Call UpdateStdElements(analysis, sample(), stdsample())
If ierror Then Exit Sub

' Reload the element arrays based on the unknown sample setup (always do this) (added 7/16/2011, v. 8.50)
Call ElementGetData(sample())
If ierror Then Exit Sub

' Reload ZAF setup and calculate oxygen channel (always do this) (added 7/20/2011, v. 8.50)
If CorrectionFlag% <> MAXCORRECTION% Then
Call ZAFSetZAF(sample())
If ierror Then Exit Sub
Else
'Call ZAFSetZAF3(sample())
'If ierror Then Exit Sub
End If

' Check to see if the sample setup has changed, or a standard data point was acquired. Save here so it only needs to be called once in this routine!!!
changedsample = UpdateChangedSample(sample())
If ierror Then Exit Sub

' Load the assigned standard count drift arrays based on the sample standard assignments and conditions. (added 7/16/2011, v. 8.50)
If UpdateChangedCorrection() Or DoNotUseFastQuantFlag Or changedsample Or AllAnalysisUpdateNeeded Or UpdateAnalyzeReloadCorrection Then
UpdateAnalyzeReloadCorrection = True    ' force update in case error occurs during update (added 10-22-2011, v.8.58)

' Initialize analysis drift arrays
If changedsample Or UpdateAnalyzeReloadDrift Then
Call InitStandards(analysis)
If ierror Then Exit Sub
Call AnalyzeInitAnalysis(analysis)
If ierror Then Exit Sub
End If

' Load MAN assignments in case sample or assigned standards or interference standard have MAN corrected elements
Call DataMAN(Int(1), sample())
If ierror Then Exit Sub

' Recalculate standard k-factors for this sample based on sample conditions (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% = 0 Or CorrectionFlag% = 5 Or CorrectionFlag% = MAXCORRECTION% Then
Call AnalyzeStatusAnal("Calculating standard k-factors...")
End If
If CorrectionFlag% > 0 And CorrectionFlag% < 5 Then
Call AnalyzeStatusAnal("Calculating standard beta-factors...")
End If

' Check for cancel
If icancelanal Then
ierror = True
Exit Sub
End If

' Update not analyzed elements for standards and oxygen by stoichiometry
Call UpdateStdElements(analysis, sample(), stdsample())
If ierror Then Exit Sub

' If PTC correction, update geometry in ZAFSetZAF, BEFORE calculating k-factors!!!
If UseParticleCorrectionFlag And iptc% Then
Call ElementGetData(sample())
If ierror Then Exit Sub
If CorrectionFlag% <> MAXCORRECTION% Then
Call ZAFSetZAF(sample())
If ierror Then Exit Sub
Else
'Call ZAFSetZAF3(sample())
'If ierror Then Exit Sub
End If
End If

' Re-calculate standard k-factors or beta-factors
Call UpdateStdKfacs(analysis, sample(), stdsample())
If ierror Then Exit Sub

' Additional elements now added to sample, next load alpha-factor arrays if necessary (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% > 0 And CorrectionFlag% < 5 Then
Call AFactorLoadFactors(analysis, sample())
If ierror Then Exit Sub
End If

' Reload the element arrays based on the unknown sample setup
Call ElementGetData(sample())
If ierror Then Exit Sub

' Update the sample setup primary intensity ZAF arrays
Call AnalyzeStatusAnal("Re-calculating primary intensities...")
If icancelanal Then
ierror = True
Exit Sub
End If

' Setup ZAF arrays (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% = 0 Or CorrectionFlag% = 5 Then
Call ZAFSetZAF(sample())
If ierror Then Exit Sub
ElseIf CorrectionFlag% = MAXCORRECTION% Then
'Call ZAFSetZAF3(sample())
'If ierror Then Exit Sub

' Or just oxygen channel
Else
Call ZAFGetOxygenChannel(sample())
If ierror Then Exit Sub
End If

' Now save calculated analysis structure for next call (in case re-calculate std k-factors and reload drift arrays not performed for next passed structure)
UpdateAnalysis = analysis

' Check to see if the sample setup has changed, or a standard data point was acquired. Load the assigned standard count drift
' arrays based on the sample standard assignments and conditions. (commented out 7/16/2011, v. 8.50)
If DoNotUseFastQuantFlag Or changedsample Or AllAnalysisUpdateNeeded Or UpdateAnalyzeReloadDrift Then
UpdateAnalyzeReloadDrift = True    ' force update in case error occurs during update (added 10-22-2011, v.8.58)

Call AnalyzeStatusAnal("Loading standard, interference standard and MAN drift correction arrays...")
If icancelanal Then
ierror = True
Exit Sub
End If

' Load drift arrys for standards and interference standards
Call UpdateGetStandards(sample())
If ierror Then Exit Sub

' Load all MAN standard drift sets for standard and unknown MAN
' background corrections. Note that MAN drift counts need to be loaded
' even if the sample is off-peak corrected because some or all of the
' standards may be MAN corrected.
Call UpdateGetMANStandards(Int(0), sample())
If ierror Then Exit Sub

Call AnalyzeStatusAnal("Calculating standard MAN corrections...")
If icancelanal Then
ierror = True
Exit Sub
End If

' Correct the assigned standard count drift arrays for MAN backgrounds
Call UpdateStdMANBackgrounds(analysis, sample())
If ierror Then Exit Sub

Call AnalyzeStatusAnal("Calculating standard interference corrections...")
If icancelanal Then
ierror = True
Exit Sub
End If

' Correct the assigned standard count drift arrays for interferences
Call UpdateStdInterferences(analysis, sample())
If ierror Then Exit Sub

' Load all standard percents (even if not assigned to any element)
Call UpdateGetStdPercents(analysis, sample(), stdsample())
If ierror Then Exit Sub

' Re-set update flags
AllAnalysisUpdateNeeded = False

' Now save calculated analysis structure for next call (in case re-calculate std k-factors and reload drift arrays not performed for next passed structure)
UpdateAnalysis = analysis
End If
End If

' Re-load analysis structure (always) for last calculated/loaded structure
analysis = UpdateAnalysis

' Re-set module level reload flags (added 10-22-2011, v.8.58)
UpdateAnalyzeReloadCorrection = False
UpdateAnalyzeReloadDrift = False

Call AnalyzeStatusAnal(vbNullString)
Exit Sub

' Errors
UpdateAnalyzeError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateAnalyze"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

End Sub

Function UpdateGetMaxSet(mode As Integer, sample() As TypeSample) As Integer
' Return the maximum number of assigned sets in the sample
'  mode = 1  do standard sets
'  mode = 2  do standard interference sets
'  mode = 3  do standard MAN sets

ierror = False
On Error GoTo UpdateGetMaxSetError

Dim i As Integer, j As Integer
Dim imax As Integer

imax% = 0

' Standards sets
If mode% = 1 Then
For i% = 1 To sample(1).LastElm%
If StdAssignsSets%(i%) > imax% Then imax% = StdAssignsSets%(i%)
Next i%
End If

' Standard interference sets
If mode% = 2 Then
For i% = 1 To sample(1).LastElm%
For j% = 1 To MAXINTF%
If StdAssignsIntfSets%(j%, i%) > imax% Then imax% = StdAssignsIntfSets%(j%, i%)
Next j%
Next i%
End If

' Standard MAN sets
If mode% = 3 Then
For i% = 1 To sample(1).LastElm%
For j% = 1 To MAXMAN%
If MANAssignsSets%(j%, i%) > imax% Then imax% = MANAssignsSets%(j%, i%)
Next j%
Next i%
End If

UpdateGetMaxSet% = imax%
Exit Function

' Errors
UpdateGetMaxSetError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateGetMaxSet"
ierror = True
Exit Function

End Function

Sub UpdateGetStandards(sample() As TypeSample)
' Routine to load all standard and interference standard sets for drift correction
' calculation of both standards and unknowns. See routine UpdateCalculateStdDrift
' for actual drift calculation.

ierror = False
On Error GoTo UpdateGetStandardsError

Dim i As Integer, j As Integer, k As Integer
Dim samplerow As Integer, ip As Integer, chan As Integer
Dim jmax As Integer, kmax As Integer

Dim average As TypeAverage
Dim averbgd As TypeAverage

Dim average1 As TypeAverage
Dim average2 As TypeAverage

If DebugMode Then
Call IOWriteLog(vbCrLf & "Entering UpdateGetStandards...")
End If

' Initialize drift arrays (note 1 to MAXSET%)
If sample(1).LastElm% < 1 Then GoTo UpdateGetStandardsNoElements

' Check for missing standard assignments and initialize drift arrays
For i% = 1 To sample(1).LastElm%
If sample(1).StdAssigns%(i%) = 0 Then GoTo UpdateGetStandardsMissingAssignment
For k% = 1 To MAXSET%
StdAssignsDriftCounts!(k%, i%) = NOTLOADEDVALUE!
StdAssignsDriftTimes!(k%, i%) = 0#
StdAssignsDriftBeams!(k%, i%) = 0#
StdAssignsDateTimes#(k%, i%) = 0#
StdAssignsDriftBgdCounts!(k%, i%) = NOTLOADEDVALUE!
StdAssignsBackgroundTypes%(k%, i%) = 0#
StdAssignsSampleRows%(k%, i%) = 0

For j% = 1 To MAXINTF%
StdAssignsIntfDriftCounts!(k%, j%, i%) = NOTLOADEDVALUE!
StdAssignsIntfDateTimes(k%, j%, i%) = 0#
StdAssignsIntfBackgroundTypes%(k%, j%, i%) = 0#
StdAssignsIntfSampleRows%(k%, j%, i%) = 0
Next j%
Next k%

' Initialize the set counters
StdAssignsSets%(i%) = 0
For j% = 1 To MAXINTF%
StdAssignsIntfSets%(j%, i%) = 0
Next j%
Next i%

' Search through samples and load all assigned standard sets
For samplerow% = 1 To NumberofSamples%

' Check if not a standard or all deleted lines
If SampleTyps%(samplerow%) <> 1 Or SampleDels%(samplerow%) Then GoTo 9100

' See if this standard is used for the standard assignments for this sample
For i% = 1 To sample(1).LastElm%
If SampleNums%(samplerow%) = sample(1).StdAssigns%(i%) Then GoTo 7100
Next i%

' See if this standard is used for the interference standard assignments for this sample
For i% = 1 To sample(1).LastElm%
For j% = 1 To MAXINTF%
If SampleNums%(samplerow%) = sample(1).StdAssignsIntfStds%(j%, i%) Then GoTo 7100
Next j%
Next i%

' Standard is not used for any standard assignments, try next sample
GoTo 9100

' Load data from disk file into "CorData" array for this standard
7100:

' Update status form
msg$ = SampleGetString$(samplerow%)
Call AnalyzeStatusAnal("Loading count data for " & msg$)
If icancelanal Then
ierror = True
Exit Sub
End If

' Load data for this standard
Call DataGetMDBSample(samplerow%, UpdateTmpSample())
If ierror Then Exit Sub

' Check for valid data points
If UpdateTmpSample(1).Datarows% < 1 Then GoTo 9100
If UpdateTmpSample(1).GoodDataRows% < 1 Then GoTo 9100

' Check that integrated intensity sample flag match (will not match if quick standard!)
'If UpdateTmpSample(1).IntegratedIntensitiesUseFlag% <> sample(1).IntegratedIntensitiesUseFlag% Then GoTo 9100

' If analytical conditions do not match selected sample, skip
If Not UpdateTmpSample(1).CombinedConditionsFlag And Not sample(1).CombinedConditionsFlag Then
If UpdateTmpSample(1).takeoff! <> sample(1).takeoff! Then GoTo 9100
If UpdateTmpSample(1).kilovolts! <> sample(1).kilovolts! Then GoTo 9100
End If

' Obtain EDS net intensities for this standard
If sample(1).EDSSpectraFlag And sample(1).EDSSpectraUseFlag Then
If UpdateEDSCheckForEDSElements(Int(1), sample()) Then
If UpdateTmpSample(1).EDSSpectraFlag Then
If UpdateEDSCheckForEDSElements(Int(1), UpdateTmpSample()) Then
Call UpdateEDSSpectraNetIntensities(UpdateTmpSample())
If ierror Then Exit Sub
End If
End If
End If
End If

' Do a Savitzky-Golay smooth to the integrated intensity data if smooth selected
If UpdateTmpSample(1).IntegratedIntensitiesUseFlag And IntegratedIntensityUseSmoothingFlag Then
For chan% = 1 To UpdateTmpSample(1).LastElm%
If UpdateTmpSample(1).IntegratedIntensitiesUseIntegratedFlags%(chan%) Then
Call DataCorrectDataIntegratedSmooth(chan%, IntegratedIntensitySmoothingPointsPerSide%, UpdateTmpSample())
If ierror Then Exit Sub
End If
Next chan%
End If

' Correct the data for dead time, beam drift and off-peak background (results returned in sample(1).CorData!())
If Not UseAggregateIntensitiesFlag Then
Call DataCorrectData(Int(0), UpdateTmpSample())
If ierror Then Exit Sub

Else
Call DataCorrectData(Int(2), UpdateTmpSample())     ' skip aggregate intensity load
If ierror Then Exit Sub
Call DataCorrectDataAggregate(Int(2), sample(), UpdateTmpSample())     ' perform aggregate intensity based on unknown
If ierror Then Exit Sub
End If

' Correct on-peak intensities (UpdateTmpSample(1).CorData!()) for TDI volatile correction
If UseVolElFlag And CorrectionFlag% <> 5 Then
Call VolatileCalculateCorrectionAll(UpdateTmpSample())
If ierror Then Exit Sub
End If

' Average the standard count data and datetime (average of sample(1).CorData!())
Call MathCountAverage(average, UpdateTmpSample())
If ierror Then Exit Sub

' Average background count data
Call MathArrayAverage(averbgd, UpdateTmpSample(1).BgdData!(), UpdateTmpSample(1).Datarows%, UpdateTmpSample(1).LastElm%, UpdateTmpSample())
If ierror Then Exit Sub

' Average beam currents
If Not sample(1).CombinedConditionsFlag Then
Call MathAverage(average1, sample(1).OnBeamCounts!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
Call MathAverage(average2, sample(1).OnBeamCounts2!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
Else
Call MathArrayAverage(average1, sample(1).OnBeamCountsArray!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
Call MathArrayAverage(average2, sample(1).OnBeamCountsArray2!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
End If

' Load counts from this standard, if assigned to the sample elements
For i% = 1 To sample(1).LastElm%
If sample(1).DisableQuantFlag%(i%) = 1 Then GoTo 7500
If sample(1).StdAssigns%(i%) <> SampleNums%(samplerow%) Then GoTo 7500

' Make sure the element is analyzed in the standard also
ip% = IPOS5(Int(0), i%, sample(), UpdateTmpSample())
If ip% = 0 Then GoTo 7500

' If analytical conditions do not match selected sample, skip
If sample(1).CombinedConditionsFlag Or UpdateTmpSample(1).CombinedConditionsFlag Then
If sample(1).TakeoffArray!(i%) <> UpdateTmpSample(1).TakeoffArray!(ip%) Then GoTo 7500
If sample(1).KilovoltsArray!(i%) <> UpdateTmpSample(1).KilovoltsArray!(ip%) Then GoTo 7500
End If

' Check for Bragg order
If sample(1).BraggOrders%(i%) <> UpdateTmpSample(1).BraggOrders%(ip%) Then GoTo 7500

' Check for matching integrated intensity flag
If sample(1).IntegratedIntensitiesUseIntegratedFlags%(i%) <> UpdateTmpSample(1).IntegratedIntensitiesUseIntegratedFlags%(ip%) Then GoTo 7500

' Check for disabled acquisition in standard
If UpdateTmpSample(1).DisableAcqFlag%(ip%) = 1 Then GoTo 7500

' Check for disabled quantification in standard (do not check standard flags)
'If UpdateTmpSample(1).DisableQuantFlag%(ip%) = 1 Then GoTo 7500

' Check for same peak position
If AnalysisCheckForSamePeakPositions Then
If Not MiscDifferenceIsSmall(sample(1).OnPeaks!(i%), UpdateTmpSample(1).OnPeaks!(ip%), 0.00005) Then GoTo 7500
End If

' Check for same PHA settings
If AnalysisCheckForSamePHASettings Then
If Not MiscDifferenceIsSmall(sample(1).Baselines!(i%), UpdateTmpSample(1).Baselines!(ip%), 0.005) Then GoTo 7500
If Not MiscDifferenceIsSmall(sample(1).Windows!(i%), UpdateTmpSample(1).Windows!(ip%), 0.005) Then GoTo 7500
If Not MiscDifferenceIsSmall(sample(1).Gains!(i%), UpdateTmpSample(1).Gains!(ip%), 0.005) Then GoTo 7500
If Not MiscDifferenceIsSmall(sample(1).Biases!(i%), UpdateTmpSample(1).Biases!(ip%), 0.005) Then GoTo 7500
If sample(1).InteDiffModes%(i%) <> UpdateTmpSample(1).InteDiffModes%(ip%) Then GoTo 7500
End If

' Matching conditions, now increment set counter
If StdAssignsSets%(i%) + 1 > MAXSET% Then GoTo UpdateGetStandardsTooManyStdSets
StdAssignsSets%(i%) = StdAssignsSets%(i%) + 1

' Load average counts
StdAssignsDriftCounts!(StdAssignsSets%(i%), i%) = average.averags!(ip%)

' Load average count times
StdAssignsDriftTimes!(StdAssignsSets%(i%), i%) = UpdateTmpSample(1).LastOnCountTimes!(ip%) ' just assume without averaging

' Load average beam currents
If Not sample(1).CombinedConditionsFlag Then
StdAssignsDriftBeams!(StdAssignsSets%(i%), i%) = average1.averags!(1)
If average2.averags!(1) <> 0# Then StdAssignsDriftBeams!(StdAssignsSets%(i%), i%) = (average1.averags!(1) + average2.averags!(1)) / 2#
Else
StdAssignsDriftBeams!(StdAssignsSets%(i%), i%) = average1.averags!(i%)
If average2.averags!(i%) <> 0# Then StdAssignsDriftBeams!(StdAssignsSets%(i%), i%) = (average1.averags!(i%) + average2.averags!(i%)) / 2#
End If

' Load average "DateTime"
StdAssignsDateTimes#(StdAssignsSets%(i%), i%) = average.AverDateTime#

' Load background type for UpdateStdMANBackgrounds
StdAssignsBackgroundTypes%(StdAssignsSets%(i%), i%) = UpdateTmpSample(1).BackgroundTypes%(ip%)  ' 0=off-peak, 1=MAN, 2=multipoint

' Load the sample row number for saving the sample to the SETUP database
StdAssignsSampleRows%(StdAssignsSets%(i%), i%) = samplerow%

' Load background counts for UpdateTypeIntensities
StdAssignsDriftBgdCounts!(StdAssignsSets%(i%), i%) = averbgd.averags!(ip%)

7500:  Next i%

' Load counts from this interference standard, if assigned to the sample elements
For i% = 1 To sample(1).LastElm%
If sample(1).DisableQuantFlag%(i%) = 0 Then

For j% = 1 To MAXINTF%
If sample(1).StdAssignsIntfStds%(j%, i%) <> SampleNums%(samplerow%) Then GoTo 7600

' Make sure the element is analyzed in the standard also
ip% = IPOS5(Int(0), i%, sample(), UpdateTmpSample())
If ip% = 0 Then GoTo 7600

' If analytical conditions do not match selected sample, skip
If sample(1).CombinedConditionsFlag Or UpdateTmpSample(1).CombinedConditionsFlag Then
If sample(1).TakeoffArray!(i%) <> UpdateTmpSample(1).TakeoffArray!(ip%) Then GoTo 7600
If sample(1).KilovoltsArray!(i%) <> UpdateTmpSample(1).KilovoltsArray!(ip%) Then GoTo 7600
End If

' Check for Bragg order
If sample(1).BraggOrders%(i%) <> UpdateTmpSample(1).BraggOrders%(ip%) Then GoTo 7600

' Check for matching integrated intensity flag
If sample(1).IntegratedIntensitiesUseIntegratedFlags%(i%) <> UpdateTmpSample(1).IntegratedIntensitiesUseIntegratedFlags%(ip%) Then GoTo 7600

' Check for disabled acquisition in standard
If UpdateTmpSample(1).DisableAcqFlag%(ip%) = 1 Then GoTo 7600

' Check for disabled quant in standard (do not check standard flags)
'If UpdateTmpSample(1).DisableQuantFlag%(ip%) = 1 Then GoTo 7600

' Check for same peak position
If AnalysisCheckForSamePeakPositions Then
If Not MiscDifferenceIsSmall(sample(1).OnPeaks!(i%), UpdateTmpSample(1).OnPeaks!(ip%), 0.00005) Then GoTo 7600
End If

' Check for same PHA settings
If AnalysisCheckForSamePHASettings Then
If Not MiscDifferenceIsSmall(sample(1).Baselines!(i%), UpdateTmpSample(1).Baselines!(ip%), 0.005) Then GoTo 7600
If Not MiscDifferenceIsSmall(sample(1).Windows!(i%), UpdateTmpSample(1).Windows!(ip%), 0.005) Then GoTo 7600
If Not MiscDifferenceIsSmall(sample(1).Gains!(i%), UpdateTmpSample(1).Gains!(ip%), 0.005) Then GoTo 7600
If Not MiscDifferenceIsSmall(sample(1).Biases!(i%), UpdateTmpSample(1).Biases!(ip%), 0.005) Then GoTo 7600
If sample(1).InteDiffModes%(i%) <> UpdateTmpSample(1).InteDiffModes%(ip%) Then GoTo 7600
End If

' Matching conditions, now increment set counter
If StdAssignsIntfSets%(j%, i%) + 1 > MAXSET% Then GoTo UpdateGetStandardsTooManyIntfSets
StdAssignsIntfSets%(j, i%) = StdAssignsIntfSets%(j, i%) + 1

' Load average counts
StdAssignsIntfDriftCounts!(StdAssignsIntfSets%(j, i%), j%, i%) = average.averags!(ip%)

' Load average "dateTime"
StdAssignsIntfDateTimes(StdAssignsIntfSets%(j, i%), j%, i%) = average.AverDateTime

' Load background type for UpdateStdMANBackgrounds
StdAssignsIntfBackgroundTypes%(StdAssignsIntfSets%(j%, i%), j%, i%) = UpdateTmpSample(1).BackgroundTypes%(ip%)  ' 0=off-peak, 1=MAN, 2=multipoint

' Load the sample row number for saving the sample to the SETUP3 database
StdAssignsIntfSampleRows%(StdAssignsIntfSets%(j%, i%), j%, i%) = samplerow%

7600:  Next j%
End If
Next i%

9100:  Next samplerow%

' Check for empty standard sets (skip if standard is a "virtual" intensity)
For i% = 1 To sample(1).LastElm%
If sample(1).StdAssignsFlag%(i%) = 0 Then   ' not virtual

' Skip if disabled quant flag
If Not UseAggregateIntensitiesFlag Then
If sample(1).DisableQuantFlag%(i%) = 0 Then
If StdAssignsSets%(i%) = 0 Then GoTo UpdateGetStandardsNoSets
End If

Else
'ip% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())
ip% = IPOS8A(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample(1).KilovoltsArray!(i%), sample())
If ip% = 0 And sample(1).DisableQuantFlag%(i%) = 0 Then
If StdAssignsSets%(i%) = 0 Then GoTo UpdateGetStandardsNoSets
End If
End If

End If

' Check for empty interference standard sets (virtual intensity not allowed)
For j% = 1 To MAXINTF%
If sample(1).DisableQuantFlag%(i%) = 0 Then
If sample(1).StdAssignsIntfStds%(j%, i%) > 0 And StdAssignsIntfSets%(j%, i%) = 0 Then GoTo UpdateGetStandardsNoIntfSets
End If
Next j%

Next i%

' Check for virtual standards and give then a pseudo date/time for interference corrections (assume set 1 always)
'For i% = 1 To sample(1).LastElm%
'If sample(1).StdAssigns%(i%) > 0 And sample(1).StdAssignsFlag%(i%) = 1 Then
'StdAssignsDateTimes#(1, i%) = Now
'End If
'Next i%

If Not DebugMode Then Exit Sub

' Debug, print standard drift array
kmax% = UpdateGetMaxSet(Int(1), sample())

' Print standard assignments and sets
msg$ = vbCrLf & "Standard drift array set numbers:"
Call IOWriteLog(msg$)

' Print standard numbers
msg$ = vbNullString
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(sample(1).StdAssigns%(i%), i80$), a80$)
Next i%
Call IOWriteLog(msg$)

' Print sets
msg$ = vbNullString
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(StdAssignsSets%(i%), i80$), a80$)
Next i%
Call IOWriteLog(msg$)

' Print standard assignments and counts
msg$ = "Standard Drift Arrays Counts (cps/" & Format$(NominalBeam!) & FaradayCurrentUnits$ & "):"
Call IOWriteLog(msg$)

' Print standard numbers
msg$ = vbNullString
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(sample(1).StdAssigns%(i%), i80$), a80$)
Next i%
Call IOWriteLog(msg$)

For k% = 1 To kmax%

' Print drift counts
msg$ = vbNullString
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(StdAssignsDriftCounts!(k%, i%), f81$), a80$)
Next i%
Call IOWriteLog(msg$)
Next k%

' Debug, print interference standard drift arrays
kmax% = UpdateGetMaxSet(Int(2), sample())
jmax% = UpdateGetMaxInterfAssign(sample())

For j% = 1 To jmax%

' Print interference standard assignments and sets
If j% = 1 Then
msg$ = vbCrLf & "Interference Standard Drift Array Set Numbers (cps/" & Format$(NominalBeam!) & FaradayCurrentUnits$ & "):"
Call IOWriteLog(msg$)
End If

' Print interference standard numbers
msg$ = vbNullString
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(sample(1).StdAssignsIntfStds%(j%, i%), i80$), a80$)
Next i%
Call IOWriteLog(msg$)

' Print interference sets
msg$ = vbNullString
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(StdAssignsIntfSets%(j%, i%), i80$), a80$)
Next i%
Call IOWriteLog(msg$)
Next j%

For k% = 1 To kmax%

' Print interference standard assignments and counts
msg$ = "Set " & Format$(k%) & " Interference Standard Drift Arrays Counts (cps/" & Format$(NominalBeam!) & FaradayCurrentUnits$ & "):"
Call IOWriteLog(msg$)
For j% = 1 To jmax%

' Print interference standard numbers
msg$ = vbNullString
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(sample(1).StdAssignsIntfStds%(j%, i%), i80$), a80$)
Next i%
Call IOWriteLog(msg$)

' Print interference drift counts
msg$ = vbNullString
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(StdAssignsIntfDriftCounts!(k%, j%, i%), f82$), a80$)
Next i%
Call IOWriteLog(msg$)

Next j%
Next k%

Exit Sub

' Errors
UpdateGetStandardsError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateGetStandards"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateGetStandardsNoElements:
msg$ = "No analyzed elements in sample " & Format$(sample(1).number%) & " " & sample(1).Name$ & ". Either click the "
msg$ = msg$ & "Acquire! | New Sample, or the Acquire! | Elements/Cations buttons "
msg$ = msg$ & "to add some analyzed elements to the run first."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateGetStandards"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateGetStandardsMissingAssignment:
msg$ = "Element " & MiscAutoUcase$(sample(1).Elsyms$(i%)) & " has not been assigned a standard. "
msg$ = msg$ & "Click the Analyze! | Standard Assignments buttons "
msg$ = msg$ & "to assign standards to the analyzed elements first and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateGetStandards"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateGetStandardsTooManyStdSets:
msg$ = "Too many assigned standard data sets for standard number " & Format$(UpdateTmpSample(1).number%) & ". "
msg$ = msg$ & "Delete some standard sets for this standard number and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateGetStandards"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateGetStandardsTooManyIntfSets:
msg$ = "Too many assigned standard interference data sets for standard number " & Format$(UpdateTmpSample(1).number%) & ". "
msg$ = msg$ & "Delete some standard interference sets for this standard number and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateGetStandards"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateGetStandardsNoSets:
msg$ = "No primary standard intensity data found for standard number " & Format$(sample(1).StdAssigns%(i%)) & " for "
msg$ = msg$ & MiscAutoUcase$(sample(1).Elsyms$(i%)) & " " & sample(1).Xrsyms$(i%) & " "
msg$ = msg$ & "on spectrometer " & Format$(sample(1).MotorNumbers%(i%)) & " crystal " & sample(1).CrystalNames$(i%) & " at "
msg$ = msg$ & Format$(sample(1).KilovoltsArray!(i%)) & " KeV for standard assignments. Either acquire count data "
msg$ = msg$ & "for the indicated element on the indicated standard at the indicated conditions or "
msg$ = msg$ & "change the standard assignment by clicking on the Analyze! | Standard Assignments buttons." & vbCrLf & vbCrLf
msg$ = msg$ & "Also, make sure that the Use Automatic Analysis flag in the Acquisition Options window is not checked, at "
msg$ = msg$ & "least until all standards have been acquired." & vbCrLf & vbCrLf
msg$ = msg$ & "In addition, if the Check For Same Peak Positions or Check For Same PHA Settings options are checked in the Analytical Options dialog, "
msg$ = msg$ & "make sure the on-peak and PHA settings are the same for standard and unknown."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateGetStandards"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateGetStandardsNoIntfSets:
msg$ = "No interference standard intensity data found for standard number " & Format$(sample(1).StdAssignsIntfStds%(j%, i%)) & " for an interference on "
msg$ = msg$ & MiscAutoUcase$(sample(1).Elsyms$(i%)) & " " & sample(1).Xrsyms$(i%) & " by " & sample(1).StdAssignsIntfElements$(j%, i%) & " "
msg$ = msg$ & "on spectrometer " & Format$(sample(1).MotorNumbers%(i%)) & " crystal " & sample(1).CrystalNames$(i%) & " at "
msg$ = msg$ & Format$(sample(1).KilovoltsArray!(i%)) & " KeV for standard interference assignments. Either acquire count data "
msg$ = msg$ & "for the indicated element on the indicated standard at the indicated conditions or "
msg$ = msg$ & "change the standard assignment by clicking on the Analyze! | Standard Assignments buttons."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateGetStandards"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

End Sub

Sub UpdateStdInterferences(analysis As TypeAnalysis, sample() As TypeSample)
' Correct the standard drift array intensities for interferences

ierror = False
On Error GoTo UpdateStdInterferencesError

Dim i As Integer, j As Integer, jmax As Integer, kmax As Integer
Dim chan As Integer, intfchan As Integer, ip As Integer
Dim intfstd As Integer, assignstd As Integer
Dim temp As Single
'Dim temp1 As Single, temp2 As Single

ReDim intcts(1 To MAXSET%, 1 To MAXINTF%, 1 To MAXCHAN%) As Single

If DebugMode Then
Call IOWriteLog(vbCrLf & "Entering UpdateStdInterferences...")
End If

' Check if using interferences
If Not UseInterfFlag Then Exit Sub

' Debug
If DebugMode Then

' Calculate actual maximum number of interference standard sets
kmax% = UpdateGetMaxSet(Int(2), sample())
jmax% = UpdateGetMaxInterfAssign(sample())

For j% = 1 To jmax%
msg$ = InterfSyms(j%) & " Interference standard drift arrays counts (cps/" & Format$(NominalBeam!) & FaradayCurrentUnits$ & ") (MAN corrected):"
Call IOWriteLog(msg$)

For i% = 1 To kmax%
msg$ = vbNullString
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(StdAssignsIntfDriftCounts!(i, j%, chan%), f81$), a80$)
Next chan%
Call IOWriteLog(msg$)
Next i%

Next j%
End If

' Calculate actual maximum number of standard sets
kmax% = UpdateGetMaxSet(Int(1), sample())
jmax% = UpdateGetMaxInterfAssign(sample())

' Correct standard count drift arrays for interferences
For chan% = 1 To sample(1).LastElm%
If sample(1).DisableQuantFlag%(chan%) = 0 Then
'ip% = IPOS8(chan%, sample(1).Elsyms$(chan%), sample(1).Xrsyms$(chan%), sample()) ' find if element is duplicated
ip% = IPOS8A(chan%, sample(1).Elsyms$(chan%), sample(1).Xrsyms$(chan%), sample(1).KilovoltsArray!(chan%), sample()) ' find if element is duplicated
If Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0) Then

assignstd% = IPOS2(NumberofStandards%, sample(1).StdAssigns%(chan%), StandardNumbers%())
If assignstd% = 0 Then GoTo UpdateStdInterferencesBadAssignStd

' Correct for each assigned interference
For j% = 1 To jmax%
If sample(1).StdAssignsIntfStds%(j%, chan%) > 0 Then

' Find the position of the interfering element in the analyzed sample arrays (skip disabled quant elements)
intfchan% = IPOSDQ(sample(1).LastElm%, sample(1).StdAssignsIntfElements$(j%, chan%), sample(1).StdAssignsIntfXrays$(j%, chan%), sample(1).Elsyms$(), sample(1).Xrsyms$(), sample(1).DisableQuantFlag%())
If intfchan% = 0 Then GoTo UpdateStdInterferencesBadIntfElement

' Find the position of the standard used for the interference correction in the standard list
intfstd% = IPOS2(NumberofStandards%, sample(1).StdAssignsIntfStds%(j%, chan%), StandardNumbers%())
If intfstd% = 0 Then GoTo UpdateStdInterferencesBadIntfStandard

' Check for valid weight percents on interference standard
If analysis.StdPercents!(intfstd%, intfchan%) <= 0.01 Then GoTo UpdateStdInterferencesBadIntfPercents

' Correct all standard sets for interferences
For i% = 1 To kmax%

' Check for valid standard drift set
If StdAssignsDateTimes(i%, chan%) <> 0# Then

' Check for valid interference standard drift set
If StdAssignsIntfDateTimes(i%, j%, chan%) <> 0# Then

' Check for valid interference drift counts on interference standard
If StdAssignsIntfDriftCounts!(i%, j%, chan%) = 0# Then GoTo UpdateStdInterferencesNoIntfCounts

' Calculate interference correction on standard drift counts
If DisableFullQuantInterferenceCorrectionFlag = 0 Then
intcts!(i%, j%, chan%) = StdAssignsIntfDriftCounts!(i%, j%, chan%) * analysis.StdPercents!(assignstd%, intfchan%) / analysis.StdPercents!(intfstd%, intfchan%)
Else            ' no need to use Gilfrich since standard concentrations are known (note: StdIntfDriftCounts!() array does not exist yet!)
'temp1! = StdAssignsIntfDriftCounts!(i%, j%, chan%) * StdIntfDriftCounts!(i%, assignstd%, intfchan%) * analysis.StdAssignsPercents(intfchan%)
'temp2! = analysis.StdPercents!(intfstd%, intfchan%) * StdAssignsDriftCounts!(i%, intfchan%)
'If temp2! <> 0# Then intcts!(i%, j%, chan%) = temp1! / temp2!            ' use approximation method of Gilfrich for educational purposes
End If

' Calculate matrix correction for standard interference drift counts (use full correction since fluorescence is same for Ka lines)
If CorrectionFlag% = 0 Then
If analysis.StdZAFCors!(4, assignstd%, chan%) <= 0# Then GoTo UpdateStdInterferencesBadStdCor
temp! = analysis.StdZAFCors!(4, intfstd%, chan%) / analysis.StdZAFCors!(4, assignstd%, chan%)
End If
        
If CorrectionFlag% > 0 And CorrectionFlag% < 5 Then
If analysis.StdBetas!(assignstd%, chan%) <= 0# Then GoTo UpdateStdInterferencesBadStdCor
temp! = analysis.StdBetas!(intfstd%, chan%) / analysis.StdBetas!(assignstd%, chan%)
End If

' Correct interfering counts on standard for matrix correction
If DisableFullQuantInterferenceCorrectionFlag = 0 And DisableMatrixCorrectionInterferenceCorrectionFlag = 0 And sample(1).StdAssignsIntfOrders%(j%, chan%) = 1 Then
intcts!(i%, j%, chan%) = intcts!(i%, j%, chan%) * temp!
End If

' Subtract interference on standard drift arrays
StdAssignsDriftCounts!(i%, chan%) = StdAssignsDriftCounts!(i%, chan%) - intcts!(i%, j%, chan%)

' Check for valid standard count drift arrays
If StdAssignsDriftCounts!(i%, chan%) <= 0# Then GoTo UpdateStdInterferencesNoStdCounts
End If
End If
Next i%

End If
Next j%

End If
End If
Next chan%

' Debug
If DebugMode Then

' Calculate actual maximum number of interference standard sets
kmax% = UpdateGetMaxSet(Int(2), sample())
jmax% = UpdateGetMaxInterfAssign(sample())

For j% = 1 To jmax%
msg$ = InterfSyms$(j%) & " Calculated interference counts from interference drift arrays (cps/" & Format$(NominalBeam!) & FaradayCurrentUnits$ & "):"
Call IOWriteLog(msg$)

For i% = 1 To kmax%
msg$ = vbNullString
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(intcts!(i%, j%, chan%), f81$), a80$)
Next chan%
Call IOWriteLog(msg$)
Next i%

Next j%

msg$ = "Standards Drift Array Counts (cps/" & Format$(NominalBeam!) & FaradayCurrentUnits$ & ") (corrected for MAN/Interference):"
Call IOWriteLog(msg$)

' Calculate actual maximum number of standard sets
kmax% = UpdateGetMaxSet(Int(1), sample())

For i% = 1 To kmax%
msg$ = vbNullString
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(StdAssignsDriftCounts!(i%, chan%), f81$), a80$)
Next chan%
Call IOWriteLog(msg$)
Next i%
End If

Exit Sub

' Errors
UpdateStdInterferencesError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateStdInterferences"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateStdInterferencesBadIntfElement:
msg$ = "Element " & sample(1).StdAssignsIntfElements$(j%, chan%) & " " & sample(1).StdAssignsIntfXrays$(j%, chan%)
msg$ = msg$ & " is an invalid or disabled interfering element on the " & sample(1).Elsyms$(chan%) & " channel"
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateStdInterferences"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateStdInterferencesBadIntfStandard:
msg$ = "Standard number " & Format$(sample(1).StdAssignsIntfStds%(j%, chan%)) & " is an invalid interference standard on " & sample(1).Elsyms$(chan%) & " channel"
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateStdInterferences"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateStdInterferencesBadIntfPercents:
msg$ = "Insufficient weight percent of " & sample(1).StdAssignsIntfElements$(j%, chan%) & " in "
msg$ = msg$ & "interference standard " & Format$(sample(1).StdAssignsIntfStds%(j%, chan%)) & " for "
msg$ = msg$ & " for sample " & Format$(sample(1).number%) & " " & sample(1).Name$ & "."
msg$ = msg$ & vbCrLf & vbCrLf
msg$ = msg$ & "Although an interference of " & sample(1).StdAssignsIntfElements$(j%, chan%) & " on "
msg$ = msg$ & sample(1).Elsyms$(chan%) & " was assigned, there is an insufficient concentration of the interfering element "
msg$ = msg$ & "present. Please change the interference standard used for the interference correction, or "
msg$ = msg$ & "disable the interference correction, using the Standard Assignments button in the ANALYZE! window."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateStdInterferences"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateStdInterferencesBadAssignStd:
msg$ = "Standard " & Format$(sample(1).StdAssigns%(chan%)) & " is an invalid assigned standard on " & sample(1).Elsyms$(chan%) & " channel. Please check the primary standard assignment for this element."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateStdInterferences"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateStdInterferencesNoIntfCounts:
msg$ = "No " & sample(1).StdAssignsIntfElements$(j%, chan%) & " interfering standard counts on interference standard " & Format$(sample(1).StdAssignsIntfStds%(j%, chan%)) & " for "
msg$ = msg$ & sample(1).Elsyms$(chan%) & " in sample " & SampleGetString2$(sample()) & "."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateStdInterferences"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateStdInterferencesBadStdCor:
msg$ = "Invalid standard correction factors on " & sample(1).Elsyms$(chan%) & " channel"
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateStdInterferences"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateStdInterferencesNoStdCounts:
msg$ = "Standard number " & Format$(sample(1).StdAssigns%(chan%)) & " " & sample(1).Elsyms$(chan%) & " counts are zero or less "
msg$ = msg$ & "(" & Format$(StdAssignsDriftCounts!(i%, chan%)) & "). "
msg$ = msg$ & "There may be an interference assigned that does not actually exist."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateStdInterferences"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

End Sub

Sub UpdateStdMANBackgrounds(analysis As TypeAnalysis, sample() As TypeSample)
' Correct standard and interference standard drift arrays for MAN background

ierror = False
On Error GoTo UpdateStdMANBackgroundsError

Dim i As Integer, j As Integer, k As Integer
Dim jmax As Integer, kmax As Integer
Dim chan As Integer, intfstd As Integer
Dim ip As Integer, ipp As Integer
Dim temp As Single
Dim tmsg As String

ReDim stdbac(1 To MAXSET%, 1 To MAXCHAN%) As Single
ReDim intbac(1 To MAXSET%, 1 To MAXINTF%, 1 To MAXCHAN%) As Single

If DebugMode Then
Call IOWriteLog(vbCrLf & "Entering UpdateStdMANBackgrounds...")
End If

' Debug
If DebugMode Then
msg$ = "Standard drift arrays counts (cps/" & Format$(NominalBeam!) & FaradayCurrentUnits$ & ") (uncorrected for MAN/Interference):"
Call IOWriteLog(msg$)

' Calculate actual maximum number of standard sets
kmax% = UpdateGetMaxSet(Int(1), sample())

For k% = 1 To kmax%
msg$ = vbNullString
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(StdAssignsDriftCounts!(k%, chan%), f81$), a80$)
Next chan%
Call IOWriteLog(msg$)
Next k%
End If

' Correct average standard counts for MAN bgd for both before and after
' standardizations. If standard type for this element standardization is an
' off-peak type, don't use MAN background correction. Background type is
' loaded in "UpdateGetNextStandard".

' Calculate actual maximum number of standard sets
kmax% = UpdateGetMaxSet(Int(1), sample())
For chan% = 1 To sample(1).LastElm%
If sample(1).DisableQuantFlag(chan%) = 0 Then

ip% = IPOS2(NumberofStandards%, sample(1).StdAssigns%(chan%), StandardNumbers%())
If ip% = 0 Then GoTo UpdateStdMANBackgroundsNoStandard

' Now load the drift corrected background counts for each standard set and
' then calculate the "analysis.MANFitCoefficients" background correction coefficients
' using least squares fit of MAN background counts vs zbar based on standard
' acquisition time, if this standard intensity is MAN corrected.
For i% = 1 To kmax%
If StdAssignsBackgroundTypes%(i%, chan%) = 1 Then

' Was: Call UpdateCalculateMANDrift(chan%, StdAssignsDateTimes(i%, chan%), analysis)
Call UpdateCalculateMANDrift(Int(1), chan%, Int(0), i%, Int(0), analysis, sample())
If ierror Then Exit Sub

' Now fit the drift corrected MAN counts for this standard
Call UpdateFitMAN(chan%, analysis, sample())
If ierror Then Exit Sub

' Determine MAN background counts on the assigned standard for this element
stdbac(i%, chan%) = analysis.MANFitCoefficients!(1, chan%) + analysis.MANFitCoefficients!(2, chan%) * analysis.StdZbars!(ip%) + analysis.MANFitCoefficients!(3, chan%) * analysis.StdZbars!(ip%) ^ 2

' If correcting for continuum absorption, uncorrect calculated MAN counts for absorption in this
' standard. See "UpdateFitMAN" for absorption correction to fitted background counts.
If UseMANAbsFlag Then
If CorrectionFlag% = 0 Then temp! = analysis.StdZAFCors!(1, ip%, chan%)
If CorrectionFlag% > 0 And CorrectionFlag% < 5 Then temp! = analysis.StdBetas!(ip%, chan%)
If sample(1).MANAbsCorFlags%(chan%) And temp! > 0# Then
stdbac!(i%, chan%) = stdbac!(i%, chan%) / temp!
End If
End If

' Perform the MAN background correction to the standard drift arrays
If StdAssignsDateTimes(i%, chan%) > 0# Then

' Load MAN background counts for UpdateTypeIntensities
StdAssignsDriftBgdCounts!(i%, chan%) = stdbac!(i%, chan%)
StdAssignsDriftCounts!(i%, chan%) = StdAssignsDriftCounts!(i%, chan%) - stdbac!(i%, chan%)

' Check for no counts in MAN corrected standard drift array
If Not UseAggregateIntensitiesFlag Then
If StdAssignsDriftCounts!(i%, chan%) <= 0# Then GoTo UpdateStdMANBackgroundsNegativeStd
Else
'ipp% = IPOS8(chan%, sample(1).Elsyms$(chan%), sample(1).Xrsyms$(chan%), sample())
ipp% = IPOS8A(chan%, sample(1).Elsyms$(chan%), sample(1).Xrsyms$(chan%), sample(1).KilovoltsArray!(chan%), sample())
If ipp% = 0 Then
If StdAssignsDriftCounts!(i%, chan%) <= 0# Then GoTo UpdateStdMANBackgroundsNegativeStd
End If
End If

End If
End If

Next i%
End If
8000:
Next chan%

' Debug
If DebugMode Then
msg$ = "Standard MAN calculated background counts:"
Call IOWriteLog(msg$)

' Calculate actual maximum number of standard sets
kmax% = UpdateGetMaxSet(Int(1), sample())

For k% = 1 To kmax%
msg$ = vbNullString
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(stdbac!(k%, chan%), f81$), a80$)
Next chan%
Call IOWriteLog(msg$)
Next k%

End If

' Debug
If DebugMode Then
msg$ = "Standard drift arrays counts (cps/" & Format$(NominalBeam!) & FaradayCurrentUnits$ & ") (MAN corrected):"
Call IOWriteLog(msg$)

For k% = 1 To kmax%
msg$ = vbNullString
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(StdAssignsDriftCounts!(k%, chan%), f81$), a80$)
Next chan%
Call IOWriteLog(msg$)
Next k%

End If

' Calculate actual maximum number of interference standard sets
kmax% = UpdateGetMaxSet(Int(2), sample())
jmax% = UpdateGetMaxInterfAssign(sample())

' Debug
If DebugMode Then
For j% = 1 To jmax%
msg$ = InterfSyms$(j%) & " Interference standard drift arrays counts (cps/" & Format$(NominalBeam!) & FaradayCurrentUnits$ & ") (uncorrected for MAN):"
Call IOWriteLog(msg$)

For k% = 1 To kmax%
msg$ = vbNullString
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(StdAssignsIntfDriftCounts!(k%, j%, chan%), f81$), a80$)
Next chan%
Call IOWriteLog(msg$)
Next k%

Next j%
End If

' Now correct interference standard drift arrys for MAN background. First see if this channel
' has any interfering elements.
For chan% = 1 To sample(1).LastElm%
If sample(1).DisableQuantFlag%(chan%) = 0 Then     ' skip disable quant

For j% = 1 To jmax%
If sample(1).StdAssignsIntfStds%(j%, chan%) > 0 Then

' Find the position of the standard used for the interference correction in the standard list
intfstd% = IPOS2(NumberofStandards%, sample(1).StdAssignsIntfStds%(j%, chan%), StandardNumbers%())
If intfstd% = 0 Then GoTo UpdateStdMANBackgroundsBadIntfStandard

' Correct interference standard drift arrays for MAN background
For i% = 1 To kmax%

' Now load the drift corrected background counts for each interference
' standard set and then calculate the "analysis.MANFitCoefficients" background
' correction coefficients using least squares fit of MAN background
' counts vs zbar, if this interference standard intensity is MAN corrected.
If StdAssignsIntfBackgroundTypes%(i%, j%, chan%) = 1 Then

' Was: Call UpdateCalculateMANDrift(chan%, StdAssignsIntfDateTimes(i%, j%, chan%), analysis)
Call UpdateCalculateMANDrift(Int(2), chan%, Int(0), i%, j%, analysis, sample())
If ierror Then Exit Sub

' Now fit the drift corrected MAN counts for this standard time. Counts are
' corrected for continuum absorption in UpdateFitMAN if flagged for this sample
' and element. Counts are UNcorrected for continuum absorption below before
' background correction.
Call UpdateFitMAN(chan%, analysis, sample())
If ierror Then Exit Sub

' Determine MAN background counts on the standard for this element
intbac(i%, j%, chan%) = analysis.MANFitCoefficients!(1, chan%) + analysis.MANFitCoefficients!(2, chan%) * analysis.StdZbars!(intfstd%) + analysis.MANFitCoefficients!(3, chan%) * analysis.StdZbars!(intfstd%) ^ 2

' If correcting for continuum absorption, uncorrect calculated MAN counts for absorption in this
' standard. See "UpdateFitMAN" for absorption correction to fitted background counts.
If UseMANAbsFlag Then
If CorrectionFlag% = 0 Then temp! = analysis.StdZAFCors!(1, intfstd%, chan%)
If CorrectionFlag% > 0 And CorrectionFlag% < 5 Then temp! = analysis.StdBetas!(intfstd%, chan%)
If sample(1).MANAbsCorFlags%(chan%) And temp! > 0# Then
intbac!(i%, j%, chan%) = intbac!(i%, j%, chan%) / temp!
End If
End If

' Perform the MAN background correction to the interference standard drift arrays
If StdAssignsIntfDateTimes(i, j%, chan%) > 0# Then
StdAssignsIntfDriftCounts!(i%, j%, chan%) = StdAssignsIntfDriftCounts!(i%, j%, chan%) - intbac!(i%, j%, chan%)

' Check for no counts in MAN corrected interference standards
If Not UseAggregateIntensitiesFlag Then
If StdAssignsIntfDriftCounts!(i%, j%, chan%) < -2# Then GoTo UpdateStdMANBackgroundsNegativeInterf
If StdAssignsIntfDriftCounts!(i%, j%, chan%) <= 0# Then
tmsg$ = "Warning- Zero or less interference counts for " & sample(1).Elsyms$(chan%) & " on standard " & Format$(sample(1).StdAssignsIntfStds%(j%, chan%))
Call IOWriteLogRichText(tmsg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
End If

Else
'ipp% = IPOS8(chan%, sample(1).Elsyms$(chan%), sample(1).Xrsyms$(chan%), sample())
ipp% = IPOS8A(chan%, sample(1).Elsyms$(chan%), sample(1).Xrsyms$(chan%), sample(1).KilovoltsArray!(chan%), sample())
If ipp% = 0 Then
If StdAssignsIntfDriftCounts!(i%, j%, chan%) < -2# Then GoTo UpdateStdMANBackgroundsNegativeInterf
If StdAssignsIntfDriftCounts!(i%, j%, chan%) <= 0# Then
tmsg$ = "Warning- Zero or less interference counts for " & sample(1).Elsyms$(chan%) & " on standard " & Format$(sample(1).StdAssignsIntfStds%(j%, chan%))
Call IOWriteLogRichText(tmsg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
End If
End If
End If

End If
End If

Next i%

End If
Next j%

End If
Next chan%

' Debug
If DebugMode Then
For j% = 1 To jmax%
msg$ = InterfSyms$(j%) & " Interference standard MAN calculated background counts:"
Call IOWriteLog(msg$)

For k% = 1 To kmax%
msg$ = vbNullString
For chan% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(intbac!(k%, j%, chan%), f81$), a80$)
Next chan%
Call IOWriteLog(msg$)
Next k%

Next j%
End If

Exit Sub

' Errors
UpdateStdMANBackgroundsError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateStdMANBackgrounds"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateStdMANBackgroundsNoStandard:
msg$ = "Standard number " & Format$(sample(1).StdAssigns%(chan%)) & " is not in the run"
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateStdMANBackgrounds"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateStdMANBackgroundsNegativeStd:
msg$ = "Standard number " & Format$(sample(1).StdAssigns%(chan%)) & " (set " & Format$(i%) & ") for " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " MAN background corrected counts are zero or negative on channel " & Format$(chan%)
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateStdMANBackgrounds"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateStdMANBackgroundsBadIntfStandard:
msg$ = "Standard number " & Format$(sample(1).StdAssignsIntfStds%(j%, chan%)) & " is an invalid interference standard on " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " on channel " & Format$(chan%)
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateStdMANBackgrounds"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateStdMANBackgroundsNegativeInterf:
msg$ = "Interference standard number " & Format$(sample(1).StdAssignsIntfStds%(j%, chan%)) & ", for "
msg$ = msg$ & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " MAN background corrected counts are very negative "
msg$ = msg$ & "(" & Format$(StdAssignsIntfDriftCounts!(i%, j%, chan%)) & ") on channel " & Format$(chan%) & ". "
msg$ = msg$ & "There may be an interference assigned that does not actually exist."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateStdMANBackgrounds"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

End Sub

Sub UpdateTypeIntensities(analysis As TypeAnalysis, sample() As TypeSample, stdsample() As TypeSample)
' Routine to type out standard intensities for a specific sample

ierror = False
On Error GoTo UpdateTypeIntensitiesError

Dim i As Integer, j As Integer, k As Integer
Dim imax As Integer, row As Integer

' Load analysis arrays (force re-load)
AllAnalysisUpdateNeeded = True
Call UpdateAnalyze(analysis, sample(), stdsample())
If ierror Then Exit Sub

' Type out assigned drift array intensities for the passed sample
row% = SampleGetRow%(sample())
msg$ = SampleGetString$(row%)
msg$ = vbCrLf & "Assigned average standard intensities for sample " & msg$
Call IOWriteLog(msg$)

msg$ = vbCrLf & "Drift array background intensities (cps/" & Format$(NominalBeam!) & FaradayCurrentUnits$ & ") for standards:"
Call IOWriteLog(msg$)

msg$ = "ELMXRY: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "MOTCRY: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).MotorNumbers%(i%), a20$) & Format$(sample(1).CrystalNames$(i%), a60$)
Next i%
Call IOWriteLog(msg$)

msg$ = "INTEGR: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).IntegratedIntensitiesUseIntegratedFlags%(i%), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "STDASS: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).StdAssigns%(i%), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "STDVIR: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).StdAssignsFlag%(i%), a80$)
Next i%
Call IOWriteLog(msg$)

' Type standard background counts for this sample
imax% = UpdateGetMaxSet%(Int(1), sample())
If imax% = 0 Then GoTo UpdateTypeIntensitiesNoStdSets

' Type all sets loaded
For k% = 1 To imax%
msg$ = Space$(8)

For i% = 1 To sample(1).LastElm%
If StdAssignsDriftBgdCounts!(k%, i%) = NOTLOADEDVALUE! Then
msg$ = msg$ & Format$("      - ", a80$)
Else
If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(StdAssignsDriftBgdCounts!(k%, i%), f81$), a80$)
Else
msg$ = msg$ & Format$(Format$(StdAssignsDriftBgdCounts!(k%, i%), f82$), a80$)
End If
End If
Next i%
Call IOWriteLog(msg$)

Next k%

' Type out assigned drift array intensities for the passed sample
If UseVolElFlag And AcquireVolatileSelfStandardIntensitiesFlag Then
msg$ = vbCrLf & "Drift array standard intensities (cps/" & Format$(NominalBeam!) & FaradayCurrentUnits$ & ") (background corrected) (TDI corrected):"
Else
msg$ = vbCrLf & "Drift array standard intensities (cps/" & Format$(NominalBeam!) & FaradayCurrentUnits$ & ") (background corrected):"
End If
Call IOWriteLog(msg$)

msg$ = "ELMXRY: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "MOTCRY: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).MotorNumbers%(i%), a20$) & Format$(sample(1).CrystalNames$(i%), a60$)
Next i%
Call IOWriteLog(msg$)

msg$ = "STDASS: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).StdAssigns%(i%), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "STDVIR: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).StdAssignsFlag%(i%), a80$)
Next i%
Call IOWriteLog(msg$)

' Type standard counts for this sample
imax% = UpdateGetMaxSet%(Int(1), sample())
If imax% = 0 Then GoTo UpdateTypeIntensitiesNoStdSets

' Type all sets loaded
For k% = 1 To imax%
msg$ = Space$(8)

For i% = 1 To sample(1).LastElm%
If StdAssignsDriftCounts!(k%, i%) = NOTLOADEDVALUE! Then
msg$ = msg$ & Format$("      - ", a80$)
Else
If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(StdAssignsDriftCounts!(k%, i%), f81$), a80$)
Else
msg$ = msg$ & Format$(Format$(StdAssignsDriftCounts!(k%, i%), f82$), a80$)
End If
End If
Next i%
Call IOWriteLog(msg$)

Next k%

' Warn if using virtual standards
Call IOWriteLog(vbNullString)
For i% = 1 To sample(1).LastElm%
If sample(1).StdAssignsFlag%(i%) = 1 Then
msg$ = "WARNING- Using Virtual Standard Intensity For " & sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%)
Call IOWriteLog(msg$)
End If
Next i%

' Get maximum sets of interference stds
imax% = UpdateGetMaxSet%(Int(2), sample())

' Type standard interference counts for this sample
If UseInterfFlag And imax% > 0 Then
msg$ = vbCrLf & "Drift array interference standard intensities (cps/" & Format$(NominalBeam!) & FaradayCurrentUnits$ & "):"
Call IOWriteLog(msg$)

For j% = 1 To MAXINTF%
If Not MiscAllZero2(j%, sample(1).LastElm%, sample(1).StdAssignsIntfStds%()) Then
msg$ = vbCrLf & InterfSyms$(j%) & " assigned interference elements"
Call IOWriteLog(msg$)

msg$ = "ELMXRY: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "INTFELM:"
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).StdAssignsIntfElements$(j%, i%) & " ", a80$)
Next i%
Call IOWriteLog(msg$)

If ProbeDataFileVersionNumber! > 6.41 And MiscIsElementDuplicated(sample()) Then
msg$ = "INTFXRY:"
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).StdAssignsIntfXrays$(j%, i%) & " ", a80$)
Next i%
Call IOWriteLog(msg$)
End If

msg$ = "INTFSTD:"
For i% = 1 To sample(1).LastElm%
If StdAssignsIntfDriftCounts!(1, j%, i%) = NOTLOADEDVALUE! Then ' use drift counts for load check
msg$ = msg$ & Format$("        ", a80$)
Else
msg$ = msg$ & Format$(sample(1).StdAssignsIntfStds%(j%, i%), a80$)
End If
Next i%
Call IOWriteLog(msg$)

' Type all sets loaded
For k% = 1 To imax%
msg$ = Space$(8)

For i% = 1 To sample(1).LastElm%
If StdAssignsIntfDriftCounts!(k%, j%, i%) = NOTLOADEDVALUE! Then
msg$ = msg$ & Format$("        ", a80$)
Else
If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(StdAssignsIntfDriftCounts!(k%, j%, i%), f81$), a80$)
Else
msg$ = msg$ & Format$(Format$(StdAssignsIntfDriftCounts!(k%, j%, i%), f82$), a80$)
End If
End If
Next i%
Call IOWriteLog(msg$)

Next k%
End If
Next j%
End If

Exit Sub

' Errors
UpdateTypeIntensitiesError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateTypeIntensities"
ierror = True
Exit Sub

UpdateTypeIntensitiesNoStdSets:
row% = SampleGetRow%(sample())
msg$ = SampleGetString$(row%)
msg$ = "No assigned standard data for sample " & msg$ & ". Assign standards first and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateTypeIntensities"
ierror = True
Exit Sub

End Sub

Sub UpdateCalculateDrift(row As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' Calculate the drift correction using linear interpolation, "row" points to the sample data line

ierror = False
On Error GoTo UpdateCalculateDriftError

Dim i As Integer, j As Integer
Dim ip As Integer

' Initialize drift corrected arrays
For i% = 1 To MAXCHAN%
analysis.StdAssignsCounts!(i%) = 0#
analysis.StdAssignsRows%(i%) = 0
For j% = 1 To MAXINTF%
analysis.StdAssignsIntfCounts!(j%, i%) = 0#
analysis.StdAssignsIntfRows%(j%, i%) = 0
Next j%
For j% = 1 To MAXMAN%
analysis.MANAssignsCounts!(j%, i%) = 0#
analysis.MANAssignsRows%(j%, i%) = 0
Next j%
Next i%

' Loop on each element and calculate standard and interference standard drift corrected counts
For i% = 1 To sample(1).LastElm%
If sample(1).DisableQuantFlag%(i%) = 0 Then     ' for aggregate printout
Call UpdateCalculateStdDrift(i%, row%, analysis, sample())
If ierror Then Exit Sub
End If
Next i%

' Check that no standard counts are zero
For i% = 1 To sample(1).LastElm%
If sample(1).DisableQuantFlag%(i%) = 0 Then ' skip check if quant is disabled

If Not UseAggregateIntensitiesFlag Then
If analysis.StdAssignsCounts!(i%) <= 0# Then GoTo UpdateCalculateDriftNoCounts
Else
'ip% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())
ip% = IPOS8A(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample(1).KilovoltsArray!(i%), sample())
If ip% = 0 Then
If analysis.StdAssignsCounts!(i%) <= 0# Then GoTo UpdateCalculateDriftNoCounts
End If
End If

End If
Next i%

' Loop on each element and calculate MAN standard drift correction
For i% = 1 To sample(1).LastElm%

' Check if this is a MAN correction at all (0=off-peak, 1=MAN, 2=multipoint)
If sample(1).BackgroundTypes%(i%) = 1 Then
Call UpdateCalculateMANDrift(Int(3), i%, row%, Int(0), Int(0), analysis, sample())
If ierror Then Exit Sub
End If

Next i%

Exit Sub

' Errors
UpdateCalculateDriftError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateCalculateDrift"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateCalculateDriftNoCounts:
msg$ = "Insufficient standard counts on standard " & Format$(sample(1).StdAssigns%(i%)) & " for "
msg$ = msg$ & sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & " on spectrometer "
msg$ = msg$ & Format$(sample(1).MotorNumbers%(i%)) & " using crystal " & sample(1).CrystalNames$(i%) & ". "
msg$ = msg$ & "Make sure that valid data for the indicated standard has been acquired at "
msg$ = msg$ & Format$(sample(1).KilovoltsArray!(i%)) & " kilovolts and " & Format$(sample(1).takeoff!) & " takeoff angle."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateCalculateDrift"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

End Sub

Sub UpdateCalculateMANDrift(mode As Integer, chan As Integer, row As Integer, setnumber As Integer, interfnumber As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' Calculate drift corrected counts for MAN correction of standards, interference stds and unknowns.
' mode = 1 calculate MAN counts for primary stds
' mode = 2 calculate MAN counts for interference stds
' mode = 3 calculate MAN counts for unknowns
' Called by UpdateStandardMANBackground for standards and interference standards and
' by UpdateCalculateDrift for unknowns. Note that because the acqusition datetimes
' of the standards can be different from the unknowns, this routine uses an array
' of all assigned MAN standard sets in the run to accurately calculate the MAN
' drift for each acquisition time.

ierror = False
On Error GoTo UpdateCalculateMANDriftError

Dim j As Integer, iset As Integer
Dim deltatime As Double, elapsedtime As Double
Dim deltacounts As Single
Dim tDateTime As Variant

ReDim set1(1 To MAXMAN%) As Integer, set2(1 To MAXMAN%) As Integer

' Determine the first set for each MAN background standard taken before this point
For j% = 1 To MAXMAN%
set1%(j%) = 1
Next j%

' Calculate drift for primary assigned standards
If mode% = 1 Then
tDateTime = StdAssignsDateTimes(setnumber%, chan%)

' Calculate drift for primary assigned interference standards
ElseIf mode% = 2 Then
tDateTime = StdAssignsIntfDateTimes(setnumber%, interfnumber%, chan%)

' Calculate drift for unknowns (point or pixels)
ElseIf mode% = 3 Then

' Load acquisition time from single point analyses time for PFE and Evaluate
If UCase$(Trim$(app.EXEName)) <> UCase$(Trim$("CalcImage")) Then
tDateTime = sample(1).DateTimes(row%)

' Load acquisition time for x-ray map pixels from image start/stop map acquisition times for CalcImage
Else
tDateTime = ConditionSampleDateTime(chan%)
End If
End If

For iset% = 2 To MAXSET%
For j% = 1 To MAXMAN%
If MANAssignsDateTimes(iset%, j%, chan%) <> 0# Then
If tDateTime >= MANAssignsDateTimes(iset%, j%, chan%) Then set1%(j%) = iset%
End If
Next j%
Next iset%

' Determine the last set of MAN background standards taken after this point
For j% = 1 To MAXMAN%
set2%(j%) = 0
Next j%

For iset% = MAXSET% To 1 Step -1
For j% = 1 To MAXMAN%
If tDateTime < MANAssignsDateTimes(iset%, j%, chan%) Then set2%(j%) = iset%
Next j%
Next iset%

' Check that valid drift sets were found
For j% = 1 To MAXMAN%
If set1%(j%) >= set2%(j%) Then set2%(j%) = 0
Next j%

' Load first (last actually) set MAN counts as default
For j% = 1 To MAXMAN%
analysis.MANAssignsCounts!(j%, chan%) = MANAssignsDriftCounts!(set1%(j%), j%, chan%)
analysis.MANAssignsRows%(j%, chan%) = MANAssignsSampleRows%(set1%(j%), j%, chan%)

analysis.MANAssignsCountTimes!(j%, chan%) = MANAssignsCountTimes!(set1%(j%), j%, chan%)
analysis.MANAssignsBeamCurrents!(j%, chan%) = MANAssignsBeamCurrents!(set1%(j%), j%, chan%)

' If a subsequent set was found, then calculate drift based on elasped time from passed DateTime parameter
If UseDriftFlag And set2%(j%) <> 0 Then
deltacounts! = MANAssignsDriftCounts!(set2%(j%), j%, chan%) - MANAssignsDriftCounts!(set1%(j%), j%, chan%)
deltatime = MANAssignsDateTimes(set2%(j%), j%, chan%) - MANAssignsDateTimes(set1%(j%), j%, chan%)
elapsedtime = tDateTime - MANAssignsDateTimes(set1%(j%), j%, chan%)

If deltatime <> 0# Then
analysis.MANAssignsCounts!(j%, chan%) = MANAssignsDriftCounts!(set1%(j%), j%, chan%) + deltacounts! * elapsedtime / deltatime
End If
End If

Next j%

Exit Sub

' Errors
UpdateCalculateMANDriftError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateCalculateMANDrift"
ierror = True
Exit Sub

End Sub

Sub UpdateCalculateStdDrift(chan As Integer, row As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' Calculate drift corrected standard and interference counts for standards and unknowns

ierror = False
On Error GoTo UpdateCalculateStdDriftError

Dim j As Integer, iset As Integer
Dim deltatime As Double, elapsedtime As Double
Dim deltacounts As Single, deltatimes As Single
Dim deltabeams As Single, temp As Single
Dim deltabgdcounts As Single
Dim std1 As Integer, std2 As Integer
Dim tmsg As String
Dim tDateTime As Variant

ReDim set1(1 To MAXINTF%) As Integer, set2(1 To MAXINTF%) As Integer

' Init the drift set numbers
std1% = 1                   ' primary standard
For j% = 1 To MAXINTF%      ' interference standards
set1%(j%) = 1
Next j%

' Load acquisition time from single point analyses time for PFE and Evaluate
If UCase$(Trim$(app.EXEName)) <> UCase$(Trim$("CalcImage")) Then
tDateTime = sample(1).DateTimes(row%)

' Load acquisition time for x-ray map pixels from image start/stop map acquisition times for CalcImage
Else
tDateTime = ConditionSampleDateTime(chan%)
End If

' Determine the first set for each standard and interference standard taken before this point
For iset% = 2 To MAXSET%
If StdAssignsDateTimes(iset%, chan%) <> 0# Then
If tDateTime >= StdAssignsDateTimes(iset%, chan%) Then std1% = iset%
End If
For j% = 1 To MAXINTF%
If StdAssignsIntfDateTimes(iset%, j%, chan%) <> 0# Then
If tDateTime >= StdAssignsIntfDateTimes(iset%, j%, chan%) Then set1%(j%) = iset%
End If
Next j%
Next iset%

' Determine the last set of standards taken after this point
std2% = 0                   ' primary standard
For j% = 1 To MAXINTF%      ' interference standards
set2%(j%) = 0
Next j%

For iset% = MAXSET% To 1 Step -1
If tDateTime < StdAssignsDateTimes(iset%, chan%) Then std2% = iset%
For j% = 1 To MAXINTF%
If tDateTime < StdAssignsIntfDateTimes(iset%, j%, chan%) Then set2%(j%) = iset%
Next j%
Next iset%

' Check that valid drift sets were found
If std1% >= std2% Then std2% = 0
For j% = 1 To MAXINTF%
If set1%(j%) >= set2%(j%) Then set2%(j%) = 0
Next j%

' Load standard counts (if not virtual standard intensity)
If sample(1).StdAssignsFlag%(chan%) = 0 Then

' Load first (last actually) set standard counts as default
analysis.StdAssignsCounts!(chan%) = StdAssignsDriftCounts!(std1%, chan%)
analysis.StdAssignsTimes!(chan%) = StdAssignsDriftTimes!(std1%, chan%)
analysis.StdAssignsBeams!(chan%) = StdAssignsDriftBeams!(std1%, chan%)
analysis.StdAssignsRows%(chan%) = StdAssignsSampleRows%(std1%, chan%)
analysis.StdAssignsBgdCounts!(chan%) = StdAssignsDriftBgdCounts!(std1%, chan%)

' If a subsequent set was found, then calculate drift based on elasped time from passed date/time for this sample row
If UseDriftFlag And std2% <> 0 Then
deltacounts! = StdAssignsDriftCounts!(std2%, chan%) - StdAssignsDriftCounts!(std1%, chan%)
deltatimes! = StdAssignsDriftTimes!(std2%, chan%) - StdAssignsDriftTimes!(std1%, chan%)
deltabeams! = StdAssignsDriftBeams!(std2%, chan%) - StdAssignsDriftBeams!(std1%, chan%)
deltatime = StdAssignsDateTimes(std2%, chan%) - StdAssignsDateTimes(std1%, chan%)
elapsedtime = tDateTime - StdAssignsDateTimes(std1%, chan%)

' Load interpolated std bgd counts
deltabgdcounts! = StdAssignsDriftBgdCounts!(std2%, chan%) - StdAssignsDriftBgdCounts!(std1%, chan%)

' Calculate drift and print warning if greater than tolerance (and more than an hour has elapsed)
If StdAssignsDriftCounts!(std2%, chan%) <> 0# And deltatime <> 0# And deltatime * HOURPERDAY# > 1# Then
temp! = (deltacounts! / StdAssignsDriftCounts!(std2%, chan%)) / (deltatime * HOURPERDAY#)
If Abs(100# * temp!) > 2# Then      ' warn if greater than 2% per hour
tmsg$ = "Warning: standard drift is " & Format$(100# * temp!, f82$) & "% per hour for " & MiscAutoUcase$(analysis.Elsyms$(chan%)) & " " & analysis.Xrsyms$(chan%) & " on spectrometer " & Format$(analysis.MotorNumbers%(chan%))
Call IOWriteLogRichText(tmsg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
End If
End If

If deltatime <> 0# Then
analysis.StdAssignsCounts!(chan%) = StdAssignsDriftCounts!(std1%, chan%) + deltacounts! * elapsedtime / deltatime
analysis.StdAssignsTimes!(chan%) = StdAssignsDriftTimes!(std1%, chan%) + deltatimes! * elapsedtime / deltatime
analysis.StdAssignsBeams!(chan%) = StdAssignsDriftBeams!(std1%, chan%) + deltatimes! * elapsedtime / deltatime
analysis.StdAssignsBgdCounts!(chan%) = StdAssignsDriftBgdCounts!(std1%, chan%) + deltabgdcounts! * elapsedtime / deltatime

' Print interpolated date/time
If DebugMode And VerboseMode Then
Call IOWriteLog("UpdateCalculateStdDrift, Std1: " & Format$(StdAssignsDriftCounts!(std1%, chan%)) & ", Std2: " & Format$(StdAssignsDriftCounts!(std2%, chan%)) & ", Interpolated: " & Format$(analysis.StdAssignsCounts!(chan%)))
End If

End If
End If

' Load virtual standard intensity from database table (moved here to allow interference counts to be loaded below)
Else
Call VirtualCalculate2(chan%, sample(), analysis.StdAssignsCounts!(chan%))
If ierror Then Exit Sub
End If

' Load first set interference standard counts as default (must not be virtual!)
For j% = 1 To MAXINTF%
analysis.StdAssignsIntfCounts!(j%, chan%) = StdAssignsIntfDriftCounts!(set1%(j%), j%, chan%)
analysis.StdAssignsIntfRows%(j%, chan%) = StdAssignsIntfSampleRows%(set1%(j%), j%, chan%)

' If a subsequent set was found, then calculate drift based on elasped time from passed data/time for this sample row
If UseDriftFlag And set2%(j%) <> 0 Then
deltacounts! = StdAssignsIntfDriftCounts!(set2%(j%), j%, chan%) - StdAssignsIntfDriftCounts!(set1%(j%), j%, chan%)
deltatime = StdAssignsIntfDateTimes(set2%(j%), j%, chan%) - StdAssignsIntfDateTimes(set1%(j%), j%, chan%)
elapsedtime = tDateTime - StdAssignsIntfDateTimes(set1%(j%), j%, chan%)

If deltatime <> 0# Then
analysis.StdAssignsIntfCounts!(j%, chan%) = StdAssignsIntfDriftCounts!(set1%(j%), j%, chan%) + deltacounts! * elapsedtime / deltatime
End If
End If

Next j%

Exit Sub

' Errors
UpdateCalculateStdDriftError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateCalculateStdDrift"
ierror = True
Exit Sub

End Sub

Function UpdateChangedSample(sample() As TypeSample) As Integer
' Checks to see if the sample parameters have changed since the last time called. Check for
' changed standard assignments (No need to check for changed MAN assignments
' since any specific element MAM assignment is global. As long as the element doesn't
' change, then the MAN assignment is the same). However, we do need to check for change
' in background correction type.

ierror = False
On Error GoTo UpdateChangedSampleError

Dim i As Integer, j As Integer, row As Integer

' Preserve global values for comparison
Static lastsamplerow As Integer
Static lastsampleline As Long
Static lastnumberofstandards As Integer

UpdateChangedSample = False

' First check if a new standard data line has been acquired
If lastsamplerow% > 0 Then
For row% = lastsamplerow% To NumberofSamples%
If lastsampleline& <> NumberofLines& And SampleTyps%(row%) = 1 Then UpdateChangedSample = True
Next row%
End If

' Check if a standard has been added
If lastnumberofstandards% <> NumberofStandards% Then UpdateChangedSample = True

' Check conditions
If Not sample(1).CombinedConditionsFlag Then
If sample(1).takeoff! <> UpdateOldSample(1).takeoff! Then UpdateChangedSample = True
If sample(1).kilovolts! <> UpdateOldSample(1).kilovolts! Then UpdateChangedSample = True
End If

' Check for number of analyzed elements changed (in case the user just deleted some from the end) (this is needed in ProbFormTypeIntensities)
If sample(1).LastElm% <> UpdateOldSample(1).LastElm% Then UpdateChangedSample = True
If sample(1).LastChan% <> UpdateOldSample(1).LastChan% Then UpdateChangedSample = True

' Check for use integrated intensity and specify matrix by analysis flags
If sample(1).IntegratedIntensitiesUseFlag% <> UpdateOldSample(1).IntegratedIntensitiesUseFlag% Then UpdateChangedSample = True
If sample(1).SpecifyMatrixByAnalysisUnknownNumber% <> UpdateOldSample(1).SpecifyMatrixByAnalysisUnknownNumber% Then UpdateChangedSample = True

' See if sample has changed since last time
For i% = 1 To sample(1).LastChan%

' Check conditions
If sample(1).CombinedConditionsFlag Then
If sample(1).TakeoffArray!(i%) <> UpdateOldSample(1).TakeoffArray!(i%) Then UpdateChangedSample = True
If sample(1).KilovoltsArray!(i%) <> UpdateOldSample(1).KilovoltsArray!(i%) Then UpdateChangedSample = True
End If

If sample(1).Elsyms$(i%) <> UpdateOldSample(1).Elsyms$(i%) Then UpdateChangedSample = True
If sample(1).Xrsyms$(i%) <> UpdateOldSample(1).Xrsyms$(i%) Then UpdateChangedSample = True
If sample(1).MotorNumbers%(i%) <> UpdateOldSample(1).MotorNumbers%(i%) Then UpdateChangedSample = True
If sample(1).CrystalNames$(i%) <> UpdateOldSample(1).CrystalNames$(i%) Then UpdateChangedSample = True

If sample(1).IntegratedIntensitiesUseIntegratedFlags%(i%) <> UpdateOldSample(1).IntegratedIntensitiesUseIntegratedFlags%(i%) Then UpdateChangedSample = True

' Check standard assignments (analyzed elements only) for changes
If sample(1).StdAssigns%(i%) <> UpdateOldSample(1).StdAssigns%(i%) Then UpdateChangedSample = True
For j% = 1 To MAXINTF%  ' check all interfering arrays
If sample(1).StdAssignsIntfElements$(j%, i%) <> UpdateOldSample(1).StdAssignsIntfElements$(j%, i%) Then UpdateChangedSample = True
If sample(1).StdAssignsIntfXrays$(j%, i%) <> UpdateOldSample(1).StdAssignsIntfXrays$(j%, i%) Then UpdateChangedSample = True
If sample(1).StdAssignsIntfStds%(j%, i%) <> UpdateOldSample(1).StdAssignsIntfStds%(j%, i%) Then UpdateChangedSample = True
Next j%

' Check for change in background correction type (0=off-peak, 1=MAN, 2=multipoint)
If sample(1).BackgroundTypes%(i%) <> UpdateOldSample(1).BackgroundTypes%(i%) Then UpdateChangedSample = True

' Check for change in peak position
If AnalysisCheckForSamePeakPositions Then
If Not MiscDifferenceIsSmall(sample(1).OnPeaks!(i%), UpdateOldSample(1).OnPeaks!(i%), 0.00005) Then UpdateChangedSample = True
If Not MiscDifferenceIsSmall(sample(1).HiPeaks!(i%), UpdateOldSample(1).HiPeaks!(i%), 0.00005) Then UpdateChangedSample = True
If Not MiscDifferenceIsSmall(sample(1).LoPeaks!(i%), UpdateOldSample(1).LoPeaks!(i%), 0.00005) Then UpdateChangedSample = True
End If

' Check for change in PHA settings
If AnalysisCheckForSamePHASettings Then
If Not MiscDifferenceIsSmall(sample(1).Baselines!(i%), UpdateOldSample(1).Baselines!(i%), 0.005) Then UpdateChangedSample = True
If Not MiscDifferenceIsSmall(sample(1).Windows!(i%), UpdateOldSample(1).Windows!(i%), 0.005) Then UpdateChangedSample = True
If Not MiscDifferenceIsSmall(sample(1).Gains!(i%), UpdateOldSample(1).Gains!(i%), 0.005) Then UpdateChangedSample = True
If Not MiscDifferenceIsSmall(sample(1).Biases!(i%), UpdateOldSample(1).Biases!(i%), 0.005) Then UpdateChangedSample = True
If sample(1).InteDiffModes%(i%) <> UpdateOldSample(1).InteDiffModes%(i%) Then UpdateChangedSample = True
End If

' Check for calculation flags
If sample(1).numcat%(i%) <> UpdateOldSample(1).numcat%(i%) Then UpdateChangedSample = True
If sample(1).numoxd%(i%) <> UpdateOldSample(1).numoxd%(i%) Then UpdateChangedSample = True

If sample(1).BlankCorrectionUnks%(i%) <> UpdateOldSample(1).BlankCorrectionUnks%(i%) Then UpdateChangedSample = True
If sample(1).BlankCorrectionLevels!(i%) <> UpdateOldSample(1).BlankCorrectionLevels!(i%) Then UpdateChangedSample = True

If sample(1).DisableAcqFlag%(i%) <> UpdateOldSample(1).DisableAcqFlag%(i%) Then UpdateChangedSample = True
If sample(1).DisableQuantFlag%(i%) <> UpdateOldSample(1).DisableQuantFlag%(i%) Then UpdateChangedSample = True
Next i%

' Save globals for next time
lastsamplerow% = NumberofSamples%
lastsampleline& = NumberofLines&
lastnumberofstandards% = NumberofStandards%

' Save sample setup for next time
UpdateOldSample(1) = sample(1)

Exit Function

' Errors
UpdateChangedSampleError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateChangedSample"
ierror = True
Exit Function

End Function

Function UpdateChangedCorrection() As Integer
' Checks to see if the correction parameters have changed since the last time called. Only used for CalculateAllMatrixCorrections flag)

ierror = False
On Error GoTo UpdateChangedCorrectionError

' Preserve global values for comparison
Static lastcorrectionflag As Integer
Static lastempiricalalphaflag As Integer
Static lastizaf As Integer

UpdateChangedCorrection = False

' 0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters
If CorrectionFlag% <> lastcorrectionflag% Then UpdateChangedCorrection = True
If EmpiricalAlphaFlag% <> lastempiricalalphaflag% Then UpdateChangedCorrection = True
If izaf% <> lastizaf Then UpdateChangedCorrection = True

' Save globals for next time
lastcorrectionflag% = CorrectionFlag%
lastempiricalalphaflag% = EmpiricalAlphaFlag%
lastizaf% = izaf%

Exit Function

' Errors
UpdateChangedCorrectionError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateChangedCorrection"
ierror = True
Exit Function

End Function

Sub UpdateFitMAN(chan As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' Routine to fit the drift corrected MAN counts

ierror = False
On Error GoTo UpdateFitMANError

Dim npts As Integer, i As Integer
Dim ip As Integer, row As Integer
Dim norder As Integer
Dim temp As Single

ReDim acoeff(1 To MAXCOEFF%) As Single
ReDim txdata(1 To MAXMAN%) As Single
ReDim tydata(1 To MAXMAN%) As Single
ReDim abscor(1 To MAXMAN%) As Single

' Load the sample counts loaded by UPDATE
npts% = 0
For i% = 1 To MAXMAN%
If sample(1).MANStdAssigns%(i%, chan%) > 0 Then

' Find MAN in standard list
ip% = IPOS2(NumberofStandards%, sample(1).MANStdAssigns%(i%, chan%), StandardNumbers%())
If ip% = 0 Then GoTo UpdateFitMANBadStandard

' Load z-bar and counts
npts% = npts% + 1
txdata!(npts%) = analysis.StdZbars!(ip%)
tydata!(npts%) = analysis.MANAssignsCounts!(i%, chan%)

' Load correction for continuum absorption (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% = 0 Then temp! = analysis.StdZAFCors!(1, ip%, chan%) ' use characteristic for continuum (MAN)
If CorrectionFlag% > 0 And CorrectionFlag% < 5 Then temp! = analysis.StdBetas!(ip%, chan%) ' use characteristic for continuum (MAN)
'temp! = analysis.StdContinuums!(ip%, chan%)
abscor!(npts%) = temp!

End If
Next i%

If DebugMode Then
msg$ = vbCrLf & "MAN fit data and coefficients for " & sample(1).Elsyms$(chan%)
Call IOWriteLog(msg$)

msg$ = "Order " & Format$(sample(1).MANLinearFitOrders%(chan%)) & ", npts " & Format$(npts%)
Call IOWriteLog(msg$)

msg$ = "StdAss: "
For i% = 1 To npts%
msg$ = msg$ & MiscAutoFormatI$(sample(1).MANStdAssigns%(i%, chan%))
Next i%
Call IOWriteLog(msg$)

msg$ = "Z-bars: "
For i% = 1 To npts%
msg$ = msg$ & MiscAutoFormat$(txdata!(i%))
Next i%
Call IOWriteLog(msg$)

msg$ = "Counts: "
For i% = 1 To npts%
msg$ = msg$ & MiscAutoFormat$(tydata!(i%))
Next i%
Call IOWriteLog(msg$)

msg$ = "AbsCor: "
For i% = 1 To npts%
msg$ = msg$ & MiscAutoFormat$(abscor!(i%))
Next i%
Call IOWriteLog(msg$)

End If

' Correct MAN count data for absorption of the continuum
If UseMANAbsFlag Then
For i% = 1 To npts%
If sample(1).MANAbsCorFlags%(chan%) And abscor!(i%) <> 0# Then
tydata!(i%) = tydata!(i%) * abscor!(i%)
End If
Next i%
End If

' Debug
If DebugMode Then
msg$ = "Counts: "
For i% = 1 To npts%
msg$ = msg$ & MiscAutoFormat$(tydata!(i%))
Next i%
Call IOWriteLog(msg$)
End If

' Do least squares fit of data in txdata and tydata
If npts% = 0 Then GoTo UpdateFitMANNoPoints
norder% = sample(1).MANLinearFitOrders%(chan%)
Call LeastSquares(norder%, npts%, txdata!(), tydata!(), acoeff!())
If ierror Then Exit Sub

analysis.MANFitCoefficients!(1, chan%) = acoeff!(1)
analysis.MANFitCoefficients!(2, chan%) = acoeff!(2)
analysis.MANFitCoefficients!(3, chan%) = acoeff!(3)

If DebugMode Then
msg$ = "Coeffs: "
msg$ = msg$ & MiscAutoFormat$(acoeff!(1)) & MiscAutoFormat$(acoeff!(2)) & MiscAutoFormat$(acoeff!(3))
Call IOWriteLog(msg$)
End If

Exit Sub

' Errors
UpdateFitMANError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateFitMAN"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateFitMANBadStandard:
msg$ = "MAN standard number " & Format$(sample(1).MANStdAssigns%(i%, chan%)) & " is not present in the current probe data file."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateFitMAN"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateFitMANNoPoints:
row% = SampleGetRow%(sample())
msg$ = SampleGetString$(row%)
msg$ = "No MAN data to fit for " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " ( spectro " & Format$(sample(1).MotorNumbers%(chan%)) & " " & sample(1).CrystalNames$(chan%) & ") at " & sample(1).KilovoltsArray!(chan%) & " keV, in sample " & msg$ & ". " & vbCrLf & vbCrLf
msg$ = msg$ & "Be sure that MAN standards have been acquired and that MAN assignments are properly assigned by using the Analytical | Assign MAN Fits menu. "
msg$ = msg$ & "If no MAN elements were acquired check the Analytical | Use Off Peak Elements for MAN Fit to have the program utilize off-peak elements for the MAN assignments."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateFitMAN"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

End Sub

Sub UpdateGetMANStandards(mode As Integer, sample() As TypeSample)
' Routine to load all MAN standard sets for MAN drift correction calculation of
' both standards and unknowns. See routine UpdateCalculateMANDrift for actual
' MAN drift calculation. Called by routines UpdateStdDriftCounts and MANLoad, etc.
'  mode = 0 display debug data
'  mode = 1 do not display debug data

ierror = False
On Error GoTo UpdateGetMANStandardsError

Dim i As Integer, j As Integer, k As Integer, n As Integer, response As Integer
Dim samplerow As Integer, ip As Integer, chan As Integer
Dim jmax As Integer, kmax As Integer
Dim astring As String
Dim alreadyasked As Boolean

Dim average As TypeAverage
Dim average2 As TypeAverage
Dim average3 As TypeAverage

If DebugMode Then
Call IOWriteLog(vbCrLf & "Entering UpdateGetMANStandards...")
End If

' Initialize MAN drift arrays (note 1 to MAXSET%)
If sample(1).LastElm% < 1 Then GoTo UpdateGetMANStandardsNoElements

For i% = 1 To sample(1).LastElm%
For j% = 1 To MAXMAN%
For k% = 1 To MAXSET%
MANAssignsDriftCounts!(k%, j%, i%) = 0#
MANAssignsDateTimes(k%, j%, i%) = 0#
MANAssignsSampleRows%(k%, j%, i%) = 0
MANAssignsCountTimes!(k%, j%, i%) = 0#
MANAssignsBeamCurrents!(k%, j%, i%) = 0#
Next k%
Next j%

' Initialize the MAN set counters
For j% = 1 To MAXMAN%
MANAssignsSets%(j%, i%) = 0
Next j%
Next i%

' Init missing MAN assignments
MissingMANAssignmentsStringsNumberOf% = 0

' Search through samples and load all assigned MAN standard sets
For samplerow% = 1 To NumberofSamples%

' Check if not a standard or all deleted lines
If SampleTyps%(samplerow%) <> 1 Or SampleDels%(samplerow%) Then GoTo 9000

' See if this standard is used for the MAN corrections for this sample
For i% = 1 To sample(1).LastElm%
If sample(1).DisableQuantFlag%(i%) = 0 And sample(1).BackgroundTypes%(i%) = 1 Then   ' 0=off-peak, 1=MAN, 2=multipoint
For j% = 1 To MAXMAN%
If SampleNums%(samplerow%) = sample(1).MANStdAssigns%(j%, i%) Then GoTo 7000
Next j%
End If
Next i%

' Standard is not used for any MAN standard assignments, try next sample
GoTo 9000

' Load data from disk file into "CorData" array for this standard
7000:

' Update status form
msg$ = SampleGetString(samplerow%)
Call AnalyzeStatusAnal("Loading count data for MAN  " & msg$ & "...")
If icancelanal Then
ierror = True
Exit Sub
End If

' Load data for this MAN standard
Call DataGetMDBSample(samplerow%, UpdateTmpSample())
If ierror Then Exit Sub

' Do a Savitzky-Golay smooth to the integrated intensity data if smooth selected
If UpdateTmpSample(1).IntegratedIntensitiesUseFlag And IntegratedIntensityUseSmoothingFlag Then
For chan% = 1 To UpdateTmpSample(1).LastElm%
If UpdateTmpSample(1).IntegratedIntensitiesUseIntegratedFlags%(chan%) Then
Call DataCorrectDataIntegratedSmooth(chan%, IntegratedIntensitySmoothingPointsPerSide%, UpdateTmpSample())
If ierror Then Exit Sub
End If
Next chan%
End If

' Correct the data for dead time, beam drift (and off-peak
' background if NOT using off-peak sample elements)
If UseOffPeakElementsForMANFlag = False Then
If Not UseAggregateIntensitiesFlag Then
Call DataCorrectData(Int(0), UpdateTmpSample())
If ierror Then Exit Sub

Else
Call DataCorrectData(Int(2), UpdateTmpSample())     ' skip aggregate intensity load
If ierror Then Exit Sub
Call DataCorrectDataAggregate(Int(2), sample(), UpdateTmpSample())     ' perform aggregate intensity based on unknown
If ierror Then Exit Sub
End If

Else
If Not UseAggregateIntensitiesFlag Then
Call DataCorrectData(Int(1), UpdateTmpSample())
If ierror Then Exit Sub
Else
Call DataCorrectData(Int(3), UpdateTmpSample())     ' skip aggregate intensity load
If ierror Then Exit Sub
Call DataCorrectDataAggregate(Int(3), sample(), UpdateTmpSample())     ' perform aggregate intensity based on unknown
If ierror Then Exit Sub
End If
End If

' Check for valid data points
If UpdateTmpSample(1).Datarows% < 1 Then GoTo 9000
If UpdateTmpSample(1).GoodDataRows% < 1 Then GoTo 9000

' Check that sample integrated intensity flag matches (this line is commented out as it prevents mixing integrated intensities and MAN elements)
'If UpdateTmpSample(1).IntegratedIntensitiesUseFlag% <> sample(1).IntegratedIntensitiesUseFlag% Then GoTo 9000

' If analytical conditions do not match selected sample, skip (if passed MAN sample has combined conditions, this check is skipped)
If Not UpdateTmpSample(1).CombinedConditionsFlag And Not sample(1).CombinedConditionsFlag Then
If UpdateTmpSample(1).takeoff! <> sample(1).takeoff! Then GoTo 9000
If UpdateTmpSample(1).kilovolts! <> sample(1).kilovolts! Then GoTo 9000
End If

' Average the standard count data and datetime
Call MathCountAverage(average, UpdateTmpSample())
If ierror Then Exit Sub

' Average the count times
Call MathArrayAverage(average2, UpdateTmpSample(1).OnTimeData!(), UpdateTmpSample(1).Datarows%, UpdateTmpSample(1).LastElm%, UpdateTmpSample())
If ierror Then Exit Sub

' Average the beam currents
If Not UpdateTmpSample(1).CombinedConditionsFlag Then
Call MathArrayAverage(average3, UpdateTmpSample(1).OnBeamData!(), UpdateTmpSample(1).Datarows%, UpdateTmpSample(1).LastElm%, UpdateTmpSample())
If ierror Then Exit Sub
Else
Call MathArrayAverage(average3, UpdateTmpSample(1).OnBeamDataArray!(), UpdateTmpSample(1).Datarows%, UpdateTmpSample(1).LastElm%, UpdateTmpSample())
If ierror Then Exit Sub
End If

If DebugMode Then
msg$ = vbCrLf & "MAN Standard " & SampleGetString2$(UpdateTmpSample()) & " intensity data (averaged):"
Call IOWriteLog(msg$)

msg$ = "ELEM: "
For i% = 1 To UpdateTmpSample(1).LastElm%
msg$ = msg$ & Format$(UpdateTmpSample(1).Elsyms$(i%), a80$)
Next i%
Call IOWriteLog(msg$)

For j% = 1 To UpdateTmpSample(1).Datarows%
msg$ = Format$(UpdateTmpSample(1).Linenumber&(j%), a60$)
For i% = 1 To UpdateTmpSample(1).LastElm%
msg$ = msg$ & MiscAutoFormat$(UpdateTmpSample(1).CorData!(j%, i%))
Next i%
Call IOWriteLog(msg$)
Next j%

msg$ = "AVER: "
For i% = 1 To UpdateTmpSample(1).LastElm%
msg$ = msg$ & MiscAutoFormat$(average.averags!(i%))
Next i%
Call IOWriteLog(msg$)
End If

' Load counts from this MAN standard, if assigned to the sample elements
For i% = 1 To sample(1).LastElm%
If sample(1).DisableQuantFlag%(i%) = 0 Then

For j% = 1 To MAXMAN%
If sample(1).MANStdAssigns%(j%, i%) <> SampleNums%(samplerow%) Then GoTo 7400

' Make sure the element is analyzed in the MAN standard also
ip% = IPOS5(Int(0), i%, sample(), UpdateTmpSample())
If ip% = 0 Then GoTo 7400

' Check that element is not off-peak corrected
If UpdateTmpSample(1).BackgroundTypes%(ip%) <> 1 Then GoTo 7400  ' 0=off-peak, 1=MAN, 2=multipoint

' If analytical conditions do not match selected sample, skip
If sample(1).CombinedConditionsFlag Or UpdateTmpSample(1).CombinedConditionsFlag Then
If sample(1).TakeoffArray!(i%) <> UpdateTmpSample(1).TakeoffArray!(ip) Then GoTo 7400
If sample(1).KilovoltsArray!(i%) <> UpdateTmpSample(1).KilovoltsArray!(ip%) Then GoTo 7400
End If

' Check for Bragg order
If sample(1).BraggOrders%(i%) <> UpdateTmpSample(1).BraggOrders%(ip%) Then GoTo 7400

' Check for matching element integrated intensity flag
If sample(1).IntegratedIntensitiesUseIntegratedFlags%(i%) <> UpdateTmpSample(1).IntegratedIntensitiesUseIntegratedFlags%(ip%) Then GoTo 7400

' Check for disabled acquisition in standard
If UpdateTmpSample(1).DisableAcqFlag%(ip%) = 1 Then GoTo 7400

' Check for disabled quant in standard (do not check standard flags)
'If UpdateTmpSample(1).DisableQuantFlag%(ip%) = 1 Then GoTo 7400

' Matching conditions, now increment set counter
If MANAssignsSets%(j, i%) + 1 <= MAXSET% Then
MANAssignsSets%(j, i%) = MANAssignsSets%(j, i%) + 1

' Load average counts
MANAssignsDriftCounts!(MANAssignsSets%(j, i%), j%, i%) = average.averags!(ip%)

' Load average count times
MANAssignsCountTimes!(MANAssignsSets%(j, i%), j%, i%) = average2.averags!(ip%)

' Load average beam currents
MANAssignsBeamCurrents!(MANAssignsSets%(j, i%), j%, i%) = average3.averags!(ip%)

' Load average "dateTime"
MANAssignsDateTimes(MANAssignsSets%(j, i%), j%, i%) = average.AverDateTime

' Load the sample row number for saving the sample to the SETUP2 (MAN) database
MANAssignsSampleRows%(MANAssignsSets%(j%, i%), j%, i%) = samplerow%

' Warn if too many MAN standardizations
Else
msg$ = "WARNING- Too many MAN standardizations on standard " & Format$(sample(1).MANStdAssigns%(j%, i%)) & " for " & sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%)
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateGetMANStandards"
End If

7400:  Next j%
End If
Next i%

9000:  Next samplerow%
Call AnalyzeStatusAnal(vbNullString)

' Check for empty MAN sets on MAN corrected elements
For i% = 1 To sample(1).LastElm%
If sample(1).BackgroundTypes%(i%) = 1 And sample(1).DisableQuantFlag(i%) = 0 Then  ' 0=off-peak, 1=MAN, 2=multipoint
For j% = 1 To MAXMAN%
If sample(1).MANStdAssigns%(j%, i%) > 0 And MANAssignsSets%(j%, i%) = 0 Then

MissingMANAssignmentsStringsNumberOf% = MissingMANAssignmentsStringsNumberOf% + 1
ReDim Preserve MissingMANAssignmentsStrings(1 To MissingMANAssignmentsStringsNumberOf%) As String

astring$ = "Warning- No MAN intensity data found for standard " & Format$(sample(1).MANStdAssigns%(j%, i%)) & " for "
astring$ = astring$ & MiscAutoUcase$(sample(1).Elsyms$(i%)) & " " & sample(1).Xrsyms$(i%) & " "
astring$ = astring$ & "on spectrometer " & Format$(sample(1).MotorNumbers%(i%)) & " crystal " & sample(1).CrystalNames$(i%) & " at "
astring$ = astring$ & Format$(sample(1).KilovoltsArray!(i%)) & " KeV for MAN assignments."
MissingMANAssignmentsStrings$(MissingMANAssignmentsStringsNumberOf%) = astring$

sample(1).MANStdAssigns%(j%, i%) = 0    ' deselect MAN assignment
End If
Next j%
End If
Next i%

' Output elements with missing MAN assignments
For n% = 1 To MissingMANAssignmentsStringsNumberOf%
If Not alreadyasked Then
msg$ = vbCrLf & "One or more MAN elements are missing standard intensities. Either acquire data for the "
msg$ = msg$ & "indicated element/standard at the indicated conditions or "
msg$ = msg$ & "remove the MAN assignment by clicking " & MiscAutoUcase$(sample(1).Elsyms$(i%)) & " from "
msg$ = msg$ & "the MAN Assignments element grid, then <ctrl> click on the Standards list box to unselect the MAN assignment and "
msg$ = msg$ & "click the Update Fit button." & vbCrLf & vbCrLf
msg$ = msg$ & "If you loaded a file setup from another run and removed or changed any standards you should "
msg$ = msg$ & "clear the MAN assignments using the Analytical | Clear All MAN Assignments menu and re-assign them "
msg$ = msg$ & "using the Analytical | Assign MAN Fits menu." & vbCrLf & vbCrLf
msg$ = msg$ & "This error can also occur when quick standards are acquired before all MAN standards have been assigned to the MAN background "
msg$ = msg$ & "calibration curve. Please acquire standards the first time without the quick standards option "
msg$ = msg$ & "so all MAN elements are acquired on each standard." & vbCrLf & vbCrLf
msg$ = msg$ & "Do you want to see the missing MAN assignments? Click Yes to see each element/standard with missing MAN intensities, "
msg$ = msg$ & "No to just output to the log window, or Cancel to exit."

' Check first time user response
response% = MsgBox(msg$, vbYesNoCancel + vbExclamation, "UpdateGetMANStandards")
alreadyasked = True
End If

' User cancels
If response% = vbCancel Then
ierror = True
Exit Sub
End If

' Output to log window
msg$ = MissingMANAssignmentsStrings$(n%)
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))

' Output msgbox
If response% = vbYes Then
msg$ = msg$ & vbCrLf & vbCrLf & "Do you want to see the missing MAN assignments? Click Yes to see each element/standard with missing MAN intensities, No to just output to the log window, or Cancel to exit."
response% = MsgBox(msg$, vbYesNoCancel + vbInformation, "UpdateGetMANStandards")
End If
Next n%

If Not DebugMode Or mode% = 1 Then Exit Sub

' Debug, print MAN drift array
jmax% = UpdateGetMaxMANAssign(sample())
kmax% = UpdateGetMaxSet(Int(3), sample())

For j% = 1 To jmax%

' Print MAN standard assignments and MAN sets
If j% = 1 Then
msg$ = vbCrLf & "MAN Standard drift array set numbers (cps/" & Format$(NominalBeam!) & FaradayCurrentUnits$ & "):"
Call IOWriteLog(msg$)
msg$ = vbCrLf & "ELEM: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).Elsyup$(i%), a80$)
Next i%
Call IOWriteLog(msg$)
End If

' Print MAN standard numbers
msg$ = "MANASS"
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(sample(1).MANStdAssigns%(j%, i%), i80$), a80$)
Next i%
Call IOWriteLog(msg$)

' Print MAN sets
msg$ = "MANSET"
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(MANAssignsSets%(j%, i%), i80$), a80$)
Next i%
Call IOWriteLog(msg$)
Next j%

' Print MAN standard counts
For k% = 1 To kmax%
msg$ = vbCrLf & "Set " & Format$(k%) & " MAN Standard Drift Arrays Counts (cps/" & Format$(NominalBeam!) & FaradayCurrentUnits$ & "):"
Call IOWriteLog(msg$)

msg$ = vbCrLf & "ELEM: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).Elsyup$(i%), a80$)
Next i%
Call IOWriteLog(msg$)

For j% = 1 To jmax%

' Print MAN standard numbers (again)
msg$ = "MANASS"
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(sample(1).MANStdAssigns%(j%, i%), i80$), a80$)
Next i%
Call IOWriteLog(msg$)

' Print MAN drift counts
msg$ = "MANCNT"
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(MANAssignsDriftCounts!(k%, j%, i%), f82$), a80$)
Next i%
Call IOWriteLog(msg$)

' Print MAN drift count times
msg$ = "MANTIM"
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(MANAssignsCountTimes!(k%, j%, i%), f82$), a80$)
Next i%
Call IOWriteLog(msg$)

' Print MAN drift beam currents
msg$ = "MANCUR"
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(MANAssignsBeamCurrents!(k%, j%, i%), f82$), a80$)
Next i%
Call IOWriteLog(msg$)

Next j%
Next k%

Exit Sub

' Errors
UpdateGetMANStandardsError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateGetMANStandards"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

UpdateGetMANStandardsNoElements:
msg$ = "Warning- No elements in this sample at " & Format$(sample(1).takeoff!) & " degrees  and " & Format$(sample(1).kilovolts!) & " KeV "
msg$ = msg$ & "for the MAN assignments. Click the Acquire | New Sample, buttons "
msg$ = msg$ & "to add some analyzed elements to the run first or change the KeV field "
msg$ = msg$ & "in the MAN Fits window and click Re-Load."
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateGetMANStandards"
Call AnalyzeStatusAnal(vbNullString)
ierror = False  ' do not set error flag for missing elements
Exit Sub

End Sub

Sub UpdateGetStdPercents(analysis As TypeAnalysis, sample() As TypeSample, stdsample() As TypeSample)
' Load the standard percents for all standards in run

ierror = False
On Error GoTo UpdateGetStdPercentsError

Dim i As Integer, j As Integer
Dim ip As Integer

' Loop on each standard
For j% = 1 To NumberofStandards%

' Get standard composition and cations from the standard database and load in stdsample arrays
Call StandardGetMDBStandard(StandardNumbers%(j%), stdsample())
If ierror Then Exit Sub

' Load standard arrays for this standard into the standard list arrays
For i% = 1 To sample(1).LastChan%
ip% = IPOS1(stdsample(1).LastChan%, sample(1).Elsyms$(i%), stdsample(1).Elsyms$())

' Load the standard percents
If ip% > 0 Then
analysis.StdPercents!(j%, i%) = stdsample(1).ElmPercents!(ip%)
End If

Next i%
Next j%

Exit Sub

' Errors
UpdateGetStdPercentsError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateGetStdPercents"
ierror = True
Exit Sub

End Sub

Sub UpdateInitSample()
' Init update structures

ierror = False
On Error GoTo UpdateInitError

' Dimension sample array for changed sample
Call InitSample(UpdateOldSample())
If ierror Then Exit Sub

Exit Sub

' Errors
UpdateInitError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateInit"
ierror = True
Exit Sub

End Sub
