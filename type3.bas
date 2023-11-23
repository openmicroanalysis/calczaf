Attribute VB_Name = "CodeTYPE3"
' (c) Copyright 1995-2023 by John J. Donovan
Option Explicit

Sub TypeNewCounts(sample() As TypeSample)
' This routine types out the new acquired xray counts to the log window

ierror = False
On Error GoTo TypeNewCountsError

Dim i As Integer, row As Integer
Dim ii As Integer, jj As Integer, n As Integer
Dim temp As Single, counttime As Single

Dim average As TypeAverage

Dim average1 As TypeAverage
Dim average2 As TypeAverage
Dim average3 As TypeAverage
Dim average4 As TypeAverage

msg$ = TypeLoadString(sample())
If ierror Then Exit Sub
Call IOWriteLogRichText(vbCrLf & msg$, vbNullString, Int(LogWindowFontSize% + 4), vbBlue, Int(FONT_ITALIC% Or FONT_UNDERLINE%), Int(0))
msg$ = TypeLoadDescription(sample())
If ierror Then Exit Sub
Call IOWriteLog(vbCrLf & msg$)

' Print EDS acquisition time (if no data yet)
If UseDetailedFlag And sample(1).EDSSpectraFlag And sample(1).Datarows% = 0 Then
If UseEDSSampleCountTimeFlag Then
Call RealTimeEDSSpectraGetSampleOrSpecifiedTime(Int(1), counttime!, sample())
If ierror Then Exit Sub
msg$ = "EDS Acquisition (sample count) Time: " & Space$(25) & Format$(Format$(counttime!, f82$), a80$)
Else
msg$ = "EDS Acquisition (user specified) Time: " & Space$(23) & Format$(Format$(sample(1).LastEDSSpecifiedCountTime!, f82$), a80$)
End If
Call IOWriteLog(msg$)
End If

' Print CL acquisition time (if no data yet)
If UseDetailedFlag And sample(1).Datarows% = 0 Then
If sample(1).CLSpectraFlag Then
msg$ = "CL Acquisition Time:  " & Space$(40) & Format$(Format$(sample(1).LastCLSpecifiedCountTime!, f82$), a80$) & vbCrLf
msg$ = msg$ & "CL Unknown Count Factor:  " & Space$(36) & Format$(Format$(sample(1).LastCLUnknownCountFactor!, f81$), a80$) & vbCrLf
msg$ = msg$ & "CL Dark Spectra Count Time Fraction:  " & Space$(24) & Format$(Format$(sample(1).LastCLDarkSpectraCountTimeFraction!, f81$), a80$)
Call IOWriteLog(msg$)
End If
End If

' Print Number of lines and number of "good" lines
If UseDetailedFlag Then
msg$ = "Number of Data Lines: " & Format$(sample(1).Datarows%, a30$)
msg$ = msg$ & Space$(12)
msg$ = msg$ & " Number of 'Good' Data Lines: " & Format$(sample(1).GoodDataRows%, a30$)
Call IOWriteLog(msg$)
End If

' Print Number of lines and number of "good" lines and date and time
If UseDetailedFlag Then
If sample(1).Datarows% > 0 Then
msg$ = "First/Last Date-Time: " & Format$(sample(1).DateTimes#(1), "mm/dd/yyyy hh:mm:ss AM/PM") & " to " & Format$(sample(1).DateTimes#(sample(1).Datarows%), "mm/dd/yyyy hh:mm:ss AM/PM")
Call IOWriteLog(msg$)
End If
End If

' Type sample setup
If sample(1).LastElm% > 0 Then          ' skip output if no WDS or EDS analyzed elements
Call TypeSampleSetup(sample())
If ierror Then Exit Sub

' If combined sample type arrays
If sample(1).CombinedConditionsFlag And UseDetailedFlag Then
Call TypeCombined(Int(0), sample())
If ierror Then Exit Sub
End If

' Obtain EDS net intensities
If sample(1).EDSSpectraFlag And sample(1).EDSSpectraUseFlag Then
If UpdateEDSCheckForEDSElements(Int(1), sample()) Then
Call UpdateEDSSpectraNetIntensities(sample())
If ierror Then Exit Sub
End If
End If

' Warn if using integrated counts
If sample(1).IntegratedIntensitiesFlag% Then
msg$ = vbCrLf & "Warning: Sample Contains One or More Elements Specified for Integrated Intensity Acquisition"
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
End If
End If

' Print type of data to Log window
If sample(1).Datarows% = 0 Then Exit Sub

' Do a Savitzky-Golay smooth to the integrated intensity data if smooth selected
'If sample(1).IntegratedIntensitiesUseFlag And IntegratedIntensityUseSmoothingFlag Then
'For chan% = 1 To sample(1).LastElm%
'If sample(1).IntegratedIntensitiesUseIntegratedFlags%(chan%) Then
'Call DataCorrectDataIntegratedSmooth(chan%, IntegratedIntensitySmoothingPointsPerSide%, sample())
'If ierror Then Exit Sub
'End If
'Next chan%
'End If

' Correct data for deadtime and beam current
Call DataCorrectData(Int(0), sample())
If ierror Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

If UseBeamDriftCorrectionFlag And NominalBeam! <> 0# Then
msg$ = vbCrLf & "Off-Peak Corrected or MAN On-Peak X-ray Counts (cps/" & Format$(NominalBeam!) & FaradayCurrentUnits$ & ") (and Faraday/Absorbed Currents):"
Else
msg$ = vbCrLf & "Off-Peak Corrected or MAN On-Peak X-ray Counts (cps):"
End If
Call IOWriteLog(msg$)

n = 0
Do Until False
n% = n% + 1
Call TypeGetRange(Int(1), n%, ii%, jj%, sample())
If ierror Then Exit Sub
If sample(1).LastElm% > 0 And ii% > sample(1).LastElm% Then Exit Do         ' this line is modified (also see end of the loop below), so the beam currents will be printed out even if there are no analyzed elements

' Type out symbols for count time data lines for analyzed elements
msg$ = "ELEM: "
For i% = ii% To jj%
If sample(1).DisableAcqFlag%(i%) = 1 Then
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & "-D", a80$)
Else
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%), a80$)
End If
Next i%

' Type beam current column labels (do not output combined condition beam/absorbed currents here)
If Not sample(1).CombinedConditionsFlag Then
msg$ = msg$ & Format$("   BEAM1", a80$)
If sample(1).OnBeamCounts2!(1) > 0# Then msg$ = msg$ & Format$("   BEAM2", a80$)
If sample(1).AbBeamCounts!(1) > 0# Then msg$ = msg$ & Format$("   ABSD1", a80$)
If sample(1).AbBeamCounts2!(1) > 0# Then msg$ = msg$ & Format$("   ABSD2", a80$)
End If
Call IOWriteLog(msg$)

' Calculate beam current averages
If Not sample(1).CombinedConditionsFlag Then
Call MathAverage(average1, sample(1).OnBeamCounts!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
Call MathAverage(average2, sample(1).OnBeamCounts2!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
Call MathAverage(average3, sample(1).AbBeamCounts!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
Call MathAverage(average4, sample(1).AbBeamCounts2!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
End If

' For each row type out counts (all lines)
For row% = 1 To sample(1).Datarows
msg$ = Format$(Format$(sample(1).Linenumber&(row%), i50$), a50$)
If sample(1).LineStatus(row%) Then msg$ = msg$ & "G"
If Not sample(1).LineStatus(row%) Then msg$ = msg$ & "B"

' Type corrected counts
For i% = ii% To jj%
If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(sample(1).CorData!(row%, i%), f81$), a80$)
Else
msg$ = msg$ & Format$(Format$(sample(1).CorData!(row%, i%), f82$), a80$)
End If
Next i%

' Add columns for beam and absorbed currents if not combined condition sample
If Not sample(1).CombinedConditionsFlag Then
msg$ = msg$ & Format$(Format$(sample(1).OnBeamCounts!(row%), FaradayCurrentFormat$), a80$)
If sample(1).OnBeamCounts2!(1) > 0# Then msg$ = msg$ & Format$(Format$(sample(1).OnBeamCounts2!(row%), FaradayCurrentFormat$), a80$)
If sample(1).AbBeamCounts!(1) > 0# Then msg$ = msg$ & Format$(Format$(sample(1).AbBeamCounts!(row%), FaradayCurrentFormat$), a80$)
If sample(1).AbBeamCounts2!(1) > 0# Then msg$ = msg$ & Format$(Format$(sample(1).AbBeamCounts2!(row%), FaradayCurrentFormat$), a80$)
End If
Call IOWriteLog(msg$)
Next row%

' Type average counts and std deviation
If sample(1).Datarows% > 0 Then
Call MathArrayAverage(average, sample(1).CorData!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub

' Print average counts
msg$ = vbCrLf & "AVER: "
For i% = ii% To jj%
If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(average.averags!(i%), f81$), a80$)
Else
msg$ = msg$ & Format$(Format$(average.averags!(i%), f82$), a80$)
End If
Next i%
If Not sample(1).CombinedConditionsFlag Then
msg$ = msg$ & Format$(Format$(average1.averags!(1), FaradayCurrentFormat$), a80$)
If average2.averags!(1) > 0# Then msg$ = msg$ & Format$(Format$(average2.averags!(1), FaradayCurrentFormat$), a80$)
If sample(1).AbBeamCounts!(1) > 0# Then msg$ = msg$ & Format$(Format$(average3.averags!(1), FaradayCurrentFormat$), a80$)
If sample(1).AbBeamCounts2!(1) > 0# Then msg$ = msg$ & Format$(Format$(average4.averags!(1), FaradayCurrentFormat$), a80$)
End If
Call IOWriteLogRichText(msg$, vbNullString, Int(0), VbDarkBlue&, Int(0), Int(0))
        
msg$ = "SDEV: "
For i% = ii% To jj%
If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(average.Stddevs!(i%), f81$), a80$)
Else
msg$ = msg$ & Format$(Format$(average.Stddevs!(i%), f82$), a80$)
End If
Next i%
If Not sample(1).CombinedConditionsFlag Then
msg$ = msg$ & Format$(Format$(average1.Stddevs!(1), FaradayCurrentFormat$), a80$)
If average2.averags!(1) > 0# Then msg$ = msg$ & Format$(Format$(average2.Stddevs!(1), FaradayCurrentFormat$), a80$)
If sample(1).AbBeamCounts!(1) > 0# Then msg$ = msg$ & Format$(Format$(average3.Stddevs!(1), FaradayCurrentFormat$), a80$)
If sample(1).AbBeamCounts2!(1) > 0# Then msg$ = msg$ & Format$(Format$(average4.Stddevs!(1), FaradayCurrentFormat$), a80$)
End If
Call IOWriteLogRichText(msg$, vbNullString, Int(0), VbDarkBlue&, Int(0), Int(0))

' Modify square-root for being normalized to cps
Call TypeCalculateOneSigma(Int(1), Int(1), sample(1).LastElm%, average, sample())
If ierror Then Exit Sub

msg$ = "1SIG: "
For i% = ii% To jj%
If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(average.Sqroots!(i%), f81$), a80$)
Else
msg$ = msg$ & Format$(Format$(average.Sqroots!(i%), f82$), a80$)
End If
Next i%
Call IOWriteLogRichText(msg$, vbNullString, Int(0), VbDarkBlue&, Int(0), Int(0))
    
msg$ = "SERR: "
For i% = ii% To jj%
If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(average.Stderrs!(i%), f81$), a80$)
Else
msg$ = msg$ & Format$(Format$(average.Stderrs!(i%), f82$), a80$)
End If
Next i%
Call IOWriteLog(msg$)

msg$ = "%RSD: "
For i% = ii% To jj%
If Abs(100# * average.Reldevs!(i%)) < MAXRELDEV% Then
If NominalBeam! > 1# Then
msg$ = msg$ & Format$(Format$(100# * average.Reldevs!(i%), f82$), a80$)
Else
msg$ = msg$ & Format$(Format$(100# * average.Reldevs!(i%), f82$), a80$)     ' leave the same
End If
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If
Next i%
Call IOWriteLog(msg$)
End If

If Not ExtendedFormat Then Call IOWriteLog(vbNullString)
If ii% > sample(1).LastElm% Then Exit Do         ' this line is so the beam currents will be printed out even if there are no analyzed elements
Loop

' Do a quick check of the sigma ratio and warn user if element is assigned as primary standard and is excessive
If sample(1).Type% = 1 Then
For i% = 1 To sample(1).LastElm%
If sample(1).number% = sample(1).StdAssigns%(i%) And sample(1).StdAssigns%(i%) <> 0 Then
temp! = 0#
If average.Sqroots!(i%) <> 0# Then
temp! = average.Stddevs!(i%) / average.Sqroots!(i%)
End If
If temp! > 4# Then
msg$ = "The assigned element " & sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & " has an excessively large sigma ratio (" & Format$(temp!) & ") for " & TypeLoadString(sample()) & ". Please check that the standard intensity data was properly acquired."
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(0), Int(0))
End If
End If
Next i%
End If

Exit Sub

' Errors
TypeNewCountsError:
MsgBox Error$, vbOKOnly + vbCritical, "TypeNewCounts"
ierror = True
Exit Sub

End Sub

Sub TypeSampleSetup(sample() As TypeSample)
' Type just the sample setup

ierror = False
On Error GoTo TypeSampleSetupError

Dim n As Integer, i As Integer
Dim k As Integer, m As Integer
Dim ii As Integer, jj As Integer
Dim temp As Single, count_time As Single, counttime As Single

' Only output peak and PHA if not all EDS
If Not MiscAreAllElementsEDS(sample()) Then

' Peak positions
If sample(1).LastElm% > 0 And UseDetailedFlag Then
msg$ = vbCrLf & "On and Off Peak Positions:"
Call IOWriteLog(msg$)

n = 0
Do Until False
n% = n% + 1
Call TypeGetRange(Int(1), n%, ii%, jj%, sample())
If ierror Then Exit Sub
If ii% > sample(1).LastElm% Then Exit Do

If n% <> 1 Then
Call IOWriteLog(vbNullString)
End If

' Type out symbols for data lines for analyzed elements
msg$ = "ELEM: "
For i% = ii% To jj%
If sample(1).DisableAcqFlag%(i%) = 1 Then
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & "-D", a80$)
Else
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%), a80$)
End If
Next i%
Call IOWriteLog(msg$)

' Type out crystal
msg$ = "CRYST:"
For i% = ii% To jj%
msg$ = msg$ & Format$(sample(1).CrystalNames$(i%), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "ONPEAK"
For i% = ii% To jj%
msg$ = msg$ & MiscAutoFormat$(sample(1).OnPeaks!(i%))
Next i%
Call IOWriteLog(msg$)
msg$ = "OFFSET"
For i% = ii% To jj%
msg$ = msg$ & MiscAutoFormat$(sample(1).Offsets!(i%))
Next i%
Call IOWriteLog(msg$)

' Normal off-peak positions
msg$ = "HIPEAK"
For i% = ii% To jj%
If sample(1).BackgroundTypes%(i%) = 0 Then
msg$ = msg$ & MiscAutoFormat$(sample(1).HiPeaks!(i%))
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If
Next i%
Call IOWriteLog(msg$)
msg$ = "LOPEAK"
For i% = ii% To jj%
If sample(1).BackgroundTypes%(i%) = 0 Then
msg$ = msg$ & MiscAutoFormat$(sample(1).LoPeaks!(i%))
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If
Next i%
Call IOWriteLog(msg$)

msg$ = "HI-OFF"
For i% = ii% To jj%
If sample(1).BackgroundTypes%(i%) = 0 Then
msg$ = msg$ & MiscAutoFormat$(sample(1).HiPeaks!(i%) - sample(1).OnPeaks!(i%))
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If
Next i%
Call IOWriteLog(msg$)
msg$ = "LO-OFF"
For i% = ii% To jj%
If sample(1).BackgroundTypes%(i%) = 0 Then
msg$ = msg$ & MiscAutoFormat$(sample(1).LoPeaks!(i%) - sample(1).OnPeaks!(i%))
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If
Next i%
Call IOWriteLog(msg$)

' Multi-point parameters and positions
If ProbeDataFileVersionNumber! > 8.31 Then
If MiscIsEqualTo(sample(1).LastElm%, sample(1).BackgroundTypes%(), Int(2)) Or MiscIsEqualTo(sample(1).LastElm%, sample(1).OffPeakCorrectionTypes%(), Int(MAXOFFBGDTYPES%)) Or DebugMode Then

msg$ = vbCrLf & "Multi-Point Background Positions and Parameters:"
Call IOWriteLog(msg$)

' Type out symbols for data lines for analyzed elements
msg$ = "ELEM: "
For i% = ii% To jj%
If sample(1).DisableAcqFlag%(i%) = 1 Then
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & "-D", a80$)
Else
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%), a80$)
End If
Next i%
Call IOWriteLog(msg$)

' Print multi-point positions
k% = MiscGetArrayMax%(MAXMULTI%, sample(1).MultiPointNumberofPointsAcquireHi%())
For m% = 1 To k%
msg$ = "MULHI:"
For i% = ii% To jj%
If (sample(1).BackgroundTypes%(i%) = 2 Or sample(1).OffPeakCorrectionTypes%(i%) = MAXOFFBGDTYPES%) And sample(1).MultiPointNumberofPointsAcquireHi%(i%) >= m% Then
msg$ = msg$ & MiscAutoFormat$(sample(1).MultiPointAcquirePositionsHi!(i%, m%))
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If
Next i%
Call IOWriteLog(msg$)
Next m%

Call IOWriteLog(vbNullString)
For m% = 1 To k%
msg$ = "MHIOFF"
For i% = ii% To jj%
If (sample(1).BackgroundTypes%(i%) = 2 Or sample(1).OffPeakCorrectionTypes%(i%) = MAXOFFBGDTYPES%) And sample(1).MultiPointNumberofPointsAcquireHi%(i%) >= m% Then
msg$ = msg$ & MiscAutoFormat$(sample(1).MultiPointAcquirePositionsHi!(i%, m%) - sample(1).OnPeaks!(i%))
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If
Next i%
Call IOWriteLog(msg$)
Next m%

' Type out symbols for data lines for analyzed elements
msg$ = vbCrLf & "ELEM: "
For i% = ii% To jj%
If sample(1).DisableAcqFlag%(i%) = 1 Then
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & "-D", a80$)
Else
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%), a80$)
End If
Next i%
Call IOWriteLog(msg$)

k% = MiscGetArrayMax%(MAXMULTI%, sample(1).MultiPointNumberofPointsAcquireLo%())
For m% = 1 To k%
msg$ = "MULLO:"
For i% = ii% To jj%
If (sample(1).BackgroundTypes%(i%) = 2 Or sample(1).OffPeakCorrectionTypes%(i%) = MAXOFFBGDTYPES%) And sample(1).MultiPointNumberofPointsAcquireLo%(i%) >= m% Then
msg$ = msg$ & MiscAutoFormat$(sample(1).MultiPointAcquirePositionsLo!(i%, m%))
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If
Next i%
Call IOWriteLog(msg$)
Next m%

Call IOWriteLog(vbNullString)
For m% = 1 To k%
msg$ = "MLOOFF"
For i% = ii% To jj%
If (sample(1).BackgroundTypes%(i%) = 2 Or sample(1).OffPeakCorrectionTypes%(i%) = MAXOFFBGDTYPES%) And sample(1).MultiPointNumberofPointsAcquireLo%(i%) >= m% Then
msg$ = msg$ & MiscAutoFormat$(sample(1).MultiPointAcquirePositionsLo!(i%, m%) - sample(1).OnPeaks!(i%))
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If
Next i%
Call IOWriteLog(msg$)
Next m%

' Type out symbols for data lines for analyzed elements
msg$ = vbCrLf & "ELEM: "
For i% = ii% To jj%
If sample(1).DisableAcqFlag%(i%) = 1 Then
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & "-D", a80$)
Else
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%), a80$)
End If
Next i%
Call IOWriteLog(msg$)

' Print multi-point parameters
msg$ = "MACQHI"
For i% = ii% To jj%
If sample(1).BackgroundTypes%(i%) = 2 Or sample(1).OffPeakCorrectionTypes%(i%) = MAXOFFBGDTYPES% Then
msg$ = msg$ & Format$(Format$(sample(1).MultiPointNumberofPointsAcquireHi%(i%), i80$), a80$)
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If
Next i%
Call IOWriteLog(msg$)

msg$ = "MACQLO"
For i% = ii% To jj%
If sample(1).BackgroundTypes%(i%) = 2 Or sample(1).OffPeakCorrectionTypes%(i%) = MAXOFFBGDTYPES% Then
msg$ = msg$ & Format$(Format$(sample(1).MultiPointNumberofPointsAcquireLo%(i%), i80$), a80$)
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If
Next i%
Call IOWriteLog(msg$)

msg$ = "MUITHI"
For i% = ii% To jj%
If sample(1).BackgroundTypes%(i%) = 2 Or sample(1).OffPeakCorrectionTypes%(i%) = MAXOFFBGDTYPES% Then
msg$ = msg$ & Format$(Format$(sample(1).MultiPointNumberofPointsIterateHi%(i%), i80$), a80$)
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If
Next i%
Call IOWriteLog(msg$)

msg$ = "MUITLO"
For i% = ii% To jj%
If sample(1).BackgroundTypes%(i%) = 2 Or sample(1).OffPeakCorrectionTypes%(i%) = MAXOFFBGDTYPES% Then
msg$ = msg$ & Format$(Format$(sample(1).MultiPointNumberofPointsIterateLo%(i%), i80$), a80$)
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If
Next i%
Call IOWriteLog(msg$)

msg$ = "MULFIT"
For i% = ii% To jj%
If sample(1).BackgroundTypes%(i%) = 2 Or sample(1).OffPeakCorrectionTypes%(i%) = MAXOFFBGDTYPES% Then
msg$ = msg$ & Format$(MultiPointBackgroundFitTypeStrings2$(sample(1).MultiPointBackgroundFitType%(i%)), a80$)  ' 0 = linear, 1 = 2nd order polynomial, 2 = exponential
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If
Next i%
Call IOWriteLog(msg$)

End If
End If

Loop
End If

' Print PHA parameters
If sample(1).LastElm% > 0 And UseDetailedFlag Then
msg$ = vbCrLf & "PHA Parameters:"
Call IOWriteLog(msg$)

n = 0
Do Until False
n% = n% + 1
Call TypeGetRange(Int(1), n%, ii%, jj%, sample())
If ierror Then Exit Sub
If ii% > sample(1).LastElm% Then Exit Do

If n% <> 1 Then
Call IOWriteLog(vbNullString)
End If

' Type out symbols for data lines for analyzed elements
msg$ = "ELEM: "
For i% = ii% To jj%
If sample(1).DisableAcqFlag%(i%) = 1 Then
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & "-D", a80$)
Else
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%), a80$)
End If
Next i%
Call IOWriteLog(msg$)

msg$ = "DEAD: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).DeadTimes!(i%) * MSPS!, f82$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "BASE: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).Baselines!(i%), f82$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "WINDOW"
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).Windows!(i%), f82$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "MODE: "
For i% = ii% To jj%
If sample(1).CrystalNames$(i%) <> EDS_CRYSTAL$ Then
If sample(1).InteDiffModes%(i%) = 0 Then
msg$ = msg$ & Format$("INTE", a80$)
Else
msg$ = msg$ & Format$("DIFF", a80$)
End If
Else
msg$ = msg$ & Format$("----", a80$)
End If
Next i%
Call IOWriteLog(msg$)

msg$ = "GAIN: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).Gains!(i%), f80$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "BIAS: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).Biases!(i%), f80$), a80$)
Next i%
Call IOWriteLog(msg$)

Loop
End If
End If

' Type LAST count times for each element
If sample(1).LastElm% > 0 And UseDetailedFlag Then
msg$ = vbCrLf & "Last (Current) On and Off Peak Count Times: "
If sample(1).Type% = 2 And ProbeDataFileVersionNumber! > 2.44 Then
If MiscIsDifferent2(sample(1).LastElm%, sample(1).LastMaxCounts&()) Or DebugMode Then
msg$ = msg$ & "(" & VbDquote$ & DASHED4$ & VbDquote$ & " indicates default max count)"
End If
End If
Call IOWriteLog(msg$)

n% = 0
Do Until False
n% = n% + 1
Call TypeGetRange(Int(1), n%, ii%, jj%, sample())
If ierror Then Exit Sub
If ii% > sample(1).LastElm% Then Exit Do

If n% <> 1 Then
Call IOWriteLog(vbNullString)
End If

' Type out symbols for data lines for analyzed elements
msg$ = "ELEM: "
For i% = ii% To jj%
If sample(1).DisableAcqFlag%(i%) = 1 Then
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & "-D", a80$)
Else
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%), a80$)
End If
Next i%
Call IOWriteLog(msg$)

' Type out background correction type (0=off-peak, 1=MAN, 2=multipoint)
If sample(1).Type% <> 3 Then
msg$ = "BGD:  "
For i% = ii% To jj%
If sample(1).CrystalNames$(i%) <> EDS_CRYSTAL$ Then
msg$ = msg$ & Format$(BgdTypeStrings$(sample(1).BackgroundTypes%(i%)), a80$)
Else
msg$ = msg$ & Format$(EDS_CRYSTAL$, a80$) ' special treatment for EDS spectrum data
End If
Next i%
Call IOWriteLog(msg$)

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
msg$ = msg$ & Format$(BgStrings$(sample(1).OffPeakCorrectionTypes%(i%)), a80$)  ' 0=linear, 1=average, 2=high only, 3=low only, 4=exponential, 5=slope hi, 6=slope lo, 7=polynomial, 8=multi-point
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
End If

If DebugMode Or MiscIsDifferent(sample(1).LastElm%, sample(1).BraggOrders%()) Then
msg$ = "BRAGG:"
For i% = ii% To jj%
msg$ = msg$ & Format$(sample(1).BraggOrders%(i%), a80$)
Next i%
Call IOWriteLog(msg$)
End If

' Type out motor
msg$ = "SPEC: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).MotorNumbers%(i%), i80$), a80$)
Next i%
Call IOWriteLog(msg$)

' Type out crystal
msg$ = "CRYST:"
For i% = ii% To jj%
msg$ = msg$ & Format$(sample(1).CrystalNames$(i%), a80$)
Next i%
Call IOWriteLog(msg$)

If DebugMode Then
msg$ = "CRY2D:"
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).Crystal2ds!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)
msg$ = "CRYK :"
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).CrystalKs!(i%), f86$), a80$)
Next i%
Call IOWriteLog(msg$)
End If

' Type out order number
msg$ = "ORDER:"
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).OrderNumbers%(i%), i80$), a80$)
Next i%
Call IOWriteLog(msg$)

If sample(1).Type% <> 3 Or DebugMode Then
msg$ = "ONTIM:"
For i% = ii% To jj%

' WDS count times
If sample(1).CrystalNames$(i%) <> EDS_CRYSTAL$ Then
msg$ = msg$ & Format$(Format$(sample(1).LastOnCountTimes!(i%), f82$), a80$)

' EDS count times
Else
If sample(1).Datarows% = 0 Then
If UseEDSSampleCountTimeFlag Then
Call RealTimeEDSSpectraGetSampleOrSpecifiedTime(Int(1), counttime!, sample())
If ierror Then Exit Sub
Else
counttime! = sample(1).LastEDSSpecifiedCountTime!
End If
msg$ = msg$ & Format$(Format$(counttime!, f82$), a80$)
Else
msg$ = msg$ & Format$(Format$(sample(1).EDSSpectraLiveTime!(sample(1).Datarows%), f82$), a80$)  ' all EDS spectra datarows use same count time
End If

End If

Next i%
Call IOWriteLog(msg$)

If Not sample(1).AllMANBgdFlag Then
msg$ = "HITIM:"
For i% = ii% To jj%
If sample(1).BackgroundTypes%(i%) = 1 Or sample(1).CrystalNames$(i%) = EDS_CRYSTAL$ Then  ' 0=off-peak, 1=MAN, 2=multipoint
msg$ = msg$ & Format$(DASHED4$, a80$)
Else
msg$ = msg$ & Format$(Format$(sample(1).LastHiCountTimes!(i%), f82$), a80$)
End If
Next i%
Call IOWriteLog(msg$)

msg$ = "LOTIM:"
For i% = ii% To jj%
If sample(1).BackgroundTypes%(i%) = 1 Or sample(1).CrystalNames$(i%) = EDS_CRYSTAL$ Then   ' 0=off-peak, 1=MAN, 2=multipoint
msg$ = msg$ & Format$(DASHED4$, a80$)
Else
msg$ = msg$ & Format$(Format$(sample(1).LastLoCountTimes!(i%), f82$), a80$)
End If
Next i%
Call IOWriteLog(msg$)
End If
End If

If sample(1).Type% = 2 And ProbeDataFileVersionNumber! > 2.44 Then
If MiscIsDifferent3(sample(1).LastElm%, sample(1).LastCountFactors!()) Or Not MiscAllOne(sample(1).LastElm%, sample(1).LastCountFactors!) Or DebugMode Then
msg$ = "UNFAC:"
For i% = ii% To jj%
If sample(1).CrystalNames$(i%) <> EDS_CRYSTAL$ Then
msg$ = msg$ & Format$(sample(1).LastCountFactors!(i%), a80$)
Else
msg$ = msg$ & Format$(sample(1).LastEDSUnknownCountFactor!, a80$)
End If
Next i%
Call IOWriteLog(msg$)
End If
End If

' Print multiplied count times if unknown count factors are not all one
If ProbeDataFileVersionNumber! > 2.44 Then
If sample(1).Type% = 1 And sample(1).UnknownCountTimeForInterferenceStandardFlag Or sample(1).Type% = 2 Then
If MiscIsDifferent3(sample(1).LastElm%, sample(1).LastCountFactors!()) Or Not MiscAllOne(sample(1).LastElm%, sample(1).LastCountFactors!) Or DebugMode Then
msg$ = "ONTIME"
For i% = ii% To jj%

If sample(1).Type% = 2 Then          ' unknown sample
If sample(1).CrystalNames$(i%) <> EDS_CRYSTAL$ Then
count_time! = sample(1).LastOnCountTimes!(i%) * sample(1).LastCountFactors!(i%)
Else
count_time! = sample(1).LastEDSSpecifiedCountTime! * sample(1).LastEDSUnknownCountFactor!
End If
End If

If sample(1).Type% = 1 Then         ' standard sample
count_time! = sample(1).LastOnCountTimes!(i%)

' Standard sample has no data
If sample(1).Datarows% = 0 Then
If AcquireIsUseUnknownCountTimeForInterferenceStandardFlag(i%, sample()) Then
count_time! = sample(1).LastOnCountTimes!(i%) * sample(1).LastCountFactors!(i%)
End If

' Standard sample has data
Else
If sample(1).UnknownCountTimeForInterferenceStandardChanFlag(i%) Then
count_time! = sample(1).LastOnCountTimes!(i%) * sample(1).LastCountFactors!(i%)
End If
End If
End If

msg$ = msg$ & Format$(Format$(count_time!, f82$), a80$)
Next i%
Call IOWriteLog(msg$)

' Output actual off-peak background count times
If Not sample(1).AllMANBgdFlag Then
msg$ = "HITIME"
For i% = ii% To jj%
If sample(1).BackgroundTypes%(i%) <> 0 Or sample(1).CrystalNames$(i%) = EDS_CRYSTAL$ Then    ' 0=off-peak, 1=MAN, 2=multipoint
msg$ = msg$ & Format$(DASHED4$, a80$)
Else

If sample(1).Type% = 2 Then          ' unknown sample
count_time! = sample(1).LastHiCountTimes!(i%) * sample(1).LastCountFactors!(i%)
End If

If sample(1).Type% = 1 Then         ' standard sample
count_time! = sample(1).LastHiCountTimes!(i%)

' Standard sample has no data
If sample(1).Datarows% = 0 Then
If AcquireIsUseUnknownCountTimeForInterferenceStandardFlag(i%, sample()) Then
count_time! = sample(1).LastHiCountTimes!(i%) * sample(1).LastCountFactors!(i%)
End If

' Standard sample has data
Else
If sample(1).UnknownCountTimeForInterferenceStandardChanFlag(i%) Then
count_time! = sample(1).LastHiCountTimes!(i%) * sample(1).LastCountFactors!(i%)
End If
End If
End If

msg$ = msg$ & Format$(Format$(count_time!, f82$), a80$)
End If
Next i%
Call IOWriteLog(msg$)

msg$ = "LOTIME"
For i% = ii% To jj%
If sample(1).BackgroundTypes%(i%) <> 0 Or sample(1).CrystalNames$(i%) = EDS_CRYSTAL$ Then   ' 0=off-peak, 1=MAN, 2=multipoint
msg$ = msg$ & Format$(DASHED4$, a80$)
Else

If sample(1).Type% = 2 Then          ' unknown sample
count_time! = sample(1).LastLoCountTimes!(i%) * sample(1).LastCountFactors!(i%)
End If

If sample(1).Type% = 1 Then         ' standard sample
count_time! = sample(1).LastLoCountTimes!(i%)

' Standard sample has no data
If sample(1).Datarows% = 0 Then
If AcquireIsUseUnknownCountTimeForInterferenceStandardFlag(i%, sample()) Then
count_time! = sample(1).LastLoCountTimes!(i%) * sample(1).LastCountFactors!(i%)
End If

' Standard sample has data
Else
If sample(1).UnknownCountTimeForInterferenceStandardChanFlag(i%) Then
count_time! = sample(1).LastLoCountTimes!(i%) * sample(1).LastCountFactors!(i%)
End If
End If
End If

msg$ = msg$ & Format$(Format$(count_time!, f82$), a80$)
End If
Next i%
Call IOWriteLog(msg$)
End If
End If

' Output actual (acquired) MPB background count times
If MiscContainsInteger(sample(1).LastElm%, sample(1).BackgroundTypes%(), Int(2)) Then
msg$ = "HIMULT"
For i% = ii% To jj%
If sample(1).BackgroundTypes%(i%) <> 2 Or sample(1).CrystalNames$(i%) = EDS_CRYSTAL$ Then    ' 0=off-peak, 1=MAN, 2=multipoint
msg$ = msg$ & Format$(DASHED4$, a80$)
Else

If sample(1).Type% = 2 Then          ' unknown sample
count_time! = sample(1).LastHiCountTimes!(i%) / 2# * sample(1).MultiPointNumberofPointsAcquireHi%(i%) * sample(1).LastCountFactors!(i%)
End If

If sample(1).Type% = 1 Then         ' standard sample
count_time! = sample(1).LastHiCountTimes!(i%) / 2# * sample(1).MultiPointNumberofPointsAcquireHi%(i%)

' Standard sample has no data
If sample(1).Datarows% = 0 Then
If AcquireIsUseUnknownCountTimeForInterferenceStandardFlag(i%, sample()) Then
count_time! = sample(1).LastHiCountTimes!(i%) / 2# * sample(1).MultiPointNumberofPointsAcquireHi%(i%) * sample(1).LastCountFactors!(i%)
End If

' Standard sample has data
Else
If sample(1).UnknownCountTimeForInterferenceStandardChanFlag(i%) Then
count_time! = sample(1).LastHiCountTimes!(i%) / 2# * sample(1).MultiPointNumberofPointsAcquireHi%(i%) * sample(1).LastCountFactors!(i%)
End If
End If
End If

msg$ = msg$ & Format$(Format$(count_time!, f82$), a80$)
End If
Next i%
Call IOWriteLog(msg$)

msg$ = "LOMULT"
For i% = ii% To jj%
If sample(1).BackgroundTypes%(i%) <> 2 Or sample(1).CrystalNames$(i%) = EDS_CRYSTAL$ Then   ' 0=off-peak, 1=MAN, 2=multipoint
msg$ = msg$ & Format$(DASHED4$, a80$)
Else

If sample(1).Type% = 2 Then          ' unknown sample
count_time! = sample(1).LastLoCountTimes!(i%) / 2# * sample(1).MultiPointNumberofPointsAcquireLo%(i%) * sample(1).LastCountFactors!(i%)
End If

If sample(1).Type% = 1 Then         ' standard sample
count_time! = sample(1).LastLoCountTimes!(i%) / 2# * sample(1).MultiPointNumberofPointsAcquireLo%(i%)

' Standard sample has no data
If sample(1).Datarows% = 0 Then
If AcquireIsUseUnknownCountTimeForInterferenceStandardFlag(i%, sample()) Then
count_time! = sample(1).LastLoCountTimes!(i%) / 2# * sample(1).MultiPointNumberofPointsAcquireLo%(i%) * sample(1).LastCountFactors!(i%)
End If

' Standard sample has data
Else
If sample(1).UnknownCountTimeForInterferenceStandardChanFlag(i%) Then
count_time! = sample(1).LastLoCountTimes!(i%) / 2# * sample(1).MultiPointNumberofPointsAcquireLo%(i%) * sample(1).LastCountFactors!(i%)
End If
End If
End If

msg$ = msg$ & Format$(Format$(count_time!, f82$), a80$)
End If
Next i%
Call IOWriteLog(msg$)
End If

End If
End If

If sample(1).Type% = 2 And ProbeDataFileVersionNumber! > 4# Then
If MiscIsDifferent2(sample(1).LastElm%, sample(1).LastMaxCounts&()) Or DebugMode Then
msg$ = "MAXCNT"
For i% = ii% To jj%
If sample(1).LastMaxCounts&(i%) = MAXCOUNT& Then
msg$ = msg$ & Format$(DASHED4$, a80$)
Else
msg$ = msg$ & Format$(Format$(sample(1).LastMaxCounts&(i%), i80$), a80$)
End If
Next i%
Call IOWriteLog(msg$)
End If

' Display number of aggregate channels
If UseAggregateIntensitiesFlag Then
msg$ = "AGGR: "
For i% = ii% To jj%
If sample(1).DisableQuantFlag%(i%) = 0 Then

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

Loop
End If

' Type out volatile correction assignments (-1 = self, 0 = none, >0 = assigned)
If UseDetailedFlag Or DebugMode Then
msg$ = vbCrLf & "Miscellaneous Sample Acquisition/Calculation Parameters: "
Call IOWriteLog(msg$)
Call ElementGetData(sample())
If ierror Then Exit Sub

n% = 0
Do Until False
n% = n% + 1
Call TypeGetRange(Int(1), n%, ii%, jj%, sample())
If ierror Then Exit Sub
If ii% > sample(1).LastElm% Then Exit Do

If n% <> 1 Then
Call IOWriteLog(vbNullString)
End If

msg$ = "ETAKOF"
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).EffectiveTakeOffs!(i%), f82$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "KILO: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).KilovoltsArray!(i%), f82$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "ENERGY"
For i% = ii% To jj%
msg$ = msg$ & MiscAutoFormatN$(sample(1).LineEnergy!(i%) / EVPERKEV#, Int(3))   ' output in KeV
Next i%
Call IOWriteLog(msg$)

msg$ = "EDGE: "
For i% = ii% To jj%
msg$ = msg$ & MiscAutoFormatN$(sample(1).LineEdge!(i%) / EVPERKEV#, Int(3))     ' output in KeV
Next i%
Call IOWriteLog(msg$)

msg$ = "Eo/Ec:"
For i% = ii% To jj%
temp! = 0#
If sample(1).LineEdge!(i%) <> 0# Then temp! = EVPERKEV# * sample(1).KilovoltsArray!(i%) / sample(1).LineEdge!(i%)
msg$ = msg$ & MiscAutoFormatN$(temp!, Int(2))
Next i%
Call IOWriteLog(msg$)

msg$ = "ATWT: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).AtomicWts!(i%), f83$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "STDS: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).StdAssigns%(i%), i50$), a80$)
Next i%
Call IOWriteLog(msg$)

' Print TDI assignment (0 = none, 1 = self, 2 = assigned)
If sample(1).VolatileAcquisitionType% <> 0 And sample(1).LastElm% > 1 Then
If MiscIsDifferent(sample(1).LastElm%, sample(1).VolatileCorrectionUnks%()) Or Not MiscAllZero(sample(1).LastElm%, sample(1).VolatileCorrectionUnks%()) Or DebugMode Then
msg$ = "TDI#: "
For i% = ii% To jj%
msg$ = msg$ & Format$(sample(1).VolatileCorrectionUnks%(i%), a80$)
Next i%
Call IOWriteLog(msg$)
End If
End If

' Type out blank correction assignment
If UseBlankCorFlag% Then
If sample(1).Type% = 2 And sample(1).LastElm% > 1 Then
If MiscIsDifferent(sample(1).LastElm%, sample(1).BlankCorrectionUnks%()) Or DebugMode Then
msg$ = "BLNK#:"
For i% = ii% To jj%
msg$ = msg$ & Format$(sample(1).BlankCorrectionUnks%(i%), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "BLNKL:"
For i% = ii% To jj%
msg$ = msg$ & MiscAutoFormat$(sample(1).BlankCorrectionLevels!(i%))
Next i%
Call IOWriteLog(msg$)
End If
End If
End If

' Type out specified APF
If UseAPFFlag% And UseAPFOption% = 1 Then
msg$ = "APF*: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).SpecifiedAreaPeakFactors!(i%), f83$), a80$)
Next i%
Call IOWriteLog(msg$)
End If

' Type integrated intensity flag
If sample(1).Type% <> 3 And sample(1).LastElm% > 1 Then
If Not MiscAllZero(sample(1).LastElm%, sample(1).IntegratedIntensitiesUseIntegratedFlags%()) Or DebugMode Then
msg$ = "INTE: "
For i% = ii% To jj%
msg$ = msg$ & Format$(sample(1).IntegratedIntensitiesUseIntegratedFlags%(i%), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "INTEIN"
For i% = ii% To jj%
If sample(1).IntegratedIntensitiesUseIntegratedFlags%(i%) Then
msg$ = msg$ & MiscAutoFormat$(sample(1).IntegratedIntensitiesInitialStepSizes!(i%))
Else
msg$ = msg$ & Format$(DASHED3$, a80$)
End If
Next i%
Call IOWriteLog(msg$)

msg$ = "INTEMI"
For i% = ii% To jj%
If sample(1).IntegratedIntensitiesUseIntegratedFlags%(i%) Then
msg$ = msg$ & MiscAutoFormat$(sample(1).IntegratedIntensitiesMinimumStepSizes!(i%))
Else
msg$ = msg$ & Format$(DASHED3$, a80$)
End If
Next i%
Call IOWriteLog(msg$)
End If
End If

' Type peaking before acquisition flag
If sample(1).Type% = 2 And sample(1).LastElm% > 1 Then
If MiscIsDifferent(sample(1).LastElm%, sample(1).PeakingBeforeAcquisitionElementFlags%()) Or DebugMode Then
msg$ = "PEAK: "
For i% = ii% To jj%
msg$ = msg$ & Format$(sample(1).PeakingBeforeAcquisitionElementFlags%(i%), a80$)
Next i%
Call IOWriteLog(msg$)
End If
End If

' Type sample Nth Point acquisition flags for each element
If (ProbeDataFileVersionNumber! > 8.25 Or sample(1).Datarows% = 0) And sample(1).LastElm% > 1 Then
If MiscIsDifferent(sample(1).LastElm%, sample(1).NthPointAcquisitionFlags%()) Or DebugMode Then
msg$ = "NthPT:"
For i% = ii% To jj%
If sample(1).NthPointAcquisitionFlags%(i%) Then
msg$ = msg$ & Format$(Format$(sample(1).NthPointAcquisitionIntervals%(i%), i80$), a80$)
Else
msg$ = msg$ & Format$(DASHED4$, a80$)
End If
Next i%
Call IOWriteLog(msg$)
End If
End If

Loop
End If

Exit Sub

' Errors
TypeSampleSetupError:
MsgBox Error$, vbOKOnly + vbCritical, "TypeSampleSetup"
ierror = True
Exit Sub

End Sub

Sub TypeCombined(mode As Integer, sample() As TypeSample)
' Type the multiple sample setup array
'  mode = 0 use sample(1).Elsyms$()
'  mode = 1 use sample(1).Elsyup$()

ierror = False
On Error GoTo TypeCombinedError

Dim i As Integer, ii As Integer, jj As Integer, n As Integer

msg$ = vbCrLf & "Combined Analytical Condition Arrays:"
Call IOWriteLog(msg$)

' Type multiple conditions if combined sample
n% = 0
Do Until False
n% = n% + 1
Call TypeGetRange(Int(1), n%, ii%, jj%, sample())
If ierror Then Exit Sub
If ii% > sample(1).LastElm% Then Exit Do

' Type out symbols for conditions of analyzed elements
If UseDetailedFlag Or PrintAnalyzedAndSpecifiedOnSameLineFlag Then
msg$ = "ELEM: "
For i% = ii% To jj%
If mode% = 0 Then
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%), a80$)
Else
msg$ = msg$ & Format$(sample(1).Elsyup$(i%), a80$)
End If
Next i%
Call IOWriteLog(msg$)

' Type out motor
If mode% = 0 Then
msg$ = "SPEC: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).MotorNumbers%(i%), i80$), a80$)
Next i%
Call IOWriteLog(msg$)
End If

msg$ = "CONDN:"
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).ConditionNumbers%(i%), i80$), a80$)     ' condition number for this element
Next i%
Call IOWriteLog(msg$)

msg$ = "CONDO:"
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).ConditionOrders%(sample(1).ConditionNumbers%(i%)), i80$), a80$)      ' combined condition acquisition order
Next i%
Call IOWriteLog(msg$)

If DebugMode Then
msg$ = "TAKE: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).TakeoffArray!(i%), f81$), a80$)
Next i%
Call IOWriteLog(msg$)
End If

msg$ = "KILO: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).KilovoltsArray!(i%), f81$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "CURR: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).BeamCurrentArray!(i%), f81$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "SIZE: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).BeamSizeArray!(i%), f81$), a80$)
Next i%
Call IOWriteLog(msg$)

If MiscIsDifferent(sample(1).LastElm%, sample(1).ColumnConditionMethodArray%()) Or Not MiscAllZero(sample(1).LastElm%, sample(1).ColumnConditionMethodArray%()) Then
msg$ = "METH: "
For i% = ii% To jj%
msg$ = msg$ & Format$(sample(1).ColumnConditionMethodArray%(i%), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "STRI: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Left$(MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(sample(1).ColumnConditionStringArray$(i%))), 8), a80$)
Next i%
Call IOWriteLog(msg$)
End If

End If
Loop

Exit Sub

' Errors
TypeCombinedError:
MsgBox Error$, vbOKOnly + vbCritical, "TypeCombined"
ierror = True
Exit Sub

End Sub

Sub TypeCalculateOneSigma(mode As Integer, ii As Integer, jj As Integer, tAverage As TypeAverage, sample() As TypeSample)
' Adjust square roots because the counts are normalized to cps
' mode = 1 calculate 1 sigmas for on-peak counts
' mode = 2 calculate 1 sigmas for hi-peak counts
' mode = 3 calculate 1 sigmas for lo-peak counts

ierror = False
On Error GoTo TypeCalculateOneSigmaError

Dim i As Integer

Dim onaverage As TypeAverage
Dim hiaverage As TypeAverage
Dim loaverage As TypeAverage

Dim bmaverage As TypeAverage
Dim ctaverage As TypeAverage

' Calculate average count times
If mode% = 1 Then
Call MathArrayAverage(onaverage, sample(1).OnTimeData!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
End If
If mode% = 2 Then
Call MathArrayAverage(hiaverage, sample(1).HiTimeData!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
End If
If mode% = 3 Then
Call MathArrayAverage(loaverage, sample(1).LoTimeData!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
End If

' Calculate average beam currents for each element
If Not sample(1).CombinedConditionsFlag Then
Call MathAverage(bmaverage, sample(1).OnBeamCounts!(), sample(1).Datarows%, sample())
If ierror Then Exit Sub
Else
Call MathArrayAverage(bmaverage, sample(1).OnBeamCountsArray!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
End If

' Calculate average count data using raw cps!!!!
If mode% = 1 Then
Call MathArrayAverage(ctaverage, sample(1).OnPeakCounts_Raw_Cps!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
End If
If mode% = 2 Then
Call MathArrayAverage(ctaverage, sample(1).HiPeakCounts_Raw_Cps!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
End If
If mode% = 3 Then
Call MathArrayAverage(ctaverage, sample(1).LoPeakCounts_Raw_Cps!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
End If

' Denormalize average counts for average count time
For i% = ii% To jj%
If mode% = 1 Then ctaverage.averags!(i%) = ctaverage.averags!(i%) * onaverage.averags!(i%)
If mode% = 2 Then ctaverage.averags!(i%) = ctaverage.averags!(i%) * hiaverage.averags!(i%)
If mode% = 3 Then ctaverage.averags!(i%) = ctaverage.averags!(i%) * loaverage.averags!(i%)
Next i%

' Type out debug data
If VerboseMode Then
msg$ = vbCrLf & "ELEM: "
For i% = ii% To jj%
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%), a80$)
Next i%
Call IOWriteLog(msg$)

If mode% = 1 Then msg$ = "AVGON:"
If mode% = 2 Then msg$ = "AVGHI:"
If mode% = 3 Then msg$ = "AVGLO:"
For i% = ii% To jj%
If mode% = 1 Then msg$ = msg$ & MiscAutoFormat$(onaverage.averags!(i%))
If mode% = 2 Then msg$ = msg$ & MiscAutoFormat$(hiaverage.averags!(i%))
If mode% = 3 Then msg$ = msg$ & MiscAutoFormat$(loaverage.averags!(i%))
Next i%
Call IOWriteLog(msg$)

' Type average beam currents
msg$ = "AVGBM:"
For i% = ii% To jj%
If Not sample(1).CombinedConditionsFlag Then
msg$ = msg$ & MiscAutoFormat$(bmaverage.averags!(1))
Else
msg$ = msg$ & MiscAutoFormat$(bmaverage.averags!(i%))
End If
Next i%
Call IOWriteLog(msg$)

' Type current raw counts
msg$ = "COUNTS"
For i% = ii% To jj%
msg$ = msg$ & MiscAutoFormat$(ctaverage.averags!(i%))
Next i%
Call IOWriteLog(msg$)
End If

' Now calculate the square root on the actual raw data
For i% = ii% To jj%
If ctaverage.averags!(i%) > 0# Then
tAverage.Sqroots!(i%) = Sqr(ctaverage.averags!(i%))
Else
tAverage.Sqroots!(i%) = 0#
End If
Next i%

' Type current square roots
If VerboseMode Then
msg$ = "SQRT1:"
For i% = ii% To jj%
msg$ = msg$ & MiscAutoFormat$(tAverage.Sqroots!(i%))
Next i%
Call IOWriteLog(msg$)
End If

' Now normalize to count time
For i% = ii% To jj%
If mode% = 1 And onaverage.averags!(i%) <> 0# Then
tAverage.Sqroots!(i%) = tAverage.Sqroots!(i%) / onaverage.averags!(i%)
End If
If mode% = 2 And hiaverage.averags!(i%) <> 0# Then
tAverage.Sqroots!(i%) = tAverage.Sqroots!(i%) / hiaverage.averags!(i%)
End If
If mode% = 3 And loaverage.averags!(i%) <> 0# Then
tAverage.Sqroots!(i%) = tAverage.Sqroots!(i%) / loaverage.averags!(i%)
End If
Next i%

' Type current square roots
If VerboseMode Then
msg$ = "SQRT2:"
For i% = ii% To jj%
msg$ = msg$ & MiscAutoFormat$(tAverage.Sqroots!(i%))
Next i%
Call IOWriteLog(msg$)
End If

' Now normalize back for deadtime and beam drift
For i% = ii% To jj%
If UseDeadtimeCorrectionFlag And sample(1).CrystalNames$(i%) <> EDS_CRYSTAL$ Then
Call DataCorrectDataDeadTime(tAverage.Sqroots!(i%), sample(1).DeadTimes!(i%))           ' normalize for deadtime
If ierror Then Exit Sub
End If

If UseBeamDriftCorrectionFlag Then
If Not sample(1).CombinedConditionsFlag Then
Call DataCorrectDataBeamDrift(tAverage.Sqroots!(i%), bmaverage.averags!(1))             ' normalize for beam
If ierror Then Exit Sub
Else
Call DataCorrectDataBeamDrift(tAverage.Sqroots!(i%), bmaverage.averags!(i%))            ' normalize for beam
If ierror Then Exit Sub
End If
End If
Next i%

' Type current square roots
If VerboseMode Then
msg$ = "SQRT3:"
For i% = ii% To jj%
msg$ = msg$ & MiscAutoFormat$(tAverage.Sqroots!(i%))
Next i%
Call IOWriteLog(msg$)
Call IOWriteLog(vbNullString)
End If

Exit Sub

' Errors
TypeCalculateOneSigmaError:
MsgBox Error$, vbOKOnly + vbCritical, "TypeCalculateOneSigma"
ierror = True
Exit Sub

End Sub

