Attribute VB_Name = "CodeINTERF2"
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

Dim Interf2Sample(1 To 1) As TypeSample

Function Interf2Calculate(method As Integer, mode As Integer, chan As Integer, rangefraction As Single, lifwidth As Single, minimumoverlap As Single, discrimination As Single, sample() As TypeSample) As String
' Calculates interferences based on a sample composition
'  method = 0 check all interfering elements
'  method = 1 check only interfering elements in sample
'  mode = 0 calculate on-peak interferences
'  mode = 1 calculate hi-peak interferences
'  mode = 2 calculate lo-peak interferences
'  chan = 0 calculate interferences on all elements in sample
'  chan > 0 calculate interferences for a single element channel
'  rangefraction = Xray database angstrom search range

ierror = False
On Error GoTo Interf2CalculateError

Dim i As Integer, n As Integer
Dim ip As Integer, ipp As Integer
Dim pos As Single, temp As Single, factor As Single
Dim interferedline As Single, interferingline As Single
Dim interferedintensity As Single, interferingintensity As Single
Dim xstart As Single, xstop As Single
Dim klm As Single, keV As Single
Dim order As Integer
Dim overlapfraction As Single, sigma As Single, overlappercent As Single
Dim xlabel As String, xedge As String, tmsg As String, astring As String
Dim esym As String, xsym As String

ReDim nominalintensity(1 To MAXRAY% - 1) As Single

nominalintensity!(1) = 150# ' Ka
nominalintensity!(2) = 15#  ' Kb
nominalintensity!(3) = 100# ' La
nominalintensity!(4) = 25#  ' Lb
nominalintensity!(5) = 200# ' Ma
nominalintensity!(6) = 100# ' Mb

' Check for valid ranges
If lifwidth! <= 0# Or lifwidth! > 10# Then GoTo Interf2CalculateBadWidth
If minimumoverlap! <= 0# Or minimumoverlap! > 100# Then GoTo Interf2CalculateBadOverlap
If discrimination! <= 0# Or discrimination! > 100# Then GoTo Interf2CalculateBadDiscrimination

If chan% < 0 Or chan% > sample(1).LastElm% Then GoTo Interf2CalculateBadChan

' Loop on all analyzed elements in sample
tmsg$ = vbNullString
For i% = 1 To sample(1).LastElm%

' Check for valid lines
If sample(1).Xrsyms$(i%) = vbNullString Then GoTo 3000

' Check if doing a single element in sample
If chan% > 0 And chan% <> i% Then GoTo 3000

' Load specified position and convert to angstroms
If mode% = 0 Then pos! = sample(1).OnPeaks!(i%)
If mode% = 1 Then pos! = sample(1).HiPeaks!(i%)
If mode% = 2 Then pos! = sample(1).LoPeaks!(i%)
interferedline! = XrayConvertSpecAng(Int(1), i%, pos!, sample(1).BraggOrders%(i%), sample())

' Load current analyzed line text
If chan% = 0 Then tmsg$ = tmsg$ & vbCrLf
tmsg$ = tmsg$ & "For " & Format$(sample(1).Elsyup$(i%), a20$) & " " & Format$(sample(1).Xrsyms$(i%), a20$) & " "
If mode% = 1 Then tmsg$ = tmsg$ & "(hi-off), "
If mode% = 2 Then tmsg$ = tmsg$ & "(lo-off), "
tmsg$ = tmsg$ & Format$(sample(1).CrystalNames$(i%), a60$) & " at " & MiscAutoFormat$(interferedline!) & " angstroms"
If mode% > 0 Then tmsg$ = tmsg$ & "(" & MiscAutoFormat$(pos!) & ")"
tmsg$ = tmsg$ & ", at an assumed concentration of " & Format$(sample(1).ElmPercents!(i%)) & " wt.% "
tmsg$ = tmsg$ & vbCrLf

' Calculate angstrom range to check
xstart! = interferedline! - (interferedline! * rangefraction!)
xstop! = interferedline! + (interferedline! * rangefraction!)

' Get Xray Database list for this interfered xray line
DefaultXrayStart! = xstart!
DefaultXrayStop! = xstop!
klm! = DefaultMinimumKLMDisplay!
keV! = sample(1).KilovoltsArray!(i%)

' Load Xray List box
Call XrayLoad(Int(2), Int(0), klm!, keV!, xstart!, xstop!)
If ierror Then Exit Function

' Calculate a spectral resolution factor for the Bragg angle of the spectrometer
temp! = MotLoLimits!(sample(1).MotorNumbers%(chan%)) + Abs(MotHiLimits!(sample(1).MotorNumbers%(chan%)) - MotLoLimits!(sample(1).MotorNumbers%(chan%)))
If temp! / sample(1).OnPeaks(chan%) < 0# Then GoTo Interf2CalculatePositionsNegative
factor! = 12# / (temp! / sample(1).OnPeaks(chan%))                      ' factor used to be a constant of 10.0, changed 04-04-2016 to a variable based on spectrometer angle

' Correct LiF width for actual crystal 2d
sigma! = lifwidth! / factor! * (sample(1).Crystal2ds!(i%) / LIF2D!) ^ 1.1

' Adjust for LDE
If sample(1).Crystal2ds!(i%) > MAXCRYSTAL2D_NOT_LDE! Then sigma! = sigma! * 3#  ' triple for LDE analyzers (changed to triple 9-28-2006)
If sample(1).Crystal2ds!(i%) > MAXCRYSTAL2D_LARGE_LDE! Then sigma! = sigma! * 2#   ' increase again for large 2d LDEs

' Loop on all lines from Xray database that match range
For n% = 0 To FormXRAY.ListXray.ListCount - 1
astring$ = FormXRAY.ListXray.List(n%)
Call XrayExtractListString(astring$, esym$, xsym$, order%, interferingline!, interferingintensity!, xlabel$, xedge$)
If ierror Then Exit Function

' Check if interfering element is in sample (if method = 1)
ip% = IPOS1(sample(1).LastChan%, esym$, sample(1).Elsyms$())
If ip% = 0 And method% = 1 Then GoTo 2000

' If on-peak interferences, check that interfering element (line) is not the interfered line
If mode% = 0 And ip% = i% Then GoTo 2000

' Get nominal overlap intensity
overlapfraction! = Interf2GetOverlap(interferingline!, interferedline!, sigma!)
If ierror Then Exit Function

' Correct intensity of interfering line for concentration and overlap
If ip% > 0 Then
interferingintensity! = interferingintensity! * sample(1).ElmPercents!(ip%) / 100# * overlapfraction!
Else
interferingintensity! = interferingintensity! * overlapfraction!    ' assume 100% concentration
End If

' Calculate interfered intensity
ipp% = IPOS1(MAXRAY% - 1, sample(1).Xrsyms$(i%), Xraylo$())
interferedintensity! = nominalintensity!(ipp%)
interferedintensity! = interferedintensity! * sample(1).ElmPercents!(i%) / 100#
If interferedintensity! <= 0# Then interferedintensity! = 0.1

' Correct interfering line intensity for Bragg order if differential mode to simulate PHA filtering
If order% > 1 Then
If sample(1).Set% = 0 Or sample(1).InteDiffModes(i%) Then       ' depending on whether called from Standard or Probewin
interferingintensity! = interferingintensity! / discrimination! * (order% - 1)
End If
End If

' Calculate actual overlap
overlappercent! = 100# * interferingintensity! / interferedintensity!

' Print if greater than specified minimum overlap
If overlappercent! > minimumoverlap! Then
tmsg$ = tmsg$ & "  Interference by"
tmsg$ = tmsg$ & Format$(xlabel$) & " "
tmsg$ = tmsg$ & "at " & MiscAutoFormat$(interferingline!) & " "
pos! = XrayConvertSpecAng(Int(2), i%, interferingline!, sample(1).BraggOrders%(i%), sample())   ' don't use actual order to be self consistant
tmsg$ = tmsg$ & "(" & MiscAutoFormat$(pos!) & ") "
If mode% = 0 Then
tmsg$ = tmsg$ & "(" & MiscAutoFormat$(pos! - sample(1).OnPeaks!(i%)) & ") "
End If
tmsg$ = tmsg$ & "= " & Format$(Format$(overlappercent!, f81$), a80$) & "%"
tmsg$ = tmsg$ & vbCrLf
End If

2000:  Next n%
3000:  Next i%

' Return string
Interf2Calculate$ = tmsg$

Exit Function

' Errors
Interf2CalculateError:
MsgBox Error$, vbOKOnly + vbCritical, "Interf2Calculate"
ierror = True
Exit Function

Interf2CalculateBadWidth:
msg$ = "Lif Peak Width is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "Interf2Calculate"
ierror = True
Exit Function

Interf2CalculateBadOverlap:
msg$ = "Minimum Overlap is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "Interf2Calculate"
ierror = True
Exit Function

Interf2CalculateBadDiscrimination:
msg$ = "PHA Discrimination is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "Interf2Calculate"
ierror = True
Exit Function

Interf2CalculateBadChan:
msg$ = "Channel number is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "Interf2Calculate"
ierror = True
Exit Function

Interf2CalculatePositionsNegative:
msg$ = "Negative result prior to square root on channel " & Format$(chan%) & ", onpos " & Format$(sample(1).OnPeaks!(chan%))
MsgBox msg$, vbOKOnly + vbExclamation, "Interf2Calculate"
ierror = True
Exit Function

End Function

Function Interf2GetOverlap(interferedline As Single, interferingline As Single, sigma As Single) As Single
' Routine to return the overlap fraction of the interfering line on the interfered line,
' based on the distance between the lines and the "sigma" gaussian peak width.

ierror = False
On Error GoTo Interf2GetOverlapError

Dim overlapfraction As Single

' Calculate
If Abs(-0.5 * ((interferingline! - interferedline!) ^ 2) > 75#) Then
overlapfraction! = 0#
Else
overlapfraction! = Exp(-0.5 * ((interferingline! - interferedline!) / sigma!) ^ 2)
End If
Interf2GetOverlap! = overlapfraction!

Exit Function

' Errors
Interf2GetOverlapError:
MsgBox Error$, vbOKOnly + vbCritical, "Interf2GetOverlap"
ierror = True
Exit Function

End Function

Function Interf2LoadElement(method%, mode As Integer, chan As Integer, sample() As TypeSample) As String
' Function to load a nominal sample for a single element interference calculation (called Probewin.exe only)
'  method = 0 check all interfering elements
'  method = 1 check only interfering elements in sample
'  mode = 0 calculate on-peak interferences
'  mode = 1 calculate hi-peak interferences
'  mode = 2 calculate lo-peak interferences

ierror = False
On Error GoTo Interf2LoadElementError

Dim i As Integer
Dim temp As Single, rangefraction As Single
Dim tmsg As String

' Load passed sample
Interf2LoadElement = vbNullString
Interf2Sample(1) = sample(1)

' Load element data
Call ElementGetData(Interf2Sample())
If ierror Then Exit Function

' Load dummy wtpercents
For i% = 1 To Interf2Sample(1).LastChan%
If i% <> chan% Then
Interf2Sample(1).ElmPercents!(i%) = 100#
Else
Interf2Sample(1).ElmPercents!(i%) = 1
End If
Next i%

' Check for non-default focal circle size
If MiscIsInstrumentStage("JEOL") Then
temp! = ROWLAND_JEOL# / ScalRolandCircleMMs!(sample(1).MotorNumbers%(chan%))
Else
temp! = ROWLAND_CAMECA# / ScalRolandCircleMMs!(sample(1).MotorNumbers%(chan%))
End If

' Get the nominal interference
rangefraction! = DefaultRangeFraction! * sample(1).Crystal2ds!(chan%) / LIF2D!
If rangefraction! > 0.5 Then rangefraction! = 0.5
tmsg$ = Interf2Calculate(method%, mode%, chan%, rangefraction!, DefaultLIFPeakWidth! * temp!, DefaultMinimumOverlap!, DefaultPHADiscrimination!, Interf2Sample())
If ierror Then Exit Function

' Return string
Interf2LoadElement = tmsg$
Exit Function

' Errors
Interf2LoadElementError:
MsgBox Error$, vbOKOnly + vbCritical, "Interf2LoadElement"
ierror = True
Exit Function

End Function

