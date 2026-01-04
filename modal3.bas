Attribute VB_Name = "CodeMODAL3"
' (c) Copyright 1995-2026 by John J. Donovan
Option Explicit

Dim ModalAnalysis As TypeAnalysis

Dim ModalTmpSample(1 To 1) As TypeSample
Dim ModalOldSample(1 To 1) As TypeSample

Sub ModalAccumulateSums(adata As Single, average As Single, stddev As Single)
' Perform sum and sum of squares for each data

ierror = False
On Error GoTo ModalAccumulateSumsError

' Calculate sum of data
average! = average! + adata!

' Calculate sum of squares of data
stddev! = stddev! + adata! ^ 2

Exit Sub

' Errors
ModalAccumulateSumsError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalAccumulateSums"
ierror = True
Exit Sub

End Sub

Sub ModalDoModal(ModalGroup As TypeModalGroup)
' Loop on each unknown composition, calculate vector and endmembers

ierror = False
On Error GoTo ModalDoModalError

Dim chan As Integer, ip As Integer, k As Integer
Dim totaloxygen As Single, sum As Single, temp As Single

Dim phanum As Integer, stdnum As Integer
Dim numelms As Integer, numstds As Integer
Dim fitcoeff As Double

Dim linecount As Long
Dim lowvector As Single
Dim phasenumber As Integer
Dim phasevector As Single

Dim pointstotal As Long, pointsvalid As Long, pointsmatch As Long

ReDim phasematched(1 To MAXPHASE%) As Integer
ReDim phaseaveragevector(1 To MAXPHASE%) As Single
ReDim phasestddevvector(1 To MAXPHASE%) As Single

ReDim phaseaveragetotal(1 To MAXPHASE%) As Single
ReDim phasestddevtotal(1 To MAXPHASE%) As Single

ReDim phasemintotal(1 To MAXPHASE%) As Single
ReDim phasemaxtotal(1 To MAXPHASE%) As Single

ReDim phaseaverageendmember(1 To MAXPHASE%, 1 To MAXEND%) As Single
ReDim phaseaverageweights(1 To MAXPHASE%, 1 To MAXCHAN%) As Single

ReDim phasestddevendmember(1 To MAXPHASE%, 1 To MAXEND%) As Single
ReDim phasestddevweights(1 To MAXPHASE%, 1 To MAXCHAN%) As Single

ReDim phaseminendmember(1 To MAXPHASE%, 1 To MAXEND%) As Single
ReDim phaseminweights(1 To MAXPHASE%, 1 To MAXCHAN%) As Single
ReDim phasemaxendmember(1 To MAXPHASE%, 1 To MAXEND%) As Single
ReDim phasemaxweights(1 To MAXPHASE%, 1 To MAXCHAN%) As Single

ReDim upercents(1 To MAXCHAN%) As Double
ReDim tPercents(1 To MAXCHAN%) As Single
ReDim apercents(1 To MAXCHAN%) As Single

ReDim spercents(1 To MAXCHAN%, 1 To MAXSTD%) As Double  ' note reversed column-row order
ReDim allpercents(1 To MAXCHAN%, 1 To MAXSTD%, 1 To MAXPHASE%) As Double  ' note reversed column-row order

Const LARGEVALUE! = 10000#

icancelauto = False

' Load and sort all standard compositions from database
For phanum% = 1 To ModalGroup.NumberofPhases%
For stdnum% = 1 To ModalGroup.NumberofStandards%(phanum%)

Call StandardGetMDBStandard(ModalGroup.StandardNumbers%(phanum%, stdnum%), ModalTmpSample())
If ierror Then Exit Sub

sum! = 0#
For chan% = 1 To ModalOldSample(1).LastChan%
ip% = IPOS1(ModalTmpSample(1).LastChan%, ModalOldSample(1).Elsyms$(chan%), ModalTmpSample(1).Elsyms$())
If ip% > 0 Then
allpercents#(chan%, stdnum%, phanum%) = ModalTmpSample(1).ElmPercents!(ip%)
End If

sum! = sum! + allpercents#(chan%, stdnum%, phanum%)
Next chan%

' Normalize to 100%
If ModalGroup.NormalizeFlag% And sum! > 0# Then
For chan% = 1 To ModalOldSample(1).LastChan%
allpercents#(chan%, stdnum%, phanum%) = 100# * allpercents#(chan%, stdnum%, phanum%) / sum!
Next chan%
End If

' Confirm standard compositions
If DebugMode Then
ip% = StandardGetRow%(ModalGroup.StandardNumbers%(phanum%, stdnum%))
msg$ = "Standard " & StandardGetString$(ip%) & vbCrLf
For chan% = 1 To ModalOldSample(1).LastChan%
msg$ = msg$ & Format$(Format$(allpercents#(chan%, stdnum%, phanum%), f83$), a80$)
Next chan%
Call IOWriteLog(msg$)
End If

Next stdnum%
Next phanum%

' Initialize
pointstotal& = 0
pointsvalid& = 0
pointsmatch& = 0

For phanum% = 1 To ModalGroup.NumberofPhases%
For chan% = 1 To ModalOldSample(1).LastChan%
phaseminweights!(phanum%, chan%) = LARGEVALUE!
Next chan%
For k% = 1 To MAXEND%
phaseminendmember!(phanum%, k%) = LARGEVALUE!
Next k%
phasemintotal!(phanum%) = LARGEVALUE!
Next phanum%

' Loop on each unknown composition in input file
linecount& = 0
Do While Not EOF(ModalInputDataFileNumber%)
linecount& = linecount& + 1

' Increment total number of points
pointstotal& = pointstotal& + 1
Call IOStatusAuto("Performing modal analysis on line " & Str$(linecount&))
DoEvents
If icancelauto Then
ierror = True
Exit Sub
End If

' Zero parameters
For chan% = 1 To MAXCHAN%
upercents#(chan%) = 0#
tPercents!(chan%) = 0#
apercents!(chan%) = 0#
Next chan%

' Load unknown weights from ASCII input data file
For chan% = 1 To ModalOldSample(1).LastElm%
If Not EOF(ModalInputDataFileNumber%) Then
Input #ModalInputDataFileNumber%, tPercents!(chan%)
End If
Next chan%

' Convert to elemental if necessary
totaloxygen! = 0#
For chan% = 1 To ModalOldSample(1).LastChan%
If ModalOldSample(1).OxideOrElemental% = 1 Then
upercents#(chan%) = ConvertOxdToElm!(tPercents!(chan%), ModalOldSample(1).Elsyms$(chan%), ModalOldSample(1).numcat%(chan%), ModalOldSample(1).numoxd%(chan%))
If ierror Then Exit Sub
totaloxygen! = totaloxygen! + (tPercents!(chan%) - upercents#(chan%))

' Just load to unknown array
Else
upercents#(chan%) = tPercents!(chan%)
End If
Next chan%

' Add excess oxygen to oxygen channel
If ModalOldSample(1).OxideOrElemental% = 1 Then
ip% = IPOS1(ModalOldSample(1).LastChan%, Symlo$(ATOMIC_NUM_OXYGEN%), ModalOldSample(1).Elsyms$())
If ip% > 0 Then
upercents#(ip%) = totaloxygen!
End If
End If

' Sum weight percents
sum! = 0#
For chan% = 1 To ModalOldSample(1).LastChan%
sum! = sum! + upercents#(chan%)
apercents!(chan%) = CSng(upercents#(chan%))
Next chan%

' Check for good sum
If sum! < ModalGroup.MinimumTotal! Then
msg$ = Format$(linecount&, a80$)
msg$ = msg$ & Format$("------", a80$)
msg$ = msg$ & Format$("------", a80$)
If ModalGroup.DoEndMember Then
msg$ = msg$ & Format$(Format$(0#, f80$), a80$)
msg$ = msg$ & Format$(Format$(0#, f80$), a80$)
msg$ = msg$ & Format$(Format$(0#, f80$), a80$)
msg$ = msg$ & Format$(Format$(0#, f80$), a80$)
End If
msg$ = msg$ & Format$(Format$(sum!, f82$), a80$)
For chan% = 1 To ModalOldSample(1).LastElm%
msg$ = msg$ & Format$(Format$(tPercents!(chan%), f82$), a80$)
Next chan%

Call IOWriteLog(msg$)
GoTo 2000
End If

' Confirm unknown composition
If DebugMode Then
msg$ = vbCrLf & "Unknown line " & Str$(linecount&) & vbCrLf
For chan% = 1 To ModalOldSample(1).LastChan%
msg$ = msg$ & Format$(ModalOldSample(1).Elsyms$(chan%), a80$)
Next chan%
Call IOWriteLog(msg$)
msg$ = vbNullString
For chan% = 1 To ModalOldSample(1).LastChan%
msg$ = msg$ & Format$(Format$(upercents#(chan%), f83$), a80$)
Next chan%
Call IOWriteLog(msg$)
End If

' Normalize to 100%
If ModalGroup.NormalizeFlag% And sum! > 0# Then
For chan% = 1 To ModalOldSample(1).LastChan%
upercents#(chan%) = 100# * upercents#(chan%) / sum!
Next chan%
End If

' Initialize variables
lowvector! = 100000#
phasevector! = 0#
phasenumber% = 0

' Loop on each phase in group
For phanum% = 1 To ModalGroup.NumberofPhases%

' Load standard compositions for this phase
For stdnum% = 1 To MAXSTD%
For chan% = 1 To MAXCHAN%
spercents#(chan%, stdnum%) = allpercents#(chan%, stdnum%, phanum%)
Next chan%
Next stdnum%

' Fit unknown weights to standard phase
numelms% = ModalOldSample(1).LastChan%
numstds% = ModalGroup.NumberofStandards%(phanum%)
Call ModalFitModal(numelms%, numstds%, upercents#(), spercents#(), fitcoeff#, ModalGroup.weightflag%)
If ierror Then Exit Sub

' Store fit coefficients for this phase
If fitcoeff# < lowvector! Then
lowvector! = fitcoeff#
If lowvector! < ModalGroup.MinimumVectors!(phanum%) Then
phasevector! = lowvector!
phasenumber% = phanum%
End If
End If

Next phanum%

' Increment valid number of points
pointsvalid& = pointsvalid& + 1

' No phase matched unknown composition
If phasenumber% = 0 Then
msg$ = Format$(linecount&, a80$)
msg$ = msg$ & Format$(Format$(lowvector!, f82$), a80$)
msg$ = msg$ & Format$("------", a80$)

If ModalGroup.DoEndMember Then
msg$ = msg$ & Format$(Format$(0#, f80$), a80$)
msg$ = msg$ & Format$(Format$(0#, f80$), a80$)
msg$ = msg$ & Format$(Format$(0#, f80$), a80$)
msg$ = msg$ & Format$(Format$(0#, f80$), a80$)
End If

msg$ = msg$ & Format$(Format$(sum!, f82$), a80$)
For chan% = 1 To ModalOldSample(1).LastElm%
msg$ = msg$ & Format$(Format$(tPercents!(chan%), f82$), a80$)
Next chan%

' Phase matched unknown composition
Else

' Increment match number of points
pointsmatch& = pointsmatch& + 1
phasematched%(phasenumber%) = phasematched%(phasenumber%) + 1

' Calculate running average and standard deviations for vectors
Call ModalAccumulateSums(phasevector!, phaseaveragevector!(phasenumber%), phasestddevvector!(phasenumber%))
If ierror Then Exit Sub

' Load data for output
msg$ = Format$(linecount&, a80$)
msg$ = msg$ & Format$(Format$(lowvector!, f82$), a80$)
msg$ = msg$ & Format$(Left$(ModalGroup.PhaseNames$(phasenumber%), 7), a80$)

' Calculate end-member composition
If ModalGroup.DoEndMember Then

If ModalGroup.EndMemberNumbers%(phasenumber%) > 0 Then
ModalOldSample(1).MineralFlag% = ModalGroup.EndMemberNumbers%(phasenumber%)
ModalOldSample(1).Datarows% = 1
ModalOldSample(1).LineStatus%(ModalOldSample(1).Datarows%) = True

' Calculate atomic percents
For chan% = 1 To ModalOldSample(1).LastChan%
ModalAnalysis.CalData!(ModalOldSample(1).Datarows%, chan%) = ConvertWeightToAtom(ModalOldSample(1).LastChan%, chan%, apercents!(), ModalOldSample(1).Elsyms$())
If ierror Then Exit Sub
Next chan%

' Calculate mineral end members
Call ConvertMinerals(ModalAnalysis, ModalOldSample())
If ierror Then Exit Sub

' Load end-member data
For k% = 1 To MAXEND%
msg$ = msg$ & ModalFormatEndMember$(ModalGroup.EndMemberNumbers%(phasenumber%), k%, ModalAnalysis.CalData!(ModalOldSample(1).Datarows%, k%))
Next k%

' Calculate running average and standard deviations (and min/max) for end-members
For k% = 1 To MAXEND%
Call ModalAccumulateSums(ModalAnalysis.CalData!(ModalOldSample(1).Datarows%, k%), phaseaverageendmember!(phasenumber%, k%), phasestddevendmember!(phasenumber%, k%))
If ierror Then Exit Sub

If ModalAnalysis.CalData!(ModalOldSample(1).Datarows%, k%) < phaseminendmember!(phasenumber%, k%) Then
phaseminendmember!(phasenumber%, k%) = ModalAnalysis.CalData!(ModalOldSample(1).Datarows%, k%)
End If
If ModalAnalysis.CalData!(ModalOldSample(1).Datarows%, k%) > phasemaxendmember!(phasenumber%, k%) Then
phasemaxendmember!(phasenumber%, k%) = ModalAnalysis.CalData!(ModalOldSample(1).Datarows%, k%)
End If
Next k%

' Just load empty space
Else
msg$ = msg$ & Format$(Format$(0#, f80$), a80$)
msg$ = msg$ & Format$(Format$(0#, f80$), a80$)
msg$ = msg$ & Format$(Format$(0#, f80$), a80$)
msg$ = msg$ & Format$(Format$(0#, f80$), a80$)
End If
End If

' Load sum
msg$ = msg$ & Format$(Format$(sum!, f82$), a80$)

' Calculate running average and standard deviations for total
Call ModalAccumulateSums(sum!, phaseaveragetotal!(phasenumber%), phasestddevtotal!(phasenumber%))
If ierror Then Exit Sub

If sum! < phasemintotal!(phasenumber%) Then
phasemintotal!(phasenumber%) = sum!
End If
If sum > phasemaxtotal!(phasenumber%) Then
phasemaxtotal!(phasenumber%) = sum!
End If

' Load original weight percents from file
For chan% = 1 To ModalOldSample(1).LastElm%
msg$ = msg$ & Format$(Format$(tPercents!(chan%), f82$), a80$)

' Calculate running average and standard deviations for weight percents
Call ModalAccumulateSums(tPercents!(chan%), phaseaverageweights!(phasenumber%, chan%), phasestddevweights!(phasenumber%, chan%))
If ierror Then Exit Sub

If tPercents!(chan%) < phaseminweights!(phasenumber%, chan%) Then
phaseminweights!(phasenumber%, chan%) = tPercents!(chan%)
End If
If tPercents!(chan%) > phasemaxweights!(phasenumber%, chan%) Then
phasemaxweights!(phasenumber%, chan%) = tPercents!(chan%)
End If

Next chan%
End If

' Output to screen and file
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$

2000:
DoEvents
Loop

' Set back to zero if not used
For phanum% = 1 To ModalGroup.NumberofPhases%
For chan% = 1 To ModalOldSample(1).LastChan%
If phaseminweights!(phanum%, chan%) = LARGEVALUE! Then phaseminweights!(phanum%, chan%) = 0#
Next chan%
For k% = 1 To MAXEND%
If phaseminendmember!(phanum%, k%) = LARGEVALUE! Then phaseminendmember!(phanum%, k%) = 0#
Next k%
If phasemintotal!(phanum%) = LARGEVALUE! Then phasemintotal!(phanum%) = 0#
Next phanum%

' Display results
msg$ = vbCrLf & "Results of Modal Analysis" & vbCrLf
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$
msg$ = "InputFile    : " & ModalInputDataFile$
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$
msg$ = "OutputFile   : " & ModalOutputDataFile$
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$
msg$ = "Date and Time: " & Now
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$

msg$ = vbCrLf & "Group Name   : " & ModalGroup.GroupName$
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$
msg$ = "Total Number of Points in File : " & Str$(pointstotal&)
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$
msg$ = "Valid Number of Points in File : " & Str$(pointsvalid&)
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$
msg$ = "Match Number of Points in File : " & Str$(pointsmatch&)
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$

msg$ = vbCrLf & "Minimum Total for Valid Points : " & Format$(Format$(ModalGroup.MinimumTotal, f82$), a80$)
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$

If pointstotal& > 0 Then
temp! = CSng(pointsvalid&) / pointstotal& * 100#
msg$ = "Percentage of Valid Points : " & Format$(Format$(temp!, f81$), a80$)
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$
temp! = CSng(pointsmatch&) / pointstotal& * 100#
msg$ = "Percentage of Match Points : " & Format$(Format$(temp!, f81$), a80$)
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$
End If

' Loop on each phase in group
For phanum% = 1 To ModalGroup.NumberofPhases%

' Phase info titles
msg$ = vbCrLf & Format$("Phase", a80$)
msg$ = msg$ & Format$("#Match", a80$)
msg$ = msg$ & Format$("%Total", a80$)
msg$ = msg$ & Format$("%Valid", a80$)
msg$ = msg$ & Format$("%Match", a80$)
msg$ = msg$ & Format$("AvgVec", a80$)
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$

' Phase info
msg$ = Format$(Left$(ModalGroup.PhaseNames$(phanum%), 8), a80$)
msg$ = msg$ & Format$(phasematched%(phanum%), a80$)

If pointstotal& > 0 Then
temp! = CSng(phasematched%(phanum%)) / pointstotal& * 100#
msg$ = msg$ & Format$(Format$(temp!, f81$), a80$)
End If

If pointsvalid& > 0 Then
temp! = CSng(phasematched%(phanum%)) / pointsvalid& * 100#
msg$ = msg$ & Format$(Format$(temp!, f81$), a80$)
End If

If pointsmatch& > 0 Then
temp! = CSng(phasematched%(phanum%)) / pointsmatch& * 100#
msg$ = msg$ & Format$(Format$(temp!, f81$), a80$)
End If

temp! = ModalGetAverage!(phasematched%(phanum%), phaseaveragevector!(phanum%))
msg$ = msg$ & Format$(Format$(temp!, f82$), a80$)
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$

' Average and standard deviation titles
msg$ = Space$(8)
If ModalGroup.DoEndMember Then
msg$ = msg$ & Format$(" ", a80$)
msg$ = msg$ & Format$("End  -", a80$)
msg$ = msg$ & Format$("Member", a80$)
msg$ = msg$ & Format$(" ", a80$)
End If

msg$ = msg$ & Format$("Sum", a80$)
For chan% = 1 To ModalOldSample(1).LastElm%
If ModalOldSample(1).OxideOrElemental% = 1 Then
msg$ = msg$ & Format$(ModalOldSample(1).Oxsyup$(chan%), a80$)
Else
msg$ = msg$ & Format$(ModalOldSample(1).Elsyup$(chan%), a80$)
End If
Next chan%
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$

' Averages
msg$ = Format$("Average:", a80$)

If ModalGroup.DoEndMember Then
For k% = 1 To MAXEND%
temp! = ModalGetAverage!(phasematched%(phanum%), phaseaverageendmember!(phanum%, k%))
msg$ = msg$ & ModalFormatEndMember$(ModalGroup.EndMemberNumbers%(phanum%), k%, temp!)
Next k%
End If

temp! = ModalGetAverage!(phasematched%(phanum%), phaseaveragetotal!(phanum%))
msg$ = msg$ & Format$(Format$(temp!, f82$), a80$)

For chan% = 1 To ModalOldSample(1).LastElm%
temp! = ModalGetAverage!(phasematched%(phanum%), phaseaverageweights!(phanum%, chan%))
msg$ = msg$ & Format$(Format$(temp!, f82$), a80$)
Next chan%
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$

' Standard deviations
msg$ = Format$("Std Dev:", a80$)

If ModalGroup.DoEndMember Then
For k% = 1 To MAXEND%
temp! = ModalGetStdDev!(phasematched%(phanum%), phaseaverageendmember!(phanum%, k%), phasestddevendmember!(phanum%, k%))
msg$ = msg$ & Format$(Format$(temp!, f80$), a80$)
Next k%
End If

temp! = ModalGetStdDev!(phasematched%(phanum%), phaseaveragetotal!(phanum%), phasestddevtotal!(phanum%))
msg$ = msg$ & Format$(Format$(temp!, f82$), a80$)

For chan% = 1 To ModalOldSample(1).LastElm%
temp! = ModalGetStdDev!(phasematched%(phanum%), phaseaverageweights!(phanum%, chan%), phasestddevweights!(phanum%, chan%))
msg$ = msg$ & Format$(Format$(temp!, f82$), a80$)
Next chan%
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$

' Minimum
msg$ = Format$("Minimum:", a80$)

If ModalGroup.DoEndMember Then
For k% = 1 To MAXEND%
temp! = phaseminendmember!(phanum%, k%)
msg$ = msg$ & Format$(Format$(temp!, f80$), a80$)
Next k%
End If

temp! = phasemintotal!(phanum%)
msg$ = msg$ & Format$(Format$(temp!, f82$), a80$)

For chan% = 1 To ModalOldSample(1).LastElm%
temp! = phaseminweights!(phanum%, chan%)
msg$ = msg$ & Format$(Format$(temp!, f82$), a80$)
Next chan%
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$

' Maximum
msg$ = Format$("Maximum:", a80$)

If ModalGroup.DoEndMember Then
For k% = 1 To MAXEND%
temp! = phasemaxendmember!(phanum%, k%)
msg$ = msg$ & Format$(Format$(temp!, f80$), a80$)
Next k%
End If

temp! = phasemaxtotal!(phanum%)
msg$ = msg$ & Format$(Format$(temp!, f82$), a80$)

For chan% = 1 To ModalOldSample(1).LastElm%
temp! = phasemaxweights!(phanum%, chan%)
msg$ = msg$ & Format$(Format$(temp!, f82$), a80$)
Next chan%
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$

Next phanum%

Exit Sub

' Errors
ModalDoModalError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalDoModal"
ierror = True
Exit Sub

End Sub

Function ModalGetAverage(npts As Integer, sum As Single) As Single
' Calculate the average based on the sum and number of observations

ierror = False
On Error GoTo ModalGetAverageError

' Calculate average
If npts% > 0 Then
ModalGetAverage! = sum! / npts%
Else
ModalGetAverage! = 0
End If

Exit Function

' Errors
ModalGetAverageError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalGetAverage"
ierror = True
Exit Function

End Function

Function ModalGetStdDev(npts As Integer, sum As Single, sumsq As Single) As Single
' Calculate the standard deviation based on the sum of squares and number of observations

ierror = False
On Error GoTo ModalGetStdDevError

' Calculate average
If npts% > 1 Then
ModalGetStdDev! = Sqr((sumsq! - npts% * (sum! / npts%) ^ 2) / (npts% - 1))
Else
ModalGetStdDev! = 0
End If

Exit Function

' Errors
ModalGetStdDevError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalGetStdDev"
ierror = True
Exit Function

End Function

Sub ModalReadColumnLabels(sample() As TypeSample)
' Read the column labels of the input data file

ierror = False
On Error GoTo ModalReadColumnLabelsError

Dim astring As String, bstring As String
Dim i As Integer, j As Integer, ip As Integer
Dim oxidefound As Integer, ncols As Integer
Dim numelms As Integer, cat As Integer, oxd As Integer
Dim icat As Integer, ioxd As Integer
Dim weight As Single

ReDim symbols(1 To MAXCHAN%) As String
ReDim fatoms(1 To MAXCHAN%) As Single

' Determine number of columns in file
ncols% = MiscGetNumberofColumns(ModalInputDataFile$)
If ierror Then Exit Sub

' Read and parse element symbols (if oxide)
oxidefound% = False
sample(1).LastElm% = 0
For j% = 1 To ncols%

' If data is separated by tabs only, then VB fails to read correctly!
Input #ModalInputDataFileNumber%, astring$
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub

' Get elements and oxygens
Call MWCalculate(bstring$, numelms%, symbols$(), fatoms!(), weight!)
If ierror Then Exit Sub

' Check for elements
If numelms% < 1 Then GoTo ModalReadColumnLabelsBadLabel

' Determine which is cation and which is oxygen
If symbols$(1) <> Symup$(ATOMIC_NUM_OXYGEN%) Then
icat% = 1
ioxd% = 2

ElseIf symbols$(1) = Symup$(ATOMIC_NUM_OXYGEN%) And symbols$(2) <> vbNullString Then
icat% = 2
ioxd% = 1

ElseIf symbols$(1) = Symup$(ATOMIC_NUM_OXYGEN%) And symbols$(2) = vbNullString Then
icat% = 1
ioxd% = 2

Else
GoTo ModalReadColumnLabelsBadCatOxd
End If

' Check for valid element cation
ip% = IPOS1(MAXELM%, symbols$(icat%), Symlo$())
If ip% = 0 Then GoTo ModalReadColumnLabelsBadElement

' Load default cation and oxygens
cat% = 1
oxd% = 0

' Check for valid oxygen
If symbols$(ioxd%) <> vbNullString Then
ip% = IPOS1(MAXELM%, symbols$(ioxd%), Symlo$())
If ip% = 0 Then GoTo ModalReadColumnLabelsBadElement

If symbols$(ioxd%) <> Symup$(ATOMIC_NUM_OXYGEN%) Then GoTo ModalReadColumnLabelsNoOxygen

' Load oxide cations
cat% = CInt(fatoms!(icat%))
oxd% = CInt(fatoms!(ioxd%))
oxidefound% = True
End If

' Load new element arrays
sample(1).LastElm% = sample(1).LastElm% + 1
sample(1).Elsyms$(sample(1).LastElm%) = LCase$(symbols$(icat%))
sample(1).numcat%(sample(1).LastElm%) = cat%
sample(1).numoxd%(sample(1).LastElm%) = oxd%

Next j%

' Save oxide or elemental flag
If oxidefound% Then
sample(1).OxideOrElemental% = 1
Else
sample(1).OxideOrElemental% = 2
End If

' Check number of elements
If sample(1).LastElm% < 1 Then GoTo ModalReadColumnLabelsNoElements
sample(1).LastChan% = sample(1).LastElm%

' If oxide (and no oxygen already), add oxygen
If ModalOldSample(1).OxideOrElemental% = 1 Then
ip% = IPOS1(ModalOldSample(1).LastChan%, Symlo$(ATOMIC_NUM_OXYGEN%), ModalOldSample(1).Elsyms$())
If ip% = 0 Then
If sample(1).LastChan% + 1 > MAXCHAN% Then GoTo ModalReadColumnLabelsNoRoom
sample(1).LastChan% = sample(1).LastChan% + 1
sample(1).Elsyms$(sample(1).LastChan%) = Symlo$(ATOMIC_NUM_OXYGEN%)
sample(1).numcat%(sample(1).LastChan%) = 1
sample(1).numoxd%(sample(1).LastChan%) = 0
End If
End If

' Load defaults
sample(1).takeoff! = DefaultTakeOff!
sample(1).kilovolts! = DefaultKiloVolts!
sample(1).beamcurrent! = DefaultBeamCurrent!
sample(1).beamsize! = DefaultBeamSize!

' Load kilovolts array
For i% = 1 To sample(1).LastChan%
sample(1).TakeoffArray!(i%) = sample(1).takeoff!
sample(1).KilovoltsArray!(i%) = sample(1).kilovolts!
sample(1).BeamCurrentArray(i%) = sample(1).beamcurrent!
sample(1).BeamSizeArray(i%) = sample(1).beamsize!
Next i%

' Load x-ray lines from element defaults
For i% = 1 To sample(1).LastChan%
ip% = IPOS1(MAXELM%, sample(1).Elsyms$(i%), Symlo$())
If ip% > 0 Then sample(1).Xrsyms$(i%) = Deflin$(ip%)

' Make hydrogen and helium absorber only
If ip% = 1 Or ip% = 2 Then sample(1).Xrsyms$(i%) = vbNullString

Next i%

' Load element parameters
Call ElementGetData(sample())
If ierror Then Exit Sub

Exit Sub

' Errors
ModalReadColumnLabelsError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalReadColumnLabels"
ierror = True
Exit Sub

ModalReadColumnLabelsBadElement:
msg$ = "Bad element in label " & bstring$
MsgBox msg$, vbOKOnly + vbExclamation, "ModalReadColumnLabels"
ierror = True
Exit Sub

ModalReadColumnLabelsBadLabel:
msg$ = "Bad element label " & bstring$
MsgBox msg$, vbOKOnly + vbExclamation, "ModalReadColumnLabels"
ierror = True
Exit Sub

ModalReadColumnLabelsBadCatOxd:
msg$ = "Bad cation or oxygen in label " & bstring$
MsgBox msg$, vbOKOnly + vbExclamation, "ModalReadColumnLabels"
ierror = True
Exit Sub

ModalReadColumnLabelsNoOxygen:
msg$ = "No oxygen in oxide formula label " & bstring$
MsgBox msg$, vbOKOnly + vbExclamation, "ModalReadColumnLabels"
ierror = True
Exit Sub

ModalReadColumnLabelsNoElements:
msg$ = "No elements in " & ModalInputDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "ModalReadColumnLabels"
ierror = True
Exit Sub

ModalReadColumnLabelsNoRoom:
msg$ = "No room to add oxygen in element list"
MsgBox msg$, vbOKOnly + vbExclamation, "ModalReadColumnLabels"
ierror = True
Exit Sub

End Sub

Sub ModalRunModal(ModalGroup As TypeModalGroup)
' Run the modal analysis

ierror = False
On Error GoTo ModalRunModalError

Dim i As Integer, j As Integer
Dim k As Integer, ip As Integer

' Init analysis arrays (for end-members)
Call InitStandards(ModalAnalysis)
If ierror Then Exit Sub

' Read columns labels of input file
Call ModalReadColumnLabels(ModalOldSample())
If ierror Then Exit Sub

' Check number of elements and standards
If ModalOldSample(1).LastChan% < 1 Then GoTo ModalRunModalNoElements
If ModalGroup.NumberofPhases% < 1 Then GoTo ModalRunModalNoPhases
For i% = 1 To ModalGroup.NumberofPhases%
If ModalGroup.NumberofStandards%(i%) < 1 Then GoTo ModalRunModalNoStandards
Next i%

' Loop on each standard and check that it is in standard database
For i% = 1 To ModalGroup.NumberofPhases%
For j% = 1 To ModalGroup.NumberofStandards%(i%)
ip% = StandardGetRow%(ModalGroup.StandardNumbers%(i%, j%))
If ip% = 0 Then GoTo ModalRunModalStdNotFound
Next j%
Next i%

' Check for elements in standards, that are not in unknown
For i% = 1 To ModalGroup.NumberofPhases%
For j% = 1 To ModalGroup.NumberofStandards%(i%)

Call StandardGetMDBStandard(ModalGroup.StandardNumbers%(i%, j%), ModalTmpSample())
If ierror Then Exit Sub

For k% = 1 To ModalTmpSample(1).LastChan%
ip% = IPOS1(ModalOldSample(1).LastChan%, ModalTmpSample(1).Elsyms$(k%), ModalOldSample(1).Elsyms$())
If ip% = 0 Then
If ModalOldSample(1).LastChan% + 1 <= MAXSTD% Then
ModalOldSample(1).LastChan% = ModalOldSample(1).LastChan% + 1
ModalOldSample(1).Elsyms$(ModalOldSample(1).LastChan%) = ModalTmpSample(1).Elsyms$(k%)
ModalOldSample(1).Xrsyms$(ModalOldSample(1).LastChan%) = vbNullString
ip% = IPOS1(MAXELM%, ModalOldSample(1).Elsyms$(ModalOldSample(1).LastChan%), Symlo$())
ModalOldSample(1).numcat%(ModalOldSample(1).LastChan%) = AllCat%(ip%)
ModalOldSample(1).numoxd%(ModalOldSample(1).LastChan%) = AllOxd%(ip%)
End If
End If
Next k%

Next j%
Next i%

' Check for more standards than elements in each phase
For i% = 1 To ModalGroup.NumberofPhases%
If ModalOldSample(1).LastChan% <= ModalGroup.NumberofStandards%(i%) Then GoTo ModalRunModalTooManyStandards
Next i%

' Load element parameters
Call ElementGetData(ModalOldSample())
If ierror Then Exit Sub

' Print group summary
Call ModalGroupPrint
If ierror Then Exit Sub

' Print column labels
msg$ = vbCrLf & Format$("Line", a80$)
msg$ = msg$ & Format$("Vector", a80$)
msg$ = msg$ & Format$("Phase", a80$)
If ModalGroup.DoEndMember Then
msg$ = msg$ & Format$(" ", a80$)
msg$ = msg$ & Format$("End  -", a80$)
msg$ = msg$ & Format$("Member", a80$)
msg$ = msg$ & Format$(" ", a80$)
End If
msg$ = msg$ & Format$("Sum", a80$)
For i% = 1 To ModalOldSample(1).LastElm%
If ModalOldSample(1).OxideOrElemental% = 1 Then
msg$ = msg$ & Format$(ModalOldSample(1).Oxsyup$(i%), a80$)
Else
msg$ = msg$ & Format$(ModalOldSample(1).Elsyup$(i%), a80$)
End If
Next i%
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$

' Loop on unknown compositions
Call IOStatusAuto(vbNullString)
Call ModalDoModal(ModalGroup)
Call IOStatusAuto(vbNullString)
If ierror Then Exit Sub

Exit Sub

' Errors
ModalRunModalError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalRunModal"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

ModalRunModalNoElements:
msg$ = "No input elements found in " & ModalInputDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "ModalRunModal"
ierror = True
Exit Sub

ModalRunModalNoPhases:
msg$ = "No modal phases specified in group" & ModalGroup.GroupName$
MsgBox msg$, vbOKOnly + vbExclamation, "ModalRunModal"
ierror = True
Exit Sub

ModalRunModalNoStandards:
msg$ = "No standards specified in phase " & ModalGroup.PhaseNames$(i%)
MsgBox msg$, vbOKOnly + vbExclamation, "ModalRunModal"
ierror = True
Exit Sub

ModalRunModalStdNotFound:
msg$ = "Standard number " & Str$(ModalGroup.StandardNumbers%(i%, j%)) & " was not found in the database"
MsgBox msg$, vbOKOnly + vbExclamation, "ModalRunModal"
ierror = True
Exit Sub

ModalRunModalTooManyStandards:
msg$ = "Too many standards defined in modal phase " & ModalGroup.PhaseNames$(i%)
MsgBox msg$, vbOKOnly + vbExclamation, "ModalRunModal"
ierror = True
Exit Sub

End Sub

Sub ModalTestModal()
' Test ModalFitModal

ierror = False
On Error GoTo ModalTestModalError

Dim numelms As Integer, numstds As Integer, weightflag As Integer
Dim fitcoeff As Double

ReDim upercents(1 To MAXCHAN%) As Double
ReDim spercents(1 To MAXCHAN%, 1 To MAXSTD%) As Double  ' note reversed column-row order

numelms% = 4 ' fe, si, o and mg
numstds% = 2    ' 273 and 263
upercents#(1) = 54#
upercents#(2) = 14#
upercents#(3) = 32#

spercents#(1, 1) = 0#
spercents#(2, 1) = 19.96
spercents#(3, 1) = 45.48
spercents#(4, 1) = 34.55

spercents#(1, 2) = 54.81
spercents#(2, 2) = 13.78
spercents#(3, 2) = 31.41
spercents#(4, 2) = 0#

Call ModalFitModal(numelms%, numstds%, upercents#(), spercents#(), fitcoeff#, weightflag%)
If ierror Then Exit Sub

msg$ = "Fit coefficient = " & Str$(fitcoeff#)
Call IOWriteLog(msg$)

Exit Sub

' Errors
ModalTestModalError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalTestModal"
ierror = True
Exit Sub

End Sub

