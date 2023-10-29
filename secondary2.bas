Attribute VB_Name = "CodeSecondary2"
' (c) Copyright 1995-2023 by John J. Donovan
Option Explicit

' Analyzed point coordinates
Dim curpnt As Long
Dim apoints As Long
Dim xcoord() As Single
Dim ycoord() As Single
Dim zcoord() As Single

' Raw data from a single kratios.DAT file (for storage)
Dim k_npts As Long
Dim k_string1 As String, k_string2 As String, k_string3 As String
Dim k_eV() As Double, k_dist() As Double
Dim k_total() As Double, k_fluor() As Double
Dim k_flach() As Double, k_flabr() As Double
Dim k_flbch() As Double, k_flbbr() As Double
Dim k_pri_int() As Double, k_std_int() As Double

' Processed K-ratio values from all Kratios.DAT files (all channels)
Dim maxpoints As Long                   ' maximum number of kratios for all channels
Dim nPoints(1 To MAXCHAN%) As Long      ' number of k-ratio data points for each channel

Dim xdist() As Double       ' linear distance (um) for each k-ratio
Dim yktotal() As Double     ' fluorescence kratio% plus primary x-rays kratio% from material A and material B
Dim ykfluor() As Double     ' fluorescence kratio% only (minus primary x-ray kratio% from material A)

Dim fluA_k() As Double        ' Mat A total fluorescence k-ratio %
Dim fluB_k() As Double        ' Mat B total fluorescence k-ratio %
Dim prix_k() As Double        ' Primary x-ray k-ratio %

Sub SecondaryInit1()
' Initialize module level variables for x, y coordinates

ierror = False
On Error GoTo SecondaryInit1Error

' Init module level data point coordinate arrays
curpnt& = 0
apoints& = 1
ReDim xcoord(1 To apoints&) As Single
ReDim ycoord(1 To apoints&) As Single
ReDim zcoord(1 To apoints&) As Single

Exit Sub

' Errors
SecondaryInit1Error:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryInit1"
ierror = True
Exit Sub

End Sub

Sub SecondaryInit2()
' Initialize module level variables for kratio data points

ierror = False
On Error GoTo SecondaryInit2Error

Dim n As Long

' Init number of k-ratio data points
maxpoints& = 1
For n& = 1 To MAXCHAN%
nPoints&(n&) = 0
Next n&

' Dimension variables
ReDim xdist(1 To MAXCHAN%, 1 To maxpoints&) As Double
ReDim yktotal(1 To MAXCHAN%, 1 To maxpoints&) As Double
ReDim ykfluor(1 To MAXCHAN%, 1 To maxpoints&) As Double

ReDim fluA_k(1 To MAXCHAN%, 1 To maxpoints&) As Double
ReDim fluB_k(1 To MAXCHAN%, 1 To maxpoints&) As Double
ReDim prix_k(1 To MAXCHAN%, 1 To maxpoints&) As Double

Exit Sub

' Errors
SecondaryInit2Error:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryInit2"
ierror = True
Exit Sub

End Sub

Sub SecondaryCorrection(sampleline As Integer, kratios() As Single, sample() As TypeSample)
' Perform the secondary boundary fluorescence correction (kratios must be elemental normalized kratios)

ierror = False
On Error GoTo SecondaryCorrectionError

Dim chan As Integer
Dim astring1 As String, astring2 As String, astring3 As String, astring4 As String

' Calculate the k-ratio for the interpolated distance for each element
For chan% = 1 To sample(1).LastElm%
If sample(1).SecondaryFluorescenceBoundaryFlag(chan%) = True Then

' Check if k-ratios are loaded for this channel
If nPoints&(chan%) <= 0 Then GoTo SecondaryCorrectionNoKratios

' Calculate boundary correction k-ratio for this data line and channel based on MatAB minus MatA
Call SecondaryCalculateKratio(sampleline%, chan%, sample())
If ierror Then Exit Sub

' Save string info
astring1$ = astring1$ & Format$(sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%), a80$)
astring2$ = astring2$ & MiscAutoFormat$(kratios!(chan%) * 100#)

' Correct the measured (elemental) k-ratio for the boundary correction intensity (in k-ratio % units)
kratios!(chan%) = kratios!(chan%) - sample(1).SecondaryFluorescenceBoundaryKratios!(sampleline%, chan%) / 100#

' Save string info
astring3$ = astring3$ & MiscAutoFormat$(sample(1).SecondaryFluorescenceBoundaryKratios!(sampleline%, chan%) / 100# * 100#)
astring4$ = astring4$ & MiscAutoFormat$(kratios!(chan%) * 100#)

End If
Next chan%

' Debug output
'If DebugMode Then
Call IOWriteLog(vbCrLf & "SecondaryCorrection: SF k-ratios * 100, Line: " & Format$(sample(1).Linenumber&(sampleline%)) & ", Dist: " & sample(1).SecondaryFluorescenceBoundaryDistance!(sampleline%))
Call IOWriteLog(Format$("Element:", a80$) & astring1$)
Call IOWriteLog(Format$("Elm. Kr:", a80$) & astring2$)
Call IOWriteLog(Format$("Cal. Kr:", a80$) & astring3$)
Call IOWriteLog(Format$("Cor. Kr:", a80$) & astring4$)
'End If

' Load data point coordinate to module level arrays (for display)
curpnt& = curpnt& + 1

If curpnt& > apoints& Then
ReDim Preserve xcoord(1 To curpnt&) As Single
ReDim Preserve ycoord(1 To curpnt&) As Single
ReDim Preserve zcoord(1 To curpnt&) As Single
apoints& = curpnt&
End If

' Add current data point to arrays
xcoord!(curpnt&) = sample(1).StagePositions!(sampleline%, 1)
ycoord!(curpnt&) = sample(1).StagePositions!(sampleline%, 2)
zcoord!(curpnt&) = sample(1).StagePositions!(sampleline%, 3)

Exit Sub

' Errors
SecondaryCorrectionError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryCorrection"
ierror = True
Exit Sub

SecondaryCorrectionNoKratios:
msg$ = "No secondary fluorescence k-ratio values are currently loaded for channel " & Format$(chan%) & ". Please select a k-ratio data file from a Fanal calculation appropriate for this situation."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryCorrection"
ierror = True
Exit Sub

End Sub

Sub SecondaryReadKratiosDATFile(tfilename As String, sample() As TypeSample)
' Reads in the selected k-ratios.dat file from Fanal calculation for the boundary correction and store at module level

ierror = False
On Error GoTo SecondaryReadKratiosDATFileError

Dim ip As Integer, n As Integer
Dim astring As String, bstring As String
Dim esym As String, trans As String
Dim keV As Single

' Check for valid k-ratio file
If Trim$(tfilename$) = vbNullString Then GoTo SecondaryReadKratiosDATFileNoFilename
If Dir$(tfilename$) = vbNullString Then GoTo SecondaryReadKratiosDATFileNoFileFound

Close #Temp2FileNumber%
DoEvents
Open tfilename$ For Input As #Temp2FileNumber%

' Read characteristic line
Line Input #Temp2FileNumber%, k_string1$

' Read electron energy line
Line Input #Temp2FileNumber%, k_string2$

' Read column labels line
Line Input #Temp2FileNumber%, k_string3$

' Parse element
n% = InStr(k_string1$, "#  Characteristic line: ")
If n% = 0 Then GoTo SecondaryReadKratiosDATFileBadFirstLine

astring$ = Mid$(k_string1$, Len("#  Characteristic line: ") + 1)
Call MiscParseStringToStringA(astring$, VbSpace, bstring$)
If ierror Then Exit Sub

esym$ = Symlo$(Val(Trim$(bstring$)))
ip% = IPOS1%(sample(1).LastElm%, esym$, sample(1).Elsyms$())
If ip% = 0 Then GoTo SecondaryReadKratiosDATFileElementNotFound

' Parse x-ray (see table 6.2 in Penelope-2006-NEA-pdf)
trans$ = Trim$(astring$)
If sample(1).XrayNums%(ip%) = 1 And Trim$(trans$) <> "K L3" Then GoTo SecondaryReadKratiosDATFileXrayIsDifferent
If sample(1).XrayNums%(ip%) = 2 And Trim$(trans$) <> "K M3" Then GoTo SecondaryReadKratiosDATFileXrayIsDifferent
If sample(1).XrayNums%(ip%) = 3 And Trim$(trans$) <> "L3 M5" Then GoTo SecondaryReadKratiosDATFileXrayIsDifferent
If sample(1).XrayNums%(ip%) = 4 And Trim$(trans$) <> "L2 M4" Then GoTo SecondaryReadKratiosDATFileXrayIsDifferent
If sample(1).XrayNums%(ip%) = 5 And Trim$(trans$) <> "M5 N7" Then GoTo SecondaryReadKratiosDATFileXrayIsDifferent
If sample(1).XrayNums%(ip%) = 6 And Trim$(trans$) <> "M4 N6" Then GoTo SecondaryReadKratiosDATFileXrayIsDifferent

'If sample(1).XrayNums%(ip%) = 7 And Trim$(trans$) <> "L2 M1" Then GoTo SecondaryReadKratiosDATFileXrayIsDifferent
'If sample(1).XrayNums%(ip%) = 8 And Trim$(trans$) <> "L2 N4" Then GoTo SecondaryReadKratiosDATFileXrayIsDifferent
'If sample(1).XrayNums%(ip%) = 9 And Trim$(trans$) <> "L2 N6" Then GoTo SecondaryReadKratiosDATFileXrayIsDifferent
'If sample(1).XrayNums%(ip%) = 10 And Trim$(trans$) <> "L3 M1" Then GoTo SecondaryReadKratiosDATFileXrayIsDifferent
'If sample(1).XrayNums%(ip%) = 11 And Trim$(trans$) <> "M3 N5" Then GoTo SecondaryReadKratiosDATFileXrayIsDifferent
'If sample(1).XrayNums%(ip%) = 12 And Trim$(trans$) <> "M5 N3" Then GoTo SecondaryReadKratiosDATFileXrayIsDifferent

' Parse keV
n% = InStr(UCase$(k_string2$), UCase$("#  e0 (eV) = "))
If n% = 0 Then GoTo SecondaryReadKratiosDATFileBadSecondLine

astring$ = Mid$(k_string2$, Len("#  e0 (eV) = ") + 1)
keV! = CDbl(astring$)      ' use double for language issues
keV! = keV! / EVPERKEV#    ' convert to keV

' Now check keV since element was found
If keV! <> sample(1).KilovoltsArray!(ip%) Then GoTo SecondaryReadKratiosDATFileDifferentKilovolts

' Set sample channel flag for secondary boundary fluorescence correction (for CalcZAF only)
If MiscStringsAreSame(app.EXEName, "CalcZAF") Then
sample(1).SecondaryFluorescenceBoundaryFlag(ip%) = True
End If

' Debugmode (print column labels to log window)
If DebugMode And VerboseMode Then
Call IOWriteLog(vbCrLf)
Call IOWriteLog(k_string1$)
Call IOWriteLog(k_string2$)
Call IOWriteLog(k_string3$)
End If

' Input raw k-ratio and intensity data
k_npts& = 0
Do Until EOF(Temp2FileNumber%)
Line Input #Temp2FileNumber%, astring$      ' data string
k_npts& = k_npts& + 1

' Dimension variables
ReDim Preserve k_eV(1 To k_npts&) As Double
ReDim Preserve k_dist(1 To k_npts&) As Double
ReDim Preserve k_total(1 To k_npts&) As Double
ReDim Preserve k_fluor(1 To k_npts&) As Double
ReDim Preserve k_flach(1 To k_npts&) As Double
ReDim Preserve k_flabr(1 To k_npts&) As Double
ReDim Preserve k_flbch(1 To k_npts&) As Double
ReDim Preserve k_flbbr(1 To k_npts&) As Double
ReDim Preserve k_pri_int(1 To k_npts&) As Double
ReDim Preserve k_std_int(1 To k_npts&) As Double

' Parse eV for storage
Call MiscParseStringToStringA(astring$, VbSpace, bstring$)
If ierror Then Exit Sub
k_eV#(k_npts&) = CDbl(bstring$)   ' use double for language issues

' Parse xdist for storage
Call MiscParseStringToStringA(astring$, VbSpace, bstring$)
If ierror Then Exit Sub
k_dist#(k_npts&) = CDbl(bstring$)   ' use double for language issues

' Parse yktotal (total intensity %) and ykfluor (total fluorescence only %)
Call MiscParseStringToStringA(astring$, VbSpace, bstring$)
If ierror Then Exit Sub
k_total#(k_npts&) = CDbl(bstring$)   ' use double for language issues
Call MiscParseStringToStringA(astring$, VbSpace, bstring$)
If ierror Then Exit Sub
k_fluor#(k_npts&) = CDbl(bstring$)   ' use double for language issues

' Parse Mat A intensities
Call MiscParseStringToStringA(astring$, VbSpace, bstring$)
If ierror Then Exit Sub
k_flach#(k_npts&) = CDbl(bstring$)   ' use double for language issues

Call MiscParseStringToStringA(astring$, VbSpace, bstring$)
If ierror Then Exit Sub
k_flabr#(k_npts&) = CDbl(bstring$)   ' use double for language issues

' Parse Mat B intensities
Call MiscParseStringToStringA(astring$, VbSpace, bstring$)
If ierror Then Exit Sub
k_flbch#(k_npts&) = CDbl(bstring$)   ' use double for language issues

Call MiscParseStringToStringA(astring$, VbSpace, bstring$)
If ierror Then Exit Sub
k_flbbr#(k_npts&) = CDbl(bstring$)   ' use double for language issues

' Parse primary intensity
Call MiscParseStringToStringA(astring$, VbSpace, bstring$)
If ierror Then Exit Sub
k_pri_int#(k_npts&) = CDbl(bstring$)   ' use double for language issues

' Parse std intensity
Call MiscParseStringToStringA(astring$, vbTab, bstring$)
If ierror Then Exit Sub
k_std_int#(k_npts&) = CDbl(bstring$)   ' use double for language issues

' Calculate intensity percents for Mat A and Mat B
If k_std_int#(k_npts&) = 0 Then GoTo SecondaryReadKratiosDATFileBadStdInt

' Debugmode
If DebugMode And VerboseMode Then
astring$ = Space$(2) & Format$(k_eV#(k_npts&), e115$)
astring$ = astring$ & Space$(2) & Format$(k_dist#(k_npts&), e115$)

astring$ = astring$ & Space$(2) & Format$(k_total#(k_npts&), e115$)
astring$ = astring$ & Space$(2) & Format$(k_fluor#(k_npts&), e115$)

astring$ = astring$ & Space$(2) & Format$(k_flach#(k_npts&), e115$)
astring$ = astring$ & Space$(2) & Format$(k_flabr#(k_npts&), e115$)
astring$ = astring$ & Space$(2) & Format$(k_flbch#(k_npts&), e115$)
astring$ = astring$ & Space$(2) & Format$(k_flbbr#(k_npts&), e115$)

astring$ = astring$ & Space$(2) & Format$(k_pri_int#(k_npts&), e115$)
astring$ = astring$ & Space$(2) & Format$(k_std_int#(k_npts&), e115$)

Call IOWriteLog(astring$)
End If
Loop

Close #Temp2FileNumber%

' Check for kratio data
If k_npts& <= 0 Then GoTo SecondaryReadKratiosDATFileNoPoints

Exit Sub

' Errors
SecondaryReadKratiosDATFileError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryReadKratiosDATFile"
Close #Temp2FileNumber%
ierror = True
Exit Sub

SecondaryReadKratiosDATFileNoFilename:
msg$ = "No k-ratio data file was specified. Please browse for an appropriate k-ratio data file for secondary boundary fluorescence corrections."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryReadKratiosDATFile"
ierror = True
Exit Sub

SecondaryReadKratiosDATFileNoFileFound:
msg$ = "The k-ratio data file," & tfilename$ & " was not found. Please browse for an appropriate k-ratio data file for secondary boundary fluorescence corrections."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryReadKratiosDATFile"
ierror = True
Exit Sub

SecondaryReadKratiosDATFileBadFirstLine:
msg$ = "The first line of the specified k-ratio data file," & tfilename$ & " was not expected (" & k_string1$ & "). Please browse for an appropriate k-ratio data file for secondary boundary fluorescence corrections."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryReadKratiosDATFile"
ierror = True
Exit Sub

SecondaryReadKratiosDATFileBadSecondLine:
msg$ = "The second line of the specified k-ratio data file," & tfilename$ & " was not expected (" & k_string2$ & "). Please browse for an appropriate k-ratio data file for secondary boundary fluorescence corrections."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryReadKratiosDATFile"
ierror = True
Exit Sub

SecondaryReadKratiosDATFileElementNotFound:
msg$ = "The element symbol in the specified k-ratio data file (" & esym$ & ") does not match any sample element symbols) (please select another k-ratio data file and try again)"
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryReadKratiosDATFile"
Close #Temp2FileNumber%
ierror = True
Exit Sub

SecondaryReadKratiosDATFileXrayIsDifferent:
msg$ = "The x-ray transition in the specified k-ratio data file (" & trans$ & ") does not match the sample x-ray symbol for the matched element (" & sample(1).Elsyms$(ip%) & " " & sample(1).Xrsyms$(ip%) & ") (please select another k-ratio data file and try again)"
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryReadKratiosDATFile"
Close #Temp2FileNumber%
ierror = True
Exit Sub

SecondaryReadKratiosDATFileDifferentKilovolts:
msg$ = "The kilovolt parameter in the specified k-ratio data file does not match the sample kilovolts (please select another k-ratio data file and try again)"
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryReadKratiosDATFile"
Close #Temp2FileNumber%
ierror = True
Exit Sub

SecondaryReadKratiosDATFileBadStdInt:
msg$ = "The standard intensity for the BStd material in the specified k-ratio data file is zero, so something went wrong with the Fanal calculation (try typing Fanal < Fanal.in from the Fanal prompt and see what happens)"
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryReadKratiosDATFile"
Close #Temp2FileNumber%
ierror = True
Exit Sub

SecondaryReadKratiosDATFileNoPoints:
msg$ = "The number of k-ratio intensities in the specified k-ratio data file are zero, so something went wrong with the Fanal calculation (try typing Fanal < Fanal.in from the Fanal prompt and see what happens)"
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryReadKratiosDATFile"
Close #Temp2FileNumber%
ierror = True
Exit Sub

End Sub

Sub SecondaryCalculateDistance(X_Pos1 As Single, Y_Pos1 As Single, X_Pos2 As Single, Y_Pos2 As Single, sampleline As Integer, sample() As TypeSample)
' Calculate the boundary distance (in um) for the specified data line based on passed line and X, Y stage positions

ierror = False
On Error GoTo SecondaryCalculateDistanceError

Dim linear_dist As Single
Dim kmax As Integer, nmax As Integer
Dim X1 As Single, Y1 As Single  ' stage coordinate of data point
Dim m As Single, b As Single    ' m is slope and b is intercept of boundary

ReDim xdata(1 To 2) As Single
ReDim ydata(1 To 2) As Single
ReDim acoeff(1 To MAXCOEFF%) As Single

' Load the stage coordinate of the point
X1! = sample(1).StagePositions!(sampleline%, 1)
Y1! = sample(1).StagePositions!(sampleline%, 2)

If VerboseMode Then
Call IOWriteLog(vbCrLf & vbCrLf & "SecondaryCalculateDistance: Stage Coordinates, X= " & MiscAutoFormat$(X1!) & ", Y= " & MiscAutoFormat$(Y1!))
End If

' Check if boundary is vertical (X_Pos1 = X_Pos2) or horizonal (Y_Pos1 = Y_Pos2)
If X_Pos1! = X_Pos2! Or Y_Pos1! = Y_Pos2! Then

' Check for pathological conditions (point is on boundary line)
If X_Pos1! = X_Pos2! And X_Pos1! = X1! Then GoTo SecondaryCalculateDistanceX_PosOnBoundary
If Y_Pos1! = Y_Pos2! And Y_Pos1! = Y1! Then GoTo SecondaryCalculateDistanceY_PosOnBoundary

' Vertical boundary
If X_Pos1! = X_Pos2! Then
linear_dist! = Abs(X1! - X_Pos1!)
End If

' Horizontal boundary
If Y_Pos1! = Y_Pos2! Then
linear_dist! = Abs(Y1! - Y_Pos1!)
End If

' Check for zero distance
If linear_dist! = 0# Then GoTo SecondaryCalculateDistanceZeroDistance

' Convert to um
linear_dist! = linear_dist! * MotUnitsToAngstromMicrons!(XMotor%)

' Store micron linear distance in sample arrays
sample(1).SecondaryFluorescenceBoundaryDistance!(sampleline%) = linear_dist!
Exit Sub
End If

' First get equation for line using two points
kmax% = 1   ' linear fit
nmax% = 2   ' 2 points

xdata!(1) = X_Pos1!
ydata!(1) = Y_Pos1!
xdata!(2) = X_Pos2!
ydata!(2) = Y_Pos2!

Call LeastSquares(kmax%, nmax%, xdata!(), ydata!(), acoeff!())
If ierror Then Exit Sub

' Save slope and intercept
m! = acoeff!(2)
b! = acoeff!(1)

' Calculate from specified boundary (specified using several methods but stored as X,Y pair)
' Given that y = mx + b and point (x1, y1)... see http://math.ucsd.edu/~wgarner/math4c/derivations/distance/pdf/distptline.pdf
linear_dist! = Abs(Y1! - m! * X1! - b!) / Sqr(m! ^ 2 + 1)

' Check for zero distance
If linear_dist! = 0# Then GoTo SecondaryCalculateDistanceZeroDistance

' Convert to um
linear_dist! = linear_dist! * MotUnitsToAngstromMicrons!(XMotor%)

' Store in sample arrays
sample(1).SecondaryFluorescenceBoundaryDistance!(sampleline%) = Abs(linear_dist!)

Exit Sub

' Errors
SecondaryCalculateDistanceError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryCalculateDistance"
ierror = True
Exit Sub

SecondaryCalculateDistanceX_PosOnBoundary:
msg$ = "The passed stage coordinate (" & Format$(X1!) & "," & Format$(Y1!) & ") is on the specified boundary line. Distance will be zero. Try using a different specified boundary."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryCalculateDistance"
ierror = True
Exit Sub

SecondaryCalculateDistanceY_PosOnBoundary:
msg$ = "The passed stage coordinate (" & Format$(X1!) & "," & Format$(Y1!) & ") is on the specified boundary line. Distance will be zero. Try using a different specified boundary."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryCalculateDistance"
ierror = True
Exit Sub

SecondaryCalculateDistanceZeroDistance:
msg$ = "The calculated distance from the stage coordinate to the specified boundary is zero. Try using a different specified boundary."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryCalculateDistance"
ierror = True
Exit Sub

End Sub

Sub SecondaryCalculateKratio(sampleline As Integer, chan As Integer, sample() As TypeSample)
' Calculate the PAR file secondary fluorescence k-ratio for the boundary phase for the specified absolute linear distance or mass distance

ierror = False
On Error GoTo SecondaryCalculateKratioError

Dim krat As Single
Dim dist As Single

Dim n As Long
Dim npts As Integer
Dim kmax As Integer, nmax As Integer
Dim dmin As Double
Dim dpnt As Long

ReDim xdata(1 To 1) As Single
ReDim ydata(1 To 1) As Single

ReDim acoeff(1 To MAXCOEFF%) As Single

' Interpolate k-ratio values using 1, 2, 3 point fit for the current distance
dist! = sample(1).SecondaryFluorescenceBoundaryDistance!(sampleline%)
If dist! = 0# Then GoTo SecondaryCalculateKratioZeroDistance

' First find the closet calculated k-ratio
dmin# = MAXMINIMUM!
For n& = 1 To nPoints&(chan%)
If Abs(xdist#(chan%, n&) - dist!) < dmin# Then
dmin# = Abs(xdist#(chan%, n&) - dist!)
dpnt& = n&
End If
Next n&

' Now load the xdata array using linear distance
npts% = 1
xdata!(1) = xdist#(chan%, dpnt&)

' Now load the ydata kratio array using the mat B characteristic and continuum fluorescence only
ydata!(1) = fluB_k#(chan%, dpnt&)

' Check for pathological conditions
If dist! > xdist#(chan%, 1) And dist! < xdist#(chan%, nPoints&(chan%)) Then

' Now load points on each side if available
If dpnt& - 1 > 0 Then
npts% = npts% + 1
ReDim Preserve xdata(1 To npts%) As Single
ReDim Preserve ydata(1 To npts%) As Single
xdata!(npts%) = xdist#(chan%, dpnt& - 1)          ' distance in um
ydata!(npts%) = fluB_k#(chan%, dpnt& - 1)         ' characteristic and continuum fluorescence from mat B (boundary phase)
End If

If dpnt& + 1 <= nPoints&(chan%) Then
npts% = npts% + 1
ReDim Preserve xdata(1 To npts%) As Single
ReDim Preserve ydata(1 To npts%) As Single
xdata!(npts%) = xdist#(chan%, dpnt& + 1)          ' distance in um
ydata!(npts%) = fluB_k#(chan%, dpnt& + 1)         ' characteristic and continuum fluorescence from mat B (boundary phase)
End If

' Debug mode
If DebugMode Then
Call IOWriteLog(vbNullString)
For n& = 1 To npts%
Call IOWriteLog("SecondaryCalculateKratio: Point " & Format$(n&) & ", X= " & Format$(xdata!(n&)) & ", Y= " & Format$(ydata!(n&)))
Next n&
End If

' Now fit the data depending on the number of points found
kmax% = 2
If npts% < 3 Then kmax% = 1   ' linear fit or parabolic fit
nmax% = npts%
Call LeastSquares(kmax%, nmax%, xdata!(), ydata!(), acoeff!())
If ierror Then Exit Sub

' Now interpolate to get the actual k-ratio % for the specified distance
krat! = acoeff!(1) + dist! * acoeff!(2) + dist! ^ 2 * acoeff!(3)

' Distance is outside k-ratio data range (just use end values)
Else
If dist! <= xdist#(chan%, 1) Then krat! = fluB_k#(chan%, 1)
If dist! >= xdist#(chan%, nPoints&(chan%)) Then krat! = fluB_k#(chan%, nPoints&(chan%))
End If

If DebugMode Then
Call IOWriteLog("SecondaryCalculateKratio: Interpolated K-ratio % is " & MiscAutoFormat$(krat!) & " at a distance " & Format$(dist!) & " um")
End If

' Store in sample arrays (as k-ratio %)
sample(1).SecondaryFluorescenceBoundaryKratios!(sampleline%, chan%) = krat!

Exit Sub

' Errors
SecondaryCalculateKratioError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryCalculateKratio"
ierror = True
Exit Sub

SecondaryCalculateKratioZeroDistance:
msg$ = "The calculated distance from the stage coordinate to the specified boundary is zero. Try using a different specified boundary."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryCalculateKratio"
ierror = True
Exit Sub

End Sub

Sub SecondaryGetCoordinates(n As Long, x() As Single, y() As Single, Z() As Single)
' Get the currently analyzed data point coordinates

ierror = False
On Error GoTo SecondaryGetCoordinatesError

Dim i As Long

' Check for valid points
If apoints& < 1 Then Exit Sub

' Dimension
ReDim x(1 To apoints&) As Single
ReDim y(1 To apoints&) As Single
ReDim Z(1 To apoints&) As Single

For i& = 1 To apoints&
x!(i&) = xcoord!(i&)
y!(i&) = ycoord!(i&)
Z!(i&) = zcoord!(i&)
Next i&
n& = apoints&

Exit Sub

' Errors
SecondaryGetCoordinatesError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryGetCoordinates"
ierror = True
Exit Sub

End Sub

Sub SecondaryReturnKratiosDAT(t_string1 As String, t_string2 As String, t_string3 As String, t_eV() As Double, t_dist() As Double, t_total() As Double, t_fluor() As Double, t_flach() As Double, t_flabr() As Double, t_flbch() As Double, t_flbbr() As Double, t_pri_int() As Double, t_std_int() As Double, t_npts As Long)
' Return the raw data variables read by SecondaryReadKratiosDATFile (for storage)

ierror = False
On Error GoTo SecondaryReturnKratiosDATError

Dim n As Long

' Check for values to return
If k_npts& <= 0 Then GoTo SecondaryReturnKratiosDATNoData

t_npts& = k_npts&
t_string1$ = k_string1$
t_string2$ = k_string2$
t_string3$ = k_string3$

ReDim t_eV(1 To t_npts&) As Double
ReDim t_dist(1 To t_npts&) As Double

ReDim t_total(1 To t_npts&) As Double
ReDim t_fluor(1 To t_npts&) As Double
ReDim t_flach(1 To t_npts&) As Double
ReDim t_flabr(1 To t_npts&) As Double
ReDim t_flbch(1 To t_npts&) As Double
ReDim t_flbbr(1 To t_npts&) As Double
ReDim t_pri_int(1 To t_npts&) As Double
ReDim t_std_int(1 To t_npts&) As Double

For n& = 1 To t_npts&
t_eV#(n&) = k_eV#(n&)
t_dist#(n&) = k_dist#(n&)

t_total#(n&) = k_total#(n&)
t_fluor#(n&) = k_fluor#(n&)
t_flach#(n&) = k_flach#(n&)
t_flabr#(n&) = k_flabr#(n&)
t_flbch#(n&) = k_flbch#(n&)
t_flbbr#(n&) = k_flbbr#(n&)
t_pri_int#(n&) = k_pri_int#(n&)
t_std_int#(n&) = k_std_int#(n&)
Next n&

Exit Sub

' Errors
SecondaryReturnKratiosDATError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryReturnKratiosDAT"
ierror = True
Exit Sub

SecondaryReturnKratiosDATNoData:
msg$ = "No k-ratio data points are available to be returned."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryReturnKratiosDAT"
ierror = True
Exit Sub

End Sub

Sub SecondaryRestoreKratiosDAT(t_string1 As String, t_string2 As String, t_string3 As String, t_eV() As Double, t_dist() As Double, t_total() As Double, t_fluor() As Double, t_flach() As Double, t_flabr() As Double, t_flbch() As Double, t_flbbr() As Double, t_pri_int() As Double, t_std_int() As Double, t_npts As Long)
' Restore the raw data variables to module level for writing with SecondaryWriteKratiosDATFile (for output)

ierror = False
On Error GoTo SecondaryRestoreKratiosDATError

Dim n As Long

k_npts& = t_npts&
k_string1$ = t_string1$
k_string2$ = t_string2$
k_string3$ = t_string3$

ReDim k_eV(1 To k_npts&) As Double
ReDim k_dist(1 To k_npts&) As Double

ReDim k_total(1 To k_npts&) As Double
ReDim k_fluor(1 To k_npts&) As Double
ReDim k_flach(1 To k_npts&) As Double
ReDim k_flabr(1 To k_npts&) As Double
ReDim k_flbch(1 To k_npts&) As Double
ReDim k_flbbr(1 To k_npts&) As Double
ReDim k_pri_int(1 To k_npts&) As Double
ReDim k_std_int(1 To k_npts&) As Double

For n& = 1 To k_npts&
k_eV#(n&) = t_eV#(n&)
k_dist#(n&) = t_dist#(n&)

k_total#(n&) = t_total#(n&)
k_fluor#(n&) = t_fluor#(n&)
k_flach#(n&) = t_flach#(n&)
k_flabr#(n&) = t_flabr#(n&)
k_flbch#(n&) = t_flbch#(n&)
k_flbbr#(n&) = t_flbbr#(n&)
k_pri_int#(n&) = t_pri_int#(n&)
k_std_int#(n&) = t_std_int#(n&)
Next n&

Exit Sub

' Errors
SecondaryRestoreKratiosDATError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryRestoreKratiosDAT"
ierror = True
Exit Sub

End Sub

Sub SecondaryProcessKratiosDAT(chan As Integer)
' Process the raw data variables read by SecondaryReadKratiosDATFile and save at module level for each channel

ierror = False
On Error GoTo SecondaryProcessKratiosDATError

Dim n As Long
Dim temp1 As Single, temp2 As Single

' Dimension variables for this channel in the kratio arrays (ip is the channel number)
If k_npts& > maxpoints& Then
ReDim Preserve xdist(1 To MAXCHAN%, 1 To k_npts&) As Double
ReDim Preserve yktotal(1 To MAXCHAN%, 1 To k_npts&) As Double
ReDim Preserve ykfluor(1 To MAXCHAN%, 1 To k_npts&) As Double

ReDim Preserve fluA_k(1 To MAXCHAN%, 1 To k_npts&) As Double
ReDim Preserve fluB_k(1 To MAXCHAN%, 1 To k_npts&) As Double
ReDim Preserve prix_k(1 To MAXCHAN%, 1 To k_npts&) As Double

maxpoints& = k_npts&
End If

' Save data to module level channel arrays
nPoints(chan%) = k_npts&      ' number of k-ratio data points for each channel

' Load kratios.DAT values into module level channel arrays for calculations (eV not used)
For n& = 1 To k_npts&
xdist#(chan%, n&) = k_dist#(n&)     ' linear distance (um) for each k-ratio

yktotal(chan%, n&) = k_total#(n&)   ' fluorescence kratio% plus primary x-rays kratio% from material A and material B
ykfluor(chan%, n&) = k_fluor#(n&)   ' fluorescence kratio% only (minus primary x-ray kratio% from material A)

fluA_k#(chan%, n&) = 100# * (k_flach#(n&) + k_flabr#(n&)) / k_std_int#(n&)       ' Mat A total fluorescence k-ratio %
fluB_k#(chan%, n&) = 100# * (k_flbch#(n&) + k_flbbr#(n&)) / k_std_int#(n&)       ' Mat B total fluorescence k-ratio %
prix_k#(chan%, n&) = k_pri_int#(n&) / k_std_int#(n&)                             ' Primary x-ray k-ratio % from Mat A
Next n&

' Check that calculated distance is enough (that total intensity at max distance is close to Mat A only intensity)
temp1! = 0#
temp2! = 0#
If yktotal#(chan%, nPoints&(chan%)) > 0.0001 Then
temp1! = yktotal#(chan%, k_npts&) - (prix_k#(chan%, k_npts&) + fluA_k#(chan%, k_npts&))   ' total intensity minus mat A only intensity at max distance
temp2! = 100# * temp1! / yktotal#(chan%, 1)                                               ' (percent) total intensity at closest distance

' Check if difference is greater than 1%
If DebugMode And VerboseMode Then
msg$ = vbCrLf & "The Fanal couple k-ratio data file maximum distance for channel " & Format$(chan%)
msg$ = msg$ & " is " & Format$(xdist#(chan%, k_npts&)) & " um and the boundary fluorescence k-ratio% intensity at this distance is " & Format$(temp1!)
msg$ = msg$ & " or " & Format$(temp2!) & " relative percent to the first k-ratio distance of " & Format$(xdist#(chan%, 1)) & " um,"
msg$ = msg$ & " with a k-ratio % intensity of " & Format$(yktotal#(chan%, 1)) & "."
Call IOWriteLog(msg$)
End If

If temp2! > 1# Then
msg$ = "It might be a good idea to calculate a longer total secondary fluorescence couple distance for channel " & Format$(chan%) & " using the Fanal GUI (in Standard.exe) for improved accuracy, especially for trace elements."
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
End If
End If

Exit Sub

' Errors
SecondaryProcessKratiosDATError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryProcessKratiosDAT"
ierror = True
Exit Sub

End Sub

Sub SecondaryWriteKratiosDATFile(tfilename As String)
' Write the module level kratios to the specified k-ratios.dat file for the boundary correction

ierror = False
On Error GoTo SecondaryWriteKratiosDATFileError

Dim n As Long
Dim astring As String

' Check for valid k-ratio file
If Trim$(tfilename$) = vbNullString Then GoTo SecondaryWriteKratiosDATFileNoFilename

' Check for kratio data
If k_npts& <= 0 Then GoTo SecondaryWriteKratiosDATFileNoPoints

Close #Temp2FileNumber%
DoEvents
Open tfilename$ For Output As #Temp2FileNumber%

' Debugmode (print column labels to log window)
If DebugMode And VerboseMode Then
Call IOWriteLog(vbCrLf)
Call IOWriteLog(k_string1$)
Call IOWriteLog(k_string2$)
Call IOWriteLog(k_string3$)
End If

' Write characteristic line
Print #Temp2FileNumber%, k_string1$

' Write electron energy line
Print #Temp2FileNumber%, k_string2$

' Write column labels line
Print #Temp2FileNumber%, k_string3$

' Output raw k-ratio and intensity data
For n& = 1 To k_npts&

astring$ = Space$(2) & Format$(k_eV#(n&), e115$)
astring$ = astring$ & Space$(2) & Format$(k_dist#(n&), e115$)

astring$ = astring$ & Space$(2) & Format$(k_total#(n&), e115$)
astring$ = astring$ & Space$(2) & Format$(k_fluor#(n&), e115$)

astring$ = astring$ & Space$(2) & Format$(k_flach#(n&), e115$)
astring$ = astring$ & Space$(2) & Format$(k_flabr#(n&), e115$)
astring$ = astring$ & Space$(2) & Format$(k_flbch#(n&), e115$)
astring$ = astring$ & Space$(2) & Format$(k_flbbr#(n&), e115$)

astring$ = astring$ & Space$(2) & Format$(k_pri_int#(n&), e115$)
astring$ = astring$ & Space$(2) & Format$(k_std_int#(n&), e115$)

' Debugmode
If DebugMode And VerboseMode Then
Call IOWriteLog(astring$)
End If

Print #Temp2FileNumber%, astring$      ' data string
Next n&

Close #Temp2FileNumber%
Exit Sub

' Errors
SecondaryWriteKratiosDATFileError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryWriteKratiosDATFile"
Close #Temp2FileNumber%
ierror = True
Exit Sub

SecondaryWriteKratiosDATFileNoFilename:
msg$ = "No k-ratio data file was specified. Please browse for an appropriate k-ratio data file for secondary boundary fluorescence corrections."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryWriteKratiosDATFile"
ierror = True
Exit Sub

SecondaryWriteKratiosDATFileNoPoints:
msg$ = "The number of k-ratio intensities in the specified k-ratio data file are zero, so something went wrong with the Fanal calculation (try typing Fanal < Fanal.in from the Fanal prompt and see what happens)"
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryWriteKratiosDATFile"
Close #Temp2FileNumber%
ierror = True
Exit Sub

End Sub

