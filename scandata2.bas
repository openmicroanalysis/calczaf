Attribute VB_Name = "CodeScanData2"
' (c) Copyright 1995-2024 by John J. Donovan
Option Explicit

Const MAXSEGMENTS% = 400

' Special arrays for spline fit data and coefficients
Dim nPoints As Integer
Dim xdata() As Single, ydata() As Single
Dim ycoeff() As Double

Sub ScanDataPlotFitCurve_PE(tGraph As Pesgo, mode As Integer, linecount As Long, acoeff1 As Single, acoeff2 As Single, acoeff3 As Single, centroid As Single, threshold As Single, currentonpeak As Single, tBold As Boolean)
' Display the ROM and parabolic peak fit centroid and threshold on the graph (Pro Essentials code)
' mode: 1 = parabolic, 2 = gaussian, 3 = maxima, 4 = maximum value, 5 = cubic spline, 6 = (multi-point) exponential

ierror = False
On Error GoTo ScanDataPlotFitCurve_PEError

Dim firstpointdone As Boolean
Dim i As Integer
Dim tX As Single, tY As Single
Dim xmin As Single, xmax As Single, ymin As Single, ymax As Single
Dim sxmin As Single, sxmax As Single, symin As Single, symax As Single
Dim tsymin As Double, tsymax As Double

tGraph.PEactions = REINITIALIZE_RESETIMAGE

' Determine min and max of graph
xmin! = tGraph.ManualMinX
xmax! = tGraph.ManualMaxX

ymin! = tGraph.ManualMinY
ymax! = tGraph.ManualMaxY

' Check for valid curve
If (acoeff1! <> 0# And acoeff2! <> 0# And acoeff3! <> 0#) Or mode% = 5 Then

' Calculate line to draw based on fit coefficients
sxmax! = xmin!
For i% = 1 To MAXSEGMENTS%

' Calculate partial line segments for x
sxmin! = sxmax!
sxmax! = sxmin! + (xmax! - xmin!) / (MAXSEGMENTS% - 1)

' Calculate partial line segments for y
If mode% = 1 Then               ' parabolic
tsymin# = acoeff1! + sxmin! * acoeff2! + sxmin! ^ 2 * acoeff3!
tsymax# = acoeff1! + sxmax! * acoeff2! + sxmax! ^ 2 * acoeff3!

ElseIf mode% = 2 Then           ' gaussian
tsymin# = acoeff1! + sxmin! * acoeff2! + sxmin! ^ 2 * acoeff3!
tsymax# = acoeff1! + sxmax! * acoeff2! + sxmax! ^ 2 * acoeff3!
If tsymin# > MAXLOGEXPD! Then tsymin# = MAXLOGEXPD!
If tsymax# > MAXLOGEXPD! Then tsymax# = MAXLOGEXPD!
tsymin# = NATURALE# ^ tsymin#
tsymax# = NATURALE# ^ tsymax#

ElseIf mode% = 5 Then           ' cubic spline
tX! = sxmin!
Call SplineInterpolate(xdata!(), ydata!(), ycoeff#(), CLng(nPoints%), tX!, tY!)
If ierror Then Exit Sub
tsymin# = tY!
tX! = sxmax!
Call SplineInterpolate(xdata!(), ydata!(), ycoeff#(), CLng(nPoints%), tX!, tY!)
If ierror Then Exit Sub
tsymax# = tY!

ElseIf mode% = 6 Then           ' multi-point exponential
tsymin# = acoeff1! + sxmin! * acoeff2! + sxmin! ^ 2 * acoeff3!
tsymax# = acoeff1! + sxmax! * acoeff2! + sxmax! ^ 2 * acoeff3!
If tsymin# > MAXLOGEXPD! Then tsymin# = MAXLOGEXPD!
If tsymax# > MAXLOGEXPD! Then tsymax# = MAXLOGEXPD!
tsymin# = NATURALE# ^ tsymin#
tsymax# = NATURALE# ^ tsymax#
End If

' Clip
If tsymin# < ymin! Then tsymin# = ymin!
If tsymax# > ymax! Then tsymax# = ymax!

If tsymin# > ymax! Then tsymin# = ymax!
If tsymax# < ymin! Then tsymax# = ymin!

' Load from temp doubles
symin! = tsymin#
symax! = tsymax#

' Plot fit
If Not firstpointdone Then
Call ScanDataPlotLine(tGraph, linecount&, sxmin!, symin!, sxmax!, symax!, False, tBold, Int(255), Int(0), Int(0), Int(255))     ' blue
If ierror Then Exit Sub
firstpointdone = True
Else
Call ScanDataPlotLine(tGraph, linecount&, sxmin!, symin!, sxmax!, symax!, True, tBold, Int(255), Int(0), Int(0), Int(255))      ' blue
If ierror Then Exit Sub
End If
Next i%
End If

' Plot centroid
If centroid! <> 0# Then
Call ScanDataPlotLine(tGraph, linecount&, centroid!, ymin!, centroid!, ymax!, False, True, Int(255), Int(255), Int(0), Int(0))              ' red
If ierror Then Exit Sub
End If

' Plot threshold
If threshold! <> 0# Then
Call ScanDataPlotLine(tGraph, linecount&, xmin!, threshold!, xmax!, threshold!, False, True, Int(255), Int(0), Int(255), Int(255))          ' cyan
If ierror Then Exit Sub
End If

' Plot current on-peak
If currentonpeak! <> 0# Then
Call ScanDataPlotLine(tGraph, linecount&, currentonpeak!, ymin!, currentonpeak!, ymax!, False, True, Int(255), Int(0), Int(255), Int(0))    ' green
If ierror Then Exit Sub
End If

Exit Sub

' Errors
ScanDataPlotFitCurve_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "ScanDataPlotFitCurve_PE"
ierror = True
Exit Sub

End Sub

Sub ScanDataPlotFitLoad(npts%, txdata() As Single, tydata() As Single, tycoeff() As Double)
' Special call to load data points and y coefficients for spline fit

ierror = False
On Error GoTo ScanDataPlotFitLoadError

Dim i As Integer

' Load to module level variables
nPoints% = npts%

ReDim xdata(1 To nPoints%) As Single
ReDim ydata(1 To nPoints%) As Single
ReDim ycoeff(1 To nPoints%) As Double

For i% = 1 To nPoints%
xdata!(i%) = txdata!(i%)
ydata!(i%) = tydata!(i%)
ycoeff#(i%) = tycoeff#(i%)
Next i%

Exit Sub

' Errors
ScanDataPlotFitLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "ScanDataPlotFitLoad"
ierror = True
Exit Sub

End Sub

Sub ScanDataPlotOnPeak_PE(tGraph As Pesgo, linecount As Long, onpeak As Single)
' Display the selected on-peak position (Pro Essentials code)

ierror = False
On Error GoTo ScanDataPlotOnPeak_PEError

Dim xmin As Single, xmax As Single, ymin As Single, ymax As Single

' Determine min and max of graph
xmin! = tGraph.ManualMinX
xmax! = tGraph.ManualMaxX

ymin! = tGraph.ManualMinY
ymax! = tGraph.ManualMaxY

' Plot on-peak position
If onpeak! <> 0# Then
Call ScanDataPlotLine(tGraph, linecount&, onpeak!, ymin!, onpeak!, ymax!, False, True, Int(255), Int(0), Int(255), Int(0))   ' light green
If ierror Then Exit Sub
End If

Exit Sub

' Errors
ScanDataPlotOnPeak_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "ScanDataPlotOnPeak_PE"
ierror = True
Exit Sub

End Sub

Sub ScanDataPlotLine(tGraph As Pesgo, linecount As Long, txmin As Single, tymin As Single, txmax As Single, tymax As Single, tContinue As Boolean, tBold As Boolean, tAlpha As Integer, tRed As Integer, tGreen As Integer, tBlue As Integer)
' Plots a line on the passed graph using the passed color

ierror = False
On Error GoTo ScanDataPlotLineError

tGraph.ShowAnnotations = True

' Start point
tGraph.GraphAnnotationX(linecount&) = txmin!
tGraph.GraphAnnotationY(linecount&) = tymin!
If Not tContinue Then
If Not tBold Then
tGraph.GraphAnnotationType(linecount&) = PEGAT_THIN_SOLIDLINE&
Else
tGraph.GraphAnnotationType(linecount&) = PEGAT_MEDIUM_SOLIDLINE&
End If
Else
tGraph.GraphAnnotationType(linecount&) = PEGAT_LINECONTINUE&
End If
tGraph.GraphAnnotationColor(linecount&) = tGraph.PEargb(tAlpha%, tRed%, tGreen%, tBlue%)

' No text (at this time)
tGraph.GraphAnnotationText(linecount&) = ""         ' vbNullString does not work here!!!

' End point
linecount& = linecount& + 1
tGraph.GraphAnnotationX(linecount&) = txmax!
tGraph.GraphAnnotationY(linecount&) = tymax!
tGraph.GraphAnnotationType(linecount&) = PEGAT_LINECONTINUE&
tGraph.GraphAnnotationColor(linecount&) = tGraph.PEargb(tAlpha%, tRed%, tGreen%, tBlue%)

' No text (at this time)
tGraph.GraphAnnotationText(linecount&) = ""         ' vbNullString does not work here!!!
linecount& = linecount& + 1

Exit Sub

' Errors
ScanDataPlotLineError:
MsgBox Error$, vbOKOnly + vbCritical, "ScanDataPlotLine"
ierror = True
Exit Sub

End Sub

Sub ScanDataPlotLineRGB(tGraph As Pesgo, linecount As Long, txmin As Single, tymin As Single, txmax As Single, tymax As Single, tContinue As Boolean, tBold As Boolean, tRGB As Long)
' Plots a line on the passed graph using the passed RGB color

ierror = False
On Error GoTo ScanDataPlotLineRGBError

tGraph.ShowAnnotations = True

' Start point
tGraph.GraphAnnotationX(linecount&) = txmin!
tGraph.GraphAnnotationY(linecount&) = tymin!
If Not tContinue Then
If Not tBold Then
tGraph.GraphAnnotationType(linecount&) = PEGAT_THIN_SOLIDLINE&
Else
tGraph.GraphAnnotationType(linecount&) = PEGAT_MEDIUM_SOLIDLINE&
End If
Else
tGraph.GraphAnnotationType(linecount&) = PEGAT_LINECONTINUE&
End If
tGraph.GraphAnnotationColor(linecount&) = tRGB&

' No text (at this time)
tGraph.GraphAnnotationText(linecount&) = ""                  ' vbNullString does not work here!!!

' End point
linecount& = linecount& + 1
tGraph.GraphAnnotationX(linecount&) = txmax!
tGraph.GraphAnnotationY(linecount&) = tymax!
tGraph.GraphAnnotationType(linecount&) = PEGAT_LINECONTINUE&
tGraph.GraphAnnotationColor(linecount&) = tRGB&

' No text (at this time)
tGraph.GraphAnnotationText(linecount&) = ""                  ' vbNullString does not work here!!!
linecount& = linecount& + 1

Exit Sub

' Errors
ScanDataPlotLineRGBError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "ScanDataPlotLineRGB"
ierror = True
Exit Sub

End Sub

Sub ScanDataPlotBar(tGraph As Pesgo, linecount As Long, txmin As Double, tymin As Double, txmax As Double, tymax As Double, tText As String, tAlpha As Integer, tRed As Integer, tGreen As Integer, tBlue As Integer)
' Plots a (scale) bar (and text striong) on the passed graph using the passed color

ierror = False
On Error GoTo ScanDataPlotBarError

tGraph.ShowAnnotations = True

' Start point
tGraph.GraphAnnotationX(linecount&) = txmin#
tGraph.GraphAnnotationY(linecount&) = tymin#
tGraph.GraphAnnotationType(linecount&) = PEGAT_THICK_SOLIDLINE&
tGraph.GraphAnnotationColor(linecount&) = tGraph.PEargb(tAlpha%, tRed%, tGreen%, tBlue%)

' End point
linecount& = linecount& + 1
tGraph.GraphAnnotationX(linecount&) = txmax#
tGraph.GraphAnnotationY(linecount&) = tymax#
tGraph.GraphAnnotationType(linecount&) = PEGAT_LINECONTINUE&
tGraph.GraphAnnotationColor(linecount&) = tGraph.PEargb(tAlpha%, tRed%, tGreen%, tBlue%)

' Load text if passed
If tText$ <> vbNullString Then tGraph.GraphAnnotationText(linecount&) = tText$

linecount& = linecount& + 1
Exit Sub

' Errors
ScanDataPlotBarError:
MsgBox Error$, vbOKOnly + vbCritical, "ScanDataPlotBar"
ierror = True
Exit Sub

End Sub

Sub ScanDataPlotPHAWindow_PE(tGraph As Pesgo, linecount As Long, baseline As Single, window As Single, intediff As Integer, tBold As Boolean)
' Display the PHA window on the graph (Pro Essentials code)

ierror = False
On Error GoTo ScanDataPlotPHAWindow_PEError

Dim xmin As Single, xmax As Single, ymin As Single, ymax As Single
Dim sxmin As Single, sxmax As Single, symin As Single, symax As Single

tGraph.PEactions = REINITIALIZE_RESETIMAGE                   ' generate new plot

' Determine min and max of graph
xmin! = tGraph.ManualMinX
xmax! = tGraph.ManualMaxX

ymin! = tGraph.ManualMinY
ymax! = tGraph.ManualMaxY

' Calculate PHA window
sxmin! = baseline!
sxmax! = baseline! + window!
symin! = ymin! + (ymax! - ymin!) / 2#
symax! = symin!

' Plot PHA window
Call ScanDataPlotLine(tGraph, linecount&, sxmin!, symin!, sxmax!, symax!, False, tBold, Int(255), Int(255), Int(0), Int(0))     ' red
If ierror Then Exit Sub

' Calculate end bars
Call ScanDataPlotLine(tGraph, linecount&, sxmin!, symin! - (ymax! - ymin!) * 0.1, sxmin!, symax! + (ymax! - ymin!) * 0.1, False, tBold, Int(255), Int(255), Int(0), Int(0))     ' red
If ierror Then Exit Sub

' Plot end bar if differential
If intediff% <> 0 Then
Call ScanDataPlotLine(tGraph, linecount&, sxmax!, symin! - (ymax! - ymin!) * 0.1, sxmax!, symax! + (ymax! - ymin!) * 0.1, False, tBold, Int(255), Int(255), Int(0), Int(0))     ' red
If ierror Then Exit Sub
End If

Exit Sub

' Errors
ScanDataPlotPHAWindow_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "ScanDataPlotPHAWindow_PE"
ierror = True
Exit Sub

End Sub

Sub ScanDataPlotPHABiasGain_PE(mode As Integer, tGraph As Pesgo, linecount As Long, bias As Single, gain As Single, tScanCurrentBiasGainOnPeak As Single, tBold As Boolean)
' Display the PHA window on the graph (Pro Essentials code)

ierror = False
On Error GoTo ScanDataPlotPHABiasGain_PEError

Dim xmin As Single, xmax As Single, ymin As Single, ymax As Single
Dim sxmin As Single, sxmax As Single, symin As Single, symax As Single

tGraph.PEactions = REINITIALIZE_RESETIMAGE                   ' generate new plot

' Determine min and max of graph
xmin! = tGraph.ManualMinX
xmax! = tGraph.ManualMaxX

ymin! = tGraph.ManualMinY
ymax! = tGraph.ManualMaxY

' Calculate scan data bias position (original)
If mode% = 2 Then
sxmin! = bias!
sxmax! = bias
End If

' Calculate scan data gain position (original)
If mode% = 3 Then
sxmin! = gain!
sxmax! = gain
End If

symin! = ymin!
symax! = ymax!

' Plot original PHA bias
If mode% = 2 Then
Call ScanDataPlotLine(tGraph, linecount&, sxmin!, symin!, sxmax!, symax!, False, tBold, Int(255), Int(255), Int(0), Int(0))     ' red
If ierror Then Exit Sub
End If

' Plot original PHA gain
If mode% = 3 Then
Call ScanDataPlotLine(tGraph, linecount&, sxmin!, symin!, sxmax!, symax!, False, tBold, Int(255), Int(255), Int(0), Int(255))     ' red
If ierror Then Exit Sub
End If

' Calculate scan data bias position (new)
If mode% = 2 Then
sxmin! = tScanCurrentBiasGainOnPeak!
sxmax! = tScanCurrentBiasGainOnPeak!
End If

' Calculate scan data gain position (new)
If mode% = 3 Then
sxmin! = tScanCurrentBiasGainOnPeak!
sxmax! = tScanCurrentBiasGainOnPeak!
End If

symin! = ymin!
symax! = ymax!

' Plot new PHA bias
If mode% = 2 Then
Call ScanDataPlotLine(tGraph, linecount&, sxmin!, symin!, sxmax!, symax!, False, tBold, Int(255), Int(0), Int(255), Int(0))     ' green
If ierror Then Exit Sub
End If

' Plot new PHA gain
If mode% = 3 Then
Call ScanDataPlotLine(tGraph, linecount&, sxmin!, symin!, sxmax!, symax!, False, tBold, Int(255), Int(0), Int(255), Int(0))     ' green
If ierror Then Exit Sub
End If

Exit Sub

' Errors
ScanDataPlotPHABiasGain_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "ScanDataPlotPHABiasGain_PE"
ierror = True
Exit Sub

End Sub


