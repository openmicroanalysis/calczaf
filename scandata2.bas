Attribute VB_Name = "CodeScanData2"
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

Const MAXSEGMENTS% = 400

' Special arrays for spline fit data and coefficients
Dim nPoints As Integer
Dim xdata() As Single, ydata() As Single
Dim ycoeff() As Double

Sub ScanDataPlotFitCurve_PE(tGraph As Pesgo, mode As Integer, linecount As Long, acoeff1 As Single, acoeff2 As Single, acoeff3 As Single, centroid As Single, threshold As Single, currentonpeak As Single)
' Display the ROM and parabolic peak fit centroid and threshold on the graph (Pro Essentials code)
' mode: 1 = parabolic, 2 = gaussian, 3 = maxima, 4 = maximum value, 5 = cubic spline, 6 = (multi-point) exponential

ierror = False
On Error GoTo ScanDataPlotFitCurve_PEError

Dim firstpointdone As Boolean
Dim i As Integer
Dim tx As Single, ty As Single
Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double
Dim sxmin As Double, sxmax As Double, symin As Double, symax As Double

' Determine min and max of graph
xmin# = tGraph.ManualMinX
xmax# = tGraph.ManualMaxX

ymin# = tGraph.ManualMinY
ymax# = tGraph.ManualMaxY

' Check for valid curve
If (acoeff1! <> 0# And acoeff2! <> 0# And acoeff3! <> 0#) Or mode% = 5 Then

' Calculate line to draw based on fit coefficients
sxmax# = xmin#
For i% = 1 To MAXSEGMENTS%

' Calculate partial line segments for x
sxmin# = sxmax#
sxmax# = sxmin# + (xmax# - xmin#) / (MAXSEGMENTS% - 1)

' Calculate partial line segments for y
If mode% = 1 Then               ' parabolic
symin# = acoeff1! + sxmin# * acoeff2! + sxmin# ^ 2 * acoeff3!
symax# = acoeff1! + sxmax# * acoeff2! + sxmax# ^ 2 * acoeff3!

ElseIf mode% = 2 Then           ' gaussian
symin# = acoeff1! + sxmin# * acoeff2! + sxmin# ^ 2 * acoeff3!
symax# = acoeff1! + sxmax# * acoeff2! + sxmax# ^ 2 * acoeff3!
If symin# > MAXLOGEXPD! Then symin# = MAXLOGEXPD!
If symax# > MAXLOGEXPD! Then symax# = MAXLOGEXPD!
symin# = NATURALE# ^ symin#
symax# = NATURALE# ^ symax#

ElseIf mode% = 5 Then           ' cubic spline
tx! = CSng(sxmin#)
Call SplineInterpolate(xdata!(), ydata!(), ycoeff#(), CLng(nPoints%), tx!, ty!)
If ierror Then Exit Sub
symin# = CDbl(ty!)
tx! = CSng(sxmax#)
Call SplineInterpolate(xdata!(), ydata!(), ycoeff#(), CLng(nPoints%), tx!, ty!)
If ierror Then Exit Sub
symax# = CDbl(ty!)

ElseIf mode% = 6 Then           ' multi-point exponential
symin# = acoeff1! + sxmin# * acoeff2! + sxmin# ^ 2 * acoeff3!
symax# = acoeff1! + sxmax# * acoeff2! + sxmax# ^ 2 * acoeff3!
If symin# > MAXLOGEXPD! Then symin# = MAXLOGEXPD!
If symax# > MAXLOGEXPD! Then symax# = MAXLOGEXPD!
symin# = NATURALE# ^ symin#
symax# = NATURALE# ^ symax#
End If

If Not firstpointdone Then
Call ScanDataPlotLine(tGraph, linecount&, sxmin#, symin#, sxmax#, symax#, False, False, Int(255), Int(128), Int(0), Int(0))     ' brown
If ierror Then Exit Sub
firstpointdone = True
Else
Call ScanDataPlotLine(tGraph, linecount&, sxmin#, symin#, sxmax#, symax#, True, False, Int(255), Int(128), Int(0), Int(0))      ' brown
If ierror Then Exit Sub
End If
Next i%
End If

' Plot centroid
If centroid! <> 0# Then
Call ScanDataPlotLine(tGraph, linecount&, CDbl(centroid!), ymin#, CDbl(centroid!), ymax#, False, True, Int(255), Int(255), Int(0), Int(0))              ' red
If ierror Then Exit Sub
End If

' Plot threshold
If threshold! <> 0# Then
Call ScanDataPlotLine(tGraph, linecount&, xmin#, CDbl(threshold!), xmax#, CDbl(threshold!), False, True, Int(255), Int(0), Int(255), Int(255))          ' cyan
If ierror Then Exit Sub
End If

' Plot current on-peak
If currentonpeak! <> 0# Then
Call ScanDataPlotLine(tGraph, linecount&, CDbl(currentonpeak!), ymin#, CDbl(currentonpeak!), ymax#, False, True, Int(255), Int(0), Int(255), Int(0))    ' green
If ierror Then Exit Sub
End If

tGraph.PEactions = REINITIALIZE_RESETIMAGE

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

Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double
Dim txmin As Double, txmax As Double, tymin As Double, tymax As Double

' Determine min and max of graph
xmin# = tGraph.ManualMinX
xmax# = tGraph.ManualMaxX

ymin# = tGraph.ManualMinY
ymax# = tGraph.ManualMaxY

' Plot on-peak position
If onpeak! <> 0# Then
Call ScanDataPlotLine(tGraph, linecount&, CDbl(onpeak!), tymin#, CDbl(onpeak!), tymax#, False, True, Int(255), Int(0), Int(255), Int(0))   ' light green
If ierror Then Exit Sub
End If

Exit Sub

' Errors
ScanDataPlotOnPeak_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "ScanDataPlotOnPeak_PE"
ierror = True
Exit Sub

End Sub

Sub ScanDataPlotLine(tGraph As Pesgo, linecount As Long, txmin As Double, tymin As Double, txmax As Double, tymax As Double, tContinue As Boolean, tBold As Boolean, tAlpha As Integer, tRed As Integer, tGreen As Integer, tBlue As Integer)
' Plots a line on the passed graph using the passed color

ierror = False
On Error GoTo ScanDataPlotLineError

tGraph.ShowAnnotations = True

' Start point
tGraph.GraphAnnotationX(linecount&) = txmin#
tGraph.GraphAnnotationY(linecount&) = tymin#
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
tGraph.GraphAnnotationText(linecount&) = ""
'tGraph.GraphAnnotationText(linecount&) = vbNullString      ' this does not work!!!

' End point
linecount& = linecount& + 1
tGraph.GraphAnnotationX(linecount&) = txmax#
tGraph.GraphAnnotationY(linecount&) = tymax#
tGraph.GraphAnnotationType(linecount&) = PEGAT_LINECONTINUE&
tGraph.GraphAnnotationColor(linecount&) = tGraph.PEargb(tAlpha%, tRed%, tGreen%, tBlue%)

' No text (at this time)
tGraph.GraphAnnotationText(linecount&) = ""
'tGraph.GraphAnnotationText(linecount&) = vbNullString      ' this does not work!!!
linecount& = linecount& + 1

Exit Sub

' Errors
ScanDataPlotLineError:
MsgBox Error$, vbOKOnly + vbCritical, "ScanDataPlotLine"
ierror = True
Exit Sub

End Sub

Sub ScanDataPlotLineRGB(tGraph As Pesgo, linecount As Long, txmin As Double, tymin As Double, txmax As Double, tymax As Double, tContinue As Boolean, tBold As Boolean, tRGB As Long)
' Plots a line on the passed graph using the passed RGB color

ierror = False
On Error GoTo ScanDataPlotLineRGBError

tGraph.ShowAnnotations = True

' Start point
tGraph.GraphAnnotationX(linecount&) = txmin#
tGraph.GraphAnnotationY(linecount&) = tymin#
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
tGraph.GraphAnnotationText(linecount&) = ""
'tGraph.GraphAnnotationText(linecount&) = vbNullString      ' this does not work!!!

' End point
linecount& = linecount& + 1
tGraph.GraphAnnotationX(linecount&) = txmax#
tGraph.GraphAnnotationY(linecount&) = tymax#
tGraph.GraphAnnotationType(linecount&) = PEGAT_LINECONTINUE&
tGraph.GraphAnnotationColor(linecount&) = tRGB&

' No text (at this time)
tGraph.GraphAnnotationText(linecount&) = ""
'tGraph.GraphAnnotationText(linecount&) = vbNullString      ' this does not work!!!
linecount& = linecount& + 1

Exit Sub

' Errors
ScanDataPlotLineRGBError:
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

