Attribute VB_Name = "CodeScanData2"
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

Const MAXSEGMENTS% = 400

' Load colors
Const CURVECOLOR% = 4           ' red
Const THRESHOLDCOLOR% = 11      ' cyan
Const CENTROIDCOLOR% = 12       ' light red
Const CURRENTCOLOR% = 10        ' light green
Const ONPEAKCOLOR% = 5          ' magenta

' Special arrays for spline fit data and coefficients
Dim nPoints As Integer
Dim xdata() As Single, ydata() As Single
Dim ycoeff() As Double

Sub ScanDataPlotFitCurve_GS(tGraph As Graph, mode As Integer, acoeff1 As Single, acoeff2 As Single, acoeff3 As Single, centroid As Single, threshold As Single, currentonpeak As Single)
' Display the ROM and parabolic peak fit centroid and threshold on the graph (Graphics Server code)
' mode = 1 = parabolic, 2 = gaussian, 3 = maxima, 4 = maximum value, 5 = cubic spline, 6 = (multi-point) exponential

ierror = False
On Error GoTo ScanDataPlotFitCurve_GSError

Dim i As Integer
Dim r As Long
Dim tx As Single, ty As Single

Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double
Dim sxmin As Double, sxmax As Double, symin As Double, symax As Double
Dim txmin As Double, txmax As Double, tymin As Double, tymax As Double

Dim gxorg As Double, gyorg As Double
Dim gxlen As Double, gylen As Double

' Determine min and max of graph (in user data units)
xmin# = tGraph.SDKInfo(2)
xmax# = tGraph.SDKInfo(1)

ymin# = tGraph.SDKInfo(4)
ymax# = tGraph.SDKInfo(3)

gxorg# = tGraph.SDKInfo(7)
gyorg# = tGraph.SDKInfo(8)

gxlen# = tGraph.SDKInfo(5)
gylen# = tGraph.SDKInfo(6)

' Do not plot curve if not full size (use toolbar as flag)
If tGraph.Toolbar = 2 Or tGraph.Toolbar = 3 Then

' Check for valid curve
If (acoeff1! <> 0# And acoeff2! <> 0# And acoeff3! <> 0#) Or mode% = 5 Then

' Calculate line to draw based on fit coefficients
sxmax# = xmin#
For i% = 1 To MAXSEGMENTS%

' Calculate partial line segments for x
sxmin# = sxmax#
sxmax# = sxmin# + (xmax# - xmin#) / (MAXSEGMENTS% - 1)

' Calculate partial line segments for y
If mode% = 1 Then        ' parabolic
symin# = acoeff1! + sxmin# * acoeff2! + sxmin# ^ 2 * acoeff3!
symax# = acoeff1! + sxmax# * acoeff2! + sxmax# ^ 2 * acoeff3!

ElseIf mode% = 2 Then    ' gaussian
symin# = acoeff1! + sxmin# * acoeff2! + sxmin# ^ 2 * acoeff3!
symax# = acoeff1! + sxmax# * acoeff2! + sxmax# ^ 2 * acoeff3!
If symin# > MAXLOGEXPD! Then symin# = MAXLOGEXPD!
If symax# > MAXLOGEXPD! Then symax# = MAXLOGEXPD!
symin# = NATURALE# ^ symin#
symax# = NATURALE# ^ symax#

ElseIf mode% = 5 Then    ' cubic spline
tx! = CSng(sxmin#)
Call SplineInterpolate(xdata!(), ydata!(), ycoeff#(), CLng(nPoints%), tx!, ty!)
If ierror Then Exit Sub
symin# = CDbl(ty!)
tx! = CSng(sxmax#)
Call SplineInterpolate(xdata!(), ydata!(), ycoeff#(), CLng(nPoints%), tx!, ty!)
If ierror Then Exit Sub
symax# = CDbl(ty!)

ElseIf mode% = 6 Then    ' multi-point exponential
symin# = acoeff1! + sxmin# * acoeff2! + sxmin# ^ 2 * acoeff3!
symax# = acoeff1! + sxmax# * acoeff2! + sxmax# ^ 2 * acoeff3!
If symin# > MAXLOGEXPD! Then symin# = MAXLOGEXPD!
If symax# > MAXLOGEXPD! Then symax# = MAXLOGEXPD!
symin# = NATURALE# ^ symin#
symax# = NATURALE# ^ symax#
End If

' Clip
If sxmin# < xmin# Then sxmin# = xmin#
If sxmax# > xmax# Then sxmax# = xmax#
If symin# < ymin# Then symin# = ymin#
If symax# > ymax# Then symax# = ymax#

If symin# > ymax# Then symin# = ymax#
If symax# < ymin# Then symax# = ymin#

' Convert to graph units
Call XrayPlotConvertGraph(tGraph, sxmin#, symin#, txmin#, tymin#)
If ierror Then Exit Sub

Call XrayPlotConvertGraph(tGraph, sxmax#, symax#, txmax#, tymax#)
If ierror Then Exit Sub

If txmin# < gxorg# Then txmin# = gxorg#
If txmax# > gxorg# + gxlen# Then txmax# = gxorg# + gxlen#
If ymin# > 0# And tymin# < gyorg# Then tymin# = gyorg#      ' clip fit plot only if y-axis minimum is positive
If tymax# > gyorg# + gylen# Then tymax# = gyorg# + gylen#

r& = GSLineAbs(txmin#, tymin#, txmax#, tymax#, 4, 2, CURVECOLOR%)
Next i%
End If
End If

' Plot centroid
If centroid! <> 0# Then
Call XrayPlotConvertGraph(tGraph, CDbl(centroid!), ymin#, txmin#, tymin#)
If ierror Then Exit Sub
Call XrayPlotConvertGraph(tGraph, CDbl(centroid!), ymax#, txmax#, tymax#)
If ierror Then Exit Sub

r& = GSLineAbs(txmin#, tymin#, txmax#, tymax#, 4, 2, CENTROIDCOLOR%)
End If

' Plot threshold
If threshold! <> 0# Then
Call XrayPlotConvertGraph(tGraph, xmin#, CDbl(threshold!), txmin#, tymin#)
If ierror Then Exit Sub
Call XrayPlotConvertGraph(tGraph, xmax#, CDbl(threshold!), txmax#, tymax#)
If ierror Then Exit Sub

r& = GSLineAbs(txmin#, tymin#, txmax#, tymax#, 4, 2, THRESHOLDCOLOR%)
End If

' Plot current on-peak
If currentonpeak! <> 0# Then
Call XrayPlotConvertGraph(tGraph, CDbl(currentonpeak!), ymin#, txmin#, tymin#)
If ierror Then Exit Sub
Call XrayPlotConvertGraph(tGraph, CDbl(currentonpeak!), ymax#, txmax#, tymax#)
If ierror Then Exit Sub

r& = GSLineAbs(txmin#, tymin#, txmax#, tymax#, 4, 2, CURRENTCOLOR%)
End If

Exit Sub

' Errors
ScanDataPlotFitCurve_GSError:
MsgBox Error$, vbOKOnly + vbCritical, "ScanDataPlotFitCurve_GS"
ierror = True
Exit Sub

End Sub

Sub ScanDataPlotFitCurve_PE(tGraph As Pesgo, mode As Integer, linecount As Long, acoeff1 As Single, acoeff2 As Single, acoeff3 As Single, centroid As Single, threshold As Single, currentonpeak As Single)
' Display the ROM and parabolic peak fit centroid and threshold on the graph (Pro Essentials code)
' mode: 1 = parabolic, 2 = gaussian, 3 = maxima, 4 = maximum value, 5 = cubic spline, 6 = (multi-point) exponential

ierror = False
On Error GoTo ScanDataPlotFitCurve_PEError

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
If mode% = 1 Then        ' parabolic
symin# = acoeff1! + sxmin# * acoeff2! + sxmin# ^ 2 * acoeff3!
symax# = acoeff1! + sxmax# * acoeff2! + sxmax# ^ 2 * acoeff3!

ElseIf mode% = 2 Then    ' gaussian
symin# = acoeff1! + sxmin# * acoeff2! + sxmin# ^ 2 * acoeff3!
symax# = acoeff1! + sxmax# * acoeff2! + sxmax# ^ 2 * acoeff3!
If symin# > MAXLOGEXPD! Then symin# = MAXLOGEXPD!
If symax# > MAXLOGEXPD! Then symax# = MAXLOGEXPD!
symin# = NATURALE# ^ symin#
symax# = NATURALE# ^ symax#

ElseIf mode% = 5 Then    ' cubic spline
tx! = CSng(sxmin#)
Call SplineInterpolate(xdata!(), ydata!(), ycoeff#(), CLng(nPoints%), tx!, ty!)
If ierror Then Exit Sub
symin# = CDbl(ty!)
tx! = CSng(sxmax#)
Call SplineInterpolate(xdata!(), ydata!(), ycoeff#(), CLng(nPoints%), tx!, ty!)
If ierror Then Exit Sub
symax# = CDbl(ty!)

ElseIf mode% = 6 Then    ' multi-point exponential
symin# = acoeff1! + sxmin# * acoeff2! + sxmin# ^ 2 * acoeff3!
symax# = acoeff1! + sxmax# * acoeff2! + sxmax# ^ 2 * acoeff3!
If symin# > MAXLOGEXPD! Then symin# = MAXLOGEXPD!
If symax# > MAXLOGEXPD! Then symax# = MAXLOGEXPD!
symin# = NATURALE# ^ symin#
symax# = NATURALE# ^ symax#
End If

' Clip to data extents
If sxmin# < xmin# Then sxmin# = xmin#
If sxmax# > xmax# Then sxmax# = xmax#
If symin# < ymin# Then symin# = ymin#
If symax# > ymax# Then symax# = ymax#

If i% = 1 Then
Call ScanDataPlotLine(tGraph, linecount&, sxmin#, symin#, sxmax#, symax#, False, True, Int(255), Int(128), Int(0), Int(0))     ' brown
If ierror Then Exit Sub
Else
Call ScanDataPlotLine(tGraph, linecount&, sxmin#, symin#, sxmax#, symax#, True, True, Int(255), Int(128), Int(0), Int(0))      ' brown
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

Sub ScanDataPlotOnPeak_GS(tGraph As Graph, onpeak As Single)
' Display the selected on-peak position (Graphics Server code)

ierror = False
On Error GoTo ScanDataPlotOnPeak_GSError

Dim r As Long

Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double
Dim txmin As Double, txmax As Double, tymin As Double, tymax As Double

Dim gxorg As Double, gyorg As Double
Dim gxlen As Double, gylen As Double

' Determine min and max of graph (in user data units)
xmin# = tGraph.SDKInfo(2)
xmax# = tGraph.SDKInfo(1)

ymin# = tGraph.SDKInfo(4)
ymax# = tGraph.SDKInfo(3)

gxorg# = tGraph.SDKInfo(7)
gyorg# = tGraph.SDKInfo(8)

gxlen# = tGraph.SDKInfo(5)
gylen# = tGraph.SDKInfo(6)

If txmin# < gxorg# Then txmin# = gxorg#
If txmax# > gxorg# + gxlen# Then txmax# = gxorg# + gxlen#
If ymin# > 0# And tymin# < gyorg# Then tymin# = gyorg#      ' clip fit plot only if y-axis minimum is positive
If tymax# > gyorg# + gylen# Then tymax# = gyorg# + gylen#

' Plot on-peak
If onpeak! <> 0# Then
Call XrayPlotConvertGraph(tGraph, CDbl(onpeak!), ymin#, txmin#, tymin#)
If ierror Then Exit Sub
Call XrayPlotConvertGraph(tGraph, CDbl(onpeak!), ymax#, txmax#, tymax#)
If ierror Then Exit Sub

r& = GSLineAbs(txmin#, tymin#, txmax#, tymax#, 4, 2, ONPEAKCOLOR%)
End If

Exit Sub

' Errors
ScanDataPlotOnPeak_GSError:
MsgBox Error$, vbOKOnly + vbCritical, "ScanDataPlotOnPeak_GS"
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
