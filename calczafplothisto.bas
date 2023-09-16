Attribute VB_Name = "CodeCalcZAFPlotHisto"
' (c) Copyright 1995-2023 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

' Histogram variables
Dim HistogramMinimum As Single
Dim HistogramMaximum As Single
Dim HistogramNumberofBuckets As Integer
Dim HistogramOutputOption As Integer

Global BinaryOutputRangeMinAbs As Integer
Global BinaryOutputRangeMinFlu As Integer
Global BinaryOutputRangeMinZed As Integer

Global BinaryOutputRangeMaxAbs As Integer
Global BinaryOutputRangeMaxFlu As Integer
Global BinaryOutputRangeMaxZed As Integer

Global BinaryOutputRangeAbsMin As Single
Global BinaryOutputRangeFluMin As Single
Global BinaryOutputRangeZedMin As Single

Global BinaryOutputRangeAbsMax As Single
Global BinaryOutputRangeFluMax As Single
Global BinaryOutputRangeZedMax As Single

Global BinaryOutputMinimumZbar As Integer
Global BinaryOutputMaximumZbar As Integer

Global BinaryOutputMinimumZbarDiff As Single
Global BinaryOutputMaximumZbarDiff As Single

Global FirstApproximationApplyAbsorption As Integer
Global FirstApproximationApplyFluorescence As Integer
Global FirstApproximationApplyAtomicNumber As Integer

Sub CalcZAFHistogram(hmin As Single, hmax As Single, nbin As Integer, nRow As Long, nCol As Integer, rarray() As Single, xdata() As Single, ydata() As Single)
' Calculate and output histogram data for the passed array
' hmin is histogram minimum
' hmax is histogram maximum
' nbin is the number of histogram buckets
' nrow is the number of elements (rows) in the data array
' ncol is the number of elements (columns) in the data array
' rarray() is the data array (two dimensional)

ierror = False
On Error GoTo CalcZAFHistogramError

Dim i As Integer, k As Integer, m As Integer
Dim n As Long
Dim hstep As Single

' Calculate bucket width
hstep! = (hmax! - hmin!) / (nbin% - 1)

For i% = 1 To nCol%

For k% = 1 To nbin%
xdata(i%, k%) = hmin! + hstep! * (k% - 1) + hstep! / 2#     ' calculate buckets centered on intervals
Next k%

' Calculate bucket number and increment bucket
For n& = 1 To nRow&
If rarray!(i%, n&) < hmin! Then
m% = 1

ElseIf rarray!(i%, n&) > hmax! Then
m% = nbin%

Else
For k% = 1 To nbin%
If rarray!(i%, n&) >= (xdata!(i%, k%) - hstep! / 2#) And rarray!(i%, n&) <= (xdata!(i%, k%) + hstep! / 2#) Then
m% = k%
Exit For
End If
Next k%
End If

' Increment data
ydata(i%, m%) = ydata(i%, m%) + 1#

Next n&
Next i%

' Output data
Open HistogramDataFile$ For Output As #HistogramDataFileNumber%

' Loop on each bin
For k% = 1 To nbin%

' Loop on each column
msg$ = vbNullString
For i% = 1 To nCol%
msg$ = msg$ & MiscAutoFormat$(xdata!(i%, k%)) & vbTab & MiscAutoFormat$(ydata!(i%, k%)) & vbTab
Next i%

Print #HistogramDataFileNumber%, msg$
Next k%

Close #HistogramDataFileNumber%
Exit Sub

' Errors
CalcZAFHistogramError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFHistogram"
Call CalcZAFImportClose
ierror = True
Exit Sub

End Sub

Sub CalcZAFHistogramLoad()
' Load histogram options

ierror = False
On Error GoTo CalcZAFHistogramLoadError

Static initialized As Integer

If Not initialized Then
HistogramMinimum! = 0.5
HistogramMaximum! = 1.5
HistogramNumberofBuckets% = 40
initialized = True
End If

' Histogram options
FormHISTO.TextHistogramMinimum.Text = Str$(HistogramMinimum!)
FormHISTO.TextHistogramMaximum.Text = Str$(HistogramMaximum!)
FormHISTO.TextHistogramNumberofBuckets.Text = Str$(HistogramNumberofBuckets%)

Exit Sub

' Errors
CalcZAFHistogramLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFHistogramLoad"
ierror = True
Exit Sub

End Sub

Sub CalcZAFHistogramSave()
' Save histogram options

ierror = False
On Error GoTo CalcZAFHistogramSaveError

' Histogram options
HistogramMinimum! = FormHISTO.TextHistogramMinimum.Text

If Val(FormHISTO.TextHistogramMaximum.Text) > HistogramMinimum! Then
HistogramMaximum! = Val(FormHISTO.TextHistogramMaximum.Text)
Else
msg$ = "Histogram maximum is less than minimum"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFHistogramSave"
End If

If Val(FormHISTO.TextHistogramNumberofBuckets.Text) > 5 Then
HistogramNumberofBuckets% = Val(FormHISTO.TextHistogramNumberofBuckets.Text)
Else
msg$ = "Histogram number of buckets is too small"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFHistogramSave"
End If

Exit Sub

' Errors
CalcZAFHistogramSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFHistogramSave"
ierror = True
Exit Sub

End Sub

Sub CalcZAFPlotHistogram_PE(CalcZAFOutputCount As Long, KratioError() As Single)
' Calculate and plot the histogram

ierror = False
On Error GoTo CalcZAFPlotHistogram_PEError

Dim i As Integer, nCol As Integer, acounter As Integer
Dim xannotation As Single, yannotation As Single, ydecrement As Single
Dim ymin As Single, ymax As Single

Dim xdata() As Single, ydata() As Single

Dim average As TypeAverage

If CalcZAFOutputCount& < 1 Then GoTo CalcZAFPlotHistogram_PENoData
If HistogramNumberofBuckets% = 0 Then GoTo CalcZAFPlotHistogram_PENoBuckets

' Init the graph
Call MiscPlotInit(FormPLOTHISTO_PE.Pesgo1, True)
If ierror Then Exit Sub

FormPLOTHISTO_PE.Pesgo1.ShowAnnotations = True
FormPLOTHISTO_PE.Pesgo1.AnnotationsInFront = True

FormPLOTHISTO_PE.Pesgo1.ShowTickMarkY = PESTM_TICKS_HIDE&
FormPLOTHISTO_PE.Pesgo1.ShowTickMarkX = PESTM_TICKS_OUTSIDE&
FormPLOTHISTO_PE.Pesgo1.ImageAdjustRight = -80                      ' axis formatting

' Plot type
FormPLOTHISTO_PE.Pesgo1.PlottingMethod = SGPM_BAR&                  ' bargraph subset (see below for setting bar width)

FormPLOTHISTO_PE.Pesgo1.ManualScaleControlX = PEMSC_MINMAX          ' manually Control X Axis
FormPLOTHISTO_PE.Pesgo1.ManualMinX = HistogramMinimum!
FormPLOTHISTO_PE.Pesgo1.ManualMaxX = HistogramMaximum!

FormPLOTHISTO_PE.Pesgo1.ManualScaleControlY = PEMSC_MIN             ' autoscale Control Y Axis max, Manual Control min
FormPLOTHISTO_PE.Pesgo1.ManualMinY = 0

FormPLOTHISTO_PE.Pesgo1.GridLineControl = PEGLC_NONE&
FormPLOTHISTO_PE.Pesgo1.GridBands = False                           ' removes color banding on background

' Define #subset and #points
FormPLOTHISTO_PE.Pesgo1.Subsets = 1
FormPLOTHISTO_PE.Pesgo1.points = HistogramNumberofBuckets%

' Load histogram data
ReDim xdata(1 To CalcZAFOutputCount&, 1 To HistogramNumberofBuckets%) As Single
ReDim ydata(1 To CalcZAFOutputCount&, 1 To HistogramNumberofBuckets%) As Single

' Special code for single binary
If CalcZAFOutputCount& < 2 Then
ReDim xdata(1 To 2, 1 To HistogramNumberofBuckets%) As Single
ReDim ydata(1 To 2, 1 To HistogramNumberofBuckets%) As Single
End If

' Calculate histogram
Call CalcZAFHistogram(HistogramMinimum!, HistogramMaximum!, HistogramNumberofBuckets%, CalcZAFOutputCount, Int(2), KratioError!(), xdata!(), ydata!())
If ierror Then Exit Sub

' Determine column to load
If FormPLOTHISTO_PE.OptionColumnNumber(0).Value Then
nCol% = 1
Else
nCol% = 2
End If

' Set bar width
FormPLOTHISTO_PE.Pesgo1.BarWidth = (HistogramMaximum! - HistogramMinimum!) / HistogramNumberofBuckets%

'FormPLOTHISTO_PE.Pesgo1.BarWidth = 0                                                         ' 0 for auto width
FormPLOTHISTO_PE.Pesgo1.AdjoinBars = True
FormPLOTHISTO_PE.Pesgo1.SubsetColors(0) = FormPLOTHISTO_PE.Pesgo1.PEargb(Int(255), Int(255), Int(0), Int(0))             ' red bars
FormPLOTHISTO_PE.Pesgo1.BarBorderColor = FormPLOTHISTO_PE.Pesgo1.PEargb(Int(255), Int(0), Int(0), Int(0))                ' black border

' Load y axis data
For i% = 1 To HistogramNumberofBuckets%
FormPLOTHISTO_PE.Pesgo1.ydata(0, i% - 1) = ydata!(nCol%, i%)
FormPLOTHISTO_PE.Pesgo1.xdata(0, i% - 1) = xdata!(nCol%, i%)
Next i%

FormPLOTHISTO_PE.Pesgo1.XAxisLabel = "Relative Error" & ", " & Str$(HistogramNumberofBuckets%) & " Bins (" & Str$(HistogramMinimum!) & " - " & Str$(HistogramMaximum!) & " )"
FormPLOTHISTO_PE.Pesgo1.YAxisLabel = "Number of Binaries" & " (" & Str$(CalcZAFOutputCount&) & ")"
FormPLOTHISTO_PE.Pesgo1.MainTitle = ImportDataFile2$

FormPLOTHISTO_PE.Pesgo1.PEactions = REINITIALIZE_RESETIMAGE    ' generate new plot

' Calculate average and standard deviation
Call MathArrayAverage3(average, KratioError!(), CalcZAFOutputCount&, 2)
If ierror Then Exit Sub

' Print average and standard deviation on plot
FormPLOTHISTO_PE.LabelAverage.Caption = MiscAutoFormat$(average.averags!(nCol%))
FormPLOTHISTO_PE.LabelStdDev.Caption = MiscAutoFormat$(average.Stddevs!(nCol%))
FormPLOTHISTO_PE.LabelMinimum.Caption = MiscAutoFormat$(average.Minimums!(nCol%))
FormPLOTHISTO_PE.LabelMaximum.Caption = MiscAutoFormat$(average.Maximums!(nCol%))

' Annotation properties
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationTextSize = 60               ' define annotation text size
FormPLOTHISTO_PE.Pesgo1.LabelFont = "Arial"                        ' define Font for annotations (and axes)
FormPLOTHISTO_PE.Pesgo1.HideIntersectingText = PEHIT_NO_HIDING&    ' or PEHIT_HIDE&

If HistogramOutputOption% = -1 Then msg$ = "First Approximation Atomic Fractions"
If HistogramOutputOption% = -2 Then msg$ = "First Approximation Weight Fractions"
If HistogramOutputOption% = -3 Then msg$ = "First Approximation Electron Fractions"
If HistogramOutputOption% = 0 Then msg$ = "Calculated Intensities (K-ratios)"
If HistogramOutputOption% = 1 Then msg$ = "Calculated Intensities (1st Approx. Atomic Fractions)"
If HistogramOutputOption% = 2 Then msg$ = "Calculated Intensities (1st Approx. Weight Fractions)"
If HistogramOutputOption% = 3 Then msg$ = "Calculated Intensities (1st Approx. Electron Fractions)"

' Load calculation options as annotations
xannotation! = FormPLOTHISTO_PE.Pesgo1.ManualMinX + (FormPLOTHISTO_PE.Pesgo1.ManualMaxX - FormPLOTHISTO_PE.Pesgo1.ManualMinX) * 0.8
yannotation! = FormPLOTHISTO_PE.Pesgo1.ManualMaxY * 0.98
ydecrement! = (FormPLOTHISTO_PE.Pesgo1.ManualMaxY - FormPLOTHISTO_PE.Pesgo1.ManualMinY) / 25#
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = msg$

' Corrections to First approximation
If HistogramOutputOption% >= 0 Then
If HistogramOutputOption% > 0 And FirstApproximationApplyAbsorption Then
acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = "First Approximations corrected for Absorption"
End If
If HistogramOutputOption% > 0 And FirstApproximationApplyFluorescence Then
acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = "First Approximations corrected for Fluorescence"
End If
If HistogramOutputOption% > 0 And FirstApproximationApplyAtomicNumber Then
acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = "First Approximations corrected for Atomic Number"
End If

' 0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters
acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = corstring$(CorrectionFlag%)

' Output flags
If BinaryOutputRangeMinAbs Then
acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = "Minimum Absorption Correction=" & Str$(BinaryOutputRangeAbsMin!)
End If
If BinaryOutputRangeMinFlu Then
acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = "Minimum Fluorescence Correction=" & Str$(BinaryOutputRangeFluMin!)
End If
If BinaryOutputRangeMinZed Then
acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = "Minimum Atomic Number Correction=" & Str$(BinaryOutputRangeZedMin!)
End If

If BinaryOutputRangeMaxAbs Then
acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = "Maximum Absorption Correction=" & Str$(BinaryOutputRangeAbsMax!)
End If

If BinaryOutputRangeMaxFlu Then
acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = "Maximum Fluorescence Correction=" & Str$(BinaryOutputRangeFluMax!)
End If

If BinaryOutputRangeMaxZed Then
acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = "Maximum Atomic Number Correction=" & Str$(BinaryOutputRangeZedMax!)
End If

' Zbar filters
If BinaryOutputMinimumZbar Then
acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = "Minimum Mass-Electron Zbar Difference" & Str$(BinaryOutputMinimumZbarDiff!) & "%"
End If

If BinaryOutputMaximumZbar Then
acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = "Maximum Mass-Electron Zbar Difference" & Str$(BinaryOutputMaximumZbarDiff!) & "%"
End If

' Correction options
If CorrectionFlag% = 0 Or (CorrectionFlag% >= 1 And CorrectionFlag% <= 4 And UsePenepmaKratiosFlag% = 1) Then
acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
If ibsc% <> 5 Then
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = bscstring$(ibsc%)
Else
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = bscstring$(ibsc%) & " (x=" & Format$(ZFractionBackscatterExponent!) & ")"
End If

acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = mipstring$(imip%)

acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = stpstring$(istp%)

acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = bksstring$(ibks%)

acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = absstring$(iabs%)

acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = flustring$(iflu%)

acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = macstring$(MACTypeFlag%)

If EmpTypeFlag% = 1 Then
acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = "Using Empirical MACs if available"
End If

' Note beta fluorescence
If Not UseFluorescenceByBetaLinesFlag Then
acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = "Fluorescence of/by beta lines not included"
Else
acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = "Fluorescence of/by beta lines included"
End If

If CorrectionFlag% >= 1 And CorrectionFlag% <= 4 Then
acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = empstring$(EmpiricalAlphaFlag%)
End If

' Using fundamental parameters
ElseIf CorrectionFlag% = 5 Then
acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = "Using Fundamental Parameters"

' Using Penepma derived k-ratios alpha factors
Else
    If Not UsePenepmaKratiosLimitFlag Then
    acounter% = acounter% + 1
    yannotation! = yannotation! - ydecrement!
    FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
    FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
    FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
    FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
    FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = "Using Penepma k-ratios if available..."
    Else
    acounter% = acounter% + 1
    yannotation! = yannotation! - ydecrement!
    FormPLOTHISTO_PE.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
    FormPLOTHISTO_PE.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
    FormPLOTHISTO_PE.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
    FormPLOTHISTO_PE.Pesgo1.GraphAnnotationColor(acounter%) = FormPLOTHISTO_PE.Pesgo1.PEargb(225, 0, 0, 0)       ' black
    FormPLOTHISTO_PE.Pesgo1.GraphAnnotationText(acounter%) = "Using Penepma k-ratios if available...(" & Format$(PenepmaKratiosLimitValue!) & " % limit)"
    End If
End If
End If

' Draw line at 1.0
acounter% = acounter% + 1
ymin! = FormPLOTHISTO_PE.Pesgo1.ManualMinY
ymax! = FormPLOTHISTO_PE.Pesgo1.ManualMaxY
Call ScanDataPlotLine(FormPLOTHISTO_PE.Pesgo1, CLng(acounter%), CSng(1#), ymin!, CSng(1#), ymax!, False, True, Int(255), Int(0), Int(0), Int(0))       ' black
If ierror Then Exit Sub

FormPLOTHISTO_PE.Pesgo1.PEactions = REINITIALIZE_RESETIMAGE    ' generate new plot

Exit Sub

' Errors
CalcZAFPlotHistogram_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFPlotHistogram_PE"
Call CalcZAFImportClose
ierror = True
Exit Sub

CalcZAFPlotHistogram_PENoData:
msg$ = "No data to plot"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFPlotHistogram_PE"
ierror = True
Exit Sub

CalcZAFPlotHistogram_PENoBuckets:
msg$ = "No histogram data to plot"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFPlotHistogram_PE"
ierror = True
Exit Sub

End Sub

Sub CalcZAFPlotHistogramConcentration_PE(CalcZAFOutputCount As Long, KratioConc() As Single, KratioError() As Single, KratioLine() As Long)
' Calculate and plot the histogram as a concentration histogram (x axis = concentration and y axis = error)

ierror = False
On Error GoTo CalcZAFPlotHistogramConcentration_PEError

Dim i As Integer, nCol As Integer

If CalcZAFOutputCount& < 1 Then GoTo CalcZAFPlotHistogramConcentration_PENoData

' Init the graph
Call MiscPlotInit(FormPlotHistoConc.Pesgo1, True)
If ierror Then Exit Sub

' Missing data points
FormPlotHistoConc.Pesgo1.NullDataValueX = 0
FormPlotHistoConc.Pesgo1.NullDataValueY = 0

FormPlotHistoConc.Pesgo1.PlottingMethod = SGPM_POINT&                ' controls point or point plus line, etc
FormPlotHistoConc.Pesgo1.PointSize = PEPS_LARGE&
FormPlotHistoConc.Pesgo1.MinimumPointSize = PEMPS_MEDIUM_LARGE&      ' helps readability if sizing

' Determine column to load from histogram form
If FormPLOTHISTO_PE.OptionColumnNumber(0).Value Then
nCol% = 1
Else
nCol% = 2
End If

' Define #subset and #points
FormPlotHistoConc.Pesgo1.Subsets = 1
FormPlotHistoConc.Pesgo1.points = CalcZAFOutputCount&

' Load y axis data
For i% = 1 To CalcZAFOutputCount&
FormPlotHistoConc.Pesgo1.ydata(0, i% - 1) = KratioError!(nCol%, i%)
FormPlotHistoConc.Pesgo1.xdata(0, i% - 1) = KratioConc!(nCol%, i%)
Next i%

' For data point labels
If FormPLOTHISTO_PE.CheckLabels.Value = vbChecked Then
FormPlotHistoConc.Pesgo1.AllowDataLabels = PEADL_DATAPOINTLABELS&
FormPlotHistoConc.Pesgo1.GraphDataLabels = True
For i% = 1 To CalcZAFOutputCount&
FormPlotHistoConc.Pesgo1.DataPointLabels(0, i% - 1) = Format$(KratioLine&(nCol%, i%))
Next i%
Else
FormPlotHistoConc.Pesgo1.GraphDataLabels = False
End If

' Load axis labels and title
FormPlotHistoConc.Pesgo1.XAxisLabel = "Concentration (Wt. %)"
FormPlotHistoConc.Pesgo1.YAxisLabel = "Relative Error"
FormPlotHistoConc.Pesgo1.MainTitle = ImportDataFile2$

FormPlotHistoConc.Show vbModal

Exit Sub

' Errors
CalcZAFPlotHistogramConcentration_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFPlotHistogramConcentration_PE"
Call CalcZAFImportClose
ierror = True
Exit Sub

CalcZAFPlotHistogramConcentration_PENoData:
msg$ = "No CalcZAF data to plot"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFPlotHistogramConcentration_PE"
ierror = True
Exit Sub

End Sub

