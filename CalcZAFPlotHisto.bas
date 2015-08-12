Attribute VB_Name = "CodeCalcZAFPlotHisto"
' (c) Copyright 1995-2015 by John J. Donovan
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
Dim smin As Single, smax As Single

' Calculate bucket width
hstep! = (hmax! - hmin!) / nbin%

For i% = 1 To nCol%
For k% = 1 To nbin%
'xdata(i%, k%) = hmin! + hstep! * (k% - 1) + hstep! / 2# ' calculate buckets centered on intervals
xdata(i%, k%) = hmin! + hstep! * (k% - 1)
ydata(i%, k%) = 0#
Next k%

' Calculate bucket number and increment bucket
For n& = 1 To nRow&
If rarray!(i%, n&) < hmin! Then
m% = 1

ElseIf rarray!(i%, n&) > hmax! Then
m% = nbin%

Else
For k% = 1 To nbin%
smin! = hmin! + hstep! * (k% - 1)
smax! = smin! + hstep!
If rarray!(i%, n&) >= smin! And rarray!(i%, n&) < smax! Then m% = k%
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

Sub CalcZAFPlotHistogram_GS(CalcZAFOutputCount As Long, KratioError() As Single)
' Calculate and plot the histogram

ierror = False
On Error GoTo CalcZAFPlotHistogram_GSError

Dim i As Integer, nCol As Integer
Dim TickGap As Single, AllowPct As Single

Dim xdata() As Single, ydata() As Single

Dim average As TypeAverage

If CalcZAFOutputCount& < 1 Then GoTo CalcZAFPlotHistogram_GSNoData
If HistogramNumberofBuckets% = 0 Then GoTo CalcZAFPlotHistogram_GSNoBuckets

ReDim xdata(1 To CalcZAFOutputCount&, 1 To HistogramNumberofBuckets%) As Single
ReDim ydata(1 To CalcZAFOutputCount&, 1 To HistogramNumberofBuckets%) As Single

' Special code for single binary
If CalcZAFOutputCount& < 2 Then
ReDim xdata(1 To 2, 1 To HistogramNumberofBuckets%) As Single
ReDim ydata(1 To 2, 1 To HistogramNumberofBuckets%) As Single
End If

' Calculate histogram
Call CalcZAFHistogram(HistogramMinimum!, HistogramMaximum!, HistogramNumberofBuckets%, CalcZAFOutputCount&, 2, KratioError!(), xdata!(), ydata!())
If ierror Then Exit Sub

' Plot data into control, clear the graph
FormPLOTHISTO_GS.Graph1.DrawMode = 1

' Display plot and fit
FormPLOTHISTO_GS.Graph1.DataReset = 9   ' reset all array based properties
FormPLOTHISTO_GS.Graph1.NumPoints = HistogramNumberofBuckets%
FormPLOTHISTO_GS.Graph1.NumSets = 1
FormPLOTHISTO_GS.Graph1.AutoInc = 0

FormPLOTHISTO_GS.Graph1.XAxisStyle = 2 ' user defined
FormPLOTHISTO_GS.Graph1.XAxisTicks = HistogramNumberofBuckets% / 4#
FormPLOTHISTO_GS.Graph1.XAxisMinorTicks = -1   ' 1 minor ticks per tick

FormPLOTHISTO_GS.Graph1.YAxisStyle = 1 ' automatic

' Printer info
FormPLOTHISTO_GS.Graph1.PrintInfo(11) = 1  ' landscape
FormPLOTHISTO_GS.Graph1.PrintInfo(12) = 1  ' fit to page

' Set aspect ratio for axes
FormPLOTHISTO_GS.Graph1.XAxisMin = HistogramMinimum!
FormPLOTHISTO_GS.Graph1.XAxisMax = HistogramMaximum!

' Calculate best bar gap
TickGap! = (FormPLOTHISTO_GS.Graph1.XAxisMax - FormPLOTHISTO_GS.Graph1.XAxisMin) / CDbl(FormPLOTHISTO_GS.Graph1.XAxisTicks)

' Each bar can use this fraction of the distance
AllowPct! = 0.025 * (FormPLOTHISTO_GS.Graph1.XAxisMax - FormPLOTHISTO_GS.Graph1.XAxisMin) / TickGap!

' Convert from fraction used, to percentage unused, and add 1 to allow for roundoff error.
FormPLOTHISTO_GS.Graph1.Bar2DGap = Int(100# * (1# - AllowPct!)) + 1

' Determine column to load
If FormPLOTHISTO_GS.OptionColumnNumber(0).Value Then
nCol% = 1
Else
nCol% = 2
End If

' Load y axis data
For i% = 1 To HistogramNumberofBuckets%
FormPLOTHISTO_GS.Graph1.Data(i%) = ydata!(nCol%, i%)
FormPLOTHISTO_GS.Graph1.xpos(i%) = xdata!(nCol%, i%)
FormPLOTHISTO_GS.Graph1.Color(i%) = 0
Next i%

FormPLOTHISTO_GS.Graph1.BottomTitle = "Relative Error" & ", " & Str$(HistogramNumberofBuckets%) & " Bins (" & Str$(HistogramMinimum!) & " - " & Str$(HistogramMaximum!) & " )"
FormPLOTHISTO_GS.Graph1.LeftTitleStyle = 1
FormPLOTHISTO_GS.Graph1.LeftTitle = "Number of Binaries" & " (" & Str$(CalcZAFOutputCount&) & ")"

If HistogramOutputOption% = -1 Then msg$ = "First Approximation Atomic Fractions"
If HistogramOutputOption% = -2 Then msg$ = "First Approximation Weight Fractions"
If HistogramOutputOption% = -3 Then msg$ = "First Approximation Electron Fractions"
If HistogramOutputOption% = 0 Then msg$ = "Calculated Intensities (K-ratios)"
If HistogramOutputOption% = 1 Then msg$ = "Calculated Intensities (1st Approx. Atomic Fractions)"
If HistogramOutputOption% = 2 Then msg$ = "Calculated Intensities (1st Approx. Weight Fractions)"
If HistogramOutputOption% = 3 Then msg$ = "Calculated Intensities (1st Approx. Electron Fractions)"

FormPLOTHISTO_GS.Graph1.GraphTitle = ImportDataFile2$

' Load legends for options
FormPLOTHISTO_GS.Graph1.AutoInc = 1
FormPLOTHISTO_GS.Graph1.LegendText = msg$

' Corrections to First approximation
If HistogramOutputOption% >= 0 Then

If HistogramOutputOption% > 0 And FirstApproximationApplyAbsorption Then
FormPLOTHISTO_GS.Graph1.LegendText = "First Approximations corrected for Absorption"
End If
If HistogramOutputOption% > 0 And FirstApproximationApplyFluorescence Then
FormPLOTHISTO_GS.Graph1.LegendText = "First Approximations corrected for Fluorescence"
End If
If HistogramOutputOption% > 0 And FirstApproximationApplyAtomicNumber Then
FormPLOTHISTO_GS.Graph1.LegendText = "First Approximations corrected for Atomic Number"
End If

' 0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters
FormPLOTHISTO_GS.Graph1.LegendText = corstring$(CorrectionFlag%)

' Output flags
If BinaryOutputRangeMinAbs Then
FormPLOTHISTO_GS.Graph1.LegendText = "Minimum Absorption Correction=" & Str$(BinaryOutputRangeAbsMin!)
End If
If BinaryOutputRangeMinFlu Then
FormPLOTHISTO_GS.Graph1.LegendText = "Minimum Fluorescence Correction=" & Str$(BinaryOutputRangeFluMin!)
End If
If BinaryOutputRangeMinZed Then
FormPLOTHISTO_GS.Graph1.LegendText = "Minimum Atomic Number Correction=" & Str$(BinaryOutputRangeZedMin!)
End If

If BinaryOutputRangeMaxAbs Then
FormPLOTHISTO_GS.Graph1.LegendText = "Maximum Absorption Correction=" & Str$(BinaryOutputRangeAbsMax!)
End If

If BinaryOutputRangeMaxFlu Then
FormPLOTHISTO_GS.Graph1.LegendText = "Maximum Fluorescence Correction=" & Str$(BinaryOutputRangeFluMax!)
End If

If BinaryOutputRangeMaxZed Then
FormPLOTHISTO_GS.Graph1.LegendText = "Maximum Atomic Number Correction=" & Str$(BinaryOutputRangeZedMax!)
End If

' Zbar filters
If BinaryOutputMinimumZbar Then
FormPLOTHISTO_GS.Graph1.LegendText = "Minimum Mass-Electron Zbar Difference" & Str$(BinaryOutputMinimumZbarDiff!) & "%"
End If

If BinaryOutputMaximumZbar Then
FormPLOTHISTO_GS.Graph1.LegendText = "Maximum Mass-Electron Zbar Difference" & Str$(BinaryOutputMaximumZbarDiff!) & "%"
End If

' Correction options
If CorrectionFlag% = 0 Or (CorrectionFlag% >= 1 And CorrectionFlag% <= 4 And UsePenepmaKratiosFlag% = 1) Then
FormPLOTHISTO_GS.Graph1.LegendText = bscstring$(ibsc%)
FormPLOTHISTO_GS.Graph1.LegendText = mipstring$(imip%)
FormPLOTHISTO_GS.Graph1.LegendText = stpstring$(istp%)
FormPLOTHISTO_GS.Graph1.LegendText = bksstring$(ibks%)
FormPLOTHISTO_GS.Graph1.LegendText = absstring$(iabs%)
FormPLOTHISTO_GS.Graph1.LegendText = flustring$(iflu%)
FormPLOTHISTO_GS.Graph1.LegendText = macstring$(MACTypeFlag%)

' Note beta fluorescence
If Not UseFluorescenceByBetaLinesFlag Then
FormPLOTHISTO_GS.Graph1.LegendText = "Fluorescence of/by beta lines not included"
Else
FormPLOTHISTO_GS.Graph1.LegendText = "Fluorescence of/by beta lines included"
End If

If CorrectionFlag% >= 1 And CorrectionFlag% <= 4 Then
FormPLOTHISTO_GS.Graph1.LegendText = empstring$(EmpiricalAlphaFlag%)
End If

' Using fundamental parameters
ElseIf CorrectionFlag% = 5 Then
FormPLOTHISTO_GS.Graph1.LegendText = "Using Fundamental Parameters"

' Using Penepma derived k-ratios alpha factors
Else
    If Not UsePenepmaKratiosLimitFlag Then
    FormPLOTHISTO_GS.Graph1.LegendText = "Using Penepma k-ratios if available..."
    Else
    FormPLOTHISTO_GS.Graph1.LegendText = "Using Penepma k-ratios if available...(" & Format$(PenepmaKratiosLimitValue!) & " % limit)"
    End If
End If
End If

' Calculate average and standard deviation
Call MathArrayAverage3(average, KratioError!(), CalcZAFOutputCount&, 2)
If ierror Then Exit Sub

' Print average and standard deviation on plot
FormPLOTHISTO_GS.LabelAverage.Caption = MiscAutoFormat$(average.averags!(nCol%))
FormPLOTHISTO_GS.LabelStdDev.Caption = MiscAutoFormat$(average.Stddevs!(nCol%))
FormPLOTHISTO_GS.LabelMinimum.Caption = MiscAutoFormat$(average.Minimums!(nCol%))
FormPLOTHISTO_GS.LabelMaximum.Caption = MiscAutoFormat$(average.Maximums!(nCol%))

' Show plot
FormPLOTHISTO_GS.Graph1.DrawMode = 2

Exit Sub

' Errors
CalcZAFPlotHistogram_GSError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFPlotHistogram_GS"
Call CalcZAFImportClose
ierror = True
Exit Sub

CalcZAFPlotHistogram_GSNoData:
msg$ = "No CalcZAF data to plot"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFPlotHistogram_GS"
ierror = True
Exit Sub

CalcZAFPlotHistogram_GSNoBuckets:
msg$ = "No histogram data to plot"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFPlotHistogram_GS"
ierror = True
Exit Sub

End Sub

Sub CalcZAFPlotHistogram_PE(CalcZAFOutputCount As Long, KratioError() As Single)
' Calculate and plot the histogram

ierror = False
On Error GoTo CalcZAFPlotHistogram_PEError

Dim i As Integer, nCol As Integer
Dim TickGap As Single, AllowPct As Single

Dim xdata() As Single, ydata() As Single

Dim average As TypeAverage

If CalcZAFOutputCount& < 1 Then GoTo CalcZAFPlotHistogram_PENoData
If HistogramNumberofBuckets% = 0 Then GoTo CalcZAFPlotHistogram_PENoBuckets

ReDim xdata(1 To CalcZAFOutputCount&, 1 To HistogramNumberofBuckets%) As Single
ReDim ydata(1 To CalcZAFOutputCount&, 1 To HistogramNumberofBuckets%) As Single

' Special code for single binary
If CalcZAFOutputCount& < 2 Then
ReDim xdata(1 To 2, 1 To HistogramNumberofBuckets%) As Single
ReDim ydata(1 To 2, 1 To HistogramNumberofBuckets%) As Single
End If

' Calculate histogram
Call CalcZAFHistogram(HistogramMinimum!, HistogramMaximum!, HistogramNumberofBuckets%, CalcZAFOutputCount, 2, KratioError!(), xdata!(), ydata!())
If ierror Then Exit Sub

' Plot data into control, clear the graph
FormPLOTHISTO_PE.Graph1.DrawMode = 1

' Display plot and fit
FormPLOTHISTO_PE.Graph1.DataReset = 9   ' reset all array based properties
FormPLOTHISTO_PE.Graph1.NumPoints = HistogramNumberofBuckets%
FormPLOTHISTO_PE.Graph1.NumSets = 1
FormPLOTHISTO_PE.Graph1.AutoInc = 0

FormPLOTHISTO_PE.Graph1.XAxisStyle = 2 ' user defined
FormPLOTHISTO_PE.Graph1.XAxisTicks = HistogramNumberofBuckets% / 4#
FormPLOTHISTO_PE.Graph1.XAxisMinorTicks = -1   ' 1 minor ticks per tick

FormPLOTHISTO_PE.Graph1.YAxisStyle = 1 ' automatic

' Printer info
FormPLOTHISTO_PE.Graph1.PrintInfo(11) = 1  ' landscape
FormPLOTHISTO_PE.Graph1.PrintInfo(12) = 1  ' fit to page

' Set aspect ratio for axes
FormPLOTHISTO_PE.Graph1.XAxisMin = HistogramMinimum!
FormPLOTHISTO_PE.Graph1.XAxisMax = HistogramMaximum!

' Calculate best bar gap
TickGap! = (FormPLOTHISTO_PE.Graph1.XAxisMax - FormPLOTHISTO_PE.Graph1.XAxisMin) / CDbl(FormPLOTHISTO_PE.Graph1.XAxisTicks)

' Each bar can use this fraction of the distance
AllowPct! = 0.025 * (FormPLOTHISTO_PE.Graph1.XAxisMax - FormPLOTHISTO_PE.Graph1.XAxisMin) / TickGap!

' Convert from fraction used, to percentage unused, and add 1 to allow for roundoff error.
FormPLOTHISTO_PE.Graph1.Bar2DGap = Int(100# * (1# - AllowPct!)) + 1

' Determine column to load
If FormPLOTHISTO_PE.OptionColumnNumber(0).Value Then
nCol% = 1
Else
nCol% = 2
End If

' Load y axis data
For i% = 1 To HistogramNumberofBuckets%
FormPLOTHISTO_PE.Graph1.Data(i%) = ydata!(nCol%, i%)
FormPLOTHISTO_PE.Graph1.xpos(i%) = xdata!(nCol%, i%)
FormPLOTHISTO_PE.Graph1.Color(i%) = 0
Next i%

FormPLOTHISTO_PE.Graph1.BottomTitle = "Relative Error" & ", " & Str$(HistogramNumberofBuckets%) & " Bins (" & Str$(HistogramMinimum!) & " - " & Str$(HistogramMaximum!) & " )"
FormPLOTHISTO_PE.Graph1.LeftTitleStyle = 1
FormPLOTHISTO_PE.Graph1.LeftTitle = "Number of Binaries" & " (" & Str$(CalcZAFOutputCount&) & ")"

If HistogramOutputOption% = -1 Then msg$ = "First Approximation Atomic Fractions"
If HistogramOutputOption% = -2 Then msg$ = "First Approximation Weight Fractions"
If HistogramOutputOption% = -3 Then msg$ = "First Approximation Electron Fractions"
If HistogramOutputOption% = 0 Then msg$ = "Calculated Intensities (K-ratios)"
If HistogramOutputOption% = 1 Then msg$ = "Calculated Intensities (1st Approx. Atomic Fractions)"
If HistogramOutputOption% = 2 Then msg$ = "Calculated Intensities (1st Approx. Weight Fractions)"
If HistogramOutputOption% = 3 Then msg$ = "Calculated Intensities (1st Approx. Electron Fractions)"

FormPLOTHISTO_PE.Graph1.GraphTitle = ImportDataFile2$

' Load legends for options
FormPLOTHISTO_PE.Graph1.AutoInc = 1
FormPLOTHISTO_PE.Graph1.LegendText = msg$

' Corrections to First approximation
If HistogramOutputOption% >= 0 Then

If HistogramOutputOption% > 0 And FirstApproximationApplyAbsorption Then
FormPLOTHISTO_PE.Graph1.LegendText = "First Approximations corrected for Absorption"
End If
If HistogramOutputOption% > 0 And FirstApproximationApplyFluorescence Then
FormPLOTHISTO_PE.Graph1.LegendText = "First Approximations corrected for Fluorescence"
End If
If HistogramOutputOption% > 0 And FirstApproximationApplyAtomicNumber Then
FormPLOTHISTO_PE.Graph1.LegendText = "First Approximations corrected for Atomic Number"
End If

' 0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters
FormPLOTHISTO_PE.Graph1.LegendText = corstring$(CorrectionFlag%)

' Output flags
If BinaryOutputRangeMinAbs Then
FormPLOTHISTO_PE.Graph1.LegendText = "Minimum Absorption Correction=" & Str$(BinaryOutputRangeAbsMin!)
End If
If BinaryOutputRangeMinFlu Then
FormPLOTHISTO_PE.Graph1.LegendText = "Minimum Fluorescence Correction=" & Str$(BinaryOutputRangeFluMin!)
End If
If BinaryOutputRangeMinZed Then
FormPLOTHISTO_PE.Graph1.LegendText = "Minimum Atomic Number Correction=" & Str$(BinaryOutputRangeZedMin!)
End If

If BinaryOutputRangeMaxAbs Then
FormPLOTHISTO_PE.Graph1.LegendText = "Maximum Absorption Correction=" & Str$(BinaryOutputRangeAbsMax!)
End If

If BinaryOutputRangeMaxFlu Then
FormPLOTHISTO_PE.Graph1.LegendText = "Maximum Fluorescence Correction=" & Str$(BinaryOutputRangeFluMax!)
End If

If BinaryOutputRangeMaxZed Then
FormPLOTHISTO_PE.Graph1.LegendText = "Maximum Atomic Number Correction=" & Str$(BinaryOutputRangeZedMax!)
End If

' Zbar filters
If BinaryOutputMinimumZbar Then
FormPLOTHISTO_PE.Graph1.LegendText = "Minimum Mass-Electron Zbar Difference" & Str$(BinaryOutputMinimumZbarDiff!) & "%"
End If
If BinaryOutputMaximumZbar Then
FormPLOTHISTO_PE.Graph1.LegendText = "Maximum Mass-Electron Zbar Difference" & Str$(BinaryOutputMaximumZbarDiff!) & "%"
End If

' Correction options
If CorrectionFlag% = 0 Or (CorrectionFlag% >= 1 And CorrectionFlag% <= 4 And UsePenepmaKratiosFlag% = 1) Then
FormPLOTHISTO_GS.Graph1.LegendText = bscstring$(ibsc%)
FormPLOTHISTO_GS.Graph1.LegendText = mipstring$(imip%)
FormPLOTHISTO_GS.Graph1.LegendText = stpstring$(istp%)
FormPLOTHISTO_GS.Graph1.LegendText = bksstring$(ibks%)
FormPLOTHISTO_GS.Graph1.LegendText = absstring$(iabs%)
FormPLOTHISTO_GS.Graph1.LegendText = flustring$(iflu%)
FormPLOTHISTO_GS.Graph1.LegendText = macstring$(MACTypeFlag%)

' Note beta fluorescence
If Not UseFluorescenceByBetaLinesFlag Then
FormPLOTHISTO_GS.Graph1.LegendText = "Fluorescence of/by beta lines not included"
Else
FormPLOTHISTO_GS.Graph1.LegendText = "Fluorescence of/by beta lines included"
End If

If CorrectionFlag% >= 1 And CorrectionFlag% <= 4 Then
FormPLOTHISTO_GS.Graph1.LegendText = empstring$(EmpiricalAlphaFlag%)
End If

' Using fundamental parameters
ElseIf CorrectionFlag% = 5 Then
FormPLOTHISTO_GS.Graph1.LegendText = "Using Fundamental Parameters"

' Using Penepma derived k-ratios alpha factors
Else
    If Not UsePenepmaKratiosLimitFlag Then
    FormPLOTHISTO_GS.Graph1.LegendText = "Using Penepma k-ratios if available..."
    Else
    FormPLOTHISTO_GS.Graph1.LegendText = "Using Penepma k-ratios if available...(" & Format$(PenepmaKratiosLimitValue!) & " % limit)"
    End If
End If
End If

' Calculate average and standard deviation
Call MathArrayAverage3(average, KratioError!(), CalcZAFOutputCount&, 2)
If ierror Then Exit Sub

' Print average and standard deviation on plot
FormPLOTHISTO_PE.LabelAverage.Caption = MiscAutoFormat$(average.averags!(nCol%))
FormPLOTHISTO_PE.LabelStdDev.Caption = MiscAutoFormat$(average.Stddevs!(nCol%))
FormPLOTHISTO_PE.LabelMinimum.Caption = MiscAutoFormat$(average.Minimums!(nCol%))
FormPLOTHISTO_PE.LabelMaximum.Caption = MiscAutoFormat$(average.Maximums!(nCol%))

' Show plot
FormPLOTHISTO_PE.Graph1.DrawMode = 2

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

