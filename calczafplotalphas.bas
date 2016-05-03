Attribute VB_Name = "CodeCalcZAFPlotAlphas"
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Dim CalcZAFAnalysis As TypeAnalysis

Dim CalcZAFTmpSample(1 To 1) As TypeSample

Dim tzaftype As Integer, tmactype As Integer

Sub CalcZAFLoadAlphas_PE(sample() As TypeSample)
' Load binaries for the current sample for alpha factor plotting (Pro Essentials graphing code)

ierror = False
On Error GoTo CalcZAFLoadAlphas_PEError

Dim i As Integer
Dim emitter As Integer, absorber As Integer
Dim inum As Integer
Dim astring As String

' Check for Bence-Albee corrections
If CorrectionFlag% < 1 Or CorrectionFlag% > 4 Then
msg$ = "Bence-Albee corrections are not selected. Changing matrix correction type to polynomial alpha-factors."
MsgBox msg$, vbOKOnly + vbInformation, "CalcZAFLoadAlphas_PE"
CorrectionFlag% = 3
End If

' Load Bence-Albee modes only
For i% = 0 To 3
FormPlotAlpha_PE.OptionBenceAlbee(i%).Caption = corstring(i% + 1)
If i% + 1 = CorrectionFlag% Then FormPlotAlpha_PE.OptionBenceAlbee(i%).Value = True
Next i%

' Calculate current sample
Call CalcZAFCalculate
If ierror Then Exit Sub

' Load to module level
CalcZAFTmpSample(1) = sample(1)

' Calculate each binary in sample
inum% = 0
FormPlotAlpha_PE.ComboPlotAlpha.Clear
For emitter% = 1 To sample(1).LastElm%
For absorber% = 1 To sample(1).LastChan%

' Skip if emitter and absorber are the same (duplicate elements)
If emitter% <> absorber% And sample(1).Elsyms$(emitter%) <> sample(1).Elsyms$(absorber%) Then
inum% = inum% + 1

astring$ = MiscAutoUcase$(sample(1).Elsyup$(emitter%)) & " " & sample(1).Xrsyms$(emitter%) & " in " & MiscAutoUcase$(sample(1).Elsyup$(absorber%))
FormPlotAlpha_PE.ComboPlotAlpha.AddItem astring$
FormPlotAlpha_PE.ComboPlotAlpha.ItemData(FormPlotAlpha_PE.ComboPlotAlpha.NewIndex) = emitter% * MAXCHAN% + absorber%
End If

Next absorber%
Next emitter%

' Check number of binaries
If inum% = 0 Then GoTo CalcZAFLoadAlphas_PENoBinaries

' Check for Penepma k-ratios flag
If UsePenepmaKratiosFlag = 2 Then
FormPlotAlpha_PE.CheckAllOptions.Enabled = False
FormPlotAlpha_PE.CheckAllMacs.Enabled = False
Else
FormPlotAlpha_PE.CheckAllOptions.Enabled = True
FormPlotAlpha_PE.CheckAllMacs.Enabled = True
End If

' Click first binary
FormPlotAlpha_PE.ComboPlotAlpha.ListIndex = 0
FormPlotAlpha_PE.Show vbModal

Exit Sub

' Errors
CalcZAFLoadAlphas_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFLoadAlphas_PE"
ierror = True
Exit Sub

CalcZAFLoadAlphas_PENoBinaries:
msg$ = "No alpha factor binaries to plot for the current sample"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFLoadAlphas_PE"
ierror = True
Exit Sub

End Sub

Sub CalcZAFPlotAlphaFactors_PE()
' Plot alpha factors for the indicated binary (Pro Essentials graphing code)

ierror = False
On Error GoTo CalcZAFPlotAlphaFactors_PEError

Dim itemp As Integer, i As Integer
Dim emitter As Integer, absorber As Integer, k As Integer
Dim astring As String

Dim npts As Integer, nsets As Integer
Dim xdata() As Single, ydata() As Single, acoeff() As Single, stddev As Single

' Init the graph
Call MiscPlotInit(FormPlotAlpha_PE.Pesgo1, True)
If ierror Then Exit Sub

' Get the selected binary
If FormPlotAlpha_PE.ComboPlotAlpha.ListIndex < 0 Then Exit Sub
If FormPlotAlpha_PE.ComboPlotAlpha.ListCount < 1 Then Exit Sub

' Missing data points
FormPlotAlpha_PE.Pesgo1.NullDataValueX = 0
FormPlotAlpha_PE.Pesgo1.NullDataValueY = 0

If FormPlotAlpha_PE.CheckAllOptions.Value = vbUnchecked And FormPlotAlpha_PE.CheckAllMacs.Value = vbUnchecked Then
FormPlotAlpha_PE.Pesgo1.PlottingMethod = SGPM_POINT&               ' controls point or point plus line, etc
FormPlotAlpha_PE.Pesgo1.PointSize = PEPS_LARGE&
FormPlotAlpha_PE.Pesgo1.MinimumPointSize = PEMPS_MEDIUM_LARGE&     ' helps readability if sizing
Else
FormPlotAlpha_PE.Pesgo1.PlottingMethod = SGPM_POINTSPLUSLINE&      ' controls point or point plus line, etc
FormPlotAlpha_PE.Pesgo1.PointSize = PEPS_LARGE&
FormPlotAlpha_PE.Pesgo1.MinimumPointSize = PEMPS_MEDIUM_LARGE&     ' helps readability if sizing
End If

' Determine which binary to calculate
itemp% = FormPlotAlpha_PE.ComboPlotAlpha.ItemData(FormPlotAlpha_PE.ComboPlotAlpha.ListIndex)
emitter% = (itemp% / MAXCHAN%)
absorber% = itemp% - emitter% * MAXCHAN%

' Save current ZAF and MAC selection
tzaftype% = izaf%
tmactype% = MACTypeFlag%

' Load number of data sets (number of binaries)
nsets% = 1
If FormPlotAlpha_PE.CheckAllOptions.Value = vbChecked Then nsets% = MAXZAF%

If FormPlotAlpha_PE.CheckAllMacs.Value = vbChecked Then
nsets% = MAXMACTYPE%
For k% = 1 To MAXMACTYPE%
MACFile$ = ApplicationCommonAppData$ & macstring2$(k%) & ".DAT"
If Dir$(MACFile$) = vbNullString Then nsets% = nsets% - 1
Next k%
End If

' Load number of sets
FormPlotAlpha_PE.Pesgo1.Subsets = nsets%

' Set symbols for each data set (use solid symbols only)
Call MiscPlotGetSymbols_PE(nsets%, FormPlotAlpha_PE.Pesgo1)
If ierror Then Exit Sub

' Add legend text
For k% = 1 To nsets%
If FormPlotAlpha_PE.CheckAllOptions.Value = vbChecked Then
FormPlotAlpha_PE.Pesgo1.SubsetLabels(k% - 1) = zafstring2$(k%)
FormPlotAlpha_PE.Pesgo1.SubsetPointTypes(k% - 1) = k%
End If
If FormPlotAlpha_PE.CheckAllMacs.Value = vbChecked Then
MACFile$ = ApplicationCommonAppData$ & macstring2$(k%) & ".DAT"
If Dir$(MACFile$) <> vbNullString Then
FormPlotAlpha_PE.Pesgo1.SubsetLabels(k% - 1) = macstring2$(k%)
End If
End If
Next k%

FormPlotAlpha_PE.Pesgo1.XAxisLabel = "Weight Fraction of Emitter"
FormPlotAlpha_PE.Pesgo1.YAxisLabel = "Elemental Alpha Factor (C/K - C)/(1 - C)"

' Calculate alpha-factors
astring$ = MiscAutoUcase$(CalcZAFTmpSample(1).Elsyup$(emitter%)) & " " & CalcZAFTmpSample(1).Xrsyms$(emitter%) & " in " & MiscAutoUcase$(CalcZAFTmpSample(1).Elsyup$(absorber%))
astring$ = astring$ & ", TO=" & Str$(CalcZAFTmpSample(1).takeoff!) & ", KeV=" & Str$(CalcZAFTmpSample(1).kilovolts!)
FormPlotAlpha_PE.Pesgo1.MainTitle = astring$

' Start loop
For k% = 1 To nsets%
If FormPlotAlpha_PE.CheckAllOptions.Value = vbChecked Then
izaf% = k%
Call InitGetZAFSetZAF2(k%)
If ierror Then Exit Sub
End If

If FormPlotAlpha_PE.CheckAllMacs.Value = vbChecked Then
MACFile$ = ApplicationCommonAppData$ & macstring2$(k%) & ".DAT"
If Dir$(MACFile$) = vbNullString Then
msg$ = "File " & MACFile$ & " was not found, therefore the calculation will be skipped..."
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
GoTo CalcZAFPlotAlphaFactors_PESkip
End If
Call GetZAFAllSaveMAC2(k%)
If ierror Then Exit Sub
MACTypeFlag% = k%       ' set after check for exist
End If

' Calculate the binary
Call AFactorCalculateKFactors(emitter%, absorber%, CalcZAFAnalysis, CalcZAFTmpSample())
If ierror Then Exit Sub

' Return the plot data (always return first emitter of binary only for plotting)
Call AFactorReturnAFactors(Int(1), npts%, xdata!(), ydata!(), acoeff!(), stddev!)
If ierror Then Exit Sub

If FormPlotAlpha_PE.CheckAllOptions.Value = vbUnchecked And FormPlotAlpha_PE.CheckAllMacs.Value = vbUnchecked Then
FormPlotAlpha_PE.LabelStdDev.Caption = MiscAutoFormat$(stddev!)
Else
FormPlotAlpha_PE.LabelStdDev.Caption = vbNullString
End If

' Plot alpha factors
If npts% < 1 Then GoTo CalcZAFPlotAlphaFactors_PENoPoints

' Display plot and fit
FormPlotAlpha_PE.Pesgo1.points = npts%

' Load y axis data (alpha)
For i% = 1 To npts%
FormPlotAlpha_PE.Pesgo1.ydata(k% - 1, i% - 1) = ydata!(i%)
Next i%

For i% = 1 To npts%
FormPlotAlpha_PE.Pesgo1.xdata(k% - 1, i% - 1) = xdata!(i%)
Next i%

CalcZAFPlotAlphaFactors_PESkip:
If FormPlotAlpha_PE.CheckAllMacs.Value = vbUnchecked And FormPlotAlpha_PE.CheckAllOptions.Value = vbUnchecked Then Exit For
Next k%

' Load caption
If CorrectionFlag% = 1 Then astring$ = "CONSTANT Alpha Factors"
If CorrectionFlag% = 2 Then astring$ = "LINEAR Alpha Factors"
If CorrectionFlag% = 3 Then astring$ = "POLYNOMIAL Alpha Factors"
If CorrectionFlag% = 4 Then astring$ = "NON-LINEAR Alpha Factors"
If FormPlotAlpha_PE.CheckAllOptions.Value = vbUnchecked And FormPlotAlpha_PE.CheckAllMacs.Value = vbUnchecked Then
astring$ = astring$ & " derived from k-ratios using: " & zafstring$(izaf%) & vbCrLf & "MAC Table: " & macstring$(MACTypeFlag%)
ElseIf FormPlotAlpha_PE.CheckAllOptions.Value = vbUnchecked And FormPlotAlpha_PE.CheckAllMacs.Value = vbChecked Then
astring$ = astring$ & " derived from k-ratios using: " & zafstring$(izaf%)
ElseIf FormPlotAlpha_PE.CheckAllOptions.Value = vbChecked And FormPlotAlpha_PE.CheckAllMacs.Value = vbUnchecked Then
astring$ = astring$ & " derived from k-ratios using: " & vbCrLf & "MAC Table: " & macstring$(MACTypeFlag%)
End If

' If using Penepma k-ratios (1 = no, 2 = yes)
If UsePenepmaKratiosFlag = 2 Then
If Not UsePenepmaKratiosLimitFlag Then
astring$ = astring$ & vbCrLf & " Using Penepma k-ratios if available..."
Else
astring$ = astring$ & "  Using Penepma k-ratios if available...(" & Format$(PenepmaKratiosLimitValue!) & " % limit)"
End If
End If
FormPlotAlpha_PE.LabelMatrixCorrection.Caption = astring$

' Restore current ZAF and MAC selection
izaf% = tzaftype%
Call InitGetZAFSetZAF2(izaf%)
If ierror Then Exit Sub
MACTypeFlag% = tmactype%
Call GetZAFAllSaveMAC2(MACTypeFlag%)
If ierror Then Exit Sub

'FormPlotAlpha_PE.Pesgo1.LegendStyle = PELS_1_LINE_INSIDE_OVERLAP&
'FormPlotAlpha_PE.Pesgo1.LegendStyle = PELS_1_LINE_INSIDE_AXIS&
'FormPlotAlpha_PE.Pesgo1.LegendLocation = PELL_TOP&
'FormPlotAlpha_PE.Pesgo1.LegendLocation = PELL_BOTTOM&
'FormPlotAlpha_PE.Pesgo1.LegendLocation = PELL_LEFT&
FormPlotAlpha_PE.Pesgo1.LegendLocation = PELL_RIGHT&
FormPlotAlpha_PE.Pesgo1.OneLegendPerLine = True                               ' put one legend per line
FormPlotAlpha_PE.Pesgo1.SimpleLineLegend = True
FormPlotAlpha_PE.Pesgo1.SimplePointLegend = True                              ' default = False encloses in a box

FormPlotAlpha_PE.Pesgo1.PEactions = REINITIALIZE_RESETIMAGE&

FormPlotAlpha_PE.Pesgo1.ManualScaleControlX = PEMSC_MINMAX&         ' manually control x axis
FormPlotAlpha_PE.Pesgo1.ManualScaleControlY = PEMSC_NONE&           ' autoscale y axis
FormPlotAlpha_PE.Pesgo1.ManualMinX = 0.01
FormPlotAlpha_PE.Pesgo1.ManualMaxX = 0.99

FormPlotAlpha_PE.Pesgo1.GraphAnnotationX(-1) = 0                    ' empty annotation array
FormPlotAlpha_PE.Pesgo1.GraphAnnotationY(-1) = 0

' Plot regression fit line
If FormPlotAlpha_PE.CheckAllOptions.Value = vbUnchecked And FormPlotAlpha_PE.CheckAllMacs.Value = vbUnchecked Then
Call CalcZAFPlotAlphaFit_PE(CorrectionFlag%, FormPlotAlpha_PE)
If ierror Then Exit Sub
End If

Call IOStatusAuto(vbNullString)
Exit Sub

' Errors
CalcZAFPlotAlphaFactors_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFPlotAlphaFactors_PE"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

CalcZAFPlotAlphaFactors_PENoPoints:
msg$ = "No alpha factors to plot for the current sample"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFPlotAlphaFactors_PE"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub CalcZAFAlphaExportData_PE(tForm As Form)
' Export alpha factor data (Pro Essentials graphing code)

ierror = False
On Error GoTo CalcZAFAlphaExportData_PEError

Dim j As Integer
Dim tfilename As String

' Load set data strings
If FormPlotAlpha_PE.CheckAllOptions.Value = vbChecked Then
ReDim sString(1 To MAXZAF%) As String
For j% = 1 To MAXZAF%
sString$(j%) = zafstring2$(j%)
Next j%
End If

If FormPlotAlpha_PE.CheckAllMacs.Value = vbChecked Then
ReDim sString(1 To MAXMACTYPE%) As String
For j% = 1 To MAXMACTYPE%
sString$(j%) = macstring2$(j%)
Next j%
End If

If FormPlotAlpha_PE.OptionBenceAlbee(0).Value Then tfilename$ = "Alpha-factors, Constant"
If FormPlotAlpha_PE.OptionBenceAlbee(1).Value Then tfilename$ = "Alpha-factors, Linear"
If FormPlotAlpha_PE.OptionBenceAlbee(2).Value Then tfilename$ = "Alpha-factors, Polynomial"
If FormPlotAlpha_PE.CheckAllOptions.Value = vbChecked Then tfilename$ = tfilename$ & ", AllZAFs"
If FormPlotAlpha_PE.CheckAllMacs.Value = vbChecked Then tfilename$ = tfilename$ & ", AllMACs"

Call MiscSaveDataSets_PE(tfilename$, FormPlotAlpha_PE.Pesgo1.MainTitle, FormPlotAlpha_PE.Pesgo1.XAxisLabel, FormPlotAlpha_PE.Pesgo1.YAxisLabel, sString$(), FormPlotAlpha_PE.Pesgo1, tForm)
If ierror Then Exit Sub

Exit Sub

' Errors
CalcZAFAlphaExportData_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFAlphaExportData_PE"
ierror = True
Exit Sub

End Sub

Sub CalcZAFPlotAlphaFit_PE(tCorrectionFlag As Integer, tForm As Form)
' Display the regression fit on the passed form

ierror = False
On Error GoTo CalcZAFPlotAlphaFit_PEError

Dim i As Integer
Dim linecount As Long

Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double
Dim sxmin As Double, sxmax As Double, symin As Double, symax As Double

Dim npts As Integer
Dim xdata() As Single, ydata() As Single, acoeff() As Single, stddev As Single

Const MAXSEGMENTS% = 400

' Return the plot data (always return first emitter of binary only for plotting)
Call AFactorReturnAFactors(Int(1), npts%, xdata!(), ydata!(), acoeff!(), stddev!)
If ierror Then Exit Sub

' Determine min and max of graph (in user data units)
xmin# = tForm.Pesgo1.ManualMinX
xmax# = tForm.Pesgo1.ManualMaxX
ymin# = tForm.Pesgo1.ManualMinY
ymax# = tForm.Pesgo1.ManualMaxY

' Calculate line to draw based on fit coefficients
sxmax# = xmin#
For i% = 1 To MAXSEGMENTS%

' Calculate partial line segments for x and y
sxmin# = sxmax#
sxmax# = sxmin# + (xmax# - xmin#) / (MAXSEGMENTS% - 1)
If sxmin# > 0# Then

' Constant fit (assume 50:50 composition only)
If tCorrectionFlag% = 1 Then
symin# = FormPlotAlpha_PE.Pesgo1.ydata(0, 5)            ' use mid point
symax# = FormPlotAlpha_PE.Pesgo1.ydata(0, 5)

' Linear fit
ElseIf tCorrectionFlag% = 2 Then
symin# = CDbl(acoeff!(1) + sxmin# * acoeff!(2))
symax# = CDbl(acoeff!(1) + sxmax# * acoeff!(2))

' Polynomial fit
ElseIf tCorrectionFlag% = 3 Then
symin# = CDbl(acoeff!(1) + sxmin# * acoeff!(2) + sxmin# ^ 2 * acoeff!(3))
symax# = CDbl(acoeff!(1) + sxmax# * acoeff!(2) + sxmax# ^ 2 * acoeff!(3))

' Non-linear fit
ElseIf tCorrectionFlag% = 4 Then
symin# = CDbl(acoeff!(1) + sxmin# * acoeff!(2) + sxmin# ^ 2 * acoeff!(3) + Exp(sxmin#) * acoeff!(4))
symax# = CDbl(acoeff!(1) + sxmax# * acoeff!(2) + sxmax# ^ 2 * acoeff!(3) + Exp(sxmax#) * acoeff!(4))
End If

' Clip
If symin# < ymin# Then symin# = ymin#
If symax# > ymax# Then symax# = ymax#

If symin# > ymax# Then symin# = ymax#
If symax# < ymin# Then symax# = ymin#

If i% = 1 Then
Call ScanDataPlotLine(tForm.Pesgo1, linecount&, sxmin#, symin#, sxmax#, symax#, False, True, Int(255), Int(0), Int(0), Int(255))     ' blue
If ierror Then Exit Sub
Else
Call ScanDataPlotLine(tForm.Pesgo1, linecount&, sxmin#, symin#, sxmax#, symax#, True, True, Int(255), Int(0), Int(0), Int(255))      ' blue
If ierror Then Exit Sub
End If

End If
Next i%

Exit Sub

' Errors
CalcZAFPlotAlphaFit_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFPlotAlphaFit_PE"
ierror = True
Exit Sub

CalcZAFPlotAlphaFit_PEZeroData:
msg$ = "Fit data contains zero values"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFPlotAlphaFit_PE"
ierror = True
Exit Sub

End Sub
