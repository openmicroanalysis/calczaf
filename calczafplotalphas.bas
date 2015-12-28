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

Dim CalcZAFTmpSample(1 To 1) As TypeSample

Dim tzaftype As Integer, tmactype As Integer

Sub CalcZAFLoadAlphas_GS(sample() As TypeSample)
' Load binaries for the current sample for alpha factor plotting (Graphics Server graphing code)

ierror = False
On Error GoTo CalcZAFLoadAlphas_GSError

Dim i As Integer
Dim emitter As Integer, absorber As Integer
Dim inum As Integer
Dim astring As String

' Check for Bence-Albee corrections
If CorrectionFlag% < 1 Or CorrectionFlag% > 4 Then
msg$ = "Bence-Albee corrections are not selected. Changing matrix correction type to polynomial alpha-factors."
MsgBox msg$, vbOKOnly + vbInformation, "CalcZAFLoadAlphas_GS"
CorrectionFlag% = 3
End If

' Load Bence-Albee modes only
For i% = 0 To 3
FormPlotAlpha_GS.OptionBenceAlbee(i%).Caption = corstring(i% + 1)
If i% + 1 = CorrectionFlag% Then FormPlotAlpha_GS.OptionBenceAlbee(i%).Value = True
Next i%

' Calculate current sample
Call CalcZAFCalculate
If ierror Then Exit Sub

' Load to module level
CalcZAFTmpSample(1) = sample(1)

' Calculate each binary in sample
inum% = 0
FormPlotAlpha_GS.ComboPlotAlpha.Clear
For emitter% = 1 To sample(1).LastElm%
For absorber% = 1 To sample(1).LastChan%

' Skip if emitter and absorber are the same (duplicate elements)
If emitter% <> absorber% And sample(1).Elsyms$(emitter%) <> sample(1).Elsyms$(absorber%) Then
inum% = inum% + 1

astring$ = MiscAutoUcase$(sample(1).Elsyup$(emitter%)) & " " & sample(1).Xrsyms$(emitter%) & " in " & MiscAutoUcase$(sample(1).Elsyup$(absorber%))
FormPlotAlpha_GS.ComboPlotAlpha.AddItem astring$
FormPlotAlpha_GS.ComboPlotAlpha.ItemData(FormPlotAlpha_GS.ComboPlotAlpha.NewIndex) = emitter% * MAXCHAN% + absorber%
End If

Next absorber%
Next emitter%

' Check number of binaries
If inum% = 0 Then GoTo CalcZAFLoadAlphas_GSNoBinaries

' Check for Penepma k-ratios flag
If UsePenepmaKratiosFlag = 2 Then
FormPlotAlpha_GS.CheckAllOptions.Enabled = False
FormPlotAlpha_GS.CheckAllMacs.Enabled = False
Else
FormPlotAlpha_GS.CheckAllOptions.Enabled = True
FormPlotAlpha_GS.CheckAllMacs.Enabled = True
End If

' Click first binary
FormPlotAlpha_GS.ComboPlotAlpha.ListIndex = 0
FormPlotAlpha_GS.Show vbModal

Exit Sub

' Errors
CalcZAFLoadAlphas_GSError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFLoadAlphas_GS"
ierror = True
Exit Sub

CalcZAFLoadAlphas_GSNoBinaries:
msg$ = "No alpha factor binaries to plot for the current sample"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFLoadAlphas_GS"
ierror = True
Exit Sub

End Sub

Sub CalcZAFPlotAlphaFactors_GS()
' Plot alpha factors for the indicated binary (Graphics Server graphing code)

ierror = False
On Error GoTo CalcZAFPlotAlphaFactors_GSError

Dim itemp As Integer, i As Integer
Dim emitter As Integer, absorber As Integer, k As Integer
Dim astring As String
Dim xmin As Single, xmax As Single

Dim npts As Integer, nsets As Integer
Dim xdata() As Single, ydata() As Single, acoeff() As Single, stddev As Single

' Get the selected binary
If FormPlotAlpha_GS.ComboPlotAlpha.ListIndex < 0 Then Exit Sub
If FormPlotAlpha_GS.ComboPlotAlpha.ListCount < 1 Then Exit Sub

' Determine which binary to calculate
itemp% = FormPlotAlpha_GS.ComboPlotAlpha.ItemData(FormPlotAlpha_GS.ComboPlotAlpha.ListIndex)
emitter% = (itemp% / MAXCHAN%)
absorber% = itemp% - emitter% * MAXCHAN%

' Specify number of sets
FormPlotAlpha_GS.Graph1.DataReset = 9   ' reset all array based properties
FormPlotAlpha_GS.Graph1.DrawMode = 1

' Save current ZAF and MAC selection
tzaftype% = izaf%
tmactype% = MACTypeFlag%

nsets% = 1
If FormPlotAlpha_GS.CheckAllOptions.Value = vbChecked Then nsets% = MAXZAF%

If FormPlotAlpha_GS.CheckAllMacs.Value = vbChecked Then
nsets% = MAXMACTYPE%
For k% = 1 To MAXMACTYPE%
MACFile$ = ApplicationCommonAppData$ & macstring2$(k%) & ".DAT"
If Dir$(MACFile$) = vbNullString Then nsets% = nsets% - 1
Next k%
End If

' Load number of sets
FormPlotAlpha_GS.Graph1.NumSets = nsets%

' Add legend text
FormPlotAlpha_GS.Graph1.AutoInc = 0
If nsets% > 1 Then
FormPlotAlpha_GS.Graph1.AutoInc = 1
For k% = 1 To nsets%
If FormPlotAlpha_GS.CheckAllOptions.Value = vbChecked Then FormPlotAlpha_GS.Graph1.LegendText = zafstring2$(k%)
If FormPlotAlpha_GS.CheckAllMacs.Value = vbChecked Then
MACFile$ = ApplicationCommonAppData$ & macstring2$(k%) & ".DAT"
If Dir$(MACFile$) <> vbNullString Then
FormPlotAlpha_GS.Graph1.LegendText = macstring2$(k%)
End If
End If
Next k%

' Set symbols for each data set (use odd numbers from 3 to 13 for solid symbols)
Call MiscPlotGetSymbols_GS(nsets%, FormPlotAlpha_GS.Graph1)
If ierror Then Exit Sub
End If

xmin! = MAXMINIMUM!
xmax! = MAXMAXIMUM!
FormPlotAlpha_GS.Graph1.GraphType = 9      ' scatter graph
If nsets% = 1 Then
FormPlotAlpha_GS.Graph1.GraphStyle = 2     ' symbols only
Else
FormPlotAlpha_GS.Graph1.GraphStyle = 3     ' lines and symbols
End If
FormPlotAlpha_GS.Graph1.SymbolSize = 100   ' 100% of default
FormPlotAlpha_GS.Graph1.SymbolData = 13    ' solid circle
FormPlotAlpha_GS.Graph1.Background = 15    ' use white background

FormPlotAlpha_GS.Graph1.EBarSource = 0
FormPlotAlpha_GS.Graph1.AutoInc = 0  ' for loading multiple independent X values
FormPlotAlpha_GS.Graph1.XAxisStyle = 2   ' user defined
FormPlotAlpha_GS.Graph1.YAxisStyle = 1   ' variable origin
FormPlotAlpha_GS.Graph1.SDKMouse = 1

FormPlotAlpha_GS.Graph1.FSize(GSR_USETITG%) = GSR_SIZE_150%
FormPlotAlpha_GS.Graph1.BottomTitle = "Weight Fraction of Emitter"
FormPlotAlpha_GS.Graph1.LeftTitleStyle = 1
FormPlotAlpha_GS.Graph1.LeftTitle = "Elemental Alpha Factor (C/K - C)/(1 - C)"
FormPlotAlpha_GS.Graph1.FSize(GSR_USETITXY%) = -GSR_SIZE_100%
FormPlotAlpha_GS.Graph1.FSize(GSR_USELABS%) = GSR_SIZE_100%

FormPlotAlpha_GS.Graph1.LineStats = 0    ' no fit for multiple plots

' Calculate alpha-factors
astring$ = MiscAutoUcase$(CalcZAFTmpSample(1).Elsyup$(emitter%)) & " " & CalcZAFTmpSample(1).Xrsyms$(emitter%) & " in " & MiscAutoUcase$(CalcZAFTmpSample(1).Elsyup$(absorber%))
astring$ = astring$ & ", TO=" & Str$(CalcZAFTmpSample(1).takeoff!) & ", KeV=" & Str$(CalcZAFTmpSample(1).kilovolts!)
FormPlotAlpha_GS.Graph1.GraphTitle = astring$

' Start loop
For k% = 1 To nsets%
If FormPlotAlpha_GS.CheckAllOptions.Value = vbChecked Then
izaf% = k%
Call InitGetZAFSetZAF2(k%)
If ierror Then Exit Sub
End If

If FormPlotAlpha_GS.CheckAllMacs.Value = vbChecked Then
MACFile$ = ApplicationCommonAppData$ & macstring2$(k%) & ".DAT"
If Dir$(MACFile$) = vbNullString Then
msg$ = "File " & MACFile$ & " was not found, therefore the calculation will be skipped..."
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
GoTo CalcZAFPlotAlphaFactors_GSSkip
End If
Call GetZAFAllSaveMAC2(k%)
If ierror Then Exit Sub
MACTypeFlag% = k%       ' set after check for exist
End If

' Calculate the binary
Call AFactorCalculateKFactors(emitter%, absorber%, CalcZAFTmpSample())
If ierror Then Exit Sub

' Return the plot data (always return first emitter of binary only for plotting)
Call AFactorReturnAFactors(Int(1), npts%, xdata!(), ydata!(), acoeff!(), stddev!)
If ierror Then Exit Sub

If FormPlotAlpha_GS.CheckAllOptions.Value = vbUnchecked And FormPlotAlpha_GS.CheckAllMacs.Value = vbUnchecked Then
FormPlotAlpha_GS.LabelStdDev.Caption = Format$(stddev!)
Else
FormPlotAlpha_GS.LabelStdDev.Caption = vbNullString
End If

' Plot alpha factors
If npts% < 1 Then GoTo CalcZAFPlotAlphaFactors_GSNoPoints

' Display plot and fit
FormPlotAlpha_GS.Graph1.ThisSet = k%
FormPlotAlpha_GS.Graph1.NumPoints = npts%

' Load y axis data (alpha)
For i% = 1 To npts%
FormPlotAlpha_GS.Graph1.ThisPoint = i%
FormPlotAlpha_GS.Graph1.GraphData = ydata!(i%)
Next i%

' Load x axis data
For i% = 1 To npts%
FormPlotAlpha_GS.Graph1.ThisPoint = i%
If xdata!(i%) < xmin! Then xmin! = xdata!(i%)
If xdata!(i%) > xmax! Then xmax! = xdata!(i%)
FormPlotAlpha_GS.Graph1.XPosData = xdata!(i%)
Next i%

FormPlotAlpha_GS.Graph1.XAxisMin = xmin!
FormPlotAlpha_GS.Graph1.XAxisMax = xmax!
If Abs(xmin!) > 10# And CLng(xmin!) <> CLng(xmax!) Then FormPlotAlpha_GS.Graph1.XAxisMin = CLng(xmin!)
If Abs(xmax!) > 10# And CLng(xmin!) <> CLng(xmax!) Then FormPlotAlpha_GS.Graph1.XAxisMax = CLng(xmax!)

FormPlotAlpha_GS.Graph1.XAxisTicks = 10      ' (0-100)
FormPlotAlpha_GS.Graph1.XAxisMinorTicks = -1   ' 1 minor ticks per tick

FormPlotAlpha_GS.Graph1.YAxisTicks = 10      ' (0-100)
FormPlotAlpha_GS.Graph1.YAxisMinorTicks = -1   ' 1 minor ticks per tick

FormPlotAlpha_GS.Graph1.ThickLines = 1  ' turn on
FormPlotAlpha_GS.Graph1.SymbolSize = 100    ' 100% of default

CalcZAFPlotAlphaFactors_GSSkip:
If FormPlotAlpha_GS.CheckAllMacs.Value = vbUnchecked And FormPlotAlpha_GS.CheckAllOptions.Value = vbUnchecked Then Exit For
Next k%

' Show plot
FormPlotAlpha_GS.Graph1.SDKPaint = 1
FormPlotAlpha_GS.Graph1.DrawMode = 2

' Load caption
If CorrectionFlag% = 1 Then astring$ = "CONSTANT Alpha Factors"
If CorrectionFlag% = 2 Then astring$ = "LINEAR Alpha Factors"
If CorrectionFlag% = 3 Then astring$ = "POLYNOMIAL Alpha Factors"
If CorrectionFlag% = 4 Then astring$ = "NON-LINEAR Alpha Factors"
If FormPlotAlpha_GS.CheckAllOptions.Value = vbUnchecked And FormPlotAlpha_GS.CheckAllMacs.Value = vbUnchecked Then
astring$ = astring$ & " derived from k-ratios using: " & zafstring$(izaf%) & vbCrLf & "MAC Table: " & macstring$(MACTypeFlag%)
ElseIf FormPlotAlpha_GS.CheckAllOptions.Value = vbUnchecked And FormPlotAlpha_GS.CheckAllMacs.Value = vbChecked Then
astring$ = astring$ & " derived from k-ratios using: " & zafstring$(izaf%)
ElseIf FormPlotAlpha_GS.CheckAllOptions.Value = vbChecked And FormPlotAlpha_GS.CheckAllMacs.Value = vbUnchecked Then
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
FormPlotAlpha_GS.LabelMatrixCorrection.Caption = astring$

' Plot regression fit line
Call CalcZAFPlotAlphaFit(Int(0))
If ierror Then Exit Sub

' Restore current ZAF and MAC selection
izaf% = tzaftype%
Call InitGetZAFSetZAF2(izaf%)
If ierror Then Exit Sub
MACTypeFlag% = tmactype%
Call GetZAFAllSaveMAC2(MACTypeFlag%)
If ierror Then Exit Sub

Call IOStatusAuto(vbNullString)
Exit Sub

' Errors
CalcZAFPlotAlphaFactors_GSError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFPlotAlphaFactors_GS"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

CalcZAFPlotAlphaFactors_GSNoPoints:
msg$ = "No alpha factors to plot for the current sample"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFPlotAlphaFactors_GS"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub CalcZAFAlphaExportData_GS(tForm As Form)
' Export alpha factor data (Graphics Server graphing code)

ierror = False
On Error GoTo CalcZAFAlphaExportData_GSError

Dim j As Integer
Dim tfilename As String

' Load set data strings
If FormPlotAlpha_GS.CheckAllOptions.Value = vbChecked Then
ReDim sString(1 To MAXZAF%) As String
For j% = 1 To MAXZAF%
sString$(j%) = zafstring2$(j%)
Next j%
End If

If FormPlotAlpha_GS.CheckAllMacs.Value = vbChecked Then
ReDim sString(1 To MAXMACTYPE%) As String
For j% = 1 To MAXMACTYPE%
sString$(j%) = macstring2$(j%)
Next j%
End If

If FormPlotAlpha_GS.OptionBenceAlbee(0).Value Then tfilename$ = "Alpha-factors, Constant"
If FormPlotAlpha_GS.OptionBenceAlbee(1).Value Then tfilename$ = "Alpha-factors, Linear"
If FormPlotAlpha_GS.OptionBenceAlbee(2).Value Then tfilename$ = "Alpha-factors, Polynomial"
If FormPlotAlpha_GS.CheckAllOptions.Value = vbChecked Then tfilename$ = tfilename$ & ", AllZAFs"
If FormPlotAlpha_GS.CheckAllMacs.Value = vbChecked Then tfilename$ = tfilename$ & ", AllMACs"

Call MiscSaveDataSets(tfilename$, FormPlotAlpha_GS.Graph1.GraphTitle, FormPlotAlpha_GS.Graph1.BottomTitle, FormPlotAlpha_GS.Graph1.LeftTitle, sString$(), FormPlotAlpha_GS.Graph1, tForm)
If ierror Then Exit Sub

Exit Sub

' Errors
CalcZAFAlphaExportData_GSError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFAlphaExportData_GS"
ierror = True
Exit Sub

End Sub

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
Dim xmin As Single, xmax As Single

Dim npts As Integer, nsets As Integer
Dim xdata() As Single, ydata() As Single, acoeff() As Single, stddev As Single

' Get the selected binary
If FormPlotAlpha_PE.ComboPlotAlpha.ListIndex < 0 Then Exit Sub
If FormPlotAlpha_PE.ComboPlotAlpha.ListCount < 1 Then Exit Sub

' Determine which binary to calculate
itemp% = FormPlotAlpha_PE.ComboPlotAlpha.ItemData(FormPlotAlpha_PE.ComboPlotAlpha.ListIndex)
emitter% = (itemp% / MAXCHAN%)
absorber% = itemp% - emitter% * MAXCHAN%

' Specify number of sets
FormPlotAlpha_PE.Graph1.DataReset = 9   ' reset all array based properties
FormPlotAlpha_PE.Graph1.DrawMode = 1

' Save current ZAF and MAC selection
tzaftype% = izaf%
tmactype% = MACTypeFlag%

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
FormPlotAlpha_PE.Graph1.NumSets = nsets%

' Add legend text
FormPlotAlpha_PE.Graph1.AutoInc = 0
If nsets% > 1 Then
FormPlotAlpha_PE.Graph1.AutoInc = 1
For k% = 1 To nsets%
If FormPlotAlpha_PE.CheckAllOptions.Value = vbChecked Then FormPlotAlpha_PE.Graph1.LegendText = zafstring2$(k%)
If FormPlotAlpha_PE.CheckAllMacs.Value = vbChecked Then
MACFile$ = ApplicationCommonAppData$ & macstring2$(k%) & ".DAT"
If Dir$(MACFile$) <> vbNullString Then
FormPlotAlpha_PE.Graph1.LegendText = macstring2$(k%)
End If
End If
Next k%

' Set symbols for each data set
Call MiscPlotGetSymbols_PE(nsets%, FormPlotAlpha_PE.Pesgo1)
If ierror Then Exit Sub
End If

xmin! = MAXMINIMUM!
xmax! = MAXMAXIMUM!
FormPlotAlpha_GS.Graph1.GraphType = 9      ' scatter graph
If nsets% = 1 Then
FormPlotAlpha_GS.Graph1.GraphStyle = 2     ' symbols only
Else
FormPlotAlpha_GS.Graph1.GraphStyle = 3     ' lines and symbols
End If
FormPlotAlpha_PE.Graph1.SymbolSize = 100    ' 100% of default
FormPlotAlpha_PE.Graph1.SymbolData = 13     ' solid circle
FormPlotAlpha_PE.Graph1.Background = 15     ' use white background

FormPlotAlpha_PE.Graph1.EBarSource = 0
FormPlotAlpha_PE.Graph1.AutoInc = 0      ' for loading multiple independent X values
FormPlotAlpha_PE.Graph1.XAxisStyle = 2   ' user defined
FormPlotAlpha_PE.Graph1.YAxisStyle = 1   ' variable origin
FormPlotAlpha_PE.Graph1.SDKMouse = 1

FormPlotAlpha_PE.Graph1.FSize(GSR_USETITG%) = GSR_SIZE_150%
FormPlotAlpha_PE.Graph1.BottomTitle = "Weight Fraction of Emitter"
FormPlotAlpha_PE.Graph1.LeftTitleStyle = 1
FormPlotAlpha_PE.Graph1.LeftTitle = "Elemental Alpha Factor (C/K - C)/(1 - C)"
FormPlotAlpha_PE.Graph1.FSize(GSR_USETITXY%) = -GSR_SIZE_100%
FormPlotAlpha_PE.Graph1.FSize(GSR_USELABS%) = GSR_SIZE_100%

FormPlotAlpha_PE.Graph1.LineStats = 0    ' no fit for multiple plots

' Calculate alpha-factors
astring$ = MiscAutoUcase$(CalcZAFTmpSample(1).Elsyup$(emitter%)) & " " & CalcZAFTmpSample(1).Xrsyms$(emitter%) & " in " & MiscAutoUcase$(CalcZAFTmpSample(1).Elsyup$(absorber%))
astring$ = astring$ & ", TO=" & Str$(CalcZAFTmpSample(1).takeoff!) & ", KeV=" & Str$(CalcZAFTmpSample(1).kilovolts!)
FormPlotAlpha_PE.Graph1.GraphTitle = astring$

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
Call AFactorCalculateKFactors(emitter%, absorber%, CalcZAFTmpSample())
If ierror Then Exit Sub

' Return the plot data (always return first emitter of binary only for plotting)
Call AFactorReturnAFactors(Int(1), npts%, xdata!(), ydata!(), acoeff!(), stddev!)
If ierror Then Exit Sub

If FormPlotAlpha_PE.CheckAllOptions.Value = vbUnchecked And FormPlotAlpha_PE.CheckAllMacs.Value = vbUnchecked Then
FormPlotAlpha_PE.LabelStdDev.Caption = Format$(stddev!)
Else
FormPlotAlpha_PE.LabelStdDev.Caption = vbNullString
End If

' Plot alpha factors
If npts% < 1 Then GoTo CalcZAFPlotAlphaFactors_PENoPoints

' Display plot and fit
FormPlotAlpha_PE.Graph1.ThisSet = k%
FormPlotAlpha_PE.Graph1.NumPoints = npts%

' Load y axis data (alpha)
For i% = 1 To npts%
FormPlotAlpha_PE.Graph1.ThisPoint = i%
FormPlotAlpha_PE.Graph1.GraphData = ydata!(i%)
Next i%

' Load x axis data
For i% = 1 To npts%
FormPlotAlpha_PE.Graph1.ThisPoint = i%
If xdata!(i%) < xmin! Then xmin! = xdata!(i%)
If xdata!(i%) > xmax! Then xmax! = xdata!(i%)
FormPlotAlpha_PE.Graph1.XPosData = xdata!(i%)
Next i%

FormPlotAlpha_PE.Graph1.XAxisMin = xmin!
FormPlotAlpha_PE.Graph1.XAxisMax = xmax!
If Abs(xmin!) > 10# And CLng(xmin!) <> CLng(xmax!) Then FormPlotAlpha_PE.Graph1.XAxisMin = CLng(xmin!)
If Abs(xmax!) > 10# And CLng(xmin!) <> CLng(xmax!) Then FormPlotAlpha_PE.Graph1.XAxisMax = CLng(xmax!)

FormPlotAlpha_PE.Graph1.XAxisTicks = 10      ' (0-100)
FormPlotAlpha_PE.Graph1.XAxisMinorTicks = -1   ' 1 minor ticks per tick

FormPlotAlpha_PE.Graph1.YAxisTicks = 10      ' (0-100)
FormPlotAlpha_PE.Graph1.YAxisMinorTicks = -1   ' 1 minor ticks per tick

FormPlotAlpha_PE.Graph1.ThickLines = 1  ' turn on
FormPlotAlpha_PE.Graph1.SymbolSize = 100    ' 100% of default

CalcZAFPlotAlphaFactors_PESkip:
If FormPlotAlpha_PE.CheckAllMacs.Value = vbUnchecked And FormPlotAlpha_PE.CheckAllOptions.Value = vbUnchecked Then Exit For
Next k%

' Show plot
FormPlotAlpha_PE.Graph1.SDKPaint = 1
FormPlotAlpha_PE.Graph1.DrawMode = 2

' Plot regression fit line
Call CalcZAFPlotAlphaFit(Int(1))
If ierror Then Exit Sub

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

Call MiscSaveDataSets(tfilename$, FormPlotAlpha_PE.Graph1.GraphTitle, FormPlotAlpha_PE.Graph1.BottomTitle, FormPlotAlpha_PE.Graph1.LeftTitle, sString$(), FormPlotAlpha_PE.Graph1, tForm)
If ierror Then Exit Sub

Exit Sub

' Errors
CalcZAFAlphaExportData_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFAlphaExportData_PE"
ierror = True
Exit Sub

End Sub

Sub CalcZAFPlotAlphaFit_GS(tCorrectionFlag As Integer, tForm As Form)
' Display the regression fit on the passed form

ierror = False
On Error GoTo CalcZAFPlotAlphaFit_GSError

Dim i As Integer
Dim r As Long

Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double
Dim sxmin As Double, sxmax As Double, symin As Double, symax As Double
Dim txmin As Double, txmax As Double, tymin As Double, tymax As Double

Dim npts As Integer
Dim xdata() As Single, ydata() As Single, acoeff() As Single, stddev As Single

Const MAXSEGMENTS% = 100

' Return the plot data (always return first emitter of binary only for plotting)
Call AFactorReturnAFactors(Int(1), npts%, xdata!(), ydata!(), acoeff!(), stddev!)
If ierror Then Exit Sub

' Determine min and max of graph (in user data units)
xmin# = tForm.Graph1.SDKInfo(2)
xmax# = tForm.Graph1.SDKInfo(1)

ymin# = tForm.Graph1.SDKInfo(4)
ymax# = tForm.Graph1.SDKInfo(3)

' Calculate line to draw based on fit coefficients
sxmax# = xmin#
For i% = 1 To MAXSEGMENTS%

' Calculate partial line segments for x
sxmin# = sxmax#
sxmax# = sxmin# + (xmax# - xmin#) / (MAXSEGMENTS% - 1)

' Calculate partial line segments for y
If sxmin# = 0# Then GoTo CalcZAFPlotAlphaFit_GSZeroData

' Constant fit (assume 50:50 composition only)
If tCorrectionFlag% = 1 Then
tForm.Graph1.ThisPoint = 6
symin# = CDbl(tForm.Graph1.GraphData)
symax# = CDbl(tForm.Graph1.GraphData)

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

' Convert to graph units
Call XrayPlotConvert(tForm, sxmin#, symin#, txmin#, tymin#)
If ierror Then Exit Sub

Call XrayPlotConvert(tForm, sxmax#, symax#, txmax#, tymax#)
If ierror Then Exit Sub

r& = GSLineAbs(txmin#, tymin#, txmax#, tymax#, 4, 2, 4)     ' draw thick red line
Next i%

Exit Sub

' Errors
CalcZAFPlotAlphaFit_GSError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFPlotAlphaFit_GS"
ierror = True
Exit Sub

CalcZAFPlotAlphaFit_GSZeroData:
msg$ = "Fit data contains zero values"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFPlotAlphaFit_GS"
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

Const MAXSEGMENTS% = 100

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

' Calculate partial line segments for x
sxmin# = sxmax#
sxmax# = sxmin# + (xmax# - xmin#) / (MAXSEGMENTS% - 1)

' Calculate partial line segments for y
If sxmin# = 0# Then GoTo CalcZAFPlotAlphaFit_PEZeroData

' Constant fit (assume 50:50 composition only)
If tCorrectionFlag% = 1 Then
symin# = CDbl(tForm.Pesgo1.ydata(0, 5))      ' use 50: 50 point
symax# = CDbl(tForm.Pesgo1.ydata(0, 5))      ' use 50: 50 point

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
symin# = CDbl(acoeff!(1) + sxmin# ^ 2 * acoeff!(2) + sxmin# ^ 2 * acoeff!(3) + Exp(sxmin#) * acoeff!(4))
symax# = CDbl(acoeff!(1) + sxmax# ^ 2 * acoeff!(2) + sxmax# ^ 2 * acoeff!(3) + Exp(sxmax#) * acoeff!(4))
End If

' Clip
If symin# < ymin# Then symin# = ymin#
If symax# > ymax# Then symax# = ymax#

If symin# > ymax# Then symin# = ymax#
If symax# < ymin# Then symax# = ymin#

If i% = 1 Then
Call ScanDataPlotLine(tForm.Pesgo1, linecount&, sxmin#, symin#, sxmax#, symax#, False, True, Int(255), Int(128), Int(0), Int(0))     ' brown
If ierror Then Exit Sub
Else
Call ScanDataPlotLine(tForm.Pesgo1, linecount&, sxmin#, symin#, sxmax#, symax#, True, True, Int(255), Int(128), Int(0), Int(0))      ' brown
If ierror Then Exit Sub
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

Sub CalcZAFPlotAlphaFit(mode As Integer)
' Plot the regression fit for both graphics packages
'  mode = 0 use GS code
'  mode = 1 use PE code

ierror = False
On Error GoTo CalcZAFPlotAlphaFitError

' Graphics Server graphics call
If mode% = 0 Then
If FormPlotAlpha_GS.CheckAllOptions.Value = vbUnchecked And FormPlotAlpha_GS.CheckAllMacs.Value = vbUnchecked Then
Call CalcZAFPlotAlphaFit_GS(CorrectionFlag%, FormPlotAlpha_GS)
If ierror Then Exit Sub
End If
End If

' Pro Essentials graphics call
If mode% = 1 Then
If FormPlotAlpha_PE.CheckAllOptions.Value = vbUnchecked And FormPlotAlpha_PE.CheckAllMacs.Value = vbUnchecked Then
Call CalcZAFPlotAlphaFit_PE(CorrectionFlag%, FormPlotAlpha_PE)
If ierror Then Exit Sub
End If
End If

Exit Sub

' Errors
CalcZAFPlotAlphaFitError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFPlotAlphaFit"
ierror = True
Exit Sub

End Sub
