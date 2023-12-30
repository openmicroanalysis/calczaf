Attribute VB_Name = "CodePenepma12Plot"
' (c) Copyright 1995-2024 by John J. Donovan
' Written by Gareth Seward under contract for Probe Software
Option Explicit

Sub Penepma12PlotKRatios_PE(nPoints As Long, nsets As Long, MaterialMeasuredEnergy As Double, MaterialMeasuredElement As Integer, MaterialMeasuredXray As Integer, _
    yktotal() As Double, yctotal() As Double, yc_prix() As Double, ycb_only() As Double, yctotal_meas() As Double, xdist() As Double)

' Load the k-ratio graph (PE code)
ierror = False
On Error GoTo Penepma12PlotKRatios_PEError

Dim i As Integer, j As Integer
Dim astring As String, bstring As String, cstring As String

' Unzoom, in case already zoomed
FormPENEPMA12.Pesgo1.PEactions = UNDO_ZOOM&

' Allow auto scaling of axes from here on  (initial plot load has fixed axes to create generic blank plot)
FormPENEPMA12.Pesgo1.ManualScaleControlY = PEMSC_NONE&
FormPENEPMA12.Pesgo1.ManualScaleControlX = PEMSC_NONE&

' Check for data
If nPoints& <= 0 Then GoTo Penepma12PlotKRatios_PENoData

' Load total intensity, fluorescence only and total concentration
FormPENEPMA12.Pesgo1.Subsets = nsets&
FormPENEPMA12.Pesgo1.Points = nPoints&

' Load y1 axis data (total fluorescence), Pro Essentials array index from 0, PFE array index 1
For j% = 0 To nsets& - 1
For i% = 0 To nPoints& - 1
If j% = 0 Then FormPENEPMA12.Pesgo1.ydata(j%, i%) = yktotal#(i% + 1)  ' total intensity
If j% = 1 Then FormPENEPMA12.Pesgo1.ydata(j%, i%) = yctotal#(i% + 1)  ' total (calculated) concentration
If j% = 2 Then
If ParameterFileA$ = ParameterFileB$ Then
FormPENEPMA12.Pesgo1.ydata(j%, i%) = yc_prix#(i% + 1)  ' (calculated) concentration from primary x-rays only
Else
FormPENEPMA12.Pesgo1.ydata(j%, i%) = ycb_only#(i% + 1) ' (calculated) concentration from B only
End If
End If
If j% = 3 Then FormPENEPMA12.Pesgo1.ydata(j%, i%) = yctotal_meas#(i% + 1)   ' total ("measured") concentration
Next i%

' Load x axis data
For i% = 0 To nPoints& - 1
FormPENEPMA12.Pesgo1.xdata(j%, i%) = xdist#(i% + 1)
Next i%
Next j%

' Controls point or point plus line, etc
FormPENEPMA12.Pesgo1.PlottingMethod = SGPM_POINT&
FormPENEPMA12.Pesgo1.PointSize = PEPS_LARGE&
FormPENEPMA12.Pesgo1.MinimumPointSize = PEMPS_MEDIUM_LARGE&     ' helps readability if sizing

' Load graph title
If ParameterFileA$ <> ParameterFileB$ Then
astring$ = MiscGetFileNameNoExtension$(ParameterFileA$) & " adjacent to " & MiscGetFileNameNoExtension$(ParameterFileB$)
Else
astring$ = MiscGetFileNameNoExtension$(ParameterFileA$)
End If
bstring$ = MiscGetFileNameNoExtension$(ParameterFileBStd$)
cstring$ = Trim$(Symup$(MaterialMeasuredElement%)) & " " & Xraylo$(MaterialMeasuredXray%) & ", in " & astring$ & " (" & Format$(MaterialMeasuredEnergy#) & " keV, " & bstring$ & " std)"
FormPENEPMA12.Pesgo1.ImageAdjustTop = 100                   ' edit for apperance
'FormPENEPMA12.Pesgo1.ImageAdjustRight = 100
FormPENEPMA12.Pesgo1.MainTitle = cstring$

' Axis labels
FormPENEPMA12.Pesgo1.YAxisLabel = "K Ratio %, or Conc %"
FormPENEPMA12.Pesgo1.XAxisLabel = "Distance um"

' Place y axis on right to match exp
FormPENEPMA12.Pesgo1.YAxisOnRight = True

' Log scale or not
FormPENEPMA12.Pesgo1.YAxisScaleControl = PEAC_NORMAL&

' Subset labels
For j% = 0 To nsets& - 1
If j% = 0 Then
FormPENEPMA12.Pesgo1.SubsetLabels(j%) = "K-Ratio %"
FormPENEPMA12.Pesgo1.SubsetPointTypes(j%) = PEPT_UPTRIANGLESOLID&
FormPENEPMA12.Pesgo1.SubsetLineTypes(j%) = PELT_THIN_SOLID&
FormPENEPMA12.Pesgo1.SubsetColors(j%) = FormPENEPMA12.Pesgo1.PEargb(Int(255), Int(255), Int(0), Int(0))     ' red
End If
If j% = 1 Then
FormPENEPMA12.Pesgo1.SubsetLabels(j%) = "Calc. Wt.% (Ideal)"
FormPENEPMA12.Pesgo1.SubsetPointTypes(j%) = PEPT_DOTSOLID&
FormPENEPMA12.Pesgo1.SubsetLineTypes(j%) = PELT_THIN_SOLID&
FormPENEPMA12.Pesgo1.SubsetColors(j%) = FormPENEPMA12.Pesgo1.PEargb(Int(255), Int(0), Int(255), Int(0))     ' green
End If
If j% = 2 Then
If ParameterFileA$ = ParameterFileB$ Then
FormPENEPMA12.Pesgo1.SubsetLabels(j%) = "Primary Wt.% (w/o Fluor.)"
Else
FormPENEPMA12.Pesgo1.SubsetLabels(j%) = "Boundary Wt.% (from Mat B)"
End If
FormPENEPMA12.Pesgo1.SubsetPointTypes(j%) = PEPT_DOWNTRIANGLESOLID&
FormPENEPMA12.Pesgo1.SubsetLineTypes(j%) = PELT_THIN_SOLID&
FormPENEPMA12.Pesgo1.SubsetColors(j%) = FormPENEPMA12.Pesgo1.PEargb(Int(255), Int(0), Int(0), Int(255))     ' blue
End If
If j% = 3 Then
astring$ = "CalcZAF Wt.%"
astring$ = astring$ & " (" & zafstring$(izaf%) & ")"
FormPENEPMA12.Pesgo1.SubsetLabels(j%) = astring$
FormPENEPMA12.Pesgo1.SubsetPointTypes(j%) = PEPT_DIAMONDSOLID&
FormPENEPMA12.Pesgo1.SubsetLineTypes(j%) = PELT_THIN_SOLID&
FormPENEPMA12.Pesgo1.SubsetColors(j%) = FormPENEPMA12.Pesgo1.PEargb(Int(255), Int(255), Int(128), Int(0))   ' orange
End If
Next j%

' Legend
If ParameterFileA$ <> ParameterFileB$ Then
FormPENEPMA12.Pesgo1.LegendStyle = PELS_1_LINE_INSIDE_OVERLAP&                ' legend inside plot
Else
FormPENEPMA12.Pesgo1.LegendStyle = PELS_1_LINE&                               ' single line per legend
FormPENEPMA12.Pesgo1.LegendLocation = PELL_BOTTOM&                            ' 0 = PELL_TOP, 1 = PELL_BOTTOM
End If
FormPENEPMA12.Pesgo1.OneLegendPerLine = True                                  ' edit to put one legend per line
FormPENEPMA12.Pesgo1.SimpleLineLegend = True
FormPENEPMA12.Pesgo1.SimplePointLegend = True                                 ' default False encloses in a box

Call Penepma12PlotUpdate_PE
If ierror Then Exit Sub

Exit Sub

' Errors
Penepma12PlotKRatios_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12PlotKRatios_PE"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma12PlotKRatios_PENoData:
msg$ = "No data to plot for specified conditions and element"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12PlotKRatios_PE"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma12PlotLoad_PE()
' Load the k-ratio graph (Pro Essentials code)

ierror = False
On Error GoTo Penepma12PlotLoad_PEError

' Init graph properties
Call MiscPlotInit(FormPENEPMA12.Pesgo1, True)
If ierror Then Exit Sub

FormPENEPMA12.Pesgo1.ManualScaleControlY = PEMSC_MINMAX&        ' manually control y axis - this requires resetting to 'NONE" in the PlotKRatio code
FormPENEPMA12.Pesgo1.ManualMinY = 0
FormPENEPMA12.Pesgo1.ManualMaxY = 100
FormPENEPMA12.Pesgo1.ManualScaleControlX = PEMSC_MINMAX&        ' manually control x axis
FormPENEPMA12.Pesgo1.ManualMinX = -50
FormPENEPMA12.Pesgo1.ManualMaxX = 0

FormPENEPMA12.Pesgo1.YAxisLabel = "K Ratio %, or Conc %"        ' axis labels
FormPENEPMA12.Pesgo1.XAxisLabel = "Distance um"

FormPENEPMA12.Pesgo1.ImageAdjustRight = 100                     ' axis formatting
FormPENEPMA12.Pesgo1.YAxisOnRight = True

Exit Sub

' Errors
Penepma12PlotLoad_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12PlotLoad_PE"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma12PlotUpdate_PE()
' Update the plot style (Pro Essentials code)

ierror = False
On Error GoTo Penepma12PlotUpdate_PEError

Dim r As Long
Dim astring As String, bstring As String, cstring As String         ' for material a/b labels
Dim alen As Integer, blen As Integer, clen As Integer               ' max length of x string$

' With or w/o gridlines
If FormPENEPMA12.CheckUseGridLines.Value = vbChecked Then
FormPENEPMA12.Pesgo1.GridLineControl = PEGLC_BOTH&          ' x and y grid
FormPENEPMA12.Pesgo1.GridBands = True                       ' adds color banding on background
Else
FormPENEPMA12.Pesgo1.GridLineControl = PEGLC_NONE&
FormPENEPMA12.Pesgo1.GridBands = False                      ' removes colour banding on background
End If

' With or w/o log scale
If FormPENEPMA12.CheckUseLogScale.Value = vbChecked Then
FormPENEPMA12.Pesgo1.YAxisScaleControl = PEAC_LOG&
Else
FormPENEPMA12.Pesgo1.YAxisScaleControl = PEAC_NORMAL&
End If

' Boundary label a<->b string parse
alen% = 8                                           ' if Mat A string > A density not added
blen% = 8
clen% = 32                                          ' max character length for cstring$
If ParameterFileA$ <> ParameterFileB$ Then
astring$ = MiscGetFileNameNoExtension$(ParameterFileA$)
r& = InStr(ParameterFileA$, VbSpace)                ' number of characters before space
If r& > 0 Then
astring$ = Left$(astring$, r&)                      ' only extract text before space
Else
astring$ = Left$(astring$, MAXTITLELENGTH%)        ' extract text upto max allowed length
End If
If Len(astring$) < alen% Then
astring$ = Left$(astring$ & " (" & Format$(MaterialDensityA#, f52$) & ")", MAXTITLELENGTH%)    ' add density to label if label not too long
End If

bstring$ = MiscGetFileNameNoExtension$(ParameterFileB$)
r& = InStr(ParameterFileB$, VbSpace)
If r& > 0 Then
bstring$ = Left$(bstring$, r&)
Else
bstring$ = Left$(bstring$, MAXTITLELENGTH%)
End If
If Len(bstring$) < blen% Then
bstring$ = Left$(bstring$ & " (" & Format$(MaterialDensityB#, f52$) & ")", MAXTITLELENGTH%)        ' add density to label if label not too long
End If

cstring$ = astring$ & " <--> " & bstring$           ' c$ is final format of text displayed
If Len(cstring$) > clen% Then
cstring$ = " Mat. A <--> Mat. B "                   ' if too long then just mat a<-> mat b
End If
FormPENEPMA12.Pesgo1.MultiSubTitles(0) = VbSpace
FormPENEPMA12.Pesgo1.MultiSubTitles(1) = VbSpace                                                        ' create some space
FormPENEPMA12.Pesgo1.VertLineAnnotation(0) = 0                                                          ' vertical line at x = 0 as place holder for annotation
FormPENEPMA12.Pesgo1.VertLineAnnotationColor(0) = FormPENEPMA12.Pesgo1.PEargb(Int(255), Int(0), Int(0), Int(0))             ' line black
FormPENEPMA12.Pesgo1.VertLineAnnotationText(0) = "|H" & cstring$                                        ' center justification for VertLine

' Annotations properties
FormPENEPMA12.Pesgo1.AnnotationsInFront = True
FormPENEPMA12.Pesgo1.LineAnnotationTextSize = 75
FormPENEPMA12.Pesgo1.ShowAnnotations = True
FormPENEPMA12.Pesgo1.LeftMargin = bstring$

' Matrix calculation e.g. Boundary and incedent materials the same
Else
astring$ = MiscGetFileNameNoExtension$(ParameterFileA$)
r& = InStr(ParameterFileA$, VbSpace)
If r& > 0 Then
astring$ = Left$(astring$, r&)
Else
astring$ = Left$(astring$, MAXTITLELENGTH%)
End If
FormPENEPMA12.Pesgo1.ShowAnnotations = False
FormPENEPMA12.Pesgo1.MultiSubTitles(0) = " "
FormPENEPMA12.Pesgo1.MultiSubTitles(1) = "|" & astring$ & "|"
FormPENEPMA12.Pesgo1.FontSizeMSCntl = 0.8
End If

FormPENEPMA12.Pesgo1.PEactions = REINITIALIZE_RESETIMAGE&

Exit Sub

' Errors
Penepma12PlotUpdate_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12PlotUpdate_PE"
ierror = True
Exit Sub

End Sub
