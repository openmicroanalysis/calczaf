Attribute VB_Name = "CodeEDSPlot"
' (c) Copyright 1995-2024 by John J. Donovan
Option Explicit

Dim TotalEnergyRange As Single

Dim GraphMinX As Single, GraphMinY As Single
Dim GraphMaxX As Single, GraphMaxY As Single

Sub EDSDisplayRedraw_PE(tForm As Form)
' Redraw the display, if graph has data (Pro Essentials)

ierror = False
On Error GoTo EDSDisplayRedraw_PEError

tForm.Pesgo1.GraphAnnotationX(-1) = 0               ' empty annotation array
tForm.Pesgo1.GraphAnnotationY(-1) = 0

Call EDSPlotKLM1(tForm)
If ierror Then Exit Sub

tForm.Pesgo1.PEactions = REINITIALIZE_RESETIMAGE
Exit Sub

' Errors
EDSDisplayRedraw_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSDisplayRedraw_PE"
ierror = True
Exit Sub

End Sub

Sub EDSDisplaySpectra_PE(tForm As Form, datarow As Integer, sample() As TypeSample)
' Display current spectrum from the interface (Pro Essentials)

ierror = False
On Error GoTo EDSDisplaySpectra_PEError

Dim i As Integer
Dim xtemp As Single, ytemp As Single, EDSeVmaxdata As Single

If EDSIntensityOption% = 0 Then
tForm.Pesgo1.YAxisLabel = "Intensity"           ' Axis labels
Else
tForm.Pesgo1.YAxisLabel = "Intensity (cps)"     ' Axis labels
End If

' Define #subset and #points
tForm.Pesgo1.Subsets = 1
tForm.Pesgo1.SubsetColors(0) = tForm.Pesgo1.PEargb(Int(255), Int(255), Int(0), Int(0))             ' red

tForm.Pesgo1.points = sample(1).EDSSpectraNumberofChannels%(datarow%)

' Load y axis data subset 0 - eds data
For i% = 1 To sample(1).EDSSpectraNumberofChannels%(datarow%)
If EDSIntensityOption% = 0 Then
ytemp! = sample(1).EDSSpectraIntensities&(datarow%, i%)                                            ' raw counts
Else
If sample(1).EDSSpectraLiveTime!(datarow%) <> 0# Then ytemp! = sample(1).EDSSpectraIntensities&(datarow%, i%) / sample(1).EDSSpectraLiveTime!(datarow%)      ' cps
End If

tForm.Pesgo1.ydata(0, i% - 1) = ytemp!

' Load x axis data
xtemp! = sample(1).EDSSpectraEVPerChannel!(datarow%) * (i% - 1) / EVPERKEV#
If EDSSpectraInterfaceTypeStored% = 2 Then
xtemp! = xtemp! + sample(1).EDSSpectraStartEnergy!(datarow%)                                    ' Bruker zero spectrum starts at plus start energy in keV
End If
If EDSSpectraInterfaceTypeStored% = 5 Then
xtemp! = xtemp! + 0.5 * sample(1).EDSSpectraEVPerChannel!(datarow%) / EVPERKEV#                 ' Thermo starts at zero to eV per channel
End If

tForm.Pesgo1.xdata(0, i% - 1) = xtemp!

' Find max eV channel that contains ydata, also = Duane-Hunt limit
If xtemp! * ytemp! > 0 Then EDSeVmaxdata! = xtemp!                              ' could use this to scale x if max eV with y data does not equal EDSSpectraAcceleratingVoltage!(datarow%)

If VerboseMode And DebugMode Then
If EDSIntensityOption% = 0 Then
Call IOWriteLog("EDS Point" & Str$(i%) & ", keV" & Str$(xtemp!) & ", counts" & Str$(ytemp!))      ' raw counts
Else
Call IOWriteLog("EDS Point" & Str$(i%) & ", keV" & Str$(xtemp!) & ", cps" & Str$(ytemp!))         ' cps
End If
End If
Next i%

' Define axis and graph properties and X extent
tForm.Pesgo1.ManualScaleControlX = PEMSC_MINMAX                                 ' Manually Control X Axis
tForm.Pesgo1.ManualMinX = sample(1).EDSSpectraStartEnergy!(datarow%)
'tForm.Pesgo1.ManualMaxX = sample(1).EDSSpectraEndEnergy!(datarow%)
'tForm.Pesgo1.ManualMaxX = EDSeVmaxdata!                                        ' max x axis is defined as last eV channel with Y data
tForm.Pesgo1.ManualMaxX = sample(1).EDSSpectraAcceleratingVoltage!(datarow%)

TotalEnergyRange! = (sample(1).EDSSpectraEndEnergy!(datarow%) - sample(1).EDSSpectraStartEnergy!(datarow%))

' Define Y extent
tForm.Pesgo1.ManualScaleControlY = PEMSC_MIN                                    ' autoscale Control Y Axis max, Manual Control min
tForm.Pesgo1.ManualMinY = 0

If EDSIntensityOption% = 0 Then
tForm.Pesgo1.ManualMaxY = sample(1).EDSSpectraMaxCounts&(datarow%)              ' raw counts
Else
If sample(1).EDSSpectraLiveTime!(datarow%) <> 0# Then tForm.Pesgo1.ManualMaxY = sample(1).EDSSpectraMaxCounts&(datarow%) / sample(1).EDSSpectraLiveTime!(datarow%)        ' cps
End If

If VerboseMode And DebugMode Then
msg$ = "EDS Display: "
msg$ = msg$ & "Start keV= " & Format$(sample(1).EDSSpectraStartEnergy!(datarow%)) & ", "
msg$ = msg$ & "Stop keV= " & Format$(sample(1).EDSSpectraEndEnergy!(datarow%)) & ", "
msg$ = msg$ & "numChan= " & Format$(sample(1).EDSSpectraNumberofChannels%(datarow%)) & ", "
msg$ = msg$ & "MaxInt= " & Format$(sample(1).EDSSpectraMaxCounts&(datarow%))
Call IOWriteLog(msg$)
End If

' Debug output
If VerboseMode Then
GraphMinX! = tForm.Pesgo1.ManualMinX
GraphMinY! = tForm.Pesgo1.ManualMinY
GraphMaxX! = tForm.Pesgo1.ManualMaxX
GraphMaxY! = tForm.Pesgo1.ManualMaxY
Call IOWriteLog("EDS Display Spectra: X Min/Max" & Str$(GraphMinX!) & "/" & Str$(GraphMaxX!) & ", Y Min/Max" & Str$(GraphMinY!) & "/" & Str$(GraphMaxY!))
End If

tForm.Pesgo1.PEactions = REINITIALIZE_RESETIMAGE    ' generate new plot
Exit Sub

' Errors
EDSDisplaySpectra_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSDisplaySpectra_PE"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub EDSInitDisplay_PE(tForm As Form, tCaption As String)
' Init spectrum display (Pro Essentials)

ierror = False
On Error GoTo EDSInitDisplay_PEError

' Init graph properties
Call MiscPlotInit(tForm.Pesgo1, True)
If ierror Then Exit Sub

' Plot type
tForm.Pesgo1.PlottingMethod = SGPM_BAR&         ' bargraph subset
'tForm.Pesgo1.BarWidth(0) = 1                   ' 0 equals auto width
tForm.Pesgo1.AdjoinBars = True                  ' bars full bin width

tForm.Pesgo1.ShowTickMarkY = PESTM_TICKS_HIDE&
tForm.Pesgo1.ShowTickMarkX = PESTM_TICKS_OUTSIDE&

' Annotation properties
tForm.Pesgo1.GraphAnnotationTextSize = 75               ' define annotation text size
tForm.Pesgo1.LabelFont = "Arial"                        ' define Font for annotations (and axes)
tForm.Pesgo1.HideIntersectingText = PEHIT_NO_HIDING&    ' or PEHIT_HIDE&

' Title Properties
tForm.Pesgo1.ImageAdjustTop = 50                ' add space above title formatting

tForm.Pesgo1.AnnotationsInFront = True          ' for KLM markers

tForm.Pesgo1.XAxisLabel = "keV"
If EDSIntensityOption% = 0 Then
tForm.Pesgo1.YAxisLabel = "Intensity"           ' y axis label
Else
tForm.Pesgo1.YAxisLabel = "Intensity (cps)"     ' y axis label
End If

tForm.TextStartkeV.Text = Format$(0)
tForm.TextStopkeV.Text = Format$(20)

Exit Sub

' Errors
EDSInitDisplay_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSInitDisplay_PE"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub EDSPlotKLM2_PE(tForm As Form, num As Long, sarray() As String, xarray() As String, iarray() As Single, earray() As Single, EDSKLMCounter As Long)
' Plot KLM lines and annotations (Pro Essentials code)

ierror = False
On Error GoTo EDSPlotKLM2_PEError

Dim n As Long
Dim temp1 As Single, temp2 As Single, ydatamin As Single

' Load graph y axis extents only on first init or during acquisition
temp1! = tForm.Pesgo1.ManualMaxY - tForm.Pesgo1.ManualMinY
If temp1! = 0# Then Exit Sub

' Draw annotations for KLM markers
For n& = 1 To num&

' Start point
tForm.Pesgo1.ShowAnnotations = True
tForm.Pesgo1.GraphAnnotationX(EDSKLMCounter&) = earray!(n&)
tForm.Pesgo1.GraphAnnotationY(EDSKLMCounter&) = 0
tForm.Pesgo1.GraphAnnotationType(EDSKLMCounter&) = PEGAT_THIN_SOLIDLINE&
tForm.Pesgo1.GraphAnnotationColor(EDSKLMCounter&) = tForm.Pesgo1.PEargb(Int(255), Int(0), Int(0), Int(255))
EDSKLMCounter& = EDSKLMCounter& + 1

' End point
tForm.Pesgo1.GraphAnnotationX(EDSKLMCounter&) = earray!(n&)
tForm.Pesgo1.GraphAnnotationType(EDSKLMCounter&) = PEGAT_LINECONTINUE&
tForm.Pesgo1.GraphAnnotationColor(EDSKLMCounter&) = tForm.Pesgo1.PEargb(Int(255), Int(0), Int(0), Int(255))

' Calculate Y axis height for KLM markers
ydatamin! = tForm.Pesgo1.ManualMinY
temp2! = ydatamin! + temp1! * iarray!(n&) / 150# ' maximum intensity
tForm.Pesgo1.GraphAnnotationY(EDSKLMCounter&) = temp2!

' Horizontal KLM (start by displaying x-ray text only for markers with intensities greater than 10)
tForm.Pesgo1.GraphAnnotationText(EDSKLMCounter&) = Trim$(sarray$(n&))
If iarray!(n&) > 10# Then tForm.Pesgo1.GraphAnnotationText(EDSKLMCounter&) = Trim$(sarray$(n&)) & " " & Trim$(xarray$(n&))

' Vertical KLM (for vertical text, prefix string with "|e")
'tForm.Pesgo1.GraphAnnotationText(EDSKLMCounter&) = "|e" & Trim$(sArray$(n&))
'If iarray!(n&) > 10# Then tForm.Pesgo1.GraphAnnotationText(EDSKLMCounter&) = "|e" & Trim$(sArray$(n&)) & " " & Trim$(xarray$(n&))

tForm.Pesgo1.GraphAnnotationBold(EDSKLMCounter&) = False
EDSKLMCounter& = EDSKLMCounter& + 1
Next n&

Exit Sub

' Errors
EDSPlotKLM2_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSPlotKLM2_PE"
ierror = True
Exit Sub

End Sub

Sub EDSRePlotKLM_PE(tForm As Form, num As Long, sarray() As String, xarray() As String, EDSKLMCounter As Long)
' Re plot specified KLM lines and modify x-ray text based on current zoom

ierror = False
On Error GoTo EDSRePlotKLM_PEError

Dim n As Long

' Re-load KLM annotations with x-ray text if KLM annotation height is larger then 1/10 current Y zoom
EDSKLMCounter& = 0
For n& = 1 To num&
EDSKLMCounter& = EDSKLMCounter& + 1
tForm.Pesgo1.GraphAnnotationText(EDSKLMCounter&) = Trim$(sarray$(n&))
If tForm.Pesgo1.GraphAnnotationY(EDSKLMCounter&) > tForm.Pesgo1.ZoomMaxY / 10# Then tForm.Pesgo1.GraphAnnotationText(EDSKLMCounter&) = Trim$(sarray$(n&)) & " " & Trim$(xarray$(n&))
EDSKLMCounter& = EDSKLMCounter& + 1
Next n&

Exit Sub

' Errors
EDSRePlotKLM_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSRePlotKLM_PE"
ierror = True
Exit Sub

End Sub

