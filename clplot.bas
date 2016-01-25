Attribute VB_Name = "CodeCLPlot"
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

Dim TotalEnergyRange As Single

Dim GraphMinX As Single, GraphMinY As Single
Dim GraphMaxX As Single, GraphMaxY As Single

Sub CLZoomFull_PE(tForm As Form)
' Zoom to origin (Pro Essentials)

ierror = False
On Error GoTo CLZoomFull_PEError

tForm.Pesgo1.PEactions = UNDO_ZOOM&

Exit Sub

' Errors
CLZoomFull_PEError:
IOMsgBox Error$, vbOKOnly + vbCritical, "CLZoomFull_PE"
ierror = True
Exit Sub

End Sub

Sub CLDisplaySpectra_PE(tCLDarkSpectra As Boolean, tForm As Form, datarow As Integer, sample() As TypeSample)
' Display current spectrum from the interface (Pro Essentials)
'  tCLDarkSpectra = false for normal CL spectrum
'  tCLDarkSpectra = true for dark CL spectrum

ierror = False
On Error GoTo CLDisplaySpectra_PEError

Dim i As Integer
Dim temp As Single
Dim temp1 As Single, temp2 As Single

' Define #subset and #points
tForm.Pesgo1.Subsets = 1
tForm.Pesgo1.SubsetColors(0) = tForm.Pesgo1.PEargb(Int(255), Int(0), Int(0), Int(255))                      ' blue
tForm.Pesgo1.points = sample(1).CLSpectraNumberofChannels%(datarow%)

' Display options for Y axis label
If CLIntensityOption% = 0 Then
tForm.Pesgo1.YAxisLabel = "Intensity (counts)"
ElseIf CLIntensityOption% = 1 Then
tForm.Pesgo1.YAxisLabel = "Intensity (cps)"
ElseIf CLIntensityOption% = 2 Then
tForm.Pesgo1.YAxisLabel = "Net Intensity (cps)"
End If

' Define axis and graph properties and X extent
tForm.Pesgo1.ManualScaleControlX = PEMSC_MINMAX ' Manually Control X Axis max and min
tForm.Pesgo1.ManualMinX = sample(1).CLSpectraStartEnergy!(datarow%)
tForm.Pesgo1.ManualMaxX = sample(1).CLSpectraEndEnergy!(datarow%)
TotalEnergyRange! = (sample(1).CLSpectraEndEnergy!(datarow%) - sample(1).CLSpectraStartEnergy!(datarow%))

' Define Y extent
tForm.Pesgo1.ManualScaleControlY = PEMSC_MIN ' Autoscale Control Y Axis max, Manual Control min
tForm.Pesgo1.ManualMinY = 0

If VerboseMode And DebugMode Then
msg$ = "CL Display: "
msg$ = msg$ & "Start nm = " & Format$(sample(1).CLSpectraStartEnergy!(datarow%)) & ", "
msg$ = msg$ & "Stop nm = " & Format$(sample(1).CLSpectraEndEnergy!(datarow%)) & ", "
msg$ = msg$ & "numChan= " & Format$(sample(1).CLSpectraNumberofChannels%(datarow%)) & ", "
Call IOWriteLog(msg$)
End If

' Load y axis data (nb PE array starts 0)
For i% = 1 To sample(1).CLSpectraNumberofChannels%(datarow%)

' Display CL spectra
If Not tCLDarkSpectra Then
If CLIntensityOption% = 0 Then
tForm.Pesgo1.ydata(0, i% - 1) = sample(1).CLSpectraIntensities&(datarow%, i%)                                                     ' raw counts
ElseIf CLIntensityOption% = 1 Then
temp1! = sample(1).CLSpectraIntensities&(datarow%, i%) / sample(1).CLAcquisitionCountTime!(datarow%)
tForm.Pesgo1.ydata(0, i% - 1) = temp1!                                                                                            ' counts/sec
ElseIf CLIntensityOption% = 2 Then
If sample(1).CLAcquisitionCountTime!(datarow%) = 0 Then GoTo CLDisplaySpectra_PEZeroAcqTime
If sample(1).CLDarkSpectraCountTimeFraction!(datarow%) = 0 Then GoTo CLDisplaySpectra_PEZeroFraction
temp1! = sample(1).CLSpectraIntensities&(datarow%, i%) / sample(1).CLAcquisitionCountTime!(datarow%)
temp2! = sample(1).CLSpectraDarkIntensities(datarow%, i%) / (sample(1).CLAcquisitionCountTime!(datarow%) * sample(1).CLDarkSpectraCountTimeFraction!(datarow%))
temp! = temp1! - temp2!
tForm.Pesgo1.ydata(0, i% - 1) = temp!                                                                                             ' net intensities
End If

' Display dark spectra
Else
If CLIntensityOption% = 0 Then
tForm.Pesgo1.ydata(0, i% - 1) = sample(1).CLSpectraDarkIntensities(datarow%, i%)                                                  ' raw counts
Else
If sample(1).CLAcquisitionCountTime!(datarow%) = 0 Then GoTo CLDisplaySpectra_PEZeroAcqTime
If sample(1).CLDarkSpectraCountTimeFraction!(datarow%) = 0 Then GoTo CLDisplaySpectra_PEZeroFraction
temp1! = sample(1).CLAcquisitionCountTime!(datarow%) * sample(1).CLDarkSpectraCountTimeFraction!(datarow%)
tForm.Pesgo1.ydata(0, i% - 1) = sample(1).CLSpectraDarkIntensities(datarow%, i%) / temp1!                                         ' counts/sec
End If
End If

' Calculate and Load x data
temp! = i% * TotalEnergyRange! / (sample(1).CLSpectraNumberofChannels%(datarow%) - 1)
temp! = sample(1).CLSpectraStartEnergy!(datarow%) + temp!
tForm.Pesgo1.xdata(0, i% - 1) = temp!

If VerboseMode And DebugMode Then
Call IOWriteLog("CL Point" & Str$(i%) & ", " & InterfaceStringCLUnitsX$(CLSpectraInterfaceTypeStored%) & Str$(temp!) & ", " & Format$(sample(1).CLSpectraIntensities&(datarow%, i%)) & " counts")      ' raw counts
End If
Next i%

' Debug output
If VerboseMode Then
GraphMinX! = tForm.Pesgo1.ManualMinY
GraphMinY! = tForm.Pesgo1.ManualMinX

GraphMinX! = tForm.Pesgo1.ManualMaxY
GraphMinY! = tForm.Pesgo1.ManualMaxX

Call IOWriteLog("CL Display Spectra: X Min/Max" & Str$(GraphMinX!) & "/" & Str$(GraphMaxX!) & ", Y Min/Max" & Str$(GraphMinY!) & "/" & Str$(GraphMaxY!))
End If

Exit Sub

' Errors
CLDisplaySpectra_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "CLDisplaySpectra_PE"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

CLDisplaySpectra_PEZeroAcqTime:
msg$ = "CL acquisitiion time is zero for datarow " & Format$(datarow%)
IOMsgBox msg$, vbOKOnly + vbExclamation, "CLDisplaySpectra_PE"
ierror = True
Exit Sub

CLDisplaySpectra_PEZeroFraction:
msg$ = "CL dark spectra fraction time is zero for datarow " & Format$(datarow%)
IOMsgBox msg$, vbOKOnly + vbExclamation, "CLDisplaySpectra_PE"
ierror = True
Exit Sub

End Sub

Sub CLInitDisplay_PE(tForm As Form, tCaption As String, datarow As Integer, sample() As TypeSample)
' Init spectrum display (Pro Essentials)

ierror = False
On Error GoTo CLInitDisplay_PEError

Dim astring As String

' Init graph properties
tForm.Pesgo1.Subsets = 1
tForm.Pesgo1.points = 1
tForm.Pesgo1.xdata(0, 0) = 0                    'for empty subset
tForm.Pesgo1.ydata(0, 0) = 0

tForm.Pesgo1.RenderEngine = PERE_GDIPLUS&       ' PERE_DIRECT2D may screw xp people?
tForm.Pesgo1.AntiAliasText = True               ' needed?
tForm.Pesgo1.DataShadows = PEDS_NONE&           ' no data shadows

' Plot type
tForm.Pesgo1.PlottingMethod = SGPM_BAR&         ' bargraph subset
tForm.Pesgo1.AdjoinBars = True                  ' bar always full width of bin, yes or no?

' Title Properties
tForm.Pesgo1.MainTitle = vbNullString
tForm.Pesgo1.SubTitle = vbNullString
tForm.Pesgo1.ImageAdjustTop = 50                ' space above title formatting
tForm.Pesgo1.BorderTypes = PETAB_SINGLE_LINE&

tForm.Pesgo1.XAxisLabel = InterfaceStringCLUnitsX$(CLSpectraInterfaceTypeStored%)
If CLIntensityOption% = 0 Then
tForm.Pesgo1.YAxisLabel = "Intensity (counts)"
ElseIf CLIntensityOption% = 1 Then
tForm.Pesgo1.YAxisLabel = "Intensity (cps)"
ElseIf CLIntensityOption% = 2 Then
tForm.Pesgo1.YAxisLabel = "Net Intensity (cps)"
End If

tForm.Pesgo1.ImageAdjustLeft = 100              ' axis formatting - create a little space on far left

' Enable zoom
tForm.Pesgo1.AllowZooming = PEAZ_HORZANDVERT&
tForm.Pesgo1.ZoomStyle = PEZS_RO2_NOT&

' Allow scroll after zoom
tForm.Pesgo1.ScrollingHorzZoom = True
tForm.Pesgo1.ScrollingVertZoom = True
tForm.Pesgo1.MouseDraggingX = True
tForm.Pesgo1.MouseDraggingY = True
tForm.Pesgo1.ZoomWindow = True

Exit Sub

' Errors
CLInitDisplay_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "CLInitDisplay_PE"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub CLDisplayRedraw_PE()
' Redraw the display, if graph has data (Pro Essentials)

ierror = False
On Error GoTo CLDisplayRedraw_PEError

FormCLDISPLAY.Pesgo1.PEactions = REINITIALIZE_RESETIMAGE ' GGES

Exit Sub

' Errors
CLDisplayRedraw_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "CLDisplayRedraw_PE"
ierror = True
Exit Sub

End Sub

Sub CLZoomGraph_PE(PressStatus%, PressX#, PressY#, PressDataX#, PressDataY#, mode As Integer, tForm As Form)
' User clicked mouse, zoom graph (Pro Essentials)

ierror = False
On Error GoTo CLZoomGraph_PEError

' No need for this as PE handles zoom
Exit Sub

' Errors
CLZoomGraph_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "CLZoomGraph_PE"
ierror = True
Exit Sub

End Sub
