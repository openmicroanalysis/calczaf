Attribute VB_Name = "CodeCLPlot"
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit

Dim TotalEnergyRange As Single

Dim GraphMinX As Single, GraphMinY As Single
Dim GraphMaxX As Single, GraphMaxY As Single

Sub CLZoomFull_GS(tForm As Form)
' Zoom to origin (Graphics Server)

ierror = False
On Error GoTo CLZoomFull_GSError

tForm.Graph1.XAxisMin = 0
tForm.Graph1.XAxisMax = GraphMaxX!

tForm.Graph1.YAxisMin = 0
tForm.Graph1.YAxisMax = GraphMaxY!

tForm.Graph1.MousePointer = 0
tForm.Graph1.DrawMode = graphDraw

Exit Sub

' Errors
CLZoomFull_GSError:
IOMsgBox Error$, vbOKOnly + vbCritical, "CLZoom_GSFull"
ierror = True
Exit Sub

End Sub

Sub CLZoomFull_PE(tForm As Form)
' Zoom to origin (Pro Essentials)

ierror = False
On Error GoTo CLZoomFull_PEError

tForm.Graph1.XAxisMin = 0
tForm.Graph1.XAxisMax = GraphMaxX!

tForm.Graph1.YAxisMin = 0
tForm.Graph1.YAxisMax = GraphMaxY!

tForm.Graph1.MousePointer = 0
tForm.Graph1.DrawMode = graphDraw

Exit Sub

' Errors
CLZoomFull_PEError:
IOMsgBox Error$, vbOKOnly + vbCritical, "CLZoomFull_PE"
ierror = True
Exit Sub

End Sub

Sub CLDisplaySpectra_GS(tCLIntensityOption As Integer, tCLDarkSpectra As Boolean, tForm As Form, datarow As Integer, sample() As TypeSample)
' Display current spectrum from the interface (Graphics Server)
'  tCLDarkSpectra = false normal CL spectrum
'  tCLDarkSpectra = true dark CL spectrum

ierror = False
On Error GoTo CLDisplaySpectra_GSError

Dim i As Integer
Dim temp As Single
Dim temp1 As Single, temp2 As Single

' Check for data
If sample(1).CLSpectraNumberofChannels%(datarow%) < 1 Then Exit Sub

' Display plot and fit
tForm.Graph1.NumPoints = sample(1).CLSpectraNumberofChannels%(datarow%)
If tCLIntensityOption% = 0 Then
tForm.Graph1.LeftTitle = "Intensity (counts)"
ElseIf tCLIntensityOption% = 1 Then
tForm.Graph1.LeftTitle = "Intensity (cps)"
ElseIf tCLIntensityOption% = 2 Then
tForm.Graph1.LeftTitle = "Net Intensity (cps)"
End If

' Set aspect ratio for axes (in keV)
tForm.Graph1.XAxisMin = sample(1).CLSpectraStartEnergy!(datarow%)
tForm.Graph1.XAxisMax = sample(1).CLSpectraEndEnergy!(datarow%)
TotalEnergyRange! = (sample(1).CLSpectraEndEnergy!(datarow%) - sample(1).CLSpectraStartEnergy!(datarow%))

If VerboseMode And DebugMode Then
msg$ = "CL Display: "
msg$ = msg$ & "Start nm = " & Format$(sample(1).CLSpectraStartEnergy!(datarow%)) & ", "
msg$ = msg$ & "Stop nm = " & Format$(sample(1).CLSpectraEndEnergy!(datarow%)) & ", "
msg$ = msg$ & "numChan= " & Format$(sample(1).CLSpectraNumberofChannels%(datarow%)) & ", "
Call IOWriteLog(msg$)
End If

' Load x and y axis data
For i% = 1 To sample(1).CLSpectraNumberofChannels%(datarow%)

' Display CL spectra
If Not tCLDarkSpectra Then
If tCLIntensityOption% = 0 Then
tForm.Graph1.Data(i%) = sample(1).CLSpectraIntensities&(datarow%, i%)                                                     ' raw counts
ElseIf tCLIntensityOption% = 1 Then
temp1! = sample(1).CLSpectraIntensities&(datarow%, i%) / sample(1).CLAcquisitionCountTime!(datarow%)
tForm.Graph1.Data(i%) = temp1!                                                                                            ' counts/sec
ElseIf tCLIntensityOption% = 2 Then
If sample(1).CLAcquisitionCountTime!(datarow%) = 0 Then GoTo CLDisplaySpectra_GSZeroAcqTime
If sample(1).CLDarkSpectraCountTimeFraction!(datarow%) = 0 Then GoTo CLDisplaySpectra_GSZeroFraction
temp1! = sample(1).CLSpectraIntensities&(datarow%, i%) / sample(1).CLAcquisitionCountTime!(datarow%)
temp2! = sample(1).CLSpectraDarkIntensities(datarow%, i%) / (sample(1).CLAcquisitionCountTime!(datarow%) * sample(1).CLDarkSpectraCountTimeFraction!(datarow%))
temp! = temp1! - temp2!
tForm.Graph1.Data(i%) = temp!                                                                                             ' net intensities
End If

' Display dark spectra
Else
If tCLIntensityOption% = 0 Then
tForm.Graph1.Data(i%) = sample(1).CLSpectraDarkIntensities(datarow%, i%)                                                  ' raw counts
Else
If sample(1).CLAcquisitionCountTime!(datarow%) = 0 Then GoTo CLDisplaySpectra_GSZeroAcqTime
If sample(1).CLDarkSpectraCountTimeFraction!(datarow%) = 0 Then GoTo CLDisplaySpectra_GSZeroFraction
temp1! = sample(1).CLAcquisitionCountTime!(datarow%) * sample(1).CLDarkSpectraCountTimeFraction!(datarow%)
tForm.Graph1.Data(i%) = sample(1).CLSpectraDarkIntensities(datarow%, i%) / temp1!                                         ' counts/sec
End If
End If

' Calculate x position
temp! = i% * TotalEnergyRange! / (sample(1).CLSpectraNumberofChannels%(datarow%) - 1)
temp! = sample(1).CLSpectraStartEnergy!(datarow%) + temp!
tForm.Graph1.xpos(i%) = temp!
tForm.Graph1.Color(i%) = 9        ' blue

If VerboseMode And DebugMode Then
Call IOWriteLog("CL Point" & Str$(i%) & ", " & InterfaceStringCLUnitsX$(CLSpectraInterfaceTypeStored%) & Str$(temp!) & ", " & Format$(sample(1).CLSpectraIntensities&(datarow%, i%)) & " counts")      ' raw counts
End If
Next i%

' Debug output
If VerboseMode Then
GraphMinX! = tForm.Graph1.XAxisMin
GraphMinY! = tForm.Graph1.YAxisMin

GraphMaxX! = tForm.Graph1.XAxisMax
GraphMaxY! = tForm.Graph1.YAxisMax

Call IOWriteLog("CL Display Spectra: X Min/Max" & Str$(GraphMinX!) & "/" & Str$(GraphMaxX!) & ", Y Min/Max" & Str$(GraphMinY!) & "/" & Str$(GraphMaxY!))
End If

' Resize the graph
Call CLSetBinSize(tForm)
If ierror Then Exit Sub

'tForm.Graph1.DrawMode = graphBlit         ' refresh the graph (DrawMode = 3) (causes problems in Win8)
tForm.Graph1.DrawMode = graphDraw         ' refresh the graph (DrawMode = 2)
Exit Sub

' Errors
CLDisplaySpectra_GSError:
MsgBox Error$, vbOKOnly + vbCritical, "CLDisplaySpectra_GS"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

CLDisplaySpectra_GSZeroAcqTime:
msg$ = "CL acquisitiion time is zero for datarow " & Format$(datarow%)
IOMsgBox msg$, vbOKOnly + vbExclamation, "CLDisplaySpectra_GS"
ierror = True
Exit Sub

CLDisplaySpectra_GSZeroFraction:
msg$ = "CL dark spectra fraction time is zero for datarow " & Format$(datarow%)
IOMsgBox msg$, vbOKOnly + vbExclamation, "CLDisplaySpectra_GS"
ierror = True
Exit Sub

End Sub

Sub CLDisplaySpectra_PE(tCLIntensityOption As Integer, tCLDarkSpectra As Boolean, tForm As Form, datarow As Integer, sample() As TypeSample)
' Display current spectrum from the interface (Pro Essentials)
'  tCLDarkSpectra = false normal CL spectrum
'  tCLDarkSpectra = true dark CL spectrum

ierror = False
On Error GoTo CLDisplaySpectra_PEError

Dim i As Integer
Dim temp As Single
Dim temp1 As Single, temp2 As Single

' Display plot and fit
tForm.Graph1.NumPoints = sample(1).CLSpectraNumberofChannels%(datarow%)
If tCLIntensityOption% = 0 Then
tForm.Graph1.LeftTitle = "Intensity (counts)"
ElseIf tCLIntensityOption% = 1 Then
tForm.Graph1.LeftTitle = "Intensity (cps)"
ElseIf tCLIntensityOption% = 2 Then
tForm.Graph1.LeftTitle = "Net Intensity (cps)"
End If

' Set aspect ratio for axes (in keV)
tForm.Graph1.XAxisMin = sample(1).CLSpectraStartEnergy!(datarow%)
tForm.Graph1.XAxisMax = sample(1).CLSpectraEndEnergy!(datarow%)
TotalEnergyRange! = (sample(1).CLSpectraEndEnergy!(datarow%) - sample(1).CLSpectraStartEnergy!(datarow%))

If VerboseMode And DebugMode Then
msg$ = "CL Display: "
msg$ = msg$ & "Start nm = " & Format$(sample(1).CLSpectraStartEnergy!(datarow%)) & ", "
msg$ = msg$ & "Stop nm = " & Format$(sample(1).CLSpectraEndEnergy!(datarow%)) & ", "
msg$ = msg$ & "numChan= " & Format$(sample(1).CLSpectraNumberofChannels%(datarow%)) & ", "
Call IOWriteLog(msg$)
End If

' Load x and y axis data
For i% = 1 To sample(1).CLSpectraNumberofChannels%(datarow%)

' Display CL spectra
If Not tCLDarkSpectra Then
If tCLIntensityOption% = 0 Then
tForm.Graph1.Data(i%) = sample(1).CLSpectraIntensities&(datarow%, i%)                                                     ' raw counts
ElseIf tCLIntensityOption% = 1 Then
temp1! = sample(1).CLSpectraIntensities&(datarow%, i%) / sample(1).CLAcquisitionCountTime!(datarow%)
tForm.Graph1.Data(i%) = temp1!                                                                                            ' counts/sec
ElseIf tCLIntensityOption% = 2 Then
If sample(1).CLAcquisitionCountTime!(datarow%) = 0 Then GoTo CLDisplaySpectra_PEZeroAcqTime
If sample(1).CLDarkSpectraCountTimeFraction!(datarow%) = 0 Then GoTo CLDisplaySpectra_PEZeroFraction
temp1! = sample(1).CLSpectraIntensities&(datarow%, i%) / sample(1).CLAcquisitionCountTime!(datarow%)
temp2! = sample(1).CLSpectraDarkIntensities(datarow%, i%) / (sample(1).CLAcquisitionCountTime!(datarow%) * sample(1).CLDarkSpectraCountTimeFraction!(datarow%))
temp! = temp1! - temp2!
tForm.Graph1.Data(i%) = temp!                                                                                             ' net intensities
End If

' Display dark spectra
Else
If tCLIntensityOption% = 0 Then
tForm.Graph1.Data(i%) = sample(1).CLSpectraDarkIntensities(datarow%, i%)                                                  ' raw counts
Else
If sample(1).CLAcquisitionCountTime!(datarow%) = 0 Then GoTo CLDisplaySpectra_PEZeroAcqTime
If sample(1).CLDarkSpectraCountTimeFraction!(datarow%) = 0 Then GoTo CLDisplaySpectra_PEZeroFraction
temp1! = sample(1).CLAcquisitionCountTime!(datarow%) * sample(1).CLDarkSpectraCountTimeFraction!(datarow%)
tForm.Graph1.Data(i%) = sample(1).CLSpectraDarkIntensities(datarow%, i%) / temp1!                                         ' counts/sec
End If
End If

' Calculate x position
temp! = i% * TotalEnergyRange! / (sample(1).CLSpectraNumberofChannels%(datarow%) - 1)
temp! = sample(1).CLSpectraStartEnergy!(datarow%) + temp!
tForm.Graph1.xpos(i%) = temp!
tForm.Graph1.Color(i%) = 9        ' blue

If VerboseMode And DebugMode Then
Call IOWriteLog("CL Point" & Str$(i%) & ", " & InterfaceStringCLUnitsX$(CLSpectraInterfaceTypeStored%) & Str$(temp!) & ", " & Format$(sample(1).CLSpectraIntensities&(datarow%, i%)) & " counts")      ' raw counts
End If
Next i%

' Debug output
If VerboseMode Then
GraphMinX! = tForm.Graph1.XAxisMin
GraphMinY! = tForm.Graph1.YAxisMin

GraphMaxX! = tForm.Graph1.XAxisMax
GraphMaxY! = tForm.Graph1.YAxisMax

Call IOWriteLog("CL Display Spectra: X Min/Max" & Str$(GraphMinX!) & "/" & Str$(GraphMaxX!) & ", Y Min/Max" & Str$(GraphMinY!) & "/" & Str$(GraphMaxY!))
End If

' Resize the graph
Call CLSetBinSize(tForm)
If ierror Then Exit Sub

'tForm.Graph1.DrawMode = graphBlit         ' refresh the graph (DrawMode = 3) (causes problems in Win8)
tForm.Graph1.DrawMode = graphDraw         ' refresh the graph (DrawMode = 2)
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

Sub CLInitDisplay_GS(tForm As Form, tCaption As String, datarow As Integer, sample() As TypeSample)
' Init spectrum display (Graphics Server)

ierror = False
On Error GoTo CLInitDisplay_GSError

Dim astring As String

' Plot data into control, clear the graph
tForm.Graph1.GraphType = graphBar2D

tForm.Graph1.NumSets = 1
tForm.Graph1.AutoInc = 0

tForm.Graph1.XAxisStyle = 2 ' user defined
tForm.Graph1.XAxisTicks = 10
tForm.Graph1.XAxisMinorTicks = -1   ' 1 minor ticks per tick

tForm.Graph1.YAxisStyle = 2 ' user defined
tForm.Graph1.YAxisTicks = 10
tForm.Graph1.YAxisMinorTicks = -1   ' 1 minor ticks per tick

tForm.Graph1.YAxisMin = 0
tForm.Graph1.YAxisMax = 0

' Printer info
tForm.Graph1.PrintInfo(11) = 1  ' landscape
tForm.Graph1.PrintInfo(12) = 1  ' fit to page

tForm.Graph1.BottomTitle = InterfaceStringCLUnitsX$(CLSpectraInterfaceTypeStored%)
tForm.Graph1.LeftTitleStyle = 1
tForm.Graph1.LeftTitle = vbNullString

tForm.Graph1.AutoInc = 1
tForm.Graph1.Hot = 0                  ' disable hot hit
tForm.Graph1.SDKMouse = 1             ' enable zoom
tForm.Graph1.SDKPaint = 0             ' enable repaint events
tForm.Graph1.DrawMode = graphClear
tForm.Graph1.SDKPaint = 1             ' enable repaint events

Exit Sub

' Errors
CLInitDisplay_GSError:
MsgBox Error$, vbOKOnly + vbCritical, "CLInitDisplay_GS"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub CLInitDisplay_PE(tForm As Form, tCaption As String, datarow As Integer, sample() As TypeSample)
' Init spectrum display (Pro Essentials)

ierror = False
On Error GoTo CLInitDisplay_PEError

Dim astring As String

' Plot data into control, clear the graph
tForm.Graph1.GraphType = graphBar2D

tForm.Graph1.NumSets = 1
tForm.Graph1.AutoInc = 0

tForm.Graph1.XAxisStyle = 2 ' user defined
tForm.Graph1.XAxisTicks = 10
tForm.Graph1.XAxisMinorTicks = -1   ' 1 minor ticks per tick

tForm.Graph1.YAxisStyle = 2 ' user defined
tForm.Graph1.YAxisTicks = 10
tForm.Graph1.YAxisMinorTicks = -1   ' 1 minor ticks per tick

tForm.Graph1.YAxisMin = 0
tForm.Graph1.YAxisMax = 0

' Printer info
tForm.Graph1.PrintInfo(11) = 1  ' landscape
tForm.Graph1.PrintInfo(12) = 1  ' fit to page

tForm.Graph1.BottomTitle = InterfaceStringCLUnitsX$(CLSpectraInterfaceTypeStored%)
tForm.Graph1.LeftTitleStyle = 1
tForm.Graph1.LeftTitle = vbNullString

tForm.Graph1.AutoInc = 1
tForm.Graph1.Hot = 0                  ' disable hot hit
tForm.Graph1.SDKMouse = 1             ' enable zoom
tForm.Graph1.SDKPaint = 0             ' enable repaint events
tForm.Graph1.DrawMode = graphClear
tForm.Graph1.SDKPaint = 1             ' enable repaint events

Exit Sub

' Errors
CLInitDisplay_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "CLInitDisplay_PE"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub CLSetBinSize_GS(tForm As Form)
' Re-size CL graph and recalculate bar sizes (Graphics Server)

ierror = False
On Error GoTo CLSetBinSize_GSError

Dim PointsInCurrentRange As Long
Dim CurrentEnergyRange As Single, CurrentEnergyFraction As Single
Dim BinEnergyWidth As Single, BinEnergyGap As Single
Dim BinEnergyPercent As Single

tForm.Graph1.Bar2DGap = 98
Exit Sub

' Need to get this code working !
If tForm.Graph1.NumPoints = 0# Then Exit Sub
CurrentEnergyRange! = tForm.Graph1.XAxisMax - tForm.Graph1.XAxisMin
If CurrentEnergyRange! = 0 Then Exit Sub

CurrentEnergyFraction! = CurrentEnergyRange! / TotalEnergyRange!
PointsInCurrentRange& = tForm.Graph1.NumPoints * CurrentEnergyFraction!

BinEnergyWidth! = CSng(CurrentEnergyRange! / PointsInCurrentRange&)
BinEnergyGap! = CurrentEnergyRange! / BinEnergyWidth!

tForm.Graph1.Bar2DGap = BinEnergyPercent!

Exit Sub

' Errors
CLSetBinSize_GSError:
MsgBox Error$, vbOKOnly + vbCritical, "CLSetBinSize_GS"
ierror = True
Exit Sub

End Sub

Sub CLSetBinSize_PE(tForm As Form)
' Re-size CL graph and recalculate bar sizes (Pro Essentials)

ierror = False
On Error GoTo CLSetBinSize_PEError

Dim PointsInCurrentRange As Long
Dim CurrentEnergyRange As Single, CurrentEnergyFraction As Single
Dim BinEnergyWidth As Single, BinEnergyGap As Single
Dim BinEnergyPercent As Single

tForm.Graph1.Bar2DGap = 98
Exit Sub

' Need to get this code working !
If tForm.Graph1.NumPoints = 0# Then Exit Sub
CurrentEnergyRange! = tForm.Graph1.XAxisMax - tForm.Graph1.XAxisMin
If CurrentEnergyRange! = 0 Then Exit Sub

CurrentEnergyFraction! = CurrentEnergyRange! / TotalEnergyRange!
PointsInCurrentRange& = tForm.Graph1.NumPoints * CurrentEnergyFraction!

BinEnergyWidth! = CSng(CurrentEnergyRange! / PointsInCurrentRange&)
BinEnergyGap! = CurrentEnergyRange! / BinEnergyWidth!

tForm.Graph1.Bar2DGap = BinEnergyPercent!

Exit Sub

' Errors
CLSetBinSize_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "CLSetBinSize_PE"
ierror = True
Exit Sub

End Sub

Sub CLDisplayRedraw_GS()
' Redraw the display, if graph has data (Graphics Server)

ierror = False
On Error GoTo CLDisplayRedraw_GSError

'FormCLDISPLAY.Graph1.DrawMode = graphBlit     ' refresh the graph (DrawMode = 3) (causes problems in Win8)
FormCLDISPLAY.Graph1.DrawMode = graphDraw     ' refresh the graph (DrawMode = 2)

Exit Sub

' Errors
CLDisplayRedraw_GSError:
MsgBox Error$, vbOKOnly + vbCritical, "CLDisplayRedraw_GS"
ierror = True
Exit Sub

End Sub

Sub CLDisplayRedraw_PE()
' Redraw the display, if graph has data (Pro Essentials)

ierror = False
On Error GoTo CLDisplayRedraw_PEError

'FormCLDISPLAY.Graph1.DrawMode = graphBlit     ' refresh the graph (DrawMode = 3) (causes problems in Win8)
FormCLDISPLAY.Graph1.DrawMode = graphDraw     ' refresh the graph (DrawMode = 2)

Exit Sub

' Errors
CLDisplayRedraw_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "CLDisplayRedraw_PE"
ierror = True
Exit Sub

End Sub

Sub CLZoomGraph_GS(PressStatus%, PressX#, PressY#, PressDataX#, PressDataY#, mode As Integer, tForm As Form)
' User clicked mouse, zoom graph (Graphics Server)

ierror = False
On Error GoTo CLZoomGraph_GSError

Call ZoomSDKPress(PressStatus%, PressX#, PressY#, PressDataX#, PressDataY#, mode%, tForm)
If ierror Then Exit Sub

Call CLSetBinSize(tForm)
If ierror Then Exit Sub

Exit Sub

' Errors
CLZoomGraph_GSError:
MsgBox Error$, vbOKOnly + vbCritical, "CLZoomGraph_GS"
ierror = True
Exit Sub

End Sub

Sub CLZoomGraph_PE(PressStatus%, PressX#, PressY#, PressDataX#, PressDataY#, mode As Integer, tForm As Form)
' User clicked mouse, zoom graph (Pro Essentials)

ierror = False
On Error GoTo CLZoomGraph_PEError

Call ZoomSDKPress(PressStatus%, PressX#, PressY#, PressDataX#, PressDataY#, mode%, tForm)
If ierror Then Exit Sub

Call CLSetBinSize(tForm)
If ierror Then Exit Sub

Exit Sub

' Errors
CLZoomGraph_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "CLZoomGraph_PE"
ierror = True
Exit Sub

End Sub
