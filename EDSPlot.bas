Attribute VB_Name = "CodeEDSPlot"
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit

Dim TotalEnergyRange As Single

Dim GraphMinX As Single, GraphMinY As Single
Dim GraphMaxX As Single, GraphMaxY As Single

Sub EDSDisplayRedraw_GS()
' Redraw the display, if graph has data (Graphics Server)

ierror = False
On Error GoTo EDSDisplayRedraw_GSError

'FormEDSDISPLAY3.Graph1.DrawMode = graphBlit     ' refresh the graph (DrawMode = 3) (causes problems in Win8)
FormEDSDISPLAY3.Graph1.DrawMode = graphDraw     ' refresh the graph (DrawMode = 2)

Exit Sub

' Errors
EDSDisplayRedraw_GSError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSDisplayRedraw_GS"
ierror = True
Exit Sub

End Sub

Sub EDSDisplayRedraw_PE()
' Redraw the display, if graph has data (Pro Essentials)

ierror = False
On Error GoTo EDSDisplayRedraw_PEError

'FormEDSDISPLAY3.Graph1.DrawMode = graphBlit     ' refresh the graph (DrawMode = 3) (causes problems in Win8)
FormEDSDISPLAY3.Graph1.DrawMode = graphDraw     ' refresh the graph (DrawMode = 2)

Exit Sub

' Errors
EDSDisplayRedraw_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSDisplayRedraw_PE"
ierror = True
Exit Sub

End Sub

Sub EDSDisplaySpectra_GS(tEDSIntensityOption As Integer, tForm As Form, datarow As Integer, sample() As TypeSample)
' Display current spectrum from the interface (Graphics Server)

ierror = False
On Error GoTo EDSDisplaySpectra_GSError

Dim i As Integer
Dim xtemp As Single, ytemp As Single

If tEDSIntensityOption% = 0 Then
tForm.Graph1.LeftTitle = "Intensity (counts)"
Else
tForm.Graph1.LeftTitle = "Intensity (cps)"
End If

' Display plot and fit
tForm.Graph1.NumPoints = sample(1).EDSSpectraNumberofChannels%(datarow%)

' Set aspect ratio for axes (in keV)
tForm.Graph1.XAxisMin = sample(1).EDSSpectraStartEnergy!(datarow%)
tForm.Graph1.XAxisMax = sample(1).EDSSpectraEndEnergy!(datarow%)
TotalEnergyRange! = (sample(1).EDSSpectraEndEnergy!(datarow%) - sample(1).EDSSpectraStartEnergy!(datarow%))

tForm.Graph1.YAxisMin = 0
If tEDSIntensityOption% = 0 Then
tForm.Graph1.YAxisMax = sample(1).EDSSpectraMaxCounts&(datarow%)                                                  ' raw counts
Else
If sample(1).EDSSpectraLiveTime!(datarow%) <> 0# Then tForm.Graph1.YAxisMax = sample(1).EDSSpectraMaxCounts&(datarow%) / sample(1).EDSSpectraLiveTime!(datarow%)        ' cps
End If

If VerboseMode And DebugMode Then
msg$ = "EDS Display: "
msg$ = msg$ & "Start keV= " & Format$(sample(1).EDSSpectraStartEnergy!(datarow%)) & ", "
msg$ = msg$ & "Stop keV= " & Format$(sample(1).EDSSpectraEndEnergy!(datarow%)) & ", "
msg$ = msg$ & "numChan= " & Format$(sample(1).EDSSpectraNumberofChannels%(datarow%)) & ", "
msg$ = msg$ & "MaxInt= " & Format$(sample(1).EDSSpectraMaxCounts&(datarow%))
Call IOWriteLog(msg$)
End If

' Load y axis data
For i% = 1 To sample(1).EDSSpectraNumberofChannels%(datarow%)
If tEDSIntensityOption% = 0 Then
ytemp! = sample(1).EDSSpectraIntensities&(datarow%, i%)                                                ' raw counts
Else
If sample(1).EDSSpectraLiveTime!(datarow%) <> 0# Then ytemp! = sample(1).EDSSpectraIntensities&(datarow%, i%) / sample(1).EDSSpectraLiveTime!(datarow%)      ' cps
End If

tForm.Graph1.Data(i%) = ytemp!

' Load x axis data
xtemp! = sample(1).EDSSpectraEVPerChannel!(datarow%) * (i% - 1) / EVPERKEV#
If EDSSpectraInterfaceType% = 2 Then
xtemp! = xtemp! + sample(1).EDSSpectraStartEnergy!(datarow%) ' Bruker zero spectrum starts at plus start energy in keV
End If

tForm.Graph1.xpos(i%) = xtemp!
tForm.Graph1.Color(i%) = 12     ' red

If VerboseMode And DebugMode Then
If tEDSIntensityOption% = 0 Then
Call IOWriteLog("EDS Point" & Str$(i%) & ", keV" & Str$(xtemp!) & ", counts" & Str$(ytemp!))      ' raw counts
Else
Call IOWriteLog("EDS Point" & Str$(i%) & ", keV" & Str$(xtemp!) & ", cps" & Str$(ytemp!))         ' cps
End If
End If
Next i%

' Debug output
If VerboseMode Then
GraphMinX! = tForm.Graph1.XAxisMin
GraphMinY! = tForm.Graph1.YAxisMin

GraphMaxX! = tForm.Graph1.XAxisMax
GraphMaxY! = tForm.Graph1.YAxisMax

Call IOWriteLog("EDS Display Spectra: X Min/Max" & Str$(GraphMinX!) & "/" & Str$(GraphMaxX!) & ", Y Min/Max" & Str$(GraphMinY!) & "/" & Str$(GraphMaxY!))
End If

' Resize the graph
Call EDSSetBinSize(tForm)
If ierror Then Exit Sub

'tForm.Graph1.DrawMode = graphBlit         ' refresh the graph (DrawMode = 3) (causes problems in Win8)
tForm.Graph1.DrawMode = graphDraw         ' refresh the graph (DrawMode = 2)
Exit Sub

' Errors
EDSDisplaySpectra_GSError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSDisplaySpectra_GS"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub EDSDisplaySpectra_PE(tEDSIntensityOption As Integer, tForm As Form, datarow As Integer, sample() As TypeSample)
' Display current spectrum from the interface (Pro Essentials)

ierror = False
On Error GoTo EDSDisplaySpectra_PEError

Dim i As Integer
Dim xtemp As Single, ytemp As Single

If tEDSIntensityOption% = 0 Then
tForm.Graph1.LeftTitle = "Intensity (counts)"
Else
tForm.Graph1.LeftTitle = "Intensity (cps)"
End If

' Display plot and fit
tForm.Graph1.NumPoints = sample(1).EDSSpectraNumberofChannels%(datarow%)

' Set aspect ratio for axes (in keV)
tForm.Graph1.XAxisMin = sample(1).EDSSpectraStartEnergy!(datarow%)
tForm.Graph1.XAxisMax = sample(1).EDSSpectraEndEnergy!(datarow%)
TotalEnergyRange! = (sample(1).EDSSpectraEndEnergy!(datarow%) - sample(1).EDSSpectraStartEnergy!(datarow%))

tForm.Graph1.YAxisMin = 0
If tEDSIntensityOption% = 0 Then
tForm.Graph1.YAxisMax = sample(1).EDSSpectraMaxCounts&(datarow%)                                                  ' raw counts
Else
If sample(1).EDSSpectraLiveTime!(datarow%) <> 0# Then tForm.Graph1.YAxisMax = sample(1).EDSSpectraMaxCounts&(datarow%) / sample(1).EDSSpectraLiveTime!(datarow%)        ' cps
End If

If VerboseMode And DebugMode Then
msg$ = "EDS Display: "
msg$ = msg$ & "Start keV= " & Format$(sample(1).EDSSpectraStartEnergy!(datarow%)) & ", "
msg$ = msg$ & "Stop keV= " & Format$(sample(1).EDSSpectraEndEnergy!(datarow%)) & ", "
msg$ = msg$ & "numChan= " & Format$(sample(1).EDSSpectraNumberofChannels%(datarow%)) & ", "
msg$ = msg$ & "MaxInt= " & Format$(sample(1).EDSSpectraMaxCounts&(datarow%))
Call IOWriteLog(msg$)
End If

' Load y axis data
For i% = 1 To sample(1).EDSSpectraNumberofChannels%(datarow%)
If tEDSIntensityOption% = 0 Then
ytemp! = sample(1).EDSSpectraIntensities&(datarow%, i%)                                                ' raw counts
Else
If sample(1).EDSSpectraLiveTime!(datarow%) <> 0# Then ytemp! = sample(1).EDSSpectraIntensities&(datarow%, i%) / sample(1).EDSSpectraLiveTime!(datarow%)      ' cps
End If

tForm.Graph1.Data(i%) = ytemp!

' Load x axis data
xtemp! = sample(1).EDSSpectraEVPerChannel!(datarow%) * (i% - 1) / EVPERKEV#
If EDSSpectraInterfaceType% = 2 Then
xtemp! = xtemp! + sample(1).EDSSpectraStartEnergy!(datarow%) ' Bruker zero spectrum starts at plus start energy in keV
End If

tForm.Graph1.xpos(i%) = xtemp!
tForm.Graph1.Color(i%) = 12     ' red

If VerboseMode And DebugMode Then
If tEDSIntensityOption% = 0 Then
Call IOWriteLog("EDS Point" & Str$(i%) & ", keV" & Str$(xtemp!) & ", counts" & Str$(ytemp!))      ' raw counts
Else
Call IOWriteLog("EDS Point" & Str$(i%) & ", keV" & Str$(xtemp!) & ", cps" & Str$(ytemp!))         ' cps
End If
End If
Next i%

' Debug output
If VerboseMode Then
GraphMinX! = tForm.Graph1.XAxisMin
GraphMinY! = tForm.Graph1.YAxisMin

GraphMaxX! = tForm.Graph1.XAxisMax
GraphMaxY! = tForm.Graph1.YAxisMax

Call IOWriteLog("EDS Display Spectra: X Min/Max" & Str$(GraphMinX!) & "/" & Str$(GraphMaxX!) & ", Y Min/Max" & Str$(GraphMinY!) & "/" & Str$(GraphMaxY!))
End If

' Resize the graph
Call EDSSetBinSize(tForm)
If ierror Then Exit Sub

'tForm.Graph1.DrawMode = graphBlit         ' refresh the graph (DrawMode = 3) (causes problems in Win8)
tForm.Graph1.DrawMode = graphDraw         ' refresh the graph (DrawMode = 2)
Exit Sub

' Errors
EDSDisplaySpectra_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSDisplaySpectra_PE"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub EDSInitDisplay_GS(tForm As Form, tCaption As String, datarow As Integer, sample() As TypeSample)
' Init spectrum display (Graphics Server)

ierror = False
On Error GoTo EDSInitDisplay_GSError

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

tForm.Graph1.BottomTitle = "keV"
tForm.Graph1.LeftTitleStyle = 1

tForm.Graph1.AutoInc = 1

tForm.Graph1.Hot = 0                  ' disable hot hit
tForm.Graph1.SDKMouse = 1             ' enable zoom
tForm.Graph1.SDKPaint = 0             ' enable repaint events
tForm.Graph1.DrawMode = graphClear
tForm.Graph1.SDKPaint = 1             ' enable repaint events

Exit Sub

' Errors
EDSInitDisplay_GSError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSInitDisplay_GS"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub EDSInitDisplay_PE(tForm As Form, tCaption As String, datarow As Integer, sample() As TypeSample)
' Init spectrum display (Pro Essentials)

ierror = False
On Error GoTo EDSInitDisplay_PEError

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

tForm.Graph1.BottomTitle = "keV"
tForm.Graph1.LeftTitleStyle = 1

tForm.Graph1.AutoInc = 1

tForm.Graph1.Hot = 0                  ' disable hot hit
tForm.Graph1.SDKMouse = 1             ' enable zoom
tForm.Graph1.SDKPaint = 0             ' enable repaint events
tForm.Graph1.DrawMode = graphClear
tForm.Graph1.SDKPaint = 1             ' enable repaint events

Exit Sub

' Errors
EDSInitDisplay_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSInitDisplay_PE"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub EDSSetBinSize_GS(tForm As Form)
' Re-size EDS graph and recalculate bar sizes (Graphics Server)

ierror = False
On Error GoTo EDSSetBinSize_GSError

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
EDSSetBinSize_GSError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSSetBinSize_GS"
ierror = True
Exit Sub

End Sub

Sub EDSSetBinSize_PE(tForm As Form)
' Re-size EDS graph and recalculate bar sizes (Pro Essentials)

ierror = False
On Error GoTo EDSSetBinSize_PEError

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
EDSSetBinSize_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSSetBinSize_PE"
ierror = True
Exit Sub

End Sub

Sub EDSZoomFull_GS(tForm As Form)
' Zoom to origin (Graphics Server)

ierror = False
On Error GoTo EDSZoomFull_GSError

tForm.Graph1.XAxisMin = 0
tForm.Graph1.XAxisMax = GraphMaxX!

tForm.Graph1.YAxisMin = 0
tForm.Graph1.YAxisMax = GraphMaxY!

tForm.Graph1.MousePointer = 0
tForm.Graph1.DrawMode = graphDraw

Exit Sub

' Errors
EDSZoomFull_GSError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSZoomFull_GS"
ierror = True
Exit Sub

End Sub

Sub EDSZoomFull_PE(tForm As Form)
' Zoom to origin (Pro Essentials)

ierror = False
On Error GoTo EDSZoomFull_PEError

tForm.Graph1.XAxisMin = 0
tForm.Graph1.XAxisMax = GraphMaxX!

tForm.Graph1.YAxisMin = 0
tForm.Graph1.YAxisMax = GraphMaxY!

tForm.Graph1.MousePointer = 0
tForm.Graph1.DrawMode = graphDraw

Exit Sub

' Errors
EDSZoomFull_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSZoomFull_PE"
ierror = True
Exit Sub

End Sub

Sub EDSZoomGraph_GS(PressStatus%, PressX#, PressY#, PressDataX#, PressDataY#, mode As Integer, tForm As Form)
' User clicked mouse, zoom graph (Graphics Server)

ierror = False
On Error GoTo EDSZoomGraph_GSError

Call ZoomSDKPress(PressStatus%, PressX#, PressY#, PressDataX#, PressDataY#, mode%, tForm)
If ierror Then Exit Sub

Call EDSSetBinSize(FormEDSDISPLAY3)
If ierror Then Exit Sub

Exit Sub

' Errors
EDSZoomGraph_GSError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSZoomGraph_GS"
ierror = True
Exit Sub

End Sub

Sub EDSZoomGraph_PE(PressStatus%, PressX#, PressY#, PressDataX#, PressDataY#, mode As Integer, tForm As Form)
' User clicked mouse, zoom graph (Pro Essentials)

ierror = False
On Error GoTo EDSZoomGraph_PEError

Call ZoomSDKPress(PressStatus%, PressX#, PressY#, PressDataX#, PressDataY#, mode%, tForm)
If ierror Then Exit Sub

Call EDSSetBinSize(FormEDSDISPLAY3)
If ierror Then Exit Sub

Exit Sub

' Errors
EDSZoomGraph_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSZoomGraph_PE"
ierror = True
Exit Sub

End Sub


