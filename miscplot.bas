Attribute VB_Name = "CodeMiscPlot"
' (c) Copyright 1995-2020 by John J. Donovan
Option Explicit

Sub MiscPlotGetSymbols_PE(nsets As Integer, tPesgo As Pesgo)
' Generate random solid symbols (Pro Essentials code)

ierror = False
On Error GoTo MiscPlotGetSymbols_PEError

Dim j As Integer

For j% = 0 To nsets% - 1
If j% Mod 5 = 0 Then
tPesgo.SubsetPointTypes(j%) = PEPT_DOTSOLID&
tPesgo.SubsetLineTypes(j%) = PELT_THIN_SOLID&
ElseIf j% Mod 5 = 1 Then
tPesgo.SubsetPointTypes(j%) = PEPT_SQUARESOLID&
tPesgo.SubsetLineTypes(j%) = PELT_THIN_SOLID&
ElseIf j% Mod 5 = 2 Then
tPesgo.SubsetPointTypes(j%) = PEPT_DIAMONDSOLID&
tPesgo.SubsetLineTypes(j%) = PELT_THIN_SOLID&
ElseIf j% Mod 5 = 3 Then
tPesgo.SubsetPointTypes(j%) = PEPT_UPTRIANGLESOLID&
tPesgo.SubsetLineTypes(j%) = PELT_THIN_SOLID&
ElseIf j% Mod 5 = 4 Then
tPesgo.SubsetPointTypes(j%) = PEPT_DOWNTRIANGLESOLID&
tPesgo.SubsetLineTypes(j%) = PELT_THIN_SOLID&
End If
Next j%

Exit Sub

' Errors
MiscPlotGetSymbols_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscPlotGetSymbols_PE"
ierror = True
Exit Sub

End Sub

Sub MiscPlotPrintGraph_PE(tGraph As Pesgo)
' Print the graph at current zoom (Pro Essentials code)

ierror = False
On Error GoTo MiscPlotPrintGraph_PEError

Dim bstatus As Boolean

' Launch print dialog
'bstatus = tGraph.PEprintgraph(CLng(0), CLng(0), CLng(0))      ' printer default
bstatus = tGraph.PEprintgraph(CLng(0), CLng(0), CLng(1))      ' print landscape
'bstatus = tGraph.PEprintgraph(CLng(0), CLng(0), CLng(2))      ' print portrait

Exit Sub

' Errors
MiscPlotPrintGraph_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscPlotPrintGraph_PE"
ierror = True
Exit Sub

End Sub

Sub MiscPlotTrack(mode As Integer, x As Single, y As Single, fX As Double, fY As Double, tGraph As Pesgo)
' Convert track data for Pro Essentials
'  mode = 0 for entire graph control
'  mode = 1 for just the plot area

ierror = False
On Error GoTo MiscPlotTrackError

Dim nA As Long, nX As Long, nY As Long
Dim nLeft As Integer, nTop As Integer
Dim nRight As Integer, nBottom As Integer
Dim pX As Integer, pY As Integer
    
' Get last mouse location within control
tGraph.GetLastMouseMove pX%, pY%
    
' Test to see if this is within grid area
tGraph.GetRectGraph nLeft%, nTop%, nRight%, nBottom%
If mode% = 0 Or (mode% = 1 And pX% > nLeft% And pX% < nRight% And pY% > nTop% And pY% < nBottom%) Then
   nA& = 0              ' initialize subset to use (if using OverlapMultiAxes)
   nX& = pX%      ' initialize nX and nY with mouse location
   nY& = pY%
   tGraph.PEconvpixeltograph nA&, nX&, nY&, fX#, fY#, 0, 0, 0
End If

Exit Sub

' Errors
MiscPlotTrackError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscPlotTrack"
ierror = True
Exit Sub

End Sub

Sub MiscPlotInit(tGraph As Pesgo, tZoom As Boolean)
' Create a basic blank graph with or without zoom feature

ierror = False
On Error GoTo MiscPlotInitError

tGraph.RenderEngine = PERE_GDIPLUS&
tGraph.AntiAliasText = True
tGraph.DataShadows = PEDS_NONE&                 ' no data shadows
tGraph.LineShadows = False
tGraph.PointGradientStyle = PEPGS_NONE&

tGraph.PrepareImages = True
tGraph.CacheBmp = True
tGraph.FixedFonts = True
tGraph.FontSize = PEFS_LARGE&

' Plot Formatting
tGraph.DataShadows = PEDS_NONE&
tGraph.LineShadows = False
tGraph.PointGradientStyle = PEPGS_NONE&

tGraph.AntiAliasGraphics = True
tGraph.AntiAliasText = True

tGraph.DpiX = 450
tGraph.DpiY = 450

tGraph.AnnotationsInFront = False
tGraph.BorderTypes = PETAB_SINGLE_LINE&
tGraph.AxisBorderType = PEABT_THIN_LINE&

tGraph.GraphAnnotationX(-1) = 0                 ' empty annotation array
tGraph.GraphAnnotationY(-1) = 0

tGraph.AxisNumericFormatX = PEANF_EXP_NOTATION&
tGraph.AxisNumericFormatY = PEANF_EXP_NOTATION&

tGraph.MainTitle = vbNullString
tGraph.SubTitle = vbNullString
tGraph.XAxisLabel = vbNullString
tGraph.YAxisLabel = vbNullString

tGraph.Subsets = 1
tGraph.points = 1
tGraph.xdata(0, 0) = 0                          ' for empty subset
tGraph.ydata(0, 0) = 0

' Enable zoom
If tZoom Then
tGraph.AllowZooming = PEAZ_HORZANDVERT&
tGraph.ZoomStyle = PEZS_RO2_NOT&

' Scrolling options (when zoomed)
tGraph.ScrollingHorzZoom = True
tGraph.ScrollingVertZoom = True
tGraph.MouseDraggingX = True
tGraph.MouseDraggingY = True
tGraph.ZoomWindow = True

' Disable zoom
Else
tGraph.AllowZooming = PEAZ_NONE&
End If

' Force modal dialog to true for maximize mode <esc> press
tGraph.ModalDialogs = True

tGraph.PEactions = REINITIALIZE_RESETIMAGE                ' generate new plot

Exit Sub

' Errors
MiscPlotInitError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscPlotInit"
ierror = True
Exit Sub

End Sub

