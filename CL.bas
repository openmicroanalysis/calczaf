Attribute VB_Name = "CodeCL"
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit

Dim TotalEnergyRange As Single

Dim GraphMinX As Single, GraphMinY As Single
Dim GraphMaxX As Single, GraphMaxY As Single

Dim CLDataRow As Integer
Dim CLOldSample(1 To 1) As TypeSample

Sub CLZoomFull(tForm As Form)
' Zoom to origin

ierror = False
On Error GoTo CLZoomFullError

tForm.Graph1.XAxisMin = 0
tForm.Graph1.XAxisMax = GraphMaxX!

tForm.Graph1.YAxisMin = 0
tForm.Graph1.YAxisMax = GraphMaxY!

tForm.Graph1.MousePointer = 0
tForm.Graph1.DrawMode = graphDraw

Exit Sub

' Errors
CLZoomFullError:
IOMsgBox Error$, vbOKOnly + vbCritical, "CLZoomFull"
ierror = True
Exit Sub

End Sub

Sub CLDisplaySpectra(tCLIntensityOption As Integer, tCLDarkSpectra As Boolean, tForm As Form, datarow As Integer, sample() As TypeSample)
' Display current spectrum from the interface
'  tCLDarkSpectra = false normal CL spectrum
'  tCLDarkSpectra = true dark CL spectrum

ierror = False
On Error GoTo CLDisplaySpectraError

Dim i As Integer
Dim temp As Single
Dim temp1 As Single, temp2 As Single

' Check for data
If datarow% = 0 Then Exit Sub

' Save
CLOldSample(1) = sample(1)
CLDataRow% = datarow%

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
If sample(1).CLAcquisitionCountTime!(datarow%) = 0 Then GoTo CLDisplaySpectraZeroAcqTime
If sample(1).CLDarkSpectraCountTimeFraction!(datarow%) = 0 Then GoTo CLDisplaySpectraZeroFraction
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
If sample(1).CLAcquisitionCountTime!(datarow%) = 0 Then GoTo CLDisplaySpectraZeroAcqTime
If sample(1).CLDarkSpectraCountTimeFraction!(datarow%) = 0 Then GoTo CLDisplaySpectraZeroFraction
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
CLDisplaySpectraError:
MsgBox Error$, vbOKOnly + vbCritical, "CLDisplaySpectra"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

CLDisplaySpectraZeroAcqTime:
msg$ = "CL acquisitiion time is zero for datarow " & Format$(datarow%)
IOMsgBox msg$, vbOKOnly + vbExclamation, "CLDisplaySpectra"
ierror = True
Exit Sub

CLDisplaySpectraZeroFraction:
msg$ = "CL dark spectra fraction time is zero for datarow " & Format$(datarow%)
IOMsgBox msg$, vbOKOnly + vbExclamation, "CLDisplaySpectra"
ierror = True
Exit Sub

End Sub

Sub CLInitDisplay(tForm As Form, tCaption As String, datarow As Integer, sample() As TypeSample)
' Init spectrum display

ierror = False
On Error GoTo CLInitDisplayError

Dim astring As String

' Check for data
If datarow% = 0 Then Exit Sub

' Load form caption
If tCaption$ <> vbNullString Then
tForm.Caption = "CL Spectrum Display" & " [" & tCaption$ & "]"
End If

' Load label
If sample(1).Linenumber&(datarow%) > 0 Then
astring$ = sample(1).Name$ & ", Line " & Format$(sample(1).Linenumber&(datarow%))                         ' PFE
Else
astring$ = sample(1).CLSpectraKilovolts!(datarow%) & " keV, " & sample(1).Name$                           ' Standard
End If
tForm.LabelSpectrumName.Caption = astring$

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
CLInitDisplayError:
MsgBox Error$, vbOKOnly + vbCritical, "CLInitDisplay"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub CLSetBinSize(tForm As Form)
' Re-size CL graph and recalculate bar sizes

ierror = False
On Error GoTo CLSetBinSizeError

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
CLSetBinSizeError:
MsgBox Error$, vbOKOnly + vbCritical, "CLSetBinSize"
ierror = True
Exit Sub

End Sub

Sub CLDisplayRedraw()
' Redraw the display, if graph has data

ierror = False
On Error GoTo CLDisplayRedrawError

If CLDataRow% < 1 Then Exit Sub
If CLOldSample(1).CLSpectraNumberofChannels%(CLDataRow%) < 1 Then Exit Sub
'FormCLDISPLAY.Graph1.DrawMode = graphBlit     ' refresh the graph (DrawMode = 3) (causes problems in Win8)
FormCLDISPLAY.Graph1.DrawMode = graphDraw     ' refresh the graph (DrawMode = 2)

Exit Sub

' Errors
CLDisplayRedrawError:
MsgBox Error$, vbOKOnly + vbCritical, "CLDisplayRedraw"
ierror = True
Exit Sub

End Sub

Sub CLZoomGraph(PressStatus%, PressX#, PressY#, PressDataX#, PressDataY#, mode As Integer, tForm As Form)
' User clicked mouse, zoom graph

ierror = False
On Error GoTo CLZoomGraphError

If CLDataRow% < 1 Then Exit Sub
If CLOldSample(1).CLSpectraNumberofChannels%(CLDataRow%) < 1 Then Exit Sub

Call ZoomSDKPress(PressStatus%, PressX#, PressY#, PressDataX#, PressDataY#, mode%, tForm)
If ierror Then Exit Sub

Call CLSetBinSize(FormCLDISPLAY)
If ierror Then Exit Sub

Exit Sub

' Errors
CLZoomGraphError:
MsgBox Error$, vbOKOnly + vbCritical, "CLZoomGraph"
ierror = True
Exit Sub

End Sub

Sub CLWriteDiskEMSA(method As Integer, datarow As Integer, sample() As TypeSample, tfilename As String, tForm As Form)
' Write an EMSA format spectrum file based on sample and datarow
'  method = 0 do not ask user to confirm filename
'  method = 1 ask user to confirm file name

ierror = False
On Error GoTo CLWriteDiskEMSAError

' Add extension
If tfilename$ = vbNullString Then tfilename$ = UserDataDirectory$ & "\untitled"

' Confirm filename
If method% = 1 Then
Call IOGetFileName(Int(1), "EMSA", tfilename$, tForm)
If ierror Then Exit Sub

Else
tfilename$ = tfilename$ & "_CL.emsa"
End If

' Export spectrum
Call EMSAWriteSpectrum(Int(1), datarow%, sample(), tfilename$)          ' mode = 1 for CL spectra
If ierror Then Exit Sub

' If no error, save UserData directory
UserDataDirectory$ = MiscGetPathOnly$(tfilename$)

Exit Sub

' Errors
CLWriteDiskEMSAError:
MsgBox Error$, vbOKOnly + vbCritical, "CLWriteDiskEMSA"
ierror = True
Exit Sub

End Sub

Sub CLReadDiskEMSA(datarow As Integer, sample() As TypeSample, tfilename As String, tForm As Form)
' Read an EMSA format spectrum file based on sample and datarow

ierror = False
On Error GoTo CLReadDiskEMSAError

' Add extension
If tfilename$ = vbNullString Then tfilename$ = UserDataDirectory$ & "\untitled"

' Confirm filename
Call IOGetFileName(Int(2), "EMSA", tfilename$, tForm)
If ierror Then Exit Sub

' Export spectrum
Call EMSAReadSpectrum(Int(1), datarow%, sample(), tfilename$)         ' mode = 1 for CL
If ierror Then Exit Sub

' If no error, save UserData directory
UserDataDirectory$ = MiscGetPathOnly$(tfilename$)

Exit Sub

' Errors
CLReadDiskEMSAError:
MsgBox Error$, vbOKOnly + vbCritical, "CLReadDiskEMSA"
ierror = True
Exit Sub

End Sub




