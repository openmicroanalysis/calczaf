Attribute VB_Name = "CodePlotPhiRhoZ"
' (c) Copyright 1995-2024 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Dim PlotTmpSample(1 To 1) As TypeSample

Sub PlotPhiRhoZCurves(sample() As TypeSample)
' Plot the phi-rhoz curves for all elements in sample

ierror = False
On Error GoTo PlotPhiRhoZCurvesError

Dim k As Integer
Dim n As Long
Dim temp As Single
Dim astring As String

Dim acounter As Integer
Dim xannotation As Single, yannotation As Single, ydecrement As Single

Dim phirhozsums() As Single
Dim phirhozareas60() As Single, phirhozareas80() As Single, phirhozareas90() As Single, phirhozareas95() As Single, phirhozareas99() As Single

If PhiRhoZPlotPoints& = 0 Or PhiRhoZPlotSets& = 0 Then GoTo PlotPhiRhoZCurvesNoPoints

ReDim phirhozsums(1 To PhiRhoZPlotSets&) As Single
ReDim phirhozareas60(1 To PhiRhoZPlotSets&) As Single
ReDim phirhozareas80(1 To PhiRhoZPlotSets&) As Single
ReDim phirhozareas90(1 To PhiRhoZPlotSets&) As Single
ReDim phirhozareas95(1 To PhiRhoZPlotSets&) As Single
ReDim phirhozareas99(1 To PhiRhoZPlotSets&) As Single

' Save sample density for export
PlotTmpSample(1) = sample(1)

' Init the graph
Call MiscPlotInit(FormPlotPhiRhoZ.Pesgo1, True)
If ierror Then Exit Sub

' Missing data points
FormPlotPhiRhoZ.Pesgo1.NullDataValueX = 0
FormPlotPhiRhoZ.Pesgo1.NullDataValueY = 0

FormPlotPhiRhoZ.Pesgo1.PlottingMethod = SGPM_LINE&                  ' lines only

' Load number of points and sets
FormPlotPhiRhoZ.Pesgo1.points = PhiRhoZPlotPoints&
FormPlotPhiRhoZ.Pesgo1.Subsets = PhiRhoZPlotSets& * 2             ' to plot both generated and emitted curves

' Set colors for each data set
For k% = 1 To PhiRhoZPlotSets&
'FormPlotPhiRhoZ.Pesgo1.SubsetLineTypes(k%-1) = PELT_THINSOLID
'FormPlotPhiRhoZ.Pesgo1.SubsetLineTypes(k% - 1) = PELT_MEDIUMTHINSOLID&
FormPlotPhiRhoZ.Pesgo1.SubsetLineTypes(k% - 1) = PELT_MEDIUMSOLID
'FormPlotPhiRhoZ.Pesgo1.SubsetLineTypes(k%-1) = PELT_THICKSOLID
If k% = 1 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(k% - 1) = FormPlotPhiRhoZ.Pesgo1.PEargb(Int(255), Int(215), Int(0), Int(0))
If k% = 2 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(k% - 1) = FormPlotPhiRhoZ.Pesgo1.PEargb(Int(255), Int(0), Int(215), Int(0))
If k% = 3 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(k% - 1) = FormPlotPhiRhoZ.Pesgo1.PEargb(Int(255), Int(0), Int(0), Int(215))
If k% = 4 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(k% - 1) = FormPlotPhiRhoZ.Pesgo1.PEargb(Int(255), Int(192), Int(192), Int(0))
If k% = 5 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(k% - 1) = FormPlotPhiRhoZ.Pesgo1.PEargb(Int(255), Int(0), Int(192), Int(192))
If k% = 6 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(k% - 1) = FormPlotPhiRhoZ.Pesgo1.PEargb(Int(255), Int(0), Int(0), Int(192))
If k% = 7 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(k% - 1) = FormPlotPhiRhoZ.Pesgo1.PEargb(Int(255), Int(192), Int(0), Int(192))
If k% = 8 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(k% - 1) = FormPlotPhiRhoZ.Pesgo1.PEargb(Int(255), Int(192), Int(128), Int(0))
If k% = 9 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(k% - 1) = FormPlotPhiRhoZ.Pesgo1.PEargb(Int(255), Int(192), Int(128), Int(192))
If FormPlotPhiRhoZ.Pesgo1.SubsetColors(k% - 1) = -1 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(k% - 1) = FormPlotPhiRhoZ.Pesgo1.PEargb(Int(255), Int(255), Int(0), Int(0)) ' change white to red for visibility

'FormPlotPhiRhoZ.Pesgo1.SubsetLineTypes(PhiRhoZPlotSets& + (k% - 1)) = PELT_THINSOLID
FormPlotPhiRhoZ.Pesgo1.SubsetLineTypes(PhiRhoZPlotSets& + (k% - 1)) = PELT_MEDIUMTHINSOLID&
'FormPlotPhiRhoZ.Pesgo1.SubsetLineTypes(PhiRhoZPlotSets& + (k% - 1)) = PELT_MEDIUMSOLID
'FormPlotPhiRhoZ.Pesgo1.SubsetLineTypes(PhiRhoZPlotSets& + (k% - 1)) = PELT_THICKSOLID
If k% = 1 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(PhiRhoZPlotSets& + (k% - 1)) = FormPlotPhiRhoZ.Pesgo1.PEargb(Int(255), Int(215), Int(0), Int(0))
If k% = 2 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(PhiRhoZPlotSets& + (k% - 1)) = FormPlotPhiRhoZ.Pesgo1.PEargb(Int(255), Int(0), Int(215), Int(0))
If k% = 3 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(PhiRhoZPlotSets& + (k% - 1)) = FormPlotPhiRhoZ.Pesgo1.PEargb(Int(255), Int(0), Int(0), Int(215))
If k% = 4 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(PhiRhoZPlotSets& + (k% - 1)) = FormPlotPhiRhoZ.Pesgo1.PEargb(Int(255), Int(192), Int(192), Int(0))
If k% = 5 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(PhiRhoZPlotSets& + (k% - 1)) = FormPlotPhiRhoZ.Pesgo1.PEargb(Int(255), Int(0), Int(192), Int(192))
If k% = 6 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(PhiRhoZPlotSets& + (k% - 1)) = FormPlotPhiRhoZ.Pesgo1.PEargb(Int(255), Int(0), Int(0), Int(192))
If k% = 7 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(PhiRhoZPlotSets& + (k% - 1)) = FormPlotPhiRhoZ.Pesgo1.PEargb(Int(255), Int(192), Int(0), Int(192))
If k% = 8 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(PhiRhoZPlotSets& + (k% - 1)) = FormPlotPhiRhoZ.Pesgo1.PEargb(Int(255), Int(192), Int(128), Int(0))
If k% = 9 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(PhiRhoZPlotSets& + (k% - 1)) = FormPlotPhiRhoZ.Pesgo1.PEargb(Int(255), Int(192), Int(128), Int(192))
If k% > 9 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(PhiRhoZPlotSets& + (k% - 1)) = FormPlotPhiRhoZ.Pesgo1.SubsetColors(k% - 1)
If FormPlotPhiRhoZ.Pesgo1.SubsetColors(PhiRhoZPlotSets& + (k% - 1)) = -1 Then FormPlotPhiRhoZ.Pesgo1.SubsetColors(PhiRhoZPlotSets& + (k% - 1)) = FormPlotPhiRhoZ.Pesgo1.PEargb(Int(255), Int(255), Int(0), Int(0))     ' change white to red for visibility
Next k%

' Add legend text
For k% = 1 To PhiRhoZPlotSets&
If Not sample(1).CombinedConditionsFlag Then
FormPlotPhiRhoZ.Pesgo1.SubsetLabels(k% - 1) = sample(1).Elsyms$(k%) & " " & sample(1).Xrsyms$(k%) & ", Generated"
FormPlotPhiRhoZ.Pesgo1.SubsetLabels(PhiRhoZPlotSets& + (k% - 1)) = sample(1).Elsyms$(k%) & " " & sample(1).Xrsyms$(k%) & ", Emitted"
Else
FormPlotPhiRhoZ.Pesgo1.SubsetLabels(k% - 1) = sample(1).Elsyms$(k%) & " " & sample(1).Xrsyms$(k%) & ", Generated" & ", TO=" & Format$(sample(1).TakeoffArray!(k%)) & ", keV=" & Format$(sample(1).KilovoltsArray!(k%))
FormPlotPhiRhoZ.Pesgo1.SubsetLabels(PhiRhoZPlotSets& + (k% - 1)) = sample(1).Elsyms$(k%) & " " & sample(1).Xrsyms$(k%) & ", Emitted"
End If
Next k%

If FormPlotPhiRhoZ.OptionDepthMassOrMicrons(0).Value Then
FormPlotPhiRhoZ.Pesgo1.XAxisLabel = "Phi-Rho-Z Depth (mass thickness, mg/cm^2)"
Else
FormPlotPhiRhoZ.Pesgo1.XAxisLabel = "Phi-Rho-Z Depth (microns, d=" & Format$(sample(1).SampleDensity!) & " gm/cm^3)"
End If
FormPlotPhiRhoZ.Pesgo1.YAxisLabel = "Normalized Intensity"

' Load plot title
If Not sample(1).CombinedConditionsFlag Then
astring$ = sample(1).Name$ & ", TO=" & Str$(sample(1).takeoff!) & ", KeV=" & Str$(sample(1).kilovolts!)
FormPlotPhiRhoZ.Pesgo1.MainTitle = astring$
Else
astring$ = sample(1).Name$
FormPlotPhiRhoZ.Pesgo1.MainTitle = astring$
End If

' Load y axis data (intensity)
For k% = 1 To PhiRhoZPlotSets&
For n& = 1 To PhiRhoZPlotPoints&
FormPlotPhiRhoZ.Pesgo1.ydata(k% - 1, n& - 1) = PhiRhoZPlotY1!(k%, n&)                               ' generated intensities
FormPlotPhiRhoZ.Pesgo1.ydata(PhiRhoZPlotSets& + (k% - 1), n& - 1) = PhiRhoZPlotY2!(k%, n&)          ' emitted intensities
Next n&

' Load x data (mass depth)
For n& = 1 To PhiRhoZPlotPoints&
If FormPlotPhiRhoZ.OptionDepthMassOrMicrons(0).Value Then
FormPlotPhiRhoZ.Pesgo1.xdata(k% - 1, n& - 1) = PhiRhoZPlotX!(k%, n&)
FormPlotPhiRhoZ.Pesgo1.xdata(PhiRhoZPlotSets& + (k% - 1), n& - 1) = PhiRhoZPlotX!(k%, n&)

' Load x data (micron depth)
Else
FormPlotPhiRhoZ.Pesgo1.xdata(k% - 1, n& - 1) = PhiRhoZPlotX!(k%, n&) / (sample(1).SampleDensity! / MICRONSPERCM& * MILLIGMPERGRAM#)
FormPlotPhiRhoZ.Pesgo1.xdata(PhiRhoZPlotSets& + (k% - 1), n& - 1) = PhiRhoZPlotX!(k%, n&) / (sample(1).SampleDensity! / MICRONSPERCM& * MILLIGMPERGRAM#)
End If
Next n&
Next k%

'FormPlotPhiRhoZ.Pesgo1.LegendStyle = PELS_1_LINE_INSIDE_OVERLAP&
'FormPlotPhiRhoZ.Pesgo1.LegendStyle = PELS_1_LINE_INSIDE_AXIS&
'FormPlotPhiRhoZ.Pesgo1.LegendLocation = PELL_TOP&
'FormPlotPhiRhoZ.Pesgo1.LegendLocation = PELL_BOTTOM&
'FormPlotPhiRhoZ.Pesgo1.LegendLocation = PELL_LEFT&
FormPlotPhiRhoZ.Pesgo1.LegendLocation = PELL_RIGHT&
FormPlotPhiRhoZ.Pesgo1.OneLegendPerLine = True                               ' put one legend per line
FormPlotPhiRhoZ.Pesgo1.SimpleLineLegend = True
FormPlotPhiRhoZ.Pesgo1.SimplePointLegend = True                              ' default = False encloses in a box

FormPlotPhiRhoZ.Pesgo1.PEactions = REINITIALIZE_RESETIMAGE&

FormPlotPhiRhoZ.Pesgo1.ManualScaleControlX = PEMSC_NONE&           ' autoscale x axis
FormPlotPhiRhoZ.Pesgo1.ManualScaleControlY = PEMSC_NONE&           ' autoscale y axis

FormPlotPhiRhoZ.Pesgo1.GraphAnnotationX(-1) = 0                    ' empty annotation array
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationY(-1) = 0

FormPlotPhiRhoZ.Pesgo1.ShowAnnotations = True
FormPlotPhiRhoZ.Pesgo1.AnnotationsInFront = True

' Annotation properties (is not working!!!!)
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationTextSize = 60               ' define annotation text size
FormPlotPhiRhoZ.Pesgo1.LabelFont = "Arial"                        ' define Font for annotations (and axes)
FormPlotPhiRhoZ.Pesgo1.HideIntersectingText = PEHIT_NO_HIDING&    ' or PEHIT_HIDE&

' Load calculation options as annotations
xannotation! = FormPlotPhiRhoZ.Pesgo1.ManualMinX + (FormPlotPhiRhoZ.Pesgo1.ManualMaxX - FormPlotPhiRhoZ.Pesgo1.ManualMinX) * 0.7
yannotation! = FormPlotPhiRhoZ.Pesgo1.ManualMaxY * 0.98
ydecrement! = (FormPlotPhiRhoZ.Pesgo1.ManualMaxY - FormPlotPhiRhoZ.Pesgo1.ManualMinY) / 25#

' 0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters
yannotation! = yannotation! - ydecrement!
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationColor(acounter%) = FormPlotPhiRhoZ.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationText(acounter%) = corstring$(CorrectionFlag%)

' Correction options
If CorrectionFlag% = 0 Then
acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationColor(acounter%) = FormPlotPhiRhoZ.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationText(acounter%) = macstring$(MACTypeFlag%)

If EmpTypeFlag% = 1 Then
acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationColor(acounter%) = FormPlotPhiRhoZ.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationText(acounter%) = "Using Empirical MACs if available"
End If

acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationColor(acounter%) = FormPlotPhiRhoZ.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationText(acounter%) = absstring$(iabs%)

acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationColor(acounter%) = FormPlotPhiRhoZ.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationText(acounter%) = bscstring$(ibsc%)

acounter% = acounter% + 1
yannotation! = yannotation! - ydecrement!
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationX(acounter%) = xannotation!
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationY(acounter%) = yannotation!
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationType(acounter%) = PEGAT_NOSYMBOL&
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationColor(acounter%) = FormPlotPhiRhoZ.Pesgo1.PEargb(225, 0, 0, 0)       ' black
FormPlotPhiRhoZ.Pesgo1.GraphAnnotationText(acounter%) = bksstring$(ibks%)
End If

' Calculate percent areas for emitted intensities and output a table to the log window
For k% = 1 To PhiRhoZPlotSets&
phirhozsums!(k%) = 0#
For n& = 1 To PhiRhoZPlotPoints&
phirhozsums!(k%) = phirhozsums!(k%) + PhiRhoZPlotY2!(k%, n&)
Next n&

' Now add emitted intensities to determine 60, 80, 90, 95 and 99% areas and corresponding depths
temp! = 0#
For n& = 1 To PhiRhoZPlotPoints&
temp! = temp! + PhiRhoZPlotY2!(k%, n&)
If temp! / phirhozsums!(k%) <= 0.6 Then phirhozareas60!(k%) = PhiRhoZPlotX!(k%, n&)
If temp! / phirhozsums!(k%) <= 0.8 Then phirhozareas80!(k%) = PhiRhoZPlotX!(k%, n&)
If temp! / phirhozsums!(k%) <= 0.9 Then phirhozareas90!(k%) = PhiRhoZPlotX!(k%, n&)
If temp! / phirhozsums!(k%) <= 0.95 Then phirhozareas95!(k%) = PhiRhoZPlotX!(k%, n&)
If temp! / phirhozsums!(k%) <= 0.99 Then phirhozareas99!(k%) = PhiRhoZPlotX!(k%, n&)
Next n&
Next k%

' Now output a table
Call IOWriteLog(vbNullString)
Call IOWriteLog(sample(1).Name$ & ", Emitted intensity area vs. depth for 60, 80, 90, 95 and 99% areas:")
Call IOWriteLog(corstring$(CorrectionFlag%))
Call IOWriteLog(macstring$(MACTypeFlag%))
If EmpTypeFlag% = 1 Then Call IOWriteLog("Using Empirical MACs if available")
Call IOWriteLog(absstring$(iabs%))
Call IOWriteLog(vbNullString)

For k% = 1 To PhiRhoZPlotSets&
Call IOWriteLog(Format$(MiscAutoUcase$(sample(1).Elsyms$(k%)) & " " & sample(1).Xrsyms$(k%), a80$) & Format$("Mass Depth", a140$) & Format$("Micron Depth", a140$) & ", TO = " & Format$(sample(1).TakeoffArray!(k%)) & ", keV = " & Format$(sample(1).KilovoltsArray!(k%)) & ", d = " & Format$(sample(1).SampleDensity!))
Call IOWriteLog(Format$("  60%", a80$) & Space(6) & MiscAutoFormat$(phirhozareas60!(k%)) & Space(6) & MiscAutoFormat$(CSng(phirhozareas60!(k%) / (sample(1).SampleDensity! / MICRONSPERCM& * MILLIGMPERGRAM#))))
Call IOWriteLog(Format$("  80%", a80$) & Space(6) & MiscAutoFormat$(phirhozareas80!(k%)) & Space(6) & MiscAutoFormat$(CSng(phirhozareas80!(k%) / (sample(1).SampleDensity! / MICRONSPERCM& * MILLIGMPERGRAM#))))
Call IOWriteLog(Format$("  90%", a80$) & Space(6) & MiscAutoFormat$(phirhozareas90!(k%)) & Space(6) & MiscAutoFormat$(CSng(phirhozareas90!(k%) / (sample(1).SampleDensity! / MICRONSPERCM& * MILLIGMPERGRAM#))))
Call IOWriteLog(Format$("  95%", a80$) & Space(6) & MiscAutoFormat$(phirhozareas95!(k%)) & Space(6) & MiscAutoFormat$(CSng(phirhozareas95!(k%) / (sample(1).SampleDensity! / MICRONSPERCM& * MILLIGMPERGRAM#))))
Call IOWriteLog(Format$("  99%", a80$) & Space(6) & MiscAutoFormat$(phirhozareas99!(k%)) & Space(6) & MiscAutoFormat$(CSng(phirhozareas99!(k%) / (sample(1).SampleDensity! / MICRONSPERCM& * MILLIGMPERGRAM#))))
Call IOWriteLog(vbNullString)
Next k%

Call IOStatusAuto(vbNullString)
Exit Sub

' Errors
PlotPhiRhoZCurvesError:
MsgBox Error$, vbOKOnly + vbCritical, "PlotPhiRhoZCurves"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

PlotPhiRhoZCurvesNoPoints:
msg$ = "No data to plot for the current sample"
MsgBox msg$, vbOKOnly + vbExclamation, "PlotPhiRhoZCurves"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub PlotPhiRhoZCurvesExport(tForm As Form)
' Export the phi-rho-z data

ierror = False
On Error GoTo PlotPhiRhoZCurvesExportError

Dim k As Integer, response As Integer
Dim n As Long
Dim tfilename As String, astring As String, bstring As String

' Ask user for output file
tfilename$ = ExportDataFile$
If Trim$(tfilename$) = vbNullString Then tfilename$ = "CalcZAF-Export_prz.dat"
Call IOGetFileName(Int(0), "DAT", tfilename$, tForm)
If ierror Then Exit Sub

' Since user wants to open file make sure it is closed
Close #ExportDataFileNumber%
DoEvents

If Dir$(tfilename$) <> vbNullString Then
msg$ = "Output File: " & vbCrLf
msg$ = msg$ & tfilename$ & vbCrLf
msg$ = msg$ & " already exists, are you sure you want to overwrite it (click No to append)?"
response% = MsgBox(msg$, vbYesNoCancel + vbQuestion + vbDefaultButton2, "PlotPhiRhoZCurvesExport")

If response% = vbCancel Then
ierror = True
Exit Sub
End If

' If user selects overwrite, erase it
If response% = vbYes Then
Kill tfilename$
End If
End If

' No errors, save file name
ExportDataFile$ = tfilename$
Open ExportDataFile$ For Append As #ExportDataFileNumber%

' Output column labels
astring$ = vbNullString
For k% = 1 To PhiRhoZPlotSets&
bstring$ = PlotTmpSample(1).Elsyms$(k%) & " " & PlotTmpSample(1).Xrsyms$(k%)
astring$ = astring$ & VbDquote$ & bstring$ & " mg/cm^2" & VbDquote$ & vbTab & VbDquote$ & bstring & " um" & VbDquote$ & vbTab & VbDquote$ & bstring & " generated" & VbDquote$ & vbTab & VbDquote$ & bstring & " emitted" & VbDquote$ & vbTab
Next k%
Print #ExportDataFileNumber%, astring$

' Export each element set
For n& = 1 To PhiRhoZPlotPoints&
astring$ = vbNullString
For k% = 1 To PhiRhoZPlotSets&

' Export x data (mass depth)
astring$ = astring$ & Format$(PhiRhoZPlotX!(k%, n&)) & vbTab

' Export x data (micron depth)
astring$ = astring$ & Format$(PhiRhoZPlotX!(k%, n&) / (PlotTmpSample(1).SampleDensity! / MICRONSPERCM& * MILLIGMPERGRAM#)) & vbTab

' Export y data (generated)
astring$ = astring$ & Format$(PhiRhoZPlotY1!(k%, n&)) & vbTab

' Export y data (emitted)
astring$ = astring$ & Format$(PhiRhoZPlotY2!(k%, n&)) & vbTab

Next k%
Print #ExportDataFileNumber%, astring$
Next n&

Close #ExportDataFileNumber%

msg$ = "Sample " & PlotTmpSample(1).Name$ & " prz data was exported to " & ExportDataFile$
MsgBox msg$, vbOKOnly + vbInformation, "PlotPhiRhoZCurvesExport"

Exit Sub

' Errors
PlotPhiRhoZCurvesExportError:
Close #ExportDataFileNumber%
MsgBox Error$, vbOKOnly + vbCritical, "PlotPhiRhoZCurvesExport"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub
