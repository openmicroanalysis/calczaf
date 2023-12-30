Attribute VB_Name = "CodeRANGE"
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

Dim XrayLastElementEmitted As Integer
Dim XrayLastXrayEmitted As Integer

Dim XrayLastDensity As Single
Dim XrayLastKev As Single
Dim XrayLastThickness As Single

Dim XrayLastXrayEnergy As Single
Dim XrayLastXrayEdge As Single

Dim RangeTmpSample(1 To 1) As TypeSample

Sub RangeLoad()
' Load Range calculation dialog

ierror = False
On Error GoTo RangeLoadError

' Load xray edits
Dim i As Integer

' Add the list box items
FormRANGE.ComboElement.Clear
For i% = 0 To MAXELM% - 1
FormRANGE.ComboElement.AddItem Symup$(i% + 1)
Next i%

FormRANGE.ComboXRay.Clear
For i% = 0 To MAXRAY% - 1
FormRANGE.ComboXRay.AddItem Xraylo$(i% + 1)
Next i%

' Set index to last element and x-ray
If XrayLastElementEmitted% > 0 Then
FormRANGE.ComboElement.ListIndex = XrayLastElementEmitted%
Else
FormRANGE.ComboElement.ListIndex = ATOMIC_NUM_OXYGEN% - 1 ' oxygen
End If

If XrayLastXrayEmitted% > 0 Then
FormRANGE.ComboXRay.ListIndex = XrayLastXrayEmitted%
Else
FormRANGE.ComboXRay.ListIndex = 0    ' Ka
End If

' Load density and kev
If XrayLastDensity! = 0# Then XrayLastDensity! = 2.7
If XrayLastKev! = 0# Then XrayLastKev! = 15#
If XrayLastThickness! = 0# Then XrayLastThickness! = 1#

If XrayLastXrayEnergy! = 0# Then XrayLastXrayEnergy! = 10#

FormRANGE.TextDensity.Text = Str$(XrayLastDensity!)
FormRANGE.TextKev.Text = Str$(XrayLastKev!)
FormRANGE.TextThickness.Text = Str$(XrayLastThickness!)

FormRANGE.TextXrayEnergy.Text = Str$(XrayLastXrayEnergy!)

' Load densities
FormRANGE.ListAtomicDensities.Clear
For i% = 0 To MAXELM% - 1
FormRANGE.ListAtomicDensities.AddItem "Density of " & Symup$(i% + 1) & " equals " & Format$(AllAtomicDensities!(i% + 1))
Next i%
FormRANGE.ListAtomicDensities.ListIndex = ATOMIC_NUM_SILICON% - 1        ' default = Si

Exit Sub

' Errors
RangeLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "RangeLoad"
ierror = True
Exit Sub

End Sub

Sub RangeSave()
' Save range parameters

ierror = False
On Error GoTo RangeSaveError

Dim elm As String, ray As String
Dim ip As Integer, ipp As Integer

elm$ = FormRANGE.ComboElement.Text
ip% = IPOS1(MAXELM%, elm$, Symlo$())
If ip% = 0 Then GoTo RangeSaveInvalidElement

ray$ = FormRANGE.ComboXRay.Text
ipp% = IPOS1(MAXRAY% - 1, ray$, Xraylo$())
If ipp% = 0 Then GoTo RangeSaveInvalidXray

XrayLastDensity! = Val(FormRANGE.TextDensity.Text)
XrayLastKev! = Val(FormRANGE.TextKev.Text)
XrayLastThickness! = Val(FormRANGE.TextThickness.Text)

XrayLastElementEmitted% = ip% - 1
XrayLastXrayEmitted% = ipp% - 1

XrayLastXrayEnergy! = Val(FormRANGE.TextXrayEnergy.Text)
Exit Sub

' Errors
RangeSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "RangeSave"
ierror = True
Exit Sub

RangeSaveInvalidElement:
msg$ = elm$ & " is an invalid element"
MsgBox msg$, vbOKOnly + vbExclamation, "RangeSave"
ierror = True
Exit Sub

RangeSaveInvalidXray:
msg$ = ray$ & " is an invalid x-ray"
MsgBox msg$, vbOKOnly + vbExclamation, "RangeSave"
ierror = True
Exit Sub

End Sub

Sub RangeGetComposition(mode As Integer)
' Get a composition
' mode = 1 get formula
' mode = 2 get weight
' mode = 3 get standard composition

ierror = False
On Error GoTo RangeGetCompositionError

Dim i As Integer
Dim astring As String

' Write space to log window for new composition
Call IOWriteLog(vbNullString)

' Get formula or weight from user
If mode% = 1 Then FormFORMULA.Show vbModal
If mode% = 2 Then FormWEIGHT.Show vbModal
If mode% = 3 Then FormSTDCOMP.Show vbModal

' If error, just clear and exit
If ierror Then
Call InitSample(RangeTmpSample())
Exit Sub
End If

' Return modified sample
Call FormulaReturnSample(RangeTmpSample())
If ierror Then Exit Sub

' Load string
For i% = 1 To RangeTmpSample(1).LastChan%
astring$ = astring$ & RangeTmpSample(1).Elsyms$(i%) & MiscAutoFormat$(RangeTmpSample(1).ElmPercents!(i%)) & " "
Next i%

FormRANGE.LabelComposition.Caption = astring$

' If from database, load density
If mode% = 3 Then
FormRANGE.TextDensity.Text = RangeTmpSample(1).SampleDensity!
End If

Exit Sub

' Errors
RangeGetCompositionError:
MsgBox Error$, vbOKOnly + vbCritical, "RangeGetComposition"
ierror = True
Exit Sub

End Sub

Sub RangeCalculate(mode As Integer)
' Calculate specified value
' mode = 1 electron range
' mode = 2 xray range (at above electron range)
' mode = 3 xray transmission (for given element and x-ray)
' mode = 4 xray transmission (at arbitrary energy)
' mode = 5 electron transmission (thickness in microns)

ierror = False
On Error GoTo RangeCalculateError

Dim radius As Single
Dim transmission As Single
Dim averagemassabsorption As Single

Dim energy As Single, edge As Single

' Check for composition
If RangeTmpSample(1).LastChan% = 0 Then GoTo RangeCalculateNoComposition

If mode% = 1 Then
Call ConvertCalculateElectronRange(radius!, XrayLastKev!, XrayLastDensity!, RangeTmpSample(1).LastChan%, RangeTmpSample(1).Elsyms$(), RangeTmpSample(1).ElmPercents!())
If ierror Then Exit Sub
msg$ = Format$(XrayLastKev!) & " keV, " & Format$(XrayLastDensity!) & " grams/cm^3"
Call IOWriteLog(msg$)
msg$ = "Electron range radius = " & Format$(radius!) & " um"
FormRANGE.LabelElectronRange.Caption = msg$
Call IOWriteLog(msg$)
End If

If mode% = 2 Then
Call ConvertCalculateXrayRange(radius!, XrayLastKev!, edge!, XrayLastDensity!, Symlo$(XrayLastElementEmitted% + 1), Xraylo$(XrayLastXrayEmitted% + 1), RangeTmpSample(1).LastChan%, RangeTmpSample(1).Elsyms$(), RangeTmpSample(1).ElmPercents!())
If ierror Then Exit Sub
msg$ = Symup$(XrayLastElementEmitted% + 1) & " " & Xraylo$(XrayLastXrayEmitted% + 1) & ", at " & Format$(XrayLastKev!) & " keV, (" & Format$(edge!) & " keV edge energy)"
Call IOWriteLog(msg$)
If edge! >= XrayLastKev! Then
msg$ = "Edge energy (" & Format$(edge!) & " keV) is greater or equal to the beam energy (" & Format$(XrayLastKev!) & " keV) for " & Symup$(XrayLastElementEmitted% + 1) & " " & Xraylo$(XrayLastXrayEmitted% + 1) & ", and therefore is invalid."
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
End If
msg$ = "X-ray production range radius = " & Format$(radius!) & " um"
FormRANGE.LabelXrayRange.Caption = msg$
Call IOWriteLog(msg$)

' Load actual x-ray energy for arbitrary transmission calculation
Call XrayGetEnergy(XrayLastElementEmitted% + 1, XrayLastXrayEmitted% + 1, energy!, edge!)
If ierror Then Exit Sub

XrayLastXrayEnergy! = energy!
XrayLastXrayEdge! = edge!
FormRANGE.Frame4.Caption = "X-ray Transmission at " & Format$(energy!) & " keV"
End If

If mode% = 3 Then
Call ConvertCalculateXrayTransmission(transmission!, averagemassabsorption!, XrayLastDensity!, XrayLastThickness!, Symlo$(XrayLastElementEmitted% + 1), Xraylo$(XrayLastXrayEmitted% + 1), RangeTmpSample(1).LastChan%, RangeTmpSample(1).Elsyms$(), RangeTmpSample(1).ElmPercents!())
If ierror Then Exit Sub
msg$ = Symup$(XrayLastElementEmitted% + 1) & " " & Xraylo$(XrayLastXrayEmitted% + 1) & ", x-ray transmission fraction through thickness " & Format$(XrayLastThickness!) & " um (average u/p = " & Format$(averagemassabsorption!) & ") = " & Format$(transmission!)
FormRANGE.LabelXrayTransmission.Caption = msg$
Call IOWriteLog(msg$)
End If

If mode% = 4 Then
Call ConvertCalculateXrayTransmission2(transmission!, averagemassabsorption!, XrayLastDensity!, XrayLastThickness!, XrayLastXrayEnergy!, RangeTmpSample(1).LastChan%, RangeTmpSample(1).Elsyms$(), RangeTmpSample(1).ElmPercents!())
If ierror Then Exit Sub
msg$ = "X-ray transmission fraction at energy " & Format$(XrayLastXrayEnergy!) & " keV, thickness of " & Format$(XrayLastThickness!) & " um (average u/p = " & Format$(averagemassabsorption!) & ") = " & Format$(transmission!)
FormRANGE.LabelXrayTransmission2.Caption = msg$
Call IOWriteLog(msg$)
End If

If mode% = 5 Then
Call ConvertCalculateElectronEnergy(energy!, XrayLastKev!, XrayLastDensity!, XrayLastThickness!, RangeTmpSample(1).LastChan%, RangeTmpSample(1).Elsyms$(), RangeTmpSample(1).ElmPercents!())
If ierror Then Exit Sub
msg$ = "Electron energy transmitted at incident electron energy of " & Format$(XrayLastKev!) & " keV, " & Format$(XrayLastDensity!) & " grams/cm^3, thickness of " & Format$(XrayLastThickness!) & " um = " & Format$(energy!) & " keV"
FormRANGE.LabelElectronEnergyFinal.Caption = msg$
Call IOWriteLog(msg$)
End If

Exit Sub

' Errors
RangeCalculateError:
MsgBox Error$, vbOKOnly + vbCritical, "RangeCalculate"
ierror = True
Exit Sub

RangeCalculateNoComposition:
msg$ = "No composition entered yet"
MsgBox msg$, vbOKOnly + vbExclamation, "RangeCalculate"
ierror = True
Exit Sub

End Sub

