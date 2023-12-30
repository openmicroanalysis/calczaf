Attribute VB_Name = "CodeCalcSpec"
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

Const MAXSPECTROTYPES% = 3  ' (0 to MAXSPECTROTYPES%)

Dim tSpecType As Integer    ' 0 = JEOL 140 mm, 1 = Cameca 160 mm, 2 = JEOL 100 mm, 3 =Cameca 180 mm
Dim tSpecMillimeters(0 To MAXSPECTROTYPES%) As Single

Dim tElement As String
Dim tXray As String
Dim tCrystal As String
Dim tBragg As Integer

Sub CalcSpecLoad()
' Loads the calculate spectro position form

ierror = False
On Error GoTo CalcSpecLoadError

Dim i As Integer, ip As Integer

' Load controls
FormCALCSPEC.ComboElement.Clear
For i% = 0 To MAXELM% - 1
FormCALCSPEC.ComboElement.AddItem Symlo$(i% + 1)
Next i%

FormCALCSPEC.ComboXRay.Clear
For i% = 0 To MAXRAY% - 2   ' do not include specified element
FormCALCSPEC.ComboXRay.AddItem Xraylo$(i% + 1)
Next i%

FormCALCSPEC.ComboCrystal.Clear
For i% = 0 To MAXCRYSTYPE% - 1
If Trim$(AllCrystalNames$(i% + 1)) <> vbNullString Then
FormCALCSPEC.ComboCrystal.AddItem AllCrystalNames$(i% + 1)
End If
Next i%

FormCALCSPEC.ComboBraggOrder.Clear
For i% = 1 To MAXKLMORDER%
FormCALCSPEC.ComboBraggOrder.AddItem Format$(i%)
Next i%

' Load defaults
If tElement$ = vbNullString Then tElement$ = "Fe"
ip% = IPOS1%(MAXELM%, tElement$, Symlo$())
If ip% <> 0 Then FormCALCSPEC.ComboElement.ListIndex% = ip% - 1

If tXray$ = vbNullString Then tXray$ = "Ka"
ip% = IPOS1%(MAXRAY% - 2, tXray$, Xraylo$())
If ip% <> 0 Then FormCALCSPEC.ComboXRay.ListIndex% = ip% - 1

If tCrystal$ = vbNullString Then tCrystal$ = "LIF"
ip% = IPOS1%(MAXCRYSTYPE%, tCrystal$, AllCrystalNames$())
If ip% <> 0 Then FormCALCSPEC.ComboCrystal.ListIndex% = ip% - 1

If tBragg% = 0 Then tBragg% = 1
FormCALCSPEC.ComboBraggOrder.ListIndex% = tBragg% - 1

' Load spectrometer type (0 to MAXSPECTROTYPES%)
If tSpecType% = 0 Then tSpecType% = 0
FormCALCSPEC.OptionInstrument(tSpecType%).Value = True

' Load millimeters for Rowland focal circles (0 to MAXSPECTROTYPES%)
tSpecMillimeters!(0) = 140#     ' JEOL 140
tSpecMillimeters!(1) = 160#     ' Cameca 160
tSpecMillimeters!(2) = 140#     ' JEOL 100 (reads the same as JEOL 140mm spectro- do not ask why!)
tSpecMillimeters!(3) = 180#     ' Cameca 180

' Show form
FormCALCSPEC.Show vbModeless

Exit Sub

' Errors
CalcSpecLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcSpecLoad"
ierror = True
Exit Sub

End Sub

Sub CalcSpecSave()
' Saves the calculate spectro position form

ierror = False
On Error GoTo CalcSpecSaveError

Dim i As Integer, ip As Integer

' Check if user typed an element in
Call CalcSpecSelect
If ierror Then Exit Sub

' Save defaults
tElement$ = FormCALCSPEC.ComboElement.List(FormCALCSPEC.ComboElement.ListIndex)
ip% = IPOS1%(MAXELM%, tElement$, Symlo$())
If ip% = 0 Then
msg$ = "Element " & tElement$ & " is not a valid element symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcSpecSave"
ierror = True
Exit Sub
End If

tXray$ = FormCALCSPEC.ComboXRay.List(FormCALCSPEC.ComboXRay.ListIndex)
ip% = IPOS1%(MAXRAY% - 1, tXray$, Xraylo$())
If ip% = 0 Then
msg$ = "X-ray " & tXray$ & " is not a valid x-ray symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcSpecSave"
ierror = True
Exit Sub
End If

tCrystal$ = FormCALCSPEC.ComboCrystal.List(FormCALCSPEC.ComboCrystal.ListIndex)
ip% = IPOS1%(MAXCRYSTYPE%, tCrystal$, AllCrystalNames$())
If ip% = 0 Then
msg$ = "Crystal " & tCrystal$ & " is not a valid crystal name"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcSpecSave"
ierror = True
Exit Sub
End If

tBragg% = Val(FormCALCSPEC.ComboBraggOrder.List(FormCALCSPEC.ComboBraggOrder.ListIndex))
If tBragg% < 1 Or tBragg% > MAXKLMORDER% Then
msg$ = "Bragg order " & Format$(tBragg%) & " is not a valid Bragg order (must be between 1 and " & Format$(MAXKLMORDER%) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcSpecSave"
ierror = True
Exit Sub
End If

' Check spectro type (0 = JEOL 140 mm, 1 = Cameca 160 mm, 2 = JEOL 100 mm, 3 = Cameca 180 mm)
For i% = 0 To MAXSPECTROTYPES%
If FormCALCSPEC.OptionInstrument(i%).Value = True Then
tSpecType% = i%
End If
Next i%

Exit Sub

' Errors
CalcSpecSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcSpecSave"
ierror = True
Exit Sub

End Sub

Sub CalcSpecCalculate()
' Subroutine to calculate the specified spectrometer position

ierror = False
On Error GoTo CalcSpecCalculateError

Dim ip As Integer
Dim keV As Single, lambda As Single
Dim x2d As Single, xk As Single
Dim spos As Single, tx2d As Single
Dim esym As String, xsym As String

' Determine the element, x-ray, crystal and Bragg order
Call CalcSpecSave
If ierror Then Exit Sub

' Check for valid crystal (ip% is used below for calculation)
ip% = IPOS1%(MAXCRYSTYPE%, tCrystal$, AllCrystalNames$())
If ip% = 0 Then GoTo CalcSpecBadCrystal

' Get the 2d and k for the crystal for position calculation below
Call MiscGetCrystalParameters(tCrystal$, x2d!, xk!, esym, xsym$)
If ierror Then Exit Sub

' Get angstroms for this element and x-ray
Call XrayGetKevLambda(tElement$, tXray$, keV!, lambda!)
If ierror Then Exit Sub

' Adjust angstroms for Bragg order
lambda! = lambda! * tBragg%

' Check for refractive index and modify 2d as necessary
If FormCALCSPEC.CheckUseRefractiveIndex.Value = vbChecked Then
tx2d! = x2d! * (1# - (xk! / tBragg% ^ 2))
Else
tx2d! = x2d!
End If

' Convert to JEOL L-Units, lambda = d/R, where d = crystal d-spacing, R = rowland circle in mm
If tSpecType% = 0 Or tSpecType% = 2 Then
spos! = lambda! / ((tx2d! / 2#) / tSpecMillimeters!(tSpecType%))    ' convert to L-units

' Convert to Cameca Sin theta, lambda = 2d * spos!
ElseIf tSpecType% = 1 Or tSpecType% = 3 Then
spos! = lambda! / tx2d!                 ' convert to sin theta
spos! = CLng(spos! * 100000#)           ' convert to sin theta * 10^5
End If

' Write to form and log window
If tBragg% = 1 Then
msg$ = tElement$ & " " & tXray$ & " on " & tCrystal$ & " (" & tSpecMillimeters!(tSpecType%) & " mm), is " & Format$(spos!)
Else
msg$ = tElement$ & " " & tXray$ & " (" & Format$(Trim$(RomanNum$(tBragg%))) & ") on " & tCrystal$ & " (" & tSpecMillimeters!(tSpecType%) & " mm), is " & Format$(spos!)
End If

If FormCALCSPEC.CheckUseRefractiveIndex.Value = vbUnchecked Then
msg$ = msg$ & " (without refractive index correction)"
Else
msg$ = msg$ & " (with refractive index correction, k= " & Format$(xk!) & ")"
End If

FormCALCSPEC.TextCalcSpec.Text = msg$
Call IOWriteLog("Spectro position for " & msg$)

Exit Sub

' Errors
CalcSpecCalculateError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "CalcSpecCalculate"
ierror = True
Exit Sub

CalcSpecBadCrystal:
msg$ = "Invalid crystal name " & tCrystal$ & " specified"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcSpecCalculate"
ierror = True
Exit Sub

End Sub

Sub CalcSpecSelect()
' Select the element as the user types

ierror = False
On Error GoTo CalcSpecSelectError

Dim ip As Integer
Dim sym As String

sym$ = Trim$(FormCALCSPEC.ComboElement.Text)
ip% = IPOS1%(MAXELM%, sym$, Symlo$())
If ip% = 0 Then Exit Sub

FormCALCSPEC.ComboElement.ListIndex = ip% - 1

Exit Sub

' Errors
CalcSpecSelectError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "CalcSpecSelect"
ierror = True
Exit Sub

End Sub
