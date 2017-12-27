Attribute VB_Name = "CodeStandardCoating"
' (c) Copyright 1995-2018 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

' All variables
Dim tFlag As Integer, tElement As Integer
Dim tDensity As Single, tThickness As Single

' Individual
Dim tStandardCoatingFlag(1 To MAXSTD%) As Integer    ' 0 = not coated, 1 = coated
Dim tStandardCoatingElement(1 To MAXSTD%) As Integer
Dim tStandardCoatingDensity(1 To MAXSTD%) As Single
Dim tStandardCoatingThickness(1 To MAXSTD%) As Single ' in angstroms

Sub StandardCoatingLoad()
' Load the standard parameters form

ierror = False
On Error GoTo StandardCoatingLoadError

Dim i As Integer

' Load variables to module level
For i% = 1 To NumberofStandards%
tStandardCoatingFlag%(i%) = StandardCoatingFlag%(i%)
tStandardCoatingElement%(i%) = StandardCoatingElement%(i%)
tStandardCoatingDensity!(i%) = StandardCoatingDensity!(i%)
tStandardCoatingThickness!(i%) = StandardCoatingThickness!(i%)
Next i%

tFlag% = DefaultStandardCoatingFlag%
tElement% = DefaultStandardCoatingElement%
tDensity! = DefaultStandardCoatingDensity!
tThickness! = DefaultStandardCoatingThickness!

' Load single element list
FormSTANDARDCOATING.ComboCoatingElement.Clear
For i% = 0 To MAXELM% - 1
FormSTANDARDCOATING.ComboCoatingElement.AddItem Symlo$(i% + 1)
Next i%

' Load global parameters
FormSTANDARDCOATING.ComboCoatingElementAll.Clear
For i% = 0 To MAXELM% - 1
FormSTANDARDCOATING.ComboCoatingElementAll.AddItem Symlo$(i% + 1)
Next i%

If DefaultStandardCoatingFlag% = 1 Then
FormSTANDARDCOATING.CheckCoatingFlagAll.value = vbChecked
Else
FormSTANDARDCOATING.CheckCoatingFlagAll.value = vbUnchecked
End If
FormSTANDARDCOATING.ComboCoatingElementAll.Text = Symlo$(DefaultStandardCoatingElement%)
FormSTANDARDCOATING.TextCoatingDensityAll.Text = Format$(DefaultStandardCoatingDensity!)
FormSTANDARDCOATING.TextCoatingThicknessAll.Text = Format$(DefaultStandardCoatingThickness!)

' Load standards
FormSTANDARDCOATING.ListStandard.Clear
For i% = 1 To NumberofStandards%
msg$ = Format$(StandardNumbers(i%), a40) & " " & StandardNames$(i%)
FormSTANDARDCOATING.ListStandard.AddItem msg$
Next i%

' Select last item in list
If FormSTANDARDCOATING.ListStandard.ListCount > 0 Then
FormSTANDARDCOATING.ListStandard.ListIndex = FormSTANDARDCOATING.ListStandard.ListCount - 1
End If

If MiscStringsAreSame(app.EXEName, "Probewin") Then
FormSTANDARDCOATING.LabelTurnOn.Caption = "To enable coating corrections, please explicitly turn on coating options in Analytical | Analysis Options"
Else
FormSTANDARDCOATING.LabelTurnOn.Caption = "To enable coating corrections, please explicitly turn on coating options in Analytical menu"
End If

FormSTANDARDCOATING.Show vbModal

Exit Sub

' Errors
StandardCoatingLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardCoatingLoad"
ierror = True
Exit Sub

End Sub

Sub StandardCoatingSave()
' Save the standard parameters form

ierror = False
On Error GoTo StandardCoatingSaveError

Dim i As Integer

' Load variables to globals
For i% = 1 To NumberofStandards%
StandardCoatingFlag%(i%) = tStandardCoatingFlag%(i%)
StandardCoatingElement%(i%) = tStandardCoatingElement%(i%)
StandardCoatingDensity!(i%) = tStandardCoatingDensity!(i%)
StandardCoatingThickness!(i%) = tStandardCoatingThickness!(i%)
If StandardCoatingFlag%(i%) = 1 Then tFlag% = 1     ' update global flag if any standard uses coating
Next i%

DefaultStandardCoatingFlag% = tFlag%        ' 0 = not coated, 1 = coated
DefaultStandardCoatingElement% = tElement%
DefaultStandardCoatingDensity! = tDensity!
DefaultStandardCoatingThickness! = tThickness!

' Save coating globals (leave commented out so user has to explicitly turn on in Analytical menu)
If DefaultStandardCoatingFlag% = 1 Then
'UseConductiveCoatingCorrectionForElectronAbsorption = True
'UseConductiveCoatingCorrectionForXrayTransmission = True
End If

' Force analysis update to reload changes to standards
AllAnalysisUpdateNeeded = True
AllAFactorUpdateNeeded = True

Exit Sub

' Errors
StandardCoatingSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardCoatingSave"
ierror = True
Exit Sub

End Sub

Sub StandardCoatingSelect()
' Select the parameter based on selected standard

ierror = False
On Error GoTo StandardCoatingSelectError

Dim ip As Integer

If FormSTANDARDCOATING.ListStandard.ListCount < 1 Then Exit Sub
If FormSTANDARDCOATING.ListStandard.ListIndex < 0 Then Exit Sub

' Calculate standard
ip% = FormSTANDARDCOATING.ListStandard.ListIndex + 1
If ip% > 0 Then
If tStandardCoatingFlag%(ip%) = 1 Then
FormSTANDARDCOATING.CheckCoatingFlag.value = vbChecked
Else
FormSTANDARDCOATING.CheckCoatingFlag.value = vbUnchecked
End If

FormSTANDARDCOATING.ComboCoatingElement.Text = Symlo$(tStandardCoatingElement%(ip%))
FormSTANDARDCOATING.TextCoatingDensity.Text = Format$(tStandardCoatingDensity!(ip%))
FormSTANDARDCOATING.TextCoatingThickness.Text = Format$(tStandardCoatingThickness!(ip%))
End If

Exit Sub

' Errors
StandardCoatingSelectError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardCoatingSelect"
ierror = True
Exit Sub

End Sub

Sub StandardCoatingApply()
' Apply the parameters to the selected standard

ierror = False
On Error GoTo StandardCoatingApplyError

Dim ip As Integer, ipp As Integer
Dim sym As String

If FormSTANDARDCOATING.ListStandard.ListCount < 1 Then Exit Sub
If FormSTANDARDCOATING.ListStandard.ListIndex < 0 Then Exit Sub

' Calculate standard
ip% = FormSTANDARDCOATING.ListStandard.ListIndex + 1
If ip% > 0 Then
If FormSTANDARDCOATING.CheckCoatingFlag.value = vbChecked Then
tStandardCoatingFlag%(ip%) = 1
Else
tStandardCoatingFlag%(ip%) = 0
End If

sym$ = FormSTANDARDCOATING.ComboCoatingElement.Text
ipp% = IPOS1(MAXELM%, sym$, Symlo$())
If ipp% = 0 Then
msg$ = "Not a valid element for the Standard Conductive Coating"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardCoatingApply"
ierror = True
Exit Sub
End If
tStandardCoatingElement%(ip%) = ipp%

If Val(FormSTANDARDCOATING.TextCoatingDensity.Text) < 0.1 Or Val(FormSTANDARDCOATING.TextCoatingDensity.Text) > 50# Then
msg$ = "Density out of range for the Standard Conductive Coating (must be between 0.1 and 50 gm/cm^3)"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardCoatingApply"
Else
tStandardCoatingDensity!(ip%) = Val(FormSTANDARDCOATING.TextCoatingDensity.Text)
End If

If Val(FormSTANDARDCOATING.TextCoatingThickness.Text) < 1 Or Val(FormSTANDARDCOATING.TextCoatingThickness.Text) > 10000# Then
msg$ = "Thickness out of range for the Standard Conductive Coating (must be between 1 and 10,000 angstroms)"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardCoatingApply"
Else
tStandardCoatingThickness!(ip%) = Val(FormSTANDARDCOATING.TextCoatingThickness.Text)
End If

End If

Exit Sub

' Errors
StandardCoatingApplyError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardCoatingApply"
ierror = True
Exit Sub

End Sub

Sub StandardCoatingAssignToAll()
' Assign all standards to same parameters

ierror = False
On Error GoTo StandardCoatingAssignToAllError

Dim i As Integer, ipp As Integer
Dim sym As String

If FormSTANDARDCOATING.CheckCoatingFlagAll.value = vbChecked Then
tFlag% = 1
Else
tFlag% = 0
End If

sym$ = FormSTANDARDCOATING.ComboCoatingElementAll.Text
ipp% = IPOS1(MAXELM%, sym$, Symlo$())
If ipp% = 0 Then
msg$ = "Not a valid element for the Standard Conductive Coating"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardCoatingAssignToAll"
ierror = True
Exit Sub
End If
tElement% = ipp%

If Val(FormSTANDARDCOATING.TextCoatingDensityAll.Text) < 0.1 Or Val(FormSTANDARDCOATING.TextCoatingDensityAll.Text) > 50# Then
msg$ = "Density out of range for the Standard Conductive Coating (must be between 0.1 and 50 gm/cm^3)"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardCoatingAssignToAll"
Else
tDensity! = Val(FormSTANDARDCOATING.TextCoatingDensityAll.Text)
End If

If Val(FormSTANDARDCOATING.TextCoatingThicknessAll.Text) < 1 Or Val(FormSTANDARDCOATING.TextCoatingThicknessAll.Text) > 10000# Then
msg$ = "Thickness out of range for the Standard Conductive Coating (must be between 1 and 10,000 angstroms)"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardCoatingAssignToAll"
Else
tThickness! = Val(FormSTANDARDCOATING.TextCoatingThicknessAll.Text)
End If

' Load variables to module level
For i% = 1 To NumberofStandards%
tStandardCoatingFlag%(i%) = tFlag%
tStandardCoatingElement%(i%) = tElement%
tStandardCoatingDensity!(i%) = tDensity!
tStandardCoatingThickness!(i%) = tThickness!
Next i%

' Update current standard selection
Call StandardCoatingSelect
If ierror Then Exit Sub

Exit Sub

' Errors
StandardCoatingAssignToAllError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardCoatingAssignToAll"
ierror = True
Exit Sub

End Sub
