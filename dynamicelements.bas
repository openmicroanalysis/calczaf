Attribute VB_Name = "CodeDynamicElements"
' (c) Copyright 1995-2024 by John J. Donovan
Option Explicit

Sub DynamicElementsLoad(sample() As TypeSample)
' Load the dynamic elements form options

ierror = False
On Error GoTo DynamicElementsLoadError

Dim i As Integer, j As Integer

' Load instructions
msg$ = "Use these options to calculate specified elements dynamically for each pixel based on the specified k-ratio criteria." & vbCrLf & vbCrLf
msg$ = msg$ & "For example to calculate carbonate by stoichiometry to oxygen, one would specify 0.333 atoms of C to O in the (previous) Calculation Options dialog, "
msg$ = msg$ & "and in this dialog, then one might specify Si < 0.01 AND Ca > 0.1 OR Mg > 0.1 and then check the Dynamically Calculate by Stoichiometry To Oxygen checkbox." & vbCrLf & vbCrLf
msg$ = msg$ & "A k-ratio value of zero means the operational criteria will be ignored."
FormDynamicElements.LabelInstructions.Caption = msg$

' Check for valid element channel selections
Call DynamicElementsCheckSelections(Int(0), sample())
If ierror Then Exit Sub

' Load dynamic element flags
If DynamicSpecifiedElementByDifferenceFlag% = True Then
FormDynamicElements.CheckDifference.value = vbChecked
Else
FormDynamicElements.CheckDifference.value = vbUnchecked
End If

If DynamicSpecifiedElementByDifferenceFormulaFlag% = True Then
FormDynamicElements.CheckDifferenceFormula.value = vbChecked
Else
FormDynamicElements.CheckDifferenceFormula.value = vbUnchecked
End If

If DynamicSpecifiedElementByStoichiometryToOxygenFlag% = True Then
FormDynamicElements.CheckStoichiometry.value = vbChecked
Else
FormDynamicElements.CheckStoichiometry.value = vbUnchecked
End If

If DynamicSpecifiedElementByStoichiometryToAnotherFlag% = True Then
FormDynamicElements.CheckRelative.value = vbChecked
Else
FormDynamicElements.CheckRelative.value = vbUnchecked
End If

If DynamicSpecifiedElementExcessOxygenByDroopFlag% = True Then
FormDynamicElements.CheckExcessOxygenByDroop.value = vbChecked
Else
FormDynamicElements.CheckExcessOxygenByDroop.value = vbUnchecked
End If
If ProbeDataFileVersionNumber! < 12.7 Then FormDynamicElements.CheckExcessOxygenByDroop.Enabled = False

' Load element combos with analyzed elements
For j% = 1 To MAXCRITERIA%
For i% = 1 To sample(1).LastElm%
FormDynamicElements.ComboDifferenceElement(j% - 1).AddItem sample(1).Elsyms$(i%)
FormDynamicElements.ComboDifferenceFormulaElement(j% - 1).AddItem sample(1).Elsyms$(i%)
FormDynamicElements.ComboStoichiometryElement(j% - 1).AddItem sample(1).Elsyms$(i%)
FormDynamicElements.ComboRelativeElement(j% - 1).AddItem sample(1).Elsyms$(i%)
FormDynamicElements.ComboExcessOxygenByDroopElement(j% - 1).AddItem sample(1).Elsyms$(i%)
Next i%
FormDynamicElements.ComboDifferenceElement(j% - 1).ListIndex = DynamicSpecifiedElementByDifferenceElement%(j%) - 1
FormDynamicElements.ComboDifferenceFormulaElement(j% - 1).ListIndex = DynamicSpecifiedElementByDifferenceFormulaElement%(j%) - 1
FormDynamicElements.ComboStoichiometryElement(j% - 1).ListIndex = DynamicSpecifiedElementByStoichiometryToOxygenElement%(j%) - 1
FormDynamicElements.ComboRelativeElement(j% - 1).ListIndex = DynamicSpecifiedElementByStoichiometryToAnotherElement%(j%) - 1
FormDynamicElements.ComboExcessOxygenByDroopElement(j% - 1).ListIndex = DynamicSpecifiedElementExcessOxygenByDroopElement%(j%) - 1
Next j%

' Load greater/less combos
For j% = 1 To MAXCRITERIA%
FormDynamicElements.ComboDifferenceGreaterLess(j% - 1).AddItem ">"
FormDynamicElements.ComboDifferenceGreaterLess(j% - 1).AddItem "<"
FormDynamicElements.ComboDifferenceGreaterLess(j% - 1).ListIndex = DynamicSpecifiedElementByDifferenceGreaterLess%(j%)
FormDynamicElements.ComboDifferenceFormulaGreaterLess(j% - 1).AddItem ">"
FormDynamicElements.ComboDifferenceFormulaGreaterLess(j% - 1).AddItem "<"
FormDynamicElements.ComboDifferenceFormulaGreaterLess(j% - 1).ListIndex = DynamicSpecifiedElementByDifferenceFormulaGreaterLess%(j%)
FormDynamicElements.ComboStoichiometryGreaterLess(j% - 1).AddItem ">"
FormDynamicElements.ComboStoichiometryGreaterLess(j% - 1).AddItem "<"
FormDynamicElements.ComboStoichiometryGreaterLess(j% - 1).ListIndex = DynamicSpecifiedElementByStoichiometryToOxygenGreaterLess%(j%)
FormDynamicElements.ComboRelativeGreaterLess(j% - 1).AddItem ">"
FormDynamicElements.ComboRelativeGreaterLess(j% - 1).AddItem "<"
FormDynamicElements.ComboRelativeGreaterLess(j% - 1).ListIndex = DynamicSpecifiedElementByStoichiometryToAnotherGreaterLess%(j%)
FormDynamicElements.ComboExcessOxygenByDroopGreaterLess(j% - 1).AddItem ">"
FormDynamicElements.ComboExcessOxygenByDroopGreaterLess(j% - 1).AddItem "<"
FormDynamicElements.ComboExcessOxygenByDroopGreaterLess(j% - 1).ListIndex = DynamicSpecifiedElementExcessOxygenByDroopGreaterLess%(j%)
Next j%

' Load value text controls
For j% = 1 To MAXCRITERIA%
FormDynamicElements.TextDifferenceValue(j% - 1).Text = DynamicSpecifiedElementByDifferenceValue!(j%)
FormDynamicElements.TextDifferenceFormulaValue(j% - 1).Text = DynamicSpecifiedElementByDifferenceFormulaValue!(j%)
FormDynamicElements.TextStoichiometryValue(j% - 1).Text = DynamicSpecifiedElementByStoichiometryToOxygenValue!(j%)
FormDynamicElements.TextRelativeValue(j% - 1).Text = DynamicSpecifiedElementByStoichiometryToAnotherValue!(j%)
FormDynamicElements.TextExcessOxygenByDroopValue(j% - 1).Text = DynamicSpecifiedElementExcessOxygenByDroopValue!(j%)
Next j%

' Load operator combos
FormDynamicElements.ComboDifferenceOperator1.AddItem "AND"
FormDynamicElements.ComboDifferenceOperator1.AddItem "OR"
FormDynamicElements.ComboDifferenceOperator2.AddItem "AND"
FormDynamicElements.ComboDifferenceOperator2.AddItem "OR"
FormDynamicElements.ComboDifferenceOperator1.ListIndex = DynamicSpecifiedElementByDifferenceOperator1%
FormDynamicElements.ComboDifferenceOperator2.ListIndex = DynamicSpecifiedElementByDifferenceOperator2%
FormDynamicElements.ComboDifferenceFormulaOperator1.AddItem "AND"
FormDynamicElements.ComboDifferenceFormulaOperator1.AddItem "OR"
FormDynamicElements.ComboDifferenceFormulaOperator2.AddItem "AND"
FormDynamicElements.ComboDifferenceFormulaOperator2.AddItem "OR"
FormDynamicElements.ComboDifferenceFormulaOperator1.ListIndex = DynamicSpecifiedElementByDifferenceFormulaOperator1%
FormDynamicElements.ComboDifferenceFormulaOperator2.ListIndex = DynamicSpecifiedElementByDifferenceFormulaOperator2%
FormDynamicElements.ComboStoichiometryOperator1.AddItem "AND"
FormDynamicElements.ComboStoichiometryOperator1.AddItem "OR"
FormDynamicElements.ComboStoichiometryOperator2.AddItem "AND"
FormDynamicElements.ComboStoichiometryOperator2.AddItem "OR"
FormDynamicElements.ComboStoichiometryOperator1.ListIndex = DynamicSpecifiedElementByStoichiometryToOxygenOperator1%
FormDynamicElements.ComboStoichiometryOperator2.ListIndex = DynamicSpecifiedElementByStoichiometryToOxygenOperator2%
FormDynamicElements.ComboRelativeOperator1.AddItem "AND"
FormDynamicElements.ComboRelativeOperator1.AddItem "OR"
FormDynamicElements.ComboRelativeOperator2.AddItem "AND"
FormDynamicElements.ComboRelativeOperator2.AddItem "OR"
FormDynamicElements.ComboRelativeOperator1.ListIndex = DynamicSpecifiedElementByStoichiometryToAnotherOperator1%
FormDynamicElements.ComboRelativeOperator2.ListIndex = DynamicSpecifiedElementByStoichiometryToAnotherOperator2%
FormDynamicElements.ComboExcessOxygenByDroopOperator1.AddItem "AND"
FormDynamicElements.ComboExcessOxygenByDroopOperator1.AddItem "OR"
FormDynamicElements.ComboExcessOxygenByDroopOperator2.AddItem "AND"
FormDynamicElements.ComboExcessOxygenByDroopOperator2.AddItem "OR"
FormDynamicElements.ComboExcessOxygenByDroopOperator1.ListIndex = DynamicSpecifiedElementExcessOxygenByDroopOperator1%
FormDynamicElements.ComboExcessOxygenByDroopOperator2.ListIndex = DynamicSpecifiedElementExcessOxygenByDroopOperator2%

FormDynamicElements.Show vbModal

Exit Sub

' Errors
DynamicElementsLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "DynamicElementsLoad"
ierror = True
Exit Sub

End Sub

Sub DynamicElementsSave()
' Save the dynamic elements form options

ierror = False
On Error GoTo DynamicElementsSaveError

Dim nonzero As Boolean
Dim j As Integer

' Save difference selections
If FormDynamicElements.CheckDifference.value = vbChecked Then
DynamicSpecifiedElementByDifferenceFlag% = True
Else
DynamicSpecifiedElementByDifferenceFlag% = False
End If

' Check for all zero k-ratios
If DynamicSpecifiedElementByDifferenceFlag% = True Then
nonzero = False
For j% = 1 To MAXCRITERIA%
If Val(FormDynamicElements.TextDifferenceValue(j% - 1).Text) <> 0 Then nonzero = True
Next j%
If nonzero = False Then GoTo DynamicElementsSaveAllZero1
End If

For j% = 1 To MAXCRITERIA%
DynamicSpecifiedElementByDifferenceElement%(j%) = FormDynamicElements.ComboDifferenceElement(j% - 1).ListIndex + 1
DynamicSpecifiedElementByDifferenceGreaterLess(j%) = FormDynamicElements.ComboDifferenceGreaterLess(j% - 1).ListIndex
DynamicSpecifiedElementByDifferenceValue!(j%) = Val(FormDynamicElements.TextDifferenceValue(j% - 1).Text)
DynamicSpecifiedElementByDifferenceOperator1% = FormDynamicElements.ComboDifferenceOperator1.ListIndex
DynamicSpecifiedElementByDifferenceOperator2% = FormDynamicElements.ComboDifferenceOperator2.ListIndex
Next j%
If DynamicSpecifiedElementByDifferenceFlag% = True Then FormZAFOPT.CheckDifference.value = vbChecked        ' to force parameter check in FormZAFOPT

' Save difference by formula selections
If FormDynamicElements.CheckDifferenceFormula.value = vbChecked Then
DynamicSpecifiedElementByDifferenceFormulaFlag% = True
Else
DynamicSpecifiedElementByDifferenceFormulaFlag% = False
End If

' Check for all zero k-ratios
If DynamicSpecifiedElementByDifferenceFormulaFlag% = True Then
nonzero = False
For j% = 1 To MAXCRITERIA%
If Val(FormDynamicElements.TextDifferenceFormulaValue(j% - 1).Text) <> 0 Then nonzero = True
Next j%
If nonzero = False Then GoTo DynamicElementsSaveAllZero2
End If

For j% = 1 To MAXCRITERIA%
DynamicSpecifiedElementByDifferenceFormulaElement%(j%) = FormDynamicElements.ComboDifferenceFormulaElement(j% - 1).ListIndex + 1
DynamicSpecifiedElementByDifferenceFormulaGreaterLess(j%) = FormDynamicElements.ComboDifferenceFormulaGreaterLess(j% - 1).ListIndex
DynamicSpecifiedElementByDifferenceFormulaValue!(j%) = Val(FormDynamicElements.TextDifferenceFormulaValue(j% - 1).Text)
DynamicSpecifiedElementByDifferenceFormulaOperator1% = FormDynamicElements.ComboDifferenceFormulaOperator1.ListIndex
DynamicSpecifiedElementByDifferenceFormulaOperator2% = FormDynamicElements.ComboDifferenceFormulaOperator2.ListIndex
Next j%
If DynamicSpecifiedElementByDifferenceFormulaFlag% = True Then FormZAFOPT.CheckDifferenceFormula.value = vbChecked        ' to force parameter check in FormZAFOPT

' Save stoichiometry to oxygen selections
If FormDynamicElements.CheckStoichiometry.value = vbChecked Then
DynamicSpecifiedElementByStoichiometryToOxygenFlag% = True
Else
DynamicSpecifiedElementByStoichiometryToOxygenFlag% = False
End If

' Check for all zero k-ratios
If DynamicSpecifiedElementByStoichiometryToOxygenFlag% = True Then
nonzero = False
For j% = 1 To MAXCRITERIA%
If Val(FormDynamicElements.TextStoichiometryValue(j% - 1).Text) <> 0 Then nonzero = True
Next j%
If nonzero = False Then GoTo DynamicElementsSaveAllZero3
End If

For j% = 1 To MAXCRITERIA%
DynamicSpecifiedElementByStoichiometryToOxygenElement%(j%) = FormDynamicElements.ComboStoichiometryElement(j% - 1).ListIndex + 1
DynamicSpecifiedElementByStoichiometryToOxygenGreaterLess(j%) = FormDynamicElements.ComboStoichiometryGreaterLess(j% - 1).ListIndex
DynamicSpecifiedElementByStoichiometryToOxygenValue!(j%) = Val(FormDynamicElements.TextStoichiometryValue(j% - 1).Text)
DynamicSpecifiedElementByStoichiometryToOxygenOperator1% = FormDynamicElements.ComboStoichiometryOperator1.ListIndex
DynamicSpecifiedElementByStoichiometryToOxygenOperator2% = FormDynamicElements.ComboStoichiometryOperator2.ListIndex
Next j%
If DynamicSpecifiedElementByStoichiometryToOxygenFlag% = True Then FormZAFOPT.CheckStoichiometry.value = vbChecked        ' to force parameter check in FormZAFOPT

' Save stoichiometry to another element selections
If FormDynamicElements.CheckRelative.value = vbChecked Then
DynamicSpecifiedElementByStoichiometryToAnotherFlag% = True
Else
DynamicSpecifiedElementByStoichiometryToAnotherFlag% = False
End If

' Check for all zero k-ratios
If DynamicSpecifiedElementByStoichiometryToAnotherFlag% = True Then
nonzero = False
For j% = 1 To MAXCRITERIA%
If Val(FormDynamicElements.TextRelativeValue(j% - 1).Text) <> 0 Then nonzero = True
Next j%
If nonzero = False Then GoTo DynamicElementsSaveAllZero4
End If

For j% = 1 To MAXCRITERIA%
DynamicSpecifiedElementByStoichiometryToAnotherElement%(j%) = FormDynamicElements.ComboRelativeElement(j% - 1).ListIndex + 1
DynamicSpecifiedElementByStoichiometryToAnotherGreaterLess(j%) = FormDynamicElements.ComboRelativeGreaterLess(j% - 1).ListIndex
DynamicSpecifiedElementByStoichiometryToAnotherValue!(j%) = Val(FormDynamicElements.TextRelativeValue(j% - 1).Text)
DynamicSpecifiedElementByStoichiometryToAnotherOperator1% = FormDynamicElements.ComboRelativeOperator1.ListIndex
DynamicSpecifiedElementByStoichiometryToAnotherOperator2% = FormDynamicElements.ComboRelativeOperator2.ListIndex
Next j%
If DynamicSpecifiedElementByStoichiometryToAnotherFlag% = True Then FormZAFOPT.CheckRelative.value = vbChecked        ' to force parameter check in FormZAFOPT

' Save excess oxygen by Droop selections
If FormDynamicElements.CheckExcessOxygenByDroop.value = vbChecked Then
DynamicSpecifiedElementExcessOxygenByDroopFlag% = True
Else
DynamicSpecifiedElementExcessOxygenByDroopFlag% = False
End If

' Check for all zero k-ratios
If DynamicSpecifiedElementExcessOxygenByDroopFlag% = True Then
nonzero = False
For j% = 1 To MAXCRITERIA%
If Val(FormDynamicElements.TextExcessOxygenByDroopValue(j% - 1).Text) <> 0 Then nonzero = True
Next j%
If nonzero = False Then GoTo DynamicElementsSaveAllZero5
End If

For j% = 1 To MAXCRITERIA%
DynamicSpecifiedElementExcessOxygenByDroopElement%(j%) = FormDynamicElements.ComboExcessOxygenByDroopElement(j% - 1).ListIndex + 1
DynamicSpecifiedElementExcessOxygenByDroopGreaterLess(j%) = FormDynamicElements.ComboExcessOxygenByDroopGreaterLess(j% - 1).ListIndex
DynamicSpecifiedElementExcessOxygenByDroopValue!(j%) = Val(FormDynamicElements.TextExcessOxygenByDroopValue(j% - 1).Text)
DynamicSpecifiedElementExcessOxygenByDroopOperator1% = FormDynamicElements.ComboExcessOxygenByDroopOperator1.ListIndex
DynamicSpecifiedElementExcessOxygenByDroopOperator2% = FormDynamicElements.ComboExcessOxygenByDroopOperator2.ListIndex
Next j%
If DynamicSpecifiedElementExcessOxygenByDroopFlag% = True Then FormZAFOPT.CheckFerrousFerricCalculation.value = vbChecked        ' to force parameter check in FormZAFOPT

Exit Sub

' Errors
DynamicElementsSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "DynamicElementsSave"
ierror = True
Exit Sub

DynamicElementsSaveAllZero1:
msg$ = "All element k-ratio values are zero for the  Element by Difference dynamic method.  You must have at least one non-zero criteria to perform dynamic calculation options on pixels."
MsgBox msg$, vbOKOnly + vbExclamation, "DynamicElementsSave"
ierror = True
Exit Sub

DynamicElementsSaveAllZero2:
msg$ = "All element k-ratio values are zero for the Formula by Difference dynamic method.  You must have at least one non-zero criteria to perform dynamic calculation options on pixels."
MsgBox msg$, vbOKOnly + vbExclamation, "DynamicElementsSave"
ierror = True
Exit Sub

DynamicElementsSaveAllZero3:
msg$ = "All element k-ratio values are zero for Element by Stoichiometry to Stoichiometric Oxygen dynamic method.  You must have at least one non-zero criteria to perform dynamic calculation options on pixels."
MsgBox msg$, vbOKOnly + vbExclamation, "DynamicElementsSave"
ierror = True
Exit Sub

DynamicElementsSaveAllZero4:
msg$ = "All element k-ratio values are zero for the Element by Stoichiometry by Another Element dynamic method.  You must have at least one non-zero criteria to perform dynamic calculation options on pixels."
MsgBox msg$, vbOKOnly + vbExclamation, "DynamicElementsSave"
ierror = True
Exit Sub

DynamicElementsSaveAllZero5:
msg$ = "All element k-ratio values are zero for the Excess Oxygen by Droop dynamic method.  You must have at least one non-zero criteria to perform dynamic calculation options on pixels."
MsgBox msg$, vbOKOnly + vbExclamation, "DynamicElementsSave"
ierror = True
Exit Sub

End Sub
