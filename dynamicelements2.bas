Attribute VB_Name = "CodeDynamicElements2"
' (c) Copyright 1995-2025 by John J. Donovan
Option Explicit

Sub DynamicElementsCriteria(zaf As TypeZAF, sample() As TypeSample)
' Check the dynamic elements criteria parameters against k-ratio values (only used by CalcImage for pixel calculations)

ierror = False
On Error GoTo DynamicElementsCriteriaError

Dim j As Integer
Dim btrue() As Boolean

' Element by difference
ReDim btrue(1 To MAXCRITERIA%) As Boolean           ' re-set to zero
If DynamicSpecifiedElementByDifferenceFlag% Then
sample(1).DifferenceElementFlag% = False
For j% = 1 To MAXCRITERIA%
If DynamicSpecifiedElementByDifferenceValue!(j%) <> 0# Then
If DynamicSpecifiedElementByDifferenceGreaterLess%(j%) = 0 Then
If zaf.krat!(DynamicSpecifiedElementByDifferenceElement%(j%)) > DynamicSpecifiedElementByDifferenceValue!(j%) Then btrue(j%) = True
Else
If zaf.krat!(DynamicSpecifiedElementByDifferenceElement%(j%)) < DynamicSpecifiedElementByDifferenceValue!(j%) Then btrue(j%) = True
End If
Else
btrue(j%) = True        ' k-ratio is zero, so set true
End If
Next j%

If DynamicSpecifiedElementByDifferenceOperator1% = 0 And DynamicSpecifiedElementByDifferenceOperator2% = 0 Then
If btrue(1) And btrue(2) And btrue(3) Then sample(1).DifferenceElementFlag% = True
ElseIf DynamicSpecifiedElementByDifferenceOperator1% = 1 And DynamicSpecifiedElementByDifferenceOperator2% = 0 Then
If btrue(1) Or btrue(2) And btrue(3) Then sample(1).DifferenceElementFlag% = True
ElseIf DynamicSpecifiedElementByDifferenceOperator1% = 0 And DynamicSpecifiedElementByDifferenceOperator2% = 1 Then
If btrue(1) And btrue(2) Or btrue(3) Then sample(1).DifferenceElementFlag% = True
ElseIf DynamicSpecifiedElementByDifferenceOperator1% = 1 And DynamicSpecifiedElementByDifferenceOperator2% = 1 Then
If btrue(1) Or btrue(2) Or btrue(3) Then sample(1).DifferenceElementFlag% = True
End If
End If

' Formula by difference
ReDim btrue(1 To MAXCRITERIA%) As Boolean           ' re-set to zero
If DynamicSpecifiedElementByDifferenceFormulaFlag% Then
sample(1).DifferenceFormulaFlag% = False
For j% = 1 To MAXCRITERIA%
If DynamicSpecifiedElementByDifferenceFormulaValue!(j%) <> 0# Then
If DynamicSpecifiedElementByDifferenceFormulaGreaterLess%(j%) = 0 Then
If zaf.krat!(DynamicSpecifiedElementByDifferenceFormulaElement%(j%)) > DynamicSpecifiedElementByDifferenceFormulaValue!(j%) Then btrue(j%) = True
Else
If zaf.krat!(DynamicSpecifiedElementByDifferenceFormulaElement%(j%)) < DynamicSpecifiedElementByDifferenceFormulaValue!(j%) Then btrue(j%) = True
End If
Else
btrue(j%) = True        ' k-ratio is zero, so set true
End If
Next j%

If DynamicSpecifiedElementByDifferenceFormulaOperator1% = 0 And DynamicSpecifiedElementByDifferenceFormulaOperator2% = 0 Then
If btrue(1) And btrue(2) And btrue(3) Then sample(1).DifferenceFormulaFlag% = True
ElseIf DynamicSpecifiedElementByDifferenceFormulaOperator1% = 1 And DynamicSpecifiedElementByDifferenceFormulaOperator2% = 0 Then
If btrue(1) Or btrue(2) And btrue(3) Then sample(1).DifferenceFormulaFlag% = True
ElseIf DynamicSpecifiedElementByDifferenceFormulaOperator1% = 0 And DynamicSpecifiedElementByDifferenceFormulaOperator2% = 1 Then
If btrue(1) And btrue(2) Or btrue(3) Then sample(1).DifferenceFormulaFlag% = True
ElseIf DynamicSpecifiedElementByDifferenceFormulaOperator1% = 1 And DynamicSpecifiedElementByDifferenceFormulaOperator2% = 1 Then
If btrue(1) Or btrue(2) Or btrue(3) Then sample(1).DifferenceFormulaFlag% = True
End If
End If

' Element by stoichiometry to stoichiometric oxygen
ReDim btrue(1 To MAXCRITERIA%) As Boolean           ' re-set to zero
If DynamicSpecifiedElementByStoichiometryToOxygenFlag% Then
sample(1).StoichiometryElementFlag = False
For j% = 1 To MAXCRITERIA%
If DynamicSpecifiedElementByStoichiometryToOxygenValue!(j%) <> 0# Then
If DynamicSpecifiedElementByStoichiometryToOxygenGreaterLess%(j%) = 0 Then
If zaf.krat!(DynamicSpecifiedElementByStoichiometryToOxygenElement%(j%)) > DynamicSpecifiedElementByStoichiometryToOxygenValue!(j%) Then btrue(j%) = True
Else
If zaf.krat!(DynamicSpecifiedElementByStoichiometryToOxygenElement%(j%)) < DynamicSpecifiedElementByStoichiometryToOxygenValue!(j%) Then btrue(j%) = True
End If
Else
btrue(j%) = True        ' k-ratio is zero, so set true
End If
Next j%

If DynamicSpecifiedElementByStoichiometryToOxygenOperator1% = 0 And DynamicSpecifiedElementByStoichiometryToOxygenOperator2% = 0 Then
If btrue(1) And btrue(2) And btrue(3) Then sample(1).StoichiometryElementFlag = True
ElseIf DynamicSpecifiedElementByStoichiometryToOxygenOperator1% = 1 And DynamicSpecifiedElementByStoichiometryToOxygenOperator2% = 0 Then
If btrue(1) Or btrue(2) And btrue(3) Then sample(1).StoichiometryElementFlag = True
ElseIf DynamicSpecifiedElementByStoichiometryToOxygenOperator1% = 0 And DynamicSpecifiedElementByStoichiometryToOxygenOperator2% = 1 Then
If btrue(1) And btrue(2) Or btrue(3) Then sample(1).StoichiometryElementFlag = True
ElseIf DynamicSpecifiedElementByStoichiometryToOxygenOperator1% = 1 And DynamicSpecifiedElementByStoichiometryToOxygenOperator2% = 1 Then
If btrue(1) Or btrue(2) Or btrue(3) Then sample(1).StoichiometryElementFlag = True
End If
End If

' Element relative to another element
ReDim btrue(1 To MAXCRITERIA%) As Boolean           ' re-set to zero
If DynamicSpecifiedElementByStoichiometryToAnotherFlag% Then
sample(1).RelativeElementFlag% = False
For j% = 1 To MAXCRITERIA%
If DynamicSpecifiedElementByStoichiometryToAnotherValue!(j%) <> 0# Then
If DynamicSpecifiedElementByStoichiometryToAnotherGreaterLess%(j%) = 0 Then
If zaf.krat!(DynamicSpecifiedElementByStoichiometryToAnotherElement%(j%)) > DynamicSpecifiedElementByStoichiometryToAnotherValue!(j%) Then btrue(j%) = True
Else
If zaf.krat!(DynamicSpecifiedElementByStoichiometryToAnotherElement%(j%)) < DynamicSpecifiedElementByStoichiometryToAnotherValue!(j%) Then btrue(j%) = True
End If
Else
btrue(j%) = True        ' k-ratio is zero, so set true
End If
Next j%

If DynamicSpecifiedElementByStoichiometryToAnotherOperator1% = 0 And DynamicSpecifiedElementByStoichiometryToAnotherOperator2% = 0 Then
If btrue(1) And btrue(2) And btrue(3) Then sample(1).RelativeElementFlag = True
ElseIf DynamicSpecifiedElementByStoichiometryToAnotherOperator1% = 1 And DynamicSpecifiedElementByStoichiometryToAnotherOperator2% = 0 Then
If btrue(1) Or btrue(2) And btrue(3) Then sample(1).RelativeElementFlag = True
ElseIf DynamicSpecifiedElementByStoichiometryToAnotherOperator1% = 0 And DynamicSpecifiedElementByStoichiometryToAnotherOperator2% = 1 Then
If btrue(1) And btrue(2) Or btrue(3) Then sample(1).RelativeElementFlag = True
ElseIf DynamicSpecifiedElementByStoichiometryToAnotherOperator1% = 1 And DynamicSpecifiedElementByStoichiometryToAnotherOperator2% = 1 Then
If btrue(1) Or btrue(2) Or btrue(3) Then sample(1).RelativeElementFlag = True
End If
End If

' Excess oxygen by Droop
ReDim btrue(1 To MAXCRITERIA%) As Boolean           ' re-set to zero
If DynamicSpecifiedElementExcessOxygenByDroopFlag% Then
sample(1).FerrousFerricCalculationFlag = False
For j% = 1 To MAXCRITERIA%
If DynamicSpecifiedElementExcessOxygenByDroopValue!(j%) <> 0# Then
If DynamicSpecifiedElementExcessOxygenByDroopGreaterLess%(j%) = 0 Then
If zaf.krat!(DynamicSpecifiedElementExcessOxygenByDroopElement%(j%)) > DynamicSpecifiedElementExcessOxygenByDroopValue!(j%) Then btrue(j%) = True
Else
If zaf.krat!(DynamicSpecifiedElementExcessOxygenByDroopElement%(j%)) < DynamicSpecifiedElementExcessOxygenByDroopValue!(j%) Then btrue(j%) = True
End If
Else
btrue(j%) = True        ' k-ratio is zero, so set true
End If
Next j%

If DynamicSpecifiedElementExcessOxygenByDroopOperator1% = 0 And DynamicSpecifiedElementExcessOxygenByDroopOperator2% = 0 Then
If btrue(1) And btrue(2) And btrue(3) Then sample(1).FerrousFerricCalculationFlag = True
ElseIf DynamicSpecifiedElementExcessOxygenByDroopOperator1% = 1 And DynamicSpecifiedElementExcessOxygenByDroopOperator2% = 0 Then
If btrue(1) Or btrue(2) And btrue(3) Then sample(1).FerrousFerricCalculationFlag = True
ElseIf DynamicSpecifiedElementExcessOxygenByDroopOperator1% = 0 And DynamicSpecifiedElementExcessOxygenByDroopOperator2% = 1 Then
If btrue(1) And btrue(2) Or btrue(3) Then sample(1).FerrousFerricCalculationFlag = True
ElseIf DynamicSpecifiedElementExcessOxygenByDroopOperator1% = 1 And DynamicSpecifiedElementExcessOxygenByDroopOperator2% = 1 Then
If btrue(1) Or btrue(2) Or btrue(3) Then sample(1).FerrousFerricCalculationFlag = True
End If
End If

Exit Sub

' Errors
DynamicElementsCriteriaError:
MsgBox Error$, vbOKOnly + vbCritical, "DynamicElementsCriteria"
ierror = True
Exit Sub

End Sub

Sub DynamicElementsCheckSelections(mode As Integer, sample() As TypeSample)
' Sanity check for dynamic element selections
' mode = 0 silent
' mode = 1 warn user

ierror = False
On Error GoTo DynamicElementsCheckSelectionsError

Dim j As Integer

' Check element channels
For j% = 1 To MAXCRITERIA%
If DynamicSpecifiedElementByDifferenceElement%(j%) < 1 Then
DynamicSpecifiedElementByDifferenceElement%(j%) = 1
If mode% = 1 Then GoTo DynamicElementsCheckSelectionsDifference
End If
If DynamicSpecifiedElementByDifferenceElement%(j%) > sample(1).LastElm% Then
DynamicSpecifiedElementByDifferenceElement%(j%) = sample(1).LastElm%
If mode% = 1 Then GoTo DynamicElementsCheckSelectionsDifference
End If

If DynamicSpecifiedElementByDifferenceFormulaElement%(j%) < 1 Then
DynamicSpecifiedElementByDifferenceFormulaElement%(j%) = 1
If mode% = 1 Then GoTo DynamicElementsCheckSelectionsFormula
End If
If DynamicSpecifiedElementByDifferenceFormulaElement%(j%) > sample(1).LastElm% Then
DynamicSpecifiedElementByDifferenceFormulaElement%(j%) = sample(1).LastElm%
If mode% = 1 Then GoTo DynamicElementsCheckSelectionsFormula
End If

If DynamicSpecifiedElementByStoichiometryToOxygenElement%(j%) < 1 Then
DynamicSpecifiedElementByStoichiometryToOxygenElement%(j%) = 1
If mode% = 1 Then GoTo DynamicElementsCheckSelectionsStoichiometry
End If
If DynamicSpecifiedElementByStoichiometryToOxygenElement%(j%) > sample(1).LastElm% Then
DynamicSpecifiedElementByStoichiometryToOxygenElement%(j%) = sample(1).LastElm%
If mode% = 1 Then GoTo DynamicElementsCheckSelectionsStoichiometry
End If

If DynamicSpecifiedElementByStoichiometryToAnotherElement%(j%) < 1 Then
DynamicSpecifiedElementByStoichiometryToAnotherElement%(j%) = 1
If mode% = 1 Then GoTo DynamicElementsCheckSelectionsRelative
End If
If DynamicSpecifiedElementByStoichiometryToAnotherElement%(j%) > sample(1).LastElm% Then
DynamicSpecifiedElementByStoichiometryToAnotherElement%(j%) = sample(1).LastElm%
If mode% = 1 Then GoTo DynamicElementsCheckSelectionsRelative
End If

If DynamicSpecifiedElementExcessOxygenByDroopElement%(j%) < 1 Then
DynamicSpecifiedElementExcessOxygenByDroopElement%(j%) = 1
If mode% = 1 Then GoTo DynamicElementsCheckSelectionsDroop
End If
If DynamicSpecifiedElementExcessOxygenByDroopElement%(j%) > sample(1).LastElm% Then
DynamicSpecifiedElementExcessOxygenByDroopElement%(j%) = sample(1).LastElm%
If mode% = 1 Then GoTo DynamicElementsCheckSelectionsDifference
End If
Next j%

Exit Sub

' Errors
DynamicElementsCheckSelectionsError:
MsgBox Error$, vbOKOnly + vbCritical, "DynamicElementsCheckSelections"
ierror = True
Exit Sub

DynamicElementsCheckSelectionsDifference:
msg$ = "The dynamic element by difference selections are out of range for the specified sample setup." & vbCrLf & vbCrLf
msg$ = msg$ & "Please check the selections by clicking the Calculate Elements Dynamically button from the Calculation Options dialog."
MsgBox msg$, vbOKOnly + vbExclamation, "DynamicElementsCheckSelections"
ierror = True
Exit Sub

DynamicElementsCheckSelectionsFormula:
msg$ = "The dynamic element by difference by formula selections are out of range for the specified sample setup." & vbCrLf & vbCrLf
msg$ = msg$ & "Please check the selections by clicking the Calculate Elements Dynamically button from the Calculation Options dialog."
MsgBox msg$, vbOKOnly + vbExclamation, "DynamicElementsCheckSelections"
ierror = True
Exit Sub

DynamicElementsCheckSelectionsStoichiometry:
msg$ = "The dynamic element by stoichiometry to oxygen selections are out of range for the specified sample setup." & vbCrLf & vbCrLf
msg$ = msg$ & "Please check the selections by clicking the Calculate Elements Dynamically button from the Calculation Options dialog."
MsgBox msg$, vbOKOnly + vbExclamation, "DynamicElementsCheckSelections"
ierror = True
Exit Sub

DynamicElementsCheckSelectionsRelative:
msg$ = "The dynamic element by stoichiomtry to another element selections are out of range for the specified sample setup." & vbCrLf & vbCrLf
msg$ = msg$ & "Please check the selections by clicking the Calculate Elements Dynamically button from the Calculation Options dialog."
MsgBox msg$, vbOKOnly + vbExclamation, "DynamicElementsCheckSelections"
ierror = True
Exit Sub

DynamicElementsCheckSelectionsDroop:
msg$ = "The dynamic element excess oxygen by Droop selections are out of range for the specified sample setup." & vbCrLf & vbCrLf
msg$ = msg$ & "Please check the selections by clicking the Calculate Elements Dynamically button from the Calculation Options dialog."
MsgBox msg$, vbOKOnly + vbExclamation, "DynamicElementsCheckSelections"
ierror = True
Exit Sub

End Sub
