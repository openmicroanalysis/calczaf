Attribute VB_Name = "CodeEMP"
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

Dim EmpElements(1 To MAXEMP%) As Integer
Dim EmpXrays(1 To MAXEMP%) As Integer
Dim EmpAbsorbers(1 To MAXEMP%) As Integer
Dim EmpValues(1 To MAXEMP%) As Single
Dim EmpStrings(1 To MAXEMP%) As String

Dim EmpReNormFactors(1 To MAXEMP%) As Single
Dim EmpReNormStandards(1 To MAXEMP%) As String

Sub EmpAddEmp()
' Add a MAC or APF value to the module array

ierror = False
On Error GoTo EmpAddEmpError

Dim empez As Integer, empxl As Integer, empaz As Integer
Dim empval As Single, empfactor As Single
Dim empstring As String, empstandard As String
Dim tlinenumber As Integer

' Save the renormalization factor and standard from user input
If EmpTypeFlag% = 2 Then
If Val(FormEMP.TextReNormalizeFactor.Text) <= 0# Then FormEMP.TextReNormalizeFactor.Text = MiscAutoFormat$(1#)
If Val(FormEMP.TextReNormalizeFactor.Text) > 2# Then GoTo EmpAddEmpOutOfRange
empfactor! = Val(FormEMP.TextReNormalizeFactor.Text)
empstandard$ = Trim$(FormEMP.TextReNormalizeStandard.Text)
End If

' Get the selected empirical value line number from available list
If FormEMP.ListAvailableEmp.ListCount < 1 Then Exit Sub
If FormEMP.ListAvailableEmp.ListIndex < 0 Then Exit Sub
tlinenumber% = FormEMP.ListAvailableEmp.ItemData(FormEMP.ListAvailableEmp.ListIndex)

' Get the EMP parameters for this line number (empfactor!, empstandard$ are from user input from dialog)
Call EmpGetEmpParameters(tlinenumber%, empez%, empxl%, empaz%, empval!, empstring$)
If ierror Then Exit Sub

' Check that element is not already loaded
Call EmpCheckEmp(empez%, empxl%, empaz%)
If ierror Then Exit Sub

' Renormalize if loading APF factors
If EmpTypeFlag% = 2 Then
empval! = empval! / empfactor!
End If

' Load module array
If EmpTypeFlag% = 2 And Trim$(empstandard$) <> vbNullString Then
Call EmpLoadArray(empez%, empxl%, empaz%, empval!, empstring$ & ", normalized to " & empstandard$, empfactor!, empstandard$)
If ierror Then Exit Sub
Else
Call EmpLoadArray(empez%, empxl%, empaz%, empval!, empstring$, empfactor!, empstandard$)
If ierror Then Exit Sub
End If

Exit Sub

' Errors
EmpAddEmpError:
MsgBox Error$, vbOKOnly + vbCritical, "EmpAddEmp"
ierror = True
Exit Sub

EmpAddEmpOutOfRange:
msg$ = "Re-Normalization factor must be greater than zero and less than 2 (default= 1.000)"
MsgBox msg$, vbOKOnly + vbExclamation, "EmpAddEmp"
ierror = True
Exit Sub

End Sub

Sub EmpCheckEmp(empez As Integer, empxl As Integer, empaz As Integer)
' Check if element is already loaded

ierror = False
On Error GoTo EmpCheckEmpError

Dim i As Integer

' Loop on all empirical values
For i% = 1 To MAXEMP%
If EmpElements%(i%) = empez% And EmpXrays%(i%) = empxl% And EmpAbsorbers%(i%) = empaz% Then GoTo EmpCheckEmpAlreadyLoaded
Next i%

Exit Sub

' Errors
EmpCheckEmpError:
MsgBox Error$, vbOKOnly + vbCritical, "EmpCheckEmp"
ierror = True
Exit Sub

EmpCheckEmpAlreadyLoaded:
msg$ = "Element, xray and absorber combination is already loaded"
MsgBox msg$, vbOKOnly + vbExclamation, "EmpCheckEmp"
ierror = True
Exit Sub

End Sub

Function EmpCheckAPF(syme As String, symx As String) As Boolean
' Check if element APF is already loaded globally for this emitter

ierror = False
On Error GoTo EmpCheckAPFError

Dim i As Integer, ip As Integer, ipp As Integer

' Get atomic and x-ray number
EmpCheckAPF = False
ip% = IPOS1(MAXELM%, syme$, Symlo$())
ipp% = IPOS1(MAXRAY% - 1, symx$, Xraylo$())
If ip% = 0 Or ipp% = 0 Then Exit Function

' Loop on all empirical values
For i% = 1 To MAXEMP%
If apfez%(i%) = ip% And apfxl%(i%) = ipp% Then EmpCheckAPF = True
Next i%

Exit Function

' Errors
EmpCheckAPFError:
MsgBox Error$, vbOKOnly + vbCritical, "EmpCheckAPF"
ierror = True
Exit Function

End Function

Sub EmpDeleteEmp()
' Delete a MAC or APF value from the current list

ierror = False
On Error GoTo EmpDeleteEmpError

Dim i As Integer

' Check for valid list item
If FormEMP.ListCurrentEmp.ListCount < 1 Then Exit Sub
If FormEMP.ListCurrentEmp.ListIndex < 0 Then Exit Sub

' Determine array row
i% = FormEMP.ListCurrentEmp.ItemData(FormEMP.ListCurrentEmp.ListIndex)
If i% < 1 Or i% > MAXEMP% Then GoTo EmpDeleteEmpBadRow

' Delete this array row
EmpElements%(i%) = 0
EmpXrays%(i%) = 0
EmpAbsorbers%(i%) = 0
EmpValues!(i%) = 0#
EmpStrings$(i%) = vbNullString

EmpReNormFactors!(i%) = 0#
EmpReNormStandards$(i%) = vbNullString

' Remove the selected empirical value from current list
FormEMP.ListCurrentEmp.RemoveItem FormEMP.ListCurrentEmp.ListIndex

Exit Sub

' Errors
EmpDeleteEmpError:
MsgBox Error$, vbOKOnly + vbCritical, "EmpDeleteEmp"
ierror = True
Exit Sub

EmpDeleteEmpBadRow:
msg$ = "Bad array row"
MsgBox msg$, vbOKOnly + vbExclamation, "EmpDeleteEmp"
ierror = True
Exit Sub

End Sub

Sub EmpGetEmpParameters(tlinenumber As Integer, empez As Integer, empxl As Integer, empaz As Integer, empval As Single, empstring As String)
' Get Emp parameters from ASCII disk file

ierror = False
On Error GoTo EmpGetEmpParametersError

Dim linecount As Integer, ip As Integer
Dim syme As String, symx As String, symA As String

' Open empirical file
If EmpTypeFlag% = 1 Then Open EmpMACFile$ For Input As #EMPFileNumber%
If EmpTypeFlag% = 2 Then Open EmpAPFFile$ For Input As #EMPFileNumber%

linecount% = 0
Do While Not EOF(EMPFileNumber%)

linecount% = linecount% + 1
Input #EMPFileNumber%, syme$, symx$, symA$, empval!, empstring$

' Check for bad values
ip% = IPOS1(MAXELM%, syme$, Symlo$())
If ip% = 0 Then GoTo EmpGetEmpParametersBadElement
empez% = ip%

ip% = IPOS1(MAXRAY% - 1, symx$, Xraylo$())
If ip% = 0 Then GoTo EmpGetEmpParametersBadElement
empxl% = ip%

ip% = IPOS1(MAXELM%, symA$, Symlo$())
If ip% = 0 Then GoTo EmpGetEmpParametersBadElement
empaz% = ip%

' If correct line number, just exit with data
If linecount% = tlinenumber% Then
Close #EMPFileNumber%
Exit Sub
End If

Loop
Close #EMPFileNumber%

' If we get to here, couldn't find "tlinenumber%"
GoTo EmpGetEmpParametersNotFound

Exit Sub

' Errors
EmpGetEmpParametersError:
MsgBox Error$, vbOKOnly + vbCritical, "EmpGetEmpParameters"
Close #EMPFileNumber%
ierror = True
Exit Sub

EmpGetEmpParametersBadElement:
If EmpTypeFlag% = 1 Then msg$ = "Bad element or x-ray symbol in " & EmpMACFile$ & " on line " & Str$(linecount%)
If EmpTypeFlag% = 2 Then msg$ = "Bad element or x-ray symbol in " & EmpAPFFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "EmpGetEmpParameters"
Close #EMPFileNumber%
ierror = True
Exit Sub

EmpGetEmpParametersNotFound:
If EmpTypeFlag% = 1 Then msg$ = "Could not find data for line " & Str$(tlinenumber%) & " in file " & EmpMACFile$
If EmpTypeFlag% = 1 Then msg$ = "Could not find data for line " & Str$(tlinenumber%) & " in file " & EmpAPFFile$
MsgBox msg$, vbOKOnly + vbExclamation, "EmpGetEmpParameters"
Close #EMPFileNumber%
ierror = True
Exit Sub

End Sub

Sub EmpLoad()
' Load EMP form (empirical MACs or APFs)

ierror = False
On Error GoTo EmpLoadError

Dim i As Integer

' Zero module level arrays
For i% = 1 To MAXEMP%
EmpElements%(i%) = 0
EmpXrays%(i%) = 0
EmpAbsorbers%(i%) = 0
EmpValues!(i%) = 0#
EmpStrings$(i%) = vbNullString
Next i%

' Set labels
If EmpTypeFlag% = 1 Then
FormEMP.Caption = "Add Empirical MACs (mass absorption coefficients) to Run"
FormEMP.LabelAvailable.Caption = "Available Empirical MACs from " & EmpMACFile$
FormEMP.LabelCurrent.Caption = "Current Empirical MACs in Run"
End If

If EmpTypeFlag% = 2 Then
FormEMP.Caption = "Add Empirical APFs (area peak factors) to Run"
FormEMP.LabelAvailable.Caption = "Available Empirical APFs from " & EmpAPFFile$
FormEMP.LabelCurrent.Caption = "Current Empirical APFs in Run"
End If

' Load available Empirical MAC or APF list box
If EmpTypeFlag% = 1 Then Open EmpMACFile$ For Input As #EMPFileNumber%
If EmpTypeFlag% = 2 Then Open EmpAPFFile$ For Input As #EMPFileNumber%
Call EmpLoadListAvailable
Close #EMPFileNumber%
If ierror Then Exit Sub

' Load current list box
FormEMP.ListCurrentEmp.Clear
For i% = 1 To MAXEMP%

' Load empirical MACs
If EmpTypeFlag% = 1 Then
If macez%(i%) > 0 Then
Call EmpLoadArray(macez%(i%), macxl%(i%), macaz%(i%), macval!(i%), macstr$(i%), macrenormfactor!(i%), macrenormstandard$(i%))
If ierror Then Exit Sub
End If

' Load empirical APFs
Else
If apfez%(i%) > 0 Then
Call EmpLoadArray(apfez%(i%), apfxl%(i%), apfaz%(i%), apfval!(i%), apfstr$(i%), apfrenormfactor!(i%), apfrenormstandard$(i%))
If ierror Then Exit Sub
End If
End If

Next i%

' If loading APF factors, make renormalize factor and standard string visible
If EmpTypeFlag% = 2 Then
FormEMP.LabelAPF.Visible = True
FormEMP.OLE2.Visible = True
FormEMP.LabelReNormalize.Visible = True
FormEMP.LabelReNormalizeFactor.Visible = True
FormEMP.TextReNormalizeFactor.Visible = True
FormEMP.LabelReNormalizeStandard.Visible = True
FormEMP.TextReNormalizeStandard.Visible = True

Else
FormEMP.LabelMAC.Visible = True
End If

Exit Sub

' Errors
EmpLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "EmpLoad"
Close #EMPFileNumber%
ierror = True
Exit Sub

End Sub

Sub EmpLoadArray(empez As Integer, empxl As Integer, empaz As Integer, empval As Single, empstring As String, empfactor As Single, empstandard As String)
' Load passed parameters to module array and add to list

ierror = False
On Error GoTo EmpLoadArrayError

Dim ip As Integer

' Find next free row in module array
ip% = IPOS2(MAXEMP%, 0, EmpElements%())
If ip% = 0 Then GoTo EmpLoadArrayTooMany

EmpElements%(ip%) = empez%
EmpXrays%(ip%) = empxl%
EmpAbsorbers%(ip%) = empaz%
EmpValues!(ip%) = empval!
EmpStrings$(ip%) = empstring$

EmpReNormFactors!(ip%) = empfactor!
EmpReNormStandards$(ip%) = empstandard$

' Add to "current" list
Call EmpLoadListCurrent(ip%, empez%, empxl%, empaz%, empval!, empstring$, empfactor!, empstandard$)
If ierror Then Exit Sub

Exit Sub

' Errors
EmpLoadArrayError:
MsgBox Error$, vbOKOnly + vbCritical, "EmpLoadArray"
ierror = True
Exit Sub

EmpLoadArrayTooMany:
If EmpTypeFlag% = 1 Then msg$ = "No more room in arrays to add another empirical MAC value"
If EmpTypeFlag% = 2 Then msg$ = "No more room in arrays to add another empirical APF value"
MsgBox msg$, vbOKOnly + vbExclamation, "EmpLoadArray"
ierror = True
Exit Sub

End Sub

Sub EmpLoadListAvailable()
' Loads the ListAvailable for FormEMP

ierror = False
On Error GoTo EmpLoadListAvailableError

Dim syme As String, symx As String, symA As String
Dim empval As Single
Dim empstring As String
Dim linecount As Integer, ip As Integer

FormEMP.ListAvailableEmp.Clear

linecount% = 0
Do While Not EOF(EMPFileNumber%)

linecount% = linecount% + 1
Input #EMPFileNumber%, syme$, symx$, symA$, empval!, empstring$

' Check for bad values
ip% = IPOS1(MAXELM%, syme$, Symlo$())
If ip% = 0 Then GoTo EmpLoadListAvailableBadElement

ip% = IPOS1(MAXRAY% - 1, symx$, Xraylo$())
If ip% = 0 Then GoTo EmpLoadListAvailableBadElement

ip% = IPOS1(MAXELM%, symA$, Symlo$())
If ip% = 0 Then GoTo EmpLoadListAvailableBadElement

msg$ = Format$(syme$, a20$) & " " & Format$(symx$, a20$) & " in " & Format$(symA$, a20$) & ", " & MiscAutoFormat$(empval!) & " " & empstring$
FormEMP.ListAvailableEmp.AddItem msg$
FormEMP.ListAvailableEmp.ItemData(FormEMP.ListAvailableEmp.NewIndex) = linecount%

Loop

Exit Sub

' Errors
EmpLoadListAvailableError:
MsgBox Error$, vbOKOnly + vbCritical, "EmpLoadListAvailable"
ierror = True
Exit Sub

EmpLoadListAvailableBadElement:
If EmpTypeFlag% = 1 Then msg$ = "Bad element or x-ray symbol in " & EmpMACFile$ & " on line " & Str$(linecount%)
If EmpTypeFlag% = 2 Then msg$ = "Bad element or x-ray symbol in " & EmpAPFFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "EmpLoadListAvailable"
ierror = True
Exit Sub

End Sub

Sub EmpLoadListCurrent(ip As Integer, empez As Integer, empxl As Integer, empaz As Integer, empval As Single, empstring As String, empfactor As Single, empstandard$)
' Loads the FormEMP.ListCurrentEmp based on passed parameters

ierror = False
On Error GoTo EmpLoadListCurrentError

msg$ = Format$(Symlo$(empez%), a20$) & " " & Format$(Xraylo$(empxl%), a20$) & " in " & Format$(Symlo$(empaz%), a20$) & ", " & MiscAutoFormat$(empval!) & " " & empstring$
FormEMP.ListCurrentEmp.AddItem msg$
FormEMP.ListCurrentEmp.ItemData(FormEMP.ListCurrentEmp.NewIndex) = ip%

FormEMP.TextReNormalizeFactor.Text = MiscAutoFormat$(empfactor!)
FormEMP.TextReNormalizeStandard.Text = Trim$(empstandard$)

Exit Sub

' Errors
EmpLoadListCurrentError:
MsgBox Error$, vbOKOnly + vbCritical, "EmpLoadListCurrent"
ierror = True
Exit Sub

End Sub

Sub EmpSave()
' Save FormEMP parameters to global arrays

ierror = False
On Error GoTo EmpSaveError

Dim i As Integer, n As Integer

' Zero global array
For i% = 1 To MAXEMP%
If EmpTypeFlag% = 1 Then
macez%(i%) = 0
macxl%(i%) = 0
macaz%(i%) = 0
macval!(i%) = 0#
macstr$(i%) = vbNullString
macrenormfactor!(i%) = 0#
macrenormstandard$(i%) = vbNullString
End If

If EmpTypeFlag% = 2 Then
apfez%(i%) = 0
apfxl%(i%) = 0
apfaz%(i%) = 0
apfval!(i%) = 0#
apfstr$(i%) = vbNullString
apfrenormfactor!(i%) = 0#
apfrenormstandard$(i%) = vbNullString
End If
Next i%

' Load global array
n% = 0
For i% = 1 To MAXEMP%

' Check if non-zero element
If EmpElements%(i%) > 0 Then
n% = n% + 1

' Load MACs
If EmpTypeFlag% = 1 Then
macez%(n%) = EmpElements%(i%)
macxl%(n%) = EmpXrays%(i%)
macaz%(n%) = EmpAbsorbers%(i%)
macval!(n%) = EmpValues!(i%)
macstr$(n%) = EmpStrings$(i%)

macrenormfactor!(i%) = EmpReNormFactors!(i%)
macrenormstandard$(i%) = EmpReNormStandards$(i%)
UseMACFlag = True    ' set flag
End If

' Load APFs
If EmpTypeFlag% = 2 Then
apfez%(n%) = EmpElements%(i%)
apfxl%(n%) = EmpXrays%(i%)
apfaz%(n%) = EmpAbsorbers%(i%)
apfval!(n%) = EmpValues!(i%)
apfstr$(n%) = EmpStrings$(i%)

apfrenormfactor!(i%) = EmpReNormFactors!(i%)
apfrenormstandard$(i%) = EmpReNormStandards$(i%)
UseAPFFlag = True    ' set flag
End If

End If
Next i%

' Force reload of a-factor or ZAF arrays
AllAFactorUpdateNeeded = True
AllAnalysisUpdateNeeded = True

Exit Sub

' Errors
EmpSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "EmpSave"
ierror = True
Exit Sub

End Sub

Sub EmpLoadReNormalization()
' Load the re-normalization factor and standard for the selected empirical APF (not used for empirical MACs)

ierror = False
On Error GoTo EmpLoadReNormalizationError

Dim ip As Integer

If FormEMP.ListCurrentEmp.ListCount < 1 Then Exit Sub
If FormEMP.ListCurrentEmp.ListIndex < 0 Then Exit Sub
ip% = FormEMP.ListCurrentEmp.ItemData(FormEMP.ListCurrentEmp.ListIndex)

' Get user selected row
FormEMP.ListCurrentEmp.ItemData(FormEMP.ListCurrentEmp.ListIndex) = ip%

FormEMP.TextReNormalizeFactor.Text = MiscAutoFormat$(EmpReNormFactors!(ip%))
FormEMP.TextReNormalizeStandard.Text = Trim$(EmpReNormStandards$(ip%))

Exit Sub

' Errors
EmpLoadReNormalizationError:
MsgBox Error$, vbOKOnly + vbCritical, "EmpLoadReNormalization"
ierror = True
Exit Sub

End Sub
