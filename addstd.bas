Attribute VB_Name = "CodeADDSTD"
' (c) Copyright 1995-2017 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Dim AddStdNumberToAdd As Integer
Dim AddStdStandardsToAdd(1 To MAXSTD%) As Integer

Dim tNumberofStandards As Integer
Dim tStandardNumbers(1 To MAXSTD%) As Integer

Sub AddStdAdd()
' Routine to add selected standard(s) to the "add to" list

ierror = False
On Error GoTo AddStdAddError

Dim i As Integer, number As Integer

' Get the selected standard(s)
For i% = 0 To FormADDSTD.ListAvailableStandards.ListCount - 1

' Check to see if standard is selected
If FormADDSTD.ListAvailableStandards.Selected(i%) Then

' See if standard is already in the run and add to "add to" list if not
number% = FormADDSTD.ListAvailableStandards.ItemData(i%)
Call AddStdCheck(number%)
If ierror Then Exit Sub

End If
Next i%

Exit Sub

' Errors
AddStdAddError:
MsgBox Error$, vbOKOnly + vbCritical, "AddStdAdd"
ierror = True
Exit Sub

End Sub

Sub AddStdAddToList(number As Integer)
' Add standard to "current" list box

ierror = False
On Error GoTo AddStdAddToListError

Dim ip As Integer

' Check for standard in available standards index
ip% = StandardGetRow%(number%)
If ip% = 0 Then GoTo AddStdAddToListNotFound

msg$ = Format$(StandardIndexNumbers%(ip%), a40) & " " & StandardIndexNames$(ip%)
FormADDSTD.ListCurrentStandards.AddItem msg$
FormADDSTD.ListCurrentStandards.ItemData(FormADDSTD.ListCurrentStandards.NewIndex) = number%

' Update label field
FormADDSTD.LabelNumberOfStds.Caption = Format$(FormADDSTD.ListCurrentStandards.ListCount)

Exit Sub

' Errors
AddStdAddToListError:
MsgBox Error$, vbOKOnly + vbCritical, "AddStdAddToList"
ierror = True
Exit Sub

AddStdAddToListNotFound:
msg$ = "Standard number " & Format$(number%) & " was not found in " & StandardDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "AddStdAddToList"
ierror = True
Exit Sub

End Sub

Sub AddStdCheck(number As Integer)
' Check if a standard can be added to the run

ierror = False
On Error GoTo AddStdCheckError

Dim ip As Integer

' Check if standard is already in the run
ip% = IPOS2(NumberofStandards%, number%, StandardNumbers%())
If ip% > 0 Then GoTo AddStdCheckAlreadyAdded

' Check if standard is already in AddStdStandardsToAdd array
ip% = IPOS2(AddStdNumberToAdd%, number%, AddStdStandardsToAdd%())
If ip% > 0 Then GoTo AddStdCheckAlreadyAdded

' Check for standard in available standards index
ip% = StandardGetRow%(number%)
If ip% = 0 Then GoTo AddStdCheckNotFound

' Check if too many standards
If NumberofStandards% + AddStdNumberToAdd% + 1 > MAXSTD% Then GoTo AddStdCheckTooMany

' Add to "add to" list
AddStdNumberToAdd% = AddStdNumberToAdd% + 1
AddStdStandardsToAdd%(AddStdNumberToAdd%) = number%

' Add to "current" list box
Call AddStdAddToList(number%)
If ierror Then Exit Sub

Exit Sub

' Errors
AddStdCheckError:
MsgBox Error$, vbOKOnly + vbCritical, "AddStdCheck"
ierror = True
Exit Sub

AddStdCheckAlreadyAdded:
msg$ = "Standard number " & Format$(number%) & " is already in the run"
MsgBox msg$, vbOKOnly + vbExclamation, "AddStdCheck"
ierror = True
Exit Sub

AddStdCheckTooMany:
msg$ = "Too many standards (" & Format$(NumberofStandards% + AddStdNumberToAdd%) & ") are already in the run"
MsgBox msg$, vbOKOnly + vbExclamation, "AddStdCheck"
ierror = True
Exit Sub

AddStdCheckNotFound:
msg$ = "Standard number " & Format$(number%) & " was not found in " & StandardDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "AddStdCheck"
ierror = True
Exit Sub

End Sub

Sub AddStdLoad()
' Routine to load the ADDSTD form

ierror = False
On Error GoTo AddStdLoadError

Dim i As Integer

' List the current standards in the run
FormADDSTD.ListCurrentStandards.Clear
tNumberofStandards% = NumberofStandards%
For i% = 1 To NumberofStandards%
tStandardNumbers%(i%) = StandardNumbers%(i%)    ' save in case user clicks cancel
Call AddStdAddToList(StandardNumbers%(i%))
If ierror Then Exit Sub
Next i%

' Get available standard names and numbers from database
Call StandardGetMDBIndex
If ierror Then Exit Sub

' List the available standards
Call StandardLoadList(FormADDSTD.ListAvailableStandards)
If ierror Then Exit Sub

' Zero add to list (for new standards)
AddStdNumberToAdd% = 0
For i% = 1 To MAXSTD%
AddStdStandardsToAdd%(i%) = 0
Next i%

Exit Sub

' Errors
AddStdLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "AddStdLoad"
ierror = True
Exit Sub

End Sub

Sub AddStdSave()
' Add standards in "add to" list to the run

ierror = False
On Error GoTo AddStdSaveError

Dim i As Integer, number As Integer

' Loop on each standard to add
For i% = 1 To AddStdNumberToAdd%
number% = AddStdStandardsToAdd%(i%)
If number% > 0 Then
Call AddStdSaveStd(number%)
If ierror Then Exit Sub
End If
Next i%

' Zero add to list (for new standards)
AddStdNumberToAdd% = 0
For i% = 1 To MAXSTD%
AddStdStandardsToAdd%(i%) = 0
Next i%

Exit Sub

' Errors
AddStdSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "AddStdSave"
ierror = True
Exit Sub

End Sub

Sub AddStdSaveStd(number As Integer)
' Add a single standard to the run

ierror = False
On Error GoTo AddStdSaveStdError

Dim ip As Integer

' See if standard  is already added
ip% = IPOS2(NumberofStandards%, number%, StandardNumbers%())
If ip% > 0 Then GoTo AddStdSaveStdAlreadyAdded

' Find standard in available standard.mdb index
ip% = StandardGetRow%(number%)
If ip% = 0 Then GoTo AddStdSaveStdNotFound

If NumberofStandards% + 1 > MAXSTD% Then GoTo AddStdSaveStdTooMany
NumberofStandards% = NumberofStandards% + 1
StandardNumbers%(NumberofStandards%) = StandardIndexNumbers%(ip%)
StandardNames$(NumberofStandards%) = StandardIndexNames$(ip%)
StandardDescriptions$(NumberofStandards%) = StandardIndexDescriptions$(ip%)
StandardDensities!(NumberofStandards%) = StandardIndexDensities!(ip%)

StandardCoatingFlag%(NumberofStandards%) = DefaultStandardCoatingFlag%    ' 0 = not coated, 1 = coated
StandardCoatingElement%(NumberofStandards%) = DefaultStandardCoatingElement%
StandardCoatingDensity!(NumberofStandards%) = DefaultStandardCoatingDensity!
StandardCoatingThickness!(NumberofStandards%) = DefaultStandardCoatingThickness!    ' in angstroms

Exit Sub

' Errors
AddStdSaveStdError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "AddStdSaveStd"
ierror = True
Exit Sub

AddStdSaveStdAlreadyAdded:
Screen.MousePointer = vbDefault
msg$ = "Standard number " & Format$(number%) & " is already in the run"
MsgBox msg$, vbOKOnly + vbExclamation, "AddStdSaveStd"
ierror = True
Exit Sub

AddStdSaveStdTooMany:
Screen.MousePointer = vbDefault
msg$ = "Too many standards are already in the run"
MsgBox msg$, vbOKOnly + vbExclamation, "AddStdSaveStd"
ierror = True
Exit Sub

AddStdSaveStdNotFound:
Screen.MousePointer = vbDefault
msg$ = "Standard number " & Format$(number%) & " was not found in " & StandardDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "AddStdSaveStd"
ierror = True
Exit Sub

End Sub

Sub AddStdRemove()
' Remove the selected standard if it is not already acquired

ierror = False
On Error GoTo AddStdRemoveError

Dim ip As Integer, number As Integer
Dim i As Integer

ReDim tstdnums(1 To MAXSTD%) As Integer

' Get the selected standard(s)
If FormADDSTD.ListCurrentStandards.ListIndex < 0 Then Exit Sub
If FormADDSTD.ListCurrentStandards.ListCount < 1 Then Exit Sub

number% = FormADDSTD.ListCurrentStandards.ItemData(FormADDSTD.ListCurrentStandards.ListIndex)

' See if standard is just in the add to list, if so just zero and exit
ip% = IPOS2(AddStdNumberToAdd%, number%, AddStdStandardsToAdd%())
If ip% > 0 Then
AddStdStandardsToAdd%(ip%) = 0
FormADDSTD.ListCurrentStandards.RemoveItem FormADDSTD.ListCurrentStandards.ListIndex

' Update label field
FormADDSTD.LabelNumberOfStds.Caption = Format$(FormADDSTD.ListCurrentStandards.ListCount)
Exit Sub
End If

' Now check if the standard is acquired in current run
ip% = SampleGetRow2(number%, Int(1), Int(1))
If ip% > 0 Then GoTo AddStdRemoveAlreadyAcquired

' Warn user if probe data file is open (not Stage.exe)
If ProbeDataFile$ <> vbNullString Then
msg$ = "Although standard " & Format$(number%) & " can be removed from the "
msg$ = msg$ & "standard list, the user should be aware that if this standard "
msg$ = msg$ & "is referenced in the current probe database, those references "
msg$ = msg$ & "must be changed to another suitable standard. This includes "
msg$ = msg$ & "assignments for standards, interference standards and MAN (mean "
msg$ = msg$ & "atomic number) background standards."
MsgBox msg$, vbOKOnly + vbInformation, "AddStdRemove"
End If

' Remove from list
FormADDSTD.ListCurrentStandards.RemoveItem FormADDSTD.ListCurrentStandards.ListIndex

' Update label field
FormADDSTD.LabelNumberOfStds.Caption = Format$(FormADDSTD.ListCurrentStandards.ListCount)

' Remove from standard numbers
For i% = 1 To NumberofStandards%
If StandardNumbers%(i%) <> number% Then
tstdnums%(i%) = StandardNumbers%(i%)
End If
Next i%

' Zero standards
Call InitStandard
If ierror Then Exit Sub

' Reload
For i% = 1 To MAXSTD%
If tstdnums%(i%) <> 0 Then
Call AddStdSaveStd(tstdnums%(i%))
End If
Next i%

Exit Sub

' Errors
AddStdRemoveError:
MsgBox Error$, vbOKOnly + vbCritical, "AddStdRemove"
ierror = True
Exit Sub

AddStdRemoveAlreadyAcquired:
msg$ = "Standard " & Format$(number%) & " cannot be removed from the standard list because it is already referenced in the current probe database"
MsgBox msg$, vbOKOnly + vbExclamation, "AddStdRemove"
ierror = True
Exit Sub

End Sub

Sub AddStdCancel()
' Reload the original standards

ierror = False
On Error GoTo AddStdCancelError

Dim i As Integer

' Zero standards
Call InitStandard
If ierror Then Exit Sub

' Reload from temporary list
For i% = 1 To tNumberofStandards%
Call AddStdSaveStd(tStandardNumbers%(i%))
Next i%

Exit Sub

' Errors
AddStdCancelError:
MsgBox Error$, vbOKOnly + vbCritical, "AddStdCancel"
ierror = True
Exit Sub

End Sub
