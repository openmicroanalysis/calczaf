Attribute VB_Name = "CodeMODAL"
' (c) Copyright 1995-2023 by John J. Donovan
Option Explicit

Global Const MAXGROUP% = 48             ' maximum number of groups
Global Const MAXPHASE% = 32             ' maximum number of phases

Global ModalInputDataFile As String
Global ModalOutputDataFile As String

' Modal analysis array
Public Type TypeModalGroup
    GroupNumber As Integer
    GroupName As String
    MinimumTotal As Single
    NormalizeFlag As Integer
    weightflag As Integer
    DoEndMember As Integer
    NumberofPhases As Integer
    PhaseNames(1 To MAXPHASE%) As String
    MinimumVectors(1 To MAXPHASE%) As Single
    EndMemberNumbers(1 To MAXPHASE%) As Integer
    NumberofStandards(1 To MAXPHASE%) As Integer
    StandardNumbers(1 To MAXPHASE%, 1 To MAXSTD%) As Integer
End Type

Dim defaultminimumtotal As Single
Dim defaultminimumvector As Single

Dim defaultnormalizeflag As Integer
Dim defaultweightflag As Integer

Dim endmembernames(0 To MAXEND%) As String
Dim endmembertext(0 To MAXEND%, 1 To MAXEND%) As String

Dim ModalGroup As TypeModalGroup

Function ModalFormatEndMember(i As Integer, j As Integer, temp As Single) As String
' Format an end member string

ierror = False
On Error GoTo ModalFormatEndMemberError

Dim astring As String

astring$ = Format$(endmembertext$(i%, j%) & Format$(Format$(temp!, f80$), a40$), a80$)
ModalFormatEndMember$ = astring$

Exit Function

' Errors
ModalFormatEndMemberError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalFormatEndMember"
ierror = True
Exit Function

End Function

Sub ModalGroupPrint()
' Print the modal group

ierror = False
On Error GoTo ModalGroupPrintError

Dim i As Integer, j As Integer, ip As Integer

msg$ = vbCrLf & "Modal Group " & Str$(ModalGroup.GroupNumber%) & " " & ModalGroup.GroupName$
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$

msg$ = "Minimum Total = " & Str$(ModalGroup.MinimumTotal!)
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$

If ModalGroup.NormalizeFlag Then
msg$ = "Normalize Concentrations = Yes"
Else
msg$ = "Normalize Concentrations = No"
End If
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$

If ModalGroup.weightflag Then
msg$ = "Weight Concentrations = Yes"
Else
msg$ = "Weight Concentrations = No"
End If
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$

For i% = 1 To ModalGroup.NumberofPhases%

msg$ = vbCrLf & "Phase " & Str$(i%) & " " & ModalGroup.PhaseNames$(i%)
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$

msg$ = Space$(2) & "Minimum Vector = " & Str$(ModalGroup.MinimumVectors!(i%))
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$

msg$ = Space$(2) & "End Member = " & endmembernames$(ModalGroup.EndMemberNumbers%(i%))
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$

For j% = 1 To ModalGroup.NumberofStandards%(i%)
ip% = StandardGetRow%(ModalGroup.StandardNumbers%(i%, j%))
msg$ = Space$(4) & StandardGetString$(ip%)
Call IOWriteLog(msg$)
Print #ModalOuputDataFileNumber%, msg$
Next j%

Next i%

Exit Sub

' Errors
ModalGroupPrintError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalGroupPrint"
ierror = True
Exit Sub

End Sub

Sub ModalInitGroup(tmodalgroup As TypeModalGroup)
' Initialize the group defaults

ierror = False
On Error GoTo ModalInitGroupError

Dim i As Integer, j As Integer

tmodalgroup.NumberofPhases% = 0
tmodalgroup.MinimumTotal! = defaultminimumtotal!
tmodalgroup.NormalizeFlag% = defaultnormalizeflag%
tmodalgroup.weightflag% = defaultweightflag%
tmodalgroup.DoEndMember = False

For i% = 1 To MAXPHASE%
tmodalgroup.PhaseNames$(i%) = vbNullString
tmodalgroup.NumberofStandards%(i%) = 0
tmodalgroup.MinimumVectors!(i%) = defaultminimumvector!
tmodalgroup.EndMemberNumbers%(i%) = 0

For j% = 1 To MAXSTD%
tmodalgroup.StandardNumbers%(i%, j%) = 0#
Next j%
Next i%

Exit Sub

' Errors
ModalInitGroupError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalInitGroup"
ierror = True
Exit Sub

End Sub

Sub ModalLoadForm()
' Load the modal analysis form

ierror = False
On Error GoTo ModalLoadFormError

' Load default input data file (do not use UserDataDirectory$)
If ModalInputDataFile$ = vbNullString Then ModalInputDataFile$ = CalcZAFDATFileDirectory$ & "\modal.dat"
FormMODAL.TextInputDataFile.Text = ModalInputDataFile$

' Load default output data file
If ModalOutputDataFile$ = vbNullString Then ModalOutputDataFile$ = CalcZAFDATFileDirectory$ & "\modal.out"
FormMODAL.TextOutputDataFile.Text = ModalOutputDataFile$

' Load defaults
If defaultminimumtotal! = 0# Then defaultminimumtotal! = 95#
If defaultminimumvector! = 0# Then defaultminimumvector! = 4#

' Load default flags
defaultnormalizeflag% = True
defaultweightflag% = True

' Load endmember strings
endmembernames$(0) = vbNullString
endmembernames$(1) = "Olivine"
endmembernames$(2) = "Feldspar"
endmembernames$(3) = "Pyroxene"
endmembernames$(4) = "Garnet"

endmembertext$(0, 1) = vbNullString
endmembertext$(1, 1) = "Fo"
endmembertext$(2, 1) = "Ab"
endmembertext$(3, 1) = "Wo"
endmembertext$(4, 1) = "Gr"

endmembertext$(0, 2) = vbNullString
endmembertext$(1, 2) = "Fa"
endmembertext$(2, 2) = "An"
endmembertext$(3, 2) = "En"
endmembertext$(4, 2) = "Py"

endmembertext$(0, 3) = vbNullString
endmembertext$(1, 3) = vbNullString
endmembertext$(2, 3) = "Or"
endmembertext$(3, 3) = "Fs"
endmembertext$(4, 3) = "Alm"

endmembertext$(0, 4) = vbNullString
endmembertext$(1, 4) = vbNullString
endmembertext$(2, 4) = vbNullString
endmembertext$(3, 4) = vbNullString
endmembertext$(4, 4) = "Sp"

' Load first available group
Call ModalUpdateListGroups(Int(0))
If ierror Then Exit Sub

Exit Sub

' Errors
ModalLoadFormError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalLoadForm"
ierror = True
Exit Sub

End Sub

Sub ModalPhaseDelete()
' Delete a phase

ierror = False
On Error GoTo ModalPhaseDeleteError

Dim phanum As Integer, j As Integer
Dim response As Integer

' Determine current phase selection
If FormMODAL.ListPhases.ListIndex > -1 Then
If Not FormMODAL.ListPhases.Selected(FormMODAL.ListPhases.ListIndex) Then GoTo ModalPhaseDeleteNoSelection

' Get number of phase to delete
phanum% = FormMODAL.ListPhases.ItemData(FormMODAL.ListPhases.ListIndex)

' Confirm with user
msg$ = "Are you sure that you want to delete modal phase " & ModalGroup.PhaseNames$(phanum%) & "?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton2, "ModalPhaseDelete")

' User selects "no", just exit
If response% = vbNo Then
Exit Sub
End If

' Delete phase
ModalGroup.PhaseNames$(phanum%) = vbNullString

' Delete standards for that phase too
For j% = 1 To MAXSTD%
ModalGroup.StandardNumbers%(phanum%, j%) = 0
Next j%

' Save and sort
Call ModalSaveGroup
If ierror Then Exit Sub

' Delete group (no confirm)
Call StandardModalDeleteGroup(ModalGroup.GroupNumber%)
If ierror Then Exit Sub

' Add back to database
Call StandardModalSetGroup(ModalGroup)
If ierror Then Exit Sub

' Get new data
Call StandardModalGetGroup(ModalGroup.GroupNumber%, ModalGroup)
If ierror Then Exit Sub
End If

' Update form lists
Call ModalUpdateListPhases(Int(0))
If ierror Then Exit Sub

Exit Sub

' Errors
ModalPhaseDeleteError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalPhaseDelete"
ierror = True
Exit Sub

ModalPhaseDeleteNoSelection:
msg$ = "No modal phase is currently selected"
MsgBox msg$, vbOKOnly + vbExclamation, "ModalPhaseDelete"
ierror = True
Exit Sub

End Sub

Sub ModalPhaseNew()
' Add a new phase

ierror = False
On Error GoTo ModalPhaseNewError

Dim phaname As String
Dim i As Integer

' Ask user for name of new phase
msg$ = "Enter a name for the new phase in modal group (for example: Feldspar or Iron Oxides)" & ModalGroup.GroupName$
phaname$ = InputBox$(msg$, "ModalPhaseNew", vbNullString)
If phaname$ = vbNullString Then Exit Sub

' Check for existing phase with same name
For i% = 1 To ModalGroup.NumberofPhases%
If MiscStringsAreSame(phaname$, ModalGroup.PhaseNames$(i%)) Then GoTo ModalPhaseNewAlreadyExists
Next i%

' Add phase
ModalGroup.NumberofPhases% = ModalGroup.NumberofPhases% + 1
ModalGroup.PhaseNames$(ModalGroup.NumberofPhases%) = phaname$

ModalGroup.MinimumVectors!(ModalGroup.NumberofPhases%) = defaultminimumvector!
ModalGroup.EndMemberNumbers%(ModalGroup.NumberofPhases%) = 0
ModalGroup.NumberofStandards%(ModalGroup.NumberofPhases%) = 0

' Save and sort
Call ModalSaveGroup
If ierror Then Exit Sub

' Delete group (no confirm)
Call StandardModalDeleteGroup(ModalGroup.GroupNumber%)
If ierror Then Exit Sub

' Add back to database
Call StandardModalSetGroup(ModalGroup)
If ierror Then Exit Sub

' Get new data
Call StandardModalGetGroup(ModalGroup.GroupNumber%, ModalGroup)
If ierror Then Exit Sub

' Update form lists
Call ModalUpdateListPhases(ModalGroup.NumberofPhases%)
If ierror Then Exit Sub

Exit Sub

' Errors
ModalPhaseNewError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalPhaseNew"
ierror = True
Exit Sub

ModalPhaseNewAlreadyExists:
msg$ = "Phase " & phaname$ & " already exists. Try again."
MsgBox msg$, vbOKOnly + vbExclamation, "ModalPhaseNew"
ierror = True
Exit Sub

End Sub

Sub ModalSaveForm()
' Save the modal analysis form

ierror = False
On Error GoTo ModalSaveFormError

Dim tfilename As String

' Save input file
tfilename$ = FormMODAL.TextInputDataFile.Text
If Trim$(tfilename$) = vbNullString Then tfilename$ = ApplicationCommonAppData$ & "modal.dat"
If Dir$(tfilename$) = vbNullString Then GoTo ModalSaveFormBadInputFile
ModalInputDataFile$ = tfilename$

' Save output file
tfilename$ = FormMODAL.TextOutputDataFile.Text
If Trim$(tfilename$) = vbNullString Then tfilename$ = ApplicationCommonAppData$ & "modal.out"
ModalOutputDataFile$ = tfilename$

Exit Sub

' Errors
ModalSaveFormError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalSaveForm"
ierror = True
Exit Sub

ModalSaveFormBadInputFile:
msg$ = "Input data file " & tfilename$ & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "ModalSaveForm"
ierror = True
Exit Sub

End Sub

Sub ModalSaveGroup()
' Save (sort) the current modal group (determine new number of phases and standards).

ierror = False
On Error GoTo ModalSaveGroupError

Dim i As Integer, j As Integer

Dim tmodalgroup As TypeModalGroup

' Save group parameters
tmodalgroup.GroupNumber% = ModalGroup.GroupNumber%
tmodalgroup.GroupName$ = ModalGroup.GroupName$
tmodalgroup.MinimumTotal! = ModalGroup.MinimumTotal!
tmodalgroup.DoEndMember% = ModalGroup.DoEndMember%
tmodalgroup.NormalizeFlag% = ModalGroup.NormalizeFlag%
tmodalgroup.weightflag% = ModalGroup.weightflag%

' Save phases parameters
tmodalgroup.NumberofPhases = 0
For i% = 1 To MAXPHASE%
If ModalGroup.PhaseNames$(i%) <> vbNullString Then
tmodalgroup.NumberofPhases = tmodalgroup.NumberofPhases + 1
tmodalgroup.PhaseNames$(tmodalgroup.NumberofPhases%) = ModalGroup.PhaseNames$(i%)
tmodalgroup.MinimumVectors!(tmodalgroup.NumberofPhases%) = ModalGroup.MinimumVectors!(i%)
tmodalgroup.EndMemberNumbers%(tmodalgroup.NumberofPhases%) = ModalGroup.EndMemberNumbers%(i%)
    
' Save standard parameters
tmodalgroup.NumberofStandards%(tmodalgroup.NumberofPhases%) = 0
For j% = 1 To MAXSTD%
If ModalGroup.StandardNumbers%(i%, j%) > 0 Then
tmodalgroup.NumberofStandards%(tmodalgroup.NumberofPhases%) = tmodalgroup.NumberofStandards%(tmodalgroup.NumberofPhases%) + 1
tmodalgroup.StandardNumbers%(tmodalgroup.NumberofPhases%, tmodalgroup.NumberofStandards%(tmodalgroup.NumberofPhases%)) = ModalGroup.StandardNumbers%(i%, j%)
End If
Next j%

End If
Next i%

' Return in module level UDT
ModalGroup = tmodalgroup

Exit Sub

' Errors
ModalSaveGroupError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalSaveGroup"
ierror = True
Exit Sub

End Sub

Sub ModalSaveOptionsGroup()
' Save the group options

ierror = False
On Error GoTo ModalSaveOptionsGroupError

' Update the group options
ModalGroup.MinimumTotal! = Val(FormMODAL.TextMinimumTotal.Text)
defaultminimumtotal! = ModalGroup.MinimumTotal!

If FormMODAL.CheckDoEndMembers.Value = vbChecked Then
ModalGroup.DoEndMember = True
Else
ModalGroup.DoEndMember = False
End If

If FormMODAL.CheckNormalize = vbChecked Then
ModalGroup.NormalizeFlag% = True
Else
ModalGroup.NormalizeFlag% = False
End If
defaultnormalizeflag% = ModalGroup.NormalizeFlag%

If FormMODAL.CheckWeight = vbChecked Then
ModalGroup.weightflag% = True
Else
ModalGroup.weightflag% = False
End If
defaultweightflag% = ModalGroup.weightflag%

' Reorder and sort
Call ModalSaveGroup
If ierror Then Exit Sub

' Delete group (no confirm)
Call StandardModalDeleteGroup(ModalGroup.GroupNumber%)
If ierror Then Exit Sub

' Add back to database
Call StandardModalSetGroup(ModalGroup)
If ierror Then Exit Sub

Exit Sub

' Errors
ModalSaveOptionsGroupError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalSaveOptionsGroup"
ierror = True
Exit Sub

End Sub

Sub ModalSaveOptionsPhase()
' Save the phase options

ierror = False
On Error GoTo ModalSaveOptionsPhaseError

Dim phanum As Integer, i As Integer

' Determine current phase
If FormMODAL.ListPhases.ListIndex > -1 Then
phanum% = FormMODAL.ListPhases.ItemData(FormMODAL.ListPhases.ListIndex)

' Save the phase options
ModalGroup.MinimumVectors!(phanum%) = Val(FormMODAL.TextMinimumVector.Text)
defaultminimumvector! = ModalGroup.MinimumVectors!(phanum%)

For i% = 0 To 4
If FormMODAL.OptionEndMember(i%).Value Then ModalGroup.EndMemberNumbers%(phanum%) = i%
Next i%
End If

' Reorder and sort
Call ModalSaveGroup
If ierror Then Exit Sub

' Delete group (no confirm)
Call StandardModalDeleteGroup(ModalGroup.GroupNumber%)
If ierror Then Exit Sub

' Add back to database
Call StandardModalSetGroup(ModalGroup)
If ierror Then Exit Sub

Exit Sub

' Errors
ModalSaveOptionsPhaseError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalSaveOptionsPhase"
ierror = True
Exit Sub

End Sub

Sub ModalStandardAdd()
' Add a standard to the modal group

ierror = False
On Error GoTo ModalStandardAddError

Dim phanum As Integer, j As Integer

' Determine current phase
If FormMODAL.ListPhases.ListIndex > -1 Then
If Not FormMODAL.ListPhases.Selected(FormMODAL.ListPhases.ListIndex) Then GoTo ModalStandardAddNoSelection

' Get number of phase to add standards to
phanum% = FormMODAL.ListPhases.ItemData(FormMODAL.ListPhases.ListIndex)

Call InitStandard
If ierror Then Exit Sub

' Load standards for this phase
NumberofStandards% = ModalGroup.NumberofStandards%(phanum%)
For j% = 1 To NumberofStandards%
StandardNumbers%(j%) = ModalGroup.StandardNumbers%(phanum%, j%)
Next j%

' Load form
Call AddStdLoad
If ierror Then Exit Sub

FormADDSTD.Show vbModal
If ierror Then Exit Sub

' Load standards to phase
For j% = 1 To NumberofStandards%
ModalGroup.StandardNumbers%(phanum%, j%) = StandardNumbers%(j%)
Next j%
End If

' Reorder and sort
Call ModalSaveGroup
If ierror Then Exit Sub

' Delete group (no confirm)
Call StandardModalDeleteGroup(ModalGroup.GroupNumber%)
If ierror Then Exit Sub

' Add back to database
Call StandardModalSetGroup(ModalGroup)
If ierror Then Exit Sub

' Get new data
Call StandardModalGetGroup(ModalGroup.GroupNumber%, ModalGroup)
If ierror Then Exit Sub

' Update form lists
Call ModalUpdateListPhases(phanum%)
If ierror Then Exit Sub

Exit Sub

' Errors
ModalStandardAddError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalStandardAdd"
ierror = True
Exit Sub

ModalStandardAddNoSelection:
msg$ = "No modal phase is currently selected"
MsgBox msg$, vbOKOnly + vbExclamation, "ModalStandardAdd"
ierror = True
Exit Sub

End Sub

Sub ModalStandardRemove()
' Remove a standard from the modal group

ierror = False
On Error GoTo ModalStandardRemoveError

Dim phanum As Integer, stdnum As Integer

' Determine current phase
If FormMODAL.ListPhases.ListIndex > -1 Then
If Not FormMODAL.ListPhases.Selected(FormMODAL.ListPhases.ListIndex) Then GoTo ModalStandardRemoveNoPhaseSelection

' Get number of phase to add standards to
phanum% = FormMODAL.ListPhases.ItemData(FormMODAL.ListPhases.ListIndex)

' Determine current standard
If FormMODAL.ListStandards.ListIndex > -1 Then
If Not FormMODAL.ListStandards.Selected(FormMODAL.ListStandards.ListIndex) Then GoTo ModalStandardRemoveNoStandardSelection

' Get number of standard to delete
stdnum% = FormMODAL.ListStandards.ItemData(FormMODAL.ListStandards.ListIndex)
 
' Delete standard
ModalGroup.StandardNumbers%(phanum%, stdnum%) = 0
End If
End If

' Reorder and sort
Call ModalSaveGroup
If ierror Then Exit Sub

' Delete group (no confirm)
Call StandardModalDeleteGroup(ModalGroup.GroupNumber%)
If ierror Then Exit Sub

' Add back to database
Call StandardModalSetGroup(ModalGroup)
If ierror Then Exit Sub

' Get new data
Call StandardModalGetGroup(ModalGroup.GroupNumber%, ModalGroup)
If ierror Then Exit Sub

' Update form lists
Call ModalUpdateListPhases(phanum%)
If ierror Then Exit Sub

Exit Sub

' Errors
ModalStandardRemoveError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalStandardRemove"
ierror = True
Exit Sub

ModalStandardRemoveNoPhaseSelection:
msg$ = "No modal phase is currently selected"
MsgBox msg$, vbOKOnly + vbExclamation, "ModalStandardRemove"
ierror = True
Exit Sub

ModalStandardRemoveNoStandardSelection:
msg$ = "No modal standard is currently selected"
MsgBox msg$, vbOKOnly + vbExclamation, "ModalStandardRemove"
ierror = True
Exit Sub

End Sub

Sub ModalStartModal()
' Start the modal analysis

ierror = False
On Error GoTo ModalStartModalError

' Open input file of unknown weight percents
Open ModalInputDataFile$ For Input As #ModalInputDataFileNumber%

' Open output file
Open ModalOutputDataFile$ For Output As #ModalOuputDataFileNumber%

' Run the modal analysis
Call ModalRunModal(ModalGroup)
Close #ModalInputDataFileNumber%
Close #ModalOuputDataFileNumber%

If ierror Then Exit Sub

' Confirm completion and output data file to user
msg$ = "Output data saved to " & ModalOutputDataFile$
MsgBox msg$, vbOKOnly + vbInformation, "ModalStartModal"

Exit Sub

' Errors
ModalStartModalError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalStartModal"
ierror = True
Exit Sub

End Sub

Sub ModalUpdateListGroups(grpnum As Integer)
' Update the lists for the above grpnum
' grpnum = 0 select first available
' grpnum > 0 select grpnum

ierror = False
On Error GoTo ModalUpdateListGroupsError

Dim i As Integer

' Clear all lists
FormMODAL.ListGroups.Clear
FormMODAL.ListPhases.Clear
FormMODAL.ListStandards.Clear

' Update group list
Call StandardModalUpdateGroupList(FormMODAL.ListGroups)
If ierror Then Exit Sub

' Get selected modal group from database
Call StandardModalGetGroup(grpnum%, ModalGroup)
If ierror Then Exit Sub

' Select the first available or specified group
If FormMODAL.ListGroups.ListCount > 0 Then
If grpnum% > 0 Then
For i% = 0 To FormMODAL.ListGroups.ListCount - 1
If FormMODAL.ListGroups.ItemData(i%) = grpnum% Then
FormMODAL.ListGroups.Selected(i%) = True
End If
Next i%

Else
FormMODAL.ListGroups.Selected(0) = True
End If
End If

Exit Sub

' Errors
ModalUpdateListGroupsError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalUpdateListGroups"
ierror = True
Exit Sub

End Sub

Sub ModalUpdateListPhases(phanum As Integer)
' Update the FormMODAL phase group
' phanum = 0 select first available
' phanum > 0 select specified phanum

ierror = False
On Error GoTo ModalUpdateListPhasesError

Dim i As Integer

' Clear lists
FormMODAL.ListPhases.Clear
FormMODAL.ListStandards.Clear

' Update phase list
FormMODAL.ListPhases.Clear
For i% = 1 To ModalGroup.NumberofPhases%
FormMODAL.ListPhases.AddItem ModalGroup.PhaseNames$(i%)
FormMODAL.ListPhases.ItemData(FormMODAL.ListPhases.NewIndex) = i%
Next i%

If FormMODAL.ListPhases.ListCount > 0 Then
If phanum% > 0 Then
For i% = 0 To FormMODAL.ListPhases.ListCount - 1
If FormMODAL.ListPhases.ItemData(i%) = phanum% Then
FormMODAL.ListPhases.Selected(i%) = True
End If
Next i%

Else
FormMODAL.ListPhases.Selected(phanum%) = True
End If
End If

Exit Sub

' Errors
ModalUpdateListPhasesError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalUpdateListPhases"
ierror = True
Exit Sub

End Sub

Sub ModalUpdateListStandards(phanum As Integer)
' Update the standard list for FormMODAL

ierror = False
On Error GoTo ModalUpdateListStandardsError

Dim j As Integer, ip As Integer

' Update standard list
FormMODAL.ListStandards.Clear
If phanum% > 0 Then
For j% = 1 To ModalGroup.NumberofStandards%(phanum%)

' Determine standard
ip% = StandardGetRow%(ModalGroup.StandardNumbers%(phanum%, j%))

' Check that standard has not been deleted
If ip% = 0 Then
msg$ = "Standard number " & Str$(ModalGroup.StandardNumbers%(phanum%, j%)) & " was not found in the database. Please delete it from the standard list"
MsgBox msg$, vbOKOnly + vbExclamation, "ModalUpdateListStandards"
msg$ = Format$(ModalGroup.StandardNumbers%(phanum%, j%), a40$)
Else
msg$ = StandardGetString$(ip%)
End If

FormMODAL.ListStandards.AddItem msg$
FormMODAL.ListStandards.ItemData(FormMODAL.ListStandards.NewIndex) = j%
Next j%
End If

Exit Sub

' Errors
ModalUpdateListStandardsError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalUpdateListStandards"
ierror = True
Exit Sub

End Sub

Sub ModalUpdatePhases()
' Update the phases list based on click in ListGroups

ierror = False
On Error GoTo ModalUpdatePhasesError

Dim grpnum As Integer

' Update list for selected group
If FormMODAL.ListGroups.ListIndex > -1 Then
grpnum% = FormMODAL.ListGroups.ItemData(FormMODAL.ListGroups.ListIndex)

' Get new data
Call StandardModalGetGroup(grpnum%, ModalGroup)
If ierror Then Exit Sub

' Update the phase list
Call ModalUpdateListPhases(Int(0))
If ierror Then Exit Sub

' Update the group options
FormMODAL.TextMinimumTotal.Text = Str$(ModalGroup.MinimumTotal!)

If ModalGroup.DoEndMember Then
FormMODAL.CheckDoEndMembers.Value = vbChecked
Else
FormMODAL.CheckDoEndMembers.Value = vbUnchecked
End If

If ModalGroup.NormalizeFlag% Then
FormMODAL.CheckNormalize.Value = vbChecked
Else
FormMODAL.CheckNormalize.Value = vbUnchecked
End If

If ModalGroup.weightflag% Then
FormMODAL.CheckWeight.Value = vbChecked
Else
FormMODAL.CheckWeight.Value = vbUnchecked
End If

End If

Exit Sub

' Errors
ModalUpdatePhasesError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalUpdatePhases"
ierror = True
Exit Sub

End Sub

Sub ModalUpdateStandards()
' Update the standards list based on click in ListPhases

ierror = False
On Error GoTo ModalUpdateStandardsError

Dim phanum As Integer

' Update list for selected phase
If FormMODAL.ListPhases.ListIndex > -1 Then
phanum% = FormMODAL.ListPhases.ItemData(FormMODAL.ListPhases.ListIndex)

' Update the standards list for this phase
Call ModalUpdateListStandards(phanum%)
If ierror Then Exit Sub
End If

' Update the phase options
FormMODAL.TextMinimumVector.Text = Str$(ModalGroup.MinimumVectors!(phanum%))
FormMODAL.OptionEndMember(ModalGroup.EndMemberNumbers%(phanum%)).Value = True

Exit Sub

' Errors
ModalUpdateStandardsError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalUpdateStandards"
ierror = True
Exit Sub

End Sub

Sub ModalGetInputDataFile(tForm As Form)
' Get the input data file (.DAT)

ierror = False
On Error GoTo ModalGetInputDataFileError

Dim tfilename As String

' Get new filename
tfilename$ = ModalInputDataFile$
Call IOGetFileName(Int(2), "DAT", tfilename$, tForm)
If ierror Then Exit Sub

ModalInputDataFile$ = tfilename$
FormMODAL.TextInputDataFile.Text = ModalInputDataFile$

Exit Sub

' Errors
ModalGetInputDataFileError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalGetInputDataFile"
ierror = True
Exit Sub

End Sub

Sub ModalGetOutputDataFile(tForm As Form)
' Get output data file (.OUT)

ierror = False
On Error GoTo ModalGetOutputDataFileError

Dim tfilename As String

' Get new filename
tfilename$ = ModalOutputDataFile$
Call IOGetFileName(Int(1), "OUT", tfilename$, tForm)
If ierror Then Exit Sub

ModalOutputDataFile$ = tfilename$
FormMODAL.TextOutputDataFile.Text = ModalOutputDataFile$

Exit Sub

' Errors
ModalGetOutputDataFileError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalGetOutputDataFile"
ierror = True
Exit Sub

End Sub

Sub ModalGroupDelete()
' Delete a modal group set of phases

ierror = False
On Error GoTo ModalGroupDeleteError

Dim grpnum As Integer, response As Integer

If FormMODAL.ListGroups.ListCount < 1 Then Exit Sub

' Determine current group selection
If FormMODAL.ListGroups.ListIndex > -1 Then
If Not FormMODAL.ListGroups.Selected(FormMODAL.ListGroups.ListIndex) Then GoTo ModalGroupDeleteNoSelection

' Get number of group to delete
grpnum% = FormMODAL.ListGroups.ItemData(FormMODAL.ListGroups.ListIndex)

' Confirm with user
msg$ = "Are you sure that you want to delete modal group " & ModalGroup.GroupName$ & "?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton2, "ModalGroupDelete")

' User selects "no", just exit
If response% = vbNo Then
Exit Sub
End If

' Delete group
Call StandardModalDeleteGroup(grpnum%)
If ierror Then Exit Sub
End If

' Load first available group
Call ModalUpdateListGroups(Int(0))
If ierror Then Exit Sub

Exit Sub

' Errors
ModalGroupDeleteError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalGroupDelete"
ierror = True
Exit Sub

ModalGroupDeleteNoSelection:
msg$ = "No modal group is currently selected"
MsgBox msg$, vbOKOnly + vbExclamation, "ModalGroupDelete"
ierror = True
Exit Sub

End Sub

Sub ModalGroupNew()
' Add a new modal group set of phases

ierror = False
On Error GoTo ModalGroupNewError

Dim grpname As String
Dim grpnum As Integer

' Ask user for name of new group
msg$ = "Enter a name for the new modal group (for example: Al-Cu Eutectic, Stony Chondrite or Eclogite Metamorphic Suite)"
grpname$ = InputBox$(msg$, "ModalGroupNew", vbNullString)
If grpname$ = vbNullString Then Exit Sub

' Check for existing group with same name
Call StandardModalCheckGroupName(grpname$)
If ierror Then Exit Sub

' Determine next free group number
grpnum% = StandardModalGetNextGroupNumber()
If ierror Then Exit Sub

' Initialize group
Call ModalInitGroup(ModalGroup)
If ierror Then Exit Sub

' Add defaults
ModalGroup.GroupNumber% = grpnum%
ModalGroup.GroupName$ = grpname$

' Save and sort
Call ModalSaveGroup
If ierror Then Exit Sub

' Add to database
Call StandardModalSetGroup(ModalGroup)
If ierror Then Exit Sub

' Update form lists
Call ModalUpdateListGroups(grpnum%)
If ierror Then Exit Sub

Exit Sub

' Errors
ModalGroupNewError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalGroupNew"
ierror = True
Exit Sub

End Sub

