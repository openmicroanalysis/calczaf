Attribute VB_Name = "CodeSTANDARD3"
' (c) Copyright 1995-2022 by John J. Donovan
Option Explicit

Sub StandardModalCheckGroupName(grpname As String)
' Determine if the group name is already used

ierror = False
On Error GoTo StandardModalCheckGroupNameError

Dim tstring As String

Dim StDb As Database
Dim StDt As Recordset

' Open the database
Screen.MousePointer = vbHourglass
If StandardDataFile$ = vbNullString Then StandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseNonExclusiveAccess%, dbReadOnly)

' Loop through Group table and check names
Set StDt = StDb.OpenRecordset("Groups", dbOpenSnapshot)

Do Until StDt.EOF
tstring$ = StDt("GroupNames")
If MiscStringsAreSame(tstring$, grpname$) Then GoTo StandardModalCheckGroupName
StDt.MoveNext
Loop

StDt.Close
StDb.Close

Screen.MousePointer = vbDefault
Exit Sub

' Errors
StandardModalCheckGroupNameError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardModalCheckGroupName"
ierror = True
Exit Sub

StandardModalCheckGroupName:
Screen.MousePointer = vbDefault
msg$ = "Group name already exists. Try again with another name."
MsgBox msg$, vbOKOnly + vbExclamation, "StandardModalCheckGroupName"
ierror = True
Exit Sub

End Sub

Sub StandardModalDeleteGroup(grpnum As Integer)
' Delete the specified modal group

ierror = False
On Error GoTo StandardModalDeleteGroupError

Dim SQLQ As String
Dim StDb As Database
Dim StDt As Recordset

' Open the database
If StandardDataFile$ = vbNullString Then StandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseExclusiveAccess%, False)

Set StDt = StDb.OpenRecordset("Groups", dbOpenTable)
StDt.Index = "Group Numbers"
StDt.Seek "=", grpnum%
If StDt.NoMatch Then GoTo StandardModalDeleteGroupNotFound

Call TransactionBegin("StandardModalDeleteGroup", StandardDataFile$)
If ierror Then Exit Sub

' Delete all tables for this modal group
StDt.Delete
StDt.MoveFirst
StDt.Close

' Delete phase table based on "grpnum"
SQLQ$ = "DELETE from Phase WHERE Phase.PhaseToRow = " & Str$(grpnum%)
StDb.Execute SQLQ$

' Delete ModalStd table based on "grpnum"
SQLQ$ = "DELETE from ModalStd WHERE ModalStd.StdToRow = " & Str$(grpnum%)
StDb.Execute SQLQ$

Call TransactionCommit("StandardModalDeleteGroup", StandardDataFile$)
If ierror Then Exit Sub

StDb.Close
Exit Sub

' Errors
StandardModalDeleteGroupError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardModalDeleteGroup"
Call TransactionRollback("StandardModalDeleteGroup", StandardDataFile$)
ierror = True
Exit Sub

StandardModalDeleteGroupNotFound:
msg$ = "Group number " & Str$(grpnum%) & " was not found in " & StandardDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardModalDeleteGroup"
ierror = True
Exit Sub

End Sub

Sub StandardModalGetGroup(grpnum As Integer, ModalGroup As TypeModalGroup)
' Get a saved modal group of phases from the standard database
' grpnum = group number to get (if zero then first available group)

ierror = False
On Error GoTo StandardGetModalGroupError

Dim phaserow As Integer, stdrow As Integer
Dim StDb As Database
Dim StDt As Recordset
Dim stds As Recordset
Dim SQLQ As String

' Initialize the group arrays
Call ModalInitGroup(ModalGroup)
If ierror Then Exit Sub

' Open the database
Screen.MousePointer = vbHourglass
If StandardDataFile$ = vbNullString Then StandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseNonExclusiveAccess%, dbReadOnly)

Set StDt = StDb.OpenRecordset("Groups", dbOpenTable)      ' use dbOpenTable to support .Seek method below
StDt.Index = "Group Numbers"

' Seek specified group
If grpnum% > 0 Then
StDt.Seek "=", grpnum%
If StDt.NoMatch Then GoTo StandardGetModalGroupNotFound

' Seek last available group
Else
If Not StDt.EOF Then
StDt.MoveLast
grpnum% = StDt("GroupNumbers")
Else
StDt.Close
StDb.Close
Screen.MousePointer = vbDefault
Exit Sub
End If

End If

' Load values from "group" table
ModalGroup.GroupNumber% = StDt("GroupNumbers")

' Check for Null values from database
ModalGroup.GroupName$ = Trim$(vbNullString & StDt("GroupNames"))
ModalGroup.MinimumTotal! = StDt("MinimumTotals")
ModalGroup.DoEndMember% = StDt("DoEndMembers")
ModalGroup.NormalizeFlag% = StDt("NormalizeFlags")
ModalGroup.weightflag% = StDt("WeightFlags")
ModalGroup.NumberofPhases% = StDt("NumberofPhases")
If ModalGroup.NumberofPhases% < 0 Then GoTo StandardModalGetGroupBadNumberofPhases
If ModalGroup.NumberofPhases% > MAXPHASE% Then GoTo StandardModalGetGroupBadNumberofPhases
StDt.Close

' Get phase data for specified group from standard database
SQLQ$ = "SELECT Phase.* FROM Phase WHERE Phase.PhaseToRow = " & Str$(grpnum%)
Set stds = StDb.OpenRecordset(SQLQ$, dbOpenSnapshot, dbReadOnly)

' Load all phases from "Phase" table that matched the group number
Do Until stds.EOF
phaserow% = stds("PhaseOrder")
If phaserow% < 1 Or phaserow% > ModalGroup.NumberofPhases% Then GoTo StandardModalGetGroupBadPhaseRow
ModalGroup.PhaseNames$(phaserow%) = Trim$(vbNullString & stds("PhaseNames"))
ModalGroup.MinimumVectors!(phaserow%) = stds("MinimumVectors")
ModalGroup.EndMemberNumbers%(phaserow%) = stds("EndMemberNumbers")
ModalGroup.NumberofStandards%(phaserow%) = stds("NumberofStandards")
If ModalGroup.NumberofStandards%(phaserow%) < 0 Then GoTo StandardModalGetGroupBadNumberofStds
If ModalGroup.NumberofStandards%(phaserow%) > MAXSTD% Then GoTo StandardModalGetGroupBadNumberofStds
stds.MoveNext
Loop
stds.Close

' Get ModalStd data for group number from standard database
SQLQ$ = "SELECT ModalStd.* FROM ModalStd WHERE ModalStd.StdToRow = " & Str$(grpnum%)
Set stds = StDb.OpenRecordset(SQLQ$, dbOpenSnapshot, dbReadOnly)

' Use "PhaseOrder" and "StdOrder" fields to load phases and standards in order
Do Until stds.EOF
phaserow% = stds("PhaseOrder")
stdrow% = stds("StdOrder")
If phaserow% < 1 Or phaserow% > ModalGroup.NumberofPhases% Then GoTo StandardModalGetGroupBadPhaseRow
If stdrow% < 1 Or stdrow% > ModalGroup.NumberofStandards%(phaserow%) Then GoTo StandardModalGetGroupBadStdRow
ModalGroup.StandardNumbers%(phaserow%, stdrow%) = stds("Numbers")
stds.MoveNext
Loop
stds.Close

' Close the standard database
StDb.Close

Screen.MousePointer = vbDefault
Exit Sub

' Errors
StandardGetModalGroupError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardGetModalGroup"
ierror = True
Exit Sub

StandardGetModalGroupNotFound:
Screen.MousePointer = vbDefault
msg$ = "Modal group number " & Str$(grpnum%) & " was not found in the standard database " & StandardDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardGetModalGroup"
ierror = True
Exit Sub

StandardModalGetGroupBadNumberofPhases:
Screen.MousePointer = vbDefault
msg$ = "Bad number of phases for modal group number " & Str$(grpnum%)
MsgBox msg$, vbOKOnly + vbExclamation, "StandardGetModalGroup"
ierror = True
Exit Sub

StandardModalGetGroupBadPhaseRow:
Screen.MousePointer = vbDefault
msg$ = "Bad phase row for modal group number " & Str$(grpnum%)
MsgBox msg$, vbOKOnly + vbExclamation, "StandardGetModalGroup"
ierror = True
Exit Sub

StandardModalGetGroupBadNumberofStds:
Screen.MousePointer = vbDefault
msg$ = "Bad number of standards for modal group number " & Str$(grpnum%) & " in phase number " & Str$(phaserow%)
MsgBox msg$, vbOKOnly + vbExclamation, "StandardGetModalGroup"
ierror = True
Exit Sub

StandardModalGetGroupBadStdRow:
Screen.MousePointer = vbDefault
msg$ = "Bad standard row for modal group number " & Str$(grpnum%)
MsgBox msg$, vbOKOnly + vbExclamation, "StandardGetModalGroup"
ierror = True
Exit Sub

End Sub

Function StandardModalGetNextGroupNumber() As Integer
' Determine the next free modal group number

ierror = False
On Error GoTo StandardModalGetNextGroupNumberError

Dim grpnum As Integer

Dim SQLQ As String
Dim StDb As Database
Dim stds As Recordset

' Open the database
Screen.MousePointer = vbHourglass
If StandardDataFile$ = vbNullString Then StandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseNonExclusiveAccess%, dbReadOnly)

SQLQ$ = "SELECT Groups.GroupNumbers from Groups WHERE GroupNumbers <> 0 "
SQLQ$ = SQLQ$ & "ORDER BY GroupNumbers DESC"
Set stds = StDb.OpenRecordset(SQLQ$, dbOpenDynaset)

' Create unique group number
If stds.BOF And stds.EOF Then
grpnum% = 1
Else
grpnum% = stds("GroupNumbers") + 1
End If

stds.Close
StDb.Close

' Return next free number
StandardModalGetNextGroupNumber% = grpnum%

Screen.MousePointer = vbDefault
Exit Function

' Errors
StandardModalGetNextGroupNumberError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardModalGetNextGroupNumber"
ierror = True
Exit Function

End Function

Sub StandardModalSetGroup(ModalGroup As TypeModalGroup)
' Set (save) a new or modified modal group of phases to the standard database

ierror = False
On Error GoTo StandardModalSetGroupError

Dim phaserow As Integer, stdrow As Integer
Dim grpnum As Integer

Dim StDb As Database
Dim StDt As Recordset
Dim SQLQ As String

' Check group number
grpnum% = ModalGroup.GroupNumber%
If grpnum% < 1 Then GoTo StandardModalSetGroupBadGroupNumber
If grpnum% > MAXGROUP% Then GoTo StandardModalSetGroupBadGroupNumber

If ModalGroup.NumberofPhases% < 0 Then GoTo StandardModalSetGroupBadNumberofPhases
If ModalGroup.NumberofPhases% > MAXPHASE% Then GoTo StandardModalSetGroupBadNumberofPhases

' Check for valid standard numbers
For phaserow% = 1 To ModalGroup.NumberofPhases%
If ModalGroup.NumberofStandards%(phaserow%) < 0 Then GoTo StandardModalSetGroupBadNumberofStds
If ModalGroup.NumberofStandards%(phaserow%) > MAXSTD% Then GoTo StandardModalSetGroupBadNumberofStds
Next phaserow%

' Open the database
Screen.MousePointer = vbHourglass
If StandardDataFile$ = vbNullString Then StandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseExclusiveAccess%, False)

Set StDt = StDb.OpenRecordset("Groups", dbOpenTable)
StDt.Index = "Group Numbers"
StDt.Seek "=", grpnum%
If Not StDt.NoMatch Then GoTo StandardModalSetGroupFound

Call TransactionBegin("StandardModalSetGroup", StandardDataFile$)
If ierror Then Exit Sub

' Save values to "group" table
StDt.AddNew
StDt("GroupNumbers") = grpnum%
StDt("GroupNames") = Left$(ModalGroup.GroupName$, 64)
StDt("MinimumTotals") = ModalGroup.MinimumTotal!
StDt("DoEndMembers") = ModalGroup.DoEndMember%
StDt("NormalizeFlags") = ModalGroup.NormalizeFlag%
StDt("WeightFlags") = ModalGroup.weightflag%
StDt("NumberofPhases") = ModalGroup.NumberofPhases%
StDt.Update

StDt.Close

' Delete phase table based on "grpnum"
SQLQ$ = "DELETE from Phase WHERE Phase.PhaseToRow = " & Str$(grpnum%)
StDb.Execute SQLQ$

' Save phases in order
Set StDt = StDb.OpenRecordset("Phase", dbOpenTable)
For phaserow% = 1 To ModalGroup.NumberofPhases%

StDt.AddNew
StDt("PhaseToRow") = grpnum%
StDt("PhaseOrder") = phaserow%
StDt("PhaseNames") = Left$(ModalGroup.PhaseNames$(phaserow%), 64)
StDt("MinimumVectors") = ModalGroup.MinimumVectors!(phaserow%)
StDt("EndMemberNumbers") = ModalGroup.EndMemberNumbers%(phaserow%)
StDt("NumberofStandards") = ModalGroup.NumberofStandards%(phaserow%)
StDt.Update
Next phaserow%

StDt.Close

' Delete ModalStd table based on "grpnum"
SQLQ$ = "DELETE from ModalStd WHERE ModalStd.StdToRow = " & Str$(grpnum%)
StDb.Execute SQLQ$

' Save standards in order
Set StDt = StDb.OpenRecordset("ModalStd", dbOpenTable)
For phaserow% = 1 To ModalGroup.NumberofPhases%
For stdrow% = 1 To ModalGroup.NumberofStandards%(phaserow%)
StDt.AddNew
StDt("StdToRow") = grpnum%
StDt("PhaseOrder") = phaserow%
StDt("StdOrder") = stdrow%
StDt("Numbers") = ModalGroup.StandardNumbers%(phaserow%, stdrow%)
StDt.Update
Next stdrow%
Next phaserow%
StDt.Close

Call TransactionCommit("StandardModalSetGroup", StandardDataFile$)
If ierror Then Exit Sub

' Close the standard database
StDb.Close
Screen.MousePointer = vbDefault
Exit Sub

' Errors
StandardModalSetGroupError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardModalSetGroup"
Call TransactionRollback("StandardModalSetGroup", StandardDataFile$)
ierror = True
Exit Sub

StandardModalSetGroupBadGroupNumber:
Screen.MousePointer = vbDefault
msg$ = "Bad modal group number"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardModalSetGroup"
ierror = True
Exit Sub

StandardModalSetGroupFound:
Screen.MousePointer = vbDefault
msg$ = "Modal group number " & Str$(grpnum%) & " already exists in the standard database " & StandardDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardModalSetGroup"
ierror = True
Exit Sub

StandardModalSetGroupBadNumberofPhases:
Screen.MousePointer = vbDefault
msg$ = "Bad number of phases for modal group number " & Str$(grpnum%)
MsgBox msg$, vbOKOnly + vbExclamation, "StandardModalSetGroup"
ierror = True
Exit Sub

StandardModalSetGroupBadNumberofStds:
Screen.MousePointer = vbDefault
msg$ = "Bad number of standards for modal group number " & Str$(grpnum%) & " in phase number " & Str$(phaserow%)
MsgBox msg$, vbOKOnly + vbExclamation, "StandardModalSetGroup"
ierror = True
Exit Sub

End Sub

Sub StandardModalUpdateGroupList(tList As ListBox)
' Update the group name list in FormMODAL

ierror = False
On Error GoTo StandardModalUpdateGroupListError

Dim StDb As Database
Dim StDt As Recordset

' Open the database
Screen.MousePointer = vbHourglass
If StandardDataFile$ = vbNullString Then StandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseNonExclusiveAccess%, dbReadOnly)
Set StDt = StDb.OpenRecordset("Groups", dbOpenSnapshot)

' Load values from "group" table
tList.Clear
Do Until StDt.EOF
tList.AddItem Trim$(vbNullString & StDt("GroupNames"))
tList.ItemData(tList.NewIndex) = StDt("GroupNumbers")
StDt.MoveNext
Loop
StDt.Close
StDb.Close

Screen.MousePointer = vbDefault
Exit Sub

' Errors
StandardModalUpdateGroupListError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardModalUpdateGroupList"
ierror = True
Exit Sub

End Sub


