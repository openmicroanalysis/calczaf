Attribute VB_Name = "CodeMATCH"
' (c) Copyright 1995-2024 by John J. Donovan
Option Explicit

Dim OriginalStandardDataFile As String

Dim MatchOldSample(1 To 1) As TypeSample
Dim MatchTmpSample(1 To 1) As TypeSample

Sub MatchLoad(sample() As TypeSample)
' Load the FormMATCH form to match the passed unknown composition

ierror = False
On Error GoTo MatchLoadError

Dim tmsg As String

' Load filename
FormMATCH.Caption = "Match Unknown To Match Database [" & DefaultMatchStandardDatabase$ & "]"

' Load default minimum vector
If Val(FormMATCH.TextMinimumVector.Text) <= 0# Then FormMATCH.TextMinimumVector.Text = Str$(40#)

' Load sample
MatchOldSample(1) = sample(1)

tmsg$ = SampleGetString2(MatchOldSample())
FormMATCH.LabelUnknown.Caption = tmsg$

' Convert standard to string
tmsg$ = TypeWeight(Int(0), MatchOldSample())
FormMATCH.TextComposition.Text = tmsg$

Exit Sub

' Errors
MatchLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "MatchLoad"
ierror = True
Exit Sub

End Sub

Sub MatchStandards(tStandardDataFile As String, minimumvector As Single, tList As ListBox)
' Routine to find all standards within a range and select them (always call from MatchSample)
' in the passed list box and pass vector

ierror = False
On Error GoTo MatchStandardsError

Dim i As Integer, n As Integer, chan As Integer
Dim ip As Integer, itemp As Integer, stdnum As Integer
Dim istd As Integer, istds As Integer
Dim temp As Single, vector As Single
Dim tmsg As String

Dim SQLQ As String
Dim StDb As Database
Dim stds As Recordset

ReDim numbers(1 To MAXINDEX%) As Integer
ReDim vectors(1 To MAXINDEX%) As Single

icancelauto = False

tmsg$ = vbCrLf & "Searching standard database for given composition..."
Call IOWriteLog(tmsg$)

' Store given composition
tmsg$ = TypeWeight(Int(1), MatchOldSample())
Call IOWriteLog(tmsg$)
If MatchOldSample(1).LastChan% = 0 Then GoTo MatchStandardsNoElements

' Open the standard database
If tStandardDataFile$ = vbNullString Then tStandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
Set StDb = OpenDatabase(tStandardDataFile$, StandardDatabaseNonExclusiveAccess%, dbReadOnly)

' Get all standards that contain the elements
SQLQ$ = "SELECT DISTINCT Element.Number FROM Element WHERE "
For i% = 1 To MatchOldSample(1).LastChan%
SQLQ$ = SQLQ$ & "Element.Symbol = " & "'" & (MatchOldSample(1).Elsyms$(i%)) & "'"
If i% <> MatchOldSample(1).LastChan% Then SQLQ$ = SQLQ$ & " or "
Next i%

' Get standard numbers
Set stds = StDb.OpenRecordset(SQLQ$, dbOpenSnapshot, dbReadOnly)

' Loop on all standards
istd% = 0
istds% = stds.RecordCount
n% = 0
Do Until stds.EOF
stdnum% = stds("Number")
istd% = istd% + 1

' Get standard from database
Call StandardGetMDBStandard(stdnum%, MatchTmpSample())
If ierror Then Exit Sub

Call IOStatusAuto("Checking standard " & Str$(stdnum%) & " (" & Str$(istd%) & " of " & Str$(istds%) & ")")
DoEvents
If icancelauto Then
ierror = True
Exit Sub
End If

' Add elements in standard but not in unknown
itemp% = MatchOldSample(1).LastChan%
For chan% = 1 To MatchTmpSample(1).LastChan%
ip% = IPOS1(MatchOldSample(1).LastChan%, MatchTmpSample(1).Elsyms$(chan%), MatchOldSample(1).Elsyms$())
If ip% = 0 Then
If itemp% + 1 <= MAXCHAN% Then
itemp% = itemp% + 1
MatchOldSample(1).Elsyms$(itemp%) = MatchTmpSample(1).Elsyms$(chan%)
MatchOldSample(1).ElmPercents!(itemp%) = 0#
End If
End If
Next chan%

' Calculate sum of differences squared for matching elements
vector! = 0#
For chan% = 1 To itemp%
ip% = IPOS1(MatchTmpSample(1).LastChan%, MatchOldSample(1).Elsyms$(chan%), MatchTmpSample(1).Elsyms$())
If ip% > 0 Then
temp! = MatchOldSample(1).ElmPercents!(chan%) - MatchTmpSample(1).ElmPercents!(ip%)
vector! = vector! + temp! * temp!
End If
Next chan%
vector! = Sqr(vector!)

' Check vector with minimum and store standard number and vector
If vector! < minimumvector! Then
n% = n% + 1
numbers%(n%) = MatchTmpSample(1).number%
vectors!(n%) = vector!
End If

stds.MoveNext
Loop

' Check for any matches
If n% = 0 Then GoTo MatchStandardsNoMatch

' Load list with matched standards (sorted list)
tList.Clear
For i% = 1 To n%
ip% = StandardGetRow%(numbers%(i%))
If ip% > 0 Then
msg$ = "v = " & Format$(Format$(vectors!(i%), f82$), a80$) & ", " & StandardGetString$(ip%)
tList.AddItem msg$
tList.ItemData(tList.NewIndex) = StandardIndexNumbers%(ip%)
End If
Next i%

Exit Sub

' Errors
MatchStandardsError:
MsgBox Error$, vbOKOnly + vbCritical, "MatchStandards"
ierror = True
Exit Sub

MatchStandardsNoMatch:
msg$ = "No standards matched with the given composition."
msg$ = msg$ & vbCrLf & vbCrLf & "Please try increasing the Minimum Vector fit parameter or choose a match standard database containing more chemically related compositions."
MsgBox msg$, vbOKOnly + vbExclamation, "MatchStandards"
ierror = True
Exit Sub

MatchStandardsNoElements:
msg$ = "No elements in given composition"
MsgBox msg$, vbOKOnly + vbExclamation, "MatchStandards"
ierror = True
Exit Sub

End Sub

Sub MatchLoadWeight()
' Get a user specified sample composition using weight string

ierror = False
On Error GoTo MatchLoadWeightError

Dim tmsg As String, astring As String

' Create default string
astring$ = MatchGetWeightString(MatchOldSample())
If ierror Then Exit Sub

' Load WEIGHT form and get user weight percents
FormWEIGHT.TextWeightPercentString.Text = astring$
FormWEIGHT.Show vbModal
If icancel Then Exit Sub

' Get user modified sample
Call FormulaReturnSample(MatchOldSample())
If ierror Then Exit Sub

' Get sample string to display
tmsg$ = TypeWeight$(Int(0), MatchOldSample())
FormMATCH.TextComposition.Text = tmsg$

Exit Sub

' Errors
MatchLoadWeightError:
MsgBox Error$, vbOKOnly + vbCritical, "MatchLoadWeight"
ierror = True
Exit Sub

End Sub

Sub MatchSample()
' Match the current sample to the match standard database

ierror = False
On Error GoTo MatchSampleError

Dim minimumvector As Single

' Get minimum vector
minimumvector! = Val(FormMATCH.TextMinimumVector.Text)

' Load default match database (normally DHZ)
Call MatchSave(Int(1))
If ierror Then Exit Sub

' Call the atual procedure to match to standards
Call MatchStandards(StandardDataFile$, minimumvector!, FormMATCH.ListStandards)

' Restore original database
Call MatchSave(Int(2))
If ierror Then Exit Sub

Exit Sub

' Errors
MatchSampleError:
MsgBox Error$, vbOKOnly + vbCritical, "MatchSample"
ierror = True
Exit Sub

End Sub

Function MatchGetWeightString(sample() As TypeSample) As String
' Load the current sample as a weight string

ierror = False
On Error GoTo MatchGetWeightStringError

Dim i As Integer
Dim astring As String

' Load all elements as integer weight percents
astring$ = vbNullString
For i% = 1 To sample(1).LastChan%
If sample(1).ElmPercents!(i%) > 0# Then
astring$ = astring$ & sample(1).Elsyms$(i%) & Int(sample(1).ElmPercents!(i%) + 0.5)
End If
Next i%

MatchGetWeightString$ = astring$
Exit Function

' Errors
MatchGetWeightStringError:
MsgBox Error$, vbOKOnly + vbCritical, "MatchGetWeightString"
ierror = True
Exit Function

End Function

Sub MatchOpenDatabase(tForm As Form)
' Allow user to select different match database

ierror = False
On Error GoTo MatchOpenDatabaseError

Dim tfilename As String

' Get old standard filename
tfilename$ = DefaultMatchStandardDatabase$
Call IOGetMDBFileName(Int(4), tfilename$, tForm)
If ierror Then Exit Sub

' Check the database type
Call FileInfoLoadData(Int(1), tfilename$)
If ierror Then Exit Sub

' Check if the standard database needs to be updated
If tfilename$ <> vbNullString Then
Call StandardUpdateMDBFile(tfilename$)
If ierror Then Exit Sub
End If

' Clear match list
FormMATCH.ListStandards.Clear

' No errors, load filename
DefaultMatchStandardDatabase$ = MiscGetFileNameOnly$(tfilename$)
FormMATCH.Caption = "Match Unknown To Match Database [" & DefaultMatchStandardDatabase$ & "]"

Exit Sub

' Errors
MatchOpenDatabaseError:
MsgBox Error$, vbOKOnly + vbCritical, "MatchOpenDatabase"
ierror = True
Exit Sub

End Sub

Sub MatchSave(mode As Integer)
' Restore default database
'  1 = load default match standard database
'  2 = load default standard database

ierror = False
On Error GoTo MatchSaveError

' Load match database to default match
If mode% = 1 Then
If StandardDataFile$ = vbNullString Then StandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
OriginalStandardDataFile$ = StandardDataFile$
StandardDataFile$ = ApplicationCommonAppData$ & DefaultMatchStandardDatabase$

' Check that it exists
If Dir$(StandardDataFile$) = vbNullString Then GoTo MatchSaveNotFound

' Update the selected match database in case if needs to be updated
If StandardDataFile$ <> vbNullString Then
Call StandardUpdateMDBFile(StandardDataFile$)
If ierror Then Exit Sub
End If

' Restore default
Else
If OriginalStandardDataFile$ <> vbNullString$ Then StandardDataFile$ = OriginalStandardDataFile$
If StandardDataFile$ = vbNullString Then StandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
End If

' Get the standard numbers and names
Call StandardGetMDBIndex
If ierror Then Exit Sub

Exit Sub

' Errors
MatchSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "MatchSave"
StandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
Call StandardGetMDBIndex
ierror = True
Exit Sub

MatchSaveNotFound:
msg$ = "The specified match database " & DefaultMatchStandardDatabase$ & " was not found in the application data folder " & ApplicationCommonAppData$
MsgBox Error$, vbOKOnly + vbCritical, "MatchSave"
StandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
Call StandardGetMDBIndex
ierror = True
Exit Sub

End Sub

Sub MatchTypeStandard()
' Type out the standard composition

ierror = False
On Error GoTo MatchTypeStandardError

' Print the standard (based on match sample flags)
Call MatchTypeStandard2(MatchOldSample())
If ierror Then Exit Sub

Exit Sub

' Errors
MatchTypeStandardError:
MsgBox Error$, vbOKOnly + vbCritical, "MatchTypeStandard"
ierror = True
Exit Sub

End Sub


