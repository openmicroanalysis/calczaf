Attribute VB_Name = "CodeSTANDAR2"
' (c) Copyright 1995-2022 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Dim stdtmpsample(1 To 1) As TypeSample

Sub StandardGetMDBIndex()
' Open the database and get the standard numbers and names

ierror = False
On Error GoTo StandardGetMDBIndexError

Dim StDb As Database
Dim stds As Recordset

' Open the standard database
Screen.MousePointer = vbHourglass
If Trim$(StandardDataFile$) = vbNullString Then GoTo StandardGetMDBIndexNoFilename
If Dir$(StandardDataFile$) = vbNullString Then GoTo StandardGetMDBIndexNotFound

Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseNonExclusiveAccess%, dbReadOnly)

' Open the "Standard" table as a Recordset
Set stds = StDb.OpenRecordset("Standard", dbOpenSnapshot)

' Loop through the table and load all standard numbers and names
Call InitStandardIndex
If ierror Then Exit Sub

Do Until stds.EOF
If NumberOfAvailableStandards% + 1 > MAXINDEX% Then GoTo StandardGetMDBIndexTooMany
NumberOfAvailableStandards% = NumberOfAvailableStandards% + 1
StandardIndexNumbers%(NumberOfAvailableStandards%) = stds("Numbers")
StandardIndexNames$(NumberOfAvailableStandards%) = Trim$(vbNullString & stds("Names"))
StandardIndexDescriptions$(NumberOfAvailableStandards%) = Trim$(vbNullString & stds("Descriptions"))
StandardIndexDensities!(NumberOfAvailableStandards%) = stds("Densities")
StandardIndexMaterialTypes$(NumberOfAvailableStandards%) = Trim$(vbNullString & stds("MaterialTypes"))
StandardIndexMountNames$(NumberOfAvailableStandards%) = Trim$(vbNullString & stds("MountNames"))

stds.MoveNext
Loop
stds.Close

' Close standard database
StDb.Close

Screen.MousePointer = vbDefault
Exit Sub

' Errors
StandardGetMDBIndexError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardGetMDBIndex"
ierror = True
Exit Sub

StandardGetMDBIndexNoFilename:
Screen.MousePointer = vbDefault
msg$ = "Standard database file name is blank. Please specify the standard database file."
MsgBox msg$, vbOKOnly + vbExclamation, "StandardGetMDBIndex"
ierror = True
Exit Sub

StandardGetMDBIndexNotFound:
Screen.MousePointer = vbDefault
msg$ = "The standard database file " & StandardDataFile$ & " was not found."
MsgBox msg$, vbOKOnly + vbExclamation, "StandardGetMDBIndex"
ierror = True
Exit Sub

StandardGetMDBIndexTooMany:
Screen.MousePointer = vbDefault
msg$ = "Too many standards were found in " & StandardDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardGetMDBIndex"
ierror = True
Exit Sub

End Sub

Sub StandardGetMDBStandard(stdnum As Integer, sample() As TypeSample)
' Load standard composition based on standard number from the standard database

ierror = False
On Error GoTo StandardGetMDBStandardError

Dim i As Integer, ip As Integer
Dim StDb As Database
Dim StDt As Recordset
Dim stds As Recordset
Dim SQLQ As String

' If stdnum equals zero then load a random composition
If stdnum% = 0 Then
Call StandardGetRandomComposition(sample())
If ierror Then Exit Sub
Exit Sub
End If

' Initialize the Tmp sample
Call InitSample(sample())
If ierror Then Exit Sub

' Open the standard database
Screen.MousePointer = vbHourglass
If Trim$(StandardDataFile$) = vbNullString Then GoTo StandardGetMDBStandardNoFilename
If Dir$(StandardDataFile$) = vbNullString Then GoTo StandardGetMDBStandardNotFound
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseNonExclusiveAccess%, dbReadOnly)
Set StDt = StDb.OpenRecordset("Standard", dbOpenTable)      ' use dbOpenTable to support .Seek method below
StDt.Index = "Standard Numbers"
StDt.Seek "=", stdnum%
If StDt.NoMatch Then GoTo StandardGetMDBStandardNumberNotFound

' Load values from "Standard" table
sample(1).number% = StDt("Numbers")

' Check for Null values from database
sample(1).Name$ = Trim$(vbNullString & StDt("Names"))
sample(1).Description$ = Trim$(vbNullString & StDt("Descriptions"))
sample(1).DisplayAsOxideFlag% = StDt("DisplayAsOxideFlags")
sample(1).SampleDensity! = StDt("Densities")
sample(1).MaterialType$ = Trim$(vbNullString & StDt("MaterialTypes"))
sample(1).MountNames$ = Trim$(vbNullString & StDt("MountNames"))

sample(1).FormulaElementFlag = StDt("FormulaFlags")
sample(1).FormulaRatio! = Val(StDt("FormulaRatios"))
sample(1).FormulaElement$ = Trim$(vbNullString & StDt("FormulaElements"))
StDt.Close

' Get standard composition data for specified standard from standard database
SQLQ$ = "SELECT Element.* FROM Element WHERE Element.Number = " & Str$(stdnum%)
Set stds = StDb.OpenRecordset(SQLQ$, dbOpenSnapshot, dbReadOnly)

' Load all elements from "Element" table that matched the standard number
sample(1).LastChan% = 0
Do Until stds.EOF
If sample(1).LastChan% + 1 > MAXCHAN% Then GoTo StandardGetMDBStandardTooManyElements
sample(1).LastChan% = sample(1).LastChan% + 1
sample(1).Elsyms$(sample(1).LastChan%) = Trim$(vbNullString & stds("Symbol"))
sample(1).ElmPercents!(sample(1).LastChan%) = stds("Percent")
sample(1).numcat%(sample(1).LastChan%) = stds("NumCat")
sample(1).numoxd%(sample(1).LastChan%) = stds("NumOxd")

ip% = IPOS1(MAXELM%, sample(1).Elsyms$(sample(1).LastChan%), Symlo$())
If ip% > 0 Then
sample(1).AtomicCharges!(sample(1).LastChan%) = AllAtomicCharges!(ip%)
End If
stds.MoveNext
Loop

' Close the standard database
stds.Close
StDb.Close

' All standards are elemental!!!!!!
sample(1).OxideOrElemental% = 2

' Update other sample parameters
sample(1).Type% = 1
sample(1).LastElm% = sample(1).LastChan%

' Load default "TakeOff" and "KiloVolts". These are overwritten
' in UpdateStdKfacs but used in StanFormCalculate
sample(1).takeoff! = DefaultTakeOff!
sample(1).kilovolts! = DefaultKiloVolts!
sample(1).beamcurrent! = DefaultBeamCurrent!
sample(1).beamsize! = DefaultBeamSize!
sample(1).ColumnConditionMethod% = DefaultColumnConditionMethod%
sample(1).ColumnConditionString$ = DefaultColumnConditionString$

' Load x-ray lines from element defaults. These are overwritten
' in UpdateStdKfacs but used in StanFormCalculate
For i% = 1 To sample(1).LastChan%
ip% = IPOS1(MAXELM%, sample(1).Elsyms$(i%), Symlo$())
If ip% > 0 Then sample(1).Xrsyms$(i%) = Deflin$(ip%)
Next i%

' Load default crystal types (only used by InterfSave)
For i% = 1 To sample(1).LastChan%
ip% = IPOS1(MAXELM%, sample(1).Elsyms$(i%), Symlo$())
If ip% > 0 Then sample(1).CrystalNames$(i%) = Defcry$(ip%)
Next i%

' Load default atomic charge and atomic numbers
For i% = 1 To sample(1).LastChan%
ip% = IPOS1(MAXELM%, sample(1).Elsyms$(i%), Symlo$())
If ip% > 0 Then sample(1).AtomicCharges!(i%) = AllAtomicCharges!(ip%)
If ip% > 0 Then sample(1).AtomicNums%(i%) = AllAtomicNums%(ip%)
If ip% > 0 Then sample(1).AtomicWts!(i%) = AllAtomicWts!(ip%)
Next i%

' Load kilovolts array
For i% = 1 To sample(1).LastChan%
sample(1).TakeoffArray!(i%) = sample(1).takeoff!
sample(1).KilovoltsArray!(i%) = sample(1).kilovolts!
sample(1).BeamCurrentArray(i%) = sample(1).beamcurrent!
sample(1).BeamSizeArray(i%) = sample(1).beamsize!
Next i%

Screen.MousePointer = vbDefault
Exit Sub

' Errors
StandardGetMDBStandardError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardGetMDBStandard"
ierror = True
Exit Sub

StandardGetMDBStandardNoFilename:
Screen.MousePointer = vbDefault
msg$ = "Standard database file name is blank. Please specify the standard database file."
MsgBox msg$, vbOKOnly + vbExclamation, "StandardGetMDBStandard"
ierror = True
Exit Sub

StandardGetMDBStandardNotFound:
Screen.MousePointer = vbDefault
msg$ = "The standard database file " & StandardDataFile$ & " was not found."
MsgBox msg$, vbOKOnly + vbExclamation, "StandardGetMDBStandard"
ierror = True
Exit Sub

StandardGetMDBStandardNumberNotFound:
Screen.MousePointer = vbDefault
msg$ = "Standard number " & Format$(stdnum%) & " was not found in the standard database"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardGetMDBStandard"
ierror = True
Exit Sub

StandardGetMDBStandardTooManyElements:
Screen.MousePointer = vbDefault
msg$ = "Too many elements in standard number " & Format$(sample(1).number%) & " in " & StandardDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardGetMDBStandard"
ierror = True
Exit Sub

End Sub

Function StandardGetRow(stdnum As Integer) As Integer
' Returns the standard index row number of this standard number

ierror = False
On Error GoTo StandardGetRowError

StandardGetRow = IPOS2%(NumberOfAvailableStandards%, stdnum%, StandardIndexNumbers%())

Exit Function

' Errors
StandardGetRowError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardGetRow"
ierror = True
Exit Function

End Function

Function StandardGetString(StdIndex As Integer) As String
' Returns as formatted standard string

ierror = False
On Error GoTo StandardGetStringError

StandardGetString$ = Format$(StandardIndexNumbers(StdIndex%), a40$) & " " & StandardIndexNames$(StdIndex%)

Exit Function

' Errors
StandardGetStringError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardGetString"
ierror = True
Exit Function

End Function

Function StandardGetString2(stdnum As Integer) As String
' Returns as formatted standard string

ierror = False
On Error GoTo StandardGetString2Error

StandardGetString2$ = Format$(StandardIndexNumbers(StandardGetRow%(stdnum%)), a40$) & " " & StandardIndexNames$(StandardGetRow%(stdnum%))

Exit Function

' Errors
StandardGetString2Error:
MsgBox Error$, vbOKOnly + vbCritical, "StandardGetString2"
ierror = True
Exit Function

End Function

Sub StandardLoadList(tList As ListBox)
' Routine to update the available standard list box

ierror = False
On Error GoTo StandardLoadListError

Dim i As Integer

' List the available standards
Screen.MousePointer = vbHourglass
tList.Clear
For i% = 1 To NumberOfAvailableStandards%
If StandardIndexNumbers%(i%) > 0 Then
tList.AddItem StandardGetString$(i%)
tList.ItemData(tList.NewIndex) = StandardIndexNumbers%(i%)
End If
Next i%

Screen.MousePointer = vbDefault
Exit Sub

' Errors
StandardLoadListError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardLoadList"
ierror = True
Exit Sub

End Sub

Sub StandardLoadList2(tList As ComboBox)
' Routine to update the available standard combo

ierror = False
On Error GoTo StandardLoadList2Error

Dim i As Integer

' List the available standards
Screen.MousePointer = vbHourglass
tList.Clear
For i% = 1 To NumberOfAvailableStandards%
If StandardIndexNumbers%(i%) > 0 Then
tList.AddItem StandardGetString$(i%)
tList.ItemData(tList.NewIndex) = StandardIndexNumbers%(i%)
End If
Next i%

Screen.MousePointer = vbDefault
Exit Sub

' Errors
StandardLoadList2Error:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardLoadList2"
ierror = True
Exit Sub

End Sub

Function StandardGetNumber() As Integer
' Get the next available free standard number

ierror = False
On Error GoTo StandardGetNumberError

Dim n As Integer, ip As Integer

n% = 0
Do
n% = n% + 1
If n% > MAXINTEGER% Then GoTo StandardGetNumberNoRoom
ip% = StandardGetRow%(n%)
Loop Until ip% = 0

StandardGetNumber% = n%
Exit Function

' Errors
StandardGetNumberError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardGetNumber"
ierror = True
Exit Function

StandardGetNumberNoRoom:
msg$ = "No room in the standard database for a new standard composition"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardGetNumber"
ierror = True
Exit Function

End Function

Sub StandardGetRandomComposition(sample() As TypeSample)
' Loads in a random composition (random normalized to 100%)

ierror = False
On Error GoTo StandardGetRandomCompositionError

Dim i As Integer
Dim sum As Single, sum2 As Single

' Create a random composition based on the analyzed elements
sum! = 0#
sample(1).LastChan% = sample(1).LastElm%
For i% = 1 To sample(1).LastElm%
sample(1).ElmPercents!(i%) = Rnd * 100#
sum! = sum! + sample(1).ElmPercents!(i%)
Next i%

' If calculating with stoichiometric oxygen, add it in
If sample(1).OxideOrElemental% = 1 And sample(1).LastElm% + 1 <= MAXCHAN1% Then
sum2! = 0#
For i% = 1 To sample(1).LastElm%
sum2! = sum2! + ConvertElmToOxd(sample(1).ElmPercents!(i%), sample(1).Elsyms$(i%), sample(1).numcat%(i%), sample(1).numoxd%(i%)) - sample(1).ElmPercents!(i%)
Next i%

' Add in stoichiometric oxygen to element arrays
sample(1).LastChan% = sample(1).LastElm% + 1
sample(1).Elsyms$(sample(1).LastChan%) = Symup$(ATOMIC_NUM_OXYGEN%)
sample(1).Xrsyms$(sample(1).LastChan%) = vbNullString
sample(1).numcat%(sample(1).LastChan%) = AllCat%(ATOMIC_NUM_OXYGEN%)
sample(1).numoxd%(sample(1).LastChan%) = AllOxd%(ATOMIC_NUM_OXYGEN%)
sample(1).AtomicCharges!(sample(1).LastChan%) = AllAtomicCharges!(ATOMIC_NUM_OXYGEN%)
sample(1).ElmPercents!(sample(1).LastChan%) = sum2!

' Add to elemental sum
sum! = sum! + sum2!
End If

' Normalize to 100 %
For i% = 1 To sample(1).LastChan%
sample(1).ElmPercents!(i%) = sample(1).ElmPercents!(i%) * 100# / sum!
Next i%

' Loop on analyzed elements and load default conditions
For i% = 1 To sample(1).LastElm%

' Load KeV array
If Not sample(1).CombinedConditionsFlag Then
sample(1).TakeoffArray!(i%) = sample(1).takeoff!
sample(1).KilovoltsArray!(i%) = sample(1).kilovolts!
sample(1).BeamCurrentArray!(i%) = sample(1).beamcurrent!
sample(1).BeamSizeArray!(i%) = sample(1).beamsize!

sample(1).ColumnConditionMethodArray%(i%) = sample(1).ColumnConditionMethod%
sample(1).ColumnConditionStringArray$(i%) = sample(1).ColumnConditionString$
End If

Next i%

' Disable stoichiometric oxygen parameters if present (not really necessary as they are added back in in ZAFStd2)
For i% = sample(1).LastElm% + 1 To sample(1).LastChan%

' Load KeV array
sample(1).TakeoffArray!(i%) = 0#
sample(1).KilovoltsArray!(i%) = 0#
sample(1).BeamCurrentArray!(i%) = 0#
sample(1).BeamSizeArray!(i%) = 0#

sample(1).ColumnConditionMethodArray%(i%) = 0
sample(1).ColumnConditionStringArray$(i%) = vbNullString
Next i%

Exit Sub

' Errors
StandardGetRandomCompositionError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardGetRandomComposition"
ierror = True
Exit Sub

End Sub

Sub StandardSelectList(stdnum As Integer, tList As ListBox)
' Select the list item based on itemdata

ierror = False
On Error GoTo StandardSelectListError

Dim i As Integer

' Unselect all
For i% = 0 To tList.ListCount - 1
tList.Selected(i%) = False
Next i%

' Look through list
For i% = 0 To tList.ListCount - 1
If tList.ItemData(i%) = stdnum% Then
tList.ListIndex = i%
tList.Selected(i%) = True
Exit Sub
End If
Next i%

Exit Sub

' Errors
StandardSelectListError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardSelectList"
ierror = True
Exit Sub

End Sub

Sub StandardTypeStandard(stdnum As Integer)
' Type the standard composition for a standard, no calculations

ierror = False
On Error GoTo StandardTypeStandardError

Dim n As Integer, i As Integer
Dim ii As Integer, jj As Integer
Dim tmsg As String
Dim sum As Single

' Get composition of standard
Call StandardGetMDBStandard(stdnum%, stdtmpsample())
If ierror Then Exit Sub

' Load standard description (no conditions)
tmsg$ = vbCrLf & StandardLoadDescription3(stdtmpsample())
If ierror Then Exit Sub

' Type out standard data
n = 0
Do Until False
n% = n% + 1
Call TypeGetRange(Int(2), n%, ii%, jj%, stdtmpsample())
If ierror Then Exit Sub
If ii% > stdtmpsample(1).LastChan% Then Exit Do

' Load standard composition
tmsg$ = tmsg$ & vbCrLf
tmsg$ = tmsg$ & vbCrLf & "ELEM: "
For i% = ii% To jj%
tmsg$ = tmsg$ & Format$(MiscAutoUcase$(stdtmpsample(1).Elsyms$(i%)), a80$)
Next i%

tmsg$ = tmsg$ & Format$("SUM", a80$)

sum! = 0#
tmsg$ = tmsg$ & vbCrLf & "ELWT: "
For i% = ii% To jj%
tmsg$ = tmsg$ & Format$(Format$(stdtmpsample(1).ElmPercents!(i%), f83$), a80$)
sum! = sum! + stdtmpsample(1).ElmPercents!(i%)
Next i%

tmsg$ = tmsg$ & Format$(Format$(sum!, f83$), a80$)
Loop

Call IOWriteLog(tmsg$)

Exit Sub

' Errors
StandardTypeStandardError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardTypeStandard"
ierror = True
Exit Sub

End Sub

Function StandardIsDemoStandardDatabaseLoaded() As Boolean
' Check whether demo standard database is loaded

ierror = False
On Error GoTo StandardIsDemoStandardDatabaseLoadedError

Dim ip As Integer

StandardIsDemoStandardDatabaseLoaded = True

Call InitSample(stdtmpsample())
If ierror Then Exit Function

' Check for Co std
Call StandardGetMDBStandard(Int(527), stdtmpsample())
If ierror Then Exit Function

ip% = IPOS1%(stdtmpsample(1).LastChan%, "co", stdtmpsample(1).Elsyms$())
If ip% = 0 Then
StandardIsDemoStandardDatabaseLoaded = False
Exit Function
End If

' Check for Cu std
Call StandardGetMDBStandard(Int(529), stdtmpsample())
If ierror Then Exit Function

ip% = IPOS1%(stdtmpsample(1).LastChan%, "cu", stdtmpsample(1).Elsyms$())
If ip% = 0 Then
StandardIsDemoStandardDatabaseLoaded = False
Exit Function
End If

' Check for TiO2 std
Call StandardGetMDBStandard(Int(22), stdtmpsample())
If ierror Then Exit Function

ip% = IPOS1%(stdtmpsample(1).LastChan%, "ti", stdtmpsample(1).Elsyms$())
If ip% = 0 Then
StandardIsDemoStandardDatabaseLoaded = False
Exit Function
End If

' Check for SiO2 std
Call StandardGetMDBStandard(Int(14), stdtmpsample())
If ierror Then Exit Function

ip% = IPOS1%(stdtmpsample(1).LastChan%, "si", stdtmpsample(1).Elsyms$())
If ip% = 0 Then
StandardIsDemoStandardDatabaseLoaded = False
Exit Function
End If

Exit Function

' Errors
StandardIsDemoStandardDatabaseLoadedError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardIsDemoStandardDatabaseLoaded"
ierror = True
Exit Function

End Function

Sub StandardAddRecord(sample() As TypeSample)
' Routine to append a standard to the standard database
' Called by StandardReadDATFile and StandardReplaceRecord

ierror = False
On Error GoTo StandardAddRecordError

Dim i As Integer
Dim StDb As Database
Dim StDt As Recordset

' Open the database and the "Standard" table
Screen.MousePointer = vbHourglass
If StandardDataFile$ = vbNullString Then StandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseExclusiveAccess%, False)
Set StDt = StDb.OpenRecordset("Standard", dbOpenTable)

' Check that standard number does not already exist
StDt.Index = "Standard Numbers"
StDt.Seek "=", sample(1).number%
If Not StDt.NoMatch Then GoTo StandardAddRecordExists

Call TransactionBegin("StandardAddRecord", StandardDataFile$)
If ierror Then Exit Sub

' Add new record to "Standard" table
StDt.AddNew
StDt("Numbers") = sample(1).number%
StDt("Names") = Left$(sample(1).Name$, DbTextNameLength%)
StDt("Descriptions") = Left$(sample(1).Description$, DbTextDescriptionLength%)
StDt("DisplayAsOxideFlags") = sample(1).DisplayAsOxideFlag%
StDt("Densities") = sample(1).SampleDensity!

StDt("FormulaFlags") = sample(1).FormulaElementFlag
StDt("FormulaRatios") = sample(1).FormulaRatio!
StDt("FormulaElements") = Left$(sample(1).FormulaElement$, DbTextElementStringLength%)

StDt("MaterialTypes") = Left$(sample(1).MaterialType$, DbTextNameLength%)
StDt("MountNames") = Left$(sample(1).MountNames$, DbTextDescriptionLength%)

StDt.Update
StDt.Close

' Add element symbols and weights to "Element" table
Set StDt = StDb.OpenRecordset("Element", dbOpenTable)
For i% = 1 To sample(1).LastChan%
StDt.AddNew
StDt("Number") = sample(1).number%
StDt("Symbol") = Left$(sample(1).Elsyms$(i%), DbTextXrayStringLength%)
StDt("Percent") = sample(1).ElmPercents!(i%)
StDt("NumCat") = sample(1).numcat%(i%)
StDt("NumOxd") = sample(1).numoxd%(i%)
StDt.Update
Next i%

Call TransactionCommit("StandardAddRecord", StandardDataFile$)
If ierror Then Exit Sub

StDt.Close
StDb.Close
Screen.MousePointer = vbDefault
Exit Sub

' Errors
StandardAddRecordError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardAddRecord"
Call TransactionRollback("StandardAddRecord", StandardDataFile$)
ierror = True
Exit Sub

StandardAddRecordExists:
Screen.MousePointer = vbDefault
msg$ = "Standard number " & Str$(sample(1).number%) & " already exists in " & StandardDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardAddRecord"
ierror = True
Exit Sub

End Sub

Function StandardGetDensity(stdnum As Integer, tStandardDataFile As String) As Single
' Routine to get the density (only) for the specified standard for the specified standard database

ierror = False
On Error GoTo StandardGetDensityError

Dim StDb As Database
Dim StRs As Recordset

Dim SQLQ As String

' Open the database and the "Standard" table
Screen.MousePointer = vbHourglass
If tStandardDataFile$ = vbNullString Then tStandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
Set StDb = OpenDatabase(tStandardDataFile$, StandardDatabaseNonExclusiveAccess%, True)

' Check that standard number already exists
SQLQ$ = "SELECT Standard.* FROM Standard WHERE Standard.Numbers = " & Format$(stdnum%)
Set StRs = StDb.OpenRecordset(SQLQ$, dbOpenSnapshot)
If StRs.BOF And StRs.EOF Then GoTo StandardGetDensityDoesNotExist
StandardGetDensity! = StRs("Densities")
StRs.Close

StDb.Close
Screen.MousePointer = vbDefault
Exit Function

' Errors
StandardGetDensityError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardGetDensity"
ierror = True
Exit Function

StandardGetDensityDoesNotExist:
Screen.MousePointer = vbDefault
msg$ = "Standard number " & Str$(stdnum%) & " does not exists in " & tStandardDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardGetDensity"
ierror = True
Exit Function

End Function

Function StandardLoadDescription(sample() As TypeSample) As String
' This routine returns the sample description string (full text version)

ierror = False
On Error GoTo StandardLoadDescriptionError

Dim tmsg As String

' Load sample name string
StandardLoadDescription$ = vbNullString
tmsg$ = TypeLoadString(sample())
If ierror Then Exit Function

' Load take off angle and kilovolts
tmsg$ = tmsg$ & vbCrLf & "TakeOff = " & Format$(Format$(sample(1).takeoff!, f41$), a40$)
tmsg$ = tmsg$ & "  KiloVolt = " & Format$(Format$(sample(1).kilovolts!, f41$), a40$)
tmsg$ = tmsg$ & "  Density = " & Format$(Format$(sample(1).SampleDensity!, f63$), a60$)
If Trim$(sample(1).MaterialType$) <> vbNullString Then tmsg$ = tmsg$ & "  Type = " & sample(1).MaterialType$
If Trim$(sample(1).MountNames$) <> vbNullString Then tmsg$ = tmsg$ & "  Mount = " & sample(1).MountNames$

' Add sample description field
If sample(1).Description$ <> vbNullString Then
tmsg$ = tmsg$ & vbCrLf & vbCrLf & sample(1).Description$
End If

StandardLoadDescription$ = tmsg$
Exit Function

' Errors
StandardLoadDescriptionError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardLoadDescription"
ierror = True
Exit Function

End Function

Function StandardLoadDescription2(sample() As TypeSample) As String
' This routine returns the sample description string (short version)

ierror = False
On Error GoTo StandardLoadDescription2Error

Dim tmsg As String

' Load sample name string
StandardLoadDescription2$ = vbNullString
tmsg$ = TypeLoadString(sample())
If ierror Then Exit Function

' Load take off angle and kilovolts
tmsg$ = tmsg$ & vbCrLf & "TO = " & Str$(sample(1).takeoff!) & ", KeV = " & Str$(sample(1).kilovolts!)

' Add sample description field
If sample(1).Description$ <> vbNullString Then
tmsg$ = tmsg$ & vbCrLf & sample(1).Description$
End If

StandardLoadDescription2$ = tmsg$
Exit Function

' Errors
StandardLoadDescription2Error:
MsgBox Error$, vbOKOnly + vbCritical, "StandardLoadDescription2"
ierror = True
Exit Function

End Function

Function StandardLoadDescription3(sample() As TypeSample) As String
' This routine returns the sample description string (no takeoff or keV)

ierror = False
On Error GoTo StandardLoadDescription3Error

Dim tmsg As String

' Load sample name string
StandardLoadDescription3$ = vbNullString
tmsg$ = TypeLoadString(sample())
If ierror Then Exit Function

' Load take off angle and kilovolts
tmsg$ = tmsg$ & "  Density = " & Format$(Format$(sample(1).SampleDensity!, f63$), a60$)
If Trim$(sample(1).MaterialType$) <> vbNullString Then tmsg$ = tmsg$ & "  Type = " & sample(1).MaterialType$
If Trim$(sample(1).MountNames$) <> vbNullString Then tmsg$ = tmsg$ & "  Mount = " & sample(1).MountNames$

' Add sample description field
If sample(1).Description$ <> vbNullString Then
tmsg$ = tmsg$ & vbCrLf & vbCrLf & sample(1).Description$
End If

StandardLoadDescription3$ = tmsg$
Exit Function

' Errors
StandardLoadDescription3Error:
MsgBox Error$, vbOKOnly + vbCritical, "StandardLoadDescription3"
ierror = True
Exit Function

End Function

Sub StandardDeleteRecord(stdnum As Integer)
' Routine to delete a standard number from the database

ierror = False
On Error GoTo StandardDeleteRecordError

Dim StDb As Database
Dim StDt As Recordset
Dim SQLQ As String

' Open the database and the "Standard" table
If StandardDataFile$ = vbNullString Then StandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseExclusiveAccess%, False)
Set StDt = StDb.OpenRecordset("Standard", dbOpenTable)

' Check that standard number already exists
StDt.Index = "Standard Numbers"
StDt.Seek "=", stdnum%
If StDt.NoMatch Then GoTo StandardDeleteRecordNotFound

' Delete the record in "Standard" table and update the database
Screen.MousePointer = vbHourglass

Call TransactionBegin("StandardDeleteRecord", StandardDataFile$)
If ierror Then Exit Sub

StDt.Delete
StDt.MoveFirst
StDt.Close

' Delete element symbols and weights to "Element" table based on "stdnum"
SQLQ$ = "DELETE from Element WHERE Element.Number = " & Str$(stdnum%)
StDb.Execute SQLQ$

Call TransactionCommit("StandardDeleteRecord", StandardDataFile$)
If ierror Then Exit Sub

StDb.Close
Screen.MousePointer = vbDefault

Exit Sub

' Errors
StandardDeleteRecordError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardDeleteRecord"
Call TransactionRollback("StandardDeleteRecord", StandardDataFile$)
ierror = True
Exit Sub

StandardDeleteRecordNotFound:
Screen.MousePointer = vbDefault
msg$ = "Standard number " & Str$(stdnum%) & " was not found in " & StandardDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardDeleteRecord"
ierror = True
Exit Sub

End Sub

Sub StandardFindString(mode As Integer, astring As String, tList As ListBox)
' Find standard based on name string
' mode = 0 start at beginning
' mode = 1 start at standard found

ierror = False
On Error GoTo StandardFindStringError

Dim i As Integer, stdnum As Integer

Static last_match As Integer

If Trim$(astring$) = vbNullString Then Exit Sub

' Find first match
If mode% = 0 Then
For i% = 1 To NumberOfAvailableStandards%
If InStr(UCase$(StandardIndexNames$(i%)), UCase$(astring$)) > 0 Then GoTo 1000
Next i%

Exit Sub

' Load substance number found
1000:
stdnum% = StandardIndexNumbers%(i%)

' Select item in list
Call StandardSelectList(stdnum%, tList)
If ierror Then Exit Sub

last_match% = i%
Exit Sub
End If

' Look for next match
If mode% = 1 Then
For i% = last_match% + 1 To NumberOfAvailableStandards%
If InStr(UCase$(StandardIndexNames$(i%)), UCase$(astring$)) > 0 Then GoTo 2000
Next i%

' Load next substance number found
2000:
stdnum% = StandardIndexNumbers%(i%)

' Select item in list
Call StandardSelectList(stdnum%, tList)
If ierror Then Exit Sub

last_match% = i%
End If

Exit Sub

' Errors
StandardFindStringError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardFindString"
ierror = True
Exit Sub

End Sub

Sub StandardFindNumber(mode As Integer, astring As String, tList As ListBox)
' Find standard based on standard number
' mode = 0 start at beginning
' mode = 1 start at standard found

ierror = False
On Error GoTo StandardFindNumberError

Dim i As Integer, stdnum As Integer

Static last_match As Integer

If Trim$(astring$) = vbNullString Then Exit Sub

' Find first match
If mode% = 0 Then
For i% = 1 To NumberOfAvailableStandards%
If InStr(Format$(StandardIndexNumbers%(i%)), astring$) > 0 Then GoTo 1000
Next i%

Exit Sub

' Load substance number found
1000:
stdnum% = StandardIndexNumbers%(i%)

' Select item in list
Call StandardSelectList(stdnum%, tList)
If ierror Then Exit Sub

last_match% = i%
Exit Sub
End If

' Look for next match
If mode% = 1 Then
For i% = last_match% + 1 To NumberOfAvailableStandards%
If InStr(Format$(StandardIndexNumbers%(i%)), astring) > 0 Then GoTo 2000
Next i%

' Load next substance number found
2000:
stdnum% = StandardIndexNumbers%(i%)

' Select item in list
Call StandardSelectList(stdnum%, tList)
If ierror Then Exit Sub

last_match% = i%
End If

Exit Sub

' Errors
StandardFindNumberError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardFindNumber"
ierror = True
Exit Sub

End Sub

Sub StandardUpdateMDBFile(tfilename As String)
' Routine to update the current standard MDB file automatically for new fields

ierror = False
On Error GoTo StandardUpdateMDBFileError

Dim StDb As Database
Dim StRs As Recordset

Dim versionnumber As Single
Dim updated As Integer
Dim SQLQ As String

Dim StdDensities As New Field

Dim EDSSpectra As TableDef              ' EDS spectra
Dim EDSSpectraIndex As New Index        ' EDS spectra index (to sample row numbers)
Dim EDSParameters As TableDef           ' EDS parameters
Dim EDSParametersIndex As New Index     ' EDS parameters index (to sample row numbers)

Dim CLSpectra As TableDef               ' CL spectra
Dim CLSpectraIndex As New Index         ' CL spectra index (to sample row numbers)
Dim CLParameters As TableDef            ' CL parameters
Dim CLParametersIndex As New Index      ' CL parameters index (to sample row numbers)

Dim EDSFileNames As New Field
Dim CLFileNames As New Field
Dim CLKilovolts As New Field

Dim StdKratios As TableDef              ' measured k-ratios

Dim StdKRatiosTakeoffs As New Field
Dim StdKRatiosKilovolts As New Field
Dim StdKRatiosElements As New Field
Dim StdKRatiosXrays As New Field
Dim StdKRatiosKRatios As New Field
Dim StdKRatiosStdAssigns As New Field
Dim StdKRatiosFileName As New Field

Dim StdMaterialTypes As New Field

Dim StdFormulaFlags As New Field
Dim StdFormulaRatios As New Field
Dim StdFormulaElements As New Field

Dim CLSpectraWavelengths As New Field
Dim CLSpectraIntensityDark As New Field

Dim MemoText As New TableDef
Dim MemoTextField As New Field
Dim MemoTextIndex As New Index      ' text memo index (to sample row numbers)

Dim StdMountNames As New Field

' Check for valid file name
If Trim$(tfilename$) = vbNullString Then
msg$ = "Standard data file name is blank"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardUpdateMDBFile"
ierror = True
Exit Sub
End If

' Check for standard database file
If Dir$(tfilename$) = vbNullString Then
msg$ = "Standard database " & tfilename$ & " was not found. Please re-start the program to re-create a new Standard database"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardUpdateMDBFile"
ierror = True
Exit Sub
End If

' Get version number
versionnumber! = FileInfoGetVersion!(tfilename$, "STANDARD")
If ierror Then Exit Sub

' If standard database version number is the same or higher than program version then just exit (no update needed)
If versionnumber! >= ProgramVersionNumber! Then Exit Sub

' Open the database
Screen.MousePointer = vbHourglass
Set StDb = OpenDatabase(tfilename$, StandardDatabaseExclusiveAccess%, False)

Call TransactionBegin("StandardUpdateMDBFile", tfilename$)
If ierror Then Exit Sub

' Flag file as not updated
updated = False

' Add standard density fields and records
If versionnumber! < 8.63 Then

' Add density field to Standard table
StdDensities.Name = "Densities"
StdDensities.Type = dbSingle
StDb.TableDefs("Standard").Fields.Append StdDensities

' Open standard table in standard database
Set StRs = StDb.OpenRecordset("Standard", dbOpenTable)

' Add default density fields to all records
Do Until StRs.EOF
StRs.Edit
StRs("Densities") = 5#      ' use this for now as default
StRs.Update
StRs.MoveNext
Loop

StRs.Close
updated = True
End If

' Add EDS and CL spectra fields and records
If versionnumber! < 10.93 Then

' Specify the standard database table "EDSSpectra" EDS spectra table
Set EDSSpectra = StDb.CreateTableDef("NewTableDef")
EDSSpectra.Name = "EDSSpectra"

With EDSSpectra
.Fields.Append .CreateField("EDSSpectraToNumber", dbInteger)            ' points back to Standard table/Numbers field
.Fields.Append .CreateField("EDSSpectraNumber", dbInteger)              ' for multiple spectra per standard
.Fields.Append .CreateField("EDSSpectraChannelOrder", dbInteger)        ' channel load order
.Fields.Append .CreateField("EDSSpectraIntensity", dbLong)              ' count data
End With

EDSSpectraIndex.Name = "EDS Spectra Numbers"
EDSSpectraIndex.Fields = "EDSSpectraToNumber"                           ' index to pointer to standard numbers
EDSSpectraIndex.Primary = False
EDSSpectra.Indexes.Append EDSSpectraIndex

StDb.TableDefs.Append EDSSpectra

' Create EDS parameters table for each data line
Set EDSParameters = StDb.CreateTableDef("NewTableDef")
EDSParameters.Name = "EDSParameters"

With EDSParameters
.Fields.Append .CreateField("EDSParametersToNumber", dbInteger)         ' points back to Standard table/Numbers field
.Fields.Append .CreateField("EDSParametersNumber", dbInteger)           ' for multiple spectral parameters per standard
.Fields.Append .CreateField("EDSParametersNumberofChannels", dbInteger)

.Fields.Append .CreateField("EDSParametersElapsedTime", dbSingle)
.Fields.Append .CreateField("EDSParametersDeadTime", dbSingle)
.Fields.Append .CreateField("EDSParametersLiveTime", dbSingle)

.Fields.Append .CreateField("EDSParametersEVPerChannel", dbSingle)
.Fields.Append .CreateField("EDSParametersStartEnergy", dbSingle)
.Fields.Append .CreateField("EDSParametersEndEnergy", dbSingle)
.Fields.Append .CreateField("EDSParametersTakeOff", dbSingle)
.Fields.Append .CreateField("EDSParametersAcceleratingVoltage", dbSingle)           ' in keV
End With

EDSParametersIndex.Name = "EDS Parameters Numbers"
EDSParametersIndex.Fields = "EDSParametersToNumber"                     ' index to pointer to sample rows
EDSParametersIndex.Primary = False
EDSParameters.Indexes.Append EDSParametersIndex

StDb.TableDefs.Append EDSParameters

' Specify the probe database table "CLSpectra" CL spectra table
Set CLSpectra = StDb.CreateTableDef("NewTableDef")
CLSpectra.Name = "CLSpectra"

With CLSpectra
.Fields.Append .CreateField("CLSpectraToNumber", dbInteger)         ' points back to Standard table/Numbers field
.Fields.Append .CreateField("CLSpectraNumber", dbInteger)           ' for multiple CL spectra per standard
.Fields.Append .CreateField("CLSpectraChannelOrder", dbInteger)     ' CL spectra channel load order
.Fields.Append .CreateField("CLSpectraIntensity", dbLong)           ' CL intensity count data
End With

CLSpectraIndex.Name = "CL Spectra Numbers"
CLSpectraIndex.Fields = "CLSpectraToNumber"                         ' index to pointer to standard number
CLSpectraIndex.Primary = False
CLSpectra.Indexes.Append CLSpectraIndex

StDb.TableDefs.Append CLSpectra

' Create CL parameters table for each data line
Set CLParameters = StDb.CreateTableDef("NewTableDef")
CLParameters.Name = "CLParameters"

With CLParameters
.Fields.Append .CreateField("CLParametersToNumber", dbInteger)      ' points back to Standard table/Numbers field
.Fields.Append .CreateField("CLParametersNumber", dbInteger)        ' for multiple CL spectral parameters per standard
.Fields.Append .CreateField("CLParametersCountTime", dbSingle)
.Fields.Append .CreateField("CLParametersNumberofChannels", dbInteger)
.Fields.Append .CreateField("CLParametersStartEnergy", dbSingle)
.Fields.Append .CreateField("CLParametersEndEnergy", dbSingle)
End With

CLParametersIndex.Name = "CL Parameters Numbers"
CLParametersIndex.Fields = "CLParametersToNumber" ' index to pointer to sample rows
CLParametersIndex.Primary = False
CLParameters.Indexes.Append CLParametersIndex

StDb.TableDefs.Append CLParameters

updated = True
End If

' Add additional EDS and CL spectra fields and records
If versionnumber! < 10.94 Then

' Create EDS parameters table
EDSFileNames.Name = "EDSParametersFileName"
EDSFileNames.Type = dbText
EDSFileNames.Size = DbTextFilenameLength%
EDSFileNames.AllowZeroLength = False

StDb.TableDefs("EDSParameters").Fields.Append EDSFileNames

' Add new fields to CL parameters table
CLFileNames.Name = "CLParametersFileName"
CLFileNames.Type = dbText
CLFileNames.Size = DbTextFilenameLength%
CLFileNames.AllowZeroLength = False

StDb.TableDefs("CLParameters").Fields.Append CLFileNames

CLKilovolts.Name = "CLParametersKilovolts"
CLKilovolts.Type = dbSingle

StDb.TableDefs("CLParameters").Fields.Append CLKilovolts

updated = True
End If

' Create StdKratios table
If versionnumber! < 11.06 Then

Set StdKratios = StDb.CreateTableDef("NewTableDef")
StdKratios.Name = "StdKratios"

With StdKratios
.Fields.Append .CreateField("StdKRatiosToNumber", dbInteger)        ' points back to Standard table/Numbers field
.Fields.Append .CreateField("StdKRatiosNumber", dbLong)             ' for k-ratio import set number (see StdKRatiosFileName field)

.Fields.Append .CreateField("StdKRatiosTakeOffs", dbSingle)
.Fields.Append .CreateField("StdKRatiosKilovolts", dbSingle)
.Fields.Append .CreateField("StdKRatiosElements", dbText, DbTextElementStringLength%)
.Fields.Append .CreateField("StdKRatiosXrays", dbText, DbTextXrayStringLength%)
.Fields.Append .CreateField("StdKRatiosKRatios", dbSingle)
.Fields.Append .CreateField("StdKRatiosStdAssigns", dbInteger)
End With

StDb.TableDefs.Append StdKratios

updated = True
End If

' Add k-ratio file name to Kratio table
If versionnumber! < 11.07 Then
StdKRatiosFileName.Name = "StdKRatiosFileName"
StdKRatiosFileName.Type = dbText
StdKRatiosFileName.Size = DbTextFilenameLengthNew%
StDb.TableDefs("StdKratios").Fields.Append StdKRatiosFileName

updated = True
End If

' Add material type to standard table
If versionnumber! < 11.89 Then
StdMaterialTypes.Name = "MaterialTypes"
StdMaterialTypes.Type = dbText
StdMaterialTypes.Size = DbTextNameLength%
StdMaterialTypes.AllowZeroLength = True
StDb.TableDefs("Standard").Fields.Append StdMaterialTypes

updated = True
End If

' Add formula parameters to standard table
If versionnumber! < 11.92 Then
StdFormulaFlags.Name = "FormulaFlags"
StdFormulaFlags.Type = dbBoolean
StDb.TableDefs("Standard").Fields.Append StdFormulaFlags

StdFormulaRatios.Name = "FormulaRatios"
StdFormulaRatios.Type = dbSingle
StDb.TableDefs("Standard").Fields.Append StdFormulaRatios

StdFormulaElements.Name = "FormulaElements"
StdFormulaElements.Type = dbText
StdFormulaElements.Size = DbTextElementStringLength%
StdFormulaElements.AllowZeroLength = True
StDb.TableDefs("Standard").Fields.Append StdFormulaElements

' Open standard table in standard database
Set StRs = StDb.OpenRecordset("Standard", dbOpenTable)

' Add default formula ratio fields to all records
Do Until StRs.EOF
StRs.Edit
StRs("FormulaRatios") = 0#      ' use this for now as default
StRs.Update
StRs.MoveNext
Loop
StRs.Close

updated = True
End If

' Add CL spectra wavelength (nanometers) field to CL Spectra table
If versionnumber! < 12.3 Then
CLSpectraWavelengths.Name = "CLSpectraNanometers"
CLSpectraWavelengths.Type = dbSingle
StDb.TableDefs("CLSpectra").Fields.Append CLSpectraWavelengths

' Open standard table in standard database and check if any CL spectra are present and delete them (sorry!)
Set StRs = StDb.OpenRecordset("CLSpectra", dbOpenTable)
If Not StRs.EOF Then
StRs.Close

SQLQ$ = "DELETE from CLSpectra WHERE CLSpectra.CLSpectraToNumber > 0"
StDb.Execute SQLQ$
SQLQ$ = "DELETE from CLParameters WHERE CLParameters.CLParametersToNumber > 0"
StDb.Execute SQLQ$

Else
StRs.Close
End If

updated = True
End If

' Add text memo table and records
If versionnumber! < 12.89 Then

' Specify the standard database table "MemoText" table
Set MemoText = StDb.CreateTableDef("NewTableDef")
MemoText.Name = "MemoText"

With MemoText
.Fields.Append .CreateField("MemoTextToNumber", dbInteger)        ' points back to Standard table/Numbers field
.Fields.Append .CreateField("MemoTextField", dbMemo, DbTextMemoStringLength&)
.Fields("MemoTextField").AllowZeroLength = True
End With

MemoTextIndex.Name = "Text Memo Numbers"
MemoTextIndex.Fields = "MemoTextToNumber"                           ' index to pointer to standard numbers
MemoTextIndex.Primary = False
MemoText.Indexes.Append MemoTextIndex
StDb.TableDefs.Append MemoText

updated = True
End If

' Add *missing* CL feild and new standard mount field
If versionnumber! < 12.98 Then

' Add "missing" CL dark intensity field only if it is missing (missing in some standard databases somehow?)
On Error Resume Next
CLSpectraIntensityDark.Name = "CLSpectraIntensityDark"
CLSpectraIntensityDark.Type = dbSingle
StDb.TableDefs("CLSpectra").Fields.Append CLSpectraIntensityDark
On Error GoTo StandardUpdateMDBFileError

' Add mount name(s) to standard table
StdMountNames.Name = "MountNames"
StdMountNames.Type = dbText
StdMountNames.Size = DbTextDescriptionLength%
StdMountNames.AllowZeroLength = True
StDb.TableDefs("Standard").Fields.Append StdMountNames

updated = True
End If

' Add new fields and records based on "versionnumber"




' Open "File" table and update data file version number
If updated Then

Set StRs = StDb.OpenRecordset("File", dbOpenTable)
StRs.Edit
StRs("Version") = ProgramVersionNumber!
StRs.Update
StRs.Close

Call IOWriteLog(tfilename$ & " was automatically updated for new database fields")
End If

StDb.Close

Call TransactionCommit("StandardUpdateMDBFile", tfilename$)
If ierror Then Exit Sub

Screen.MousePointer = vbDefault
Exit Sub

' Errors
StandardUpdateMDBFileError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardUpdateMDBFile"
Call TransactionRollback("StandardUpdateMDBFile", tfilename$)
ierror = True
Exit Sub

End Sub

