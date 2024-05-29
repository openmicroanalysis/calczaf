Attribute VB_Name = "CodeSTANDARD"
' (c) Copyright 1995-2024 by John J. Donovan
Option Explicit

Dim StandardTmpSample(1 To 1) As TypeSample

Sub StandardFindStandard(sym As String, low As Single, high As Single, tList As ListBox)
' Load passed list box with selected standards

ierror = False
On Error GoTo StandardFindStandardError

Dim ip As Integer, num As Integer
Dim temp As Single

Dim StDb As Database
Dim StRs As Recordset
Dim SQLQ As String

' Initialize the Tmp sample
Screen.MousePointer = vbHourglass

' Open the standard database
If StandardDataFile$ = vbNullString Then StandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseNonExclusiveAccess%, dbReadOnly)

' Get standard composition data for specified standard element and range from standard database
SQLQ$ = "SELECT Element.* FROM Element WHERE Element.Symbol = '" & Trim$(sym$) & "'"
SQLQ$ = SQLQ$ & " AND Element.Percent >= " & Str$(low!)
SQLQ$ = SQLQ$ & " AND Element.Percent <= " & Str$(high!)

Set StRs = StDb.OpenRecordset(SQLQ$, dbOpenDynaset)
If StRs.BOF And StRs.EOF Then GoTo StandardFindStandardNone

' Load all standards that matched the element symbol and composition range
tList.Clear
Do Until StRs.EOF
num% = StRs("Number")
ip% = StandardGetRow%(num%)
temp! = StRs("Percent")
msg$ = sym$ & " = " & Format$(Format$(temp!, f83$), a80$) & " " & StandardGetString$(ip%)
tList.AddItem msg$
tList.ItemData(tList.NewIndex) = StandardIndexNumbers%(ip%)     ' load standard number to ItemData
StRs.MoveNext
Loop

' Close the standard database
StRs.Close
StDb.Close

Screen.MousePointer = vbDefault
Exit Sub

' Errors
StandardFindStandardError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardFindStandard"
ierror = True
Exit Sub

StandardFindStandardNone:
Screen.MousePointer = vbDefault
msg$ = "No standards matching the element symbol and composition range were found"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardFindStandard"
ierror = True
Exit Sub

End Sub

Sub StandardOpenMDBFile(tfilename As String, tForm As Form)
' The routine opens an existing standard database file

ierror = False
On Error GoTo StandardOpenMDBFileError

ImportDataFile$ = vbNullString
ExportDataFile$ = vbNullString

' Get old standard filename
If tfilename$ = vbNullString Then tfilename$ = "standard.mdb"
Call IOGetMDBFileName(Int(4), tfilename$, tForm)
If ierror Then
StandardDataFile$ = vbNullString
Call StanFormUpdate
FormMAIN.Caption = "Standard (Compositional Database)"
Exit Sub
End If

' Check the database type
Call FileInfoLoadData(Int(1), tfilename$)
If ierror Then
StandardDataFile$ = vbNullString
Call StanFormUpdate
FormMAIN.Caption = "Standard (Compositional Database)"
Exit Sub
End If

' No errors, load filename
StandardDataFile$ = tfilename$
Call StanFormUpdate
FormMAIN.Caption = "Standard Database [" & StandardDataFile$ & "]"

' Load window positions
Call InitWindow(Int(2), MDBUserName$, FormMAIN)

' Check if the standard database needs to be updated
Call StandardUpdateMDBFile(StandardDataFile$)
If ierror Then Exit Sub

' Get the standard numbers and names
Call StandardGetMDBIndex
If ierror Then Exit Sub

' Load the standard list box
Call StandardLoadList(FormMAIN.ListAvailableStandards)
If ierror Then Exit Sub

Exit Sub

' Errors
StandardOpenMDBFileError:
StandardDataFile$ = vbNullString
MsgBox Error$, vbOKOnly + vbCritical, "StandardOpenMDBFile"
ierror = True
Exit Sub

End Sub

Sub StandardOpenNEWFile(tfilename As String, tForm As Form)
' The routine opens a new standard database file

ierror = False
On Error GoTo StandardOpenNewFileError

Dim StDb As Database

' Specify the standard database variables
Dim Std As New TableDef
Dim StdNumbers As New Field
Dim StdNames As New Field
Dim StdDescriptions As New Field
Dim StdDisplayAsOxideFlags As New Field
Dim StdDensities As New Field

Dim StdIndex As New Index

Dim elm As New TableDef
Dim ElmNumber As New Field
Dim ElmSymbol As New Field
Dim ElmPercent As New Field
Dim ElmNumCat As New Field
Dim ElmNumOxd As New Field

Dim ElmIndex As New Index

' New tables and fields
Dim Group As New TableDef
Dim GroupGroupNumbers As New Field    ' unique index for groups
Dim GroupGroupNames As New Field
Dim GroupNumberofPhases As New Field
Dim GroupMinimumTotals As New Field
Dim GroupDoEndMembers As New Field
Dim GroupNormalizeFlags As New Field
Dim GroupWeightFlags As New Field

Dim GroupIndex As New Index   ' index on group numbers

Dim Phase As New TableDef
Dim PhasePhaseToRow As New Field
Dim PhasePhaseOrder As New Field
Dim PhasePhaseNames As New Field
Dim PhaseNumberofStandards As New Field
Dim PhaseMinimumVectors As New Field
Dim PhaseEndMemberNumbers As New Field

Dim ModalStd As New TableDef
Dim ModalStdStdToRow As New Field        ' points to group table (group number)
Dim ModalStdPhaseOrder As New Field      ' phase load order
Dim ModalStdStdOrder As New Field        ' standard load order
' 1 to MAXPHASE%, 1 to MAXSTD%
Dim ModalStdStandardNumbers As New Field

Dim EDSSpectra As TableDef  ' EDS spectra
Dim EDSSpectraIndex As New Index  ' EDS spectra index (to sample row numbers)
Dim EDSParameters As TableDef ' EDS parameters
Dim EDSParametersIndex As New Index  ' EDS parameters index (to sample row numbers)

Dim CLSpectra As TableDef  ' CL spectra
Dim CLSpectraIndex As New Index  ' CL spectra index (to sample row numbers)
Dim CLParameters As TableDef ' CL parameters
Dim CLParametersIndex As New Index  ' CL parameters index (to sample row numbers)

Dim StdKratios As New TableDef

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

Dim ElmAtomicChargesStd As New Field
Dim ElmAtomicWtsStd As New Field

ImportDataFile$ = vbNullString
ExportDataFile$ = vbNullString

' Get new standard filename
If tfilename$ = vbNullString Then tfilename$ = "standard.mdb"
Call IOGetMDBFileName(Int(3), tfilename$, tForm)
If ierror Then Exit Sub

' Init standard arrays
Call InitStandardIndex
If ierror Then Exit Sub

' Open the new database and create the tables and index
Screen.MousePointer = vbHourglass
'Set StDb = CreateDatabase(tfilename$, dbLangGeneral)
'If StDb Is Nothing Or Err <> 0 Then GoTo StandardOpenNewFileError

' Open a new database by copying from existing MDB template
Call FileInfoCreateDatabase(tfilename$)
If ierror Then Exit Sub

' Open as existing database
Set StDb = OpenDatabase(tfilename$, DatabaseExclusiveAccess%, False)

Call TransactionBegin("StandardOpenNewFile", StandardDataFile$)
If ierror Then Exit Sub

' Specify the standard database "Standard" table
Std.Name = "Standard"
StdNumbers.Name = "Numbers"
StdNumbers.Type = dbInteger
Std.Fields.Append StdNumbers

StdNames.Name = "Names"
StdNames.Type = dbText
StdNames.Size = DbTextNameLength%
StdNames.AllowZeroLength = True
Std.Fields.Append StdNames

StdDescriptions.Name = "Descriptions"
StdDescriptions.Type = dbText
StdDescriptions.Size = DbTextDescriptionLength%
StdDescriptions.AllowZeroLength = True
Std.Fields.Append StdDescriptions

StdDisplayAsOxideFlags.Name = "DisplayAsOxideFlags"
StdDisplayAsOxideFlags.Type = dbInteger
Std.Fields.Append StdDisplayAsOxideFlags

StdDensities.Name = "Densities"
StdDensities.Type = dbSingle
Std.Fields.Append StdDensities

' Specify the standard database "StdIndex" index
StdIndex.Name = "Standard Numbers"
StdIndex.Fields = "Numbers"
StdIndex.Primary = True
Std.Indexes.Append StdIndex

StDb.TableDefs.Append Std

' Specify the standard database "Elm" table
elm.Name = "Element"
ElmNumber.Name = "Number"   ' pointer to "Standard" table
ElmNumber.Type = dbInteger
elm.Fields.Append ElmNumber

ElmSymbol.Name = "Symbol"
ElmSymbol.Type = dbText
ElmSymbol.Size = 2
ElmSymbol.AllowZeroLength = True
elm.Fields.Append ElmSymbol

ElmPercent.Name = "Percent"
ElmPercent.Type = dbSingle
elm.Fields.Append ElmPercent

ElmNumCat.Name = "NumCat"
ElmNumCat.Type = dbInteger
elm.Fields.Append ElmNumCat

ElmNumOxd.Name = "NumOxd"
ElmNumOxd.Type = dbInteger
elm.Fields.Append ElmNumOxd

ElmAtomicChargesStd.Name = "AtomicChargesStd"       ' v. 13.2.2
ElmAtomicChargesStd.Type = dbSingle
elm.Fields.Append ElmAtomicChargesStd

ElmAtomicWtsStd.Name = "AtomicWtsStd"       ' v. 13.2.2
ElmAtomicWtsStd.Type = dbSingle
elm.Fields.Append ElmAtomicWtsStd

' Specify the pointer to standard number index
ElmIndex.Name = "Element Numbers"
ElmIndex.Fields = "Number"
ElmIndex.Primary = False
elm.Indexes.Append ElmIndex

StDb.TableDefs.Append elm

Group.Name = "Groups"
GroupGroupNumbers.Name = "GroupNumbers"   ' unique row number for group index
GroupGroupNumbers.Type = dbInteger
Group.Fields.Append GroupGroupNumbers

GroupGroupNames.Name = "GroupNames"
GroupGroupNames.Type = dbText
GroupGroupNames.Size = DbTextNameLength%
GroupGroupNames.AllowZeroLength = True
Group.Fields.Append GroupGroupNames

GroupNumberofPhases.Name = "NumberofPhases"
GroupNumberofPhases.Type = dbInteger
Group.Fields.Append GroupNumberofPhases

GroupMinimumTotals.Name = "MinimumTotals"
GroupMinimumTotals.Type = dbSingle
Group.Fields.Append GroupMinimumTotals

GroupDoEndMembers.Name = "DoEndMembers"
GroupDoEndMembers.Type = dbInteger
Group.Fields.Append GroupDoEndMembers

GroupNormalizeFlags.Name = "NormalizeFlags"
GroupNormalizeFlags.Type = dbInteger
Group.Fields.Append GroupNormalizeFlags

GroupWeightFlags.Name = "WeightFlags"
GroupWeightFlags.Type = dbInteger
Group.Fields.Append GroupWeightFlags

' Specify the group row number index
GroupIndex.Name = "Group Numbers"
GroupIndex.Fields = "GroupNumbers"    ' index to group numbers field
GroupIndex.Primary = False
Group.Indexes.Append GroupIndex

StDb.TableDefs.Append Group

Phase.Name = "Phase"
PhasePhaseToRow.Name = "PhaseToRow"   ' pointer to group row
PhasePhaseToRow.Type = dbInteger
Phase.Fields.Append PhasePhaseToRow

PhasePhaseOrder.Name = "PhaseOrder"   ' load order
PhasePhaseOrder.Type = dbInteger
Phase.Fields.Append PhasePhaseOrder

PhasePhaseNames.Name = "PhaseNames"
PhasePhaseNames.Type = dbText
PhasePhaseNames.Size = DbTextNameLength%
PhasePhaseNames.AllowZeroLength = True
Phase.Fields.Append PhasePhaseNames

PhaseNumberofStandards.Name = "NumberofStandards"
PhaseNumberofStandards.Type = dbInteger
Phase.Fields.Append PhaseNumberofStandards

PhaseMinimumVectors.Name = "MinimumVectors"
PhaseMinimumVectors.Type = dbSingle
Phase.Fields.Append PhaseMinimumVectors

PhaseEndMemberNumbers.Name = "EndMemberNumbers"
PhaseEndMemberNumbers.Type = dbInteger
Phase.Fields.Append PhaseEndMemberNumbers

StDb.TableDefs.Append Phase

ModalStd.Name = "ModalStd"
ModalStdStdToRow.Name = "StdToRow"   ' pointer to group number
ModalStdStdToRow.Type = dbInteger
ModalStd.Fields.Append ModalStdStdToRow

ModalStdPhaseOrder.Name = "PhaseOrder"   ' phase load order
ModalStdPhaseOrder.Type = dbInteger
ModalStd.Fields.Append ModalStdPhaseOrder

ModalStdStdOrder.Name = "StdOrder"       ' standard load order
ModalStdStdOrder.Type = dbInteger
ModalStd.Fields.Append ModalStdStdOrder

ModalStdStandardNumbers.Name = "Numbers"    ' standard numbers
ModalStdStandardNumbers.Type = dbInteger
ModalStd.Fields.Append ModalStdStandardNumbers

StDb.TableDefs.Append ModalStd

' Specify the standard database table "EDSSpectra" EDS spectra table
Set EDSSpectra = StDb.CreateTableDef("NewTableDef")
EDSSpectra.Name = "EDSSpectra"

With EDSSpectra
.Fields.Append .CreateField("EDSSpectraToNumber", dbInteger) ' points back to Standard table/Numbers field
.Fields.Append .CreateField("EDSSpectraNumber", dbInteger) ' for multiple spectra per standard
.Fields.Append .CreateField("EDSSpectraChannelOrder", dbInteger) ' channel load order
.Fields.Append .CreateField("EDSSpectraIntensity", dbLong)         ' count data
End With

EDSSpectraIndex.Name = "EDS Spectra Numbers"
EDSSpectraIndex.Fields = "EDSSpectraToNumber" ' index to pointer to standard numbers
EDSSpectraIndex.Primary = False
EDSSpectra.Indexes.Append EDSSpectraIndex

StDb.TableDefs.Append EDSSpectra

' Create EDS parameters table for each data line
Set EDSParameters = StDb.CreateTableDef("NewTableDef")
EDSParameters.Name = "EDSParameters"

With EDSParameters
.Fields.Append .CreateField("EDSParametersToNumber", dbInteger) ' points back to Standard table/Numbers field
.Fields.Append .CreateField("EDSParametersNumber", dbInteger) ' for multiple spectral parameters per standard
.Fields.Append .CreateField("EDSParametersNumberofChannels", dbInteger)

.Fields.Append .CreateField("EDSParametersElapsedTime", dbSingle)
.Fields.Append .CreateField("EDSParametersDeadTime", dbSingle)
.Fields.Append .CreateField("EDSParametersLiveTime", dbSingle)

.Fields.Append .CreateField("EDSParametersEVPerChannel", dbSingle)
.Fields.Append .CreateField("EDSParametersStartEnergy", dbSingle)
.Fields.Append .CreateField("EDSParametersEndEnergy", dbSingle)
.Fields.Append .CreateField("EDSParametersTakeOff", dbSingle)
.Fields.Append .CreateField("EDSParametersAcceleratingVoltage", dbSingle)                   ' in keV

.Fields.Append .CreateField("EDSParametersFileName", dbText, DbTextFilenameLength%)         ' import filename
.Fields("EDSParametersFileName").AllowZeroLength = False
End With

EDSParametersIndex.Name = "EDS Parameters Numbers"
EDSParametersIndex.Fields = "EDSParametersToNumber" ' index to pointer to sample rows
EDSParametersIndex.Primary = False
EDSParameters.Indexes.Append EDSParametersIndex

StDb.TableDefs.Append EDSParameters

' Specify the probe database table "CLSpectra" CL spectra table
Set CLSpectra = StDb.CreateTableDef("NewTableDef")
CLSpectra.Name = "CLSpectra"

With CLSpectra
.Fields.Append .CreateField("CLSpectraToNumber", dbInteger) ' points back to Standard table/Numbers field
.Fields.Append .CreateField("CLSpectraNumber", dbInteger) ' for multiple CL spectra per standard
.Fields.Append .CreateField("CLSpectraChannelOrder", dbInteger) ' CL spectra channel load order
.Fields.Append .CreateField("CLSpectraIntensity", dbLong)         ' CL intensity count data
End With

CLSpectraIndex.Name = "CL Spectra Numbers"
CLSpectraIndex.Fields = "CLSpectraToNumber" ' index to pointer to standard number
CLSpectraIndex.Primary = False
CLSpectra.Indexes.Append CLSpectraIndex

StDb.TableDefs.Append CLSpectra

' Create CL parameters table for each data line
Set CLParameters = StDb.CreateTableDef("NewTableDef")
CLParameters.Name = "CLParameters"

With CLParameters
.Fields.Append .CreateField("CLParametersToNumber", dbInteger) ' points back to Sample table/RowOrder field
.Fields.Append .CreateField("CLParametersNumber", dbInteger) ' for multiple CL spectral parameters per standard
.Fields.Append .CreateField("CLParametersNumberofChannels", dbInteger)

.Fields.Append .CreateField("CLParametersCountTime", dbSingle)
.Fields.Append .CreateField("CLParametersStartEnergy", dbSingle)
.Fields.Append .CreateField("CLParametersEndEnergy", dbSingle)
.Fields.Append .CreateField("CLParametersKilovolts", dbSingle)

.Fields.Append .CreateField("CLParametersFileName", dbText, DbTextFilenameLength%)         ' import filename
.Fields("CLParametersFileName").AllowZeroLength = False
End With

CLParametersIndex.Name = "CL Parameters Numbers"
CLParametersIndex.Fields = "CLParametersToNumber" ' index to pointer to standard number
CLParametersIndex.Primary = False
CLParameters.Indexes.Append CLParametersIndex

StDb.TableDefs.Append CLParameters

' Create StdKratios table (new code 06/15/2017)
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

.Fields.Append .CreateField("StdKRatiosFileName", dbText, DbTextFilenameLengthNew%)
End With

StDb.TableDefs.Append StdKratios

' Add material type to standard table (new code 06/15/2017)
StdMaterialTypes.Name = "MaterialTypes"
StdMaterialTypes.Type = dbText
StdMaterialTypes.Size = DbTextNameLength%
StdMaterialTypes.AllowZeroLength = True
StDb.TableDefs("Standard").Fields.Append StdMaterialTypes

' Add formula parameters to standard table (new code 06/17/2017)
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

' Add CL spectra wavelength (nanometers and dark intensities) field to CL Spectra table
CLSpectraWavelengths.Name = "CLSpectraNanometers"
CLSpectraWavelengths.Type = dbSingle
StDb.TableDefs("CLSpectra").Fields.Append CLSpectraWavelengths

CLSpectraIntensityDark.Name = "CLSpectraIntensityDark"
CLSpectraIntensityDark.Type = dbSingle
StDb.TableDefs("CLSpectra").Fields.Append CLSpectraIntensityDark

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

' Add mount name(s) to standard table
StdMountNames.Name = "MountNames"
StdMountNames.Type = dbText
StdMountNames.Size = DbTextDescriptionLength%
StdMountNames.AllowZeroLength = True
StDb.TableDefs("Standard").Fields.Append StdMountNames

Call TransactionCommit("StandardOpenNewFile", StandardDataFile$)
If ierror Then Exit Sub

' Close the standard database
StDb.Close
Screen.MousePointer = vbDefault

' Create new File table for standard database
Call FileInfoMakeNewTable(Int(1), tfilename$)
If ierror Then Exit Sub

' No errors, load filename
StandardDataFile$ = tfilename$

' Set MAIN window title
FormMAIN.Caption = "Standard Database [" & StandardDataFile$ & "]"
Exit Sub

' Errors
StandardOpenNewFileError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardOpenNewFile"
Call TransactionRollback("StandardOpenNewFile", StandardDataFile$)
ierror = True
Exit Sub

End Sub

Sub StandardReplaceRecord(sample() As TypeSample)
' Routine to modify a standard in the standard database

ierror = False
On Error GoTo StandardReplaceRecordError

Dim response As Integer
Dim StDb As Database
Dim StDt As Recordset
Dim SQLQ As String

' Open the database and the "Standard" table
If StandardDataFile$ = vbNullString Then StandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseExclusiveAccess%, False)
Set StDt = StDb.OpenRecordset("Standard", dbOpenTable)

' Check that standard number already exists
StDt.Index = "Standard Numbers"
StDt.Seek "=", sample(1).number%
If StDt.NoMatch Then GoTo StandardReplaceRecordNotFound

' Confirm modify with user
msg$ = "Are you sure that you want to replace/modify standard :"
msg$ = msg$ & vbCrLf & StDt("Numbers") & " " & StDt("Names")
msg$ = msg$ & vbCrLf & " from database :"
msg$ = msg$ & vbCrLf & StandardDataFile$ & "?"
response% = MsgBox(msg$, vbYesNoCancel + vbQuestion + vbDefaultButton2, "StandardReplaceRecord")

' User selects "no", just close database and exit without error
If response% = vbNo Then
StDt.Close
StDb.Close
Exit Sub
End If

' User selects "cancel", close database and exit with error
If response% = vbCancel Then
StDt.Close
StDb.Close
ierror = True
Exit Sub
End If

' First delete sample record
Screen.MousePointer = vbHourglass
StDt.Delete
StDt.MoveFirst
StDt.Close

Call TransactionBegin("StandardReplaceRecord", StandardDataFile$)
If ierror Then Exit Sub

' Delete element symbols and weights to "Element" table based on "number"
SQLQ$ = "DELETE from Element WHERE Element.Number = " & Str$(sample(1).number%)
StDb.Execute SQLQ$

Call TransactionCommit("StandardReplaceRecord", StandardDataFile$)
If ierror Then Exit Sub

StDb.Close
Screen.MousePointer = vbDefault

' Now simply add the standard as a new record to the database
Call StandardAddRecord(sample())
If ierror Then Exit Sub

Exit Sub

' Errors
StandardReplaceRecordError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardReplaceRecord"
Call TransactionRollback("StandardReplaceRecord", StandardDataFile$)
ierror = True
Exit Sub

StandardReplaceRecordNotFound:
Screen.MousePointer = vbDefault
msg$ = "Standard number " & Str$(sample(1).number%) & " is not found in " & StandardDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardReplaceRecord"
ierror = True
Exit Sub

End Sub

Function StandardGetNumberofStandards(mode As Integer) As Integer
' Return the number of standards indicated
' mode = 0 all standards
' mode = 1 elemental standards
' mode = 2 oxide standards

ierror = False
On Error GoTo StandardGetNumberofStandardsError

Dim StDb As Database
Dim StRs As Recordset
Dim SQLQ As String

' Initialize the Tmp sample
Screen.MousePointer = vbHourglass

' Open the standard database
If StandardDataFile$ = vbNullString Then StandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseNonExclusiveAccess%, dbReadOnly)

' Get standard composition data for specified standard element and range from standard database
If mode% = 0 Then SQLQ$ = "SELECT Standard.* FROM Standard"
If mode% = 1 Then SQLQ$ = "SELECT Standard.* FROM Standard WHERE Standard.DisplayAsOxideFlags = 0"
If mode% = 2 Then SQLQ$ = "SELECT Standard.* FROM Standard WHERE Standard.DisplayAsOxideFlags = -1"
Set StRs = StDb.OpenRecordset(SQLQ$, dbOpenDynaset, dbReadOnly)

StRs.MoveLast
StandardGetNumberofStandards% = StRs.RecordCount

' Close the standard database
StRs.Close
StDb.Close

Screen.MousePointer = vbDefault
Exit Function

' Errors
StandardGetNumberofStandardsError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardGetNumberofStandards"
ierror = True
Exit Function

End Function

Function StandardIsStandardOxide(stdnum As Integer) As Boolean
' Return true if the standard is an oxide standard

ierror = False
On Error GoTo StandardIsStandardOxideError

Dim StDb As Database
Dim StRs As Recordset
Dim SQLQ As String

' Initialize the Tmp sample
Screen.MousePointer = vbHourglass

' Open the standard database
If StandardDataFile$ = vbNullString Then StandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseNonExclusiveAccess%, dbReadOnly)

' See if standard is oxide
SQLQ$ = "SELECT Standard.* FROM Standard WHERE Standard.Numbers = " & Str$(stdnum%)
Set StRs = StDb.OpenRecordset(SQLQ$, dbOpenDynaset, dbReadOnly)

If StRs.BOF And StRs.EOF Then GoTo StandardIsStandardOxideNotFound

' Load return flag
If StRs("DisplayAsOxideFlags") = 0 Then
StandardIsStandardOxide = False
Else
StandardIsStandardOxide = True
End If

' Close the standard database
StRs.Close
StDb.Close

Screen.MousePointer = vbDefault
Exit Function

' Errors
StandardIsStandardOxideError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardIsStandardOxide"
ierror = True
Exit Function

StandardIsStandardOxideNotFound:
Screen.MousePointer = vbDefault
msg$ = "Standard number " & Str$(stdnum%) & " is not found in " & StandardDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardIsStandardOxide"
ierror = True
Exit Function

End Function

Sub StandardMatchLoad()
' Load the FormMATCH form to match the selected standard composition

ierror = False
On Error GoTo StandardMatchLoadError

Dim number As Integer

' Get standard from listbox
If FormMAIN.ListAvailableStandards.ListIndex < 0 Then Exit Sub
number% = FormMAIN.ListAvailableStandards.ItemData(FormMAIN.ListAvailableStandards.ListIndex)

' Get standard from database
Call StandardGetMDBStandard(number%, StandardTmpSample())
If ierror Then Exit Sub

' Load form
DefaultMatchStandardDatabase$ = MiscGetFileNameOnly$(StandardDataFile$)
Call MatchLoad(StandardTmpSample())
If ierror Then Exit Sub

Exit Sub

' Errors
StandardMatchLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardMatchLoad"
ierror = True
Exit Sub

End Sub

Sub StandardCLSpectraGetData(stdnum As Integer, specnum As Integer, tfilename As String, sample() As TypeSample)
' Get the CL spectrum data for the specified sample (all datarows)
' stdnum = standard number (unique)
' specnum = spectrum number (for multiple spectra per standard)
' tfilename = import filename

ierror = False
On Error GoTo StandardCLSpectraGetDataError

Dim channel As Integer
Dim SQLQ As String

Dim StDb As Database
Dim StRs As Recordset

' Open the standard database and get spectrum
Screen.MousePointer = vbHourglass
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseNonExclusiveAccess%, dbReadOnly)

SQLQ$ = "SELECT CLSpectra.* FROM CLSpectra WHERE CLSpectra.CLSpectraToNumber = " & Str$(stdnum%) & " AND CLSpectra.CLSpectraNumber = " & Str$(specnum%)
Set StRs = StDb.OpenRecordset(SQLQ$, dbOpenSnapshot)

' Use "CLSpectraChannelOrder" field to load channels in order
Do Until StRs.EOF
channel% = StRs("CLSpectraChannelOrder")  ' 1 to MAXSPECTRA_CL
If channel% < 1 Or channel% > MAXSPECTRA_CL% Then GoTo StandardCLSpectraGetDataBadSpectra
sample(1).CLSpectraIntensities&(1, channel%) = StRs("CLSpectraIntensity")
sample(1).CLSpectraNanometers!(1, channel%) = StRs("CLSpectraNanometers")
sample(1).CLSpectraDarkIntensities&(1, channel%) = StRs("CLSpectraIntensityDark")
StRs.MoveNext
Loop

StRs.Close

' Next get the CL spectrum parameters
SQLQ$ = "SELECT CLParameters.* FROM CLParameters WHERE CLParameters.CLParametersToNumber = " & Str$(stdnum%) & " AND CLParameters.CLParametersNumber = " & Str$(specnum%)
Set StRs = StDb.OpenRecordset(SQLQ$, dbOpenSnapshot)

' Load parameters for the specified spectrum
Do Until StRs.EOF
sample(1).CLSpectraNumberofChannels%(1) = StRs("CLParametersNumberofChannels")

sample(1).CLAcquisitionCountTime!(1) = StRs("CLParametersCountTime")          ' live time (actual count integration time)
sample(1).CLSpectraStartEnergy!(1) = StRs("CLParametersStartEnergy")
sample(1).CLSpectraEndEnergy!(1) = StRs("CLParametersEndEnergy")

tfilename$ = Trim$(vbNullString & StRs("CLParametersFileName"))
sample(1).CLSpectraKilovolts!(1) = StRs("CLParametersKilovolts")
StRs.MoveNext
Loop

StRs.Close

' Close the standard database
Screen.MousePointer = vbDefault
StDb.Close

' Make sure objects are deallocated
If Not StRs Is Nothing Then Set StRs = Nothing
If Not StDb Is Nothing Then Set StDb = Nothing

Exit Sub

' Errors
StandardCLSpectraGetDataError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardCLSpectraGetData"
ierror = True
Exit Sub

StandardCLSpectraGetDataBadSpectra:
msg$ = "CL spectra channel row is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardCLSpectraGetData"
ierror = True
Exit Sub

End Sub

Sub StandardCLSpectraSendData(stdnum As Integer, specnum As Integer, tfilename As String, sample() As TypeSample)
' Write the CL spectrum data to the specified sample for the specified line
' stdnum = standard number (unique)
' specnum = spectrum number (for multiple spectra per standard)

ierror = False
On Error GoTo StandardCLSpectraSendDataError

Dim i As Integer

Dim StDb As Database
Dim StRs As Recordset

' Check limits
If stdnum% < 1 Or stdnum% > MAXINTEGER% Then GoTo StandardCLSpectraSendDataBadStandard
If specnum% < 1 Or specnum% > MAXINTEGER% Then GoTo StandardCLSpectraSendDataBadNumber

' Open the Standard database and write new data to CL table
Screen.MousePointer = vbHourglass
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseExclusiveAccess%, False)

' Open "CLSpectra" table
Set StRs = StDb.OpenRecordset("CLSpectra", dbOpenTable)

Call TransactionBegin("StandardCLSpectraSendData", StandardDataFile$)
If ierror Then Exit Sub

' Save into arrays
For i% = 1 To sample(1).CLSpectraNumberofChannels%(1)
StRs.AddNew
StRs("CLSpectraToNumber") = stdnum%
StRs("CLSpectraNumber") = specnum%              ' for more than one CL spectra per standard
StRs("CLSpectraChannelOrder") = i%
StRs("CLSpectraIntensity") = sample(1).CLSpectraIntensities&(1, i%)
StRs("CLSpectraNanometers") = sample(1).CLSpectraNanometers!(1, i%)                 ' new array because wavelengths are non-linear (v. 12.3)
StRs("CLSpectraIntensityDark") = sample(1).CLSpectraDarkIntensities&(1, i%)         ' save dark spectra to same table as light intensities
StRs.Update
Next i%

Call TransactionCommit("StandardCLSpectraSendData", StandardDataFile$)
If ierror Then Exit Sub

StRs.Close

' Open "CLParameters" table
Set StRs = StDb.OpenRecordset("CLParameters", dbOpenTable)
Call TransactionBegin("StandardCLSpectraSendData", StandardDataFile$)
If ierror Then Exit Sub

' Save CL spectrum parameters
StRs.AddNew
StRs("CLParametersToNumber") = stdnum%
StRs("CLParametersNumber") = specnum%
StRs("CLParametersNumberofChannels") = sample(1).CLSpectraNumberofChannels%(1)

StRs("CLParametersCountTime") = sample(1).CLAcquisitionCountTime!(1)          ' actual count integration time
StRs("CLParametersStartEnergy") = sample(1).CLSpectraStartEnergy!(1)
StRs("CLParametersEndEnergy") = sample(1).CLSpectraEndEnergy!(1)

StRs("CLParametersKilovolts") = sample(1).CLSpectraKilovolts!(1)
StRs("CLParametersFileName") = Trim$(Left$(tfilename$, DbTextFilenameLength%))
StRs.Update

Call TransactionCommit("StandardCLSpectraSendData", StandardDataFile$)
If ierror Then Exit Sub

StRs.Close

' Close the standard database
Screen.MousePointer = vbDefault
StDb.Close

' Make sure objects are deallocated
If Not StRs Is Nothing Then Set StRs = Nothing
If Not StDb Is Nothing Then Set StDb = Nothing
Exit Sub

' Errors
StandardCLSpectraSendDataError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardCLSpectraSendData"
Call TransactionRollback("StandardCLSpectraSendData", StandardDataFile$)
ierror = True
Exit Sub

StandardCLSpectraSendDataBadStandard:
msg$ = "CL spectra standard number is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardCLSpectraSendData"
ierror = True
Exit Sub

StandardCLSpectraSendDataBadNumber:
msg$ = "CL spectra number is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardCLSpectraSendData"
ierror = True
Exit Sub

End Sub

Sub StandardEDSSpectraGetData(stdnum As Integer, specnum As Integer, tfilename As String, sample() As TypeSample)
' Get the EDS spectrum data for the specified sample (all datarows)
' stdnum = standard number (unique)
' specnum = spectrum number (for multiple spectra per standard)

ierror = False
On Error GoTo StandardEDSSpectraGetDataError

Dim channel As Integer
Dim SQLQ As String

Dim PrDb As Database
Dim PrRs As Recordset

' Open the Standard database and read spectra table
Screen.MousePointer = vbHourglass
Set PrDb = OpenDatabase(StandardDataFile$, StandardDatabaseNonExclusiveAccess%, dbReadOnly)

SQLQ$ = "SELECT EDSSpectra.* FROM EDSSpectra WHERE EDSSpectra.EDSSpectraToNumber = " & Str$(stdnum%) & " AND EDSSpectra.EDSSpectraNumber = " & Str$(specnum%)
Set PrRs = PrDb.OpenRecordset(SQLQ$, dbOpenSnapshot)

' Use "EDSSpectraChannelOrder" field to load channels
Do Until PrRs.EOF
channel% = PrRs("EDSSpectraChannelOrder")  ' 1 to MAXSPECTRA
If channel% < 1 Or channel% > MAXSPECTRA% Then GoTo StandardEDSSpectraGetDataBadSpectra
sample(1).EDSSpectraIntensities&(1, channel%) = PrRs("EDSSpectraIntensity")
PrRs.MoveNext
Loop

PrRs.Close

' Next get the EDS spectrum parameters
SQLQ$ = "SELECT EDSParameters.* FROM EDSParameters WHERE EDSParameters.EDSParametersToNumber = " & Str$(stdnum%) & " AND EDSParameters.EDSParametersNumber = " & Str$(specnum%)
Set PrRs = PrDb.OpenRecordset(SQLQ$, dbOpenSnapshot)

' Load parameters for each data line
Do Until PrRs.EOF
sample(1).EDSSpectraNumberofChannels%(1) = PrRs("EDSParametersNumberofChannels")

sample(1).EDSSpectraElapsedTime!(1) = PrRs("EDSParametersElapsedTime")    ' elapsed (real) time
sample(1).EDSSpectraDeadTime!(1) = PrRs("EDSParametersDeadTime")          ' deadtime (in percentage)
sample(1).EDSSpectraLiveTime!(1) = PrRs("EDSParametersLiveTime")          ' live time (actual count integration time)

sample(1).EDSSpectraEVPerChannel!(1) = PrRs("EDSParametersEVPerChannel")
sample(1).EDSSpectraStartEnergy!(1) = PrRs("EDSParametersStartEnergy")
sample(1).EDSSpectraEndEnergy!(1) = PrRs("EDSParametersEndEnergy")
sample(1).EDSSpectraTakeOff!(1) = PrRs("EDSParametersTakeOff")
sample(1).EDSSpectraAcceleratingVoltage!(1) = PrRs("EDSParametersAcceleratingVoltage")      ' keV

tfilename$ = Trim$(vbNullString & PrRs("EDSParametersFileName"))      ' import filename
PrRs.MoveNext
Loop

PrRs.Close

' Close the Standard database
Screen.MousePointer = vbDefault
PrDb.Close

' Make sure objects are deallocated
If Not PrRs Is Nothing Then Set PrRs = Nothing
If Not PrDb Is Nothing Then Set PrDb = Nothing

Exit Sub

' Errors
StandardEDSSpectraGetDataError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardEDSSpectraGetData"
ierror = True
Exit Sub

StandardEDSSpectraGetDataBadSpectra:
msg$ = "EDS spectra channel row is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardEDSSpectraGetData"
ierror = True
Exit Sub

StandardEDSSpectraGetDataBadNumber:
msg$ = "EDS spectra number is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardEDSSpectraGetData"
ierror = True
Exit Sub

StandardEDSSpectraGetDataBadStrobe:
msg$ = "EDS strobe data channel is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardEDSSpectraGetData"
ierror = True
Exit Sub

End Sub

Sub StandardEDSSpectraSendData(stdnum As Integer, specnum As Integer, tfilename As String, sample() As TypeSample)
' Write the EDS spectrum data to the specified sample for the specified line
' stdnum = standard number (unique)
' specnum = spectrum number (for multiple spectra per standard)
' tfilename = import filename

ierror = False
On Error GoTo StandardEDSSpectraSendDataError

Dim i As Integer

Dim PrDb As Database
Dim PrRs As Recordset

' Check limits
If stdnum% < 1 Or stdnum% > MAXINTEGER% Then GoTo StandardEDSSpectraSendDataBadStandard
If specnum% < 1 Or specnum% > MAXINTEGER% Then GoTo StandardEDSSpectraSendDataBadNumber

' Open the Standard database and write new data to spectra table
Screen.MousePointer = vbHourglass
Set PrDb = OpenDatabase(StandardDataFile$, StandardDatabaseExclusiveAccess%, False)

' Open "EDSSpectra" table
Set PrRs = PrDb.OpenRecordset("EDSSpectra", dbOpenTable)

Call TransactionBegin("StandardEDSSpectraSendData", StandardDataFile$)
If ierror Then Exit Sub

' Save into arrays
For i% = 1 To sample(1).EDSSpectraNumberofChannels%(1)
PrRs.AddNew
PrRs("EDSSpectraToNumber") = stdnum%
PrRs("EDSSpectraNumber") = specnum%             ' for more than one EDS spectra per standard
PrRs("EDSSpectraChannelOrder") = i%
PrRs("EDSSpectraIntensity") = sample(1).EDSSpectraIntensities&(1, i%)
PrRs.Update
Next i%

Call TransactionCommit("StandardEDSSpectraSendData", StandardDataFile$)
If ierror Then Exit Sub

PrRs.Close

' Open "EDSParameters" table
Set PrRs = PrDb.OpenRecordset("EDSParameters", dbOpenTable)

Call TransactionBegin("StandardEDSSpectraSendData", StandardDataFile$)
If ierror Then Exit Sub

' Save spectrum parameters for this data line
PrRs.AddNew
PrRs("EDSParametersToNumber") = stdnum%
PrRs("EDSParametersNumber") = specnum%
PrRs("EDSParametersNumberofChannels") = sample(1).EDSSpectraNumberofChannels%(1)

PrRs("EDSParametersElapsedTime") = sample(1).EDSSpectraElapsedTime!(1)    ' elapsed time (real time)
PrRs("EDSParametersDeadTime") = sample(1).EDSSpectraDeadTime!(1)          ' deadtime in percent
PrRs("EDSParametersLiveTime") = sample(1).EDSSpectraLiveTime!(1)          ' actual count integration time

PrRs("EDSParametersEVPerChannel") = sample(1).EDSSpectraEVPerChannel!(1)
PrRs("EDSParametersStartEnergy") = sample(1).EDSSpectraStartEnergy!(1)
PrRs("EDSParametersEndEnergy") = sample(1).EDSSpectraEndEnergy!(1)
PrRs("EDSParametersTakeOff") = sample(1).EDSSpectraTakeOff!(1)
PrRs("EDSParametersAcceleratingVoltage") = sample(1).EDSSpectraAcceleratingVoltage!(1)      ' keV

PrRs("EDSParametersFileName") = Trim$(Left$(tfilename$, DbTextFilenameLength%))
PrRs.Update

Call TransactionCommit("StandardEDSSpectraSendData", StandardDataFile$)
If ierror Then Exit Sub

PrRs.Close

' Close the Standard database
Screen.MousePointer = vbDefault
PrDb.Close

' Make sure objects are deallocated
If Not PrRs Is Nothing Then Set PrRs = Nothing
If Not PrDb Is Nothing Then Set PrDb = Nothing
Exit Sub

' Errors
StandardEDSSpectraSendDataError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardEDSSpectraSendData"
Call TransactionRollback("StandardEDSSpectraSendData", StandardDataFile$)
ierror = True
Exit Sub

StandardEDSSpectraSendDataBadStandard:
msg$ = "EDS spectra sample row is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardEDSSpectraSendData"
ierror = True
Exit Sub

StandardEDSSpectraSendDataBadNumber:
msg$ = "EDS spectra number is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardEDSSpectraSendData"
ierror = True
Exit Sub

StandardEDSSpectraSendDataBadRow:
msg$ = "EDS spectra data row is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardEDSSpectraSendData"
ierror = True
Exit Sub

End Sub

Sub StandardImportEDSSpectrum(stdnum As Integer, tForm As Form, sample() As TypeSample)
' Import an EDS EMSA spectrum to the current standard

ierror = False
On Error GoTo StandardImportEDSSpectrumError

Dim specnum As Integer, chan As Integer
Dim response As Integer

Static tfilename As String

' Get new filename
Call IOGetFileName(Int(2), "EMSA", tfilename$, tForm)
If ierror Then Exit Sub

' Init the sample (for EDS and CL arrays)
Call InitSample(StandardTmpSample())
If ierror Then Exit Sub

ReDim StandardTmpSample(1).EDSSpectraIntensities(1 To MAXROW%, 1 To MAXSPECTRA%) As Long
ReDim StandardTmpSample(1).EDSSpectraStrobes(1 To MAXROW%, 1 To MAXSTROBE%) As Long    ' only for Oxford EDS

' Read the EMSA file (check that it is an EDS spectrum)
Call EMSAReadSpectrum(Int(0), Int(1), StandardTmpSample(), tfilename$)
If ierror Then Exit Sub

' Show the user the EDS spectrum and allow user to confirm
Call EDSInitDisplay(FormEDSDISPLAY3, tfilename$, Int(1), StandardTmpSample())
If ierror Then Exit Sub

' Load elements from standard compositional sample to EDS spectrum sample
For chan% = 1 To sample(1).LastChan%
StandardTmpSample(1).Elsyms$(chan%) = sample(1).Elsyms$(chan%)
Next chan%
StandardTmpSample(1).LastChan% = sample(1).LastChan%

' Display the EDS spectra in cps
Call EDSDisplaySpectra(FormEDSDISPLAY3, Int(1), StandardTmpSample())
If ierror Then Exit Sub

FormEDSDISPLAY3.Show vbModal

msg$ = "Do you want to store this EDS spectrum with the current standard composition?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton1, "StandardImportEDSSpectrum")
If response% = vbNo Then
Unload FormEDSDISPLAY3
ierror = True
Exit Sub
End If

' Get the total number of EDS spectra already stored for this standard
Call StandardImportGetTotalSpectra(Int(0), stdnum%, specnum%)
If ierror Then Exit Sub

' Save the new EDS spectrum to the standard database
Call StandardEDSSpectraSendData(stdnum%, specnum% + 1, tfilename$, StandardTmpSample())
If ierror Then Exit Sub

Unload FormEDSDISPLAY3
Exit Sub

' Errors
StandardImportEDSSpectrumError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardImportEDSSpectrum"
ierror = True
Exit Sub

End Sub

Sub StandardImportCLSpectrum(stdnum As Integer, tForm As Form, sample() As TypeSample)
' Import a CL spectrum to the current standard

ierror = False
On Error GoTo StandardImportCLSpectrumError

Dim specnum As Integer
Dim tfilename As String
Dim response As Integer

' Get new filename
Call IOGetFileName(Int(2), "EMSA", tfilename$, tForm)
If ierror Then Exit Sub

' Init the sample (for CL and CL arrays)
Call InitSample(StandardTmpSample())
If ierror Then Exit Sub

ReDim StandardTmpSample(1).CLSpectraIntensities(1 To MAXROW%, 1 To MAXSPECTRA_CL%) As Long
ReDim StandardTmpSample(1).CLSpectraDarkIntensities(1 To MAXROW%, 1 To MAXSPECTRA_CL%) As Long
ReDim StandardTmpSample(1).CLSpectraNanometers(1 To MAXROW%, 1 To MAXSPECTRA_CL%) As Single

' Read the EMSA file (check that it is a CL spectrum)
Call EMSAReadSpectrum(Int(1), Int(1), StandardTmpSample(), tfilename$)
If ierror Then Exit Sub

' Show the user the CL spectrum and allow user to confirm
Call CLInitDisplay(FormCLDISPLAY, tfilename$, Int(1), StandardTmpSample())
If ierror Then Exit Sub

' Display the CL spectra in cps (tCLIntensityOption% = 1)
Call CLDisplaySpectra(False, FormCLDISPLAY, Int(1), StandardTmpSample())
If ierror Then Exit Sub

FormCLDISPLAY.Show vbModal

msg$ = "Do you want to store this CL spectrum with the current standard composition?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton1, "StandardImportCLSpectrum")
If response% = vbNo Then
Unload FormCLDISPLAY
ierror = True
Exit Sub
End If

' Get the total number of CL spectra already stored for this standard
Call StandardImportGetTotalSpectra(Int(1), stdnum%, specnum%)
If ierror Then Exit Sub

' Save the new CL spectrum to the standard database
Call StandardCLSpectraSendData(stdnum%, specnum% + 1, tfilename$, StandardTmpSample())
If ierror Then Exit Sub

Unload FormCLDISPLAY

Exit Sub

' Errors
StandardImportCLSpectrumError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardImportCLSpectrum"
ierror = True
Exit Sub

End Sub

Sub StandardImportGetTotalSpectra(mode As Integer, stdnum As Integer, specnum As Integer)
' Determine the toal number of spectra already stored for this standard
'  mode = 0 EDS spectra
'  mode = 1 CL spectra

ierror = False
On Error GoTo StandardImportGetTotalSpectraError

Dim StDb As Database
Dim StRs As Recordset
Dim SQLQ As String

' Assume no spectra
specnum% = 0

' Open standard database
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseNonExclusiveAccess%, dbReadOnly)

' Get number of EDS spectra for this standard
If mode% = 0 Then
SQLQ$ = "SELECT EDSParameters.* FROM EDSParameters WHERE EDSParameters.EDSParametersToNumber = " & Str$(stdnum%)
Set StRs = StDb.OpenRecordset(SQLQ$, dbOpenSnapshot)
If StRs.BOF And StRs.EOF Then Exit Sub
StRs.MoveLast
specnum% = StRs.RecordCount
End If

' Get number of CL spectra for this standard
If mode% = 1 Then
SQLQ$ = "SELECT CLparameters.* FROM CLParameters WHERE CLParameters.CLParametersToNumber = " & Str$(stdnum%)
Set StRs = StDb.OpenRecordset(SQLQ$, dbOpenSnapshot)
If StRs.BOF And StRs.EOF Then Exit Sub
StRs.MoveLast
specnum% = StRs.RecordCount
End If

StRs.Close
StDb.Close

Exit Sub

' Errors
StandardImportGetTotalSpectraError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardImportGetTotalSpectra"
ierror = True
Exit Sub

End Sub

Sub StandardLoadSpectra(mode As Integer, stdnum As Integer, tList As ListBox)
' Load the specified spectra type to the passed listbox

ierror = False
On Error GoTo StandardLoadSpectraError

Dim SQLQ As String, astring As String, tfilename As String

Dim StDb As Database
Dim StRs As Recordset

' Open the Standard database and read spectra table
Screen.MousePointer = vbHourglass
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseNonExclusiveAccess%, dbReadOnly)

' Get the EDS spectrum parameters
If mode% = 0 Then
SQLQ$ = "SELECT EDSParameters.* FROM EDSParameters WHERE EDSParameters.EDSParametersToNumber = " & Str$(stdnum%) & " AND EDSParameters.EDSParametersNumber <> 0"
Set StRs = StDb.OpenRecordset(SQLQ$, dbOpenSnapshot)

' Load parameters for each data line
tList.Clear
Do Until StRs.EOF
tfilename$ = MiscGetFileNameOnly$(Trim$(vbNullString & StRs("EDSParametersFileName")))
astring$ = StRs("EDSParametersAcceleratingVoltage") & " keV, " & tfilename$
tList.AddItem astring$
tList.ItemData(tList.NewIndex) = StRs("EDSParametersNumber")
StRs.MoveNext
Loop

StRs.Close
End If

' Get the CL spectrum parameters
If mode% = 1 Then
SQLQ$ = "SELECT CLParameters.* FROM CLParameters WHERE CLParameters.CLParametersToNumber = " & Str$(stdnum%) & " AND CLParameters.CLParametersNumber <> 0"
Set StRs = StDb.OpenRecordset(SQLQ$, dbOpenSnapshot)

' Load parameters for each data line
tList.Clear
Do Until StRs.EOF
tfilename$ = MiscGetFileNameOnly$(Trim$(vbNullString & StRs("CLParametersFileName")))
astring$ = StRs("CLParametersKilovolts") & " keV, " & tfilename$
tList.AddItem astring$
tList.ItemData(tList.NewIndex) = StRs("CLParametersNumber")
StRs.MoveNext
Loop

StRs.Close
End If

' Close the Standard database
Screen.MousePointer = vbDefault
StDb.Close

' Make sure objects are deallocated
If Not StRs Is Nothing Then Set StRs = Nothing
If Not StDb Is Nothing Then Set StDb = Nothing

Exit Sub

' Errors
StandardLoadSpectraError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardLoadSpectra"
Call TransactionRollback("StandardLoadSpectra", StandardDataFile$)
ierror = True
Exit Sub

StandardLoadSpectraBadSpectra:
msg$ = "EDS spectra channel row is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardLoadSpectra"
ierror = True
Exit Sub

StandardLoadSpectraBadNumber:
msg$ = "EDS spectra number is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardLoadSpectra"
ierror = True
Exit Sub

StandardLoadSpectraBadStrobe:
msg$ = "EDS strobe data channel is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardLoadSpectra"
ierror = True
Exit Sub

End Sub

Sub StandardDeleteEDSSpectrum(stdnum As Integer, specnum As Integer)
' Delete the specified EDS spectrum

ierror = False
On Error GoTo StandardDeleteEDSSpectrumError

Dim StDb As Database
Dim SQLQ As String

' Open the database and the "Standard" table
If StandardDataFile$ = vbNullString Then StandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseExclusiveAccess%, False)

Call TransactionBegin("StandardDeleteEDSSpectrum", StandardDataFile$)
If ierror Then Exit Sub

' Delete EDSSpectra table based on "stdnum" and "specnum"
SQLQ$ = "DELETE from EDSSpectra WHERE EDSSpectra.EDSSpectraToNumber = " & Str$(stdnum%) & " AND EDSSpectra.EDSSpectraNumber = " & Str$(specnum%)
StDb.Execute SQLQ$

' Delete EDSParameters table based on "stdnum" and "specnum"
SQLQ$ = "DELETE from EDSParameters WHERE EDSParameters.EDSParametersToNumber = " & Str$(stdnum%) & " AND EDSParameters.EDSParametersNumber = " & Str$(specnum%)
StDb.Execute SQLQ$

Call TransactionCommit("StandardDeleteEDSSpectrum", StandardDataFile$)
If ierror Then Exit Sub

StDb.Close
Screen.MousePointer = vbDefault

Exit Sub

' Errors
StandardDeleteEDSSpectrumError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardDeleteEDSSpectrum"
ierror = True
Exit Sub

End Sub

Sub StandardDeleteCLSpectrum(stdnum As Integer, specnum As Integer)
' Delete the specified CL spectrum

ierror = False
On Error GoTo StandardDeleteCLSpectrumError

Dim StDb As Database
Dim SQLQ As String

' Open the database and the "Standard" table
If StandardDataFile$ = vbNullString Then StandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseExclusiveAccess%, False)

Call TransactionBegin("StandardDeleteCLSpectrum", StandardDataFile$)
If ierror Then Exit Sub

' Delete CLSpectra table based on "stdnum" and "specnum"
SQLQ$ = "DELETE from CLSpectra WHERE CLSpectra.CLSpectraToNumber = " & Str$(stdnum%) & " AND CLSpectra.CLSpectraNumber = " & Str$(specnum%)
StDb.Execute SQLQ$

' Delete CLParameters table based on "stdnum" and "specnum"
SQLQ$ = "DELETE from CLParameters WHERE CLParameters.CLParametersToNumber = " & Str$(stdnum%) & " AND CLParameters.CLParametersNumber = " & Str$(specnum%)
StDb.Execute SQLQ$

Call TransactionCommit("StandardDeleteCLSpectrum", StandardDataFile$)
If ierror Then Exit Sub

StDb.Close
Screen.MousePointer = vbDefault

Exit Sub

' Errors
StandardDeleteCLSpectrumError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardDeleteCLSpectrum"
ierror = True
Exit Sub

End Sub

Sub StandardDisplayEDSSpectrum(stdnum As Integer, specnum As Integer, sample() As TypeSample)
' Display the specified EDS spectrum

ierror = False
On Error GoTo StandardDisplayEDSSpectrumError

Dim chan As Integer
Dim tfilename As String

' Init the sample (for EDS and CL arrays)
Call InitSample(StandardTmpSample())
If ierror Then Exit Sub

ReDim StandardTmpSample(1).EDSSpectraIntensities(1 To MAXROW%, 1 To MAXSPECTRA%) As Long
ReDim StandardTmpSample(1).EDSSpectraStrobes(1 To MAXROW%, 1 To MAXSTROBE%) As Long    ' only for Oxford EDS

' Get EDS spectrum data for the standard
Call StandardEDSSpectraGetData(stdnum%, specnum%, tfilename$, StandardTmpSample())
If ierror Then Exit Sub

StandardTmpSample(1).Name$ = MiscGetFileNameOnly$(tfilename$)

' Display the EDS spectra in cps
Call EDSInitDisplay(FormEDSDISPLAY3, tfilename$, Int(1), StandardTmpSample())
If ierror Then Exit Sub

' Load elements from standard compositional sample to EDS spectrum sample
For chan% = 1 To sample(1).LastChan%
StandardTmpSample(1).Elsyms$(chan%) = sample(1).Elsyms$(chan%)
Next chan%
StandardTmpSample(1).LastChan% = sample(1).LastChan%

' Display the EDS spectra in cps
Call EDSDisplaySpectra(FormEDSDISPLAY3, Int(1), StandardTmpSample())
If ierror Then Exit Sub

FormEDSDISPLAY3.Show vbModal
Exit Sub

' Errors
StandardDisplayEDSSpectrumError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardDisplayEDSSpectrum"
ierror = True
Exit Sub

End Sub

Sub StandardDisplayCLSpectrum(stdnum As Integer, specnum As Integer, sample() As TypeSample)
' Display the specified CL spectrum

ierror = False
On Error GoTo StandardDisplayCLSpectrumError

Dim tfilename As String

' Init the sample (for EDS and CL arrays)
Call InitSample(StandardTmpSample())
If ierror Then Exit Sub

ReDim StandardTmpSample(1).CLSpectraIntensities(1 To MAXROW%, 1 To MAXSPECTRA_CL%) As Long
ReDim StandardTmpSample(1).CLSpectraDarkIntensities(1 To MAXROW%, 1 To MAXSPECTRA_CL%) As Long

' Get CL spectrum data for the standard
Call StandardCLSpectraGetData(stdnum%, specnum%, tfilename$, StandardTmpSample())
If ierror Then Exit Sub

' Show the user the EDS spectrum
StandardTmpSample(1).Name$ = MiscGetFileNameOnly$(tfilename$)
Call CLInitDisplay(FormCLDISPLAY, tfilename$, Int(1), StandardTmpSample())
If ierror Then Exit Sub

' Display the CL spectra in cps (tCLIntensityOption% = 1)
Call CLDisplaySpectra(False, FormCLDISPLAY, Int(1), StandardTmpSample())
If ierror Then Exit Sub

FormCLDISPLAY.Show vbModal

Exit Sub

' Errors
StandardDisplayCLSpectrumError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardDisplayCLSpectrum"
ierror = True
Exit Sub

End Sub

Sub StandardMemoTextGet(stdnum As Integer, memostring As String)
' Get memo text from the specified standard
' stdnum = standard number (unique)
' memostring = memo text

ierror = False
On Error GoTo StandardMemoTextGetError

Dim SQLQ As String

Dim StDb As Database
Dim StRs As Recordset

' Open the standard database and get memo text
Screen.MousePointer = vbHourglass
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseNonExclusiveAccess%, dbReadOnly)

SQLQ$ = "SELECT MemoText.* FROM MemoText WHERE MemoText.MemoTextToNumber = " & Str$(stdnum%)
Set StRs = StDb.OpenRecordset(SQLQ$, dbOpenSnapshot)

' Load memo text field
memostring$ = vbNullString
If Not StRs.EOF Then memostring$ = Trim$(vbNullString & StRs("MemoTextField"))
StRs.Close

' Close the standard database
Screen.MousePointer = vbDefault
StDb.Close

' Make sure objects are deallocated
If Not StRs Is Nothing Then Set StRs = Nothing
If Not StDb Is Nothing Then Set StDb = Nothing

Exit Sub

' Errors
StandardMemoTextGetError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardMemoTextGet"
ierror = True
Exit Sub

End Sub

Sub StandardMemoTextSend(stdnum As Integer, memostring As String)
' Write the memmo string to the memo text field
' stdnum = standard number (unique)
' memostring = memo text string

ierror = False
On Error GoTo StandardMemoTextSendError

Dim SQLQ As String
Dim StDb As Database
Dim StRs As Recordset

' Check limits
If stdnum% < 1 Or stdnum% > MAXINTEGER% Then GoTo StandardMemoTextSendBadStandard

' Open the Standard database and write new data to MemoText table
Screen.MousePointer = vbHourglass
Set StDb = OpenDatabase(StandardDataFile$, StandardDatabaseExclusiveAccess%, False)

' Open "MemoText" table
Set StRs = StDb.OpenRecordset("MemoText", dbOpenTable)

Call TransactionBegin("StandardMemoTextSend", StandardDataFile$)
If ierror Then Exit Sub

' Delete memo text table based on "stdnum"
SQLQ$ = "DELETE from MemoText WHERE MemoText.MemoTextToNumber = " & Str$(stdnum%)
StDb.Execute SQLQ$

' Save memo text
StRs.AddNew
StRs("MemoTextToNumber") = stdnum%
StRs("MemoTextField") = Left$(Trim$(memostring$), DbTextMemoStringLength&)
StRs.Update

Call TransactionCommit("StandardMemoTextSend", StandardDataFile$)
If ierror Then Exit Sub

StRs.Close

' Close the standard database
Screen.MousePointer = vbDefault
StDb.Close

' Make sure objects are deallocated
If Not StRs Is Nothing Then Set StRs = Nothing
If Not StDb Is Nothing Then Set StDb = Nothing
Exit Sub

' Errors
StandardMemoTextSendError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardMemoTextSend"
Call TransactionRollback("StandardMemoTextSend", StandardDataFile$)
ierror = True
Exit Sub

StandardMemoTextSendBadStandard:
msg$ = "Memo text standard number is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardMemoTextSend"
ierror = True
Exit Sub

End Sub

