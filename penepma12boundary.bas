Attribute VB_Name = "CodePenepma12Boundary"
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

' Boundary globals
Global BinaryElementsSwappedA As Boolean, BinaryElementsSwappedB As Boolean

Global Boundary_ZAF_Kratios() As Double
Global Boundary_ZAF_Factors() As Single
Global Boundary_ZAF_Betas() As Single

Global Boundary_Linear_Distances() As Single
Global Boundary_Mass_Distances() As Single

Global Boundary_Material_A_Densities() As Single
Global Boundary_Material_B_Densities() As Single

Sub Penepma12BoundaryNewMDB()
' This routine create a new the boundary.mdb file

ierror = False
On Error GoTo Penepma12BoundaryNewMDBError

Dim response As Integer

Dim MtDb As Database

' Specify the boundary database variables
Dim Boundary As TableDef
Dim BoundaryKRatio As TableDef
Dim BoundaryMassDistance As TableDef
Dim BoundaryMaterialDensity As TableDef

Dim BoundaryIndex As New Index

' If file already exists, warn user
If Dir$(BoundaryMDBFile$) <> vbNullString Then
msg$ = "Boundary Database: " & vbCrLf
msg$ = msg$ & BoundaryMDBFile$ & vbCrLf
msg$ = msg$ & " already exists, are you sure you want to overwrite it?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton2, "Penepma12BoundaryNewMDB")
If response% = vbNo Then
ierror = True
Exit Sub
End If

' If boundary database exists, delete it
If Dir$(BoundaryMDBFile$) <> vbNullString Then
Kill BoundaryMDBFile$

' Else inform user
Else
msg$ = "Creating a new Boundary K-ratio database: " & BoundaryMDBFile$
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12BoundaryNewMDB"
End If
End If

' Open a new database by copying from existing MDB template
Screen.MousePointer = vbHourglass
Call FileInfoCreateDatabase(BoundaryMDBFile$)
If ierror Then Exit Sub

' Open as existing database
Set MtDb = OpenDatabase(BoundaryMDBFile$, DatabaseExclusiveAccess%, False)

' Specify the Boundary database "Boundary" table
Set Boundary = MtDb.CreateTableDef("NewTableDef")
Boundary.Name = "Boundary"

' Create Boundary table fields
With Boundary
.Fields.Append .CreateField("BeamTakeOff", dbSingle)
.Fields.Append .CreateField("BeamEnergy", dbSingle)
.Fields.Append .CreateField("EmittingElement", dbInteger)
.Fields.Append .CreateField("EmittingXray", dbInteger)

.Fields.Append .CreateField("MatrixElementA1", dbInteger)
.Fields.Append .CreateField("MatrixElementA2", dbInteger)
.Fields.Append .CreateField("BoundaryElementB1", dbInteger)
.Fields.Append .CreateField("BoundaryElementB2", dbInteger)

.Fields.Append .CreateField("PointNumber", dbLong)     ' the distance point number (1 to npoints&)
.Fields.Append .CreateField("LinearDistance", dbSingle)

' Add unique record number for data tables
.Fields.Append .CreateField("BoundaryNumber", dbLong)
End With

' Specify the Boundary database "BoundaryIndex" index
BoundaryIndex.Name = "BoundaryNumbers"
BoundaryIndex.Fields = "BoundaryNumber"
BoundaryIndex.Primary = True
Boundary.Indexes.Append BoundaryIndex

MtDb.TableDefs.Append Boundary

' Make mass distance table (for material A only)
Set BoundaryMassDistance = MtDb.CreateTableDef("NewTableDef")
BoundaryMassDistance.Name = "BoundaryMassDistance"

With BoundaryMassDistance
.Fields.Append .CreateField("BoundaryMassDistanceNumber", dbLong)     ' unique record number pointing to BoundaryNumbers in Boundary table
.Fields.Append .CreateField("BoundaryMassDistanceOrder", dbInteger)   ' load order (1 to MAXBINARY%) for material A only
.Fields.Append .CreateField("BoundaryMassDistance", dbSingle)
End With

MtDb.TableDefs.Append BoundaryMassDistance

' Make density table (materials A and B)
Set BoundaryMaterialDensity = MtDb.CreateTableDef("NewTableDef")
BoundaryMaterialDensity.Name = "BoundaryMaterialDensity"

With BoundaryMaterialDensity
.Fields.Append .CreateField("BoundaryMaterialDensityNumber", dbLong)     ' unique record number pointing to BoundaryNumbers in Boundary table
.Fields.Append .CreateField("BoundaryMaterialDensityOrder", dbInteger)   ' load order (1 to MAXBINARY%) for material A and B
.Fields.Append .CreateField("BoundaryMaterialDensityMaterialA", dbSingle)
.Fields.Append .CreateField("BoundaryMaterialDensityMaterialB", dbSingle)
End With

MtDb.TableDefs.Append BoundaryMaterialDensity

' Make k-ratio table
Set BoundaryKRatio = MtDb.CreateTableDef("NewTableDef")
BoundaryKRatio.Name = "BoundaryKratio"

' Create Boundary k-ratio table fields
With BoundaryKRatio
.Fields.Append .CreateField("BoundaryKRatioNumber", dbLong)     ' unique record number pointing to BoundaryNumbers in Boundary table
.Fields.Append .CreateField("BoundaryKRatioOrderA", dbInteger)   ' load order (1 to MAXBINARY%)
.Fields.Append .CreateField("BoundaryKRatioOrderB", dbInteger)   ' load order (1 to MAXBINARY%)
.Fields.Append .CreateField("BoundaryKRatio_ZAF_KRatio", dbDouble)
End With

MtDb.TableDefs.Append BoundaryKRatio

' Specify the boundary database "BoundaryIndex" index
Set BoundaryIndex = Boundary.CreateIndex("MatrixIndexSecondary")

With BoundaryIndex
.Fields.Append .CreateField("BoundaryKRatioNumber")      ' unique record number pointing to BoundaryKRatio table
End With

BoundaryKRatio.Indexes.Append BoundaryIndex

' Close the database
MtDb.Close
Screen.MousePointer = vbDefault

' Create new File table for Boundary database
Call FileInfoMakeNewTable(Int(9), BoundaryMDBFile$)
If ierror Then Exit Sub

msg$ = "New Boundary.MDB has been created"
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12BoundaryNewMDB"

Exit Sub

' Errors
Penepma12BoundaryNewMDBError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12BoundaryNewMDB"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12BoundaryReadMDB(tTakeoff As Single, tKilovolt As Single, tEmitter As Integer, tXray As Integer, tMatrixElementA1 As Integer, tMatrixElementA2 As Integer, tBoundaryElementB1 As Integer, tBoundaryElementB2 As Integer, tKratios() As Double, tLinearDistances() As Single, tMassDistances() As Single, tMaterialDensitiesA() As Single, tMaterialDensitiesB() As Single, nPoints As Long, notfound As Boolean)
' This routine reads the Boundary.mdb file for the specified beam energy, emitter, x-ray, matrix.
'  tKratios#(1 to MAXBINARY%, 1 to MAXBINARY%, 1 to npoints&)  are the k-ratios for this x-ray and binary composition
'  tLinearDistances(1 to npoints&) are the linear distances
'
'  tMassDistances(1 to MAXBINARY%, 1 to npoints&) are the mass distances (for material A)
'  tMaterialDensitiesA(1 to MAXBINARY%) are the densities (for material A)
'  tMaterialDensitiesB(1 to MAXBINARY%) are the densities (for material B)
'
'  npoints&  is the total number of distance points for each matrix-boundary binary pair
'  notfound  is a boolean whether the specified matrix-boundary was found for the meitting element and x-ray

ierror = False
On Error GoTo Penepma12BoundaryReadMDBError

Dim j As Integer, k As Integer
Dim n As Long
Dim nrec As Long
Dim astring As String

Dim SQLQ As String
Dim MtDb As Database
Dim MtDs As Recordset
Dim MtRs As Recordset

' Check for file
If Dir$(BoundaryMDBFile$) = vbNullString Then GoTo Penepma12BoundaryReadMDBNoBoundaryMDBFile

' Load string for matrix and boundary binaries
astring$ = Trim$(Symup$(tMatrixElementA1%)) & "-" & Trim$(Symup$(tMatrixElementA2%)) & " adjacent to " & Trim$(Symup$(tBoundaryElementB1%)) & "-" & Trim$(Symup$(tBoundaryElementB2%))

' Open boundary database (exclusive and read only)
Screen.MousePointer = vbHourglass
Set MtDb = OpenDatabase(BoundaryMDBFile$, BoundaryDatabaseNonExclusiveAccess%, dbReadOnly)

' Try to find requested emitter, matrix binary, boundary binary, etc
SQLQ$ = "SELECT Boundary.* FROM Boundary WHERE BeamTakeOff = " & Format$(tTakeoff!)
SQLQ$ = SQLQ$ & " AND BeamEnergy = " & Format$(tKilovolt!) & " AND EmittingElement = " & Format$(tEmitter%)
SQLQ$ = SQLQ$ & " AND EmittingXray = " & Format$(tXray%)
SQLQ$ = SQLQ$ & " AND MatrixElementA1 = " & Format$(tMatrixElementA1%)
SQLQ$ = SQLQ$ & " AND MatrixElementA2 = " & Format$(tMatrixElementA2%)
SQLQ$ = SQLQ$ & " AND BoundaryElementB1 = " & Format$(tBoundaryElementB1%)
SQLQ$ = SQLQ$ & " AND BoundaryElementB2 = " & Format$(tBoundaryElementB2%)
SQLQ$ = SQLQ$ & " ORDER BY PointNumber"

Set MtDs = MtDb.OpenRecordset(SQLQ$, dbOpenSnapshot)

' If record not found, return notfound
If MtDs.BOF And MtDs.EOF Then
notfound = True
Screen.MousePointer = vbDefault
Exit Sub
End If

' Loop on all records for all distances and load return values based on "BoundaryNumber"
nPoints& = 0
Do Until MtDs.EOF
nrec& = MtDs("BoundaryNumber")  ' pointer to mass distance, density, k-ratio, afactor and coeff tables
nPoints& = nPoints& + 1

' Load data for this distance
n& = MtDs("PointNumber")

' Check point number
If n& <> nPoints& Then GoTo Penepma12BoundaryReadMDBBadPoint

' Confirm to status
Call IOStatusAuto("Reading " & Symup$(tEmitter%) & " " & Xraylo$(tXray%) & " (in " & Symup$(tMatrixElementA1%) & "-" & Symup$(tMatrixElementA2%) & " adjacent to " & Symup$(tBoundaryElementB1%) & "-" & Symup$(tBoundaryElementB2%) & ") record " & Format$(nrec&) & " point " & Format$(n&) & "...")

' Dimension passed variables
ReDim Preserve tLinearDistances(1 To n&) As Single
tLinearDistances(n&) = MtDs("LinearDistance")

' Search for mass distance records
SQLQ$ = "SELECT BoundaryMassDistance.* FROM BoundaryMassDistance WHERE BoundaryMassDistanceNumber = " & Format$(nrec&)
Set MtRs = MtDb.OpenRecordset(SQLQ$, dbOpenSnapshot)
If MtRs.BOF And MtRs.EOF Then GoTo Penepma12BoundaryReadMDBNoMassDistances

' Dimension passed variables
ReDim Preserve tMassDistances(1 To MAXBINARY%, 1 To n&) As Single

' Load mass distance array
Do Until MtRs.EOF
j% = MtRs("BoundaryMassDistanceOrder")          ' load order (1 to MAXBINARY% for Material A only)
tMassDistances(j%, n&) = MtRs("BoundaryMassDistance")
MtRs.MoveNext
Loop
MtRs.Close

' Search for density records (material A and material B) (no need to dimension variables)
SQLQ$ = "SELECT BoundaryMaterialDensity.* FROM BoundaryMaterialdensity WHERE BoundaryMaterialDensityNumber = " & Format$(nrec&)
Set MtRs = MtDb.OpenRecordset(SQLQ$, dbOpenSnapshot)
If MtRs.BOF And MtRs.EOF Then GoTo Penepma12BoundaryReadMDBNoMaterialDensities

Do Until MtRs.EOF
j% = MtRs("BoundaryMaterialDensityOrder")          ' load order (1 to MAXBINARY% for material A and B)
tMaterialDensitiesA!(j%) = MtRs("BoundaryMaterialDensityMaterialA")
tMaterialDensitiesB!(j%) = MtRs("BoundaryMaterialDensityMaterialB")
MtRs.MoveNext
Loop
MtRs.Close

' Search for records
SQLQ$ = "SELECT BoundaryKRatio.* FROM BoundaryKRatio WHERE BoundaryKRatioNumber = " & Format$(nrec&)
Set MtRs = MtDb.OpenRecordset(SQLQ$, dbOpenSnapshot)
If MtRs.BOF And MtRs.EOF Then GoTo Penepma12BoundaryReadMDBNoKRatios

' Dimension passed variables
ReDim Preserve tKratios(1 To MAXBINARY%, 1 To MAXBINARY%, 1 To n&) As Double

' Load kratio array
Do Until MtRs.EOF
j% = MtRs("BoundaryKRatioOrderA")          ' load order (1 to MAXBINARY%)
k% = MtRs("BoundaryKRatioOrderB")          ' load order (1 to MAXBINARY%)
tKratios#(k%, j%, n&) = MtRs("BoundaryKRatio_ZAF_KRatio")
MtRs.MoveNext
Loop
MtRs.Close

MtDs.MoveNext
Loop
Call IOStatusAuto(vbNullString)

MtDs.Close
If nPoints& > 0 Then notfound = False
MtDb.Close

Screen.MousePointer = vbDefault
Exit Sub

' Errors
Penepma12BoundaryReadMDBError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12BoundaryReadMDB"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12BoundaryReadMDBNoBoundaryMDBFile:
msg$ = "File " & BoundaryMDBFile$ & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12BoundaryReadMDB"
ierror = True
Exit Sub

Penepma12BoundaryReadMDBBadPoint:
msg$ = "Point loading was out of order, expected point number " & Format$(nPoints&) & " but found point number " & Format$(n&) & "."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12BoundaryReadMDB"
ierror = True
Exit Sub

Penepma12BoundaryReadMDBNoMassDistances:
msg$ = "File " & BoundaryMDBFile$ & " did not contain any mass distance records for " & Format$(tTakeoff!) & " degrees, " & Format$(tKilovolt!) & " keV, " & Symup$(tEmitter%) & " " & Xraylo$(tXray%) & " in " & astring$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12BoundaryReadMDB"
ierror = True
Exit Sub

Penepma12BoundaryReadMDBNoMaterialDensities:
msg$ = "File " & BoundaryMDBFile$ & " did not contain any material density records for " & Format$(tTakeoff!) & " degrees, " & Format$(tKilovolt!) & " keV, " & Symup$(tEmitter%) & " " & Xraylo$(tXray%) & " in " & astring$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12BoundaryReadMDB"
ierror = True
Exit Sub

Penepma12BoundaryReadMDBNoKRatios:
msg$ = "File " & BoundaryMDBFile$ & " did not contain any k-ratio records for " & Format$(tTakeoff!) & " degrees, " & Format$(tKilovolt!) & " keV, " & Symup$(tEmitter%) & " " & Xraylo$(tXray%) & " in " & astring$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12BoundaryReadMDB"
ierror = True
Exit Sub

Penepma12BoundaryReadMDBNoFactors:
msg$ = "File " & BoundaryMDBFile$ & " did not contain any alpha factor records for " & Format$(tTakeoff!) & " degrees, " & Format$(tKilovolt!) & " keV, " & Symup$(tEmitter%) & " " & Xraylo$(tXray%) & " in " & astring$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12BoundaryReadMDB"
ierror = True
Exit Sub

End Sub

Sub Penepma12CalculateReadWriteBinaryDataBoundary(mode As Integer, tfolder As String, tfilename As String, keV As Single, nPoints As Long)
' Reads or write the binary fluorescence boundary k-ratio data to or from a data file for a specified beam energy and distance
'  mode = 0 create file and write column labels only
'  mode = 1 read data
'  mode = 2 write data
'
'  tfolder$ is the full path of the binary compositional data file to read or write
'  tfilename$ is the filename of the binary compositional data file to read or write
'  keV is the specified beam energy
'  npoints& is the number of distance points read or to write
'
'  Boundary_ZAF_Kratios#(1 to MAXBINARY%, 1 to MAXBINARY%, 1 to npoints%)  are the k-ratios from Fanal in k-ratio % for each x distance
'  Boundary_ZAF_Factors!(1 to MAXBINARY%, 1 to MAXBINARY%, 1 to npoints%)  are the alpha factors for each boundary composition, alpha = (C/K - C)/(1 - C)
'
'  Boundary_Linear_Distances!(1 to npoints%) are linear distances for material A
'  Boundary_Mass_Distances!(1 to MAXBINARY%, 1 to npoints%)   are mass distances for material A binary compositions
'
'  Boundary_Material_A_Densities!(1 to MAXBINARY%)    are densities for material A
'  Boundary_Material_B_Densities!(1 to MAXBINARY%)    are densities for material B

ierror = False
On Error GoTo Penepma12CalculateReadWriteBinaryDataBoundaryError

Dim n As Integer
Dim j As Integer, k As Integer
Dim tkeV As Single
Dim txdist As Single, tmdist As Single
Dim astring As String, ttfilename As String
Dim jstring As String, kstring As String
Dim temp2 As Single
Dim tdensityA As Single, tdensityB As Single
Dim temp1 As Double

' Write column labels only
If mode% = 0 Then
Close #Temp1FileNumber%
ttfilename$ = tfolder$ & "\" & tfilename$
Open ttfilename$ For Output As #Temp1FileNumber%

' Load output string for keV
astring$ = VbDquote$ & "keV" & VbDquote$ & vbTab

' Load output string for linear distances
astring$ = astring$ & VbDquote$ & "Xdist" & VbDquote$ & vbTab

' Load data column labels (note that these labels should always be unswapped)
For j% = 1 To MAXBINARY%    ' material A
For k% = 1 To MAXBINARY%    ' material B
jstring$ = Format$(BinaryRanges!(j%)) & "-" & Format$(100# - BinaryRanges!(j%))
kstring$ = Format$(BinaryRanges!(k%)) & "-" & Format$(100# - BinaryRanges!(k%))

astring$ = astring$ & VbDquote$ & jstring$ & "_" & kstring$ & "_Krat" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & jstring$ & "_" & kstring$ & "_Alpha" & VbDquote$ & vbTab
Next k%
Next j%

' Load output string for mass distances and densities
For j% = 1 To MAXBINARY%    ' material A only
If Not BinaryElementsSwappedA Then
jstring$ = Format$(BinaryRanges!(j%)) & "-" & Format$(100# - BinaryRanges!(j%))
Else
jstring$ = Format$(BinaryRanges!(MAXBINARY - (j% - 1))) & "-" & Format$(100# - MAXBINARY - (j% - 1))
End If

astring$ = astring$ & VbDquote$ & "Mdist_" & jstring$ & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "DensityA_" & jstring$ & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "DensityB_" & kstring$ & VbDquote$ & vbTab
Next j%

' Output column labels
Print #Temp1FileNumber%, astring$
Close #Temp1FileNumber%
End If

' Read data for specified beam energy
If mode% = 1 Then
Close #Temp1FileNumber%
ttfilename$ = tfolder$ & "\" & tfilename$
If Dir$(ttfilename$) = vbNullString Then GoTo Penepma12CalculateReadWriteBinaryDataBoundaryFileNotFound
Open ttfilename$ For Input As #Temp1FileNumber%

' Read the column labels
Line Input #Temp1FileNumber%, astring$

' Loop on file until desired voltage is found
nPoints& = 0
Do Until EOF(Temp1FileNumber%)

' Read temp keV and x distance into temp variables
Input #Temp1FileNumber%, tkeV!, txdist!

' If keV matches then load x distance
If keV! = tkeV! Then
nPoints& = nPoints& + 1
ReDim Preserve Boundary_Linear_Distances(1 To nPoints&) As Single
Boundary_Linear_Distances!(nPoints&) = txdist!
End If

' Input 2 dimensional k-ratios and alpha factors
For j% = 1 To MAXBINARY%        ' material A
For k% = 1 To MAXBINARY%        ' material B
Input #Temp1FileNumber%, temp1#
Input #Temp1FileNumber%, temp2!

' If keV matches then load k-ratios and alpha-factors
If keV! = tkeV! Then
ReDim Preserve Boundary_ZAF_Kratios(1 To MAXBINARY%, 1 To MAXBINARY%, 1 To nPoints&) As Double
ReDim Preserve Boundary_ZAF_Factors(1 To MAXBINARY%, 1 To MAXBINARY%, 1 To nPoints&) As Single
Boundary_ZAF_Kratios#(k%, j%, nPoints&) = temp1#
Boundary_ZAF_Factors!(k%, j%, nPoints&) = temp2!
End If

Next k%
Next j%

' Input mass distances for material A only and densities for material A and material B
For j% = 1 To MAXBINARY%        ' material A and B

' Read temp mass distance into temp variable
Input #Temp1FileNumber%, tmdist!, tdensityA!, tdensityB!

' If keV matches then load mass distance and density (densities are the same for all points-distances)
If keV! = tkeV! Then
ReDim Preserve Boundary_Mass_Distances(1 To MAXBINARY%, 1 To nPoints&) As Single
Boundary_Mass_Distances!(j%, nPoints&) = tmdist!
Boundary_Material_A_Densities!(j%) = tdensityA!
Boundary_Material_B_Densities!(j%) = tdensityB!
End If
Next j%

Loop

Close #Temp1FileNumber%
End If

' Write data for specified beam energy (must be written in consecutive keV order)
If mode% = 2 Then
Open tfolder$ & "\" & tfilename$ For Append As #Temp1FileNumber%

' Calculate boundary alpha factors and output k-ratios, factors and fit coefficients for each x distance
For n% = 1 To nPoints&

' Load output string for keV and linear distance
astring$ = Format$(keV!) & vbTab & Format$(Boundary_Linear_Distances!(n%)) & vbTab

' Output k-ratios, factors and fit coefficients
For j% = 1 To MAXBINARY%    ' material A
For k% = 1 To MAXBINARY%    ' material B

' Create ZAF output string
astring$ = astring$ & Format$(Boundary_ZAF_Kratios#(k%, j%, n%)) & vbTab
astring$ = astring$ & Format$(Boundary_ZAF_Factors!(k%, j%, n%)) & vbTab
Next k%
Next j%

' Add string for mass distances
For j% = 1 To MAXBINARY%    ' material A and B
astring$ = astring$ & Format$(Boundary_Mass_Distances!(j%, n%)) & vbTab

' Add string for densities (always the same for all points)
astring$ = astring$ & Format$(Boundary_Material_A_Densities!(j%)) & vbTab
astring$ = astring$ & Format$(Boundary_Material_B_Densities!(j%)) & vbTab
Next j%

' Output data for this distance
Print #Temp1FileNumber%, astring$
Next n%

Close #Temp1FileNumber%
End If

Exit Sub

' Errors
Penepma12CalculateReadWriteBinaryDataBoundaryError:
MsgBox Error$ & ", " & tfolder$ & "\" & tfilename$, vbOKOnly + vbCritical, "Penepma12CalculateReadWriteBinaryDataBoundary"
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12CalculateReadWriteBinaryDataBoundaryFileNotFound:
msg$ = "The binary composition file " & tfolder$ & "\" & tfilename$ & " was not found."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CalculateReadWriteBinaryDataBoundary"
ierror = True
Exit Sub

End Sub

Sub Penepma12BoundaryScanMDB()
' This routine scans for all k-ratio input files and adds them to a new Boundary.mdb file
'  with new k-ratios, alpha factors and fit coefficients

ierror = False
On Error GoTo Penepma12BoundaryScanMDBError

Dim j As Integer, k As Integer
Dim m As Integer
Dim nrec As Long
Dim i As Long, ii As Long
Dim tfilename As String, tfolder As String
Dim astring As String, bstring As String
Dim eng As Single, edg As Single
Dim tovervoltage As Single

Dim filearray() As String

Dim n As Long, nPoints As Long
Dim BeamTakeOff As Single
Dim BeamEnergy As Single
Dim EmittingElement As Integer
Dim EmittingXray As Integer

Dim MatrixElementA1 As Integer
Dim MatrixElementA2 As Integer
Dim BoundaryElementB1 As Integer
Dim BoundaryElementB2 As Integer

Dim MtDb As Database
Dim MtDt As Recordset

icancelauto = False

' If file does not exist, warn user
If Dir$(BoundaryMDBFile$) = vbNullString Then
msg$ = "Boundary Database: " & vbCrLf
msg$ = msg$ & BoundaryMDBFile$ & vbCrLf
msg$ = msg$ & " does not exist. Please create a new Boundary.MDB file try updating again."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12BoundaryScanMDB"
ierror = True
Exit Sub
End If

' Check for Fanal\boundary folder
tfolder$ = PENEPMA_Root$ & "\Fanal\boundary"
If Dir$(tfolder$, vbDirectory) = vbNullString Then GoTo Penepma12BoundaryScanMDBNoDirectory

' Make a list of all input files (must do this way to avoid reentrant Dir$ calls)
tfilename$ = Dir$(PENEPMA_Root$ & "\Fanal\Boundary\" & "\*.TXT")  ' get first file
ii& = 0
Do While tfilename$ <> vbNullString
ii& = ii& + 1
ReDim Preserve filearray(1 To ii&) As String
filearray$(ii&) = tfilename$
tfilename$ = Dir$
Loop

' Delete Standard.txt and standard.err file if present
If Dir$(ProbeTextLogFile$) <> vbNullString Then Kill ProbeTextLogFile$
If Dir$(ProbeErrorLogFile$) <> vbNullString Then Kill ProbeErrorLogFile$

' Open the Boundary.mdb
Set MtDb = OpenDatabase(BoundaryMDBFile$, BoundaryDatabaseExclusiveAccess%, False)

' Check if database already has entries
Set MtDt = MtDb.OpenRecordset("Boundary", dbOpenTable)
If Not (MtDt.BOF And MtDt.EOF) Then GoTo Penepma12BoundaryScanMDBNotEmpty
MtDt.Close

' Loop through all input files
nrec& = 0
For i& = 1 To ii&
tfilename$ = filearray$(i&)

' Determine the emitting element and Boundary element from the filename
astring$ = MiscGetFileNameOnly$(tfilename$)

Call MiscParseStringToStringA(astring$, "-", bstring$)
MatrixElementA1% = Val(bstring$)
Call MiscParseStringToStringA(astring$, "_", bstring$)
MatrixElementA2% = Val(bstring$)

Call MiscParseStringToStringA(astring$, "-", bstring$)
BoundaryElementB1% = Val(bstring$)
Call MiscParseStringToStringA(astring$, "_", bstring$)
BoundaryElementB2% = Val(bstring$)

Call MiscParseStringToStringA(astring$, "_", bstring$)
BeamTakeOff! = Val(bstring$)

Call MiscParseStringToStringA(astring$, "-", bstring$)
EmittingElement% = Val(bstring$)
Call MiscParseStringToStringA(astring$, ".", bstring$)
EmittingXray% = Val(bstring$)

' Get emitting energy and edge energy
Call XrayGetEnergy(EmittingElement%, EmittingXray%, eng!, edg!)
If ierror Then Exit Sub

' Loop on each possible energy
For m% = 1 To 50
'For m% = 15 To 15       ' testing purposes (15 keV only)
'For m% = 15 To 16       ' testing purposes (15 and 16 keV only)
BeamEnergy! = CSng(m%)

Call IOWriteLog("Reading data at " & Format$(m%) & " keV, for " & Symup$(EmittingElement%) & " " & Xraylo$(EmittingXray%) & ", from input file " & tfilename$ & "...")
DoEvents

' Read binary k-ratio fluorescence data to file for the specified beam energy
Call Penepma12CalculateReadWriteBinaryDataBoundary(Int(1), tfolder$, tfilename$, CSng(m%), nPoints&)
If ierror Then Exit Sub

' Load minimum overvoltage, 0 = 2%, 1 = 10%, 2 = 20%, 3 = 40%
If MinimumOverVoltageType% = 0 Then tovervoltage! = MINIMUMOVERVOLTFRACTION_02!
If MinimumOverVoltageType% = 1 Then tovervoltage! = MINIMUMOVERVOLTFRACTION_10!
If MinimumOverVoltageType% = 2 Then tovervoltage! = MINIMUMOVERVOLTFRACTION_20!
If MinimumOverVoltageType% = 3 Then tovervoltage! = MINIMUMOVERVOLTFRACTION_40!

' Check for valid x-ray line (excitation energy (plus a buffer to avoid ultra low overvoltage issues) must be less than beam energy) (and greater than PenepmaMinimumElectronEnergy!)
If eng! <> 0# And edg! <> 0# And (edg! * (1# + tovervoltage!) < BeamEnergy!) And edg! > PenepmaMinimumElectronEnergy! And nPoints& > 0 Then

' Add each distance for this beam energy and specified x-ray
For n& = 1 To nPoints&

' Add new records to "Boundary" table
Set MtDt = MtDb.OpenRecordset("Boundary", dbOpenTable)
Call IOStatusAuto("Adding record " & Format$(nrec& + 1) & ", " & Format$(m%) & " keV, " & Symup$(EmittingElement%) & " " & Xraylo$(EmittingXray%) & " (npoint=" & Format$(n&) & ") to Boundary.MDB with input file, " & tfilename$ & "...")
DoEvents

' Add new record
nrec& = nrec& + 1

MtDt.AddNew
MtDt("BeamTakeOff") = BeamTakeOff!
MtDt("BeamEnergy") = BeamEnergy!

MtDt("EmittingElement") = EmittingElement%
MtDt("EmittingXray") = EmittingXray%

MtDt("MatrixElementA1") = MatrixElementA1%
MtDt("MatrixElementA2") = MatrixElementA2%
MtDt("BoundaryElementB1") = BoundaryElementB1%
MtDt("BoundaryElementB2") = BoundaryElementB2%

MtDt("LinearDistance") = Abs(Boundary_Linear_Distances!(n&))
MtDt("PointNumber") = n&    ' loading order

' Add unique record number for other tables
MtDt("BoundaryNumber") = nrec&
MtDt.Update
MtDt.Close

' Add records to mass distance table
Set MtDt = MtDb.OpenRecordset("BoundaryMassDistance", dbOpenTable)
For j% = 1 To MAXBINARY%        ' material A
MtDt.AddNew
MtDt("BoundaryMassDistanceNumber") = nrec&
MtDt("BoundaryMassDistanceOrder") = j%
MtDt("BoundaryMassDistance") = Boundary_Mass_Distances!(j%, n&)
MtDt.Update
Next j%
MtDt.Close

' Add records to density table
Set MtDt = MtDb.OpenRecordset("BoundaryMaterialDensity", dbOpenTable)
For j% = 1 To MAXBINARY%        ' material A and B
MtDt.AddNew
MtDt("BoundaryMaterialDensityNumber") = nrec&
MtDt("BoundaryMaterialDensityOrder") = j%
MtDt("BoundaryMaterialDensityMaterialA") = Boundary_Material_A_Densities!(j%)
MtDt("BoundaryMaterialDensityMaterialB") = Boundary_Material_B_Densities!(j%)
MtDt.Update
Next j%
MtDt.Close

' Add new records to "Kratios" table
Set MtDt = MtDb.OpenRecordset("BoundaryKratio", dbOpenTable)
For j% = 1 To MAXBINARY%        ' material A
For k% = 1 To MAXBINARY%        ' material B
MtDt.AddNew
MtDt("BoundaryKRatioNumber") = nrec&         ' unique record number pointing to Boundary table
MtDt("BoundaryKRatioOrderA") = j%          ' load order (1 to MAXBINARY%)
MtDt("BoundaryKRatioOrderB") = k%          ' load order (1 to MAXBINARY%)
MtDt("BoundaryKRatio_ZAF_KRatio") = CSng(Boundary_ZAF_Kratios#(k%, j%, n&))
MtDt.Update
Next k%
Next j%
MtDt.Close

' Check for user cancel
DoEvents
If icancelauto Then
ierror = True
Exit Sub
End If

Next n&
End If
Next m%

' Get next input filename
Next i&

MtDb.Close
Call IOStatusAuto(vbNullString)

If nrec& > 0 Then
msg$ = "Boundary.MDB has been updated with " & Format$(nrec&) & " boundary records"
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12BoundaryScanMDB"

Else
msg$ = "No Boundary.MDB k-ratio input files were found"
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12BoundaryScanMDB"
End If

Exit Sub

' Errors
Penepma12BoundaryScanMDBError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12BoundaryScanMDB"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12BoundaryScanMDBNoDirectory:
msg$ = "The boundary data folder " & tfolder$ & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12BoundaryScanMDB"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12BoundaryScanMDBNotEmpty:
msg$ = "The boundary database already contains intensity entires. Please create a new Boundary.mdb file and try updating it again."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12BoundaryScanMDB"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12BoundaryReadMDB2(tTakeoff As Single, tKilovolt As Single, tEmitter As Integer, tXray As Integer, tMatrixElementA1 As Integer, tMatrixElementA2 As Integer, tBoundaryElementB1 As Integer, tBoundaryElementB2 As Integer, tKratios() As Double, tMassDistances() As Single, nPoints As Long, notfound As Boolean)
' This routine reads the Boundary.mdb file for the specified beam energy, emitter, x-ray, matrix and returns
' the k-ratios and mass distances only for mass distance interpolation (not linear distances).
'   tKratios#(1 to MAXBINARY%, 1 to MAXBINARY%, 1 To npoints&)  are the k-ratios for all points
'   tMassDistances(1 to MAXBINARY%, 1 To npoints&) are the mass distances for material A for all points
'   npoints& is the number of returned points
'   notfound  is a boolean whether the specified matrix-boundary was found for the meitting element and x-ray

ierror = False
On Error GoTo Penepma12BoundaryReadMDB2Error

Dim j As Integer, k As Integer
Dim n As Long
Dim nrec As Long
Dim astring As String

Dim SQLQ As String
Dim MtDb As Database
Dim MtDs As Recordset
Dim MtRs As Recordset

' Check for file
If Dir$(BoundaryMDBFile$) = vbNullString Then GoTo Penepma12BoundaryReadMDB2NoBoundaryMDBFile

' Load string for matrix and boundary binaries
astring$ = Trim$(Symup$(tMatrixElementA1%)) & "-" & Trim$(Symup$(tMatrixElementA2%)) & " adjacent to " & Trim$(Symup$(tBoundaryElementB1%)) & "-" & Trim$(Symup$(tBoundaryElementB2%))

' Open boundary database (exclusive and read only)
Screen.MousePointer = vbHourglass
Set MtDb = OpenDatabase(BoundaryMDBFile$, BoundaryDatabaseNonExclusiveAccess%, dbReadOnly)

' Try to find requested emitter, matrix binary, boundary binary, etc
SQLQ$ = "SELECT Boundary.BoundaryNumber, Boundary.PointNumber FROM Boundary WHERE BeamTakeOff = " & Format$(tTakeoff!)
SQLQ$ = SQLQ$ & " AND BeamEnergy = " & Format$(tKilovolt!) & " AND EmittingElement = " & Format$(tEmitter%)
SQLQ$ = SQLQ$ & " AND EmittingXray = " & Format$(tXray%)
SQLQ$ = SQLQ$ & " AND MatrixElementA1 = " & Format$(tMatrixElementA1%)
SQLQ$ = SQLQ$ & " AND MatrixElementA2 = " & Format$(tMatrixElementA2%)
SQLQ$ = SQLQ$ & " AND BoundaryElementB1 = " & Format$(tBoundaryElementB1%)
SQLQ$ = SQLQ$ & " AND BoundaryElementB2 = " & Format$(tBoundaryElementB2%)
SQLQ$ = SQLQ$ & " ORDER BY PointNumber"

Set MtDs = MtDb.OpenRecordset(SQLQ$, dbOpenSnapshot)

' If record not found, return notfound
If MtDs.BOF And MtDs.EOF Then
notfound = True
Screen.MousePointer = vbDefault
Exit Sub
End If

' Loop on all records for all distances and load return values based on "BoundaryNumber"
nPoints& = 0
Do Until MtDs.EOF
nrec& = MtDs("BoundaryNumber")  ' pointer to mass distance, density, k-ratio, afactor and coeff tables
nPoints& = nPoints& + 1

' Load data for this distance
n& = MtDs("PointNumber")

' Check point number
If n& <> nPoints& Then GoTo Penepma12BoundaryReadMDB2BadPoint

' Confirm to status
Call IOStatusAuto("Reading " & Symup$(tEmitter%) & " " & Xraylo$(tXray%) & " (in " & Symup$(tMatrixElementA1%) & "-" & Symup$(tMatrixElementA2%) & " adjacent to " & Symup$(tBoundaryElementB1%) & "-" & Symup$(tBoundaryElementB2%) & ") record " & Format$(nrec&) & " point " & Format$(n&) & "...")

' Dimension passed variables
'ReDim Preserve tLinearDistances(1 To n&) As Single
'tLinearDistances(n&) = MtDs("LinearDistance")

' Search for mass distance records
SQLQ$ = "SELECT BoundaryMassDistance.* FROM BoundaryMassDistance WHERE BoundaryMassDistanceNumber = " & Format$(nrec&)
Set MtRs = MtDb.OpenRecordset(SQLQ$, dbOpenSnapshot)
If MtRs.BOF And MtRs.EOF Then GoTo Penepma12BoundaryReadMDB2NoMassDistances

' Dimension passed variables
ReDim Preserve tMassDistances(1 To MAXBINARY%, 1 To n&) As Single

' Load mass distance array
Do Until MtRs.EOF
j% = MtRs("BoundaryMassDistanceOrder")          ' load order (1 to MAXBINARY% for Material A only)
tMassDistances(j%, n&) = MtRs("BoundaryMassDistance")
MtRs.MoveNext
Loop
MtRs.Close

' Search for density records (material A and material B) (no need to dimension variables)
'SQLQ$ = "SELECT BoundaryMaterialDensity.* FROM BoundaryMaterialdensity WHERE BoundaryMaterialDensityNumber = " & Format$(nrec&)
'Set MtRs = MtDb.OpenRecordset(SQLQ$, dbOpenSnapshot)
'If MtRs.BOF And MtRs.EOF Then GoTo Penepma12BoundaryReadMDB2NoMaterialDensities

'Do Until MtRs.EOF
'j% = MtRs("BoundaryMaterialDensityOrder")          ' load order (1 to MAXBINARY% for material A and B)
'tMaterialDensitiesA!(j%) = MtRs("BoundaryMaterialDensityMaterialA")
'tMaterialDensitiesB!(j%) = MtRs("BoundaryMaterialDensityMaterialB")
'MtRs.MoveNext
'Loop
'MtRs.Close

' Search for records
SQLQ$ = "SELECT BoundaryKRatio.* FROM BoundaryKRatio WHERE BoundaryKRatioNumber = " & Format$(nrec&)
Set MtRs = MtDb.OpenRecordset(SQLQ$, dbOpenSnapshot)
If MtRs.BOF And MtRs.EOF Then GoTo Penepma12BoundaryReadMDB2NoKRatios

' Dimension passed variables
ReDim Preserve tKratios(1 To MAXBINARY%, 1 To MAXBINARY%, 1 To n&) As Double

' Load kratio array
Do Until MtRs.EOF
j% = MtRs("BoundaryKRatioOrderA")          ' load order (1 to MAXBINARY%)
k% = MtRs("BoundaryKRatioOrderB")          ' load order (1 to MAXBINARY%)
tKratios#(k%, j%, n&) = MtRs("BoundaryKRatio_ZAF_KRatio")
MtRs.MoveNext
Loop
MtRs.Close

MtDs.MoveNext
Loop
Call IOStatusAuto(vbNullString)

MtDs.Close
If nPoints& > 0 Then notfound = False
MtDb.Close

Screen.MousePointer = vbDefault
Exit Sub

' Errors
Penepma12BoundaryReadMDB2Error:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12BoundaryReadMDB2"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12BoundaryReadMDB2NoBoundaryMDBFile:
msg$ = "File " & BoundaryMDBFile$ & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12BoundaryReadMDB2"
ierror = True
Exit Sub

Penepma12BoundaryReadMDB2BadPoint:
msg$ = "Point loading was out of order, expected point number " & Format$(nPoints&) & " but found point number " & Format$(n&) & "."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12BoundaryReadMDB2"
ierror = True
Exit Sub

Penepma12BoundaryReadMDB2NoMassDistances:
msg$ = "File " & BoundaryMDBFile$ & " did not contain any mass distance records for " & Format$(tTakeoff!) & " degrees, " & Format$(tKilovolt!) & " keV, " & Symup$(tEmitter%) & " " & Xraylo$(tXray%) & " in " & astring$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12BoundaryReadMDB2"
ierror = True
Exit Sub

Penepma12BoundaryReadMDB2NoMaterialDensities:
msg$ = "File " & BoundaryMDBFile$ & " did not contain any material density records for " & Format$(tTakeoff!) & " degrees, " & Format$(tKilovolt!) & " keV, " & Symup$(tEmitter%) & " " & Xraylo$(tXray%) & " in " & astring$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12BoundaryReadMDB2"
ierror = True
Exit Sub

Penepma12BoundaryReadMDB2NoKRatios:
msg$ = "File " & BoundaryMDBFile$ & " did not contain any k-ratio records for " & Format$(tTakeoff!) & " degrees, " & Format$(tKilovolt!) & " keV, " & Symup$(tEmitter%) & " " & Xraylo$(tXray%) & " in " & astring$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12BoundaryReadMDB2"
ierror = True
Exit Sub

End Sub

Sub Penepma12BoundaryInterpolate(tDistanceMass As Single, tKratios2() As Double, tMassDistances() As Single, nPoints As Long, tKratios3() As Single)
' Fit the k-ratios for the specified mass distance and return mass distance interpolated k-ratios
'   tDistanceMass! is the specified mass distance to interpolate to
'   tKratios2#(1 To MAXBINARY%, 1 To MAXBINARY%, 1 To npoints&) are the k-ratios for all points
'   tMassDistances!(1 to MAXBINARY%, 1 To npoints&) are the mass distances for material A for all points
'   npoints& is the number of points
'   tKratios3!(1 to MAXBINARY%, 1 to MAXBINARY%)  are the interpolated k-ratios for the specified mass distance

ierror = False
On Error GoTo Penepma12BoundaryInterpolateError

Dim mmin As Single
Dim j As Integer, k As Integer
Dim n As Long, nn As Long
Dim krat As Single

Dim npts As Integer, kmax As Integer, nmax As Integer
Dim xdata() As Single, ydata() As Single
Dim acoeff(1 To MAXCOEFF%) As Single

' Find the closest mass distance point
For j% = 1 To MAXBINARY%    ' material A

' Loop on all points passed
mmin! = MAXMINIMUM!
For n& = 1 To nPoints&
If Abs(tDistanceMass! - tMassDistances!(j%, n&)) < mmin! Then
mmin! = Abs(tDistanceMass! - tMassDistances!(j%, n&))
nn& = n&
End If
Next n&

' Load k-ratios for this binary for fitting
For k% = 1 To MAXBINARY%    ' material B

' Now load the xdata array using closest mass distance
npts% = 1
ReDim Preserve xdata(1 To npts%) As Single
ReDim Preserve ydata(1 To npts%) As Single
xdata!(1) = tMassDistances!(j%, nn&)

' Now load the ydata array using k ratio
ydata!(1) = CSng(tKratios2#(k%, j%, nn&))

' Check for pathological conditions
If tDistanceMass! > tMassDistances!(j%, 1) And tDistanceMass! < tMassDistances!(j%, nPoints&) Then

' Now load other points if available
If nn& - 1 > 0 Then
npts% = npts% + 1
ReDim Preserve xdata(1 To npts%) As Single
ReDim Preserve ydata(1 To npts%) As Single
xdata!(npts%) = tMassDistances!(j%, nn& - 1)
ydata!(npts%) = CSng(tKratios2#(k%, j%, nn& - 1))
End If

If nn& + 1 <= nPoints& Then
npts% = npts% + 1
ReDim Preserve xdata(1 To npts%) As Single
ReDim Preserve ydata(1 To npts%) As Single
xdata!(npts%) = tMassDistances!(j%, nn& + 1)
ydata!(npts%) = CSng(tKratios2#(k%, j%, nn& + 1))
End If

' Debug mode
If DebugMode Then
Call IOWriteLog(vbNullString)
For n& = 1 To npts%
Call IOWriteLog("Penepma12BoundaryInterpolate: Point " & Format$(n&) & ", X= " & Format$(xdata!(n&)) & ", Y= " & Format$(ydata!(n&)))
Next n&
End If

' Now fit the data depending on the number of points found
kmax% = 2
If npts% < 3 Then kmax% = 1   ' linear fit or parabolic fit
nmax% = npts%
Call LeastSquares(kmax%, nmax%, xdata!(), ydata!(), acoeff!())
If ierror Then Exit Sub

' Now interpolate to get the actual k-ratio % for the specified distance
krat! = acoeff!(1) + tDistanceMass! * acoeff!(2) + tDistanceMass! ^ 2 * acoeff!(3)

' Distance is outside k-ratio data range (just use end value)
Else
If tDistanceMass! < tMassDistances!(j%, 1) Then krat! = CSng(tKratios2#(k%, j%, 1))
If tDistanceMass! > tMassDistances!(j%, nPoints&) Then krat! = CSng(tKratios2#(k%, j%, nPoints&))
End If

If DebugMode Then Call IOWriteLog("Penepma12BoundaryInterpolate: Interpolated K-ratio % is " & MiscAutoFormat$(krat!) & " at a mass distance " & Format$(tDistanceMass!) & " ug/cm^2")

' Load the interpolated k-ratios for the specified mass distance
tKratios3!(k%, j%) = krat!

Next k%
Next j%

Exit Sub

' Errors
Penepma12BoundaryInterpolateError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12BoundaryInterpolate"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub
