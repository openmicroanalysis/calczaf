Attribute VB_Name = "CodePENEPMA12B"
' (c) Copyright 1995-2026 by John J. Donovan
Option Explicit

' Matrix database arrays
Dim tKratios(1 To MAXBINARY%) As Double

' Boundary database read arrays
Dim tKratios2() As Double       ' 1 To MAXBINARY%, 1 To MAXBINARY%, 1 To npoints&

Dim tLinearDistances() As Single  ' 1 to npoints&
Dim tMassDistances() As Single    ' 1 To MAXBINARY%, 1 to npoints&

Dim tMaterialDensitiesA(1 To MAXBINARY%) As Single
Dim tMaterialDensitiesB(1 To MAXBINARY%) As Single

' Fitting arrays (for mass distance interpolation)
Dim tKratios3() As Single       ' 1 To MAXBINARY%, 1 To MAXBINARY%

' Matrix test parameters
Dim EmitterTakeOff As Single
Dim EmitterKilovolts As Single
Dim EmitterElement As Integer
Dim EmitterXray As Integer
Dim MatrixElement As Integer

' Boundary test parameters
Dim EmitterTakeOff2 As Single
Dim EmitterKilovolts2 As Single
Dim EmitterElement2 As Integer
Dim EmitterXray2 As Integer

Dim MatrixElementA1 As Integer
Dim MatrixElementA2 As Integer
Dim BoundaryElementB1 As Integer
Dim BoundaryElementB2 As Integer

Dim MatrixConcA1 As Single
Dim MatrixConcA2 As Single
Dim BoundaryConcB1 As Single
Dim BoundaryConcB2 As Single

Dim DistanceMode As Integer
Dim DistanceMicrons As Single
Dim DistanceMass As Single

Dim DensityA As Single
Dim DensityB As Single

Sub Penepma12RandomLoad()
' Load the Penfluor and Fanal instance options

ierror = False
On Error GoTo Penepma12RandomLoadError

Dim i As Integer

' Load matrix database test parameters
If EmitterTakeOff! = 0# Then EmitterTakeOff! = DefaultTakeOff!
FormPenepma12Random.TextBeamTakeoff.Text = EmitterTakeOff!
If EmitterKilovolts! = 0# Then EmitterKilovolts! = DefaultKiloVolts!
FormPenepma12Random.TextBeamEnergy.Text = EmitterKilovolts!

FormPenepma12Random.ComboEmitterElement.Clear
For i% = 0 To MAXELM% - 1
FormPenepma12Random.ComboEmitterElement.AddItem Symup$(i% + 1)
Next i%
If EmitterElement% = 0 Then EmitterElement% = 12    ' Mg
FormPenepma12Random.ComboEmitterElement.ListIndex = EmitterElement% - 1

FormPenepma12Random.ComboEmitterXRay.Clear
For i% = 0 To MAXRAY% - 2
FormPenepma12Random.ComboEmitterXRay.AddItem Xraylo$(i% + 1)
Next i%
If EmitterXray% = 0 Then EmitterXray% = 1       ' Ka
FormPenepma12Random.ComboEmitterXRay.ListIndex = EmitterXray% - 1

FormPenepma12Random.ComboMatrixElement.Clear
For i% = 0 To MAXELM% - 1
FormPenepma12Random.ComboMatrixElement.AddItem Symup$(i% + 1)
Next i%
If MatrixElement% = 0 Then MatrixElement% = 26    ' Fe
FormPenepma12Random.ComboMatrixElement.ListIndex = MatrixElement% - 1

' Load boundary database test parameters
If EmitterTakeOff2! = 0# Then EmitterTakeOff2! = DefaultTakeOff!
FormPenepma12Random.TextBeamTakeoff2.Text = EmitterTakeOff2!
If EmitterKilovolts2! = 0# Then EmitterKilovolts2! = DefaultKiloVolts!
FormPenepma12Random.TextBeamEnergy2.Text = EmitterKilovolts2!

FormPenepma12Random.ComboEmitterElement2.Clear
For i% = 0 To MAXELM% - 1
FormPenepma12Random.ComboEmitterElement2.AddItem Symup$(i% + 1)
Next i%
If EmitterElement2% = 0 Then EmitterElement2% = 26    ' Fe
FormPenepma12Random.ComboEmitterElement2.ListIndex = EmitterElement2% - 1

FormPenepma12Random.ComboEmitterXRay2.Clear
For i% = 0 To MAXRAY% - 2
FormPenepma12Random.ComboEmitterXRay2.AddItem Xraylo$(i% + 1)
Next i%
If EmitterXray2% = 0 Then EmitterXray2% = 1       ' Ka
FormPenepma12Random.ComboEmitterXRay2.ListIndex = EmitterXray2% - 1

FormPenepma12Random.ComboMatrixA1.Clear
For i% = 0 To MAXELM% - 1
FormPenepma12Random.ComboMatrixA1.AddItem Symup$(i% + 1)
Next i%
If MatrixElementA1% = 0 Then MatrixElementA1% = 26    ' Fe
FormPenepma12Random.ComboMatrixA1.ListIndex = MatrixElementA1% - 1

FormPenepma12Random.ComboMatrixA2.Clear
For i% = 0 To MAXELM% - 1
FormPenepma12Random.ComboMatrixA2.AddItem Symup$(i% + 1)
Next i%
If MatrixElementA2% = 0 Then MatrixElementA2% = 28    ' Ni
FormPenepma12Random.ComboMatrixA2.ListIndex = MatrixElementA2% - 1

FormPenepma12Random.ComboBoundaryB1.Clear
For i% = 0 To MAXELM% - 1
FormPenepma12Random.ComboBoundaryB1.AddItem Symup$(i% + 1)
Next i%
If BoundaryElementB1% = 0 Then BoundaryElementB1% = 28    ' Ni
FormPenepma12Random.ComboBoundaryB1.ListIndex = BoundaryElementB1% - 1

FormPenepma12Random.ComboBoundaryB2.Clear
For i% = 0 To MAXELM% - 1
FormPenepma12Random.ComboBoundaryB2.AddItem Symup$(i% + 1)
Next i%
If BoundaryElementB2% = 0 Then BoundaryElementB2% = 26    ' Fe
FormPenepma12Random.ComboBoundaryB2.ListIndex = BoundaryElementB2% - 1

' Load defaults
If MatrixConcA1! = 0# Then MatrixConcA1! = 1#
If MatrixConcA2! = 0# Then MatrixConcA2! = 99#
If BoundaryConcB1! = 0# Then BoundaryConcB1! = 1#
If BoundaryConcB2! = 0# Then BoundaryConcB2! = 99#
FormPenepma12Random.TextMatrixA1.Text = MatrixConcA1!
FormPenepma12Random.TextMatrixA2.Text = MatrixConcA2!
FormPenepma12Random.TextBoundaryB1.Text = BoundaryConcB1!
FormPenepma12Random.TextBoundaryB2.Text = BoundaryConcB2!

' Load distances
If DistanceMode% = 0 Then
FormPenepma12Random.OptionDistance(0).Value = True
Else
FormPenepma12Random.OptionDistance(1).Value = True
End If

If DistanceMicrons! = 0# Then DistanceMicrons! = 1.09951
FormPenepma12Random.TextDistanceMicrons.Text = Format$(DistanceMicrons!)
If DistanceMass! = 0# Then DistanceMass! = 865.3034
FormPenepma12Random.TextDistanceMass.Text = Format$(DistanceMass!)

' Load incident and boundary densities
If DensityA! = 0# Then DensityA! = 8.96 ' Cu
FormPenepma12Random.TextDensityA.Text = DensityA!
If DensityB! = 0# Then DensityB! = 8.9  ' Co
FormPenepma12Random.TextDensityB.Text = DensityB!

Exit Sub

' Errors
Penepma12RandomLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12RandomLoad"
ierror = True
Exit Sub

End Sub

Sub Penepma12RandomSave()
' Save the Penfluor and Fanal instance options

ierror = False
On Error GoTo Penepma12RandomSaveError

Dim ip As Integer
Dim esym As String, xsym As String, msym As String

' Save matrix takeoff and beam energy
If Val(FormPenepma12Random.TextBeamTakeoff.Text) < MINTAKEOFF! Or Val(FormPenepma12Random.TextBeamTakeoff.Text) > MAXTAKEOFF! Then
msg$ = "Matrix Takeoff Angle is out of range (must be between " & Format$(MINTAKEOFF!) & " and " & Format$(MAXTAKEOFF!) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RandomSave"
ierror = True
Exit Sub
Else
EmitterTakeOff! = Val(FormPenepma12Random.TextBeamTakeoff.Text)
End If

If Val(FormPenepma12Random.TextBeamEnergy.Text) < MINKILOVOLTS! Or Val(FormPenepma12Random.TextBeamEnergy.Text) > MAXKILOVOLTS! Then
msg$ = "Matrix Beam Energy is out of range (must be between " & Format$(MINKILOVOLTS!) & " and " & Format$(MAXKILOVOLTS!) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RandomSave"
ierror = True
Exit Sub
Else
EmitterKilovolts! = Val(FormPenepma12Random.TextBeamEnergy.Text)
End If

' Save element std element and x-ray
esym$ = FormPenepma12Random.ComboEmitterElement.Text
ip% = IPOS1(MAXELM%, esym$, Symlo$())
If ip% = 0 Then GoTo Penepma12RandomSaveBadElement
EmitterElement% = ip%

' Check for a valid x-ray symbol
xsym$ = FormPenepma12Random.ComboEmitterXRay.Text
ip% = IPOS1(MAXRAY% - 1, xsym$, Xraylo$())
If ip% = 0 Then GoTo Penepma12RandomSaveBadXray
EmitterXray% = ip%

' Check for a valid x-ray symbol
msym$ = FormPenepma12Random.ComboMatrixElement.Text
ip% = IPOS1(MAXELM%, msym$, Symlo$())
If ip% = 0 Then GoTo Penepma12RandomSaveBadMatrix
MatrixElement% = ip%

' Save boundary takeoff and beam energy
If Val(FormPenepma12Random.TextBeamTakeoff2.Text) < MINTAKEOFF! Or Val(FormPenepma12Random.TextBeamTakeoff2.Text) > MAXTAKEOFF! Then
msg$ = "Boundary Takeoff Angle is out of range (must be between " & Format$(MINTAKEOFF!) & " and " & Format$(MAXTAKEOFF!) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RandomSave"
ierror = True
Exit Sub
Else
EmitterTakeOff2! = Val(FormPenepma12Random.TextBeamTakeoff2.Text)
End If

If Val(FormPenepma12Random.TextBeamEnergy2.Text) < MINKILOVOLTS! Or Val(FormPenepma12Random.TextBeamEnergy2.Text) > MAXKILOVOLTS! Then
msg$ = "Boundary Beam Energy is out of range (must be between " & Format$(MINKILOVOLTS!) & " and " & Format$(MAXKILOVOLTS!) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RandomSave"
ierror = True
Exit Sub
Else
EmitterKilovolts2! = Val(FormPenepma12Random.TextBeamEnergy2.Text)
End If

' Save element std element and x-ray
esym$ = FormPenepma12Random.ComboEmitterElement2.Text
ip% = IPOS1(MAXELM%, esym$, Symlo$())
If ip% = 0 Then GoTo Penepma12RandomSaveBadElement2
EmitterElement2% = ip%

' Check for a valid x-ray symbol
xsym$ = FormPenepma12Random.ComboEmitterXRay2.Text
ip% = IPOS1(MAXRAY% - 1, xsym$, Xraylo$())
If ip% = 0 Then GoTo Penepma12RandomSaveBadXray2
EmitterXray2% = ip%

' Check for a valid matrix and boundary binaries
msym$ = FormPenepma12Random.ComboMatrixA1.Text
ip% = IPOS1(MAXELM%, msym$, Symlo$())
If ip% = 0 Then GoTo Penepma12RandomSaveBadMatrix2
MatrixElementA1% = ip%

msym$ = FormPenepma12Random.ComboMatrixA2.Text
ip% = IPOS1(MAXELM%, msym$, Symlo$())
If ip% = 0 Then GoTo Penepma12RandomSaveBadMatrix2
MatrixElementA2% = ip%

msym$ = FormPenepma12Random.ComboBoundaryB1.Text
ip% = IPOS1(MAXELM%, msym$, Symlo$())
If ip% = 0 Then GoTo Penepma12RandomSaveBadBoundary
BoundaryElementB1% = ip%

msym$ = FormPenepma12Random.ComboBoundaryB2.Text
ip% = IPOS1(MAXELM%, msym$, Symlo$())
If ip% = 0 Then GoTo Penepma12RandomSaveBadBoundary
BoundaryElementB2% = ip%

' Load concentratiosn for interpolation
If Val(FormPenepma12Random.TextMatrixA1.Text) < 0# Or Val(FormPenepma12Random.TextMatrixA1.Text) > 100# Then
msg$ = "Matrix element concentration is out of range (must be between 0 and 100)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RandomSave"
ierror = True
Exit Sub
Else
MatrixConcA1! = Val(FormPenepma12Random.TextMatrixA1.Text)
End If

If Val(FormPenepma12Random.TextMatrixA2.Text) < 0# Or Val(FormPenepma12Random.TextMatrixA2.Text) > 100# Then
msg$ = "Matrix element concentration is out of range (must be between 0 and 100)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RandomSave"
ierror = True
Exit Sub
Else
MatrixConcA2! = Val(FormPenepma12Random.TextMatrixA2.Text)
End If

If Val(FormPenepma12Random.TextBoundaryB1.Text) < 0# Or Val(FormPenepma12Random.TextBoundaryB1.Text) > 100# Then
msg$ = "Boundary element concentration is out of range (must be between 0 and 100)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RandomSave"
ierror = True
Exit Sub
Else
BoundaryConcB1! = Val(FormPenepma12Random.TextBoundaryB1.Text)
End If

If Val(FormPenepma12Random.TextBoundaryB2.Text) < 0# Or Val(FormPenepma12Random.TextBoundaryB2.Text) > 100# Then
msg$ = "Boundary element concentration is out of range (must be between 0 and 100)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RandomSave"
ierror = True
Exit Sub
Else
BoundaryConcB2! = Val(FormPenepma12Random.TextBoundaryB2.Text)
End If

' Save distances
If FormPenepma12Random.OptionDistance(0).Value = True Then
DistanceMode% = 0
Else
DistanceMode% = 1
End If

If Val(FormPenepma12Random.TextDistanceMicrons.Text) < 0.00001 Or Val(FormPenepma12Random.TextDistanceMicrons.Text) > 1000# Then
msg$ = "Boundary distance (um) is out of range (must be between 0.00001 and 1000.0)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RandomSave"
ierror = True
Exit Sub
Else
DistanceMicrons! = Val(FormPenepma12Random.TextDistanceMicrons.Text)
End If

If Val(FormPenepma12Random.TextDistanceMass.Text) < 0.001 Or Val(FormPenepma12Random.TextDistanceMass.Text) > 100000# Then
msg$ = "Boundary distance (ug/cm2) is out of range (must be between 0.001 and 100000.0)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RandomSave"
ierror = True
Exit Sub
Else
DistanceMass! = Val(FormPenepma12Random.TextDistanceMass.Text)
End If

' Save incident and boundary densities
If Val(FormPenepma12Random.TextDensityA.Text) <= 0# Or Val(FormPenepma12Random.TextDensityA.Text) > MAXDENSITY# Then
msg$ = "Incident material density is out of range (must be greater than 0 and less than " & Format$(MAXDENSITY#) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RandomSave"
ierror = True
Exit Sub
Else
DensityA! = Val(FormPenepma12Random.TextDensityA.Text)
End If

If Val(FormPenepma12Random.TextDensityB.Text) <= 0# Or Val(FormPenepma12Random.TextDensityB.Text) > MAXDENSITY# Then
msg$ = "Boundary material density is out of range (must be greater than 0 and less than " & Format$(MAXDENSITY#) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RandomSave"
ierror = True
Exit Sub
Else
DensityB! = Val(FormPenepma12Random.TextDensityB.Text)
End If

Exit Sub

' Errors
Penepma12RandomSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12RandomSave"
ierror = True
Exit Sub

Penepma12RandomSaveBadElement:
msg$ = "The specified matrix emitter element " & esym$ & ", is invalid"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RandomSave"
ierror = True
Exit Sub

Penepma12RandomSaveBadXray:
msg$ = "The specified matrix x-ray " & esym$ & ", is invalid"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RandomSave"
ierror = True
Exit Sub

Penepma12RandomSaveBadMatrix:
msg$ = "The specified matrix matrix element " & esym$ & ", is invalid"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RandomSave"
ierror = True
Exit Sub

Penepma12RandomSaveBadElement2:
msg$ = "The specified boundary emitter element " & esym$ & ", is invalid"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RandomSave"
ierror = True
Exit Sub

Penepma12RandomSaveBadXray2:
msg$ = "The specified boundary x-ray " & esym$ & ", is invalid"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RandomSave"
ierror = True
Exit Sub

Penepma12RandomSaveBadMatrix2:
msg$ = "The specified boundary matrix element " & esym$ & ", is invalid"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RandomSave"
ierror = True
Exit Sub

Penepma12RandomSaveBadBoundary:
msg$ = "The specified boundary element " & esym$ & ", is invalid"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RandomSave"
ierror = True
Exit Sub

End Sub

Sub Penepma12RandomReadMatrix()
' Read the matrix database for the specified energy, emitter, xray and matrix element (for testing purposes)

ierror = False
On Error GoTo Penepma12RandomReadMatrixError

Dim notfound As Boolean
Dim i As Integer
Dim astring As String

' Get the specified data
Call Penepma12MatrixReadMDB2(EmitterTakeOff!, EmitterKilovolts!, EmitterElement%, EmitterXray%, MatrixElement%, tKratios#(), notfound)
If ierror Then Exit Sub

' Check if found
If notfound Then
FormPenepma12Random.LabelMatrixDisplay.Caption = "Matrix values were not found"
Exit Sub
End If

' Display to user
astring$ = Symup$(EmitterElement%) & " " & Xraylo$(EmitterXray%) & " in " & Symup$(MatrixElement%)
FormPenepma12Random.LabelMatrixDisplay.Caption = "Matrix values for " & astring$ & " were found and output to the log window"

' Output to log
Call IOWriteLog$(vbCrLf & astring$ & " at " & Format$(EmitterTakeOff!) & " degrees and " & Format$(EmitterKilovolts!) & " keV")
astring$ = Format$(vbTab & "Conc%", a08$) & vbTab & Format$("Kratio%", a08$)
Call IOWriteLog$(astring$)

For i% = 1 To MAXBINARY%
astring$ = vbTab & MiscAutoFormat$(BinaryRanges!(i%)) & vbTab & MiscAutoFormatD$(tKratios#(i%))
Call IOWriteLog$(astring$)
Next i%

Exit Sub

' Errors
Penepma12RandomReadMatrixError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12RandomReadMatrix"
ierror = True
Exit Sub

End Sub

Sub Penepma12RandomReadBoundary()
' Read the boundary MDB database for the specified energy, emitter, xray and matrix element (for testing purposes)

ierror = False
On Error GoTo Penepma12RandomReadBoundaryError

Dim notfound As Boolean
Dim j As Integer, k As Integer
Dim astring As String, bstring As String, jstring As String, kstring As String

Dim n As Long, nPoints As Long

bstring$ = Symup$(EmitterElement2%) & " " & Xraylo$(EmitterXray2%) & " in "
bstring$ = bstring$ & Trim$(Symup$(MatrixElementA1%)) & "-" & Trim$(Symup$(MatrixElementA2%)) & " adjacent to "
bstring$ = bstring$ & Trim$(Symup$(BoundaryElementB1%)) & "-" & Trim$(Symup$(BoundaryElementB2%))

' Get the specified data
Call Penepma12BoundaryReadMDB(EmitterTakeOff2!, EmitterKilovolts2!, EmitterElement2%, EmitterXray2%, MatrixElementA1%, MatrixElementA2%, BoundaryElementB1%, BoundaryElementB2%, tKratios2#(), tLinearDistances!(), tMassDistances!(), tMaterialDensitiesA!(), tMaterialDensitiesB!(), nPoints&, notfound)
If ierror Then Exit Sub

' Check if found
If notfound Then
FormPenepma12Random.LabelBoundaryDisplay.Caption = "Boundary values were not found for " & bstring$
Exit Sub
End If

' Display to user
FormPenepma12Random.LabelBoundaryDisplay.Caption = "Boundary values for " & bstring$ & " were found and output to log"

' Loop on all distances (only print out specified linear or mass distances?)
Call IOWriteLog$(vbNullString)
For n& = 1 To nPoints&

' Check distance mode and specified distances (0 = linear distance, 1 = mass distance)
If DistanceMode = 0 And DistanceMicrons! = tLinearDistances!(n&) Or DistanceMode% = 1 Then

' Output to log
Call IOWriteLog$(vbCrLf & bstring$ & " at " & Format$(EmitterTakeOff2!) & " degrees and " & Format$(EmitterKilovolts2!) & " keV, at linear distance of " & Format$(tLinearDistances!(n&)) & " um")

' Create column labels (note that these labels are always unswapped)
astring$ = Format$("ConcA%", a80$) & Format$("ConcB%", a80$) & Format$("Kratio%", a80$) & Format$("DensA", a80$) & Format$("DensB", a80$) & Format$("MDistA", a80$)
Call IOWriteLog$(astring$)

For j% = 1 To MAXBINARY%    ' material A
'If DistanceMode% = 1 And DistanceMass! = tMassDistances!(j%, n&) Then
For k% = 1 To MAXBINARY%    ' material B

jstring$ = Format$(BinaryRanges!(j%)) & "-" & Format$(100# - BinaryRanges!(j%))
kstring$ = Format$(BinaryRanges!(k%)) & "-" & Format$(100# - BinaryRanges!(k%))

astring$ = Format$(Format$(tKratios2#(k%, j%, n&), f84$), a80$) & Format$(Format$(tMaterialDensitiesA!(j%), f83$), a80$) & Format$(Format$(tMaterialDensitiesB!(k%), f83$), a80$) & Format$(Format$(tMassDistances!(j%, n&), f82$), a80$)
astring$ = Format$(jstring$, a80$) & Format$(kstring$, a80$) & astring$
Call IOWriteLog$(astring$)

Next k%
'End If
Next j%

End If
Next n&

Exit Sub

' Errors
Penepma12RandomReadBoundaryError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12RandomReadBoundary"
ierror = True
Exit Sub

End Sub

Sub Penepma12RandomBoundaryInterpolate()
' Obtain k-ratios for the specified emitter, beam incident and boundary material at the specified mass distance
'   tKratios2() As Double       ' 1 To MAXBINARY%, 1 To MAXBINARY%, 1 To npoints&

ierror = False
On Error GoTo Penepma12RandomBoundaryInterpolateError

Dim notfound As Boolean
Dim j As Integer, k As Integer
Dim astring As String, bstring As String, jstring As String, kstring As String

Dim nPoints As Long

bstring$ = Symup$(EmitterElement2%) & " " & Xraylo$(EmitterXray2%) & " in "
bstring$ = bstring$ & Trim$(Symup$(MatrixElementA1%)) & "-" & Trim$(Symup$(MatrixElementA2%)) & " adjacent to "
bstring$ = bstring$ & Trim$(Symup$(BoundaryElementB1%)) & "-" & Trim$(Symup$(BoundaryElementB2%))

' Get the specified data for all mass distances
Call Penepma12BoundaryReadMDB2(EmitterTakeOff2!, EmitterKilovolts2!, EmitterElement2%, EmitterXray2%, MatrixElementA1%, MatrixElementA2%, BoundaryElementB1%, BoundaryElementB2%, tKratios2#(), tMassDistances!(), nPoints&, notfound)
If ierror Then Exit Sub

' Check if found
If notfound Then
FormPenepma12Random.LabelBoundaryDisplay.Caption = "Boundary values were not found for " & bstring$
Exit Sub
End If

' Now fit values by first finding the nearest neighbor for the specified mass distance
ReDim tKratios3(1 To MAXBINARY%, 1 To MAXBINARY%) As Single
Call Penepma12BoundaryInterpolate(DistanceMass!, tKratios2#(), tMassDistances!(), nPoints&, tKratios3!())
If ierror Then Exit Sub

' Display to user
FormPenepma12Random.LabelBoundaryDisplay.Caption = "Boundary values for " & bstring$ & " at a mass distance of " & Format$(DistanceMass!) & " ug/cm^2 were found and output to log"

Call IOWriteLog(vbCrLf & "K-ratios for " & bstring$ & " at a mass distance of " & Format$(DistanceMass!) & " ug/cm^2")

' Create column labels (note that these labels are always unswapped)
astring$ = Format$("ConcA", a80$) & Format$("ConcB", a80$) & Format$("Kratios", a80$)
astring$ = astring$ & Format$("KratA", a80$)         ' bulk calculation
astring$ = astring$ & Format$("AB-A", a80$)      ' boundary minus bulk
Call IOWriteLog$(astring$)

For j% = 1 To MAXBINARY%    ' material A
For k% = 1 To MAXBINARY%    ' material B

jstring$ = Format$(BinaryRanges!(j%)) & "-" & Format$(100# - BinaryRanges!(j%))
kstring$ = Format$(BinaryRanges!(k%)) & "-" & Format$(100# - BinaryRanges!(k%))


astring$ = Format$(jstring$, a80$) & Format$(kstring$, a80$)
astring$ = astring$ & Format$(Format$(tKratios3!(k%, j%), f84$), a80$)
astring$ = astring$ & Format$(Format$(tKratios2#(k%, j%, nPoints&), f84$), a80$) & vbTab     ' k-ratio of material A at extreme distance
astring$ = astring$ & Format$(Format$(tKratios3!(k%, j%) - tKratios2#(k%, j%, nPoints&), f106$), a80$) & vbTab     ' k-ratio of material AB - A
Call IOWriteLog$(astring$)

Next k%
Next j%

' Output to file
Call Penepma12RandomBoundaryOutput(DistanceMass!, tKratios2#(), nPoints&, tKratios3!())
If ierror Then Exit Sub

Exit Sub

' Errors
Penepma12RandomBoundaryInterpolateError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12RandomBoundaryInterpolate"
ierror = True
Exit Sub

End Sub

Sub Penepma12RandomBoundaryOutput(tDistanceMass As Single, tKratios2() As Double, nPoints As Long, tKratios3() As Single)
' Output specified boundary k-ratios for the specified mass distance

ierror = False
On Error GoTo Penepma12RandomBoundaryOutputError

Dim j As Integer, k As Integer
Dim astring As String, jstring As String, kstring As String

Dim tfolder As String, tfilename As String

' Output the data as desired
Close #Temp1FileNumber%

' Specify folder
tfolder$ = PENEPMA_Root$ & "\Fanal\boundary"

' Load output filename based on distance
tfilename$ = tfolder$ & "\" & Trim$(Symup$(MatrixElementA1%)) & "-" & Trim$(Symup$(MatrixElementA2%)) & "_" & Trim$(Symup$(BoundaryElementB1%)) & "-" & Trim$(Symup$(BoundaryElementB2%)) & "_" & Format$(EmitterTakeOff2!) & "_" & Format$(EmitterKilovolts2!) & "_" & Trim$(Symup$(EmitterElement2%)) & " " & Xraylo$(EmitterXray2%) & "_" & Format$(tDistanceMass!, "Fixed") & "ug-cm2" & ".dat"
Open tfilename$ For Output As #Temp1FileNumber%

' Output column labels
astring$ = VbDquote$ & "keV" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "ConcA1" & VbDquote$ & vbTab & VbDquote$ & "ConcA2" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "ConcB1" & VbDquote$ & vbTab & VbDquote$ & "ConcB2" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "KratioAB%" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "KratioA%" & VbDquote$ & vbTab        ' bulk calculation
astring$ = astring$ & VbDquote$ & "KratioAB-A%" & VbDquote$ & vbTab     ' boundary minus bulk
Print #Temp1FileNumber%, astring

' Output data for this energy by concentration
For j% = 1 To MAXBINARY%    ' material A
For k% = 1 To MAXBINARY%    ' material B

jstring$ = Format$(BinaryRanges!(j%)) & vbTab & Format$(100# - BinaryRanges!(j%))
kstring$ = Format$(BinaryRanges!(k%)) & vbTab & Format$(100# - BinaryRanges!(k%))

' Output calculations
astring$ = CSng(EmitterKilovolts2!) & vbTab & jstring$ & vbTab & kstring$ & vbTab      ' composition of A and B
astring$ = astring$ & Format$(tKratios3!(k%, j%)) & vbTab     ' k-ratio of AB
astring$ = astring$ & Format$(tKratios2#(k%, j%, nPoints&)) & vbTab     ' k-ratio of material A at extreme distance
astring$ = astring$ & Format$(tKratios3!(k%, j%) - tKratios2#(k%, j%, nPoints&), f127$) & vbTab     ' k-ratio of material AB - A

' Output string
Print #Temp1FileNumber%, astring$

Next k%
Next j%

Close #Temp1FileNumber%

msg$ = "The specified boundary k-ratio plot data was output based on " & Format$(EmitterKilovolts2!) & " keV, " & Trim$(Symup$(EmitterElement2%)) & " " & Trim$(Xraylo$(EmitterXray2%)) & " in " & Trim$(Symup$(MatrixElementA1%)) & "-" & Trim$(Symup$(MatrixElementA2%)) & " adjacent to " & Trim$(Symup$(BoundaryElementB1%)) & "-" & Trim$(Symup$(BoundaryElementB2%)) & " at a mass distance of " & Format$(tDistanceMass!) & " ug/cm^2"
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12RandomBoundaryOutput"

Exit Sub

' Errors
Penepma12RandomBoundaryOutputError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12RandomBoundaryOutput"
ierror = True
Exit Sub

End Sub
