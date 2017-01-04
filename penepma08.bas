Attribute VB_Name = "CodePENEPMA08"
' (c) Copyright 1995-2017 by John J. Donovan
Option Explicit

Global Const MAXMATOUTPUT% = 2      ' up to 2 materials (1 for bulk, 2 for couple, bilayer or thin film)
Const MAXDENSITY# = 50#
Const PENEPMA_DISPLAY_SEC# = 20     ' 20 seconds refresh display time
Const MAXTRIES% = 10

Const COL7% = 7

'The maximum number of channels can be changed in Penepma.f where (NEDCM) can be changed at line 464:
'PARAMETER (NEDM=25,NEDCM=1000)     ' changed to 32000 12-02-2016
'Also the output format changed on line 1121 from I4 to I5

' Global
Global Const MAXPRODUCTION% = 4
Global GraphDisplayOption As Integer     ' 0 = total spectrum, 1 = characteristic spectrum, 2  = backscatter spectrum

' Pendbase options
Dim MaterialsSelected(1 To MAXMATOUTPUT%) As Integer
Dim MaterialFiles(1 To MAXMATOUTPUT%) As String     ' filename only, no path
Dim MaterialDensity As Double

' Penepma08 options
Dim PenepmaTaskID As Long

Dim BeamTitle As String
Dim BeamTakeOff As Double, BeamEnergy As Double
Dim BeamPosition(1 To 3) As Double
Dim BeamDirection(1 To 2) As Double, BeamAperture As Double
Dim BeamDumpPeriod As Double

Dim BeamNumberSimulatedShowers As Double, BeamSimulationTimePeriod As Double

Dim BeamMinimumEnergyRange As Double, BeamMaximumEnergyRange As Double
Dim BeamNumberOfEnergyChannels As Long

Dim BeamProductionIndex As Long
Dim BeamProductionFilename(0 To MAXPRODUCTION%) As String

Dim InputEABS1(1 To MAXMATOUTPUT%) As Double    ' electron absorption energy
Dim InputEABS2(1 To MAXMATOUTPUT%) As Double    ' photon absorption energy

Dim InputEABS3(1 To MAXMATOUTPUT%) As Double    ' use file default
Dim InputC1(1 To MAXMATOUTPUT%) As Double    ' use file default
Dim InputC2(1 To MAXMATOUTPUT%) As Double    ' use file default
Dim InputWCC(1 To MAXMATOUTPUT%) As Double    ' use file default
Dim InputWCR(1 To MAXMATOUTPUT%) As Double    ' use file default

Dim InputTheta1 As Double, InputTheta2 As Double, InputPhi1 As Double, InputPhi2 As Double
Dim InputISPF As Long   ' generally always zero

Dim InputFile As String
Dim GeometryFile As String

' Re-load form options
Dim SimulationInProgress As Boolean
Dim UseGridLines As Boolean
Dim UseLogScale As Boolean

Dim PENEPMA_BATCH_FOLDER As String

' Calculated MC spectrum data
Dim nPoints As Long
Dim xdata() As Double, ydata() As Double

Dim InputFiles() As String
Dim InputDates() As Variant

Dim PenepmaTimeStart As Variant

Dim BinaryElement1 As Integer
Dim BinaryElement2 As Integer

Dim ExtractElement As Integer
Dim ExtractXray As Integer
Dim ExtractFolder As String
Dim ExtractStdFolder As String

Dim PureElement1 As Integer
Dim PureElement2 As Integer

Dim DetectorGeometryType As Integer    ' 0 = annular, 1 = north, 2 = east, 3 = south, 4 = west

' Penepma k-ratio calculations
Dim pri_int(1 To 2, 1 To MAXRAY% - 1) As Single
Dim flch_int(1 To 2, 1 To MAXRAY% - 1) As Single
Dim flbr_int(1 To 2, 1 To MAXRAY% - 1) As Single
Dim flu_int(1 To 2, 1 To MAXRAY% - 1) As Single
Dim tot_int(1 To 2, 1 To MAXRAY% - 1) As Single         ' total intensity
Dim tot_int_var(1 To 2, 1 To MAXRAY% - 1) As Single     ' total intensity uncertainty

Dim std_int(1 To 2, 1 To MAXRAY% - 1) As Single
Dim unk_pri_int(1 To 2, 1 To MAXBINARY%, 1 To MAXRAY% - 1) As Single
Dim unk_flu_int(1 To 2, 1 To MAXBINARY%, 1 To MAXRAY% - 1) As Single
Dim unk_tot_int(1 To 2, 1 To MAXBINARY%, 1 To MAXRAY% - 1) As Single

Dim unk_krat(1 To 2, 1 To MAXBINARY%, 1 To MAXRAY% - 1) As Single
Dim unk_afac(1 To 2, 1 To MAXBINARY%, 1 To MAXRAY% - 1) As Single

' Electron-photon ranges
Dim XrayAdjustNumber(1 To MAXMATOUTPUT%) As Integer     ' -1 (decrease) or 1 (increase)

' Plotting
Dim CalcZAF_ZAF_Factors() As Single
Dim Binary_ZAF_Factors() As Single

Dim Binary_ZAF_Coeffs() As Single
Dim CalcZAF_ZAF_Coeffs() As Single

Dim PENEPMASample(1 To 1) As TypeSample

Sub Penepma08SaveInput(tForm As Form)
' Save the form for the Input parameters (*.IN file)

ierror = False
On Error GoTo Penepma08SaveInputError

Static userwarned As Boolean

Dim i As Integer
Dim tfilename As String

icancelauto = False

' Input file title
BeamTitle$ = Trim$(tForm.TextInputTitle.Text)

' Beam takeoff (in degrees)
If Val(tForm.TextBeamTakeoff.Text) < 10# Or Val(tForm.TextBeamTakeoff.Text) > 80# Then
msg$ = "Beam takeoff angle is out of range (must be between 10 and 80 degrees)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveInput"
ierror = True
Exit Sub
Else
BeamTakeOff# = Val(tForm.TextBeamTakeoff.Text)
InputTheta1# = (90# - BeamTakeOff#) - 5
InputTheta2# = (90# - BeamTakeOff#) + 5
End If

' Beam energy (in eV)
If Val(tForm.TextBeamEnergy.Text) < 50# Or Val(tForm.TextBeamEnergy.Text) > MAXKILOVOLTS! * EVPERKEV# Then
msg$ = "Beam energy is out of range (must be between 50 and " & Format$(MAXKILOVOLTS! * EVPERKEV#) & " eV)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveInput"
ierror = True
Exit Sub
Else
BeamEnergy# = Val(tForm.TextBeamEnergy.Text)
End If

' Beam position (in cm)
For i% = 1 To 3
If Val(tForm.TextBeamPosition(i% - 1).Text) < -1 Or Val(tForm.TextBeamPosition(i% - 1).Text) > 1 Then
If i% = 1 Then msg$ = "Beam position X is out of range (must be between -1 and 1 cm)"
If i% = 2 Then msg$ = "Beam position Y is out of range (must be between -1 and 1 cm)"
If i% = 3 Then msg$ = "Beam position Z is out of range (must be between -1 and 1 cm)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveInput"
ierror = True
Exit Sub
Else
BeamPosition#(i%) = Val(tForm.TextBeamPosition(i% - 1).Text)
End If
Next i%

' Beam direction and aperture
If Val(tForm.TextBeamDirection(0).Text) < 0 Or Val(tForm.TextBeamDirection(0).Text) > 180 Then
msg$ = "Beam direction theta is out of range (must be between 0 and 180)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveInput"
ierror = True
Exit Sub
Else
BeamDirection#(1) = Val(tForm.TextBeamDirection(0).Text)
End If

If Val(tForm.TextBeamDirection(1).Text) < 0 Or Val(tForm.TextBeamDirection(1).Text) > 180 Then
msg$ = "Beam direction phi is out of range (must be between 0 and 180)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveInput"
ierror = True
Exit Sub
Else
BeamDirection#(2) = Val(tForm.TextBeamDirection(1).Text)
End If

If Val(tForm.TextBeamAperture.Text) < 0 Or Val(tForm.TextBeamAperture.Text) > 180 Then
msg$ = "Beam aperture is out of range (must be between 0 and 180)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveInput"
ierror = True
Exit Sub
Else
BeamAperture# = Val(tForm.TextBeamAperture.Text)
End If

If Val(tForm.TextDumpPeriod.Text) < 10 Or Val(tForm.TextDumpPeriod.Text) > 100000 Then
msg$ = "Dump Period is out of range (must be between 10 and 100,000 seconds)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveInput"
ierror = True
Exit Sub
Else
BeamDumpPeriod# = Val(tForm.TextDumpPeriod.Text)
End If

If Val(tForm.TextEnergyRangeMinMaxNumber(0).Text) < 0 Or Val(tForm.TextEnergyRangeMinMaxNumber(0).Text) > 1000 Then
msg$ = "Minimum Energy Range is out of range (must be between 0 and 1000)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveInput"
ierror = True
Exit Sub
Else
BeamMinimumEnergyRange# = Val(tForm.TextEnergyRangeMinMaxNumber(0).Text)
End If

If Val(tForm.TextEnergyRangeMinMaxNumber(1).Text) < 10000 Or Val(tForm.TextEnergyRangeMinMaxNumber(1).Text) > 100000 Then
msg$ = "Maximum Energy Range is out of range (must be between 10,000 and 100,000)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveInput"
ierror = True
Exit Sub
Else
BeamMaximumEnergyRange# = Val(tForm.TextEnergyRangeMinMaxNumber(1).Text)
End If

If Val(tForm.TextEnergyRangeMinMaxNumber(2).Text) < 1000 Or Val(tForm.TextEnergyRangeMinMaxNumber(2).Text) > 100000 Then
msg$ = "Number Of Energy Channels is out of range (must be between 1000 and 100,000)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveInput"
ierror = True
Exit Sub
Else
BeamNumberOfEnergyChannels& = Val(tForm.TextEnergyRangeMinMaxNumber(2).Text)
End If

' Check for new Penepma.exe if more than 1000 channels
If BeamNumberOfEnergyChannels& > 1000 Then
If Dir$(PENEPMA_Path$ & "\Penepma.exe") <> vbNullString Then
If FileDateTime(PENEPMA_Path$ & "\Penepma.exe") < CDate("12/01/2016") Then
msg$ = "Number of energy channels requires an update to the Penepma12 distribution. Please use the Help | Update Probe for EPMA or Help | Update CalcZAF menus to update your Penepma12 distribution."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveInput"
ierror = True
Exit Sub
End If
End If
End If

' Check other parameters
If Val(tForm.TextNumberSimulatedShowers.Text) < 100 Or Val(tForm.TextNumberSimulatedShowers.Text) > 1E+20 Then
msg$ = "Number Of Simulated Showers is out of range (must be between 100 and 1E+20)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveInput"
ierror = True
Exit Sub
Else
BeamNumberSimulatedShowers# = Val(tForm.TextNumberSimulatedShowers.Text)
End If
 
If Val(tForm.TextSimulationTimePeriod.Text) < 10 Or Val(tForm.TextSimulationTimePeriod.Text) > 10000000000# Then
msg$ = "Simulation Time Period is out of range (must be between 10 and 1E+10)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveInput"
ierror = True
Exit Sub
Else
BeamSimulationTimePeriod# = Val(tForm.TextSimulationTimePeriod.Text)
End If

' Save optimize production index
For i% = 0 To MAXPRODUCTION%
If tForm.OptionProduction(i%).Value = True Then BeamProductionIndex& = i%
Next i%

' Check for valid materials files (must have one file at least)
If Trim$(tForm.TextMaterialFiles(0).Text) = vbNullString Then GoTo Penepma08SaveInputNoMaterialFile
If Dir$(PENEPMA_Path$ & "\" & tForm.TextMaterialFiles(0).Text) = vbNullString Then GoTo Penepma08SaveInputMaterialFile2NotFound
MaterialFiles$(1) = Trim$(tForm.TextMaterialFiles(0).Text)

' Check for additional material file
MaterialFiles$(MAXMATOUTPUT%) = vbNullString
If Trim$(tForm.TextMaterialFiles(MAXMATOUTPUT% - 1).Text) <> vbNullString Then
If Dir$(PENEPMA_Path$ & "\" & tForm.TextMaterialFiles(MAXMATOUTPUT% - 1).Text) = vbNullString Then GoTo Penepma08SaveInputMaterialFile2NotFound2
MaterialFiles$(MAXMATOUTPUT%) = Trim$(tForm.TextMaterialFiles(MAXMATOUTPUT% - 1).Text)
End If

For i% = 1 To MAXMATOUTPUT%
InputEABS1#(i%) = Val(tForm.TextEABS1(i% - 1).Text)
InputEABS2#(i%) = Val(tForm.TextEABS2(i% - 1).Text)
Next i%

' Check detector geometry type (0 = annular, 1 = north, 2 = east, 3 = south, 4 = west)
For i% = 0 To 4
If tForm.OptionDetectorGeometry(i%).Value = True Then DetectorGeometryType% = i%
Next i%

' Load theta values for detector type (0 = annular, 1 = north, 2 = east, 3 = south, 4 = west)
If DetectorGeometryType% = 0 Then
InputPhi1# = 0#
InputPhi2# = 360#
ElseIf DetectorGeometryType% = 1 Then
InputPhi1# = -20#
InputPhi2# = 20#
ElseIf DetectorGeometryType% = 2 Then
InputPhi1# = 70#
InputPhi2# = 110#
ElseIf DetectorGeometryType% = 3 Then
InputPhi1# = 160#
InputPhi2# = 200#
ElseIf DetectorGeometryType% = 4 Then
InputPhi1# = 250#
InputPhi2# = 290#
End If

' Check geometry file
GeometryFile$ = Trim$(tForm.TextGeometryFile.Text)
If Trim$(GeometryFile$) = vbNullString Then GoTo Penepma08GeometryFileBlank

' If geo file is not in Penepma folder, copy it there
If Dir$(PENEPMA_Path$ & "\" & GeometryFile$) = vbNullString Then
FileCopy PENEPMA_Root$ & "\" & GeometryFile$, PENEPMA_Path$ & "\" & GeometryFile$
End If
If Dir$(PENEPMA_Path$ & "\" & GeometryFile$) = vbNullString Then GoTo Penepma08GeometryFileNotFound

' Save input file name with .in extension
tfilename$ = Trim$(tForm.TextInputFile.Text)
Call MiscModifyStringToFilename$(tfilename$)
tfilename$ = MiscGetFileNameNoExtension$(tfilename$) & ".in"
tForm.TextInputFile.Text = tfilename$
InputFile$ = tfilename$

' Check that the sample production files exist
For i% = 0 To MAXPRODUCTION%
If Dir$(Trim$(BeamProductionFilename$(i%))) = vbNullString Then GoTo Penepma08SaveInputProductionNotFound
Next i%

' Check that secondary fluorescence production file is specified if additional material files are specified
If MaterialFiles$(2) <> vbNullString And BeamProductionIndex& < 3 Then
msg$ = "Note that additional material files are specified but the Optimize Secondary Fluorescence Production or Thin Film On Substrate Production option was not selected"
MsgBox msg$, vbOKOnly + vbInformation, "Penepma08SaveInput"
End If

If MaterialFiles$(1) = vbNullString Or MaterialFiles$(2) = vbNullString And BeamProductionIndex& >= 3 Then
msg$ = "Note that two material files are required for the Optimize Secondary Fluorescence Production or Thin Film On Substrate Production options"
MsgBox msg$, vbOKOnly + vbInformation, "Penepma08SaveInput"
End If

' Warn using about maximum step length if using bi-layer model (should be 1/10th of thin film thickness)
If BeamProductionIndex& = 4 And Not userwarned Then
msg$ = "Note that if the thin film thickness is less than a few nm, the DSMAX (maximum step length) parameter in the Penepma input file (" & InputFile$ & "), may need to be manually edited to be less than 1/10th of the film thickness (in cm)." & vbCrLf & vbCrLf
msg$ = msg$ & "This warning will only be given once per application execution session."
MsgBox msg$, vbOKOnly + vbInformation, "Penepma08SaveInput"
userwarned = True
End If

Exit Sub

' Errors
Penepma08SaveInputError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08SaveInput"
ierror = True
Exit Sub

Penepma08SaveInputNoMaterialFile:
msg$ = "Material file is blank (must have at least one material file specified)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveInput"
ierror = True
Exit Sub

Penepma08SaveInputMaterialFile2NotFound:
msg$ = "The specified material file (" & PENEPMA_Path$ & "\" & tForm.TextMaterialFiles(0).Text & ") was not found (check that file actually exists)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveInput"
ierror = True
Exit Sub

Penepma08SaveInputMaterialFile2NotFound2:
msg$ = "The specified material file (" & PENEPMA_Path$ & "\" & tForm.TextMaterialFiles(MAXMATOUTPUT% - 1).Text & ") was not found (check that file actually exists)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveInput"
ierror = True
Exit Sub

Penepma08GeometryFileBlank:
msg$ = "Geometry file is blank"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveInput"
ierror = True
Exit Sub

Penepma08GeometryFileNotFound:
msg$ = "Geometry file " & GeometryFile$ & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveInput"
ierror = True
Exit Sub

Penepma08GeometryFileNameTooLong:
msg$ = "Geometry file name " & GeometryFile$ & " is too long (maximum length 20 characters)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveInput"
ierror = True
Exit Sub

Penepma08SaveInputProductionNotFound:
msg$ = "Production Optimization Input File " & BeamProductionFilename$(i%) & " was not found (please contact Probe Software for assistance)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveInput"
ierror = True
Exit Sub

End Sub

Sub Penepma08SaveMaterial(tForm As Form)
' Save the form

ierror = False
On Error GoTo Penepma08SaveMaterialError

Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer

icancelauto = False

If tForm.ListAvailableStandards.ListIndex < 0 Then Exit Sub
If tForm.ListAvailableStandards.ListCount < 1 Then Exit Sub

' Save selected materials
For i% = 1 To MAXMATOUTPUT%
MaterialsSelected%(i%) = 0
Next i%

' Load order
If tForm.CheckMaterialLoadOrder.Value = vbUnchecked Then
m% = 0
n% = tForm.ListAvailableStandards.ListCount - 1
j% = 1
Else
m% = tForm.ListAvailableStandards.ListCount - 1
n% = 0
j% = -1
End If

k% = 0
For i% = m% To n% Step j%
If tForm.ListAvailableStandards.Selected(i%) Then
k% = k% + 1
If k% > MAXMATOUTPUT% Then GoTo Penepma08SaveMaterialTooMany
MaterialsSelected%(k%) = tForm.ListAvailableStandards.ItemData(i%)
MaterialFiles$(k%) = Penepma08ReturnMaterialFile$(MaterialsSelected%(k%))
End If
Next i%

If Val(tForm.TextMaterialDensity.Text) <= 0# Or Val(tForm.TextMaterialDensity.Text) > MAXDENSITY# Then
msg$ = "Material Density is out of range (must be greater than 0 and less than " & Format$(MAXDENSITY#) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveMaterial"
ierror = True
Exit Sub
Else
MaterialDensity# = Val(tForm.TextMaterialDensity.Text)
End If

If Val(tForm.TextMaterialFcb.Text) < 0# Or Val(tForm.TextMaterialFcb.Text) > 100# Then
msg$ = "Oscillator Strength is out of range (must be between 0 and 100)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveMaterial"
ierror = True
Exit Sub
Else
MaterialFcb# = Val(tForm.TextMaterialFcb.Text)
End If

If Val(tForm.TextMaterialWcb.Text) < 0# Or Val(tForm.TextMaterialWcb.Text) > 1000# Then
msg$ = "Oscillator Energy is out of range (must be between 0 and 1000)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveMaterial"
ierror = True
Exit Sub
Else
MaterialWcb# = Val(tForm.TextMaterialWcb.Text)
End If

Exit Sub

' Errors
Penepma08SaveMaterialError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08SaveMaterial"
ierror = True
Exit Sub

Penepma08SaveMaterialTooMany:
msg$ = "Too many materials selected for output"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08SaveMaterial"
ierror = True
Exit Sub

End Sub

Sub Penepma08Init()
' Initialize the variables

ierror = False
On Error GoTo Penepma08InitError

Dim i As Integer

icancelauto = False

' Close temp file handles in case they are already open
Close #Temp1FileNumber%
Close #Temp2FileNumber%

' Pendbase init
If MaterialDensity# = 0# Then MaterialDensity# = 2.7
If MaterialFcb# = 0# Then MaterialFcb# = 0#     ' zero means use calculated default
If MaterialWcb# = 0# Then MaterialWcb# = 0#     ' zero means use calculated default

' Penepma08 init
If BeamTitle$ = vbNullString Then BeamTitle$ = "PENEPMA Input File Title"
If BeamTakeOff# = 0# Then BeamTakeOff# = CDbl(DefaultTakeOff!)
If BeamEnergy# = 0# Then BeamEnergy# = CDbl(DefaultKiloVolts! * EVPERKEV#)
If BeamPosition#(1) = 0# Then BeamPosition#(1) = 0#     ' X
If BeamPosition#(2) = 0# Then BeamPosition#(2) = 0#     ' Y
If BeamPosition#(3) = 0# Then BeamPosition#(3) = 1#      ' Z (1 cm WD)

If BeamDirection#(1) = 0# Then BeamDirection#(1) = 180#
If BeamDirection#(2) = 0# Then BeamDirection#(2) = 0#
If BeamAperture# = 0# Then BeamAperture# = 0#
If BeamDumpPeriod# = 0# Then BeamDumpPeriod# = 15#

If BeamNumberSimulatedShowers# = 0 Then BeamNumberSimulatedShowers# = 2000000000#
If BeamSimulationTimePeriod# = 0 Then BeamSimulationTimePeriod# = 100000#
If BeamProductionIndex& = 0 Then BeamProductionIndex& = 0

If Trim$(MaterialFiles$(1)) = vbNullString Then MaterialFiles$(1) = "Copper metal.mat"
If Trim$(InputFile$) = vbNullString Then InputFile$ = "Copper metal.in"

BeamProductionFilename$(0) = PENEPMA_Root$ & "\Cu_cha.in"
BeamProductionFilename$(1) = PENEPMA_Root$ & "\Cu_back.in"
BeamProductionFilename$(2) = PENEPMA_Root$ & "\Cu_cont.in"
BeamProductionFilename$(3) = PENEPMA_Root$ & "\CuFe_sec.in"
BeamProductionFilename$(4) = PENEPMA_Root$ & "\Bilayer.in"

For i% = 1 To MAXMATOUTPUT%
InputEABS1#(i%) = 1000#     ' 1 keV
InputEABS2#(i%) = 1000#     ' 1 keV
Next i%

' Binary calculations
If BinaryElement1% = 0 Then BinaryElement1% = 27
If BinaryElement2% = 0 Then BinaryElement2% = 29

If PureElement1% = 0 Then PureElement1% = 26    ' Fe
If PureElement2% = 0 Then PureElement2% = 30    ' Zn

' Load Penepma08/12 atomic weights (for self consistent calculations)
Call Penepma12AtomicWeights
If ierror Then Exit Sub

Exit Sub

' Errors
Penepma08InitError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08Init"
ierror = True
Exit Sub

Penepma08InitBadVersion:
msg$ = "Problem with PENEPMA path statements"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08Init"
ierror = True
Exit Sub

End Sub

Sub Penepma08CreateInput(mode As Integer)
' Create input file (*.IN)
'  mode = 0 normal MsgBox
'  mode = 1 no MsgBox

ierror = False
On Error GoTo Penepma08CreateInputError

Dim k As Integer
Dim astring As String, bstring As String, cstring As String, dstring As String

' Loop through sample production file and copy to new file with modified parameters
Open BeamProductionFilename$(BeamProductionIndex&) For Input As #Temp1FileNumber%
Open PENEPMA_Path$ & "\" & InputFile$ For Output As #Temp2FileNumber%

Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$
bstring$ = astring$

If InStr(astring$, "TITLE") > 0 Then bstring$ = Left$(astring, COL7%) & Left$(BeamTitle$, 120)

cstring$ = Format$(Format$(BeamEnergy#, "Scientific"), a10$)
If InStr(astring$, "SENERG") > 0 Then Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

cstring$ = Format$(BeamPosition#(1), "Scientific") & " " & Format$(BeamPosition#(2)) & " " & Format$(BeamPosition#(3))
If InStr(astring$, "SPOSIT") > 0 Then Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

cstring$ = Format$(BeamDirection#(1)) & " " & Format$(BeamDirection#(2))
If InStr(astring$, "SDIREC") > 0 Then Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

cstring$ = Format$(BeamAperture#)
If InStr(astring$, "SAPERT") > 0 Then Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

' Specify detector geometry (0 = annular, 1 = north, 2 = east, 3 = south, 4 = west)
If InStr(astring$, "PDANGL") > 0 Then
cstring$ = Format$(InputTheta1#, "0.0") & " " & Format$(InputTheta2#, "0.0") & " "
cstring$ = cstring$ & Format$(InputPhi1#, "0.0") & " " & Format$(InputPhi2#, "0.0") & " "
cstring$ = cstring$ & Format$(InputISPF&, "0")
Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub
End If

cstring$ = Format$(BeamDumpPeriod#)
If InStr(astring$, "DUMPP") > 0 Then Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

' Load each material file and simulation parameters
If InStr(astring$, "MFNAME") > 0 Then
k% = k% + 1
cstring$ = MiscGetFileNameOnly$(MaterialFiles$(k%))
Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub
End If

If InStr(astring$, "MSIMPA") > 0 Then
cstring$ = Format$(InputEABS1#(k%), "0.0E+0") & " " & Format$(InputEABS2#(k%), "0.0E+0") & " " & Format$(InputEABS3#(k%), "0E+0") & " "
If InputC1#(k%) = 0# And InputC2#(k%) = 0# Then
cstring$ = cstring$ & Format$(InputC1#(k%), "0") & " " & Format$(InputC2#(k%), "0") & " "
Else
cstring$ = cstring$ & Format$(InputC1#(k%), "0.0") & " " & Format$(InputC2#(k%), "0.0") & " "
End If
If InputWCC#(k%) = 0# Then
cstring$ = cstring$ & Format$(InputWCC#(k%), "0") & " " & Format$(InputWCR#(k%), "0E+0")
Else
cstring$ = cstring$ & Format$(InputWCC#(k%), "0E+0") & " " & Format$(InputWCR#(k%), "0E+0")
End If
Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub
End If

' Load energy range parameters
cstring$ = Format$(BeamMinimumEnergyRange#, f41$) & " " & Format$(BeamMaximumEnergyRange#, e61$) & " "
cstring$ = cstring$ & Format$(BeamNumberOfEnergyChannels&, "0")
If InStr(astring$, "PDENER") > 0 Then Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

' Load geometry file
cstring$ = MiscGetFileNameOnly$(GeometryFile$)
If InStr(astring$, "GEOMFN") > 0 Then Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

cstring$ = Format$(BeamNumberSimulatedShowers#, e71$)
If InStr(astring$, "NSIMSH") > 0 Then Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

cstring$ = Format$(BeamSimulationTimePeriod#)
If InStr(astring$, "TIME") > 0 Then Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

Print #Temp2FileNumber%, bstring$
Loop

Close #Temp1FileNumber%
Close #Temp2FileNumber%

' Confirm with user
Call IOStatusAuto(vbNullString)
If mode% = 0 Then
msg$ = "Input file " & InputFile$ & " containing PENEPMA input parameters was saved to " & PENEPMA_Path$ & "\" & " folder. " & vbCrLf & vbCrLf
msg$ = msg$ & "Use this input file for Penelope Monte-carlo calculations and execute the PENEPMA application by clicking the "
msg$ = msg$ & "Run Input File button or by clicking the PENEPMA Prompt button and executing the program from the command prompt."
MsgBox msg$, vbOKOnly + vbInformation, "Penepma08CreateInput"

Else
Call MiscDelay(CDbl(0.5), Now)    ' add delay for batch to run
End If

Exit Sub

' Errors
Penepma08CreateInputError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08CreateInput"
Close #Temp1FileNumber%
Close #Temp2FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08CreateMaterial(tForm As Form)
' Create material file (*.MAT)(PENEPMA08/12 version)

ierror = False
On Error GoTo Penepma08CreateMaterialError

Dim n As Integer, k As Integer

' Check that at least one material is selected
k% = 0
For n% = 1 To MAXMATOUTPUT%
If MaterialsSelected%(n%) > 0 Then k% = k% + 1
Next n%
If k% = 0 Then GoTo Penepma08CreateInputNoMaterials

' Loop on each selected material and create an input file (material1.inp, material2.inp, etc.) for each selected standard
For n% = 1 To MAXMATOUTPUT%
If MaterialsSelected%(n%) > 0 Then

' Get composition based on standard number
Call StandardGetMDBStandard(MaterialsSelected%(n%), PENEPMASample())
If ierror Then Exit Sub

Call IOStatusAuto("Creating material input file based on standard " & Str$(PENEPMASample(1).number%) & " " & PENEPMASample(1).Name$ & "...")
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

' Make material INP file
Call Penepma08CreateMaterialINP(n%, PENEPMASample())
If ierror Then Exit Sub
End If
Next n%

' Create and run the necessary batch files
Call Penepma08CreateMaterialBatch(Int(0), tForm)
If ierror Then Exit Sub

Exit Sub

' Errors
Penepma08CreateMaterialError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08CreateMaterial"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08CreateInputNoMaterials:
msg$ = "No materials selected for output"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08CreateMaterial"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08CreateMaterialINP(n As Integer, sample() As TypeSample)
' Create a single INP (redirected keyboard input) file based on the specified standard material

ierror = False
On Error GoTo Penepma08CreateMaterialINPError

Dim i As Integer
Dim tfilename As String, astring As String

' Load file name
astring$ = "material" & Format$(n%) & ".inp"
tfilename$ = PENDBASE_Path$ & "\" & astring$
Open tfilename$ For Output As #Temp1FileNumber%

' Output configuration
Print #Temp1FileNumber%, "1"                                    ' enter composition from keyboard
Print #Temp1FileNumber%, Left$(sample(1).Name$, 60)      ' material name
Print #Temp1FileNumber%, Format$(sample(1).LastChan%)    ' number of elements in composition

' If more than one element enter composition
If sample(1).LastChan% = 1 Then
Print #Temp1FileNumber%, Format$(sample(1).AtomicNums%(1))

Else
Print #Temp1FileNumber%, "2"   ' enter by weight fraction

' Output composition of film
For i% = 1 To sample(1).LastChan%
If sample(1).ElmPercents!(i%) < PENEPMA_MINPERCENT! Then sample(1).ElmPercents!(i%) = PENEPMA_MINPERCENT!
astring$ = Format$(sample(1).AtomicNums%(i%)) & VbComma$ & Trim$(MiscAutoFormat$(sample(1).ElmPercents!(i%) / 100#))
Print #Temp1FileNumber%, astring$
Next i%
End If

Print #Temp1FileNumber%, "2"                                    ' do not change mean excitation energy
Print #Temp1FileNumber%, Trim$(Str$(MaterialDensity#))          ' density of material

' Load default oscillator energy and strength
If MaterialFcb# = 0# And MaterialWcb# = 0# Then
Print #Temp1FileNumber%, "2"
Else
If MaterialFcb# = 0# Or MaterialWcb# = 0# Then GoTo Penepma08CreateMaterialINPZero
Print #Temp1FileNumber%, "1"
Print #Temp1FileNumber%, Trim$(Str$(MaterialFcb#)) & VbComma$ & Trim$(Str$(MaterialWcb#))
End If

astring$ = "material" & Format$(n%) & ".mat"                    ' use same folder as MATERIAL.EXE
Print #Temp1FileNumber%, Left$(astring$, 80)                    ' material filename
Close #Temp1FileNumber%

Exit Sub

' Errors
Penepma08CreateMaterialINPError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08CreateMaterialINP"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08CreateMaterialINPZero:
msg$ = "One of the material oscillator parameters is zero, please enter zero values for both or non-zero values for both and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08CreateMaterialINP"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08CreateMaterialBatch(mode As Integer, tForm As Form)
' Create and run material batch files
'  mode = 0 normal MsgBox
'  mode = 1 no MsgBox

ierror = False
On Error GoTo Penepma08CreateMaterialBatchError

Dim n As Integer
Dim bfilename As String, astring As String, bstring As String

' Run each material[n].inp file
For n% = 1 To MAXMATOUTPUT%
If MaterialsSelected%(n%) > 0 Then
Call IOStatusAuto("Creating material file " & Format$(n%) & " by running MATERIAL.EXE (this may take a while)...")
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

' Delete exisitng batch file
bfilename$ = PENDBASE_Path$ & "\temp.bat"
If Dir$(bfilename$) <> vbNullString Then
Kill bfilename$
DoEvents
End If

' Write new batch file
Open bfilename$ For Output As #Temp1FileNumber%
astring$ = Left$(PENDBASE_Path$, 2)                             ' change to drive
Print #Temp1FileNumber%, astring$
astring$ = "cd " & VbDquote$ & PENDBASE_Path$ & VbDquote$       ' change to folder
Print #Temp1FileNumber%, astring$
astring$ = "material.exe < " & "material" & Format$(n%) & ".inp"
Print #Temp1FileNumber%, astring$
Close #Temp1FileNumber%

' Run each material batch file synchronously
astring$ = PENDBASE_Path$ & "\temp.bat"
Call ExecRun(astring$)
If ierror Then Exit Sub
End If

DoEvents
Next n%

' Now copy all files to original material names (up to 20 characters)
Call IOStatusAuto("Copying temp material files to target material file...")
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

For n% = 1 To MAXMATOUTPUT%
If MaterialsSelected%(n%) > 0 Then
astring$ = "material" & Format$(n%) & ".mat"
bstring$ = MiscGetFileNameNoExtension$(MaterialFiles$(n%)) & ".mat"
FileCopy PENDBASE_Path$ & "\" & astring$, PENDBASE_Path$ & "\" & bstring$
DoEvents
FileCopy PENDBASE_Path$ & "\" & astring$, PENEPMA_Path$ & "\" & bstring$
DoEvents
End If
Next n%

' Update material input with modified filename
For n% = 1 To MAXMATOUTPUT%
If MaterialsSelected%(n%) > 0 Then
tForm.TextMaterialFiles(n% - 1).Text = MaterialFiles$(n%)
End If
Next n%

' Confirm with user
Call IOStatusAuto(vbNullString)
If mode% = 0 Then
msg$ = "Material files containing compositions were created and saved to " & PENDBASE_Path$ & "\" & " folder. " & vbCrLf & vbCrLf
msg$ = msg$ & "Use these material files for creating a PENEPMA input file for Penepma Monte-Carlo calculations."
MsgBox msg$, vbOKOnly + vbInformation, "Penepma08CreateMaterialBatch"
End If

' Update input field with material name (and conditions)
InputFile$ = vbNullString
For n% = 1 To MAXMATOUTPUT%
If MaterialsSelected%(n%) > 0 Then
If n% <> 1 Then InputFile$ = InputFile$ & "_"
InputFile$ = InputFile$ & MiscGetFileNameNoExtension$(MaterialFiles$(n%))
End If
Next n%
If InputFile$ = vbNullString Then InputFile$ = "Input.in"
tForm.TextInputFile.Text = InputFile$

Exit Sub

' Errors
Penepma08CreateMaterialBatchError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08CreateMaterialBatch"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08CreateInputNoMaterials:
msg$ = "No materials selected for output"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08CreateMaterialBatch"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08BrowseMaterialFile(n As Integer, tForm As Form)
' Browse to a selected material file (*.MAT) for creating an input file

ierror = False
On Error GoTo Penepma08BrowseMaterialFileError

Dim tfilename As String, ioextension As String

If Trim$(MaterialFiles$(n%)) = vbNullString Then MaterialFiles$(n%) = "epma.mat"
ioextension$ = "MAT"
tfilename$ = PENEPMA_Path$ & "\" & MaterialFiles$(n%)
Call IOGetFileName(Int(2), ioextension$, tfilename$, tForm)
If ierror Then Exit Sub

' Copy file to Penepma folder if default folder not selected
If Not MiscStringsAreSame(tfilename$, PENEPMA_Path$ & "\" & MiscGetFileNameOnly$(tfilename$)) Then
Kill PENEPMA_Path$ & "\" & MiscGetFileNameOnly$(tfilename$)
FileCopy tfilename$, PENEPMA_Path$ & "\" & MiscGetFileNameOnly$(tfilename$)
End If

' Load to module and dialog
MaterialFiles$(n%) = MiscGetFileNameOnly$(tfilename$)
tForm.TextMaterialFiles(n% - 1).Text = MaterialFiles$(n%)

If n% = 1 Then
InputFile$ = MiscGetFileNameNoExtension$(MaterialFiles$(n%)) & ".in"
tForm.TextInputFile.Text = InputFile$
End If

Exit Sub

' Errors
Penepma08BrowseMaterialFileError:
MsgBox Error$, vbOKOnly + vbCritical, "BrowseMaterialFile"
ierror = True
Exit Sub

End Sub

Sub Penepma08BrowseGeometryFile(tForm As Form)
' Browse to a selected geometry file (*.GEO)

ierror = False
On Error GoTo Penepma08BrowseGeometryFileError

Dim tfilename As String, ioextension As String

Static initialized As Boolean

If Trim$(GeometryFile$) = vbNullString Then GeometryFile$ = PENEPMA_Root$ & "\bulk.geo"
ioextension$ = "GEO"
tfilename$ = GeometryFile$
Call IOGetFileName(Int(2), ioextension$, tfilename$, tForm)
If ierror Then Exit Sub

' Copy selected GEO file to Penepma folder
If Not MiscStringsAreSame(MiscGetPathOnly2(tfilename$), PENEPMA_Path$) Then
If Dir$(PENEPMA_Path$ & "\" & MiscGetFileNameOnly$(tfilename$)) <> vbNullString Then
Kill PENEPMA_Path$ & "\" & MiscGetFileNameOnly$(tfilename$)
End If
FileCopy tfilename$, PENEPMA_Path$ & "\" & MiscGetFileNameOnly$(tfilename$)
End If

' Load to module and dialog
GeometryFile$ = MiscGetFileNameOnly$(tfilename$)
tForm.TextGeometryFile.Text = GeometryFile$

' Warn user (once) if using couple/sphere production file and *sphere geo file and X beam position is not sero
If Not initialized And tForm.OptionProduction(3).Value = True Then
If Val(tForm.TextBeamPosition(0).Text) <> 0# And InStr(tForm.TextGeometryFile.Text, "sphere") > 0 Then
msg$ = "When using a hemisphere geometry file (*sphere.geo), be sure the X beam position is properly specified."
msg$ = msg$ & vbCrLf & vbCrLf & "The beam position is in the center of the hemisphere only when both the X and Y Beam Position parameters above are zero."
MsgBox msg$, vbOKOnly + vbInformation, "Penepma08BrowseGeometryFile"
End If
initialized = True
End If

Exit Sub

' Errors
Penepma08BrowseGeometryFileError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08BrowseGeometryFile"
ierror = True
Exit Sub

End Sub

Sub Penepma08BrowseInputFile(tForm As Form)
' Browse to a selected Penepma Input file (*.IN)

ierror = False
On Error GoTo Penepma08BrowseInputFileError

Dim tfilename As String, ioextension As String

If Trim$(InputFile$) = vbNullString Then InputFile$ = "untitled.in"
ioextension$ = "IN"
tfilename$ = PENEPMA_Path$ & "\" & InputFile$
Call IOGetFileName(Int(2), ioextension$, tfilename$, tForm)
If ierror Then Exit Sub

' Copy to input file to Penepma folder if input file folder in not the Penepma folder
If Not MiscStringsAreSame(MiscGetPathOnly2$(tfilename$), PENEPMA_Path$) Then
If Dir$(PENEPMA_Path$ & "\" & MiscGetFileNameOnly$(tfilename$)) <> vbNullString Then
Kill PENEPMA_Path$ & "\" & MiscGetFileNameOnly$(tfilename$)
End If
FileCopy tfilename$, PENEPMA_Path$ & "\" & MiscGetFileNameOnly$(tfilename$)
End If

' Load controls
Call Penepma08LoadInputFile(tfilename$, tForm)
If ierror Then Exit Sub

' Load to module and dialog
Call Penepma08SaveInput(tForm)
If ierror Then Exit Sub

' Re-load input file to text field
InputFile$ = MiscGetFileNameOnly$(tfilename$)
tForm.TextInputFile.Text = InputFile$

' Copy material file (if found) to Penepma folder
If Not MiscStringsAreSame(MiscGetPathOnly2$(tfilename$), PENEPMA_Path$) Then
If Trim$(tForm.TextMaterialFiles(0).Text) <> vbNullString Then
If Dir$(PENEPMA_Path$ & "\" & tForm.TextMaterialFiles(0).Text) <> vbNullString Then
Kill PENEPMA_Path$ & "\" & tForm.TextMaterialFiles(0).Text
End If
FileCopy MiscGetPathOnly2$(tfilename$) & "\" & tForm.TextMaterialFiles(0).Text, PENEPMA_Path$ & "\" & tForm.TextMaterialFiles(0).Text
End If

If Trim$(tForm.TextMaterialFiles(MAXMATOUTPUT% - 1).Text) <> vbNullString Then
If Dir$(PENEPMA_Path$ & "\" & tForm.TextMaterialFiles(MAXMATOUTPUT% - 1).Text) <> vbNullString Then
Kill PENEPMA_Path$ & "\" & tForm.TextMaterialFiles(MAXMATOUTPUT% - 1).Text
End If
FileCopy MiscGetPathOnly2$(tfilename$) & "\" & tForm.TextMaterialFiles(MAXMATOUTPUT% - 1).Text, PENEPMA_Path$ & "\" & tForm.TextMaterialFiles(MAXMATOUTPUT% - 1).Text
End If
End If

' Copy geometry file (if found) to Penepma folder
If Not MiscStringsAreSame(MiscGetPathOnly2$(tfilename$), PENEPMA_Path$) Then
If Trim$(tForm.TextGeometryFile.Text) <> vbNullString Then
If Dir$(PENEPMA_Path$ & "\" & tForm.TextGeometryFile.Text) <> vbNullString Then
Kill PENEPMA_Path$ & "\" & tForm.TextGeometryFile.Text
End If
FileCopy MiscGetPathOnly2$(tfilename$) & "\" & tForm.TextGeometryFile.Text, PENEPMA_Path$ & "\" & tForm.TextGeometryFile.Text
End If
End If

Exit Sub

' Errors
Penepma08BrowseInputFileError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08BrowseInputFile"
ierror = True
Exit Sub

End Sub

Sub Penepma08BatchBrowseFolder()
' Browse to a batch project folder

ierror = False
On Error GoTo Penepma08BatchBrowseFolderError

Dim tpath As String, tstring As String

' Load to module and dialog
tstring$ = "Browse PENEPMA Batch Project Folder"
If PENEPMA_BATCH_FOLDER$ = vbNullString Then PENEPMA_BATCH_FOLDER$ = PENEPMA_Path$
tpath$ = IOBrowseForFolderByPath(True, PENEPMA_BATCH_FOLDER$, tstring$, FormPENEPMA08Batch)
If ierror Then Exit Sub

If Trim$(tpath$) <> vbNullString Then PENEPMA_BATCH_FOLDER$ = tpath$
FormPENEPMA08Batch.TextBatchFolder.Text = PENEPMA_BATCH_FOLDER$

Exit Sub

' Errors
Penepma08BatchBrowseFolderError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08BatchBrowseFolder"
ierror = True
Exit Sub

End Sub

Sub Penepma08BatchLoad()
' Browse to a batch project folder

ierror = False
On Error GoTo Penepma08BatchLoadError

Dim i As Integer
Dim astring As String, tfilename As String
Dim n As Integer

' Load window
If PENEPMA_BATCH_FOLDER$ = vbNullString Then PENEPMA_BATCH_FOLDER$ = PENEPMA_Path$ & "\Batch"
FormPENEPMA08Batch.TextBatchFolder.Text = PENEPMA_BATCH_FOLDER$

' Load binary elements
FormPENEPMA08Batch.ComboBinaryElement1.Clear
For i% = 0 To MAXELM% - 1
FormPENEPMA08Batch.ComboBinaryElement1.AddItem Symup$(i% + 1)
Next i%
FormPENEPMA08Batch.ComboBinaryElement1.ListIndex = BinaryElement1% - 1

FormPENEPMA08Batch.ComboBinaryElement2.Clear
For i% = 0 To MAXELM% - 1
FormPENEPMA08Batch.ComboBinaryElement2.AddItem Symup$(i% + 1)
Next i%
FormPENEPMA08Batch.ComboBinaryElement2.ListIndex = BinaryElement2% - 1

' Load bulk pure elements
FormPENEPMA08Batch.ComboPureElement1.Clear
For i% = 0 To MAXELM% - 1
FormPENEPMA08Batch.ComboPureElement1.AddItem Symup$(i% + 1)
Next i%
FormPENEPMA08Batch.ComboPureElement1.ListIndex = PureElement1% - 1

FormPENEPMA08Batch.ComboPureElement2.Clear
For i% = 0 To MAXELM% - 1
FormPENEPMA08Batch.ComboPureElement2.AddItem Symup$(i% + 1)
Next i%
FormPENEPMA08Batch.ComboPureElement2.ListIndex = PureElement2% - 1

' Load list of input files in PENEPMA_Path$
ReDim InputFiles(1 To 1) As String
ReDim InputDates(1 To 1) As Variant

FormPENEPMA08Batch.ListInputFiles.Clear
tfilename$ = Dir$(PENEPMA_Path$ & "\*.in") ' get first file
n% = 0
Do While tfilename$ <> vbNullString

' Add to file list
n% = n% + 1
ReDim Preserve InputFiles(1 To n%) As String
ReDim Preserve InputDates(1 To n%) As Variant
InputFiles$(n%) = tfilename$
InputDates(n%) = FileDateTime(PENEPMA_Path$ & "\" & InputFiles$(n%))

' Get next filename
tfilename$ = Dir$
Loop

' Call for sorted directory
If FormPENEPMA08Batch.CheckSortByDate.Value = vbChecked Then
Call MiscDirectorySort(PENEPMA_Path$ & "\*.in", InputFiles$(), InputDates())
If ierror Then Exit Sub
End If

' Load list box with files (check for pe-layout.in file in Penepma08LoadInputFile)
For i% = 1 To UBound(InputFiles$())
astring$ = Format$(Left$(InputFiles$(i%), 31), a32$) & " " & Format$(Format$(InputDates(i%), "General Date"), a22$)
FormPENEPMA08Batch.ListInputFiles.AddItem astring$
FormPENEPMA08Batch.ListInputFiles.ItemData(FormPENEPMA08Batch.ListInputFiles.NewIndex) = i%
Next i%

' Load element and xray to extract
If ExtractElement% = 0 Then ExtractElement% = 16    ' sulfur
If ExtractXray% = 0 Then ExtractXray% = 1           ' ka

FormPENEPMA08Batch.ComboElm.Clear
For i% = 0 To MAXELM% - 1
FormPENEPMA08Batch.ComboElm.AddItem Symup$(i% + 1)
Next i%
FormPENEPMA08Batch.ComboElm.ListIndex = ExtractElement% - 1

FormPENEPMA08Batch.ComboXray.Clear
For i% = 0 To MAXRAY% - 2
FormPENEPMA08Batch.ComboXray.AddItem Xraylo$(i% + 1)
Next i%
FormPENEPMA08Batch.ComboXray.ListIndex = ExtractXray% - 1

' Select last file
If FormPENEPMA08Batch.ListInputFiles.ListCount > 0 Then
FormPENEPMA08Batch.ListInputFiles.ListIndex = FormPENEPMA08Batch.ListInputFiles.ListCount - 1
FormPENEPMA08Batch.ListInputFiles.Selected(FormPENEPMA08Batch.ListInputFiles.ListIndex) = True
Else
FormPENEPMA08Batch.CommandRunBatch.Enabled = False
End If

FormPENEPMA08Batch.Show vbModeless

Exit Sub

' Errors
Penepma08BatchLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08BatchLoad"
ierror = True
Exit Sub

End Sub

Sub Penepma08BatchGetInputParameters(Index As Integer)
' Load values from selected input file

ierror = False
On Error GoTo Penepma08BatchGetInputParametersError

Dim astring As String, bstring As String

' Check that input files exist
If Index% < 0 Then Exit Sub

' Skip file pe-layout.in
If InStr(InputFiles$(Index% + 1), "pe-layout.in") > 0 Then Exit Sub

' Check input file parameters
If Dir$(Trim$(PENEPMA_Path$ & "\" & InputFiles$(Index% + 1))) = vbNullString Then Exit Sub

' Load frame title
FormPENEPMA08Batch.FrameInputFileParameters.Caption = "Selected File Properties [" & InputFiles$(Index% + 1) & "]"

' Open file and load values
Open PENEPMA_Path$ & "\" & InputFiles$(Index% + 1) For Input As #Temp1FileNumber%

Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$

If InStr(astring$, "TITLE") > 0 Then Call Penepma08LoadProduction1(astring$, bstring$, FormPENEPMA08Batch.TextInputTitle)
If ierror Then
MsgBox "Problem reading file " & PENEPMA_Path$ & "\" & InputFiles$(Index% + 1), vbOKOnly + vbExclamation, "Penepma08BatchGetInputParameters [Penepma08LoadProduction1]"
Exit Sub
End If
If InStr(astring$, "SENERG") > 0 Then Call Penepma08LoadProduction2(astring$, bstring$, FormPENEPMA08Batch.TextBeamEnergy)
If ierror Then
MsgBox "Problem reading file " & PENEPMA_Path$ & "\" & InputFiles$(Index% + 1), vbOKOnly + vbExclamation, "Penepma08BatchGetInputParameters [Penepma08LoadProduction2]"
Exit Sub
End If
If InStr(astring$, "DUMPP") > 0 Then Call Penepma08LoadProduction2(astring$, bstring$, FormPENEPMA08Batch.TextDumpPeriod)
If ierror Then
MsgBox "Problem reading file " & PENEPMA_Path$ & "\" & InputFiles$(Index% + 1), vbOKOnly + vbExclamation, "Penepma08BatchGetInputParameters [Penepma08LoadProduction2]"
Exit Sub
End If
If InStr(astring$, "NSIMSH") > 0 Then Call Penepma08LoadProduction2(astring$, bstring$, FormPENEPMA08Batch.TextNumberSimulatedShowers)
If ierror Then
MsgBox "Problem reading file " & PENEPMA_Path$ & "\" & InputFiles$(Index% + 1), vbOKOnly + vbExclamation, "Penepma08BatchGetInputParameters [Penepma08LoadProduction2]"
Exit Sub
End If
If InStr(astring$, "TIME") > 0 Then Call Penepma08LoadProduction2(astring$, bstring$, FormPENEPMA08Batch.TextSimulationTimePeriod)
If ierror Then
MsgBox "Problem reading file " & PENEPMA_Path$ & "\" & InputFiles$(Index% + 1), vbOKOnly + vbExclamation, "Penepma08BatchGetInputParameters [Penepma08LoadProduction2]"
Exit Sub
End If
If InStr(astring$, "MSIMPA") > 0 Then Call Penepma08LoadProduction4(astring$, bstring$, Int(1), FormPENEPMA08Batch)
If ierror Then
MsgBox "Problem reading file " & PENEPMA_Path$ & "\" & InputFiles$(Index% + 1), vbOKOnly + vbExclamation, "Penepma08BatchGetInputParameters [Penepma08LoadProduction4]"
Exit Sub
End If
If InStr(astring$, "PDENER") > 0 Then Call Penepma08LoadProduction3(astring$, bstring$, FormPENEPMA08Batch)
If ierror Then
MsgBox "Problem reading file " & PENEPMA_Path$ & "\" & InputFiles$(Index% + 1), vbOKOnly + vbExclamation, "Penepma08BatchGetInputParameters [Penepma08LoadProduction3]"
Exit Sub
End If

Loop

Close #Temp1FileNumber%
Exit Sub

' Errors
Penepma08BatchGetInputParametersError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08BatchGetInputParameters"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08LoadProduction(Index As Integer, tForm As Form)
' Load values from selected production file

ierror = False
On Error GoTo Penepma08LoadProductionError

Dim k As Integer
Dim astring As String, bstring As String

' Check production file
If Dir$(Trim$(BeamProductionFilename$(Index%))) = vbNullString Then GoTo Penepma08LoadProductionNotFound

' Open file and load values
Open BeamProductionFilename$(Index%) For Input As #Temp1FileNumber%

Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$

If InStr(astring$, "TITLE") > 0 Then
If tForm.OptionProduction(0).Value = True Then tForm.TextInputTitle.Text = "Characteristic X-ray Production Model"
If tForm.OptionProduction(1).Value = True Then tForm.TextInputTitle.Text = "Backscatter Electron Production Model"
If tForm.OptionProduction(2).Value = True Then tForm.TextInputTitle.Text = "Continuum X-ray Production Model"
If tForm.OptionProduction(3).Value = True Then tForm.TextInputTitle.Text = "Secondary Fluorescence Couple X-ray Production Model"
If tForm.OptionProduction(4).Value = True Then tForm.TextInputTitle.Text = "Thin Film On Substrate X-ray Production Model"
End If

If InStr(astring$, "SENERG") > 0 Then Call Penepma08LoadProduction2(astring$, bstring$, tForm.TextBeamEnergy)
If ierror Then
MsgBox "Problem reading file " & BeamProductionFilename$(Index%), vbOKOnly + vbExclamation, "Penepma08LoadProduction [Penepma08LoadProduction2]"
Exit Sub
End If
If InStr(astring$, "SPOSIT") > 0 Then Call Penepma08LoadProduction3(astring$, bstring$, tForm)
If ierror Then
MsgBox "Problem reading file " & BeamProductionFilename$(Index%), vbOKOnly + vbExclamation, "Penepma08LoadProduction [Penepma08LoadProduction3]"
Exit Sub
End If
If InStr(astring$, "SDIREC") > 0 Then Call Penepma08LoadProduction3(astring$, bstring$, tForm)
If ierror Then
MsgBox "Problem reading file " & BeamProductionFilename$(Index%), vbOKOnly + vbExclamation, "Penepma08LoadProduction [Penepma08LoadProduction3]"
Exit Sub
End If
If InStr(astring$, "SAPERT") > 0 Then Call Penepma08LoadProduction2(astring$, bstring$, tForm.TextBeamAperture)
If ierror Then
MsgBox "Problem reading file " & BeamProductionFilename$(Index%), vbOKOnly + vbExclamation, "Penepma08LoadProduction [Penepma08LoadProduction2]"
Exit Sub
End If

If InStr(astring$, "MFNAME") > 0 Then k% = k% + 1       ' do not load material file, just simulation parameters
If InStr(astring$, "MSIMPA") > 0 Then Call Penepma08LoadProduction4(astring$, bstring$, k%, tForm)
If ierror Then
MsgBox "Problem reading file " & BeamProductionFilename$(Index%), vbOKOnly + vbExclamation, "Penepma08LoadProduction [Penepma08LoadProduction4]"
Exit Sub
End If

If InStr(astring$, "PDANGL") > 0 Then Call Penepma08LoadProduction5(astring$, bstring$)
If ierror Then
MsgBox "Problem reading file " & BeamProductionFilename$(Index%), vbOKOnly + vbExclamation, "Penepma08LoadProduction [Penepma08LoadProduction5]"
Exit Sub
End If

If InStr(astring$, "GEOMFN") > 0 Then Call Penepma08LoadProduction2(astring$, bstring$, tForm.TextGeometryFile)
If ierror Then
MsgBox "Problem reading file " & BeamProductionFilename$(Index%), vbOKOnly + vbExclamation, "Penepma08LoadProduction [Penepma08LoadProduction2]"
Exit Sub
End If
If InStr(astring$, "DUMPP") > 0 Then Call Penepma08LoadProduction2(astring$, bstring$, tForm.TextDumpPeriod)
If ierror Then
MsgBox "Problem reading file " & BeamProductionFilename$(Index%), vbOKOnly + vbExclamation, "Penepma08LoadProduction [Penepma08LoadProduction2]"
Exit Sub
End If
If InStr(astring$, "NSIMSH") > 0 Then Call Penepma08LoadProduction2(astring$, bstring$, tForm.TextNumberSimulatedShowers)
If ierror Then
MsgBox "Problem reading file " & BeamProductionFilename$(Index%), vbOKOnly + vbExclamation, "Penepma08LoadProduction [Penepma08LoadProduction2]"
Exit Sub
End If
If InStr(astring$, "TIME") > 0 Then Call Penepma08LoadProduction2(astring$, bstring$, tForm.TextSimulationTimePeriod)
If ierror Then
MsgBox "Problem reading file " & BeamProductionFilename$(Index%), vbOKOnly + vbExclamation, "Penepma08LoadProduction [Penepma08LoadProduction2]"
Exit Sub
End If
If InStr(astring$, "PDENER") > 0 Then Call Penepma08LoadProduction3(astring$, bstring$, tForm)
If ierror Then
MsgBox "Problem reading file " & BeamProductionFilename$(Index%), vbOKOnly + vbExclamation, "Penepma08LoadProduction [Penepma08LoadProduction3]"
Exit Sub
End If

Loop

Close #Temp1FileNumber%
Exit Sub

' Errors
Penepma08LoadProductionError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08LoadProduction"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08LoadProductionNotFound:
msg$ = "Penepma production file (" & BeamProductionFilename$(Index%) & ") was not found. Please download an up to date Penepma12.zip file and extract to your Penepma12 folder or contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbCritical, "Penepma08LoadProduction"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08LoadInput(tfilename As String, cstring As String, bstring As String, k As Integer)
' Just return the specified string from the specified input file

ierror = False
On Error GoTo Penepma08LoadInputError

Dim n As Integer
Dim astring As String

' Open file and load values
Open tfilename$ For Input As #Temp1FileNumber%

bstring$ = vbNullString         ' initialize
Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$

' Read material filename
If InStr(astring$, "MFNAME") > 0 And InStr(astring$, cstring$) > 0 Then
n% = n% + 1
If n% = k% Then
bstring$ = Mid$(astring, COL7% + 1, InStr(astring$, "[") - (COL7% + 1))
bstring$ = Trim$(bstring$)
End If
End If

' Read geometry filename
If InStr(astring$, "GEOMFN") > 0 And InStr(astring$, cstring$) > 0 Then
bstring$ = Mid$(astring, COL7% + 1, InStr(astring$, "[") - (COL7% + 1))
bstring$ = Trim$(bstring$)
End If

Loop
Close #Temp1FileNumber%

Exit Sub

' Errors
Penepma08LoadInputError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08LoadInput"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08LoadProduction1(astring As String, bstring As String, tText As TextBox)
' Load the text to the passed text box (single control) (for reading title only)

ierror = False
On Error GoTo Penepma08LoadProduction1Error

If astring$ = vbNullString Then GoTo Penepma08LoadProduction1EmptyString
If Len(astring$) < COL7% + 1 Then GoTo Penepma08LoadProduction1ShortString

' Load the parameter text
bstring$ = Mid$(astring, COL7% + 1)
tText.Text = Trim$(bstring$)

Exit Sub

' Errors
Penepma08LoadProduction1Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08LoadProduction1"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08LoadProduction1EmptyString:
msg$ = "Unable to parse Penepma production file string for text control " & tText.Name & ". Please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08LoadProduction1"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08LoadProduction1ShortString:
msg$ = "Unexpectedly short string parsing Penepma production file string (" & astring$ & ") for text control " & tText.Name & ". Please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08LoadProduction1"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08LoadProduction2(astring As String, bstring As String, tText As TextBox)
' Load the text to the passed text box (single control)

ierror = False
On Error GoTo Penepma08LoadProduction2Error

If astring$ = vbNullString Then GoTo Penepma08LoadProduction2EmptyString
If Len(astring$) < COL7% + 1 Then GoTo Penepma08LoadProduction2ShortString
If InStr(astring$, "[") = 0 Then GoTo Penepma08LoadProduction2MissingBracket

' Load the parameter text
bstring$ = Mid$(astring$, COL7% + 1, InStr(astring$, "[") - (COL7% + 1))

If InStr(astring$, "GEOMFN") > 0 Then
tText.Text = Trim$(bstring$)
ElseIf InStr(astring$, "DUMPP") > 0 Then
tText.Text = Trim$(bstring$)

Else
tText.Text = Trim$(bstring$)
End If

Exit Sub

' Errors
Penepma08LoadProduction2Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08LoadProduction2"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08LoadProduction2EmptyString:
msg$ = "Unable to parse Penepma production file string for text control " & tText.Name & ". Please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08LoadProduction2"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08LoadProduction2ShortString:
msg$ = "Unexpectedly short string parsing Penepma production file string (" & astring$ & ") for text control " & tText.Name & ". Please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08LoadProduction2"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08LoadProduction2MissingBracket:
msg$ = "Missing square bracket in Penepma production file string (" & astring$ & ") for text control " & tText.Name & ". Please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08LoadProduction2"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08LoadProduction3(astring As String, bstring As String, tForm As Form)
' Load the text to the passed form (for control array of text boxes)

ierror = False
On Error GoTo Penepma08LoadProduction3Error

Dim cstring As String

If astring$ = vbNullString Then GoTo Penepma08LoadProduction3EmptyString
If Len(astring$) < COL7% + 1 Then GoTo Penepma08LoadProduction3ShortString
If InStr(astring$, "[") = 0 Then GoTo Penepma08LoadProduction3MissingBracket

' Load the parameter text
bstring$ = Mid$(astring, COL7% + 1, InStr(astring$, "[") - (COL7% + 1))

If InStr(astring$, "SPOSIT") > 0 Then
Call MiscParseStringToString(bstring$, cstring$)
tForm.TextBeamPosition(0).Text = Trim$(cstring$)
Call MiscParseStringToString(bstring$, cstring$)
tForm.TextBeamPosition(1).Text = Trim$(cstring$)
Call MiscParseStringToString(bstring$, cstring$)
tForm.TextBeamPosition(2).Text = Trim$(cstring$)

ElseIf InStr(astring$, "SDIREC") > 0 Then
Call MiscParseStringToString(bstring$, cstring$)
tForm.TextBeamDirection(0).Text = Trim$(cstring$)
Call MiscParseStringToString(bstring$, cstring$)
tForm.TextBeamDirection(1).Text = Trim$(cstring$)

ElseIf InStr(astring$, "PDENER") > 0 Then
Call MiscParseStringToString(bstring$, cstring$)            ' energy range (min)
tForm.TextEnergyRangeMinMaxNumber(0).Text = Trim$(cstring$)
Call MiscParseStringToString(bstring$, cstring$)            ' energy range (max)
tForm.TextEnergyRangeMinMaxNumber(1).Text = Trim$(cstring$)
Call MiscParseStringToString(bstring$, cstring$)            ' number of energy channels
tForm.TextEnergyRangeMinMaxNumber(2).Text = Trim$(cstring$)
End If

Exit Sub

' Errors
Penepma08LoadProduction3Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08LoadProduction3"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08LoadProduction3EmptyString:
msg$ = "Unable to parse Penepma production file string for form " & tForm.Name & ". Please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08LoadProduction3"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08LoadProduction3ShortString:
msg$ = "Unexpectedly short string parsing Penepma production file string (" & astring$ & ") for form " & tForm.Name & ". Please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08LoadProduction3"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08LoadProduction3MissingBracket:
msg$ = "Missing square bracket in Penepma production file string (" & astring$ & ") for form " & tForm.Name & ". Please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08LoadProduction3"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08LoadProduction4(astring As String, bstring As String, k As Integer, tForm As Form)
' Load the text to the passed form (for 2 dimensional control array of text boxes)

ierror = False
On Error GoTo Penepma08LoadProduction4Error

Dim cstring As String

If astring$ = vbNullString Then GoTo Penepma08LoadProduction4EmptyString
If Len(astring$) < COL7% + 1 Then GoTo Penepma08LoadProduction4ShortString
If InStr(astring$, "[") = 0 Then GoTo Penepma08LoadProduction4MissingBracket

' Load the parameter text
bstring$ = Mid$(astring$, COL7% + 1, InStr(astring$, "[") - (COL7% + 1))

If InStr(astring$, "MFNAME") > 0 Then
tForm.TextMaterialFiles(k% - 1).Text = Trim$(Left$(bstring$, 20))
MaterialFiles$(k%) = Trim$(cstring$)
End If

If InStr(astring$, "MSIMPA") > 0 Then
Call MiscParseStringToString(bstring$, cstring$)
tForm.TextEABS1(k% - 1).Text = Trim$(cstring$)
Call MiscParseStringToString(bstring$, cstring$)
tForm.TextEABS2(k% - 1).Text = Trim$(cstring$)

' Load other production variables for output (forcing parameters)
Call MiscParseStringToString(bstring$, cstring$)
InputEABS3#(k%) = Trim$(cstring$)

Call MiscParseStringToString(bstring$, cstring$)
InputC1#(k%) = Trim$(cstring$)
Call MiscParseStringToString(bstring$, cstring$)
InputC2#(k%) = Trim$(cstring$)

Call MiscParseStringToString(bstring$, cstring$)
InputWCC#(k%) = Trim$(cstring$)
Call MiscParseStringToString(bstring$, cstring$)
InputWCR#(k%) = Trim$(cstring$)
End If

Exit Sub

' Errors
Penepma08LoadProduction4Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08LoadProduction4"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08LoadProduction4EmptyString:
msg$ = "Unable to parse Penepma production file string for form " & tForm.Name & ". Please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08LoadProduction4"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08LoadProduction4ShortString:
msg$ = "Unexpectedly short string parsing Penepma production file string (" & astring$ & ") for form " & tForm.Name & ". Please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08LoadProduction4"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08LoadProduction4MissingBracket:
msg$ = "Missing square bracket in Penepma production file string (" & astring$ & ") for form " & tForm.Name & ". Please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08LoadProduction4"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08LoadProduction5(astring As String, bstring As String)
' Load the parameters to module level variables only

ierror = False
On Error GoTo Penepma08LoadProduction5Error

Dim cstring As String

If astring$ = vbNullString Then GoTo Penepma08LoadProduction5EmptyString
If Len(astring$) < COL7% + 1 Then GoTo Penepma08LoadProduction5ShortString
If InStr(astring$, "[") = 0 Then GoTo Penepma08LoadProduction5MissingBracket

' Load the parameters
bstring$ = Mid$(astring$, COL7% + 1, InStr(astring$, "[") - (COL7% + 1))

If InStr(astring$, "PDANGL") > 0 Then
Call MiscParseStringToString(bstring$, cstring$)
InputTheta1# = Val(Trim$(cstring$))
Call MiscParseStringToString(bstring$, cstring$)
InputTheta2# = Val(Trim$(cstring$))

Call MiscParseStringToString(bstring$, cstring$)
InputPhi1# = Val(Trim$(cstring$))
Call MiscParseStringToString(bstring$, cstring$)
InputPhi2# = Val(Trim$(cstring$))

Call MiscParseStringToString(bstring$, cstring$)
InputISPF& = Val(Trim$(cstring$))
End If

Exit Sub

' Errors
Penepma08LoadProduction5Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08LoadProduction5"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08LoadProduction5EmptyString:
msg$ = "Unable to parse Penepma production file string. Please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08LoadProduction5"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08LoadProduction5ShortString:
msg$ = "Unexpectedly short string parsing Penepma production file string. Please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08LoadProduction5"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08LoadProduction5MissingBracket:
msg$ = "Missing square bracket in Penepma production file string. Please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08LoadProduction5"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08EditInput()
' Edit the input file

ierror = False
On Error GoTo Penepma08EditInputError

Call IOTextViewer2(InputFile$)
If ierror Then Exit Sub

Exit Sub

' Errors
Penepma08EditInputError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08EditInput"
ierror = True
Exit Sub

End Sub

Sub Penepma08RunPenepma(tForm As Form)
' Run the selected input file in Penepma08

ierror = False
On Error GoTo Penepma08RunPenepmaError

Dim bfilename As String, astring As String

icancelauto = False

' Clear the data arrays
nPoints& = 0

' Delete dump files
Call Penepma08DeleteDumpFiles
If ierror Then Exit Sub

' Delete existing output files
If Dir$(PENEPMA_DAT_File$) <> vbNullString Then Kill PENEPMA_DAT_File$
If Dir$(PENEPMA_SPEC_File$) <> vbNullString Then Kill PENEPMA_SPEC_File$
If Dir$(PENEPMA_CHAR_File$) <> vbNullString Then Kill PENEPMA_CHAR_File$
If Dir$(PENEPMA_EL_TRANS_File$) <> vbNullString Then Kill PENEPMA_EL_TRANS_File$
Sleep (1000)

' Create batch file to run Penepma
bfilename$ = PENEPMA_Path$ & "\temp.bat"
Open bfilename$ For Output As #Temp1FileNumber%

astring$ = Left$(PENEPMA_Path$, 2)                             ' change to drive
Print #Temp1FileNumber%, astring$
astring$ = "cd " & VbDquote$ & PENEPMA_Path$ & VbDquote$       ' change to folder
Print #Temp1FileNumber%, astring$
astring$ = "Penepma " & Chr$(60) & " " & VbDquote$ & InputFile$ & VbDquote$
Print #Temp1FileNumber%, astring$
Close #Temp1FileNumber%

' Start Penepma (/k executes but window remains, /c executes but terminates)
'PenepmaTaskID& = Shell("cmd.exe /k " & VbDquote$ & bfilename$ & VbDquote$, vbNormalFocus)
PenepmaTaskID& = Shell("cmd.exe /c " & VbDquote$ & bfilename$ & VbDquote$, vbNormalFocus)

' Clear fields
tForm.TextElapsedTime.Text = vbNullString
tForm.TextElapsedShowers.Text = vbNullString
tForm.TextElapsedBSE.Text = vbNullString
Call Penepma08GraphClear
If ierror Then Exit Sub
DoEvents

' Set timer to update data from output files
tForm.Timer1.Interval = PENEPMA_DISPLAY_SEC# * MSECPERSEC#

' Set enables
tForm.CommandOK.Enabled = False
tForm.CommandClose.Enabled = False
tForm.CommandRunPENEPMA.Enabled = False
tForm.CommandOutputMaterial.Enabled = False
tForm.CommandOutputFormula.Enabled = False
tForm.CommandOutputWeight.Enabled = False
tForm.CommandOutputInputFile.Enabled = False
tForm.CommandDeleteDumpFiles.Enabled = False
tForm.CommandBatch.Enabled = False

tForm.TextInputTitle.Enabled = False
tForm.TextBeamTakeoff.Enabled = False
tForm.TextBeamEnergy.Enabled = False
tForm.TextBeamPosition(0).Enabled = False
tForm.TextBeamPosition(1).Enabled = False
tForm.TextBeamPosition(2).Enabled = False
tForm.TextBeamDirection(0).Enabled = False
tForm.TextBeamDirection(1).Enabled = False
tForm.TextBeamAperture.Enabled = False
tForm.TextDumpPeriod.Enabled = False
tForm.TextEnergyRangeMinMaxNumber(0).Enabled = False
tForm.TextEnergyRangeMinMaxNumber(1).Enabled = False
tForm.TextEnergyRangeMinMaxNumber(2).Enabled = False
tForm.TextNumberSimulatedShowers.Enabled = False
tForm.TextSimulationTimePeriod.Enabled = False
tForm.OptionProduction(0).Enabled = False
tForm.OptionProduction(1).Enabled = False
tForm.OptionProduction(2).Enabled = False
tForm.OptionProduction(3).Enabled = False
tForm.OptionProduction(4).Enabled = False
tForm.TextMaterialFiles(0).Enabled = False
tForm.TextMaterialFiles(1).Enabled = False
tForm.TextEABS1(0).Enabled = False
tForm.TextEABS1(1).Enabled = False
tForm.TextEABS2(0).Enabled = False
tForm.TextEABS2(1).Enabled = False
tForm.CommandBrowseMaterialFiles(0).Enabled = False
tForm.CommandBrowseMaterialFiles(1).Enabled = False
tForm.UpDownXray(0).Enabled = False
tForm.UpDownXray(1).Enabled = False
tForm.CommandAdjust(0).Enabled = False
tForm.CommandAdjust(1).Enabled = False
tForm.CommandElement(0).Enabled = False
tForm.CommandElement(1).Enabled = False
tForm.TextGeometryFile.Enabled = False
tForm.CommandBrowseGeometry.Enabled = False
tForm.TextInputFile.Enabled = False
tForm.CommandBrowseInputFiles.Enabled = False

FormPENEPMA08Batch.CommandClose.Enabled = False
FormPENEPMA08Batch.ListInputFiles.Enabled = False
FormPENEPMA08Batch.CommandRunBatch.Enabled = False
FormPENEPMA08Batch.CommandReload.Enabled = False
FormPENEPMA08Batch.CommandBrowseBatchFolder.Enabled = False
DoEvents

SimulationInProgress = True
PenepmaTimeStart = Now
tForm.LabelProgress.Caption = "Simulation In Progress!"
DoEvents
Exit Sub

' Errors
Penepma08RunPenepmaError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08RunPenepma"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08RunPenepmaBatch(tForm As Form)
' Run the selected input files in Penepma0812 (batch mode)

ierror = False
On Error GoTo Penepma08RunPenepmaBatchError

Dim i As Integer, response As Integer

icancelauto = False

' Save batch project folder
PENEPMA_BATCH_FOLDER$ = Trim$(FormPENEPMA08Batch.TextBatchFolder.Text)
If Dir$(PENEPMA_BATCH_FOLDER$, vbDirectory) = vbNullString Then
msg$ = "Batch Project Directory " & PENEPMA_BATCH_FOLDER$ & " is invalid. Would you like Standard to create the folder for you?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton1, "Penepma08RunPenepmaBatch")
If response% = vbYes Then
MkDir PENEPMA_BATCH_FOLDER$
Else
msg$ = "Please use the Browse button to create the folder " & PENEPMA_BATCH_FOLDER$ & " manually and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08RunPenepmaBatch"
ierror = True
Exit Sub
End If
End If

' Go through list of input files and create batch file for each
For i% = 0 To FormPENEPMA08Batch.ListInputFiles.ListCount - 1
If FormPENEPMA08Batch.ListInputFiles.Selected(i%) Then

' Load input file
InputFile$ = InputFiles$(i% + 1)
FormPENEPMA08Batch.LabelCurrentInputFile.Caption = "Currently running batch simulation file " & InputFile$ & "..."
DoEvents

' Load parameters to main form
Call Penepma08LoadInputFile(PENEPMA_Path$ & "\" & InputFile$, tForm)
If ierror Then Exit Sub
tForm.TextInputFile.Text = InputFile$

' Save the new parameters as default
Call Penepma08SaveInput(FormPENEPMA08_PE)
If ierror Then Exit Sub

' Load parameters to batch form
Call Penepma08BatchGetInputParameters(i%)
If ierror Then Exit Sub

' Run Penepma
Call Penepma08RunPenepma(tForm)
If ierror Then Exit Sub

' Wait for simulation to finish
Do While SimulationInProgress
Call Penepma08CheckTermination(tForm)
If ierror Then Exit Sub
Call MiscDelay5(PENEPMA_DISPLAY_SEC#, Now)      ' no hourglass
If icancelauto Or ierror Then Exit Sub
Loop

tForm.LabelElapsedTime.Caption = vbNullString

' Save the output to the project folder and sub-folder based on the input file name
If Dir$(PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$), vbDirectory) = vbNullString Then  ' make sure it exists
MkDir PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$)
End If

' Copy files to project folder\sub folder
Call Penepma08RunPenepmaCopy
If ierror Then Exit Sub

' Confirm for user
msg$ = "Penepma input file " & InputFile$ & " simulation is complete and output saved to the " & PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & " folder."
Call IOWriteLog(msg$)

End If
Next i%

FormPENEPMA08Batch.LabelCurrentInputFile.Caption = "All batch simulations completed!"
DoEvents

' Confirm with user
Call IOStatusAuto(vbNullString)
msg$ = "All batch mode input files were simulated and output saved to the " & PENEPMA_BATCH_FOLDER$ & " folders."
MsgBox msg$, vbOKOnly + vbInformation, "Penepma08RunPenepmaBatch"

Exit Sub

' Errors
Penepma08RunPenepmaBatchError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08RunPenepmaBatch"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08RunPenepmaCopy()
' Copy the output (and input) files in Penepma08/12 to the project folder

ierror = False
On Error GoTo Penepma08RunPenepmaCopyError

Dim i As Integer
Dim tfilename As String

' Copy files to project folder\sub folder
tfilename$ = "pe-spect-01.dat"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$
tfilename$ = "penepma.dat"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$

' Photons (Penepma08)
tfilename$ = "pe-energy-ph-trans.dat"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$
tfilename$ = "pe-energy-ph-back.dat"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$

' Photons (Penepma12)
tfilename$ = "pe-energy-ph-up.dat"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$
tfilename$ = "pe-energy-ph-down.dat"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$

' Electrons (Penepma08)
tfilename$ = "pe-energy-el-trans.dat"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$
tfilename$ = "pe-energy-el-back.dat"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$

' Electrons (Penepma12)
tfilename$ = "pe-energy-el-up.dat"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$
tfilename$ = "pe-energy-el-down.dat"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$

' Characteristic spectrum
tfilename$ = "pe-charact-01.dat"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$

' Generated spectrums
tfilename$ = "pe-gen-bremms.dat"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$
tfilename$ = "pe-gen-ph.dat"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$

tfilename$ = "pe-angle-ph.dat"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$
tfilename$ = "pe-angle-el.dat"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$
tfilename$ = "pe-anga.dat"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$
tfilename$ = "pe-anel.dat"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$

tfilename$ = "pengeom-tree.rep"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$
tfilename$ = "pe-material.dat"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$
tfilename$ = "pe-geometry.rep"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$

' Continuum
tfilename$ = "pe-gen-bremss.dat"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$

' Copy net intensity file also (for net k-ratios for all emission lines) (changed from pe-inten-01.dat to pe-intens-01.dat 11-2-2012)
tfilename$ = "pe-intens-01.dat"
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$

' Copy the input file too
tfilename$ = InputFile$
If Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$

' Read the geometry and material files to copy
Call Penepma08LoadInput(PENEPMA_Path$ & "\" & InputFile$, "GEOMFN", tfilename$, Int(0))
If ierror Then Exit Sub
If tfilename$ <> vbNullString And Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then
FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$
End If

For i% = 1 To MAXMATOUTPUT%
Call Penepma08LoadInput(PENEPMA_Path$ & "\" & InputFile$, "MFNAME", tfilename$, i%)
If ierror Then Exit Sub
If tfilename$ <> vbNullString And Dir$(PENEPMA_Path$ & "\" & tfilename$) <> vbNullString Then
FileCopy PENEPMA_Path$ & "\" & tfilename$, PENEPMA_BATCH_FOLDER$ & "\" & MiscGetFileNameNoExtension$(InputFile$) & "\" & tfilename$
End If
Next i%

Exit Sub

' Errors
Penepma08RunPenepmaCopyError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08RunPenepmaCopy"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Function Penepma08ReturnMaterialFile(stdnum As Integer) As String
' Return standard material file name based on standard database number

ierror = False
On Error GoTo Penepma08ReturnMaterialFileError

' Get standard from database
Call StandardGetMDBStandard(stdnum%, PENEPMASample())
If ierror Then Exit Function

Call MiscModifyStringToFilename(PENEPMASample(1).Name$)
If ierror Then Exit Function
Penepma08ReturnMaterialFile$ = Trim$(Left$(PENEPMASample(1).Name$, 16)) & ".MAT"   ' filename maximum 20 characters

Exit Function

' Errors
Penepma08ReturnMaterialFileError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08ReturnMaterialFile"
ierror = True
Exit Function

End Function

Sub Penepma08ReadMaterialFile(tMaterialFile As String, sample() As TypeSample)
' Read the material file and return composition

ierror = False
On Error GoTo Penepma08ReadMaterialFileError

Dim i As Integer, n As Integer
Dim tfilename As String, astring As String

ReDim atoms(1 To MAXCHAN%) As Single

' Clear sample
Call InitSample(sample())
If ierror Then Exit Sub

If tMaterialFile$ = vbNullString Then GoTo Penepma08ReadMaterialFileBlankFile
tfilename$ = PENEPMA_Path$ & "\" & tMaterialFile$
If Dir$(tfilename$) = vbNullString Then GoTo Penepma08ReadMaterialFileNoFIle
Open tfilename$ For Input As #Temp1FileNumber%

' Locate number of elements
Do Until InStr(astring$, "Number of elements") > 0
Line Input #Temp1FileNumber%, astring$
Loop

' Get number of elements
sample(1).LastChan% = Val(Right$(astring$, 2))
For i% = 1 To sample(1).LastChan%
Line Input #Temp1FileNumber%, astring$

' Locate atomic number
n% = InStr(astring$, "atomic number = ") + Len("atomic number = ")
sample(1).Elsyms$(i%) = Symlo$(Val(Mid$(astring$, n%, 2)))
atoms!(i%) = Val(Right$(astring$, 14))
Next i%
Close #Temp1FileNumber%

' Convert to weight percent
For i% = 1 To sample(1).LastChan%
sample(1).ElmPercents!(i%) = ConvertAtomToWeight!(sample(1).LastChan%, i%, atoms!(), sample(1).Elsyms$())
Next i%

Exit Sub

' Errors
Penepma08ReadMaterialFileError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08ReadMaterialFile"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08ReadMaterialFileBlankFile:
msg$ = "No material file selected for output"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08ReadMaterialFile"
ierror = True
Exit Sub

Penepma08ReadMaterialFileNoFIle:
msg$ = "Specified material file was not found in the path " & PENEPMA_Path$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08ReadMaterialFile"
ierror = True
Exit Sub

End Sub

Sub Penepma08AdjustEABS(mode As Integer, Index As Integer, tForm As Form)
' Load electron/photon absorption parameters based on mode (production option) and current material file

ierror = False
On Error GoTo Penepma08AdjustEABSError

Dim ip As Integer, ipp As Integer
Dim iwmin As Long, iwmax As Long
Dim eabs1 As Single, eabs2 As Single
Dim twmin As Single, twmax As Single

Dim ielm As Integer, iray As Integer
Dim energy As Single, edge As Single

Const JUSTUNDER! = 0.98
Const JUSTUNDER2! = 0.9

' Get composition
If mode% = MAXPRODUCTION% + 1 Then

' Get user element selections from periodic table
Call Penepma08GetElement(PENEPMASample())
If ierror Then Exit Sub

' Get elements from current material file
Else
Call Penepma08ReadMaterialFile(MaterialFiles$(Index% + 1), PENEPMASample())
If ierror Then Exit Sub
End If

' Check for at least one element
If PENEPMASample(1).LastChan% < 1 Then GoTo Penepma08AdjustEABSZeroElements

' Get the weight percent min and max
Call MiscGetArrayMinMax(CLng(PENEPMASample(1).LastChan%), PENEPMASample(1).ElmPercents!(), iwmin&, iwmax&, twmin!, twmax!)
If ierror Then Exit Sub

' Check that at least a major element was found
If iwmax& = 0 Then GoTo Penepma08AdjustEABSNoMajorElement

' Get the default x-ray for the major element
ip% = IPOS1%(MAXELM%, PENEPMASample(1).Elsyms$(iwmax&), Symlo$())
If ip% = 0 Then GoTo Penepma08AdjustEABSElementNotFound
ielm% = ip%
ipp% = IPOS1%(MAXRAY% - 1, Deflin$(ip%), Xraylo$())
iray% = ipp%

' Check xray adjust number
If iray% + XrayAdjustNumber%(Index% + 1) < 1 Then XrayAdjustNumber%(Index% + 1) = 1 - iray%
If iray% + XrayAdjustNumber%(Index% + 1) > MAXRAY% - 1 Then XrayAdjustNumber%(Index% + 1) = (MAXRAY% - 1) - iray%
iray% = iray% + XrayAdjustNumber%(Index% + 1)

' Get energy and edge of x-ray
Call XrayGetEnergy(ielm%, iray%, energy!, edge!)
If ierror Then Exit Sub

' Check for valid energey and edge
If energy! = 0# Or edge! = 0# Then Exit Sub

' Mode = 0 (characteristic x-rays)
If mode% = 0 Then
eabs1! = energy! * EVPERKEV# * JUSTUNDER!       ' minimum photon absorption energy (set equal in penepma)
eabs2! = energy! * EVPERKEV# * JUSTUNDER!       ' minimum electron energy (set equal in penepma)

' Mode = 1 (backscatter electrons)
ElseIf mode% = 1 Then
eabs1! = 50#
eabs2! = Val(tForm.TextBeamEnergy.Text) * JUSTUNDER2!

' Mode = 2 (continuum x-ray)
ElseIf mode% = 2 Then
eabs1! = 1000#
eabs2! = 1000#

' Mode = 3 (secondary fluorescent x-rays)
ElseIf mode% = 3 Then
eabs1! = 1000#
eabs2! = 1000#

' Mode = 4 (thin film)
ElseIf mode% = 4 Then
eabs1! = 1000#
eabs2! = 1000#

' Mode = 5 (selected element(s)) (skipped during Adjust button clicks)
ElseIf mode% = 5 Then
eabs1! = energy! * EVPERKEV# * JUSTUNDER!
eabs2! = energy! * EVPERKEV# * JUSTUNDER!

End If

' Load parameters
tForm.TextEABS1(Index%).Text = Format$(eabs1!, "Scientific")
tForm.TextEABS2(Index%).Text = Format$(eabs2!, "Scientific")

Exit Sub

' Errors
Penepma08AdjustEABSError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08AdjustEABS"
ierror = True
Exit Sub

Penepma08AdjustEABSZeroElements:
msg$ = "No elements specified"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08AdjustEABS"
ierror = True
Exit Sub

Penepma08AdjustEABSNoMajorElement:
msg$ = "No major element is specified"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08AdjustEABS"
ierror = True
Exit Sub

Penepma08AdjustEABSElementNotFound:
msg$ = "Element " & PENEPMASample(1).Elsyms$(iwmax&) & " is invalid"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08AdjustEABS"
ierror = True
Exit Sub

End Sub

Sub Penepma08GetElement(sample() As TypeSample)
' Return the user selected element(s) from the periodic table

ierror = False
On Error GoTo Penepma08GetElementError

Dim i As Integer
Dim elmarray(1 To MAXELM) As Boolean

' Clear sample
Call InitSample(sample())
If ierror Then Exit Sub

' Load form
icancelload = False
Call Periodic2Load
If icancelload = True Then
ierror = True
Exit Sub
End If

' Get selected elements
Call Periodic2Return(elmarray())
If ierror Then Exit Sub

' Load sample
For i% = 1 To MAXCHAN%
If elmarray(i%) Then
sample(1).LastChan% = sample(1).LastChan% + 1
sample(1).Elsyms$(sample(1).LastChan%) = Symlo$(i%)
End If
Next i%

' Check for at least one element
If sample(1).LastChan% < 1 Then
ierror = True
Exit Sub
End If

' Load nominal weight percents
For i% = 1 To sample(1).LastChan%
sample(1).ElmPercents!(i%) = 100# / sample(1).LastChan%
Next i%

Exit Sub

' Errors
Penepma08GetElementError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08GetElement"
ierror = True
Exit Sub

End Sub

Sub Penepma08XrayAdjust(n As Integer, k As Integer)
' Increment or decrement the xray adjust number

ierror = False
On Error GoTo Penepma08XrayAdjustError

XrayAdjustNumber%(k%) = XrayAdjustNumber%(k%) + n%

Exit Sub

' Errors
Penepma08XrayAdjustError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08XrayAdjust"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08LoadPenepmaDAT(tForm As Form)
' Load specified values from PENEPMA.DAT file for display

ierror = False
On Error GoTo Penepma08LoadPenepmaDATError

Dim astring As String, bstring As String
Dim ntry As Integer
Dim nlinenumber As Long
Dim cstring As String

' Check production file
Close #Temp2FileNumber%
DoEvents

cstring$ = "Penepma08LoadPenepmaDAT: Checking " & PENEPMA_DAT_File$ & "..."
If Trim$(PENEPMA_DAT_File$) = vbNullString Then Exit Sub
If Dir$(Trim$(PENEPMA_DAT_File$)) = vbNullString Then Exit Sub

' Set tolerant file open handling
On Error GoTo Penepma08LoadPenepmaDATOpenWait
cstring$ = "Penepma08LoadPenepmaDAT: Opening " & PENEPMA_DAT_File$ & "..."

' Open file and load values
Penepma08LoadPenepmaDATOpenTryAgain:
ntry% = ntry% + 1
Open PENEPMA_DAT_File$ For Input As #Temp2FileNumber%
GoTo Penepma08LoadPenepmaDATOpenProceed

Penepma08LoadPenepmaDATOpenWait:
Call MiscDelay3(FormMAIN.StatusBarAuto, "next Penepma data read...", CDbl(3#), Now)      ' wait 3 seconds and try again
If ierror Then Exit Sub

' Check for too many tries
If ntry% > MAXTRIES% Then
msg$ = vbCrLf & "Penepma08LoadPenepmaDATOpen: Unable to open file " & PENEPMA_DAT_File$ & " after " & Format$(ntry%) & " attempts."
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
ierror = True
Exit Sub
End If

' Try again
Resume Penepma08LoadPenepmaDATOpenTryAgain

' Database opened, go ahead and exit normally
Penepma08LoadPenepmaDATOpenProceed:
On Error GoTo Penepma08LoadPenepmaDATError
cstring$ = "Penepma08LoadPenepmaDAT: Reading " & PENEPMA_DAT_File$ & "..."

nlinenumber& = 0
Do Until EOF(Temp2FileNumber%)
Line Input #Temp2FileNumber%, astring$
nlinenumber& = nlinenumber& + 1
cstring$ = "Penepma08LoadPenepmaDAT: Reading " & PENEPMA_DAT_File$ & " (line= " & Format$(nlinenumber&) & ")..."

' Load specific data for current graph
If nlinenumber& = 9 Then BeamTitle$ = Trim$(astring$)
If InStr(astring$, "Initial energy =") > 0 Then BeamEnergy# = Val(Mid$(astring$, InStr(astring$, "Initial energy =") + Len("Initial energy =") + 2, 14))

' Load parameters to text fields
If InStr(astring$, "Simulation time") > 0 Then Call Penepma08LoadPenepmaDAT2(astring$, bstring$, tForm.TextElapsedTime)
If InStr(astring$, "Simulated primary showers") > 0 Then Call Penepma08LoadPenepmaDAT2(astring$, bstring$, tForm.TextElapsedShowers)

' Load BSE
If Penepma08CheckPenepmaVersion%() = 8 Then
If InStr(astring$, "Fractional transmission") > 0 Then Call Penepma08LoadPenepmaDAT2(astring$, bstring$, tForm.TextElapsedBSE)
Else
If InStr(astring$, "Upbound fraction") > 0 Then Call Penepma08LoadPenepmaDAT2(astring$, bstring$, tForm.TextElapsedBSE)
End If

DoEvents        ' to allow text fields to update
Loop

Close #Temp2FileNumber%
DoEvents        ' to allow text fields to update
Exit Sub

' Errors
Penepma08LoadPenepmaDATError:
MsgBox Error$ & ", " & cstring$, vbOKOnly + vbCritical, "Penepma08LoadPenepmaDAT"
Close #Temp2FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08LoadPenepmaDAT2(astring As String, bstring As String, tText As TextBox)
' Load the text to the passed text box (single control)

ierror = False
On Error GoTo Penepma08LoadPenepmaDAT2Error

' Load the parameter text (starting after last period and 14 characters long)
bstring$ = Mid$(astring$, MiscInstr(astring$, "..") + 2, 14)
tText.Text = Trim$(bstring$)

Exit Sub

' Errors
Penepma08LoadPenepmaDAT2Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08LoadPenepmaDAT2"
Close #Temp2FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08GetPenepmaDAT()
' Load specified values from PENEPMA.DAT file for graph title and x axis range

ierror = False
On Error GoTo Penepma08GetPenepmaDATError

Dim numtries As Integer
Dim astring As String
Dim nlinenumber As Long

' Check production file
If Dir$(Trim$(PENEPMA_DAT_File$)) = vbNullString Then Exit Sub

' Allow retry on file operations
Penepma08GetPenepmaDATTryAgain:
numtries% = numtries% + 1
DoEvents
Sleep 100
If numtries% > MAXTRIES% Then GoTo Penepma08GetPenepmaDATBadFile
On Error GoTo Penepma08GetPenepmaDATTryAgain

' Open file and load values
Open PENEPMA_DAT_File$ For Input As #Temp1FileNumber%

nlinenumber& = 0
Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$
nlinenumber& = nlinenumber& + 1

' Load specific data for current graph
If nlinenumber& = 9 Then BeamTitle$ = Trim$(astring$)
If InStr(astring$, "Initial energy =") > 0 Then BeamEnergy# = Val(Mid$(astring$, InStr(astring$, "Initial energy =") + Len("Initial energy =") + 2, 14))

Loop
Close #Temp1FileNumber%

' Restore normal error trap
On Error GoTo Penepma08GetPenepmaDATError
Exit Sub

' Errors
Penepma08GetPenepmaDATError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08GetPenepmaDAT"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08GetPenepmaDATBadFile:
msg$ = "Unable to open and/or read " & PENEPMA_DAT_File$ & " after " & Format$(numtries%) & " attempts. Will exit with error."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08GetPenepmaDAT"
ierror = True
Exit Sub

End Sub

Sub Penepma08GraphGetData(Index As Integer)
' Load spectrum data for the specified graph

ierror = False
On Error GoTo Penepma08GraphGetDataError

Dim astring As String, bstring As String

' Check production file
If Index% = 0 And Dir$(Trim$(PENEPMA_SPEC_File$)) = vbNullString Then Exit Sub
If Index% = 1 And Dir$(Trim$(PENEPMA_CHAR_File$)) = vbNullString Then Exit Sub
If Index% = 2 And Dir$(Trim$(PENEPMA_EL_TRANS_File$)) = vbNullString Then Exit Sub

' Close file in case it is already open
Close #Temp1FileNumber%
DoEvents

' Open file and load values
If Index% = 0 Then Open PENEPMA_SPEC_File$ For Input As #Temp1FileNumber%
If Index% = 1 Then Open PENEPMA_CHAR_File$ For Input As #Temp1FileNumber%
If Index% = 2 Then Open PENEPMA_EL_TRANS_File$ For Input As #Temp1FileNumber%

If DebugMode Then
If Index% = 0 Then Call IOWriteLog(vbCrLf & "Spectrum File: " & PENEPMA_SPEC_File$)
If Index% = 1 Then Call IOWriteLog(vbCrLf & "Spectrum File: " & PENEPMA_CHAR_File$)
If Index% = 2 Then Call IOWriteLog(vbCrLf & "Spectrum File: " & PENEPMA_EL_TRANS_File$)
End If

' Load array (npts&, xdata#(), ydata#())
nPoints& = 0
Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$
If Len(Trim$(astring$)) > 0 And InStr(astring$, "#") = 0 Then            ' skip to first data line

' Load total spectrum
If Index% = 0 Then
nPoints& = nPoints& + 1
Call MiscParseStringToStringT(astring$, bstring$)
If ierror Then
Close #Temp1FileNumber%
ierror = False
Exit Sub
End If
ReDim Preserve xdata(1 To nPoints&) As Double
ReDim Preserve ydata(1 To nPoints&) As Double
xdata#(nPoints&) = Val(Trim$(bstring$))
Call MiscParseStringToStringT(astring$, bstring$)
If ierror Then
Close #Temp1FileNumber%
ierror = False
Exit Sub
End If
ydata#(nPoints&) = Val(Trim$(bstring$))

' Load characteristic spectrum
ElseIf Index% = 1 Then
nPoints& = nPoints& + 1
Call MiscParseStringToStringT(astring$, bstring$)
If ierror Then
Close #Temp1FileNumber%
ierror = False
Exit Sub
End If
ReDim Preserve xdata(1 To nPoints&) As Double
ReDim Preserve ydata(1 To nPoints&) As Double
xdata#(nPoints&) = Val(Trim$(bstring$))
Call MiscParseStringToStringT(astring$, bstring$)
If ierror Then
Close #Temp1FileNumber%
ierror = False
Exit Sub
End If
ydata#(nPoints&) = Val(Trim$(bstring$))

' Load backscatter energy spectrum
ElseIf Index% = 2 Then
nPoints& = nPoints& + 1
If nPoints& > 1 Then        ' skip first blank line
Call MiscParseStringToStringT(astring$, bstring$)
If ierror Then
Close #Temp1FileNumber%
ierror = False
Exit Sub
End If
ReDim Preserve xdata(1 To nPoints&) As Double
ReDim Preserve ydata(1 To nPoints&) As Double
xdata#(nPoints&) = Val(Trim$(bstring$))
Call MiscParseStringToStringT(astring$, bstring$)
If ierror Then
Close #Temp1FileNumber%
ierror = False
Exit Sub
End If
ydata#(nPoints&) = Val(Trim$(bstring$))
End If
End If

If DebugMode Then
Call IOWriteLog("N=" & Format$(nPoints&) & ", X=" & Format$(xdata#(nPoints&)) & ", Y=" & Format$(ydata#(nPoints&), e104$))
End If

End If
Loop

Close #Temp1FileNumber%
Exit Sub

' Errors
Penepma08GraphGetDataError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08GraphGetData"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08SaveDisplay(tForm As Form)
' Save the form for the graph display options

ierror = False
On Error GoTo Penepma08SaveDisplayError

Dim i As Integer

icancelauto = False

' Save display graph index
For i% = 0 To 2
If tForm.OptionDisplayGraph(i%).Value = True Then GraphDisplayOption% = i%
Next i%

' Save gridlines
If tForm.CheckUseGridLines.Value = vbChecked Then
UseGridLines = True
Else
UseGridLines = False
End If

If tForm.CheckUseLogScale.Value = vbChecked Then
UseLogScale = True
Else
UseLogScale = False
End If

Exit Sub

' Errors
Penepma08SaveDisplayError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08SaveDisplay"
ierror = True
Exit Sub

End Sub

Sub Penepma08CheckTermination(tForm As Form)
' Check if Penepma process is terminated

ierror = False
On Error GoTo Penepma08CheckTerminationError

If SimulationInProgress Then

' Check for termination
If IOIsProcessTerminated(PenepmaTaskID&) Then
SimulationInProgress = False
tForm.LabelProgress.Caption = "Simulation Completed!"
tForm.LabelElapsedTime.Caption = vbNullString
tForm.Timer1.Interval = 0
DoEvents

' Delete dump files and update form
Call Penepma08CheckTermination2(tForm)
If ierror Then Exit Sub

' Still running
Else

' Check for user cancel
DoEvents
If icancelauto Then
Call Penepma08CheckTermination2(tForm)
If ierror Then Exit Sub
Call IOStatusAuto(vbNullString)
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

tForm.LabelElapsedTime.Caption = "Time: " & MiscConvertTimeToClockString$(Now - PenepmaTimeStart)
DoEvents
End If
End If

Exit Sub

' Errors
Penepma08CheckTerminationError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08CheckTermination"
ierror = True
Exit Sub

End Sub

Sub Penepma08CheckTermination2(tForm As Form)
' Clean up after completion or user cancel

ierror = False
On Error GoTo Penepma08CheckTermination2Error

' Delete dump files
Call Penepma08DeleteDumpFiles
If ierror Then Exit Sub

' Final display update
Call Penepma08LoadPenepmaDAT(tForm)
If ierror Then Exit Sub
Call Penepma08GraphUpdate(GraphDisplayOption%)
If ierror Then Exit Sub

' Set enables
tForm.CommandOK.Enabled = True
tForm.CommandClose.Enabled = True
tForm.CommandRunPENEPMA.Enabled = True
tForm.CommandOutputMaterial.Enabled = True
tForm.CommandOutputFormula.Enabled = True
tForm.CommandOutputWeight.Enabled = True
tForm.CommandOutputInputFile.Enabled = True
tForm.CommandDeleteDumpFiles.Enabled = True
tForm.CommandBatch.Enabled = True

tForm.TextInputTitle.Enabled = True
tForm.TextBeamTakeoff.Enabled = True
tForm.TextBeamEnergy.Enabled = True
tForm.TextBeamPosition(0).Enabled = True
tForm.TextBeamPosition(1).Enabled = True
tForm.TextBeamPosition(2).Enabled = True
tForm.TextBeamDirection(0).Enabled = True
tForm.TextBeamDirection(1).Enabled = True
tForm.TextBeamAperture.Enabled = True
tForm.TextDumpPeriod.Enabled = True
tForm.TextEnergyRangeMinMaxNumber(0).Enabled = True
tForm.TextEnergyRangeMinMaxNumber(1).Enabled = True
tForm.TextEnergyRangeMinMaxNumber(2).Enabled = True
tForm.TextNumberSimulatedShowers.Enabled = True
tForm.TextSimulationTimePeriod.Enabled = True

tForm.OptionProduction(0).Enabled = True
tForm.OptionProduction(1).Enabled = True
tForm.OptionProduction(2).Enabled = True
tForm.OptionProduction(3).Enabled = True
tForm.OptionProduction(4).Enabled = True
tForm.TextMaterialFiles(0).Enabled = True
tForm.TextMaterialFiles(1).Enabled = True
tForm.CommandBrowseMaterialFiles(0).Enabled = True
tForm.CommandBrowseMaterialFiles(1).Enabled = True
tForm.TextGeometryFile.Enabled = True
tForm.CommandBrowseGeometry.Enabled = True
tForm.TextInputFile.Enabled = True
tForm.CommandBrowseInputFiles.Enabled = True

FormPENEPMA08Batch.CommandClose.Enabled = True
FormPENEPMA08Batch.ListInputFiles.Enabled = True
FormPENEPMA08Batch.CommandRunBatch.Enabled = True
FormPENEPMA08Batch.CommandReload.Enabled = True
FormPENEPMA08Batch.CommandBrowseBatchFolder.Enabled = True

tForm.LabelElapsedTime.Caption = vbNullString
tForm.LabelProgress.Caption = vbNullString

'tForm.TextEABS1(0).Enabled = True
'tForm.TextEABS1(1).Enabled = True
'tForm.TextEABS2(0).Enabled = True
'tForm.TextEABS2(1).Enabled = True
'tForm.UpDownXray(0).Enabled = True
'tForm.UpDownXray(1).Enabled = True
'tForm.CommandAdjust(0).Enabled = True
'tForm.CommandAdjust(1).Enabled = True
'tForm.CommandElement(0).Enabled = True
'tForm.CommandElement(1).Enabled = True

' Set production control enables
Call Penepma08SetOptionProductionEnables(CInt(BeamProductionIndex&))
If ierror Then Exit Sub

Exit Sub

' Errors
Penepma08CheckTermination2Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08CheckTermination2"
ierror = True
Exit Sub

End Sub

Sub Penepma08CreateMaterialFormula(mode As Integer, tForm As Form)
' Creates a material file from a formula or weight composition
'  mode = 1 = formula string
'  mode = 2 = weight string

ierror = False
On Error GoTo Penepma08CreateMaterialFormulaError

Dim n As Integer, i As Integer

icancelauto = False

' Load FormFORMULA
If mode% = 1 Then FormFORMULA.Frame1.Caption = "Enter Formula String For PENEPMA Material File:"
If mode% = 2 Then FormWEIGHT.Frame1.Caption = "Enter Weight percent String For PENEPMA Material File:"

' Get formula from user
If mode% = 1 Then
FormFORMULA.Show vbModal
If icancel Then Exit Sub
End If

' Get weight string from user
If mode% = 2 Then
FormWEIGHT.Show vbModal
If icancel Then Exit Sub
End If

' Return modified sample (in weight percents)
Call FormulaReturnSample(PENEPMASample())
If ierror Then Exit Sub

' Load atm numbers, etc
Call ElementGetData(PENEPMASample())
If ierror Then Exit Sub

If DebugMode Then
Call IOWriteLog("Penepma08CreateMaterialFormula: " & PENEPMASample(1).Name$)
For i% = 1 To PENEPMASample(1).LastChan%
Call IOWriteLog("Penepma08CreateMaterialFormula: " & PENEPMASample(1).Elsyms$(i%) & " at " & PENEPMASample(1).ElmPercents!(i%) & " wt. %")
Next i%
End If

' Specify a single material
For n% = 1 To MAXMATOUTPUT%
MaterialFiles$(n%) = vbNullString
MaterialsSelected%(n%) = 0
Next n%

' Load name and number for this formula
Call MiscModifyStringToFilename(PENEPMASample(1).Name$)
If ierror Then Exit Sub

' Load suffix if not pure element
If IPOS1%(MAXELM%, PENEPMASample(1).Name$, Symlo$()) > 0 Then
MaterialFiles$(1) = Trim$(PENEPMASample(1).Name$) & "_100.MAT"
Else
If mode% = 1 Then MaterialFiles$(1) = Trim$(PENEPMASample(1).Name$) & "_atom.MAT"
If mode% = 2 Then MaterialFiles$(1) = Trim$(PENEPMASample(1).Name$) & "_weight.MAT"
End If
MaterialsSelected%(1) = MAXINTEGER%     ' any non-zero number

Call IOStatusAuto("Creating material input file based on formula " & PENEPMASample(1).Name$ & "...")
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

' Make material INP file (always a single file)
Call Penepma08CreateMaterialINP(Int(1), PENEPMASample())
If ierror Then Exit Sub

' Create and run the necessary batch files
Call Penepma08CreateMaterialBatch(Int(0), tForm)
If ierror Then Exit Sub

Exit Sub

' Errors
Penepma08CreateMaterialFormulaError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08CreateMaterialFormula"
ierror = True
Exit Sub

End Sub

Sub Penepma08BatchBinaryCreate(tForm As Form)
' Create a series of binary composition input files based on the two elements

ierror = False
On Error GoTo Penepma08BatchBinaryCreateError

Dim ip As Integer, response As Integer
Dim esym As String

Dim k As Integer, n As Integer
Dim pfilename As String

Dim binarynames(1 To MAXBINARY%) As String
Dim atoms(1 To 2) As Single

icancelauto = False

' Check that user has all Penepma input parameters properly specified
msg$ = "Before creating the binary composition input files, be sure to check that not only is the correct batch folder for saving results properly specified, but that all Penepma input parameters are properly specified, for example, title of run (e.g., Bulk Fe-Ni, 15 keV), beam energy, number of showers, simulation time, etc..." & vbCrLf & vbCrLf
msg$ = msg$ & Space(8) & "Penepma Batch Simulation Title:  " & BeamTitle$ & vbCrLf
msg$ = msg$ & Space(8) & "Penepma Batch Save To Folder:    " & PENEPMA_BATCH_FOLDER$ & vbCrLf
msg$ = msg$ & Space(8) & "Penepma Beam Energy (eV) :       " & Format$(BeamEnergy#, e71$) & " (" & Format$(BeamEnergy# / EVPERKEV#) & " keV)" & vbCrLf
msg$ = msg$ & Space(8) & "Penepma Number of Showers:       " & Format$(BeamNumberSimulatedShowers#, e71$) & vbCrLf
msg$ = msg$ & Space(8) & "Penepma Simulation Time (sec):   " & Format$(BeamSimulationTimePeriod#, e71$) & " (" & Format$(BeamSimulationTimePeriod# / SECPERHOUR#, f82$) & " hours per binary or " & Format$((MAXBINARY% + 2) * BeamSimulationTimePeriod# / SECPERDAY#, f82$) & " days for all " & Format$(MAXBINARY%) & " binaries plus two end-members)" & vbCrLf & vbCrLf
msg$ = msg$ & "Is everything ready to create the binary composition Penepma input files?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton1, "Penepma08BatchBinaryCreate")
If response% = vbNo Then
ierror = True
Exit Sub
End If

' Specify a single material
For n% = 1 To MAXMATOUTPUT%
MaterialFiles$(n%) = vbNullString
MaterialsSelected%(n%) = 0
Next n%

' Save the elements
esym$ = FormPENEPMA08Batch.ComboBinaryElement1.Text
ip% = IPOS1(MAXELM%, esym$, Symlo$())
If ip% = 0 Then GoTo Penepma08BatchBinaryCreateBadElement
BinaryElement1% = ip%

esym$ = FormPENEPMA08Batch.ComboBinaryElement2.Text
ip% = IPOS1(MAXELM%, esym$, Symlo$())
If ip% = 0 Then GoTo Penepma08BatchBinaryCreateBadElement
BinaryElement2% = ip%

' Check that binary elements are different
If BinaryElement1% = BinaryElement2% Then GoTo Penepma08BatchBinaryCreateSame

' Create pure element material files
For k% = 1 To 2
PENEPMASample(1).LastElm% = 1
PENEPMASample(1).LastChan% = PENEPMASample(1).LastElm%
If k% = 1 Then PENEPMASample(1).Elsyms$(1) = Symlo$(BinaryElement1%)
If k% = 1 Then PENEPMASample(1).Xrsyms$(1) = Deflin$(BinaryElement1%)  ' just load defaults here

If k% = 2 Then PENEPMASample(1).Elsyms$(1) = Symlo$(BinaryElement2%)
If k% = 2 Then PENEPMASample(1).Xrsyms$(1) = Deflin$(BinaryElement2%)  ' just load defaults here

' Load element data
Call ElementGetData(PENEPMASample())
If ierror Then Exit Sub

' Overload with Penepma08/12 atomic weights for self consistency in calculations
If k% = 1 Then PENEPMASample(1).AtomicWts!(1) = pAllAtomicWts!(BinaryElement1%)
If k% = 2 Then PENEPMASample(1).AtomicWts!(1) = pAllAtomicWts!(BinaryElement2%)

' Load element composition based on binary number
PENEPMASample(1).ElmPercents!(1) = 100#
If k% = 1 Then PENEPMASample(1).SampleDensity! = AllAtomicDensities!(BinaryElement1%)
If k% = 2 Then PENEPMASample(1).SampleDensity! = AllAtomicDensities!(BinaryElement2%)

' Load name and number for this binary
If k% = 1 Then pfilename$ = Trim$(Symup$(BinaryElement1%)) & "_" & Format$(PENEPMASample(1).ElmPercents!(1))
If k% = 2 Then pfilename$ = Trim$(Symup$(BinaryElement2%)) & "_" & Format$(PENEPMASample(1).ElmPercents!(1))
PENEPMASample(1).Name$ = pfilename$
MaterialFiles$(1) = PENEPMASample(1).Name$ & ".MAT"
MaterialsSelected%(1) = MAXINTEGER%     ' any non-zero number
MaterialDensity# = PENEPMASample(1).SampleDensity!

msg$ = "Creating material input file based on " & PENEPMASample(1).Name$ & "..."
Call IOWriteLog(msg$)
Call IOStatusAuto(msg$)
FormPENEPMA08Batch.LabelCurrentInputFile.Caption = msg$
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

' Make material INP file (always a single file)
Call Penepma08CreateMaterialINP(Int(1), PENEPMASample())
If ierror Then Exit Sub

' Create and run the necessary batch files
Call Penepma08CreateMaterialBatch(Int(1), tForm)
If ierror Then Exit Sub
Next k%

' Treat all binary materials as material A for .par file calculations and calculate
' 99:1, 95:5, 90:10, 80:20, 60:40, 50:50, 40:60, 20:80, 10:90, 5:95, 1:99 binaries

' Create binary sample
PENEPMASample(1).LastElm% = 2
PENEPMASample(1).LastChan% = PENEPMASample(1).LastElm%

PENEPMASample(1).Elsyms$(1) = Symlo$(BinaryElement1%)
PENEPMASample(1).Xrsyms$(1) = Deflin$(BinaryElement1%)  ' just load defaults here

PENEPMASample(1).Elsyms$(2) = Symlo$(BinaryElement2%)
PENEPMASample(1).Xrsyms$(2) = Deflin$(BinaryElement2%)  ' just load defaults here

' Load element data
Call ElementGetData(PENEPMASample())
If ierror Then Exit Sub

' Overload with Penepma08/12 atomic weights for self consistency in calculations
PENEPMASample(1).AtomicWts!(1) = pAllAtomicWts!(BinaryElement1%)
PENEPMASample(1).AtomicWts!(2) = pAllAtomicWts!(BinaryElement2%)

' Specify a single material
For n% = 1 To MAXMATOUTPUT%
MaterialFiles$(n%) = vbNullString
MaterialsSelected%(n%) = 0
Next n%

' Calculate material file for each compositional binary
For k% = 1 To MAXBINARY%

' Load element composition based on binary number
PENEPMASample(1).ElmPercents!(1) = BinaryRanges!(k%)
PENEPMASample(1).ElmPercents!(2) = 100# - BinaryRanges!(k%)

' Calculate density based on composition
Call ConvertWeightToAtomic(PENEPMASample(1).LastChan%, PENEPMASample(1).AtomicWts!(), PENEPMASample(1).ElmPercents!(), atoms!())
If ierror Then Exit Sub
PENEPMASample(1).SampleDensity! = atoms!(1) * AllAtomicDensities!(BinaryElement1%) + atoms!(2) * AllAtomicDensities!(BinaryElement2%)

' Load name and number for this binary
binarynames$(k%) = Trim$(Symup$(BinaryElement1%)) & "-" & Trim$(Symup$(BinaryElement2%)) & "_" & Format$(PENEPMASample(1).ElmPercents!(1)) & "-" & Format$(PENEPMASample(1).ElmPercents!(2))
PENEPMASample(1).Name$ = binarynames$(k%)
MaterialFiles$(1) = PENEPMASample(1).Name$ & ".MAT"
MaterialsSelected%(1) = MAXINTEGER%     ' any non-zero number
MaterialDensity# = PENEPMASample(1).SampleDensity!

msg$ = "Creating material input file based on binary " & PENEPMASample(1).Name$ & "..."
Call IOWriteLog(msg$)
Call IOStatusAuto(msg$)
FormPENEPMA08Batch.LabelCurrentInputFile.Caption = msg$
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

' Make material INP file (always a single file)
Call Penepma08CreateMaterialINP(Int(1), PENEPMASample())
If ierror Then Exit Sub

' Create and run the necessary batch files
Call Penepma08CreateMaterialBatch(Int(1), tForm)
If ierror Then Exit Sub

DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If
Next k%

' Confirm with user
msg$ = "All " & Format$(MAXBINARY%) & " MAT file calculations are complete"
Call IOWriteLog(msg$)
DoEvents

' Create pure element input files for Penepma08/12
For k% = 1 To 2
PENEPMASample(1).LastElm% = 1
PENEPMASample(1).LastChan% = PENEPMASample(1).LastElm%
PENEPMASample(1).ElmPercents!(1) = 100#
If k% = 1 Then pfilename$ = Trim$(Symup$(BinaryElement1%)) & "_" & Format$(PENEPMASample(1).ElmPercents!(1))
If k% = 2 Then pfilename$ = Trim$(Symup$(BinaryElement2%)) & "_" & Format$(PENEPMASample(1).ElmPercents!(1))
tForm.TextMaterialFiles(0).Text = pfilename$ & ".mat"
tForm.TextInputFile.Text = pfilename$ & ".in"

' Check input file parameters
Call Penepma08SaveInput(tForm)
If ierror Then Exit Sub

' Create .in file
Call Penepma08CreateInput(Int(1))
If ierror Then Exit Sub

msg$ = "Creating Penepma input file based on pure element " & tForm.TextInputFile.Text & "..."
Call IOWriteLog(msg$)
Call IOStatusAuto(msg$)
FormPENEPMA08Batch.LabelCurrentInputFile.Caption = msg$
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

Next k%

' Create binary input files for Penepma08/12
For k% = 1 To MAXBINARY%
tForm.TextMaterialFiles(0).Text = binarynames$(k%) & ".mat"
tForm.TextInputFile.Text = binarynames$(k%) & ".in"

' Check input file parameters
Call Penepma08SaveInput(tForm)
If ierror Then Exit Sub

' Create .in file
Call Penepma08CreateInput(Int(1))
If ierror Then Exit Sub

msg$ = "Creating Penepma input file based on binary " & tForm.TextInputFile.Text & "..."
Call IOWriteLog(msg$)
Call IOStatusAuto(msg$)
FormPENEPMA08Batch.LabelCurrentInputFile.Caption = msg$
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

Next k%

Call IOStatusAuto(vbNullString)
msg$ = "All Penepma input file calculations are complete. Now select the input files in the file list and click the Run Selected Input Files In Batch Mode button to start the Penepma simulations."
MsgBox msg$, vbOKOnly + vbInformation, "Penepma08BatchBinaryCreate"

Exit Sub

' Errors
Penepma08BatchBinaryCreateError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08BatchBinaryCreate"
ierror = True
Exit Sub

Penepma08BatchBinaryCreateSame:
msg$ = "The binary elements (" & Symup$(BinaryElement1%) & " and " & Symup$(BinaryElement2%) & ") are the same, but must be different for calculating a compositional range"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08BatchBinaryCreate"
ierror = True
Exit Sub

Penepma08BatchBinaryCreateBadElement:
msg$ = "Binary calculation element " & esym$ & " is not a valid element symbol for a binary element"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08BatchBinaryCreate"
ierror = True
Exit Sub

End Sub

Sub Penepma08BatchBinaryExtract()
' Extract the k-ratios from a series of binary composition Penepma input files based on the two elements

ierror = False
On Error GoTo Penepma08BatchBinaryExtractError

Dim ip As Integer
Dim esym As String

Dim n As Integer, k As Integer, l As Integer
Dim pfilename As String, tfilename As String, astring As String

Dim binarynames(1 To MAXBINARY%) As String

icancelauto = False

' Check that the batch folder exists
If Dir$(PENEPMA_BATCH_FOLDER$, vbDirectory) = vbNullString Then GoTo Penepma08BatchBinaryExtractFolderNotFound

Call InitKratios
If ierror Then Exit Sub

' Save the elements
esym$ = FormPENEPMA08Batch.ComboBinaryElement1.Text
ip% = IPOS1(MAXELM%, esym$, Symlo$())
If ip% = 0 Then GoTo Penepma08BatchBinaryExtractBadElement
BinaryElement1% = ip%

esym$ = FormPENEPMA08Batch.ComboBinaryElement2.Text
ip% = IPOS1(MAXELM%, esym$, Symlo$())
If ip% = 0 Then GoTo Penepma08BatchBinaryExtractBadElement
BinaryElement2% = ip%

' Check that binary elements are different
If BinaryElement1% = BinaryElement2% Then GoTo Penepma08BatchBinaryExtractSame

' Extract the pure element k-ratios
For n% = 1 To 2
PENEPMASample(1).LastElm% = 1
PENEPMASample(1).LastChan% = PENEPMASample(1).LastElm%
PENEPMASample(1).ElmPercents!(1) = 100#

If n% = 1 Then PENEPMASample(1).Elsyms$(1) = Symlo$(BinaryElement1%)
If n% = 2 Then PENEPMASample(1).Elsyms$(1) = Symlo$(BinaryElement2%)

If n% = 1 Then PENEPMASample(1).Xrsyms$(1) = Deflin$(BinaryElement1%)  ' just load defaults here
If n% = 2 Then PENEPMASample(1).Xrsyms$(1) = Deflin$(BinaryElement2%)  ' just load defaults here

' Overload with Penepma08/12 atomic weights for self consistency in calculations
If n% = 1 Then PENEPMASample(1).AtomicWts!(1) = pAllAtomicWts!(BinaryElement1%)
If n% = 2 Then PENEPMASample(1).AtomicWts!(1) = pAllAtomicWts!(BinaryElement2%)

' Load name and number for this binary
If n% = 1 Then pfilename$ = Trim$(Symup$(BinaryElement1%)) & "_" & Format$(PENEPMASample(1).ElmPercents!(1))
If n% = 2 Then pfilename$ = Trim$(Symup$(BinaryElement2%)) & "_" & Format$(PENEPMASample(1).ElmPercents!(1))
PENEPMASample(1).Name$ = pfilename$

' Check for Penepma intensity file (changed from pe-inten-01.dat to pe-intens-01.dat 11-2-2012)
tfilename$ = PENEPMA_BATCH_FOLDER$ & "\" & pfilename$ & "\pe-intens-01.dat"
If Dir$(tfilename$) = vbNullString Then GoTo Penepma08BatchBinaryExtractFileNotFound

msg$ = "Extracting k-ratios based on " & PENEPMASample(1).Name$ & "..."
Call IOWriteLog(msg$)
Call IOStatusAuto(msg$)
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

' Extract all MAXRAY emission lines from net intensity file
Call Penepma08BatchBinaryExtract2(n%, tfilename$)
If ierror Then Exit Sub

' Load standard intensities
For l% = 1 To MAXRAY% - 1
std_int!(n%, l%) = tot_int!(n%, l%)
Next l%

Next n%

' Create binary sample
PENEPMASample(1).LastElm% = 2
PENEPMASample(1).LastChan% = PENEPMASample(1).LastElm%

PENEPMASample(1).Elsyms$(1) = Symlo$(BinaryElement1%)
PENEPMASample(1).Elsyms$(2) = Symlo$(BinaryElement2%)

PENEPMASample(1).Xrsyms$(1) = Deflin$(BinaryElement1%)  ' just load defaults here
PENEPMASample(1).Xrsyms$(2) = Deflin$(BinaryElement2%)  ' just load defaults here

' Calculate material file for each compositional binary
For k% = 1 To MAXBINARY%

' Load element composition based on binary number
PENEPMASample(1).ElmPercents!(1) = BinaryRanges!(k%)
PENEPMASample(1).ElmPercents!(2) = 100# - BinaryRanges!(k%)

' Load name and number for this binary
binarynames$(k%) = Trim$(Symup$(BinaryElement1%)) & "-" & Trim$(Symup$(BinaryElement2%)) & "_" & Format$(PENEPMASample(1).ElmPercents!(1)) & "-" & Format$(PENEPMASample(1).ElmPercents!(2))
PENEPMASample(1).Name$ = binarynames$(k%)

' Check for Penepma intensity file (changed from pe-inten-01.dat to pe-intens-01.dat 11-2-2012)
tfilename$ = PENEPMA_BATCH_FOLDER$ & "\" & binarynames$(k%) & "\pe-intens-01.dat"
If Dir$(tfilename$) = vbNullString Then GoTo Penepma08BatchBinaryExtractFileNotFound

msg$ = "Extracting k-ratios based on binary " & PENEPMASample(1).Name$ & "..."
Call IOWriteLog(msg$)
Call IOStatusAuto(msg$)
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

' Extract from net intensity file
For n% = 1 To 2
Call Penepma08BatchBinaryExtract2(n%, tfilename$)
If ierror Then Exit Sub

' Load unknown (binary) intensities
For l% = 1 To MAXRAY% - 1
unk_pri_int!(n%, k%, l%) = pri_int!(n%, l%)
unk_flu_int!(n%, k%, l%) = flu_int!(n%, l%)
unk_tot_int!(n%, k%, l%) = tot_int!(n%, l%)
Next l%
Next n%
Next k%

' Output binary k-ratios (and alpha factors) to MAXRAY emission line output file
For l% = 1 To MAXRAY% - 1
astring$ = Trim$(Symup$(BinaryElement1%)) & "-" & Trim$(Symup$(BinaryElement2%))
tfilename$ = PENEPMA_BATCH_FOLDER$ & "\" & astring$ & "_k-ratios-" & Xraylo$(l%) & ".dat"
Call Penepma08BatchBinaryOutput(l%, tfilename$)
If ierror Then Exit Sub
Next l%

' Confirm with user
Call IOStatusAuto(vbNullString)
msg$ = "All " & Format$(MAXBINARY%) & " k-ratio extractions are complete"
Call IOWriteLog(msg$)
DoEvents

Exit Sub

' Errors
Penepma08BatchBinaryExtractError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08BatchBinaryExtract"
ierror = True
Exit Sub

Penepma08BatchBinaryExtractFolderNotFound:
msg$ = "The specified batch folder (" & PENEPMA_BATCH_FOLDER$ & ") was not found. Please specify the folder containing the Penepma pure element and binary k-ratio output files."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08BatchBinaryExtract"
ierror = True
Exit Sub

Penepma08BatchBinaryExtractSame:
msg$ = "The binary elements (" & Symup$(BinaryElement1%) & " and " & Symup$(BinaryElement2%) & ") are the same, but must be different for calculating a compositional range"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08BatchBinaryExtract"
ierror = True
Exit Sub

Penepma08BatchBinaryExtractBadElement:
msg$ = "Binary calculation element " & esym$ & " is not a valid element symbol for a binary element"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08BatchBinaryExtract"
ierror = True
Exit Sub

Penepma08BatchBinaryExtractFileNotFound:
msg$ = "The specified pure element or binary k-ratio output file (" & tfilename$ & ") was not found. Please check that the specified folder containing the Penepma pure element and binary k-ratio output files exists."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08BatchBinaryExtract"
ierror = True
Exit Sub

End Sub

Sub Penepma08BatchBinaryExtract2(n As Integer, tfilename As String)
' Load the net intensities from the Penepma output
' n = 1  BinaryElement1
' n = 2  BinaryElement2
'
' Sample input file:
' #  Results from PENEPMA. Output from photon detector #  1
' #
' #  Angular intervals : theta_1 = 4.500000E+01,  theta_2 = 5.500000E+01
' #                        phi_1 = 0.000000E+00,    phi_2 = 3.600000E+02
' #
' #  Intensities of characteristic lines. All in 1/(sr*electron).
' #    P = primary photons (from electron interactions);
' #    C = flourescence from characteristic x rays;
' #    B = flourescence from bremsstrahlung quanta;
' #   TF = C+B, total fluorescence;
' #  unc = statistical uncertainty (3 sigma).
' #
' # IZ S0 S1  E (eV)      P            unc       C            unc       B            unc       TF           unc       T            unc
'   29  K M3  8.9054E+03  6.622094E-06 1.80E-07  0.000000E+00 0.00E+00  3.641611E-07 1.99E-07  3.641611E-07 1.99E-07  6.986255E-06 2.68E-07
'   29  K M2  8.9054E+03  3.352315E-06 1.26E-07  0.000000E+00 0.00E+00  2.427741E-07 1.63E-07  2.427741E-07 1.63E-07  3.595089E-06 2.06E-07
'   29  K M5  8.9771E+03  7.045851E-09 5.65E-09  0.000000E+00 0.00E+00  0.000000E+00 0.00E+00  0.000000E+00 0.00E+00  7.045851E-09 5.65E-09
'   29  K M4  8.9771E+03  3.522926E-09 3.99E-09  0.000000E+00 0.00E+00  0.000000E+00 0.00E+00  0.000000E+00 0.00E+00  3.522926E-09 3.99E-09

ierror = False
On Error GoTo Penepma08BatchBinaryExtract2Error

Dim elementfound As Boolean
Dim l As Integer
Dim astring As String, bstring As String, tstring As String
Dim atnum As Integer
Dim S0 As String, s1 As String
Dim eV As Single

' Init
For l% = 1 To MAXRAY% - 1
pri_int!(n%, l%) = 0#
flch_int(n%, l%) = 0#
flbr_int(n%, l%) = 0#
flu_int(n%, l%) = 0#
tot_int(n%, l%) = 0#
tot_int_var(n%, l%) = 0#
Next l%

' Make sure input file is closed
Close #Temp1FileNumber%

If Dir$(Trim$(tfilename$)) = vbNullString Then Exit Sub
Open tfilename$ For Input As #Temp1FileNumber%

' Load array (IZ, SO, S1, E (eV), P, unc, etc.
Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$
If Len(Trim$(astring$)) > 0 And InStr(astring$, "#") = 0 Then            ' skip to first data line (also skips if value is -1.#IND0E+000)

' Load k-ratio data
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
atnum% = Val(Trim$(bstring$))                        ' IZ (atomic number)

' Check binary number against atomic number
If n% <> 1 And n% <> 2 Then GoTo Penepma08BatchBinaryExtract2BadParameter
If n% = 1 And atnum% = BinaryElement1% Or n% = 2 And atnum% = BinaryElement2% Then

' Load transition strings
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
S0$ = Trim$(bstring$)                                ' S0 (inner transition)
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
s1$ = Trim$(bstring$)                                ' S1 (outer transition)

Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
eV! = Val(Trim$(bstring$))                           ' E (eV)

' Load x-ray index for x-ray line
l% = 0
tstring$ = S0$ & " " & s1$
If tstring$ = "K L3" Then l% = 1          ' (Ka) (see table 6.2 in Penelope-2006-NEA-pdf)
If tstring$ = "K M3" Then l% = 2          ' (Kb)
If tstring$ = "L3 M5" Then l% = 3         ' (La)
If tstring$ = "L2 M4" Then l% = 4         ' (Lb)
If tstring$ = "M5 N7" Then l% = 5         ' (Ma)
If tstring$ = "M4 N6" Then l% = 6         ' (Mb)

'If tstring$ = "L2-M1" Then l% = 7         ' (Ln)
'If tstring$ = "L2-N4" Then l% = 8         ' (Lg)
'If tstring$ = "L2-N6" Then l% = 9         ' (Lv)
'If tstring$ = "L3-M1" Then l% = 10        ' (Ll)
'If tstring$ = "M3-N5" Then l% = 11        ' (Mg)
'If tstring$ = "M5-N3" Then l% = 12        ' (Mz)

' Skip if not one of the primary lines
If l% > 0 Then
elementfound = True
If DebugMode Then
If n% = 1 Then Call IOWriteLog("Penepma08BatchBinaryExtract2: Element " & Symup$(BinaryElement1%) & " found in " & tfilename$ & "...")
If n% = 2 Then Call IOWriteLog("Penepma08BatchBinaryExtract2: Element " & Symup$(BinaryElement2%) & " found in " & tfilename$ & "...")
End If

' Parse primary intensity
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
pri_int!(n%, l%) = Val(Trim$(bstring$))

' Parse primary intensity uncertainty
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub

' Parse characteristic fluorescence
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
flch_int!(n%, l%) = Val(Trim$(bstring$))

' Parse characteristic fluorescence intensity uncertainty
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub

' Parse continuum fluorescence
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
flbr_int!(n%, l%) = Val(Trim$(bstring$))

' Parse characteristic fluorescence uncertainty
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub

' Parse total fluorescence
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
flu_int!(n%, l%) = Val(Trim$(bstring$))

' Parse total fluorescence uncertainty
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub

' Parse total intensity
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
tot_int!(n%, l%) = Val(Trim$(bstring$))

' Parse total intensity uncertainty
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
tot_int_var!(n%, l%) = Val(Trim$(bstring$))
End If

End If
End If
Loop

Close #Temp1FileNumber%

' Check for element not found
If Not elementfound Then
If n% = 1 Then msg$ = "Warning in Penepma08BatchBinaryExtract2: The passed element (" & Symup$(BinaryElement1%) & ") and/or x-ray was not found in " & tfilename$ & ". "
If n% = 2 Then msg$ = "Warning in Penepma08BatchBinaryExtract2: The passed element (" & Symup$(BinaryElement2%) & ") and/or x-ray was not found in " & tfilename$ & ". "
msg$ = msg$ & "Please make sure the correct element and x-ray line were specified for the k-ratio extraction (generally ignore this warning for standard intensity folders)."
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
End If

Exit Sub

' Errors
Penepma08BatchBinaryExtract2Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08BatchBinaryExtract2"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08BatchBinaryExtract2BadParameter:
msg$ = "The passed parameter was not 1 or 2. This error should not occur, please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08BatchBinaryExtract2"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08BatchBinaryOutput(l As Integer, tfilename As String)
' Output binary k-ratios to output file for both emitters of the binary
' l% = x-ray line number (1 to MAXRAY% - 1)

ierror = False
On Error GoTo Penepma08BatchBinaryOutputError

Dim k As Integer, n As Integer, response As Integer
Dim astring As String

ReDim Binary_ZAF_Factors(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Single
ReDim Binary_ZAF_Coeffs(1 To MAXRAY%, 1 To MAXCOEFF4%) As Single

' Calculate k-ratios and a-factors first
For n% = 1 To 2
For k% = 1 To MAXBINARY%
unk_krat!(n%, k%, l%) = 0#
unk_afac!(n%, k%, l%) = 0#
Next k%
Next n%

' Calculate k-ratio and afactors (l is x-ray line to calculate)
For n% = 1 To 2
For k% = 1 To MAXBINARY%

' Calculate k-ratios
If std_int!(n%, l%) <> 0# Then
unk_krat!(n%, k%, l%) = 100# * unk_tot_int!(n%, k%, l%) / std_int!(n%, l%)
End If
If n% = 1 Then
Binary_ZAF_Kratios#(l%, k%) = CDbl(unk_krat!(n%, k%, l%))
Else
Binary_ZAF_Kratios#(l%, MAXBINARY% - (k% - 1)) = CDbl(unk_krat!(n%, k%, l%))
End If
Next k%

' Calculate alpha factors for all binaries
Call Penepma12CalculateAlphaFactors(l%, BinaryRanges!(), Binary_ZAF_Kratios#(), Binary_ZAF_Factors!(), Binary_ZAF_Coeffs!())
If ierror Then Exit Sub

' Save alpha factors
For k% = 1 To MAXBINARY%
If n% = 1 Then
unk_afac!(n%, k%, l%) = Binary_ZAF_Factors!(l%, k%)
Else
unk_afac!(n%, k%, l%) = Binary_ZAF_Factors!(l%, MAXBINARY% - (k% - 1))
End If
Next k%
Next n%

' Open output file
Open tfilename$ For Output As #Temp1FileNumber%

' Create column labels
astring$ = vbNullString
If Penepma08CheckPenepmaVersion%() = 8 Then
astring$ = astring$ & VbDquote$ & "Penepma 2008" & VbDquote$ & vbTab
Else
astring$ = astring$ & VbDquote$ & "Penepma 2012" & VbDquote$ & vbTab
End If

astring$ = astring$ & VbDquote$ & "A Conc. %" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "B Conc. %" & VbDquote$ & vbTab

astring$ = astring$ & VbDquote$ & "A Std. Int. %" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "B Std. Int. %" & VbDquote$ & vbTab

astring$ = astring$ & VbDquote$ & "A Pri. Int." & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "A Flu. Int." & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "A Tot. Int." & VbDquote$ & vbTab

astring$ = astring$ & VbDquote$ & "A K-ratio%" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "A A-Factor" & VbDquote$ & vbTab

astring$ = astring$ & VbDquote$ & "B Pri. Int." & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "B Flu. Int." & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "B Tot. Int." & VbDquote$ & vbTab

astring$ = astring$ & VbDquote$ & "B K-ratio%" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "B A-Factor" & VbDquote$ & vbTab

Print #Temp1FileNumber%, astring$

' Output each binary
For k% = 1 To MAXBINARY%
astring$ = vbNullString
astring$ = astring$ & VbDquote$ & Symup$(BinaryElement1%) & "-" & Symup$(BinaryElement2%) & VbDquote$ & vbTab

astring$ = astring$ & MiscAutoFormat$(BinaryRanges!(k%)) & vbTab
astring$ = astring$ & MiscAutoFormat$(100# - BinaryRanges!(k%)) & vbTab

astring$ = astring$ & Format$(std_int!(1, l%), e82$) & vbTab
astring$ = astring$ & Format$(std_int!(2, l%), e82$) & vbTab

For n% = 1 To 2
astring$ = astring$ & Format$(unk_pri_int!(n%, k%, l%), e82$) & vbTab
astring$ = astring$ & Format$(unk_flu_int!(n%, k%, l%), e82$) & vbTab
astring$ = astring$ & Format$(unk_tot_int!(n%, k%, l%), e82$) & vbTab

astring$ = astring$ & Format$(unk_krat!(n%, k%, l%)) & vbTab
astring$ = astring$ & Format$(unk_afac!(n%, k%, l%)) & vbTab
Next n%

Print #Temp1FileNumber%, astring$
Next k%

Close #Temp1FileNumber%

' Ask user whether to update Matrix.mdb update txt file
msg$ = "Do you want to add these k-ratios to the matrix.mdb update text file?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton1, "Penepma08BatchBinaryOutput")
If response% = vbNo Then Exit Sub

' Create/add to MatrixMDBUpdate.txt
tfilename$ = PENEPMA_Path$ & "\MatrixMDBUpdate.txt"


Exit Sub

' Errors
Penepma08BatchBinaryOutputError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08BatchBinaryOutput"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08ListStandard(mode As Integer, tForm As Form)
' Load the standard composition
' mode = 0 just load density to form
' mode = 1 calculate and output standard to log

ierror = False
On Error GoTo Penepma08ListStandardError

Dim stdnum As Integer

' Get standard from listbox
If tForm.ListAvailableStandards.ListIndex < 0 Then Exit Sub
stdnum% = tForm.ListAvailableStandards.ItemData(tForm.ListAvailableStandards.ListIndex)

' Just load standard density
If mode% = 0 Then

' Get composition based on standard number
Call StandardGetMDBStandard(stdnum, PENEPMASample())
If ierror Then Exit Sub

tForm.TextMaterialDensity.Text = Format$(PENEPMASample(1).SampleDensity!)

' Recalculate and display standard data
Else
If stdnum% > 0 Then Call StanFormCalculate(stdnum%, Int(0))
If ierror Then Exit Sub
End If

Exit Sub

' Errors
Penepma08ListStandardError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08ListStandard"
ierror = True
Exit Sub

End Sub

Sub Penepma08LoadInputFile(tfilename As String, tForm As Form)
' Load values from specified input file (*.in)

ierror = False
On Error GoTo Penepma08LoadInputFileError

Dim k As Integer
Dim astring As String, bstring As String

' Check production file
If Dir$(tfilename$) = vbNullString Then GoTo Penepma08LoadInputFileNotFound
If InStr(tfilename$, "pe-layout.in") > 0 Then GoTo Penepma08LoadInputFileInvalidInput

' First open input file and check geometry type only to load default values
Open tfilename$ For Input As #Temp1FileNumber%

Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$

If InStr(astring$, "GEOMFN") > 0 Then Call Penepma08LoadProduction2(astring$, bstring$, tForm.TextGeometryFile)
If ierror Then
MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFile [Penepma08LoadProduction2]"
Exit Sub
End If

Loop
Close #Temp1FileNumber%

' Check for geometry file and set OptionProduction based on name
bstring$ = tForm.TextGeometryFile.Text
If InStr(bstring$, "couple") > 0 Or InStr(bstring$, "sphere") > 0 Then tForm.OptionProduction(3).Value = True
If InStr(bstring$, "bilayer") > 0 Then tForm.OptionProduction(4).Value = True

' Re-open file and load values to overwrite defaults from production file
Open tfilename$ For Input As #Temp1FileNumber%

Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$

If InStr(astring$, "TITLE") > 0 Then
tForm.TextInputTitle.Text = Mid$(astring$, InStr(astring$, "TITLE") + Len("TITLE") + 2)
End If

If InStr(astring$, "SENERG") > 0 Then Call Penepma08LoadProduction2(astring$, bstring$, tForm.TextBeamEnergy)
If ierror Then
MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFile [Penepma08LoadProduction2]"
Exit Sub
End If
If InStr(astring$, "SPOSIT") > 0 Then Call Penepma08LoadProduction3(astring$, bstring$, tForm)
If ierror Then
MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFile [Penepma08LoadProduction3]"
Exit Sub
End If
If InStr(astring$, "SDIREC") > 0 Then Call Penepma08LoadProduction3(astring$, bstring$, tForm)
If ierror Then
MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFile [Penepma08LoadProduction3]"
Exit Sub
End If
If InStr(astring$, "SAPERT") > 0 Then Call Penepma08LoadProduction2(astring$, bstring$, tForm.TextBeamAperture)
If ierror Then
MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFile [Penepma08LoadProduction2]"
Exit Sub
End If

If InStr(astring$, "MFNAME") > 0 Then k% = k% + 1
If InStr(astring$, "MFNAME") > 0 Then Call Penepma08LoadProduction4(astring$, bstring$, k%, tForm)
If ierror Then
MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFile [Penepma08LoadProduction4]"
Exit Sub
End If

If InStr(astring$, "MSIMPA") > 0 Then Call Penepma08LoadProduction4(astring$, bstring$, k%, tForm)
If ierror Then
MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFile [Penepma08LoadProduction4]"
Exit Sub
End If

If InStr(astring$, "PDANGL") > 0 Then Call Penepma08LoadProduction5(astring$, bstring$)
If ierror Then
MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFile [Penepma08LoadProduction5]"
Exit Sub
End If

If InStr(astring$, "GEOMFN") > 0 Then Call Penepma08LoadProduction2(astring$, bstring$, tForm.TextGeometryFile)
If ierror Then
MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFile [Penepma08LoadProduction2]"
Exit Sub
End If

If InStr(astring$, "DUMPP") > 0 Then Call Penepma08LoadProduction2(astring$, bstring$, tForm.TextDumpPeriod)
If ierror Then
MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFile [Penepma08LoadProduction2]"
Exit Sub
End If

If InStr(astring$, "NSIMSH") > 0 Then Call Penepma08LoadProduction2(astring$, bstring$, tForm.TextNumberSimulatedShowers)
If ierror Then
MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFile [Penepma08LoadProduction2]"
Exit Sub
End If

If InStr(astring$, "TIME") > 0 Then Call Penepma08LoadProduction2(astring$, bstring$, tForm.TextSimulationTimePeriod)
If ierror Then
MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFile [Penepma08LoadProduction2]"
Exit Sub
End If

If InStr(astring$, "PDENER") > 0 Then Call Penepma08LoadProduction3(astring$, bstring$, tForm)
If ierror Then
MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFile [Penepma08LoadProduction3]"
Exit Sub
End If

Loop
Close #Temp1FileNumber%

' Because detector geometry parameters are not stored visibly, load detector geometry type based on parameters
tForm.OptionDetectorGeometry(0).Value = True    ' default is annular detector (0 to 360)
If InputPhi1# = -20# And InputPhi2# = 20# Then
tForm.OptionDetectorGeometry(1).Value = True    ' north
ElseIf InputPhi1# = 70# And InputPhi2# = 110# Then
tForm.OptionDetectorGeometry(2).Value = True    ' east
ElseIf InputPhi1# = 160# And InputPhi2# = 200# Then
tForm.OptionDetectorGeometry(3).Value = True    ' south
ElseIf InputPhi1# = 250# And InputPhi2# = 290# Then
tForm.OptionDetectorGeometry(4).Value = True    ' west
End If

Exit Sub

' Errors
Penepma08LoadInputFileError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08LoadInputFile"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08LoadInputFileNotFound:
msg$ = "Penepma input file (" & tfilename$ & ") was not found. Please download an up to date Penepma12.zip file and extract to your Penepma12 folder or contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08LoadInputFile"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08LoadInputFileInvalidInput:
msg$ = "The pe-layout.in file is not a valid Penepma input file. It is provided for documentation purposes only, so please select another input file."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08LoadInputFile"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08LoadInputFileParameter(tfilename As String, pstring As String, ptext As String, tForm As Form)
' Load specified value from specified input file (*.in)

ierror = False
On Error GoTo Penepma08LoadInputFileParameterError

'Dim k As Integer
Dim astring As String, bstring As String

' Check production file
If Dir$(tfilename$) = vbNullString Then GoTo Penepma08LoadInputFileParameterNotFound
ptext$ = vbNullString

' Open file and load specified value
Open tfilename$ For Input As #Temp1FileNumber%

Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$

If pstring$ = "TITLE" And InStr(astring$, pstring$) > 0 Then
tForm.TextTemp.Text = Mid$(astring$, InStr(astring$, "TITLE") + Len("TITLE") + 2)
ptext$ = Trim$(tForm.TextTemp.Text)
Close #Temp1FileNumber%
Exit Sub
End If

If pstring$ = "SENERG" And InStr(astring$, pstring$) > 0 Then
Call Penepma08LoadProduction2(astring$, bstring$, tForm.TextTemp)
If ierror Then
MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFileParameter [Penepma08LoadProduction2]"
Exit Sub
End If
ptext$ = Trim$(tForm.TextTemp.Text)
Close #Temp1FileNumber%
Exit Sub
End If

If pstring$ = "SPOSIT" And InStr(astring$, pstring$) > 0 Then
Call Penepma08LoadProduction3(astring$, bstring$, tForm)
If ierror Then
MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFileParameter [Penepma08LoadProduction3]"
Exit Sub
End If
ptext$ = Trim$(tForm.TextBeamPosition(0).Text)  ' X position only
Close #Temp1FileNumber%
Exit Sub
End If

'If pstring$ = "SDIREC" And InStr(astring$, pstring$) > 0 Then
'Call Penepma08LoadProduction3(astring$, bstring$, tForm)
'If ierror Then
'MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFileParameter [Penepma08LoadProduction3]"
'Exit Sub
'End If

If pstring$ = "SAPERT" And InStr(astring$, pstring$) > 0 Then
Call Penepma08LoadProduction2(astring$, bstring$, tForm.TextTemp)
If ierror Then
MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFileParameter [Penepma08LoadProduction2]"
Exit Sub
End If
ptext$ = Trim$(tForm.TextTemp.Text)
Close #Temp1FileNumber%
Exit Sub
End If

'If InStr(astring$, "MFNAME") > 0 Then k% = k% + 1
'If InStr(astring$, "MFNAME") > 0 Then Call Penepma08LoadProduction4(astring$, bstring$, k%, tForm)
'If ierror Then
'MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFileParameter [Penepma08LoadProduction4]"
'Exit Sub
'End If

'If InStr(astring$, "MSIMPA") > 0 Then Call Penepma08LoadProduction4(astring$, bstring$, k%, tForm)
'If ierror Then
'MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFileParameter [Penepma08LoadProduction4]"
'Exit Sub
'End If

'If InStr(astring$, "PDANGL") > 0 Then Call Penepma08LoadProduction5(astring$, bstring$)
'If ierror Then
'MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFileParameter [Penepma08LoadProduction5]"
'Exit Sub
'End If

If pstring$ = "GEOMFN" And InStr(astring$, pstring$) > 0 Then
Call Penepma08LoadProduction2(astring$, bstring$, tForm.TextTemp)
If ierror Then
MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFileParameter [Penepma08LoadProduction2]"
Exit Sub
End If
ptext$ = Trim$(tForm.TextTemp.Text)
Close #Temp1FileNumber%
Exit Sub
End If

If pstring$ = "DUMPP" And InStr(astring$, pstring$) > 0 Then
Call Penepma08LoadProduction2(astring$, bstring$, tForm.TextTemp)
If ierror Then
MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFileParameter [Penepma08LoadProduction2]"
Exit Sub
End If
ptext$ = Trim$(tForm.TextTemp.Text)
Close #Temp1FileNumber%
Exit Sub
End If

If pstring$ = "NSIMSH" And InStr(astring$, pstring$) > 0 Then
Call Penepma08LoadProduction2(astring$, bstring$, tForm.TextTemp)
If ierror Then
MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFileParameter [Penepma08LoadProduction2]"
Exit Sub
End If
ptext$ = Trim$(tForm.TextTemp.Text)
Close #Temp1FileNumber%
Exit Sub
End If

If pstring$ = "TIME" And InStr(astring$, pstring$) > 0 Then
Call Penepma08LoadProduction2(astring$, bstring$, tForm.TextTemp)
If ierror Then
MsgBox "Problem reading file " & tfilename$, vbOKOnly + vbExclamation, "Penepma08LoadInputFileParameter [Penepma08LoadProduction2]"
Exit Sub
End If
ptext$ = Trim$(tForm.TextTemp.Text)
Close #Temp1FileNumber%
Exit Sub
End If

Loop
Close #Temp1FileNumber%

Exit Sub

' Errors
Penepma08LoadInputFileParameterError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08LoadInputFileParameter"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08LoadInputFileParameterNotFound:
msg$ = "Penepma production file (" & tfilename$ & ") was not found. Please download an up to date Penepma12.zip file and extract to your Penepma12 folder or contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08LoadInputFileParameter"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08BatchExtractKratios(tForm As Form)
' Extract k-ratios from folders based on specified element and xray

ierror = False
On Error GoTo Penepma08BatchExtractKratiosError

Dim astring As String, tfilename As String, tfilename2 As String, kratiofile As String, bstring As String

Dim nCount As Long, n As Long
Dim sAllFiles() As String

Dim ip As Integer
Dim esym As String, xsym As String

Dim pKratio As Single, pKratio_var As Single
Dim dist_or_rad As Single
Dim ptext1 As String, ptext2 As String

' Ask user for unknown folder and standard folder
astring$ = "Browse Penepma Batch Folder For Standard Intensity (" & ExtractStdFolder$ & ")"
If ExtractStdFolder$ = vbNullString Then ExtractStdFolder$ = PENEPMA_Path$ & "\Batch"
ExtractStdFolder$ = IOBrowseForFolderByPath(True, ExtractStdFolder$, astring$, FormPENEPMA08Batch)
If Trim$(ExtractStdFolder$) = vbNullString Then GoTo Penepma08BatchExtractKratiosNoStdFolder

astring$ = "Browse Penepma Batch Folder For Unknown Intensities (" & ExtractFolder$ & ")"
If ExtractFolder$ = vbNullString Then ExtractFolder$ = PENEPMA_Path$ & "\Batch"
ExtractFolder$ = IOBrowseForFolderByPath(True, ExtractFolder$, astring$, FormPENEPMA08Batch)
If Trim$(ExtractFolder$) = vbNullString Then GoTo Penepma08BatchExtractKratiosNoFolder

If DebugMode Then
Call IOWriteLog("Penepma08ExtractKratios: Extract standard folder= " & ExtractStdFolder$)
Call IOWriteLog("Penepma08ExtractKratios: Extract unknown folder= " & ExtractFolder$)
End If

' Save the element and x-ray to extract k-ratios from
esym$ = FormPENEPMA08Batch.ComboElm.Text
ip% = IPOS1(MAXELM%, esym$, Symlo$())
If ip% = 0 Then GoTo Penepma08BatchExtractKratiosBadElement
ExtractElement% = ip%

xsym$ = FormPENEPMA08Batch.ComboXray.Text
ip% = IPOS1(MAXRAY% - 1, xsym$, Xraylo$())
If ip% = 0 Then GoTo Penepma08BatchExtractKratiosBadXray
ExtractXray% = ip%

If DebugMode Then
Call IOWriteLog("Penepma08ExtractKratios: Extract Element/Xray= " & Trim$(Symlo$(ExtractElement%)) & "/" & Xraylo$(ExtractXray%))
End If

' Check that the folders exist
If Dir$(ExtractStdFolder$, vbDirectory) = vbNullString Then GoTo Penepma08BatchExtractKratiosStdFolderNotFound
If Dir$(ExtractFolder$, vbDirectory) = vbNullString Then GoTo Penepma08BatchExtractKratiosFolderNotFound

' First extract the standard intensities (just use n% = 1 index) (check for both Penepma 2008 and Penepma 2012 output file names)
tfilename$ = ExtractStdFolder$ & "\pe-intens-01.dat"    ' check for Penepma 2012 output file
If Dir$(Trim$(tfilename$)) = vbNullString Then
tfilename$ = ExtractStdFolder$ & "\pe-inten-01.dat"    ' check for Penepma 2008 output file
If Dir$(Trim$(tfilename$)) = vbNullString Then GoTo Penepma08BatchExtractKratiosStdFileNotFound
End If

' Extract all MAXRAY emission lines from net intensity file
BinaryElement1% = ExtractElement%
Call Penepma08BatchBinaryExtract2(Int(1), tfilename$)
If ierror Then Exit Sub

' Store standard intensities
std_int!(Int(1), ExtractXray%) = tot_int!(Int(1), ExtractXray%)
If std_int!(Int(1), ExtractXray%) = 0# Then GoTo Penepma08BatchExtractKratiosStdIntZero

' Open k-ratio output file
Close #Temp2FileNumber%
kratiofile$ = ExtractFolder$ & "\Penepma_Kratios_" & Symlo$(ExtractElement%) & "_" & Xraylo$(ExtractXray%) & ".dat"
Open kratiofile$ For Output As #Temp2FileNumber%

' Create column labels
astring$ = VbDquote$ & Symup$(ExtractElement%) & " " & Xraylo$(ExtractXray%) & " in Sample" & VbDquote$ & vbTab & VbDquote$ & "Distance or Radius (um)" & VbDquote$ & vbTab & VbDquote$ & "Std.Tot.Int." & VbDquote$ & vbTab & VbDquote$ & "Unk.Tot.Int." & VbDquote$ & vbTab & VbDquote & "Unk.Tot.Int.Var." & VbDquote$ & vbTab & VbDquote$ & "K-ratio" & VbDquote$ & vbTab & VbDquote$ & "K-ratio Var." & VbDquote$
Print #Temp2FileNumber%, astring$
Call IOWriteLog(vbCrLf & astring$)

' Extract the unknown element k-ratios for all files in folder and sub folders
Call DirectorySearch("*intens-01.dat", ExtractFolder$, True, nCount&, sAllFiles$())
If ierror Then Exit Sub

' If no Penepma 2012 output files found, try using Penepma 2008 output file
If nCount& = 0 Then
Call DirectorySearch("*inten-01.dat", ExtractFolder$, True, nCount&, sAllFiles$())
If ierror Then Exit Sub
If nCount& = 0 Then GoTo Penepma08BatchExtractKratiosUnkFilesNotFound
End If

' Loop through all recursively found files
For n& = 1 To nCount&
tfilename$ = sAllFiles$(n&)
If Trim$(tfilename$) <> vbNullString Then

' Extract emission lines from current unknown intensity file (includes standard folder)
Call Penepma08BatchBinaryExtract2(Int(1), tfilename$)
If ierror Then Exit Sub

' Calculate k-ratios for specified element and line
If std_int!(Int(1), ExtractXray%) <> 0# Then
pKratio! = tot_int!(Int(1), ExtractXray%) / std_int!(Int(1), ExtractXray%)
pKratio_var! = tot_int_var!(Int(1), ExtractXray%) / std_int!(Int(1), ExtractXray%)
End If

' Obtain the geo file parameter to determine distance or radius
dist_or_rad! = 0#
tfilename2$ = Dir$(MiscGetPathOnly2$(tfilename$) & "\*.in")
tfilename2$ = MiscGetPathOnly2$(tfilename$) & "\" & tfilename2$
Call Penepma08LoadInputFileParameter(tfilename2$, "GEOMFN", ptext1$, tForm)
If ierror Then Exit Sub

' Determine the distance if couple file
If InStr(ptext1$, "couple") > 0 Then
Call Penepma08LoadInputFileParameter(tfilename2$, "SPOSIT", ptext2$, tForm)
If ierror Then Exit Sub
dist_or_rad! = Abs(Val(ptext2$) * MICRONSPERCM&)         ' x position only (from center)
End If

' Determine the radius if hemisphere file
If InStr(ptext1$, "sphere") > 0 Then
If InStr(ptext1$, "mic") > 0 Then
dist_or_rad! = Val(Left$(ptext1$, InStr(ptext1$, "mic"))) / 2#      ' radius of hemisphere
End If
End If

astring$ = vbNullString
astring$ = astring$ & VbDquote$ & MiscGetLastFolderOnly$(tfilename$) & VbDquote$ & vbTab
astring$ = astring$ & Format$(dist_or_rad!, f82$) & vbTab
astring$ = astring$ & Format$(std_int!(1, ExtractXray%), e82$) & vbTab
astring$ = astring$ & Format$(tot_int!(1, ExtractXray%), e82$) & vbTab
astring$ = astring$ & Format$(tot_int_var!(1, ExtractXray%), e82$) & vbTab
astring$ = astring$ & MiscAutoFormat$(pKratio!) & vbTab
astring$ = astring$ & MiscAutoFormat$(pKratio_var!)

Print #Temp2FileNumber%, astring$
Call IOWriteLog(astring$)
End If
Next n&

Close #Temp2FileNumber%

' Confirm with user
Call IOStatusAuto(vbNullString)
msg$ = "All " & Symlo$(ExtractElement%) & " " & Xraylo$(ExtractXray%) & " k-ratios were calculated and output to folder " & ExtractFolder$
Call IOWriteLog(msg$)
DoEvents

Exit Sub

' Errors
Penepma08BatchExtractKratiosError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08BatchExtractKratios"
Close #Temp2FileNumber%
ierror = True
Exit Sub

Penepma08BatchExtractKratiosNoStdFolder:
msg$ = "Please select a folder that contains the standard intensity files."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08BatchExtractKratios"
Close #Temp2FileNumber%
ierror = True
Exit Sub

Penepma08BatchExtractKratiosNoFolder:
msg$ = "Please select a folder that contains the unknown intensity files (extraction will include all subfolders containing a pe-intens-01.dat file)."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08BatchExtractKratios"
Close #Temp2FileNumber%
ierror = True
Exit Sub

Penepma08BatchExtractKratiosStdFolderNotFound:
msg$ = "Please select a folder that contains the standard intensity files."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08BatchExtractKratios"
Close #Temp2FileNumber%
ierror = True
Exit Sub

Penepma08BatchExtractKratiosFolderNotFound:
msg$ = "Please select a folder that contains the unknown intensity files (extraction will include all subfolders containing a pe-intens-01.dat file)."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08BatchExtractKratios"
Close #Temp2FileNumber%
ierror = True
Exit Sub

Penepma08BatchExtractKratiosStdFileNotFound:
msg$ = "Neither the pe-intens-01.dat nor the pe-inten-01.dat Penepma output files for the standard intensities were found"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08BatchExtractKratios"
Close #Temp2FileNumber%
ierror = True
Exit Sub

Penepma08BatchExtractKratiosStdIntZero:
msg$ = "Standard intensity for " & Symup$(ExtractElement%) & " " & Xraylo$(ExtractXray%) & " in " & tfilename$ & ", is zero"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08BatchExtractKratios"
Close #Temp2FileNumber%
ierror = True
Exit Sub

Penepma08BatchExtractKratiosUnkFilesNotFound:
msg$ = "No pe-intens-01.dat nor pe-inten-01.dat Penepma output files for unknowns were found"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08BatchExtractKratios"
Close #Temp2FileNumber%
ierror = True
Exit Sub

Penepma08BatchExtractKratiosBadElement:
msg$ = "The element specified is not a valid element symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08BatchExtractKratios"
Close #Temp2FileNumber%
ierror = True
Exit Sub

Penepma08BatchExtractKratiosBadXray:
msg$ = "The x-ray line specified is not a valid x-ray symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08BatchExtractKratios"
Close #Temp2FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08BatchBulkPureElementCreate(tForm As Form)
' Create a series of bulk pure element composition input files based on the range of two elements

ierror = False
On Error GoTo Penepma08BatchBulkPureElementCreateError

Dim ip As Integer
Dim esym As String

Dim n As Integer
Dim pfilename As String

icancelauto = False

' Specify a single material
For n% = 1 To MAXMATOUTPUT%
MaterialFiles$(n%) = vbNullString
MaterialsSelected%(n%) = 0
Next n%

' Save the elements
esym$ = FormPENEPMA08Batch.ComboPureElement1.Text
ip% = IPOS1(MAXELM%, esym$, Symlo$())
If ip% = 0 Then GoTo Penepma08BatchBulkPureElementCreateBadElement
PureElement1% = ip%

esym$ = FormPENEPMA08Batch.ComboPureElement2.Text
ip% = IPOS1(MAXELM%, esym$, Symlo$())
If ip% = 0 Then GoTo Penepma08BatchBulkPureElementCreateBadElement
PureElement2% = ip%

' Check that geometry file is bulk.geo
If Not MiscStringsAreSame(MiscGetFileNameOnly$(GeometryFile$), "bulk.geo") Then GoTo Penepma08BatchBulkPureElementsCreateBadGeometryFile

' Create pure element material files
For n% = PureElement1% To PureElement2%
PENEPMASample(1).LastElm% = 1
PENEPMASample(1).LastChan% = PENEPMASample(1).LastElm%
PENEPMASample(1).Elsyms$(1) = Symlo$(n%)
PENEPMASample(1).Xrsyms$(1) = Deflin$(n%)  ' just load defaults here

' Load element data
Call ElementGetData(PENEPMASample())
If ierror Then Exit Sub

' Overload with Penepma08/12 atomic weights for self consistency in calculations
PENEPMASample(1).AtomicWts!(1) = pAllAtomicWts!(PureElement1%)

' Load element composition based on binary number
PENEPMASample(1).ElmPercents!(1) = 100#
PENEPMASample(1).SampleDensity! = AllAtomicDensities!(PureElement1%)

' Load name and number for this binary
pfilename$ = Trim$(Symup$(n%)) & "_" & Format$(PENEPMASample(1).ElmPercents!(1)) & "_" & Format$(BeamEnergy# / EVPERKEV#) & "keV"
PENEPMASample(1).Name$ = pfilename$
MaterialFiles$(1) = PENEPMASample(1).Name$ & ".MAT"

' Check if material file exists already
If Dir$(PENEPMA_Path$ & "\" & MaterialFiles$(1)) = vbNullString Or (Dir$(PENEPMA_Path$ & "\" & MaterialFiles$(1)) <> vbNullString And FormPENEPMA08Batch.CheckDoNotOverwriteExisting.Value = vbUnchecked) Then
MaterialsSelected%(1) = MAXINTEGER%     ' any non-zero number
MaterialDensity# = PENEPMASample(1).SampleDensity!

msg$ = "Creating material input file based on " & PENEPMASample(1).Name$ & "..."
Call IOWriteLog(msg$)
Call IOStatusAuto(msg$)
FormPENEPMA08Batch.LabelCurrentInputFile.Caption = msg$
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

' Make material INP file (always a single file)
Call Penepma08CreateMaterialINP(Int(1), PENEPMASample())
If ierror Then Exit Sub

' Create and run the necessary batch files
Call Penepma08CreateMaterialBatch(Int(1), tForm)
If ierror Then Exit Sub

End If
Next n%

' Confirm with user
n% = (PureElement2% - PureElement1%) + 1
If n% < 0 Then n% = 0
msg$ = "All " & Format$(n%) & " MAT file calculations are complete"
Call IOWriteLog(msg$)
DoEvents

' Create pure element input files for Penepma08/12
For n% = PureElement1% To PureElement2%
PENEPMASample(1).LastElm% = 1
PENEPMASample(1).LastChan% = PENEPMASample(1).LastElm%
PENEPMASample(1).ElmPercents!(1) = 100#
pfilename$ = Trim$(Symup$(n%)) & "_" & Format$(PENEPMASample(1).ElmPercents!(1)) & "_" & Format$(BeamEnergy# / EVPERKEV#) & "keV"
tForm.TextMaterialFiles(0).Text = pfilename$ & ".mat"
tForm.TextInputFile.Text = pfilename$ & ".in"

' Check if input file exists already
If Dir$(PENEPMA_Path$ & "\" & pfilename$ & ".in") = vbNullString Or (Dir$(PENEPMA_Path$ & "\" & pfilename$ & ".in") <> vbNullString And FormPENEPMA08Batch.CheckDoNotOverwriteExisting.Value = vbUnchecked) Then

' Check input file parameters
Call Penepma08SaveInput(tForm)
If ierror Then Exit Sub
DoEvents

' Create .in file
Call Penepma08CreateInput(Int(1))
If ierror Then Exit Sub

msg$ = "Creating Penepma input file based on pure element " & tForm.TextInputFile.Text & "..."
Call IOWriteLog(msg$)
Call IOStatusAuto(msg$)
FormPENEPMA08Batch.LabelCurrentInputFile.Caption = msg$
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

End If
Next n%

Call IOStatusAuto(vbNullString)
msg$ = "All Penepma input file calculations are complete. Now select the input files in the file list and click the Run Selected Input Files In Batch Mode button to start the Penepma simulations."
MsgBox msg$, vbOKOnly + vbInformation, "Penepma08BatchBulkPureElementCreate"

Exit Sub

' Errors
Penepma08BatchBulkPureElementCreateError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08BatchBulkPureElementCreate"
ierror = True
Exit Sub

Penepma08BatchBulkPureElementCreateBadElement:
msg$ = "Bulk pure calculation element " & esym$ & " is not a valid element symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08BatchBulkPureElementCreate"
ierror = True
Exit Sub

Penepma08BatchBulkPureElementsCreateBadGeometryFile:
msg$ = "The specified geometry file for bulk pure element calculations must be the bulk.geo file"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08BatchBulkPureElementCreate"
ierror = True
Exit Sub

End Sub

Sub Penepma08Load(tForm As Form)
' Load the form (works for Penepma08 and Penepma12)

ierror = False
On Error GoTo Penepma08LoadError

Dim i As Integer, n As Integer
Dim astring As String

' Load filename and delete dump file
Call Penepma08Init0
If ierror Then Exit Sub

' Clear graph
icancelauto = False
Call Penepma08GraphClear
If ierror Then Exit Sub

' Update form caption
If Penepma08CheckPenepmaVersion = 8 Then
tForm.Caption = "Create PENEPMA Material and Input Files (using Penepma08)"
Else
tForm.Caption = "Create PENEPMA Material and Input Files (using Penepma12)"
End If

' Check for valid folders
astring$ = Dir$(PENDBASE_Path$, vbDirectory)
If astring$ = vbNullString Then
msg$ = "The Penelope (Pendbase) application files are not found in the specified folder: " & PENDBASE_Path$ & vbCrLf
msg$ = msg$ & "Please contact Probe Software, Inc to obtain the Penelope application files, copy them to the specified location and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08Load"
ierror = True
Exit Sub
End If

astring$ = Dir$(PENEPMA_Path$, vbDirectory)
If astring$ = vbNullString Then
msg$ = "The Penelope (Penepma08 or Penepma12) application files are not found in the specified folder: " & PENEPMA_Path$ & vbCrLf
msg$ = msg$ & "Please contact Probe Software, Inc to obtain the Penelope application files, copy them to the specified location and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08Load"
ierror = True
Exit Sub
End If

' Init
Call Penepma08Init
If ierror Then Exit Sub

' Load the standard list box (even if there was an error importing)
Call StandardLoadList(tForm.ListAvailableStandards)
If ierror Then Exit Sub

' Load selected materials (if any)
For n% = 1 To MAXMATOUTPUT%
If MaterialsSelected%(n%) > 0 Then
For i% = 0 To tForm.ListAvailableStandards.ListCount - 1
If tForm.ListAvailableStandards.ItemData(i%) = MaterialsSelected%(n%) Then
tForm.ListAvailableStandards.Selected(i%) = True
End If
Next i%
End If
Next n%

' Load PENDBASE options
tForm.TextMaterialDensity.Text = Format$(MaterialDensity#)
tForm.TextMaterialFcb.Text = Format$(MaterialFcb#)
tForm.TextMaterialWcb.Text = Format$(MaterialWcb#)

' Load PENEPMA options
tForm.TextInputTitle.Text = BeamTitle$
tForm.TextBeamTakeoff.Text = Format$(BeamTakeOff#)
tForm.TextBeamEnergy.Text = Format$(BeamEnergy#)

For i% = 1 To 3
tForm.TextBeamPosition(i% - 1).Text = Format$(BeamPosition#(i%))
Next i%

tForm.TextBeamDirection(0).Text = Format$(BeamDirection#(1))
tForm.TextBeamDirection(1).Text = Format$(BeamDirection#(2))
tForm.TextBeamAperture.Text = Format$(BeamAperture#)
tForm.TextDumpPeriod.Text = Format$(BeamDumpPeriod#)

tForm.TextEnergyRangeMinMaxNumber(0).Text = Format$(BeamMinimumEnergyRange#)
tForm.TextEnergyRangeMinMaxNumber(1).Text = Format$(BeamMaximumEnergyRange#)
tForm.TextEnergyRangeMinMaxNumber(2).Text = Format$(BeamNumberOfEnergyChannels)

tForm.TextNumberSimulatedShowers.Text = Format$(BeamNumberSimulatedShowers#, "Scientific")
tForm.TextSimulationTimePeriod.Text = Format$(BeamSimulationTimePeriod#, "Scientific")

tForm.OptionDetectorGeometry(DetectorGeometryType%).Value = True

' Load default materials and geometry files
If GeometryFile$ = vbNullString Then GeometryFile$ = PENEPMA_Root$ & "\bulk.geo"

For i% = 1 To MAXMATOUTPUT%
tForm.TextMaterialFiles(i% - 1).Text = MaterialFiles$(i%)
tForm.TextEABS1(i% - 1).Text = Format$(InputEABS1#(i%), "Scientific")
tForm.TextEABS2(i% - 1).Text = Format$(InputEABS2#(i%), "Scientific")
Next i%

tForm.TextGeometryFile.Text = MiscGetFileNameOnly$(GeometryFile$)
tForm.TextInputFile.Text = MiscGetFileNameOnly$(InputFile$)

' Force loading of production file values
tForm.OptionProduction(BeamProductionIndex&).Value = True

If UseGridLines Then
tForm.CheckUseGridLines.Value = vbChecked
Else
tForm.CheckUseGridLines.Value = vbUnchecked
End If

If UseLogScale Then
tForm.CheckUseLogScale.Value = vbChecked
Else
tForm.CheckUseLogScale.Value = vbUnchecked
End If

' Set timer to update data from output files (if simulation still running)
If SimulationInProgress Then tForm.Timer1.Interval = PENEPMA_DISPLAY_SEC# * MSECPERSEC#

' Load graph (no data yet)
Call Penepma08GetPenepmaDAT        ' load current file specific data
If ierror Then Exit Sub

Call Penepma08GraphLoad_PE(GraphDisplayOption%, UseGridLines, UseLogScale, BeamTitle$)       ' load graph PE parameters
If ierror Then Exit Sub

tForm.OptionDisplayGraph(GraphDisplayOption%).Value = True
DoEvents
Call Penepma08GraphClear
If ierror Then Exit Sub

tForm.Show vbModeless
If ierror Then Exit Sub

Exit Sub

' Errors
Penepma08LoadError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08Load"
ierror = True
Exit Sub

End Sub

Sub Penepma08GraphUpdate(Index As Integer)
' Update the graph control

ierror = False
On Error GoTo Penepma08GraphUpdateError

GraphDisplayOption% = Index% ' update if user clicked graph option
Call Penepma08GraphUpdate_PE(Index%, BeamEnergy#, BeamTitle$, nPoints&, xdata#(), ydata#())
If ierror Then Exit Sub

Exit Sub

' Errors
Penepma08GraphUpdateError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08GraphUpdate"
ierror = True
Exit Sub

End Sub

Sub Penepma08GraphClear()
' Clear the Penepma graph

ierror = False
On Error GoTo Penepma08GraphClearError

' Clear graph
Call Penepma08GraphClear_PE
If ierror Then Exit Sub

Exit Sub

' Errors
Penepma08GraphClearError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08GraphClear"
ierror = True
Exit Sub

End Sub

Sub Penepma08PlotLog()
' Plot y axis as log

ierror = False
On Error GoTo Penepma08PlotLogError

If FormPENEPMA08_PE.CheckUseLogScale.Value = vbChecked Then
UseLogScale = True
Else
UseLogScale = False
End If

Exit Sub

' Errors
Penepma08PlotLogError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08PlotLog"
ierror = True
Exit Sub

End Sub

Sub Penepma08SetOptionProductionEnables(Index As Integer)
' Set the enables for the OptionProduction controls

ierror = False
On Error GoTo Penepma08SetOptionProductionEnablesError

If Index% = 0 Then  ' optimize x-rays
FormPENEPMA08_PE.CommandBrowseMaterialFiles(0).Enabled = True
FormPENEPMA08_PE.CommandBrowseMaterialFiles(1).Enabled = False
FormPENEPMA08_PE.UpDownXray(0).Enabled = True
FormPENEPMA08_PE.CommandAdjust(0).Enabled = True
FormPENEPMA08_PE.CommandElement(0).Enabled = True
FormPENEPMA08_PE.UpDownXray(1).Enabled = False
FormPENEPMA08_PE.CommandAdjust(1).Enabled = False
FormPENEPMA08_PE.CommandElement(1).Enabled = False
FormPENEPMA08_PE.TextEABS1(0).Enabled = True
FormPENEPMA08_PE.TextEABS1(1).Enabled = False
FormPENEPMA08_PE.TextEABS2(0).Enabled = True
FormPENEPMA08_PE.TextEABS2(1).Enabled = False

ElseIf Index% = 1 Then  ' optimize backscatter
FormPENEPMA08_PE.CommandBrowseMaterialFiles(0).Enabled = True
FormPENEPMA08_PE.CommandBrowseMaterialFiles(1).Enabled = False
FormPENEPMA08_PE.UpDownXray(0).Enabled = True
FormPENEPMA08_PE.CommandAdjust(0).Enabled = True
FormPENEPMA08_PE.CommandElement(0).Enabled = True
FormPENEPMA08_PE.UpDownXray(1).Enabled = False
FormPENEPMA08_PE.CommandAdjust(1).Enabled = False
FormPENEPMA08_PE.CommandElement(1).Enabled = False
FormPENEPMA08_PE.TextEABS1(0).Enabled = True
FormPENEPMA08_PE.TextEABS1(1).Enabled = False
FormPENEPMA08_PE.TextEABS2(0).Enabled = True
FormPENEPMA08_PE.TextEABS2(1).Enabled = False

ElseIf Index% = 2 Then  ' optimize continuum
FormPENEPMA08_PE.CommandBrowseMaterialFiles(0).Enabled = True
FormPENEPMA08_PE.CommandBrowseMaterialFiles(1).Enabled = False
FormPENEPMA08_PE.UpDownXray(0).Enabled = True
FormPENEPMA08_PE.CommandAdjust(0).Enabled = True
FormPENEPMA08_PE.CommandElement(0).Enabled = True
FormPENEPMA08_PE.UpDownXray(1).Enabled = False
FormPENEPMA08_PE.CommandAdjust(1).Enabled = False
FormPENEPMA08_PE.CommandElement(1).Enabled = False
FormPENEPMA08_PE.TextEABS1(0).Enabled = True
FormPENEPMA08_PE.TextEABS1(1).Enabled = False
FormPENEPMA08_PE.TextEABS2(0).Enabled = True
FormPENEPMA08_PE.TextEABS2(1).Enabled = False

ElseIf Index% = 3 Then  ' optimize couple or hemisphere
FormPENEPMA08_PE.CommandBrowseMaterialFiles(0).Enabled = True
FormPENEPMA08_PE.CommandBrowseMaterialFiles(1).Enabled = True
FormPENEPMA08_PE.UpDownXray(0).Enabled = True
FormPENEPMA08_PE.CommandAdjust(0).Enabled = True
FormPENEPMA08_PE.CommandElement(0).Enabled = True
FormPENEPMA08_PE.UpDownXray(1).Enabled = True
FormPENEPMA08_PE.CommandAdjust(1).Enabled = True
FormPENEPMA08_PE.CommandElement(1).Enabled = True
FormPENEPMA08_PE.TextEABS1(0).Enabled = True
FormPENEPMA08_PE.TextEABS1(1).Enabled = True
FormPENEPMA08_PE.TextEABS2(0).Enabled = True
FormPENEPMA08_PE.TextEABS2(1).Enabled = True

ElseIf Index% = 4 Then  ' optimize bilayer (thin film)
FormPENEPMA08_PE.CommandBrowseMaterialFiles(0).Enabled = True
FormPENEPMA08_PE.CommandBrowseMaterialFiles(1).Enabled = True
FormPENEPMA08_PE.UpDownXray(0).Enabled = True
FormPENEPMA08_PE.CommandAdjust(0).Enabled = True
FormPENEPMA08_PE.CommandElement(0).Enabled = True
FormPENEPMA08_PE.UpDownXray(1).Enabled = True
FormPENEPMA08_PE.CommandAdjust(1).Enabled = True
FormPENEPMA08_PE.CommandElement(1).Enabled = True
FormPENEPMA08_PE.TextEABS1(0).Enabled = True
FormPENEPMA08_PE.TextEABS1(1).Enabled = True
FormPENEPMA08_PE.TextEABS2(0).Enabled = True
FormPENEPMA08_PE.TextEABS2(1).Enabled = True
End If

Exit Sub

' Errors
Penepma08SetOptionProductionEnablesError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08SetOptionProductionEnables"
ierror = True
Exit Sub

End Sub

Sub Penepma08BatchCopyRename()
' Copy and rename pure element files

ierror = False
On Error GoTo Penepma08BatchCopyRenameError

Dim tpath As String, tstring As String
Dim tfilename As String, tsym As String
Dim tkeV As Integer, m As Integer

Dim nCount As Long, n As Long
Dim sAllFiles() As String

' Load to module and dialog
tstring$ = "Browse PENEPMA Batch Project Folder For Pure Elements Copy/Rename"
If PENEPMA_BATCH_FOLDER$ = vbNullString Then PENEPMA_BATCH_FOLDER$ = PENEPMA_Path$
tpath$ = IOBrowseForFolderByPath(True, PENEPMA_BATCH_FOLDER$, tstring$, FormPENEPMA08Batch)
If ierror Then Exit Sub

If Trim$(tpath$) <> vbNullString Then PENEPMA_BATCH_FOLDER$ = tpath$
FormPENEPMA08Batch.TextBatchFolder.Text = PENEPMA_BATCH_FOLDER$

' Now go through each folder and copy files as PENEPMA_Path$ & "\pure\" & Format$(tkeV%) & "keV\pe-spect-01" & "_" & Trim$(tSym$) & ".dat"
Call DirectorySearch("pe-spect-01.dat", PENEPMA_BATCH_FOLDER$, True, nCount&, sAllFiles$())     ' get all pe-spect-01.dat files recursively
If ierror Then Exit Sub

If nCount& < 1 Then GoTo Penepma08BatchCopyRenameNoFiles

' Check for pure and keV sub folders and if not found, create them
If Dir$(PENEPMA_Path$ & "\pure\", vbDirectory) = vbNullString Then MkDir PENEPMA_Path$ & "\pure\"
If Dir$(PENEPMA_Path$ & "\pure\10keV", vbDirectory) = vbNullString Then MkDir PENEPMA_Path$ & "\pure\10keV"
If Dir$(PENEPMA_Path$ & "\pure\15keV", vbDirectory) = vbNullString Then MkDir PENEPMA_Path$ & "\pure\15keV"
If Dir$(PENEPMA_Path$ & "\pure\20keV", vbDirectory) = vbNullString Then MkDir PENEPMA_Path$ & "\pure\20keV"

' Loop on all files
Screen.MousePointer = vbHourglass
For n& = 1 To nCount&
tfilename$ = sAllFiles$(n&)

' Determine keV for folder
m% = InStr(tfilename$, "(")
If m% > 0 Then
tkeV% = Val(Mid$(tfilename$, m% + 1, 2))

' Determine element symbol for folder
m% = InStr(tfilename$, ")\")
If m% > 0 Then
tsym$ = Mid$(tfilename$, m% + 2, 2)

' Remove trailing underscore if present (single character element symbols)
If Right$(tsym$, 1) = "_" Then tsym$ = Left$(tsym$, Len(tsym$) - 1)

' Copy and rename spectrum file
FileCopy MiscGetPathOnly$(tfilename$) & "pe-spect-01.dat", PENEPMA_Path$ & "\pure\" & Format$(tkeV%) & "keV\pe-spect-01" & "_" & Trim$(tsym$) & ".dat"

' Copy and rename net intensity file
FileCopy MiscGetPathOnly$(tfilename$) & "pe-intens-01.dat", PENEPMA_Path$ & "\pure\" & Format$(tkeV%) & "keV\pe-intens-01" & "_" & Trim$(tsym$) & ".dat"

End If
End If
Next n&
Screen.MousePointer = vbDefault

' Confirm with user
msg$ = "All pure element files were copied to the Penepma12\Penepma\pure folder."
MsgBox msg$, vbOKOnly + vbInformation, "Penepma08BatchCopyRename"

Exit Sub

' Errors
Penepma08BatchCopyRenameError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08BatchCopyRename"
ierror = True
Exit Sub

Penepma08BatchCopyRenameNoFiles:
Screen.MousePointer = vbDefault
msg$ = "No files of the type specified were found."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08BatchCopyRename"
ierror = True
Exit Sub

End Sub
