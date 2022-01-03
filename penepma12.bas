Attribute VB_Name = "CodePENEPMA12"
' (c) Copyright 1995-2022 by John J. Donovan
Option Explicit

Global Const MICROGRAMSPERGRAM& = 1000000            ' micrograms per gram

Global MaterialDensityA As Double
Global MaterialDensityB As Double
Global MaterialDensityBStd As Double            ' assume constant (just for k-ratio calculation)

Global ParameterFileA As String                 ' filename only, no path
Global ParameterFileB As String                 ' filename only, no path
Global ParameterFileBStd As String              ' filename only, no path

' Globals
Global TotalNumberOfSimulations As Long
Global CurrentSimulationsNumber As Long

Global BinaryMethod As Integer                  ' need to be global for PenPFE
Global ExtractMethod As Integer                 ' need to be global for PenPFE

Global CalculateRandomTable() As Integer
Global ExtractRandomTable() As Integer

Global pAllAtomicWts(1 To MAXELM%) As Single    ' Penepma08/12 atomic weights

' Graph variables
Dim UseLogScale As Boolean
Dim UseGridLines As Boolean

Const PENEPMA_MINPERCENT! = 0.0001
Const COL7% = 7
Const NUMSIM& = 10                      ' number of beam energy simulations per Penfluor calculation

' Module level variables
Dim PenepmaTaskID As Long
Dim MaterialInProgress As Boolean       ' for Material
Dim SimulationInProgress As Boolean     ' for Penfluor
Dim FitParametersInProgress As Boolean  ' for Fitall
Dim CalculationInProgress As Boolean    ' for Fanal

Dim MaterialMeasuredElement As Integer  ' atomic number of element
Dim MaterialMeasuredXray As Integer     ' x-ray number of element (1 to 6)

Dim MaterialMeasuredTakeoff As Double   ' in degrees
Dim MaterialMeasuredEnergy As Double    ' in kilovolts

Dim MaterialMeasuredDistance As Double
Dim MaterialMeasuredGridPoints As Integer

Dim SendToExcel As Boolean

' Pendbase options
Dim MaterialSelectedA As Integer
Dim MaterialFileA As String     ' filename only, no path
Dim MaterialFcbA As Double, MaterialWcbA As Double

Dim MaterialSelectedB As Integer
Dim MaterialFileB As String     ' filename only, no path
Dim MaterialFcbB As Double, MaterialWcbB As Double

Dim MaterialSelectedBStd As Integer
Dim MaterialFileBStd As String
Dim MaterialFcbBStd As Double, MaterialWcbBStd As Double

Dim MaterialSimulationOverhead As Double
Dim FitParametersTime As Double

Dim CurrentSimulationModeNumber As Integer  ' 1 = MAT/PAR A, 2 = MAT/PAR B, 3 = MAT/PAR B Std
Dim AdditionalElementSimulationTime(1 To 3) As Double

Dim MaterialSimulationStart As Variant
Dim FitParametersStart As Variant
Dim CalculationStart As Variant

Dim MaterialSimulationTime As Double
Dim MaterialSimulationShowers As Double

Dim PenfluorInputFile As String
Dim PenfluorOutputFile As String
Dim FanalInputFile As String

Dim FANAL_IN_File As String
Dim VACS_DAT_File As String
Dim RANGES_DAT_File As String
Dim MIXED_DAT_File As String
Dim KRATIOS_DAT_File As String
Dim FLUORMAT1_PAR_File As String
Dim FLUORMAT2_PAR_File As String
Dim FLUORMAT3_PAR_File As String
Dim ATCOEFFS_DAT_File As String

Dim KRATIOS_DAT_File2 As String

Dim nPoints As Long, nsets As Long
Dim xdist() As Double       ' linear distance (um)
Dim mdist() As Double       ' mass distance (ug/cm2)
Dim yktotal() As Double     ' fluorescence kratio% plus primary x-rays kratio% from material A and material B
Dim ykfluor() As Double     ' fluorescence kratio% only (minus primary x-ray kratio% from material A)

Dim yctotal() As Double      ' "apparent" concentration from total intensity (fluor plus primary)
Dim ycA_only() As Double     ' "apparent" concentration % from A fluorescence only
Dim ycb_only() As Double     ' "apparent" concentration % from B (boundary) fluorescence only
Dim yc_prix() As Double      ' "apparent" concentration % from primary x-ray only

Dim yktotal_meas() As Double ' "measured" intensity % from total intensity
Dim yztotal_meas() As Double ' "measured" ZAF correction from total intensity
Dim yctotal_meas() As Double ' "measured" concentration % from total intensity

Dim flach() As Double        ' Mat A characteristic fluorescence
Dim flabr() As Double        ' Mat A continuum fluorescence
Dim flbch() As Double        ' Mat B characteristic fluorescence
Dim flbbr() As Double        ' Mat B continuum fluorescence
Dim pri_int() As Double      ' primary x-ray intensity
Dim std_int() As Double      ' standard intensity

Dim fluA_k() As Double        ' Mat A total fluorescence k-ratio %
Dim fluB_k() As Double        ' Mat B total fluorescence k-ratio %
Dim prix_k() As Double        ' Primary x-ray k-ratio %

' CalcZAF matrix correction factors for each material
Dim MatA_ZAFCors(1 To MAXZAFCOR%, 1 To MAXCHAN%) As Single
Dim MatB_ZAFCors(1 To MAXZAFCOR%, 1 To MAXCHAN%) As Single
Dim MatBStd_ZAFCors(1 To MAXZAFCOR%, 1 To MAXCHAN%) As Single

Dim MatA_Krats(1 To MAXCHAN%) As Single
Dim MatB_Krats(1 To MAXCHAN%) As Single
Dim MatBStd_Krats(1 To MAXCHAN%) As Single

Dim MatA_StdPercents(1 To MAXCHAN%) As Single
Dim MatB_StdPercents(1 To MAXCHAN%) As Single
Dim MatBStd_StdPercents(1 To MAXCHAN%) As Single

' Fanal matrix correction factors for material A (note that absorption and atomic number terms are actually the combined ZA terms)
Dim Fanal_ZAFCors(1 To MAXZAFCOR%) As Single    ' 1 = A only, 2 = F only, 3 = Z only, 4 = ZAF
Dim Fanal_Krats As Single

' Binary calculations
Dim BinaryElement1 As Integer
Dim BinaryElement2 As Integer

Dim CalculateDoNotOverwritePAR As Boolean
Dim CalculateOnlyOverwriteLowerPrecisionPAR As Boolean
Dim CalculateOnlyOverwriteHigherMinimumEnergyPAR As Boolean

Dim CalculateDoNotOverwriteTXT As Boolean
Dim CalculateForMatrixRange As Boolean

Dim CalculateFromFormulaOrStandard As Integer

' Extract calculates for all valid x-ray lines
Dim ExtractElement As Integer       ' for matrix and boundary calculations (the measured element)
Dim ExtractMatrix As Integer        ' only for matrix calculations

Dim ExtractMatrixA1 As Integer      ' for boundary calculations (beam incident material)
Dim ExtractMatrixA2 As Integer      ' for boundary calculations (beam incident material)
Dim ExtractMatrixB1 As Integer      ' for boundary calculations (boundary material)
Dim ExtractMatrixB2 As Integer      ' for boundary calculations (boundary material)

Dim ExtractForSpecifiedRange As Boolean

' Density calculations
Dim DensityElementA As Integer, DensityElementB As Integer
Dim DensityConcA As Single, DensityConcB As Single

' Alpha factor calculations (manual)
Dim OptionEnterFraction As Boolean
Dim ConcA As Single, ConcB As Single
Dim KratA As Single, KratB As Single

' Plotting
Dim CalcZAF_ZAF_Factors() As Single
Dim CalcZAF_ZA_Factors() As Single
Dim CalcZAF_F_Factors() As Single

Dim Binary_ZAF_Factors() As Single
Dim Binary_ZA_Factors() As Single
Dim Binary_F_Factors() As Single

Dim Binary_ZAF_Coeffs() As Single
Dim CalcZAF_ZAF_Coeffs() As Single

Dim Binary_ZA_Coeffs() As Single
Dim CalcZAF_ZA_Coeffs() As Single

Dim Binary_F_Coeffs() As Single
Dim CalcZAF_F_Coeffs() As Single

Dim Binary_ZAF_Betas() As Single
Dim CalcZAF_ZAF_Betas() As Single

Dim Binary_ZAF_Devs() As Single
Dim CalcZAF_ZAF_Devs() As Single

Dim Binary_ZA_Devs() As Single
Dim CalcZAF_ZA_Devs() As Single

Dim Binary_F_Devs() As Single
Dim CalcZAF_F_Devs() As Single

Dim PENEPMA_Analysis As TypeAnalysis
Dim PENEPMA_Sample(1 To 1) As TypeSample
Dim Penepma_TmpSample(1 To 1) As TypeSample
Dim PENEPMA_OldSample(1 To 1) As TypeSample

Dim PENEPMA_SampleA(1 To 1) As TypeSample
Dim PENEPMA_SampleB(1 To 1) As TypeSample
Dim PENEPMA_SampleBStd(1 To 1) As TypeSample

Sub Penepma12Load()
' Load the form

ierror = False
On Error GoTo Penepma12LoadError

Dim i As Integer
Dim astring As String

Static initialized As Boolean

Call InitSample(PENEPMA_Sample())
Call InitSample(Penepma_TmpSample())

Call InitSample(PENEPMA_SampleA())
Call InitSample(PENEPMA_SampleB())
Call InitSample(PENEPMA_SampleBStd())

' Check for valid folders
astring$ = Dir$(PENDBASE_Path$, vbDirectory)
If astring$ = vbNullString Then
msg$ = "The Pendbase data files are not found in the specified folder: " & PENDBASE_Path$ & vbCrLf
msg$ = msg$ & "Please contact Probe Software to obtain the Penelope application files, copy them to the specified location and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Load"
ierror = True
Exit Sub
End If

astring$ = Dir$(PENEPMA_Path$, vbDirectory)
If astring$ = vbNullString Then
msg$ = "The Penepma application files are not found in the specified folder: " & PENEPMA_Path$ & vbCrLf
msg$ = msg$ & "Please contact Probe Software to obtain the Penelope application files, copy them to the specified location and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Load"
ierror = True
Exit Sub
End If

astring$ = Dir$(PENEPMA_Root$ & "\Penfluor", vbDirectory)
If astring$ = vbNullString Then
msg$ = "The Penepma Penfluor application files are not found in the specified folder: " & PENEPMA_Root$ & "\Penfluor" & vbCrLf
msg$ = msg$ & "Please contact Probe Software to obtain the Penelope application files, copy them to the specified location and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Load"
ierror = True
Exit Sub
End If

astring$ = Dir$(PENEPMA_Root$ & "\Fanal", vbDirectory)
If astring$ = vbNullString Then
msg$ = "The Penepma Fanal application files are not found in the specified folder: " & PENEPMA_Root$ & "\Fanal" & vbCrLf
msg$ = msg$ & "Please contact Probe Software to obtain the Penelope application files, copy them to the specified location and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Load"
ierror = True
Exit Sub
End If

' Make Fanal sub folders if necessary
If Dir$(PENEPMA_Root$ & "\Fanal\couple", vbDirectory) = vbNullString Then
MkDir PENEPMA_Root$ & "\Fanal\couple"
End If

If Dir$(PENEPMA_Root$ & "\Fanal\matrix", vbDirectory) = vbNullString Then
MkDir PENEPMA_Root$ & "\Fanal\matrix"
End If

If Dir$(PENEPMA_Root$ & "\Fanal\boundary", vbDirectory) = vbNullString Then
MkDir PENEPMA_Root$ & "\Fanal\boundary"
End If

' Init default values
If Not initialized Then
Call Penepma12Init
If ierror Then Exit Sub
initialized = True
End If

' Load the standard list boxes
Call StandardLoadList(FormPENEPMA12.ListAvailableStandardsA)
If ierror Then Exit Sub

Call StandardLoadList(FormPENEPMA12.ListAvailableStandardsB)
If ierror Then Exit Sub

Call StandardLoadList(FormPENEPMA12.ListAvailableStandardsBStd)
If ierror Then Exit Sub

' Load selected materials (if any)
If MaterialSelectedA% > 0 Then
For i% = 0 To FormPENEPMA12.ListAvailableStandardsA.ListCount - 1
If FormPENEPMA12.ListAvailableStandardsA.ItemData(i%) = MaterialSelectedA% Then
FormPENEPMA12.ListAvailableStandardsA.Selected(i%) = True
End If
Next i%
End If

If MaterialSelectedB% > 0 Then
For i% = 0 To FormPENEPMA12.ListAvailableStandardsB.ListCount - 1
If FormPENEPMA12.ListAvailableStandardsB.ItemData(i%) = MaterialSelectedB% Then
FormPENEPMA12.ListAvailableStandardsB.Selected(i%) = True
End If
Next i%
End If

If MaterialSelectedBStd% > 0 Then
For i% = 0 To FormPENEPMA12.ListAvailableStandardsBStd.ListCount - 1
If FormPENEPMA12.ListAvailableStandardsBStd.ItemData(i%) = MaterialSelectedBStd% Then
FormPENEPMA12.ListAvailableStandardsBStd.Selected(i%) = True
End If
Next i%
End If

' Load PENDBASE options
' Fcb is the number of effective electrons that participate in a plasmon excitation
' Wcb is the plasmon energy, for this reason they are both 0 for insulators
FormPENEPMA12.TextMaterialDensityA.Text = Format$(MaterialDensityA#)
FormPENEPMA12.TextMaterialFcbA.Text = Format$(MaterialFcbA#)
FormPENEPMA12.TextMaterialWcbA.Text = Format$(MaterialWcbA#)

FormPENEPMA12.TextMaterialDensityB.Text = Format$(MaterialDensityB#)
FormPENEPMA12.TextMaterialFcbB.Text = Format$(MaterialFcbB#)
FormPENEPMA12.TextMaterialWcbB.Text = Format$(MaterialWcbB#)

FormPENEPMA12.TextMaterialDensityBStd.Text = Format$(MaterialDensityBStd#)
FormPENEPMA12.TextMaterialFcbBStd.Text = Format$(MaterialFcbBStd#)
FormPENEPMA12.TextMaterialWcbBStd.Text = Format$(MaterialWcbBStd#)

' Load material file names (filenames only, no path)
FormPENEPMA12.TextMaterialFileA.Text = MaterialFileA$
FormPENEPMA12.TextMaterialFileB.Text = MaterialFileB$
FormPENEPMA12.TextMaterialFileBStd.Text = MaterialFileBStd$

' Load densities
FormPENEPMA12.ListAtomicDensitiesA.Clear
For i% = 0 To MAXELM% - 1
FormPENEPMA12.ListAtomicDensitiesA.AddItem "Density of " & Symup$(i% + 1) & " equals " & Format$(AllAtomicDensities!(i% + 1))
Next i%
FormPENEPMA12.ListAtomicDensitiesA.ListIndex = 29 - 1        ' default = Cu

FormPENEPMA12.ListAtomicDensitiesB.Clear
For i% = 0 To MAXELM% - 1
FormPENEPMA12.ListAtomicDensitiesB.AddItem "Density of " & Symup$(i% + 1) & " equals " & Format$(AllAtomicDensities!(i% + 1))
Next i%
FormPENEPMA12.ListAtomicDensitiesB.ListIndex = 27 - 1        ' default = Co

FormPENEPMA12.ListAtomicDensitiesBStd.Clear
For i% = 0 To MAXELM% - 1
FormPENEPMA12.ListAtomicDensitiesBStd.AddItem "Density of " & Symup$(i% + 1) & " equals " & Format$(AllAtomicDensities!(i% + 1))
Next i%
FormPENEPMA12.ListAtomicDensitiesBStd.ListIndex = 27 - 1        ' default = Co

' Add the list box items for running Penfluor, Fitall and Fanal
FormPENEPMA12.ComboElementStd.Clear
For i% = 0 To MAXELM% - 1
FormPENEPMA12.ComboElementStd.AddItem Symup$(i% + 1)
Next i%
FormPENEPMA12.ComboElementStd.ListIndex = MaterialMeasuredElement% - 1

FormPENEPMA12.ComboXRayStd.Clear
For i% = 0 To MAXRAY% - 2
FormPENEPMA12.ComboXRayStd.AddItem Xraylo$(i% + 1)
Next i%
FormPENEPMA12.ComboXRayStd.ListIndex = MaterialMeasuredXray% - 1

FormPENEPMA12.TextBeamTakeoff.Text = MaterialMeasuredTakeoff#
FormPENEPMA12.TextBeamEnergy.Text = MaterialMeasuredEnergy#

' Load minimum energy
FormPENEPMA12.TextPenepmaMinimumElectronEnergy.Text = Format$(PenepmaMinimumElectronEnergy!)

FormPENEPMA12.TextSimulationTime.Text = MaterialSimulationTime#
FormPENEPMA12.TextSimulationShowers.Text = Format$(MaterialSimulationShowers#, e71$)

' Load parameter file names (filenames only, no path)
FormPENEPMA12.TextParameterFileA.Text = ParameterFileA$
FormPENEPMA12.TextParameterFileB.Text = ParameterFileB$
FormPENEPMA12.TextParameterFileBStd.Text = ParameterFileBStd$

FormPENEPMA12.TextMeasuredMicrons.Text = MaterialMeasuredDistance#
FormPENEPMA12.TextMeasuredPoints.Text = MaterialMeasuredGridPoints%

' Excel output option
If SendToExcel Then
FormPENEPMA12.CheckSendToExcel.Value = vbChecked
Else
FormPENEPMA12.CheckSendToExcel.Value = vbUnchecked
End If

If UseLogScale Then
FormPENEPMA12.CheckUseLogScale.Value = vbChecked
Else
FormPENEPMA12.CheckUseLogScale.Value = vbUnchecked
End If

If UseGridLines Then
FormPENEPMA12.CheckUseGridLines.Value = vbChecked
Else
FormPENEPMA12.CheckUseGridLines.Value = vbUnchecked
End If

' Set timer to check on Penfluor progress (if simulation still running)
If SimulationInProgress Then
FormPENEPMA12.Timer1.Interval = BIT16&
End If

' Load graph defaults
Call Penepma12PlotLoad_PE
If ierror Then Exit Sub

Exit Sub

' Errors
Penepma12LoadError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12Load"
ierror = True
Exit Sub

End Sub

Sub Penepma12Save()
' Save the form

ierror = False
On Error GoTo Penepma12SaveError

Dim ip As Integer, ipp As Integer
Dim esym As String, xsym As String

Static userwarned As Boolean

' Save selected materials (if selected)
If FormPENEPMA12.ListAvailableStandardsA.ListIndex >= 0 And FormPENEPMA12.ListAvailableStandardsA.ListCount >= 1 Then
MaterialSelectedA% = FormPENEPMA12.ListAvailableStandardsA.ItemData(FormPENEPMA12.ListAvailableStandardsA.ListIndex)
End If

If FormPENEPMA12.ListAvailableStandardsB.ListIndex >= 0 And FormPENEPMA12.ListAvailableStandardsB.ListCount >= 1 Then
MaterialSelectedB% = FormPENEPMA12.ListAvailableStandardsB.ItemData(FormPENEPMA12.ListAvailableStandardsB.ListIndex)
End If

If FormPENEPMA12.ListAvailableStandardsBStd.ListIndex >= 0 And FormPENEPMA12.ListAvailableStandardsBStd.ListCount >= 1 Then
MaterialSelectedBStd% = FormPENEPMA12.ListAvailableStandardsBStd.ItemData(FormPENEPMA12.ListAvailableStandardsBStd.ListIndex)
End If

If Val(FormPENEPMA12.TextMaterialDensityA.Text) <= 0# Or Val(FormPENEPMA12.TextMaterialDensityA.Text) > MAXDENSITY# Then
msg$ = "Material A Density is out of range (must be greater than 0 and less than " & Format$(MAXDENSITY#) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Save"
ierror = True
Exit Sub
Else
MaterialDensityA# = Val(FormPENEPMA12.TextMaterialDensityA.Text)
End If

If Val(FormPENEPMA12.TextMaterialDensityB.Text) <= 0# Or Val(FormPENEPMA12.TextMaterialDensityB.Text) > MAXDENSITY# Then
msg$ = "Material B Density is out of range (must be greater than 0 and less than " & Format$(MAXDENSITY#) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Save"
ierror = True
Exit Sub
Else
MaterialDensityB# = Val(FormPENEPMA12.TextMaterialDensityB.Text)
End If

If Val(FormPENEPMA12.TextMaterialDensityBStd.Text) <= 0# Or Val(FormPENEPMA12.TextMaterialDensityBStd.Text) > MAXDENSITY# Then
msg$ = "Material B Std Density is out of range (must be greater than 0 and less than " & Format$(MAXDENSITY#) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Save"
ierror = True
Exit Sub
Else
MaterialDensityBStd# = Val(FormPENEPMA12.TextMaterialDensityBStd.Text)
End If

If Val(FormPENEPMA12.TextMaterialFcbA.Text) < 0# Or Val(FormPENEPMA12.TextMaterialFcbA.Text) > 100# Then
msg$ = "Material A Oscillator Strength is out of range (must be between 0 and 100)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Save"
ierror = True
Exit Sub
Else
MaterialFcbA# = Val(FormPENEPMA12.TextMaterialFcbA.Text)
End If

If Val(FormPENEPMA12.TextMaterialFcbB.Text) < 0# Or Val(FormPENEPMA12.TextMaterialFcbB.Text) > 100# Then
msg$ = "Material B Oscillator Strength is out of range (must be between 0 and 100)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Save"
ierror = True
Exit Sub
Else
MaterialFcbB# = Val(FormPENEPMA12.TextMaterialFcbB.Text)
End If

If Val(FormPENEPMA12.TextMaterialFcbBStd.Text) < 0# Or Val(FormPENEPMA12.TextMaterialFcbBStd.Text) > 100# Then
msg$ = "Material B Std Oscillator Strength is out of range (must be between 0 and 100)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Save"
ierror = True
Exit Sub
Else
MaterialFcbBStd# = Val(FormPENEPMA12.TextMaterialFcbBStd.Text)
End If

If Val(FormPENEPMA12.TextMaterialWcbA.Text) < 0# Or Val(FormPENEPMA12.TextMaterialWcbA.Text) > 1000# Then
msg$ = "Material A Oscillator Energy is out of range (must be between 0 and 1000)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Save"
ierror = True
Exit Sub
Else
MaterialWcbA# = Val(FormPENEPMA12.TextMaterialWcbA.Text)
End If

If Val(FormPENEPMA12.TextMaterialWcbB.Text) < 0# Or Val(FormPENEPMA12.TextMaterialWcbB.Text) > 1000# Then
msg$ = "Material B Oscillator Energy is out of range (must be between 0 and 1000)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Save"
ierror = True
Exit Sub
Else
MaterialWcbB# = Val(FormPENEPMA12.TextMaterialWcbB.Text)
End If

If Val(FormPENEPMA12.TextMaterialWcbBStd.Text) < 0# Or Val(FormPENEPMA12.TextMaterialWcbBStd.Text) > 1000# Then
msg$ = "Material B Std Oscillator Energy is out of range (must be between 0 and 1000)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Save"
ierror = True
Exit Sub
Else
MaterialWcbBStd# = Val(FormPENEPMA12.TextMaterialWcbBStd.Text)
End If

' Save the material files
MaterialFileA$ = FormPENEPMA12.TextMaterialFileA.Text
MaterialFileB$ = FormPENEPMA12.TextMaterialFileB.Text
MaterialFileBStd$ = FormPENEPMA12.TextMaterialFileBStd.Text

' Save pure element std element and x-ray
esym$ = FormPENEPMA12.ComboElementStd.Text
ip% = IPOS1(MAXELM%, esym$, Symlo$())
If ip% = 0 Then GoTo Penepma12SaveBadElement
MaterialMeasuredElement% = ip%

' Check for a valid x-ray symbol
xsym$ = FormPENEPMA12.ComboXRayStd.Text
ipp% = IPOS1(MAXRAY% - 1, xsym$, Xraylo$())
If ipp% = 0 Then GoTo Penepma12SaveBadXray
MaterialMeasuredXray% = ipp%

' Save the xray line as a default for this element (for this session)
Deflin$(ip%) = xsym$

If Val(FormPENEPMA12.TextBeamTakeoff.Text) < MINTAKEOFF! Or Val(FormPENEPMA12.TextBeamTakeoff.Text) > MAXTAKEOFF! Then
msg$ = "Takeoff angle is out of range (must be between " & Format$(MINTAKEOFF!) & " and " & Format$(MAXTAKEOFF!) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Save"
ierror = True
Exit Sub
Else
MaterialMeasuredTakeoff# = Val(FormPENEPMA12.TextBeamTakeoff.Text)
End If

' Save minimum energy
If Val(FormPENEPMA12.TextPenepmaMinimumElectronEnergy.Text) < 0.001 Or Val(FormPENEPMA12.TextPenepmaMinimumElectronEnergy.Text) > 10# Then
msg$ = "Penepma minimum electron energy (for monte-carlo simulations) is out of range (must be between 0.001 and 10 keV, default=1 keV)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Save"
ierror = True
Exit Sub
Else
PenepmaMinimumElectronEnergy! = Val(FormPENEPMA12.TextPenepmaMinimumElectronEnergy.Text)
End If

If Val(FormPENEPMA12.TextBeamEnergy.Text) < MINKILOVOLTS! Or Val(FormPENEPMA12.TextBeamEnergy.Text) > MAXKILOVOLTS! Then
msg$ = "Beam Energy is out of range (must be between " & Format$(MINKILOVOLTS!) & " and " & Format$(MAXKILOVOLTS!) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Save"
ierror = True
Exit Sub
Else
MaterialMeasuredEnergy# = Val(FormPENEPMA12.TextBeamEnergy.Text)
If MaterialMeasuredEnergy# > 50 And Not userwarned Then
msg$ = "Beam Energy is beyond default PAR file modeling range (5 to 50 keV). Be aware that extrapolation outside this range may result in decreased accuracy."
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12Save"
userwarned = True
End If
End If

If Val(FormPENEPMA12.TextSimulationTime.Text) < 100# Or Val(FormPENEPMA12.TextSimulationTime.Text) > SECPERDAY# Then
msg$ = "Simulation Time (in sec) is out of range (must be between 100 and " & Format$(SECPERDAY#) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Save"
ierror = True
Exit Sub
Else
MaterialSimulationTime# = Val(FormPENEPMA12.TextSimulationTime.Text)
End If

If Val(FormPENEPMA12.TextSimulationShowers.Text) < 100# Or Val(FormPENEPMA12.TextSimulationShowers.Text) > MAXLONG& Then
msg$ = "Simulation Showers (in electrons) is out of range (must be between 100 and " & Format$(MAXLONG&) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Save"
ierror = True
Exit Sub
Else
MaterialSimulationShowers# = Val(FormPENEPMA12.TextSimulationShowers.Text)
End If

' Save the parameter files
ParameterFileA$ = FormPENEPMA12.TextParameterFileA.Text
ParameterFileB$ = FormPENEPMA12.TextParameterFileB.Text
ParameterFileBStd$ = FormPENEPMA12.TextParameterFileBStd.Text

' Note: if MaterialMeasuredDistance# = 0 then the modified Fanal will output exponential distances starting at 10 nm
If Val(FormPENEPMA12.TextMeasuredMicrons.Text) < 0# Or Val(FormPENEPMA12.TextMeasuredMicrons.Text) > 5000 Then
msg$ = "Measured microns distance is out of range (must be between 0 and 5000 microns, enter 0 for exponential distances)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Save"
ierror = True
Exit Sub
Else
MaterialMeasuredDistance# = Val(FormPENEPMA12.TextMeasuredMicrons.Text)
End If

If Val(FormPENEPMA12.TextMeasuredPoints.Text) < 10# Or Val(FormPENEPMA12.TextMeasuredPoints.Text) > 2000 Then
msg$ = "Measured grid points is out of range (must be between 10 and 2000 points)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Save"
ierror = True
Exit Sub
Else
MaterialMeasuredGridPoints% = Val(FormPENEPMA12.TextMeasuredPoints.Text)
End If

' Excel output option
If FormPENEPMA12.CheckSendToExcel.Value = vbChecked Then
SendToExcel = True
Else
SendToExcel = False
End If

If FormPENEPMA12.CheckUseLogScale.Value = vbChecked Then
UseLogScale = True
Else
UseLogScale = False
End If

If FormPENEPMA12.CheckUseGridLines.Value = vbChecked Then
UseGridLines = True
Else
UseGridLines = False
End If

Exit Sub

' Errors
Penepma12SaveError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12Save"
ierror = True
Exit Sub

msg$ = "Too many materials selected for output"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Save"
ierror = True
Exit Sub

Penepma12SaveBadElement:
msg$ = "Element " & esym$ & " is not a valid element symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Save"
ierror = True
Exit Sub

Penepma12SaveBadXray:
msg$ = "Xray " & xsym$ & " is not a valid element symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12Save"
ierror = True
Exit Sub

End Sub

Sub Penepma12Init()
' Initialize the variables

ierror = False
On Error GoTo Penepma12InitError

icancelauto = False

' Pendbase init
If MaterialSelectedA% = 0 Then MaterialSelectedA% = 529   ' Cu
If MaterialFileA$ = vbNullString Then MaterialFileA$ = "Cu.mat"
If MaterialDensityA# = 0# Then MaterialDensityA# = 8.96

If MaterialSelectedB% = 0 Then MaterialSelectedB% = 527   ' Co
If MaterialFileB$ = vbNullString Then MaterialFileB$ = "Co.mat"
If MaterialDensityB# = 0# Then MaterialDensityB# = 8.9

If MaterialSelectedBStd% = 0 Then MaterialSelectedBStd% = 527   ' Co
If MaterialFileBStd$ = vbNullString Then MaterialFileBStd$ = "Co.mat"
If MaterialDensityBStd# = 0# Then MaterialDensityBStd# = 8.9

If MaterialFcbA# = 0# Then MaterialFcbA# = 0#     ' zero means use calculated default
If MaterialWcbA# = 0# Then MaterialWcbA# = 0#     ' zero means use calculated default
If MaterialFcbB# = 0# Then MaterialFcbB# = 0#     ' zero means use calculated default
If MaterialWcbB# = 0# Then MaterialWcbB# = 0#     ' zero means use calculated default
If MaterialFcbBStd# = 0# Then MaterialFcbBStd# = 0#     ' zero means use calculated default
If MaterialWcbBStd# = 0# Then MaterialWcbBStd# = 0#     ' zero means use calculated default

If MaterialMeasuredElement% = 0 Then MaterialMeasuredElement% = 27  ' Co
If MaterialMeasuredXray% = 0 Then MaterialMeasuredXray% = 1  ' Ka

If MaterialMeasuredTakeoff# = 0# Then MaterialMeasuredTakeoff# = DefaultTakeOff!
If MaterialMeasuredEnergy# = 0# Then MaterialMeasuredEnergy# = DefaultKiloVolts!
If MaterialSimulationTime# = 0# Then MaterialSimulationTime# = 3600#               ' in sec
If MaterialSimulationShowers# = 0# Then MaterialSimulationShowers# = 2000000000#  ' in electron trajectories (50K is typical for single Xenon, processor, pers. comm. Llovet, 2012)

If MaterialSimulationOverhead# = 0# Then MaterialSimulationOverhead# = 70#     ' assume Penfluor overhead of 70 sec for each voltage (70 sec for 2 elements)
If FitParametersTime# = 0 Then FitParametersTime# = 50    ' in sec for for FitAll calculation

If Trim$(PenfluorInputFile$) = vbNullString Then PenfluorInputFile$ = PENEPMA_Root$ & "\Penfluor\Penfluor.inp"
If Trim$(PenfluorOutputFile$) = vbNullString Then PenfluorOutputFile$ = PENEPMA_Root$ & "\Penfluor\Penfluor.in"
If Trim$(FanalInputFile$) = vbNullString Then FanalInputFile$ = PENEPMA_Root$ & "\Fanal\Fanal.in"

If ParameterFileA$ = vbNullString Then ParameterFileA$ = "Cu.par"
If ParameterFileB$ = vbNullString Then ParameterFileB$ = "Co.par"
If ParameterFileBStd$ = vbNullString Then ParameterFileBStd$ = "Co.par"

If MaterialMeasuredDistance# = 0# Then MaterialMeasuredDistance# = 50#      ' in microns
If MaterialMeasuredGridPoints% = 0 Then MaterialMeasuredGridPoints% = 50      ' number of points (first point at dist/points)

If Trim$(FANAL_IN_File$) = vbNullString Then FANAL_IN_File$ = PENEPMA_Root$ & "\Fanal\fanal.in"
If Trim$(VACS_DAT_File$) = vbNullString Then VACS_DAT_File$ = PENEPMA_Root$ & "\Fanal\vacs.dat"
If Trim$(RANGES_DAT_File$) = vbNullString Then RANGES_DAT_File$ = PENEPMA_Root$ & "\Fanal\ranges.dat"
If Trim$(MIXED_DAT_File$) = vbNullString Then MIXED_DAT_File$ = PENEPMA_Root$ & "\Fanal\mixed.dat"
If Trim$(KRATIOS_DAT_File$) = vbNullString Then KRATIOS_DAT_File$ = PENEPMA_Root$ & "\Fanal\k-ratios.dat"
If Trim$(FLUORMAT1_PAR_File$) = vbNullString Then FLUORMAT1_PAR_File$ = PENEPMA_Root$ & "\Fanal\fluormat1.par"
If Trim$(FLUORMAT2_PAR_File$) = vbNullString Then FLUORMAT2_PAR_File$ = PENEPMA_Root$ & "\Fanal\fluormat2.par"
If Trim$(FLUORMAT3_PAR_File$) = vbNullString Then FLUORMAT3_PAR_File$ = PENEPMA_Root$ & "\Fanal\fluormat3.par"
If Trim$(ATCOEFFS_DAT_File$) = vbNullString Then ATCOEFFS_DAT_File$ = PENEPMA_Root$ & "\Fanal\atcoeffs.dat"

If BinaryMethod% = 0 Then BinaryMethod% = 0
If BinaryElement1% = 0 Then BinaryElement1% = 26    ' Fe
If BinaryElement2% = 0 Then BinaryElement2% = 29    ' Cu
If CalculateDoNotOverwritePAR = 0 Then CalculateDoNotOverwritePAR = True
If CalculateOnlyOverwriteLowerPrecisionPAR = 0 Then CalculateOnlyOverwriteLowerPrecisionPAR = True
If CalculateOnlyOverwriteHigherMinimumEnergyPAR = 0 Then CalculateOnlyOverwriteHigherMinimumEnergyPAR = True

If CalculateDoNotOverwriteTXT = 0 Then CalculateDoNotOverwriteTXT = True
If CalculateForMatrixRange = 0 Then CalculateForMatrixRange = False
If CalculateFromFormulaOrStandard% = 0 Then CalculateFromFormulaOrStandard% = 1

If ExtractMethod% = 0 Then ExtractMethod% = 1       ' default to matrix extraction
If ExtractElement% = 0 Then ExtractElement% = 22    ' Ti
If ExtractMatrix% = 0 Then ExtractMatrix% = 14      ' Si

' Boundary binary
If ExtractMatrixA1% = 0 Then ExtractMatrixA1% = 8     ' O
If ExtractMatrixA2% = 0 Then ExtractMatrixA2% = 14    ' Si
If ExtractMatrixB1% = 0 Then ExtractMatrixB1% = 8     ' O
If ExtractMatrixB2% = 0 Then ExtractMatrixB2% = 22    ' Ti

If ExtractForSpecifiedRange = 0 Then ExtractForSpecifiedRange = False

UseGridLines = True     ' first time only

FormPENEPMA12.LabelProgress.Caption = vbNullString
FormPENEPMA12.LabelRemainingTime.Caption = vbNullString

' Load Penepma08/12 atomic weights (for self consistent calculations)
Call Penepma12AtomicWeights
If ierror Then Exit Sub

' Load default minimum electron energy (now read from INI file)
If PenepmaMinimumElectronEnergy! = 0# Then PenepmaMinimumElectronEnergy! = 1#     ' 1 keV is default

Exit Sub

' Errors
Penepma12InitError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12Init"
ierror = True
Exit Sub

End Sub

Sub Penepma12CreateMaterial(mode As Integer)
' Create material file (*.mat)(Penepma12 version) for material A and B (and pure element standard for material B)
'  mode = 1 = material A
'  mode = 2 = material B
'  mode = 3 = material B (std)

ierror = False
On Error GoTo Penepma12CreateMaterialError

Dim stdnum As Integer
Dim astring As String

icancelauto = False

' Check that at least one material is selected for each
If mode% = 1 Then
If FormPENEPMA12.ListAvailableStandardsA.ListIndex < 0 Then GoTo Penepma12CreateMaterialNoMaterialA
If FormPENEPMA12.ListAvailableStandardsA.ListCount < 1 Then GoTo Penepma12CreateMaterialNoMaterialA
MaterialSelectedA% = FormPENEPMA12.ListAvailableStandardsA.ItemData(FormPENEPMA12.ListAvailableStandardsA.ListIndex)
If MaterialSelectedA% = 0 Then GoTo Penepma12CreateMaterialNoMaterialA
stdnum% = MaterialSelectedA%
End If

If mode% = 2 Then
If FormPENEPMA12.ListAvailableStandardsB.ListIndex < 0 Then GoTo Penepma12CreateMaterialNoMaterialB
If FormPENEPMA12.ListAvailableStandardsB.ListCount < 1 Then GoTo Penepma12CreateMaterialNoMaterialB
MaterialSelectedB% = FormPENEPMA12.ListAvailableStandardsB.ItemData(FormPENEPMA12.ListAvailableStandardsB.ListIndex)
If MaterialSelectedB% = 0 Then GoTo Penepma12CreateMaterialNoMaterialB
stdnum% = MaterialSelectedB%
End If

If mode% = 3 Then
If FormPENEPMA12.ListAvailableStandardsBStd.ListIndex < 0 Then GoTo Penepma12CreateMaterialNoMaterialBStd
If FormPENEPMA12.ListAvailableStandardsBStd.ListCount < 1 Then GoTo Penepma12CreateMaterialNoMaterialBStd
MaterialSelectedBStd% = FormPENEPMA12.ListAvailableStandardsBStd.ItemData(FormPENEPMA12.ListAvailableStandardsBStd.ListIndex)
If MaterialSelectedBStd% = 0 Then GoTo Penepma12CreateMaterialNoMaterialBStd
stdnum% = MaterialSelectedBStd%
End If

' Get composition based on standard number
Call StandardGetMDBStandard(stdnum%, PENEPMA_Sample())
If ierror Then Exit Sub

' Update status
Call IOStatusAuto("Creating material input files based on standard " & Str$(PENEPMA_Sample(1).number%) & " " & PENEPMA_Sample(1).Name$ & "...")
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(0), MaterialInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

' Save material filenames
astring$ = PENEPMA_Sample(1).Name$
Call MiscModifyStringToFilename(astring$)
If mode% = 1 Then MaterialFileA$ = astring$ & ".mat"
If mode% = 2 Then MaterialFileB$ = astring$ & ".mat"
If mode% = 3 Then MaterialFileBStd$ = astring$ & ".mat"

FormPENEPMA12.LabelProgress.Caption = "Creating Material File " & astring$ & ".mat"
FormPENEPMA12.LabelRemainingTime.Caption = vbNullString

' Make material INP file
Screen.MousePointer = vbHourglass
Call Penepma12CreateMaterialINP(mode%, PENEPMA_Sample())
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

' Create and run the necessary batch files
Screen.MousePointer = vbHourglass
Call Penepma12CreateMaterialBatch(mode%, Int(0))
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

Exit Sub

' Errors
Penepma12CreateMaterialError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CreateMaterial"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma12CreateMaterialNoMaterialA:
msg$ = "No Material A selected for output"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CreateMaterial"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma12CreateMaterialNoMaterialB:
msg$ = "No Material B selected for output"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CreateMaterial"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma12CreateMaterialNoMaterialBStd:
msg$ = "No Material B Std selected for output"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CreateMaterial"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma12CreateMaterialINP(n As Integer, sample() As TypeSample)
' Create a single INP (redirected keyboard input) file based on the specified standard material
'  n = 1 = Material A
'  n = 2 = Material B
'  n = 3 = Material B Std

ierror = False
On Error GoTo Penepma12CreateMaterialINPError

Dim i As Integer
Dim tfilename As String, astring As String

' Load file name
astring$ = "material" & Format$(n%) & ".INP"
tfilename$ = PENDBASE_Path$ & "\" & astring$
Open tfilename$ For Output As #Temp1FileNumber%

' Output configuration
Print #Temp1FileNumber%, "1"                             ' enter composition from keyboard
Print #Temp1FileNumber%, Left$(sample(1).Name$, 60)      ' material name
Print #Temp1FileNumber%, Format$(sample(1).LastChan%)    ' number of elements in composition

' If more than one element enter composition
If sample(1).LastChan% = 1 Then
Print #Temp1FileNumber%, Format$(sample(1).AtomicNums%(1))

Else
Print #Temp1FileNumber%, "2"   ' enter by weight fraction

' Output composition of material
For i% = 1 To sample(1).LastChan%
If sample(1).ElmPercents!(i%) < PENEPMA_MINPERCENT! Then sample(1).ElmPercents!(i%) = PENEPMA_MINPERCENT!
astring$ = Format$(sample(1).AtomicNums%(i%)) & VbComma$ & Trim$(MiscAutoFormat$(sample(1).ElmPercents!(i%) / 100#))
Print #Temp1FileNumber%, astring$
Next i%
End If

' Do not change mean excitation energy
Print #Temp1FileNumber%, "2"

' Output densities
If n% = 1 Then Print #Temp1FileNumber%, Trim$(Str$(MaterialDensityA#))                  ' density of material A
If n% = 2 Then Print #Temp1FileNumber%, Trim$(Str$(MaterialDensityB#))                  ' density of material B
If n% = 3 Then Print #Temp1FileNumber%, Trim$(Str$(MaterialDensityBStd#))               ' density of material B Std

' If values are zero, have Material.exe calculate oscillator strength and energy
If n% = 1 Then
If MaterialFcbA# = 0# And MaterialWcbA# = 0# Then
Print #Temp1FileNumber%, "2"
Else
If MaterialFcbA# = 0# Or MaterialWcbA# = 0# Then GoTo Penepma12CreatematerialINPZero
Print #Temp1FileNumber%, "1"
Print #Temp1FileNumber%, Trim$(Str$(MaterialFcbA#)) & VbComma$ & Trim$(Str$(MaterialWcbA#))
End If
End If

If n% = 2 Then
If MaterialFcbB# = 0# And MaterialWcbB# = 0# Then
Print #Temp1FileNumber%, "2"
Else
If MaterialFcbB# = 0# Or MaterialWcbB# = 0# Then GoTo Penepma12CreatematerialINPZero
Print #Temp1FileNumber%, "1"
Print #Temp1FileNumber%, Trim$(Str$(MaterialFcbB#)) & VbComma$ & Trim$(Str$(MaterialWcbB#))
End If
End If

If n% = 3 Then
If MaterialFcbBStd# = 0# And MaterialWcbBStd# = 0# Then
Print #Temp1FileNumber%, "2"
Else
If MaterialFcbBStd# = 0# Or MaterialWcbBStd# = 0# Then GoTo Penepma12CreatematerialINPZero
Print #Temp1FileNumber%, "1"
Print #Temp1FileNumber%, Trim$(Str$(MaterialFcbBStd#)) & VbComma$ & Trim$(Str$(MaterialWcbBStd#))
End If
End If

astring$ = "material" & Format$(n%) & ".mat"                    ' use same folder as MATERIAL.EXE
Print #Temp1FileNumber%, Left$(astring$, 80)                    ' material filename
Close #Temp1FileNumber%

Exit Sub

' Errors
Penepma12CreateMaterialINPError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CreateMaterialINP"
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12CreatematerialINPZero:
msg$ = "One of the material oscillator parameters is zero, please enter non-zero values for both or zero values for both and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CreateMaterialINP"
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12CreateMaterialBatch(n As Integer, method As Integer)
' Create and run material batch files
'  n = 1 = Material A
'  n = 2 = Material B
'  n = 3 = Material B Std
' method = 0 run normal
' method = 1 no message box

ierror = False
On Error GoTo Penepma12CreateMaterialBatchError

Dim bfilename As String, astring As String, bstring As String

icancelauto = False

' Run each material[n].inp file
If n% = 1 Then Call IOStatusAuto("Creating Material A file by running MATERIAL.EXE (this may take a while)...")
If n% = 2 Then Call IOStatusAuto("Creating Material B file by running MATERIAL.EXE (this may take a while)...")
If n% = 3 Then Call IOStatusAuto("Creating Material B Std file by running MATERIAL.EXE (this may take a while)...")
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(0), MaterialInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

bstring$ = "Checking for previous (material) batch file..."
bfilename$ = PENDBASE_Path$ & "\temp.bat"
If Dir$(bfilename$) <> vbNullString Then
bstring$ = "Deleting previous batch file..."
Kill bfilename$
End If

' Delete existing material file if it exists (to allow for proper check if created below)
bstring$ = "Checking for previous material file..."
If Dir$(PENDBASE_Path$ & "\" & "material" & Format$(n%) & ".mat") <> vbNullString Then
bstring$ = "Deleting previous material file..."
Kill PENDBASE_Path$ & "\" & "material" & Format$(n%) & ".mat"
End If

' Write batch file for running material.exe
Call MiscDelay(CDbl(0.2), Now)
bstring$ = "Creating new (material) batch file..."
Open bfilename$ For Append As #Temp1FileNumber%

astring$ = Left$(PENDBASE_Path$, 2)                             ' change to drive
Print #Temp1FileNumber%, astring$
astring$ = "cd " & VbDquote$ & PENDBASE_Path$ & VbDquote$       ' change to folder
Print #Temp1FileNumber%, astring$
astring$ = "material.exe < " & "material" & Format$(n%) & ".inp"
Print #Temp1FileNumber%, astring$
Close #Temp1FileNumber%

FormPENEPMA12.Timer1.Interval = 0.5 * MSECPERSEC#     ' update every 0.5 seconds
MaterialInProgress = True
FormPENEPMA12.LabelProgress.Caption = "Material In Progress!"
DoEvents

bfilename$ = PENDBASE_Path$ & "\temp.bat"
bstring$ = "Running new (material) batch file..."

' Start Material (/k executes but window remains, /c executes but terminates)
'PenepmaTaskID& = Shell("cmd.exe /k " & VbDquote$ & bfilename$ & VbDquote$, vbMinimizedNoFocus)
PenepmaTaskID& = Shell("cmd.exe /c " & VbDquote$ & bfilename$ & VbDquote$, vbMinimizedNoFocus)

' Now wait for the Material calculation to finish
Do Until Not MaterialInProgress
DoEvents
If icancelauto Then
MaterialInProgress = False
Call Penepma12UpdateForm
If ierror Then Exit Sub
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(0), MaterialInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If
Loop

' Check for created material file
Call MiscDelay(CDbl(0.2), Now)
bstring$ = "Checking for new material file..."
If Dir$(PENDBASE_Path$ & "\" & "material" & Format$(n%) & ".mat") = vbNullString Then GoTo Penepma12CreateMaterialBatchNoMaterialCreated

' Now copy all files to original material names (up to 20 characters)
Call IOStatusAuto("Copying temp material files to target material file...")
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(0), MaterialInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

' Delete existing temp batch file
bfilename$ = PENDBASE_Path$ & "\temp.bat"
bstring$ = "Deleting existing (material) batch file..."
If Dir$(bfilename$) <> vbNullString Then
Kill bfilename$
End If

' Copy defaults names to original names
Call MiscDelay(CDbl(0.2), Now)
bstring$ = "Creating new (copy) batch file..."
Open bfilename$ For Append As #Temp1FileNumber%

astring$ = Left$(PENDBASE_Path$, 2)                             ' change to drive
Print #Temp1FileNumber%, astring$
astring$ = "cd " & VbDquote$ & PENDBASE_Path$ & VbDquote$       ' change to folder
Print #Temp1FileNumber%, astring$

astring$ = "copy material" & Format$(n%) & ".mat "
If n% = 1 Then astring$ = astring$ & " " & VbDquote$ & MiscGetFileNameNoExtension$(MaterialFileA$) & ".mat" & VbDquote$
If n% = 2 Then astring$ = astring$ & " " & VbDquote$ & MiscGetFileNameNoExtension$(MaterialFileB$) & ".mat" & VbDquote$
If n% = 3 Then astring$ = astring$ & " " & VbDquote$ & MiscGetFileNameNoExtension$(MaterialFileBStd$) & ".mat" & VbDquote$
Print #Temp1FileNumber%, astring$
If n% = 1 Then astring$ = "copy " & VbDquote$ & PENDBASE_Path$ & "\" & MaterialFileA$ & VbDquote$ & " " & VbDquote$ & PENEPMA_Path$ & "\" & MaterialFileA$ & VbDquote$
If n% = 2 Then astring$ = "copy " & VbDquote$ & PENDBASE_Path$ & "\" & MaterialFileB$ & VbDquote$ & " " & VbDquote$ & PENEPMA_Path$ & "\" & MaterialFileB$ & VbDquote$
If n% = 3 Then astring$ = "copy " & VbDquote$ & PENDBASE_Path$ & "\" & MaterialFileBStd$ & VbDquote$ & " " & VbDquote$ & PENEPMA_Path$ & "\" & MaterialFileBStd$ & VbDquote$
Print #Temp1FileNumber%, astring$
Close #Temp1FileNumber%

Call IOStatusAuto("Copying material files to original file names...")
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(0), MaterialInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

' Run batch file synchronously to copy files
bstring$ = "Running new (copy) batch file..."
astring$ = VbDquote$ & PENDBASE_Path$ & "\temp.bat" & VbDquote$
Call ExecRun(astring$)
If ierror Then Exit Sub
Call MiscDelay(CDbl(0.2), Now)

' Confirm with user
Call IOStatusAuto(vbNullString)
If n% = 1 Then msg$ = "Material file " & MaterialFileA$ & " was created and saved to " & PENDBASE_Path$
If n% = 2 Then msg$ = "Material file " & MaterialFileB$ & " was created and saved to " & PENDBASE_Path$
If n% = 3 Then msg$ = "Material file " & MaterialFileBStd$ & " was created and saved to " & PENDBASE_Path$
Call IOWriteLog(msg$)

If n% = 1 Then msg$ = "Material file " & MaterialFileA$ & " was copied to " & PENEPMA_Path$
If n% = 2 Then msg$ = "Material file " & MaterialFileB$ & " was copied to " & PENEPMA_Path$
If n% = 3 Then msg$ = "Material file " & MaterialFileBStd$ & " was copied to " & PENEPMA_Path$
Call IOWriteLog(msg$)

' Update material file fields with material name
If n% = 1 Then FormPENEPMA12.TextMaterialFileA.Text = MaterialFileA$
If n% = 2 Then FormPENEPMA12.TextMaterialFileB.Text = MaterialFileB$
If n% = 3 Then FormPENEPMA12.TextMaterialFileBStd.Text = MaterialFileBStd$

If n% = 1 Then FormPENEPMA12.LabelProgress.Caption = "Material File " & MaterialFileA$ & " was created!"
If n% = 2 Then FormPENEPMA12.LabelProgress.Caption = "Material File " & MaterialFileB$ & " was created!"
If n% = 3 Then FormPENEPMA12.LabelProgress.Caption = "Material File " & MaterialFileBStd$ & " was created!"
FormPENEPMA12.LabelRemainingTime.Caption = vbNullString

' Confirm with user
If method% = 0 Then
If n% = 1 Then msg$ = "Material file " & MaterialFileA$ & " was copied to " & PENEPMA_Path$
If n% = 2 Then msg$ = "Material file " & MaterialFileB$ & " was copied to " & PENEPMA_Path$
If n% = 3 Then msg$ = "Material file " & MaterialFileBStd$ & " was copied to " & PENEPMA_Path$
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12CreateMaterialBatch"
End If

Exit Sub

' Errors
Penepma12CreateMaterialBatchError:
MsgBox Error$ & ", during " & bstring$, vbOKOnly + vbCritical, "Penepma12CreateMaterialBatch"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma12CreateMaterialBatchNoMaterials:
msg$ = "No materials selected for output"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CreateMaterialBatch"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma12CreateMaterialBatchNoMaterialCreated:
msg$ = "The specified material file was not created properly. Please click the PENDBASE Prompt button and type the following command to see the error: material.exe < " & "material" & Format$(n%) & ".inp"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CreateMaterialBatch"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma12BrowseMaterialFile(n As Integer, tForm As Form)
' Browse to a selected material file (*.mat) for creating an input file
'  n = 1 = Material A
'  n = 2 = Material B
'  n = 3 = Material B Std

ierror = False
On Error GoTo Penepma12BrowseMaterialFileError

Dim tfilename As String, ioextension As String

If n% = 1 Then tfilename$ = PENDBASE_Path$ & "\" & MaterialFileA$
If n% = 2 Then tfilename$ = PENDBASE_Path$ & "\" & MaterialFileB$
If n% = 3 Then tfilename$ = PENDBASE_Path$ & "\" & MaterialFileBStd$

ioextension$ = "MAT"
Call IOGetFileName(Int(2), ioextension$, tfilename$, tForm)
If ierror Then Exit Sub

' Check that the user did select a .mat file
If UCase$(MiscGetFileNameExtensionOnly$(tfilename$)) <> UCase$(".mat") Then GoTo Penepma12BrowseMaterialFileNotMAT

' Check that selected file is in Penepma path already, if not then copy there
If Trim$(UCase$(tfilename$)) <> Trim$(UCase$(PENEPMA_Path$ & "\" & MiscGetFileNameOnly$(tfilename$))) Then
FileCopy tfilename$, PENEPMA_Path$ & "\" & MiscGetFileNameOnly$(tfilename$)
End If

' Load to module and dialog
If n% = 1 Then MaterialFileA$ = MiscGetFileNameOnly$(tfilename$)
If n% = 2 Then MaterialFileB$ = MiscGetFileNameOnly$(tfilename$)
If n% = 3 Then MaterialFileBStd$ = MiscGetFileNameOnly$(tfilename$)

If n% = 1 Then FormPENEPMA12.TextMaterialFileA.Text = MaterialFileA$
If n% = 2 Then FormPENEPMA12.TextMaterialFileB.Text = MaterialFileB$
If n% = 3 Then FormPENEPMA12.TextMaterialFileBStd.Text = MaterialFileBStd$

Exit Sub

' Errors
Penepma12BrowseMaterialFileError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12BrowseMaterialFile"
ierror = True
Exit Sub

Penepma12BrowseMaterialFileNotMAT:
msg$ = "The selected file is not a MAT file. Please try again and select a file with the extension .MAT"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12BrowseParameterFile"
ierror = True
Exit Sub

End Sub

Sub Penepma12BrowseParameterFile(n As Integer, tForm As Form)
' Browse to a selected parameter file (*.par) for creating an input file for Fanal
'  n = 1 = Parameter A
'  n = 2 = Parameter B
'  n = 3 = Parameter B Std

ierror = False
On Error GoTo Penepma12BrowseParameterFileError

Dim tfilename As String, ioextension As String

If n% = 1 Then tfilename$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileA$
If n% = 2 Then tfilename$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileB$
If n% = 3 Then tfilename$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileBStd$

ioextension$ = "PAR"
Call IOGetFileName(Int(2), ioextension$, tfilename$, tForm)
If ierror Then Exit Sub

' Check that the user did select a .par file
If UCase$(MiscGetFileNameExtensionOnly$(tfilename$)) <> UCase$(".par") Then GoTo Penepma12BrowseParameterFileNotPAR

' Check that selected file is in Penfluor path already, if not then copy there
If Trim$(UCase$(tfilename$)) <> Trim$(UCase$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameOnly$(tfilename$))) Then
FileCopy tfilename$, PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameOnly$(tfilename$)

' Check if the <penfluor>.in file is present also, and if so, copy that too
If Dir$(MiscGetFileNameNoExtension$(tfilename$) & ".in") <> vbNullString Then
FileCopy MiscGetFileNameNoExtension$(tfilename$) & ".in", PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(tfilename$)) & ".in"
End If
End If

' Load to module and dialog
If n% = 1 Then ParameterFileA$ = MiscGetFileNameOnly$(tfilename$)
If n% = 2 Then ParameterFileB$ = MiscGetFileNameOnly$(tfilename$)
If n% = 3 Then ParameterFileBStd$ = MiscGetFileNameOnly$(tfilename$)

If n% = 1 Then FormPENEPMA12.TextParameterFileA.Text = ParameterFileA$
If n% = 2 Then FormPENEPMA12.TextParameterFileB.Text = ParameterFileB$
If n% = 3 Then FormPENEPMA12.TextParameterFileBStd.Text = ParameterFileBStd$

Exit Sub

' Errors
Penepma12BrowseParameterFileError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12BrowseParameterFile"
ierror = True
Exit Sub

Penepma12BrowseParameterFileNotPAR:
msg$ = "The selected file is not a PAR file. Please try again and select a file with the extension .PAR"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12BrowseParameterFile"
ierror = True
Exit Sub

End Sub

Sub Penepma12CreateMaterialFormula(n As Integer)
' Creates a material file from a formula composition
'  n = 1 = Material A
'  n = 2 = Material B
'  n = 3 = Material B Std

ierror = False
On Error GoTo Penepma12CreateMaterialFormulaError

icancelauto = False

' Load FormFORMULA
If n% = 1 Then FormFORMULA.Frame1.Caption = "Enter Formula String For PENEPMA Material A File:"
If n% = 2 Then FormFORMULA.Frame1.Caption = "Enter Formula String For PENEPMA Material B File:"
If n% = 3 Then FormFORMULA.Frame1.Caption = "Enter Formula String For PENEPMA Material Std File:"

' Get formula from user
FormFORMULA.Show vbModal
If icancel Then Exit Sub

' Return modified sample
Call FormulaReturnSample(PENEPMA_Sample())
If ierror Then Exit Sub

' Load name and number for this formula
Call MiscModifyStringToFilename(PENEPMA_Sample(1).Name$)
If ierror Then Exit Sub
If n% = 1 Then MaterialFileA$ = PENEPMA_Sample(1).Name$ & ".mat"
If n% = 2 Then MaterialFileB$ = PENEPMA_Sample(1).Name$ & ".mat"
If n% = 3 Then MaterialFileBStd$ = PENEPMA_Sample(1).Name$ & ".mat"

If n% = 1 Then MaterialSelectedA% = MAXINTEGER%     ' any non-zero number
If n% = 2 Then MaterialSelectedB% = MAXINTEGER%     ' any non-zero number
If n% = 3 Then MaterialSelectedBStd% = MAXINTEGER%     ' any non-zero number

Call IOStatusAuto("Creating material input file based on formula " & PENEPMA_Sample(1).Name$ & "...")
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(0), MaterialInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

FormPENEPMA12.LabelProgress.Caption = "Creating Material File " & PENEPMA_Sample(1).Name$ & ".mat"
FormPENEPMA12.LabelRemainingTime.Caption = vbNullString

' Make material INP file (always a single file)
Screen.MousePointer = vbHourglass
Call Penepma12CreateMaterialINP(n%, PENEPMA_Sample())
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

' Create and run the necessary batch files
Screen.MousePointer = vbHourglass
Call Penepma12CreateMaterialBatch(n%, Int(0))
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

Exit Sub

' Errors
Penepma12CreateMaterialFormulaError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CreateMaterialFormula"
ierror = True
Exit Sub

End Sub

Sub Penepma12RunPenFluor(n As Integer)
' Run Penfluor and Fitall for all three materials
'  n = 1 = Material A
'  n = 2 = Material B
'  n = 3 = Material B Std

ierror = False
On Error GoTo Penepma12RunPenfluorError

Dim tfilename As String

icancelauto = False

' Save the current mode number for time estimate
CurrentSimulationModeNumber% = n%

' Check for files
If Dir$(PENEPMA_Root$ & "\Penfluor\penfluor.exe") = vbNullString Then GoTo Penepma12RunPenfluorNoEXE
If Dir$(PENEPMA_Root$ & "\Penfluor\penfluor.geo") = vbNullString Then GoTo Penepma12RunPenfluorNoGEO

' Delete existing input files (to check for proper creation by Penfluor.exe)
If Dir$(PENEPMA_Root$ & "\Penfluor\gphcoxs.dat") <> vbNullString Then Kill PENEPMA_Root$ & "\Penfluor\gphcoxs.dat"
If Dir$(PENEPMA_Root$ & "\Penfluor\br.dat") <> vbNullString Then Kill PENEPMA_Root$ & "\Penfluor\br.dat"
If Dir$(PENEPMA_Root$ & "\Penfluor\xr.dat") <> vbNullString Then Kill PENEPMA_Root$ & "\Penfluor\xr.dat"
If Dir$(PENEPMA_Root$ & "\Penfluor\penepma.dat") <> vbNullString Then Kill PENEPMA_Root$ & "\Penfluor\penepma.dat"

' Delete existing parameter file (to check for proper creation by Fitall.exe)
If n% = 1 Then
If Dir$(PENEPMA_Root$ & "\Penfluor\material1.par") <> vbNullString Then Kill PENEPMA_Root$ & "\Penfluor\material1.par"
End If
If n% = 2 Then
If Dir$(PENEPMA_Root$ & "\Penfluor\material2.par") <> vbNullString Then Kill PENEPMA_Root$ & "\Penfluor\material2.par"
End If
If n% = 3 Then
If Dir$(PENEPMA_Root$ & "\Penfluor\material3.par") <> vbNullString Then Kill PENEPMA_Root$ & "\Penfluor\material3.par"
End If

' Check for Li, Be, B, C, N, O, F or Ne and adjust minimum energy if so
If FormPENEPMA12.CheckAutoAdjustMinimumEnergy.Value = vbChecked Then
If n% = 1 Then Call Penepma12AdjustMinimumEnergy(PENEPMA_Path$ & "\" & MaterialFileA$)
If n% = 2 Then Call Penepma12AdjustMinimumEnergy(PENEPMA_Path$ & "\" & MaterialFileB$)
If n% = 3 Then Call Penepma12AdjustMinimumEnergy(PENEPMA_Path$ & "\" & MaterialFileBStd$)
If ierror Then Exit Sub
End If

' Check for each material .mat file
If n% = 1 Then tfilename$ = PENEPMA_Path$ & "\" & MaterialFileA$
If n% = 2 Then tfilename$ = PENEPMA_Path$ & "\" & MaterialFileB$
If n% = 3 Then tfilename$ = PENEPMA_Path$ & "\" & MaterialFileBStd$

' Copy to temp name
If Dir$(tfilename$) = vbNullString Then GoTo Penepma12RunPenFluorMaterialFileNotFound
If n% = 1 Then FileCopy tfilename$, PENEPMA_Root$ & "\Penfluor\material1.mat"
If n% = 2 Then FileCopy tfilename$, PENEPMA_Root$ & "\Penfluor\material2.mat"
If n% = 3 Then FileCopy tfilename$, PENEPMA_Root$ & "\Penfluor\material3.mat"

' Create Penfluor input file from Penfluor.inp (0 to 90 takeoff, 5 to 50 keV, default minimum energy 1 keV)
Call Penepma12CreatePenfluorInput(n%)
If ierror Then Exit Sub

' Copy Penfluor.in to <binaryname$>.in (do not need filenames with blanks to be in double quotes for FileCopy or Kill statements)
If Dir$(PENEPMA_Root$ & "\Penfluor\Penfluor.in") <> vbNullString Then
If n% = 1 And Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(MaterialFileA$) & ".in") <> vbNullString Then Kill PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(MaterialFileA$) & ".in"
If n% = 2 And Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(MaterialFileB$) & ".in") <> vbNullString Then Kill PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(MaterialFileB$) & ".in"
If n% = 3 And Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(MaterialFileBStd$) & ".in") <> vbNullString Then Kill PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(MaterialFileBStd$) & ".in"

If n% = 1 Then FileCopy PENEPMA_Root$ & "\Penfluor\Penfluor.in", PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(MaterialFileA$) & ".in"
If n% = 2 Then FileCopy PENEPMA_Root$ & "\Penfluor\Penfluor.in", PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(MaterialFileB$) & ".in"
If n% = 3 Then FileCopy PENEPMA_Root$ & "\Penfluor\Penfluor.in", PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(MaterialFileBStd$) & ".in"
Else
GoTo Penepma12RunPenfluorINFileNotFound:
End If

' Run Penfluor on material
Call Penepma12RunPenfluor2
If ierror Then Exit Sub

' Now wait for the Penfluor calculation to finish
Do Until Not SimulationInProgress
DoEvents
If icancelauto Then
SimulationInProgress = False
FitParametersInProgress = False
Call Penepma12UpdateForm
If ierror Then Exit Sub
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(1), SimulationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If
Loop

' Check for .dat files from Penfluor.exe
DoEvents
tfilename$ = PENEPMA_Root$ & "\Penfluor\gphcoxs.dat"
If Dir$(tfilename$) = vbNullString Then GoTo Penepma12RunPenFluorDATFileNotFound
tfilename$ = PENEPMA_Root$ & "\Penfluor\br.dat"
If Dir$(tfilename$) = vbNullString Then GoTo Penepma12RunPenFluorDATFileNotFound
tfilename$ = PENEPMA_Root$ & "\Penfluor\xr.dat"
If Dir$(tfilename$) = vbNullString Then GoTo Penepma12RunPenFluorDATFileNotFound
tfilename$ = PENEPMA_Root$ & "\Penfluor\penepma.dat"
If Dir$(tfilename$) = vbNullString Then GoTo Penepma12RunPenFluorDATFileNotFound

' Run Fitall on material
Call Penepma12RunFitall
If ierror Then Exit Sub

' Now wait for the Fitall calculation to finish
Do Until Not FitParametersInProgress
DoEvents
If icancelauto Then
SimulationInProgress = False
FitParametersInProgress = False
Call Penepma12UpdateForm
If ierror Then Exit Sub
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(2), FitParametersInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If
Loop

' Check for .par file from Fitall.exe
If n% = 1 Then tfilename$ = PENEPMA_Root$ & "\Penfluor\material1.par"
If n% = 2 Then tfilename$ = PENEPMA_Root$ & "\Penfluor\material2.par"
If n% = 3 Then tfilename$ = PENEPMA_Root$ & "\Penfluor\material3.par"
If Dir$(tfilename$) = vbNullString Then GoTo Penepma12RunPenFluorParameterFileNotFound

' Load new parameter file names
If n% = 1 Then tfilename$ = PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(MaterialFileA$) & ".par"
If n% = 2 Then tfilename$ = PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(MaterialFileB$) & ".par"
If n% = 3 Then tfilename$ = PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(MaterialFileBStd$) & ".par"

' Update new parameter filenames (without path)
If n% = 1 Then ParameterFileA$ = MiscGetFileNameOnly$(tfilename$)
If n% = 2 Then ParameterFileB$ = MiscGetFileNameOnly$(tfilename$)
If n% = 3 Then ParameterFileBStd$ = MiscGetFileNameOnly$(tfilename$)

' Modify .par files back to original filenames
If n% = 1 Then Call Penepma12ModifyParFiles(Int(1), Int(0))
If n% = 2 Then Call Penepma12ModifyParFiles(Int(2), Int(0))
If n% = 3 Then Call Penepma12ModifyParFiles(Int(3), Int(0))
If ierror Then Exit Sub

' Update parameter file fields with parameter name
If n% = 1 Then FormPENEPMA12.TextParameterFileA.Text = ParameterFileA$
If n% = 2 Then FormPENEPMA12.TextParameterFileB.Text = ParameterFileB$
If n% = 3 Then FormPENEPMA12.TextParameterFileBStd.Text = ParameterFileBStd$

If n% = 1 Then msg$ = "Parameter File A " & ParameterFileA$ & " saved to " & PENEPMA_Root$ & "\Penfluor"
If n% = 2 Then msg$ = "Parameter File B " & ParameterFileB$ & " saved to " & PENEPMA_Root$ & "\Penfluor"
If n% = 3 Then msg$ = "Parameter File B Std " & ParameterFileBStd$ & " saved to " & PENEPMA_Root$ & "\Penfluor"
Call IOWriteLog(msg$)

Call Penepma12UpdateForm
If ierror Then Exit Sub

Exit Sub

' Errors
Penepma12RunPenfluorError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12RunPenfluor"
ierror = True
Exit Sub

Penepma12RunPenfluorNoEXE:
msg$ = "Penfluor.exe was not found in the folder " & PENEPMA_Root$ & "\Penfluor"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunPenfluor"
ierror = True
Exit Sub

Penepma12RunPenfluorNoGEO:
msg$ = "Penfluor.geo was not found in the folder " & PENEPMA_Root$ & "\Penfluor"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunPenfluor"
ierror = True
Exit Sub

Penepma12RunPenFluorMaterialFileNotFound:
msg$ = "Material file " & tfilename$ & " was not found in the folder " & PENEPMA_Root$ & "\Penfluor"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunPenfluor"
ierror = True
Exit Sub

Penepma12RunPenfluorINFileNotFound:
msg$ = "Penfluor input file (Penfluor.in) was not found in the folder " & PENEPMA_Root$ & "\Penfluor"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunPenfluor"
ierror = True
Exit Sub

Penepma12RunPenFluorDATFileNotFound:
msg$ = "Data file " & tfilename$ & " was not created properly in the folder " & PENEPMA_Root$ & "\Penfluor\. "
msg$ = msg$ & "Please click the Penfluor Prompt button and type the following command to see the actual error: penfluor.exe"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunPenfluor"
ierror = True
Exit Sub

Penepma12RunPenFluorParameterFileNotFound:
msg$ = "Parameter file " & tfilename$ & " was not created properly in the folder " & PENEPMA_Root$ & "\Penfluor\. "
msg$ = msg$ & "Please click the Penfluor Prompt button and type the following command to see the actual error: fitall.exe."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunPenfluor"
ierror = True
Exit Sub

End Sub

Sub Penepma12RunPenfluor2()
' Run Penfluor batch file

ierror = False
On Error GoTo Penepma12RunPenfluor2Error

Dim bfilename As String, astring As String

' Create batch file to run Penfluor
bfilename$ = PENEPMA_Root$ & "\Penfluor\temp.bat"
Open bfilename$ For Output As #Temp1FileNumber%

astring$ = Left$(PENEPMA_Root$, 2)                                          ' change to drive
Print #Temp1FileNumber%, astring$
astring$ = "cd " & VbDquote$ & PENEPMA_Root$ & "\Penfluor" & VbDquote$      ' change to folder
Print #Temp1FileNumber%, astring$
astring$ = "Penfluor"
Print #Temp1FileNumber%, astring$
Close #Temp1FileNumber%

' Start Penfluor (/k executes but window remains, /c executes but terminates)
'PenepmaTaskID& = Shell("cmd.exe /k " & VbDquote$ & bfilename$ & VbDquote$, vbNormalFocus)
PenepmaTaskID& = Shell("cmd.exe /c " & VbDquote$ & bfilename$ & VbDquote$, vbNormalFocus)

FormPENEPMA12.Timer1.Interval = 4 * MSECPERSEC#     ' Update every 4 seconds
SimulationInProgress = True
FormPENEPMA12.LabelProgress.Caption = "Simulation In Progress!"
MaterialSimulationStart = Now
DoEvents

Call Penepma12UpdateForm
If ierror Then Exit Sub
Exit Sub

' Errors
Penepma12RunPenfluor2Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12RunPenfluor2"
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12RunPenfluorCheck(n As Integer)
' Check with the user before running Penfluor calculations
'  n = 0 = All three materials
'  n = 1 = Material A
'  n = 2 = Material B
'  n = 3 = Material B Std

ierror = False
On Error GoTo Penepma12RunPenfluorCheckError

Dim astring As String
Dim response As Integer

' Check if target parameter files already exist and ask use whether to overwrite
astring$ = vbNullString
If n% = 0 Then
If Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(MaterialFileA$) & ".par") <> vbNullString Then
astring$ = MiscGetFileNameNoExtension$(MaterialFileA$) & ".par"
End If
If Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(MaterialFileB$) & ".par") <> vbNullString Then
astring$ = astring$ & ", " & MiscGetFileNameNoExtension$(MaterialFileB$) & ".par"
End If
If Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(MaterialFileBStd$) & ".par") <> vbNullString Then
astring$ = astring$ & ", " & MiscGetFileNameNoExtension$(MaterialFileBStd$) & ".par"
End If

ElseIf n% = 1 Then
astring$ = vbNullString
If Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(MaterialFileA$) & ".par") <> vbNullString Then
astring$ = MiscGetFileNameNoExtension$(MaterialFileA$) & ".par"
End If

ElseIf n% = 2 Then
If Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(MaterialFileB$) & ".par") <> vbNullString Then
astring$ = MiscGetFileNameNoExtension$(MaterialFileB$) & ".par"
End If

ElseIf n% = 3 Then
If Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(MaterialFileBStd$) & ".par") <> vbNullString Then
astring$ = MiscGetFileNameNoExtension$(MaterialFileBStd$) & ".par"
End If
End If

' If not overwriting files, check if any are present
If CalculateDoNotOverwritePAR Then
If astring$ <> vbNullString Then
msg$ = "Parameter file(s) " & astring$ & " already exist. Are you sure you want to overwrite them?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton2, "Penepma12RunPenfluorCheck")
If response% = vbNo Then
ierror = True
Exit Sub
End If
End If
End If

Exit Sub

' Errors
Penepma12RunPenfluorCheckError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12RunPenfluorCheck"
ierror = True
Exit Sub

End Sub

Sub Penepma12RunPenfluorCheck2(n As Integer)
' Check with the user before running Penfluor calculations (calculate material simulation time)
'  n = 0 = All three materials
'  n = 1 = Material A
'  n = 2 = Material B
'  n = 3 = Material B Std

ierror = False
On Error GoTo Penepma12RunPenfluorCheck2Error

Dim response As Integer
Dim tAdditionalElementSimulationTime As Double
Dim atime As Variant

' Check for number of elements in each .MAT composition (if they exist)
If n% = 1 Or n% = 0 Then
If Dir$(PENDBASE_Path$ & "\" & MaterialFileA$) <> vbNullString Then
Call Penepma12GetMatFileComposition(PENDBASE_Path$ & "\" & MaterialFileA$, PENEPMA_SampleA())
If ierror Then Exit Sub
Else
PENEPMA_SampleA(1).LastChan% = 2    ' assume two elements if not found (probably running full range calculations)
End If
End If

If n% = 2 Or n% = 0 Then
If Dir$(PENDBASE_Path$ & "\" & MaterialFileB$) <> vbNullString Then
Call Penepma12GetMatFileComposition(PENDBASE_Path$ & "\" & MaterialFileB$, PENEPMA_SampleB())
If ierror Then Exit Sub
Else
PENEPMA_SampleB(1).LastChan% = 2    ' assume two elements if not found (probably running full range calculations)
End If
End If

If n% = 3 Or n% = 0 Then
If Dir$(PENDBASE_Path$ & "\" & MaterialFileBStd$) <> vbNullString Then
Call Penepma12GetMatFileComposition(PENDBASE_Path$ & "\" & MaterialFileBStd$, PENEPMA_SampleBStd())
If ierror Then Exit Sub
Else
PENEPMA_SampleBStd(1).LastChan% = 2    ' assume two elements if not found (probably running full range calculations)
End If
End If

' Calculate additional element time (add 1% extra time per element)
If (n% = 1 Or n% = 0) Then AdditionalElementSimulationTime#(1) = MaterialSimulationTime# * PENEPMA_SampleA(1).LastChan% / 100#
If (n% = 2 Or n% = 0) Then AdditionalElementSimulationTime#(2) = MaterialSimulationTime# * PENEPMA_SampleB(1).LastChan% / 100#
If (n% = 3 Or n% = 0) Then AdditionalElementSimulationTime#(3) = MaterialSimulationTime# * PENEPMA_SampleBStd(1).LastChan% / 100#

' Average additional element time if running all three
If n% = 0 Then
tAdditionalElementSimulationTime# = (AdditionalElementSimulationTime#(1) + AdditionalElementSimulationTime#(2) + AdditionalElementSimulationTime#(3)) / 3#
End If

' Confirm calculation time (multiply by NUMSIM& because Penfluor uses NUMSIM& simulations)
atime = MaterialSimulationOverhead# * NUMSIM& / SECPERDAY# + MaterialSimulationTime# * NUMSIM& / SECPERDAY#   ' total simulation time

If n% = 1 Then atime = atime + AdditionalElementSimulationTime#(1) / SECPERDAY#     ' additional element time
If n% = 2 Then atime = atime + AdditionalElementSimulationTime#(2) / SECPERDAY#     ' additional element time
If n% = 3 Then atime = atime + AdditionalElementSimulationTime#(3) / SECPERDAY#     ' additional element time
If n% = 0 Then atime = atime + tAdditionalElementSimulationTime# / SECPERDAY#     ' additional element time

atime = atime + FitParametersTime# / SECPERDAY#     ' total fit all parameters time
atime = atime * SECPERDAY#  ' convert to seconds
atime = atime * TotalNumberOfSimulations&   ' multiple times number of par files
atime = atime / SECPERDAY#  ' convert back to days

msg$ = "The " & Format$(TotalNumberOfSimulations&) & " specified Penfluor simulation calculation(s) will take " & MiscConvertTimeToClockString$(atime) & " to complete. Are you sure you want to start the calculations?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton2, "Penepma12RunPenfluorCheck2")
If response% = vbNo Then
ierror = True
Exit Sub
End If

Exit Sub

' Errors
Penepma12RunPenfluorCheck2Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12RunPenfluorCheck2"
ierror = True
Exit Sub

Penepma12RunPenfluorCheck2NoFileA:
msg$ = "Material file " & MaterialFileA$ & " was not found in the folder " & PENDBASE_Path$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunPenfluorCheck2"
ierror = True
Exit Sub

Penepma12RunPenfluorCheck2NoFileB:
msg$ = "Material file " & MaterialFileB$ & " was not found in the folder " & PENDBASE_Path$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunPenfluorCheck2"
ierror = True
Exit Sub

Penepma12RunPenfluorCheck2NoFileBStd:
msg$ = "Material file " & MaterialFileBStd$ & " was not found in the folder " & PENDBASE_Path$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunPenfluorCheck2"
ierror = True
Exit Sub

End Sub

Sub Penepma12RunFitall()
' Run Fitall batch file

ierror = False
On Error GoTo Penepma12RunFitallError

Dim bfilename As String, astring As String

' Create batch file to run Penepma
bfilename$ = PENEPMA_Root$ & "\Penfluor\temp.bat"
Open bfilename$ For Output As #Temp1FileNumber%

astring$ = Left$(PENEPMA_Root$, 2)                                          ' change to drive
Print #Temp1FileNumber%, astring$
astring$ = "cd " & VbDquote$ & PENEPMA_Root$ & "\Penfluor" & VbDquote$      ' change to folder
Print #Temp1FileNumber%, astring$
astring$ = "Fitall"
Print #Temp1FileNumber%, astring$
Close #Temp1FileNumber%

' Start Fitall (/k executes but window remains, /c executes but terminates)
'PenepmaTaskID& = Shell("cmd.exe /k " & VbDquote$ & bfilename$ & VbDquote$, vbNormalFocus)
PenepmaTaskID& = Shell("cmd.exe /c " & VbDquote$ & bfilename$ & VbDquote$, vbNormalFocus)

FormPENEPMA12.Timer1.Interval = 4 * MSECPERSEC#     ' Update every 4 seconds
FitParametersInProgress = True
FitParametersStart = Now
FormPENEPMA12.LabelProgress.Caption = "Fit All Parameters In Progress!"
DoEvents

Exit Sub

' Errors
Penepma12RunFitallError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12RunFitall"
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12CreatePenfluorInput(n As Integer)
' Create a Penfluor.in file based on the specified material and the simulation time
'  n = 1 = Material A
'  n = 2 = Material B
'  n = 3 = Material B Std
'
'    Sample Penfluor.in file (based on Penfluor.inp):
'
'TITLE  A thick cylindrical target
'       .  Change only the material filename. Keep the rest unaltered.
'       >>>>>>>> Electron beam definition.
'SENERG 100e3                     [Energy of the electron beam, in eV]
'SPOSIT 0 0 1                     [Coordinates of the electron source]
'SDIREC 180 0              [Direction angles of the beam axis, in deg]
'SAPERT 0                                      [Beam aperture, in deg]
'       .
'       >>>>>>>> Material data and simulation parameters.
'                Up to 10 materials; 2 lines for each material.
'MFNAME Fe.mat                         [Material file, up to 20 chars]
'MSIMPA 1e3 1e3 1e3 0.2 0.2 1e3 1e3          [EABS(1:3),C1,C2,WCC,WCR]
'       .
'       >>>>>>>> Geometry of the sample.
'GEOMFN penfluor.geo              [Geometry definition file, 20 chars]
'DSMAX  1 1.0e-4             [IB, Maximum step length (cm) in body IB]
'       .
'       >>>>>>>> Interaction forcing.
'IFORCE 1 1 4 -10    0.1 1.0           [KB,KPAR,ICOL,FORCER,WLOW,WHIG]
'IFORCE 1 1 5 -400   0.1 1.0           [KB,KPAR,ICOL,FORCER,WLOW,WHIG]
'       .
'       >>>>>>>> Emerging particles. Energy and angular distributions.
'NBE    0.0 0.0 100                [E-interval and no. of energy bins]
'NBTH   45                     [No. of bins for the polar angle THETA]
'NBPH   30                   [No. of bins for the azimuthal angle PHI]
'       .
'       >>>>>>>> Photon detectors (up to 25 different detectors).
'                IPSF=0, do not create a phase-space file.
'                IPSF=1, creates a phase-space file.
'PDANGL 0 90  0 360 0                   [Angular window, in deg, IPSF]
'PDENER 0   20e3 1000                 [Energy window, no. of channels]
'       .
'NSIMSH 2.0e9                    [Desired number of simulated showers]
'TIME   3600                        [Allotted simulation time, in sec]

ierror = False
On Error GoTo Penepma12CreatePenfluorInputError

Dim astring As String, bstring As String, cstring As String, dstring As String

Dim estring As String, fstring As String
Dim EABS(1 To 3) As Double, c1 As Double, c2 As Double, WCC As Double, WCR As Double

' Loop through sample input file and copy to new file with modified parameters
Open PenfluorInputFile$ For Input As #Temp1FileNumber%
Open PenfluorOutputFile$ For Output As #Temp2FileNumber%

Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$
bstring$ = astring$
cstring$ = vbNullString

If InStr(astring$, "TITLE") > 0 Then Call Penepma12CreateInput2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

If InStr(astring$, "SENERG") > 0 Then Call Penepma12CreateInput2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

If InStr(astring$, "SPOSIT") > 0 Then Call Penepma12CreateInput2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

If InStr(astring$, "SDIREC") > 0 Then Call Penepma12CreateInput2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

If InStr(astring$, "SAPERT") > 0 Then Call Penepma12CreateInput2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

' Load each material file and simulation parameters
If InStr(astring$, "MFNAME") > 0 Then
If n% = 1 Then cstring$ = "material1.mat"
If n% = 2 Then cstring$ = "material2.mat"
If n% = 3 Then cstring$ = "material3.mat"
Call Penepma12CreateInput2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub
End If

' Parse out MSIMPA parameters: MSIMPA 1e3 1e3 1e3 0.2 0.2 1e3 1e3          [EABS(1:3),C1,C2,WCC,WCR]
If InStr(astring$, "MSIMPA") > 0 Then
estring$ = astring$
Call MiscParseStringToString(estring$, fstring$)    ' read "MSIMPA" keyword (not used)
If ierror Then Exit Sub

Call MiscParseStringToString(estring$, fstring$)
If ierror Then Exit Sub
EABS#(1) = Val(Trim$(fstring$))     ' not used
EABS#(1) = PenepmaMinimumElectronEnergy! * EVPERKEV#    ' overwrite with value from FormPenepma12
Call MiscParseStringToString(estring$, fstring$)
If ierror Then Exit Sub
EABS#(2) = Val(Trim$(fstring$))     ' not used
EABS#(2) = PenepmaMinimumElectronEnergy! * EVPERKEV#    ' overwrite with value from FormPenepma12
Call MiscParseStringToString(estring$, fstring$)
If ierror Then Exit Sub
EABS#(3) = Val(Trim$(fstring$))     ' not used
EABS#(3) = PenepmaMinimumElectronEnergy! * EVPERKEV#   ' overwrite with value from FormPenepma12

Call MiscParseStringToString(estring$, fstring$)
If ierror Then Exit Sub
c1# = Val(Trim$(fstring$))          ' save for re-use
Call MiscParseStringToString(estring$, fstring$)
If ierror Then Exit Sub
c2# = Val(Trim$(fstring$))          ' save for re-use

Call MiscParseStringToString(estring$, fstring$)
If ierror Then Exit Sub
WCC# = Val(Trim$(fstring$))          ' not used
WCC# = PenepmaMinimumElectronEnergy! * EVPERKEV#    ' overwrite with value from FormPenepma12
Call MiscParseStringToString(estring$, fstring$)
If ierror Then Exit Sub
WCR# = Val(Trim$(fstring$))          ' not used
WCR# = PenepmaMinimumElectronEnergy! * EVPERKEV#    ' overwrite with value from FormPenepma12

cstring$ = Format$(EABS#(1), "0.0E+0") & " " & Format$(EABS#(2), "0.0E+0") & " " & Format$(EABS#(3), "0E+0") & " "
cstring$ = cstring$ & Format$(c1#, "0.0") & " " & Format$(c2#, "0.0") & " "
cstring$ = cstring$ & Format$(WCC#, "0E+0") & " " & Format$(WCR#, "0E+0")

Call Penepma12CreateInput2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub
End If

' Load geometry file name (no changes)
If InStr(astring$, "GEOMFN") > 0 Then Call Penepma12CreateInput2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

' Load maximum step length (no changes)
If InStr(astring$, "DSMAX") > 0 Then Call Penepma12CreateInput2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

' Load forcing parameters (can have multiple occurances) (no changes)
If InStr(astring$, "IFORCE") > 0 Then Call Penepma12CreateInput2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

If InStr(astring$, "NBE") > 0 Then Call Penepma12CreateInput2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

If InStr(astring$, "NBTH") > 0 Then Call Penepma12CreateInput2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

If InStr(astring$, "NBPH") > 0 Then Call Penepma12CreateInput2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

If InStr(astring$, "PDANGL") > 0 Then Call Penepma12CreateInput2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

If InStr(astring$, "PDENER") > 0 Then Call Penepma12CreateInput2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

cstring$ = Format$(MaterialSimulationShowers#, e71$)
If InStr(astring$, "NSIMSH") > 0 Then Call Penepma12CreateInput2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

cstring$ = Format$(MaterialSimulationTime#)
If InStr(astring$, "TIME") > 0 Then Call Penepma12CreateInput2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

Print #Temp2FileNumber%, bstring$
Loop

Close #Temp1FileNumber%
Close #Temp2FileNumber%

' Confirm with user
If n% = 1 Then msg$ = "Penfluor input file " & PenfluorOutputFile$ & " for " & MiscGetFileNameOnly$(MaterialFileA$) & " was saved to " & PENEPMA_Root$ & "\Penfluor"
If n% = 2 Then msg$ = "Penfluor input file " & PenfluorOutputFile$ & " for " & MiscGetFileNameOnly$(MaterialFileB$) & " was saved to " & PENEPMA_Root$ & "\Penfluor"
If n% = 3 Then msg$ = "Penfluor input file " & PenfluorOutputFile$ & " for " & MiscGetFileNameOnly$(MaterialFileBStd$) & " was saved to " & PENEPMA_Root$ & "\Penfluor"
Call IOWriteLog(msg$)

Exit Sub

' Errors
Penepma12CreatePenfluorInputError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CreatePenfluorInput"
Close #Temp1FileNumber%
Close #Temp2FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma12CreateInput2(astring As String, bstring As String, cstring As String, dstring As String)
' Make the output string based on new value and current string

ierror = False
On Error GoTo Penepma12CreateInput2Error

' If cstring is blank then no editing required
If cstring$ = vbNullString Then
bstring$ = astring$
Exit Sub
End If

' Look for bracketed information to append
If InStr(astring$, "[") > 0 Then
dstring$ = Mid$(astring$, InStr(astring$, "["))
bstring$ = Left$(astring, COL7%) & cstring$ & Space$(Len(astring$) - COL7% - Len(cstring$) - Len(dstring$)) & dstring$

Else
bstring$ = astring$                 ' comment string, leave as is
End If

Exit Sub

' Errors
Penepma12CreateInput2Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CreateInput2"
Close #Temp1FileNumber%
Close #Temp2FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma12CheckTermination()
' Check if Material, Penfluor, Fitall or Fanal process is terminated

ierror = False
On Error GoTo Penepma12CheckTerminationError

Dim atime As Variant

' Check if running material calculation
If MaterialInProgress Then
FormPENEPMA12.CommandOK.Enabled = False
FormPENEPMA12.CommandClose.Enabled = False

If FormPenepma12Binary.Visible Then
FormPenepma12Binary.CommandOK.Enabled = False
FormPenepma12Binary.CommandClose.Enabled = False
End If

FormPENEPMA12.LabelRemainingTime.Caption = "..."

If IOIsProcessTerminated(PenepmaTaskID&) Then
Call Penepma12CheckTermination2(Int(0), MaterialInProgress)
If ierror Then Exit Sub
End If

DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(0), MaterialInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If
End If

' Check if running Penfluor simulation
If SimulationInProgress Then
FormPENEPMA12.CommandOK.Enabled = False
FormPENEPMA12.CommandClose.Enabled = False

If FormPenepma12Binary.Visible Then
FormPenepma12Binary.CommandOK.Enabled = False
FormPenepma12Binary.CommandClose.Enabled = False
End If

' Simulation still running, display remaining time (NUMSIM& is number of beam energy simulations per Penfluor run)
atime = MaterialSimulationOverhead# * NUMSIM& / SECPERDAY# + MaterialSimulationTime# * NUMSIM& / SECPERDAY#   ' total simulation time
atime = atime + AdditionalElementSimulationTime#(CurrentSimulationModeNumber%) / SECPERDAY#     ' additional element time
atime = (MaterialSimulationStart + atime) - Now
FormPENEPMA12.LabelRemainingTime.Caption = "Simulation: " & Format$(CurrentSimulationsNumber&) & " (of " & Format$(TotalNumberOfSimulations&) & ") Remaining Time: " & MiscConvertTimeToClockString$(atime)

If IOIsProcessTerminated(PenepmaTaskID&) Then
Call Penepma12CheckTermination2(Int(1), SimulationInProgress)
If ierror Then Exit Sub
End If

DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(1), SimulationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If
End If

' Check if running Fitall
If FitParametersInProgress Then
FormPENEPMA12.CommandOK.Enabled = False
FormPENEPMA12.CommandClose.Enabled = False

If FormPenepma12Binary.Visible Then
FormPenepma12Binary.CommandOK.Enabled = False
FormPenepma12Binary.CommandClose.Enabled = False
End If

' Fit parameters still running, display message
atime = FitParametersTime# / SECPERDAY#
atime = (FitParametersStart + atime) - Now
FormPENEPMA12.LabelRemainingTime.Caption = "Fit All Parameters: " & Format$(CurrentSimulationsNumber&) & " (of " & Format$(TotalNumberOfSimulations&) & ") Remaining Time: " & MiscConvertTimeToClockString$(atime)

If IOIsProcessTerminated(PenepmaTaskID&) Then
Call Penepma12CheckTermination2(Int(2), FitParametersInProgress)
If ierror Then Exit Sub
End If

DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(2), FitParametersInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If
End If

' Check if running Fanal
If CalculationInProgress Then
FormPENEPMA12.CommandOK.Enabled = False
FormPENEPMA12.CommandClose.Enabled = False

If FormPenepma12Binary.Visible Then
FormPenepma12Binary.CommandOK.Enabled = False
FormPenepma12Binary.CommandClose.Enabled = False
End If

' Fit parameters still running, display message
atime = Now - CalculationStart
FormPENEPMA12.LabelRemainingTime.Caption = "Elapsed Time: " & MiscConvertTimeToClockString$(atime)

If IOIsProcessTerminated(PenepmaTaskID&) Then
Call Penepma12CheckTermination2(Int(3), CalculationInProgress)
If ierror Then Exit Sub
End If

DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(3), CalculationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If
End If

' Update form
Call Penepma12UpdateForm
If ierror Then Exit Sub

Exit Sub

' Errors
Penepma12CheckTerminationError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CheckTermination"
ierror = True
Exit Sub

End Sub

Sub Penepma12CheckTermination2(mode As Integer, tBoolean As Boolean)
' Clean up after completion or user cancel

ierror = False
On Error GoTo Penepma12CheckTermination2Error

tBoolean = False
If mode% = 0 Then FormPENEPMA12.LabelProgress.Caption = "Material Completed!"
If mode% = 1 Then FormPENEPMA12.LabelProgress.Caption = "Simulation Completed!"
If mode% = 2 Then FormPENEPMA12.LabelProgress.Caption = "Fit All Parameters Completed!"
If mode% = 3 Then FormPENEPMA12.LabelProgress.Caption = "K-ratio Calculations Completed!"

FormPENEPMA12.LabelRemainingTime.Caption = vbNullString
FormPENEPMA12.Timer1.Interval = 0
DoEvents

' Set enables
FormPENEPMA12.CommandOK.Enabled = True
FormPENEPMA12.CommandClose.Enabled = True

If FormPenepma12Binary.Visible Then
FormPenepma12Binary.CommandOK.Enabled = True
FormPenepma12Binary.CommandClose.Enabled = True
End If

Exit Sub

' Errors
Penepma12CheckTermination2Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CheckTermination2"
ierror = True
Exit Sub

End Sub

Sub Penepma12RunFanal()
' Run Fanal for all three previously specified parameter files

ierror = False
On Error GoTo Penepma12RunFanalError

Dim ip As Integer
Dim tfilename As String, pstring As String, astring As String
Dim eng As Single, edg As Single, temp As Single

Dim t1 As Single, t2 As Single

Static userwarned As Boolean

' Check that the overvoltage is ok for the selected line
PENEPMA_Sample(1).LastElm% = 1
PENEPMA_Sample(1).LastChan% = 1
PENEPMA_Sample(1).DisableQuantFlag%(1) = 0
PENEPMA_Sample(1).takeoff! = MaterialMeasuredTakeoff#
PENEPMA_Sample(1).TakeoffArray!(1) = MaterialMeasuredTakeoff#
PENEPMA_Sample(1).kilovolts! = MaterialMeasuredEnergy#
PENEPMA_Sample(1).KilovoltsArray!(1) = MaterialMeasuredEnergy#
PENEPMA_Sample(1).Elsyms$(1) = Symlo$(MaterialMeasuredElement%)
PENEPMA_Sample(1).Xrsyms$(1) = Xraylo$(MaterialMeasuredXray%)

' Fill element arrays
astring = "loading element arrays"
Call ElementLoadArrays(PENEPMA_Sample())
If ierror Then Exit Sub

' Get x-ray data
astring = "getting x-ray data"
Call XrayGetEnergy(MaterialMeasuredElement%, MaterialMeasuredXray%, eng!, edg!)
If ierror Then Exit Sub

' Check for valid x-ray line (excitation energy must be less than beam energy) (and greater than PenepmaMinimumElectronEnergy!)
If eng! = 0# Then GoTo Penepma12RunFanalNoXrayData
If edg! = 0# Then GoTo Penepma12RunFanalNoEdgeData
If edg! > MaterialMeasuredEnergy# Then GoTo Penepma12RunFanalBelowEdge

' Check for .IN file in all three files and if found check MSIMPA parameters (minimum electron/photon energy)
astring = "checking parameter A .in file"
If Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileA$) & ".in") <> vbNullString Then
Call Penepma12RunFanalCheckINFile("MSIMPA", PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileA$) & ".in", pstring$)
If ierror Then Exit Sub
temp! = Val(pstring$)
temp! = temp! / EVPERKEV#
If edg! < temp! Then GoTo Penepma12RunFanalBadMinimumEnergyA
End If

astring = "checking parameter B .in file"
If Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileB$) & ".in") <> vbNullString Then
Call Penepma12RunFanalCheckINFile("MSIMPA", PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileB$) & ".in", pstring$)
If ierror Then Exit Sub
temp! = Val(pstring$)
temp! = temp! / EVPERKEV#
If edg! < temp! Then GoTo Penepma12RunFanalBadMinimumEnergyB
End If

astring = "checking parameter BStd .in file"
If Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileBStd$) & ".in") <> vbNullString Then
Call Penepma12RunFanalCheckINFile("MSIMPA", PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileBStd$) & ".in", pstring$)
If ierror Then Exit Sub
temp! = Val(pstring$)
temp! = temp! / EVPERKEV#
If edg! < temp! Then GoTo Penepma12RunFanalBadMinimumEnergyBStd
End If

' Check if edge is less than Penfluor default modeling energy
astring = "checking edge energy and Penfluor minimum energy"
If Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileA$) & ".in") = vbNullString Then
If Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileB$) & ".in") = vbNullString Then
If Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileBStd$) & ".in") = vbNullString Then
If Not userwarned And edg! < 1# Then
msg$ = "The measured element and x-ray (" & Symlo$(MaterialMeasuredElement%) & " " & Xraylo$(MaterialMeasuredXray%) & ") edge energy (" & Format$(edg!) & " keV) is less "
msg$ = msg$ & "than the default Penfluor modeling energy of 1 keV." & vbCrLf & vbCrLf
msg$ = msg$ & "Please make sure that you are utilizing PAR files that have been properly calculated to the edge energy of the x-ray line you wish to produce." & vbCrLf & vbCrLf
msg$ = msg$ & "Some element and compound .PAR files have been run down to 500 eV (and even 200 or 100 eV) and are available from Probe Software or you may re-calculate them yourself using the Run Penfluor buttons after "
msg$ = msg$ & "adjusting the Minimum Electron Energy parameter (in the Penepma 2012 | Primary Intensity Calculations area)." & vbCrLf & vbCrLf & "This warning will be given just this once."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunFanal"
userwarned = True
End If
End If
End If
End If

' Double check that specific transition exists (see table 6.2 in Penelope-2006-NEA-pdf)
Call PenepmaGetPDATCONFTransition(MaterialMeasuredElement%, MaterialMeasuredXray%, t1!, t2!)
If ierror Then Exit Sub

' If both shells have ionization energies, it is ok to calculate
If t1! = 0# Or t2! = 0# Then GoTo Penepma12RunFanalBadTransitions

' Check overvoltage (again)
astring = "checking overvoltage"
Call ElementCheckXray(Int(1), PENEPMA_Sample())
If ierror Then Exit Sub

' Check for each material .par file in penfluor folder (copy original .par file even though the temp is copied)
astring = "checking parameter A file"
tfilename$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileA$
If Dir$(tfilename$) = vbNullString Then GoTo Penepma12RunFanalParameterFileNotFound
FileCopy tfilename$, PENEPMA_Root$ & "\Fanal\db\" & ParameterFileA$
DoEvents
If Dir$(MiscGetFileNameNoExtension$(tfilename$) & ".in") <> vbNullString Then FileCopy MiscGetFileNameNoExtension$(tfilename$) & ".in", PENEPMA_Root$ & "\Fanal\db\" & MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(ParameterFileA$)) & ".in"

astring = "checking parameter B file"
tfilename$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileB$
If Dir$(tfilename$) = vbNullString Then GoTo Penepma12RunFanalParameterFileNotFound
FileCopy tfilename$, PENEPMA_Root$ & "\Fanal\db\" & ParameterFileB$
DoEvents
If Dir$(MiscGetFileNameNoExtension$(tfilename$) & ".in") <> vbNullString Then FileCopy MiscGetFileNameNoExtension$(tfilename$) & ".in", PENEPMA_Root$ & "\Fanal\db\" & MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(ParameterFileB$)) & ".in"

astring = "checking parameter BStd file"
tfilename$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileBStd$
If Dir$(tfilename$) = vbNullString Then GoTo Penepma12RunFanalParameterFileNotFound
FileCopy tfilename$, PENEPMA_Root$ & "\Fanal\db\" & ParameterFileBStd$
DoEvents
If Dir$(MiscGetFileNameNoExtension$(tfilename$) & ".in") <> vbNullString Then FileCopy MiscGetFileNameNoExtension$(tfilename$) & ".in", PENEPMA_Root$ & "\Fanal\db\" & MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(ParameterFileBStd$)) & ".in"

' Modify .par files and copy as temp files to fanal\db folder
astring = "modifying parameter A file"
Call Penepma12ModifyParFiles(Int(1), Int(1))
If ierror Then Exit Sub
DoEvents
astring = "modifying parameter B file"
Call Penepma12ModifyParFiles(Int(2), Int(1))
If ierror Then Exit Sub
DoEvents
astring = "modifying parameter BStd file"
Call Penepma12ModifyParFiles(Int(3), Int(1))
If ierror Then Exit Sub
DoEvents

' Check that measured element is present in material B (not absolutely necessary)
'tfilename$ = PENEPMA_Root$ & "\Fanal\db\" & "material2.par"        ' material B
'Call Penepma12GetParFileComposition(Int(2), tfilename$, PENEPMA_Sample())
'If ierror Then Exit Sub
'ip% = IPOS1%(PENEPMA_Sample(1).LastElm%, Symlo$(MaterialMeasuredElement%), PENEPMA_Sample(1).Elsyms$())
'If ip% = 0 Then GoTo Penepma12RunFanalNotFoundMaterialB

' Check that measured element is present in material BStd
astring = "checking measured element in BStd file"
tfilename$ = PENEPMA_Root$ & "\Fanal\db\" & "material3.par"        ' material B Std
Call Penepma12GetParFileComposition(Int(3), tfilename$, PENEPMA_Sample())
If ierror Then Exit Sub
ip% = IPOS1%(PENEPMA_Sample(1).LastElm%, Symlo$(MaterialMeasuredElement%), PENEPMA_Sample(1).Elsyms$())
If ip% = 0 Then GoTo Penepma12RunFanalNotFoundMaterialBStd

Exit Sub

' Errors
Penepma12RunFanalError:
MsgBox Error$ & ", during, " & astring$ & " (the Fanal\db folder may need to be cleared of .PAR and .IN files on some systems)", vbOKOnly + vbCritical, "Penepma12RunFanal"
ierror = True
Exit Sub

Penepma12RunFanalParameterFileNotFound:
msg$ = "Parameter file " & tfilename$ & " was not found in the folder " & PENEPMA_Root$ & "\Penfluor\"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunFanal"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12RunFanalNotFoundMaterialB:
msg$ = "The measured element " & Symlo$(MaterialMeasuredElement%) & " was not found in parameter file for material B " & ParameterFileB$ & " in the folder " & PENEPMA_Root$ & "\Fanal\db"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunFanal"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12RunFanalNotFoundMaterialBStd:
msg$ = "The measured element " & Symlo$(MaterialMeasuredElement%) & " was not found in parameter file for material B Std " & ParameterFileBStd$ & " in the folder " & PENEPMA_Root$ & "\Fanal\db"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunFanal"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12RunFanalNoXrayData:
msg$ = "No x-ray emission data was found for the measured element " & Symlo$(MaterialMeasuredElement%) & " " & Xraylo$(MaterialMeasuredXray%) & " in x-ray databases. Please check that the specified element and x-ray line are valid (and/or update the " & XLineFile$ & " database)."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunFanal"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12RunFanalNoEdgeData:
msg$ = "No x-ray edge data was found for the measured element " & Symlo$(MaterialMeasuredElement%) & " " & Xraylo$(MaterialMeasuredXray%) & " in x-ray databases. Please check that the specified element and x-ray line are valid (and/or update the " & XEdgeFile$ & " database)."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunFanal"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12RunFanalBelowEdge:
msg$ = "The specified beam energy (" & Format$(MaterialMeasuredEnergy#) & " keV) is below the x-ray edge energy for the measured element " & Symlo$(MaterialMeasuredElement%) & " " & Xraylo$(MaterialMeasuredXray%) & ". Please increase the specified beam energy or try another x-ray line."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunFanal"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12RunFanalBadMinimumEnergyA:
msg$ = "The edge energy for the specified element and x-ray (" & Symlo$(MaterialMeasuredElement%) & " " & Xraylo$(MaterialMeasuredXray%) & ") is less than the minimum electron/photon energy that the MaterialA.PAR file was calculated using." & vbCrLf & vbCrLf & "Please choose another x-ray emission line for this element or re-calculate the .PAR file with a lower minimum energy."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunFanal"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12RunFanalBadMinimumEnergyB:
msg$ = "The edge energy for the specified element and x-ray (" & Symlo$(MaterialMeasuredElement%) & " " & Xraylo$(MaterialMeasuredXray%) & ") is less than the minimum electron/photon energy that the MaterialB.PAR file was calculated using." & vbCrLf & vbCrLf & "Please choose another x-ray emission line for this element or re-calculate the .PAR file with a lower minimum energy."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunFanal"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12RunFanalBadMinimumEnergyBStd:
msg$ = "The edge energy for the specified element and x-ray (" & Symlo$(MaterialMeasuredElement%) & " " & Xraylo$(MaterialMeasuredXray%) & ") is less than the minimum electron/photon energy that the MaterialBStd.PAR file was calculated using." & vbCrLf & vbCrLf & "Please choose another x-ray emission line for this element or re-calculate the .PAR file with a lower minimum energy."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunFanal"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12RunFanalBadTransitions:
msg$ = "One or more shell transitions for the specified element and x-ray (" & Symlo$(MaterialMeasuredElement%) & " " & Xraylo$(MaterialMeasuredXray%) & ") are not in the Penepma pdatconf.pen file. Please choose another x-ray emission line for this element."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunFanal"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12RunFanalCreateInput()
' Create a Fanal input file based on the specified parameter files
'
'    Sample Fanal input file:
'
' Cu.mat                  Material A
' Co.mat                  Material B
' Co.mat                  Standard material
' 50                      detector angle, theta_d(deg) (90 - takeoff)
' 20e3                    electron beam energy (eV)
' 27 1 4                  characteristic line (IZ S1 S2)
' 50.0e-4  50             maximum distance, no. of grid points

ierror = False
On Error GoTo Penepma12RunFanalCreateInputError

Dim astring As String, tfilename As String

' Write each line
Open FanalInputFile$ For Output As #Temp2FileNumber%

' Load each parameter file (change extension to .mat- bless Cesc's physicist's heart)
tfilename$ = "material1.mat"      ' Fanal is case sensitive!
astring$ = Format$(tfilename$, a24$) & "Material A"
Print #Temp2FileNumber%, astring$

tfilename$ = "material2.mat"      ' Fanal is case sensitive!
astring$ = Format$(tfilename$, a24$) & "Material B"
Print #Temp2FileNumber%, astring$

tfilename$ = "material3.mat"      ' Fanal is case sensitive!
astring$ = Format$(tfilename$, a24$) & "Standard material"
Print #Temp2FileNumber%, astring$

astring$ = Format$(Format$(90# - MaterialMeasuredTakeoff#, f52$), a24$) & "detector angle, theta_d (deg)"    ' fixed 09/09/2014
Print #Temp2FileNumber%, astring$

astring$ = Format$(Format$(MaterialMeasuredEnergy# * EVPERKEV#, e82$), a24$) & "electron beam energy (eV)"
Print #Temp2FileNumber%, astring$

' Load element
astring$ = Format$(MaterialMeasuredElement%, i30$)

' Load x-ray transitions
If MaterialMeasuredXray% = 1 Then astring$ = astring$ & " 1 4"      ' K L3  (Ka) (see table 6.2 in Penelope-2006-NEA-pdf)
If MaterialMeasuredXray% = 2 Then astring$ = astring$ & " 1 7"      ' K M3  (Kb)
If MaterialMeasuredXray% = 3 Then astring$ = astring$ & " 4 9"      ' L3 M5 (La)
If MaterialMeasuredXray% = 4 Then astring$ = astring$ & " 3 8"      ' L2 M4 (Lb)
If MaterialMeasuredXray% = 5 Then astring$ = astring$ & " 9 16"     ' M5 N7 (Ma)
If MaterialMeasuredXray% = 6 Then astring$ = astring$ & " 8 15"     ' M4 N6 (Mb)

If MaterialMeasuredXray% = 7 Then astring$ = astring$ & " 3 5"      ' K L3  (Ln) (see table 6.2 in Penelope-2006-NEA-pdf)
If MaterialMeasuredXray% = 8 Then astring$ = astring$ & " 3 13"     ' K M3  (Lg)
If MaterialMeasuredXray% = 9 Then astring$ = astring$ & " 3 15"     ' L3 M5 (Lv)
If MaterialMeasuredXray% = 10 Then astring$ = astring$ & " 4 5"     ' L2 M4 (Ll)
If MaterialMeasuredXray% = 11 Then astring$ = astring$ & " 7 14"    ' M5 N7 (Mg)
If MaterialMeasuredXray% = 12 Then astring$ = astring$ & " 9 12"    ' M4 N6 (Mz)

astring$ = Format$(astring$, a24$)
astring$ = astring$ & "characteristic line (IZ S1 S2)"
Print #Temp2FileNumber%, astring$

' Note: if MaterialMeasuredDistance# = 0 then the modified Fanal will output exponential distances starting at 10 nm
astring$ = Format$(MaterialMeasuredDistance# * CMPERMICRON#, e82$)   ' convert from microns to cm
astring$ = astring$ & " " & Format$(MaterialMeasuredGridPoints%, i50$)
astring$ = Format$(astring$, a24$)
astring$ = astring$ & "maximum distance, no. of grid points"
Print #Temp2FileNumber%, astring$

Close #Temp2FileNumber%

' Confirm with user
msg$ = MiscGetFileNameOnly$(ParameterFileA$)
msg$ = msg$ & ", " & MiscGetFileNameOnly$(ParameterFileB$)
msg$ = msg$ & ", " & MiscGetFileNameOnly$(ParameterFileBStd$)
msg$ = "Fanal input file " & FanalInputFile$ & " for " & msg$ & " saved to " & PENEPMA_Root$ & "\Fanal" & " folder..."
Call IOWriteLog(msg$)

Exit Sub

' Errors
Penepma12RunFanalCreateInputError:
Close #Temp2FileNumber%
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12RunFanalCreateInput"
ierror = True
Exit Sub

End Sub

Sub Penepma12RunFanal2()
' Run Fanal batch file

ierror = False
On Error GoTo Penepma12RunFanal2Error

Dim bfilename As String, astring As String

' Create batch file to run Fanal
bfilename$ = PENEPMA_Root$ & "\Fanal\temp.bat"
Open bfilename$ For Output As #Temp1FileNumber%

astring$ = Left$(PENEPMA_Root$, 2)                                      ' change to drive
Print #Temp1FileNumber%, astring$
astring$ = "cd " & VbDquote$ & PENEPMA_Root$ & "\Fanal" & VbDquote$     ' change to folder
Print #Temp1FileNumber%, astring$
astring$ = "Fanal < " & VbDquote$ & FanalInputFile$ & VbDquote$
Print #Temp1FileNumber%, astring$
Close #Temp1FileNumber%

' Start Fanal (/k executes but window remains, /c executes but terminates)
'PenepmaTaskID& = Shell("cmd.exe /k " & VbDquote$ & bfilename$ & VbDquote$, vbMinimizedNoFocus)
PenepmaTaskID& = Shell("cmd.exe /c " & VbDquote$ & bfilename$ & VbDquote$, vbMinimizedNoFocus)

FormPENEPMA12.Timer1.Interval = 4 * MSECPERSEC#     ' Update every 4 seconds
CalculationInProgress = True
FormPENEPMA12.LabelProgress.Caption = "Calculation In Progress!"
CalculationStart = Now
DoEvents

Call Penepma12UpdateForm
If ierror Then Exit Sub
Exit Sub

' Errors
Penepma12RunFanal2Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12RunFanal2"
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12RunFanalOutput(tForm As Form)
' Create Fanal k-ratio output and plot k-ratio data

ierror = False
On Error GoTo Penepma12RunFanalOutputError

Dim tfilename As String, tfolder As String
Dim filenamearray(1 To 1) As String

' Run the bounary couple in Fanal
Call Penepma12RunFanal1
If ierror Then Exit Sub

' Make a new sub folder in Fanal\couple based on par files and takeoff, energy and measured element and xray
tfolder$ = PENEPMA_Root$ & "\Fanal\couple\"
If Dir$(tfolder$, vbDirectory) = vbNullString Then MkDir tfolder$
tfolder$ = tfolder$ & Format$(MaterialMeasuredTakeoff#)
tfolder$ = tfolder$ & "_" & Format$(MaterialMeasuredEnergy#)
tfolder$ = tfolder$ & "_" & MiscGetFileNameNoExtension$(ParameterFileA$)
tfolder$ = tfolder$ & "_" & MiscGetFileNameNoExtension$(ParameterFileB$)
tfolder$ = tfolder$ & "_" & MiscGetFileNameNoExtension$(ParameterFileBStd$)
tfolder$ = tfolder$ & "_" & Format$(MaterialMeasuredElement%)
tfolder$ = tfolder$ & "_" & Format$(MaterialMeasuredXray%)
If Dir$(tfolder$, vbDirectory) = vbNullString Then MkDir tfolder$

' Write the Fanal data to couple\tfolder$
FileCopy FANAL_IN_File$, tfolder$ & "\" & MiscGetFileNameOnly$(FANAL_IN_File$)
FileCopy VACS_DAT_File$, tfolder$ & "\" & MiscGetFileNameOnly$(VACS_DAT_File$)
FileCopy RANGES_DAT_File$, tfolder$ & "\" & MiscGetFileNameOnly$(RANGES_DAT_File$)
FileCopy MIXED_DAT_File$, tfolder$ & "\" & MiscGetFileNameOnly$(MIXED_DAT_File$)
FileCopy KRATIOS_DAT_File$, tfolder$ & "\" & MiscGetFileNameOnly$(KRATIOS_DAT_File$)
FileCopy FLUORMAT1_PAR_File$, tfolder$ & "\" & MiscGetFileNameOnly$(FLUORMAT1_PAR_File$)
FileCopy FLUORMAT2_PAR_File$, tfolder$ & "\" & MiscGetFileNameOnly$(FLUORMAT2_PAR_File$)
FileCopy FLUORMAT3_PAR_File$, tfolder$ & "\" & MiscGetFileNameOnly$(FLUORMAT3_PAR_File$)
FileCopy ATCOEFFS_DAT_File$, tfolder$ & "\" & MiscGetFileNameOnly$(ATCOEFFS_DAT_File$)

' Output a Fanal.txt file containing the actual PAR files and conditions (as of 10/05/2014)
tfilename$ = tfolder$ & "\fanal.txt"

Open tfilename$ For Output As #Temp1FileNumber%

Write #Temp1FileNumber%, ParameterFileA$
Write #Temp1FileNumber%, ParameterFileB$
Write #Temp1FileNumber%, ParameterFileBStd$
Write #Temp1FileNumber%, MaterialMeasuredTakeoff#
Write #Temp1FileNumber%, MaterialMeasuredEnergy#
Write #Temp1FileNumber%, MaterialMeasuredElement%
Write #Temp1FileNumber%, MaterialMeasuredXray%

Close #Temp1FileNumber%

' Confirm with user
msg$ = "Fanal output files were saved to " & tfolder$
Call IOWriteLog(msg$)

' Get k-ratio data from k-ratio file
Call Penepma12LoadPlotData
If ierror Then Exit Sub

' Ouput modified values to file
Call Penepma12OutputKratios(tfolder$)
If ierror Then Exit Sub

' Plot the k-ratio and modified results
Call Penepma12PlotKRatios_PE(nPoints&, nsets&, MaterialMeasuredEnergy#, MaterialMeasuredElement%, MaterialMeasuredXray%, _
yktotal#(), yctotal#(), yc_prix#(), ycb_only#(), yctotal_meas#(), xdist#())
If ierror Then Exit Sub

' Check if user wants to send modified data to Excel
If FormPENEPMA12.CheckSendToExcel.Value = vbChecked Then
filenamearray$(1) = KRATIOS_DAT_File2$
Call ExcelSendFileListToExcel(Int(1), filenamearray$(), tForm)
If ierror Then Exit Sub
End If

Exit Sub

' Errors
Penepma12RunFanalOutputError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12RunFanalOutput"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma12CalculateMatrix(analysis As TypeAnalysis, sample() As TypeSample, tmpsample() As TypeSample)
' Calculate the matrix corrections for the passed composition using CalcZAF methods

ierror = False
On Error GoTo Penepma12CalculateMatrixError

Dim i As Integer, ip As Integer
Dim eng As Single, edg As Single
Dim ielm As Integer, iray As Integer

' Init ZAF arrays
Call ZAFInitZAF
If ierror Then Exit Sub

' Initialize arrays
Call InitStandards(analysis)
If ierror Then Exit Sub

' Initialize the analysis
Call InitLine(analysis)
If ierror Then Exit Sub

' Make sure that new condition arrays are loaded
For i% = 1 To sample(1).LastElm%
sample(1).TakeoffArray!(i%) = sample(1).takeoff!
sample(1).KilovoltsArray!(i%) = sample(1).kilovolts!
sample(1).BeamCurrentArray!(i%) = DefaultBeamCurrent!
sample(1).BeamSizeArray!(i%) = DefaultBeamSize!

ip% = IPOS1(MAXELM%, sample(1).Elsyms$(i%), Symlo$())
ielm% = ip%
ip% = IPOS1(MAXRAY% - 1, sample(1).Xrsyms$(i%), Xraylo$())
iray% = ip%

' Get x-ray data
Call XrayGetEnergy(ielm%, iray%, eng!, edg!)
If ierror Then Exit Sub

' Check overvoltage and set elements as absorber only if necessary (also disable H and He)
If edg! >= sample(1).kilovolts! Or UCase$(sample(1).Elsyms$(i%)) = UCase$(Symlo$(ATOMIC_NUM_HYDROGEN%)) Or UCase$(sample(1).Elsyms$(i%)) = UCase$(Symlo$(ATOMIC_NUM_HELIUM%)) Then
ip% = IPOS1(sample(1).LastElm%, sample(1).Elsyms$(i%), sample(1).Elsyms$())
If ip% > 0 Then
sample(1).Xrsyms$(i%) = Xraylo$(MAXRAY%)
analysis.WtPercents!(i%) = sample(1).ElmPercents!(ip%)   ' move specified concentration to analysis structure for ZAF calculation
End If
End If
Next i%

' Re-load in case absorber only was set
Call GetElmSaveSampleOnly(Int(0), sample(), Int(0), Int(0))
If ierror Then Exit Sub

' Force unknown if type not specified
If sample(1).Type% = 0 Then sample(1).Type% = 2
sample(1).Datarows% = 1   ' always a single data point
sample(1).GoodDataRows% = 1
sample(1).LineStatus(1) = True      ' force status flag always true (good data point)
sample(1).AtomicPercentFlag% = True

' Set sample standard assignments so that ZAF and K-factors get loaded
For i% = 1 To sample(1).LastElm%
sample(1).StdAssigns%(i%) = sample(1).number%
Next i%

' Set TmpSample equal to OldSample so k factors and ZAF corrections get loaded in ZAFStd
tmpsample(1) = sample(1)

' Reload the element arrays
Call ElementGetData(sample())
If ierror Then Exit Sub

' Initialize calculations (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% <> MAXCORRECTION% Then
Call ZAFSetZAF(sample())
If ierror Then Exit Sub
Else
'Call ZAFSetZAF3(sample())
'If ierror Then Exit Sub
End If

' Force standard assignment for intensity calculation
For i% = 1 To sample(1).LastElm%
sample(1).StdAssigns%(i%) = MAXINTEGER%     ' fake standard assignment
Next i%

' Fake sample coating for ZAFStd calculation
If UseConductiveCoatingCorrectionForXrayTransmission Or UseConductiveCoatingCorrectionForElectronAbsorption Then                   ' fake standard coating
StandardCoatingFlag%(1) = sample(1).CoatingFlag%
StandardCoatingDensity!(1) = sample(1).CoatingDensity!
StandardCoatingThickness!(1) = sample(1).CoatingThickness!
StandardCoatingElement%(1) = sample(1).CoatingElement%
End If

' Run the intensity from concentration calculations on the "standard"
'Call ZAFStd(Int(1), analysis, sample(), tmpsample())
Call ZAFStd2(Int(1), analysis, sample(), tmpsample())
If ierror Then Exit Sub

Exit Sub

' Errors
Penepma12CalculateMatrixError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CalculateMatrix"
ierror = True
Exit Sub

End Sub

Sub Penepma12GetParFileComposition(n As Integer, tfilename As String, sample() As TypeSample)
' Extract the composition (and density) from the passed .par file
'
' Sample .par file composition
'
' Calcite.mat
' 6                         NELEM
'  12  6.00000E-02          IZ, atoms/mol
'  20  1.98320E+01          IZ, atoms/mol
'  25  6.10000E-02          IZ, atoms/mol
'  26  4.70000E-02          IZ, atoms/mol
'   8  6.00000E+01          IZ, atoms/mol
'   6  2.00000E+01          IZ, atoms/mol
'      2.71000E+00          Mass density (g/cm**3)

ierror = False
On Error GoTo Penepma12GetParFileCompositionError

Dim i As Integer, ip As Integer, atnum As Integer
Dim astring As String, bstring As String

Dim atoms(1 To MAXCHAN%) As Single

' Init sample
Call InitSample(sample())
If ierror Then Exit Sub

' Check for file
If Dir$(tfilename$) = vbNullString Then GoTo Penepma12GetParFileCompositionPARFileNotFound

' Open file and parse
Close #Temp1FileNumber%
Open tfilename$ For Input As #Temp1FileNumber%

' Read number of elements
Line Input #Temp1FileNumber%, astring$      ' read material filename line
sample(1).Name$ = Trim$(astring$)

Line Input #Temp1FileNumber%, astring$      ' read number of elements line
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub

sample(1).LastElm% = Val(Trim$(bstring$))
sample(1).LastChan% = sample(1).LastElm%

' Load atomic numbers and concentrations
For i% = 1 To sample(1).LastElm%
Line Input #Temp1FileNumber%, astring$

' Load atomic number (symbol)
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
atnum% = Val(bstring$)
If atnum% < 1 Or atnum% > MAXELM% Then GoTo Penepma12GetParFileCompositionBadAtomicNumber

sample(1).Elsyms$(i%) = Symlo$(atnum%)
sample(1).Xrsyms$(i%) = Deflin$(atnum%) ' just load defaults here (updated in Penepma12Save)

sample(1).numcat%(i%) = AllCat%(atnum%)  ' just load defaults here
sample(1).numoxd%(i%) = AllOxd%(atnum%)  ' just load defaults here

' Load molecules
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
atoms!(i%) = Val(Trim$(bstring$))

' Overload with Penepma08/12 atomic weights for self consistency in calculations
Penepma_TmpSample(1).AtomicWts!(i%) = pAllAtomicWts!(atnum%)
Next i%

' Input sample density
Line Input #Temp1FileNumber%, astring$
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
sample(1).SampleDensity! = Val(bstring$)

Close #Temp1FileNumber%

' Convert to weight fraction
Penepma_TmpSample(1).LastChan% = sample(1).LastChan%
Call ConvertAtomicToWeight(Penepma_TmpSample(1).LastChan%, Penepma_TmpSample(1).AtomicWts!(), Penepma_TmpSample(1).ElmPercents!(), atoms!())
If ierror Then Exit Sub

' Convert to weight percent
For i% = 1 To sample(1).LastChan%
sample(1).ElmPercents!(i%) = Penepma_TmpSample(1).ElmPercents!(i%) * 100#
Next i%

' Load other necessary values
sample(1).number% = MAXINTEGER%
sample(1).Set% = 1
sample(1).Type% = 1

' Load the measured element and x-ray into material A if not already there
If n% = 1 Then
ip% = IPOS1%(sample(1).LastElm%, Symlo$(MaterialMeasuredElement%), sample(1).Elsyms$())
If ip% = 0 Then
sample(1).LastElm% = sample(1).LastElm% + 1
sample(1).LastChan% = sample(1).LastElm%
sample(1).Elsyms$(sample(1).LastElm%) = Symlo$(MaterialMeasuredElement%)
sample(1).Xrsyms$(sample(1).LastElm%) = Xraylo$(MaterialMeasuredXray%)
sample(1).numcat%(i%) = AllCat%(MaterialMeasuredElement%)
sample(1).numoxd%(i%) = AllOxd%(MaterialMeasuredElement%)
sample(1).ElmPercents!(i%) = 0#
End If
End If

' Load the takeoff and kilovolts
sample(1).takeoff! = CSng(MaterialMeasuredTakeoff#)
sample(1).kilovolts! = CSng(MaterialMeasuredEnergy#)

Exit Sub

' Errors
Penepma12GetParFileCompositionError:
MsgBox Error$ & ", reading file " & tfilename$, vbOKOnly + vbCritical, "Penepma12GetParFileComposition"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma12GetParFileCompositionPARFileNotFound:
msg$ = "The specified .PAR file (" & tfilename$ & ") was not found. Please calculate the specified parameter file and try again"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12GetParFileComposition"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12GetParFileCompositionBadAtomicNumber:
msg$ = "Invalid atomic number (" & Format$(atnum%) & ") read from the specified .PAR file (" & tfilename$ & ")."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12GetParFileComposition"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12GetMatFileComposition(tfilename As String, sample() As TypeSample)
' Extract the composition (and density) from the passed .mat file
'
' Sample .mat file composition
'
' PENELOPE (v. 2012)  Material data file ...............
' Material:  Fe2SiO4
' Mass density = 8.96000000E+00 g/cm**3
' Number of elements in the molecule =  3
'   atomic number = 26, atoms / molecule = 0.499969314
'   atomic number = 14, atoms / molecule = 0.249984332
'   atomic number = 8, atoms / molecule = 1#

ierror = False
On Error GoTo Penepma12GetMatFileCompositionError

Dim i As Integer, atnum As Integer
Dim astring As String, bstring As String

Dim atoms(1 To MAXCHAN%) As Single

' Check for file
If Dir$(tfilename$) = vbNullString Then GoTo Penepma12GetMatFileCompositionMATFileNotFound

' Open file and parse
Close #Temp1FileNumber%
DoEvents
Open tfilename$ For Input As #Temp1FileNumber%

' Skip title (1st line)
Line Input #Temp1FileNumber%, astring$

' Read material name (2nd line)
Line Input #Temp1FileNumber%, astring$
bstring$ = Mid$(astring$, 12)
sample(1).Name$ = Trim$(bstring$)

' Read mass density line (3rd line)
Line Input #Temp1FileNumber%, astring$
bstring$ = Mid$(astring$, 16, 15)
sample(1).SampleDensity! = Val(Trim$(bstring$))

' Read number of elements
Line Input #Temp1FileNumber%, astring$
bstring$ = Mid$(astring$, 38)
sample(1).LastElm% = Val(Trim$(bstring$))
sample(1).LastChan% = sample(1).LastElm%

' Load atomic numbers and cocentrations
For i% = 1 To sample(1).LastElm%
Line Input #Temp1FileNumber%, astring$
bstring$ = Mid$(astring$, 19, 3)
atnum% = Val(bstring$)
If atnum% < 1 Or atnum% > MAXELM% Then GoTo Penepma12GetMatFileCompositionBadAtomicNumber

' Load atomic number (symbol)
sample(1).Elsyms$(i%) = Symlo$(atnum%)
sample(1).Xrsyms$(i%) = Deflin$(atnum%) ' just load defaults here (updated in Penepma12Save)

sample(1).numcat%(i%) = AllCat%(atnum%)  ' just load defaults here
sample(1).numoxd%(i%) = AllOxd%(atnum%)  ' just load defaults here

' Load molecule concentration
bstring$ = Mid$(astring$, 42, 14)
atoms!(i%) = Val(Trim$(bstring$))    ' convert to weight percent below

' Overload with Penepma08/12 atomic weights for self consistency in calculations
Penepma_TmpSample(1).AtomicWts!(i%) = pAllAtomicWts!(atnum%)
Next i%

Close #Temp1FileNumber%

' Convert to weight fraction
Penepma_TmpSample(1).LastChan% = sample(1).LastChan%
Call ConvertAtomicToWeight(Penepma_TmpSample(1).LastChan%, Penepma_TmpSample(1).AtomicWts!(), Penepma_TmpSample(1).ElmPercents!(), atoms!())
If ierror Then Exit Sub

' Convert to weight percent
For i% = 1 To sample(1).LastChan%
sample(1).ElmPercents!(i%) = Penepma_TmpSample(1).ElmPercents!(i%) * 100#
Next i%

' Load other necessary values
sample(1).number% = MAXINTEGER%
sample(1).Set% = 1
sample(1).Type% = 1

Exit Sub

' Errors
Penepma12GetMatFileCompositionError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12GetMatFileComposition"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma12GetMatFileCompositionMATFileNotFound:
msg$ = "The specified .MAT file (" & tfilename$ & ") was not found. Please calculate the specified material file and try again"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12GetMatFileComposition"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12GetMatFileCompositionBadAtomicNumber:
msg$ = "Invalid atomic number (" & Format$(atnum%) & ") read from the specified .MAT file (" & tfilename$ & ")."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12GetMatFileComposition"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12OutputKratios(tfolder As String)
' Ouput the original k-ratios and the modified k-ratios, "apparent" concentrations and correction factors
'  analysis.StdAssignsZAFCors(1,chan%) = absorption correction
'  analysis.StdAssignsZAFCors(2,chan%) = fluorescence correction
'  analysis.StdAssignsZAFCors(3,chan%) = atomic number correction
'  analysis.StdAssignsZAFCors(4,chan%) = ZAF correction (abscor*flucor*zedcor)
'  analysis.StdAssignsZAFCors!(5, i%) = stopping power
'  analysis.StdAssignsZAFCors!(6, i%) = backscatter
'  analysis.StdAssignsZAFCors!(7, i%) = std intensity
'  analysis.StdAssignsZAFCors!(8, i%) = unk intensity

ierror = False
On Error GoTo Penepma12OutputKratiosError

Dim i As Integer, j As Integer
Dim ipA As Integer, ipB As Integer, ipBStd As Integer
Dim unk_int_pri As Double, unk_int_flu As Double, unk_int_all As Double
Dim tzafA(1 To MAXZAFCOR%) As Double

Dim unk_krat_meas As Double, unk_zaf_meas As Double, unk_conc_meas As Double

' Initialize module values if case they are not loaded (Mat A does not contain the measured element)
Fanal_Krats! = 0#
Fanal_ZAFCors!(1) = 0#
Fanal_ZAFCors!(2) = 0#
Fanal_ZAFCors!(3) = 0#
Fanal_ZAFCors!(4) = 0#
unk_krat_meas# = 0#
unk_zaf_meas# = 0#
unk_conc_meas# = 0#

' Get composition from each .par file (loads MaterialMeasuredTakeoff# and MaterialMeasuredEnergy# for Penepma12CalculateMatrix)
Call Penepma12GetParFileComposition(Int(3), PENEPMA_Root$ & "\Fanal\db\" & ParameterFileBStd$, PENEPMA_SampleBStd())
If ierror Then Exit Sub
Call Penepma12CalculateMatrix(PENEPMA_Analysis, PENEPMA_SampleBStd(), Penepma_TmpSample())
If ierror Then Exit Sub

' Load BStd parameters for calculations below
For i% = 1 To MAXCHAN%
MatBStd_Krats!(i%) = PENEPMA_Analysis.StdAssignsKfactors!(i%)
MatBStd_StdPercents!(i%) = PENEPMA_Analysis.StdAssignsPercents!(i%)
For j% = 1 To MAXZAFCOR%
MatBStd_ZAFCors!(j%, i%) = PENEPMA_Analysis.StdAssignsZAFCors!(j%, i%)
Next j%
Next i%

' Do material B
Call Penepma12GetParFileComposition(Int(2), PENEPMA_Root$ & "\Fanal\db\" & ParameterFileB$, PENEPMA_SampleB())
If ierror Then Exit Sub
Call Penepma12CalculateMatrix(PENEPMA_Analysis, PENEPMA_SampleB(), Penepma_TmpSample())
If ierror Then Exit Sub

' Load B parameters for calculations below
For i% = 1 To MAXCHAN%
MatB_Krats!(i%) = PENEPMA_Analysis.StdAssignsKfactors!(i%)
MatB_StdPercents!(i%) = PENEPMA_Analysis.StdAssignsPercents!(i%)
For j% = 1 To MAXZAFCOR%
MatB_ZAFCors!(j%, i%) = PENEPMA_Analysis.StdAssignsZAFCors!(j%, i%)
Next j%
Next i%

' Do material A last!!!! (to keep specified element wt%s for non-absorbers in Penepma12OutputKratios3 calculation)
Call Penepma12GetParFileComposition(Int(1), PENEPMA_Root$ & "\Fanal\db\" & ParameterFileA$, PENEPMA_SampleA())
If ierror Then Exit Sub
Call Penepma12CalculateMatrix(PENEPMA_Analysis, PENEPMA_SampleA(), Penepma_TmpSample())
If ierror Then Exit Sub

' Load A parameters for calculations below
For i% = 1 To MAXCHAN%
MatA_Krats!(i%) = PENEPMA_Analysis.StdAssignsKfactors!(i%)
MatA_StdPercents!(i%) = PENEPMA_Analysis.StdAssignsPercents!(i%)

For j% = 1 To MAXZAFCOR%
MatA_ZAFCors!(j%, i%) = PENEPMA_Analysis.StdAssignsZAFCors!(j%, i%)
Next j%
Next i%

' Determine measured element in sample arrays
ipA% = IPOS1%(PENEPMA_SampleA(1).LastChan%, Symlo$(MaterialMeasuredElement%), PENEPMA_SampleA(1).Elsyms$())
ipB% = IPOS1%(PENEPMA_SampleB(1).LastChan%, Symlo$(MaterialMeasuredElement%), PENEPMA_SampleB(1).Elsyms$())
ipBStd% = IPOS1%(PENEPMA_SampleBStd(1).LastChan%, Symlo$(MaterialMeasuredElement%), PENEPMA_SampleBStd(1).Elsyms$())

' Correct intensity data for matrix effects (k-ratio is in %)
If ipA% > 0 And ipBStd% > 0 Then

' First calculate total fluorescence from each material
Call IOWriteLog(vbNullString)
For i% = 1 To nPoints&
If std_int#(i%) = 0# Then GoTo Penepma12OutputKratiosStdIntZero

' Calculate and load total fluorescence and primary k-ratios
fluA_k#(i%) = 100# * (flach#(i%) + flabr#(i%)) / std_int#(i%)
fluB_k#(i%) = 100# * (flbch#(i%) + flbbr#(i%)) / std_int#(i%)
prix_k#(i%) = 100# * pri_int#(i%) / std_int#(i%)
Next i%

' Calculate CalcZAF matrix correction from current ZAF selection (in case it is needed when measured element is not in Mat A)
tzafA#(1) = MatA_ZAFCors!(1, ipA%) / MatBStd_ZAFCors!(1, ipBStd%)    ' 1 = A only, 2 = F only, 3 = Z only, 4 = ZAF
tzafA#(2) = MatA_ZAFCors!(2, ipA%) / MatBStd_ZAFCors!(2, ipBStd%)    ' 1 = A only, 2 = F only, 3 = Z only, 4 = ZAF
tzafA#(3) = MatA_ZAFCors!(3, ipA%) / MatBStd_ZAFCors!(3, ipBStd%)    ' 1 = A only, 2 = F only, 3 = Z only, 4 = ZAF
tzafA#(4) = MatA_ZAFCors!(4, ipA%) / MatBStd_ZAFCors!(4, ipBStd%)    ' 1 = A only, 2 = F only, 3 = Z only, 4 = ZAF

' Calculate Mat A matrix correction based on known intensities and concentrations from Fanal (to be self consistant)
For i% = 1 To nPoints&
unk_int_pri# = pri_int#(i%)                                           ' calculate Mat A primary intensity
unk_int_flu# = flach#(i%) + flabr#(i%)                                ' calculate Mat A fluorescence intensity

' If performing bulk matrix correction, load fluorescence contribution from material B also
If ParameterFileA$ = ParameterFileB$ Then unk_int_flu# = unk_int_flu# + flbch#(i%) + flbbr#(i%)

' Calculate Mat A total intensity (or Mat A and Mat B total intensity if bulk matrix)
unk_int_all# = unk_int_flu# + unk_int_pri#

' Load material A matrix corrections with Fanal values (1 = A only, 2 = F only, 3 = Z only, 4 = ZAF)
If unk_int_all# <> 0# Then
Fanal_Krats! = unk_int_all# / std_int#(i%) * MatBStd_StdPercents!(ipBStd%) / 100#
Fanal_ZAFCors!(4) = MatA_StdPercents!(ipA%) / MatBStd_StdPercents!(ipBStd%) * std_int#(i%) / unk_int_all#

' Overload tzafA ZAF value for self consistent "apparent" concentration calculations below
tzafA#(4) = Fanal_ZAFCors!(4)

' If B std is a pure element and primary intensity is non-zero, then calculate fluorescence and combined ZA terms
If MatBStd_StdPercents!(ipBStd%) = 100# And pri_int#(i%) > 0# Then
Fanal_ZAFCors!(2) = 1# / (1# + unk_int_flu# / unk_int_all#)     ' calculate fluorescence only
'Fanal_ZAFCors!(3) = MatA_StdPercents!(ipA%) / MatBStd_StdPercents!(ipBStd%) * std_int#(i%) / unk_int_pri#
'Fanal_ZAFCors!(1) = MatA_StdPercents!(ipA%) / MatBStd_StdPercents!(ipBStd%) * std_int#(i%) / unk_int_pri#
Fanal_ZAFCors!(3) = Fanal_ZAFCors!(4) / Fanal_ZAFCors!(2)       ' calculate combined ZA
Fanal_ZAFCors!(1) = Fanal_ZAFCors!(4) / Fanal_ZAFCors!(2)       ' calculate combined ZA
End If
End If

' Calculate "measured" concentration using CalcZAF matrix correction based on actual apparent intensities
If MatBStd_StdPercents!(ipBStd%) = 100# Then
Call Penepma12OutputKratios3(ipA%, PENEPMA_Analysis, PENEPMA_SampleA(), yktotal#(i%), unk_krat_meas#, unk_zaf_meas#, unk_conc_meas#)
If ierror Then Exit Sub

' Load "measured" concentration
yktotal_meas#(i%) = unk_krat_meas#
yztotal_meas#(i%) = unk_zaf_meas#
yctotal_meas#(i%) = unk_conc_meas#
nsets& = 4

Else
nsets& = 3
End If

' Apply full material A correction to total intensity to obtain "apparent" concentration in material A
yctotal#(i%) = yktotal#(i%) * tzafA#(4) * MatBStd_StdPercents!(ipBStd%) / 100#

' Apply full material A correction to material A fluorescence to obtain "apparent" concentration from material A fluorescence
ycA_only#(i%) = fluA_k#(i%) * tzafA#(4) * MatBStd_StdPercents!(ipBStd%) / 100#

' Apply full material A correction to material B fluorescence (boundary) to obtain "apparent" concentration from material B fluorescence
ycb_only#(i%) = fluB_k#(i%) * tzafA#(4) * MatBStd_StdPercents!(ipBStd%) / 100#

' Apply full material A correction to primary x-ray to obtain "apparent" concentration
yc_prix#(i%) = prix_k#(i%) * tzafA#(4) * MatBStd_StdPercents!(ipBStd%) / 100#

' Output dist and k-ratio to log window
If ParameterFileA$ <> ParameterFileB$ Then
Call IOWriteLog("Penepma12OutputKratios: dist=" & Format$(xdist(i%)) & " um, kratio%= " & Format$(yktotal(i%)))
Else
If i% = nPoints& Then Call IOWriteLog("Penepma12OutputKratios: matrix kratio%= " & Format$(yktotal(i%)))
End If
Next i%

Else
GoTo Penepma12OutputKratiosElementNotFound
End If

' Create modified k-ratio output file
Call Penepma12OutputKratios2(tfolder$)
If ierror Then Exit Sub

Exit Sub

' Errors
Penepma12OutputKratiosError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12OutputKratios"
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12OutputKratiosElementNotFound:
msg$ = "The measured element " & Symlo$(MaterialMeasuredElement%) & " was not found in the material A or the material B Std compositions. This error should not occur, please contact Probe Software with details."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12OutputKratio"
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12OutputKratiosStdIntZero:
msg$ = "The standard intensity for the measured element " & Symlo$(MaterialMeasuredElement%) & " " & Xraylo$(MaterialMeasuredXray%) & " was zero for the material B Std composition (" & ParameterFileBStd$ & ") at " & Format$(MaterialMeasuredEnergy#) & " keV. This error should not occur, please contact Probe Software with details (and check the Fanal\k-ratios.dat file)."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12OutputKratio"
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12OutputKratios2(tfolder As String)
' Output k-ratio and modified results to a text file for off-line plotting
'  analysis.StdAssignsZAFCors(1,chan%) = absorption correction
'  analysis.StdAssignsZAFCors(2,chan%) = fluorescence correction
'  analysis.StdAssignsZAFCors(3,chan%) = atomic number correction
'  analysis.StdAssignsZAFCors(4,chan%) = ZAF correction (abscor*flucor*zedcor)

ierror = False
On Error GoTo Penepma12OutputKratios2Error

Dim astring As String
Dim i As Integer
Dim ipA As Integer, ipB As Integer, ipBStd As Integer
Dim temp1 As Single, temp2 As Single

 ' Load the modified k-ratio output file name
KRATIOS_DAT_File2$ = tfolder$ & "\k-ratios2.dat"

Open KRATIOS_DAT_File2$ For Output As #Temp1FileNumber%

' Create column labels
astring$ = vbNullString
astring$ = astring$ & VbDquote$ & "TO/keV/Elm/Xray" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "Distance (um)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "Total Inten. %" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "Fluor. Only %" & VbDquote$ & vbTab

' Calculate and load total fluorescence and primary k-ratios
astring$ = astring$ & VbDquote$ & "Fluor. A Only %" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "Fluor. B Only %" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "Pri. Int. Only %" & VbDquote$ & vbTab

astring$ = astring$ & VbDquote$ & "Calc. Total Conc. %" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "Calc. A Flu Conc. %" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "Calc. B Flu Conc. %" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "Calc. Pri. Conc. %" & VbDquote$ & vbTab

astring$ = astring$ & VbDquote$ & "Meas. Total Inten. %" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "Meas. ZAF" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "Meas. Total Conc. %" & VbDquote$ & vbTab

astring$ = astring$ & VbDquote$ & "Actual A Conc. %" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "Actual B Conc. %" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "Actual B Std Conc. %" & VbDquote$ & vbTab

' Output Fanal matrix factors (unkZAF/stdZAF)
astring$ = astring$ & VbDquote$ & "ZA u/s (Fanal)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "Flu u/s (Fanal)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "ZAF u/s (Fanal)" & VbDquote$ & vbTab

' Output related CalcZAF factors
astring$ = astring$ & VbDquote$ & "ZA u/s (CalcZAF)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "Flu u/s (CalcZAF)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "ZAF u/s (CalcZAF)" & VbDquote$ & vbTab

' Output normal CalcZAF matrix factors
astring$ = astring$ & VbDquote$ & "A Abs (CalcZAF)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "A Flu (CalcZAF)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "A Zed (CalcZAF)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "A ZAF (CalcZAF)" & VbDquote$ & vbTab

astring$ = astring$ & VbDquote$ & "B Abs (CalcZAF)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "B Flu (CalcZAF)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "B Zed (CalcZAF)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "B ZAF (CalcZAF)" & VbDquote$ & vbTab

astring$ = astring$ & VbDquote$ & "B Std Abs (CalcZAF)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "B Std Flu (CalcZAF)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "B Std Zed (CalcZAF)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "B Std ZAF (CalcZAF)" & VbDquote$ & vbTab

astring$ = astring$ & VbDquote$ & "Distance (ug/cm2)" & VbDquote$ & vbTab
Print #Temp1FileNumber%, astring$

' Determine measured element in sample arrays (only ipB can be zero)
ipA% = IPOS1%(PENEPMA_SampleA(1).LastChan%, Symlo$(MaterialMeasuredElement%), PENEPMA_SampleA(1).Elsyms$())
ipB% = IPOS1%(PENEPMA_SampleB(1).LastChan%, Symlo$(MaterialMeasuredElement%), PENEPMA_SampleB(1).Elsyms$())
ipBStd% = IPOS1%(PENEPMA_SampleBStd(1).LastChan%, Symlo$(MaterialMeasuredElement%), PENEPMA_SampleBStd(1).Elsyms$())
If ipA% = 0 Or ipBStd% = 0 Then GoTo Penepma12OutputKratios2NoEmittingElement

' Check for number of points
If nPoints& < 1 Then GoTo Penepma12OutputKratios2NoPoints

' Output data
For i% = 1 To nPoints&
astring$ = vbNullString
astring$ = astring$ & VbDquote$ & Format$(MaterialMeasuredTakeoff#) & "/" & Format$(MaterialMeasuredEnergy#) & "/" & Trim$(Symup$(MaterialMeasuredElement%)) & "/" & Xraylo$(MaterialMeasuredXray%) & VbDquote$ & vbTab
astring$ = astring$ & MiscAutoFormatD$(xdist#(i%)) & vbTab
astring$ = astring$ & MiscAutoFormatD$(yktotal#(i%)) & vbTab
astring$ = astring$ & MiscAutoFormatD$(ykfluor#(i%)) & vbTab

astring$ = astring$ & MiscAutoFormatD$(fluA_k#(i%)) & vbTab
astring$ = astring$ & MiscAutoFormatD$(fluB_k#(i%)) & vbTab
astring$ = astring$ & MiscAutoFormatD$(prix_k#(i%)) & vbTab

astring$ = astring$ & MiscAutoFormatD$(yctotal#(i%)) & vbTab
astring$ = astring$ & MiscAutoFormatD$(ycA_only#(i%)) & vbTab
astring$ = astring$ & MiscAutoFormatD$(ycb_only#(i%)) & vbTab
astring$ = astring$ & MiscAutoFormatD$(yc_prix#(i%)) & vbTab

astring$ = astring$ & MiscAutoFormatD$(yktotal_meas#(i%)) & vbTab
astring$ = astring$ & MiscAutoFormatD$(yztotal_meas#(i%)) & vbTab
astring$ = astring$ & MiscAutoFormatD$(yctotal_meas#(i%)) & vbTab

' Only output material B concentration if measured element is present in material B
astring$ = astring$ & MiscAutoFormat$(MatA_StdPercents!(ipA%)) & vbTab
If ipB% > 0 Then
astring$ = astring$ & MiscAutoFormat$(MatB_StdPercents!(ipB%)) & vbTab
Else
astring$ = astring$ & MiscAutoFormat$(INT_ZERO%) & vbTab
End If
astring$ = astring$ & MiscAutoFormat$(MatBStd_StdPercents!(ipBStd%)) & vbTab

' Output Fanal matrix factors (unkZAF/stdZAF)
astring$ = astring$ & MiscAutoFormat$(Fanal_ZAFCors!(1)) & vbTab     ' combined ZA terms
astring$ = astring$ & MiscAutoFormat$(Fanal_ZAFCors!(2)) & vbTab     ' fluorescence only
astring$ = astring$ & MiscAutoFormat$(Fanal_ZAFCors!(4)) & vbTab     ' ZAF correction

' Output related CalcZAF matrix factors (unkZAF/stdZAF)
temp1! = MatA_ZAFCors!(1, ipA%) * MatA_ZAFCors!(3, ipA%)
temp2! = MatBStd_ZAFCors!(1, ipBStd%) * MatBStd_ZAFCors!(3, ipBStd%)
astring$ = astring$ & MiscAutoFormat$(temp1! / temp2!) & vbTab                                        ' combined ZA term
astring$ = astring$ & MiscAutoFormat$(MatA_ZAFCors!(2, ipA%) / MatBStd_ZAFCors!(2, ipBStd%)) & vbTab  ' fluorescence only
astring$ = astring$ & MiscAutoFormat$(MatA_ZAFCors!(4, ipA%) / MatBStd_ZAFCors!(4, ipBStd%)) & vbTab  ' ZAF correction

' Output CalcZAF matrix factors
astring$ = astring$ & MiscAutoFormat$(MatA_ZAFCors!(1, ipA%)) & vbTab    ' absorption
astring$ = astring$ & MiscAutoFormat$(MatA_ZAFCors!(2, ipA%)) & vbTab    ' fluorescence
astring$ = astring$ & MiscAutoFormat$(MatA_ZAFCors!(3, ipA%)) & vbTab    ' atomic number
astring$ = astring$ & MiscAutoFormat$(MatA_ZAFCors!(4, ipA%)) & vbTab    ' ZAF correction

' Only output material B matrix corrections if measured element is present in material B
If ipB% > 0 Then
astring$ = astring$ & MiscAutoFormat$(MatB_ZAFCors!(1, ipB%)) & vbTab    ' absorption
astring$ = astring$ & MiscAutoFormat$(MatB_ZAFCors!(2, ipB%)) & vbTab    ' fluorescence
astring$ = astring$ & MiscAutoFormat$(MatB_ZAFCors!(3, ipB%)) & vbTab    ' atomic number
astring$ = astring$ & MiscAutoFormat$(MatB_ZAFCors!(4, ipB%)) & vbTab    ' ZAF correction
Else
astring$ = astring$ & MiscAutoFormat$(INT_ZERO%) & vbTab
astring$ = astring$ & MiscAutoFormat$(INT_ZERO%) & vbTab
astring$ = astring$ & MiscAutoFormat$(INT_ZERO%) & vbTab
astring$ = astring$ & MiscAutoFormat$(INT_ZERO%) & vbTab
End If

astring$ = astring$ & MiscAutoFormat$(MatBStd_ZAFCors!(1, ipBStd%)) & vbTab    ' absorption
astring$ = astring$ & MiscAutoFormat$(MatBStd_ZAFCors!(2, ipBStd%)) & vbTab    ' fluorescence
astring$ = astring$ & MiscAutoFormat$(MatBStd_ZAFCors!(3, ipBStd%)) & vbTab    ' atomic number
astring$ = astring$ & MiscAutoFormat$(MatBStd_ZAFCors!(4, ipBStd%)) & vbTab    ' ZAF correction

' Calculate mass distance (thickness in ug/cm2)
mdist#(i%) = PENEPMA_SampleA(1).SampleDensity! * Abs(xdist#(i%)) * CMPERMICRON# * MICROGRAMSPERGRAM&
astring$ = astring$ & MiscAutoFormatD$(mdist#(i%)) & vbTab

Print #Temp1FileNumber%, astring$
Next i%

Close #Temp1FileNumber%

' Confirm with user
msg$ = "Modified K-ratio file " & MiscGetFileNameOnly$(KRATIOS_DAT_File2$) & " was saved to " & tfolder$
Call IOWriteLog(msg$)

' Calculate hemispheric intensities for all radii
ReDim yhemi(1 To nPoints&) As Single
Call Penepma12CalculateHemisphere(nPoints&, xdist#(), ykfluor#(), yhemi!())
If ierror Then Exit Sub

Exit Sub

' Errors
Penepma12OutputKratios2Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12OutputKratios2"
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12OutputKratios2NoEmittingElement:
msg$ = "The emitting element is not in paramter A or in parameter BStd (this error should not occur)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12OutputKratios2"
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12OutputKratios2NoPoints:
msg$ = "The number of intensity points output is less than one so something went wrong with the Fanal calculation (try typing Fanal < Fanal.in from the Fanal prompt and see what happens)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12OutputKratios2"
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12UpdateForm()
' Update form enables based on flags

ierror = False
On Error GoTo Penepma12UpdateFormError

' Check if running Penfluor/Fitall/Fanal simulation/calculations
If SimulationInProgress Or FitParametersInProgress Or CalculationInProgress Then
FormPENEPMA12.CommandOK.Enabled = False
FormPENEPMA12.CommandClose.Enabled = False

FormPENEPMA12.CommandOutputMaterialA.Enabled = False
FormPENEPMA12.CommandOutputMaterialB.Enabled = False
FormPENEPMA12.CommandOutputMaterialBStd.Enabled = False

FormPENEPMA12.CommandOutputFormulaA.Enabled = False
FormPENEPMA12.CommandOutputFormulaB.Enabled = False
FormPENEPMA12.CommandOutputFormulaBStd.Enabled = False

FormPENEPMA12.CommandRunPenfluorRunAll.Enabled = False
FormPENEPMA12.CommandRunPenfluorA.Enabled = False
FormPENEPMA12.CommandRunPenfluorB.Enabled = False
FormPENEPMA12.CommandRunPenfluorBStd.Enabled = False

FormPENEPMA12.CommandBrowseMatA.Enabled = False
FormPENEPMA12.CommandBrowseMatB.Enabled = False
FormPENEPMA12.CommandBrowseMatBStd.Enabled = False

FormPENEPMA12.CommandBrowseParA.Enabled = False
FormPENEPMA12.CommandBrowseParB.Enabled = False
FormPENEPMA12.CommandBrowseParBStd.Enabled = False

FormPENEPMA12.CommandRunFanal.Enabled = False

' Binary calculations
FormPenepma12Binary.CommandOK.Enabled = False
FormPenepma12Binary.CommandClose.Enabled = False

FormPenepma12Binary.CommandBinaryCalculate.Enabled = False
FormPenepma12Binary.CommandCalculateComposition.Enabled = False
FormPenepma12Binary.CommandCalculateRandom.Enabled = False

FormPenepma12Binary.CommandExtract.Enabled = False
FormPenepma12Binary.CommandExtractRandom.Enabled = False
FormPenepma12Binary.CommandOutputPlotData.Enabled = False

FormPenepma12Binary.CommandPenPFE.Enabled = False
FormPenepma12Binary.CommandCalculateAlphas.Enabled = False
FormPenepma12Binary.CommandCalculateDensity.Enabled = False
FormPenepma12Binary.CommandTestLoadPenepmaAtomicWeights.Enabled = False
FormPenepma12Binary.CommandCalculateKratios.Enabled = False

Else
FormPENEPMA12.CommandOK.Enabled = True
FormPENEPMA12.CommandClose.Enabled = True

FormPENEPMA12.CommandOutputMaterialA.Enabled = True
FormPENEPMA12.CommandOutputMaterialB.Enabled = True
FormPENEPMA12.CommandOutputMaterialBStd.Enabled = True

FormPENEPMA12.CommandOutputFormulaA.Enabled = True
FormPENEPMA12.CommandOutputFormulaB.Enabled = True
FormPENEPMA12.CommandOutputFormulaBStd.Enabled = True

FormPENEPMA12.CommandRunPenfluorRunAll.Enabled = True
FormPENEPMA12.CommandRunPenfluorA.Enabled = True
FormPENEPMA12.CommandRunPenfluorB.Enabled = True
FormPENEPMA12.CommandRunPenfluorBStd.Enabled = True

FormPENEPMA12.CommandBrowseMatA.Enabled = True
FormPENEPMA12.CommandBrowseMatB.Enabled = True
FormPENEPMA12.CommandBrowseMatBStd.Enabled = True

FormPENEPMA12.CommandBrowseParA.Enabled = True
FormPENEPMA12.CommandBrowseParB.Enabled = True
FormPENEPMA12.CommandBrowseParBStd.Enabled = True

FormPENEPMA12.CommandRunFanal.Enabled = True

FormPenepma12Binary.CommandOK.Enabled = True
FormPenepma12Binary.CommandClose.Enabled = True

FormPenepma12Binary.CommandBinaryCalculate.Enabled = True
FormPenepma12Binary.CommandCalculateComposition.Enabled = True
FormPenepma12Binary.CommandCalculateRandom.Enabled = True

FormPenepma12Binary.CommandExtract.Enabled = True
FormPenepma12Binary.CommandExtractRandom.Enabled = True
FormPenepma12Binary.CommandOutputPlotData.Enabled = True

FormPenepma12Binary.CommandPenPFE.Enabled = True
FormPenepma12Binary.CommandCalculateAlphas.Enabled = True
FormPenepma12Binary.CommandCalculateDensity.Enabled = True
FormPenepma12Binary.CommandTestLoadPenepmaAtomicWeights.Enabled = True
FormPenepma12Binary.CommandCalculateKratios.Enabled = True
End If

Exit Sub

' Errors
Penepma12UpdateFormError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12UpdateForm"
ierror = True
Exit Sub

End Sub

Sub Penepma12ModifyParFiles(n As Integer, method As Integer)
' Modify the par files for the temp file names
'  n = 1 = Material A
'  n = 2 = Material B
'  n = 3 = Material B Std
' method = 0 modify for original file names (for after running Penfluor)
' method = 1 modify for temp file names (for before running Fanal)

ierror = False
On Error GoTo Penepma12ModifyParFilesError

Dim astring As String, bstring As String
Dim tfilename1 As String, tfilename2 As String

' Modify for after Penfluor
If method% = 0 Then
If n% = 1 Then tfilename1$ = PENEPMA_Root$ & "\Penfluor\material1.par"
If n% = 2 Then tfilename1$ = PENEPMA_Root$ & "\Penfluor\material2.par"
If n% = 3 Then tfilename1$ = PENEPMA_Root$ & "\Penfluor\material3.par"

If n% = 1 Then tfilename2$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileA$
If n% = 2 Then tfilename2$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileB$
If n% = 3 Then tfilename2$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileBStd$
End If

' Modify for before Fanal
If method% = 1 Then
If n% = 1 Then tfilename1$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileA$
If n% = 2 Then tfilename1$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileB$
If n% = 3 Then tfilename1$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileBStd$

If n% = 1 Then tfilename2$ = PENEPMA_Root$ & "\Fanal\db\material1.par"
If n% = 2 Then tfilename2$ = PENEPMA_Root$ & "\Fanal\db\material2.par"
If n% = 3 Then tfilename2$ = PENEPMA_Root$ & "\Fanal\db\material3.par"
End If

' Loop through sample input file and copy to new file with modified parameters
Close #Temp1FileNumber%
Close #Temp2FileNumber%
DoEvents
Open tfilename1$ For Input As #Temp1FileNumber%
Open tfilename2$ For Output As #Temp2FileNumber%

Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$
bstring$ = astring$

' Update the .mat name in the .par file for the temp name
If InStr(astring$, ".mat") > 0 Then

' After Penfluor
If method% = 0 Then
If n% = 1 Then bstring$ = MiscGetFileNameNoExtension$(ParameterFileA$) & ".mat"
If n% = 2 Then bstring$ = MiscGetFileNameNoExtension$(ParameterFileB$) & ".mat"
If n% = 3 Then bstring$ = MiscGetFileNameNoExtension$(ParameterFileBStd$) & ".mat"
End If

' Before Fanal
If method% = 1 Then
If n% = 1 Then bstring$ = "material1.mat"
If n% = 2 Then bstring$ = "material2.mat"
If n% = 3 Then bstring$ = "material3.mat"
End If
End If

Print #Temp2FileNumber%, bstring$
Loop

Close #Temp1FileNumber%
Close #Temp2FileNumber%

Exit Sub

' Errors
Penepma12ModifyParFilesError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12ModifyParFiles"
Close #Temp1FileNumber%
Close #Temp2FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma12UpdateCombo()
' Update the x-ray line combo control for the new element

ierror = False
On Error GoTo Penepma12UpdateComboError

Dim ip As Integer
Dim sym As String

sym$ = FormPENEPMA12.ComboElementStd.Text
ip% = IPOS1(MAXELM%, sym$, Symlo$())

If ip% > 0 Then
FormPENEPMA12.ComboXRayStd.Text = Deflin$(ip%)
End If

Exit Sub

' Errors
Penepma12UpdateComboError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12UpdateCombo"
ierror = True
Exit Sub

End Sub

Sub Penepma12Calculate()
' Calculate a single pair or a range of pure elements or binary composition .par files for the specified elements

ierror = False
On Error GoTo Penepma12CalculateError

Dim i As Integer, j As Integer, m As Integer, n As Integer
Dim response As Integer
Dim atime1 As Single
Dim tBinaryElement1 As Integer, tBinaryElement2 As Integer

icancelauto = False

' Warn if less than 1.0 keV minimum energy
If PenepmaMinimumElectronEnergy! < 1# Then
msg$ = "The minimum electron energy for Penepma kratio extractions is less than 1 keV. Penfluor by default only calculates down to 1 keV. Do you want to continue?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "Penepma12Calculate")
If response% = vbCancel Then Exit Sub
End If

' Calculating entire matrix range
If CalculateForMatrixRange Then

' BinaryMethod = 0  Calculate binary element compositional ranges for the range of the two elements
If BinaryMethod% = 0 Then

' Calculate number of binaries for range
m% = 0
For i% = BinaryElement1% To BinaryElement2%
For j% = i% To BinaryElement2%  ' do not duplicate binary pairs in reverse order
If i% <> j% Then m% = m% + 1
Next j%
Next i%

' Calculate estimated time for all binaries (always 99 to 1 wt%)
atime1! = m% * MaterialSimulationTime# * 8# * MAXBINARY%
atime1! = atime1! / SECPERDAY#

msg$ = "The complete calculation will take a long time to complete (at " & Format$(MaterialSimulationTime#) & " sec per simulation, " & Format$(m%) & " binaries of " & Format$(MAXBINARY%) & " compositions will take approximately " & MiscAutoFormat4$(atime1!) & " days to calculate). Though the calculation can be interrupted and restarted using the Do Not Overwrite Existing .PAR Files option. Are you sure you want to proceed?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "Penepma12Calculate")
If response% = vbCancel Then Exit Sub

TotalNumberOfSimulations& = CLng(m%) * MAXBINARY%    ' specify number of PAR files to create (always 99 to 1 wt%)
CurrentSimulationsNumber& = 1

n% = 0
tBinaryElement1% = BinaryElement1%      ' save
tBinaryElement2% = BinaryElement2%      ' save
For i% = tBinaryElement1% To tBinaryElement2%
For j% = i% To tBinaryElement2%

' Skip if sample element
If i% <> j% Then

n% = n% + 1
msg$ = vbCrLf & vbCrLf & "Calculating binary " & Format$(n%) & " of " & Format$(m%) & ": " & Trim$(Symup$(i%)) & "-" & Trim$(Symup$(j%)) & "..."
Call IOWriteLog(msg$)
Call IOStatusAuto(msg$)
DoEvents

BinaryElement1% = i%                    ' load matrix 1
BinaryElement2% = j%                    ' load matrix 2
Call Penepma12CalculateBinaries
BinaryElement1% = tBinaryElement1%      ' restore
BinaryElement2% = tBinaryElement2%      ' restore

If ierror Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

End If

Penepma12CalculateSkippingBinary:
Next j%
Next i%

msg$ = "All " & Format$(TotalNumberOfSimulations&) & " PAR file calculations are complete"
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12Calculate"
End If

' BinaryMethod = 1  Calculate pure element materials over the range of the two elements
If BinaryMethod% = 1 Then

m% = Abs(BinaryElement1% - BinaryElement2%) + 1
atime1! = m% * MaterialSimulationTime# * 8#
atime1! = atime1! / SECPERDAY#

msg$ = "The complete calculation of " & Format$(m%) & " pure element PAR files will take approximately " & MiscAutoFormat4$(atime1!) & " days to complete. Though it can be interrupted and restarted using the Do Not Overwrite Existing .PAR Files option). Are you sure you want to proceed?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "Penepma12Calculate")
If response% = vbCancel Then Exit Sub

TotalNumberOfSimulations& = m%    ' specify number of PAR files to create
CurrentSimulationsNumber& = 1

n% = 0
tBinaryElement1% = BinaryElement1%      ' save
tBinaryElement2% = BinaryElement2%      ' save
For i% = tBinaryElement1% To tBinaryElement2%

n% = n% + 1
msg$ = vbCrLf & vbCrLf & "Calculating binary " & Format$(n%) & " of " & Format$(m%) & ": " & Trim$(Symup$(i%)) & "..."
Call IOWriteLog(msg$)
Call IOStatusAuto(msg$)
DoEvents

BinaryElement1% = i%    ' calculate one at a time
BinaryElement2% = i%    ' calculate one at a time
Call Penepma12CalculateElements
BinaryElement1% = tBinaryElement1%      ' restore
BinaryElement2% = tBinaryElement2%      ' restore

If ierror Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

Next i%

msg$ = "All " & Format$(TotalNumberOfSimulations&) & " PAR file calculations are complete"
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12Calculate"
End If

' Calculate a single element pair
Else

' BinaryMethod = 0  Calculate binary element compositional ranges for the specified element pair
If BinaryMethod% = 0 Then
TotalNumberOfSimulations& = MAXBINARY%   ' specify number of PAR files to create
CurrentSimulationsNumber& = 1
Call Penepma12CalculateBinaries
If ierror Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If
End If

' BinaryMethod = 1  Calculate pure element compositions for the specified two elements
If BinaryMethod% = 1 Then
If BinaryElement1% = BinaryElement2% Then
TotalNumberOfSimulations& = 1   ' specify number of PAR files to create
Else
TotalNumberOfSimulations& = 2   ' specify number of PAR files to create
End If
CurrentSimulationsNumber& = 1
Call Penepma12CalculateElements
If ierror Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If
End If
End If

Call IOStatusAuto(vbNullString)
Exit Sub

' Errors
Penepma12CalculateError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12Calculate"
ierror = True
Exit Sub

End Sub

Sub Penepma12CalculateBinaries()
' Calculate a range of MAXBINARY% binary composition .par files for the specified element pair

ierror = False
On Error GoTo Penepma12CalculateBinariesError

Dim notfound1 As Boolean, notfound2 As Boolean
Dim k As Integer, ip As Integer
Dim pfilename As String

Dim binarynames(1 To MAXBINARY%) As String
Dim atoms(1 To 2) As Single

icancelauto = False

' Check that binary elements are different
If BinaryElement1% = BinaryElement2% Then GoTo Penepma12CalculateBinariesSame

' Check that BinaryElement1 is always less than BinaryElement2 (allow this calculation)
'If BinaryElement1% > BinaryElement2% Then GoTo Penepma12CalculateBinariesMore

' Check with user before calculating .mat or .par files (only if not calculating a range of binaries)
If Not CalculateForMatrixRange Then
For k% = 1 To MAXBINARY%

' Load element composition based on binary number (always 99 to 1 wt%)
PENEPMA_Sample(1).ElmPercents!(1) = BinaryRanges!(k%)
PENEPMA_Sample(1).ElmPercents!(2) = 100# - BinaryRanges!(k%)

' Load name and number for this binary
binarynames$(k%) = Trim$(Symup$(BinaryElement1%)) & "-" & Trim$(Symup$(BinaryElement2%)) & "_" & Format$(PENEPMA_Sample(1).ElmPercents!(1)) & "-" & Format$(PENEPMA_Sample(1).ElmPercents!(2))
MaterialFileA$ = binarynames$(k%) & ".mat"

' Check for existing .PAR file
pfilename$ = PENEPMA_Root$ & "\Penfluor\" & binarynames$(k%) & ".par"
If Not CalculateDoNotOverwritePAR Or (CalculateDoNotOverwritePAR And Dir$(pfilename$) = vbNullString) Then

' Check if existing PAR is lower precision and check if existing PAR is higher minimum energy- if either is true, run calculation
If (Not CalculateOnlyOverwriteLowerPrecisionPAR Or (CalculateOnlyOverwriteLowerPrecisionPAR And Dir$(pfilename$) = vbNullString) Or (CalculateOnlyOverwriteLowerPrecisionPAR And Penepma12PARLowerPrecision(pfilename$, MaterialSimulationTime#))) Or (Not CalculateOnlyOverwriteHigherMinimumEnergyPAR Or (CalculateOnlyOverwriteHigherMinimumEnergyPAR And Dir$(pfilename$) = vbNullString) Or (CalculateOnlyOverwriteHigherMinimumEnergyPAR And Penepma12PARHigherMinimumEnergy(pfilename$, PenepmaMinimumElectronEnergy!))) Then

' Check for existing PAR files with user
Call Penepma12RunPenfluorCheck(Int(1))
If ierror Then Exit Sub

End If
End If
Next k%

' Check calculation time
Call Penepma12RunPenfluorCheck2(Int(1))
If ierror Then Exit Sub
End If

' Treat all binary materials as material A for .par file calculations and calculate
' 99:1, 95:5, 90:10, 80:20, 60:40, 50:50, 40:60, 20:80, 10:90, 5:95, 1:99 binaries

' Create binary sample
PENEPMA_Sample(1).LastElm% = 2
PENEPMA_Sample(1).LastChan% = PENEPMA_Sample(1).LastElm%

PENEPMA_Sample(1).Elsyms$(1) = Symlo$(BinaryElement1%)
PENEPMA_Sample(1).Xrsyms$(1) = Deflin$(BinaryElement1%)  ' just load defaults here

PENEPMA_Sample(1).Elsyms$(2) = Symlo$(BinaryElement2%)
PENEPMA_Sample(1).Xrsyms$(2) = Deflin$(BinaryElement2%)  ' just load defaults here

PENEPMA_Sample(1).AtomicNums%(1) = AllAtomicNums%(BinaryElement1%)
PENEPMA_Sample(1).AtomicNums%(2) = AllAtomicNums%(BinaryElement2%)

ip% = IPOS1(MAXRAY% - 1, PENEPMA_Sample(1).Xrsyms$(1), Xraylo$())
PENEPMA_Sample(1).XrayNums%(1) = ip%
ip% = IPOS1(MAXRAY% - 1, PENEPMA_Sample(1).Xrsyms$(2), Xraylo$())
PENEPMA_Sample(1).XrayNums%(2) = ip%

' Check if x-ray is in database
Call Penepma12CheckXray(PENEPMA_Sample(), notfound1, notfound2)
If notfound1 Then
msg$ = "Penepma12CalculateBinaries: No x-ray data found for " & PENEPMA_Sample(1).Elsyms$(1) & " " & PENEPMA_Sample(1).Xrsyms$(1) & ". The " & PENEPMA_Sample(1).Elsyms$(1) & "-" & PENEPMA_Sample(1).Elsyms$(2) & " binary calculation will be skipped."
Call IOWriteLog(msg$)
Exit Sub
End If
If notfound2 Then
msg$ = "Penepma12CalculateBinaries: No x-ray data found for " & PENEPMA_Sample(1).Elsyms$(2) & " " & PENEPMA_Sample(1).Xrsyms$(2) & ". The " & PENEPMA_Sample(1).Elsyms$(1) & "-" & PENEPMA_Sample(1).Elsyms$(2) & " binary calculation will be skipped."
Call IOWriteLog(msg$)
Exit Sub
End If

' Load element data
PENEPMA_Sample(1).KilovoltsArray!(1) = CSng(MaterialMeasuredEnergy#)    ' set voltage just for over voltage check (not used by Penfluor)
PENEPMA_Sample(1).KilovoltsArray!(2) = CSng(MaterialMeasuredEnergy#)    ' set voltage just for over voltage check (not used by Penfluor)
Call ElementGetData(PENEPMA_Sample())
If ierror Then Exit Sub

' Overload with Penepma08/12 atomic weights for self consistency in calculations
PENEPMA_Sample(1).AtomicWts!(1) = pAllAtomicWts!(BinaryElement1%)
PENEPMA_Sample(1).AtomicWts!(2) = pAllAtomicWts!(BinaryElement2%)

' Calculate material file for each compositional binary (always 99 to 1 wt%)
For k% = 1 To MAXBINARY%

' Load element composition based on binary number (always 99 to 1 wt%)
PENEPMA_Sample(1).ElmPercents!(1) = BinaryRanges!(k%)
PENEPMA_Sample(1).ElmPercents!(2) = 100# - BinaryRanges!(k%)

' Calculate density based on composition
Call ConvertWeightToAtomic(PENEPMA_Sample(1).LastChan%, PENEPMA_Sample(1).AtomicWts!(), PENEPMA_Sample(1).ElmPercents!(), atoms!())
If ierror Then Exit Sub
PENEPMA_Sample(1).SampleDensity! = atoms!(1) * AllAtomicDensities!(BinaryElement1%) + atoms!(2) * AllAtomicDensities!(BinaryElement2%)

' Load name and number for this binary
binarynames$(k%) = Trim$(Symup$(BinaryElement1%)) & "-" & Trim$(Symup$(BinaryElement2%)) & "_" & Format$(PENEPMA_Sample(1).ElmPercents!(1)) & "-" & Format$(PENEPMA_Sample(1).ElmPercents!(2))
PENEPMA_Sample(1).Name$ = binarynames$(k%)
MaterialFileA$ = PENEPMA_Sample(1).Name$ & ".mat"
MaterialSelectedA% = MAXINTEGER%     ' any non-zero number
MaterialDensityA# = PENEPMA_Sample(1).SampleDensity!
If MaterialDensityA# < 1# Then MaterialDensityA# = 1#             ' force density to 1.0 in case the binary contains a gaseous element (to avoid detector geometry issues)

' Check for existing .PAR file
pfilename$ = PENEPMA_Root$ & "\Penfluor\" & binarynames$(k%) & ".par"
If Not CalculateDoNotOverwritePAR Or (CalculateDoNotOverwritePAR And Dir$(pfilename$) = vbNullString) Then

' Check if existing PAR is lower precision and check if existing PAR is higher minimum energy- if either is true, run calculation
If (Not CalculateOnlyOverwriteLowerPrecisionPAR Or (CalculateOnlyOverwriteLowerPrecisionPAR And Dir$(pfilename$) = vbNullString) Or (CalculateOnlyOverwriteLowerPrecisionPAR And Penepma12PARLowerPrecision(pfilename$, MaterialSimulationTime#))) Or (Not CalculateOnlyOverwriteHigherMinimumEnergyPAR Or (CalculateOnlyOverwriteHigherMinimumEnergyPAR And Dir$(pfilename$) = vbNullString) Or (CalculateOnlyOverwriteHigherMinimumEnergyPAR And Penepma12PARHigherMinimumEnergy(pfilename$, PenepmaMinimumElectronEnergy!))) Then

Call IOStatusAuto("Creating material input file based on formula " & PENEPMA_Sample(1).Name$ & "...")
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(0), MaterialInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

FormPENEPMA12.LabelProgress.Caption = "Creating Material File " & PENEPMA_Sample(1).Name$ & ".mat"
FormPENEPMA12.LabelRemainingTime.Caption = vbNullString

' Make material INP file (always a single file)
Screen.MousePointer = vbHourglass
Call Penepma12CreateMaterialINP(Int(1), PENEPMA_Sample())
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

' Create and run the necessary batch files
Screen.MousePointer = vbHourglass
Call Penepma12CreateMaterialBatch(Int(1), Int(1))
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

' Existing PAR file is same or higher precision or existing PAR file is same or lower minimum energy
Else
If CalculateOnlyOverwriteLowerPrecisionPAR Then msg$ = pfilename$ & " file is same or higher precision and Material calculations will be skipped..."
If CalculateOnlyOverwriteHigherMinimumEnergyPAR Then msg$ = pfilename$ & " file is same or lower minimum energy and Material calculations will be skipped..."
If CalculateOnlyOverwriteLowerPrecisionPAR And CalculateOnlyOverwriteHigherMinimumEnergyPAR Then msg$ = pfilename$ & " file is same or higher precision or same or higher minimum energy and Material calculations will be skipped..."
Call IOWriteLog(msg$)
End If

' PAR file already exists
Else
msg$ = pfilename$ & " file already exists and Material calculations will be skipped..."
Call IOWriteLog(msg$)
End If

DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(0), MaterialInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

Next k%

' Confirm with user
If Not CalculateForMatrixRange Then
msg$ = "All " & Format$(MAXBINARY%) & " MAT file calculations are complete"
Call IOWriteLog(msg$)
DoEvents
End If

' Create parameter files (always 99 to 1 wt%)
For k% = 1 To MAXBINARY%
MaterialFileA$ = binarynames$(k%) & ".mat"

' Check for existing .PAR file
pfilename$ = PENEPMA_Root$ & "\Penfluor\" & binarynames$(k%) & ".par"
If Not CalculateDoNotOverwritePAR Or (CalculateDoNotOverwritePAR And Dir$(pfilename$) = vbNullString) Then

' Check if existing PAR is lower precision and check if existing PAR is higher minimum energy- if either is true, run calculation
If (Not CalculateOnlyOverwriteLowerPrecisionPAR Or (CalculateOnlyOverwriteLowerPrecisionPAR And Dir$(pfilename$) = vbNullString) Or (CalculateOnlyOverwriteLowerPrecisionPAR And Penepma12PARLowerPrecision(pfilename$, MaterialSimulationTime#))) Or (Not CalculateOnlyOverwriteHigherMinimumEnergyPAR Or (CalculateOnlyOverwriteHigherMinimumEnergyPAR And Dir$(pfilename$) = vbNullString) Or (CalculateOnlyOverwriteHigherMinimumEnergyPAR And Penepma12PARHigherMinimumEnergy(pfilename$, PenepmaMinimumElectronEnergy!))) Then

' Run Penfluor and Fitall on material A
Call Penepma12RunPenFluor(Int(1))
If ierror Then Exit Sub

' Existing PAR file is same or higher precision or existing PAR file is same or lower minimum energy
Else
If CalculateOnlyOverwriteLowerPrecisionPAR Then msg$ = pfilename$ & " file is same or higher precision and Penfluor calculations will be skipped..."
If CalculateOnlyOverwriteHigherMinimumEnergyPAR Then msg$ = pfilename$ & " file is same or lower minimum energy and Penfluor calculations will be skipped..."
If CalculateOnlyOverwriteLowerPrecisionPAR And CalculateOnlyOverwriteHigherMinimumEnergyPAR Then msg$ = pfilename$ & " file is same or higher precision or same or higher minimum energy and Penfluor calculations will be skipped..."
Call IOWriteLog(msg$)
End If

' PAR file already exists
Else
msg$ = pfilename$ & " file already exists and Penfluor calculations will be skipped..."
Call IOWriteLog(msg$)
End If

CurrentSimulationsNumber& = CurrentSimulationsNumber& + 1
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(0), MaterialInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If
Next k%

If Not CalculateForMatrixRange Then
msg$ = "All " & Format$(TotalNumberOfSimulations&) & " PAR file calculations are complete"
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12CalculateBinaries"
End If

Exit Sub

' Errors
Penepma12CalculateBinariesError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CalculateBinaries"
ierror = True
Exit Sub

Penepma12CalculateBinariesSame:
msg$ = "The binary elements (" & Trim$(Symup$(BinaryElement1%)) & " and " & Trim$(Symup$(BinaryElement2%)) & ") are the same, but must be different for calculating a compositional range"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CalculateBinaries"
ierror = True
Exit Sub

Penepma12CalculateBinariesMore:
msg$ = "Binary matrix element 1 (" & Trim$(Symup$(BinaryElement1%)) & ") must precede binary matrix element 2 (" & Trim$(Symup$(BinaryElement2%)) & ") in the Periodic Table (don't ask why!)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CalculateBinaries"
ierror = True
Exit Sub

End Sub

Sub Penepma12CalculateElements()
' Calculate pure element composition .par files for the specified element pair range

ierror = False
On Error GoTo Penepma12CalculateElementsError

Dim n As Integer
Dim pfilename As String

icancelauto = False

' Check that BinaryElement1 is always equal or less than BinaryElement2 (they can be equal for a single element calculation)
If BinaryElement1% > BinaryElement2% Then GoTo Penepma12CalculateElementsMore

' Check with user before calculating .mat or .par files (only if not calculating a range of binaries)
If Not CalculateForMatrixRange Then
For n% = BinaryElement1% To BinaryElement2%
MaterialFileA$ = Trim$(Symup$(n%)) & ".mat"

' Check for existing .PAR file
pfilename$ = PENEPMA_Root$ & "\Penfluor\" & Trim$(Symup$(n%)) & ".par"
If Not CalculateDoNotOverwritePAR Or (CalculateDoNotOverwritePAR And Dir$(pfilename$) = vbNullString) Then

' Check for existing PAR files
Call Penepma12RunPenfluorCheck(Int(1))
If ierror Then Exit Sub

' Check calculation time
Call Penepma12RunPenfluorCheck2(Int(1))
If ierror Then Exit Sub

End If
Next n%
End If

' Create material files
For n% = BinaryElement1% To BinaryElement2%
MaterialFileA$ = Trim$(Symup$(n%)) & ".mat"

PENEPMA_Sample(1).LastElm% = 1      ' pure element
PENEPMA_Sample(1).LastChan% = PENEPMA_Sample(1).LastElm%
PENEPMA_Sample(1).Elsyms$(1) = Symlo$(n%)
PENEPMA_Sample(1).Xrsyms$(1) = Deflin$(n%)  ' just load defaults here

' Load element composition based on binary number
PENEPMA_Sample(1).ElmPercents!(1) = 100#

' Calculate density based on composition
PENEPMA_Sample(1).SampleDensity! = AllAtomicDensities!(n%)

' Load element data
Call ElementGetData(PENEPMA_Sample())
If ierror Then Exit Sub

' Overload with Penepma08/12 atomic weights for self consistency in calculations
PENEPMA_Sample(1).AtomicWts!(1) = pAllAtomicWts!(n%)

' Load name and number for this pure element
PENEPMA_Sample(1).Name$ = Trim$(Symup$(n%))
MaterialFileA$ = PENEPMA_Sample(1).Name$ & ".mat"
MaterialSelectedA% = MAXINTEGER%     ' any non-zero number
MaterialDensityA# = PENEPMA_Sample(1).SampleDensity!
If MaterialDensityA# < 1# Then MaterialDensityA# = 1#             ' force density to 1.0 in case the pure element is a gas (to avoid detector geometry issues)

' Check for existing .PAR file
pfilename$ = PENEPMA_Root$ & "\Penfluor\" & Trim$(Symup$(n%)) & ".par"
If Not CalculateDoNotOverwritePAR Or (CalculateDoNotOverwritePAR And Dir$(pfilename$) = vbNullString) Then

Call IOStatusAuto("Creating material input file based on formula " & PENEPMA_Sample(1).Name$ & "...")
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(0), MaterialInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

FormPENEPMA12.LabelProgress.Caption = "Creating Material File " & PENEPMA_Sample(1).Name$ & ".mat"
FormPENEPMA12.LabelRemainingTime.Caption = vbNullString

' Make material INP file (always a single file)
Screen.MousePointer = vbHourglass
Call Penepma12CreateMaterialINP(Int(1), PENEPMA_Sample())
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

' Create and run the necessary batch files
Screen.MousePointer = vbHourglass
Call Penepma12CreateMaterialBatch(Int(1), Int(1))
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

' PAR file already exists
Else
msg$ = pfilename$ & " file already exists and Material calculations will be skipped..."
Call IOWriteLog(msg$)
End If

DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(0), MaterialInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

Next n%

' Confirm with user
If Not CalculateForMatrixRange Then
msg$ = "Material file calculations are complete"
Call IOWriteLog(msg$)
DoEvents
End If

' Create PAR files
For n% = BinaryElement1% To BinaryElement2%
MaterialFileA$ = Trim$(Symup$(n%)) & ".mat"

' Check for existing .PAR file
pfilename$ = PENEPMA_Root$ & "\Penfluor\" & Trim$(Symup$(n%)) & ".par"
If Not CalculateDoNotOverwritePAR Or (CalculateDoNotOverwritePAR And Dir$(pfilename$) = vbNullString) Then

' Run Penfluor and Fitall on material A
Call Penepma12RunPenFluor(Int(1))
If ierror Then Exit Sub

' PAR file already exists
Else
msg$ = pfilename$ & " file already exists and Penfluor calculations will be skipped..."
Call IOWriteLog(msg$)
End If

CurrentSimulationsNumber& = CurrentSimulationsNumber& + 1
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(0), MaterialInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If
Next n%

' Confirm with user
If Not CalculateForMatrixRange Then
msg$ = "Parameter file calculations are complete"
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12CalculateElements"
End If

Exit Sub

' Errors
Penepma12CalculateElementsError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CalculateElements"
ierror = True
Exit Sub

Penepma12CalculateElementsSame:
msg$ = "The 2 pure elements (" & Trim$(Symup$(BinaryElement1%)) & " and " & Trim$(Symup$(BinaryElement2%)) & " ) are the same, but should be different or you are just repeating the same calculation twice"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CalculateElements"
ierror = True
Exit Sub

Penepma12CalculateElementsMore:
msg$ = "Binary matrix element 1 (" & Trim$(Symup$(BinaryElement1%)) & ") must precede binary matrix element 2 (" & Trim$(Symup$(BinaryElement2%)) & ") in the Periodic Table (don't ask why!)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CalculateElements"
ierror = True
Exit Sub

End Sub

Sub Penepma12CalculateRandom()
' Calculate binary or pure element composition .par files for the periodic table.
' The binary or pure element calculation is selected randomly and existing .PAR files are skipped if they are being calculated based on a shared look up table.

ierror = False
On Error GoTo Penepma12CalculateRandomError

Dim done As Boolean
Dim i As Integer, j As Integer, k As Integer, m As Integer
Dim im As Integer, mm As Integer
Dim response As Integer
Dim tBinaryElement1 As Integer, tBinaryElement2 As Integer
Dim tfilename As String

icancelauto = False

' Calculating entire matrix range
CalculateForMatrixRange = True  ' to skip user warnings

' BinaryMethod = 0  Calculate binary compositions over the entire periodic table
If BinaryMethod% = 0 Then

msg$ = "This binary compositional range calculation is designed to be performed by executing many multiple applications running in parallel utilizing a shared network folder for the Penepma12_PAR_Path to facilitate calculation of all binaries for the entire periodic table. The total calculation time will be approximately 50 years divided by the number of parallel applications running simultaneously (100 applications running in parallel will take approximately 6 months)! Are you sure you want to proceed?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "Penepma12CalculateRandom")
If response% = vbCancel Then Exit Sub

' Make sure PAR share folder exists
If Dir$(PENEPMA_PAR_Path$, vbDirectory) = vbNullString Then GoTo Penepma12CalculateRandomNoPARSharePath

' Calculate number of binaries for entire periodic table
m% = 0
ReDim CalculateRandomTable(1 To 3, 1 To 1) As Integer
For i% = 1 To MAXELM%
For j% = i% To MAXELM%  ' do not duplicate binary pairs in reverse order
If i% <> j% Then
m% = m% + 1
ReDim Preserve CalculateRandomTable(1 To 3, 1 To m%) As Integer
CalculateRandomTable%(1, m%) = i%   ' BinaryElement1
CalculateRandomTable%(2, m%) = j%   ' BinaryElement2
CalculateRandomTable%(3, m%) = m%   ' binary number (1 to m%)
End If
Next j%
Next i%

' Try to create a new PAR share file
Call Penepma12CalculateRandomCheck(Int(0), Int(3), m%, CalculateRandomTable%(), im%, mm%, done)
If ierror Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

TotalNumberOfSimulations& = CLng(m%) * MAXBINARY%    ' specify number of PAR files to create (should be 4950 binaries)
CurrentSimulationsNumber& = 1

tBinaryElement1% = BinaryElement1%      ' save
tBinaryElement2% = BinaryElement2%      ' save

' Check if randomly selected binary is being calculated already (im is selected binary, mm is binaries calculated so far)
Do Until done
Call Penepma12CalculateRandomCheck(Int(1), Int(3), m%, CalculateRandomTable%(), im%, mm%, done)
If ierror Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If
If done Then Exit Do

' Load next calculation
i% = CalculateRandomTable%(1, im%)
j% = CalculateRandomTable%(2, im%)

msg$ = vbCrLf & vbCrLf & "Calculating binary " & Format$(mm% + 1) & " of " & Format$(m%) & ": " & Trim$(Symup$(i%)) & "-" & Trim$(Symup$(j%)) & "..."
Call IOWriteLog(msg$)
Call IOStatusAuto(msg$)
DoEvents

BinaryElement1% = i%                    ' load matrix 1
BinaryElement2% = j%                    ' load matrix 2
Call Penepma12CalculateBinaries
BinaryElement1% = tBinaryElement1%      ' restore
BinaryElement2% = tBinaryElement2%      ' restore

' Update complete status for this calculation
Call Penepma12CalculateRandomCheck(Int(2), Int(3), m%, CalculateRandomTable%(), im%, mm%, done)
If ierror Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

' Copy completed PAR files to the Penepma PAR share folder (if folders are different)
If UCase$(PENEPMA_PAR_Path$) <> UCase$(PENEPMA_Root$ & "\penfluor") Then
For k% = 1 To MAXBINARY%
tfilename$ = Trim$(Symup$(i%)) & "-" & Trim$(Symup$(j%)) & "_" & Format$(BinaryRanges!(k%)) & "-" & Format$(100# - BinaryRanges!(k%)) & ".par"
If Dir$(PENEPMA_Root$ & "\penfluor" & "\" & tfilename$) <> vbNullString Then
FileCopy PENEPMA_Root$ & "\penfluor" & "\" & tfilename$, PENEPMA_PAR_Path$ & "\" & tfilename$
FileCopy PENEPMA_Root$ & "\penfluor" & "\" & MiscGetFileNameNoExtension$(tfilename$) & ".in", PENEPMA_PAR_Path$ & "\" & MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(tfilename$)) & ".in"
Call IOWriteLog("Copied " & tfilename$ & " PAR file to " & PENEPMA_PAR_Path$)
End If
Next k%
End If

DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(0), MaterialInProgress)
If ierror Then Exit Sub
Call Penepma12CheckTermination2(Int(1), SimulationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If
Loop

msg$ = "All " & Format$(TotalNumberOfSimulations&) & " PAR file calculations are complete"
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12CalculateRandom"
End If

' BinaryMethod = 1  Calculate pure element materials over the periodic table
If BinaryMethod% = 1 Then

msg$ = "This pure element calculation is designed to be performed by executing many multiple applications running in parallel utilizing a shared network folder for the Penepma12_PAR_Path to facilitate calculation of all binaries for the entire periodic table. The total calculation time will be approximately 33 days divided by the number of parallel applications running simultaneously (10 applications running in parallel will take approximately 3.3 days)! Are you sure you want to proceed?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "Penepma12CalculateRandom")
If response% = vbCancel Then Exit Sub

m% = 0
ReDim CalculateRandomTable(1 To 2, 1 To 1) As Integer
For i% = 1 To MAXELM%
m% = m% + 1
ReDim Preserve CalculateRandomTable(1 To 2, 1 To m%) As Integer
CalculateRandomTable%(1, m%) = i%   ' BinaryElement1 and BinaryElement2
CalculateRandomTable%(2, m%) = m%   ' Sequence number
Next i%

' Try to create a new PAR share file
Call Penepma12CalculateRandomCheck(Int(0), Int(2), m%, CalculateRandomTable%(), im%, mm%, done)
If ierror Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

TotalNumberOfSimulations& = m%    ' specify number of PAR files to create
CurrentSimulationsNumber& = 1

tBinaryElement1% = BinaryElement1%      ' save
tBinaryElement2% = BinaryElement2%      ' save

' Check if randomly selected pure element is being calculated already (im is selected element, mm is elements calculated so far)
Do Until done
Call Penepma12CalculateRandomCheck(Int(1), Int(2), m%, CalculateRandomTable%(), im%, mm%, done)
If ierror Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If
If done Then Exit Do

' Load next calculation
i% = CalculateRandomTable%(1, im%)

msg$ = vbCrLf & vbCrLf & "Calculating pure element " & Format$(mm%) & " of " & Format$(m%) & ": " & Trim$(Symup$(i%)) & "..."
Call IOWriteLog(msg$)
Call IOStatusAuto(msg$)
DoEvents

BinaryElement1% = i%    ' load same element for both
BinaryElement2% = i%    ' load same element for both
Call Penepma12CalculateElements
BinaryElement1% = tBinaryElement1%      ' restore
BinaryElement2% = tBinaryElement2%      ' restore

' Update complete status for this calculation
Call Penepma12CalculateRandomCheck(Int(2), Int(2), m%, CalculateRandomTable%(), im%, mm%, done)
If ierror Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(0), MaterialInProgress)
If ierror Then Exit Sub
Call Penepma12CheckTermination2(Int(1), SimulationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If
Loop

msg$ = "All " & Format$(TotalNumberOfSimulations&) & " PAR file calculations are complete"
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12CalculateRandom"
End If

Call IOStatusAuto(vbNullString)
Exit Sub

' Errors
Penepma12CalculateRandomError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CalculateRandom"
ierror = True
Exit Sub

Penepma12CalculateRandomNoPARSharePath:
msg$ = "The specified Penepma PAR Share Path " & PENEPMA_PAR_Path$ & " does not exist. Please create the specified folder and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CalculateRandom"
ierror = True
Exit Sub

End Sub

Sub Penepma12Extract()
' Extract the k-ratios from .par files for a matrix or boundary fluorescence model
'  matrix correction when Mat A = Mat B
'  boundary correction is when Mat A <> Mat B

ierror = False
On Error GoTo Penepma12ExtractError

Dim i As Integer, j As Integer
Dim response As Integer

Dim tMaterialMeasuredGridPoints As Integer
Dim tMaterialMeasuredDistance As Double
Dim tExtractElement As Integer
Dim tExtractMatrix As Integer
Dim tExtractMatrixA1 As Integer
Dim tExtractMatrixA2 As Integer
Dim tExtractMatrixB1 As Integer
Dim tExtractMatrixB2 As Integer

icancelauto = False

' Warn if less than 1.0 keV minimum energy and not auto adjust minimum energy
If PenepmaMinimumElectronEnergy! < 1# And FormPENEPMA12.CheckAutoAdjustMinimumEnergy.Value = vbUnchecked Then
msg$ = "The minimum electron energy for Penepma kratio extractions is less than 1 keV. Since Penfluor usually only calculates down to 1 keV, this might be problematic. Do you want to continue?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "Penepma12Extract")
If response% = vbCancel Then Exit Sub
End If

' Calculate for the entire range
If ExtractForSpecifiedRange Then

If ExtractMethod% = 0 Then  ' boundary extract disabled in dialog for specified range
msg$ = "The specified range boundary extract k-ratio calculations will take several months to complete (assuming all binary .PAR files necessary are present). Are you sure you want to proceed?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "Penepma12Extract")
If response% = vbCancel Then Exit Sub
End If

If ExtractMethod% = 1 Then
msg$ = "The specified range matrix extract k-ratio calculations will take several days to complete (assuming all binary .PAR files necessary are present). Are you sure you want to proceed?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "Penepma12Extract")
If response% = vbCancel Then Exit Sub
End If

' ExtractMethod = 0  Extract k-ratios for boundary fluorescence using a range of elements
If ExtractMethod% = 0 Then
tExtractMatrixA1% = ExtractMatrixA1%      ' save original
tExtractMatrixA2% = ExtractMatrixA2%      ' save original
tExtractMatrixB1% = ExtractMatrixB1%      ' save original
tExtractMatrixB2% = ExtractMatrixB2%      ' save original

msg$ = "Feature not implemented yet"
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12Extract"
Exit Sub

For i% = tExtractMatrixA1% To tExtractMatrixA2%
For j% = tExtractMatrixB1% To tExtractMatrixB2%
If i% <> j% Then

ExtractMatrixA1% = j%                   ' load (this is not right?!)
ExtractMatrixA2% = i%                   ' load
ExtractMatrixB1% = j%                   ' load
ExtractMatrixB2% = i%                   ' load

tMaterialMeasuredDistance# = MaterialMeasuredDistance#      ' save
MaterialMeasuredDistance# = 0#                              ' to force exponential distance calculations in Fanal
Call Penepma12ExtractBoundary
MaterialMeasuredDistance# = tMaterialMeasuredDistance#      ' restore
If ierror Then Exit Sub

End If
Next j%
Next i%

ExtractMatrixA1% = tExtractMatrixA1%      ' restore original
ExtractMatrixA2% = tExtractMatrixA2%      ' restore original
ExtractMatrixB1% = tExtractMatrixB1%      ' restore original
ExtractMatrixB2% = tExtractMatrixB2%      ' restore original
End If

' ExtractMethod = 1  Extract k-ratios for matrix corrections using a range of elements
If ExtractMethod% = 1 Then
tExtractElement% = ExtractElement%      ' save original emitting element
tExtractMatrix% = ExtractMatrix%        ' save original matrix element

For i% = tExtractElement% To tExtractMatrix%
For j% = tExtractElement% To tExtractMatrix%    ' need to calculate each emitter separately for each binary

' Skip if same element
If i% <> j% Then

ExtractElement% = i%      ' load emitting element
ExtractMatrix% = j%       ' load matrix element

tMaterialMeasuredGridPoints% = MaterialMeasuredGridPoints%      ' save
MaterialMeasuredGridPoints% = 1     ' use a single point for matrix calculations
Call Penepma12ExtractMatrix
MaterialMeasuredGridPoints% = tMaterialMeasuredGridPoints%      ' restore

If ierror Then
Exit Sub
Call IOStatusAuto(vbNullString)
End If

End If

Penepma12ExtractSkippingBinary:
Next j%
Next i%

ExtractElement% = tExtractElement%      ' restore original emitting element
ExtractMatrix% = tExtractMatrix%        ' restore original matrix element
End If

' Calculate a single boundary or matrix
Else

If ExtractMethod% = 0 Then msg$ = "The single boundary extract k-ratio calculations will take 10 to 20 hours to complete (assuming all binary .PAR files necessary are present). Are you sure you want to proceed?"
If ExtractMethod% = 1 Then msg$ = "The single matrix extract k-ratio calculations will take several hours to complete (assuming all binary .PAR files necessary are present). Are you sure you want to proceed?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "Penepma12Extract")
If response% = vbCancel Then Exit Sub

' ExtractMethod = 0  Extract k-ratios for boundary fluorescence
If ExtractMethod% = 0 Then
tMaterialMeasuredDistance# = MaterialMeasuredDistance#      ' save
MaterialMeasuredDistance# = 0#                              ' to force exponential distance calculations in Fanal
Call Penepma12ExtractBoundary
MaterialMeasuredDistance# = tMaterialMeasuredDistance#      ' restore
If ierror Then
Call IOStatusAuto(vbNullString)
Exit Sub
End If
End If

' ExtractMethod = 1  Extract k-ratios for matrix corrections
If ExtractMethod% = 1 Then
tMaterialMeasuredGridPoints% = MaterialMeasuredGridPoints%
MaterialMeasuredGridPoints% = 1     ' use a single point for matrix calculations
Call Penepma12ExtractMatrix
MaterialMeasuredGridPoints% = tMaterialMeasuredGridPoints%
If ierror Then
Call IOStatusAuto(vbNullString)
Exit Sub
End If
End If
End If

Call IOStatusAuto(vbNullString)
Exit Sub

' Errors
Penepma12ExtractError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12Extract"
ierror = True
Exit Sub

End Sub

Sub Penepma12Extract2()
' Extract the k-ratios from .par files for a matrix or boundary fluorescence model based on a user specified formula
'  matrix correction when Mat A = Mat B
'  boundary correction is when Mat A <> Mat B

ierror = False
On Error GoTo Penepma12Extract2Error

Dim i As Integer, j As Integer
Dim response As Integer

Dim tMaterialMeasuredGridPoints As Integer
Dim tMaterialMeasuredDistance As Double

icancelauto = False

' Get a formula composition from the user
Call Penepma12CalculateGetComposition(Int(1), PENEPMA_OldSample())
If ierror Then Exit Sub

' Check for at least one element
If PENEPMA_OldSample(1).LastChan% < 1 Then Exit Sub

' Warn if less than 1.0 keV minimum energy and not auto adjuect minimum energy
If PenepmaMinimumElectronEnergy! < 1# And FormPENEPMA12.CheckAutoAdjustMinimumEnergy.Value = vbUnchecked Then
msg$ = "The minimum electron energy for Penepma kratio extractions is less than 1 keV. Since Penfluor usually only calculates down to 1 keV, this might be problematic. Do you want to continue?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "Penepma12Extract2")
If response% = vbCancel Then Exit Sub
End If

If ExtractMethod% = 0 Then msg$ = "The boundary extract k-ratio calculations will take several weeks to complete (assuming all binary .PAR files necessary are present). Are you sure you want to proceed?"
If ExtractMethod% = 1 Then msg$ = "The matrix extract k-ratio calculations will take several days to complete (assuming all binary .PAR files necessary are present). Are you sure you want to proceed?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "Penepma12Extract2")
If response% = vbCancel Then Exit Sub

For i% = 1 To PENEPMA_OldSample(1).LastChan%
For j% = 1 To PENEPMA_OldSample(1).LastChan%

' Skip if same element
If i% <> j% Then

ExtractElement% = PENEPMA_OldSample(1).AtomicNums%(i%)
ExtractMatrix% = PENEPMA_OldSample(1).AtomicNums%(j%)

' ExtractMethod = 0  Extract k-ratios for boundary fluorescence
If ExtractMethod% = 0 Then
tMaterialMeasuredDistance# = MaterialMeasuredDistance#      ' save
MaterialMeasuredDistance# = 0#                              ' to force exponential distance calculations in Fanal
Call Penepma12ExtractBoundary
MaterialMeasuredDistance# = tMaterialMeasuredDistance#      ' restore
If ierror Then
Call IOStatusAuto(vbNullString)
Exit Sub
End If
End If

' ExtractMethod = 1  Extract k-ratios for matrix corrections
If ExtractMethod% = 1 Then
tMaterialMeasuredGridPoints% = MaterialMeasuredGridPoints%
MaterialMeasuredGridPoints% = 1     ' use a single point for matrix calculations
Call Penepma12ExtractMatrix
MaterialMeasuredGridPoints% = tMaterialMeasuredGridPoints%
If ierror Then
Call IOStatusAuto(vbNullString)
Exit Sub
End If
End If

End If
Next j%
Next i%


Call IOStatusAuto(vbNullString)
Exit Sub

' Errors
Penepma12Extract2Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12Extract2"
ierror = True
Exit Sub

End Sub

Sub Penepma12ExtractBoundary()
' Extract k-ratios for boundary fluorescence
' Material A and Material B are two different binary compositions
' Material B Std is the pure element for the emitting element

ierror = False
On Error GoTo Penepma12ExtractBoundaryError

Dim l As Integer, m As Integer, n As Integer, ipA As Integer, ipB As Integer
Dim j As Integer, k As Integer
Dim response As Integer
Dim eng As Single, edg As Single
Dim unk_int_pri As Double, unk_int_flu As Double, unk_int_all As Double
Dim notfound As Boolean

Dim tfolder As String, tfilename As String
Dim pfilename1 As String, pfilename2 As String, pfilename3 As String

Dim binarynamesA(1 To MAXBINARY%) As String
Dim binarynamesB(1 To MAXBINARY%) As String

Dim t1 As Single, t2 As Single

icancelauto = False

' Dimension k-ratio and alpha factor arrays
Call InitKratios
If ierror Then Exit Sub

' If extract element is not in matrix A or B, calculation cannot be done
notfound = True
If ExtractElement% = ExtractMatrixA1% Then notfound = False
If ExtractElement% = ExtractMatrixA2% Then notfound = False
If ExtractElement% = ExtractMatrixB1% Then notfound = False
If ExtractElement% = ExtractMatrixB2% Then notfound = False
If notfound Then GoTo Penepma12ExtractBoundaryNotFound

' If extract is not in material A, alpha factors cannot be calculated
If ExtractElement% <> ExtractMatrixA1% And ExtractElement% <> ExtractMatrixA2% Then
msg$ = "The extract element " & Trim$(Symup$(ExtractElement%)) & " is not in material A. Boundary alpha-factors will not be able to be calculated. Do you still want to proceed?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "Penepma12ExtractBoundary")
If response% = vbCancel Then Exit Sub
End If

' Check that emitting and matrix elements are different
If ExtractMatrixA1% = ExtractMatrixA2% Then GoTo Penepma12ExtractBoundarySameA
If ExtractMatrixB1% = ExtractMatrixB2% Then GoTo Penepma12ExtractBoundarySameB

msg$ = vbCrLf & "Extracting Boundary K-Ratios for " & Trim$(Symup$(ExtractElement%)) & " in " & Trim$(Symup$(ExtractMatrixA1%)) & "-" & Trim$(Symup$(ExtractMatrixA2%)) & " adjacent to " & Trim$(Symup$(ExtractMatrixB1%)) & "-" & Trim$(Symup$(ExtractMatrixB2%))
Call IOWriteLog(msg$)

' Check for Fanal\boundary folder
tfolder$ = PENEPMA_Root$ & "\Fanal\boundary"
If Dir$(tfolder$, vbDirectory) = vbNullString Then MkDir tfolder$

' Determine which PAR files are available for extraction (all subsequent input/output assumes unswapped elements)
Call Penepma12ExtractSwapped(ExtractMatrixA1%, ExtractMatrixA2%, ExtractMatrixB1%, ExtractMatrixB2%)
If ierror Then Exit Sub

' Check that .PAR files already exist (if not then just skip calculation)
For j% = 1 To MAXBINARY%    ' material A and material B
DoEvents

' Load sample A element composition based on binary number (always 99 to 1 wt%)
If Not BinaryElementsSwappedA Then
PENEPMA_SampleA(1).ElmPercents!(1) = BinaryRanges!(j%)
PENEPMA_SampleA(1).ElmPercents!(2) = 100# - BinaryRanges!(j%)
Else
PENEPMA_SampleA(1).ElmPercents!(1) = BinaryRanges!(MAXBINARY - (j% - 1))
PENEPMA_SampleA(1).ElmPercents!(2) = 100# - BinaryRanges!(MAXBINARY% - (j% - 1))
End If

' Load sample B element composition based on binary number
If Not BinaryElementsSwappedB Then
PENEPMA_SampleB(1).ElmPercents!(1) = BinaryRanges!(j%)
PENEPMA_SampleB(1).ElmPercents!(2) = 100# - BinaryRanges!(j%)
Else
PENEPMA_SampleB(1).ElmPercents!(1) = BinaryRanges!(MAXBINARY - (j% - 1))
PENEPMA_SampleB(1).ElmPercents!(2) = 100# - BinaryRanges!(MAXBINARY% - (j% - 1))
End If

' Load name and number for this binary (swap element symbols for PAR file if necessary as either will do)
If Not BinaryElementsSwappedA Then
binarynamesA$(j%) = Trim$(Symup$(ExtractMatrixA1%)) & "-" & Trim$(Symup$(ExtractMatrixA2%)) & "_" & Format$(PENEPMA_SampleA(1).ElmPercents!(1)) & "-" & Format$(PENEPMA_SampleA(1).ElmPercents!(2))
Else
binarynamesA$(j%) = Trim$(Symup$(ExtractMatrixA2%)) & "-" & Trim$(Symup$(ExtractMatrixA1%)) & "_" & Format$(PENEPMA_SampleA(1).ElmPercents!(1)) & "-" & Format$(PENEPMA_SampleA(1).ElmPercents!(2))
End If
PENEPMA_SampleA(1).Name$ = binarynamesA$(j%)

If Not BinaryElementsSwappedB Then
binarynamesB$(j%) = Trim$(Symup$(ExtractMatrixB1%)) & "-" & Trim$(Symup$(ExtractMatrixB2%)) & "_" & Format$(PENEPMA_SampleB(1).ElmPercents!(1)) & "-" & Format$(PENEPMA_SampleB(1).ElmPercents!(2))
Else
binarynamesB$(j%) = Trim$(Symup$(ExtractMatrixB2%)) & "-" & Trim$(Symup$(ExtractMatrixB1%)) & "_" & Format$(PENEPMA_SampleB(1).ElmPercents!(1)) & "-" & Format$(PENEPMA_SampleB(1).ElmPercents!(2))
End If
PENEPMA_SampleB(1).Name$ = binarynamesB$(j%)

' Load PAR file for this binary
ParameterFileA$ = PENEPMA_SampleA(1).Name$ & ".par"
ParameterFileB$ = PENEPMA_SampleB(1).Name$ & ".par"
ParameterFileBStd$ = Trim$(Symup$(ExtractElement%)) & ".par"

' Check that A and B files are not the same
If (Not BinaryElementsSwappedA And Not BinaryElementsSwappedB) Or (BinaryElementsSwappedA And BinaryElementsSwappedB) Then
If ParameterFileA$ = ParameterFileB$ Then GoTo Penepma12ExtractBoundaryFilesAreSame
End If

' Check .PAR files for boundary k-ratio extraction
pfilename1$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileA$
pfilename2$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileB$
pfilename3$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileBStd$

' Check in Binary and Pure folders if not found in Penfluor
If Dir$(pfilename1$) = vbNullString Then
tfilename$ = PENEPMA_Root$ & "\Penfluor\Binary\" & ParameterFileA$
If Dir$(tfilename$) <> vbNullString Then FileCopy tfilename$, PENEPMA_Root$ & "\Penfluor\" & ParameterFileA$
If Dir$(MiscGetFileNameNoExtension$(tfilename$) & ".in") <> vbNullString Then FileCopy MiscGetFileNameNoExtension$(tfilename$) & ".in", PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(ParameterFileA$)) & ".in"
End If
If Dir$(pfilename2$) = vbNullString Then
tfilename$ = PENEPMA_Root$ & "\Penfluor\Binary\" & ParameterFileB$
If Dir$(tfilename$) <> vbNullString Then FileCopy tfilename$, PENEPMA_Root$ & "\Penfluor\" & ParameterFileB$
If Dir$(MiscGetFileNameNoExtension$(tfilename$) & ".in") <> vbNullString Then FileCopy MiscGetFileNameNoExtension$(tfilename$) & ".in", PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(ParameterFileB$)) & ".in"
End If
If Dir$(pfilename3$) = vbNullString Then
tfilename$ = PENEPMA_Root$ & "\Penfluor\Pure\" & ParameterFileBStd$
If Dir$(tfilename$) <> vbNullString Then FileCopy tfilename$, PENEPMA_Root$ & "\Penfluor\" & ParameterFileBStd$
If Dir$(MiscGetFileNameNoExtension$(tfilename$) & ".in") <> vbNullString Then FileCopy MiscGetFileNameNoExtension$(tfilename$) & ".in", PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(ParameterFileBStd$)) & ".in"
End If

' Check if binary composition is present
If Dir$(pfilename1$) = vbNullString Or Dir$(pfilename2$) = vbNullString Or Dir$(pfilename3$) = vbNullString Then
msg$ = "Parameter files " & ParameterFileA$ & " or " & ParameterFileB$ & " or " & ParameterFileBStd$ & " are not found and the k-ratio extract calculation will be skipped..."
Call IOWriteLog(msg$)
Exit Sub
End If

DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(3), CalculationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If
Next j%     ' material A and material B

' Check only the last A and B PAR file that emitting element is present in material A or material B and give warning if not
tfilename$ = pfilename1$        ' material A
Call Penepma12GetParFileComposition(Int(1), tfilename$, PENEPMA_SampleA())
If ierror Then Exit Sub
ipA% = IPOS1%(PENEPMA_SampleA(1).LastElm%, Symlo$(ExtractElement%), PENEPMA_SampleA(1).Elsyms$())

tfilename$ = pfilename2$        ' material B
Call Penepma12GetParFileComposition(Int(2), tfilename$, PENEPMA_SampleB())
If ierror Then Exit Sub
ipB% = IPOS1%(PENEPMA_SampleB(1).LastElm%, Symlo$(ExtractElement%), PENEPMA_SampleB(1).Elsyms$())

If ipA% = 0 And ipB% = 0 Then
msg$ = "Neither material A (" & ParameterFileA$ & ") nor material B (" & ParameterFileB$ & ") contains the measured element " & Trim$(Symup$(ExtractElement%)) & ". Are you sure you want to proceed?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "Penepma12ExtractBoundary")
If response% = vbCancel Then Exit Sub
End If

' Load measured element
MaterialMeasuredElement% = ExtractElement%

' Loop on each valid x-ray
For l% = 1 To MAXRAY% - 1
'For l% = 1 To 1         ' testing purposes (Ka only)
MaterialMeasuredXray% = l%

' Write column labels for first energy only (note that elements are unswapped)
tfilename$ = Format$(ExtractMatrixA1%) & "-" & Format$(ExtractMatrixA2%) & "_" & Format$(ExtractMatrixB1%) & "-" & Format$(ExtractMatrixB2%) & "_" & Format$(MaterialMeasuredTakeoff#) & "_" & Format$(MaterialMeasuredElement%) & "-" & Format$(MaterialMeasuredXray%) & ".txt"
If Not CalculateDoNotOverwriteTXT Or (CalculateDoNotOverwriteTXT And Dir$(PENEPMA_Root$ & "\Fanal\boundary\" & tfilename$) = vbNullString) Then

Call Penepma12CalculateReadWriteBinaryDataBoundary(Int(0), tfolder$, tfilename$, CSng(MaterialMeasuredEnergy#), CLng(0))
If ierror Then Exit Sub

msg$ = vbCrLf & "Extracting Boundary Fluorescence K-Ratios for " & Trim$(Symup$(MaterialMeasuredElement%)) & " " & Xraylo$(MaterialMeasuredXray%) & "..."
Call IOWriteLog(msg$)

' Loop on each beam voltage from 1 to 50 keV
For m% = 1 To 50
'For m% = 15 To 15       ' testing purposes (15 keV only)
'For m% = 15 To 16       ' testing purposes (15 and 16 keV only)
MaterialMeasuredEnergy# = m%

' Get x-ray data
Call XrayGetEnergy(MaterialMeasuredElement%, MaterialMeasuredXray%, eng!, edg!)
If ierror Then Exit Sub

' Check for valid x-ray line (excitation energy must be less than beam energy) (and greater than PenepmaMinimumElectronEnergy!)
If eng! <> 0# And edg! <> 0# And edg! < MaterialMeasuredEnergy# And edg! > PenepmaMinimumElectronEnergy! Then

' Double check that specific transition exists (see table 6.2 in Penelope-2006-NEA-pdf)
Call PenepmaGetPDATCONFTransition(MaterialMeasuredElement%, MaterialMeasuredXray%, t1!, t2!)
If ierror Then Exit Sub

' If both shells have ionization energies, it is ok to calculate
If t1! <> 0# And t2! <> 0# Then

msg$ = vbCrLf & "Calculating Boundary Fluorescence K-Ratios for " & Trim$(Symup$(MaterialMeasuredElement%)) & " " & Xraylo$(MaterialMeasuredXray%) & " in " & Trim$(Symup$(ExtractMatrixA1%)) & "-" & Trim$(Symup$(ExtractMatrixA2%)) & " adjacent to " & Trim$(Symup$(ExtractMatrixB1%)) & "-" & Trim$(Symup$(ExtractMatrixB2%)) & " at " & Format$(MaterialMeasuredEnergy#) & " keV..."
Call IOWriteLog(msg$)

' Load each combination of binary compositional parameter files
For j% = 1 To MAXBINARY%    ' material A
'For j% = 1 To MAXBINARY% Step MAXBINARY% - 1 ' material A (testing only)
DoEvents
For k% = 1 To MAXBINARY%    ' material B
'For k% = 1 To MAXBINARY% Step MAXBINARY% - 1  ' material B (testing only)
DoEvents
PENEPMA_SampleA(1).LastElm% = 2
PENEPMA_SampleA(1).LastChan% = 2
PENEPMA_SampleB(1).LastElm% = 2
PENEPMA_SampleB(1).LastChan% = 2

' Load sample A element composition based on binary number (need to reload everything because of call to Penepma12GetPARFileComposition above)
If Not BinaryElementsSwappedA Then
PENEPMA_SampleA(1).ElmPercents!(1) = BinaryRanges!(j%)
PENEPMA_SampleA(1).ElmPercents!(2) = 100# - BinaryRanges!(j%)
Else
PENEPMA_SampleA(1).ElmPercents!(1) = BinaryRanges!(MAXBINARY - (j% - 1))
PENEPMA_SampleA(1).ElmPercents!(2) = 100# - BinaryRanges!(MAXBINARY% - (j% - 1))
End If

' Load sample B element composition based on binary number (always 99 to 1 wt%)
If Not BinaryElementsSwappedB Then
PENEPMA_SampleB(1).ElmPercents!(1) = BinaryRanges!(k%)
PENEPMA_SampleB(1).ElmPercents!(2) = 100# - BinaryRanges!(k%)
Else
PENEPMA_SampleB(1).ElmPercents!(1) = BinaryRanges!(MAXBINARY - (k% - 1))
PENEPMA_SampleB(1).ElmPercents!(2) = 100# - BinaryRanges!(MAXBINARY% - (k% - 1))
End If

' Load name and number for this binary (swap element symbols for PAR file if necessary as either will do)
If Not BinaryElementsSwappedA Then
binarynamesA$(j%) = Trim$(Symup$(ExtractMatrixA1%)) & "-" & Trim$(Symup$(ExtractMatrixA2%)) & "_" & Format$(PENEPMA_SampleA(1).ElmPercents!(1)) & "-" & Format$(PENEPMA_SampleA(1).ElmPercents!(2))
Else
binarynamesA$(j%) = Trim$(Symup$(ExtractMatrixA2%)) & "-" & Trim$(Symup$(ExtractMatrixA1%)) & "_" & Format$(PENEPMA_SampleA(1).ElmPercents!(1)) & "-" & Format$(PENEPMA_SampleA(1).ElmPercents!(2))
End If
PENEPMA_SampleA(1).Name$ = binarynamesA$(j%)

If Not BinaryElementsSwappedB Then
binarynamesB$(k%) = Trim$(Symup$(ExtractMatrixB1%)) & "-" & Trim$(Symup$(ExtractMatrixB2%)) & "_" & Format$(PENEPMA_SampleB(1).ElmPercents!(1)) & "-" & Format$(PENEPMA_SampleB(1).ElmPercents!(2))
Else
binarynamesB$(k%) = Trim$(Symup$(ExtractMatrixB2%)) & "-" & Trim$(Symup$(ExtractMatrixB1%)) & "_" & Format$(PENEPMA_SampleB(1).ElmPercents!(1)) & "-" & Format$(PENEPMA_SampleB(1).ElmPercents!(2))
End If
PENEPMA_SampleB(1).Name$ = binarynamesB$(k%)

' Load parameter files
ParameterFileA$ = PENEPMA_SampleA(1).Name$ & ".par"
ParameterFileB$ = PENEPMA_SampleB(1).Name$ & ".par"
ParameterFileBStd$ = Trim$(Symup$(ExtractElement%)) & ".par"

' Load density for sample A (beam incident material) and B (boundary material)
PENEPMA_SampleA(1).SampleDensity! = Penepma12GetParFileDensityOnly(PENEPMA_Root$ & "\Fanal\db\" & ParameterFileA$)
If ierror Then Exit Sub
PENEPMA_SampleB(1).SampleDensity! = Penepma12GetParFileDensityOnly(PENEPMA_Root$ & "\Fanal\db\" & ParameterFileB$)
If ierror Then Exit Sub

' Overload with Penepma08/12 atomic weights for self consistency in calculations
If Not BinaryElementsSwappedA Then
PENEPMA_SampleA(1).AtomicWts!(1) = pAllAtomicWts!(ExtractMatrixA1%)
PENEPMA_SampleA(1).AtomicWts!(2) = pAllAtomicWts!(ExtractMatrixA2%)
Else
PENEPMA_SampleA(1).AtomicWts!(1) = pAllAtomicWts!(ExtractMatrixA2%)
PENEPMA_SampleA(1).AtomicWts!(2) = pAllAtomicWts!(ExtractMatrixA1%)
End If
If Not BinaryElementsSwappedB Then
PENEPMA_SampleB(1).AtomicWts!(1) = pAllAtomicWts!(ExtractMatrixB1%)
PENEPMA_SampleB(1).AtomicWts!(2) = pAllAtomicWts!(ExtractMatrixB2%)
Else
PENEPMA_SampleB(1).AtomicWts!(1) = pAllAtomicWts!(ExtractMatrixB2%)
PENEPMA_SampleB(1).AtomicWts!(2) = pAllAtomicWts!(ExtractMatrixB1%)
End If

' Double check that par file is in db folder (check penfluor folder in case manually copied)
If Dir$(PENEPMA_Root$ & "\Fanal\db\" & ParameterFileA$) = vbNullString Then
If Dir$(PENEPMA_Root$ & "\Penfluor\" & ParameterFileA$) <> vbNullString Then
FileCopy PENEPMA_Root$ & "\Penfluor\" & ParameterFileA$, PENEPMA_Root$ & "\Fanal\db\" & ParameterFileA$
If Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileA$) & ".in") <> vbNullString Then FileCopy PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileA$) & ".in", PENEPMA_Root$ & "\Fanal\db\" & MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(ParameterFileA$)) & ".in"
Else
GoTo Penepma12ExtractBoundaryPARFilesNotFoundA
End If
End If

If Dir$(PENEPMA_Root$ & "\Fanal\db\" & ParameterFileB$) = vbNullString Then
If Dir$(PENEPMA_Root$ & "\Penfluor\" & ParameterFileB$) <> vbNullString Then
FileCopy PENEPMA_Root$ & "\Penfluor\" & ParameterFileB$, PENEPMA_Root$ & "\Fanal\db\" & ParameterFileB$
If Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileA$) & ".in") <> vbNullString Then FileCopy PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileA$) & ".in", PENEPMA_Root$ & "\Fanal\db\" & MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(ParameterFileB$)) & ".in"
Else
GoTo Penepma12ExtractBoundaryPARFilesNotFoundB
End If
End If

If Dir$(PENEPMA_Root$ & "\Fanal\db\" & ParameterFileBStd$) = vbNullString Then
If Dir$(PENEPMA_Root$ & "\Penfluor\" & ParameterFileBStd$) <> vbNullString Then
FileCopy PENEPMA_Root$ & "\Penfluor\" & ParameterFileBStd$, PENEPMA_Root$ & "\Fanal\db\" & ParameterFileBStd$
If Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileA$) & ".in") <> vbNullString Then FileCopy PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileA$) & ".in", PENEPMA_Root$ & "\Fanal\db\" & MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(ParameterFileBStd$)) & ".in"
Else
GoTo Penepma12ExtractBoundaryPARFilesNotFoundBStd
End If
End If

msg$ = "Calculating Boundary Fluorescence K-Ratios for " & Trim$(Symup$(MaterialMeasuredElement%)) & " " & Xraylo$(MaterialMeasuredXray%) & " in " & Trim$(Symup$(ExtractMatrixA1%)) & "-" & Trim$(Symup$(ExtractMatrixA2%)) & " adjacent to " & Trim$(Symup$(ExtractMatrixB1%)) & "-" & Trim$(Symup$(ExtractMatrixB2%)) & " at " & Format$(MaterialMeasuredEnergy#) & " keV..."
Call IOStatusAuto(msg$)
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(3), CalculationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

FormPENEPMA12.LabelProgress.Caption = "Extracting boundary fluorescence k-ratios from " & PENEPMA_SampleA(1).Name$ & ".par" & " adjacent to " & PENEPMA_SampleB(1).Name$ & ".par..."
FormPENEPMA12.LabelRemainingTime.Caption = vbNullString

' Check the parameters files
Call Penepma12RunFanal
If ierror Then Exit Sub

' Run the Fanal program
Call Penepma12RunFanal1
If ierror Then Exit Sub

' Get k-ratio data from Fanal k-ratio file
Call Penepma12LoadPlotData
If ierror Then Exit Sub

' Check that Fanal produced good data
If nPoints& > 0 Then

' Dimension boundary arrays (npoints should always be the same)
ReDim Preserve Boundary_ZAF_Kratios(1 To MAXBINARY%, 1 To MAXBINARY%, 1 To nPoints&) As Double
ReDim Preserve Boundary_ZAF_Factors(1 To MAXBINARY%, 1 To MAXBINARY%, 1 To nPoints&) As Single

ReDim Preserve Boundary_Linear_Distances(1 To nPoints&) As Single
ReDim Preserve Boundary_Mass_Distances(1 To MAXBINARY%, 1 To nPoints&) As Single

' Store essential boundary k-ratio data to data array
For n% = 1 To nPoints&
If DebugMode Then
msg$ = "K-ratio%= " & yktotal#(n%) & " at distance " & Format$(xdist#(n%)) & ", for " & Trim$(Symup$(ExtractElement%)) & " " & Xraylo$(MaterialMeasuredXray%) & " in " & ParameterFileA$ & " adjacent to " & ParameterFileB$ & " using standard " & ParameterFileBStd$
Call IOWriteLog(msg$)
End If

unk_int_pri# = pri_int#(n%)                                                ' calculate Mat A primary intensity
unk_int_flu# = flach#(n%) + flabr#(n%) + flbch#(n%) + flbbr#(n%)           ' calculate Mat A and Mat B fluorescence intensity
unk_int_all# = unk_int_flu# + pri_int#(nPoints&)                           ' calculate total intensity

' Load k-ratio arrays
'Boundary_ZAF_Kratios#(k%, j%, n%) = yktotal#(n%)                            ' calculate total k-ratio
Boundary_ZAF_Kratios#(k%, j%, n%) = 100# * unk_int_all# / std_int#(n%)       ' calculate total k-ratio

' Load linear distances
Boundary_Linear_Distances!(n%) = CSng(xdist#(n%))

' Load mass distances (ug/cm^2)
Boundary_Mass_Distances!(j%, n%) = PENEPMA_SampleA(1).SampleDensity! * Abs(xdist#(n%)) * CMPERMICRON# * MICROGRAMSPERGRAM&
Next n%

' Load sample A and B densities
Boundary_Material_A_Densities!(j%) = PENEPMA_SampleA(1).SampleDensity!
Boundary_Material_B_Densities!(k%) = PENEPMA_SampleB(1).SampleDensity!

DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(3), CalculationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

' Check for Pause button
Do Until Not RealTimePauseAutomation
DoEvents
Sleep 200
Loop

' Nothing to output
Else
msg$ = "No intensity data to output for " & Trim$(Symup$(MaterialMeasuredElement%)) & " " & Xraylo$(MaterialMeasuredXray%) & " in " & ParameterFileA$ & " adjacent to " & ParameterFileB$ & " using standard " & ParameterFileBStd$
Call IOWriteLog(msg$)
End If
Next k%
Next j%

msg$ = "All " & Format$(MAXBINARY% * MAXBINARY%) & " boundary k-ratio calculations are complete for " & Trim$(Symup$(MaterialMeasuredElement%)) & " " & Xraylo$(MaterialMeasuredXray%) & "..."
Call IOWriteLog(msg$)
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(3), CalculationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

' Calculate alpha factors and fits for boundary data
Call Penepma12ExtractBoundaryCalculate(nPoints&)
If ierror Then Exit Sub

' Write binary k-ratio fluorescence data to file for the current beam energy and specified element and x-ray
tfilename$ = Format$(ExtractMatrixA1%) & "-" & Format$(ExtractMatrixA2%) & "_" & Format$(ExtractMatrixB1%) & "-" & Format$(ExtractMatrixB2%) & "_" & Format$(MaterialMeasuredTakeoff#) & "_" & Format$(MaterialMeasuredElement%) & "-" & Format$(MaterialMeasuredXray%) & ".txt"
Call Penepma12CalculateReadWriteBinaryDataBoundary(Int(2), tfolder$, tfilename$, CSng(MaterialMeasuredEnergy#), nPoints&)
If ierror Then Exit Sub

' Valid transitions for this element
End If

' Valid x-ray line at this energy
End If

DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(3), CalculationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If
Next m%

msg$ = "Boundary calculations are complete for " & Trim$(Symup$(MaterialMeasuredElement%)) & " " & Trim$(Xraylo$(MaterialMeasuredXray%)) & " in " & Trim$(Symup$(ExtractMatrixA1%)) & "-" & Trim$(Symup$(ExtractMatrixA2%)) & " adjacent to " & Trim$(Symup$(ExtractMatrixB1%)) & "-" & Trim$(Symup$(ExtractMatrixB2%)) & " at " & Format$(MaterialMeasuredEnergy#) & " keV"
Call IOWriteLog(msg$)

Else
msg$ = "Skipping k-ratio boundary extraction for " & tfilename$ & "..."
Call IOWriteLog(msg$)
End If

DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(3), CalculationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If
Next l%

Call IOStatusAuto(vbNullString)
msg$ = "All boundary fluorescence k-ratio extractions and alpha factor calculations are complete"
Call IOWriteLog(msg$)
DoEvents

Exit Sub

' Errors
Penepma12ExtractBoundaryError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12ExtractBoundary"
ierror = True
Exit Sub

Penepma12ExtractBoundaryNotFound:
msg$ = "The extract element " & Trim$(Symup$(ExtractElement%)) & " was not found in either material A binary (" & Trim$(Symup$(ExtractMatrixA1%)) & "-" & Trim$(Symup$(ExtractMatrixA2%)) & ") or material B binary  (" & Trim$(Symup$(ExtractMatrixB1%)) & "-" & Trim$(Symup$(ExtractMatrixB2%)) & "). The calculation cannot be performed."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12ExtractBoundary"
ierror = True
Exit Sub

Penepma12ExtractBoundarySameA:
msg$ = "The matrix binary elements (" & Trim$(Symup$(ExtractMatrixA1%)) & " and " & Trim$(Symup$(ExtractMatrixA2%)) & ") are the same, but must be different for calculating binary fluorescence effects"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12ExtractBoundary"
ierror = True
Exit Sub

Penepma12ExtractBoundarySameB:
msg$ = "The matrix binary elements (" & Trim$(Symup$(ExtractMatrixB1%)) & " and " & Trim$(Symup$(ExtractMatrixB2%)) & ") are the same, but must be different for calculating binary fluorescence effects"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12ExtractBoundary"
ierror = True
Exit Sub

Penepma12ExtractBoundaryPARFilesNotFoundA:
msg$ = "The specified .PAR file (" & ParameterFileA$ & ") was not found in the Fanal\db or Penfluor folders. Please calculate the specified parameter file and try again"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12ExtractBoundary"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12ExtractBoundaryPARFilesNotFoundB:
msg$ = "The specified .PAR file (" & ParameterFileB$ & ") was not found in the Fanal\db or Penfluor folders. Please calculate the specified parameter file and try again"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12ExtractBoundary"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12ExtractBoundaryPARFilesNotFoundBStd:
msg$ = "The specified .PAR file (" & ParameterFileBStd$ & ") was not found in the Fanal\db or Penfluor folders. Please calculate the specified parameter file and try again"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12ExtractBoundary"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12ExtractBoundaryFilesAreSame:
msg$ = "The parameter files (" & ParameterFileA$ & " and " & ParameterFileB$ & ") are the same. You cannot calculate boundary effects if the two materials are the same."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12ExtractBoundary"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12ExtractBoundaryFilesNotFound:
msg$ = "The Fanal output files were not found. Check Fanal.exe by running the input file manually using a command prompt from the Fanal folder and typing Fanal < Fanal.in"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12ExtractBoundary"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12ExtractMatrix()
' Extract k-ratios for matrix correction
' Material A and Material B are the same binary compositions and material B Std is the pure element
'  analysis.StdAssignsZAFCors(1,chan%) = absorption correction
'  analysis.StdAssignsZAFCors(2,chan%) = fluorescence correction
'  analysis.StdAssignsZAFCors(3,chan%) = atomic number correction
'  analysis.StdAssignsZAFCors(4,chan%) = ZAF correction (abscor*flucor*zedcor)
'  analysis.StdAssignsZAFCors!(5, i%) = stopping power
'  analysis.StdAssignsZAFCors!(6, i%) = backscatter
'  analysis.StdAssignsZAFCors!(7, i%) = std intensity
'  analysis.StdAssignsZAFCors!(8, i%) = unk intensity

ierror = False
On Error GoTo Penepma12ExtractMatrixError

Dim BinaryElementsSwapped As Boolean
Dim k As Integer, l As Integer, m As Integer, ip As Integer, ipA As Integer
Dim inum1 As Integer, inum2 As Integer, inum3 As Integer
Dim eng As Single, edg As Single, tovervoltage As Single
Dim tempF As Double, tempZA As Double
Dim unk_int_pri As Double, unk_int_flu As Double, unk_int_all As Double

Dim tfolder As String, tfilename As String
Dim pfilename As String
Dim pfilename1 As String, pfilename2 As String, pfilename3 As String

Dim pvalue As Single
Dim pstring As String

Dim FanalIntensitiesOutput As Boolean

Dim binarynames(1 To MAXBINARY%) As String

Dim t1 As Single, t2 As Single

icancelauto = False

' Dimension k-ratio and alpha factor arrays
Call InitKratios
If ierror Then Exit Sub

' Check that emitting and matrix elements are different
If ExtractElement% = ExtractMatrix% Then GoTo Penepma12ExtractMatrixSame

' Check for Fanal\matrix folder
tfolder$ = PENEPMA_Root$ & "\Fanal\matrix"
If Dir$(tfolder$, vbDirectory) = vbNullString Then MkDir tfolder$

' Check for pure element PAR file in Penfluor\Pure folder
pfilename$ = Trim$(Symup$(ExtractElement%)) & ".par"
If Dir$(PENEPMA_Root$ & "\Penfluor\" & pfilename$) = vbNullString Then
tfilename$ = PENEPMA_Root$ & "\Penfluor\Pure\" & pfilename$
If Dir$(tfilename$) <> vbNullString Then FileCopy tfilename$, PENEPMA_Root$ & "\Penfluor\" & pfilename$
If Dir$(MiscGetFileNameNoExtension$(tfilename$) & ".in") <> vbNullString Then FileCopy MiscGetFileNameNoExtension$(tfilename$) & ".in", PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(pfilename$)) & ".in"
End If

' Check which first binary PAR file is available (e.g., Fe-Ni_99-1 or Ni-Fe_1-99 as either will work)
pfilename = Trim$(Symup$(ExtractElement%)) & "-" & Trim$(Symup$(ExtractMatrix%)) & "_" & Format$(BinaryRanges!(1)) & "-" & Format$(100# - BinaryRanges!(1)) & ".par"

' Check in Binary folder if not found in Penfluor
'If Dir$(PENEPMA_Root$ & "\Penfluor\" & pfilename$) = vbNullString Then
'tfilename$ = PENEPMA_Root$ & "\Penfluor\Binary\" & pfilename$    ' A and B are the same
'If Dir$(tfilename$) <> vbNullString Then FileCopy tfilename$, PENEPMA_Root$ & "\Penfluor\" & pfilename$
'If Dir$(MiscGetFileNameNoExtension$(tfilename$) & ".in") <> vbNullString Then FileCopy MiscGetFileNameNoExtension$(tfilename$) & ".in", PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(pfilename$)) & ".in"
'End If

If Dir$(PENEPMA_Root$ & "\Penfluor\" & pfilename$) <> vbNullString Then
BinaryElementsSwapped = False

' Try swapping the elements and concentrations (e.g., Fe-Ni_99-1 or Ni-Fe_1-99 as either will work)
Else
pfilename = Trim$(Symup$(ExtractMatrix%)) & "-" & Trim$(Symup$(ExtractElement%)) & "_" & Format$(BinaryRanges!(MAXBINARY%)) & "-" & Format$(100# - BinaryRanges!(MAXBINARY%)) & ".par"

' Check in Binary folder if not found in Penfluor
'If Dir$(PENEPMA_Root$ & "\Penfluor\" & pfilename$) = vbNullString Then
'tfilename$ = PENEPMA_Root$ & "\Penfluor\Binary\" & pfilename$    ' A and B are the same
'If Dir$(tfilename$) <> vbNullString Then FileCopy tfilename$, PENEPMA_Root$ & "\Penfluor\" & pfilename$
'If Dir$(MiscGetFileNameNoExtension$(tfilename$) & ".in") <> vbNullString Then FileCopy MiscGetFileNameNoExtension$(tfilename$) & ".in", PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(pfilename$)) & ".in"
'End If

If Dir$(PENEPMA_Root$ & "\Penfluor\" & pfilename$) <> vbNullString Then
BinaryElementsSwapped = True
Else
' If neither found then let it check below, as probably no valid x-ray lines exist for this binary anyway
End If
End If

' Load binary loop depending on filenames present
If Not BinaryElementsSwapped Then
inum1% = 1
inum2% = MAXBINARY%
inum3% = 1  ' loop step polarity
Else
inum1% = MAXBINARY%
inum2% = 1
inum3% = -1  ' loop step polarity
End If

' Check that .PAR files already exist (if not then just skip extraction)
For k% = inum1% To inum2% Step inum3%
DoEvents

' Load element composition based on binary number (always 99 to 1 wt%)
PENEPMA_Sample(1).ElmPercents!(1) = BinaryRanges!(k%)
PENEPMA_Sample(1).ElmPercents!(2) = 100# - BinaryRanges!(k%)

' Load name and number for this binary (swap element symbols for PAR file if necessary as either will do)
If Not BinaryElementsSwapped Then
binarynames$(k%) = Trim$(Symup$(ExtractElement%)) & "-" & Trim$(Symup$(ExtractMatrix%)) & "_" & Format$(PENEPMA_Sample(1).ElmPercents!(1)) & "-" & Format$(PENEPMA_Sample(1).ElmPercents!(2))
Else
binarynames$(k%) = Trim$(Symup$(ExtractMatrix%)) & "-" & Trim$(Symup$(ExtractElement%)) & "_" & Format$(PENEPMA_Sample(1).ElmPercents!(1)) & "-" & Format$(PENEPMA_Sample(1).ElmPercents!(2))
End If
PENEPMA_Sample(1).Name$ = binarynames$(k%)

' Load PAR file for this pure element for both material A and material B
ParameterFileA$ = PENEPMA_Sample(1).Name$ & ".par"
ParameterFileB$ = PENEPMA_Sample(1).Name$ & ".par"
ParameterFileBStd$ = Trim$(Symup$(ExtractElement%)) & ".par"

' Check .PAR files for matrix k-ratio extraction
pfilename1$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileA$    ' A and B are the same
pfilename2$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileB$    ' A and B are the same
pfilename3$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileBStd$

' Check in Binary folder if not found in Penfluor
'If Dir$(PENEPMA_Root$ & "\Penfluor\" & ParameterFileA$) = vbNullString Then
'tfilename$ = PENEPMA_Root$ & "\Penfluor\Binary\" & ParameterFileA$    ' A and B are the same for matrix calculations
'If Dir$(tfilename$) <> vbNullString Then FileCopy tfilename$, PENEPMA_Root$ & "\Penfluor\" & ParameterFileA$
'If Dir$(MiscGetFileNameNoExtension$(tfilename$) & ".in") <> vbNullString Then FileCopy MiscGetFileNameNoExtension$(tfilename$) & ".in", PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(ParameterFileA$)) & ".in"
'End If

' Check if binary composition is present
If Dir$(pfilename1$) = vbNullString Or Dir$(pfilename3$) = vbNullString Then
msg$ = "Parameter files " & ParameterFileA$ & " or " & ParameterFileBStd$ & " are not found and the k-ratio extraction will be skipped..."
Call IOWriteLog(msg$)
Exit Sub
End If

DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(3), CalculationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If
Next k%

' Load measured element, but for matrix fluorescence calculations just use default distance
MaterialMeasuredElement% = ExtractElement%

' Check for Li, Be, B, C, N, O, F or Ne and adjust minimum energy if so
If FormPENEPMA12.CheckAutoAdjustMinimumEnergy.Value = vbChecked Then
Call Penepma12AdjustMinimumEnergy2(Symlo$(MaterialMeasuredElement%))
End If

' Check for existing .TXT file
tfilename$ = Format$(ExtractElement%) & "-" & Format$(ExtractMatrix%) & "_" & Format$(MaterialMeasuredTakeoff#) & ".txt"
If Not CalculateDoNotOverwriteTXT Or (CalculateDoNotOverwriteTXT And Dir$(PENEPMA_Root$ & "\Fanal\matrix\" & tfilename$) = vbNullString) Then

' Write column labels
Call Penepma12CalculateReadWriteBinaryDataMatrix(Int(0), tfolder$, tfilename$, CSng(0))
If ierror Then Exit Sub

' Reset non-zero intensities output flag
FanalIntensitiesOutput = False

' Loop on each beam voltage from 1 to 50 keV
For m% = 1 To 50
'For m% = 50 To 50       ' testing purposes (50 keV only)
'For m% = 29 To 30       ' testing purposes (49 and 30 keV only)
'For m% = 15 To 15       ' testing purposes (15 keV only)
'For m% = 15 To 16       ' testing purposes (15 and 16 keV only)
'For m% = 19 To 19       ' testing purposes (19 keV only)
'For m% = 28 To 28       ' testing purposes (28 keV only for In ka in Na)
MaterialMeasuredEnergy# = m%

msg$ = vbCrLf & "Extracting Matrix K-Ratios for " & Trim$(Symup$(ExtractElement%)) & " in " & Trim$(Symup$(ExtractMatrix%)) & " at " & Format$(MaterialMeasuredEnergy#) & " keV..."
Call IOWriteLog(msg$)

' Init the globals
Call InitKratios
If ierror Then Exit Sub

' Loop on each valid x-ray line (Ka, Kb, La, Lb, Ma, Mb only at this time)
For l% = 1 To MAXRAY_OLD%
'For l% = 1 To 1         ' testing purposes (Ka only)
'For l% = 2 To 2         ' testing purposes (Kb only)
Call XrayGetEnergy(MaterialMeasuredElement%, l%, eng!, edg!)
If ierror Then Exit Sub

' Load minimum overvoltage percent, 0 = 2%, 1 = 10%, 2 = 20%, 3 = 40%
If MinimumOverVoltageType% = 0 Then tovervoltage! = MINIMUMOVERVOLTFRACTION_02!
If MinimumOverVoltageType% = 1 Then tovervoltage! = MINIMUMOVERVOLTFRACTION_10!
If MinimumOverVoltageType% = 2 Then tovervoltage! = MINIMUMOVERVOLTFRACTION_20!
If MinimumOverVoltageType% = 3 Then tovervoltage! = MINIMUMOVERVOLTFRACTION_40!

' Check for valid x-ray line (excitation energy (plus a buffer to avoid ultra low overvoltage issues) must be less than beam energy) (and greater than PenepmaMinimumElectronEnergy!)
If eng! <> 0# And edg! <> 0# And (edg! * (1# + tovervoltage!) < MaterialMeasuredEnergy#) And edg! > PenepmaMinimumElectronEnergy! Then

' Double check that specific transition exists
Call PenepmaGetPDATCONFTransition(MaterialMeasuredElement%, l%, t1!, t2!)
If ierror Then Exit Sub

' If both shells have ionization energies, it is ok to calculate
If t1! <> 0# And t2! <> 0# Then

' Load measured x-ray line
MaterialMeasuredXray% = l%

PENEPMA_Sample(1).LastElm% = 1
PENEPMA_Sample(1).LastChan% = 2
PENEPMA_Sample(1).Xrsyms$(1) = MaterialMeasuredXray%
PENEPMA_Sample(1).Xrsyms$(2) = Xraylo$(MAXRAY%)                ' make matrix element absorber only
Deflin$(ExtractElement%) = Xraylo$(MaterialMeasuredXray%)      ' save default xray for Penepma12GetParFileComposition

msg$ = vbCrLf & "Extracting Matrix Fluorescence K-Ratios for " & Trim$(Symup$(MaterialMeasuredElement%)) & " " & Xraylo$(MaterialMeasuredXray%) & "..."
Call IOWriteLog(msg$)

' Load each set of binary compositional parameter files
For k% = inum1% To inum2% Step inum3%

' Load element composition based on binary number (always 99 to 1 wt%)
PENEPMA_Sample(1).ElmPercents!(1) = BinaryRanges!(k%)
PENEPMA_Sample(1).ElmPercents!(2) = 100# - BinaryRanges!(k%)

' Load name and number for this binary (swap element symbols for PAR file if necessary as either will do)
If Not BinaryElementsSwapped Then
binarynames$(k%) = Trim$(Symup$(ExtractElement%)) & "-" & Trim$(Symup$(ExtractMatrix%)) & "_" & Format$(PENEPMA_Sample(1).ElmPercents!(1)) & "-" & Format$(PENEPMA_Sample(1).ElmPercents!(2))
Else
binarynames$(k%) = Trim$(Symup$(ExtractMatrix%)) & "-" & Trim$(Symup$(ExtractElement%)) & "_" & Format$(PENEPMA_Sample(1).ElmPercents!(1)) & "-" & Format$(PENEPMA_Sample(1).ElmPercents!(2))
End If
PENEPMA_Sample(1).Name$ = binarynames$(k%)

' For matrix correction calculation, parameter files A and B are the same
ParameterFileA$ = PENEPMA_Sample(1).Name$ & ".par"
ParameterFileB$ = PENEPMA_Sample(1).Name$ & ".par"
ParameterFileBStd$ = Trim$(Symup$(ExtractElement%)) & ".par"

' Double check that PAR file is in db folder (check penfluor folder in case manually copied)
If Dir$(PENEPMA_Root$ & "\Fanal\db\" & ParameterFileA$) = vbNullString Then
If Dir$(PENEPMA_Root$ & "\Penfluor\" & ParameterFileA$) <> vbNullString Then
FileCopy PENEPMA_Root$ & "\Penfluor\" & ParameterFileA$, PENEPMA_Root$ & "\Fanal\db\" & ParameterFileA$
If Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileA$) & ".in") <> vbNullString Then FileCopy PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileA$) & ".in", PENEPMA_Root$ & "\Fanal\db\" & MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(ParameterFileA$)) & ".in"
Else
GoTo Penepma12ExtractMatrixPARFilesNotFound:
End If
End If

' Calculate CalcZAF matrix corrections for this material (use Material A only) (loads MaterialMeasuredTakeoff# and MaterialMeasuredEnergy# for Penepma12CalculateMatrix)
Sleep 200   ' to make sure above file copy completes
Call Penepma12GetParFileComposition(Int(1), PENEPMA_Root$ & "\Fanal\db\" & ParameterFileA$, PENEPMA_SampleA())
If ierror Then Exit Sub

' Make matrix element absorber only for CalcZAF calculations
ip% = IPOS1(PENEPMA_SampleA(1).LastChan%, Symlo$(ExtractMatrix%), PENEPMA_SampleA(1).Elsyms$())
If ip% > 0 Then
PENEPMA_SampleA(1).Xrsyms$(ip%) = Xraylo$(MAXRAY%)
Call GetElmSaveSampleOnly(Int(0), PENEPMA_SampleA(), Int(0), Int(0))
If ierror Then Exit Sub
End If

' Calculate CalcZAF matrix effects for "measured concentrations" calculations
Call Penepma12CalculateMatrix(PENEPMA_Analysis, PENEPMA_SampleA(), Penepma_TmpSample())
If ierror Then Exit Sub

' Load CalcZAF matrix k-ratios for this material
ipA% = IPOS1(PENEPMA_SampleA(1).LastChan%, Symlo$(MaterialMeasuredElement%), PENEPMA_SampleA(1).Elsyms$())
If ipA% = 0 Then GoTo Penepma12ExtractMatrixNotFound

' Load CalcZAF output values
If Not BinaryElementsSwapped Then
CalcZAF_ZAF_Kratios#(l%, k%) = 100# * PENEPMA_SampleA(1).ElmPercents!(ipA%) / 100# / PENEPMA_Analysis.StdAssignsZAFCors!(3, ipA%) * 1# / PENEPMA_Analysis.StdAssignsZAFCors!(2, ipA%) * (1# / PENEPMA_Analysis.StdAssignsZAFCors!(7, ipA%)) / (1# / PENEPMA_Analysis.StdAssignsZAFCors!(8, ipA%))
CalcZAF_ZA_Kratios#(l%, k%) = 100# * PENEPMA_SampleA(1).ElmPercents!(ipA%) / 100# / PENEPMA_Analysis.StdAssignsZAFCors!(3, ipA%) * (1# / PENEPMA_Analysis.StdAssignsZAFCors!(7, ipA%)) / (1# / PENEPMA_Analysis.StdAssignsZAFCors!(8, ipA%))
CalcZAF_F_Kratios#(l%, k%) = 100# * PENEPMA_SampleA(1).ElmPercents!(ipA%) / 100# * 1# / PENEPMA_Analysis.StdAssignsZAFCors!(2, ipA%)
Else
CalcZAF_ZAF_Kratios#(l%, MAXBINARY% - (k% - 1)) = 100# * PENEPMA_SampleA(1).ElmPercents!(ipA%) / 100# / PENEPMA_Analysis.StdAssignsZAFCors!(3, ipA%) * 1# / PENEPMA_Analysis.StdAssignsZAFCors!(2, ipA%) * (1# / PENEPMA_Analysis.StdAssignsZAFCors!(7, ipA%)) / (1# / PENEPMA_Analysis.StdAssignsZAFCors!(8, ipA%))
CalcZAF_ZA_Kratios#(l%, MAXBINARY% - (k% - 1)) = 100# * PENEPMA_SampleA(1).ElmPercents!(ipA%) / 100# / PENEPMA_Analysis.StdAssignsZAFCors!(3, ipA%) * (1# / PENEPMA_Analysis.StdAssignsZAFCors!(7, ipA%)) / (1# / PENEPMA_Analysis.StdAssignsZAFCors!(8, ipA%))
CalcZAF_F_Kratios#(l%, MAXBINARY% - (k% - 1)) = 100# * PENEPMA_SampleA(1).ElmPercents!(ipA%) / 100# * 1# / PENEPMA_Analysis.StdAssignsZAFCors!(2, ipA%)
End If

Call IOStatusAuto("Extracting binary matrix k-ratios based on " & PENEPMA_Sample(1).Name$ & "...")
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(3), CalculationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

FormPENEPMA12.LabelProgress.Caption = "Extracting binary matrix k-ratios from " & PENEPMA_Sample(1).Name$ & ".par"
FormPENEPMA12.LabelRemainingTime.Caption = vbNullString

' Check for .IN file and if found check MSIMPA parameters (minimum electron/photon energy)
If Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileA$) & ".in") <> vbNullString Then
Call Penepma12RunFanalCheckINFile("MSIMPA", PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileA$) & ".in", pstring$)
If ierror Then Exit Sub
pvalue! = Val(pstring$)
pvalue! = pvalue! / EVPERKEV#

' If necessary skip this beam energy (empty file will deleted below automatically)
If edg! < pvalue! Then
msg$ = PENEPMA_Sample(1).Name$ & ".par" & " was only calculated down to " & Format$(pvalue!) & "keV. Skipping k-ratio extraction for " & Trim$(Symup$(MaterialMeasuredElement%)) & " " & Xraylo$(MaterialMeasuredXray%) & " in " & Symup$(ExtractMatrix%) & "..."
Call IOWriteLog(msg$)
GoTo Penepma12ExtractMatrixSkip
End If
End If

' Check the parameters files
Call Penepma12RunFanal
If ierror Then Exit Sub

' Run the Fanal program
Call Penepma12RunFanal1
If ierror Then Exit Sub

' Get k-ratio data from k-ratio file
Call Penepma12LoadPlotData
If ierror Then Exit Sub

' Check that Fanal produced good data
If nPoints& > 0 Then

' Debug
If DebugMode Then
msg$ = "K-ratio%= " & yktotal#(1) & " for " & Trim$(Symup$(ExtractElement%)) & " " & Xraylo$(MaterialMeasuredXray%) & " in " & ParameterFileA$ & " using standard " & ParameterFileBStd$
Call IOWriteLog(msg$)
End If

' Store essential fluorescent k-ratio data to data array (only need to store first or last data point for matrix calculations)
unk_int_pri# = pri_int#(nPoints&)                                                           ' calculate Mat A primary intensity
unk_int_flu# = flach#(nPoints&) + flabr#(nPoints&) + flbch#(nPoints&) + flbbr#(nPoints&)    ' calculate Mat A and Mat B fluorescence intensity
unk_int_all# = unk_int_flu# + pri_int#(nPoints&)                                            ' calculate total intensity

' Skip if primary intensity not calculated
If pri_int#(nPoints&) > 0# Then
tempF# = 1# / (1# + unk_int_flu# / unk_int_all#)     ' calculate fluorescence only
tempZA# = 1# / (unk_int_pri# / unk_int_all#)         ' calculate ZA correction only
End If

' Check for valid std intensity
If std_int#(nPoints&) <= 0# Then GoTo Penepma12ExtractMatrixZeroStdInt

If Not BinaryElementsSwapped Then
Binary_ZAF_Kratios#(l%, k%) = 100# * unk_int_all# / std_int#(nPoints&)       ' calculate total k-ratio

' Skip if primary intensity not calculated
If pri_int#(nPoints&) > 0# Then
Binary_F_Kratios#(l%, k%) = 100# * PENEPMA_SampleA(1).ElmPercents!(ipA%) / 100# * 1# / tempF#
Binary_ZA_Kratios#(l%, k%) = 100# * PENEPMA_SampleA(1).ElmPercents!(ipA%) / 100# / tempZA#
End If

Else
Binary_ZAF_Kratios#(l%, MAXBINARY% - (k% - 1)) = 100# * unk_int_all# / std_int#(nPoints&)       ' calculate total k-ratio

' Skip if primary intensity not calculated
If pri_int#(nPoints&) > 0# Then
Binary_F_Kratios#(l%, MAXBINARY% - (k% - 1)) = 100# * PENEPMA_SampleA(1).ElmPercents!(ipA%) / 100# * 1# / tempF#
Binary_ZA_Kratios#(l%, MAXBINARY% - (k% - 1)) = 100# * PENEPMA_SampleA(1).ElmPercents!(ipA%) / 100# / tempZA#
End If
End If

DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(3), CalculationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

' Check for Pause button
Do Until Not RealTimePauseAutomation
DoEvents
Sleep 200
Loop

' If we get to here, non-zero intensities were calculated. Set flag to not erase output file
FanalIntensitiesOutput = True

' Nothing to output
Else
msg$ = "No intensity data to output for " & Trim$(Symup$(ExtractElement%)) & " " & Xraylo$(MaterialMeasuredXray%) & " in " & ParameterFileA$ & " using standard " & ParameterFileBStd$
Call IOWriteLog(msg$)
End If
Next k%

msg$ = "All " & Format$(MAXBINARY%) & " matrix k-ratio extractions are complete for " & Trim$(Symup$(MaterialMeasuredElement%)) & " " & Xraylo$(MaterialMeasuredXray%) & "..."
Call IOWriteLog(msg$)
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(3), CalculationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

End If

' Overvoltage too low (zero arrays)
Else
If Not BinaryElementsSwapped Then
For k% = 1 To MAXBINARY%
Binary_ZAF_Kratios#(l%, k%) = 0#
Binary_F_Kratios#(l%, k%) = 0#
Binary_ZA_Kratios#(l%, k%) = 0#
Next k%
Else
For k% = MAXBINARY% To 1 Step -1
Binary_ZAF_Kratios#(l%, MAXBINARY% - (k% - 1)) = 0#
Binary_F_Kratios#(l%, MAXBINARY% - (k% - 1)) = 0#
Binary_ZA_Kratios#(l%, MAXBINARY% - (k% - 1)) = 0#
Next k%
End If

End If
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(3), CalculationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If
Next l%

msg$ = "All x-ray line extractions are complete for " & Trim$(Symup$(MaterialMeasuredElement%)) & " in " & Trim$(Symup$(ExtractMatrix%)) & " at " & Format$(MaterialMeasuredEnergy#) & " keV"
Call IOWriteLog(msg$)

' Write binary k-ratio fluorescence data to file for the specified beam energy
tfilename$ = Format$(ExtractElement%) & "-" & Format$(ExtractMatrix%) & "_" & Format$(MaterialMeasuredTakeoff#) & ".txt"
Call Penepma12CalculateReadWriteBinaryDataMatrix(Int(2), tfolder$, tfilename$, CSng(MaterialMeasuredEnergy#))
If ierror Then Exit Sub

DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(3), CalculationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If
Next m%

' PAR file not calculated with sufficiently low minimum energy
Penepma12ExtractMatrixSkip:

' Check if non-zero intensities were actually output. If not, delete the TXT file
If Not FanalIntensitiesOutput Then
Kill tfolder$ & "\" & tfilename$
Exit Sub
End If

Else
msg$ = "Skipping k-ratio matrix extraction for " & tfilename$ & "..."
Call IOWriteLog(msg$)
End If

Call IOStatusAuto(vbNullString)
msg$ = "All matrix fluorescence k-ratio extractions are complete"
Call IOWriteLog(msg$)
DoEvents

Exit Sub

' Errors
Penepma12ExtractMatrixError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12ExtractMatrix"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12ExtractMatrixSame:
msg$ = "The extract and matrix elements (" & Trim$(Symup$(ExtractElement%)) & " and " & Trim$(Symup$(ExtractMatrix%)) & ") are the same, but must be different for calculating matrix effects"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12ExtractMatrix"
ierror = True
Exit Sub

Penepma12ExtractMatrixMore:
msg$ = "The emitting element (" & Trim$(Symup$(ExtractElement%)) & ") must precede specified matrix element (" & Trim$(Symup$(ExtractMatrix%)) & ") in the Periodic Table (don't ask why!)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12ExtractMatrix"
ierror = True
Exit Sub

Penepma12ExtractMatrixNotFound:
msg$ = "The measured element (" & Trim$(Symup$(MaterialMeasuredElement%)) & ") was not found in the Penepma (CalcZAF) sample"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12ExtractMatrix"
ierror = True
Exit Sub

Penepma12ExtractMatrixPARFilesNotFound:
msg$ = "The specified .PAR file (" & ParameterFileA$ & ") was not found in the Fanal\db or Penfluor folders. Please calculate the specified .PAR parameter file and try again"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12ExtractMatrix"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12ExtractMatrixZeroStdInt:
msg$ = "The standard intensity for the measured element " & Symlo$(MaterialMeasuredElement%) & " " & Xraylo$(MaterialMeasuredXray%) & " was zero for the material B Std composition (" & ParameterFileBStd$ & ") at " & Format$(MaterialMeasuredEnergy#) & " keV. This error should not occur, please contact Probe Software with details (and check the Fanal\k-ratios.dat file)."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12ExtractMatrix"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12ExtractRandom()
' Extract boundary or matrix k-ratios for the periodic table. The boundary or matrix extraction is selected randomly
' and existing output files are skipped if they are being extracted, based on a shared look up table.

ierror = False
On Error GoTo Penepma12ExtractRandomError

Dim done As Boolean
Dim i As Integer, j As Integer, m As Integer
Dim im As Integer, mm As Integer
Dim response As Integer
Dim tExtractElement As Integer, tExtractMatrix As Integer
Dim tMaterialMeasuredGridPoints As Integer

icancelauto = False

' Extracting entire matrix range
ExtractForSpecifiedRange = True  ' to skip
icancelauto = False

' ExtractMethod = 0  Extract boundary k-ratios over the entire periodic table
If ExtractMethod% = 0 Then

msg$ = "This boundary k-ratio extraction is designed to be performed by executing many multiple applications running in parallel utilizing a shared network folder for the Penepma12_PAR_Path to facilitate calculation of all binaries for the entire periodic table. The total calculation time will be approximately 50 years divided by the number of parallel applications running simultaneously (100 applications running in parallel will take approximately 6 months)! Are you sure you want to proceed?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "Penepma12ExtractRandom")
If response% = vbCancel Then Exit Sub

' Calculate number of boundaries for entire periodic table
m% = 0
ReDim ExtractRandomTable(1 To 3, 1 To 1) As Integer
For i% = 1 To MAXELM%
For j% = i% To MAXELM%
If i% <> j% Then
m% = m% + 1
ReDim Preserve ExtractRandomTable(1 To 3, 1 To m%) As Integer
ExtractRandomTable%(1, m%) = i%   ' ExtractElement1
ExtractRandomTable%(2, m%) = j%   ' ExtractElement2
ExtractRandomTable%(3, m%) = m%   ' Extract number (1 to m%)
End If
Next j%
Next i%

' Try to create a new PAR share file
'Call Penepma12ExtractRandomCheck(Int(0), Int(3), m%, ExtractRandomTable%(), im%, mm%, done)
'If ierror Then Exit Sub

tExtractElement% = ExtractElement%      ' save
tExtractMatrix% = ExtractMatrix%      ' save

' Check if randomly selected extraction is being calculated already (im is selected extraction, mm is extractions calculated so far)
Do Until done
'Call Penepma12ExtractRandomCheck(Int(1), Int(3), m%, ExtractRandomTable%(), im%, mm%, done)
'If ierror Then Exit Sub
If done Then Exit Do

' Load next calculation
i% = ExtractRandomTable%(1, im%)
j% = ExtractRandomTable%(2, im%)

msg$ = vbCrLf & vbCrLf & "Calculating boundary extract " & Format$(mm%) & " of " & Format$(m%) & ": " & Trim$(Symup$(i%)) & "-" & Trim$(Symup$(j%)) & "..."
Call IOWriteLog(msg$)
Call IOStatusAuto(msg$)
DoEvents

ExtractElement% = i%                    ' load matrix 1
ExtractMatrix% = j%                    ' load matrix 2
Call Penepma12ExtractBoundary
ExtractElement% = tExtractElement%      ' restore
ExtractMatrix% = tExtractMatrix%      ' restore

' Update complete status for this calculation
'Call Penepma12ExtractRandomCheck(Int(2), Int(3), m%, ExtractRandomTable%(), im%, mm%, done)
'If ierror Then Exit Sub

If ierror Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

Loop

msg$ = "All boundary k-ratio extraction calculations are complete"
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12ExtractRandom"
End If

' ExtractMethod = 1  Extract matrix k-ratios over the periodic table
If ExtractMethod% = 1 Then

msg$ = "This matrix k-ratio extraction is designed to be performed by executing many multiple applications running in parallel utilizing a shared network folder for the Penepma12_PAR_Path to facilitate calculation of all matrix binaries for the entire periodic table. The total calculation time will be approximately several months divided by the number of parallel applications running simultaneously (10 applications running in parallel will take weeks)! Are you sure you want to proceed?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "Penepma12ExtractRandom")
If response% = vbCancel Then Exit Sub

m% = 0
ReDim ExtractRandomTable(1 To 2, 1 To 1) As Integer
For i% = 1 To MAXELM%
m% = m% + 1
ReDim Preserve ExtractRandomTable(1 To 2, 1 To m%) As Integer
ExtractRandomTable%(1, m%) = i%   ' ExtractElement1 and ExtractElement2
ExtractRandomTable%(2, m%) = m%   ' Sequence number
Next i%

' Try to create a new PAR share file
'Call Penepma12ExtractRandomCheck(Int(0), Int(2), m%, ExtractRandomTable%(), im%, mm%, done)
'If ierror Then Exit Sub

tExtractElement% = ExtractElement%      ' save
tExtractMatrix% = ExtractMatrix%      ' save

' Check if randomly selected pure element is being calculated already (im is selected element, mm is elements calculated so far)
Do Until done
'Call Penepma12ExtractRandomCheck(Int(1), Int(2), m%, ExtractRandomTable%(), im%, mm%, done)
'If ierror Then Exit Sub
If done Then Exit Do

' Load next calculation
i% = ExtractRandomTable%(1, im%)

msg$ = vbCrLf & vbCrLf & "Calculating matrix extract " & Format$(mm%) & " of " & Format$(m%) & ": " & Trim$(Symup$(i%)) & "..."
Call IOWriteLog(msg$)
Call IOStatusAuto(msg$)
DoEvents

ExtractElement% = i%
ExtractMatrix% = i%
tMaterialMeasuredGridPoints% = MaterialMeasuredGridPoints%      ' save
MaterialMeasuredGridPoints% = 1     ' use a single point for matrix calculations
Call Penepma12ExtractMatrix
MaterialMeasuredGridPoints% = tMaterialMeasuredGridPoints%      ' restore
ExtractElement% = tExtractElement%      ' restore
ExtractMatrix% = tExtractMatrix%      ' restore

' Update complete status for this calculation
'Call Penepma12ExtractRandomCheck(Int(2), Int(2), m%, ExtractRandomTable%(), im%, mm%, done)
'If ierror Then Exit Sub

If ierror Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

Loop

msg$ = "All matrix k-ratio extraction file calculations are complete"
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12ExtractRandom"
End If

Call IOStatusAuto(vbNullString)
Exit Sub

' Errors
Penepma12ExtractRandomError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12ExtractRandom"
ierror = True
Exit Sub

End Sub

Sub Penepma12Random()
' Load the Penepma12 Calculate/Extract form (used to create multiple instances of PenPFE.exe)

ierror = False
On Error GoTo Penepma12RandomError

' Load the form
Call Penepma12RandomLoad
If ierror Then Exit Sub

' Show the form
FormPenepma12Random.Show vbModeless

Exit Sub

' Errors
Penepma12RandomError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12Random"
ierror = True
Exit Sub

End Sub

Sub Penepma12OutputPlotData()
' This routine outputs data for plotting the k-ratios or alpha factors for a specified keV, element and x-ray

ierror = False
On Error GoTo Penepma12OutputPlotDataError

ReDim CalcZAF_ZAF_Factors(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Single
ReDim CalcZAF_ZA_Factors(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Single
ReDim CalcZAF_F_Factors(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Single

ReDim Binary_ZAF_Factors(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Single
ReDim Binary_ZA_Factors(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Single
ReDim Binary_F_Factors(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Single

ReDim Binary_ZAF_Coeffs(1 To MAXRAY%, 1 To MAXCOEFF4%) As Single
ReDim CalcZAF_ZAF_Coeffs(1 To MAXRAY%, 1 To MAXCOEFF4%) As Single

ReDim Binary_ZA_Coeffs(1 To MAXRAY%, 1 To MAXCOEFF4%) As Single
ReDim CalcZAF_ZA_Coeffs(1 To MAXRAY%, 1 To MAXCOEFF4%) As Single

ReDim Binary_F_Coeffs(1 To MAXRAY%, 1 To MAXCOEFF4%) As Single
ReDim CalcZAF_F_Coeffs(1 To MAXRAY%, 1 To MAXCOEFF4%) As Single

ReDim Binary_ZAF_Betas(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Single
ReDim CalcZAF_ZAF_Betas(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Single

ReDim Binary_ZAF_Devs(1 To MAXRAY%) As Single
ReDim CalcZAF_ZAF_Devs(1 To MAXRAY%) As Single

ReDim Binary_ZA_Devs(1 To MAXRAY%) As Single
ReDim CalcZAF_ZA_Devs(1 To MAXRAY%) As Single

ReDim Binary_F_Devs(1 To MAXRAY%) As Single
ReDim CalcZAF_F_Devs(1 To MAXRAY%) As Single

' ExtractMethod = 0  Extract k-ratios for boundary fluorescence
If ExtractMethod% = 0 Then
Call Penepma12OutputPlotDataBoundary
If ierror Then Exit Sub
End If

' ExtractMethod = 1  Extract k-ratios for matrix corrections
If ExtractMethod% = 1 Then
Call Penepma12OutputPlotDataMatrix
If ierror Then Exit Sub
End If

Exit Sub

' Errors
Penepma12OutputPlotDataError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12OutputPlotData"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma12OutputPlotDataBoundary()
' This routine outputs data for plotting the k-ratios or alpha factors for a specified keV, element and x-ray
'  from the output from Penepma12ExtractBoundary.
'  Boundary_ZAF_Kratios#(1 to MAXBINARY%, 1 to MAXBINARY%, 1 to npoints&)  are the k-ratios from Fanal in k-ratio % for each x distance
'  Boundary_ZAF_Factors!(1 to MAXBINARY%, 1 to MAXBINARY%, 1 to npoints&)  are the alpha factors for each boundary composition, alpha = (C/K - C)/(1 - C)
'
'  Boundary_Linear_Distances(1 to npoints&)                     ' linear distances (not used here)
'  Boundary_Mass_Distances(1 to MAXBINARY%, 1 to npoints&)      ' mass distances (not used here)
'
'  Boundary_Material_A_Densities(1 to MAXBINARY%)  ' material A densities (not used here)
'  Boundary_Material_B_Densities(1 to MAXBINARY%)  ' material B densities (not used here)

ierror = False
On Error GoTo Penepma12OutputPlotDataBoundaryError

Dim j As Integer, k As Integer
Dim n As Long
Dim tfilename As String, tfolder As String
Dim astring As String, jstring As String, kstring As String
Dim eng As Single, edg As Single

Dim bulkcalculationindex As Integer
Dim tfilename2 As String, tfolder2 As String
Dim relativepercent As Single

' Specify folder
tfolder$ = PENEPMA_Root$ & "\Fanal\boundary"

' Specify extract element (x-ray is based on MeasuredMaterialXray%, beam energy is based on MeasuredMaterialEnergy)
MaterialMeasuredElement% = ExtractElement%

' Load input filename based on specified materials
jstring$ = Format$(ExtractMatrixA1%) & "-" & Format$(ExtractMatrixA2%)
kstring$ = Format$(ExtractMatrixB1%) & "-" & Format$(ExtractMatrixB2%)
tfilename$ = jstring$ & "_" & kstring$ & "_" & Format$(MaterialMeasuredTakeoff#) & "_" & Format$(MaterialMeasuredElement%) & "-" & Format$(MaterialMeasuredXray%) & ".txt"
If Dir$(tfolder$ & "\" & tfilename$) <> vbNullString Then
BinaryElementsSwappedA = False
BinaryElementsSwappedB = False
GoTo Penepma12OutputPlotDataBoundaryProceed
End If

jstring$ = Format$(ExtractMatrixA1%) & "-" & Format$(ExtractMatrixA2%)
kstring$ = Format$(ExtractMatrixB2%) & "-" & Format$(ExtractMatrixB1)
tfilename$ = jstring$ & "_" & kstring$ & "_" & Format$(MaterialMeasuredTakeoff#) & "_" & Format$(MaterialMeasuredElement%) & "-" & Format$(MaterialMeasuredXray%) & ".txt"
If Dir$(tfolder$ & "\" & tfilename$) <> vbNullString Then
BinaryElementsSwappedA = False
BinaryElementsSwappedB = True
GoTo Penepma12OutputPlotDataBoundaryProceed
End If

jstring$ = Format$(ExtractMatrixA2%) & "-" & Format$(ExtractMatrixA1%)
kstring$ = Format$(ExtractMatrixB1%) & "-" & Format$(ExtractMatrixB2)
tfilename$ = jstring$ & "_" & kstring$ & "_" & Format$(MaterialMeasuredTakeoff#) & "_" & Format$(MaterialMeasuredElement%) & "-" & Format$(MaterialMeasuredXray%) & ".txt"
If Dir$(tfolder$ & "\" & tfilename$) <> vbNullString Then
BinaryElementsSwappedA = True
BinaryElementsSwappedB = False
GoTo Penepma12OutputPlotDataBoundaryProceed
End If

jstring$ = Format$(ExtractMatrixA2%) & "-" & Format$(ExtractMatrixA1%)
kstring$ = Format$(ExtractMatrixB2%) & "-" & Format$(ExtractMatrixB1)
tfilename$ = jstring$ & "_" & kstring$ & "_" & Format$(MaterialMeasuredTakeoff#) & "_" & Format$(MaterialMeasuredElement%) & "-" & Format$(MaterialMeasuredXray%) & ".txt"
If Dir$(tfolder$ & "\" & tfilename$) <> vbNullString Then
BinaryElementsSwappedA = True
BinaryElementsSwappedB = True
GoTo Penepma12OutputPlotDataBoundaryProceed
End If

' If we get to here, no suitable extraction files were found
GoTo Penepma12OutputPlotDataBoundaryNoFile

' Check the specified x-ray
Penepma12OutputPlotDataBoundaryProceed:
Call XrayGetEnergy(MaterialMeasuredElement%, MaterialMeasuredXray%, eng!, edg!)
If ierror Then Exit Sub

' Check for valid x-ray line (excitation energy must be less than beam energy)
If eng! = 0# Or edg! = 0# Or edg! > MaterialMeasuredEnergy# Then GoTo Penepma12OutputPlotDataBoundaryNoXray

' Read the data from the boundary extract file
Call Penepma12CalculateReadWriteBinaryDataBoundary(Int(1), tfolder$, tfilename$, CSng(MaterialMeasuredEnergy#), nPoints&)
If ierror Then Exit Sub

' Check if measured element is one of the matrix elements in boundary calculation
bulkcalculationindex% = 0
If Not BinaryElementsSwappedA And Not BinaryElementsSwappedB Then
If MaterialMeasuredElement% = ExtractMatrixA1% Then bulkcalculationindex% = 1
If MaterialMeasuredElement% = ExtractMatrixB1% Then bulkcalculationindex% = 3
End If

' Read the bulk matrix k-ratios for material A using Binary_ZAF_Kratios#(MaterialMeasuredXray%, j%)
If bulkcalculationindex% > 0 Then
tfolder2$ = PENEPMA_Root$ & "\Fanal\matrix"
If bulkcalculationindex% = 1 Then
tfilename2$ = Format$(ExtractMatrixA1%) & "-" & Format$(ExtractMatrixA2%) & "_" & Format$(MaterialMeasuredTakeoff#) & ".txt"
If Dir$(tfolder2$ & "\" & tfilename2$) = vbNullString Then GoTo Penepma12OutputPlotDataBoundaryBulkNotFound
End If

If bulkcalculationindex% = 3 Then
tfilename2$ = Format$(ExtractMatrixB1%) & "-" & Format$(ExtractMatrixB2%) & "_" & Format$(MaterialMeasuredTakeoff#) & ".txt"
If Dir$(tfolder2$ & "\" & tfilename2$) = vbNullString Then GoTo Penepma12OutputPlotDataBoundaryBulkNotFound
End If

Call Penepma12CalculateReadWriteBinaryDataMatrix(Int(1), tfolder2$, tfilename2$, CSng(MaterialMeasuredEnergy#))
If ierror Then Exit Sub
End If

' Output file for each distance
For n& = 1 To nPoints&

' Output the data as desired (conc vs k-ratio/alphas etc.)
Close #Temp1FileNumber%
DoEvents

' Load output filename based on distance
tfilename$ = tfolder$ & "\" & Trim$(Symup$(ExtractMatrixA1%)) & "-" & Trim$(Symup$(ExtractMatrixA2%)) & "_" & Trim$(Symup$(ExtractMatrixB1%)) & "-" & Trim$(Symup$(ExtractMatrixB2%)) & "_" & Format$(MaterialMeasuredTakeoff#) & "_" & Format$(MaterialMeasuredEnergy#) & "_" & Trim$(Symup$(ExtractElement%)) & " " & Xraylo$(MaterialMeasuredXray%) & "_" & Format$(Abs(Boundary_Linear_Distances!(n&)), "Fixed") & "um" & ".dat"
Open tfilename$ For Output As #Temp1FileNumber%

' Output column labels
astring$ = VbDquote$ & "keV" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "ConcA1" & VbDquote$ & vbTab & VbDquote$ & "ConcA2" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "ConcB1" & VbDquote$ & vbTab & VbDquote$ & "ConcB2" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "KratioAB%" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "Alpha" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "KratioA%" & VbDquote$ & vbTab        ' bulk calculation
astring$ = astring$ & VbDquote$ & "KratioAB-A%" & VbDquote$ & vbTab     ' boundary minus bulk
astring$ = astring$ & VbDquote$ & "KratioAB-A% (relative)" & VbDquote$ & vbTab     ' boundary minus bulk relative error
astring$ = astring$ & VbDquote$ & "MatA Mass Distance" & VbDquote$ & vbTab      ' material A mass distances
astring$ = astring$ & VbDquote$ & "MatA Density" & VbDquote$ & vbTab            ' material A densities
astring$ = astring$ & VbDquote$ & "MatB Density" & VbDquote$ & vbTab            ' material B densities
'astring$ = astring$ & VbDquote$ & "_Gamma" & VbDquote$ & vbTab
Print #Temp1FileNumber%, astring

' Calculate the gamma factor for each binary composition for the specified binary and xray
'Call Penepma12CalculateBinaryGamma(MaterialMeasuredXray%, n&, BinaryRanges!(), Boundary_ZAF_Coeffs#(), Boundary_ZAF_Gammas!())
'If ierror Then Exit Sub

' Output data for this energy by concentration
For j% = 1 To MAXBINARY%    ' material A
For k% = 1 To MAXBINARY%    ' material B

If Not BinaryElementsSwappedA Then
jstring$ = Format$(BinaryRanges!(j%)) & vbTab & Format$(100# - BinaryRanges!(j%))
Else
jstring$ = Format$(BinaryRanges!(MAXBINARY - (j% - 1))) & vbTab & Format$(100# - BinaryRanges!(MAXBINARY - (j% - 1)))
End If

If Not BinaryElementsSwappedB Then
kstring$ = Format$(BinaryRanges!(k%)) & vbTab & Format$(100# - BinaryRanges!(k%))
Else
kstring$ = Format$(BinaryRanges!(MAXBINARY - (k% - 1))) & vbTab & Format$(100# - BinaryRanges!(MAXBINARY - (k% - 1)))
End If

' Calculate relative percent error from bulk k-ratio
If bulkcalculationindex% = 1 Then
If Binary_ZAF_Kratios#(MaterialMeasuredXray%, j%) <> 0# Then relativepercent! = 100# * (Boundary_ZAF_Kratios#(k%, j%, n&) - Binary_ZAF_Kratios#(MaterialMeasuredXray%, j%)) / Binary_ZAF_Kratios#(MaterialMeasuredXray%, j%)
ElseIf bulkcalculationindex% = 3 Then
If Binary_ZAF_Kratios#(MaterialMeasuredXray%, MAXBINARY - (j% - 1)) <> 0# Then relativepercent! = 100# * (Boundary_ZAF_Kratios#(k%, j%, n&) - Binary_ZAF_Kratios#(MaterialMeasuredXray%, MAXBINARY - (j% - 1))) / Binary_ZAF_Kratios#(MaterialMeasuredXray%, MAXBINARY - (j% - 1))
Else
relativepercent! = 0#
End If

' Output calculations
astring$ = CSng(MaterialMeasuredEnergy#) & vbTab & jstring$ & vbTab & kstring$ & vbTab      ' composition of A and B
astring$ = astring$ & Boundary_ZAF_Kratios#(k%, j%, n&) & vbTab     ' k-ratio of boundary extraction
astring$ = astring$ & Boundary_ZAF_Factors!(k%, j%, n&) & vbTab     ' alpha factor of boundary extraction (only if emitter is in material A)

' Load bulk k-ratio depending on availability of bulk extraction file (e.g., Si-Ti versus Ti-Si)
If bulkcalculationindex% = 1 Then
astring$ = astring$ & Binary_ZAF_Kratios#(MaterialMeasuredXray%, j%) & vbTab    ' k-ratio of bulk material A
astring$ = astring$ & (Boundary_ZAF_Kratios#(k%, j%, n&) - Binary_ZAF_Kratios#(MaterialMeasuredXray%, j%)) & vbTab  ' k-ratioAB-k-ratioA
ElseIf bulkcalculationindex% = 3 Then
astring$ = astring$ & Binary_ZAF_Kratios#(MaterialMeasuredXray%, MAXBINARY - (j% - 1)) & vbTab    ' k-ratio of bulk material A
astring$ = astring$ & (Boundary_ZAF_Kratios#(k%, j%, n&) - Binary_ZAF_Kratios#(MaterialMeasuredXray%, MAXBINARY - (j% - 1))) & vbTab  ' k-ratioAB-k-ratioA
Else
astring$ = astring$ & CSng(0#) & vbTab & CSng(0#) & vbTab
End If

' Add relative percent% difference between boundary and bulk
astring$ = astring$ & relativepercent! & vbTab

' Add mass distances and density of material A and material B (densities are the same for all distances)
astring$ = astring$ & Boundary_Mass_Distances(j%, n&) & vbTab
astring$ = astring$ & Boundary_Material_A_Densities(j%) & vbTab
astring$ = astring$ & Boundary_Material_B_Densities(k%) & vbTab

' Output string
Print #Temp1FileNumber%, astring$

Next k%
Next j%

Close #Temp1FileNumber%
Next n&

msg$ = "The specified boundary k-ratio plot data was output based on " & Format$(MaterialMeasuredTakeoff#) & " degrees, " & Format$(MaterialMeasuredEnergy#) & " keV, " & Trim$(Symup$(ExtractElement%)) & " " & Trim$(Xraylo$(MaterialMeasuredXray%)) & " in " & Trim$(Symup$(ExtractMatrixA1%)) & "-" & Trim$(Symup$(ExtractMatrixA2%)) & " adjacent to " & Trim$(Symup$(ExtractMatrixB1%)) & "-" & Trim$(Symup$(ExtractMatrixB2%))
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12OutputPlotDataBoundary"

Exit Sub

' Errors
Penepma12OutputPlotDataBoundaryError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12OutputPlotDataBoundary"
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12OutputPlotDataBoundaryNoXray:
msg$ = "The specified element and x-ray (" & Trim$(Symup$(MaterialMeasuredElement%)) & " " & Xraylo$(MaterialMeasuredXray%) & ") is not valid for the specified beam energy (" & Format$(MaterialMeasuredEnergy#) & " keV)"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12OutputPlotDataBoundary"
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12OutputPlotDataBoundaryNoFile:
msg$ = "The specified Penepma12ExtractBoundary file (" & tfilename$ & ") or other combination of " & Symup$(ExtractMatrixA1%) & ", " & Symup$(ExtractMatrixA2%) & " and " & Symup$(ExtractMatrixB1%) & ", " & Symup$(ExtractMatrixB2%) & " was not found. Please run the Extract Boundary calculations first for the specified element and matrix A and B"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12OutputPlotDataBoundary"
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12OutputPlotDataBoundaryBulkNotFound:
msg$ = "The specified Penepma12ExtractMatrix file (" & tfilename2$ & ") was not found. Please run the Extract Matrix calculations first for the specified emitting element (" & " and matrix elements"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12OutputPlotDataBoundary"
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12OutputPlotDataMatrix()
' This routine outputs data for plotting the k-ratios or alpha factors for a specified keV, element and x-ray
'  from the output from Penepma12ExtractMatrix.
'  Binary_ZAF_Factors!(1 to MAXRAY%, 1 to MAXBINARY%)  are the full alpha factors for each x-ray and binary composition, alpha = (C/K - C)/(1 - C)
'  CalcZAF_ZAF_Factors!(1 to MAXRAY%, 1 to MAXBINARY%)  are the full alpha factors for each x-ray and binary composition, alpha = (C/K - C)/(1 - C)
'  Binary_ZA_Kratios#(1 to MAXRAY%, 1 to MAXBINARY%)  are the ZA only k-ratios from Fanal in k-ratio % for each x-ray and binary composition
'  CalcZAF_ZA_Kratios#(1 to MAXRAY%, 1 to MAXBINARY%)  are the ZA only k-ratios from CalcZAF in k-ratio % for each x-ray and binary composition
'  Binary_F_Kratios#(1 to MAXRAY%, 1 to MAXBINARY%)  are the fluorescence only k-ratios from Fanal in k-ratio % for each x-ray and binary composition
'  CalcZAF_F_Kratios#(1 to MAXRAY%, 1 to MAXBINARY%)  are the fluorescence only k-ratios from CalcZAF in k-ratio % for each x-ray and binary composition

ierror = False
On Error GoTo Penepma12OutputPlotDataMatrixError

Dim n As Integer, l As Integer, npts As Integer
Dim tfilename As String, tfolder As String
Dim astring As String
Dim eng As Single, edg As Single

tfolder$ = PENEPMA_Root$ & "\Fanal\matrix"
tfilename$ = Format$(ExtractElement%) & "-" & Format$(ExtractMatrix%) & "_" & Format$(MaterialMeasuredTakeoff#) & ".txt"
If Dir$(tfolder$ & "\" & tfilename$) = vbNullString Then GoTo Penepma12OutputPlotDataMatrixNotFound

' Read the specified data from binary text file
Call Penepma12CalculateReadWriteBinaryDataMatrix(Int(1), tfolder$, tfilename$, CSng(MaterialMeasuredEnergy#))
If ierror Then Exit Sub

' Output for each valid x-ray
For l% = 1 To MAXRAY_OLD%
Call XrayGetEnergy(ExtractElement%, l%, eng!, edg!)
If ierror Then Exit Sub

' Check for valid x-ray line (excitation energy must be less than beam energy)
If eng! <> 0# And edg! <> 0# And edg! < MaterialMeasuredEnergy# Then

' Calculate Fanal ZAF matrix correction alpha factors for these binary compositions
Call Penepma12CalculateAlphaFactors(l%, BinaryRanges!(), Binary_ZAF_Kratios#(), Binary_ZAF_Factors!(), Binary_ZAF_Coeffs!(), Binary_ZAF_Devs!(), npts%)
If ierror Then Exit Sub

' Calculate CalcZAF ZAF matrix correction alpha factors for these binary compositions
Call Penepma12CalculateAlphaFactors(l%, BinaryRanges!(), CalcZAF_ZAF_Kratios#(), CalcZAF_ZAF_Factors!(), CalcZAF_ZAF_Coeffs!(), CalcZAF_ZAF_Devs!(), npts%)
If ierror Then Exit Sub

' Calculate Fanal ZA matrix correction alpha factors for these binary compositions
Call Penepma12CalculateAlphaFactors(l%, BinaryRanges!(), Binary_ZA_Kratios#(), Binary_ZA_Factors!(), Binary_ZA_Coeffs!(), Binary_ZA_Devs!(), npts%)
If ierror Then Exit Sub

' Calculate CalcZAF ZA matrix correction alpha factors for these binary compositions
Call Penepma12CalculateAlphaFactors(l%, BinaryRanges!(), CalcZAF_ZA_Kratios#(), CalcZAF_ZA_Factors!(), CalcZAF_ZA_Coeffs!(), CalcZAF_ZA_Devs!(), npts%)
If ierror Then Exit Sub

' Calculate Fanal F matrix correction alpha factors for these binary compositions
Call Penepma12CalculateAlphaFactors(l%, BinaryRanges!(), Binary_F_Kratios#(), Binary_F_Factors!(), Binary_F_Coeffs!(), Binary_F_Devs!(), npts%)
If ierror Then Exit Sub

' Calculate CalcZAF F matrix correction alpha factors for these binary compositions
Call Penepma12CalculateAlphaFactors(l%, BinaryRanges!(), CalcZAF_F_Kratios#(), CalcZAF_F_Factors!(), CalcZAF_F_Coeffs!(), CalcZAF_F_Devs!(), npts%)
If ierror Then Exit Sub

' Output the data as desired (conc vs k-ratio/alphas etc.)
Close #Temp1FileNumber%
DoEvents
tfilename$ = tfolder$ & "\" & Format$(ExtractElement%) & "-" & Format$(ExtractMatrix%) & "_" & Format$(MaterialMeasuredTakeoff#) & "_" & Format$(MaterialMeasuredEnergy#) & "_" & Trim$(Symup$(ExtractElement%)) & " " & Xraylo$(l%) & ".dat"
Open tfilename$ For Output As #Temp1FileNumber%

' Output column labels
astring$ = VbDquote$ & "keV" & VbDquote$ & vbTab & VbDquote$ & "Conc" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "ZAF_Krat" & "_(Fanal)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "ZAF_Alpha" & "_(Fanal)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "ZAF_Beta" & "_(Fanal)" & VbDquote$ & vbTab

astring$ = astring$ & VbDquote$ & "ZAF_Krat" & "_(CalcZAF)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "ZAF_Alpha" & "_(CalcZAF)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "ZAF_Beta" & "_(CalcZAF)" & VbDquote$ & vbTab

astring$ = astring$ & VbDquote$ & "ZA_Krat" & "_(Fanal)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "ZA_Alpha" & "_(Fanal)" & VbDquote$ & vbTab

astring$ = astring$ & VbDquote$ & "ZA_Krat" & "_(CalcZAF)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "ZA_Alpha" & "_(CalcZAF)" & VbDquote$ & vbTab

astring$ = astring$ & VbDquote$ & "F_Krat" & "_(Fanal)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "F_Alpha" & "_(Fanal)" & VbDquote$ & vbTab

astring$ = astring$ & VbDquote$ & "F_Krat" & "_(CalcZAF)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "F_Alpha" & "_(CalcZAF)" & VbDquote$ & vbTab
Print #Temp1FileNumber%, astring

' Output data for this energy by concentration (always 99 to 1 wt%)
For n% = 1 To MAXBINARY%

' Calculate the beta factor for each binary composition for the specified binary and xray
Call Penepma12CalculateBinaryBeta(l%, n%, BinaryRanges!(), Binary_ZAF_Coeffs!(), Binary_ZAF_Betas!())
If ierror Then
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

' Calculate the beta factor for each binary composition for the specified binary and xray
Call Penepma12CalculateBinaryBeta(l%, n%, BinaryRanges!(), CalcZAF_ZAF_Coeffs!(), CalcZAF_ZAF_Betas!())
If ierror Then
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

' Output calculations
Print #Temp1FileNumber%, CSng(MaterialMeasuredEnergy#), vbTab, BinaryRanges!(n%), vbTab, _
    CSng(Binary_ZAF_Kratios#(l%, n%)), vbTab, Binary_ZAF_Factors!(l%, n%), vbTab, Binary_ZAF_Betas!(l%, n%), vbTab, _
    CSng(CalcZAF_ZAF_Kratios#(l%, n%)), vbTab, CalcZAF_ZAF_Factors!(l%, n%), vbTab, CalcZAF_ZAF_Betas!(l%, n%), vbTab, _
    CSng(Binary_ZA_Kratios#(l%, n%)), vbTab, Binary_ZA_Factors!(l%, n%), vbTab, _
    CSng(CalcZAF_ZA_Kratios#(l%, n%)), vbTab, CalcZAF_ZA_Factors!(l%, n%), vbTab, _
    CSng(Binary_F_Kratios#(l%, n%)), vbTab, Binary_F_Factors!(l%, n%), vbTab, _
    CSng(CalcZAF_F_Kratios#(l%, n%)), vbTab, CalcZAF_F_Factors!(l%, n%)
    
Next n%

Close #Temp1FileNumber%

msg$ = "The specified " & Trim$(Symup$(ExtractElement%)) & " " & Xraylo$(l%) & " k-ratio plot data was output to " & tfilename$
Call IOWriteLog(msg$)
End If
Next l%

msg$ = "The specified matrix k-ratio plot data (from the available matrix .TXT files in " & tfolder$ & ") was output based on " & Format$(MaterialMeasuredTakeoff#) & " degrees, " & Format$(MaterialMeasuredEnergy#) & " keV, " & Trim$(Symup$(ExtractElement%)) & " in " & Trim$(Symup$(ExtractMatrix%))
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12OutputPlotDataMatrix"

Exit Sub

' Errors
Penepma12OutputPlotDataMatrixError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12OutputPlotDataMatrix"
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12OutputPlotDataMatrixNotFound:
msg$ = "The specified Penepma12ExtractMatrix file (" & tfilename$ & ") was not found. Please run the Extract Matrix calculations first for the specified element and matrix"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12OutputPlotDataMatrix"
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12CalculateComposition()
' Calculate binary or pure element composition .par files for a specified composition.
' mode = 1 get formula
' mode = 2 get weight string
' mode = 3 get standard composition

ierror = False
On Error GoTo Penepma12CalculateCompositionError

Dim ip1 As Integer, ip2 As Integer, mode As Integer
Dim i As Integer, j As Integer, m As Integer
Dim mm As Integer
Dim response As Integer
Dim tBinaryElement1 As Integer, tBinaryElement2 As Integer

' Load mode
If CalculateFromFormulaOrStandard% = 1 Then mode% = 1
If CalculateFromFormulaOrStandard% = 2 Then mode% = 3

' Get a formula composition or standard composition from the user
Call Penepma12CalculateGetComposition(mode%, PENEPMA_OldSample())
If ierror Then Exit Sub

' Check for at least one element
If PENEPMA_OldSample(1).LastChan% < 1 Then Exit Sub

' BinaryMethod = 0  Calculate binary compositions for the specified composition
If BinaryMethod% = 0 Then

' Calculate number of binaries for the specified composition
m% = 0
For i% = 1 To PENEPMA_OldSample(1).LastChan%
For j% = i% To PENEPMA_OldSample(1).LastChan%  ' do not duplicate binary pairs in reverse order
If i% <> j% Then
m% = m% + 1
End If
Next j%
Next i%

TotalNumberOfSimulations& = CLng(m%) * MAXBINARY%    ' specify number of PAR files to create (should be 4950 binaries)
CurrentSimulationsNumber& = 1

msg$ = "This binary compositional range calculation based on the specified composition will require " & Format$(TotalNumberOfSimulations&) & " PAR file calculations. Are you sure you want to proceed?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "Penepma12CalculateComposition")
If response% = vbCancel Then Exit Sub

' Calculating entire matrix range
CalculateForMatrixRange = True  ' to skip user warnings

tBinaryElement1% = BinaryElement1%      ' save
tBinaryElement2% = BinaryElement2%      ' save

mm% = 0
For i% = 1 To PENEPMA_OldSample(1).LastChan%
For j% = i% To PENEPMA_OldSample(1).LastChan%  ' do not duplicate binary pairs in reverse order

' Load next calculation
ip1% = IPOS1(MAXELM%, PENEPMA_OldSample(1).Elsyms$(i%), Symlo$())
If ip1% = 0 Then GoTo Penepma12CalculateCompositionBadElement1
ip2% = IPOS1(MAXELM%, PENEPMA_OldSample(1).Elsyms$(j%), Symlo$())
If ip2% = 0 Then GoTo Penepma12CalculateCompositionBadElement2

' Check for H or He in binary
If ip1% = 1 Or ip1% = 2 Then
msg$ = vbCrLf & "Skipping binary: " & Trim$(Symup$(ip1%)) & "-" & Trim$(Symup$(ip2%)) & "..." & vbCrLf
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
GoTo Penepma12CalculateCompositionSkippingBinary
End If
If ip2% = 1 Or ip2% = 2 Then
msg$ = vbCrLf & "Skipping binary: " & Trim$(Symup$(ip1%)) & "-" & Trim$(Symup$(ip2%)) & "..." & vbCrLf
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
GoTo Penepma12CalculateCompositionSkippingBinary
End If

' Calculate binary
If i% <> j% Then
mm% = mm% + 1

msg$ = vbCrLf & vbCrLf & "Calculating binary " & Format$(mm%) & " of " & Format$(m%) & ": " & Trim$(PENEPMA_OldSample(1).Elsyms$(i%)) & "-" & Trim$(PENEPMA_OldSample(1).Elsyms$(j%)) & "..."
Call IOWriteLog(msg$)
Call IOStatusAuto(msg$)
DoEvents

BinaryElement1% = ip1%                    ' load matrix 1
BinaryElement2% = ip2%                    ' load matrix 2
Call Penepma12CalculateBinaries
BinaryElement1% = tBinaryElement1%      ' restore
BinaryElement2% = tBinaryElement2%      ' restore

If ierror Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

End If
Penepma12CalculateCompositionSkippingBinary:
Next j%
Next i%

msg$ = "All " & Format$(TotalNumberOfSimulations&) & " PAR file calculations are complete"
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12CalculateComposition"
End If

' BinaryMethod = 1  Calculate pure element materials for the specified composition
If BinaryMethod% = 1 Then

TotalNumberOfSimulations& = CLng(PENEPMA_OldSample(1).LastChan%)    ' specify number of PAR files to create
CurrentSimulationsNumber& = 1

msg$ = "This pure element calculation based on the specified composition will require " & Format$(TotalNumberOfSimulations&) & " PAR file calculations. Are you sure you want to proceed?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "Penepma12CalculateComposition")
If response% = vbCancel Then Exit Sub

m% = 0
For i% = 1 To PENEPMA_OldSample(1).LastChan%
m% = m% + 1
Next i%

TotalNumberOfSimulations& = m%    ' specify number of PAR files to create
CurrentSimulationsNumber& = 1

tBinaryElement1% = BinaryElement1%      ' save
tBinaryElement2% = BinaryElement2%      ' save

mm% = 0
For i% = 1 To PENEPMA_OldSample(1).LastElm%
mm% = mm% + 1

msg$ = vbCrLf & vbCrLf & "Calculating pure element " & Format$(mm%) & " of " & Format$(m%) & ": " & Trim$(PENEPMA_OldSample(1).Elsyms$(i%)) & "..."
Call IOWriteLog(msg$)
Call IOStatusAuto(msg$)
DoEvents

' Load next calculation
ip1% = IPOS1(MAXELM%, PENEPMA_OldSample(1).Elsyms$(i%), Symlo$())
If ip1% = 0 Then GoTo Penepma12CalculateCompositionBadElement1
BinaryElement1% = ip1%                       ' load same element for both
BinaryElement2% = ip1%                       ' load same element for both
Call Penepma12CalculateElements
BinaryElement1% = tBinaryElement1%      ' restore
BinaryElement2% = tBinaryElement2%      ' restore

If ierror Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

Next i%

msg$ = "All " & Format$(TotalNumberOfSimulations&) & " PAR file calculations are complete"
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12CalculateComposition"
End If

Call IOStatusAuto(vbNullString)
Exit Sub

' Errors
Penepma12CalculateCompositionError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CalculateComposition"
ierror = True
Exit Sub

Penepma12CalculateCompositionBadElement1:
msg$ = "Element " & PENEPMA_OldSample(1).Elsyms$(i%) & " is not a valid element symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CalculateComposition"
ierror = True
Exit Sub

Penepma12CalculateCompositionBadElement2:
msg$ = "Element " & PENEPMA_OldSample(1).Elsyms$(j%) & " is not a valid element symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CalculateComposition"
ierror = True
Exit Sub

End Sub

Sub Penepma12BinaryLoad()
' Load the form Penepma12Binary

ierror = False
On Error GoTo Penepma12BinaryLoadError

Dim i As Integer

' Binary calculations
If BinaryMethod% = 0 Then
FormPenepma12Binary.OptionBinaryMethod(0).Value = True
Else
FormPenepma12Binary.OptionBinaryMethod(1).Value = True
End If

FormPenepma12Binary.ComboBinaryElement1.Clear
For i% = 0 To MAXELM% - 1
FormPenepma12Binary.ComboBinaryElement1.AddItem Symup$(i% + 1)
Next i%
FormPenepma12Binary.ComboBinaryElement1.ListIndex = BinaryElement1% - 1

FormPenepma12Binary.ComboBinaryElement2.Clear
For i% = 0 To MAXELM% - 1
FormPenepma12Binary.ComboBinaryElement2.AddItem Symup$(i% + 1)
Next i%
FormPenepma12Binary.ComboBinaryElement2.ListIndex = BinaryElement2% - 1

If CalculateDoNotOverwritePAR Then
FormPenepma12Binary.CheckDoNotOverwritePAR.Value = vbChecked
Else
FormPenepma12Binary.CheckDoNotOverwritePAR.Value = vbUnchecked
End If

If CalculateOnlyOverwriteLowerPrecisionPAR Then
FormPenepma12Binary.CheckOverwriteLowerPrecisionPAR.Value = vbChecked
Else
FormPenepma12Binary.CheckOverwriteLowerPrecisionPAR.Value = vbUnchecked
End If

If CalculateOnlyOverwriteHigherMinimumEnergyPAR Then
FormPenepma12Binary.CheckOverwriteHigherMinimumEnergyPAR.Value = vbChecked
Else
FormPenepma12Binary.CheckOverwriteHigherMinimumEnergyPAR.Value = vbUnchecked
End If

If CalculateDoNotOverwriteTXT Then
FormPenepma12Binary.CheckDoNotOverwriteTXT.Value = vbChecked
Else
FormPenepma12Binary.CheckDoNotOverwriteTXT.Value = vbUnchecked
End If

If CalculateForMatrixRange Then
FormPenepma12Binary.CheckCalculateForMatrixRange.Value = vbChecked
Else
FormPenepma12Binary.CheckCalculateForMatrixRange.Value = vbUnchecked
End If

If CalculateFromFormulaOrStandard% = 1 Then
FormPenepma12Binary.OptionFromFormula.Value = True
Else
FormPenepma12Binary.OptionFromStandard.Value = True
End If

If ExtractMethod% = 0 Then
FormPenepma12Binary.OptionExtractMethod(0).Value = True
Else
FormPenepma12Binary.OptionExtractMethod(1).Value = True
End If

FormPenepma12Binary.ComboExtractElement.Clear
For i% = 0 To MAXELM% - 1
FormPenepma12Binary.ComboExtractElement.AddItem Symup$(i% + 1)
Next i%
FormPenepma12Binary.ComboExtractElement.ListIndex = ExtractElement% - 1

FormPenepma12Binary.ComboExtractMatrix.Clear
For i% = 0 To MAXELM% - 1
FormPenepma12Binary.ComboExtractMatrix.AddItem Symup$(i% + 1)
Next i%
FormPenepma12Binary.ComboExtractMatrix.ListIndex = ExtractMatrix% - 1

FormPenepma12Binary.ComboExtractMatrixA1.Clear
For i% = 0 To MAXELM% - 1
FormPenepma12Binary.ComboExtractMatrixA1.AddItem Symup$(i% + 1)
Next i%
FormPenepma12Binary.ComboExtractMatrixA1.ListIndex = ExtractMatrixA1% - 1

FormPenepma12Binary.ComboExtractMatrixA2.Clear
For i% = 0 To MAXELM% - 1
FormPenepma12Binary.ComboExtractMatrixA2.AddItem Symup$(i% + 1)
Next i%
FormPenepma12Binary.ComboExtractMatrixA2.ListIndex = ExtractMatrixA2% - 1

FormPenepma12Binary.ComboExtractMatrixB1.Clear
For i% = 0 To MAXELM% - 1
FormPenepma12Binary.ComboExtractMatrixB1.AddItem Symup$(i% + 1)
Next i%
FormPenepma12Binary.ComboExtractMatrixB1.ListIndex = ExtractMatrixB1% - 1

FormPenepma12Binary.ComboExtractMatrixB2.Clear
For i% = 0 To MAXELM% - 1
FormPenepma12Binary.ComboExtractMatrixB2.AddItem Symup$(i% + 1)
Next i%
FormPenepma12Binary.ComboExtractMatrixB2.ListIndex = ExtractMatrixB2% - 1

If ExtractForSpecifiedRange Then
FormPenepma12Binary.CheckExtractForSpecifiedRange.Value = vbChecked
Else
FormPenepma12Binary.CheckExtractForSpecifiedRange.Value = vbUnchecked
End If

' Load alpha factor calculation fields
If OptionEnterFraction Then
FormPenepma12Binary.OptionEnter(0).Value = True
Else
FormPenepma12Binary.OptionEnter(1).Value = True
End If

If ConcA! = 0# Then ConcA! = 1#
If ConcB! = 0# Then ConcB! = 99#
If KratA! = 0# Then KratA! = 1.030796
If KratB! = 0# Then KratB! = 98.97312
FormPenepma12Binary.TextConcA.Text = Format$(ConcA!)        ' Al ka in Si
FormPenepma12Binary.TextConcB.Text = Format$(ConcB!)
FormPenepma12Binary.TextKratA.Text = Format$(KratA!)        ' Al ka in Si (15 keV)
FormPenepma12Binary.TextKratB.Text = Format$(KratB!)

' Load density calculation fields
If DensityElementA% = 0 Then DensityElementA% = 14
If DensityElementB% = 0 Then DensityElementB% = 8
FormPenepma12Binary.TextDensityElementA.Text = Symup$(DensityElementA%)
FormPenepma12Binary.TextDensityElementB.Text = Symup$(DensityElementB%)

If DensityConcA! = 0# Then DensityConcA! = 46.734
If DensityConcB! = 0# Then DensityConcB! = 53.257
FormPenepma12Binary.TextConcAA.Text = Format$(DensityConcA!)
FormPenepma12Binary.TextConcBB.Text = Format$(DensityConcB!)

FormPenepma12Binary.OptionMinimumOvervoltagePercent(MinimumOverVoltageType%).Value = True

Exit Sub

' Errors
Penepma12BinaryLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12BinaryLoad"
ierror = True
Exit Sub

End Sub

Sub Penepma12BinarySave()
' Save the form Penepma12Binary

ierror = False
On Error GoTo Penepma12BinarySaveError

Dim ip As Integer
Dim esym As String

' Save binary elements
If FormPenepma12Binary.OptionBinaryMethod(0).Value = True Then
BinaryMethod% = 0
Else
BinaryMethod% = 1
End If

esym$ = FormPenepma12Binary.ComboBinaryElement1.Text
ip% = IPOS1(MAXELM%, esym$, Symlo$())
If ip% = 0 Then GoTo Penepma12BinarySaveBadBinaryElement
BinaryElement1% = ip%

esym$ = FormPenepma12Binary.ComboBinaryElement2.Text
ip% = IPOS1(MAXELM%, esym$, Symlo$())
If ip% = 0 Then GoTo Penepma12BinarySaveBadBinaryElement
BinaryElement2% = ip%

If FormPenepma12Binary.CheckDoNotOverwritePAR.Value = vbChecked Then
CalculateDoNotOverwritePAR = True
Else
CalculateDoNotOverwritePAR = False
End If

If FormPenepma12Binary.CheckOverwriteLowerPrecisionPAR.Value = vbChecked Then
CalculateOnlyOverwriteLowerPrecisionPAR = True
Else
CalculateOnlyOverwriteLowerPrecisionPAR = False
End If

If FormPenepma12Binary.CheckOverwriteHigherMinimumEnergyPAR.Value = vbChecked Then
CalculateOnlyOverwriteHigherMinimumEnergyPAR = True
Else
CalculateOnlyOverwriteHigherMinimumEnergyPAR = False
End If

If FormPenepma12Binary.CheckDoNotOverwriteTXT.Value = vbChecked Then
CalculateDoNotOverwriteTXT = True
Else
CalculateDoNotOverwriteTXT = False
End If

If FormPenepma12Binary.CheckCalculateForMatrixRange.Value = vbChecked Then
CalculateForMatrixRange = True
Else
CalculateForMatrixRange = False
End If

If FormPenepma12Binary.OptionExtractMethod(0).Value = True Then
ExtractMethod% = 0
Else
ExtractMethod% = 1
End If

If FormPenepma12Binary.OptionFromFormula.Value = True Then CalculateFromFormulaOrStandard% = 1
If FormPenepma12Binary.OptionFromStandard.Value = True Then CalculateFromFormulaOrStandard% = 2

' Save extract element and matrix
esym$ = FormPenepma12Binary.ComboExtractElement.Text
ip% = IPOS1(MAXELM%, esym$, Symlo$())
If ip% = 0 Then GoTo Penepma12BinarySaveBadExtractElement
ExtractElement% = ip%

esym$ = FormPenepma12Binary.ComboExtractMatrix.Text
ip% = IPOS1(MAXELM%, esym$, Symlo$())
If ip% = 0 Then GoTo Penepma12BinarySaveBadExtractElement
ExtractMatrix% = ip%

esym$ = FormPenepma12Binary.ComboExtractMatrixA1.Text
ip% = IPOS1(MAXELM%, esym$, Symlo$())
If ip% = 0 Then GoTo Penepma12BinarySaveBadExtractElement
ExtractMatrixA1% = ip%

esym$ = FormPenepma12Binary.ComboExtractMatrixA2.Text
ip% = IPOS1(MAXELM%, esym$, Symlo$())
If ip% = 0 Then GoTo Penepma12BinarySaveBadExtractElement
ExtractMatrixA2% = ip%

esym$ = FormPenepma12Binary.ComboExtractMatrixB1.Text
ip% = IPOS1(MAXELM%, esym$, Symlo$())
If ip% = 0 Then GoTo Penepma12BinarySaveBadExtractElement
ExtractMatrixB1% = ip%

esym$ = FormPenepma12Binary.ComboExtractMatrixB2.Text
ip% = IPOS1(MAXELM%, esym$, Symlo$())
If ip% = 0 Then GoTo Penepma12BinarySaveBadExtractElement
ExtractMatrixB2% = ip%

If FormPenepma12Binary.CheckExtractForSpecifiedRange.Value = vbChecked Then
ExtractForSpecifiedRange = True
Else
ExtractForSpecifiedRange = False
End If

' Save alpha factor calculation fields
If FormPenepma12Binary.OptionEnter(0).Value = True Then
OptionEnterFraction = True
Else
OptionEnterFraction = False
End If

ConcA! = Val(FormPenepma12Binary.TextConcA.Text)
ConcB! = Val(FormPenepma12Binary.TextConcB.Text)
KratA! = Val(FormPenepma12Binary.TextKratA.Text)
KratB! = Val(FormPenepma12Binary.TextKratB.Text)

' Save density calculation fields
esym$ = FormPenepma12Binary.TextDensityElementA.Text
ip% = IPOS1%(MAXELM%, esym$, Symlo$())
If ip% = 0 Then GoTo Penepma12BinarySaveBadBinaryDensityElement
DensityElementA% = ip%

esym$ = FormPenepma12Binary.TextDensityElementB.Text
ip% = IPOS1%(MAXELM%, esym$, Symlo$())
If ip% = 0 Then GoTo Penepma12BinarySaveBadBinaryDensityElement
DensityElementB% = ip%

DensityConcA! = Val(FormPenepma12Binary.TextConcAA.Text)
DensityConcB! = Val(FormPenepma12Binary.TextConcBB.Text)

If FormPenepma12Binary.OptionMinimumOvervoltagePercent(0).Value = True Then MinimumOverVoltageType% = 0
If FormPenepma12Binary.OptionMinimumOvervoltagePercent(1).Value = True Then MinimumOverVoltageType% = 1
If FormPenepma12Binary.OptionMinimumOvervoltagePercent(2).Value = True Then MinimumOverVoltageType% = 2
If FormPenepma12Binary.OptionMinimumOvervoltagePercent(3).Value = True Then MinimumOverVoltageType% = 3

Exit Sub

' Errors
Penepma12BinarySaveError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12BinarySave"
ierror = True
Exit Sub

Penepma12BinarySaveBadExtractElement:
msg$ = "Binary extraction element " & esym$ & " is not a valid element symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12BinarySave"
ierror = True
Exit Sub

Penepma12BinarySaveBadBinaryElement:
msg$ = "Binary calculation element " & esym$ & " is not a valid element symbol for a binary element"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12BinarySave"
ierror = True
Exit Sub

Penepma12BinarySaveBadBinaryDensityElement:
msg$ = "Density calculation element " & esym$ & " is not a valid element symbol for a density element"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12BinarySave"
ierror = True
Exit Sub

End Sub

Sub Penepma12AtomicWeights()
' Read and write the Penepma atomic weight file (for self consistency in calculations)

ierror = False
On Error GoTo Penepma12AtomicWeightsError

Dim tfilename As String, astring As String, bstring As String
Dim i As Integer, linecount As Integer
Dim idnum As Long, atnum As Long, numelm As Long
Dim ZbyA As Double, wtfrac As Double

' Check Penepma 2008 or 2012 atomic weights
If Penepma08CheckPenepmaVersion%() <= 12 Then
tfilename$ = PENDBASE_Path$ & "\pdfiles\" & "pdcompos.p08"

' Penepma 2014 or 2016 atomic weights
Else
tfilename$ = PENDBASE_Path$ & "\pdfiles\" & "pdcompos.pen"
End If

If Dir$(tfilename$) = vbNullString Then
msg$ = "Warning in Penepma12AtomicWeights: Penepma 2008/2012 atomic weights file " & tfilename$ & ", was not found. CalcZAF atomic weights will be utilized for calculations."
Call IOWriteLog(msg$)

' Load CalcZAF atomic weights
For i% = 1 To MAXELM%
pAllAtomicWts!(i%) = AllAtomicWts!(i%)
Next i%
Exit Sub
End If

' Open pdcompos.p08 or pdcompos.pen
linecount% = 1
Open tfilename$ For Input As #Temp1FileNumber%

' Skip comments lines
Do Until InStr(astring$, "*********************************") > 0
Line Input #Temp1FileNumber%, astring$
linecount% = linecount% + 1
Loop

' Skip blank line
Line Input #Temp1FileNumber%, astring$
linecount% = linecount% + 1

' Penepma only goes up to element 99
For i% = 1 To MAXELM% - 1

' Read element line
Line Input #Temp1FileNumber%, astring$
linecount% = linecount% + 1
Call MiscParseStringToString(astring, bstring$)    ' parse atomic number
idnum& = Val(bstring$)

' Read Z/A line
Line Input #Temp1FileNumber%, astring$
linecount% = linecount% + 1
Call MiscParseStringToString(astring, bstring$)    ' parse number of elements (should be one)
numelm& = Val(bstring$)
If numelm& <> 1 Then GoTo Penepma12AtomicWeightsBadNumberofElements

Call MiscParseStringToString(astring, bstring$)    ' parse Z/A
ZbyA# = Val(bstring$)

' Read weight fraction line line
Line Input #Temp1FileNumber%, astring$
linecount% = linecount% + 1
Call MiscParseStringToString(astring, bstring$)    ' parse atomic number
atnum& = Val(bstring$)

Call MiscParseStringToString(astring, bstring$)    ' parse weight fraction (should be 1.000)
wtfrac# = Val(bstring$)
If wtfrac# <> 1# Then GoTo Penepma12AtomicWeightsBadWeightFraction

' Calculate actual atomic weight
pAllAtomicWts!(i%) = CSng(atnum& / ZbyA#)
If pAllAtomicWts!(i%) < 1# Or pAllAtomicWts!(i%) > 254# Then GoTo Penepma12AtomicWeightsInvalidData
Next i%

Close #Temp1FileNumber%

' Load last atomic weight (not in pdcompos.p08 or pdcompos.pen file)
pAllAtomicWts!(MAXELM%) = AllAtomicWts!(MAXELM%)

Exit Sub

' Errors
Penepma12AtomicWeightsError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12AtomicWeights"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma12AtomicWeightsBadNumberofElements:
msg$ = "Invalid number of elements (" & Format$(numelm&) & ") in " & tfilename$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12AtomicWeights"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma12AtomicWeightsBadWeightFraction:
msg$ = "Invalid weight fraction (" & Format$(wtfrac#) & ") in " & tfilename$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12AtomicWeights"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma12AtomicWeightsInvalidData:
msg$ = "Invalid atomic weight (" & Format$(pAllAtomicWts!(i%)) & ") in " & tfilename$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12AtomicWeights"
Close #Temp1FileNumber%
ierror = True
Exit Sub


End Sub

Sub Penepma12RunFanal1()
' Run Fanal based on current parameter files specified

ierror = False
On Error GoTo Penepma12RunFanal1Error

' Delete Fanal output files if found
'If Dir$(FANAL_IN_File$) <> vbNullString Then Kill FANAL_IN_File$    ' do not delete input file!
If Dir$(VACS_DAT_File$) <> vbNullString Then Kill VACS_DAT_File$
If Dir$(RANGES_DAT_File$) <> vbNullString Then Kill RANGES_DAT_File$
If Dir$(MIXED_DAT_File$) <> vbNullString Then Kill MIXED_DAT_File$
If Dir$(KRATIOS_DAT_File$) <> vbNullString Then Kill KRATIOS_DAT_File$
If Dir$(FLUORMAT1_PAR_File$) <> vbNullString Then Kill FLUORMAT1_PAR_File$
If Dir$(FLUORMAT2_PAR_File$) <> vbNullString Then Kill FLUORMAT2_PAR_File$
If Dir$(FLUORMAT3_PAR_File$) <> vbNullString Then Kill FLUORMAT3_PAR_File$
If Dir$(ATCOEFFS_DAT_File$) <> vbNullString Then Kill ATCOEFFS_DAT_File$

' Create Fanal input file
Call Penepma12RunFanalCreateInput
If ierror Then Exit Sub

' Run Fanal on input file
Call Penepma12RunFanal2
If ierror Then Exit Sub

' Now wait for the Fanal calculation to finish
Do Until Not CalculationInProgress
DoEvents
Loop

Call Penepma12UpdateForm
If ierror Then Exit Sub

' Check for Fanal output files
If Dir$(VACS_DAT_File$) = vbNullString Then GoTo Penepma12RunFanal1FilesNotFound
If Dir$(RANGES_DAT_File$) = vbNullString Then GoTo Penepma12RunFanal1FilesNotFound
If Dir$(MIXED_DAT_File$) = vbNullString Then GoTo Penepma12RunFanal1FilesNotFound
If Dir$(KRATIOS_DAT_File$) = vbNullString Then GoTo Penepma12RunFanal1FilesNotFound
If Dir$(FLUORMAT1_PAR_File$) = vbNullString Then GoTo Penepma12RunFanal1FilesNotFound
If Dir$(FLUORMAT2_PAR_File$) = vbNullString Then GoTo Penepma12RunFanal1FilesNotFound
If Dir$(FLUORMAT3_PAR_File$) = vbNullString Then GoTo Penepma12RunFanal1FilesNotFound
If Dir$(ATCOEFFS_DAT_File$) = vbNullString Then GoTo Penepma12RunFanal1FilesNotFound

Exit Sub

' Errors
Penepma12RunFanal1Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12RunFanal1"
ierror = True
Exit Sub

Penepma12RunFanal1FilesNotFound:
msg$ = "The Fanal output files were not found. Check Fanal.exe by running the input file manually using a command prompt from the Fanal folder and typing Fanal < Fanal.in"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunFanal1"
ierror = True
Exit Sub

End Sub

Sub Penepma12OutputKratios3(ipA As Integer, analysis As TypeAnalysis, sample() As TypeSample, kratio#, unk_krat_meas#, unk_zaf_meas#, unk_conc_meas#)
' Perform an iteration to calculate the "apparent" concentration (B Std must be a pure element)

ierror = False
On Error GoTo Penepma12OutputKratios3Error

Dim i As Integer
Dim zerror As Integer
Dim unk_krats(1 To MAXCHAN%) As Single

' Load all elements as elemental k-ratios
For i% = 1 To sample(1).LastChan%
unk_krats!(i%) = MatA_Krats!(i%)
If i% = ipA% Then unk_krats!(i%) = kratio# / 100#    ' load calculated (measured) k-ratio
Next i%

' Reload the element arrays based on the unknown sample setup
Call ElementGetData(sample())
If ierror Then Exit Sub

' Initialize calculations (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% <> MAXCORRECTION% Then
Call ZAFSetZAF(sample())
If ierror Then Exit Sub
Else
'Call ZAFSetZAF3(sample())
'If ierror Then Exit Sub
End If

' No assigned standard used in k-ratio calculation
For i% = 1 To sample(1).LastElm%
sample(1).StdAssigns%(i%) = MAXINTEGER%     ' fake standard assignment
Next i%

' Init intensities for unknown and standard
For i% = 1 To sample(1).LastChan%
analysis.StdAssignsCounts!(i%) = 1#
analysis.StdAssignsKfactors!(i%) = 1#
analysis.StdAssignsPercents!(i%) = 100#
Next i%

' Calculate ZAF weights
Call ZAFSmp(Int(1), unk_krats!(), zerror%, analysis, sample())
If ierror Then Exit Sub

' Return k-ratios and concentrations
unk_krat_meas# = analysis.UnkKrats!(ipA%) * 100#
unk_zaf_meas# = analysis.UnkZAFCors!(4, ipA%)
unk_conc_meas# = analysis.WtPercents!(ipA%)

Exit Sub

' Errors
Penepma12OutputKratios3Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12OutputKratios3"
ierror = True
Exit Sub

End Sub

Sub Penepma12ExtractBoundaryCalculate(nPoints As Long)
' Calculate alpha factors and fit coefficients for boundary k-ratios
'  npoints& is the number of distance points read or to write
'
'  Boundary_ZAF_Kratios#(1 to MAXBINARY%, 1 to MAXBINARY%, 1 to npoints&)  are the k-ratios from Fanal in k-ratio % for each x distance
'  Boundary_ZAF_Factors!(1 to MAXBINARY%, 1 to MAXBINARY%, 1 to npoints&)  are the alpha factors for each boundary composition, alpha = (C/K - C)/(1 - C)
'
'  Boundary_Linear_Distances(1 to npoints&)                     ' linear distances (not used here)
'  Boundary_Mass_Distances(1 to MAXBINARY%, 1 to npoints&)      ' mass distances (not used here)
'
'  Boundary_Material_A_Densities(1 to MAXBINARY%)  ' material A densities (not used here)
'  Boundary_Material_B_Densities(1 to MAXBINARY%)  ' material B densities (not used here)

Dim n As Long
Dim k As Integer, j As Integer

ierror = False
On Error GoTo Penepma12ExtractBoundaryCalculateError

Dim tBinaryRanges(1 To MAXBINARY%) As Single
Dim tBinaryKratios(1 To MAXBINARY%) As Double
Dim tBinaryFactors(1 To MAXBINARY%) As Single

' If element is not present in material A, then skip alpha factor and fit calcuations
If ExtractElement% = ExtractMatrixA1% Or ExtractElement% = ExtractMatrixA2% Then

' Calculate for all distances
For n& = 1 To nPoints&

' Load binary arrays from boundary arrays
For j% = 1 To MAXBINARY%    ' material A
For k% = 1 To MAXBINARY%    ' material B
If ExtractElement% = ExtractMatrixA1% Then
tBinaryRanges!(k%) = BinaryRanges!(j%)              ' concentration of emitted element in material A
Else
tBinaryRanges!(k%) = 100# - BinaryRanges!(j%)       ' concentration of emitted element in material A
End If

tBinaryKratios#(k%) = Boundary_ZAF_Kratios#(k%, j%, n&)
Next k%

' Calculate ZAF boundary correction alpha factors for these boundary binary compositions
Call Penepma12CalculateAlphaFactors2(tBinaryRanges!(), tBinaryKratios#(), tBinaryFactors!())
If ierror Then Exit Sub

' Load boundary arrays from binary arrays
For k% = 1 To MAXBINARY%
Boundary_ZAF_Factors!(k%, j%, n&) = tBinaryFactors!(k%)
Next k%
Next j%

Next n&

' Skip alpha calculations
Else
msg$ = "Skipping boundary alpha factor calculations for " & Trim$(Symup$(ExtractElement%)) & " in " & Trim$(Symup$(ExtractMatrixA1%)) & "-" & Trim$(Symup$(ExtractMatrixA2%)) & " adjacent to " & Trim$(Symup$(ExtractMatrixB1%)) & "-" & Trim$(Symup$(ExtractMatrixB2%))
Call IOWriteLog(msg$)

' Zero binary arrays
For n& = 1 To nPoints&
For j% = 1 To MAXBINARY%    ' material A
For k% = 1 To MAXBINARY%    ' material B
Boundary_ZAF_Factors!(k%, j%, n&) = 0#
Next k%
Next j%

Next n&
End If

Exit Sub

' Errors
Penepma12ExtractBoundaryCalculateError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12ExtractBoundaryCalculate"
ierror = True
Exit Sub

End Sub

Sub Penepma12BinaryCalculateAlphaFactor()
' Calculate an alpha factor based on user input

ierror = False
On Error GoTo Penepma12BinaryCalculateAlphaFactorError

Dim k As Single, c As Single
Dim alphaA As Single, alphaB As Single
Dim astring As String

' Determine input
If OptionEnterFraction Then
c! = ConcA!
k! = KratA!
Else
c! = ConcA! / 100#
k! = KratA! / 100#
End If

' Calculate alpha factor for this binary composition
alphaA! = ((c! / k!) - c!) / (1 - c!)        ' calculate binary alpha factors

' Output to log
msg$ = vbCrLf & "Penepma12BinaryCalculateAlphaFactor: C = " & MiscAutoFormat$(c!) & ", K = " & MiscAutoFormat$(k!) & ", AlphaA = " & MiscAutoFormat$(alphaA!)
Call IOWriteLog(msg$)
astring$ = msg$

' Determine input
If OptionEnterFraction Then
c! = ConcB!
k! = KratB!
Else
c! = ConcB! / 100#
k! = KratB! / 100#
End If

' Calculate alpha factor for this binary composition
alphaB! = ((c! / k!) - c!) / (1 - c!)        ' calculate binary alpha factors

' Output to log
msg$ = "Penepma12BinaryCalculateAlphaFactor: C = " & MiscAutoFormat$(c!) & ", K = " & MiscAutoFormat$(k!) & ", AlphaB = " & MiscAutoFormat$(alphaB!)
Call IOWriteLog(msg$)
astring$ = astring$ & vbCrLf & msg$

MsgBox astring$, vbOKOnly + vbInformation, "Penepma12BinaryCalculateAlphaFactor"
Exit Sub

' Errors
Penepma12BinaryCalculateAlphaFactorError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12BinaryCalculateAlphaFactor"
ierror = True
Exit Sub

End Sub

Sub Penepma12GetStandard(method As Integer, mode As Integer)
' Get the standard number and print it out
'  method = 0 single click, just update density
'  method = 1 double click, also output composition to log window
'  mode = 1 = Mat A
'  mode = 2 = Mat B
'  mode = 3 = Mat B Std

Dim stdnum As Integer

' Get standard from listbox
If mode% = 1 Then
If FormPENEPMA12.ListAvailableStandardsA.ListIndex < 0 Then Exit Sub
stdnum% = FormPENEPMA12.ListAvailableStandardsA.ItemData(FormPENEPMA12.ListAvailableStandardsA.ListIndex)
End If
If mode% = 2 Then
If FormPENEPMA12.ListAvailableStandardsB.ListIndex < 0 Then Exit Sub
stdnum% = FormPENEPMA12.ListAvailableStandardsB.ItemData(FormPENEPMA12.ListAvailableStandardsB.ListIndex)
End If
If mode% = 3 Then
If FormPENEPMA12.ListAvailableStandardsBStd.ListIndex < 0 Then Exit Sub
stdnum% = FormPENEPMA12.ListAvailableStandardsBStd.ItemData(FormPENEPMA12.ListAvailableStandardsBStd.ListIndex)
End If

' Update density field
If mode% = 1 Then
Call StandardGetMDBStandard(stdnum%, PENEPMA_SampleA())
If ierror Then Exit Sub
FormPENEPMA12.TextMaterialDensityA.Text = Format$(PENEPMA_SampleA(1).SampleDensity!)
End If
If mode% = 2 Then
Call StandardGetMDBStandard(stdnum%, PENEPMA_SampleB())
If ierror Then Exit Sub
FormPENEPMA12.TextMaterialDensityB.Text = Format$(PENEPMA_SampleB(1).SampleDensity!)
End If
If mode% = 3 Then
Call StandardGetMDBStandard(stdnum%, PENEPMA_SampleBStd())
If ierror Then Exit Sub
FormPENEPMA12.TextMaterialDensityBStd.Text = Format$(PENEPMA_SampleBStd(1).SampleDensity!)
End If

' Display standard data
If method% = 1 Then
If stdnum% > 0 Then Call StandardTypeStandard(stdnum%)
If ierror Then Exit Sub
End If

Exit Sub

' Errors
Penepma12GetStandardError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12GetStandard"
ierror = True
Exit Sub

End Sub

Function Penepma12GetParFileDensityOnly(tfilename As String) As Single
' Extract the density only from the passed .par file
'
' Sample .par file composition
'
' Calcite.mat
' 6                         NELEM
'  12  6.00000E-02          IZ, atoms/mol
'  20  1.98320E+01          IZ, atoms/mol
'  25  6.10000E-02          IZ, atoms/mol
'  26  4.70000E-02          IZ, atoms/mol
'   8  6.00000E+01          IZ, atoms/mol
'   6  2.00000E+01          IZ, atoms/mol
'      2.71000E+00          Mass density (g/cm**3)

ierror = False
On Error GoTo Penepma12GetParFileDensityOnlyError

Dim i As Integer, n As Integer, atnum As Integer
Dim astring As String, bstring As String

Dim atoms(1 To MAXCHAN%) As Single

' Check for file
If Trim$(tfilename$) = vbNullString Then GoTo Penepma12GetParFileDensityOnlyPARFileNoFile
If Dir$(tfilename$) = vbNullString Then GoTo Penepma12GetParFileDensityOnlyPARFileNotFound

' Open file and parse
Close #Temp1FileNumber%
Open tfilename$ For Input As #Temp1FileNumber%

' Read number of elements
Line Input #Temp1FileNumber%, astring$      ' read material filename line

Line Input #Temp1FileNumber%, astring$      ' read number of elements line
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Function
n% = Val(Trim$(bstring$))

' Load atomic numbers and concentrations
For i% = 1 To n%
Line Input #Temp1FileNumber%, astring$

' Load atomic number (symbol)
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Function

atnum% = Val(bstring$)
If atnum% < 1 Or atnum% > MAXELM% Then GoTo Penepma12GetParFileDensityOnlyBadAtomicNumber

' Load molecules
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Function
atoms!(i%) = Val(Trim$(bstring$))
Next i%

' Input sample density
Line Input #Temp1FileNumber%, astring$
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Function
Penepma12GetParFileDensityOnly! = Val(bstring$)

Close #Temp1FileNumber%

Exit Function

' Errors
Penepma12GetParFileDensityOnlyError:
MsgBox Error$ & ", reading file " & tfilename$, vbOKOnly + vbCritical, "Penepma12GetParFileDensityOnly"
Close #Temp1FileNumber%
ierror = True
Exit Function

Penepma12GetParFileDensityOnlyPARFileNoFile:
msg$ = "The specified .PAR file was blank"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12GetParFileDensityOnly"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Function

Penepma12GetParFileDensityOnlyPARFileNotFound:
msg$ = "The specified .PAR file (" & tfilename$ & ") was not found. Please calculate the specified parameter file and try again"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12GetParFileDensityOnly"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Function

Penepma12GetParFileDensityOnlyBadAtomicNumber:
msg$ = "Invalid atomic number (" & Format$(atnum%) & ") read from the specified .PAR file (" & tfilename$ & ")."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12GetParFileDensityOnly"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Function

End Function

Sub Penepma12CalculateBinaryDensity()
' Procedure to calculate pure element normalized density for a specified binary composition

ierror = False
On Error GoTo Penepma12CalculateBinaryDensityError

Dim atoms(1 To MAXCHAN%) As Single

' Load sample
PENEPMA_Sample(1).LastChan% = 2

' Calculate the pure element normalized density based on composition
PENEPMA_Sample(1).ElmPercents!(1) = DensityConcA!
PENEPMA_Sample(1).ElmPercents!(2) = DensityConcB!

' Overload with Penepma08/12 atomic weights for self consistency in calculations
PENEPMA_Sample(1).AtomicWts!(1) = pAllAtomicWts!(DensityElementA%)
PENEPMA_Sample(1).AtomicWts!(2) = pAllAtomicWts!(DensityElementB%)

Call ConvertWeightToAtomic(PENEPMA_Sample(1).LastChan%, PENEPMA_Sample(1).AtomicWts!(), PENEPMA_Sample(1).ElmPercents!(), atoms!())
If ierror Then Exit Sub

PENEPMA_Sample(1).SampleDensity! = atoms!(1) * AllAtomicDensities!(DensityElementA%) + atoms!(2) * AllAtomicDensities!(DensityElementB%)

FormPenepma12Binary.LabelDensity.Caption = Format$(PENEPMA_Sample(1).SampleDensity!)
Exit Sub

' Errors
Penepma12CalculateBinaryDensityError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CalculateBinaryDensity"
ierror = True
Exit Sub

End Sub

Sub Penepma12ExtractSwapped(tExtractMatrixA1 As Integer, tExtractMatrixA2 As Integer, tExtractMatrixB1 As Integer, tExtractMatrixB2 As Integer)
' Determine which PAR files are present and set "swap" flags

ierror = False
On Error GoTo Penepma12ExtractSwappedError

Dim pfilename1 As String, pfilename2 As String

' Check which first binary PAR file is available (e.g., Fe-Ni_99-1 or Ni-Fe_1-99 as either will work)
pfilename1$ = Trim$(Symup$(tExtractMatrixA1%)) & "-" & Trim$(Symup$(tExtractMatrixA2%)) & "_" & Format$(BinaryRanges!(1)) & "-" & Format$(100# - BinaryRanges!(1))
If Dir$(PENEPMA_Root$ & "\Penfluor\" & pfilename1$ & ".par") <> vbNullString Then
BinaryElementsSwappedA = False

' Try swapping the elements and concentrations
Else
pfilename1$ = Trim$(Symup$(tExtractMatrixA2%)) & "-" & Trim$(Symup$(tExtractMatrixA1%)) & "_" & Format$(BinaryRanges!(MAXBINARY%)) & "-" & Format$(100# - BinaryRanges!(MAXBINARY%))
If Dir$(PENEPMA_Root$ & "\Penfluor\" & pfilename1$ & ".par") <> vbNullString Then
BinaryElementsSwappedA = True

' If neither found then warn user and give up
Else
msg$ = "Cannot find either variant of material A .PAR file " & pfilename1$ & ". The extraction cannot be performed."
MsgBox msg$, vbOKOnly + vbQuestion + vbDefaultButton2, "Penepma12ExtractSwapped"
ierror = True
Exit Sub
End If

End If

' Check which second binary PAR file is available (e.g., Fe-Ni_99-1 or Ni-Fe_1-99 as either will work)
pfilename2$ = Trim$(Symup$(tExtractMatrixB1%)) & "-" & Trim$(Symup$(tExtractMatrixB2%)) & "_" & Format$(BinaryRanges!(1)) & "-" & Format$(100# - BinaryRanges!(1))
If Dir$(PENEPMA_Root$ & "\Penfluor\" & pfilename2$ & ".par") <> vbNullString Then
BinaryElementsSwappedB = False

' Try swapping the elements and concentrations
Else
pfilename2$ = Trim$(Symup$(tExtractMatrixB2%)) & "-" & Trim$(Symup$(tExtractMatrixB1%)) & "_" & Format$(BinaryRanges!(MAXBINARY%)) & "-" & Format$(100# - BinaryRanges!(MAXBINARY%))
If Dir$(PENEPMA_Root$ & "\Penfluor\" & pfilename2$ & ".par") <> vbNullString Then
BinaryElementsSwappedB = True

' If neither found then warn user and give up
Else
msg$ = "Cannot find either variant of material B .PAR file " & pfilename2$ & ". The extraction cannot be performed."
MsgBox msg$, vbOKOnly + vbQuestion + vbDefaultButton2, "Penepma12ExtractSwapped"
ierror = True
Exit Sub
End If

End If

Exit Sub

' Errors
Penepma12ExtractSwappedError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12ExtractSwapped"
ierror = True
Exit Sub

End Sub

Sub Penepma12CalculateKratios(tForm As Form)
' Calculate k-ratios from binary data from .DAT file and compare to experimental kratios to calculate error histogram
'
' Data file format assumes one line for each binary. The first two
' columns are the atomic numbers of the two binary components
' to be calculated. The second two columns are the xray lines to use
' (1 = Ka, 2 = Kb, 3 = La, 4 = Lb, 5 = Ma, 6 = Mb, 7 = Ln, 8 = Lg,
' 9 = Lv, 10 = Ll, 11 = Mg, 12 = Mz, 13 = by difference). The next
' two columns are the operating voltage and take-off angle. The next
' two columns are the wt. fractions of the binary components. The
' last two columns contains the k-exp values for calculation of k-calc/k-exp.
'
'       79     29     5    13    15.     52.5    .8015   .1983   .7400   .0
'       79     29     5    13    15.     52.5    .6036   .3964   .5110   .0
'       79     29     5    13    15.     52.5    .4010   .5992   .3120   .0
'       79     29     5    13    15.     52.5    .2012   .7985   .1450   .0

ierror = False
On Error GoTo Penepma12CalculateKratiosError

Dim tMaterialMeasuredGridPoints As Integer

Dim n As Integer, ibin As Integer
Dim tfilename As String, pfilename As String
Dim binaryname As String
Dim eO As Single, TOA As Single
Dim eng As Single, edg As Single

Dim average As TypeAverage

Dim ncmp As Integer
Dim ksym() As Integer
Dim kcmp() As Single

Static lastfilename As String

ReDim isym(1 To 2) As Integer
ReDim iray(1 To 2) As Integer
ReDim conc(1 To 2) As Single
ReDim kexp(1 To 2) As Single

Dim BinaryLineCount As Long
Dim BinaryLineTotal As Long

Dim BinaryLine() As Integer
Dim BinaryEsym() As String
Dim BinaryXsym() As String
Dim BinaryT0E0() As Single
Dim BinaryConc() As Single
Dim BinaryKexp() As Single
Dim BinaryKratio() As Single
Dim BinaryKerror() As Single

icancelauto = False

Close #ImportDataFileNumber%
Close #ExportDataFileNumber%

' Get import filename from user
If lastfilename$ = vbNullString Then lastfilename$ = CalcZAFDATFileDirectory$ & "\Pouchou2_Au,Cu,Ag_only.dat"
tfilename$ = lastfilename$
Call IOGetFileName(Int(2), "DAT", tfilename$, tForm)
If ierror Then
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
Exit Sub
End If

' Save current path
CalcZAFDATFileDirectory$ = CurDir$

' No errors, save file name
lastfilename$ = tfilename$
ImportDataFile$ = lastfilename$

' Get export filename from user
tfilename$ = MiscGetFileNameNoExtension(tfilename$) & ".out"
Call IOGetFileName(Int(1), "OUT", tfilename$, tForm)
If ierror Then
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
Exit Sub
End If

' No errors, save file name
ExportDataFile$ = tfilename$

' Open files
Open ImportDataFile$ For Input As #ImportDataFileNumber%
BinaryLineCount& = 0
BinaryLineTotal& = 0
Call IOStatusAuto(vbNullString)

' Loop through and find unique compositions first
ncmp% = 0
Do While Not EOF(ImportDataFileNumber%)
BinaryLineCount& = BinaryLineCount& + 1

msg$ = "Checking binary composition " & Str$(BinaryLineCount&) & "..."
Call IOStatusAuto(msg$)
If icancelauto Then
Call IOStatusAuto(vbNullString)
Close #ImportDataFileNumber%
ierror = True
Exit Sub
End If

' Read binary elements, kilovolts and takeoff
Input #ImportDataFileNumber%, isym%(1), isym%(2), iray%(1), iray%(2), eO!, TOA!, conc!(1), conc!(2), kexp!(1), kexp!(2)

' Check limits
If isym%(1) < 1 Or isym%(1) > MAXELM% Then GoTo Penepma12CalculateKratiosOutofLimits
If isym%(2) < 1 Or isym%(2) > MAXELM% Then GoTo Penepma12CalculateKratiosOutofLimits
If iray%(1) < 1 Or iray%(1) > MAXRAY% Then GoTo Penepma12CalculateKratiosOutofLimits
If iray%(2) < 1 Or iray%(2) > MAXRAY% Then GoTo Penepma12CalculateKratiosOutofLimits
If eO! < 1# Or eO! > 100# Then GoTo Penepma12CalculateKratiosOutofLimits
If TOA! < 1# Or TOA! > 90# Then GoTo Penepma12CalculateKratiosOutofLimits
If conc!(1) < 0# Or conc!(1) > 1# Then GoTo Penepma12CalculateKratiosOutofLimits
If conc!(2) < 0# Or conc!(2) > 1# Then GoTo Penepma12CalculateKratiosOutofLimits
If kexp!(1) < 0# Or kexp!(1) > 1# Then GoTo Penepma12CalculateKratiosOutofLimits
If kexp!(2) < 0# Or kexp!(2) > 1# Then GoTo Penepma12CalculateKratiosOutofLimits

' Check that both elements are not by difference
If iray%(1) = MAXRAY% And iray%(2) = MAXRAY% Then GoTo Penepma12CalculateKratiosBothByDifference

' Check that at least one concentration is entered
If conc!(1) = 0# And conc!(2) = 0# Then GoTo Penepma12CalculateKratiosNoConcData

' Check for valid kexp data if x-ray used
If iray%(1) <= MAXRAY% - 1 And kexp!(1) = 0# Then GoTo Penepma12CalculateKratiosNoKexpData
If iray%(2) <= MAXRAY% - 1 And kexp!(2) = 0# Then GoTo Penepma12CalculateKratiosNoKexpData

' Just load first composition
If ncmp% = 0 Then
ncmp% = ncmp% + 1
ReDim ksym(1 To 2, ncmp%) As Integer
ReDim kcmp(1 To 2, ncmp%) As Single
ksym%(1, ncmp%) = isym%(1)
ksym%(2, ncmp%) = isym%(2)
kcmp!(1, ncmp%) = conc!(1)
kcmp!(2, ncmp%) = conc!(2)

' See if composition is already loaded
Else
For n% = 1 To ncmp%
If isym%(1) = ksym%(1, n%) And conc!(1) = kcmp!(1, n%) And isym%(2) = ksym%(2, n%) And conc!(2) = kcmp!(2, n%) Then
GoTo Penepma12CalculateKratiosNextLine
End If
Next n%

' Composition was not found, so load in comp array
ncmp% = ncmp% + 1
ReDim Preserve ksym(1 To 2, ncmp%) As Integer
ReDim Preserve kcmp(1 To 2, ncmp%) As Single
ksym%(1, ncmp%) = isym%(1)
ksym%(2, ncmp%) = isym%(2)
kcmp!(1, ncmp%) = conc!(1)
kcmp!(2, ncmp%) = conc!(2)
End If

Penepma12CalculateKratiosNextLine:
Loop
Close #ImportDataFileNumber%

' Save total number of binaries
BinaryLineTotal& = BinaryLineCount&

' Save parameters from form
Call Penepma12Save
If ierror Then
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
Exit Sub
End If

' Check for existing PAR files
For n% = 1 To ncmp%

PENEPMA_Sample(1).ElmPercents!(1) = 100# * kcmp!(1, n%)
If kcmp!(2, n%) = 0# Then kcmp!(2, n%) = 1# - kcmp!(1, n%)
PENEPMA_Sample(1).ElmPercents!(2) = 100# * kcmp!(2, n%)

' Create material (and PAR) file name
binaryname$ = Trim$(Symup$(ksym%(1, n%))) & "-" & Format$(PENEPMA_Sample(1).ElmPercents!(1)) & "_" & Trim$(Symup$(ksym%(2, n%))) & "-" & Format$(PENEPMA_Sample(1).ElmPercents!(2))
pfilename$ = PENEPMA_Root$ & "\Penfluor\" & binaryname$ & ".par"

If Not CalculateDoNotOverwritePAR Or (CalculateDoNotOverwritePAR And Dir$(pfilename$) = vbNullString) Then
MaterialFileA$ = binaryname$ & ".mat"
Call Penepma12RunPenfluorCheck(Int(1))
If ierror Then
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
Exit Sub
End If
End If

Next n%

TotalNumberOfSimulations& = ncmp%
CurrentSimulationsNumber& = 1

' Check calculation time with user
Call Penepma12RunPenfluorCheck2(Int(1))
If ierror Then
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
Exit Sub
End If

' Now calculate unique compositions using Material.exe and Penfluor.exe
For n% = 1 To ncmp%

' Load in PENEPMA_Sample()
PENEPMA_Sample(1).LastElm% = 2
PENEPMA_Sample(1).LastChan% = 2

PENEPMA_Sample(1).ElmPercents!(1) = 100# * kcmp!(1, n%)
If kcmp!(2, n%) = 0# Then kcmp!(2, n%) = 1# - kcmp!(1, n%)
PENEPMA_Sample(1).ElmPercents!(2) = 100# * kcmp!(2, n%)

PENEPMA_Sample(1).AtomicNums%(1) = ksym%(1, n%)
PENEPMA_Sample(1).AtomicNums%(2) = ksym%(2, n%)

' Create material (and PAR) file name
binaryname$ = Trim$(Symup$(ksym%(1, n%))) & "-" & Format$(PENEPMA_Sample(1).ElmPercents!(1)) & "_" & Trim$(Symup$(ksym%(2, n%))) & "-" & Format$(PENEPMA_Sample(1).ElmPercents!(2))
MaterialFileA$ = binaryname$ & ".mat"

PENEPMA_Sample(1).Name$ = binaryname$

' Check for existing .PAR file
pfilename$ = PENEPMA_Root$ & "\Penfluor\" & binaryname$ & ".par"
If Not CalculateDoNotOverwritePAR Or (CalculateDoNotOverwritePAR And Dir$(pfilename$) = vbNullString) Then

FormPENEPMA12.LabelProgress.Caption = "Creating Material File " & MaterialFileA$
FormPENEPMA12.LabelRemainingTime.Caption = vbNullString

msg$ = "Calculating material file (" & MaterialFileA$ & ") for binary composition (" & Str$(n%) & " of " & Format$(ncmp%) & ")..."
Call IOStatusAuto(msg$)
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(0), MaterialInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

' Make material INP file
MaterialDensityA# = 5#      ' bulk material so density doesn't matter
Screen.MousePointer = vbHourglass
Call Penepma12CreateMaterialINP(Int(1), PENEPMA_Sample())
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

' Create and run the necessary batch files
Screen.MousePointer = vbHourglass
Call Penepma12CreateMaterialBatch(Int(1), Int(1))
Screen.MousePointer = vbDefault
If ierror Then
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
Exit Sub
End If

' PAR file already exists
Else
msg$ = pfilename$ & " file already exists and Material calculations will be skipped..."
Call IOWriteLog(msg$)
End If

' Check for existing .PAR file
pfilename$ = PENEPMA_Root$ & "\Penfluor\" & binaryname$ & ".par"
If Not CalculateDoNotOverwritePAR Or (CalculateDoNotOverwritePAR And Dir$(pfilename$) = vbNullString) Then

' Run Penfluor and Fitall on material A
Call Penepma12RunPenFluor(Int(1))
If ierror Then
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
Exit Sub
End If

' PAR file already exists
Else
msg$ = pfilename$ & " file already exists and Penfluor calculations will be skipped..."
Call IOWriteLog(msg$)
End If

CurrentSimulationsNumber& = CurrentSimulationsNumber& + 1
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(0), MaterialInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

Next n%

' Now extract k-ratios
msg$ = vbCrLf & "Penepma12CalculateKratios: All PAR file calculations for " & ImportDataFile$ & " are complete, now extracting k-ratios..."
Call IOWriteLog(msg$)

' Output file and write column labels for Fanal results
Open ExportDataFile$ For Output As #ExportDataFileNumber%
    Print #ExportDataFileNumber%, " ", VbDquote$ & "Line" & VbDquote$, vbTab, VbDquote$ & "SymA" & VbDquote$, vbTab, VbDquote$ & "SymB" & VbDquote$, vbTab, VbDquote$ & "RayA" & VbDquote$, vbTab, VbDquote$ & "RayB" & VbDquote$, vbTab, _
    VbDquote$ & "KeV" & VbDquote$, vbTab, VbDquote$ & "Takeoff" & VbDquote$, vbTab, VbDquote$ & "ConcA" & VbDquote$, vbTab, VbDquote$ & "ConcB" & VbDquote$, vbTab, _
    VbDquote$ & "KexpA" & VbDquote$, vbTab, VbDquote$ & "KexpB" & VbDquote$, vbTab, VbDquote$ & "KratA" & VbDquote$, vbTab, VbDquote$ & "KratB" & VbDquote$, vbTab, _
    VbDquote$ & "KerrA" & VbDquote$, vbTab, VbDquote$ & "KerrB" & VbDquote$

' Open input file again
Open ImportDataFile$ For Input As #ImportDataFileNumber%
BinaryLineCount& = 0
Call IOStatusAuto(vbNullString)

' Extract k-ratios for all lines in input file
Do While Not EOF(ImportDataFileNumber%)
BinaryLineCount& = BinaryLineCount& + 1

' Read binary elements, kilovolts and takeoff
Input #ImportDataFileNumber%, isym%(1), isym%(2), iray%(1), iray%(2), eO!, TOA!, conc!(1), conc!(2), kexp!(1), kexp!(2)
If conc!(2) = 0# Then conc!(2) = 1# - conc!(1)

' Find matching composition previously calculated (PAR file)
For n% = 1 To ncmp%
If isym%(1) = ksym%(1, n%) And conc!(1) = kcmp!(1, n%) And isym%(2) = ksym%(2, n%) And conc!(2) = kcmp!(2, n%) Then
GoTo Penepma12CalculateKratiosNextLine2:
End If
Next n%
GoTo Penepma12CalculateKratiosCompNotFound

' Extract k-ratio using Fanal
Penepma12CalculateKratiosNextLine2:

' Calculate for both emitters
For ibin% = 1 To 2
If iray%(ibin%) < MAXRAY% Then  ' skip if x-ray is "not analyzed"

Call XrayGetEnergy(isym%(ibin%), iray%(ibin%), eng!, edg!)
If ierror Then Exit Sub

If ibin% = 1 Then msg$ = "Extracting k-ratios from binary " & Format$(BinaryLineCount&) & " of " & Format$(BinaryLineTotal&) & " (" & Symlo$(isym%(ibin%)) & " " & Xraylo$(iray%(ibin%)) & " in " & Symlo$(isym%(2)) & " at " & Format$(eO!) & " kev)..."
If ibin% = 2 Then msg$ = "Extracting k-ratios from binary " & Format$(BinaryLineCount&) & " of " & Format$(BinaryLineTotal&) & " (" & Symlo$(isym%(ibin%)) & " " & Xraylo$(iray%(ibin%)) & " in " & Symlo$(isym%(1)) & " at " & Format$(eO!) & " kev)..."
Call IOStatusAuto(msg$)
If icancelauto Then
Call IOStatusAuto(vbNullString)
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
ierror = True
Exit Sub
End If

' Specify the Fanal parameters
MaterialMeasuredTakeoff# = TOA!
MaterialMeasuredEnergy# = eO!

MaterialMeasuredElement% = isym%(ibin%)
MaterialMeasuredXray% = iray%(ibin%)

' Create parameter file names
PENEPMA_Sample(1).ElmPercents!(1) = 100# * kcmp!(1, n%)
If kcmp!(2, n%) = 0# Then kcmp!(2, n%) = 1# - kcmp!(1, n%)
PENEPMA_Sample(1).ElmPercents!(2) = 100# * kcmp!(1, n%)
PENEPMA_Sample(1).ElmPercents!(2) = 100# * kcmp!(2, n%)

' Create material (and PAR) file name
binaryname$ = Trim$(Symup$(ksym%(1, n%))) & "-" & Format$(PENEPMA_Sample(1).ElmPercents!(1)) & "_" & Trim$(Symup$(ksym%(2, n%))) & "-" & Format$(PENEPMA_Sample(1).ElmPercents!(2))

ParameterFileA$ = binaryname$ & ".par"
ParameterFileB$ = binaryname$ & ".par"                          ' same as A for matrix calculations
ParameterFileBStd$ = Trim$(Symup$(isym%(ibin%))) & ".par"       ' use pure element always (use Trim$ for single letter elements)

' Check for pure element PAR file in Penfluor\Pure folder
If Dir$(PENEPMA_Root$ & "\Penfluor\" & ParameterFileBStd$) = vbNullString Then
tfilename$ = PENEPMA_Root$ & "\Penfluor\Pure\" & ParameterFileBStd$
If Dir$(tfilename$) <> vbNullString Then FileCopy tfilename$, PENEPMA_Root$ & "\Penfluor\" & ParameterFileBStd$
If Dir$(MiscGetFileNameNoExtension$(tfilename$) & ".in") <> vbNullString Then FileCopy MiscGetFileNameNoExtension$(tfilename$) & ".in", PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(ParameterFileBStd$)) & ".in"
End If

' Check the parameters files
Call Penepma12RunFanal
If ierror Then
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
Exit Sub
End If

' Run the Fanal program
tMaterialMeasuredGridPoints% = MaterialMeasuredGridPoints%      ' save
MaterialMeasuredGridPoints% = 1     ' use a single point for matrix calculations
Call Penepma12RunFanal1
MaterialMeasuredGridPoints% = tMaterialMeasuredGridPoints%      ' restore
If ierror Then
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
Exit Sub
End If

' Get k-ratio data from Fanal k-ratio file
Call Penepma12LoadPlotData
If ierror Then
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
Exit Sub
End If

' Save data (only dimension arrays for first binary)
If ibin% = 1 Then
ReDim Preserve BinaryLine(1 To 2, 1 To BinaryLineCount&) As Integer
ReDim Preserve BinaryEsym(1 To 2, 1 To BinaryLineCount&) As String
ReDim Preserve BinaryXsym(1 To 2, 1 To BinaryLineCount&) As String
ReDim Preserve BinaryConc(1 To 2, 1 To BinaryLineCount&) As Single
ReDim Preserve BinaryT0E0(1 To 2, 1 To BinaryLineCount&) As Single
ReDim Preserve BinaryKexp(1 To 2, 1 To BinaryLineCount&) As Single
ReDim Preserve BinaryKratio(1 To 2, 1 To BinaryLineCount&) As Single
ReDim Preserve BinaryKerror(1 To 2, 1 To BinaryLineCount&) As Single
End If

BinaryLine%(ibin%, BinaryLineCount&) = BinaryLineCount&

BinaryEsym$(1, BinaryLineCount&) = Symlo$(isym%(1)) ' same for both binaries
BinaryEsym$(2, BinaryLineCount&) = Symlo$(isym%(2)) ' same for both binaries

BinaryXsym$(1, BinaryLineCount&) = Xraylo$(iray%(1)) ' same for both binaries
BinaryXsym$(2, BinaryLineCount&) = Xraylo$(iray%(2)) ' same for both binaries

BinaryConc!(1, BinaryLineCount&) = kcmp!(1, n%) ' same for both binaries
BinaryConc!(2, BinaryLineCount&) = kcmp!(2, n%) ' same for both binaries

BinaryT0E0!(1, BinaryLineCount&) = TOA!  ' same for both binaries
BinaryT0E0!(2, BinaryLineCount&) = eO!  ' same for both binaries

BinaryKexp!(ibin%, BinaryLineCount&) = kexp!(ibin%)
BinaryKratio!(ibin%, BinaryLineCount&) = yktotal#(1) / 100#  ' from bulk Fanal calculation (convert to normal k-ratio)

' Calculate k-ratio error
If kexp!(ibin%) <> 0# Then BinaryKerror!(ibin%, BinaryLineCount&) = BinaryKratio!(ibin%, BinaryLineCount&) / kexp!(ibin%)

' Output binary k-ratio results
    Print #ExportDataFileNumber%, BinaryLineCount&, vbTab, isym%(1), vbTab, isym%(2), vbTab, iray%(1), vbTab, iray%(2), vbTab, _
    eO!, vbTab, TOA!, vbTab, conc!(1), vbTab, conc!(2), vbTab, kexp!(1), vbTab, kexp!(2), vbTab, _
    BinaryKratio!(1, BinaryLineCount&), vbTab, BinaryKratio!(2, BinaryLineCount&), vbTab, _
    BinaryKerror!(1, BinaryLineCount&), vbTab, BinaryKerror!(2, BinaryLineCount&)
            
If DebugMode Then
Call IOWriteLog(vbNullString)
msg$ = Space$(4) & Format$("Conc", a80$) & Format$("K-Exp", a80$) & Format$("K-Rat", a80$) & Format$("K-Err", a80$)
Call IOWriteLog(msg$)
msg$ = Symup$(isym%(1)) & " " & Xraylo$(iray%(1)) & MiscAutoFormat$(conc!(1)) & MiscAutoFormat$(kexp!(1)) & MiscAutoFormat$(BinaryKratio!(1)) & MiscAutoFormat$(BinaryKerror!(1, BinaryLineCount&))
Call IOWriteLog(msg$)

If iray%(2) <> MAXRAY% Then
msg$ = Symup$(isym%(2)) & " " & Xraylo$(iray%(2)) & MiscAutoFormat$(conc!(2)) & MiscAutoFormat$(kexp!(2)) & MiscAutoFormat$(BinaryKratio!(2)) & MiscAutoFormat$(BinaryKerror!(2, BinaryLineCount&))
Call IOWriteLog(msg$)
End If
End If

End If

' Check for Pause button
Do Until Not RealTimePauseAutomation
DoEvents
Sleep 200
Loop

Next ibin%
Loop

' Store line total (again since skipbelow1keV flag may have been utilized)
BinaryLineTotal& = BinaryLineCount&

' Check if any binaries were calculated
If BinaryLineCount& < 1 Then
msg$ = "No binaries were calculated."
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12CalculateKratios"
Exit Sub
End If

' Print out problematic lines
Call IOWriteLog(vbNullString)
msg$ = "Problematic k-ratio errors (< 0.8 or > 1.2)"
Call IOWriteLog(msg$)
msg$ = Format$("Line", a60$) & Space$(10) & Format$("ConcA", a80$) & Format$("ConcB", a80$) & Format$("TOA", a80$) & Format$("eO", a80$) & Format$("K-Exp", a80$) & Format$("K-Cal", a80$) & Format$("K-Err", a80$)
Call IOWriteLog(msg$)

For n% = 1 To BinaryLineTotal&
If BinaryKexp!(1, n%) <> 0# Then
If VerboseMode Or (BinaryKerror!(1, n%) < 0.8 Or BinaryKerror!(1, n%) > 1.2) Then
msg$ = Format$(BinaryLine%(1, n%), "!@@@@@@") & BinaryEsym$(1, n%) & " " & BinaryXsym$(1, n%) & " in " & BinaryEsym$(2, n%) & MiscAutoFormat$(BinaryConc!(1, n%)) & MiscAutoFormat$(BinaryConc!(2, n%)) & MiscAutoFormat$(BinaryT0E0!(1, n%)) & MiscAutoFormat$(BinaryT0E0!(2, n%)) & MiscAutoFormat$(BinaryKexp!(1, n%)) & MiscAutoFormat$(BinaryKratio!(1, n%)) & MiscAutoFormat$(BinaryKerror!(1, n%))
Call IOWriteLog(msg$)
End If
End If

If BinaryKexp!(2, n%) <> 0# Then
If VerboseMode Or (BinaryKerror!(2, n%) < 0.8 Or BinaryKerror!(2, n%) > 1.2) Then
msg$ = Format$(BinaryLine%(2, n%), "!@@@@@@") & BinaryEsym$(2, n%) & " " & BinaryXsym$(2, n%) & " in " & BinaryEsym$(1, n%) & MiscAutoFormat$(BinaryConc!(1, n%)) & MiscAutoFormat$(BinaryConc!(2, n%)) & MiscAutoFormat$(BinaryT0E0!(1, n%)) & MiscAutoFormat$(BinaryT0E0!(2, n%)) & MiscAutoFormat$(BinaryKexp!(2, n%)) & MiscAutoFormat$(BinaryKratio!(2, n%)) & MiscAutoFormat$(BinaryKerror!(2, n%))
Call IOWriteLog(msg$)
End If
End If
Next n%

' Calculate average and standard deviation
Call MathArrayAverage3(average, BinaryKerror!(), BinaryLineCount&, 2)
If ierror Then
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
Exit Sub
End If

' Write to file
Print #ExportDataFileNumber%, " "
Print #ExportDataFileNumber%, VbDquote$ & "AverageA" & VbDquote$, vbTab, MiscAutoFormat$(average.averags!(1))
Print #ExportDataFileNumber%, VbDquote$ & "StdDevA" & VbDquote$, vbTab, MiscAutoFormat$(average.Stddevs!(1))
Print #ExportDataFileNumber%, VbDquote$ & "MinimumA" & VbDquote$, vbTab, MiscAutoFormat$(average.Minimums!(1))
Print #ExportDataFileNumber%, VbDquote$ & "MaximumA" & VbDquote$, vbTab, MiscAutoFormat$(average.Maximums!(1))
Print #ExportDataFileNumber%, VbDquote$ & "AverageB" & VbDquote$, vbTab, MiscAutoFormat$(average.averags!(2))
Print #ExportDataFileNumber%, VbDquote$ & "StdDevB" & VbDquote$, vbTab, MiscAutoFormat$(average.Stddevs!(2))
Print #ExportDataFileNumber%, VbDquote$ & "MinimumB" & VbDquote$, vbTab, MiscAutoFormat$(average.Minimums!(2))
Print #ExportDataFileNumber%, VbDquote$ & "MaximumB" & VbDquote$, vbTab, MiscAutoFormat$(average.Maximums!(2))

Call IOWriteLog(vbNullString)
Call IOWriteLog("AverageA" & MiscAutoFormat$(average.averags!(1)))
Call IOWriteLog("StdDevA" & MiscAutoFormat$(average.Stddevs!(1)))
Call IOWriteLog("MinimumA" & MiscAutoFormat$(average.Minimums!(1)))
Call IOWriteLog("MaximumA" & MiscAutoFormat$(average.Maximums!(1)))
Call IOWriteLog("AverageB" & MiscAutoFormat$(average.averags!(2)))
Call IOWriteLog("StdDevB" & MiscAutoFormat$(average.Stddevs!(2)))
Call IOWriteLog("MinimumB" & MiscAutoFormat$(average.Minimums!(2)))
Call IOWriteLog("MaximumB" & MiscAutoFormat$(average.Maximums!(2)))

' Close files
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%

Call IOStatusAuto(vbNullString)
msg$ = "Binary Penfluor/Fanal calculations completed on file " & ImportDataFile$ & vbCrLf
msg$ = msg$ & "Data output saved to " & ExportDataFile$ & vbCrLf
Call IOWriteLog(vbCrLf & vbCrLf & msg$)
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12CalculateKratios"

Exit Sub

' Errors
Penepma12CalculateKratiosError:
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CalculateKratios"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12CalculateKratiosOutofLimits:
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
msg$ = "Bad data on line " & Str$(BinaryLineCount&) & " in " & ImportDataFile$ & " (file format may be wrong)."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CalculateKratios"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12CalculateKratiosBothByDifference:
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
msg$ = "Both elements are by difference on line " & Str$(BinaryLineCount&) & " in " & ImportDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CalculateKratios"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12CalculateKratiosNoConcData:
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
msg$ = "No Conc data on line " & Str$(BinaryLineCount&) & " in " & ImportDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CalculateKratios"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12CalculateKratiosNoKexpData:
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
msg$ = "No K-exp data on line " & Str$(BinaryLineCount&) & " in " & ImportDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CalculateKratios"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12CalculateKratiosCompNotFound:
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
msg$ = "No matching composition for Fanal binary extraction was found for line " & Str$(BinaryLineCount&) & " in " & ImportDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CalculateKratios"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12AdjustMinimumEnergy(mfilename As String)
' Check for Li, Be, B, C, N, O, F or Ne and adjust minimum energy if so

ierror = False
On Error GoTo Penepma12AdjustMinimumEnergyError

Dim ip As Integer

' Check .MAT file for light elements
Call Penepma12GetMatFileComposition(mfilename$, Penepma_TmpSample())
If ierror Then Exit Sub

' Load adjustment
PenepmaMinimumElectronEnergy! = 1#      ' assume default 1 keV to start
ip% = IPOS1%(Penepma_TmpSample(1).LastChan%, Symlo$(ATOMIC_NUM_NEON%), Penepma_TmpSample(1).Elsyms$())    ' Ne
If ip% > 0 Then PenepmaMinimumElectronEnergy! = 0.8
ip% = IPOS1%(Penepma_TmpSample(1).LastChan%, Symlo$(ATOMIC_NUM_FLUORINE%), Penepma_TmpSample(1).Elsyms$())     ' F
If ip% > 0 Then PenepmaMinimumElectronEnergy! = 0.6
ip% = IPOS1%(Penepma_TmpSample(1).LastChan%, Symlo$(ATOMIC_NUM_OXYGEN%), Penepma_TmpSample(1).Elsyms$())     ' O
If ip% > 0 Then PenepmaMinimumElectronEnergy! = 0.5
ip% = IPOS1%(Penepma_TmpSample(1).LastChan%, Symlo$(ATOMIC_NUM_NITROGEN%), Penepma_TmpSample(1).Elsyms$())     ' N
If ip% > 0 Then PenepmaMinimumElectronEnergy! = 0.3
ip% = IPOS1%(Penepma_TmpSample(1).LastChan%, Symlo$(ATOMIC_NUM_CARBON%), Penepma_TmpSample(1).Elsyms$())     ' C
If ip% > 0 Then PenepmaMinimumElectronEnergy! = 0.2
ip% = IPOS1%(Penepma_TmpSample(1).LastChan%, Symlo$(ATOMIC_NUM_BORON%), Penepma_TmpSample(1).Elsyms$())     ' B
If ip% > 0 Then PenepmaMinimumElectronEnergy! = 0.1
ip% = IPOS1%(Penepma_TmpSample(1).LastChan%, Symlo$(ATOMIC_NUM_BERYLLIUM%), Penepma_TmpSample(1).Elsyms$())     ' Be
If ip% > 0 Then PenepmaMinimumElectronEnergy! = 0.05
ip% = IPOS1%(Penepma_TmpSample(1).LastChan%, Symlo$(ATOMIC_NUM_LITHIUM%), Penepma_TmpSample(1).Elsyms$())     ' Li
If ip% > 0 Then PenepmaMinimumElectronEnergy! = 0.02

'ip% = IPOS1%(Penepma_TmpSample(1).LastChan%, Symlo$(ATOMIC_NUM_HELIUM%), Penepma_TmpSample(1).Elsyms$())    ' He
'If ip% > 0 Then PenepmaMinimumElectronEnergy! = 0.01
'ip% = IPOS1%(Penepma_TmpSample(1).LastChan%, Symlo$(ATOMIC_NUM_HYDROGEN%), Penepma_TmpSample(1).Elsyms$())    ' H
'If ip% > 0 Then PenepmaMinimumElectronEnergy! = 0.00

' Update minimum energy field on form
FormPENEPMA12.TextPenepmaMinimumElectronEnergy.Text = Format$(PenepmaMinimumElectronEnergy!)

Exit Sub

' Errors
Penepma12AdjustMinimumEnergyError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12AdjustMinimumEnergy"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12AdjustMinimumEnergy2(esym As String)
' Check for Li, Be, B, C, N, O, F or Ne and adjust minimum energy if so

ierror = False
On Error GoTo Penepma12AdjustMinimumEnergy2Error

' Load adjustment
PenepmaMinimumElectronEnergy! = 1#      ' assume default 1 keV to start
If MiscStringsAreSame(esym$, Symlo$(ATOMIC_NUM_NEON%)) Then PenepmaMinimumElectronEnergy! = 0.8   ' Ne
If MiscStringsAreSame(esym$, Symlo$(ATOMIC_NUM_FLUORINE%)) Then PenepmaMinimumElectronEnergy! = 0.6   ' F
If MiscStringsAreSame(esym$, Symlo$(ATOMIC_NUM_OXYGEN%)) Then PenepmaMinimumElectronEnergy! = 0.5   ' O
If MiscStringsAreSame(esym$, Symlo$(ATOMIC_NUM_NITROGEN%)) Then PenepmaMinimumElectronEnergy! = 0.3   ' N
If MiscStringsAreSame(esym$, Symlo$(ATOMIC_NUM_CARBON%)) Then PenepmaMinimumElectronEnergy! = 0.2   ' C
If MiscStringsAreSame(esym$, Symlo$(ATOMIC_NUM_BORON%)) Then PenepmaMinimumElectronEnergy! = 0.1   ' B
If MiscStringsAreSame(esym$, Symlo$(ATOMIC_NUM_BERYLLIUM%)) Then PenepmaMinimumElectronEnergy! = 0.05   ' Be
If MiscStringsAreSame(esym$, Symlo$(ATOMIC_NUM_LITHIUM%)) Then PenepmaMinimumElectronEnergy! = 0.02   ' Li

'If MiscStringsAreSame(esym$, Symlo$(ATOMIC_NUM_HELIUM%)) Then PenepmaMinimumElectronEnergy! = 0.01   ' He
'If MiscStringsAreSame(esym$, Symlo$(ATOMIC_NUM_HYDROGEN%)) Then PenepmaMinimumElectronEnergy! = 0.00   ' H

' Update minimum energy field on form
FormPENEPMA12.TextPenepmaMinimumElectronEnergy.Text = Format$(PenepmaMinimumElectronEnergy!)

Exit Sub

' Errors
Penepma12AdjustMinimumEnergy2Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12AdjustMinimumEnergy2"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12RunFanalCheckINFile(mstring As String, ifilename As String, pstring As String)
' Check the .PAR file's corresponding .IN file and return the specified parameter
'  mstring$ = "MSIMPA" check minimum electron photon energy
'  mstring$ = "TIME" check material simulation time
'  ifilename$ = penfluor.in parameter file
'  pstring$ = returned parameter as string
'
'    Sample Penfluor.in file (based on Penfluor.inp):
'
'TITLE  A thick cylindrical target
'       .  Change only the material filename. Keep the rest unaltered.
'       >>>>>>>> Electron beam definition.
'SENERG 100e3                     [Energy of the electron beam, in eV]
'SPOSIT 0 0 1                     [Coordinates of the electron source]
'SDIREC 180 0              [Direction angles of the beam axis, in deg]
'SAPERT 0                                      [Beam aperture, in deg]
'       .
'       >>>>>>>> Material data and simulation parameters.
'                Up to 10 materials; 2 lines for each material.
'MFNAME Fe.mat                         [Material file, up to 20 chars]
'MSIMPA 1e3 1e3 1e3 0.2 0.2 1e3 1e3          [EABS(1:3),C1,C2,WCC,WCR]
'       .
'       >>>>>>>> Geometry of the sample.
'GEOMFN penfluor.geo              [Geometry definition file, 20 chars]
'DSMAX  1 1.0e-4             [IB, Maximum step length (cm) in body IB]
'       .
'       >>>>>>>> Interaction forcing.
'IFORCE 1 1 4 -10    0.1 1.0           [KB,KPAR,ICOL,FORCER,WLOW,WHIG]
'IFORCE 1 1 5 -400   0.1 1.0           [KB,KPAR,ICOL,FORCER,WLOW,WHIG]
'       .
'       >>>>>>>> Emerging particles. Energy and angular distributions.
'NBE    0.0 0.0 100                [E-interval and no. of energy bins]
'NBTH   45                     [No. of bins for the polar angle THETA]
'NBPH   30                   [No. of bins for the azimuthal angle PHI]
'       .
'       >>>>>>>> Photon detectors (up to 25 different detectors).
'                IPSF=0, do not create a phase-space file.
'                IPSF=1, creates a phase-space file.
'PDANGL 0 90  0 360 0                   [Angular window, in deg, IPSF]
'PDENER 0.0 0.0 100                   [Energy window, no. of channels]
'       .
'NSIMSH 2.0e9                    [Desired number of simulated showers]
'TIME   3600                        [Allotted simulation time, in sec]

ierror = False
On Error GoTo Penepma12RunFanalCheckINFileError

Dim astring As String, bstring As String, cstring As String
Dim EABS(1 To 3) As Double, c1 As Double, c2 As Double, WCC As Double, WCR As Double

' Check if file exists (if not use defaults)
pstring$ = vbNullString
If Dir$(ifilename$) = vbNullString Then
If mstring$ = "MSIMPA" Then pstring$ = Format$(PenepmaMinimumElectronEnergy!)
If mstring$ = "TIME" Then pstring$ = Format$(MaterialSimulationTime#)
If pstring$ = vbNullString Then GoTo Penepma12RunFanalCheckINFileBadParameter
Exit Sub
End If

' Loop through sample input file and find specified parameter
Open ifilename$ For Input As #Temp1FileNumber%

Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$
If astring$ <> vbNullString Then
Call MiscParseStringToString(astring$, bstring$)    ' get parameter type
If ierror Then Exit Sub

' Parse out MSIMPA parameters: MSIMPA 1e3 1e3 1e3 0.2 0.2 1e3 1e3          [EABS(1:3),C1,C2,WCC,WCR]
If mstring$ = "MSIMPA" Then
If MiscStringsAreSame(bstring$, mstring$) Then
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
pstring$ = bstring$     ' just return the first parameter for now

EABS#(1) = Val(Trim$(bstring$))
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
EABS#(2) = Val(Trim$(bstring$))
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
EABS#(3) = Val(Trim$(bstring$))

Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
c1# = Val(Trim$(bstring$))
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
c2# = Val(Trim$(bstring$))

Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
WCC# = Val(Trim$(bstring$))
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
WCR# = Val(Trim$(bstring$))

' Exit with module level variables loaded
cstring$ = Format$(EABS#(1), "0.0E+0") & " " & Format$(EABS#(2), "0.0E+0") & " " & Format$(EABS#(3), "0E+0") & " "
cstring$ = cstring$ & Format$(c1#, "0.0") & " " & Format$(c2#, "0.0") & " "
cstring$ = cstring$ & Format$(WCC#, "0E+0") & " " & Format$(WCR#, "0E+0")
If DebugMode Then Call IOWriteLog(cstring$)
Close #Temp1FileNumber%
Exit Sub
End If
End If

' Parse out TIME parameter: TIME   3600                        [Allotted simulation time, in sec]
If mstring$ = "TIME" Then
If MiscStringsAreSame(bstring$, mstring$) Then
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
pstring$ = bstring$     ' return time parameter
Close #Temp1FileNumber%
Exit Sub
End If
End If

End If
Loop

' If we get to here, we did not find the indicated parameter
msg$ = "Passed input file parameter (" & mstring$ & ") was not found in " & ifilename$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12RunFanalCheckINFile"

Close #Temp1FileNumber%
Exit Sub

' Errors
Penepma12RunFanalCheckINFileError:
Close #Temp1FileNumber%
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12RunFanalCheckINFile"
ierror = True
Exit Sub

Penepma12RunFanalCheckINFileBadParameter:
MsgBox "Bad parameter (" & mstring$ & ") was passed", vbOKOnly + vbExclamation, "Penepma12RunFanalCheckINFile"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12LoadPlotData()
' Load the k-ratio graph data for the k-ratio plot

ierror = False
On Error GoTo Penepma12LoadPlotDataError

Dim astring As String, bstring As String

' Make sure input file is closed
Close #Temp1FileNumber%

ReDim xdist(1 To 1) As Double       ' linear distance (um)
ReDim mdist(1 To 1) As Double       ' mass distance (ug/cm^2)
ReDim yktotal(1 To 1) As Double      ' total fluorescence kratio% plus primary x-ray kratio%
ReDim ykfluor(1 To 1) As Double      ' fluorescence kratio% only

ReDim yctotal(1 To 1) As Double      ' apparent concentration from fluorescence and primary x-ray (just dimension)
ReDim ycA_only(1 To 1) As Double     ' apparent concentration from A fluorescence only (just dimension)
ReDim ycb_only(1 To 1) As Double     ' apparent concentration from B fluorescence only (just dimension)
ReDim yc_prix(1 To 1) As Double      ' apparent concentration from primary x-rays only (just dimension)

ReDim yktotal_meas(1 To 1) As Double  ' "measured" k-ratio % from total intensity (just dimension)
ReDim yztotal_meas(1 To 1) As Double  ' "measured" ZAF correction from total intensity (just dimension)
ReDim yctotal_meas(1 To 1) As Double  ' "measured" concentration % from total intensity (just dimension)

ReDim flach(1 To 1) As Double        ' Mat A characteristic fluorescence
ReDim flabr(1 To 1) As Double        ' Mat A continuum fluorescence
ReDim flbch(1 To 1) As Double        ' Mat B characteristic fluorescence
ReDim flbbr(1 To 1) As Double        ' Mat B continuum fluorescence
ReDim pri_int(1 To 1) As Double      ' primary x-ray intensity
ReDim std_int(1 To 1) As Double      ' standard intensity

ReDim fluA_k(1 To 1) As Double        ' Mat A total fluorescence k-ratio % (just dimension)
ReDim fluB_k(1 To 1) As Double        ' Mat B total fluorescence k-ratio % (just dimension)
ReDim prix_k(1 To 1) As Double        ' Primary x-ray k-ratio % (just dimension)

If Dir$(Trim$(KRATIOS_DAT_File$)) = vbNullString Then Exit Sub
Open KRATIOS_DAT_File$ For Input As #Temp1FileNumber%

' Load array (npoints&, xdist#(), yktotal#(), ykfluor#())
nPoints& = 0
Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$
If Len(Trim$(astring$)) > 0 And InStr(astring$, "#") = 0 Then            ' skip to first data line (also skips if value is -1.#IND0E+000)

' Load k-ratio data
nPoints& = nPoints& + 1
Call MiscParseStringToString(astring$, bstring$)    ' throw away eV data (1st column)
If ierror Then Exit Sub

Call MiscParseStringToString(astring$, bstring$)    ' get distance data (microns)
If ierror Then Exit Sub

ReDim Preserve xdist(1 To nPoints&) As Double       ' linear distance (um)
ReDim Preserve mdist(1 To nPoints&) As Double       ' mass distance (ug/cm^2)
ReDim Preserve yktotal(1 To nPoints&) As Double      ' total fluorescence kratio% plus primary x-ray kratio%
ReDim Preserve ykfluor(1 To nPoints&) As Double      ' fluorescence kratio% only

ReDim Preserve yctotal(1 To nPoints&) As Double      ' apparent concentration from fluorescence and primary x-ray (just dimension)
ReDim Preserve ycA_only(1 To nPoints&) As Double     ' apparent concentration from A fluorescence only (just dimension)
ReDim Preserve ycb_only(1 To nPoints&) As Double     ' apparent concentration from B fluorescence only (just dimension)
ReDim Preserve yc_prix(1 To nPoints&) As Double      ' apparent concentration from primary x-rays only (just dimension)

ReDim Preserve yktotal_meas(1 To nPoints&) As Double  ' "measured" k-ratio % from total intensity (just dimension)
ReDim Preserve yztotal_meas(1 To nPoints&) As Double  ' "measured" ZAF correction from total intensity (just dimension)
ReDim Preserve yctotal_meas(1 To nPoints&) As Double  ' "measured" concentration % from total intensity (just dimension)

ReDim Preserve flach(1 To nPoints&) As Double        ' Mat A characteristic fluorescence
ReDim Preserve flabr(1 To nPoints&) As Double        ' Mat A continuum fluorescence
ReDim Preserve flbch(1 To nPoints&) As Double        ' Mat B characteristic fluorescence
ReDim Preserve flbbr(1 To nPoints&) As Double        ' Mat B continuum fluorescence
ReDim Preserve pri_int(1 To nPoints&) As Double      ' primary x-ray intensity
ReDim Preserve std_int(1 To nPoints&) As Double      ' standard intensity

ReDim Preserve fluA_k(1 To nPoints&) As Double        ' Mat A total fluorescence k-ratio % (just dimension)
ReDim Preserve fluB_k(1 To nPoints&) As Double        ' Mat B total fluorescence k-ratio % (just dimension)
ReDim Preserve prix_k(1 To nPoints&) As Double        ' Primary x-ray k-ratio % (just dimension)

' Load the distance
xdist#(nPoints&) = Val(Trim$(bstring$))
xdist#(nPoints&) = -xdist#(nPoints&)                ' reverse axis polarity

Call MiscParseStringToString(astring$, bstring$)    ' get total fluorescence plus primary x-ray
If ierror Then Exit Sub
yktotal#(nPoints&) = MathDVal#(Trim$(bstring$))           ' load total fluorescence plus primary x-ray

Call MiscParseStringToString(astring$, bstring$)    ' get fluorescence only
If ierror Then Exit Sub
ykfluor#(nPoints&) = MathDVal#(Trim$(bstring$))           ' load fluorescence only

Call MiscParseStringToString(astring$, bstring$)    ' get Mat A characteristic fluorescence
If ierror Then Exit Sub
flach#(nPoints&) = MathDVal#(Trim$(bstring$))             ' load Mat A characteristic fluorescence

Call MiscParseStringToString(astring$, bstring$)    ' get Mat A continuum fluorescence
If ierror Then Exit Sub
flabr#(nPoints&) = MathDVal#(Trim$(bstring$))             ' load Mat A continuum fluorescence

Call MiscParseStringToString(astring$, bstring$)    ' get Mat B characteristic fluorescence
If ierror Then Exit Sub
flbch#(nPoints&) = MathDVal#(Trim$(bstring$))             ' load Mat B characteristic fluorescence

Call MiscParseStringToString(astring$, bstring$)    ' get Mat B continuum fluorescence
If ierror Then Exit Sub
flbbr#(nPoints&) = MathDVal#(Trim$(bstring$))             ' load Mat B continuum fluorescence

Call MiscParseStringToString(astring$, bstring$)    ' get primary x-ray intensity
If ierror Then Exit Sub
pri_int#(nPoints&) = MathDVal#(Trim$(bstring$))            ' load primary x-ray intensity

Call MiscParseStringToString(astring$, bstring$)    ' get std x-ray intensity
If ierror Then Exit Sub
std_int#(nPoints&) = MathDVal#(Trim$(bstring$))            ' load std x-ray intensity

If VerboseMode Then
Call IOWriteLog("N=" & Format$(nPoints&) & ", X=" & Format$(xdist#(nPoints&)) & ", Y1=" & Format$(yktotal#(nPoints&), e104$) & ", Y2=" & Format$(ykfluor#(nPoints&), e104$))
End If

End If
Loop

Close #Temp1FileNumber%

' Get PAR file densities
MaterialDensityA# = Penepma12GetParFileDensityOnly!(PENEPMA_Root$ & "\Penfluor\" & ParameterFileA$)
If ierror Then Exit Sub
MaterialDensityB# = Penepma12GetParFileDensityOnly!(PENEPMA_Root$ & "\Penfluor\" & ParameterFileB$)
If ierror Then Exit Sub

Exit Sub

' Errors
Penepma12LoadPlotDataError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12LoadPlotData"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma12PlotGrid()
' Plot grid lines

ierror = False
On Error GoTo Penepma12PlotGridError

If FormPENEPMA12.CheckUseGridLines Then
UseGridLines = True
Else
UseGridLines = False
End If

Exit Sub

' Errors
Penepma12PlotGridError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12PlotGrid"
ierror = True
Exit Sub

End Sub

Sub Penepma12PlotLog()
' Plot y axis as log

ierror = False
On Error GoTo Penepma12PlotLogError

If FormPENEPMA12.CheckUseLogScale Then
UseLogScale = True
Else
UseLogScale = False
End If

Exit Sub

' Errors
Penepma12PlotLogError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12PlotLog"
ierror = True
Exit Sub

End Sub

Function Penepma12PARLowerPrecision(pfilename As String, tMaterialSimulationTime As Double) As Boolean
' Check if the input file for the passed PAR file is lower precision that the current precision value

ierror = False
On Error GoTo Penepma12PARLowerPrecisionError

Dim pstring As String
Dim temp As Single

Penepma12PARLowerPrecision = False

' Check if input file exists (it should if PAR file was previously calculated)
If Dir$(MiscGetFileNameNoExtension$(pfilename$) & ".in") = vbNullString Then Exit Function

' Read the simulation time from the input file
Call Penepma12RunFanalCheckINFile("TIME", MiscGetFileNameNoExtension$(pfilename$) & ".in", pstring$)
If ierror Then Exit Function

' Convert
temp! = Val(pstring$)

' Check
If CDbl(temp!) < tMaterialSimulationTime# Then Penepma12PARLowerPrecision = True

Exit Function

' Errors
Penepma12PARLowerPrecisionError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12PARLowerPrecision"
ierror = True
Exit Function

End Function

Function Penepma12PARHigherMinimumEnergy(pfilename As String, tPenepmaMinimumElectronEnergy As Single) As Boolean
' Check if the input file for the passed PAR file is higher minimum energy that the current minimum energy value

ierror = False
On Error GoTo Penepma12PARHigherMinimumEnergyError

Dim pstring As String
Dim temp As Single

Penepma12PARHigherMinimumEnergy = False

' Check if input file exists (it should if PAR file was previously calculated)
If Dir$(MiscGetFileNameNoExtension$(pfilename$) & ".in") = vbNullString Then Exit Function

' Check for adjustment of minimum energy for Li, Be, B, C, N, O, F or Ne
If FormPENEPMA12.CheckAutoAdjustMinimumEnergy.Value = vbChecked Then
Call Penepma12AdjustMinimumEnergy(PENEPMA_Path$ & "\" & MaterialFileA$)
tPenepmaMinimumElectronEnergy! = PenepmaMinimumElectronEnergy!
End If

' Read the simulation time from the input file
Call Penepma12RunFanalCheckINFile("MSIMPA", MiscGetFileNameNoExtension$(pfilename$) & ".in", pstring$)
If ierror Then Exit Function

' Convert
temp! = Val(pstring$)

' Check
If CDbl(temp! / EVPERKEV#) > tPenepmaMinimumElectronEnergy! Then Penepma12PARHigherMinimumEnergy = True

Exit Function

' Errors
Penepma12PARHigherMinimumEnergyError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12PARHigherMinimumEnergy"
ierror = True
Exit Function

End Function

Sub Penepma12ExtractPure()
' Extract the pure element intensities from pure element .par files for all elements, all x-rays

ierror = False
On Error GoTo Penepma12ExtractPureError

Dim i As Integer
Dim response As Integer

Dim tMaterialMeasuredGridPoints As Integer
Dim tExtractElement As Integer
Dim tExtractMatrix As Integer

icancelauto = False

' Warn if less than 1.0 keV minimum energy and not auto adjust minimum energy
If PenepmaMinimumElectronEnergy! < 1# And FormPENEPMA12.CheckAutoAdjustMinimumEnergy.Value = vbUnchecked Then
msg$ = "The minimum electron energy for Penepma kratio extractions is less than 1 keV. Since Penfluor usually only calculates down to 1 keV, this might be problematic. Do you want to continue?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "Penepma12ExtractPure")
If response% = vbCancel Then Exit Sub
End If

' Pure elements are always calculated for the entire range
msg$ = "The specified range extract pure element intensity calculations will take several days to complete (assuming all pure element .PAR files necessary are present). Are you sure you want to proceed?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "Penepma12ExtractPure")
If response% = vbCancel Then Exit Sub

' Extract pure elements
tExtractElement% = ExtractElement%      ' save original emitting element
tExtractMatrix% = ExtractMatrix%        ' save original matrix element

If tExtractElement% > tExtractMatrix% Then GoTo Penepma12ExtractPureNoDo

For i% = tExtractElement% To tExtractMatrix%
ExtractElement% = i%      ' load emitting element
ExtractMatrix% = i%       ' load matrix element

tMaterialMeasuredGridPoints% = MaterialMeasuredGridPoints%      ' save
MaterialMeasuredGridPoints% = 1     ' use a single point for matrix calculations
Call Penepma12ExtractPureIntensity
MaterialMeasuredGridPoints% = tMaterialMeasuredGridPoints%      ' restore

If ierror Then
Exit Sub
Call IOStatusAuto(vbNullString)
End If

Next i%

ExtractElement% = tExtractElement%      ' restore original emitting element
ExtractMatrix% = tExtractMatrix%        ' restore original matrix element

Call IOStatusAuto(vbNullString)
Exit Sub

' Errors
Penepma12ExtractPureError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12ExtractPure"
ierror = True
Exit Sub

Penepma12ExtractPureNoDo:
msg$ = "The specified emitter element is greater than the specified matrix element. There is nothing to do! Please change the element range and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12ExtractPureIntensity"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12ExtractPureIntensity()
' Extract pure element intensities (generated and emitted)
' Material A and Material B and material B Std is the same pure element

ierror = False
On Error GoTo Penepma12ExtractPureIntensityError

Dim l As Integer, m As Integer
Dim eng As Single, edg As Single, tovervoltage As Single
Dim unk_int_pri As Double, unk_int_flu As Double, unk_int_all As Double

Dim tfolder As String, tfilename As String
Dim pfilename As String, pfilename1 As String, pfilename2 As String, pfilename3 As String

Dim pvalue As Single
Dim pstring As String

Dim FanalIntensitiesOutput As Boolean

Dim t1 As Single, t2 As Single

icancelauto = False

' Check for Fanal\pure folder
tfolder$ = PENEPMA_Root$ & "\Fanal\pure"
If Dir$(tfolder$, vbDirectory) = vbNullString Then MkDir tfolder$

' Check for pure element PAR file in Penfluor\Pure folder
pfilename$ = Trim$(Symup$(ExtractElement%)) & ".par"
If Dir$(PENEPMA_Root$ & "\Penfluor\" & pfilename$) = vbNullString Then
tfilename$ = PENEPMA_Root$ & "\Penfluor\Pure\" & pfilename$
If Dir$(tfilename$) <> vbNullString Then FileCopy tfilename$, PENEPMA_Root$ & "\Penfluor\" & pfilename$
If Dir$(MiscGetFileNameNoExtension$(tfilename$) & ".in") <> vbNullString Then FileCopy MiscGetFileNameNoExtension$(tfilename$) & ".in", PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(pfilename$)) & ".in"
End If

' Load PAR file for this pure element for both material A and material B
ParameterFileA$ = pfilename$
ParameterFileB$ = pfilename$
ParameterFileBStd$ = pfilename$

pfilename1$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileA$
pfilename2$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileB$
pfilename3$ = PENEPMA_Root$ & "\Penfluor\" & ParameterFileBStd$

' Load measured element, but for matrix fluorescence calculations just use default distance
MaterialMeasuredElement% = ExtractElement%

' Check for Li, Be, B, C, N, O, F or Ne and adjust minimum energy if so
If FormPENEPMA12.CheckAutoAdjustMinimumEnergy.Value = vbChecked Then
Call Penepma12AdjustMinimumEnergy2(Symlo$(MaterialMeasuredElement%))
End If

' Check for existing .TXT file
tfilename$ = Format$(ExtractElement%) & "_" & Format$(MaterialMeasuredTakeoff#) & ".txt"
If Not CalculateDoNotOverwriteTXT Or (CalculateDoNotOverwriteTXT And Dir$(PENEPMA_Root$ & "\Fanal\pure\" & tfilename$) = vbNullString) Then

' Write column labels
Call Penepma12CalculateReadWritePureElement(Int(0), tfolder$, tfilename$, CSng(0))
If ierror Then Exit Sub

' Reset non-zero intensities output flag
FanalIntensitiesOutput = False

' Loop on each beam voltage from 1 to 50 keV
For m% = 1 To 50
'For m% = 50 To 50       ' testing purposes (50 keV only)
'For m% = 29 To 30       ' testing purposes (49 and 30 keV only)
'For m% = 15 To 15       ' testing purposes (15 keV only)
'For m% = 15 To 16       ' testing purposes (15 and 16 keV only)
'For m% = 19 To 19       ' testing purposes (19 keV only)
'For m% = 28 To 28       ' testing purposes (28 keV only for In ka in Na)
MaterialMeasuredEnergy# = m%

msg$ = vbCrLf & "Extracting Matrix K-Ratios for " & Trim$(Symup$(ExtractElement%)) & " in " & Trim$(Symup$(ExtractMatrix%)) & " at " & Format$(MaterialMeasuredEnergy#) & " keV..."
Call IOWriteLog(msg$)

' Init the globals
Call InitKratios
If ierror Then Exit Sub

' Loop on each valid x-ray
For l% = 1 To MAXRAY% - 1
'For l% = 1 To 1         ' testing purposes (Ka only)
'For l% = 2 To 2         ' testing purposes (Kb only)
Call XrayGetEnergy(MaterialMeasuredElement%, l%, eng!, edg!)
If ierror Then Exit Sub

' Load minimum overvoltage percent, 0 = 2%, 1 = 10%, 2 = 20%, 3 = 40%
If MinimumOverVoltageType% = 0 Then tovervoltage! = MINIMUMOVERVOLTFRACTION_02!
If MinimumOverVoltageType% = 1 Then tovervoltage! = MINIMUMOVERVOLTFRACTION_10!
If MinimumOverVoltageType% = 2 Then tovervoltage! = MINIMUMOVERVOLTFRACTION_20!
If MinimumOverVoltageType% = 3 Then tovervoltage! = MINIMUMOVERVOLTFRACTION_40!

' Check for valid x-ray line (excitation energy (plus a buffer to avoid ultra low overvoltage issues) must be less than beam energy) (and greater than PenepmaMinimumElectronEnergy!)
If eng! <> 0# And edg! <> 0# And (edg! * (1# + tovervoltage!) < MaterialMeasuredEnergy#) And edg! > PenepmaMinimumElectronEnergy! Then

' Double check that specific transition exists (see table 6.2 in Penelope-2006-NEA-pdf)
Call PenepmaGetPDATCONFTransition(MaterialMeasuredElement%, l%, t1!, t2!)
If ierror Then Exit Sub

' If both shells have ionization energies, it is ok to calculate
If t1! <> 0# And t2! <> 0# Then

' Load measured x-ray line
MaterialMeasuredXray% = l%

' Double check that PAR file is in db folder (check penfluor folder in case manually copied)
If Dir$(PENEPMA_Root$ & "\Fanal\db\" & ParameterFileA$) = vbNullString Then
If Dir$(PENEPMA_Root$ & "\Penfluor\" & ParameterFileA$) <> vbNullString Then
FileCopy PENEPMA_Root$ & "\Penfluor\" & ParameterFileA$, PENEPMA_Root$ & "\Fanal\db\" & ParameterFileA$
If Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileA$) & ".in") <> vbNullString Then FileCopy PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileA$) & ".in", PENEPMA_Root$ & "\Fanal\db\" & MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(ParameterFileA$)) & ".in"
Else
GoTo Penepma12ExtractPureIntensityPARFilesNotFound:
End If
End If

Call IOStatusAuto("Extracting pure element intensities based on " & pfilename$ & "...")
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(3), CalculationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

FormPENEPMA12.LabelProgress.Caption = "Extracting pure element intensities from " & pfilename$
FormPENEPMA12.LabelRemainingTime.Caption = vbNullString

' Check for .IN file and if found check MSIMPA parameters (minimum electron/photon energy)
If Dir$(PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileA$) & ".in") <> vbNullString Then
Call Penepma12RunFanalCheckINFile("MSIMPA", PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameNoExtension$(ParameterFileA$) & ".in", pstring$)
If ierror Then Exit Sub
pvalue! = Val(pstring$)
pvalue! = pvalue! / EVPERKEV#

' If necessary skip this beam energy (empty file will deleted below automatically)
If edg! < pvalue! Then
msg$ = pfilename$ & " was only calculated down to " & Format$(pvalue!) & "keV. Skipping pure element intensity extraction for " & Trim$(Symup$(MaterialMeasuredElement%)) & " " & Xraylo$(MaterialMeasuredXray%) & " in " & Symup$(ExtractMatrix%) & "..."
Call IOWriteLog(msg$)
GoTo Penepma12ExtractPureIntensitySkip
End If
End If

' Check the parameters files
Call Penepma12RunFanal
If ierror Then Exit Sub

' Run the Fanal program
Call Penepma12RunFanal1
If ierror Then Exit Sub

' Get k-ratio data from k-ratio file
Call Penepma12LoadPlotData
If ierror Then Exit Sub

' Check that Fanal produced good data
If nPoints& > 0 Then

' Check for valid std intensity
If std_int#(nPoints&) <= 0# Then GoTo Penepma12ExtractPureIntensityZeroStdInt

' Debug
If DebugMode Then
msg$ = "STD-INT= " & std_int#(1) & " for " & Trim$(Symup$(ExtractElement%)) & " " & Xraylo$(MaterialMeasuredXray%) & " in " & ParameterFileA$ & " using standard " & ParameterFileBStd$
Call IOWriteLog(msg$)
End If

' Store essential fluorescent k-ratio data to data array (only need to store first or last data point for matrix calculations)
unk_int_pri# = pri_int#(nPoints&)                                                           ' calculate Mat A/B primary intensity
unk_int_flu# = flach#(nPoints&) + flabr#(nPoints&) + flbch#(nPoints&) + flbbr#(nPoints&)    ' calculate Mat A/B fluorescence intensity
unk_int_all# = unk_int_flu# + pri_int#(nPoints&)                                            ' calculate total intensity

PureGenerated_Intensities#(l%) = std_int(1)
PureEmitted_Intensities#(l%) = std_int(1)

DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(3), CalculationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

' Check for Pause button
Do Until Not RealTimePauseAutomation
DoEvents
Sleep 200
Loop

' If we get to here, non-zero intensities were calculated. Set flag to not erase output file
FanalIntensitiesOutput = True

' Nothing to output
Else
msg$ = "No intensity data to output for " & Trim$(Symup$(ExtractElement%)) & " " & Xraylo$(MaterialMeasuredXray%) & " in " & ParameterFileA$ & " using standard " & ParameterFileBStd$
Call IOWriteLog(msg$)
End If

msg$ = "All pure element intensity extractions are complete for " & Trim$(Symup$(MaterialMeasuredElement%)) & " " & Xraylo$(MaterialMeasuredXray%) & "..."
Call IOWriteLog(msg$)
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(3), CalculationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If

End If

' Overvoltage too low (zero arrays)
Else
PureGenerated_Intensities#(l%) = 0#
PureEmitted_Intensities#(l%) = 0#
End If

DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(3), CalculationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If
Next l%

msg$ = "All x-ray line extractions are complete for " & Trim$(Symup$(MaterialMeasuredElement%)) & " in " & Trim$(Symup$(ExtractMatrix%)) & " at " & Format$(MaterialMeasuredEnergy#) & " keV"
Call IOWriteLog(msg$)

' Write binary k-ratio fluorescence data to file for the specified beam energy
tfilename$ = Format$(ExtractElement%) & "_" & Format$(MaterialMeasuredTakeoff#) & ".txt"
Call Penepma12CalculateReadWritePureElement(Int(2), tfolder$, tfilename$, CSng(MaterialMeasuredEnergy#))
If ierror Then Exit Sub

DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Call Penepma12CheckTermination2(Int(3), CalculationInProgress)
If ierror Then Exit Sub
Call IOShellTerminateTask(PenepmaTaskID&)
If ierror Then Exit Sub
ierror = True
Exit Sub
End If
Next m%

' PAR file not calculated with sufficiently low minimum energy
Penepma12ExtractPureIntensitySkip:

' Check if non-zero intensities were actually output. If not, delete the TXT file
If Not FanalIntensitiesOutput Then
Kill tfolder$ & "\" & tfilename$
Exit Sub
End If

Else
msg$ = "Skipping pure element intensity extraction for " & tfilename$ & "..."
Call IOWriteLog(msg$)
End If

Call IOStatusAuto(vbNullString)
msg$ = "All pure element intensity extractions are complete"
Call IOWriteLog(msg$)
DoEvents

Exit Sub

' Errors
Penepma12ExtractPureIntensityError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12ExtractPureIntensity"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12ExtractPureIntensityPARFilesNotFound:
msg$ = "The specified .PAR file (" & ParameterFileA$ & ") was not found in the Fanal\db or Penfluor folders. Please calculate the specified .PAR parameter file and try again"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12ExtractPureIntensity"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12ExtractPureIntensityZeroStdInt:
msg$ = "The standard intensity for the measured element " & Symlo$(MaterialMeasuredElement%) & " " & Xraylo$(MaterialMeasuredXray%) & " was zero for the material B Std composition (" & ParameterFileBStd$ & ") at " & Format$(MaterialMeasuredEnergy#) & " keV. This error should not occur, please contact Probe Software with details (and check the Fanal\k-ratios.dat file)."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12ExtractPureIntensity"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12CheckXray(sample() As TypeSample, notfound1 As Boolean, notfound2 As Boolean)
' Check for valid x-rays

ierror = False
On Error GoTo Penepma12CheckXrayError

Dim chan As Integer, nrec As Integer, jnum As Integer
Dim tenergy As Single
Dim sym As String
 
Dim engrow As TypeEnergy
Dim edgrow As TypeEdge

notfound1 = False
notfound2 = False

' Open x-ray edge file
Open XEdgeFile$ For Random Access Read As #XEdgeFileNumber% Len = XRAY_FILE_RECORD_LENGTH%

' Open x-ray line file
Open XLineFile$ For Random Access Read As #XLineFileNumber% Len = XRAY_FILE_RECORD_LENGTH%

' Open x-ray line file
If Dir$(XLineFile2$) = vbNullString Then GoTo Penepma12CheckXrayNotFoundXLINE2DAT
If FileLen(XLineFile2$) = 0 Then GoTo Penepma12CheckXrayZeroSizeXLINE2DAT
Open XLineFile2$ For Random Access Read As #XLineFileNumber2% Len = XRAY_FILE_RECORD_LENGTH%

' Check xray line selections, load as absorber if specified or a problem is found
For chan% = 1 To sample(1).LastChan%
sym$ = sample(1).Xrsyms$(chan%)

' Read original x-ray lines
nrec% = sample(1).AtomicNums%(chan%) + 2
If sample(1).XrayNums%(chan%) <= MAXRAY_OLD% Then
Get #XLineFileNumber%, nrec%, engrow
tenergy! = engrow.energy!(sample(1).XrayNums%(chan%))

' Read additional x-ray lines
Else
Get #XLineFileNumber2%, nrec%, engrow
tenergy! = engrow.energy!(sample(1).XrayNums%(chan%) - MAXRAY_OLD%)
End If

' Check for bad xray lines (no data)
If chan% = 1 And tenergy! <= 0# Then notfound1 = True
If chan% = 2 And tenergy! <= 0# Then notfound2 = True

' Now read edge energies
Get #XEdgeFileNumber%, nrec%, edgrow

' Calculate edge index for this x-ray
If sample(1).XrayNums%(chan%) = 1 Then jnum% = 1   ' Ka
If sample(1).XrayNums%(chan%) = 2 Then jnum% = 1   ' Kb
If sample(1).XrayNums%(chan%) = 3 Then jnum% = 4   ' La
If sample(1).XrayNums%(chan%) = 4 Then jnum% = 3   ' Lb
If sample(1).XrayNums%(chan%) = 5 Then jnum% = 9   ' Ma
If sample(1).XrayNums%(chan%) = 6 Then jnum% = 8   ' Mb

If sample(1).XrayNums%(chan%) = 7 Then jnum% = 3   ' Ln
If sample(1).XrayNums%(chan%) = 8 Then jnum% = 3   ' Lg
If sample(1).XrayNums%(chan%) = 9 Then jnum% = 3   ' Lv
If sample(1).XrayNums%(chan%) = 10 Then jnum% = 4   ' Ll
If sample(1).XrayNums%(chan%) = 11 Then jnum% = 7   ' Mg
If sample(1).XrayNums%(chan%) = 12 Then jnum% = 9   ' Mz

' Check for missing absorption edge energy
If chan% = 1 And edgrow.energy!(jnum%) <= 0# Then notfound1 = True
If chan% = 2 And edgrow.energy!(jnum%) <= 0# Then notfound2 = True
Next chan%

Close #XEdgeFileNumber%
Close #XLineFileNumber%
Close #XLineFileNumber2%

Exit Sub

' Errors
Penepma12CheckXrayError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CheckXray"
Close #XEdgeFileNumber%
Close #XLineFileNumber%
Close #XLineFileNumber2%
ierror = True
Exit Sub

Penepma12CheckXrayNotFoundXLINE2DAT:
msg$ = "The " & XLineFile2$ & " was not found." & vbCrLf & vbCrLf
msg$ = msg$ & "Please run the latest CalcZAF.msi installer to obtain this additional x-ray line file."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CheckXray"
Close #XEdgeFileNumber%
Close #XLineFileNumber%
Close #XLineFileNumber2%
ierror = True
Exit Sub

Penepma12CheckXrayZeroSizeXLINE2DAT:
Kill XLineFile2$
msg$ = "The " & XLineFile2$ & " was not found." & vbCrLf & vbCrLf
msg$ = msg$ & "Please run the latest CalcZAF.msi installer to obtain this additional x-ray line file."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CheckXray"
Close #XEdgeFileNumber%
Close #XLineFileNumber%
Close #XLineFileNumber2%
ierror = True
Exit Sub

End Sub

Sub Penepma12CheckPenfluorInputFiles()
' Check Penfluor input files for self consistency in selected folder

ierror = False
On Error GoTo Penepma12CheckPenfluorInputFilesError

Dim EmittingElement As Integer, ip As Integer
Dim tstring As String, pstring As String, bstring As String
Dim tfilename As String
Dim i As Long, ii As Long

Dim filearray() As String

Static tpath As String

icancelauto = False

' Browse to a specified folder containing the Penfluor .in files
tstring$ = "Browse PENFLUOR Input File(s) Folder"
If tpath$ = vbNullString Then tpath$ = PENEPMA_Root$ & "\Penfluor\"
tpath$ = IOBrowseForFolderByPath(True, tpath$, tstring$, FormPenepma12Binary)
If ierror Then Exit Sub
If Trim$(tpath$) = vbNullString Then Exit Sub

' Save current Penfluor form parameters (simulation time and minimum energy)
Call Penepma12Save
If ierror Then Exit Sub

' Make a list of all input files (must do this way to avoid reentrant Dir$ calls)
tfilename$ = Dir$(tpath$ & "\*.in")  ' get first file
ii& = 0
Do While tfilename$ <> vbNullString
ii& = ii& + 1
ReDim Preserve filearray(1 To ii&) As String
filearray$(ii&) = tfilename$
tfilename$ = Dir$
Loop

' Check for any .in files
If ii& < 1 Then GoTo Penepma12CheckPenfluorInputFilesNotFound

' Loop through all input files and check against form parameters
Call IOWriteLog(vbCrLf & "Penepma12CheckPenfluorInputFiles: Checking Penfluor input file parameters in folder " & tpath$)
Screen.MousePointer = vbHourglass
For i& = 1 To ii&
tfilename$ = tpath$ & "\" & filearray$(i&)
Call IOStatusAuto("Penepma12CheckPenfluorInputFiles: checking Penfluor input file " & tfilename$ & " (" & Format$(i&) & " of " & Format$(ii&) & ")")

' Read the simulation time from the input file
Call Penepma12RunFanalCheckINFile("TIME", tfilename, pstring$)
If ierror Then
Screen.MousePointer = vbDefault
Exit Sub
End If
If Val(pstring$) <> MaterialSimulationTime# Then
Call IOWriteLog("Penepma12CheckPenfluorInputFiles: Penfluor input file " & tfilename$ & ", does not match current TIME parameter (" & pstring$ & " vs. " & Format$(MaterialSimulationTime#) & ").")
End If

' Load minimum energy if auto adjust is checked
If FormPENEPMA12.CheckAutoAdjustMinimumEnergy.Value = vbChecked Then
Call MiscParseStringToStringA(filearray$(i&), "-", bstring$)
If ierror Then
Screen.MousePointer = vbDefault
Exit Sub
End If
ip% = IPOS1%(MAXELM%, bstring$, Symup$())
If ip% = 0 Then GoTo Penepma12CheckPenfluorInputFilesBadSymbol
EmittingElement% = ip%
Call Penepma12AdjustMinimumEnergy2(Symlo$(EmittingElement%))
If ierror Then
Screen.MousePointer = vbDefault
Exit Sub
End If
End If

' Read the minimum energy from the input file
Call Penepma12RunFanalCheckINFile("MSIMPA", tfilename$, pstring$)
If ierror Then
Screen.MousePointer = vbDefault
Exit Sub
End If
If Val(pstring$) <> CSng(PenepmaMinimumElectronEnergy! * EVPERKEV#) Then    ' need to cast to single precision
Call IOWriteLog("Penepma12CheckPenfluorInputFiles: Penfluor input file " & tfilename$ & ", does not match current MSIMPA parameter (" & pstring$ & " vs. " & Format$(CSng(PenepmaMinimumElectronEnergy! * EVPERKEV#)) & ").")
End If

DoEvents
If icancelauto Or ierror Then
Screen.MousePointer = vbDefault
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

' Get next input filename
Next i&
Screen.MousePointer = vbDefault

Call IOWriteLog("Penepma12CheckPenfluorInputFiles: Checking Penfluor input file parameters in folder " & tpath$ & " is complete!")
Call IOStatusAuto(vbNullString)
Exit Sub

' Errors
Penepma12CheckPenfluorInputFilesError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CheckPenfluorInputFiles"
ierror = True
Exit Sub

Penepma12CheckPenfluorInputFilesNotFound:
Screen.MousePointer = vbDefault
msg$ = "No Penfluor input files found in folder " & tpath$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CheckPenfluorInputFiles"
ierror = True
Exit Sub

Penepma12CheckPenfluorInputFilesBadSymbol:
Screen.MousePointer = vbDefault
msg$ = "Element symbol " & bstring$ & " in file " & filearray$(i&) & " is not a valid element symbol."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CheckPenfluorInputFiles"
ierror = True
Exit Sub

End Sub

Sub Penepma12MatrixCheckDeviations(tTakeoff As Single)
' This routine reads the Matrix.mdb file for beam energy, emitter, x-ray, matrix and fits them as alpha factors and check for large variances in the fit.
'  tBinaryRanges!(1 to MAXBINARY%)  are the compositional binaries in weight percent
'  tBinary_Kratios#(1 to MAXRAY_OLD%, 1 to MAXBINARY%)  are the k-ratios for each x-ray and binary composition
'  tBinary_Factors!(1 to MAXRAY_OLD%, 1 to MAXBINARY%)  are the alpha factors for each x-ray and binary composition, alpha = (C/K - C)/(1 - C)
'  tBinary_Coeffs!(1 to MAXRAY_OLD%, 1 to MAXCOEFF4%)  are the polynomial/non-linear alpha factors fit coefficients for each x-ray
'  tBinary_Devs!(1 to MAXRAY_OLD%)  are the alpha factor fit standard deviations for each x-ray

ierror = False
On Error GoTo Penepma12MatrixCheckDeviationsError

Dim notfound As Boolean
Dim i As Integer, j As Integer, n As Integer, l As Integer, m As Integer, npts As Integer
Dim temp As Single
Dim tKilovolts As Single
Dim tEmitter As Integer, tXray As Integer, tMatrix As Integer
Dim astring As String
Dim eng As Single, edg As Single

ReDim tKratios(1 To MAXBINARY%) As Double

ReDim CalcZAF_ZAF_Factors(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Single
ReDim CalcZAF_ZA_Factors(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Single
ReDim CalcZAF_F_Factors(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Single

ReDim Binary_ZAF_Factors(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Single
ReDim Binary_ZA_Factors(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Single
ReDim Binary_F_Factors(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Single

ReDim Binary_ZAF_Coeffs(1 To MAXRAY%, 1 To MAXCOEFF4%) As Single
ReDim CalcZAF_ZAF_Coeffs(1 To MAXRAY%, 1 To MAXCOEFF4%) As Single

ReDim Binary_ZA_Coeffs(1 To MAXRAY%, 1 To MAXCOEFF4%) As Single
ReDim CalcZAF_ZA_Coeffs(1 To MAXRAY%, 1 To MAXCOEFF4%) As Single

ReDim Binary_F_Coeffs(1 To MAXRAY%, 1 To MAXCOEFF4%) As Single
ReDim CalcZAF_F_Coeffs(1 To MAXRAY%, 1 To MAXCOEFF4%) As Single

ReDim Binary_ZAF_Betas(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Single
ReDim CalcZAF_ZAF_Betas(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Single

ReDim Binary_ZAF_Devs(1 To MAXRAY%) As Single
ReDim CalcZAF_ZAF_Devs(1 To MAXRAY%) As Single

ReDim Binary_ZA_Devs(1 To MAXRAY%) As Single
ReDim CalcZAF_ZA_Devs(1 To MAXRAY%) As Single

ReDim Binary_F_Devs(1 To MAXRAY%) As Single
ReDim CalcZAF_F_Devs(1 To MAXRAY%) As Single

Const maxdev! = 20#            ' check for more than 20% average deviation in alpha fit

icancelauto = False

' Check for file
If Dir$(MatrixMDBFile$) = vbNullString Then GoTo Penepma12MatrixCheckDeviationsNoMatrixMDBFile

' Delete Standard.txt and Standard.err file if present
If Dir$(ProbeTextLogFile$) <> vbNullString Then Kill ProbeTextLogFile$
If Dir$(ProbeErrorLogFile$) <> vbNullString Then Kill ProbeErrorLogFile$

' Loop on each possible energy
Screen.MousePointer = vbHourglass
'For m% = 4 To 50         ' Fanal calculations are only good down to 5 keV at this time but start at 4 keV for Pouchou database calculations!
For m% = 5 To 50
tKilovolts! = CSng(m%)

' Loop on each possible emitter and absorber
For i% = 6 To MAXELM%                       ' emitters from carbon
For j% = 3 To MAXELM%                       ' absorbers from lithium

' Skip if same element
If i% <> j% Then

tEmitter% = i%      ' load emitting element
tMatrix% = j%       ' load matrix element

' Loop on each valid x-ray
For l% = 1 To MAXRAY_OLD%       'only original x-ray lines for now
'For l% = 1 To 1                 ' testing purposes (Ka only)
tXray% = l%

Call XrayGetEnergy(tEmitter%, l%, eng!, edg!)
If ierror Then Exit Sub

' Check for valid x-ray line
If eng! <> 0# And edg! <> 0# Then

' Get the next emitter
Call Penepma12MatrixReadMDB2(tTakeoff!, tKilovolts!, tEmitter%, tXray%, tMatrix%, tKratios#(), notfound)
If ierror Then Exit Sub

Call IOStatusAuto("Checking binary " & Symlo$(tEmitter%) & " " & Xraylo$(tXray%) & " in " & Symlo$(tMatrix%) & " at TO= " & Format$(tTakeoff!) & ", keV= " & Format$(tKilovolts!))
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

' Load and check if found
If Not notfound Then
For n% = 1 To MAXBINARY%
Binary_ZAF_Kratios#(l%, n%) = tKratios#(n%)
If UsePenepmaKratiosLimitFlag And BinaryRanges(n%) > PenepmaKratiosLimitValue! Then
Binary_ZAF_Kratios#(l%, n%) = 0#
End If
Next n%

' Fit alpha (assume polynomial) for this x-ray
Call Penepma12CalculateAlphaFactors(l%, BinaryRanges!(), Binary_ZAF_Kratios#(), Binary_ZAF_Factors!(), Binary_ZAF_Coeffs!(), Binary_ZAF_Devs!(), npts%)
If ierror Then Exit Sub

' Check calculated standard deviation
If Binary_ZAF_Devs!(l%) > maxdev! Then

' Output to log window and log file
If tKilovolts! >= 5 And tKilovolts! <= 30 Then    ' only print problematic matrix corrections if between 5 and 30 keV
msg$ = vbCrLf & "Average standard deviation for alpha fit (npts=" & Format$(npts%) & ") is " & Format$(Binary_ZAF_Devs!(l%)) & "% for " & Symlo$(tEmitter%) & " " & Xraylo$(tXray%) & " in " & Symlo$(tMatrix%) & " at TO= " & Format$(tTakeoff!) & ", keV= " & Format$(tKilovolts!)
Call IOWriteLog(msg$)
Call IOWriteError(msg$, "Penepma12MatrixCheckDeviations")
If ierror Then Exit Sub

' Output to log
astring$ = Symup$(tEmitter%) & " " & Xraylo$(tXray%) & " in " & Symup$(tMatrix%)
Call IOWriteLog$(vbCrLf & astring$ & " at " & Format$(tTakeoff!) & " degrees and " & Format$(tKilovolts!) & " keV")
astring$ = Format$(vbTab & "Conc%", a08$) & vbTab & Format$("Kratio%", a08$) & vbTab & Format$("Alpha", a08$)
Call IOWriteLog$(astring$)

' Output all binaries
For n% = 1 To MAXBINARY%
astring$ = vbTab & MiscAutoFormat$(BinaryRanges!(n%)) & vbTab & MiscAutoFormatD$(Binary_ZAF_Kratios#(l%, n%)) & vbTab & MiscAutoFormat$(Binary_ZAF_Factors!(l%, n%))
Call IOWriteLog$(astring$)
Next n%
End If
End If

End If

End If
Next l%

End If
Next j%
Next i%

Next m%

Screen.MousePointer = vbDefault
Exit Sub

' Errors
Penepma12MatrixCheckDeviationsError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12MatrixCheckDeviations"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12MatrixCheckDeviationsNoMatrixMDBFile:
Screen.MousePointer = vbDefault
msg$ = "File " & MatrixMDBFile$ & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12MatrixCheckDeviations"
ierror = True
Exit Sub

Penepma12MatrixCheckDeviationsNoMatrixRecords:
Screen.MousePointer = vbDefault
msg$ = "No matrix records found in matrix database " & MatrixMDBFile$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12MatrixCheckDeviations"
ierror = True
Exit Sub

End Sub

Sub Penepma12ExtractBinary(tForm As Form)
' Extract k-ratios from binary compositions in file

ierror = False
On Error GoTo Penepma12ExtractBinaryError

Dim ibin As Integer, n As Integer
Dim tMaterialMeasuredGridPoints As Integer

Dim eO As Single, TOA As Single

Dim tfilename As String
Dim binaryname As String

Dim BinarySwapped As Boolean

Dim kratios(1 To MAXBINARY%) As Single

ReDim isym(1 To 2) As Integer
ReDim iray(1 To 2) As Integer
ReDim conc(1 To 2) As Single
ReDim kexp(1 To 2) As Single

Dim BinaryLineCount As Long

Static lastfilename As String

' Ask user for input file
icancelauto = False

Close #ImportDataFileNumber%

' Get import filename from user
If lastfilename$ = vbNullString Then lastfilename$ = CalcZAFDATFileDirectory$ & "\Pouchou2_Au,Cu,Ag_only.dat"
tfilename$ = lastfilename$
Call IOGetFileName(Int(2), "DAT", tfilename$, tForm)
If ierror Then
Close #ImportDataFileNumber%
Exit Sub
End If

' Save current path
CalcZAFDATFileDirectory$ = CurDir$

' No errors, save file name
lastfilename$ = tfilename$
ImportDataFile$ = lastfilename$

' Set keV rounding flag to false for input file calculations
Penepma12UseKeVRoundingFlag = False

msg$ = "Setting keV rounding flag to false..."
Call IOWriteLog(vbCrLf & msg$ & vbCrLf)

' Open input file
Open ImportDataFile$ For Input As #ImportDataFileNumber%
BinaryLineCount& = 0
Call IOStatusAuto(vbNullString)

' Extract k-ratios for all lines in input file
Do While Not EOF(ImportDataFileNumber%)
BinaryLineCount& = BinaryLineCount& + 1

' Read binary elements, kilovolts and takeoff
Input #ImportDataFileNumber%, isym%(1), isym%(2), iray%(1), iray%(2), eO!, TOA!, conc!(1), conc!(2), kexp!(1), kexp!(2)
If conc!(2) = 0# Then conc!(2) = 1# - conc!(1)

' Check limits
If isym%(1) < 1 Or isym%(1) > MAXELM% Then GoTo Penepma12ExtractBinaryOutofLimits
If isym%(2) < 1 Or isym%(2) > MAXELM% Then GoTo Penepma12ExtractBinaryOutofLimits
If iray%(1) < 1 Or iray%(1) > MAXRAY% Then GoTo Penepma12ExtractBinaryOutofLimits
If iray%(2) < 1 Or iray%(2) > MAXRAY% Then GoTo Penepma12ExtractBinaryOutofLimits
If eO! < 1# Or eO! > 100# Then GoTo Penepma12ExtractBinaryOutofLimits
If TOA! < 1# Or TOA! > 90# Then GoTo Penepma12ExtractBinaryOutofLimits
If conc!(1) < 0# Or conc!(1) > 1# Then GoTo Penepma12ExtractBinaryOutofLimits
If conc!(2) < 0# Or conc!(2) > 1# Then GoTo Penepma12ExtractBinaryOutofLimits
If kexp!(1) < 0# Or kexp!(1) > 1# Then GoTo Penepma12ExtractBinaryOutofLimits
If kexp!(2) < 0# Or kexp!(2) > 1# Then GoTo Penepma12ExtractBinaryOutofLimits

' Check that both elements are not by difference
If iray%(1) = MAXRAY% And iray%(2) = MAXRAY% Then GoTo Penepma12ExtractBinaryBothByDifference

' Check that at least one concentration is entered
If conc!(1) = 0# And conc!(2) = 0# Then GoTo Penepma12ExtractBinaryNoConcData

' Extract k-ratios for both binaries
For ibin% = 1 To 2
If iray%(ibin%) < MAXRAY% Then  ' skip if x-ray is "not analyzed"

If ibin% = 1 Then msg$ = "Extracting k-ratios from binary " & Format$(BinaryLineCount&) & " (" & Symlo$(isym%(ibin%)) & " " & Xraylo$(iray%(ibin%)) & " in " & Symlo$(isym%(2)) & " at " & Format$(TOA!) & " deg, " & Format$(eO!) & " kev)..."
If ibin% = 2 Then msg$ = "Extracting k-ratios from binary " & Format$(BinaryLineCount&) & " (" & Symlo$(isym%(ibin%)) & " " & Xraylo$(iray%(ibin%)) & " in " & Symlo$(isym%(1)) & " at " & Format$(TOA!) & " deg, " & Format$(eO!) & " kev)..."
Call IOWriteLog(vbCrLf & msg$)
Call IOStatusAuto(msg$)
If icancelauto Then
Call IOStatusAuto(vbNullString)
Close #ImportDataFileNumber%
ierror = True
Exit Sub
End If

' Specify the Fanal parameters
MaterialMeasuredTakeoff# = TOA!
MaterialMeasuredEnergy# = eO!

MaterialMeasuredElement% = isym%(ibin%)
MaterialMeasuredXray% = iray%(ibin%)

' Loop on each binary composition (1:99, 5:95, 90:10, etc.)
For n% = 1 To MAXBINARY%
PENEPMA_Sample(1).ElmPercents!(1) = BinaryRanges!(n%)
PENEPMA_Sample(1).ElmPercents!(2) = 100# - BinaryRanges!(n%)

' Create material (and PAR) file name
If isym%(1) < isym%(2) Then
binaryname$ = Trim$(Symup$(isym%(1))) & "-" & Trim$(Symup$(isym%(2))) & "_" & Format$(PENEPMA_Sample(1).ElmPercents!(1)) & "-" & Format$(PENEPMA_Sample(1).ElmPercents!(2))
BinarySwapped = False
Else
binaryname$ = Trim$(Symup$(isym%(2))) & "-" & Trim$(Symup$(isym%(1))) & "_" & Format$(PENEPMA_Sample(1).ElmPercents!(1)) & "-" & Format$(PENEPMA_Sample(1).ElmPercents!(2))
BinarySwapped = True
End If

ParameterFileA$ = binaryname$ & ".par"
ParameterFileB$ = binaryname$ & ".par"                          ' same as A for matrix calculations
ParameterFileBStd$ = Trim$(Symup$(isym%(ibin%))) & ".par"       ' use pure element always (use Trim$ for single letter elements)

' Check for pure element PAR file in Penfluor\Pure folder
If Dir$(PENEPMA_Root$ & "\Penfluor\" & ParameterFileBStd$) = vbNullString Then
tfilename$ = PENEPMA_Root$ & "\Penfluor\Pure\" & ParameterFileBStd$
If Dir$(tfilename$) <> vbNullString Then FileCopy tfilename$, PENEPMA_Root$ & "\Penfluor\" & ParameterFileBStd$
If Dir$(MiscGetFileNameNoExtension$(tfilename$) & ".in") <> vbNullString Then FileCopy MiscGetFileNameNoExtension$(tfilename$) & ".in", PENEPMA_Root$ & "\Penfluor\" & MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(ParameterFileBStd$)) & ".in"
End If

' Check the parameters files
Call Penepma12RunFanal
If ierror Then
Close #ImportDataFileNumber%
Exit Sub
End If

' Run the Fanal program
tMaterialMeasuredGridPoints% = MaterialMeasuredGridPoints%      ' save
MaterialMeasuredGridPoints% = 1     ' use a single point for matrix calculations
Call Penepma12RunFanal1
MaterialMeasuredGridPoints% = tMaterialMeasuredGridPoints%      ' restore
If ierror Then
Close #ImportDataFileNumber%
Exit Sub
End If

' Get k-ratio data from Fanal k-ratio file
Call Penepma12LoadPlotData
If ierror Then
Close #ImportDataFileNumber%
Exit Sub
End If

' Save k-ratio to array
kratios!(n%) = yktotal#(1)   ' from bulk Fanal calculation (store in k-ratio percent)

' Check for Pause button
Do Until Not RealTimePauseAutomation
DoEvents
Sleep 200
Loop

' Check for cancel
Call IOStatusAuto(msg$)
If icancelauto Then
Call IOStatusAuto(vbNullString)
Close #ImportDataFileNumber%
ierror = True
Exit Sub
End If

' Next binary composition
Next n%

' Store calculated binary alpha k-ratios to matrix.mdb file
If ibin% = 1 Then
If Not BinarySwapped Then
Call Penepma12MatrixUpdateMDB(Int(1), TOA!, eO!, isym%(1), iray%(1), isym%(2), kratios!())
If ierror Then Exit Sub

Else
Call Penepma12MatrixUpdateMDB(Int(2), TOA!, eO!, isym%(1), iray%(1), isym%(2), kratios!())
If ierror Then Exit Sub
End If

Else
If Not BinarySwapped Then
Call Penepma12MatrixUpdateMDB(Int(1), TOA!, eO!, isym%(2), iray%(2), isym%(1), kratios!())
If ierror Then Exit Sub

Else
Call Penepma12MatrixUpdateMDB(Int(2), TOA!, eO!, isym%(2), iray%(2), isym%(1), kratios!())
If ierror Then Exit Sub
End If
End If

' Next element of binary
End If
Next ibin%

' Next import line from file
Loop

' Close files
Close #ImportDataFileNumber%

Call IOStatusAuto(vbNullString)
msg$ = "Binary Fanal alpha calculations completed on file " & ImportDataFile$ & ", and k-ratios saved to matrix.mdb file."
MsgBox msg$, vbOKOnly + vbInformation, "Penepma12ExtractBinary"

Screen.MousePointer = vbDefault
Exit Sub

' Errors
Penepma12ExtractBinaryError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12ExtractBinaryError"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12ExtractBinaryBothByDifference:
Close #ImportDataFileNumber%
msg$ = "Both elements are by difference on line " & Str$(BinaryLineCount&) & " in " & ImportDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12ExtractBinary"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12ExtractBinaryOutofLimits:
Close #ImportDataFileNumber%
msg$ = "Bad data on line " & Str$(BinaryLineCount&) & " in " & ImportDataFile$ & " (file format may be wrong)."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12ExtractBinary"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12ExtractBinaryNoConcData:
Close #ImportDataFileNumber%
msg$ = "No Conc data on line " & Str$(BinaryLineCount&) & " in " & ImportDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12ExtractBinary"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub
