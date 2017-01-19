Attribute VB_Name = "CodePenepma08c"
' (c) Copyright 1995-2017 by John J. Donovan
Option Explicit

Const COL7% = 7

Sub Penepma08CreatePenepmaFile(tLiveTime As Single, sample() As TypeSample)
' Create Penepma input file (*.IN) from the passed sample (only for demo EDS acquisition)

ierror = False
On Error GoTo Penepma08CreatePenepmaFileError

Dim astring As String, bstring As String, cstring As String, dstring As String
Dim tfilename As String, tfilename2 As String
Dim iseed1 As Integer, iseed2 As Integer

Const tSimulatedShowers# = 20000000000#

Dim tBeamMinimumEnergyRange As Double
Dim tBeamMaximumEnergyRange As Double
Dim tBeamNumberOfEnergyChannels As Long

' Loop through sample production file and copy to new file with modified parameters
tfilename$ = PENEPMA_Root$ & "\Cu_cha.in"
Open tfilename$ For Input As #Temp1FileNumber%
tfilename2$ = PENEPMA_Path$ & "\material.in"
Open tfilename2$ For Output As #Temp2FileNumber%

Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$
bstring$ = astring$

If InStr(astring$, "TITLE") > 0 Then bstring$ = Left$(astring, COL7%) & Left$(sample(1).Name$, 120)

' Beam energy
If InStr(astring$, "SENERG") > 0 Then
cstring$ = Format$(Format$(sample(1).kilovolts! * EVPERKEV#, "Scientific"), a10$)
Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub
End If

If InStr(astring$, "SPOSIT") > 0 Then
cstring$ = Format$(0#) & " " & Format$(0#) & " " & Format$(1#)
Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub
End If

If InStr(astring$, "SDIREC") > 0 Then
cstring$ = Format$(180#) & " " & Format$(0#)
Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub
End If

If InStr(astring$, "SAPERT") > 0 Then
cstring$ = Format$(0#)
Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub
End If

' Load material file and simulation parameters
If InStr(astring$, "MFNAME") > 0 Then
cstring$ = "material.mat"
Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub
End If

' Electron and photon minimum energies (1.0E+3 1.0E+3 1E+3 0.1 0.1 1E+3 1E+3)
If InStr(astring$, "MSIMPA") > 0 Then
cstring$ = Format$(PenepmaMinimumElectronEnergy! * EVPERKEV#, "0.0E+0") & " " & Format$(PenepmaMinimumElectronEnergy! * EVPERKEV#, "0.0E+0") & " " & Format$(PenepmaMinimumElectronEnergy! * EVPERKEV#, "0E+0") & " "
cstring$ = cstring$ & Format$(0.1, "0.0") & " " & Format$(0.1, "0.0") & " "
cstring$ = cstring$ & Format$(PenepmaMinimumElectronEnergy! * EVPERKEV#, "0E+0") & " " & Format$(PenepmaMinimumElectronEnergy! * EVPERKEV#, "0E+0")
Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub
End If

' Load geometry file
If InStr(astring$, "GEOMFN") > 0 Then
If Dir$(PENEPMA_Path$ & "\bulk.geo") = vbNullString Then
If Dir$(PENEPMA_Root$ & "\bulk.geo") = vbNullString Then GoTo Penepma08CreatePenepmaFileNotFoundGEO
FileCopy PENEPMA_Root$ & "\bulk.geo", PENEPMA_Path$ & "\bulk.geo"
End If
cstring$ = MiscGetFileNameOnly$(PENEPMA_Path$ & "\bulk.geo")
Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub
End If

' Detector angles (45.0, 55.0, 0.0, 360.0, 0)
If InStr(astring$, "PDANGL") > 0 Then
cstring$ = Format$(90# - (sample(1).takeoff! + 5#), "0.0") & " " & Format$(90# - (sample(1).takeoff! - 5#), "0.0") & " "
cstring$ = cstring$ & Format$(0#, "0.0") & " " & Format$(360#, "0.0") & " "
cstring$ = cstring$ & Format$(0, "0")
Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub
End If

' EDS spectrum energy range and channels
If InStr(astring$, "PDENER") > 0 Then
tBeamMinimumEnergyRange# = 0
tBeamMaximumEnergyRange# = sample(1).kilovolts! * EVPERKEV#
tBeamNumberOfEnergyChannels& = CInt((tBeamMaximumEnergyRange# - tBeamMinimumEnergyRange#) / DEMO_EDS_EVPERCHANNEL!)
cstring$ = Format$(tBeamMinimumEnergyRange#, "0.0") & " " & Format$(tBeamMaximumEnergyRange#, "0.0") & " "
cstring$ = cstring$ & Format$(tBeamNumberOfEnergyChannels&, "0")
Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
End If

' Dump time (15 seconds for now)
If InStr(astring$, "DUMPP") > 0 Then
cstring$ = Format$(15#)
Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub
End If

' Number of simulated showers (make arbitrarily large number)
If InStr(astring$, "NSIMSH") > 0 Then
cstring$ = Format$(tSimulatedShowers#, e71$)
Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub
End If

' Add random seed for realistic replicate EDS intensities!
If InStr(astring$, "RSEED") > 0 Then
iseed1% = (Rnd() - 1#) * 1000
iseed2% = 1
cstring$ = Format$(iseed1%) & " " & Format$(iseed2%)
Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub
End If

If InStr(astring$, "TIME") > 0 Then
cstring$ = Format$(tLiveTime!)
Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub
End If

Print #Temp2FileNumber%, bstring$
Loop

Close #Temp1FileNumber%
Close #Temp2FileNumber%

Call IOStatusAuto(vbNullString)
Exit Sub

' Errors
Penepma08CreatePenepmaFileError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08CreatePenepmaFile"
Close #Temp1FileNumber%
Close #Temp2FileNumber%
ierror = True
Exit Sub

Penepma08CreatePenepmaFileNotFoundGEO:
msg$ = "The specified geometry file (" & PENEPMA_Root$ & "\bulk.geo" & ") was not found. Please download the Penepma12.ZIP file, extract and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08CreatePenepmaFile"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08RunConvolgEXE(mode As Integer)
' Run the Convolg program to convolve the Penepma spectrum
'  mode = 0  use default EDS resolution (~140 eV)
'  mode = 1  use WDS LIF resolution (~4 eV)
'  mode = 2  use WDS PET resolution (~2 eV)
'  mode = 3  use WDS TAP resolution (~8 eV)
'  mode = 4  use WDS LDE resolution (~20 eV)

ierror = False
On Error GoTo Penepma08RunConvolgEXEError

Dim tfilenumber As Integer
Dim taskID As Long
Dim bfilename As String, astring As String

' First check if Penepma has had time to create a spectrum file
If PENEPMA_SPEC_File$ = vbNullString Then GoTo Penepma08RunConvolgEXENotSpecified
If Dir$(PENEPMA_SPEC_File$) = vbNullString Then Exit Sub

' Check for valid mode
If mode% < 0 Or mode > 4 Then GoTo Penepma08RunConvolgEXEBadMode

' Check for blank convolg output file name
If Trim$(PENEPMA_CONVOLG_File$) = vbNullString Then GoTo Penepma08RunConvolgEXENoConvolgFilename

' Create Convolg input file
Call Penepma08CreateConvolgFile(mode%)
If ierror Then Exit Sub

Sleep 200

' Delete existing temp batch file (use temp2.bat because Penepma is using temp.bat)
bfilename$ = PENDBASE_Path$ & "\temp2.bat"
If Dir$(bfilename$) <> vbNullString Then
Kill bfilename$
DoEvents
End If

Sleep 200

' Write batch file for running convolg.exe
tfilenumber% = FreeFile()
Open bfilename$ For Output As #tfilenumber%
astring$ = Left$(PENEPMA_Path$, 2)                             ' change to drive
Print #tfilenumber%, astring$
astring$ = "cd " & VbDquote$ & PENEPMA_Path$ & VbDquote$       ' change to folder
Print #tfilenumber%, astring$
If mode% = 0 Then
astring$ = "convolg.exe < " & "convolg.in"
If Dir$(PENEPMA_Path$ & "\Convolg.exe") = vbNullString Then GoTo Penepma08RunConvolgEXENotFound
ElseIf mode% = 1 Then
astring$ = "convolg_LIF.exe < " & "convolg.in"
If Dir$(PENEPMA_Path$ & "\Convolg_LIF.exe") = vbNullString Then GoTo Penepma08RunConvolgEXENotFound
ElseIf mode% = 2 Then
astring$ = "convolg_PET.exe < " & "convolg.in"
If Dir$(PENEPMA_Path$ & "\Convolg_PET.exe") = vbNullString Then GoTo Penepma08RunConvolgEXENotFound
ElseIf mode% = 3 Then
astring$ = "convolg_TAP.exe < " & "convolg.in"
If Dir$(PENEPMA_Path$ & "\Convolg_TAP.exe") = vbNullString Then GoTo Penepma08RunConvolgEXENotFound
ElseIf mode% = 4 Then
astring$ = "convolg_LDE.exe < " & "convolg.in"
If Dir$(PENEPMA_Path$ & "\Convolg_LDE.exe") = vbNullString Then GoTo Penepma08RunConvolgEXENotFound
End If
Print #tfilenumber%, astring$
Close #tfilenumber%

Sleep 200

' Run batch file asynchronously (/k executes but window remains, /c executes but terminates)
'taskID& = Shell("cmd.exe /k " & VbDquote$ & bfilename$ & VbDquote$, vbNormalFocus)
taskID& = Shell("cmd.exe /c " & VbDquote$ & bfilename$ & VbDquote$, vbMinimizedNoFocus)

' Loop until complete
Do Until IOIsProcessTerminated(taskID&)
Sleep 200
DoEvents
Loop

' Check for created convolved file
If Dir$(PENEPMA_CONVOLG_File$) = vbNullString Then GoTo Penepma08RunConvolgEXENoConvolgCreated

Exit Sub

' Errors
Penepma08RunConvolgEXEError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08RunConvolgEXE"
Close #tfilenumber%
ierror = True
Exit Sub

Penepma08RunConvolgEXENotSpecified:
msg$ = "The Convolg.exe input file was not specified."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08RunConvolgEXE"
ierror = True
Exit Sub

Penepma08RunConvolgEXEBadMode:
msg$ = "Invalid mode (" & Format$(mode%) & ") passed to procedure."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08RunConvolgEXE"
ierror = True
Exit Sub

Penepma08RunConvolgEXENotFound:
msg$ = "The specified Convolg*.exe file (mode= " & Format$(mode%) & ") was not found. " & vbCrLf & vbCrLf
msg$ = msg$ & "Please go to the Probe for EPMA Help | Update Probe for EPMA menu, and select the Update Penepma Monte Carlo Files Only option, and update yopur Penepma distribution for EDS and WDS spectrum simulation."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08RunConvolgEXE"
Close #tfilenumber%
ierror = True
Exit Sub

Penepma08RunConvolgEXENoConvolgCreated:
msg$ = "The Convolg.exe output file (" & PENEPMA_CONVOLG_File$ & ") was not created properly. Go to the " & PENEPMA_Path$ & " command prompt and type the following command to see the error: convolg.exe < " & "convolg.in"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08RunConvolgEXE"
Close #tfilenumber%
ierror = True
Exit Sub

Penepma08RunConvolgEXENoConvolgFilename:
msg$ = "The Convolg.exe output file name, PENEPMA_CONVOLG_File$, is blank.  This error should not occur, please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08RunConvolgEXE"
Close #tfilenumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08CreateConvolgFile(mode As Integer)
' Create Convolg input file (*.IN) from the passed parameters
'  mode = 0  use default EDS resolution (~140 eV)
'  mode = 1  use WDS LIF resolution (~4 eV)
'  mode = 2  use WDS PET resolution (~6 eV)
'  mode = 3  use WDS TAP resolution (~8 eV)
'  mode = 4  use WDS LDE resolution (~20 eV)

ierror = False
On Error GoTo Penepma08CreateConvolgFileError

Dim tfilename As String
Dim tfilenumber As Integer

' Create input file for Convolg.exe
tfilename$ = PENEPMA_Path$ & "\convolg.in"
tfilenumber% = FreeFile()
Open tfilename$ For Output As #tfilenumber%

' Output configuration
Print #tfilenumber%, MiscGetFileNameOnly$(PENEPMA_SPEC_File$)
If mode% = 0 Then
Print #tfilenumber%, Format$(DEMO_EDS_EVPERCHANNEL!)        ' always assume 20 eV per channel for EDS (0 to 20 keV range)
Else
Print #tfilenumber%, Format$(DEMO_WDS_EVPERCHANNEL!)        ' always assume 1 eV per channel for WDS (range can vary depending on simulation!)
End If
Print #tfilenumber%, vbNullString                           ' needs a CrLf

Close #tfilenumber%
Exit Sub

' Errors
Penepma08CreateConvolgFileError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08CreateConvolgFile"
Close #tfilenumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08GetNetIntensities(nlines As Long, penepma_iz() As Integer, penepma_s0() As String, penepma_s1() As String, penepma_eV() As Single, penepma_total() As Single, penepma_unc() As Single, penepma_bgd() As Single)
' Extract the net intensites from the EDS pe-intens-01.dat file (for demo EDS acquisition) (penepma_bgd!() is only dimensioned here)
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
On Error GoTo Penepma08GetNetIntensitiesError

Dim tfilename As String
Dim tfilenumber As Integer
Dim temp As Single
Dim astring As String, bstring As String

nlines& = 0

tfilename$ = PENEPMA_Path$ & "\pe-intens-01.dat"
tfilenumber% = FreeFile()
Open tfilename$ For Input As #tfilenumber%

' Load array (IZ, SO, S1, E (eV), P, unc, etc.
Do Until EOF(tfilenumber%)
Line Input #tfilenumber%, astring$
If Len(Trim$(astring$)) > 0 And InStr(astring$, "#") = 0 Then            ' skip to first data line (also skips if value is -1.#IND0E+000)

' Dimension arrays
nlines& = nlines& + 1
ReDim Preserve penepma_iz(1 To nlines&) As Integer
ReDim Preserve penepma_s0(1 To nlines&) As String
ReDim Preserve penepma_s1(1 To nlines&) As String
ReDim Preserve penepma_eV(1 To nlines&) As Single
ReDim Preserve penepma_total(1 To nlines&) As Single
ReDim Preserve penepma_bgd(1 To nlines&) As Single
ReDim Preserve penepma_unc(1 To nlines&) As Single

' Load k-ratio data
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
penepma_iz%(nlines&) = Val(Trim$(bstring$))                           ' IZ (atomic number)

' Load transition strings
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
penepma_s0$(nlines&) = Trim$(bstring$)                                ' S0 (inner transition)
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
penepma_s1$(nlines&) = Trim$(bstring$)                                ' S1 (outer transition)

' Load energy in eV
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
penepma_eV!(nlines&) = Val(Trim$(bstring$))                           ' E (eV)

' Parse primary intensity
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
temp! = Val(Trim$(bstring$))

' Parse primary intensity uncertainty
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
temp! = Val(Trim$(bstring$))

' Parse characteristic fluorescence
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
temp! = Val(Trim$(bstring$))

' Parse characteristic fluorescence intensity uncertainty
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
temp! = Val(Trim$(bstring$))

' Parse continuum fluorescence
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
temp! = Val(Trim$(bstring$))

' Parse characteristic fluorescence uncertainty
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
temp! = Val(Trim$(bstring$))

' Parse total fluorescence
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
temp! = Val(Trim$(bstring$))

' Parse total fluorescence uncertainty
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
temp! = Val(Trim$(bstring$))

' Parse total intensity
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
penepma_total!(nlines&) = Val(Trim$(bstring$))

' Parse total intensity uncertainty
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
penepma_unc!(nlines&) = Val(Trim$(bstring$))

End If
Loop

Close #tfilenumber%

Exit Sub

' Errors
Penepma08GetNetIntensitiesError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08GetNetIntensities"
Close #tfilenumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08GetBgdIntensities(nlines As Long, penepma_eV() As Single, penepma_bgd() As Single)
' Extract the generated bgd intensites from the EDS pe-gen-bremss.dat file (for demo EDS acquisition)
'
' Sample input file:
' #  Results from PENEPMA.
' #  Probability of emission of bremmstrahlung photons.
' #  1st column: E (eV).
' #  2nd and 3rd columns: probability density and STU (1/(eV*sr*electron)).
'
'   4.125000E+02  1.187693E-07  3.098989E-08
'   4.275000E+02  1.232177E-07  2.439716E-08
'   4.425000E+02  1.332971E-07  3.209410E-08
'   4.575000E+02  1.092279E-07  2.268812E-08

ierror = False
On Error GoTo Penepma08GetBgdIntensitiesError

Dim tfilename As String
Dim ncont As Long, n As Long
Dim tfilenumber As Integer
Dim astring As String, bstring As String

Dim energy_array() As Single
Dim bremms_array() As Single
Dim uncert_array() As Single

ncont& = 0

tfilename$ = PENEPMA_Path$ & "\pe-gen-bremss.dat"
tfilenumber% = FreeFile()
Open tfilename$ For Input As #tfilenumber%

' Load array of generated bgd continuum intensities
Do Until EOF(tfilenumber%)
Line Input #tfilenumber%, astring$
If Len(Trim$(astring$)) > 0 And InStr(astring$, "#") = 0 Then            ' skip to first data line (also skips if value is -1.#IND0E+000)

' Dimension interpolation arrays
ncont& = ncont& + 1
ReDim Preserve energy_array(1 To ncont&) As Single
ReDim Preserve bremms_array(1 To ncont&) As Single
ReDim Preserve uncert_array(1 To ncont&) As Single

' Load continuum energy
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
energy_array!(ncont&) = Val(Trim$(bstring$))                              ' generated bgd intensity

' Load generated continuum intensity
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
bremms_array!(ncont&) = Val(Trim$(bstring$))                               ' generated bgd intensity

' Load generated continuum intensity uncertainty
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
uncert_array!(ncont&) = Val(Trim$(bstring$))                              ' generated bgd intensity variance

End If
Loop

Close #tfilenumber%

' Now interpolate to get bgd intensity at each emission line energy
For n& = 1 To nlines&

' Return intensity for specified spectrometer position
penepma_bgd!(n&) = MathGetInterpolatedYValue(penepma_eV!(n&), CInt(ncont&), energy_array!(), bremms_array!())
If ierror Then Exit Sub

Next n&

Exit Sub

' Errors
Penepma08GetBgdIntensitiesError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08GetBgdIntensities"
Close #tfilenumber%
ierror = True
Exit Sub

End Sub



