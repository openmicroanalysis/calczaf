Attribute VB_Name = "CodePenepma08a"
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Const COL7% = 7

'The maximum number of channels can be changed in Penepma.f where (NEDCM) can be changed at line 464:
'PARAMETER (NEDM=25,NEDCM=1000)     ' changed to 20000 12-02-2016

Global Const PENEPMA_MINPERCENT! = 0.0001

' Display (output) files
Global PENEPMA_DAT_File As String
Global PENEPMA_SPEC_File As String
Global PENEPMA_CHAR_File As String
Global PENEPMA_EL_TRANS_File As String

Global PENEPMA_CONVOLG_File As String

Sub Penepma08Init0()
' Initialie the Penepma display files and delete dump files

ierror = False
On Error GoTo Penepma08Init0Error

' Load Penepma output (display) files
PENEPMA_DAT_File$ = PENEPMA_Path$ & "\PENEPMA.DAT"
PENEPMA_SPEC_File$ = PENEPMA_Path$ & "\PE-SPECT-01.DAT"
PENEPMA_CHAR_File$ = PENEPMA_Path$ & "\PE-CHARACT-01.DAT"

If Penepma08CheckPenepmaVersion%() = 8 Then
PENEPMA_EL_TRANS_File$ = PENEPMA_Path$ & "\PE-ENERGY-EL-TRANS.DAT"
Else
PENEPMA_EL_TRANS_File$ = PENEPMA_Path$ & "\PE-ENERGY-EL-UP.DAT"
End If

PENEPMA_CONVOLG_File$ = PENEPMA_Path$ & "\CHSPECT.DAT"

' Delete dump files
Call Penepma08DeleteDumpFiles
If ierror Then Exit Sub

' Close temp file handles in case they are already open
Close #Temp1FileNumber%
Close #Temp2FileNumber%

' Delete existing output files
If Dir$(PENEPMA_DAT_File$) <> vbNullString Then Kill PENEPMA_DAT_File$
If Dir$(PENEPMA_SPEC_File$) <> vbNullString Then Kill PENEPMA_SPEC_File$
If Dir$(PENEPMA_CHAR_File$) <> vbNullString Then Kill PENEPMA_CHAR_File$
If Dir$(PENEPMA_EL_TRANS_File$) <> vbNullString Then Kill PENEPMA_EL_TRANS_File$
If Dir$(PENEPMA_CONVOLG_File$) <> vbNullString Then Kill PENEPMA_CONVOLG_File$
Sleep (400)

Exit Sub

' Errors
Penepma08Init0Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08Init0"
ierror = True
Exit Sub
End Sub

Sub Penepma08CreateMaterialFile(sample() As TypeSample)
' Create a Penepma material file based on the passed sample

ierror = False
On Error GoTo Penepma08CreateMaterialFileError

Dim i As Integer
Dim tfilename As String
Dim astring As String

' Load input file name
tfilename$ = PENDBASE_Path$ & "\" & "material" & ".inp"
Open tfilename$ For Output As #Temp1FileNumber%

' Output configuration
Print #Temp1FileNumber%, "1"                             ' enter composition from keyboard
Print #Temp1FileNumber%, Left$(sample(1).Name$, 60)      ' material name

If sample(1).LastChan% = 0 Then
Print #Temp1FileNumber%, Format$("1")    ' number of elements in composition
Else
Print #Temp1FileNumber%, Format$(sample(1).LastChan%)    ' number of elements in composition
End If

' If no elements, just enter Si
If sample(1).LastChan% = 0 Then
Print #Temp1FileNumber%, Format$("14") & VbComma$ & Format$("1")

' If more than one element enter composition
ElseIf sample(1).LastChan% = 1 Then
Print #Temp1FileNumber%, Format$(sample(1).AtomicNums%(1)) & VbComma$ & Format$("1")

Else
Print #Temp1FileNumber%, "2"   ' enter by weight fraction

' Output composition of material
For i% = 1 To sample(1).LastChan%
If sample(1).ElmPercents!(i%) < PENEPMA_MINPERCENT! Then sample(1).ElmPercents!(i%) = PENEPMA_MINPERCENT!
astring$ = Format$(sample(1).AtomicNums%(i%)) & VbComma$ & Trim$(MiscAutoFormat$(sample(1).ElmPercents!(i%) / 100#))
Print #Temp1FileNumber%, astring$
Next i%
End If

Print #Temp1FileNumber%, "2"                                    ' do not change mean excitation energy
Print #Temp1FileNumber%, Trim$(Str$(sample(1).SampleDensity!))
Print #Temp1FileNumber%, "2"                                    ' do not change oscillator strength and energy

astring$ = "material" & ".mat"                                  ' use same folder as MATERIAL.EXE
Print #Temp1FileNumber%, Left$(astring$, 80)                    ' material filename
Close #Temp1FileNumber%

Exit Sub

' Errors
Penepma08CreateMaterialFileError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08CreateMaterialFile"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08RunMaterialFile()
' Create and run material batch file

ierror = False
On Error GoTo Penepma08RunMaterialFileError

Dim taskID As Long
Dim bfilename As String, astring As String

' Run material.inp file
Call IOStatusAuto("Creating material file for Penepma Monte-Carlo calculations...")
DoEvents

' Delete existing temp batch file (use temp1.bat since other Penepma programs use temp.bat)
bfilename$ = PENDBASE_Path$ & "\temp0.bat"
If Dir$(bfilename$) <> vbNullString Then
Kill bfilename$
DoEvents
End If

' Delete existing material file if it exists (to allow for proper check if created below)
If Dir$(PENDBASE_Path$ & "\" & "material" & ".mat") <> vbNullString Then
Kill PENDBASE_Path$ & "\" & "material" & ".mat"
DoEvents
End If

' Write batch file for running material.exe
Open bfilename$ For Output As #Temp1FileNumber%

astring$ = Left$(PENDBASE_Path$, 2)                             ' change to drive
Print #Temp1FileNumber%, astring$
astring$ = "cd " & VbDquote$ & PENDBASE_Path$ & VbDquote$       ' change to folder
Print #Temp1FileNumber%, astring$
astring$ = "material.exe < " & "material" & ".inp"
Print #Temp1FileNumber%, astring$
Close #Temp1FileNumber%

' Run batch file asynchronously (/k executes but window remains, /c executes but terminates)
'taskID& = Shell("cmd.exe /k " & VbDquote$ & bfilename$ & VbDquote$, vbNormalFocus)
taskID& = Shell("cmd.exe /c " & VbDquote$ & bfilename$ & VbDquote$, vbMinimizedNoFocus)

' Loop until complete
Do Until IOIsProcessTerminated(taskID&)
Sleep 100
DoEvents
Loop

' Check for created material file
If Dir$(PENDBASE_Path$ & "\" & "material.mat") = vbNullString Then GoTo Penepma08RunMaterialFileNoMaterialCreated

Call IOStatusAuto(vbNullString)
Exit Sub

' Errors
Penepma08RunMaterialFileError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08RunMaterialFile"
Close #Temp1FileNumber%
ierror = True
Exit Sub

Penepma08RunMaterialFileNoMaterialCreated:
msg$ = "The specified material file was not created properly. Go to the " & PENDBASE_Path$ & " command prompt and type the following command to see the error: material.exe < " & "material.inp"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08RunMaterialFile"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08CreateInputFile2(astring As String, bstring As String, cstring As String, dstring As String)
' Make the output string based on new value and current string

ierror = False
On Error GoTo Penepma08CreateInputFile2Error

dstring$ = Mid$(astring$, InStr(astring$, "["))
bstring$ = Left$(astring, COL7%) & cstring$ & Space$(Len(astring$) - COL7% - Len(cstring$) - Len(dstring$)) & dstring$

Exit Sub

' Errors
Penepma08CreateInputFile2Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08CreateInputFile2"
Close #Temp1FileNumber%
Close #Temp2FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08CreatePenepmaFile(tLiveTime As Single, sample() As TypeSample)
' Create Penepma input file (*.IN) from the passed sample

ierror = False
On Error GoTo Penepma08CreatePenepmaFileError

Dim astring As String, bstring As String, cstring As String, dstring As String
Dim tfilename As String, tfilename2 As String

Const tSimulatedShowers# = 2000000000#

Dim tBeamMinimumEnergyRange As Double
Dim tBeamMaximumEnergyRange As Double
Dim tBeamNumberOfEnergyChannels As Long

tBeamMinimumEnergyRange# = 0#
tBeamMaximumEnergyRange# = 20000#
tBeamNumberOfEnergyChannels& = 1000

' Loop through sample production file and copy to new file with modified parameters
tfilename$ = PENEPMA_Root$ & "\Cu_cha.in"
Open tfilename$ For Input As #Temp1FileNumber%
tfilename2$ = PENEPMA_Path$ & "\material.in"
Open tfilename2$ For Output As #Temp2FileNumber%

Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$
bstring$ = astring$

If InStr(astring$, "TITLE") > 0 Then bstring$ = Left$(astring, COL7%) & Left$(sample(1).Name$, 120)

cstring$ = Format$(Format$(sample(1).kilovolts! * EVPERKEV#, "Scientific"), a10$)
If InStr(astring$, "SENERG") > 0 Then Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

cstring$ = Format$(0#) & " " & Format$(0#) & " " & Format$(1#)
If InStr(astring$, "SPOSIT") > 0 Then Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

cstring$ = Format$(180#) & " " & Format$(0#)
If InStr(astring$, "SDIREC") > 0 Then Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

cstring$ = Format$(0#)
If InStr(astring$, "SAPERT") > 0 Then Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

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
If Dir$(PENEPMA_Path$ & "\bulk.geo") = vbNullString Then
If Dir$(PENEPMA_Root$ & "\bulk.geo") = vbNullString Then GoTo Penepma08CreatePenepmaFileNotFoundGEO
FileCopy PENEPMA_Root$ & "\bulk.geo", PENEPMA_Path$ & "\bulk.geo"
End If
cstring$ = MiscGetFileNameOnly$(PENEPMA_Path$ & "\bulk.geo")
If InStr(astring$, "GEOMFN") > 0 Then Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

' Detector angles (45.0, 55.0, 0.0, 360.0, 0)
If InStr(astring$, "PDANGL") > 0 Then
cstring$ = Format$(90# - (sample(1).takeoff! + 5#), "0.0") & " " & Format$(90# - (sample(1).takeoff! - 5#), "0.0") & " "
cstring$ = cstring$ & Format$(0#, "0.0") & " " & Format$(360#, "0.0") & " "
cstring$ = cstring$ & Format$(0, "0")
Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub
End If

' Spectrum energy range and channels
If InStr(astring$, "PDENER") > 0 Then
If sample(1).kilovolts! * EVPERKEV# > tBeamMaximumEnergyRange# Then tBeamMaximumEnergyRange# = 30# * EVPERKEV#
cstring$ = Format$(tBeamMinimumEnergyRange#, "0.0") & " " & Format$(tBeamMaximumEnergyRange#, "0.0") & " "
cstring$ = cstring$ & Format$(tBeamNumberOfEnergyChannels&, "0")
Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
End If

' Dump time (15 seconds for now)
cstring$ = Format$(15#)
If InStr(astring$, "DUMPP") > 0 Then Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

' Number of simulated showers (make arbitrarily large number)
cstring$ = Format$(tSimulatedShowers#, e71$)
If InStr(astring$, "NSIMSH") > 0 Then Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

cstring$ = Format$(tLiveTime!)
If InStr(astring$, "TIME") > 0 Then Call Penepma08CreateInputFile2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

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

Sub Penepma08CopyMaterialFile(mode As Integer, tfilename As String)
' Copy the newly created material file to the user probe data folder (and Penepma folder)

ierror = False
On Error GoTo Penepma08CopyMaterialFileError

' Now copy all files to original material names (up to 20 characters)
Call IOStatusAuto("Copying material file to " & PENEPMA_Path$ & " folder...")
DoEvents

' Copy material file to PENEPMA path using passed name
FileCopy PENDBASE_Path$ & "\material.mat", PENEPMA_Path$ & "\" & MiscGetFileNameOnly$(tfilename$)

' Copy material file to probe data path
If mode% = 0 Then
Call IOStatusAuto("Copying material file to probe data folder...")
DoEvents

FileCopy PENDBASE_Path$ & "\material.mat", tfilename$

' Confirm with user
msg$ = "Material file " & MiscGetFileNameOnly$(tfilename$) & " was copied to " & MiscGetPathOnly$(tfilename$)
MsgBox msg$, vbOKOnly + vbInformation, "Penepma08CopyMaterialFile"
End If

Call IOStatusAuto(vbNullString)
Exit Sub

' Errors
Penepma08CopyMaterialFileError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08CopyMaterialFile"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08RunPenepmaFile(PenepmaTaskID As Long)
' Run the Penepma input file (and monitor using returned taskID)

ierror = False
On Error GoTo Penepma08RunPenepmaFileError

Dim bfilename As String, astring As String

' Run material.inp file
If DebugMode Then
Call IOStatusAuto("Calculating Penepma Monte-Carlo EDS spectrum...")
DoEvents
End If

' Init file names and delete dump files
Call Penepma08Init0
If ierror Then Exit Sub

' Delete existing output files
If Dir$(PENEPMA_DAT_File$) <> vbNullString Then Kill PENEPMA_DAT_File$
If Dir$(PENEPMA_SPEC_File$) <> vbNullString Then Kill PENEPMA_SPEC_File$
If Dir$(PENEPMA_CHAR_File$) <> vbNullString Then Kill PENEPMA_CHAR_File$
If Dir$(PENEPMA_EL_TRANS_File$) <> vbNullString Then Kill PENEPMA_EL_TRANS_File$
Sleep (400)

' Create batch file to run Penepma (use temp1.bat since other Penepma programs use temp.bat)
bfilename$ = PENEPMA_Path$ & "\temp1.bat"
Open bfilename$ For Output As #Temp2FileNumber%

astring$ = Left$(PENEPMA_Path$, 2)                             ' change to drive
Print #Temp2FileNumber%, astring$
astring$ = "cd " & VbDquote$ & PENEPMA_Path$ & VbDquote$       ' change to folder
Print #Temp2FileNumber%, astring$
astring$ = "Penepma " & ChrW$(60) & " material.in"
Print #Temp2FileNumber%, astring$
Close #Temp2FileNumber%

' Start Penepma (/k executes but window remains, /c executes but terminates)
'PenepmaTaskID& = Shell("cmd.exe /k " & VbDquote$ & bfilename$ & VbDquote$, vbNormalFocus)
PenepmaTaskID& = Shell("cmd.exe /c " & VbDquote$ & bfilename$ & VbDquote$, vbMinimizedNoFocus)

Exit Sub

' Errors
Penepma08RunPenepmaFileError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08RunPenepmaFile"
Close #Temp2FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08DeleteDumpFiles()
' Delete the dump files (if present)

ierror = False
On Error GoTo Penepma08DeleteDumpFilesError

If Dir$(PENEPMA_Path$ & "\dump*.dat") <> vbNullString Then Kill PENEPMA_Path$ & "\dump*.dat"

Exit Sub

' Errors
Penepma08DeleteDumpFilesError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08DeleteDumpFiles"
ierror = True
Exit Sub

End Sub

Sub Penepma08RunConvolgEXE(eVPerChannel As Single)
' Run the Convolg program to convolve the Penepma spectrum

ierror = False
On Error GoTo Penepma08RunConvolgEXEError

Dim tfilenumber As Integer
Dim taskID As Long
Dim bfilename As String, astring As String

' First check if Penepma has had time to create a spectrum file
If PENEPMA_SPEC_File$ = vbNullString Then GoTo Penepma08RunConvolgEXENotSpecified
If Dir$(PENEPMA_SPEC_File$) = vbNullString Then Exit Sub

' Create Convg input file
Call Penepma08CreateConvolgFile(eVPerChannel!)
If ierror Then Exit Sub

' Delete existing temp batch file (use temp2.bat because Penepma is using temp.bat)
bfilename$ = PENDBASE_Path$ & "\temp2.bat"
If Dir$(bfilename$) <> vbNullString Then
Kill bfilename$
DoEvents
End If

' Write batch file for running convolg.exe
tfilenumber% = FreeFile()
Open bfilename$ For Output As #tfilenumber%
astring$ = Left$(PENEPMA_Path$, 2)                             ' change to drive
Print #tfilenumber%, astring$
astring$ = "cd " & VbDquote$ & PENEPMA_Path$ & VbDquote$       ' change to folder
Print #tfilenumber%, astring$
astring$ = "convolg.exe < " & "convolg.in"
Print #tfilenumber%, astring$
Close #tfilenumber%

' Run batch file asynchronously (/k executes but window remains, /c executes but terminates)
'taskID& = Shell("cmd.exe /k " & VbDquote$ & bfilename$ & VbDquote$, vbNormalFocus)
taskID& = Shell("cmd.exe /c " & VbDquote$ & bfilename$ & VbDquote$, vbMinimizedNoFocus)

' Loop until complete
Do Until IOIsProcessTerminated(taskID&)
Sleep 100
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
Close #tfilenumber%
ierror = True
Exit Sub

Penepma08RunConvolgEXENoConvolgCreated:
msg$ = "The Convolg.exe output file (" & PENEPMA_CONVOLG_File$ & ") was not created properly. Go to the " & PENEPMA_Path$ & " command prompt and type the following command to see the error: convolg.exe < " & "convolg.in"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08RunConvolgEXE"
Close #tfilenumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08GraphGetPenepma(nPoints As Long, xdata() As Double, ydata() As Double)
' Load spectrum data from spectrum file (PENEPMA_CONVOLG_File$)

ierror = False
On Error GoTo Penepma08GraphGetPenepmaError

Dim astring As String, bstring As String

' Check production file
If Trim$(PENEPMA_CONVOLG_File$) = vbNullString Then Exit Sub
If Dir$(Trim$(PENEPMA_CONVOLG_File$)) = vbNullString Then Exit Sub
Open PENEPMA_CONVOLG_File$ For Input As #Temp1FileNumber%

' Load array (npts&, xdata#(), ydata#())
nPoints& = 0
Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$
If Len(Trim$(astring$)) > 0 And InStr(astring$, "#") = 0 Then            ' skip to first data line

' Load total spectrum
nPoints& = nPoints& + 1
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
ReDim Preserve xdata(1 To nPoints&) As Double
ReDim Preserve ydata(1 To nPoints&) As Double
xdata#(nPoints&) = Val(Trim$(bstring$))
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
ydata#(nPoints&) = Val(Trim$(bstring$))

If VerboseMode And DebugMode Then
Call IOWriteLog("N=" & Format$(nPoints&) & ", X=" & Format$(xdata#(nPoints&)) & ", Y=" & Format$(ydata#(nPoints&), e104$))
End If

End If
Loop

Close #Temp1FileNumber%

Exit Sub

' Errors
Penepma08GraphGetPenepmaError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08GraphGetPenepma"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08CreateConvolgFile(eVPerChannel As Single)
' Create Convolg input file (*.IN) from the passed parameters

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
Print #tfilenumber%, Format$(eVPerChannel!)
Print #tfilenumber%, vbNullString               ' needs a crlf

Close #tfilenumber%
Exit Sub

' Errors
Penepma08CreateConvolgFileError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08CreateConvolgFile"
Close #tfilenumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08PenepmaSpectrumRead(tfilename As String, nPoints As Long, xdata() As Double, ydata() As Double, zdata() As Double, xmin As Double, xmax As Double, xnum As Long, xwidth As Double, theta1 As Double, theta2 As Double, phi1 As Double, phi2 As Double)
' Load spectrum data from passed Penepma spectrum file (for synthetic wavescan data in demo mode)
'
' Default header shown here:
' #  Results from PENEPMA. Output from photon detector #  1
' #
' #  Angular intervals : theta_1 = 4.500000E+01,  theta_2 = 5.500000E+01
' #                        phi_1 = 0.000000E+00,    phi_2 = 3.600000E+02
' #  Energy window = ( 0.00000E+00, 2.00000E+04) eV, no. of channels = 1000
' #  Channel width = 2.000000E+01 eV
' #
' #  Whole spectrum. Characteristic peaks and background.
' #  1st column: photon energy (eV).
' #  2nd column: probability density (1/(eV*sr*electron)).
' #  3rd column: statistical uncertainty (3 sigma).
' #
'   1.000000E+01  1.000000E-35  1.000000E-35
'   3.000000E+01  1.000000E-35  1.000000E-35
'   5.000001E+01  1.000000E-35  1.000000E-35
'   7.000001E+01  1.000000E-35  1.000000E-35
'   9.000001E+01  1.000000E-35  1.000000E-35
' ....

ierror = False
On Error GoTo Penepma08PenepmaSpectrumReadError

Dim astring As String, bstring As String
Dim cstring As String, dstring As String
Dim tfilenumber As Integer

' Check production file
If Dir$(Trim$(tfilename$)) = vbNullString Then Exit Sub
tfilenumber% = FreeFile()
Open tfilename$ For Input As #tfilenumber%

' Load array (nPoints&, xdata#(), ydata#())
nPoints& = 0
Do Until EOF(tfilenumber%)
Line Input #tfilenumber%, astring$
cstring$ = astring$

' Load theta parameters
If Len(Trim$(astring$)) > 0 And InStr(astring$, "theta_1") > 0 Then
Call MiscParseStringToStringA(cstring$, ",", dstring$)
If ierror Then Exit Sub
theta1# = Val(Trim$(Right$(dstring$, 12)))
theta2# = Val(Trim$(Right$(cstring$, 12)))

' Load phi parameters
ElseIf Len(Trim$(cstring$)) > 0 And InStr(cstring$, "phi_1") > 0 Then
Call MiscParseStringToStringA(cstring$, ",", dstring$)
If ierror Then Exit Sub
phi1# = Val(Trim$(Right$(dstring$, 12)))
phi2# = Val(Trim$(Right$(cstring$, 12)))

' Load energy window parameters
ElseIf Len(Trim$(cstring$)) > 0 And InStr(cstring$, "Energy window") > 0 Then
Call MiscParseStringToStringA(cstring$, ",", dstring$)
If ierror Then Exit Sub
xmin# = Val(Trim$(Right$(dstring$, 11)))
Call MiscParseStringToStringA(cstring$, ")", dstring$)
If ierror Then Exit Sub
xmax# = Val(Trim$(Right$(dstring$, 11)))
Call MiscParseStringToStringA(cstring$, "=", dstring$)
If ierror Then Exit Sub
xnum& = Val(Trim$(cstring$))

' Load channel resolution parameter
ElseIf Len(Trim$(cstring$)) > 0 And InStr(cstring$, "Channel width") > 0 Then
Call MiscParseStringToStringA(cstring$, "=", dstring$)
If ierror Then Exit Sub
xwidth# = Val(Trim$(Left$(cstring$, 13)))
End If

' Skip until comments are read
If Len(Trim$(astring$)) > 0 And InStr(astring$, "#") = 0 Then            ' skip to first data line

' Load total spectrum
nPoints& = nPoints& + 1
ReDim Preserve xdata(1 To nPoints&) As Double
ReDim Preserve ydata(1 To nPoints&) As Double
ReDim Preserve zdata(1 To nPoints&) As Double

Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
xdata#(nPoints&) = Val(Trim$(bstring$))

Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
ydata#(nPoints&) = Val(Trim$(bstring$))

Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
zdata#(nPoints&) = Val(Trim$(bstring$))

If VerboseMode And DebugMode Then
Call IOWriteLog("N=" & Format$(nPoints&) & ", X=" & Format$(xdata#(nPoints&)) & ", Y=" & Format$(ydata#(nPoints&), e104$) & ", Z=" & Format$(zdata#(nPoints&), e104$))
End If

End If
Loop

Close #tfilenumber%

Exit Sub

' Errors
Penepma08PenepmaSpectrumReadError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08PenepmaSpectrumRead"
Close #tfilenumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08PenepmaSpectrumWrite(tfilename As String, chan As Integer, npts() As Long, xdata() As Double, ydata() As Double, zdata() As Double, xmin As Double, xmax As Double, xnum As Long, xwid As Double, theta1 As Double, theta2 As Double, phi1 As Double, phi2 As Double)
' Save spectrum data to passed Penepma spectrum file (for synthetic wavescan data in demo mode)
'
' Default header shown here:
' #  Results from PENEPMA. Output from photon detector #  1
' #
' #  Angular intervals : theta_1 = 4.500000E+01,  theta_2 = 5.500000E+01
' #                        phi_1 = 0.000000E+00,    phi_2 = 3.600000E+02
' #  Energy window = ( 0.00000E+00, 2.00000E+04) eV, no. of channels = 1000
' #  Channel width = 2.000000E+01 eV
' #
' #  Whole spectrum. Characteristic peaks and background.
' #  1st column: photon energy (eV).
' #  2nd column: probability density (1/(eV*sr*electron)).
' #  3rd column: statistical uncertainty (3 sigma).
' #
'   1.000000E+01  1.000000E-35  1.000000E-35
'   3.000000E+01  1.000000E-35  1.000000E-35
'   5.000001E+01  1.000000E-35  1.000000E-35
'   7.000001E+01  1.000000E-35  1.000000E-35
'   9.000001E+01  1.000000E-35  1.000000E-35
' ....

ierror = False
On Error GoTo Penepma08PenepmaSpectrumWriteError

Dim tfilenumber As Integer
Dim n As Long

tfilenumber% = FreeFile()
Open tfilename$ For Output As #tfilenumber%

' Write header information
Print #tfilenumber%, " #  Results from PENEPMA. Output from photon detector #  1"
Print #tfilenumber%, " #"
Print #tfilenumber%, " #  Angular intervals : theta_1 = " & Format$(theta1#, e125$) & ",  theta_2 = " & Format$(theta2#, e125$)
Print #tfilenumber%, " #                        phi_1 = " & Format$(phi1#, e125$) & ",    phi_2 = " & Format$(phi2#, e125$)
Print #tfilenumber%, " #  Energy window = ( " & Format$(xmin#, e115$) & ", " & Format$(xmax, e115$) & ") eV, no. of channels = " & Format$(xnum&)
Print #tfilenumber%, " #  Channel width = "; Format$(xwid#, e115$) & " eV"
Print #tfilenumber%, " #"
Print #tfilenumber%, " #  Whole spectrum. Characteristic peaks and background."
Print #tfilenumber%, " #  1st column: photon energy (eV)."
Print #tfilenumber%, " #  2nd column: probability density (1/(eV*sr*electron))."
Print #tfilenumber%, " #  3rd column: statistical uncertainty (3 sigma)."
Print #tfilenumber%, " #"

' Save array to disk (nPoints&, xdata#(), ydata#(), zdata#())
For n& = 1 To npts&(chan%)
Print #tfilenumber%, " ", Format$(Format$(xdata#(chan%, n&), e104$), a14$), Format$(Format$(ydata#(chan%, n&), e104$), a14$), Format$(Format$(zdata#(chan%, n&), e104$), a14$)
Next n&

Close #tfilenumber%

Exit Sub

' Errors
Penepma08PenepmaSpectrumWriteError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08PenepmaSpectrumWrite"
Close #tfilenumber%
ierror = True
Exit Sub

End Sub

