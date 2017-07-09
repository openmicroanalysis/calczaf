Attribute VB_Name = "CodePenepma08a"
' (c) Copyright 1995-2017 by John J. Donovan
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
'PARAMETER (NEDM=25,NEDCM=1000)     ' changed to 32000 12-02-2016

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
' astring$ = original input line
' bstring$ = modified input line

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

Sub Penepma08GraphGetPenepma2(nPoints As Long, xdata() As Single, ydata() As Single)
' Load spectrum data from spectrum file (PENEPMA_CONVOLG_File$) (single precision version)

ierror = False
On Error GoTo Penepma08GraphGetPenepma2Error

Dim astring As String, bstring As String

' Check production file
If Trim$(PENEPMA_CONVOLG_File$) = vbNullString Then Exit Sub
If Dir$(Trim$(PENEPMA_CONVOLG_File$)) = vbNullString Then Exit Sub
Open PENEPMA_CONVOLG_File$ For Input As #Temp1FileNumber%

' Load array (npts&, xdata!(), ydata!())
nPoints& = 0
Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$
If Len(Trim$(astring$)) > 0 And InStr(astring$, "#") = 0 Then            ' skip to first data line

' Load total spectrum
nPoints& = nPoints& + 1
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
ReDim Preserve xdata(1 To nPoints&) As Single
ReDim Preserve ydata(1 To nPoints&) As Single
xdata!(nPoints&) = Val(Trim$(bstring$))
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
ydata!(nPoints&) = Val(Trim$(bstring$))

If VerboseMode And DebugMode Then
Call IOWriteLog("N=" & Format$(nPoints&) & ", X=" & Format$(xdata!(nPoints&)) & ", Y=" & Format$(ydata!(nPoints&), e104$))
End If

End If
Loop

Close #Temp1FileNumber%

Exit Sub

' Errors
Penepma08GraphGetPenepma2Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08GraphGetPenepma2"
Close #Temp1FileNumber%
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
If Dir$(Trim$(tfilename$)) = vbNullString Then GoTo Penepma08PenepmaSpectrumReadFileNotFound

' Open file
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

Penepma08PenepmaSpectrumReadFileNotFound:
msg$ = "The specified file (" & tfilename$ & ") was not found."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08PenepmaSpectrumRead"
ierror = True
Exit Sub

End Sub

Sub Penepma08PenepmaSpectrumRead2(tfilename As String, nPoints As Long, xdata() As Single, ydata() As Single, zdata() As Single, xmin As Single, xmax As Single, xnum As Long, xwidth As Single, theta1 As Single, theta2 As Single, phi1 As Single, phi2 As Single)
' Load spectrum data from passed Penepma spectrum file (for synthetic wavescan data in demo mode) (single precision version for increased speed)
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
On Error GoTo Penepma08PenepmaSpectrumRead2Error

Dim astring As String, bstring As String
Dim cstring As String, dstring As String
Dim tfilenumber As Integer

' Check production file
If Dir$(Trim$(tfilename$)) = vbNullString Then GoTo Penepma08PenepmaSpectrumRead2FileNotFound

' Open file
tfilenumber% = FreeFile()
Open tfilename$ For Input As #tfilenumber%

' Load array (nPoints&, xdata!(), ydata!())
nPoints& = 0
Do Until EOF(tfilenumber%)
Line Input #tfilenumber%, astring$
cstring$ = astring$

' Load theta parameters
If Len(Trim$(astring$)) > 0 And InStr(astring$, "theta_1") > 0 Then
Call MiscParseStringToStringA(cstring$, ",", dstring$)
If ierror Then Exit Sub
theta1! = Val(Trim$(Right$(dstring$, 12)))
theta2! = Val(Trim$(Right$(cstring$, 12)))

' Load phi parameters
ElseIf Len(Trim$(cstring$)) > 0 And InStr(cstring$, "phi_1") > 0 Then
Call MiscParseStringToStringA(cstring$, ",", dstring$)
If ierror Then Exit Sub
phi1! = Val(Trim$(Right$(dstring$, 12)))
phi2! = Val(Trim$(Right$(cstring$, 12)))

' Load energy window parameters
ElseIf Len(Trim$(cstring$)) > 0 And InStr(cstring$, "Energy window") > 0 Then
Call MiscParseStringToStringA(cstring$, ",", dstring$)
If ierror Then Exit Sub
xmin! = Val(Trim$(Right$(dstring$, 11)))
Call MiscParseStringToStringA(cstring$, ")", dstring$)
If ierror Then Exit Sub
xmax! = Val(Trim$(Right$(dstring$, 11)))
Call MiscParseStringToStringA(cstring$, "=", dstring$)
If ierror Then Exit Sub
xnum& = Val(Trim$(cstring$))

' Load channel resolution parameter
ElseIf Len(Trim$(cstring$)) > 0 And InStr(cstring$, "Channel width") > 0 Then
Call MiscParseStringToStringA(cstring$, "=", dstring$)
If ierror Then Exit Sub
xwidth! = Val(Trim$(Left$(cstring$, 13)))
End If

' Skip until comments are read
If Len(Trim$(astring$)) > 0 And InStr(astring$, "#") = 0 Then            ' skip to first data line

' Load total spectrum
nPoints& = nPoints& + 1
ReDim Preserve xdata(1 To nPoints&) As Single
ReDim Preserve ydata(1 To nPoints&) As Single
ReDim Preserve zdata(1 To nPoints&) As Single

Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
xdata!(nPoints&) = Val(Trim$(bstring$))

Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
ydata!(nPoints&) = Val(Trim$(bstring$))

Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
zdata!(nPoints&) = Val(Trim$(bstring$))

If VerboseMode And DebugMode Then
Call IOWriteLog("N=" & Format$(nPoints&) & ", X=" & Format$(xdata!(nPoints&)) & ", Y=" & Format$(ydata!(nPoints&), e104$) & ", Z=" & Format$(zdata!(nPoints&), e104$))
End If

End If
Loop

Close #tfilenumber%

Exit Sub

' Errors
Penepma08PenepmaSpectrumRead2Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08PenepmaSpectrumRead2"
Close #tfilenumber%
ierror = True
Exit Sub

Penepma08PenepmaSpectrumRead2FileNotFound:
msg$ = "The specified file (" & tfilename$ & ") was not found."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08PenepmaSpectrumRead2"
ierror = True
Exit Sub

End Sub

Sub Penepma08PenepmaSpectrumWrite(tfilename As String, npts As Long, xdata() As Double, ydata() As Double, zdata() As Double, xmin As Double, xmax As Double, xnum As Long, xwid As Double, theta1 As Double, theta2 As Double, phi1 As Double, phi2 As Double)
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

' Save array to disk (npts&, xdata#(), ydata#(), zdata#())
For n& = 1 To npts&
Print #tfilenumber%, " ", Format$(Format$(xdata#(n&), e104$), a14$), Format$(Format$(ydata#(n&), e104$), a14$), Format$(Format$(zdata#(n&), e104$), a14$)
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

Sub Penepma08PenepmaSpectrumWrite2(tfilename As String, npts As Long, xdata() As Single, ydata() As Single, zdata() As Single, xmin As Single, xmax As Single, xnum As Long, xwid As Single, theta1 As Single, theta2 As Single, phi1 As Single, phi2 As Single)
' Save spectrum data to passed Penepma spectrum file (for synthetic wavescan data in demo mode) (single precision version for increased speed)
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
On Error GoTo Penepma08PenepmaSpectrumWrite2Error

Dim tfilenumber As Integer
Dim n As Long

tfilenumber% = FreeFile()
Open tfilename$ For Output As #tfilenumber%

' Write header information
Print #tfilenumber%, " #  Results from PENEPMA. Output from photon detector #  1"
Print #tfilenumber%, " #"
Print #tfilenumber%, " #  Angular intervals : theta_1 = " & Format$(theta1!, e125$) & ",  theta_2 = " & Format$(theta2!, e125$)
Print #tfilenumber%, " #                        phi_1 = " & Format$(phi1!, e125$) & ",    phi_2 = " & Format$(phi2!, e125$)
Print #tfilenumber%, " #  Energy window = ( " & Format$(xmin!, e115$) & ", " & Format$(xmax!, e115$) & ") eV, no. of channels = " & Format$(xnum&)
Print #tfilenumber%, " #  Channel width = "; Format$(xwid!, e115$) & " eV"
Print #tfilenumber%, " #"
Print #tfilenumber%, " #  Whole spectrum. Characteristic peaks and background."
Print #tfilenumber%, " #  1st column: photon energy (eV)."
Print #tfilenumber%, " #  2nd column: probability density (1/(eV*sr*electron))."
Print #tfilenumber%, " #  3rd column: statistical uncertainty (3 sigma)."
Print #tfilenumber%, " #"

' Save array to disk (npts&, xdata!(), ydata!(), zdata!())
For n& = 1 To npts&
Print #tfilenumber%, " ", Format$(Format$(xdata!(n&), e104$), a14$), Format$(Format$(ydata!(n&), e104$), a14$), Format$(Format$(zdata!(n&), e104$), a14$)
Next n&

Close #tfilenumber%

Exit Sub

' Errors
Penepma08PenepmaSpectrumWrite2Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08PenepmaSpectrumWrite2"
Close #tfilenumber%
ierror = True
Exit Sub

End Sub

Sub Penepma08ExtractTransitionEnergy(ielm As Integer, iray As Integer, tfilename As String, tenergy As Double)
' Return the specified emission energy from the Penepma output (for synthetic wavescan data in demo mode)
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
On Error GoTo Penepma08ExtractTransitionEnergyError

Dim elementfound As Boolean
Dim l As Integer, tfilenumber As Integer
Dim astring As String, bstring As String, tstring As String
Dim atnum As Integer
Dim s0 As String, s1 As String
Dim eV As Single

tenergy# = 0#

' Check for element and xray number
If ielm% = 0 Then GoTo Penepma08ExtractTransitionEnergyZeroAtomicNumber
If iray% = 0 Then GoTo Penepma08ExtractTransitionEnergyZeroXrayNumber

' Check for file
If Dir$(Trim$(tfilename$)) = vbNullString Then GoTo Penepma08ExtractTransitionEnergyFileNotFound

' Open file and read (specified emission energy may not be present if the specified element is not in the composition)
tfilenumber% = FreeFile()
Open tfilename$ For Input As #tfilenumber%

' Load array (IZ, SO, S1, E (eV), P, unc, etc.
Do Until EOF(tfilenumber%)
Line Input #tfilenumber%, astring$
If Len(Trim$(astring$)) > 0 And InStr(astring$, "#") = 0 Then            ' skip to first data line (also skips if value is -1.#IND0E+000)

' Load k-ratio data
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
atnum% = Val(Trim$(bstring$))                        ' IZ (atomic number)

' Check atomic number
If atnum% = ielm% Then

' Load transition strings
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
s0$ = Trim$(bstring$)                                ' S0 (inner transition)
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
s1$ = Trim$(bstring$)                                ' S1 (outer transition)

Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
eV! = Val(Trim$(bstring$))                           ' E (eV)

' Load x-ray index for x-ray line
l% = 0
tstring$ = s0$ & " " & s1$
If tstring$ = "K L3" Then l% = 1          ' (Ka) (see table 6.2 in Penelope-2006-NEA-pdf)
If tstring$ = "K M3" Then l% = 2          ' (Kb)
If tstring$ = "L3 M5" Then l% = 3         ' (La)
If tstring$ = "L2 M4" Then l% = 4         ' (Lb)
If tstring$ = "M5 N7" Then l% = 5         ' (Ma)
If tstring$ = "M4 N6" Then l% = 6         ' (Mb)

If tstring$ = "L2 M1" Then l% = 7         ' (Ln)
If tstring$ = "L2 N4" Then l% = 8         ' (Lg)
If tstring$ = "L2 N6" Then l% = 9         ' (Lv)
If tstring$ = "L3 M1" Then l% = 10        ' (Ll)
If tstring$ = "M3 N5" Then l% = 11        ' (Mg)
If tstring$ = "M5 N3" Then l% = 12        ' (Mz)

' Skip if not the specified primary line
If l% > 0 And l% = iray% Then
tenergy# = eV!
Close #tfilenumber%
Exit Sub

' Parse primary intensity
'Call MiscParseStringToString(astring$, bstring$)
'If ierror Then Exit Function
'pri_int!(n%, l%) = Val(Trim$(bstring$))

' Parse primary intensity uncertainty
'Call MiscParseStringToString(astring$, bstring$)
'If ierror Then Exit Function

' Parse characteristic fluorescence
'Call MiscParseStringToString(astring$, bstring$)
'If ierror Then Exit Function
'flch_int!(n%, l%) = Val(Trim$(bstring$))

' Parse characteristic fluorescence intensity uncertainty
'Call MiscParseStringToString(astring$, bstring$)
'If ierror Then Exit Function

' Parse continuum fluorescence
'Call MiscParseStringToString(astring$, bstring$)
'If ierror Then Exit Function
'flbr_int!(n%, l%) = Val(Trim$(bstring$))

' Parse characteristic fluorescence uncertainty
'Call MiscParseStringToString(astring$, bstring$)
'If ierror Then Exit Function

' Parse total fluorescence
'Call MiscParseStringToString(astring$, bstring$)
'If ierror Then Exit Function
'flu_int!(n%, l%) = Val(Trim$(bstring$))

' Parse total fluorescence uncertainty
'Call MiscParseStringToString(astring$, bstring$)
'If ierror Then Exit Function

' Parse total intensity
'Call MiscParseStringToString(astring$, bstring$)
'If ierror Then Exit Function
'tot_int!(n%, l%) = Val(Trim$(bstring$))

' Parse total intensity uncertainty
'Call MiscParseStringToString(astring$, bstring$)
'If ierror Then Exit Function
'tot_int_var!(n%, l%) = Val(Trim$(bstring$))
End If

End If
End If
Loop

' Specified element and x-ray not found (not present in file composition)
Close #tfilenumber%

Exit Sub

' Errors
Penepma08ExtractTransitionEnergyError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08ExtractTransitionEnergy"
Close #tfilenumber%
ierror = True
Exit Sub

Penepma08ExtractTransitionEnergyZeroAtomicNumber:
msg$ = "Zero atomic number passed for file " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08ExtractTransitionEnergy"
ierror = True
Exit Sub

Penepma08ExtractTransitionEnergyZeroXrayNumber:
msg$ = "Zero xray number passed for file " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08ExtractTransitionEnergy"
ierror = True
Exit Sub

Penepma08ExtractTransitionEnergyFileNotFound:
msg$ = "The specified file (" & tfilename$ & ") was not found."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08ExtractTransitionEnergy"
ierror = True
Exit Sub

End Sub
