Attribute VB_Name = "CodePenepma12Matrix"
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Global Const MINIMUMOVERVOLTFRACTION_02! = 0.02     ' 2 percent overvoltage
Global Const MINIMUMOVERVOLTFRACTION_10! = 0.1      ' 10 percent overvoltage
Global Const MINIMUMOVERVOLTFRACTION_20! = 0.2      ' 20 percent overvoltage
Global Const MINIMUMOVERVOLTFRACTION_40! = 0.4      ' 40 percent overvoltage

' Matrix globals
Global CalcZAF_ZAF_Kratios() As Double
Global CalcZAF_ZA_Kratios() As Double
Global CalcZAF_F_Kratios() As Double

Global Binary_ZAF_Kratios() As Double
Global Binary_ZA_Kratios() As Double
Global Binary_F_Kratios() As Double

' Pure element globals
Global PureGenerated_Intensities() As Double
Global PureEmitted_Intensities() As Double

' Temporary flag to round keV to nearest integer
Const Penepma12UseKeVRoundingFlag = True

Sub Penepma12CalculateReadWriteBinaryDataMatrix(mode As Integer, tfolder As String, tfilename As String, keV As Single)
' Reads or write the binary fluorescence matrix k-ratio data to or from a data file for a specified beam energy
'  mode = 0 create file and write column labels only
'  mode = 1 read data
'  mode = 2 write data
'  tfolder$ is the full path of the binary compositional data file to read or write
'  tfilename$ is the filename of the binary compositional data file to read or write
'  keV is the specified beam energy
'
'  BinaryRanges!(1 to MAXBINARY%)  are the compositional binaries (always 99 to 1 wt%)
'
'  Binary_ZAF_Kratios#(1 to MAXRAY%, 1 to MAXBINARY%)  are the full k-ratios from Fanal in k-ratio % for each x-ray and binary composition
'  CalcZAF_ZAF_Kratios#(1 to MAXRAY%, 1 to MAXBINARY%)  are the full k-ratios from CalcZAF in k-ratio % for each x-ray and binary composition
'
'  Binary_ZA_Kratios#(1 to MAXRAY%, 1 to MAXBINARY%)  are the ZA only k-ratios from Fanal in k-ratio % for each x-ray and binary composition
'  CalcZAF_ZA_Kratios#(1 to MAXRAY%, 1 to MAXBINARY%)  are the ZA only k-ratios from CalcZAF in k-ratio % for each x-ray and binary composition
'
'  Binary_F_Kratios#(1 to MAXRAY%, 1 to MAXBINARY%)  are the fluorescence only k-ratios from Fanal in k-ratio % for each x-ray and binary composition
'  CalcZAF_F_Kratios#(1 to MAXRAY%, 1 to MAXBINARY%)  are the fluorescence only k-ratios from CalcZAF in k-ratio % for each x-ray and binary composition

ierror = False
On Error GoTo Penepma12CalculateReadWriteBinaryDataMatrixError

Dim l As Integer, n As Integer, i As Integer
Dim tfilenumber As Integer
Dim tkeV As Single
Dim astring As String, ttfilename As String

' Init k-ratio globals not writing (to dimension arrays)
If mode% <> 2 Then
Call InitKratios
If ierror Then Exit Sub
End If

' Write column labels only
If mode% = 0 Then
ttfilename$ = tfolder$ & "\" & tfilename$
tfilenumber% = FreeFile()
Open tfilename$ For Output As #tfilenumber%

' Load output string for keV
astring$ = VbDquote$ & "keV" & VbDquote$ & vbTab

' Load data column labels for k-ratios and binary alpha factors
For l% = 1 To MAXRAY% - 1
For n% = 1 To MAXBINARY%

' Create ZAF output strings
astring$ = astring$ & VbDquote$ & Format$(BinaryRanges!(n%)) & "_ZAF_Krat-" & Xraylo$(l%) & "_(Fanal)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & Format$(BinaryRanges!(n%)) & "_ZAF_Krat-" & Xraylo$(l%) & "_(CalcZAF)" & VbDquote$ & vbTab

' Create ZA only output strings
astring$ = astring$ & VbDquote$ & Format$(BinaryRanges!(n%)) & "_ZA_Krat-" & Xraylo$(l%) & "_(Fanal)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & Format$(BinaryRanges!(n%)) & "_ZA_Krat-" & Xraylo$(l%) & "_(CalcZAF)" & VbDquote$ & vbTab

' Create F only output strings
astring$ = astring$ & VbDquote$ & Format$(BinaryRanges!(n%)) & "_F_Krat-" & Xraylo$(l%) & "_(Fanal)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & Format$(BinaryRanges!(n%)) & "_F_Krat-" & Xraylo$(l%) & "_(CalcZAF)" & VbDquote$ & vbTab
Next n%
Next l%

' Output column labels
Print #tfilenumber%, astring$
Close #tfilenumber%
End If

' Read data for specified beam energy
If mode% = 1 Then
ttfilename$ = tfolder$ & "\" & tfilename$
If Dir$(ttfilename$) = vbNullString Then GoTo Penepma12CalculateReadWriteBinaryDataMatrixFileNotFound
tfilenumber% = FreeFile()
Open ttfilename$ For Input As #tfilenumber%

' Read the column labels
Line Input #tfilenumber%, astring$

' Loop on file until desired voltage is found
For i% = 1 To 50

' Read temp keV
Input #tfilenumber%, tkeV!
If tkeV! <> i% Then GoTo Penepma12CalculateReadWriteBinaryDataMatrixWrongkeV

For l% = 1 To MAXRAY% - 1
For n% = 1 To MAXBINARY%

' Input ZAF values
Input #tfilenumber%, Binary_ZAF_Kratios#(l%, n%)
Input #tfilenumber%, CalcZAF_ZAF_Kratios#(l%, n%)

' Input ZA values
Input #tfilenumber%, Binary_ZA_Kratios#(l%, n%)
Input #tfilenumber%, CalcZAF_ZA_Kratios#(l%, n%)

' Input F values
Input #tfilenumber%, Binary_F_Kratios#(l%, n%)
Input #tfilenumber%, CalcZAF_F_Kratios#(l%, n%)
Next n%
Next l%

' Exit for loop if desired voltage was read
If keV! = tkeV! Then Exit For
Next i%

Close #tfilenumber%
End If

' Write data for specified beam energy (must be written in consecutive keV order)
If mode% = 2 Then
tfilenumber% = FreeFile()
Open tfolder$ & "\" & tfilename$ For Append As #tfilenumber%

' Load output string for keV
astring$ = Format$(keV!) & vbTab

' Output binary (Penepma) and CalcZAF k-ratios
For l% = 1 To MAXRAY% - 1
For n% = 1 To MAXBINARY%

' Create output string for k-ratios (%) and binary alpha factors
astring$ = astring$ & Format$(Binary_ZAF_Kratios#(l%, n%)) & vbTab
astring$ = astring$ & Format$(CalcZAF_ZAF_Kratios#(l%, n%)) & vbTab

' Create ZA only output string
astring$ = astring$ & Format$(Binary_ZA_Kratios#(l%, n%)) & vbTab
astring$ = astring$ & Format$(CalcZAF_ZA_Kratios#(l%, n%)) & vbTab

' Create F only output string
astring$ = astring$ & Format$(Binary_F_Kratios#(l%, n%)) & vbTab
astring$ = astring$ & Format$(CalcZAF_F_Kratios#(l%, n%)) & vbTab
Next n%
Next l%

' Output data for this keV
Print #tfilenumber%, astring$
Close #tfilenumber%
End If

Exit Sub

' Errors
Penepma12CalculateReadWriteBinaryDataMatrixError:
If mode% = 0 Then MsgBox Error$ & ", creating file " & tfilename$, vbOKOnly + vbCritical, "Penepma12CalculateReadWriteBinaryDataMatrix"
If mode% = 1 Then MsgBox Error$ & ", keV(read)= " & Format$(i%) & ", keV(desired)= " & Format$(keV!) & ", ray= " & Format$(l%) & ", n= " & Format$(n%) & ", reading file " & tfilename$, vbOKOnly + vbCritical, "Penepma12CalculateReadWriteBinaryDataMatrix"
If mode% = 2 Then MsgBox Error$ & ", writing file " & tfilename$, vbOKOnly + vbCritical, "Penepma12CalculateReadWriteBinaryDataMatrix"
Close #tfilenumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12CalculateReadWriteBinaryDataMatrixFileNotFound:
msg$ = "The binary composition matrix database file " & tfolder$ & "\" & tfilename$ & " was not found. Unable to perform Penepma k-ratio quantification."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CalculateReadWriteBinaryDataMatrix"
Close #tfilenumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12CalculateReadWriteBinaryDataMatrixWrongkeV:
msg$ = "The keV value (" & Format$(tkeV!) & ") read from file " & tfilename$ & ", is not the expected keV value (" & Format$(i%) & ")."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CalculateReadWriteBinaryDataMatrix"
Close #tfilenumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12CalculateAlphaFactors(l As Integer, tBinaryRanges() As Single, tBinary_Kratios() As Double, tBinary_Factors() As Single, tBinary_Coeffs() As Single)
' Calculate the alpha factors and fit coefficients for the passed k-ratios (matrix alpha factors)
'  l% is the x-ray line being calculated
'  tBinaryRanges!(1 to MAXBINARY%)  are the compositional binaries in weight percent
'  tBinary_Kratios#(1 to MAXRAY%, 1 to MAXBINARY%)  are the k-ratios for each x-ray and binary composition
'  tBinary_Factors!(1 to MAXRAY%, 1 to MAXBINARY%)  are the alpha factors for each x-ray and binary composition, alpha = (C/K - C)/(1 - C)
'  tBinary_Coeffs!(1 to MAXRAY%, 1 to MAXCOEFF4%)  are the polynomial/non-linear alpha factors fit coefficients for each x-ray and MAXBINARY% alpha factors

ierror = False
On Error GoTo Penepma12CalculateAlphaFactorsError

Dim n As Integer, kmax As Integer, nmax As Integer
Dim k As Single, c As Single

' Fit calculations
ReDim xdata(1 To MAXBINARY%) As Single
ReDim ydata(1 To MAXBINARY%) As Single

ReDim acoeff(1 To MAXCOEFF4%) As Single

nmax% = 0
For n% = 1 To MAXBINARY%
If CSng(tBinary_Kratios#(l%, n%)) > 0# And tBinaryRanges!(n%) > 0# Then

' Calculate alpha factor for this binary composition
c! = tBinaryRanges!(n%) / 100#
k! = CSng(tBinary_Kratios#(l%, n%)) / 100#    ' k-ratios are in k-ratio percent
tBinary_Factors!(l%, n%) = ((c! / k!) - c!) / (1 - c!)        ' calculate binary alpha factors

' Increment number of data points
nmax% = nmax% + 1

' Load data arrays for polynomial fit of concentration versus alpha
xdata!(nmax%) = c!
ydata!(nmax%) = tBinary_Factors!(l%, n%)

If DebugMode Then
If nmax% = 1 Then
msg$ = vbCrLf & "Calculating alpha factor fits..."
Call IOWriteLog(msg$)
End If
msg$ = "P=" & Format$(n%) & ", C=" & Format$(c!, f84$) & ", K=" & Format$(k!, f84$) & ", Alpha=" & Format$(ydata!(nmax%), f84$)
Call IOWriteLog(msg$)
End If

' No data to fit, just zero the alpha factor
Else
tBinary_Factors!(l%, n%) = 0#
If VerboseMode Then
msg$ = "Penepma12CalculateAlphaFactors: Zero or negative k-ratio for binary number " & Format$(n%) & " for " & Xraylo$(l%) & " x-ray line"
Call IOWriteLog(msg$)
End If
End If
Next n%

' Calculate alpha factors coefficients if more than two k-ratios are present
If nmax% > 0 Then
kmax% = 2

' Calculate polynomial fit to alpha factors
If CorrectionFlag% < 4 Then
Call LeastSquares(kmax%, nmax%, xdata!(), ydata!(), acoeff!())
If ierror Then Exit Sub

' Calculate non-linear alpha factors
Else
Call LeastMathNonLinear(nmax%, xdata!(), ydata!(), acoeff!())
If ierror Then Exit Sub
End If

tBinary_Coeffs!(l%, 1) = acoeff!(1)
tBinary_Coeffs!(l%, 2) = acoeff!(2)
tBinary_Coeffs!(l%, 3) = acoeff!(3)
tBinary_Coeffs!(l%, 4) = acoeff!(4)

' Display results
If DebugMode Then
Call IOWriteLog("  Alpha1  Alpha2  Alpha3  Alpha4")
msg$ = Format$(Format$(acoeff!(1), f84$), a80$) & Format$(Format$(acoeff!(2), f84$), a80$) & Format$(Format$(acoeff!(3), f84$), a80$) & Format$(Format$(acoeff!(4), f84$), a80$)
Call IOWriteLog(msg$)
End If

' No data to fit, just load default values
Else
tBinary_Coeffs!(l%, 1) = CSng(INT_ONE%)
tBinary_Coeffs!(l%, 2) = CSng(INT_ZERO%)
tBinary_Coeffs!(l%, 3) = CSng(INT_ZERO%)
tBinary_Coeffs!(l%, 4) = CSng(INT_ZERO%)
End If

Exit Sub

' Errors
Penepma12CalculateAlphaFactorsError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CalculateAlphaFactors"
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub PenepmaGetPDATCONFTransition(n As Integer, l As Integer, t1 As Single, t2 As Single)
' Returns the ionization (edge) energies of the specified transition
'  n = atomic number
'  l = PFE x-ray line
'     Ka = 1  = "K L3"   = 1 and 4
'     Kb = 2  = "K M3"   = 1 and 7
'     La = 3  = "L3 M5"  = 4 and 9
'     Lb = 4  = "L2 M4"  = 3 and 8
'     Ma = 5  = "M5 N7"  = 9 and 16
'     Mb = 6  = "M4 N6"  = 8 and 15
'  t1 = first transition edge
'  t2 = second transition edge

ierror = False
On Error GoTo PenepmaGetPDATCONFTransitionError

Dim tfilename As String
Dim astring As String, bstring As String

Dim i As Integer, i1 As Integer, i2 As Integer
Dim atnum As Integer, ishell As Integer
Dim ashell As String, sshell As String
Dim oshell As Integer
Dim ienergy As Single, cprofile As Single

' Init return variables
t1! = 0#
t2! = 0#

' Load test variables
If n% > 6 Then
If l% = 1 Then i1% = 1: i2% = 4
If l% = 2 Then i1% = 1: i2% = 7
If l% = 3 Then i1% = 4: i2% = 9
If l% = 4 Then i1% = 3: i2% = 8
If l% = 5 Then i1% = 9: i2% = 16
If l% = 6 Then i1% = 8: i2% = 15

' Special sub shells for carbon, boron, beryllium
Else
If n% = 4 And l% = 1 Then i1% = 1: i2% = 2  ' Be (K L1)
If n% = 5 And l% = 1 Then i1% = 1: i2% = 3  ' B  (K L2)
If n% = 6 And l% = 1 Then i1% = 1: i2% = 3  ' C  (K L2)
End If

' Open file
Close #Temp1FileNumber%
tfilename$ = PENDBASE_Path$ & "\pdfiles\pdatconf.pen"
If Dir$(tfilename$) = vbNullString Then GoTo PenepmaGetPDATCONFTransitionFileNotFound
Open tfilename$ For Input As #Temp1FileNumber%

' Read text lines (1 to 22)
For i% = 1 To 22
Line Input #Temp1FileNumber%, astring$
Next i%

Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$

' Parse out atomic number
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
atnum% = Val(bstring$)
If atnum% < 1 Or atnum% > MAXELM% Then GoTo PenepmaGetPDATCONFTransitionNotValidElement

' Check for atomic number
If atnum% = n% Then

' Parse out shell number
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
ishell% = Val(bstring$)

' Parse out shell string (K1, L3, etc)
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
ashell$ = bstring$

' Parse out spectroscopic notation (1s1/2, 2p3/2, etc)
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
sshell$ = bstring$

' Parse out shell occupation number
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
oshell% = Val(bstring$)

' Parse out ionization energy
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
ienergy! = Val(bstring$)

' Parse out compton profile
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
cprofile! = Val(bstring$)

' Check for transition numbers and load ionization energies if matched
If ishell% = i1% Then t1! = ienergy! / EVPERKEV#
If ishell% = i2% Then t2! = ienergy! / EVPERKEV#

End If
Loop

Close #Temp1FileNumber%
Exit Sub

' Errors
PenepmaGetPDATCONFTransitionError:
MsgBox Error$, vbOKOnly + vbCritical, "PenepmaGetPDATCONFTransition"
Close #Temp1FileNumber%
ierror = True
Exit Sub

PenepmaGetPDATCONFTransitionFileNotFound:
msg$ = "File " & tfilename$ & " was not found."
MsgBox msg$, vbOKOnly + vbExclamation, "PenepmaGetPDATCONFTransition"
ierror = True
Exit Sub

PenepmaGetPDATCONFTransitionNotValidElement:
msg$ = "Atomic number " & Format$(atnum%) & " is not a valid atomic number"
MsgBox msg$, vbOKOnly + vbExclamation, "PenepmaGetPDATCONFTransition"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma12MatrixReadMDB2(tTakeoff As Single, tKilovolt As Single, tEmitter As Integer, tXray As Integer, tMatrix As Integer, tKratios() As Double, notfound As Boolean)
' This routine reads the Matrix.mdb file for the specified beam energy, emitter, x-ray, matrix.
'  tKratios#(1 to MAXBINARY%)  are the k-ratios for this x-ray and binary composition

ierror = False
On Error GoTo Penepma12MatrixReadMDB2Error

Dim i As Integer
Dim nrec As Long

Dim SQLQ As String
Dim MtDb As Database
Dim MtDs As Recordset

' Check for file
If Dir$(MatrixMDBFile$) = vbNullString Then GoTo Penepma12MatrixReadMDBNoMatrixMDB2File

' Check for use keV rounding flag for fractional keVs
If Penepma12UseKeVRoundingFlag Then
tKilovolt! = Int(tKilovolt! + 0.5)
End If

' Open matrix database (non exclusive and read only)
Screen.MousePointer = vbHourglass
Set MtDb = OpenDatabase(MatrixMDBFile$, MatrixDatabaseNonExclusiveAccess%, dbReadOnly)

' Try to find requested emitter, matrix, etc
SQLQ$ = "SELECT Matrix.BeamTakeOff, Matrix.BeamEnergy, Matrix.EmittingElement, Matrix.EmittingXray, Matrix.MatrixElement, Matrix.MatrixNumber FROM Matrix WHERE"
SQLQ$ = SQLQ$ & " BeamTakeOff = " & Format$(tTakeoff!) & " AND BeamEnergy = " & Format$(tKilovolt!) & " AND"
SQLQ$ = SQLQ$ & " EmittingElement = " & Format$(tEmitter%) & " AND EmittingXray = " & Format$(tXray%) & " AND"
SQLQ$ = SQLQ$ & " MatrixElement = " & Format$(tMatrix%)
Set MtDs = MtDb.OpenRecordset(SQLQ$, dbOpenSnapshot)

' If record not found, return notfound
If MtDs.BOF And MtDs.EOF Then
notfound = True
Screen.MousePointer = vbDefault
Exit Sub
End If

' Load return values based on "MatrixNumber"
nrec& = MtDs("MatrixNumber")
MtDs.Close

' Search for records
SQLQ$ = "SELECT MatrixKRatio.* FROM MatrixKRatio WHERE MatrixKRatioNumber = " & Format$(nrec&)
Set MtDs = MtDb.OpenRecordset(SQLQ$, dbOpenSnapshot)
If MtDs.BOF And MtDs.EOF Then GoTo Penepma12MatrixReadMDB2NoKRatios

' Load kratio array (only!)
Do Until MtDs.EOF
i% = MtDs("MatrixKRatioOrder")          ' load order (1 to MAXBINARY%)
tKratios#(i%) = CDbl(MtDs("MatrixKRatio_ZAF_KRatio"))
MtDs.MoveNext
Loop
MtDs.Close

notfound = False
MtDb.Close

Screen.MousePointer = vbDefault
Exit Sub

' Errors
Penepma12MatrixReadMDB2Error:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12MatrixReadMDB2"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12MatrixReadMDBNoMatrixMDB2File:
msg$ = "File " & MatrixMDBFile$ & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12MatrixReadMDB2"
ierror = True
Exit Sub

Penepma12MatrixReadMDB2NoKRatios:
msg$ = "File " & MatrixMDBFile$ & " did not contain any k-ratio records for " & Format$(tTakeoff!) & " degrees, " & Format$(tKilovolt!) & " keV, " & Symup$(tEmitter%) & " " & Xraylo$(tXray%) & " in " & Symup$(tMatrix%)
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12MatrixReadMDB2"
ierror = True
Exit Sub

End Sub

Sub Penepma12CalculateReadWritePureElement(mode As Integer, tfolder As String, tfilename As String, keV As Single)
' Reads or write the binary fluorescence matrix k-ratio data to or from a data file for a specified beam energy
'  mode = 0 create file and write column labels only
'  mode = 1 read data
'  mode = 2 write data
'  tfolder$ is the full path of the pure element data file to read or write
'  tfilename$ is the filename of the pure element data file to read or write
'  keV is the specified beam energy

ierror = False
On Error GoTo Penepma12CalculateReadWritePureElementError

Dim tfilenumber As Integer
Dim l As Integer, i As Integer
Dim tkeV As Single
Dim astring As String, ttfilename As String

' Write column labels only
If mode% = 0 Then
ttfilename$ = tfolder$ & "\" & tfilename$
tfilenumber% = FreeFile()
Open ttfilename$ For Output As #tfilenumber%

' Load output string for keV
astring$ = VbDquote$ & "keV" & VbDquote$ & vbTab

' Load data column labels for generated and emitted pure element intensities
For l% = 1 To MAXRAY% - 1
astring$ = astring$ & VbDquote$ & "Gene_" & Xraylo$(l%) & "_(Fanal)" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "Emit_" & Xraylo$(l%) & "_(Fanal)" & VbDquote$ & vbTab
Next l%

' Output column labels
Print #tfilenumber%, astring$
Close #tfilenumber%
End If

' Read data for specified beam energy
If mode% = 1 Then
ttfilename$ = tfolder$ & "\" & tfilename$
If Dir$(ttfilename$) = vbNullString Then GoTo Penepma12CalculateReadWritePureElementFileNotFound
tfilenumber% = FreeFile()
Open ttfilename$ For Input As #tfilenumber%

' Read the column labels
Line Input #tfilenumber%, astring$

' Loop on file until desired voltage is found
For i% = 1 To 50

' Read temp keV
Input #tfilenumber%, tkeV!
If tkeV! <> i% Then GoTo Penepma12CalculateReadWritePureElementWrongkeV

For l% = 1 To MAXRAY% - 1
Input #tfilenumber%, PureGenerated_Intensities#(l%)
Input #tfilenumber%, PureEmitted_Intensities#(l%)
Next l%

' Exit for loop if desired voltage was read
If keV! = tkeV! Then Exit For
Next i%

Close #tfilenumber%
End If

' Write data for specified beam energy (must be written in consecutive keV order)
If mode% = 2 Then
tfilenumber% = FreeFile()
Open tfolder$ & "\" & tfilename$ For Append As #tfilenumber%

' Load output string for keV
astring$ = Format$(keV!) & vbTab

' Create output string for generated and pure element intensities
For l% = 1 To MAXRAY% - 1
astring$ = astring$ & Format$(PureGenerated_Intensities#(l%)) & vbTab
astring$ = astring$ & Format$(PureEmitted_Intensities#(l%)) & vbTab
Next l%

' Output data for this keV
Print #tfilenumber%, astring$
Close #tfilenumber%
End If

Exit Sub

' Errors
Penepma12CalculateReadWritePureElementError:
If mode% = 0 Then MsgBox Error$ & ", creating file " & tfilename$, vbOKOnly + vbCritical, "Penepma12CalculateReadWritePureElement"
If mode% = 1 Then MsgBox Error$ & ", keV(read)= " & Format$(i%) & ", keV(desired)= " & Format$(keV!) & ", ray= " & Format$(l%) & ", reading file " & tfilename$, vbOKOnly + vbCritical, "Penepma12CalculateReadWritePureElement"
If mode% = 2 Then MsgBox Error$ & ", writing file " & tfilename$, vbOKOnly + vbCritical, "Penepma12CalculateReadWritePureElement"
Close #tfilenumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12CalculateReadWritePureElementFileNotFound:
msg$ = "The binary composition matrix database file " & tfolder$ & "\" & tfilename$ & " was not found. Unable to perform Penepma k-ratio quantification."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CalculateReadWritePureElement"
Close #tfilenumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12CalculateReadWritePureElementWrongkeV:
msg$ = "The keV value (" & Format$(tkeV!) & ") read from file " & tfilename$ & ", is not the expected keV value (" & Format$(i%) & ")."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CalculateReadWritePureElement"
Close #tfilenumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12PureReadMDB2(tTakeoff As Single, tKilovolt As Single, tEmitter As Integer, tXray As Integer, tIntensityGenerated As Double, tIntensityEmitted As Double, notfound As Boolean)
' This routine reads the Pure.mdb file for the specified beam energy, emitter and x-ray and returns the generated and emitted pure element intensities.

ierror = False
On Error GoTo Penepma12PureReadMDB2Error

Dim i As Integer
Dim nrec As Long

Dim SQLQ As String
Dim MtDb As Database
Dim MtDs As Recordset

' Check for file
If Dir$(PureMDBFile$) = vbNullString Then GoTo Penepma12PureReadMDBNoPureMDB2File

' Check for use keV rounding flag for fractional keVs
If Penepma12UseKeVRoundingFlag Then
tKilovolt! = Int(tKilovolt! + 0.5)
End If

' Open Pure database (non exclusive and read only)
Screen.MousePointer = vbHourglass
Set MtDb = OpenDatabase(PureMDBFile$, PureDatabaseNonExclusiveAccess%, dbReadOnly)

' Try to find requested emitter, etc
SQLQ$ = "SELECT Pure.BeamTakeOff, Pure.BeamEnergy, Pure.EmittingElement, Pure.EmittingXray, Pure.PureNumber FROM Pure WHERE"
SQLQ$ = SQLQ$ & " BeamTakeOff = " & Format$(tTakeoff!) & " AND BeamEnergy = " & Format$(tKilovolt!) & " AND"
SQLQ$ = SQLQ$ & " EmittingElement = " & Format$(tEmitter%) & " AND EmittingXray = " & Format$(tXray%)
Set MtDs = MtDb.OpenRecordset(SQLQ$, dbOpenSnapshot)

' If record not found, return notfound
If MtDs.BOF And MtDs.EOF Then
notfound = True
Screen.MousePointer = vbDefault
Exit Sub
End If

' Load return values based on "PureNumber"
nrec& = MtDs("PureNumber")
MtDs.Close

' Search for records
SQLQ$ = "SELECT PureIntensity.PureIntensityGenerated, PureIntensity.PureIntensityEmitted FROM PureIntensity WHERE PureIntensityNumber = " & Format$(nrec&)
Set MtDs = MtDb.OpenRecordset(SQLQ$, dbOpenSnapshot)
If MtDs.BOF And MtDs.EOF Then GoTo Penepma12PureReadMDB2NoIntensities

' Load pure element intensities
tIntensityGenerated# = CDbl(MtDs("PureIntensityGenerated"))
tIntensityEmitted# = CDbl(MtDs("PureIntensityEmitted"))
MtDs.Close

notfound = False
MtDb.Close

Screen.MousePointer = vbDefault
Exit Sub

' Errors
Penepma12PureReadMDB2Error:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12PureReadMDB2"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

Penepma12PureReadMDBNoPureMDB2File:
msg$ = "File " & PureMDBFile$ & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12PureReadMDB2"
ierror = True
Exit Sub

Penepma12PureReadMDB2NoIntensities:
msg$ = "File " & PureMDBFile$ & " did not contain any pure element intensity records for " & Format$(tTakeoff!) & " degrees, " & Format$(tKilovolt!) & " keV, " & Symup$(tEmitter%) & " " & Xraylo$(tXray%)
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12PureReadMDB2"
ierror = True
Exit Sub

End Sub

