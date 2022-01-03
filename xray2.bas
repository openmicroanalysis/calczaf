Attribute VB_Name = "CodeXRAY2"
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

Function XrayGetEdgeTable(mode As Integer) As String
' Subroutine to return a string of a table of edge energies or angstroms
'  mode = 1 return energy
'  mode = 2 return angstroms

ierror = False
On Error GoTo XrayGetEdgeTableError

Dim ielm As Integer, iedg As Integer
Dim nrec As Integer
Dim lam As Single, temp As Single
Dim tmsg As String

Dim edgrow As TypeEdge

' Determine the element number
tmsg$ = vbCrLf
If mode% = 1 Then tmsg$ = tmsg$ & "Table of Edge Energies (KeV)"
If mode% = 2 Then tmsg$ = tmsg$ & "Table of Edge Angstroms"
tmsg$ = tmsg$ & vbCrLf
tmsg$ = tmsg$ & Format$("Element", a80$) & " "
For iedg% = 1 To MAXEDG%
tmsg$ = tmsg$ & Format$(Edglo$(iedg%), a80$) & " "
Next iedg%
tmsg$ = tmsg$ & vbCrLf

For ielm% = 1 To MAXELM%

' Read x-ray edge file
nrec% = ielm% + 2
Open XEdgeFile$ For Random Access Read As #XEdgeFileNumber% Len = XRAY_FILE_RECORD_LENGTH%
Get #XEdgeFileNumber%, nrec%, edgrow
Close #XEdgeFileNumber%

tmsg$ = tmsg$ & Format$(Symup$(ielm%), a80$) & " "

' Loop on all xrays
For iedg% = 1 To MAXEDG%

' Check for non-zero entry
If edgrow.energy!(iedg%) <> 0# Then

' Load energy (in keV)
If mode% = 1 Then
temp! = edgrow.energy!(iedg%) / EVPERKEV#
tmsg$ = tmsg$ & Format$(Format$(temp!, f85$), a80$) & " "

' Calculate angstroms and load
Else
lam! = ANGEV! / edgrow.energy!(iedg%)
tmsg$ = tmsg$ & Format$(Format$(lam!, f84$), a80$) & " "
End If

' Load space if zero
Else
tmsg$ = tmsg$ & Space$(8) & " "
End If

Next iedg%
tmsg$ = tmsg$ & vbCrLf
Next ielm%

XrayGetEdgeTable$ = tmsg$

Exit Function

' Errors
XrayGetEdgeTableError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "XrayGetEdgeTable"
ierror = True
Exit Function

End Function

Function XrayGetEmissionTable(mode As Integer, method As Integer) As String
' Subroutine to return a string of a table of emission energies or angstroms
'  mode = 1 original x-ray lines
'  mode = 2 additional x-ray lines
'  method = 1 return energy
'  method = 2 return angstroms

ierror = False
On Error GoTo XrayGetEmissionTableError

Dim ielm As Integer, iray As Integer
Dim nrec As Integer
Dim lam As Single, temp As Single
Dim tmsg As String

Dim engrow As TypeEnergy

' Check for new MAXRAY constant
If mode% = 2 And MAXRAY% - 1 = MAXRAY_OLD% Then Exit Function

' Determine the element number
tmsg$ = vbCrLf
If method% = 1 Then tmsg$ = tmsg$ & "Table of Emission Energies (KeV)"
If method% = 2 Then tmsg$ = tmsg$ & "Table of Emission Angstroms"
tmsg$ = tmsg$ & vbCrLf
tmsg$ = tmsg$ & Format$("Element", a80$) & " "

' Load column labels
For iray% = 1 To MAXRAY_OLD%

' original x-rays
If mode% = 1 Then
tmsg$ = tmsg$ & Format$(Xraylo$(iray%), a80$) & " "

' Additional x-rays
Else
tmsg$ = tmsg$ & Format$(Xraylo$(iray% + MAXRAY_OLD%), a80$) & " "
End If
Next iray%
tmsg$ = tmsg$ & vbCrLf

' All elements
For ielm% = 1 To MAXELM%

' Load element row record
nrec% = ielm% + 2

' Original x-ray lines
If mode% = 1 Then
Open XLineFile$ For Random Access Read As #XLineFileNumber% Len = XRAY_FILE_RECORD_LENGTH%
Get #XLineFileNumber%, nrec%, engrow
Close #XLineFileNumber%

' Additional x-ray lines
Else
If Dir$(XLineFile2$) = vbNullString Then GoTo XrayGetEmissionTableNotFoundXLINE2DAT
If FileLen(XLineFile2$) = 0 Then GoTo XrayGetEmissionTableZeroSizeXLINE2DAT
Open XLineFile2$ For Random Access Read As #XLineFileNumber2% Len = XRAY_FILE_RECORD_LENGTH%
Get #XLineFileNumber2%, nrec%, engrow
Close #XLineFileNumber2%
End If

tmsg$ = tmsg$ & Format$(Symup$(ielm%), a80$) & " "

' Loop on all xrays
For iray% = 1 To MAXRAY_OLD%

' Check for non-zero entry
If Not MiscDifferenceIsSmall(engrow.energy!(iray%), 0#, 0.000001) Then

' Load energy (in keV)
If method% = 1 Then
temp! = engrow.energy!(iray%) / EVPERKEV#
tmsg$ = tmsg$ & Format$(Format$(temp!, f85$), a80$) & " "

' Calculate angstroms and load
Else
lam! = ANGEV! / engrow.energy!(iray%)
tmsg$ = tmsg$ & Format$(Format$(lam!, f84$), a80$) & " "
End If

' Load space if zero
Else
tmsg$ = tmsg$ & Space$(8) & " "
End If

Next iray%
tmsg$ = tmsg$ & vbCrLf
Next ielm%

XrayGetEmissionTable$ = tmsg$

Exit Function

' Errors
XrayGetEmissionTableError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "XrayGetEmissionTable"
Close #XLineFileNumber%
Close #XLineFileNumber2%
ierror = True
Exit Function

XrayGetEmissionTableNotFoundXLINE2DAT:
msg$ = "The " & XLineFile2$ & " was not found." & vbCrLf & vbCrLf
msg$ = msg$ & "Please run the latest CalcZAF.msi installer to obtain this additional x-ray line file."
MsgBox msg$, vbOKOnly + vbExclamation, "XrayGetEmissionTable"
Close #XLineFileNumber%
Close #XLineFileNumber2%
ierror = True
Exit Function

XrayGetEmissionTableZeroSizeXLINE2DAT:
Kill XLineFile2$
msg$ = "The " & XLineFile2$ & " was not found." & vbCrLf & vbCrLf
msg$ = msg$ & "Please run the latest CalcZAF.msi installer to obtain this additional x-ray line file."
MsgBox msg$, vbOKOnly + vbExclamation, "XrayGetEmissionTable"
Close #XLineFileNumber%
Close #XLineFileNumber2%
ierror = True
Exit Function

End Function

Function XrayGetFluorescentYieldTable(mode As Integer) As String
' Subroutine to return a string of a table of fluorescent yields
'  mode = 1 original x-ray lines
'  mode = 2 additional x-ray lines

ierror = False
On Error GoTo XrayGetFluorescentYieldTableError

Dim ielm As Integer, iray As Integer
Dim nrec As Integer
Dim tmsg As String

Dim flurow As TypeFlur

' Check for new MAXRAY constant
If mode% = 2 And MAXRAY% - 1 = MAXRAY_OLD% Then Exit Function

' Determine the element number
tmsg$ = vbCrLf
tmsg$ = tmsg$ & "Table of Fluorescent Yields"
tmsg$ = tmsg$ & vbCrLf
tmsg$ = tmsg$ & Format$("Element", a80$) & " "

' Load column labels
For iray% = 1 To MAXRAY_OLD%

' Original x-rays
If mode% = 1 Then
tmsg$ = tmsg$ & Format$(Xraylo$(iray%), a80$) & " "

' Additional x-rays
Else
tmsg$ = tmsg$ & Format$(Xraylo$(iray% + MAXRAY_OLD%), a80$) & " "
End If
Next iray%
tmsg$ = tmsg$ & vbCrLf

For ielm% = 1 To MAXELM%

' Set element record number
nrec% = ielm% + 2

' Load original x-ray lines
If mode% = 1 Then
Open XFlurFile$ For Random Access Read As #XFlurFileNumber% Len = XRAY_FILE_RECORD_LENGTH%
Get #XFlurFileNumber%, nrec%, flurow
Close #XFlurFileNumber%

' Load additional x-ray lines
Else
If Dir$(XFlurFile2$) = vbNullString Then GoTo XrayGetFluorescentTableNotFoundXFLUR2DAT
If FileLen(XFlurFile2$) = 0 Then GoTo XrayGetFluorescentTableZeroSizeXFLUR2DAT
Open XFlurFile2$ For Random Access Read As #XFlurFileNumber2% Len = XRAY_FILE_RECORD_LENGTH%
Get #XFlurFileNumber2%, nrec%, flurow
Close #XFlurFileNumber2%
End If

tmsg$ = tmsg$ & Format$(Symup$(ielm%), a80$) & " "

' Loop on all xrays
For iray% = 1 To MAXRAY_OLD%

' Check for non-zero entry
If flurow.fraction!(iray%) <> 0# Then
tmsg$ = tmsg$ & Format$(Format$(flurow.fraction!(iray%), f84$), a80$) & " "

' Load space if zero
Else
tmsg$ = tmsg$ & Space$(8) & " "
End If

Next iray%
tmsg$ = tmsg$ & vbCrLf
Next ielm%

XrayGetFluorescentYieldTable$ = tmsg$

Exit Function

' Errors
XrayGetFluorescentYieldTableError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "XrayGetFluorescentYieldTable"
Close #XFlurFileNumber%
Close #XFlurFileNumber2%
ierror = True
Exit Function

XrayGetFluorescentTableNotFoundXFLUR2DAT:
msg$ = "The " & XFlurFile2$ & " was not found." & vbCrLf & vbCrLf
msg$ = msg$ & "Please run the latest CalcZAF.msi installer to obtain this additional x-ray line file."
MsgBox msg$, vbOKOnly + vbExclamation, "XrayGetFluorescentTable"
Close #XFlurFileNumber%
Close #XFlurFileNumber2%
ierror = True
Exit Function

XrayGetFluorescentTableZeroSizeXFLUR2DAT:
Kill XFlurFile2$
msg$ = "The " & XFlurFile2$ & " was not found." & vbCrLf & vbCrLf
msg$ = msg$ & "Please run the latest CalcZAF.msi installer to obtain this additional x-ray line file."
MsgBox msg$, vbOKOnly + vbExclamation, "XrayGetFluorescentTable"
Close #XFlurFileNumber%
Close #XFlurFileNumber2%
ierror = True
Exit Function

End Function

Function XrayGetMACTable(mode As Integer, sym As String) As String
' Subroutine to return a string of a table of MACs
'  mode = 1 original x-ray lines
'  mode = 2 additional x-ray lines

ierror = False
On Error GoTo XrayGetMACTableError

Dim ielm As Integer, iray As Integer, ip As Integer
Dim nrec As Integer, num As Integer
Dim tmsg As String

Dim macrow As TypeMu

' Check for new MAXRAY constant
If mode% = 2 And MAXRAY% - 1 = MAXRAY_OLD% Then Exit Function

' Check for additional lines. If found, check that FFAST2.DAT, PENEPMAMAC or POUCHOUMAC files exist.
If mode% = 2 Then
If MACTypeFlag% < 6 Then GoTo XrayGetMACTableNewMACFileNotSpecified
MACFile$ = ApplicationCommonAppData$ & macstring2$(MACTypeFlag%) & "2.DAT"
If Dir$(MACFile$) = vbNullString Then GoTo XrayGetMACTableFFAST2DATNotFound
End If

If DebugMode Then Call IOWriteLog("Now loading MACs...")

If mode% = 1 Then
MACFile$ = ApplicationCommonAppData$ & macstring2$(MACTypeFlag%) & ".DAT"
Else
MACFile$ = ApplicationCommonAppData$ & macstring2$(MACTypeFlag%) & "2.DAT"
End If

Open MACFile$ For Random Access Read As #MACFileNumber% Len = MAC_FILE_RECORD_LENGTH%

' Determine the element number
ip% = IPOS1(MAXELM%, sym$, Symlo$())
If ip% = 0 Then GoTo XrayGetMACTableBadElement

' Read MAC file element emitter record
nrec% = AllAtomicNums%(ip%)
Get #MACFileNumber%, nrec%, macrow
Close #MACFileNumber%

tmsg$ = vbCrLf
tmsg$ = tmsg$ & VbDquote$ & "Table of MACs (mass absorption coefficients) from " & MACFile$ & VbDquote$ & vbCrLf
tmsg$ = tmsg$ & VbDquote$ & "Emitting Element: " & Symup$(ip%) & VbDquote$ & vbCrLf
tmsg$ = tmsg$ & Format$("Absorber", a80$) & vbTab

' Load column labels
For iray% = 1 To MAXRAY_OLD%

' Original x-rays
If mode% = 1 Then
tmsg$ = tmsg$ & Format$(Xraylo$(iray%), a80$) & vbTab

' Additional x-rays
Else
tmsg$ = tmsg$ & Format$(Xraylo$(iray% + MAXRAY_OLD%), a80$) & vbTab
End If
Next iray%
tmsg$ = tmsg$ & vbCrLf

For ielm% = 1 To MAXELM%
tmsg$ = tmsg$ & Format$(Symup$(ielm%), a80$) & vbTab

' Loop on all xrays
For iray% = 1 To MAXRAY_OLD%
num% = iray% + (ielm% - 1) * (MAXRAY_OLD%)
If macrow.mac!(num%) > 0# Then
tmsg$ = tmsg$ & Format$(Format$(macrow.mac!(num%), f82$), a80$) & vbTab

' Load space if zero
Else
tmsg$ = tmsg$ & Space$(8) & vbTab
End If

Next iray%
tmsg$ = tmsg$ & vbCrLf
Next ielm%

Close #MACFileNumber%
XrayGetMACTable$ = tmsg$

Exit Function

' Errors
XrayGetMACTableError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "XrayGetMACTable"
Close #MACFileNumber%
ierror = True
Exit Function

XrayGetMACTableBadElement:
Screen.MousePointer = vbDefault
msg$ = "Invalid element symbol in MAC table"
MsgBox msg$, vbOKOnly + vbExclamation, "XrayGetMACTable"
Close #MACFileNumber%
ierror = True
Exit Function

XrayGetMACTableNewMACFileNotSpecified:
msg$ = "For quantification of additional x-ray lines, you must specify the FFAST, PENEPMAMAC or POUCHOUMAC MAC databases. See the Analytical | ZAF, Phi-Rho-Z, Alpha Factor and Calibration Curve Selections menu and select the MACs button."
MsgBox msg$, vbOKOnly + vbExclamation, "XrayGetMACTable"
ierror = True
Exit Function

XrayGetMACTableFFAST2DATNotFound:
msg$ = "File " & MACFile$ & " was not found. You will need to re-run the CalcZAF.msi installer to obtain the new x-ray tables for additional x-ray lines."
MsgBox msg$, vbOKOnly + vbExclamation, "XrayGetMACTable"
ierror = True
Exit Function

End Function

Function XrayGetSpectrometerTable() As String
' Subroutine to return a string of a table of spectrometer positions (Ka, La, Ma, Kb, Lb, Mb only)

ierror = False
On Error GoTo XrayGetSpectrometerTableError

Dim ielm As Integer, iray As Integer
Dim nrec As Integer, ip As Integer, n As Integer
Dim i As Integer, j As Integer
Dim lam As Single, temp As Single
Dim x2d As Single, k As Single
Dim esym As String, xsym As String
Dim tmsg As String

Dim engrow As TypeEnergy

Dim xnames(1 To MAXCRYSTYPE%) As String, txnames(1 To MAXCRYSTYPE%) As String
Dim x2ds(1 To MAXCRYSTYPE%) As Single, tx2ds(1 To MAXCRYSTYPE%) As Single
Dim xks(1 To MAXCRYSTYPE%) As Single
Dim astring(1 To MAXCRYSTYPE%) As String

' Determine the crystals in this probe setup
n% = 0
For i% = 1 To NumberOfTunableSpecs%
For j% = 1 To ScalNumberOfCrystals(i%)
ip% = IPOS1(n%, ScalCrystalNames$(j%, i%), txnames$())

' If new crystal, load
If ip% = 0 Then
n% = n% + 1
If n% <= MAXCRYSTYPE% Then
txnames$(n%) = ScalCrystalNames$(j%, i%)

' Get the 2d for each crystal to sort by below
Call MiscGetCrystalParameters(txnames$(n%), x2d!, k!, esym$, xsym$)
If ierror Then Exit Function
tx2ds!(n%) = x2d!
End If
End If

Next j%
Next i%

' Sort the crystal names based on 2d
Call MiscSortStringArray(Int(1), n%, txnames$(), xnames$(), tx2ds!(), x2ds!())
If ierror Then Exit Function

' Get the 2d and k for each sorted crystal for position calculation below
For i% = 1 To n%
Call MiscGetCrystalParameters(xnames$(i%), x2d!, k!, esym$, xsym$)
If ierror Then Exit Function

x2ds!(i%) = x2d!
xks!(i%) = k!
Next i%

' Line title
tmsg$ = vbCrLf
tmsg$ = tmsg$ & "Table of Emission Line Spectrometer Positions"
tmsg$ = tmsg$ & vbCrLf
tmsg$ = tmsg$ & Space$(8) & " "
tmsg$ = tmsg$ & Format$("Alpha", a80$) & " " & Space$((n% - 1) * 9)
tmsg$ = tmsg$ & "  "
tmsg$ = tmsg$ & Format$("Beta", a80$) & " " & Space$((n% - 1) * 9)

' Load column labels
tmsg$ = tmsg$ & vbCrLf
tmsg$ = tmsg$ & Format$("Element", a80$) & " "

' Alpha series
For i% = 1 To n%
tmsg$ = tmsg$ & Format$(xnames$(i%), a80$) & " "
Next i%

' Divider
tmsg$ = tmsg$ & "  "

' Beta series
For i% = 1 To n%
tmsg$ = tmsg$ & Format$(xnames$(i%), a80$) & " "
Next i%
tmsg$ = tmsg$ & vbCrLf

' Determine the element number
For ielm% = 1 To MAXELM%
tmsg$ = tmsg$ & Format$(Symup$(ielm%), a80$) & " "

' Read x-ray line file for this line
nrec% = ielm% + 2
Open XLineFile$ For Random Access Read As #XLineFileNumber% Len = XRAY_FILE_RECORD_LENGTH%
Get #XLineFileNumber%, nrec%, engrow
Close #XLineFileNumber%

' Loop on each crystal type
For i% = 1 To n%
astring$(i%) = Space$(8) & " "

' Loop on alpha xrays
For iray% = 1 To MAXRAY_OLD% - 1 Step 2

' Load xray symbol and get energy and angstroms
If engrow.energy(iray%) > 0 Then

' Calculate angstroms
lam! = ANGEV! / engrow.energy!(iray%)

' Convert angstroms to spectrometer position
temp! = XrayCalculatePositions!(Int(0), Int(1), Int(1), x2ds(i%), xks(i%), lam!)
If ierror Then Exit Function

' If inbounds, load string
If MiscMotorInBounds(Int(1), temp!) Then
astring$(i%) = MiscAutoFormat$(temp!) & " "
End If
End If

Next iray%

' Add string
tmsg$ = tmsg$ & astring$(i%)
Next i%

' Divider
tmsg$ = tmsg$ & "  "

' Loop on each crystal type
For i% = 1 To n%
astring$(i%) = Space$(8) & " "

' Loop on beta xrays
For iray% = 2 To (MAXRAY_OLD%) Step 2

' Load xray symbol and get energy and angstroms
If engrow.energy(iray%) > 0 Then

' Calculate angstroms
lam! = ANGEV! / engrow.energy!(iray%)

' Convert angstroms to spectrometer position
temp! = XrayCalculatePositions!(Int(0), Int(1), Int(1), x2ds(i%), xks(i%), lam!)
If ierror Then Exit Function

' If inbounds, load string
If MiscMotorInBounds(Int(1), temp!) Then
astring$(i%) = MiscAutoFormat$(temp!) & " "
End If
End If

Next iray%

' Add string
tmsg$ = tmsg$ & astring$(i%)
Next i%

tmsg$ = tmsg$ & vbCrLf
Next ielm%

XrayGetSpectrometerTable$ = tmsg$

Exit Function

' Errors
XrayGetSpectrometerTableError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "XrayGetSpectrometerTable"
Close #XLineFileNumber%
ierror = True
Exit Function

End Function

Sub XrayGetTable(mode As Integer)
' Display xray table
'  mode = 1 display emission table
'  mode = 2 display edge table
'  mode = 3 display fluorescent yield table
'
'  mode = 4 display emission table (additional x-ray lines)
'  mode = 5 not used
'  mode = 6 display fluorescent yield table (additional x-ray lines)
'
'  mode = 7  display MAC table
'  mode = 8  display MAC table (complete)
'  mode = 9  display MAC table (additional x-ray lines)
'  mode = 10 display MAC table (complete) (additional x-ray lines)

ierror = False
On Error GoTo XrayGetTableError

Dim response As Integer, i As Integer
Dim tfilename As String

Static sym As String

If mode% = 1 Or mode% = 2 Or mode% = 4 Then
msg$ = "Display the x-ray table values in angstrom units?"
response% = MsgBox(msg$, vbYesNoCancel + vbQuestion + vbDefaultButton2, "XrayGetTable")
If response% = vbCancel Then Exit Sub

' Angstrom
If response% = vbYes Then
If mode% = 1 Then
Screen.MousePointer = vbHourglass
msg$ = XrayGetEmissionTable(Int(1), Int(2))
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End If

If mode% = 2 Then
Screen.MousePointer = vbHourglass
msg$ = XrayGetEdgeTable(Int(2))
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End If

If mode% = 4 Then
Screen.MousePointer = vbHourglass
msg$ = XrayGetEmissionTable(Int(2), Int(2))
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End If

' Energy
Else
If mode% = 1 Then
Screen.MousePointer = vbHourglass
msg$ = XrayGetEmissionTable(Int(1), Int(1))
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End If

If mode% = 2 Then
Screen.MousePointer = vbHourglass
msg$ = XrayGetEdgeTable(Int(1))
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End If

If mode% = 4 Then
Screen.MousePointer = vbHourglass
msg$ = XrayGetEmissionTable(Int(2), Int(1))
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End If
End If
End If

' Fluorescent yield
If mode% = 3 Then
Screen.MousePointer = vbHourglass
msg$ = XrayGetFluorescentYieldTable(Int(1))
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End If

If mode% = 6 Then
Screen.MousePointer = vbHourglass
msg$ = XrayGetFluorescentYieldTable(Int(2))
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End If

' MAC table
If mode% = 7 Or mode% = 9 Then

' MAC element emitter
msg$ = "Enter the emitter element symbol for MAC table"
If sym$ = vbNullString Then sym$ = Symup$(ATOMIC_NUM_OXYGEN%)
sym$ = InputBox$(msg$, "XrayGetTable", sym$)
If sym$ = vbNullString Then Exit Sub

If mode% = 7 Then MACFile$ = ApplicationCommonAppData$ & macstring2$(MACTypeFlag%) & ".DAT"
If mode% = 9 Then MACFile$ = ApplicationCommonAppData$ & macstring2$(MACTypeFlag%) & "2.DAT"
tfilename$ = MiscGetFileNameNoExtension$(MACFile$) & "_" & sym$ & ".TXT"
Open tfilename$ For Output As #Temp1FileNumber%

Screen.MousePointer = vbHourglass
If mode% = 7 Then msg$ = XrayGetMACTable(Int(1), sym$)
If mode% = 9 Then msg$ = XrayGetMACTable(Int(2), sym$)
Screen.MousePointer = vbDefault
If ierror Then
Close (Temp1FileNumber%)
Exit Sub
End If

Print #Temp1FileNumber, msg$
Close (Temp1FileNumber%)
End If

' MAC table (complete)
If mode% = 8 Or mode% = 10 Then

' MAC element emitter
msg$ = "Are you sure that you want to see the complete MAC table?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton2, "XrayGetTable")
If response% = vbNo Then
ierror = True
Exit Sub
End If

If mode% = 8 Then MACFile$ = ApplicationCommonAppData$ & macstring2$(MACTypeFlag%) & ".DAT"
If mode% = 10 Then MACFile$ = ApplicationCommonAppData$ & macstring2$(MACTypeFlag%) & "2.DAT"
tfilename$ = MiscGetFileNameNoExtension$(MACFile$) & "_OUTPUT.TXT"
Open tfilename$ For Output As #Temp1FileNumber%

For i% = 1 To MAXELM%
Screen.MousePointer = vbHourglass
If mode% = 8 Then msg$ = XrayGetMACTable(Int(1), Symlo$(i%))
If mode% = 10 Then msg$ = XrayGetMACTable(Int(2), Symlo$(i%))
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

Call IOWriteLog(msg$)
Print #Temp1FileNumber, msg$
Next i%
Close (Temp1FileNumber%)

msg$ = vbNullString   ' blank line last
End If

' Load into log window
Call IOWriteLog(msg$)   ' write string to log file (all except complete MAC table)
If mode% = 8 Or mode% = 10 Then MsgBox "Tab delimited output also saved to " & tfilename$, vbOKOnly + vbInformation, "XrayGetTable"
Exit Sub

' Errors
XrayGetTableError:
Close (Temp1FileNumber%)
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "XrayGetTable"
ierror = True
Exit Sub

End Sub
