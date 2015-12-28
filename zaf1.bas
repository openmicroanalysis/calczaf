Attribute VB_Name = "CodeZAF1"
' (c) Copyright 1995-2016 by John J. Donovan (credit to John Armstrong for original code)
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

' Declare global ZAF structure (for passing to other modules)
Type TypeZAF
    n8 As Long          ' sample number (can be line number)
    in1 As Integer      ' number of elements
    in0 As Integer      ' number of elements (minus stoichiometric oxygen)
    iter As Integer     ' number of iterations
    ksum As Single      ' sum of k-ratios

    TOA As Single                   ' takeoff angle
    m1(1 To MAXCHAN1%) As Single    ' takeoff parameter
    eO(1 To MAXCHAN1%) As Single    ' array for kilovolts

    p1(1 To MAXCHAN1%) As Single    ' stoichiometric proportion
    il(1 To MAXCHAN1%) As Integer   ' x-ray line number (0 = stoichiometric element, 1 = ka, 2 = kb, 3 = la, 4 = lb, 5 = ma, 6 = mb)
    Z(1 To MAXCHAN1%) As Integer    ' atomic number

    atwts(1 To MAXCHAN1%) As Single     ' atomic weights
    conc(1 To MAXCHAN1%) As Single      ' concentrations
    krat(1 To MAXCHAN1%) As Single      ' k-ratio (normalized)
    kraw(1 To MAXCHAN1%) As Single      ' raw k-ratio
    genstd(1 To MAXCHAN1%) As Single    ' pure element intensities
    gensmp(1 To MAXCHAN1%) As Single    ' sample intensities

    eng(1 To MAXRAY% - 1, 1 To MAXCHAN1%) As Single     ' emission energies
    flu(1 To MAXRAY% - 1, 1 To MAXCHAN1%) As Single     ' fluorescent yields
    edg(1 To MAXEDG%, 1 To MAXCHAN1%) As Single         ' edge energies

    eC(1 To MAXCHAN1%) As Single                    ' critical excitation energy
    v(1 To MAXCHAN1%) As Single                     ' overvoltage
    vv(1 To MAXCHAN1%) As Single                    ' fluorescence term (used to be shared with "v")

    mup() As Single   ' mass absorption coefficients (allocated in ZAFInitZAF (1 To MAXCHAN1%, 1 To MAXCHAN1%)
    r() As Single     ' pure element backscatter loss (allocated in ZAFInitZAF (1 To MAXCHAN1%, 1 To MAXCHAN1%)
    s() As Single     ' pure element sample stopping power (allocated in ZAFInitZAF (1 To MAXCHAN1%, 1 To MAXCHAN1%)
    
    bks(1 To MAXCHAN1%) As Single                   ' sample backscatter loss
    stp(1 To MAXCHAN1%) As Single                   ' sample stopping power
    zed(1 To MAXCHAN1%) As Single                   ' atomic number correction
    
    t1 As Single ' used in ZAFPtc
    t2 As Single
    t3 As Single
    t4 As Single
    g As Single
    
    model As Integer                    ' current model
    imodel As Integer                   ' number of models
    models(1 To MAXMODELS%) As Integer  ' list of geometric models
    diam As Single                      ' current diameter (in microns)
    idiam As Integer                    ' number of diameters
    diams(1 To MAXDIAMS%) As Single     ' list of diameters
    
    d As Single                         ' diameter (in cm, divide microns by MICRONSPERCM&)
    rho As Single                       ' density (in g/cm^3)
    j9 As Single                        ' particle thickness factor
    x1 As Single                        ' integration step length (in g/cm^2)
    
    ' Particle calculations
    intnum(1 To MAXCHAN1%) As Long
    erange(1 To MAXCHAN1%) As Single
    
    ' Oxides, formulas, etc. (dimensioned MAXCHAN1%, but loaded based on MAXCHAN% dimensions, without zaf oxygen channel)
    TotalCations As Single
    totalatoms As Single
    OxPercents(1 To MAXCHAN1%) As Single
    AtPercents(1 To MAXCHAN1%) As Single
    Formulas(1 To MAXCHAN1%) As Single
    NormElPercents(1 To MAXCHAN1%) As Single
    NormOxPercents(1 To MAXCHAN1%) As Single
    
    ' Coating corrections (for emitting x-ray absorption and electron absorption)
    coating_flag As Integer
    coating_sin_thickness As Single                               ' x-ray absorption path length
    coating_actual_kilovolts(1 To MAXCHAN%) As Single             ' actual beam energy (after electron absorption beam energy loss)
    
    coating_trans_smp(1 To MAXCHAN1%) As Single                   ' x-ray transmission for sample coating
    coating_trans_std() As Single                                 ' x-ray transmission for standard coating (allocated in ZAFInitZAF (1 To MAXSTD%, 1 To MAXCHAN1%)
    coating_trans_std_assigns(1 To MAXCHAN1%) As Single           ' x-ray transmission for assigned standard coating
    
    coating_absorbs_smp(1 To MAXCHAN1%) As Single                 ' electron absorption for sample coating
    coating_absorbs_std() As Single                               ' electron absorption for standard coating (allocated in ZAFInitZAF (1 To MAXSTD%, 1 To MAXCHAN1%)
    coating_absorbs_std_assigns(1 To MAXCHAN1%) As Single         ' electron absorption for assigned standard coating
    
    coating_std_assigns_element(1 To MAXCHAN1%) As Integer        ' coating element for assigned standard coating
    coating_std_assigns_density(1 To MAXCHAN1%) As Single         ' coating density for assigned standard coating
    coating_std_assigns_thickness(1 To MAXCHAN1%) As Single       ' coating thickness for assigned standard coating
    coating_std_assigns_sinthickness(1 To MAXCHAN1%) As Single    ' coating sin thickness for assigned standard coating
    
    MACs(1 To MAXCHAN1%) As Single                        ' individual emitting element MACs
End Type

Sub ZAFReadMu(zaf As TypeZAF)
' This routine reads MAC's for the primary analytical lines

ierror = False
On Error GoTo ZAFReadMuError

Dim im4 As Integer, i As Integer, i1 As Integer, i3 As Integer
Dim mac As Single

' Check for lines close to edges
For i% = 1 To zaf.in1%
If zaf.il%(i%) > MAXRAY% - 1 Then GoTo 1020     ' skip absorber only elements
im4% = zaf.il%(i%)

' Check for additional lines. If found, check that FFAST2.DAT files exists.
If im4% > MAXRAY_OLD% Then
If MACTypeFlag% <> 6 Then GoTo ZAFReadMuFFASTNotSpecified
MACFile$ = ApplicationCommonAppData$ & macstring2$(MACTypeFlag%) & "2.DAT"
If Dir$(MACFile$) = vbNullString Then GoTo ZAFReadMuFFAST2DATNotFound
End If

' Loop on all emitters, all edges
For i1% = 1 To zaf.in0%
For i3% = 1 To MAXEDG%
If zaf.edg!(i3, i1%) = 0# Then GoTo 990
If zaf.edg!(i3, i1%) <= (zaf.eng!(im4%, i%) - 0.1) Then GoTo 990
If zaf.edg!(i3, i1%) >= (zaf.eng!(im4%, i%) + 0.03) Then GoTo 990

If VerboseMode Then
msg$ = "WARNING in ZAFReadMu- " & Format$(Symup$(zaf.Z%(i%)), a20$) & " " & Format$(Xraylo$(im4%), a20$) & " is close to the " & Format$(Edglo$(i3%), a20$) & " edge of " & Format$(Symup$(zaf.Z%(i1%)), a20$)
Call IOWriteLog(msg$)
End If
990:   Next i3%
1010:  Next i1%
1020:  Next i%

' zaf.z%(i%)  is atomic # of emitting element
' zaf.z%(i1%) is atomic # of absorbing element
' zaf.mup!(i1%,i%) is mass absorption coefficient (MACs)
If VerboseMode Then Call IOWriteLog("Now loading MACs for primary lines...")

' Load MAC values for primary analytical lines
For i% = 1 To zaf.in1%
If zaf.il%(i%) > MAXRAY% - 1 Then GoTo 1060     ' skip absorber only elements
im4% = zaf.il%(i%)

' Load MAC values for all emitter absorber pairs
For i1% = 1 To zaf.in0%
Call ZAFLoadMac(i%, im4%, i1%, mac!, zaf)
If ierror Then Exit Sub

' Unable to load MAC, load rough MAC and type error message
If mac! = 0# Then
If im4% = 1 Or im4% = 2 Then
zaf.mup!(i1%, i%) = 10#
If zaf.Z%(i%) < 30 Then zaf.mup!(i1%, i%) = 100#
If zaf.Z%(i%) < 20 Then zaf.mup!(i1%, i%) = 1000#
If zaf.Z%(i%) < 10 Then zaf.mup!(i1%, i%) = 10000#
End If
If im4% = 3 Or im4% = 4 Then zaf.mup!(i1%, i%) = 100#
If im4% = 5 Or im4% = 6 Then zaf.mup!(i1%, i%) = 1000#
If im4% > MAXRAY_OLD% Then zaf.mup!(i1%, i%) = 2000#
msg$ = "WARNING in ZAFReadMu- MAC not loaded for " & Symup$(zaf.Z%(i%)) & " " & Xraylo$(im4%) & " in " & Symup$(zaf.Z%(i1%)) & ", will assume a MAC value of " & Format$(zaf.mup!(i1%, i%)) & " for this line."
Call IOWriteLog(msg$)
Else
zaf.mup!(i1%, i%) = mac!
End If

1050:  Next i1%
1060:  Next i%

Exit Sub

' Errors
ZAFReadMuError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFReadMu"
ierror = True
Exit Sub

ZAFReadMuFFASTNotSpecified:
msg$ = "For quantification of additional x-ray lines, you must specify the FFAST MAC database." & vbCrLf & vbCrLf
msg$ = msg$ & "See the Analytical | ZAF, Phi-Rho-Z, Alpha Factor and Calibration Curve Selections menu and select the MACs button."
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFReadMu"
ierror = True
Exit Sub

ZAFReadMuFFAST2DATNotFound:
msg$ = "File " & MACFile$ & " was not found. You will need to re-run the CalcZAF.msi installer to obtain the new x-ray tables for additional x-ray lines."
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFReadMu"
ierror = True
Exit Sub

End Sub

Sub ZAFLoadMac(i1 As Integer, i2 As Integer, i4 As Integer, mac As Single, zaf As TypeZAF)
' Load alternative mass absorption coefficient using zaf array channel numbers
'  i1 = emitting channel (1 to MAXCHAN%)
'  i2 = emitting x-ray line (1 to MAXRAY%)
'  i4 = absorbing channel (1 to MAXCHAN%)

ierror = False
On Error GoTo ZAFLoadMacError

Dim num As Integer, emtz As Integer, absz As Integer
Dim energy As Single
Dim tstring As String
Dim aelastic As Single, ainelastic As Single, aphoto As Single
Dim tfactor As Single, tstandard As String

Dim macrow As TypeMu

' Initialize MAC
mac! = 0#
emtz% = zaf.Z%(i1%)
absz% = zaf.Z%(i4%)

' Try to load MAC from empirical array
If UseMACFlag Then
Call EmpLoadMACAPF(Int(1), emtz%, i2%, absz%, mac!, tstring$, tfactor!, tstandard$)
If ierror Then Exit Sub
If mac! > 0# Then
If DebugMode Then
msg$ = "WARNING in ZAFLoadMac- Loading empirical MAC (" & Format$(mac!) & ") for " & Format$(Symup$(emtz%), a20$) & " " & Format$(Xraylo$(i2%), a20$) & " in " & Format$(Symup$(absz%), a20$)
Call IOWriteLog(msg$)
End If
Exit Sub
End If
End If

' Now try to load from disk file if not available from empirical array
If i2% <= MAXRAY_OLD% Then
MACFile$ = ApplicationCommonAppData$ & macstring2$(MACTypeFlag%) & ".DAT"
If Dir$(MACFile$) = vbNullString Then GoTo ZAFLoadMACNotFound
Open MACFile$ For Random Access Read As #MACFileNumber% Len = MAC_FILE_RECORD_LENGTH%
Get #MACFileNumber%, emtz%, macrow
Close #MACFileNumber%

num% = i2% + (absz% - 1) * (MAXRAY_OLD%)
mac! = macrow.mac!(num%)
If mac! > 0# Then
Exit Sub
End If

' Load from additional lines MAC file
Else
MACFile$ = ApplicationCommonAppData$ & macstring2$(MACTypeFlag%) & "2.DAT"
If Dir$(MACFile$) <> vbNullString Then
Open MACFile$ For Random Access Read As #MACFileNumber% Len = MAC_FILE_RECORD_LENGTH%
Get #MACFileNumber%, emtz%, macrow
Close #MACFileNumber%

num% = (i2% - MAXRAY_OLD%) + (absz% - 1) * (MAXRAY_OLD%)
mac! = macrow.mac!(num%)
If mac! > 0# Then
Exit Sub
End If
End If
End If

' Now try to calculate using McMaster if not available from disk file
energy! = zaf.eng!(i2%, i1%)
Call AbsorbGetMAC(absz%, energy!, aphoto!, aelastic!, ainelastic!, mac!)
If ierror Then
Exit Sub
End If
If mac! > 0# Then
If VerboseMode Then
msg$ = "WARNING in ZAFLoadMac- MAC (" & Format$(mac!) & ") calculated for " & Format$(Symup$(emtz%), a20$) & " " & Format$(Xraylo$(i2%), a20$) & " in " & Format$(Symup$(absz%), a20$) & " using McMaster expressions."
Call IOWriteLog(msg$)
End If
Exit Sub
End If

Exit Sub

' Errors
ZAFLoadMacError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFLoadMac"
Close #MACFileNumber%
ierror = True
Exit Sub

ZAFLoadMACNotFound:
msg$ = "File " & MACFile$ & " was not found, please choose another MAC file or create the missing file using the CalcZAF Xray menu items"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFLoadMAC"
ierror = True
Exit Sub

End Sub

Sub ZAFLoadMac2(i1 As Integer, i2 As Integer, i4 As Integer, mac As Single)
' Load alternative mass absorption coefficient (LINEMU.DAT, CITZMU.DAT, etc) using atomic numbers
'  i1 = emitting atomic number (1 to MAXELM%)
'  i2 = emitting x-ray line (1 to MAXRAY%)
'  i4 = absorbing atomic number (1 to MAXELM%)

ierror = False
On Error GoTo ZAFLoadMac2Error

Dim num As Integer, emtz As Integer, absz As Integer
Dim aenergy As Single, aflur As Single
Dim tstring As String
Dim aelastic As Single, ainelastic As Single, aphoto As Single
Dim tfactor As Single, tstandard As String

Dim macrow As TypeMu

' Initialize MAC
mac! = 0#
emtz% = i1%
absz% = i4%

' Try to load MAC from empirical array
If UseMACFlag Then
Call EmpLoadMACAPF(Int(1), emtz%, i2%, absz%, mac!, tstring$, tfactor!, tstandard$)
If ierror Then Exit Sub
If mac! > 0# Then
If DebugMode Then
msg$ = "WARNING in ZAFLoadMac2- Loading empirical MAC for " & Format$(Symup$(emtz%), a20$) & " " & Format$(Xraylo$(i2%), a20$) & " in " & Format$(Symup$(absz%), a20$)
Call IOWriteLog(msg$)
End If
Exit Sub
End If
End If

' Now try to load from disk file if not available from empirical array
If i2% <= MAXRAY_OLD% Then
MACFile$ = ApplicationCommonAppData$ & macstring2$(MACTypeFlag%) & ".DAT"
If Dir$(MACFile$) = vbNullString Then GoTo ZAFLoadMAC2NotFound
Open MACFile$ For Random Access Read As #MACFileNumber% Len = MAC_FILE_RECORD_LENGTH%
Get #MACFileNumber%, emtz%, macrow
Close #MACFileNumber%

num% = i2% + (absz% - 1) * (MAXRAY_OLD%)
mac! = macrow.mac!(num%)
If mac! > 0# Then
Exit Sub
End If

' Load MAC from additional lines file
Else
MACFile$ = ApplicationCommonAppData$ & macstring2$(MACTypeFlag%) & "2.DAT"
If Dir$(MACFile$) = vbNullString Then GoTo ZAFLoadMAC2NotFound
Open MACFile$ For Random Access Read As #MACFileNumber% Len = MAC_FILE_RECORD_LENGTH%
Get #MACFileNumber%, emtz%, macrow
Close #MACFileNumber%

num% = (i2% - MAXRAY_OLD%) + (absz% - 1) * (MAXRAY_OLD%)
mac! = macrow.mac!(num%)
If mac! > 0# Then
Exit Sub
End If
End If

' Now try to calculate using McMaster if not available from disk file
Call ZAFLoadXray(i1%, i2%, aenergy!, aflur!)
If ierror Then Exit Sub
Call AbsorbGetMAC(absz%, aenergy!, aphoto!, aelastic!, ainelastic!, mac!)
If ierror Then
Exit Sub
End If
If mac! > 0# Then
If DebugMode Then
msg$ = "WARNING in ZAFLoadMac2- MAC calculated for " & Format$(Symup$(emtz%), a20$) & " " & Format$(Xraylo$(i2%), a20$) & " in " & Format$(Symup$(absz%), a20$) & " using McMaster expressions."
Call IOWriteLog(msg$)
End If
Exit Sub
End If

Exit Sub

' Errors
ZAFLoadMac2Error:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFLoadMac2"
Close #MACFileNumber%
ierror = True
Exit Sub

ZAFLoadMAC2NotFound:
msg$ = "File " & MACFile$ & " was not found, please choose another MAC file or create the missing file using the CalcZAF Xray menu items"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFLoadMAC2"
ierror = True
Exit Sub

End Sub

Sub ZAFLoadXray(i As Integer, i2 As Integer, aenergy As Single, aflur As Single)
' This routine reads the x-ray line energy and x-ray fluorescence yield for a specified atomic number emitter

ierror = False
On Error GoTo ZAFLoadXrayError

Dim nrec As Integer

Dim engrow As TypeEnergy
Dim flurow As TypeFlur

' Load record number
nrec% = i% + 2

' Open x-ray files (for original lines)
If i2% <= MAXRAY_OLD% Then

' Open x-ray line file
Open XLineFile$ For Random Access Read As #XLineFileNumber% Len = XRAY_FILE_RECORD_LENGTH%

' Open x-ray flur file
Open XFlurFile$ For Random Access Read As #XFlurFileNumber% Len = XRAY_FILE_RECORD_LENGTH%

' Read emission lines (convert to keV)
Get #XLineFileNumber%, nrec%, engrow
aenergy! = engrow.energy!(i2%) / EVPERKEV#

' Read fluorescent yields
Get #XFlurFileNumber%, nrec%, flurow
aflur! = flurow.fraction!(i2%)

Close #XLineFileNumber%
Close #XFlurFileNumber%

' Open x-ray edge file (for additional lines)
Else

' Open x-ray line file
If Dir$(XLineFile2$) = vbNullString Then GoTo ZAFLoadXrayNotFoundXLINE2DAT
If FileLen(XLineFile2$) = 0 Then GoTo ZAFLoadXrayZeroSizeXLINE2DAT
Open XLineFile2$ For Random Access Read As #XLineFileNumber2% Len = XRAY_FILE_RECORD_LENGTH%

' Open x-ray flur file
If Dir$(XFlurFile2$) = vbNullString Then GoTo ZAFLoadXrayNotFoundXFLUR2DAT
If FileLen(XFlurFile2$) = 0 Then GoTo ZAFLoadXrayZeroSizeXFLUR2DAT
Open XFlurFile2$ For Random Access Read As #XFlurFileNumber2% Len = XRAY_FILE_RECORD_LENGTH%

' Read emission lines (convert to keV)
Get #XLineFileNumber2%, nrec%, engrow
aenergy! = engrow.energy!(i2% - MAXRAY_OLD%) / EVPERKEV#

' Read fluorescent yields
Get #XFlurFileNumber2%, nrec%, flurow
aflur! = flurow.fraction!(i2% - MAXRAY_OLD%)

Close #XLineFileNumber2%
Close #XFlurFileNumber2%
End If

Exit Sub

' Errors
ZAFLoadXrayError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFLoadXray"
Close #XLineFileNumber%
Close #XFlurFileNumber%
Close #XLineFileNumber2%
Close #XFlurFileNumber2%
ierror = True
Exit Sub

ZAFLoadXrayNotFoundXLINE2DAT:
msg$ = "The " & XLineFile2$ & " was not found." & vbCrLf & vbCrLf
msg$ = msg$ & "Please run the latest CalcZAF.msi installer to obtain this additional x-ray line file."
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFLoadXray"
Close #XLineFileNumber%
Close #XFlurFileNumber%
Close #XLineFileNumber2%
Close #XFlurFileNumber2%
ierror = True
Exit Sub

ZAFLoadXrayNotFoundXFLUR2DAT:
msg$ = "The " & XFlurFile2$ & " was not found." & vbCrLf & vbCrLf
msg$ = msg$ & "Please run the latest CalcZAF.msi installer to obtain this additional x-ray line file."
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFLoadXray"
Close #XLineFileNumber%
Close #XFlurFileNumber%
Close #XLineFileNumber2%
Close #XFlurFileNumber2%
ierror = True
Exit Sub

ZAFLoadXrayZeroSizeXLINE2DAT:
Kill XLineFile2$
msg$ = "The " & XLineFile2$ & " was not found." & vbCrLf & vbCrLf
msg$ = msg$ & "Please run the latest CalcZAF.msi installer to obtain this additional x-ray line file."
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFLoadXray"
Close #XLineFileNumber%
Close #XFlurFileNumber%
Close #XLineFileNumber2%
Close #XFlurFileNumber2%
ierror = True
Exit Sub

ZAFLoadXrayZeroSizeXFLUR2DAT:
Kill XFlurFile2$
msg$ = "The " & XFlurFile2$ & " was not found." & vbCrLf & vbCrLf
msg$ = msg$ & "Please run the latest CalcZAF.msi installer to obtain this additional x-ray line file."
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFLoadXray"
Close #XLineFileNumber%
Close #XFlurFileNumber%
Close #XLineFileNumber2%
Close #XFlurFileNumber2%
ierror = True
Exit Sub

End Sub

Function ZAFErrorFunction(erfx As Single) As Single
' Error function

ierror = False
On Error GoTo ZAFErrorFunctionError

Dim i1 As Integer
Dim temp As Single
Dim erfp As Single, erft As Single

ReDim Erf(1 To 5) As Single

erfp! = 0.3275911
Erf!(1) = 0.254829592
Erf!(2) = -0.284496736
Erf!(3) = 1.421413741
Erf!(4) = -1.453152027
Erf!(5) = 1.061405429

erft! = 1# / (1# + erfp! * erfx!)

temp! = 0#
For i1% = 1 To 5
temp! = temp! + Erf!(i1%) * erft! ^ i1%
Next i1%
ZAFErrorFunction = temp!

Exit Function

' Errors
ZAFErrorFunctionError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFErrorFunction"
ierror = True
Exit Function

End Function

Function ZAFMACCal(i As Integer, zaf As TypeZAF) As Single
' Sums the weight fraction MACs (i is the emitting x-ray, i1 is the absorber)

Dim i1 As Integer
Dim m7 As Single

m7! = 0#
For i1% = 1 To zaf.in0%
m7! = m7! + zaf.conc!(i1%) * zaf.mup!(i1%, i%)
Next i1%

' Save in ZAF array also
zaf.MACs!(i%) = m7!

ZAFMACCal! = m7!
Exit Function

End Function

Function ZAFCalculateFlurYield(iz As Integer, ix As Integer) As Single
' Function to calculate fluorescent yield
'  iz = atomic number
'  ix = x-ray line (1=ka, 2=kb, 3=La, 4=lb, 5=ma, 6=mb)

ierror = False
On Error GoTo ZAFCalculateFlurYieldError

Dim temp As Single

' Ka or kb
If ix% = 1 Or ix% = 2 Then
temp! = -0.02805120355
temp! = temp! + 0.02238036235 * iz%
temp! = temp! - 0.004761377724 * iz% ^ 2
temp! = temp! + 0.0004119476258 * iz% ^ 3
temp! = temp! - 0.00001465794432 * iz% ^ 4
temp! = temp! + 2.713177255E-07 * iz% ^ 5
temp! = temp! - 2.775617967E-09 * iz% ^ 6
temp! = temp! + 1.493469918E-11 * iz% ^ 7
temp! = temp! - 3.308086081E-14 * iz% ^ 8

' La
ElseIf ix% = 3 Then
temp! = -0.06742420934
temp! = temp! + 0.008123020763 * iz%
temp! = temp! - 0.0003246758117 * iz% ^ 2
temp! = temp! + 0.000005481143052 * iz% ^ 3
temp! = temp! - 2.392032369E-08 * iz% ^ 4

' Lb
ElseIf ix% = 4 Then
temp! = 0.5630566911
temp! = temp! - 0.1299455199 * iz%
temp! = temp! + 0.01210454893 * iz% ^ 2
temp! = temp! - 0.000598756947 * iz% ^ 3
temp! = temp! + 0.00001734156803 * iz% ^ 4
temp! = temp! - 3.023765803E-07 * iz% ^ 5
temp! = temp! + 3.12096042E-09 * iz% ^ 6
temp! = temp! - 1.750253258E-11 * iz% ^ 7
temp! = temp! + 4.095475997E-14 * iz% ^ 8

' Ma
ElseIf ix% = 5 Then
temp! = 0.00237 * iz% - 0.163

' Mb
ElseIf ix% = 6 Then
temp! = 0.00237 * iz% - 0.163   ' need to get a mb specific fit

End If

ZAFCalculateFlurYield! = temp!
Exit Function

' Errors
ZAFCalculateFlurYieldError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFCalculateFlurYield"
ierror = True
Exit Function

End Function

Sub ZAFReadLn(zaf As TypeZAF)
' This routine reads the x-ray line, x-ray edge and x-ray fluorescense yield data for this run. The xray emission line energies are read
' in electron volts in the order KA, KB, LA, LB, MA, MB. The edge energies are read in electron volts in the order K, L-I, L-II,
' L-III, M-I, M-II, M-III, M-IV, and M-V. Fluorescent yields are read in the same order as the emission line energies.

ierror = False
On Error GoTo ZAFReadLnError

Dim nrec As Integer, im4 As Integer, im5 As Integer
Dim i As Integer, i2 As Integer

Dim engrow As TypeEnergy
Dim edgrow As TypeEdge
Dim flurow As TypeFlur

' Open x-ray edge file
Open XEdgeFile$ For Random Access Read As #XEdgeFileNumber% Len = XRAY_FILE_RECORD_LENGTH%

' Open x-ray line file
Open XLineFile$ For Random Access Read As #XLineFileNumber% Len = XRAY_FILE_RECORD_LENGTH%

' Open x-ray flur file
Open XFlurFile$ For Random Access Read As #XFlurFileNumber% Len = XRAY_FILE_RECORD_LENGTH%

' Open x-ray line file for additional x-rays
If Dir$(XLineFile2$) = vbNullString Then GoTo ZAFReadLnNotFoundXLINE2DAT
If FileLen(XLineFile2$) = 0 Then GoTo ZAFReadLnZeroSizeXLINE2DAT
Open XLineFile2$ For Random Access Read As #XLineFileNumber2% Len = XRAY_FILE_RECORD_LENGTH%

' Open x-ray flur file for additional x-rays
If Dir$(XFlurFile2$) = vbNullString Then GoTo ZAFReadLnNotFoundXFLUR2DAT
If FileLen(XFlurFile2$) = 0 Then GoTo ZAFReadLnZeroSizeXFLUR2DAT
Open XFlurFile2$ For Random Access Read As #XFlurFileNumber2% Len = XRAY_FILE_RECORD_LENGTH%

' Loop on each emitter in matrix
For i% = 1 To zaf.in0%
nrec% = zaf.Z%(i%) + 2

' Read all absorption edges for this element (convert to keV)
Get #XEdgeFileNumber%, nrec%, edgrow
For i2% = 1 To MAXEDG%
zaf.edg!(i2%, i%) = edgrow.energy!(i2%) / EVPERKEV#
Next i2%

' Read all original emission lines for this element (convert to keV)
Get #XLineFileNumber%, nrec%, engrow
For i2% = 1 To MAXRAY_OLD%
zaf.eng!(i2%, i%) = engrow.energy!(i2%) / EVPERKEV#
Next i2%

' Read all original fluorescent yields for this element
Get #XFlurFileNumber%, nrec%, flurow
For i2% = 1 To MAXRAY_OLD%
zaf.flu!(i2%, i%) = flurow.fraction!(i2%)
Next i2%

' Read all additional emission lines for this element (convert to keV)
Get #XLineFileNumber2%, nrec%, engrow
For i2% = MAXRAY_OLD% + 1 To MAXRAY% - 1
zaf.eng!(i2%, i%) = engrow.energy!(i2% - MAXRAY_OLD%) / EVPERKEV#
Next i2%

' Read all additional fluorescent yields for this element
Get #XFlurFileNumber2%, nrec%, flurow
For i2% = MAXRAY_OLD% + 1 To MAXRAY% - 1
zaf.flu!(i2%, i%) = flurow.fraction!(i2% - MAXRAY_OLD%)
Next i2%

' Calculate yields if zero value loaded form original x-ray lines
If zaf.flu!(1, i%) = 0# Then zaf.flu!(1, i%) = ZAFCalculateFlurYield(zaf.Z%(i%), Int(1))   ' calculate Ka yields
If zaf.flu!(1, i%) < 0# Then zaf.flu!(1, i%) = 0#

If zaf.flu!(2, i%) = 0# Then zaf.flu!(2, i%) = ZAFCalculateFlurYield(zaf.Z%(i%), Int(2))   ' calculate Kb yields
If zaf.flu!(2, i%) < 0# Then zaf.flu!(2, i%) = 0#

If zaf.flu!(3, i%) = 0# Then zaf.flu!(3, i%) = ZAFCalculateFlurYield(zaf.Z%(i%), Int(3))   ' calculate La yields
If zaf.flu!(3, i%) < 0# Then zaf.flu!(3, i%) = 0#

If zaf.flu!(4, i%) = 0# Then zaf.flu!(4, i%) = ZAFCalculateFlurYield(zaf.Z%(i%), Int(4))   ' calculate Lb yields
If zaf.flu!(4, i%) < 0# Then zaf.flu!(4, i%) = 0#

If zaf.flu!(5, i%) = 0# Then zaf.flu!(5, i%) = ZAFCalculateFlurYield(zaf.Z%(i%), Int(5))   ' calculate Ma yields
If zaf.flu!(5, i%) < 0# Then zaf.flu!(5, i%) = 0#

If zaf.flu!(6, i%) = 0# Then zaf.flu!(6, i%) = ZAFCalculateFlurYield(zaf.Z%(i%), Int(6))   ' calculate Mb yields
If zaf.flu!(6, i%) < 0# Then zaf.flu!(6, i%) = 0#

' Calculate yields if zero value loaded for additional x-ray lines
If MAXRAY% - 1 > MAXRAY_OLD% Then
If zaf.flu!(7, i%) = 0# Then zaf.flu!(7, i%) = ZAFCalculateFlurYield(zaf.Z%(i%), Int(7))   ' calculate Ln yields
If zaf.flu!(7, i%) < 0# Then zaf.flu!(7, i%) = 0#

If zaf.flu!(8, i%) = 0# Then zaf.flu!(8, i%) = ZAFCalculateFlurYield(zaf.Z%(i%), Int(8))   ' calculate Lg yields
If zaf.flu!(8, i%) < 0# Then zaf.flu!(8, i%) = 0#

If zaf.flu!(9, i%) = 0# Then zaf.flu!(9, i%) = ZAFCalculateFlurYield(zaf.Z%(i%), Int(9))   ' calculate Lv yields
If zaf.flu!(9, i%) < 0# Then zaf.flu!(9, i%) = 0#

If zaf.flu!(10, i%) = 0# Then zaf.flu!(10, i%) = ZAFCalculateFlurYield(zaf.Z%(i%), Int(10))   ' calculate Ll yields
If zaf.flu!(10, i%) < 0# Then zaf.flu!(10, i%) = 0#

If zaf.flu!(11, i%) = 0# Then zaf.flu!(11, i%) = ZAFCalculateFlurYield(zaf.Z%(i%), Int(11))   ' calculate Mg yields
If zaf.flu!(11, i%) < 0# Then zaf.flu!(11, i%) = 0#

If zaf.flu!(12, i%) = 0# Then zaf.flu!(12, i%) = ZAFCalculateFlurYield(zaf.Z%(i%), Int(12))   ' calculate Mz yields
If zaf.flu!(12, i%) < 0# Then zaf.flu!(12, i%) = 0#
End If
Next i%

' Loop on each emitter in matrix and check for bad values
For i% = 1 To zaf.in0%

' Now load x-ray line types for all x-rays if emitting line (skip absorber only)
If zaf.il%(i%) >= 1 And zaf.il%(i%) <= MAXRAY% - 1 Then
im4% = zaf.il%(i%)

' Calculate edge index for each line overvoltage (K, L-I, L-II, L-III, M-I, M-II, M-III, M-IV, and M-V)
If im4% = 1 Then im5% = 1   ' Ka
If im4% = 2 Then im5% = 1   ' Kb
If im4% = 3 Then im5% = 4   ' La
If im4% = 4 Then im5% = 3   ' Lb
If im4% = 5 Then im5% = 9   ' Ma
If im4% = 6 Then im5% = 8   ' Mb

If im4% = 7 Then im5% = 3    ' Ln
If im4% = 8 Then im5% = 3    ' Lg
If im4% = 9 Then im5% = 3    ' Lv
If im4% = 10 Then im5% = 4   ' Ll
If im4% = 11 Then im5% = 7   ' Mg
If im4% = 12 Then im5% = 9   ' Mz

' Calculate overvoltages
zaf.eC!(i%) = zaf.edg!(im5%, i%)
zaf.gensmp!(i%) = zaf.eng!(im4%, i%)
If zaf.eC!(i%) = 0# Then
msg$ = "Warning in ZAFReadLn: Edge energy of " & Symup$(zaf.Z%(i%)) & " " & Xraylo$(zaf.il%(i%)) & " is zero. Overvoltage was not calculated."
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFReadLn"
Else
zaf.v!(i%) = zaf.eO!(i%) / zaf.eC!(i%)
End If

' Check overvoltage values
If zaf.il%(i%) < 1 Or zaf.il%(i%) > MAXRAY% - 1 Then GoTo 750

'If zaf.v!(i%) <= 1# Then
'msg$ = "Error in ZAFReadLn: Operating voltage is less than or equal to the " & Symup$(zaf.z%(i%)) & " " & Xraylo$(zaf.il%(i%)) & " absorption edge at " & Str$(zaf.ec!(i%)) & " KeV"
'MsgBox msg$, vbOKOnly + vbExclamation, "ZAFReadLn"
'Close #XEdgeFileNumber%
'Close #XLineFileNumber%
'Close #XFlurFileNumber%
'ierror = True
'Exit Sub
'End If

If zaf.v!(i%) < 1.1 Then
msg$ = "Warning in ZAFReadLn: Overvoltage of " & Symup$(zaf.Z%(i%)) & " " & Xraylo$(zaf.il%(i%)) & " is only " & Str$(zaf.v!(i%))
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
End If

End If      ' skip absorber only
750:  Next i%

Close #XEdgeFileNumber%
Close #XLineFileNumber%
Close #XFlurFileNumber%
Close #XLineFileNumber2%
Close #XFlurFileNumber2%

Exit Sub

' Errors
ZAFReadLnError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFReadLn"
Close #XEdgeFileNumber%
Close #XLineFileNumber%
Close #XFlurFileNumber%
ierror = True
Exit Sub

ZAFReadLnNotFoundXLINE2DAT:
msg$ = "The " & XLineFile2$ & " was not found." & vbCrLf & vbCrLf
msg$ = msg$ & "Please run the latest CalcZAF.msi installer to obtain this additional x-ray line file."
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFReadLn"
Close #XEdgeFileNumber%
Close #XLineFileNumber%
Close #XFlurFileNumber%
Close #XLineFileNumber2%
Close #XFlurFileNumber2%
ierror = True
Exit Sub

ZAFReadLnNotFoundXFLUR2DAT:
msg$ = "The " & XFlurFile2$ & " was not found." & vbCrLf & vbCrLf
msg$ = msg$ & "Please run the latest CalcZAF.msi installer to obtain this additional x-ray line file."
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFReadLn"
Close #XEdgeFileNumber%
Close #XLineFileNumber%
Close #XFlurFileNumber%
Close #XLineFileNumber2%
Close #XFlurFileNumber2%
ierror = True
Exit Sub

ZAFReadLnZeroSizeXLINE2DAT:
Kill XLineFile2$
msg$ = "The " & XLineFile2$ & " was not found." & vbCrLf & vbCrLf
msg$ = msg$ & "Please run the latest CalcZAF.msi installer to obtain this additional x-ray line file."
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFReadLn"
Close #XEdgeFileNumber%
Close #XLineFileNumber%
Close #XFlurFileNumber%
Close #XLineFileNumber2%
Close #XFlurFileNumber2%
ierror = True
Exit Sub

ZAFReadLnZeroSizeXFLUR2DAT:
Kill XFlurFile2$
msg$ = "The " & XFlurFile2$ & " was not found." & vbCrLf & vbCrLf
msg$ = msg$ & "Please run the latest CalcZAF.msi installer to obtain this additional x-ray line file."
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFReadLn"
Close #XEdgeFileNumber%
Close #XLineFileNumber%
Close #XFlurFileNumber%
Close #XLineFileNumber2%
Close #XFlurFileNumber2%
ierror = True
Exit Sub

End Sub

Sub ZAFGetContinuumAbsorption(continuum_absorbtion() As Single, zaf As TypeZAF)
' Calculate continuum absorption (modified Heinrich from Myklebust) for all emitters

ierror = False
On Error GoTo ZAFGetContinuumAbsorptionError

Dim i As Integer

ReDim m7(1 To MAXCHAN1%) As Single
ReDim h(1 To MAXCHAN1%) As Single
ReDim gsmp(1 To MAXCHAN1%) As Single
ReDim gstd(1 To MAXCHAN1%) As Single

' Calculate standard (pure element) absorption
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
h!(i%) = 0.0000012 * (zaf.eO!(i%) ^ 1.65 - zaf.eC!(i%) ^ 1.65)
gstd!(i%) = (1# + h!(i%) * zaf.mup!(i%, i%) * zaf.m1!(i%)) ^ 2

' Modify intensity using depth production and anisotropy from Small and Myklebust
gstd!(i%) = gstd!(i%) * 1.15 - 0.15 * 1# / gstd!(i%)
End If
Next i%

' Calculate sample (actual standard composition) absorption
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
m7!(i%) = ZAFMACCal(i%, zaf)
gsmp!(i%) = (1# + h!(i%) * m7!(i%) * zaf.m1!(i%)) ^ 2

' Modify intensity using depth production and anisotropy from Small and Myklebust
gsmp!(i%) = gsmp!(i%) * 1.15 - 0.15 * 1# / gsmp!(i%)
End If
Next i%

' Create continuum absorption correction factors
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
continuum_absorbtion!(i%) = gsmp!(i%) / gstd!(i%)
End If
Next i%

Exit Sub

' Errors
ZAFGetContinuumAbsorptionError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFGetContinuumAbsorption"
ierror = True
Exit Sub

End Sub

Function ZAFFNLint(X9 As Single) As Single
'LOGARITHMIC INTEGRAL CALCULATION FUNCTION

ierror = False
On Error GoTo ZAFFNLintError

Dim X7 As Single, X8 As Single
Dim xx1 As Single, XX3 As Single, XX4 As Single, XX5 As Single

xx1! = Log(X9!)
XX4! = 0#
XX3! = 1#
XX5! = 1#
X8! = Log(Abs(xx1!)) + 0.577216
      
Do Until (XX3! < 0.00005) And (XX4! > 2#)
XX4! = XX4! + 1#
XX5! = XX5! * XX4!
X7! = X8!
X8! = X8! + (xx1! ^ XX4!) / (XX4! * XX5!)
XX3! = Abs(X8! - X7!)
Loop
      
ZAFFNLint! = X8!
Exit Function

' Errors
ZAFFNLintError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFFNLint"
ierror = True
Exit Function

End Function

Sub ZAFQsCalc(zafinit As Integer, zaf As TypeZAF)
' Numerical integration

ierror = False
On Error GoTo ZAFQsCalcError

msg$ = "Numerical integration of stopping power is not implemented"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFQsCalc"
ierror = True

Exit Sub

' Errors
ZAFQsCalcError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFQsCalc"
ierror = True
Exit Sub

End Sub

Function ZAFCalculateEnergyLoss(chan As Integer, sample() As TypeSample) As Single
' Calculate electron beam energy loss for the specified coating, density and thickness

ierror = False
On Error GoTo ZAFCalculateEnergyLossError

Dim energy As Single

' Electron energy (final)
ZAFCalculateEnergyLoss! = sample(1).kilovolts!

' Calculate in microns
If Not sample(1).CombinedConditionsFlag Then
energy! = (sample(1).kilovolts! ^ 1.67 - (sample(1).CoatingDensity! * sample(1).CoatingThickness! / ANGPERMICRON& * AllAtomicNums%(sample(1).CoatingElement%) ^ 0.89) / (0.276 * AllAtomicWts!(sample(1).CoatingElement%))) ^ (1# / 1.67)
Else
energy! = (sample(1).KilovoltsArray!(chan%) ^ 1.67 - (sample(1).CoatingDensity! * sample(1).CoatingThickness! / ANGPERMICRON& * AllAtomicNums%(sample(1).CoatingElement%) ^ 0.89) / (0.276 * AllAtomicWts!(sample(1).CoatingElement%))) ^ (1# / 1.67)
End If

ZAFCalculateEnergyLoss! = energy!
Exit Function

' Errors
ZAFCalculateEnergyLossError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFCalculateEnergyLoss"
ierror = True
Exit Function

End Function
