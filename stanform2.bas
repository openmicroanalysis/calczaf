Attribute VB_Name = "CodeSTANFORM2"
' (c) Copyright 1995-2024 by John J. Donovan
Option Explicit

Const MAXCONTINUUM% = 5
Const MAXZBAR% = 10

Sub StanFormCalculateZbars(maszbar As Single, zedzbar As Single, sample() As TypeSample)
' Calculate alternative z-bars

ierror = False
On Error GoTo StanFormCalculateZbarsError

Dim i As Integer, j As Integer
Dim n As Integer, ii As Integer, jj As Integer
Dim aexp As Single, temp As Single
Dim temp1 As Single, temp2 As Single

Dim atmzbar As Single    ' atomic fraction
Dim elazbar As Single    ' elastic scattering fraction

Dim salzbar As Single    ' Saldick and Allen
Dim joyzbar As Single    ' Joyet et al., Hohn and Niedrig, Buchner
Dim evezbar As Single    ' Everhart, Danguy and Quivy
Dim donozbar As Single   ' Donovan (for continuum)
Dim donob667zbar As Single  ' Donovan (for backscatter)
Dim donob70zbar As Single  ' Donovan (for backscatter)
Dim donob707zbar As Single  ' Donovan (for backscatter)
Dim donob80zbar As Single  ' Donovan (for backscatter)
Dim donob85zbar As Single  ' Donovan (for backscatter)
Dim donob90zbar As Single  ' Donovan (for backscatter)
Dim bhzbar As Single     ' Bocker and Hehenkamp (for continuum)
Dim logzbar As Single    ' Duncumb log mass z-bar (for continuum)

ReDim atmfrac(1 To MAXCHAN%) As Single
ReDim elafrac(1 To MAXCHAN%) As Single
ReDim zedfrac(1 To MAXCHAN%) As Single
ReDim masfrac(1 To MAXCHAN%) As Single

ReDim elastic(1 To MAXCHAN%) As Single

ReDim masabars(1 To MAXZBAR%) As Single
ReDim masaexps(1 To MAXZBAR%) As Single
ReDim masfracs(1 To MAXZBAR%, 1 To MAXCHAN%) As Single

ReDim zedzbars(1 To MAXZBAR%) As Single
ReDim zedzexps(1 To MAXZBAR%) As Single
ReDim zedfracs(1 To MAXZBAR%, 1 To MAXCHAN%) As Single

ReDim atemp1(1 To MAXCHAN%) As Integer
ReDim atemp2(1 To MAXCHAN%) As Single

' Calculate atomic fraction and z-bar
atmzbar! = 0#
For i% = 1 To sample(1).LastChan%
atmfrac!(i%) = AtPercents!(i%) / 100#
atmzbar! = atmzbar! + atmfrac!(i%) * sample(1).AtomicNums%(i%)
Next i%

' Calculate various mass and zed fractions with range of exponents (0.5 to 1.5)
For j% = 1 To MAXZBAR%
aexp! = 0.5 + 0.1 * (j% - 1)

' Mass fractions
Call StanFormCalculateZbarFrac(Int(1), sample(1).LastChan%, atmfrac!(), sample(1).AtomicNums%(), atemp1%(), sample(1).AtomicWts!(), aexp!, masfrac!(), maszbar!)
If ierror Then Exit Sub

For i% = 1 To sample(1).LastChan%
masfracs!(j%, i%) = masfrac!(i%)
Next i%

' Zed fractions
Call StanFormCalculateZbarFrac(Int(0), sample(1).LastChan%, atmfrac!(), sample(1).AtomicNums%(), sample(1).AtomicNums%(), atemp2!(), aexp!, zedfrac!(), zedzbar!)
If ierror Then Exit Sub

For i% = 1 To sample(1).LastChan%
zedfracs!(j%, i%) = zedfrac!(i%)
Next i%

' Z-bars (exponent 0.7 is index 3)
masaexps!(j%) = aexp!
zedzexps!(j%) = aexp!
masabars!(j%) = maszbar!
zedzbars!(j%) = zedzbar!
Next j%

' Calculate mass fraction and zbar
Call StanFormCalculateZbarFrac(Int(1), sample(1).LastChan%, atmfrac!(), sample(1).AtomicNums%(), atemp1%(), sample(1).AtomicWts!(), CSng(1#), masfrac!(), maszbar!)
If ierror Then Exit Sub

' Calculate zed (electron) fractions and zbar
Call StanFormCalculateZbarFrac(Int(0), sample(1).LastChan%, atmfrac!(), sample(1).AtomicNums%(), sample(1).AtomicNums%(), atemp2!(), CSng(1#), zedfrac!(), zedzbar!)
If ierror Then Exit Sub

' Calculate elastic scattering
Call StanFormCalculateElastic(elastic!(), sample())
If ierror Then Exit Sub

' Calculate elastic fractions and zbar
Call StanFormCalculateZbarFrac(Int(1), sample(1).LastChan%, atmfrac!(), sample(1).AtomicNums%(), atemp1%(), elastic!(), CSng(1#), elafrac!(), elazbar!)
If ierror Then Exit Sub

' Calculate Saldick and Allen z-bar (same as zed or electron fraction)
salzbar! = 0
temp1! = 0#
temp2! = 0#
For i% = 1 To sample(1).LastChan%
temp1! = temp1! + atmfrac!(i%) * sample(1).AtomicNums%(i%) ^ 2
temp2! = temp2! + atmfrac!(i%) * sample(1).AtomicNums%(i%)
Next i%
salzbar! = temp1! / temp2!

' Calculate Joyet et al., Hohn and Niedrig, Buchner zbar
joyzbar! = 0
For i% = 1 To sample(1).LastChan%
joyzbar! = joyzbar! + atmfrac!(i%) * sample(1).AtomicNums%(i%) ^ 2
Next i%
joyzbar! = Sqr(joyzbar!)

' Calculate Everhart, Danguy and Quivy zbar
evezbar! = 0
temp1! = 0#
temp2! = 0#
For i% = 1 To sample(1).LastChan%
temp1! = temp1! + sample(1).ElmPercents!(i%) / 100# * sample(1).AtomicNums%(i%) ^ 2
temp2! = temp2! + sample(1).ElmPercents!(i%) / 100# * sample(1).AtomicNums%(i%)
Next i%
evezbar! = temp1! / temp2!

' Calculate Donovan z-bar (for continuum)
Call StanFormCalculateZbarFrac(Int(0), sample(1).LastChan%, atmfrac!(), sample(1).AtomicNums%(), sample(1).AtomicNums%(), atemp2!(), 0.5!, masfrac!(), donozbar!)
If ierror Then Exit Sub

' Calculate Donovan z-bar (for backscatter)
Call StanFormCalculateZbarFrac(Int(0), sample(1).LastChan%, atmfrac!(), sample(1).AtomicNums%(), sample(1).AtomicNums%(), atemp2!(), 0.667!, masfrac!(), donob667zbar!)
If ierror Then Exit Sub
Call StanFormCalculateZbarFrac(Int(0), sample(1).LastChan%, atmfrac!(), sample(1).AtomicNums%(), sample(1).AtomicNums%(), atemp2!(), 0.7!, masfrac!(), donob70zbar!)
If ierror Then Exit Sub
Call StanFormCalculateZbarFrac(Int(0), sample(1).LastChan%, atmfrac!(), sample(1).AtomicNums%(), sample(1).AtomicNums%(), atemp2!(), 0.707!, masfrac!(), donob707zbar!)
If ierror Then Exit Sub
Call StanFormCalculateZbarFrac(Int(0), sample(1).LastChan%, atmfrac!(), sample(1).AtomicNums%(), sample(1).AtomicNums%(), atemp2!(), 0.8!, masfrac!(), donob80zbar!)
If ierror Then Exit Sub
Call StanFormCalculateZbarFrac(Int(0), sample(1).LastChan%, atmfrac!(), sample(1).AtomicNums%(), sample(1).AtomicNums%(), atemp2!(), 0.85!, masfrac!(), donob85zbar!)
If ierror Then Exit Sub
Call StanFormCalculateZbarFrac(Int(0), sample(1).LastChan%, atmfrac!(), sample(1).AtomicNums%(), sample(1).AtomicNums%(), atemp2!(), 0.9!, masfrac!(), donob90zbar!)
If ierror Then Exit Sub

' Calculate Bocker and Hehenkamp z-bar (for continuum)
Call StanFormCalculateZbarFrac(Int(1), sample(1).LastChan%, atmfrac!(), sample(1).AtomicNums%(), atemp1%(), sample(1).AtomicWts!(), 0.25!, masfrac!(), bhzbar!)
If ierror Then Exit Sub

' Log mass z-bar (Duncumb)
logzbar! = 0
temp1! = 0#
For i% = 1 To sample(1).LastChan%
temp1! = temp1! + sample(1).ElmPercents!(i%) / 100# * Log(sample(1).AtomicNums%(i%))
Next i%
logzbar! = Exp(temp1!)

' Type out data for the sample
n = 0
Do Until False
n% = n% + 1
Call TypeGetRange(Int(2), n%, ii%, jj%, sample())
If ierror Then Exit Sub
If ii% > sample(1).LastChan% Then Exit Do

msg$ = vbCrLf & "ELEMENT:  "
For i% = ii% To jj%
msg$ = msg$ & Format$(sample(1).Elsyup$(i%), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "CONC FRAC:"
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).ElmPercents!(i%) / 100#, f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = vbCrLf & "ZFRAC 1.0:"
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(zedfrac!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "C/Z %DIF: "
For i% = ii% To jj%
temp! = 0#
If sample(1).ElmPercents!(i%) <> 0# Then
temp! = 100# * (zedfrac!(i%) - (sample(1).ElmPercents!(i%) / 100#)) / (sample(1).ElmPercents!(i%) / 100#)
End If
msg$ = msg$ & Format$(Format$(temp!, f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = vbCrLf & "ZFRAC 0.7:"
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(zedfracs!(3, i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "C/Z %DIF: "
For i% = ii% To jj%
temp! = 0#
If sample(1).ElmPercents!(i%) <> 0# Then
temp! = 100# * (zedfracs!(3, i%) - (sample(1).ElmPercents!(i%) / 100#)) / (sample(1).ElmPercents!(i%) / 100#)
End If
msg$ = msg$ & Format$(Format$(temp!, f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = vbCrLf & "ATOM FRAC:"
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(atmfrac!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "ELAS FRAC:"
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(elafrac!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "A/Z Ratio:"
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).AtomicWts!(i%) / sample(1).AtomicNums%(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

Loop

Call IOWriteLog(vbNullString)
Call IOWriteLog("Zbar (Mass fraction) = " & MiscAutoFormat$(maszbar!))
Call IOWriteLog("Zbar (Electron (Z^1.0) fraction) = " & MiscAutoFormat$(zedzbar!))
Call IOWriteLog("Zbar (Mass/Electron (Z^1.0) fraction Zbar % difference) = " & MiscAutoFormat$((maszbar! - zedzbar!) / maszbar! * 100#))

Call IOWriteLog(vbNullString)
Call IOWriteLog("Zbar (Electron (Z^0.7) fraction) = " & MiscAutoFormat$(zedzbars!(3)))
Call IOWriteLog("Zbar (Mass/Electron (Z^0.7) fraction Zbar % difference) = " & MiscAutoFormat$((maszbar! - zedzbars!(3)) / maszbar! * 100#))

Call IOWriteLog(vbNullString)
Call IOWriteLog("Zbar (Elastic fraction) = " & MiscAutoFormat$(elazbar!))
Call IOWriteLog("Zbar (Atomic fraction) = " & MiscAutoFormat$(atmzbar!))

Call IOWriteLog(vbNullString)
Call IOWriteLog("Zbar (Saldick and Allen, for backscatter) = " & MiscAutoFormat$(salzbar!))
Call IOWriteLog("Zbar (Joyet et al.) = " & MiscAutoFormat$(joyzbar!))
Call IOWriteLog("Zbar (Everhart) = " & MiscAutoFormat$(evezbar!))

Call IOWriteLog(vbNullString)
Call IOWriteLog("Zbar (Donovan Z^0.5) = " & MiscAutoFormat$(donozbar!))
Call IOWriteLog("Zbar (Donovan Z^0.667, Yukawa Potential, Z^2/3) = " & MiscAutoFormat$(donob667zbar!))
Call IOWriteLog("Zbar (Donovan Z^0.70) = " & MiscAutoFormat$(donob70zbar!))
Call IOWriteLog("Zbar (Donovan Z^0.707, 1/SQRT(2)) = " & MiscAutoFormat$(donob707zbar!))
Call IOWriteLog("Zbar (Donovan Z^0.80) = " & MiscAutoFormat$(donob80zbar!))
Call IOWriteLog("Zbar (Donovan Z^0.85) = " & MiscAutoFormat$(donob85zbar!))
Call IOWriteLog("Zbar (Donovan Z^0.90) = " & MiscAutoFormat$(donob90zbar!))
Call IOWriteLog("Zbar (Bocker and Hehenkamp for continuum) = " & MiscAutoFormat$(bhzbar!))
Call IOWriteLog("Zbar (Duncumb Log(Mass) for continuum) = " & MiscAutoFormat$(logzbar!))

' Output mass fractions (range of exponent) to file
Call StanFormCalculateSendFracsToFile(Int(0), sample(1).LastChan%, masaexps!(), masfracs!())
If ierror Then Exit Sub

' Output electron fractions (range of exponent) to file
Call StanFormCalculateSendFracsToFile(Int(1), sample(1).LastChan%, zedzexps!(), zedfracs!())
If ierror Then Exit Sub

' Output mass z-bars (range of exponent) to file
Call StanFormCalculateSendBarsToFile(Int(0), masaexps!(), masabars!())
If ierror Then Exit Sub

' Output electron z-bars (range of exponent) to file
Call StanFormCalculateSendBarsToFile(Int(1), zedzexps!(), zedzbars!())
If ierror Then Exit Sub

Exit Sub

' Errors
StanFormCalculateZbarsError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StanFormCalculateZbars"
ierror = True
Exit Sub

End Sub

Sub StanFormCalculateZbarFrac(mode As Integer, lchan As Integer, atomfrac() As Single, atomnums() As Integer, zdata1() As Integer, zdata2() As Single, exponent As Single, fracdata() As Single, zbar As Single)
' Calculate a Z bar based on passed data (integer)
'  mode = 0 for use integer zdata or 1 for use real zdata
'  lchan = number of elements in arrays
'  atomfrac = atomic fractions
'  atomnums = atomic numbers
'  zdata1 = integer data for fraction calculation   ' usually atomic numbers
'  zdata2 = real data for fraction calculation      ' usually atomic weights
'  exponent = exponent for fraction calculation
'  fracdata = returned fraction based on zdata and exponent
'  zbar = returned zbar based on fraction

ierror = False
On Error GoTo StanFormCalculateZbarFracError

Dim i As Integer
Dim sum As Single

' Calculate sum for fraction
sum! = 0#
For i% = 1 To lchan%
If mode% = 0 Then
sum! = sum! + atomfrac!(i%) * zdata1%(i%) ^ exponent!
Else
sum! = sum! + atomfrac!(i%) * zdata2!(i%) ^ exponent!
End If
Next i%
If sum! = 0# Then GoTo StanFormCalculateZbarFracBadSum

' Calculate fractions
For i% = 1 To lchan%
If mode% = 0 Then
fracdata!(i%) = (atomfrac!(i%) * zdata1%(i%) ^ exponent!) / sum!
Else
fracdata!(i%) = (atomfrac!(i%) * zdata2!(i%) ^ exponent!) / sum!
End If
Next i%

' Calculate Z bar
zbar! = 0
For i% = 1 To lchan%
zbar! = zbar! + fracdata!(i%) * atomnums%(i%)
Next i%

Exit Sub

' Errors
StanFormCalculateZbarFracError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StanFormCalculateZbarFrac"
ierror = True
Exit Sub

StanFormCalculateZbarFracBadSum:
Screen.MousePointer = vbDefault
msg$ = "Bad sum in fraction calculation"
MsgBox msg$, vbOKOnly + vbExclamation, "StanFormCalculateZbarFrac"
ierror = True
Exit Sub

End Sub


Sub StanFormCalculateContinuum(sample() As TypeSample)
' Calculate continuum absorption at various wavelengths

ierror = False
On Error GoTo StanFormCalculateContinuumError

Dim i As Integer, j As Integer
Dim ii As Integer, jj As Integer

Dim sum As Single, tt As Single, m1 As Single
Dim aelastic As Single, ainelastic As Single, aphoto As Single, atotal As Single

ReDim elements(1 To MAXCONTINUUM%) As String
ReDim sinthetas(1 To MAXCONTINUUM%) As Single
ReDim characteristicangstroms(1 To MAXCONTINUUM%) As Single
ReDim characteristicenergies(1 To MAXCONTINUUM%) As Single
ReDim continuumangstroms(1 To MAXCONTINUUM%) As Single
ReDim continuumenergies(1 To MAXCONTINUUM%) As Single

ReDim xtalnames(1 To MAXCONTINUUM%) As String
ReDim xtal2ds(1 To MAXCONTINUUM%) As Single

ReDim Z(1 To MAXCONTINUUM%) As Integer
ReDim eC(1 To MAXCONTINUUM%) As Single
ReDim h(1 To MAXCONTINUUM%) As Single

ReDim characteristicabsorptions(1 To MAXCONTINUUM%) As Single
ReDim continuumabsorptions(1 To MAXCONTINUUM%) As Single
ReDim characteristicgenstd(1 To MAXCONTINUUM%) As Single
ReDim continuumgenstd(1 To MAXCONTINUUM%) As Single
ReDim characteristicgensmp(1 To MAXCONTINUUM%) As Single
ReDim continuumgensmp(1 To MAXCONTINUUM%) As Single

ReDim characteristicmac(1 To MAXCHAN%) As Single  ' pure element
ReDim characteristicmacs(1 To MAXCHAN%) As Single ' multi-element
ReDim continuummac(1 To MAXCHAN%) As Single  ' pure element
ReDim continuummacs(1 To MAXCHAN%) As Single ' multi-element

elements$(1) = "Fe"
elements$(2) = "Ti"
elements$(3) = "K"
elements$(4) = "Al"
elements$(5) = "Na"

characteristicangstroms!(1) = 1.9373
characteristicangstroms!(2) = 2.7497
characteristicangstroms!(3) = 3.7424
characteristicangstroms!(4) = 8.3401
characteristicangstroms!(5) = 11.9101

xtalnames$(1) = "LiF"
xtalnames$(2) = "LiF"
xtalnames$(3) = "PET"
xtalnames$(4) = "TAP"
xtalnames$(5) = "PET"

xtal2ds!(1) = 4.0267
xtal2ds!(2) = 4.0267
xtal2ds!(3) = 8.648
xtal2ds!(4) = 25.745
xtal2ds!(5) = 25.745

Z%(1) = 26
Z%(2) = 22
Z%(3) = 19
Z%(4) = 13
Z%(5) = 11

eC!(1) = 7.112
eC!(2) = 4.967
eC!(3) = 3.608
eC!(4) = 1.56
eC!(5) = 1.073

' Calculate at 0.01 sin theta above the emission line wavelengths
' of Fe ka (LiF), Ti ka (LiF), K ka (PET), Al ka (TAP) and Na ka (TAP)
For i% = 1 To MAXCONTINUUM%
characteristicenergies!(i%) = ANGKEV! / characteristicangstroms!(i%)
sinthetas!(i%) = characteristicangstroms!(i%) / xtal2ds!(i%) + 0.01
Next i%

' Calculate continuum angstroms (0.01 sintheta above the characteristic emission lines)
For i% = 1 To MAXCONTINUUM%
continuumangstroms!(i%) = xtal2ds!(i%) * sinthetas!(i%)
continuumenergies!(i%) = ANGKEV! / continuumangstroms!(i%)
Next i%

' Sum weight percents
sum! = 0#
For j% = 1 To sample(1).LastChan%
sum! = sum! + sample(1).ElmPercents!(j%)
Next j%

' Get total characteristic MACs for this composition
For i% = 1 To MAXCONTINUUM%
Call AbsorbGetMAC(Z%(i%), characteristicenergies!(i%), aphoto!, aelastic!, ainelastic!, atotal!)
If ierror Then Exit Sub
characteristicmac!(i%) = atotal!
characteristicmacs!(i%) = 0#
For j% = 1 To sample(1).LastChan%
Call AbsorbGetMAC(sample(1).AtomicNums%(j%), characteristicenergies!(i%), aphoto!, aelastic!, ainelastic!, atotal!)
If ierror Then Exit Sub
characteristicmacs!(i%) = characteristicmacs!(i%) + atotal! * sample(1).ElmPercents!(j%) / sum!
Next j%
Next i%

' Calculate characteristic absorption correction (Heinrich, Anal. Chem.)
tt! = sample(1).takeoff * 3.14159 / 180#
m1! = 1# / Sin(tt!)

' Pure element absorption
For i% = 1 To MAXCONTINUUM%
h!(i%) = 0.0000012 * (sample(1).kilovolts! ^ 1.65 - eC!(i%) ^ 1.65)
characteristicgenstd!(i%) = (1# + h!(i%) * characteristicmac!(i%) * m1!) ^ 2

' Multi-element absorption
h!(i%) = 0.0000012 * (sample(1).kilovolts! ^ 1.65 - eC!(i%) ^ 1.65)
characteristicgensmp!(i%) = (1# + h!(i%) * characteristicmacs!(i%) * m1!) ^ 2

characteristicabsorptions!(i%) = characteristicgensmp!(i%) / characteristicgenstd!(i%)
Next i%

' Get total continuum MACs for this composition
For i% = 1 To MAXCONTINUUM%
Call AbsorbGetMAC(Z%(i%), continuumenergies!(i%), aphoto!, aelastic!, ainelastic!, atotal!)
If ierror Then Exit Sub
continuummac!(i%) = atotal!
continuummacs!(i%) = 0#
For j% = 1 To sample(1).LastChan%
Call AbsorbGetMAC(sample(1).AtomicNums%(j%), continuumenergies!(i%), aphoto!, aelastic!, ainelastic!, atotal!)
If ierror Then Exit Sub
continuummacs!(i%) = continuummacs!(i%) + atotal! * sample(1).ElmPercents!(j%) / sum!
Next j%
Next i%

' Calculate continuum absorption
For i% = 1 To MAXCONTINUUM%

' Pure element
h!(i%) = 0.0000012 * (sample(1).kilovolts! ^ 1.65 - continuumenergies!(i%) ^ 1.65)
continuumgenstd!(i%) = (1# + h!(i%) * continuummac!(i%) * m1!) ^ 2

' Modified depth production and anisotropy correction (Small et al., 1987)
continuumgenstd!(i%) = continuumgenstd!(i%) * 1.15 - 0.15 * 1# / continuumgenstd!(i%)

' Multi-element
continuumgensmp!(i%) = (1# + h!(i%) * continuummacs!(i%) * m1!) ^ 2

' Modified depth production and anisotropy correction (Small et al., 1987)
continuumgensmp!(i%) = continuumgensmp!(i%) * 1.15 - 0.15 * 1# / continuumgensmp!(i%)

continuumabsorptions!(i%) = continuumgensmp!(i%) / continuumgenstd!(i%)
Next i%

ii% = 1
jj% = MAXCONTINUUM%

' Characteristic calculations
Call IOWriteLog(vbNullString)
Call IOWriteLog("Characteristic calculations at" & Str$(MAXCONTINUUM%) & " wavelengths (McMaster MACs, Heinrich Anal. Chem.):")

msg$ = "ELEM: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(elements$(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "ANGS: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(characteristicangstroms!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "KEV : "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(characteristicenergies!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "MAC : "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(characteristicmacs!(i%), e71$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "GSTD: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(characteristicgenstd!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "GSMP: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(characteristicgensmp!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "ABSC: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(characteristicabsorptions!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)


' Continuum calculations
Call IOWriteLog(vbNullString)
Call IOWriteLog("Continuum calculations at" & Str$(MAXCONTINUUM%) & " wavelengths (0.01 sin0 above emission lines):")

msg$ = "ANGS: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(continuumangstroms!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "KEV : "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(continuumenergies!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "MAC : "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(continuummacs!(i%), e71$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "GSTD: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(continuumgenstd!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "GSMP: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(continuumgensmp!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "ABSC: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(continuumabsorptions!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

Exit Sub

' Errors
StanFormCalculateContinuumError:
MsgBox Error$, vbOKOnly + vbCritical, "StanFormCalculateContinuum"
ierror = True
Exit Sub

End Sub

Sub StanFormCalculateContinuumToFile(elements() As String, contstd() As Single, contsmp() As Single, contabs() As Single)
' Output the passed data to files

ierror = False
On Error GoTo StanFormCalculateContinuumToFileError

Dim tfilename1 As String
Dim tfilename2 As String
Dim tfilename3 As String

Dim i As Integer
Dim astring As String, bstring As String
Dim cstring As String, lstring As String

Static initialized As Boolean

tfilename1$ = ApplicationCommonAppData$ & "CONTSTD.DAT"
tfilename2$ = ApplicationCommonAppData$ & "CONTSMP.DAT"
tfilename3$ = ApplicationCommonAppData$ & "CONTABS.DAT"

' Concatanate
astring$ = vbNullString
bstring$ = vbNullString
cstring$ = vbNullString
For i% = 1 To MAXCONTINUUM%
lstring$ = lstring & elements$(i%) & vbTab
astring$ = astring & contstd!(i%) & vbTab
bstring$ = bstring & contsmp!(i%) & vbTab
cstring$ = cstring & contabs!(i%) & vbTab
Next i%

' Delete files if not initialized
If Not initialized Then
If Dir$(tfilename1$) <> vbNullString Then Kill tfilename1$
If Dir$(tfilename2$) <> vbNullString Then Kill tfilename2$
If Dir$(tfilename3$) <> vbNullString Then Kill tfilename3$

' Write labels if first time
Open tfilename1$ For Append As #Temp1FileNumber%
Print #Temp1FileNumber%, lstring$
Close #Temp1FileNumber%
Open tfilename2$ For Append As #Temp1FileNumber%
Print #Temp1FileNumber%, lstring$
Close #Temp1FileNumber%
Open tfilename3$ For Append As #Temp1FileNumber%
Print #Temp1FileNumber%, lstring$
Close #Temp1FileNumber%

initialized = True
End If

' Output continuum absorption for pure elements
Open tfilename1$ For Append As #Temp1FileNumber%
Print #Temp1FileNumber%, astring$
Close #Temp1FileNumber%

' Output continuum absorption for multi-element sample
Open tfilename2$ For Append As #Temp1FileNumber%
Print #Temp1FileNumber%, bstring$
Close #Temp1FileNumber%

' Output absorption relative to pure element
Open tfilename3$ For Append As #Temp1FileNumber%
Print #Temp1FileNumber%, cstring$
Close #Temp1FileNumber%

Exit Sub

' Errors
StanFormCalculateContinuumToFileError:
MsgBox Error$, vbOKOnly + vbCritical, "StanFormCalculateContinuumToFile"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub StanFormCalculateSendBarsToFile(mode As Integer, texps() As Single, tbars() As Single)
' Output the passed data to file
' mode = 0 Send data to MASZBAR.DAT
' mode = 1 Send data to ZEDZBAR.DAT

ierror = False
On Error GoTo StanFormCalculateSendBarsToFileError

Dim i As Integer
Dim tfilename As String
Dim astring As String, bstring As String

Static initialized(0 To 1) As Boolean

If mode% = 0 Then
tfilename$ = ApplicationCommonAppData$ & "MASZBAR.DAT"
Else
tfilename$ = ApplicationCommonAppData$ & "ZEDZBAR.DAT"
End If

' Concatanate
astring$ = vbNullString
bstring$ = vbNullString
For i% = 1 To MAXZBAR%
astring$ = astring$ & "n=" & Trim$(MiscAutoFormat$(texps!(i%))) & vbTab
bstring$ = bstring$ & MiscAutoFormat$(tbars!(i%)) & vbTab
Next i%

' Delete file if not initialized
If Not initialized(mode%) Then
If Dir$(tfilename$) <> vbNullString Then Kill tfilename$

' Output titles first time
Open tfilename$ For Output As #Temp1FileNumber%
Print #Temp1FileNumber%, astring$
Close #Temp1FileNumber%

initialized(mode%) = True
End If

' Output zed zbars
Open tfilename$ For Append As #Temp1FileNumber%
Print #Temp1FileNumber%, bstring$
Close #Temp1FileNumber%

Exit Sub

' Errors
StanFormCalculateSendBarsToFileError:
MsgBox Error$, vbOKOnly + vbCritical, "StanFormCalculateSendBarsToFile"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub StanFormCalculateSendFracsToFile(mode As Integer, nchan As Integer, texps() As Single, tfracs() As Single)
' Output the passed data to file
' mode = 0 Send data to ZEDZFRAC.DAT
' mode = 1 Send data to MASAFRAC.DAT

ierror = False
On Error GoTo StanFormCalculateSendFracsToFileError

Dim i As Integer, j As Integer
Dim tfilename As String
Dim astring As String, bstring As String

Static initialized(0 To 1) As Boolean

If mode% = 0 Then
tfilename$ = ApplicationCommonAppData$ & "MASAFRAC.DAT"
Else
tfilename$ = ApplicationCommonAppData$ & "ZEDZFRAC.DAT"
End If

' Concatanate
astring$ = vbNullString
bstring$ = vbNullString
For j% = 1 To MAXZBAR%
For i% = 1 To nchan%
astring$ = astring$ & "n=" & Trim$(MiscAutoFormat$(texps!(j%))) & vbTab
Next i%
For i% = 1 To nchan%
bstring$ = bstring$ & MiscAutoFormat$(tfracs!(j%, i%)) & vbTab
Next i%
Next j%

' Delete file if not initialized
If Not initialized(mode%) Then
If Dir$(tfilename$) <> vbNullString Then Kill tfilename$

' Output titles first time
Open tfilename$ For Output As #Temp1FileNumber%
Print #Temp1FileNumber%, astring$
Close #Temp1FileNumber%

initialized(mode%) = True
End If

' Output mass or zed zbars
Open tfilename$ For Append As #Temp1FileNumber%
Print #Temp1FileNumber%, bstring$
Close #Temp1FileNumber%

Exit Sub

' Errors
StanFormCalculateSendFracsToFileError:
MsgBox Error$, vbOKOnly + vbCritical, "StanFormCalculateSendFracsToFile"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub


Sub StanFormCalculateElastic(elastic() As Single, sample() As TypeSample)
' Calculate elastic scattering for this sample

ierror = False
On Error GoTo StanFormCalculateElasticError

Dim i As Integer
Dim Alpha As Single, m0c2 As Single, fourpi As Single, k As Single
Dim keV As Single, zed As Single

' Calculate elastic (Rutherford) scattering for each element
k! = 5.21E-21
fourpi! = 4# * 3.14159
m0c2! = 511#

For i% = 1 To sample(1).LastChan%
zed! = sample(1).AtomicNums%(i%)
keV! = sample(1).kilovolts!
Alpha! = 0.0034 * (zed! ^ 0.67) / keV!
elastic!(i%) = k! * (zed! ^ 2) / (keV! ^ 2)
elastic!(i%) = elastic!(i%) * fourpi! / (Alpha! * (1 + Alpha!))
elastic!(i%) = elastic!(i%) * ((keV! + m0c2!) / (keV! + 2 * m0c2!)) ^ 2
Next i%

Exit Sub

' Errors
StanFormCalculateElasticError:
MsgBox Error$, vbOKOnly + vbCritical, "StanFormCalculateElastic"
ierror = True
Exit Sub
End Sub

