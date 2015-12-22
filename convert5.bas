Attribute VB_Name = "CodeCONVERT5"
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit

Const MAXAMPHI% = 12

Dim tfilenumber As Integer, tfilenumber2 As Integer

Sub ConvertMinerals2(method As Integer, percents() As Single, sample() As TypeSample)
' Alternative calculations for mineral calculations
'  method = 6 amphibole calculations
'  method = 7 biotite calculations
'  percents() is oxide percents (dimensioned 1 to sample(1).LastChan%)

ierror = False
On Error GoTo ConvertMinerals2Error

Dim i As Integer
Dim sum As Single

Dim tfilename As String, tfilename2 As String
Dim response As Integer

Static initialized6 As Boolean
Static initialized7 As Boolean

' Load filenames
If method% = 6 Then tfilename$ = UserDataDirectory$ & "\AMPHI.OUT"
'If method% = 6 Then tfilename$ = UserDataDirectory$ & "\AMPHI2.OUT"
'If method% = 7 Then tfilename$ = UserDataDirectory$ & "\BIOTITE.OUT"
If method% = 7 Then tfilename$ = UserDataDirectory$ & "\HALOG.OUT"

If method% = 6 Then tfilename2$ = UserDataDirectory$ & "\AMPHI.DAT"
'If method% = 6 Then tfilename2$ = UserDataDirectory$ & "\AMPHI2.DAT"
'If method% = 7 Then tfilename2$ = UserDataDirectory$ & "\BIOTITE.DAT"
If method% = 7 Then tfilename2$ = UserDataDirectory$ & "\HALOG.DAT"

' Check for existing file if first output
If method% = 6 And Not initialized6 Then
If Not Dir$(tfilename$) = vbNullString Then
msg$ = tfilename$ & " already exists. Do you want to delete the existing output file and start a new AMPHI.OUT output file?"
response% = MsgBox(msg$, vbYesNoCancel + vbQuestion + vbDefaultButton1, "ConvertMinerals2")
If response% = vbYes Then Kill tfilename$
If response% = vbYes And Not Dir$(tfilename2$) = vbNullString Then Kill tfilename2$
If response% = vbCancel Then Exit Sub
End If
initialized6 = True
End If

If method% = 7 And Not initialized7 Then
If Not Dir$(tfilename$) = vbNullString Then
msg$ = tfilename$ & " already exists. Do you want to delete the existing output file and start a new BIOTITE.OUT output file?"
response% = MsgBox(msg$, vbYesNoCancel + vbQuestion + vbDefaultButton1, "ConvertMinerals2")
If response% = vbYes Then Kill tfilename$
If response% = vbYes And Not Dir$(tfilename2$) = vbNullString Then Kill tfilename2$
If response% = vbCancel Then Exit Sub
End If
initialized7 = True
End If

' Check for valid total
sum! = 0#
For i% = 1 To sample(1).LastChan%
sum! = sum! + percents!(i%)
Next i%
If sum! < 80# Then GoTo ConvertMinerals2LowTotal

' Amphibole
If method% = 6 Then
tfilenumber% = FreeFile()
Open tfilename$ For Append As #tfilenumber%
tfilenumber2% = FreeFile()
Open tfilename2$ For Append As #tfilenumber2%
Call ConvertAmphibole(Int(13), percents!(), sample())
'Call ConvertAmphibole2(CSng(0#), percents!(), sample())
Close #tfilenumber%
Close #tfilenumber2%
If ierror Then Exit Sub
End If

' Biotite
If method% = 7 Then
tfilenumber% = FreeFile()
Open tfilename$ For Append As #tfilenumber%
tfilenumber2% = FreeFile()
Open tfilename$ For Append As #tfilenumber2%
'Call ConvertBiotite(percents!(), sample())
Call ConvertHalog(percents!(), sample())
Close #tfilenumber%
Close #tfilenumber2%
If ierror Then Exit Sub
End If

Exit Sub

' Errors
ConvertMinerals2Error:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertMinerals2"
Close #tfilenumber%
Close #tfilenumber2%
ierror = True
Exit Sub

ConvertMinerals2LowTotal:
msg$ = "Total is too low to calculate a structural formula"
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertMinerals2"
Close #tfilenumber%
Close #tfilenumber2%
ierror = True
Exit Sub

End Sub

Sub ConvertAmphibole(INORM As Integer, percents() As Single, sample() As TypeSample)
' Amphibole calculation (COMPUTES CALCIC AMPHIBOLE STRUCTURAL FORMULAS)
'  Originally written in FORTRAN by JAY AGUE, translated to Visual Basic by John Donovan
'  Calls ConvertAmphiboleNORM, ConvertAmphiboleAVER, ConvertAmphiboleSUM, ConvertAmphiboleGETFE and ConvertAmphiboleAUTOAV routines

ierror = False
On Error GoTo ConvertAmphiboleError

Const SSTRING1$ = "-----------------------------------------------------------"
Const SSTRING2$ = " ------ "

Dim j As Integer, ip As Integer
Dim IFE As Integer  ' 1 = Fe2O3 analyzed, 2 = Fe2O3 not analyzed
Dim IDEBUG As Integer, IBIG As Integer, NCODE As Integer

Dim PRESS2 As Single, PRESS23 As Single, ATOPT As Single
Dim MNX As Single, MGMFT As Single, MGB As Single, MGTOFE As Single, MGFET As Single
Dim WTPCT As Single, WTT As Single, RAMGFE  As Single
Dim TETAL As Single, OCTA As Single, OCTAL As Single, ANAM4 As Single, ANA12 As Single
Dim FE2X As Single, FE3X As Single, XALVI As Single, tix As Single, CAX As Single, ANAM4X As Single
Dim FERAT As Single, HALMO As Single, HALMF As Single, HALCL As Single, ASITE As Single
Dim WTPCF As Single, WTPCCL As Single, XFOXOH As Single, HALOG As Single
Dim FELOG  As Single, OC1 As Single, OC2 As Single

Dim d As Double

Dim astring As String, bstring As String

Dim WTPC(1 To MAXAMPHI%) As Single, ATOP(1 To MAXAMPHI%) As Single, PMOL(1 To MAXAMPHI%) As Single
Dim ANIO(1 To MAXAMPHI%) As Single, CAT1(1 To MAXAMPHI%) As Single, ANSFO(1 To MAXAMPHI%) As Single
Dim CAT2(1 To MAXAMPHI%) As Single, CAT3(1 To MAXAMPHI%) As Single, CATF(1 To MAXAMPHI%) As Single
Dim CAT4(1 To MAXAMPHI%) As Single, CAT5(1 To MAXAMPHI%) As Single

Dim esym(1 To MAXAMPHI%) As String
      
' Pre-load small values
For j% = 1 To MAXAMPHI%
    WTPC!(j%) = NotAnalyzedValue!
Next j%

IDEBUG% = 0
IBIG% = 0
      
' Print calculation
Call IOWriteLog(vbCrLf & "Amphibole Formula Calculations (from Jay Ague AMPHI.F code)...")
      
' Assume Fe2O3 not analyzed
IFE% = 2        ' 1 = Fe2O3 analyzed, 2 = Fe2O3 not analyzed
WTPCT! = 0#
WTT! = 0#

' Load oxide percents
For j% = 1 To sample(1).LastChan%
    ip% = IPOS1%(MAXELM%, sample(1).Elsyms$(j%), Symlo$())
    If ip% <> 0 Then
        If ip% = 14 Then WTPC!(1) = percents!(j%)   ' SiO2
        If ip% = 22 Then WTPC!(2) = percents!(j%)   ' TiO2
        If ip% = 13 Then WTPC!(3) = percents!(j%)   ' Al2O3
        If ip% = 26 Then
            If IFE% = 1 Then
            WTPC!(4) = 0#              ' Fe2O3 (just zero out for now)
            WTPC!(5) = percents!(j%)   ' FeO
            Else
            WTPC!(4) = 0#              ' Fe2O3
            WTPC!(5) = percents!(j%)   ' FeO
            End If
        End If
        If ip% = 12 Then WTPC!(6) = percents!(j%)   ' MgO
        If ip% = 25 Then WTPC!(7) = percents!(j%)   ' MnO
        If ip% = 20 Then WTPC!(8) = percents!(j%)   ' CaO
        If ip% = 11 Then WTPC!(9) = percents!(j%)   ' Na2O
        If ip% = 19 Then WTPC!(10) = percents!(j%)  ' K2O
        If ip% = 9 Then WTPC!(11) = percents!(j%)  ' F
        If ip% = 17 Then WTPC!(12) = percents!(j%)   ' Cl
    End If
Next j%

' Sum total
WTT! = 0#
For j% = 1 To MAXAMPHI%
    WTT! = WTT! + WTPC!(j%)
Next j%

' Begin calculations
For j% = 1 To MAXAMPHI%
    PMOL!(1) = WTPC!(1) / 60.09: esym$(1) = "Si"
    PMOL!(2) = WTPC!(2) / 79.9: esym$(2) = "Ti"
    PMOL!(3) = WTPC!(3) / 101.94: esym$(3) = "Al"
    PMOL!(4) = WTPC!(4) / 159.7: esym$(4) = "Fe+3"
    PMOL!(5) = WTPC!(5) / 71.85: esym$(5) = "Fe+2"
    PMOL!(6) = WTPC!(6) / 40.32: esym$(6) = "Mg"
    PMOL!(7) = WTPC!(7) / 70.94: esym$(7) = "Mn"
    PMOL!(8) = WTPC!(8) / 56.08: esym$(8) = "Ca"
    PMOL!(9) = WTPC!(9) / 61.982: esym$(9) = "Na"
    PMOL!(10) = WTPC!(10) / 94.2: esym$(10) = "K"
    PMOL!(11) = WTPC!(11) / 19#: esym$(11) = "F"
    PMOL!(12) = WTPC!(12) / 35.457: esym$(12) = "Cl"
    If j% <= 2 Then GoTo 102
    If j% = 3 Or j% = 4 Then GoTo 103
    If j% > 4 Then GoTo 104

' 2 OXYGENS (equivalence)
102:
    ATOP!(j%) = PMOL!(j%) * 2#
    GoTo 105

' 3 OXYGENS
103:
    ATOP!(j%) = PMOL!(j%) * 3#
    GoTo 105

' 1 OXYGEN, F, CL
104:
    ATOP!(j%) = PMOL!(j%) * 1#

' Next element
105:
    WTPCT = WTPCT + WTPC!(j%)
Next j%
      
WTPCF = WTPC(11) * 0.4211
WTPCCL = WTPC(12) * 0.2256
WTPCT = WTPCT - WTPCF - WTPCCL
ATOPT = 0#
      
For j% = 1 To MAXAMPHI% - 2
    ATOPT = ATOPT + ATOP!(j%)
Next j%

' USE 46 NEGATIVE CHARGES
d# = 23# / ATOPT!
For j% = 1 To MAXAMPHI%
    ANIO!(j%) = d * ATOP!(j%)
Next j%

For j% = 1 To MAXAMPHI%
      If j% <= 2 Then ANSFO!(j%) = ANIO!(j%) / 2#
      If j% = 3 Or j% = 4 Then ANSFO!(j%) = ANIO!(j%) * 0.6666
      If j% > 4 And j% <= 8 Then ANSFO!(j%) = ANIO!(j%)
      If j% = 9 Or j% = 10 Then ANSFO!(j%) = ANIO!(j%) * 2#
      If j% > 10 Then ANSFO!(j%) = ANIO!(j%)
Next j%

' Call routine NORM to perform the structural formula calculations
If IFE% <> 1 Then
Call ConvertAmphiboleNorm(ANSFO!(), CAT1!(), CAT2!(), CAT3!(), CAT4!(), CAT5!())
If ierror Then Exit Sub
End If

' END NORMALIZATION PROCEDURE (FORMULA IS STORED IN ARRAY CAT1-5())
If DebugMode Then
Call IOWriteLog(vbCrLf & "Entered Amphibole Analysis:")
For j% = 1 To MAXAMPHI%
    If j% = 1 Then astring$ = Format$("SiO2", a80$)
    If j% = 2 Then astring$ = Format$("TiO2", a80$)
    If j% = 3 Then astring$ = Format$("Al2O3", a80$)
    If j% = 4 Then astring$ = Format$("Fe2O3", a80$)
    If j% = 5 Then astring$ = Format$("FeO", a80$)
    If j% = 6 Then astring$ = Format$("MgO", a80$)
    If j% = 7 Then astring$ = Format$("MnO", a80$)
    If j% = 8 Then astring$ = Format$("CaO", a80$)
    If j% = 9 Then astring$ = Format$("Na2O", a80$)
    If j% = 10 Then astring$ = Format$("K2O", a80$)
    If j% = 11 Then astring$ = Format$("F", a80$)
    If j% = 12 Then astring$ = Format$("Cl", a80$)
    If j% <> 4 Then bstring$ = Format$(Format$(WTPC!(j%), f83$), a80)
    If j% = 4 And IFE% = 1 Then bstring$ = Format$(Format$(WTPC!(j%), f83$), a80)
    If j% = 4 And IFE% = 2 Then bstring$ = Format$(Format$(" ---- "), a80)
Call IOWriteLog(bstring$ & a4x$ & astring$)
Next j%
      
Call IOWriteLog(SSTRING2$)
      
astring$ = Format$("TOTAL", a80$)
bstring$ = Format$(Format$(WTT!, f83$), a80)
Call IOWriteLog(bstring$ & a4x$ & astring$)
      
astring$ = Format$("TOTAL-OXYGEN EQUIV. OF F,CL")
bstring$ = Format$(Format$(WTPCT!, f83$), a80)
Call IOWriteLog(bstring$ & a4x$ & astring$)
Call IOWriteLog(SSTRING1$)
End If
      
' Do Fe2O3/FeO calculations (Fe2O3 analyzed)
If IFE% = 1 Then
    NCODE = 0
    For j% = 1 To MAXAMPHI%
        CATF!(j%) = ANSFO!(j%)
    Next j%
End If
      
' Do Fe2O3/FeO calculations (Fe2O3 not analyzed), print out candidate normalizations
If IFE% <> 1 Then
astring$ = a6x$ & Format$("ALL FE2", a80$) & Format$("NORM 1", a80$) & Format$("NORM 2", a80$) & Format$("NORM 3", a80$) & Format$("NORM 4", a80$) & Format$("NORM 5", a80$)
bstring$ = a6x$ & Format$(SSTRING2$, a80$) & Format$(SSTRING2$, a80$) & Format$(SSTRING2$, a80$) & Format$(SSTRING2$, a80$) & Format$(SSTRING2$, a80$) & Format$(SSTRING2$, a80$)
Call IOWriteLog(astring$)
Call IOWriteLog(bstring$)

For j% = 1 To MAXAMPHI%
    If j% = 1 Then astring$ = Format$("Si", a60$)
    If j% = 2 Then astring$ = Format$("Ti", a60$)
    If j% = 3 Then astring$ = Format$("Al", a60$)
    If j% = 4 Then astring$ = Format$("Fe3+", a60$)
    If j% = 5 Then astring$ = Format$("Fe2+", a60$)
    If j% = 6 Then astring$ = Format$("Mg", a60$)
    If j% = 7 Then astring$ = Format$("Mn", a60$)
    If j% = 8 Then astring$ = Format$("Ca", a60$)
    If j% = 9 Then astring$ = Format$("Na", a60$)
    If j% = 10 Then astring$ = Format$("K", a60$)
    If j% = 11 Then astring$ = Format$("F", a60$)
    If j% = 12 Then astring$ = Format$("Cl", a60$)
    bstring$ = Format$(Format$(ANSFO!(j%), f84$), a80$) & Format$(Format$(CAT1!(j%), f84$), a80$)
    bstring$ = bstring$ & Format$(Format$(CAT2!(j%), f84$), a80$) & Format$(Format$(CAT3!(j%), f84$), a80$)
    bstring$ = bstring$ & Format$(Format$(CAT4!(j%), f84$), a80$) & Format$(Format$(CAT5!(j%), f84$), a80$)
Call IOWriteLog$(astring$ & bstring$)
Next j%
      
Call IOWriteLog(vbNullString)
astring$ = a6x$ & "NORM 1: TOTAL-(NA+K)=15   " & "NORM 2: TOTAL-(NA+CA+K)=13"
Call IOWriteLog$(astring$)
astring$ = a6x$ & "NORM 3: TOTAL-K=15        " & "NORM 4: SI+AL=8.0"
Call IOWriteLog$(astring$)
astring$ = a6x$ & "NORM 5: TOTAL=15.8"
Call IOWriteLog$(astring$)
Call IOWriteLog$(SSTRING1$)

' Call averaging routine to determine final structure
Call ConvertAmphiboleAver(ANSFO!(), CAT1!(), CAT2!(), CAT3!(), CAT4!(), CAT5!(), CATF!(), INORM%)
If ierror Then Exit Sub
End If

' Round values in structural formula (array CATF) to 4 decimal places. If Mg = 0, set to 0.0001 to prevent errors.
For j% = 1 To MAXAMPHI%
    CATF!(j%) = MiscSetRounding2!(CATF!(j%), Int(4))
Next j%
If CATF!(6) = 0# Then CATF!(6) = 0.0001

' Compute mole fractions, molar ratios and pressure of crystallazation using
'  Schmidt (1992) calibration of the Hammarstrom and Zen (1986) barometer

' TETRAHEDRAL AL
      If CATF!(1) < 8 Then
        TETAL = 8# - CATF!(1)
      Else
        TETAL = 0#
      End If

' OCTAHEDRAL AL
      OCTAL = CATF(3) - TETAL
      If (OCTAL < 0#) Then TETAL = CATF(3)
      If (OCTAL < 0#) Then OCTAL = 0#

' NA ON M4 SITE
      ANAM4 = 7# - (CATF(8) + CATF(4) + CATF(5) + CATF(6) + CATF(7) + OCTAL + CATF(2))
      If (ANAM4 < 0#) Then Call IOWriteLog(a6x$ & "***OCTAHEDRAL NA IS NEGATIVE***")

' 12-FOLD NA
      If (ANAM4 < 0#) Then ANA12 = CATF(9)
      If (ANAM4 >= 0#) Then ANA12 = CATF(9) - 1# * ANAM4
      If ((ANAM4 >= 0#) And (ANAM4 > CATF(9))) Then ANA12 = 0#
      If (ANAM4 < 0#) Then Call IOWriteLog(a6x$ & "***ALL NA ASSIGNED TO 12-FOLD SITE***")
      If (ANAM4 < 0#) Then ANAM4 = 0#
      If (ANAM4 > CATF(9)) Then ANAM4 = CATF(9)

' TOTAL OCTAHEDRAL ATOMS
      OCTA = ANAM4 + OCTAL + CATF(8) + CATF(4) + CATF(5) + CATF(6) + CATF(2) + CATF(7)

' X - ALVI
      XALVI = OCTAL / OCTA

' X-FE2+
      FE2X = CATF(5) / OCTA

' X-FE3+
      FE3X = CATF(4) / OCTA

' X - MG
      MGMFT = CATF(6) / OCTA

' X - TI
      tix = CATF(2) / OCTA

' X - MN
      MNX = CATF(7) / OCTA

' X - CA
      CAX = CATF(8) / OCTA

'X - NAM4
      ANAM4X = ANAM4 / OCTA

' A-SITE OCCUPANCY
      ASITE = ANA12 + CATF(10)

' X - F
      HALMF = CATF(11) / 2#
      If (HALMF = 0#) Then HALMF = 0.00001

' X - CL
      HALCL = CATF(12) / 2#
      If (HALCL = 0#) Then HALCL = 0.00001

' X - OH
      HALMO = 1# - HALMF - HALCL
      If (HALMO <= 0#) Then HALMO = 0.00001

' COMPUTE LOG X-F/X-OH
      If HALMF / HALMO > 0# Then XFOXOH = MiscConvertLog10#(CDbl(HALMF / HALMO))

' COMPUTE LOG X-F/X-CL
      If HALMF / HALCL > 0# Then HALOG = MiscConvertLog10#(CDbl(HALMF / HALCL))

' COMPUTE LOG MG/FE2+
      If MGMFT / FE2X > 0# Then RAMGFE = MiscConvertLog10#(CDbl(MGMFT / FE2X))

' COMPUTE LOG (MG/(FE2+ + FE3+))
      If MGMFT / (FE2X + FE3X) > 0# Then MGFET = MiscConvertLog10#(CDbl(MGMFT / (FE2X + FE3X)))

' COMPUTE LOG FE2+/FE3+
      If (FE3X = 0#) Then FE3X = 0.0001
      If FE2X / FE3X > 0# Then FELOG = MiscConvertLog10#(CDbl(FE2X / FE3X))

' COMPUTE MG/(MG+FE2+)
      MGB = CATF(6) / (CATF(6) + CATF(5))

' COMPUTE ALVI+FE3+ + 2TI+ASITE
      OC1 = OCTAL + CATF(4) + 2# * CATF(2) + ASITE

' COMPUTE ALVI+FE3+ +2TI
      OC2 = OCTAL + CATF(4) + 2# * CATF(2)

' COMPUTE FE2+/(FE2+ + FE3+)
      FERAT = CATF(5) / (CATF(5) + CATF(4))

' COMPUTE MG/(MG+FE TOTAL)
      MGTOFE = CATF(6) / (CATF(6) + CATF(4) + CATF(5))

' Compute pressure, both with total Al (all Fe2+) and total Al (Fe2+ - Fe3+)
PRESS2! = 5.03 * ANSFO!(3) - 3.92
PRESS23! = 5.03 * CATF!(3) - 3.92
If PRESS2! < 0# Or IFE = 1 Then PRESS2! = 0#
If PRESS23! < 0# Then PRESS23! = 0#

Call IOWriteLog(a6x$ & "STRUCTURAL FORMULA:")
Call IOWriteLog(a8x$ & "SI   " & Format$(Format$(CATF(1), f84$), a80$))
Call IOWriteLog(a8x$ & "TI   " & Format$(Format$(CATF(2), f84$), a80$))
Call IOWriteLog(a8x$ & "AL IV" & Format$(Format$(TETAL!, f84$), a80$) & a8x$ & "AL VI" & Format$(Format$(OCTAL!, f84$), a80$))
Call IOWriteLog(a8x$ & "FE3+ " & Format$(Format$(CATF(4), f84$), a80$))
Call IOWriteLog(a8x$ & "FE2+ " & Format$(Format$(CATF(5), f84$), a80$))
Call IOWriteLog(a8x$ & "MG   " & Format$(Format$(CATF(6), f84$), a80$))
Call IOWriteLog(a8x$ & "MN   " & Format$(Format$(CATF(7), f84$), a80$))
Call IOWriteLog(a8x$ & "CA   " & Format$(Format$(CATF(8), f84$), a80$))
Call IOWriteLog(a8x$ & "NA A " & Format$(Format$(ANAM4!, f84$), a80$) & a8x$ & "NA B " & Format$(Format$(ANA12!, f84$), a80$))
Call IOWriteLog(a8x$ & "K    " & Format$(Format$(CATF(10), f84$), a80$))
Call IOWriteLog(a8x$ & "F    " & Format$(Format$(CATF(11), f84$), a80$))
Call IOWriteLog(a8x$ & "CL   " & Format$(Format$(CATF(12), f84$), a80$))
Call IOWriteLog(a8x$ & "OH   " & Format$(Format$(2# - (CATF(11) + CATF(12)), f84$), a80$))

Call IOWriteLog(SSTRING1$)
Call IOWriteLog("MOLE FRACTIONS AND LOGARITHMS OF ATOMIC RATIOS:")

Call IOWriteLog(a6x$ & "X-FE2+= " & Format$(Format$(FE2X!, f83$), a80$) & a6x$ & "X-MG=   " & Format$(Format$(MGMFT!, f83$), a80$))
Call IOWriteLog(a6x$ & "X-FE3+= " & Format$(Format$(FE3X!, f83$), a80$) & a6x$ & "X-ALVI= " & Format$(Format$(XALVI, f83$), a80$))
Call IOWriteLog(a6x$ & "X-MN=   " & Format$(Format$(MNX!, f84$), a80$) & a6x$ & "X-TI=   " & Format$(Format$(tix!, f84$), a80$))
Call IOWriteLog(a6x$ & "X-CA=   " & Format$(Format$(CAX!, f83$), a80$) & a6x$ & "X-NAM4= " & Format$(Format$(ANAM4X!, f84$), a80$) & vbCrLf)

Call IOWriteLog(a6x$ & "MG / (MG + FE2+) =  " & Format$(Format$(MGB!, f83$), a80$))
Call IOWriteLog(a6x$ & "FE2+/(FE2+ + FE3+)= " & Format$(Format$(FERAT!, f83$), a80$))
Call IOWriteLog(a6x$ & "MG/(MG+FE2+ + FE3+)=" & Format$(Format$(MGTOFE!, f83$), a80$) & vbCrLf)

Call IOWriteLog(a6x$ & "X-OH=                     " & Format$(Format$(HALMO!, f83$), a80$) & a4x$ & "X-F=               " & Format$(Format$(HALMF!, f83$), a80$) & a4x & "X-CL=         " & Format$(Format$(HALCL!, f83$), a80$))
Call IOWriteLog(a6x$ & "LOG(X-MG/X-FE2+)=         " & Format$(Format$(RAMGFE!, f83$), a80$) & a4x$ & "LOG(X-F/X-CL)=     " & Format$(Format$(HALOG!, f83$), a80$) & a4x$ & "LOG(X-F/X-OH)=" & Format$(Format$(XFOXOH!, f83$), a80$))
Call IOWriteLog(a6x$ & "LOG(X-MG/(X-FE2+ + FE3+))=" & Format$(Format$(MGFET!, f83$), a80$) & a4x$ & "LOG(X-FE2+/X-FE3+)=" & Format$(Format$(FELOG!, f83$), a80$) & vbCrLf)
Call IOWriteLog(a6x$ & "A-SITE=                   " & Format$(Format$(ASITE!, f83$), a80$) & a4x$ & "TOTAL VI=          " & Format$(Format$(OCTA!, f83$), a80$))
Call IOWriteLog(a6x$ & "ALVI+2TI+A-SITE+FE3+=     " & Format$(Format$(OC1!, f83$), a80$) & a4x$ & "ALVI+2TI+FE3+=     " & Format$(Format$(OC2!, f83$), a80$))
      
Call IOWriteLog(vbNullString)
Call IOWriteLog(a6x$ & "Schmidt (1992) Pressure (All FE2+): " & Format$(Format$(PRESS2!, f42$), a80$) & " KBar, " & a4x$ & "(FE2+ -FE3+): " & Format$(Format$(PRESS23!, f42$), a80$) & " KBar")
      
' Output to file (AMPHI.OUT)
astring$ = vbCrLf & "Sample " & VbDquote$ & sample(1).number% & VbDquote$ & vbTab & VbDquote$ & sample(1).Name$ & VbDquote$
Print #tfilenumber%, astring$

astring$ = vbNullString
For j% = 1 To MAXAMPHI%
    If j% = 3 Then
    astring$ = astring$ + MiscAutoFormat$(WTPC!(j%)) & vbTab$ & MiscAutoFormat$(ANSFO!(j%)) & vbTab$ & MiscAutoFormat$(CATF!(j%)) & vbTab$ & MiscAutoFormat$(TETAL!) & vbTab$ & MiscAutoFormat$(OCTAL!) & vbTab & esym$(j%) & vbTab$ & MiscAutoFormat$(PRESS2!) & vbTab$ & MiscAutoFormat$(PRESS23!) & vbCrLf
    ElseIf j = 9 Then
    astring$ = astring$ + MiscAutoFormat$(WTPC!(j%)) & vbTab$ & MiscAutoFormat$(ANSFO!(j%)) & vbTab$ & MiscAutoFormat$(CATF!(j%)) & vbTab$ & MiscAutoFormat$(ANAM4!) & vbTab$ & MiscAutoFormat$(ANA12!) & vbTab & esym$(j%) & vbCrLf
    ElseIf j = 12 Then
    astring$ = astring$ + MiscAutoFormat$(WTPC!(j%)) & vbTab$ & MiscAutoFormat$(ANSFO!(j%)) & vbTab$ & MiscAutoFormat$(CATF!(j%)) & vbTab$ & MiscAutoFormat$(2# - (CATF!(11) + CATF!(12))) & vbTab & esym$(j%) & vbCrLf
    Else
    astring$ = astring$ + MiscAutoFormat$(WTPC!(j%)) & vbTab$ & MiscAutoFormat$(ANSFO!(j%)) & vbTab$ & MiscAutoFormat$(CATF!(j%)) & vbTab & esym$(j%) & vbCrLf
    End If
Next j%
Print #tfilenumber%, astring$
    
astring$ = vbNullString
astring$ = astring$ & MiscAutoFormat$(HALOG!) & vbTab$ & MiscAutoFormat$(RAMGFE!) & vbTab$ & MiscAutoFormat$(MGB!) & vbTab$
astring$ = astring$ & MiscAutoFormat$(FE2X!) & vbTab$ & MiscAutoFormat$(tix!) & vbTab$ & MiscAutoFormat$(XALVI!) & vbTab$
astring$ = astring$ & MiscAutoFormat$(MNX!) & vbTab$ & MiscAutoFormat$(XFOXOH!) & vbTab$ & MiscAutoFormat$(OC1!) & vbTab$
astring$ = astring$ & MiscAutoFormat$(OC2!) & vbTab$ & MiscAutoFormat$(CATF(3)) & vbTab$ & MiscAutoFormat$(ASITE!) & vbTab$
astring$ = astring$ & MiscAutoFormat$(FERAT!) & vbTab$ & MiscAutoFormat$(CATF(1)) & vbCrLf
astring$ = astring$ & MiscAutoFormat$(ANAM4!) & vbTab$ & MiscAutoFormat$(ANA12!) & vbTab$ & MiscAutoFormat$(OCTA!) & vbTab$
astring$ = astring$ & MiscAutoFormat$(CATF!(8)) & vbTab & MiscAutoFormatI$(NCODE%)
Print #tfilenumber%, astring$

' Output to file (AMPHI.DAT)
astring$ = vbCrLf & "Sample " & vbTab & VbDquote$ & sample(1).number% & VbDquote$ & vbTab & VbDquote$ & sample(1).Name$ & VbDquote$
Print #tfilenumber2%, astring$

' Output oxide labels
astring$ = vbNullString
For j% = 1 To MAXAMPHI%
    If j% = 1 Then astring$ = astring$ & "SiO2" & vbTab
    If j% = 2 Then astring$ = astring$ & "TiO2" & vbTab
    If j% = 3 Then
        astring$ = astring$ & "Al2O3" & vbTab
        astring$ = astring$ & "-----" & vbTab
    End If
    If j% = 4 Then astring$ = astring$ & "Fe2O3" & vbTab
    If j% = 5 Then astring$ = astring$ & "FeO" & vbTab
    If j% = 6 Then astring$ = astring$ & "MgO" & vbTab
    If j% = 7 Then astring$ = astring$ & "MnO" & vbTab
    If j% = 8 Then astring$ = astring$ & "CaO" & vbTab
    If j% = 9 Then
        astring$ = astring$ & "Na2O" & vbTab
        astring$ = astring$ & "----" & vbTab
    End If
    If j% = 10 Then astring$ = astring$ & "K2O" & vbTab
    If j% = 11 Then astring$ = astring$ & "F" & vbTab
    If j% = 12 Then astring$ = astring$ & "Cl" & vbTab
Next j%
astring$ = astring$ & "OH" & vbTab
Print #tfilenumber2%, astring$

' Output oxide wt%
astring$ = vbNullString
For j% = 1 To MAXAMPHI%
    If j% = 1 Then astring$ = astring$ & Format$(WTPC!(j%)) & vbTab
    If j% = 2 Then astring$ = astring$ & Format$(WTPC!(j%)) & vbTab
    If j% = 3 Then
        astring$ = astring$ & Format$(WTPC!(j%)) & vbTab
        astring$ = astring$ & Format$(0#) & vbTab
    End If
    If j% = 4 Then astring$ = astring$ & Format$(WTPC!(j%)) & vbTab
    If j% = 5 Then astring$ = astring$ & Format$(WTPC!(j%)) & vbTab
    If j% = 6 Then astring$ = astring$ & Format$(WTPC!(j%)) & vbTab
    If j% = 7 Then astring$ = astring$ & Format$(WTPC!(j%)) & vbTab
    If j% = 8 Then astring$ = astring$ & Format$(WTPC!(j%)) & vbTab
    If j% = 9 Then
        astring$ = astring$ & Format$(WTPC!(j%)) & vbTab
        astring$ = astring$ & Format$(0#) & vbTab
    End If
    If j% = 10 Then astring$ = astring$ & Format$(WTPC!(j%)) & vbTab
    If j% = 11 Then astring$ = astring$ & Format$(WTPC!(j%)) & vbTab
    If j% = 12 Then astring$ = astring$ & Format$(WTPC!(j%)) & vbTab
Next j%
astring$ = astring$ & Format$(2# - (CATF(11) + CATF(12))) & vbTab
Print #tfilenumber2%, astring$

' Output structural formula labels
astring$ = vbNullString
For j% = 1 To MAXAMPHI%
    If j% = 1 Then astring$ = astring$ & VbDquote$ & "SI" & VbDquote$ & vbTab
    If j% = 2 Then astring$ = astring$ & VbDquote$ & "TI" & VbDquote$ & vbTab
    If j% = 3 Then
        astring$ = astring$ & "AL IV" & vbTab
        astring$ = astring$ & "AL VI" & vbTab
    End If
    If j% = 4 Then astring$ = astring$ & VbDquote$ & "FE 3+" & VbDquote$ & vbTab
    If j% = 5 Then astring$ = astring$ & VbDquote$ & "FE 2+" & VbDquote$ & vbTab
    If j% = 6 Then astring$ = astring$ & VbDquote$ & "MG" & VbDquote$ & vbTab
    If j% = 7 Then astring$ = astring$ & VbDquote$ & "Mn" & VbDquote$ & vbTab
    If j% = 8 Then astring$ = astring$ & VbDquote$ & "CA" & VbDquote$ & vbTab
    If j% = 9 Then
        astring$ = astring$ & VbDquote$ & "NA A" & VbDquote$ & vbTab
        astring$ = astring$ & VbDquote$ & "NA B" & VbDquote$ & vbTab
    End If
    If j% = 10 Then astring$ = astring$ & VbDquote$ & "K" & VbDquote$ & vbTab
    If j% = 11 Then astring$ = astring$ & VbDquote$ & "F" & VbDquote$ & vbTab
    If j% = 12 Then astring$ = astring$ & VbDquote$ & "CL" & VbDquote$ & vbTab
Next j%
astring$ = astring$ & VbDquote$ & "OH" & VbDquote$ & vbTab
Print #tfilenumber2%, astring$
    
' Output structural formulas
astring$ = vbNullString
For j% = 1 To MAXAMPHI%
    If j% = 1 Then astring$ = astring$ & Format$(CATF(1)) & vbTab
    If j% = 2 Then astring$ = astring$ & Format$(CATF(2)) & vbTab
    If j% = 3 Then
        astring$ = astring$ & Format$(TETAL!) & vbTab
        astring$ = astring$ & Format$(OCTAL!) & vbTab
    End If
    If j% = 4 Then astring$ = astring$ & Format$(CATF(4)) & vbTab
    If j% = 5 Then astring$ = astring$ & Format$(CATF(5)) & vbTab
    If j% = 6 Then astring$ = astring$ & Format$(CATF(6)) & vbTab
    If j% = 7 Then astring$ = astring$ & Format$(CATF(7)) & vbTab
    If j% = 8 Then astring$ = astring$ & Format$(CATF(8)) & vbTab
    If j% = 9 Then
        astring$ = astring$ & Format$(ANAM4!) & vbTab
        astring$ = astring$ & Format$(ANA12!) & vbTab
    End If
    If j% = 10 Then astring$ = astring$ & Format$(CATF(10)) & vbTab
    If j% = 11 Then astring$ = astring$ & Format$(CATF(11)) & vbTab
    If j% = 12 Then astring$ = astring$ & Format$(CATF(12)) & vbTab
Next j%
astring$ = astring$ & Format$(2# - (CATF(11) + CATF(12))) & vbTab
Print #tfilenumber2%, astring$
    
Exit Sub

' Errors
ConvertAmphiboleError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertAmphibole"
Close #tfilenumber%
Close #tfilenumber2%
ierror = True
Exit Sub

End Sub

Sub ConvertAmphibole2(FIXRAT As Single, percents() As Single, sample() As TypeSample)
' Amphibole calculation (COMPUTES CALCIC AMPHIBOLE STRUCTURAL FORMULAS) with Fe3+/Fe total modification
'  Originally written in FORTRAN by JAY AGUE, translated to Visual Basic by John Donovan

ierror = False
On Error GoTo ConvertAmphibole2Error

Const SSTRING1$ = "-----------------------------------------------------------"
Const SSTRING2$ = " ------ "

Dim j As Integer, ip As Integer
Dim IFE As Integer  ' 1 = Fe2O3 analyzed, 2 = Fe2O3 not analyzed
Dim IDEBUG As Integer, IBIG As Integer, NCODE As Integer

Dim MNX As Single, MGMFT As Single, MGB As Single, MGTOFE As Single, MGFET As Single
Dim PRESS2 As Single, PRESS23 As Single, ATOPT As Single
Dim WTT As Single, WTPCT As Single, WTPCF As Single, WTPCCL As Single
Dim TETAL As Single, OCTA As Single, OCTAL As Single, ANAM4 As Single, ANA12 As Single
Dim FE2X As Single, FE3X As Single, XALVI As Single, tix As Single, CAX As Single, ANAM4X As Single
Dim FERAT As Single, HALMO As Single, HALMF As Single, HALCL As Single, ASITE As Single
Dim XFOXOH As Single, HALOG As Single, RAMGFE As Single, FELOG As Single, OC1 As Single, OC2 As Single

Dim d As Double

Dim WTPC(1 To MAXAMPHI%) As Single, ATOP(1 To MAXAMPHI%) As Single, PMOL(1 To MAXAMPHI%) As Single
Dim ANIO(1 To MAXAMPHI%) As Single, ANSFO(1 To MAXAMPHI%) As Single, CATF(1 To MAXAMPHI%) As Single

Dim esym(1 To MAXAMPHI%) As String

Dim astring As String, bstring As String

' Pre-load small values
For j% = 1 To MAXAMPHI%
    WTPC!(j%) = NotAnalyzedValue!
Next j%

' Print calculation
Call IOWriteLog(vbCrLf & "Amphibole Formula Calculations (from Jay Ague AMPHI2.F code)...")
      
' Assume Fe3+/Fe total ratio is 0.3
If FIXRAT! = 0# Then FIXRAT! = 0.3
IDEBUG% = 0
IBIG% = 0
IFE% = 1    ' Fe2O3 analyzed
'IFE% = 2    ' Fe2O3 not analyzed (not usable)

' Load oxide percents
For j% = 1 To sample(1).LastChan%
    ip% = IPOS1%(MAXELM%, sample(1).Elsyms$(j%), Symlo$())
    If ip% <> 0 Then
        If ip% = 14 Then WTPC!(1) = percents!(j%)   ' SiO2
        If ip% = 22 Then WTPC!(2) = percents!(j%)   ' TiO2
        If ip% = 13 Then WTPC!(3) = percents!(j%)   ' Al2O3
        If ip% = 26 And IFE% = 1 Then
            WTPC!(4) = 0#              ' Fe2O3 (just zero out for now)
        End If
        If ip% = 26 Then WTPC!(5) = percents!(j%)   ' FeO (Total)
        If ip% = 12 Then WTPC!(6) = percents!(j%)   ' MgO
        If ip% = 25 Then WTPC!(7) = percents!(j%)   ' MnO
        If ip% = 20 Then WTPC!(8) = percents!(j%)   ' CaO
        If ip% = 11 Then WTPC!(9) = percents!(j%)   ' Na2O
        If ip% = 19 Then WTPC!(10) = percents!(j%)  ' K2O
        If ip% = 9 Then WTPC!(11) = percents!(j%)  ' F
        If ip% = 17 Then WTPC!(12) = percents!(j%)   ' Cl
    End If
Next j%

' Sum total
WTT! = 0#
For j% = 1 To MAXAMPHI%
    WTT! = WTT! + WTPC!(j%)
Next j%

' Begin calculations
For j% = 1 To MAXAMPHI%
    PMOL!(1) = WTPC!(1) / 60.09: esym$(1) = "Si"
    PMOL!(2) = WTPC!(2) / 79.9: esym$(2) = "Ti"
    PMOL!(3) = WTPC!(3) / 101.94: esym$(3) = "Al"
    PMOL!(4) = WTPC!(4) / 159.7: esym$(4) = "Fe+3"
    PMOL!(5) = WTPC!(5) / 71.85: esym$(5) = "Fe+2"
    PMOL!(6) = WTPC!(6) / 40.32: esym$(6) = "Mg"
    PMOL!(7) = WTPC!(7) / 70.94: esym$(7) = "Mn"
    PMOL!(8) = WTPC!(8) / 56.08: esym$(8) = "Ca"
    PMOL!(9) = WTPC!(9) / 61.982: esym$(9) = "Na"
    PMOL!(10) = WTPC!(10) / 94.2: esym$(10) = "K"
    PMOL!(11) = WTPC!(11) / 19#: esym$(11) = "F"
    PMOL!(12) = WTPC!(12) / 35.457: esym$(12) = "Cl"
    
' Normalize to FIXRAT = Fe3+/FeT following Schmidt (1992)
    PMOL!(4) = PMOL!(5) * FIXRAT!       ' moles of Fe3+
    PMOL!(5) = PMOL!(5) - PMOL!(4)      ' moles of FeO
    PMOL!(4) = PMOL!(4) * 0.5           ' moles of Fe2O3
    
    If j% <= 2 Then GoTo 102
    If j% = 3 Or j% = 4 Then GoTo 103
    If j% > 4 Then GoTo 104

' 2 OXYGENS (equivalence)
102:
    ATOP!(j%) = PMOL!(j%) * 2#
    GoTo 105

' 3 OXYGENS
103:
    ATOP!(j%) = PMOL!(j%) * 3#
    GoTo 105

' 1 OXYGEN, F, CL
104:
    ATOP!(j%) = PMOL!(j%) * 1#

' Next element
105:
    WTPCT = WTPCT + WTPC!(j%)
Next j%
      
WTPCF = WTPC(11) * 0.4211
WTPCCL = WTPC(12) * 0.2256
WTPCT = WTPCT - WTPCF - WTPCCL
ATOPT = 0#
      
For j% = 1 To MAXAMPHI% - 2
    ATOPT = ATOPT + ATOP!(j%)
Next j%

' USE 46 NEGATIVE CHARGES
d# = 23# / ATOPT!
For j% = 1 To MAXAMPHI%
    ANIO!(j%) = d * ATOP!(j%)
Next j%

For j% = 1 To MAXAMPHI%
      If j% <= 2 Then ANSFO!(j%) = ANIO!(j%) / 2#
      If j% = 3 Or j% = 4 Then ANSFO!(j%) = ANIO!(j%) * 0.6666
      If j% > 4 And j% <= 8 Then ANSFO!(j%) = ANIO!(j%)
      If j% = 9 Or j% = 10 Then ANSFO!(j%) = ANIO!(j%) * 2#
      If j% > 10 Then ANSFO!(j%) = ANIO!(j%)
Next j%

If DebugMode Then
Call IOWriteLog(vbCrLf & "Entered Amphibole Analysis:")
For j% = 1 To MAXAMPHI%
    If j% = 1 Then astring$ = Format$("SiO2", a80$)
    If j% = 2 Then astring$ = Format$("TiO2", a80$)
    If j% = 3 Then astring$ = Format$("Al2O3", a80$)
    If j% = 4 Then astring$ = Format$("Fe2O3", a80$)
    If j% = 5 Then astring$ = Format$("FeO", a80$)
    If j% = 6 Then astring$ = Format$("MgO", a80$)
    If j% = 7 Then astring$ = Format$("MnO", a80$)
    If j% = 8 Then astring$ = Format$("CaO", a80$)
    If j% = 9 Then astring$ = Format$("Na2O", a80$)
    If j% = 10 Then astring$ = Format$("K2O", a80$)
    If j% = 11 Then astring$ = Format$("F", a80$)
    If j% = 12 Then astring$ = Format$("Cl", a80$)
    If j% <> 4 Then bstring$ = Format$(Format$(WTPC!(j%), f83$), a80)
    If j% = 4 And IFE% = 1 Then bstring$ = Format$(Format$(WTPC!(j%), f83$), a80)
    If j% = 4 And IFE% <> 1 Then bstring$ = Format$(Format$(" ---- "), a80)
Call IOWriteLog(bstring$ & a4x$ & astring$)
Next j%
Call IOWriteLog(SSTRING2$)

astring$ = Format$("TOTAL", a80$)
bstring$ = Format$(Format$(WTT!, f83$), a80)
Call IOWriteLog(bstring$ & a4x$ & astring$)
      
astring$ = Format$("TOTAL-OXYGEN EQUIV. OF F,CL")
bstring$ = Format$(Format$(WTPCT!, f83$), a80)
Call IOWriteLog(bstring$ & a4x$ & astring$)
Call IOWriteLog(SSTRING1$)
End If
      
' Problem here if IFE% = 2 (CATF() array does not get loaded)
If IFE% = 1 Then
    NCODE% = 0
    For j% = 1 To MAXAMPHI%
        CATF!(j%) = ANSFO!(j%)
    Next j%
End If

' Round values in structural formula (array CATF) to 4 decimal places. If Mg = 0, set to 0.0001 to prevent errors.
For j% = 1 To MAXAMPHI%
    CATF!(j%) = MiscSetRounding2!(CATF!(j%), Int(4))
Next j%
If CATF!(6) = 0# Then CATF!(6) = 0.0001

' Compute mole fractions, molar ratios and pressure of crystallazation using
'  Schmidt (1992) calibration of the Hammarstrom and Zen (1986) barometer

' TETRAHEDRAL AL
      If CATF!(1) < 8# Then
      TETAL = 8# - CATF!(1)
      Else
      TETAL = 0#
      End If

' OCTAHEDRAL AL
      OCTAL = CATF(3) - TETAL
      If (OCTAL < 0#) Then TETAL = CATF(3)
      If (OCTAL < 0#) Then OCTAL = 0#

' NA ON M4 SITE
      ANAM4 = 7# - (CATF(8) + CATF(4) + CATF(5) + CATF(6) + CATF(7) + OCTAL + CATF(2))
      If (ANAM4 < 0#) Then Call IOWriteLog(a6x$ & "***OCTAHEDRAL NA IS NEGATIVE***")

' 12-FOLD NA
      If (ANAM4 < 0#) Then ANA12 = CATF(9)
      If (ANAM4 >= 0#) Then ANA12 = CATF(9) - ANAM4
      If ((ANAM4 >= 0#) And (ANAM4 > CATF(9))) Then ANA12 = 0#
      If (ANAM4 < 0#) Then Call IOWriteLog(a6x$ & "***ALL NA ASSIGNED TO 12-FOLD SITE***")
      If (ANAM4 < 0#) Then ANAM4 = 0#
      If (ANAM4 > CATF(9)) Then ANAM4 = CATF(9)

' TOTAL OCTAHEDRAL ATOMS
      OCTA = ANAM4 + OCTAL + CATF(8) + CATF(4) + CATF(5) + CATF(6) + CATF(2) + CATF(7)

' X - ALVI
      XALVI = OCTAL / OCTA

' X-FE2+
      FE2X = CATF(5) / OCTA

' X-FE3+
      FE3X = CATF(4) / OCTA

' X - MG
      MGMFT = CATF(6) / OCTA

' X - TI
      tix = CATF(2) / OCTA

' X - MN
      MNX = CATF(7) / OCTA

' X - CA
      CAX = CATF(8) / OCTA

'X - NAM4
      ANAM4X = ANAM4 / OCTA

' A-SITE OCCUPANCY
      ASITE = ANA12 + CATF(10)

' X - F
      HALMF = CATF(11) / 2#
      If HALMF = 0# Then HALMF = 0.00001

' X - CL
      HALCL = CATF(12) / 2#
      If HALCL = 0# Then HALCL = 0.00001

' X - OH
      HALMO = 1# - HALMF - HALCL
      If HALMO <= 0# Then HALMO = 0.00001

' COMPUTE LOG X-F/X-OH
      XFOXOH = MiscConvertLog10#(CDbl(HALMF / HALMO))

' COMPUTE LOG X-F/X-CL
      HALOG = MiscConvertLog10#(CDbl(HALMF / HALCL))

' COMPUTE LOG MG/FE2+
      RAMGFE = MiscConvertLog10#(CDbl(MGMFT / FE2X))

' COMPUTE LOG (MG/(FE2+ + FE3+))
      MGFET = MiscConvertLog10#(CDbl(MGMFT / (FE2X + FE3X)))

' COMPUTE LOG FE2+/FE3+
      If (FE3X = 0#) Then FE3X = 0.0001
      FELOG = MiscConvertLog10#(CDbl(FE2X / FE3X))

' COMPUTE MG/(MG+FE2+)
      MGB = CATF(6) / (CATF(6) + CATF(5))

' COMPUTE ALVI+FE3+ + 2TI+ASITE
      OC1 = OCTAL + CATF(4) + 2# * CATF(2) + ASITE

' COMPUTE ALVI+FE3+ +2TI
      OC2 = OCTAL + CATF(4) + 2# * CATF(2)

' COMPUTE FE2+/(FE2+ + FE3+)
      FERAT = CATF(5) / (CATF(5) + CATF(4))

' COMPUTE MG/(MG+FE TOTAL)
      MGTOFE = CATF(6) / (CATF(6) + CATF(4) + CATF(5))
      
' Compute pressure, both with total Al (all Fe2+) and total Al (Fe2+ - Fe3+)
PRESS2! = 4.76 * ANSFO!(3) - 3.01
PRESS23! = 4.76 * CATF!(3) - 3.01
If PRESS2! < 0# Or IFE = 1 Then PRESS2! = 0#
If PRESS23! < 0# Then PRESS23! = 0#

Call IOWriteLog(a6x$ & "STRUCTURAL FORMULA:")
Call IOWriteLog(a8x$ & "SI   " & Format$(Format$(CATF(1), f84$), a80$))
Call IOWriteLog(a8x$ & "ALIV " & Format$(Format$(TETAL!, f84$), a80$) & a8x$ & "ALVI" & Format$(Format$(OCTAL!, f84$), a80$))
Call IOWriteLog(a8x$ & "TI   " & Format$(Format$(CATF(2), f84$), a80$))
Call IOWriteLog(a8x$ & "FE3+ " & Format$(Format$(CATF(4), f84$), a80$))
Call IOWriteLog(a8x$ & "FE2+ " & Format$(Format$(CATF(5), f84$), a80$))
Call IOWriteLog(a8x$ & "MG   " & Format$(Format$(CATF(6), f84$), a80$))
Call IOWriteLog(a8x$ & "MN   " & Format$(Format$(CATF(7), f84$), a80$))
Call IOWriteLog(a8x$ & "CA   " & Format$(Format$(CATF(8), f84$), a80$))
Call IOWriteLog(a8x$ & "NA   " & Format$(Format$(ANAM4!, f84$), a80$) & a8x$ & "NA  " & Format$(Format$(ANA12!, f84$), a80$))
Call IOWriteLog(a8x$ & "K    " & Format$(Format$(CATF(10), f84$), a80$))
Call IOWriteLog(a8x$ & "F    " & Format$(Format$(CATF(11), f84$), a80$))
Call IOWriteLog(a8x$ & "CL   " & Format$(Format$(CATF(12), f84$), a80$))
Call IOWriteLog(a8x$ & "OH   " & Format$(Format$(2# - (CATF(11) + CATF(12)), f84$), a80$))

Call IOWriteLog(SSTRING1$)
Call IOWriteLog("MOLE FRACTIONS AND LOGARITHMS OF ATOMIC RATIOS:")

Call IOWriteLog(a6x$ & "X-FE2+= " & Format$(Format$(FE2X!, f83$), a80$) & a6x$ & "X-MG=   " & Format$(Format$(MGMFT!, f83$), a80$))
Call IOWriteLog(a6x$ & "X-FE3+= " & Format$(Format$(FE3X!, f83$), a80$) & a6x$ & "X-ALVI= " & Format$(Format$(XALVI, f83$), a80$))
Call IOWriteLog(a6x$ & "X-MN=   " & Format$(Format$(MNX!, f84$), a80$) & a6x$ & "X-TI=   " & Format$(Format$(tix!, f84$), a80$))
Call IOWriteLog(a6x$ & "X-CA=   " & Format$(Format$(CAX!, f84$), a80$) & a6x$ & "X-NAM4= " & Format$(Format$(ANAM4X!, f84$), a80$) & vbCrLf)

Call IOWriteLog(a6x$ & "MG / (MG + FE2+) =  " & Format$(Format$(MGB!, f83$), a80$))
Call IOWriteLog(a6x$ & "FE2+/(FE2+ + FE3+)= " & Format$(Format$(FERAT!, f83$), a80$))
Call IOWriteLog(a6x$ & "MG/(MG+FE2+ + FE3+)=" & Format$(Format$(MGTOFE!, f83$), a80$) & vbCrLf)

Call IOWriteLog(a6x$ & "X-OH=                     " & Format$(Format$(HALMO!, f83$), a80$) & a4x$ & "X-F=               " & Format$(Format$(HALMF!, f83$), a80$) & a4x & "X-CL=         " & Format$(Format$(HALCL!, f84$), a80$))
Call IOWriteLog(a6x$ & "LOG(X-MG/X-FE2+)=         " & Format$(Format$(RAMGFE!, f83$), a80$) & a4x$ & "LOG(X-F/X-CL)=     " & Format$(Format$(HALOG!, f83$), a80$) & a4x$ & "LOG(X-F/X-OH)=" & Format$(Format$(XFOXOH!, f83$), a80$))
Call IOWriteLog(a6x$ & "LOG(X-MG/(X-FE2+ + FE3+))=" & Format$(Format$(MGFET!, f83$), a80$) & a4x$ & "LOG(X-FE2+/X-FE3+)=" & Format$(Format$(FELOG!, f83$), a80$) & vbCrLf)
Call IOWriteLog(a6x$ & "A-SITE=                   " & Format$(Format$(ASITE!, f83$), a80$) & a4x$ & "TOTAL VI=          " & Format$(Format$(OCTA!, f83$), a80$))
Call IOWriteLog(a6x$ & "ALVI+2TI+A-SITE+FE3+=     " & Format$(Format$(OC1!, f83$), a80$) & a4x$ & "ALVI+2TI+FE3+=     " & Format$(Format$(OC2!, f83$), a80$))
      
Call IOWriteLog(vbNullString)
Call IOWriteLog(a6x$ & "Schmidt (1992) Pressure (All FE2+): " & Format$(Format$(PRESS2!, f42$), a80$) & " KBar, " & a4x$ & "(FE2+ -FE3+): " & Format$(Format$(PRESS23!, f42$), a80$) & " KBar")
      
' Output to file
astring$ = "Sample " & VbDquote$ & sample(1).number% & VbDquote$
Print #tfilenumber%, astring$

astring$ = vbNullString
For j% = 1 To MAXAMPHI%
    If j% = 3 Then
    astring$ = astring$ + MiscAutoFormat$(WTPC!(j%)) & vbTab$ & MiscAutoFormat$(ANSFO!(j%)) & vbTab$ & MiscAutoFormat$(CATF!(j%)) & vbTab$ & MiscAutoFormat$(TETAL!) & vbTab$ & MiscAutoFormat$(OCTAL!) & vbTab & esym$(j%) & vbTab$ & MiscAutoFormat$(PRESS2!) & vbTab$ & MiscAutoFormat$(PRESS23!) & vbCrLf
    ElseIf j = 9 Then
    astring$ = astring$ + MiscAutoFormat$(WTPC!(j%)) & vbTab$ & MiscAutoFormat$(ANSFO!(j%)) & vbTab$ & MiscAutoFormat$(CATF!(j%)) & vbTab$ & MiscAutoFormat$(ANAM4!) & vbTab$ & MiscAutoFormat$(ANA12!) & vbTab & esym$(j%) & vbCrLf
    ElseIf j = 12 Then
    astring$ = astring$ + MiscAutoFormat$(WTPC!(j%)) & vbTab$ & MiscAutoFormat$(ANSFO!(j%)) & vbTab$ & MiscAutoFormat$(CATF!(j%)) & vbTab$ & MiscAutoFormat$(2# - (CATF!(11) + CATF!(12))) & vbTab$ & MiscAutoFormat$(ANA12!) & vbTab & esym$(j%) & vbCrLf
    Else
    astring$ = astring$ + MiscAutoFormat$(WTPC!(j%)) & vbTab$ & MiscAutoFormat$(ANSFO!(j%)) & vbTab$ & MiscAutoFormat$(CATF!(j%)) & vbTab & esym$(j%) & vbCrLf
    End If
Next j%
Print #tfilenumber%, astring$

astring$ = vbNullString
astring$ = astring$ & MiscAutoFormat$(HALOG!) & vbTab$ & MiscAutoFormat$(RAMGFE!) & vbTab$ & MiscAutoFormat$(MGB!) & vbTab$ & MiscAutoFormat$(FE2X!) & vbTab$ & MiscAutoFormat$(tix!) & vbTab$ & MiscAutoFormat$(XALVI!) & vbTab$ & MiscAutoFormat$(MNX!) & vbTab$ & MiscAutoFormat$(XFOXOH!) & vbCrLf
astring$ = astring$ & MiscAutoFormat$(OC1!) & vbTab$ & MiscAutoFormat$(OC2!) & vbTab$ & MiscAutoFormat$(CATF(3)) & vbTab$ & MiscAutoFormat$(ASITE!) & vbTab$ & MiscAutoFormat$(FERAT!) & vbTab$ & MiscAutoFormat$(CATF(1)) & vbCrLf
astring$ = astring$ & MiscAutoFormat$(ANAM4!) & vbTab$ & MiscAutoFormat$(ANA12!) & vbTab$ & MiscAutoFormat$(OCTA!) & vbTab$ & MiscAutoFormat$(CATF!(8)) & vbTab$ & MiscAutoFormat$(FIXRAT!) & vbTab$ & MiscAutoFormat$(WTT!)
Print #tfilenumber%, astring$

Exit Sub

' Errors
ConvertAmphibole2Error:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertAmphibole2"
Close #tfilenumber%
Close #tfilenumber2%
ierror = True
Exit Sub

End Sub

Sub ConvertAmphiboleNorm(ANSFO() As Single, CAT1() As Single, CAT2() As Single, CAT3() As Single, CAT4() As Single, CAT5() As Single)
'  Normalize amphibole structural formulas (written by Jay Ague, converted to VB by John Donovan)
'  The specific normalizations performed are:
'
'      NORM 1: TOTAL CATIONS - (Na+K) = 15
'      NORM 2: TOTAL CATIONS - (Na+Ca+K) = 13
'      NORM 3: TOTAL CATIONS - (K) = 15
'      NORM 4: Si + Al = 8
'      NORM 5: TOTAL CATIONS = 15.8 (Best value for arvfredsonite compositions in DHZ.)
'
'  The calculations begin with a normalization of cation sums according
'  to each of the above schemes. All Fe is here taken to be divalent.
'  For each normalization, the total charge of all cations, exclusive
'  of Fe, is summed. The excess charge is 46 - this sum. The distribution
'  of Fe2+ and Fe3+ is then described by two equations in two unknowns:
'
'      2Fe2+ + 3Fe3+ = excess charge
'       Fe2+ +  Fe3+ = total Fe from normalization

ierror = False
On Error GoTo ConvertAmphiboleNormError

Dim j As Integer
Dim AFRAC As Single, FET As Single
Dim ATOM1 As Single, ATOM2 As Single, ATOM3 As Single, ATOM4 As Single, ATOM5 As Single
Dim TOTAL1 As Single, TOTAL2 As Single, TOTAL3 As Single, TOTAL4 As Single, TOTAL5 As Single

ATOM1! = 0#
ATOM2! = 0#
ATOM3! = 0#
ATOM4! = 0#
ATOM5! = 0#

TOTAL1! = 0#
TOTAL2! = 0#
TOTAL3! = 0#
TOTAL4! = 0#
TOTAL5! = 0#

' NORMALIZATION 1: TOTAL CATIONS - (NA+K)=15
For j% = 1 To MAXAMPHI% - 4
      ATOM1 = ATOM1 + ANSFO!(j%)
Next j%
    AFRAC = 15# / ATOM1

' SUM CATIONS AND INITIALIZE TOTAL FE
Call ConvertAmphiboleSum(CAT1!(), TOTAL1!, ANSFO!(), AFRAC!)
If ierror Then Exit Sub

' CALCULATE FE2+, FE3+, FOR NORM 1
    FET! = CAT1!(5)
Call ConvertAmphiboleGetFe(TOTAL1!, FET!, CAT1(4), CAT1(5))
If ierror Then Exit Sub

' HALOGENS FOR NORM 1
    CAT1(11) = ANSFO(11) * AFRAC
    CAT1(12) = ANSFO(12) * AFRAC

' NORMALIZATION 2: TOTAL CATIONS - (NA+K+CA)=13
For j% = 1 To MAXAMPHI% - 5
      ATOM2 = ATOM2 + ANSFO!(j%)
Next j%
      AFRAC = 13# / ATOM2

' SUM CATIONS AND INITIALIZE TOTAL FE
Call ConvertAmphiboleSum(CAT2!(), TOTAL2!, ANSFO!(), AFRAC!)
If ierror Then Exit Sub

' CALCULATE FE2+, FE3+, FOR NORM 2
    FET! = CAT2!(5)
Call ConvertAmphiboleGetFe(TOTAL2!, FET!, CAT2(4), CAT2(5))
If ierror Then Exit Sub

' HALOGENS FOR NORM 2
    CAT2(11) = ANSFO(11) * AFRAC
    CAT2(12) = ANSFO(12) * AFRAC

' NORMALIZATION 3: TOTAL-K = 15
For j% = 1 To MAXAMPHI% - 3
      ATOM3 = ATOM3 + ANSFO!(j%)
Next j%
    AFRAC = 15# / ATOM3

' SUM CATIONS AND INITIALIZE TOTAL FE
Call ConvertAmphiboleSum(CAT3!(), TOTAL3!, ANSFO!(), AFRAC!)
If ierror Then Exit Sub

' CALCULATE FE2+, FE3+, FOR NORM 3
    FET! = CAT3!(5)
Call ConvertAmphiboleGetFe(TOTAL3!, FET!, CAT3(4), CAT3(5))
If ierror Then Exit Sub

' HALOGENS FOR NORM 3
    CAT3(11) = ANSFO(11) * AFRAC
    CAT3(12) = ANSFO(12) * AFRAC

' NORMALIZATION 4: SI+AL=8.0
      ATOM4 = ANSFO(1) + ANSFO(3)
      AFRAC = 8# / ATOM4

' SUM CATIONS AND INITIALIZE TOTAL FE
Call ConvertAmphiboleSum(CAT4!(), TOTAL4!, ANSFO!(), AFRAC!)
If ierror Then Exit Sub

' CALCULATE FE2+, FE3+, FOR NORM 4
    FET! = CAT4!(5)
Call ConvertAmphiboleGetFe(TOTAL4!, FET!, CAT4(4), CAT4(5))
If ierror Then Exit Sub

' HALOGENS FOR NORM 4
    CAT4(11) = ANSFO(11) * AFRAC
    CAT4(12) = ANSFO(12) * AFRAC

' NORMALIZATION 5: TOTAL CATIONS = 15.8
For j% = 1 To MAXAMPHI% - 2
      ATOM5 = ATOM5 + ANSFO!(j%)
Next j%
    AFRAC = 15.8 / ATOM5

' SUM CATIONS AND INITIALIZE TOTAL FE
Call ConvertAmphiboleSum(CAT5!(), TOTAL5!, ANSFO!(), AFRAC!)
If ierror Then Exit Sub

' CALCULATE FE2+, FE3+, FOR NORM 5
    FET! = CAT5!(5)
Call ConvertAmphiboleGetFe(TOTAL5!, FET!, CAT5(4), CAT5(5))
If ierror Then Exit Sub

' HALOGENS FOR NORM 5
    CAT5(11) = ANSFO(11) * AFRAC
    CAT5(12) = ANSFO(12) * AFRAC

Exit Sub

' Errors
ConvertAmphiboleNormError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertAmphiboleNorm"
Close #tfilenumber%
Close #tfilenumber2%
ierror = True
Exit Sub

End Sub

Sub ConvertAmphiboleGetFe(total As Single, FET As Single, FE3 As Single, FE2 As Single)
'  Calculate the amphibole Fe2+ and Fe3+ for a given normalization.

ierror = False
On Error GoTo ConvertAmphiboleGetFeError

Dim excess As Single

excess = 46# - total
FE3 = excess - 2# * FET
If FE3 < 0# Then
    FE2 = -1#
Else
    FE2 = FET - FE3
End If

'    If CAT2(4) < 0# Then CAT2(5) = -1#
'    If CAT2(4) > ANSFO(5) Then CAT2(5) = -1#
    If FE3 < 0# Then FE2 = -1#
    If FE3 > FET Then FE2 = -1#

Exit Sub

' Errors
ConvertAmphiboleGetFeError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertAmphiboleGetFe"
Close #tfilenumber%
Close #tfilenumber2%
ierror = True
Exit Sub

End Sub

Sub ConvertAmphiboleSum(cat() As Single, total As Single, ANSFO() As Single, AFRAC As Single)
'  Sum up the charges on cations, exclusive of Fe, for a given normalization

ierror = False
On Error GoTo ConvertAmphiboleSumError

Dim j As Integer

ReDim TOT(1 To MAXAMPHI% - 2) As Single

For j% = 1 To MAXAMPHI% - 2
      cat!(j%) = AFRAC * ANSFO!(j%)
      If j% <= 2 Then TOT!(j%) = cat!(j%) * 4#
      If j% = 3 Then TOT!(j%) = cat!(j%) * 3#
      If j% = 4 Or j% = 5 Then TOT!(j%) = 0#
      If j% > 5 And j% <= 8 Then TOT!(j%) = cat!(j%) * 2#
      If j% > 8 Then TOT!(j%) = cat!(j%)
      total = total + TOT!(j%)
Next j%

Exit Sub

' Errors
ConvertAmphiboleSumError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertAmphiboleSum"
Close #tfilenumber%
Close #tfilenumber2%
ierror = True
Exit Sub

End Sub

Sub ConvertAmphiboleAver(ANSFO() As Single, CAT1() As Single, CAT2() As Single, CAT3() As Single, CAT4() As Single, CAT5() As Single, CATF() As Single, INORM As Integer)
'  Provide "best estimate" of amphibole structural formula. INORM selects the desired normalization or
'  average of normalizations
'   1) Norm 1
'   2) Norm 2
'   3) Norm 3
'   4) Norm 4
'   5) All Fe2+
'   6) Norms 1, 2
'   7) All Fe2+ - Norm 1
'   8) All Fe2+ - Norm 2
'   9) Norms 3, 4
'  10) Norms 2, 3
'  11) Norms 2,3,4
'  12) Norm 5
'  13) Auto

ierror = False
On Error GoTo ConvertAmphiboleAverError

Dim i As Integer

' Default is all as Fe2+
If INORM% = 0 Then INORM = 5

' Load selected norm
For i% = 1 To MAXAMPHI%
If INORM = 1 Then CATF(i) = CAT1(i)
If INORM = 2 Then CATF(i) = CAT2(i)
If INORM = 3 Then CATF(i) = CAT3(i)
If INORM = 4 Then CATF(i) = CAT4(i)
If INORM = 5 Then CATF(i) = ANSFO(i)
If INORM = 6 Then CATF(i) = (CAT1(i) + CAT2(i)) / 2#
If INORM = 7 Then CATF(i) = (ANSFO(i) + CAT1(i)) / 2#
If INORM = 8 Then CATF(i) = (ANSFO(i) + CAT2(i)) / 2#
If INORM = 9 Then CATF(i) = (CAT3(i) + CAT4(i)) / 2#
If INORM = 10 Then CATF(i) = (CAT2(i) + CAT3(i)) / 2#
If INORM = 11 Then CATF(i) = (CAT2(i) + CAT3(i) + CAT4(i)) / 3#
If INORM = 12 Then CATF(i) = CAT5(i)
If INORM = 13 Then Call ConvertAmphiboleAutoAv(ANSFO!(), CAT1!(), CAT2!(), CAT3!(), CAT4!(), CATF!())
Next i%
      
Exit Sub

' Errors
ConvertAmphiboleAverError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertAmphiboleAver"
Close #tfilenumber%
Close #tfilenumber2%
ierror = True
Exit Sub

End Sub
      
Sub ConvertAmphiboleAutoAv(ANSFO() As Single, CAT1() As Single, CAT2() As Single, CAT3() As Single, CAT4() As Single, CATF() As Single)
'  Automatic Fe3+ averaging routine. Performs averaging of normalizations such that structural
'  formulas with greatest lower bound and least upper bound on Fe3+ are averaged. Results are
'  consistent with the Papike algorithm, as used, for example, by Czamanske et al. (1981).

ierror = False
On Error GoTo ConvertAmphiboleAutoAvError

Dim j As Integer

' Start averaging
        If CAT4(5) < 0# Then GoTo 904:
        If CAT2(5) < 0# Or CAT3(5) < 0# Then GoTo 905:
        If CAT2(4) < CAT3(4) And CAT2(4) < CAT4(4) Then GoTo 906:
        If CAT3(4) < CAT2(4) And CAT3(4) < CAT4(4) Then GoTo 907
        If CAT4(4) < CAT2(4) And CAT4(4) < CAT3(4) Then GoTo 908

' Iterate
906:    For j% = 1 To MAXAMPHI%
          If CAT1(5) < 0# Then CATF(j) = (CAT2(j) + ANSFO(j)) / 2#
          If CAT1(5) >= 0# Then CATF(j) = (CAT2(j) + CAT1(j)) / 2#
Next j%
        GoTo Success
        
' Iterate
907     For j% = 1 To MAXAMPHI%
          If CAT1(5) < 0# Then CATF(j) = (CAT3(j) + ANSFO(j)) / 2#
          If CAT1(5) >= 0# Then CATF(j) = (CAT3(j) + CAT1(j)) / 2#
        Next j%
        GoTo Success

' Iterate
908:    For j% = 1 To MAXAMPHI%
          If CAT1(5) < 0# Then CATF(j) = (CAT4(j) + ANSFO(j)) / 2#
          If CAT1(5) >= 0# Then CATF(j) = (CAT4(j) + CAT1(j)) / 2#
        Next j%
        GoTo Success
        
904:     If CAT2(5) < 0# And CAT3(5) < 0# Then GoTo Failure
        If CAT2(5) < 0# Then GoTo 907:
        If CAT3(5) < 0# Then GoTo 906:
        If CAT2(4) > CAT3(4) Then GoTo 907:
        If CAT2(4) < CAT3(4) Then GoTo 906:
        GoTo Failure
        
905:    If CAT2(5) < 0# And CAT3(5) < 0# Then GoTo 908:
        If CAT2(5) < 0# Then GoTo 913:
        If CAT2(4) < CAT4(4) Then GoTo 906:
        GoTo 908
913:    If CAT3(4) < CAT4(4) Then GoTo 907:
        GoTo 908
  
Failure:
Call IOWriteLog("ConvertAmphiboleAutoAv: NORMALIZATION PROCEDURE HAS FAILED")
        
' Load return array
For j% = 1 To MAXAMPHI%
        CATF(j) = ANSFO(j)
Next j%
        ANSFO(4) = 0.0001

Success:
Exit Sub

' Errors
ConvertAmphiboleAutoAvError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertAmphiboleAutoAv"
Close #tfilenumber%
Close #tfilenumber2%
ierror = True
Exit Sub

End Sub
