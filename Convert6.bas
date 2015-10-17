Attribute VB_Name = "CodeCONVERT6"
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit

Sub ConvertBiotite(percents() As Single, sample() As TypeSample)
' Biotite calculation (COMPUTES BIOTITE FORMULAS AND SITE OCCUPANCIES AND MOLE FRACTION
'    RATIOS FOR USE IN HALOGEN CHEMISTRY AND CHARACTERIZATION OF MINERALIZATION ENVIRONMENTS)
'  Originally written by George Brimhall, Jay Ague and John Donovan in FORTRAN
'  Translated to Visual Basic by John Donovan

ierror = False
On Error GoTo ConvertBiotiteError

Const MAXBIOT% = 15

Dim i As Integer, ip As Integer, iter As Integer

Dim TOT As Single, MTOT As Single, NOAN As Single, NOVOL As Single, NOWAT As Single
Dim XANN As Single, LOGOHF As Single, FOH As Single
Dim WTPCT As Single, WTPCF As Single, WTPCCL As Single, ATOPT As Single, PTOT As Single
Dim TETAL As Single, OCTAL As Single, temp1 As Single
Dim XALVI As Single, ALKM As Single, MGMFT As Single, FEX As Single, tix As Single, MNX As Single
Dim HALMF As Single, HALMC As Single, HALMO As Single, XFOXOH As Single, HALOG1 As Single, SI As Single
Dim AFEF As Single, AMGF As Single, AFEOH As Single, AMGOH As Single, RAMGFE As Single, MGB As Single
Dim sum1 As Single, sum2 As Single, sum3 As Single, sum4 As Single, temp As Single

Dim numcat(1 To MAXBIOT%) As Single, NUMOXG(1 To MAXBIOT%) As Single
Dim mw(1 To MAXBIOT%) As Single, mp(1 To MAXBIOT%) As Single
Dim AP(1 To MAXBIOT%) As Single, AN(1 To MAXBIOT%) As Single, num(1 To MAXBIOT%) As Single
Dim ANSFO(1 To MAXBIOT%) As Single, ANIO(1 To MAXBIOT%) As Single
Dim WTPC(1 To MAXBIOT%) As Single, MOPC(1 To MAXBIOT%) As Single

Dim esym(1 To MAXBIOT%) As String

Dim d As Double

Dim OXIDE(1 To MAXBIOT%) As String * 6
Dim ATYPE As String * 2
Dim astring As String

For i% = 1 To MAXBIOT%
OXIDE$(1) = "SIO2": numcat!(1) = 1: NUMOXG!(1) = 2: mw!(1) = 60.09: WTPC!(1) = NotAnalyzedValue!
OXIDE$(2) = "TIO2": numcat!(2) = 1: NUMOXG!(2) = 2: mw!(2) = 79.9: WTPC!(2) = NotAnalyzedValue!
OXIDE$(3) = "AL2O3": numcat!(3) = 2: NUMOXG!(3) = 3: mw!(3) = 101.96: WTPC!(3) = NotAnalyzedValue!
OXIDE$(4) = "FEO": numcat!(4) = 1: NUMOXG!(4) = 1: mw!(4) = 71.85: WTPC!(4) = NotAnalyzedValue!
OXIDE$(5) = "MGO": numcat!(5) = 1: NUMOXG!(5) = 1: mw!(5) = 40.31: WTPC!(5) = NotAnalyzedValue!
OXIDE$(6) = "CAO": numcat!(6) = 1: NUMOXG!(6) = 1: mw!(6) = 56.08: WTPC!(6) = NotAnalyzedValue!
OXIDE$(7) = "NA2O": numcat!(7) = 2: NUMOXG!(7) = 1: mw!(7) = 61.982: WTPC!(7) = NotAnalyzedValue!
OXIDE$(8) = "BAO": numcat!(8) = 1: NUMOXG!(8) = 1: mw!(8) = 153.36: WTPC!(8) = NotAnalyzedValue!
OXIDE$(9) = "K2O": numcat!(9) = 2: NUMOXG!(9) = 1: mw!(9) = 94.2: WTPC!(9) = NotAnalyzedValue!
OXIDE$(10) = "F": numcat!(10) = 1: NUMOXG!(10) = 1: mw!(10) = 18.998: WTPC!(10) = NotAnalyzedValue!
OXIDE$(11) = "CL": numcat!(11) = 1: NUMOXG!(11) = 1: mw!(11) = 35.457: WTPC!(11) = NotAnalyzedValue!
OXIDE$(12) = "MNO": numcat!(12) = 1: NUMOXG!(12) = 1: mw!(12) = 70.94: WTPC!(12) = NotAnalyzedValue!
OXIDE$(13) = "CR2O3": numcat!(13) = 2: NUMOXG!(13) = 3: mw!(13) = 152.02: WTPC!(13) = NotAnalyzedValue!
OXIDE$(14) = "NIO": numcat!(14) = 1: NUMOXG!(14) = 1: mw!(14) = 74.71: WTPC!(14) = NotAnalyzedValue!
OXIDE$(15) = "H2O": numcat!(15) = 2: NUMOXG!(15) = 1: mw!(15) = 18.016: WTPC!(15) = NotAnalyzedValue!
Next i%

' Print calculation
Call IOWriteLog(vbCrLf & "Biotite Formula Calculations (from Brimhall, et. al. BIOTITE.F code)...")
      
' Load oxide percents
For i% = 1 To sample(1).LastChan%
    ip% = IPOS1%(MAXELM%, sample(1).Elsyms$(i%), Symlo$())
    If ip% <> 0 Then
        If ip% = 14 Then WTPC!(1) = percents!(i%): esym$(1) = "Si"   ' SiO2
        If ip% = 22 Then WTPC!(2) = percents!(i%): esym$(2) = "Ti"   ' TiO2
        If ip% = 13 Then WTPC!(3) = percents!(i%): esym$(3) = "Al"   ' Al2O3
        If ip% = 26 Then WTPC!(4) = percents!(i%): esym$(4) = "Fe"   ' FeO
        If ip% = 12 Then WTPC!(5) = percents!(i%): esym$(5) = "Mg"   ' MgO
        If ip% = 20 Then WTPC!(6) = percents!(i%): esym$(6) = "Ca"   ' CaO
        If ip% = 11 Then WTPC!(7) = percents!(i%): esym$(7) = "Na"   ' Na2O
        If ip% = 56 Then WTPC!(8) = percents!(i%): esym$(8) = "Ba"   ' BaO
        If ip% = 19 Then WTPC!(9) = percents!(i%): esym$(9) = "K"   ' K2O
        If ip% = 9 Then WTPC!(10) = percents!(i%): esym$(10) = "F"   ' F
        If ip% = 17 Then WTPC!(11) = percents!(i%): esym$(11) = "Cl"  ' Cl
        If ip% = 25 Then WTPC!(12) = percents!(i%): esym$(12) = "Mn"  ' MnO
        If ip% = 24 Then WTPC!(13) = percents!(i%): esym$(13) = "Cr"  ' Cr2O3
        If ip% = 28 Then WTPC!(14) = percents!(i%): esym$(14) = "Ni"  ' NiO
        If ip% = 1 Then WTPC!(15) = percents!(i%): esym$(15) = "H"   ' H2O
    End If
Next i%
        
' START CALCULATIONS
NOVOL! = 0#
NOWAT! = 0#
MTOT! = 0#
TOT! = 0#
iter% = 0

' CALCULATE TOTAL WEIGHT PERCENT (without H2O)
For i% = 1 To MAXBIOT%
    TOT! = TOT! + WTPC!(i%)
Next i%

' CALCULATE FLUORINE AND OXYGEN EQUIVALENCE. SEE DHZ, APPENDIX 1. STORE ORIGINAL TOTAL IN "PTOT"
PTOT! = TOT!
WTPCT! = TOT!
WTPCF! = WTPC!(10) * 15.9994 / (2 * mw!(10))
WTPCCL! = WTPC!(11) * 15.9994 / (2 * mw!(11))
WTPCT! = WTPCT! - 1! * WTPCF! - 1# * WTPCCL!
TOT! = WTPCT!

' CALCULATE MOLE PERCENT
For i% = 1 To MAXBIOT%
    mp!(i%) = WTPC!(i%) / mw!(i%)
    MOPC!(i%) = mp!(i%)
    AP!(i%) = mp!(i%) * NUMOXG!(i%)
    MTOT! = MTOT! + AP!(i%)
Next i%

' CORRECT FOR FORMULA ATOMS FOR FLUORINE AND CHLORINE
        ATOPT! = MTOT!
        ATOPT! = ATOPT! - 1! * AP!(10) - 1! * AP!(11)
        MTOT! = ATOPT!

' BRIMHALL CALCULATIONS, BASED ON 22 NEGATIVE CHARGES
    d# = 11! / ATOPT
For i% = 1 To MAXBIOT%
    ANIO!(i%) = d# * AP!(i%)
    ANSFO!(i%) = ANIO!(i%) * numcat!(i%) / NUMOXG!(i%)
Next i%

' ITER LOOP BEGIN HERE
1400:   NOAN = 24# / MTOT

' CALCULATE FORMULA ATOMS
For i% = 1 To MAXBIOT%
    AN!(i%) = NOAN * AP!(i%)
    num!(i%) = AN!(i%) * numcat!(i%) / NUMOXG!(i%)
Next i%

' PRINT OUT WEIGHT PERCENTS AND FORMULA ATOMS
If iter% > 0 Then GoTo 1900

astring$ = vbCrLf
For i% = 1 To 8
astring$ = astring$ & Format$(OXIDE$(i%), a80$)
Next i%
astring$ = astring$ & vbCrLf
For i% = 1 To 8
astring$ = astring$ & Format$(Format$(WTPC!(i%), f83$), a80$)
Next i%
astring$ = astring$ & vbCrLf
For i% = 1 To 8
astring$ = astring$ & Format$(Format$(ANSFO!(i%), f85$), a80$)
Next i%
Call IOWriteLog(astring$)

astring$ = vbCrLf
For i% = 9 To MAXBIOT%
astring$ = astring$ & Format$(OXIDE$(i%), a80$)
Next i%
astring$ = astring$ & Format$("TOTAL", a80$) & vbCrLf
For i% = 9 To MAXBIOT%
astring$ = astring$ & Format$(Format$(WTPC!(i%), f83$), a80$)
Next i%
astring$ = astring$ & Format$(Format$(PTOT!, f83$), a80$) & vbCrLf
For i% = 9 To MAXBIOT%
astring$ = astring$ & Format$(Format$(ANSFO!(i%), f85$), a80$)
Next i%
astring$ = astring$ & Format$(Format$(ATOPT!, f85$), a80$)
Call IOWriteLog(astring$)

' BRIMHALL CALCULATIONS
        TETAL = 4# - 1# * ANSFO(1)
        OCTAL = ANSFO(3) - 1# * TETAL
        If (OCTAL < 0#) Then
        TETAL = ANSFO(3)
        OCTAL = 0#
        End If

        XALVI = OCTAL / (ANSFO(5) + ANSFO(2) + OCTAL + ANSFO(4) + ANSFO(12))

' K20 + NA20 + BA0 + CA0
        ALKM = ANSFO(9) + ANSFO(7) + ANSFO(8) + ANSFO(6)

' X MG (FULL OCTAHEDRAL)
        MGMFT = ANSFO(5) / (ANSFO(5) + ANSFO(2) + OCTAL + ANSFO(4) + ANSFO(12))

' X FE++ (FULL OCTAHEDRAL ANNITE)
        FEX = ANSFO(4) / (ANSFO(5) + ANSFO(2) + OCTAL + ANSFO(4) + ANSFO(12))

' X TI-BIOTITE (FULL OCTAHEDRAL)
        tix = ANSFO(2) / (ANSFO(5) + ANSFO(2) + OCTAL + ANSFO(4) + ANSFO(12))

' X MN- BIOTITE (FULL OCTAHEDRAL)
        MNX = ANSFO(12) / (ANSFO(5) + ANSFO(2) + OCTAL + ANSFO(4) + ANSFO(12))

' X-F, X-CL, X-OH, LOG X-F/X-OH, LOG X-F/X-CL
        HALMF = ANSFO(10) / 2#
        HALMC = ANSFO(11) / 2#
        HALMO = 1# - HALMF - HALMC
        If HALMF = 0# Then HALMF = NotAnalyzedValue!
        If HALMC = 0# Then HALMC = NotAnalyzedValue!
        If HALMO <= 0# Then HALMO = NotAnalyzedValue!
        XFOXOH = MiscConvertLog10#(HALMF / HALMO)
        HALOG1 = MiscConvertLog10#(HALMF / HALMC)

' BIOTITE COMPONENT ACTIVITIES
        SI = ANSFO(1) - 2#
        AFEF = MiscConvertLog10#(ANSFO(9) * ((ANSFO(4) / 3#) ^ 3) * TETAL * SI * (HALMF ^ 2))
        AMGF = MiscConvertLog10#(ANSFO(9) * ((ANSFO(5) / 3#) ^ 3) * TETAL * SI * (HALMF ^ 2))
        AFEOH = MiscConvertLog10#(ANSFO(9) * ((ANSFO(4) / 3#) ^ 3) * TETAL * SI * (HALMO ^ 2))
        AMGOH = MiscConvertLog10#(ANSFO(9) * ((ANSFO(5) / 3#) ^ 3) * TETAL * SI * (HALMO ^ 2))

        RAMGFE = MiscConvertLog10#(MGMFT / FEX)
        MGB = ANSFO(5) / (ANSFO(5) + ANSFO(4))

' PRINT OUT BRIMHALL CALCULATIONS
astring$ = vbCrLf
astring$ = astring$ & Format$("LOG(XF/XCL)", a14$) & Format$("X-MG++(OCT)", a14$)
astring$ = astring$ & Format$("X-MG++(BIN)", a14$) & Format$("X-FE++(OCT)", a14$)
astring$ = astring$ & Format$("X-TI", a14$) & Format$("AL-VI", a14$)
Call IOWriteLog(astring$)

astring$ = vbNullString
astring$ = astring$ & Format$(Format$(HALOG1!, f83$), a14$) & Format$(Format$(MGMFT!, f83$), a14$)
astring$ = astring$ & Format$(Format$(MGB!, f83$), a14$) & Format$(Format$(FEX!, f83$), a14$)
astring$ = astring$ & Format$(Format$(tix!, f83$), a14$) & Format$(Format$(XALVI!, f83$), a14$)
Call IOWriteLog(astring$)

astring$ = vbCrLf
astring$ = astring$ & Format$("LOG(X-MG/X-FE)", a14$) & Format$("X-MN(OCT)", a14$)
astring$ = astring$ & Format$("LOG(X-F/X-OH)", a14$) & Format$("X-OH", a14$)
astring$ = astring$ & Format$("X-F", a14$) & Format$("X-CL", a14$)
Call IOWriteLog(astring$)

astring$ = vbNullString
astring$ = astring$ & Format$(Format$(RAMGFE!, f83$), a14$) & Format$(Format$(MNX!, f83$), a14$)
astring$ = astring$ & Format$(Format$(XFOXOH!, f83$), a14$) & Format$(Format$(HALMO!, f83$), a14$)
astring$ = astring$ & Format$(Format$(HALMF!, f83$), a14$) & Format$(Format$(HALMC!, f83$), a14$)
Call IOWriteLog(astring$)

' ROUTINE TO CLASSIFY BIOTITE BY MG/FE AND F/OH RATIOS
        If RAMGFE <= -1# Then
        ATYPE$ = "SR"
        GoTo 1700
        End If
        If RAMGFE <= -0.2 Then
        ATYPE$ = "sr"
        GoTo 1700
        End If
        If XFOXOH <= -1.5 Then
        ATYPE$ = "WC"
        GoTo 1700
        End If
        If XFOXOH <= -1# Then
        ATYPE$ = "MC"
        GoTo 1700
        End If
        If XFOXOH > -1# Then
        ATYPE$ = "SC"
        GoTo 1700
        End If

' SUM FORMULA UNITS IN 12-FOLD SITE, 6-FOLD, 4-FOLD, HYDROXL SITE
1700:   sum1 = ANSFO(9) + ANSFO(7) + ANSFO(6) + ANSFO(8)
        sum2 = ANSFO(2) + OCTAL + ANSFO(4) + ANSFO(5) + ANSFO(12)
        sum3 = TETAL + ANSFO(1)
        sum4 = ANSFO(10) + ANSFO(11)

astring$ = vbCrLf
astring$ = astring$ & Format$("AL(IV)= ", a80$) & Format$(Format$(TETAL!, f83$), a80$)
astring$ = astring$ & Format$("   SI = ", a80$) & Format$(Format$(ANSFO!(1), f83$), a80$)
astring$ = astring$ & "         ****** BIOTITE TYPE = " & ATYPE$ & " ******"
Call IOWriteLog(astring$)

astring$ = vbNullString
astring$ = astring$ & Format$("(K+NA+CA+BA)", a22$) & Format$("(TI+AL(VI)+FE+MG+MN)", a22$)
astring$ = astring$ & Format$("(AL(IV)+SI)", a22$) & Format$("(F+CL)", a22$)
Call IOWriteLog(astring$)

astring$ = vbNullString
astring$ = astring$ & Format$(Format$(sum1!, f84$), a22$) & Format$(Format$(sum2!, f84$), a22$)
astring$ = astring$ & Format$(Format$(sum3!, f84$), a22$) & Format$(Format$(sum4!, f84$), a22$)
Call IOWriteLog(astring$)

astring$ = vbCrLf
astring$ = astring$ & Format$("LOG FE-OH ACTIVITY= ", a22$) & Format$(Format$(AFEOH!, f83$), a12$)
astring$ = astring$ & Format$("LOG MG-OH ACTIVITY= ", a22$) & Format$(Format$(AMGOH!, f83$), a12$) & vbCrLf
        
astring$ = astring$ & Format$("LOG FE-F ACTIVITY= ", a22$) & Format$(Format$(AFEF!, f83$), a12$)
astring$ = astring$ & Format$("LOG MG-F ACTIVITY= ", a22$) & Format$(Format$(AMGF!, f83$), a12$)
Call IOWriteLog(astring$)
        
' OUTPUT TO COLUMN FORMAT
astring$ = VbDquote$ & sample(1).number% & VbDquote$ & vbTab
For i% = 1 To MAXBIOT%
astring$ = astring$ & MiscAutoFormat$(WTPC!(i%)) & vbTab$
Next i%
For i% = 1 To MAXBIOT%
astring$ = astring$ & MiscAutoFormat$(ANSFO!(i%)) & vbTab$
Next i%
astring$ = astring$ & MiscAutoFormat$(HALOG1!) & vbTab$ & MiscAutoFormat$(MGMFT!) & vbTab$
astring$ = astring$ & MiscAutoFormat$(MGB!) & vbTab$ & MiscAutoFormat$(FEX!) & vbTab$
astring$ = astring$ & MiscAutoFormat$(tix!) & vbTab$ & MiscAutoFormat$(XALVI!) & vbTab$
astring$ = astring$ & MiscAutoFormat$(RAMGFE!) & vbTab$ & MiscAutoFormat$(MNX!) & vbTab$
astring$ = astring$ & MiscAutoFormat$(XFOXOH!) & vbTab$ & MiscAutoFormat$(HALMO!) & vbTab$
astring$ = astring$ & MiscAutoFormat$(HALMF!) & vbTab$ & MiscAutoFormat$(HALMC!) & vbTab$
astring$ = astring$ & MiscAutoFormat$(TETAL!) & vbTab$
astring$ = astring$ & MiscAutoFormat$(ANSFO!(1)) & vbTab$ & VbDquote$ & ATYPE$ & VbDquote$ & vbTab$
astring$ = astring$ & MiscAutoFormat$(sum1!) & vbTab$ & MiscAutoFormat$(sum2!) & vbTab$
astring$ = astring$ & MiscAutoFormat$(sum3!) & vbTab$ & MiscAutoFormat$(sum4!) & vbTab$
astring$ = astring$ & MiscAutoFormat$(AFEOH!) & vbTab$ & MiscAutoFormat$(AMGOH!) & vbTab$
astring$ = astring$ & MiscAutoFormat$(AFEF!) & vbTab$ & MiscAutoFormat$(AMGF!)
Print #Temp2FileNumber%, astring$

' ITER CALCULATIONS, SUM FLUORINE, CHLORINE, WATER
1900:   NOVOL = num(10) + num(15)
        If (NOVOL > 3.98) Then GoTo 2000
        iter = iter + 1
        NOWAT = 4# - num(10)
        NOWAT = NOWAT / (2 * NOAN)
        MTOT! = MTOT! + NOWAT - mp!(15)
        WTPC!(15) = 18# * NOWAT
        mp!(15) = NOWAT
        AP(15) = mp!(15)
        NOVOL = 0
        If (iter > 100) Then GoTo 2000
        GoTo 1400

' CALCULATE FINAL RESULTS
2000:   TOT! = TOT! + WTPC(15)
        temp! = (838# / (1.0337 - mp!(2) / mp!(4))) - 273
        XANN = mp!(4) / (mp!(5) + mp!(4))
        FOH = mp!(10) / (mp!(10) + mp!(15) * 2#)

        LOGOHF = 0#
        temp1 = 2 * mp!(15) / mp!(10)
        If temp1! > 0# Then LOGOHF! = MiscConvertLog10#(CDbl(temp1!))

' Output final results
astring$ = vbCrLf
astring$ = astring$ & Format$("VOL OCC", a12$) & Format$("H2O WT.%", a12$)
astring$ = astring$ & Format$("TOTAL", a12$) & Format$("ITER", a12$)
astring$ = astring$ & Format$("TEMP(C)", a12$) & Format$("X-ANN", a12$)
astring$ = astring$ & Format$("LOG OH/F", a12$) & Format$("F/F+OH", a12$)
Call IOWriteLog(astring$)

astring$ = vbNullString
astring$ = astring$ & Format$(Format$(NOVOL!, f83$), a12$) & Format$(Format$(WTPC!(MAXBIOT%), f83$), a12$)
astring$ = astring$ & Format$(Format$(TOT!, f83$), a12$) & Format$(Format$(iter%, i80$), a12$)
astring$ = astring$ & Format$(Format$(temp!, f83$), a12$) & Format$(Format$(XANN!, f83$), a12$)
astring$ = astring$ & Format$(Format$(LOGOHF!, f83$), a12$) & Format$(Format$(FOH!, f83$), a12$)
Call IOWriteLog(astring$)

Exit Sub

' Errors
ConvertBiotiteError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertBiotite"
Close #Temp2FileNumber%
ierror = True
Exit Sub

End Sub

Sub ConvertHalog(percents() As Single, sample() As TypeSample)
' v. 1.0; WRITTEN BY G. BRIMHALL
'
' v. 1.1; MODIFIED BY JAY AGUE (8-1-84) TO COMPUTE BIOTITE
' COMPONENT ACTIVITIES.  RECALCULATION PROCEDURE ALSO MODIFIED.
'
' v. 1.2; Further modifications to calculation procedure and
' output formats by Jay J. Ague 4/89. One of the more
' important changes has been the incorporation of code
' which rounds the values in the structural formula to
' 4 decimal places before any computations of mole fractions
' etc. takes place. This is to insure consistency between
' the printed (rounded) structural formula and computed
' quantities.
'
' v.1.3, 19 May 98; Jay J. Ague. Updated with atomic and
'   molecular weights from DHZ `92. Output format also modified.
'
' COMPUTES BIOTITE FORMULAS AND SITE OCCUPANCIES AND MOLE FRACTION RATIOS FOR USE IN HALOGEN
' CHEMISTRY AND CHARACTERIZATION OF MINERALIZATION ENVIRONMENTS

ierror = False
On Error GoTo ConvertHalogError

Const MAXBIOT% = 12
Const SSTRING1$ = "-----------------------------------------------------------"
Const SSTRING2$ = " ------ "

Dim IDEBUG As Integer, IBIG As Integer
Dim i As Integer, ip As Integer
     
Dim MNX As Single, MGMFT As Single, MGB As Single, YM32X As Single
Dim SIDER As Single, ANNIT As Single, PHLOG As Single, WTPCT As Single, WTT  As Single
Dim sum1 As Single, sum2 As Single, sum3 As Single, sum4 As Single
Dim WTPCF As Single, WTPCCL As Single, ATOPT As Single, TETAL As Single, OCTAL As Single
Dim XALVI As Single, ALKM As Single, FEX As Single, tix As Single
Dim HALMF As Single, HALMC As Single, HALMO As Single, XFOXOH As Single, HALOG1 As Single, SI As Single
Dim AFEF As Single, AMGF As Single, AFEOH As Single, AMGOH As Single, RAMGFE As Single
Dim XMG As Single, XSID As Single, XAN As Single

Dim d As Double

Dim astring As String, bstring As String

Dim WTPC(1 To MAXBIOT%) As Single, ATOP(1 To MAXBIOT%) As Single, ANSFO(1 To MAXBIOT%) As Single
Dim PMOL(1 To MAXBIOT%) As Single, ANIO(1 To MAXBIOT%) As Single

Dim esym(1 To MAXBIOT%) As String

' Print calculation
Call IOWriteLog(vbCrLf & "Biotite Formula Calculations (from Brimhal and Ague, v. 1.3, HALOG.F code)...")
      
' Load oxide percents
For i% = 1 To sample(1).LastChan%
    ip% = IPOS1%(MAXELM%, sample(1).Elsyms$(i%), Symlo$())
    If ip% <> 0 Then
        If ip% = 14 Then WTPC!(1) = percents!(i%)  ' SiO2
        If ip% = 22 Then WTPC!(2) = percents!(i%)  ' TiO2
        If ip% = 13 Then WTPC!(3) = percents!(i%)  ' Al2O3
        If ip% = 26 Then WTPC!(4) = percents!(i%)  ' FeO
        If ip% = 12 Then WTPC!(5) = percents!(i%)  ' MgO
        If ip% = 20 Then WTPC!(6) = percents!(i%)  ' CaO
        If ip% = 11 Then WTPC!(7) = percents!(i%)  ' Na2O
        If ip% = 56 Then WTPC!(8) = percents!(i%)  ' BaO
        If ip% = 19 Then WTPC!(9) = percents!(i%)  ' K2O
        If ip% = 9 Then WTPC!(10) = percents!(i%)  ' F
        If ip% = 17 Then WTPC!(11) = percents!(i%) ' Cl
        If ip% = 25 Then WTPC!(12) = percents!(i%) ' MnO
    End If
Next i%
        
IDEBUG = 0
IBIG = 0
WTPCT = 0#
WTT = 0#

For i% = 1 To MAXBIOT%
WTT = WTT + WTPC(i%)
Next i%
        
If DebugMode Then
Call IOWriteLog(vbCrLf & "Entered Biotite Analysis:")
For i% = 1 To MAXBIOT%
    If i% = 1 Then astring$ = Format$("SiO2", a80$)
    If i% = 2 Then astring$ = Format$("TiO2", a80$)
    If i% = 3 Then astring$ = Format$("Al2O3", a80$)
    If i% = 4 Then astring$ = Format$("FeO", a80$)
    If i% = 5 Then astring$ = Format$("MgO", a80$)
    If i% = 6 Then astring$ = Format$("MnO", a80$)
    If i% = 7 Then astring$ = Format$("CaO", a80$)
    If i% = 8 Then astring$ = Format$("Na2O", a80$)
    If i% = 9 Then astring$ = Format$("BaO", a80$)
    If i% = 10 Then astring$ = Format$("K2O", a80$)
    If i% = 11 Then astring$ = Format$("F", a80$)
    If i% = 12 Then astring$ = Format$("Cl", a80$)
    bstring$ = Format$(Format$(WTPC!(i%), f83$), a80)
Call IOWriteLog(bstring$ & a4x$ & astring$)
Next i%
Call IOWriteLog(SSTRING2$)
      
astring$ = Format$("TOTAL", a80$)
bstring$ = Format$(Format$(WTT!, f83$), a80)
Call IOWriteLog(bstring$ & a4x$ & astring$)
Call IOWriteLog(SSTRING1$)
End If

' Updated using DHZ '92
For i% = 1 To MAXBIOT%
      PMOL(1) = WTPC(1) / 60.08: esym$(1) = "Si"
      PMOL(2) = WTPC(2) / 79.88: esym$(2) = "Ti"
      PMOL(3) = WTPC(3) / 101.96: esym$(3) = "Al"
      PMOL(4) = WTPC(4) / 71.85: esym$(4) = "Fe+2"
      PMOL(5) = WTPC(5) / 40.3: esym$(5) = "Mg"
      PMOL(6) = WTPC(6) / 56.08: esym$(6) = "Ca"
      PMOL(7) = WTPC(7) / 61.98: esym$(7) = "Na"
      PMOL(8) = WTPC(8) / 153.33: esym$(8) = "Ba"
      PMOL(9) = WTPC(9) / 94.2: esym$(9) = "K"
      PMOL(10) = WTPC(10) / 19#: esym$(10) = "F"
      PMOL(11) = WTPC(11) / 35.45: esym$(11) = "Cl"
      PMOL(12) = WTPC(12) / 70.94: esym$(12) = "Mn"
      If i% <= 2 Then GoTo 102
      If i% = 3 Then GoTo 103
      If i% > 3 Then GoTo 104

' 2 OXYGENS
102:  ATOP(i%) = PMOL(i%) * 2#
      GoTo 105

' 3 OXYGENS
103:  ATOP(i%) = PMOL(i%) * 3#
      GoTo 105

' 1 OXYGENS
104:  ATOP(i%) = PMOL(i%) * 1#

' Total wt percents
105:  WTPCT = WTPCT + WTPC(i%)
Next i%
      
      WTPCF = WTPC(10) * 0.4211
      WTPCCL = WTPC(11) * 0.2256
      WTPCT = WTPCT - 1# * WTPCF - 1# * WTPCCL
      ATOPT = 0#
For i% = 1 To MAXBIOT%
      ATOPT = ATOPT + ATOP(i%)
Next i%
      ATOPT = ATOPT - 1# * ATOP(10) - 1# * ATOP(11)

' SPECIFY NUMBER OF OXYGENS, USE A TOTAL OF 22 NEGATIVE CHARGES
      d# = 11# / ATOPT

For i% = 1 To MAXBIOT%
      ANIO(i%) = d# * ATOP(i%)
Next i%
     
For i% = 1 To MAXBIOT%
      If i% <= 2 Then ANSFO(i%) = ANIO(i%) / 2#
      If i% = 3 Then ANSFO(i%) = ANIO(i%) * (2# / 3#)
      If i% > 3 And i% <= 6 Then ANSFO(i%) = ANIO(i%)
      If i% = 7 Or i% = 9 Then ANSFO(i%) = ANIO(i%) * 2#
      If i% = 8 Then ANSFO(i%) = ANIO(i%)
      If i% > 9 Then ANSFO(i%) = ANIO(i%)
Next i%

' Round values in the structural formula (array ANSFO) to four decimal places
For i% = 1 To MAXBIOT%
    ANSFO!(i%) = MiscSetRounding2!(ANSFO!(i%), Int(4))
Next i%

' TETRAHEDRAL AL
      TETAL = 4# - ANSFO(1)

' OCTAHEDRAL AL
      OCTAL = ANSFO(3) - TETAL
      If OCTAL < 0# Then TETAL = ANSFO(3)
      If OCTAL < 0# Then OCTAL = 0#
      XALVI = OCTAL / (ANSFO(5) + ANSFO(2) + OCTAL + ANSFO(4) + ANSFO(12))

' K20 + NA20 + BA0 + CA0
      ALKM = ANSFO(9) + ANSFO(7) + ANSFO(8) + ANSFO(6)

' X MG (FULL OCTAHEDRAL)
      MGMFT = ANSFO(5) / (ANSFO(5) + ANSFO(2) + OCTAL + ANSFO(4) + ANSFO(12))

' Mg / (Mg + FE)
      MGB = ANSFO(5) / (ANSFO(5) + ANSFO(4))

' X FE++ (FULL OCTAHEDRAL ANNITE)
      FEX = ANSFO(4) / (ANSFO(5) + ANSFO(2) + OCTAL + ANSFO(4) + ANSFO(12))

'  X TI-BIOTITE (FULL OCTAHEDRAL)
      tix = ANSFO(2) / (ANSFO(5) + ANSFO(2) + OCTAL + ANSFO(4) + ANSFO(12))

' X MN- BIOTITE (FULL OCTAHEDRAL)
      MNX = ANSFO(12) / (ANSFO(5) + ANSFO(2) + OCTAL + ANSFO(4) + ANSFO(12))

' X   F
      HALMF = ANSFO(10) / 2#

' X  CL
      HALMC = ANSFO(11) / 2#

' COMPUTE X-OH
      HALMO = 1# - HALMF - HALMC
      If HALMF = 0# Then HALMF = 0.00001
      If HALMC = 0# Then HALMC = 0.00001
      If HALMO <= 0# Then HALMO = 0.00001

' COMPUTE LOG X-F/X-OH
      If HALMF / HALMO > 0# Then XFOXOH = MiscConvertLog10#(CDbl(HALMF / HALMO))
      If HALMF / HALMC > 0# Then HALOG1 = MiscConvertLog10#(CDbl(HALMF / HALMC))

      sum1 = ANSFO(9) + ANSFO(7) + ANSFO(6) + ANSFO(8)
      sum2 = ANSFO(2) + OCTAL + ANSFO(4) + ANSFO(5) + ANSFO(12)
      sum3 = TETAL + ANSFO(1)
      sum4 = ANSFO(10) + ANSFO(11)

' BIOTITE COMPONENT ACTIVITIES
      SI = ANSFO(1)
      AFEF = MiscConvertLog10#(CDbl(ANSFO(9) * ((ANSFO(4) / 3#) ^ 3) * TETAL * ((SI / 3#) ^ 3) * (HALMF ^ 2)))
      AMGF = MiscConvertLog10#(CDbl(ANSFO(9) * ((ANSFO(5) / 3#) ^ 3) * TETAL * ((SI / 3#) ^ 3) * (HALMF ^ 2)))
      AFEOH = MiscConvertLog10#(CDbl(ANSFO(9) * ((ANSFO(4) / 3#) ^ 3) * TETAL * ((SI / 3#) ^ 3) * (HALMO ^ 2)))
      AMGOH = MiscConvertLog10#(CDbl(ANSFO(9) * ((ANSFO(5) / 3#) ^ 3) * TETAL * ((SI / 3#) ^ 3) * (HALMO ^ 2)))

      XMG = (ANSFO(5) / 3#)
      XSID = (((3# - ANSFO(1) / ANSFO(3)) / 1.75) * (1# - XMG))
      XAN = 1# - (XMG + XSID)
      PHLOG = XMG * 100#
      SIDER = XSID * 100#
      ANNIT = XAN * 100#

' COMPUTE Y-INTERCEPT
      RAMGFE = MiscConvertLog10#(CDbl(MGMFT / FEX))
      YM32X = XFOXOH - 1.5 * RAMGFE

Call IOWriteLog(SSTRING2$)
Call IOWriteLog("NUMBER OF ATOMS:")
Call IOWriteLog(a8x$ & "SI   " & Format$(Format$(ANSFO!(1), f84$), a80$))
Call IOWriteLog(a8x$ & "ALIV " & Format$(Format$(TETAL!, f84$), a80$) & a8x$ & "ALVI" & Format$(Format$(OCTAL!, f84$), a80$))
Call IOWriteLog(a8x$ & "TI   " & Format$(Format$(ANSFO!(2), f84$), a80$))
Call IOWriteLog(a8x$ & "FE   " & Format$(Format$(ANSFO!(4), f84$), a80$))
Call IOWriteLog(a8x$ & "MG   " & Format$(Format$(ANSFO!(5), f84$), a80$))
Call IOWriteLog(a8x$ & "MN   " & Format$(Format$(ANSFO!(12), f84$), a80$))
Call IOWriteLog(a8x$ & "CA   " & Format$(Format$(ANSFO!(6), f84$), a80$))
Call IOWriteLog(a8x$ & "NA   " & Format$(Format$(ANSFO!(7), f84$), a80$))
Call IOWriteLog(a8x$ & "BA   " & Format$(Format$(ANSFO!(8), f84$), a80$))
Call IOWriteLog(a8x$ & "K    " & Format$(Format$(ANSFO!(9), f84$), a80$))
Call IOWriteLog(a8x$ & "F    " & Format$(Format$(ANSFO!(10), f84$), a80$))
Call IOWriteLog(a8x$ & "CL   " & Format$(Format$(ANSFO!(11), f84$), a80$))
Call IOWriteLog(a8x$ & "OH   " & Format$(Format$(2# - (ANSFO(11) + ANSFO(10)), f84$), a80$) & "        CALCULATED")

Call IOWriteLog(vbCrLf & "SUMMARY OF BIOTITE GEOCHEMISTRY:")
      
' PRINT OUT BRIMHALL CALCULATIONS
astring$ = vbCrLf
astring$ = astring$ & Format$("LOG(XF/XCL) = ", a18$) & Format$(Format$(HALOG1!, f84$), a10$)
astring$ = astring$ & Format$("LOG(X-F/X-OH) = ", a18$) & Format$(Format$(XFOXOH!, f84$), a10$)
astring$ = astring$ & Format$("LOG(X-MG/X-FE) = ", a18$) & Format$(Format$(RAMGFE!, f84$), a10$)
Call IOWriteLog(astring$)

astring$ = vbNullString
astring$ = astring$ & Format$("X-MG = ", a18$) & Format$(Format$(MGMFT!, f84$), a10$)
astring$ = astring$ & Format$("X-FE = ", a18$) & Format$(Format$(FEX!, f84$), a10$)
astring$ = astring$ & Format$("X-TI = ", a18$) & Format$(Format$(tix!, f84$), a10$)
Call IOWriteLog(astring$)

astring$ = vbNullString
astring$ = astring$ & Format$("X-MN = ", a18$) & Format$(Format$(MNX!, f84$), a10$)
astring$ = astring$ & Format$("X-AL VI = ", a18$) & Format$(Format$(XALVI!, f84$), a10$)
astring$ = astring$ & Format$("MG/(MG+FE) = ", a18$) & Format$(Format$(MGB!, f84$), a10$)
Call IOWriteLog(astring$)
            
astring$ = vbCrLf
astring$ = astring$ & Format$("X-OH = ", a18$) & Format$(Format$(HALMO!, f84$), a10$)
astring$ = astring$ & Format$("X-F = ", a18$) & Format$(Format$(HALMF!, f84$), a10$)
astring$ = astring$ & Format$("X-CL = ", a18$) & Format$(Format$(HALMC!, f84$), a10$)
astring$ = astring$ & Format$("LOG (X-F/X-OH) = ", a18$) & Format$(Format$(XFOXOH!, f84$), a10$)
Call IOWriteLog(astring$)
            
astring$ = vbCrLf
astring$ = astring$ & Format$("TRIANGULAR PLOT        LOG X-F/X-OH  -1.5 * LOG X-MG/X-FE = ") & Format$(Format$(YM32X!, f84$), a10$)
Call IOWriteLog(astring$)
      
astring$ = vbCrLf
astring$ = astring$ & Format$("X-SID: = ", a18$) & Format$(Format$(SIDER!, f84$), a10$)
astring$ = astring$ & Format$("X-ANN: = ", a18$) & Format$(Format$(ANNIT!, f84$), a10$)
astring$ = astring$ & Format$("X-PHLOG: = ", a18$) & Format$(Format$(PHLOG!, f84$), a10$)
Call IOWriteLog(astring$)

astring$ = vbCrLf
astring$ = astring$ & Format$("(K+Na+Ca+Ba)", a22$) & Format$("(Ti+Al(VI)+Fe+Mg+Mn)", a22$) & Format$("(Al(IV)+Si)", a22$) & Format$("(F+CL)", a22$) & vbCrLf
astring$ = astring$ & Format$(Format$(sum1!, f84$), a22$) & Format$(Format$(sum2!, f84$), a22$) & Format$(Format$(sum3!, f84$), a22$) & Format$(Format$(sum4!, f84$), a22$)
Call IOWriteLog(astring$)
            
' WRITE OUTPUT FILE FOR HALOG.OUT (COMPOSITONAL FRAMES)
astring$ = "Sample " & vbTab & VbDquote$ & sample(1).number% & VbDquote$ & vbTab & VbDquote$ & sample(1).Name$ & VbDquote$
Print #Temp2FileNumber%, astring$

astring$ = vbNullString
For i% = 1 To MAXBIOT%
    If i% = 3 Then
        astring$ = astring$ + MiscAutoFormat$(WTPC!(i%)) & vbTab$ & MiscAutoFormat$(ANSFO!(i%)) & vbTab$ & MiscAutoFormat$(TETAL!) & vbTab$ & MiscAutoFormat$(OCTAL!) & vbTab & esym$(i%) & vbCrLf
    ElseIf i% = 11 Then
        astring$ = astring$ + MiscAutoFormat$(WTPC!(i%)) & vbTab$ & MiscAutoFormat$(ANSFO!(i%)) & vbTab$ & MiscAutoFormat$(2# - (ANSFO!(10) + ANSFO!(11))) & vbTab$ & esym$(i%) & vbCrLf
    Else
        astring$ = astring$ + MiscAutoFormat$(WTPC!(i%)) & vbTab$ & MiscAutoFormat$(ANSFO!(i%)) & vbTab$ & esym$(i%) & vbCrLf
    End If
Next i%
Print #Temp2FileNumber%, astring$

astring$ = vbNullString
astring$ = astring$ & MiscAutoFormat$(WTT!) & vbTab$ & MiscAutoFormat$(HALOG1!) & vbTab$ & MiscAutoFormat$(RAMGFE!) & vbTab$ & MiscAutoFormat$(MGMFT!) & vbTab$ & MiscAutoFormat$(FEX!) & vbTab$ & MiscAutoFormat$(tix!) & vbTab$ & MiscAutoFormat$(XALVI!) & vbTab$ & MiscAutoFormat$(MNX!) & vbTab$ & MiscAutoFormat$(XFOXOH!) & vbTab$
astring$ = astring$ & MiscAutoFormat$(SIDER!) & vbTab$ & MiscAutoFormat$(ANNIT!) & vbTab$ & MiscAutoFormat$(PHLOG!) & vbTab$ & MiscAutoFormat$(sum1!) & vbTab$ & MiscAutoFormat$(sum2!)
Print #Temp2FileNumber%, astring$
 
Exit Sub

' Errors
ConvertHalogError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertHalog"
Close #Temp2FileNumber%
ierror = True
Exit Sub

End Sub


