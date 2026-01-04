Attribute VB_Name = "CodeCONVERT6"
' (c) Copyright 1995-2026 by John J. Donovan
Option Explicit

Sub ConvertHalog(tfilenumber As Integer, percents() As Single, sample() As TypeSample)
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
Call IOWriteLog(vbCrLf & "Biotite Formula Calculations (from Brimhall and Ague, v. 1.3, HALOG.F code)...")
      
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
Print #tfilenumber%, astring$

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
Print #tfilenumber%, astring$

astring$ = vbNullString
astring$ = astring$ & MiscAutoFormat$(WTT!) & vbTab$ & MiscAutoFormat$(HALOG1!) & vbTab$ & MiscAutoFormat$(RAMGFE!) & vbTab$ & MiscAutoFormat$(MGMFT!) & vbTab$ & MiscAutoFormat$(FEX!) & vbTab$ & MiscAutoFormat$(tix!) & vbTab$ & MiscAutoFormat$(XALVI!) & vbTab$ & MiscAutoFormat$(MNX!) & vbTab$ & MiscAutoFormat$(XFOXOH!) & vbTab$
astring$ = astring$ & MiscAutoFormat$(SIDER!) & vbTab$ & MiscAutoFormat$(ANNIT!) & vbTab$ & MiscAutoFormat$(PHLOG!) & vbTab$ & MiscAutoFormat$(sum1!) & vbTab$ & MiscAutoFormat$(sum2!)
Print #tfilenumber%, astring$
 
Exit Sub

' Errors
ConvertHalogError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertHalog"
Close #tfilenumber%
ierror = True
Exit Sub

End Sub


