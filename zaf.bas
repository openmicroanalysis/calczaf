Attribute VB_Name = "CodeZAF"
' (c) Copyright 1995-2017 by John J. Donovan (credit to John Armstrong for original code)
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

' Declare ZAF structure
Dim zaf As TypeZAF

Const UseBrianJoyModifications = True
'Const UseBrianJoyModifications = False

' Module level ZAF arrays
Dim phi(1 To MAXCHAN1%) As Single
Dim pz(1 To MAXCHAN1%, 1 To MAXCHAN1%) As Single

Dim hb(1 To MAXCHAN1%) As Single    ' pure element backscatter coefficients
Dim eta(1 To MAXCHAN1%) As Single   ' sample backscatter coefficients
Dim jm(1 To MAXCHAN1%) As Single    ' mean ionization

' Used in PAP and Proza corrections
Dim zz As Single, zn As Single
Dim xp As Single, hh As Single
Dim xi As Single, meanz As Single, FP As Single
Dim sumatom As Single, jbar As Single

Dim dp(1 To 3) As Single, pp(1 To 3) As Single

Dim em(1 To MAXCHAN1%) As Single
Dim zipi(1 To MAXCHAN1%) As Single

Dim FormulaTmpSample(1 To 1) As TypeSample

Sub ZAFAbs(zafinit As Integer)
' When iabs% equals:
' 1 = "Absorption of Philibert (FRAME)"
' 2 = "Absorption of Heinrich (Quadratic Anal. Chem.)"
' 3 = "Absorption of Heinrich (Duplex 1989 MAS)"
' 4 = "Absorption of Love/Scott (1983 J. Phys. D.)"
' 5 = "Absorption of Sewell/Love/Scott (1985-I J. Phys. D.)"
' 6 = "Absorption of Sewell/Love/Scott (1985-II J. Phys. D.)"
' 7 = "Phi(PZ) Absorption of Packwood/Brown 1982/XRS Alpha"
' 8 = "Phi(PZ) Absorption of Bastin 1984/XRS Alpha"
' 9 = "Phi(PZ) Absorption of Armstrong (Packwood/Brown) 1981 MAS"
'10 = "Phi(PZ) Absorption of Bastin 1986/Scanning"
'11 = "Phi(PZ) Absorption of Riveros 1987/XRS"
'12 = "Phi(PZ) Absorption of Pouchou & Pichoir (Full)"
'13 = "Phi(PZ) Absorption of Pouchou and Pichoir (Simplified)"
'14 = "Phi(PZ) Absorption of Packwood (New)"
'15 = "Phi(PZ) Absorption of Bastin Proza (EPQ-91)"

ierror = False
On Error GoTo ZAFAbsError

Dim a1 As Single, a2 As Single
Dim i As Integer, i1 As Integer

Dim ax As Single, ha As Single, hal As Single
Dim hx As Single, a3 As Single, xxx As Single
Dim lu As Single, ppz As Single, ps As Single

Dim zed As Single, zm As Single, zr As Single, zra As Single
Dim zrb As Single, zrc As Single
Dim m5 As Single, m6 As Single, m7 As Single

Dim avez As Single

' Preserve values for "zafinit" modes
Static f1(1 To MAXCHAN1%) As Single
Static h(1 To MAXCHAN1%) As Single

' Sample calculations
If zafinit% = 1 Then GoTo 7200

' STDABS1 / PHILIBERT ABSORPTION CORRECTION FOR STANDARDS
If iabs% = 1 Then
For i% = 1 To zaf.in0%
f1!(i%) = 1.2 * zaf.atwts!(i%) / zaf.Z%(i%) ^ 2
Next i%

For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
h!(i%) = 450000# / (zaf.eO!(i%) ^ 1.65 - zaf.eC!(i%) ^ 1.65)
zaf.genstd!(i%) = 1# + f1!(i%) / (1# + f1!(i%)) * zaf.mup!(i%, i%) * zaf.m1!(i%) / h!(i%)
zaf.genstd!(i%) = zaf.genstd!(i%) * (1# + zaf.mup!(i%, i%) * zaf.m1!(i%) / h!(i%))
End If
Next i%

' STDABS2 / HEINRICH/AN. CHEM. ABSORPTION CORRECTION FOR STANDARDS
ElseIf iabs% = 2 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
h!(i%) = 0.0000012 * (zaf.eO!(i%) ^ 1.65 - zaf.eC!(i%) ^ 1.65)
zaf.genstd!(i%) = (1# + h!(i%) * zaf.mup!(i%, i%) * zaf.m1!(i%)) ^ 2
End If
Next i%

' STDABS3 / HEINRICH/1989 MAS ABSORPTION CORRECTION FOR STANDARDS
ElseIf iabs% = 3 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
h!(i%) = zaf.eO!(i%) ^ 1.65 - zaf.eC!(i%) ^ 1.65
ha! = 0.00000165
hal! = 0.18 + 2# / h!(i%) + 0.000008 * zaf.eC!(i%) + 0.005 * Sqr(zaf.Z%(i%))
zaf.genstd!(i%) = (1# + ha! * h!(i%) * zaf.mup!(i%, i%) * zaf.m1!(i%)) ^ 2 / (1# + hal! * (ha! * h!(i%) * zaf.mup!(i%, i%) * zaf.m1!(i%)))
End If
Next i%

' STDABS4 / LOVE/SCOTT ABSORPTION CORRECTION FOR STANDARDS
ElseIf iabs% = 4 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
ps! = (0.000007870001 * jm!(i%) ^ 0.5 * zaf.eO!(i%) ^ 1.5 + 0.000000735 * zaf.eO!(i%) * zaf.eO!(i%)) / (zaf.Z%(i%) / zaf.atwts!(i%))
ppz! = ps! * (0.49269 - 1.0987 * hb!(i%) + 0.78557 * hb!(i%) * hb!(i%)) * Log(zaf.v!(i%))
ppz! = ppz! / (0.70256 - 1.09865 * hb!(i%) + 1.0046 * hb!(i%) * hb!(i%) + Log(zaf.v!(i%)))
zaf.genstd!(i%) = (1# - Exp(-2# * zaf.mup!(i%, i%) * zaf.m1!(i%) * ppz!)) / (2# * zaf.mup!(i%, i%) * zaf.m1!(i%) * ppz!)
zaf.genstd!(i%) = 1# / zaf.genstd!(i%)
End If
Next i%

' STDABS5 / LOVE/SCOTT I ABSORPTION CORRECTION FOR STANDARDS
ElseIf iabs% = 5 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
For i1% = 1 To zaf.in0%
ps! = (0.000007870001 * jm!(i1%) ^ 0.5 * zaf.eO!(i%) ^ 1.5 + 0.000000735 * zaf.eO!(i%) * zaf.eO!(i%)) / (zaf.Z%(i1%) / zaf.atwts!(i1%))
lu! = Log(zaf.v!(i%))
pz!(i%, i1%) = ps! * lu! / ((2.4 + 0.07 * zaf.Z%(i1%)) * lu! + 1.04 + 0.48 * hb!(i1%)) ' use pure element backscatter
Next i1%

If i% < zaf.in0 Or zaf.in1% = zaf.in0 Then
zaf.genstd!(i%) = (1# - Exp(-2# * zaf.mup!(i%, i%) * zaf.m1!(i%) * pz!(i%, i%))) / (2# * zaf.mup!(i%, i%) * zaf.m1!(i%) * pz!(i%, i%))
zaf.genstd!(i%) = 1# / zaf.genstd!(i%)
End If
End If
Next i%

' STDABS6 / LOVE/SCOTT II ABSORPTION CORRECTION FOR STANDARDS
ElseIf iabs% = 6 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
For i1% = 1 To zaf.in0%
lu! = Log(zaf.v!(i%))
ps! = (0.000007870001 * jm!(i1%) ^ 0.5 * zaf.eO!(i%) ^ 1.5 + 0.000000735 * zaf.eO!(i%) * zaf.eO!(i%)) / (zaf.Z%(i1%) / zaf.atwts!(i1%))
pz!(i%, i1%) = ps! * lu! / ((2.4 + 0.07 * zaf.Z%(i1%)) * lu! + 1.04 + 0.48 * hb!(i1%))
Next i1%

If i% < zaf.in0 Or zaf.in1% = zaf.in0 Then
a1! = 2.2 + 0.00188 * zaf.Z%(i%)
a3! = 0.01 + 0.00719 * zaf.Z%(i%)
a2! = (a1! - 1#) * Exp(a3!)
ax! = 1.23 - 1.25 * hb!(i%)
hx! = a1! - a2! * Exp(-a3! * zaf.v!(i%) ^ ax!)

zm! = pz!(i%, i%) * (0.29 + (0.662 + 0.443 * zaf.v!(i%) ^ 0.2) / Sqr(zaf.Z%(i%)))
zra! = hx!
zrb! = hx! * (zm! - 3# * pz!(i%, i%))
zrc! = zm! * (zm! - 3# * pz!(i%, i%))
zr! = (-zrb! + Sqr(zrb! * zrb! - 4# * zra! * zrc!)) / (2# * zra!)
xxx! = zaf.mup!(i%, i%) * zaf.m1!(i%)
zaf.genstd!(i%) = (Exp(-xxx! * zm!) * (zr! - hx! * zr!) + hx! * zr! - zr!) / zm!
zaf.genstd!(i%) = zaf.genstd!(i%) - Exp(-xxx! * zm!) + hx! * Exp(-xxx! * zr) + xxx! * (zr! - zm!) - hx! + 1#
zaf.genstd!(i%) = zaf.genstd!(i%) * 2# / ((zr! - zm!) * (zm! + hx! * zr!) * xxx! ^ 2)
zaf.genstd!(i%) = 1# / zaf.genstd!(i%)
End If
End If
Next i%

' Phi-rho-z calculations for pure elements
ElseIf iabs% = 7 Or iabs% = 8 Or iabs% = 9 Or iabs% = 10 Or iabs% = 11 Then
Call ZAFPhiCal(zafinit%)
If ierror Then Exit Sub

' STDABS12 / POUCHOU and PICHOIR (Full) for Pure Elements
ElseIf iabs% = 12 Then
dp!(1) = 0.0000066
pp!(1) = 0.78
pp!(2) = 0.1
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
zz! = zaf.Z%(i%)
zn! = zaf.Z%(i%)
xp! = zaf.r!(i%, i%) / zaf.s!(i%, i%)
jbar! = jm!(i%)
hh! = hb!(i)
xi! = zaf.mup!(i%, i%) * zaf.m1!(i%)
sumatom! = zaf.Z%(i%) / zaf.atwts!(i%)

' Calculate PAP absorption
Call ZAFPap(Int(1), i%)
If ierror Then Exit Sub

zaf.genstd!(i%) = 1# / FP!
End If
Next i%

' STDABS13 / POUCHOU and PICHOIR (Simplified) for Pure Elements
ElseIf iabs% = 13 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
zz! = zaf.Z%(i%)
zn! = zaf.Z%(i%)
meanz! = zaf.Z%(i%)
xp! = zaf.r!(i%, i%) / zaf.s!(i%, i%)
jbar! = jm!(i%)
hh! = hb!(i%)
xi! = zaf.mup!(i%, i%) * zaf.m1!(i%)
sumatom! = 1#
        
' Calculate PAP absorption
Call ZAFPap(Int(2), i%)
If ierror Then Exit Sub
        
zaf.genstd!(i%) = 1# / FP!
End If
Next i%

' STDABS14 / Packwood (New) for Pure Elements
ElseIf iabs% = 14 Then
Call ZAFPhiCal(zafinit%)
If ierror Then Exit Sub
  
' STDABS15 / Bastin Proza for Pure Elements
ElseIf iabs% = 15 Then
For i% = 1 To zaf.in1%
For i1% = 1 To zaf.in0%
If zaf.il%(i%) <= MAXRAY% - 1 Then
pz!(i%, i1%) = 216140# * zaf.Z%(i1%) ^ 1.163 / ((zaf.eO!(i%) / zaf.eC!(i%) - 1) ^ 0.5 * zaf.eO!(i%) ^ 1.25 * zaf.atwts!(i1%))
If Not UseBrianJoyModifications Then
pz!(i%, i1%) = pz!(i%, i1%) * (Log(1.166 * zaf.eO!(i%) / jm!(i%)) / zaf.eC!(i%)) ^ 0.5       ' original CITZAF code
Else
pz!(i%, i1%) = pz!(i%, i1%) * (Log(1.166 * zaf.eO!(i%) / jm!(i1%)) / zaf.eC!(i%)) ^ 0.5      ' corrected by Brian Joy (02-2016)
End If
End If
Next i1%
Next i%

Call ZAFPhiCal(zafinit%)
If ierror Then Exit Sub
End If

Exit Sub
  
' SMPABS1 / PHILIBERT ABSORPTION CORRECTION FOR SAMPLE
7200:
If iabs% = 1 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
hh! = 0#
For i1% = 1 To zaf.in0%
hh! = hh! + zaf.conc!(i1%) * f1!(i1%)
Next i1%
m7! = ZAFMACCal(i%, zaf)
zaf.gensmp!(i%) = (1# + hh! / (1# + hh!) * m7! * zaf.m1!(i%) / h!(i%)) * (1# + m7! * zaf.m1!(i%) / h!(i%))
End If
Next i%

' SMPABS2 / HEINRICH/AN. CHEM. ABSORPTION CORRECTION FOR SAMPLE
ElseIf iabs% = 2 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
m7! = ZAFMACCal(i%, zaf)
zaf.gensmp!(i%) = (1# + h!(i%) * m7! * zaf.m1!(i%)) ^ 2
End If
Next i%

' SMPABS3 / HEINRICH/1989 MAS ABSORPTION CORRECTION FOR SAMPLE
ElseIf iabs% = 3 Then
ha! = 0.00000165
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
zed! = 0#
For i1% = 1 To zaf.in0%
zed! = zed! + zaf.conc!(i1%) * zaf.Z%(i1%)
Next i1%
hal! = 0.18 + 2# / h!(i%) + 0.000008 * zaf.eC!(i%) + 0.005 * Sqr(zed!)
m7! = ZAFMACCal(i%, zaf)
zaf.gensmp!(i%) = (1# + ha! * h!(i%) * m7! * zaf.m1!(i%)) ^ 2 / (1# + hal! * (ha! * h!(i%) * m7! * zaf.m1!(i%)))
End If
Next i%

' SMPABS4 / LOVE/SCOTT ABSORPTION CORRECTION FOR SAMPLE
ElseIf iabs% = 4 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
m5! = 0#
m6! = 0#
        
For i1% = 1 To zaf.in0%
m5! = m5! + zaf.conc!(i1%) * zaf.Z%(i1%) / zaf.atwts!(i1%)
m6! = m6! + zaf.conc!(i1%) * zaf.Z%(i1%) * Log(jm!(i1%)) / zaf.atwts!(i1%)
Next i1%
        
m6! = Exp(m6! / m5!)
m7! = ZAFMACCal(i%, zaf)
ps! = (0.000007870001 * m6! ^ 0.5 * zaf.eO!(i%) ^ 1.5 + 0.000000735 * zaf.eO!(i%) * zaf.eO!(i%)) / m5!
ppz! = ps! * (0.49269 - 1.0987 * eta!(i%) + 0.78557 * eta!(i%) * eta!(i%)) * Log(zaf.v!(i%))
ppz! = ppz! / (0.70256 - 1.09865 * eta!(i%) + 1.0046 * eta!(i%) * eta!(i%) + Log(zaf.v!(i%)))
zaf.gensmp!(i%) = (1# - Exp(-2# * m7! * zaf.m1!(i%) * ppz!)) / (2# * m7! * zaf.m1!(i%) * ppz!)
zaf.gensmp!(i%) = 1# / zaf.gensmp!(i%)
End If
Next i%

' SMPABS5 / LOVE/SCOTT I ABSORPTION CORRECTION FOR SAMPLE
ElseIf iabs% = 5 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
ppz! = 0#

For i1% = 1 To zaf.in0%
ppz! = ppz! + zaf.conc!(i1%) * pz!(i%, i1%)
Next i1%

m7! = ZAFMACCal(i%, zaf)
zaf.gensmp!(i%) = (1# - Exp(-2# * m7! * zaf.m1!(i%) * ppz!)) / (2# * m7! * zaf.m1!(i%) * ppz!)
zaf.gensmp!(i%) = 1# / zaf.gensmp!(i%)
End If
Next i%

' SMPABS6 / LOVE/SCOTT II ABSORPTION CORRECTION FOR SAMPLE
ElseIf iabs% = 6 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
m7! = ZAFMACCal(i%, zaf)
xxx! = m7! * zaf.m1!(i%)
m5! = 0#
ppz! = 0#

For i1% = 1 To zaf.in0%
m5! = m5! + zaf.conc!(i1%) * zaf.Z%(i1%)
ppz! = ppz! + zaf.conc!(i1%) * pz!(i%, i1%)
Next i1%

a1! = 2.2 + 0.00188 * m5!
a3! = 0.01 + 0.00719 * m5!
a2! = (a1! - 1#) * Exp(a3!)
ax! = 1.23 - 1.25 * eta!(i%)
hx! = a1! - a2! * Exp(-a3! * zaf.v!(i%) ^ ax!)

zm! = ppz! * (0.29 + (0.662 + 0.443 * zaf.v!(i%) ^ 0.2) / Sqr(m5!))
zra! = hx!
zrb! = hx! * (zm! - 3# * ppz!)
zrc! = zm! * (zm! - 3# * ppz!)
zr! = (-zrb! + Sqr(zrb! * zrb! - 4# * zra! * zrc!)) / (2# * zra!)
zaf.gensmp!(i%) = (Exp(-xxx! * zm!) * (zr! - hx! * zr!) + hx! * zr! - zr!) / zm!
zaf.gensmp!(i%) = zaf.gensmp!(i%) - Exp(-xxx! * zm!) + hx! * Exp(-xxx! * zr!) + xxx! * (zr! - zm!) - hx! + 1#
zaf.gensmp!(i%) = zaf.gensmp!(i%) * 2# / ((zr! - zm!) * (zm! + hx! * zr!) * xxx! ^ 2)
zaf.gensmp!(i%) = 1# / zaf.gensmp!(i%)
End If
Next i%

' Phi-rho-z calculations for pure elements
ElseIf iabs% = 7 Or iabs% = 8 Or iabs% = 9 Or iabs% = 10 Or iabs% = 11 Then
Call ZAFPhiCal(zafinit%)
If ierror Then Exit Sub

' SMPABS12 / POUCHOU and PICHOIR (Full) for Sample
ElseIf iabs% = 12 Then
dp!(1) = 0.0000066
pp!(1) = 0.78
pp!(2) = 0.1
zz! = 0#
zn! = 0#

For i1% = 1 To zaf.in0%
zz! = zz! + zaf.conc!(i1%) * zaf.Z%(i1%)
zn! = zn! + zaf.conc!(i1%) * Log(zaf.Z%(i1%))
Next i1%

If zz! < 0# Then GoTo ZAFAbsBadZZ

zn! = Exp(zn!)
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
hh! = eta!(i%)
xp! = zaf.bks!(i%) / zaf.stp!(i%)
m7! = ZAFMACCal(i%, zaf)
xi! = m7! * zaf.m1!(i%)
        
' Calculate PAP absorption for sample
Call ZAFPap(Int(1), i%)
If ierror Then Exit Sub

zaf.gensmp!(i%) = 1# / FP!
End If
Next i%

' SMPABS13 / POUCHOU and PICHOIR (Simplified) for Sample
ElseIf iabs% = 13 Then
zz! = 0#
zn! = 0#
avez! = 0#
For i1% = 1 To zaf.in0%
zz! = zz! + zaf.conc!(i1%) * zaf.Z%(i1%)
zn! = zn! + zaf.conc!(i1%) * Log(zaf.Z%(i1%))
avez! = avez! + zaf.conc!(i1%) * Sqr(zaf.Z%(i1%))
Next i1%

If zz! < 0# Then GoTo ZAFAbsBadZZ

zn! = Exp(zn!)
meanz! = avez! * avez!

For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
hh! = eta!(i%)
xp! = zaf.bks!(i%) / zaf.stp!(i%)
m7! = ZAFMACCal(i%, zaf)
xi! = m7! * zaf.m1!(i%)

' Calculate PAP absorption for sample
Call ZAFPap(Int(2), i%)
If ierror Then Exit Sub
        
zaf.gensmp!(i%) = 1# / FP!
End If
Next i%

' SMPABS14 / Packwood (New) for Sample
ElseIf iabs% = 14 Then
Call ZAFPhiCal(zafinit%)
If ierror Then Exit Sub

' SMPABS15 / Bastin Proza for Sample
ElseIf iabs% = 15 Then
Call ZAFPhiCal(zafinit%)
If ierror Then Exit Sub
End If

Exit Sub

' Errors
ZAFAbsError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFAbs"
ierror = True
Exit Sub

ZAFAbsBadZZ:
msg$ = "Bad (negative) zz parameter calculated for the sample analysis. This usually indicates negative concentrations so you should check that you are not analyzing epoxy." & vbCrLf & vbCrLf
msg$ = msg$ & "You should also make sure your off-peak background and interference corrections are not overcorrecting, or perhaps you have assigned a blank correction to a major or minor element and you did not enter the correct blank level in the Standard Assignments dialog."
If Not CalcImageQuantFlag Then
MiscMsgBoxTim FormMSGBOXTIME, "ZAFAbs", msg$, 20#
Call IOWriteLog(msg$)
Else
Call IOWriteLog(msg$)
End If
'ierror = True
Exit Sub

End Sub

Sub ZAFAFactor(wout() As Single, rout() As Single, eout() As String, xout() As String, zout() As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' Calculate the k-ratios for the passed concentrations for alpha-factor calculations

ierror = False
On Error GoTo ZAFAFactorError

Dim i As Integer
Dim n As Integer

ReDim amount(1 To MAXBINARY%) As Single

' Calculate MAXBIN compositions (see GLOBAL.BAS)
For i% = 1 To MAXBINARY%
amount!(i%) = BinaryRanges!(i%) / 100#   ' for concentration
Next i%

' Check that zaf.in0% = 2 (binary sample)
If zaf.in0% <> 2 Then GoTo ZAFAFactorBadBinary

' Calculate for each binary composition
For n% = 1 To MAXBINARY%
If Not UsePenepmaKratiosLimitFlag Or (UsePenepmaKratiosLimitFlag And amount!(n%) < PenepmaKratiosLimitValue! / 100#) Then
zaf.n8& = n%

zaf.ksum! = 1#     ' must sum to 1.000
zaf.krat!(1) = amount!(n%)
zaf.krat!(2) = CSng(1# - amount!(n%))
zaf.conc!(1) = zaf.krat!(1)
zaf.conc!(2) = zaf.krat!(2)

zaf.kraw!(1) = 0#
zaf.kraw!(2) = 0#

zaf.AtPercents!(1) = 0#
zaf.AtPercents!(2) = 0#

zaf.Formulas!(1) = 0#
zaf.Formulas!(2) = 0#

' Calculate ZAFCORS and standard K-factors
Call ZAFMip(Int(1))
If ierror Then Exit Sub

Call ZAFBsc(Int(1))
If ierror Then Exit Sub

If istp% = 6 Then
Call ZAFAbs(Int(1))
If ierror Then Exit Sub
Call ZAFStp(Int(1))
If ierror Then Exit Sub
Call ZAFBks(Int(1))
If ierror Then Exit Sub

Else
Call ZAFStp(Int(1))
If ierror Then Exit Sub
Call ZAFBks(Int(1))
If ierror Then Exit Sub
Call ZAFAbs(Int(1))
If ierror Then Exit Sub
End If

' Calculate fluorescence correction
If iflu% < 5 Then
Call ZAFFlu(Int(1), zaf)
If ierror Then Exit Sub
Else
Call ZAFFlu3(Int(1), zaf)
If ierror Then Exit Sub
End If

' Calculate the atomic number correction
If zaf.il%(1) <= MAXRAY% - 1 And zaf.krat!(1) > 0# Then
zaf.stp!(1) = zaf.stp!(1) / zaf.s!(1, 1)
zaf.bks!(1) = zaf.r!(1, 1) / zaf.bks!(1)
zaf.zed!(1) = zaf.stp!(1) * zaf.bks!(1)
End If

If zaf.il%(2) <= MAXRAY% - 1 And zaf.krat!(2) > 0# Then
zaf.stp!(2) = zaf.stp!(2) / zaf.s!(2, 2)
zaf.bks!(2) = zaf.r!(2, 2) / zaf.bks!(2)
zaf.zed!(2) = zaf.stp!(2) * zaf.bks!(2)
End If

' Calculate the K-ratio for all emitting lines
If zaf.krat!(1) > 0# And zaf.il%(1) <= MAXRAY% - 1 Then
zaf.krat!(1) = zaf.conc!(1) / zaf.zed!(1) * (1# + zaf.vv!(1)) * zaf.genstd!(1) / zaf.gensmp!(1)
End If
If zaf.krat!(2) > 0# And zaf.il%(2) <= MAXRAY% - 1 Then
zaf.krat!(2) = zaf.conc!(2) / zaf.zed!(2) * (1# + zaf.vv!(2)) * zaf.genstd!(2) / zaf.gensmp!(2)
End If

' Print out the calculation
If DebugMode Then
'Call ZAFPrintCalculate(zaf, analysis, sample())
'If ierror Then Exit Sub
Call ZAFPrintSmp(zaf, analysis, CInt(0))
If ierror Then Exit Sub
End If

' Load weight percents and k-ratios for fit
wout!(2 * n% - 1) = 100# * zaf.conc!(1)
wout!(2 * n%) = 100# * zaf.conc!(2)

rout!(2 * n% - 1) = zaf.krat!(1)
rout!(2 * n%) = zaf.krat!(2)

' Load symbols for print out
eout$(2 * n% - 1) = sample(1).Elsyms$(1)
eout$(2 * n%) = sample(1).Elsyms$(2)

xout$(2 * n% - 1) = sample(1).Xrsyms$(1)
xout$(2 * n%) = sample(1).Xrsyms$(2)

' Load atomic numbers for look-up tables
zout%(2 * n% - 1) = sample(1).AtomicNums%(1)
zout%(2 * n%) = sample(1).AtomicNums%(2)

End If
Next n%

Exit Sub

' Errors
ZAFAFactorError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFAFactor"
ierror = True
Exit Sub

ZAFAFactorBadBinary:
msg$ = "Binary sample does not contain two elements"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFAFactor"
ierror = True
Exit Sub

End Sub

Sub ZAFAFactorOxide(wout() As Single, rout() As Single, eout() As String, xout() As String, zout() As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' Calculate the k-ratios for the passed oxide end-member concentrations for oxide alpha-factor calculations

ierror = False
On Error GoTo ZAFAFactorOxideError

ReDim atoms(1 To 2) As Single
ReDim amount(1 To 2) As Single
ReDim syms(1 To 2) As String

' Calculate MAXBIN compositions
syms$(1) = sample(1).Elsyms$(1)
syms$(2) = sample(1).Elsyms$(2)
If sample(1).AtomicNums%(1) <> 8 Then
atoms!(1) = AllCat%(sample(1).AtomicNums%(1))
atoms!(2) = AllOxd%(sample(1).AtomicNums%(1))

Else
atoms!(1) = AllOxd%(sample(1).AtomicNums%(2))
atoms!(2) = AllCat%(sample(1).AtomicNums%(2))
End If

amount!(1) = ConvertAtomToWeight!(Int(2), Int(1), atoms!(), syms$()) / 100#
amount!(2) = ConvertAtomToWeight!(Int(2), Int(2), atoms!(), syms$()) / 100#

' Check that zaf.in0% = 2 (binary sample)
If zaf.in0% <> 2 Then GoTo ZAFAFactorOxideBadBinary

' Calculate for each binary composition
zaf.n8& = 0

zaf.ksum! = 1#     ' must sum to 1.000
zaf.krat!(1) = amount!(1)
zaf.krat!(2) = CSng(1# - amount!(1))
zaf.conc!(1) = zaf.krat!(1)
zaf.conc!(2) = zaf.krat!(2)

' Calculate ZAFCORS and standard K-factors
Call ZAFMip(Int(1))
If ierror Then Exit Sub

Call ZAFBsc(Int(1))
If ierror Then Exit Sub

If istp% = 6 Then
Call ZAFAbs(Int(1))
If ierror Then Exit Sub
Call ZAFStp(Int(1))
If ierror Then Exit Sub
Call ZAFBks(Int(1))
If ierror Then Exit Sub

Else
Call ZAFStp(Int(1))
If ierror Then Exit Sub
Call ZAFBks(Int(1))
If ierror Then Exit Sub
Call ZAFAbs(Int(1))
If ierror Then Exit Sub
End If

' Calculate fluorescence correction
If iflu% < 5 Then
Call ZAFFlu(Int(1), zaf)
If ierror Then Exit Sub
Else
Call ZAFFlu3(Int(1), zaf)
If ierror Then Exit Sub
End If

' Calculate the atomic number correction
If zaf.il%(1) <= MAXRAY% - 1 And zaf.krat!(1) > 0# Then
zaf.stp!(1) = zaf.stp!(1) / zaf.s!(1, 1)
zaf.bks!(1) = zaf.r!(1, 1) / zaf.bks!(1)
zaf.zed!(1) = zaf.stp!(1) * zaf.bks!(1)
End If

If zaf.il%(2) <= MAXRAY% - 1 And zaf.krat!(2) > 0# Then
zaf.stp!(2) = zaf.stp!(2) / zaf.s!(2, 2)
zaf.bks!(2) = zaf.r!(2, 2) / zaf.bks!(2)
zaf.zed!(2) = zaf.stp!(2) * zaf.bks!(2)
End If

' Calculate the K-ratio for all emitting lines
If zaf.krat!(1) > 0# And zaf.il%(1) <= MAXRAY% - 1 Then
zaf.krat!(1) = zaf.conc!(1) / zaf.zed!(1) * (1# + zaf.vv!(1)) * zaf.genstd!(1) / zaf.gensmp!(1)
End If
If zaf.krat!(2) > 0# And zaf.il%(2) <= MAXRAY% - 1 Then
zaf.krat!(2) = zaf.conc!(2) / zaf.zed!(2) * (1# + zaf.vv!(2)) * zaf.genstd!(2) / zaf.gensmp!(2)
End If

' Print out the calculation
If DebugMode Then
'Call ZAFPrintCalculate(zaf, analysis, sample())
'If ierror Then Exit Sub
Call ZAFPrintSmp(zaf, analysis, CInt(0))
If ierror Then Exit Sub
End If

' Load weight percents and k-ratios for fit
wout!(1) = 100# * zaf.conc!(1)
wout!(2) = 100# * zaf.conc!(2)

rout!(1) = zaf.krat!(1)
rout!(2) = zaf.krat!(2)

' Load symbols for print out
eout$(1) = sample(1).Elsyms$(1)
eout$(2) = sample(1).Elsyms$(2)

xout$(1) = sample(1).Xrsyms$(1)
xout$(2) = sample(1).Xrsyms$(2)

' Load atomic numbers for look-up tables
zout%(1) = sample(1).AtomicNums%(1)
zout%(2) = sample(1).AtomicNums%(2)

Exit Sub

' Errors
ZAFAFactorOxideError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFAFactorOxide"
ierror = True
Exit Sub

ZAFAFactorOxideBadBinary:
msg$ = "Binary sample does not contain two elements"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFAFactorOxide"
ierror = True
Exit Sub

End Sub

Sub ZAFBsc(zafinit As Integer)
' When ibsc equals:
' 1 = "Backscatter Coefficient of Heinrich"
' 2 = "Backscatter Coefficient of Love/Scott"
' 3 = "Backscatter Coefficient of Pouchou & Pichoir"
' 4 = "Backscatter Coefficient of Hungler/Kuchler (August/Wernisch Mod.)"

ierror = False
On Error GoTo ZAFBscError

Dim i As Integer, i1 As Integer
Dim yy As Single, h1 As Single, h2 As Single
Dim Zbar As Single
        
If zafinit% = 1 Then GoTo 6900

' BSC1 / CALCULATION OF HEINRICH BACKSCATTER COEFFICIENTS FOR PURE ELEMENTS
If ibsc% = 1 Then
For i% = 1 To zaf.in0%
yy! = zaf.Z%(i%)
If zaf.eO!(i%) < 11# Then GoTo 6630
If zaf.eO!(i%) < 16# Then GoTo 6600
If zaf.eO!(i%) < 21# Then GoTo 6570
h1! = -0.01629 + 0.009371999 * yy! + 0.0004598 * yy! ^ 2 - 0.0000249 * yy! ^ 3
h1! = h1! + 0.0000004944 * yy! ^ 4 - 0.000000004478 * yy! ^ 5 + 0.0000000000153 * yy! ^ 6
GoTo 6650

6570:
h1! = -0.01392 + 0.01104 * yy! + 0.0003201 * yy! ^ 2 - 0.00001848 * yy! ^ 3
h1! = h1! + 0.000000336 * yy! ^ 4 - 0.00000000268 * yy! ^ 5 + 0.000000000007907 * yy! ^ 6
GoTo 6650

6600:
h1! = -0.01322 + 0.01191 * yy! + 0.0002676 * yy! ^ 2 - 0.00001765 * yy! ^ 3
h1! = h1! + 0.0000003426 * yy! ^ 4 - 0.000000002931 * yy! ^ 5 + 0.000000000009364 * yy! ^ 6
GoTo 6650

6630:
h1! = -0.007447 + 0.0113 * yy! + 0.0003298 * yy! ^ 2 - 0.00002045 * yy! ^ 3
h1! = h1! + 0.0000003994 * yy! ^ 4 - 0.000000003468 * yy! ^ 5 + 0.00000000001128 * yy! ^ 6

6650:
If h1! <= 0# Then h1! = 0.000001
hb!(i%) = h1!
6670:  Next i%

' BSC2 / CALCULATION OF LOVE/SCOTT BACKSCATTER COEFFICIENTS FOR PURE ELEMENTS
ElseIf ibsc% = 2 Then
For i% = 1 To zaf.in0%
yy! = zaf.Z%(i%)
h1! = (-52.3791 + 150.48371 * yy! - 1.67373 * yy! ^ 2 + 0.00716 * yy! ^ 3) / 10000#
h2! = (-1112.8 + 30.289 * yy! - 0.15498 * yy! ^ 2) / 10000#
hb!(i%) = h1! * (1# + h2! * Log(zaf.eO!(i%) / 20#))
' hb!(i%) = 0.891 * ((hb!(i%) / 0.891) ^ (Cos(Alpha!)))  ' "alpha" is inclination of beam
Next i%

' BSC3 / CALCULATION OF POUCHOU/PICHOIR BACKSCATTER COEFFICIENTS FOR PURE ELEMENTS
ElseIf ibsc% = 3 Then
For i% = 1 To zaf.in0%
yy! = zaf.Z%(i%)
h1! = 0.00175 * yy! + 0.37 * (1# - Exp(-0.015 * Exp(1.3 * Log(yy!))))
hb!(i%) = h1!
Next i%

' BSC4 / CALCULATION OF HUNGLER & KUCHLER (AUGUST & WERNISCH) BACKSCATTER COEFFICIENTS FOR PURE ELEMENTS
ElseIf ibsc% = 4 Then
For i% = 1 To zaf.in0%
yy! = zaf.Z%(i%)
h1! = 0.1904 - 0.2236 * Log(yy!) + 0.1292 * (Log(yy!)) ^ 2 - 0.01491 * (Log(yy!)) ^ 3
h2! = 0.0002167 * yy! + 0.9987
h1! = h1! * h2! * zaf.eO!(i) ^ (0.1382 - 0.9211 / Sqr(yy!))
hb!(i%) = h1!
Next i%
End If

Exit Sub

' BSC1, BSC2, BSC4 / SAMPLE CALCULATION OF HEINRICH, LOVE/SCOTT AND HUNGLER BACKSCATTER COEFFICIENTS
6900:
If ibsc% = 1 Or ibsc% = 2 Or ibsc% = 4 Then
For i% = 1 To zaf.in0%
eta!(i%) = 0#
For i1% = 1 To zaf.in0%
eta!(i%) = eta!(i%) + zaf.conc!(i1%) * hb!(i1%)
Next i1%
Next i%

' BSC3 / SAMPLE CALCULATION OF POUCHOU and PICHOIR BACKSCATTER
ElseIf ibsc% = 3 Then
Zbar! = 0#
For i1% = 1 To zaf.in0%
Zbar! = Zbar! + zaf.conc!(i1%) * Sqr(zaf.Z%(i1%))
Next i1%
Zbar! = Zbar! * Zbar!

For i% = 1 To zaf.in0%
eta!(i%) = 0.00175 * Zbar! + 0.37 * (1 - Exp(-0.015 * Exp(1.3 * Log(Zbar!))))
Next i%
End If

Exit Sub

' Errors
ZAFBscError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFBsc"
ierror = True
Exit Sub

End Sub

Sub ZAFMip(zafinit As Integer)
' When imip% equals:
' 1 = "Mean Ionization of Berger/Seltzer"
' 2 = "Mean Ionization of Duncumb/Da Casa"
' 3 = "Mean Ionization of Ruste & Zeller"
' 4 = "Mean Ionization of Springer"
' 5 = "Mean Ionization of Wilson"
' 6 = "Mean Ionization of Heinrich"
' 7 = "Mean Ionization of Bloch (Love/Scott)"
' 8 = "Mean Ionization of Armstrong (Springer/Berger)"
' 9 = "Mean Ionization of Joy (Wilson/Berger)"

ierror = False
On Error GoTo ZAFMipError

Dim i As Integer
Dim m5 As Single, m6 As Single

If zafinit% = 1 Then GoTo 5400  ' calculate sample intensities

' STDMIP1 / BERGER/SELTZER MEAN IONIZATION POTENTIALS FOR STANDARDS
If imip% = 1 Then
For i% = 1 To zaf.in0%
jm!(i%) = (9.76 * zaf.Z%(i%) + 58.5 * (1# / zaf.Z%(i%) ^ 0.19)) / 1000#
' was jm!(i%) = (9.76 * zaf.z%(i%) + 58.5 * zaf.z%(i%) ^ (-.19)) / 1000#
Next i%

' STDMIP2 / DUNCUMB/DA CASA MEAN IONIZATION POTENTIALS FOR STANDARDS
ElseIf imip% = 2 Then
For i% = 1 To zaf.in0%
jm!(i%) = 14# * (1# - Exp(-0.1 * zaf.Z%(i%))) + 75.5 / zaf.Z%(i%) ^ (zaf.Z%(i%) / 7.5) - zaf.Z%(i%) / (100# + zaf.Z%(i%))
jm!(i%) = zaf.Z%(i%) * jm!(i%) / 1000#
Next i%

' STDMIP3 / RUSTE & ZELLER MEAN IONIZATION POTENTIALS FOR STANDARDS
ElseIf imip% = 3 Then
For i% = 1 To zaf.in0%
jm!(i%) = 10.04 + 8.25 * Exp(-zaf.Z%(i%) / 11.22)   ' changed 10.4 to 10.04 5/26/04 (typo)
jm!(i%) = zaf.Z%(i%) * jm!(i%) / 1000#
Next i%

' STDMIP4 / SPRINGER MEAN IONIZATION POTENTIALS FOR STANDARDS
ElseIf imip% = 4 Then
For i% = 1 To zaf.in0%
jm!(i%) = 9# * (1 + zaf.Z%(i%) ^ (-2 / 3)) + 0.03 * zaf.Z%(i%)
jm!(i%) = zaf.Z%(i%) * jm!(i%) / 1000#
Next i%

' STDMIP5 / WILSON MEAN IONIZATION POTENTIALS FOR STANDARDS
ElseIf imip% = 5 Then
For i% = 1 To zaf.in0%
jm!(i%) = 11.5
jm!(i%) = zaf.Z%(i%) * jm!(i%) / 1000#
Next i%

' STDMIP6 / HEINRICH MEAN IONIZATION POTENTIALS FOR STANDARDS
ElseIf imip% = 6 Then
For i% = 1 To zaf.in0%
jm!(i%) = 9.94 + 19.52 / zaf.Z%(i%)
jm!(i%) = zaf.Z%(i%) * jm!(i%) / 1000#
Next i%

' STDMIP7 / BLOCH (LOVE/SCOTT) MEAN IONIZATION POTENTIALS FOR STANDARDS
ElseIf imip% = 7 Then
For i% = 1 To zaf.in0%
jm!(i%) = 13.5
jm!(i%) = zaf.Z%(i%) * jm!(i%) / 1000#
Next i%

' STDMIP8 / ARMSTRONG (SPRINGER/BERGER) MEAN IONIZATION POTENTIALS FOR STANDARDS
ElseIf imip% = 8 Then
For i% = 1 To zaf.in0%
If zaf.Z%(i%) < 30# Then
  jm!(i%) = 9# * (1# + zaf.Z%(i%) ^ (-2 / 3)) + 0.03 * zaf.Z%(i%)
Else
  jm!(i%) = 9.76 + 58.5 * zaf.Z%(i%) ^ (-1.19)
End If
jm!(i%) = zaf.Z%(i%) * jm!(i%) / 1000#
Next i%

' STDMIP9 / JOY (WILSON/BERGER) MEAN IONIZATION POTENTIALS FOR STANDARDS
ElseIf imip% = 9 Then
For i% = 1 To zaf.in0%
If zaf.Z%(i%) < 13# Then
  jm!(i%) = 11.5
Else
  jm!(i%) = 9.76 + 58.5 * zaf.Z%(i%) ^ (-1.19)
End If
jm!(i%) = zaf.Z%(i%) * jm!(i%) / 1000#
Next i%
End If

Exit Sub

' SAMPLE MIP CALCULATION
5400:
m5! = 0#
m6! = 0#
For i% = 1 To zaf.in0%
m5! = m5! + zaf.conc!(i%) * zaf.Z%(i%) / zaf.atwts!(i%)
m6! = m6! + zaf.conc!(i%) * zaf.Z%(i%) * Log(jm!(i%)) / zaf.atwts!(i%)
Next i%

sumatom! = m5!
If m6! / m5! > MAXLOGEXPS! Then GoTo ZAFMipBadJbar
jbar! = Exp(m6! / m5!)

Exit Sub

' Errors
ZAFMipError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFMip"
ierror = True
Exit Sub

ZAFMipBadJbar:
msg$ = "Bad jbar parameter calculated for the sample analysis. This usually indicates negative concentrations so you should check that you are not analyzing epoxy." & vbCrLf & vbCrLf
msg$ = msg$ & "You should also make sure your off-peak background and interference corrections are not overcorrecting, or perhaps you have assigned a blank correction to a major or minor element and you did not enter the correct blank level in the Standard Assignments dialog."
If Not CalcImageQuantFlag Then
MiscMsgBoxTim FormMSGBOXTIME, "ZAFMip", msg$, 20#
Call IOWriteLog(msg$)
Else
Call IOWriteLog(msg$)
End If
'ierror = True
Exit Sub

End Sub

Sub ZAFMip2(ii As Integer, zz As Single, ww As Single)
' Simplified Mean J for use in Phi(PZ) absorption correction

ierror = False
On Error GoTo ZAFMip2Error

' SMPMIP1 / BERGER/SELTZER MEAN IONIZATION POTENTIAL FOR SAMPLE
If imip% = 1 Then
ww! = (9.76 * zz! + 58.5 * (1# / zz! ^ 0.19)) / 1000#
' was ww! = (9.76 * zz! + 58.5 * zz! ^ (-.19)) / 1000#

' SMPMIP2 / DUNCUMB/DA CASA MEAN IONIZATION POTENTIAL FOR SAMPLE
ElseIf imip% = 2 Then
ww! = 14# * (1# - Exp(-0.1 * zz!)) + 75.5 / zz! ^ (zz! / 7.5) - zz! / (100# + zz!)
ww! = zz! * ww! / 1000#

' SMPMIP3 / RUSTE & ZELLER MEAN IONIZATION POTENTIAL FOR SAMPLE
ElseIf imip% = 3 Then
ww! = 10.04 + 8.25 * Exp(-zz! / 11.22)
ww! = zz! * ww! / 1000#

' SMPMIP4 / SPRINGER MEAN IONIZATION POTENTIAL FOR SAMPLE
ElseIf imip% = 4 Then
ww! = 9# * (1# + zz! ^ (-2 / 3)) + 0.03 * zz!
ww! = zz! * ww! / 1000#

' SMPMIP5 / WILSON MEAN IONIZATION POTENTIAL FOR SAMPLE
ElseIf imip% = 5 Then
ww! = 11.5
ww! = zz! * ww! / 1000#

' SMPMIP6 / HEINRICH MEAN IONIZATION POTENTIAL FOR SAMPLE
ElseIf imip% = 6 Then
ww! = 9.94 + 19.52 / zz!
ww! = zz! * ww! / 1000#

' SMPMIP7 / BLOCH (LOVE/SCOTT) MEAN IONIZATION POTENTIAL FOR SAMPLE
ElseIf imip% = 7 Then
ww! = 13.5
ww! = zz! * ww! / 1000#

' SMPMIP8 / ARMSTRONG (SPRINGER/BERGER) MEAN IONIZATION POTENTIAL FOR SAMPLE
ElseIf imip% = 8 Then
If zaf.Z%(ii%) < 30# Then
  ww! = 9# * (1# + zz! ^ (-2 / 3)) + 0.03 * zz!
Else
  ww! = 9.76 + 58.5 * zz! ^ (-1.19)
End If
ww! = zz! * ww! / 1000#

' SMPMIP9 / JOY (WILSON/BERGER) MEAN IONIZATION POTENTIAL FOR SAMPLE
ElseIf imip% = 9 Then
If zaf.Z%(ii%) < 13# Then
  ww! = 11.5
Else
  ww! = 9.76 + 58.5 * zz! ^ (-1.19)
End If
ww! = zz! * ww! / 1000#
End If

Exit Sub

' Errors
ZAFMip2Error:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFMip2"
ierror = True
Exit Sub

End Sub

Sub ZAFPhiCal(zafinit As Integer)
' Phi-Rho-Z calculation
' When iabs% equals:
'  7, then use  PHI(PZ) ABSORPTION of BROWN 1982/XRS ALPHA
'  8, then use  PHI(PZ) ABSORPTION of BASTIN 1984/XRS ALPHA
'  9, then use  PHI(PZ) ABSORPTION of BROWN 1981/JTA ALPHA
'  10, then use  PHI(PZ) ABSORPTION of Bastin (Scanning)
'  11, then use  PHI(PZ) ABSORPTION of Riveros
'
'
'  14, then use  PHI(PZ) ABSORPTION of Packwood (New)
'  15, then use  PHI(PZ) ABSORPTION of Bastin (Proza)

ierror = False
On Error GoTo ZAFPhiCalError

Dim i As Integer, i1 As Integer

Dim X2 As Single, x3 As Single, x4 As Single, x5 As Single
Dim zz As Single, za As Single, spi As Single, tx4 As Single

Dim a1 As Single, a2 As Single, uu As Single, v0 As Single, aa As Single
Dim erfx As Single, er1 As Single, er2 As Single, er3 As Single
Dim chi As Single, xx As Single, m7 As Single, pp As Single

Dim ww As Single, rr As Single, hh As Single, v1 As Single
Dim qeO As Single, ff As Single, rbas As Single

Dim beta_iter As Integer
Dim y1 As Single, Y2 As Single, beta0 As Single, beta1 As Single    ' variables for beta refinement by Brian Joy

' Calculate square root of pi (used to be  Sqr(3.14159) / 2, fixed 7/14/2011, Carpenter)
spi! = Sqr(3.14159)

' Calculate for each absorber
For i% = 1 To zaf.in1%
v0! = zaf.eO!(i%)

If zaf.il%(i%) > MAXRAY% - 1 Then GoTo 10610
If zafinit% = 1 Then GoTo 10200

' PHI-STD
zz! = zaf.Z%(i%)
aa! = zaf.atwts!(i%)
ww! = jm!(i%)

hh! = hb!(i%)
m7! = zaf.mup!(i%, i%)  ' load pure element MAC
GoTo 10320

' PHI-SMP
10200:
'If zaf.conc!(i%) = 0# Then GoTo 10610  ' do not skip for zero concentration
zz! = 0#
aa! = 0#
pp! = 0#

hh! = eta!(i%)

For i1% = 1 To zaf.in0%
zz! = zz! + zaf.conc!(i1%) * zaf.Z%(i1%) / zaf.atwts!(i1%)
aa! = aa! + zaf.conc!(i1%)
pp! = pp! + zaf.conc!(i1%) / zaf.atwts!(i1%)
Next i1%

zz! = zz! / pp!
aa! = aa! / pp!
m7! = ZAFMACCal(i%, zaf)    ' load sample MAC

' Calculate mean ionization for sample
If zz! < 0# Then GoTo ZAFPhiCalBadZZ
Call ZAFMip2(i%, zz!, ww!)
If ierror Then Exit Sub

' PHI-RUN
10320:
v1! = zaf.eC!(i%)
xx! = m7!   ' loaded by ZAFMACCal above
uu! = zaf.eO!(i%) / zaf.eC!(i%)

' x2! = Gamma(0), x3! = Beta, x4! = Alpha, R! = Phi(0)
X2! = 5# * 3.14159 * uu! / (Log(uu!) * (uu! - 1#)) * (Log(uu!) - 5# + 5# * (1# / uu! ^ 0.2))
' was x2! = 5# * 3.14159 * uu! / (Log(uu!) * (uu! - 1#)) * (Log(uu!) - 5# + 5# * uu! ^ (-.2))

' PACKWOOD-BROWN 1982 XRS PHI(PZ) ALPHA EXPRESSION
If iabs% = 7 Then
tx4! = Log(1.166 * v0! / ww!) / (v0! - v1!)
If tx4! < 0# Then GoTo ZAFPhiCalBadX4
x4! = 395000# * zz! ^ 0.95 / (aa! * v0! ^ 1.25) * (tx4!) ^ 0.5
x3! = 0.4 * x4! * zz! ^ 0.6

' BASTIN 1984 XRS PHI(PZ) ALPHA EXPRESSION
ElseIf iabs% = 8 Then
tx4! = Log(1.166 * v0! / ww!) / v1!
If tx4! < 0# Then GoTo ZAFPhiCalBadX4
x4! = 175000# / (v0! ^ 1.25 * (uu! - 1#) ^ 0.55) * (tx4!) ^ 0.5
X2! = 0.98 * X2! * Exp(0.001 * zz!)
x3! = 0.4 * x4! * (zz! ^ 1.7 / aa!) * (uu! - 1#) ^ 0.3

' BROWN 1981 JTA PHI(PZ) ALPHA EXPRESSION
ElseIf iabs% = 9 Then
tx4! = Log(1.166 * v0! / ww!) / (v0! - v1!)
If tx4! < 0# Then GoTo ZAFPhiCalBadX4
x4! = 297000# * zz! ^ 1.05 / (aa! * v0! ^ 1.25) * (tx4!) ^ 0.5
x3! = 850000# * zz! * zz! / (aa! * v0! * v0! * (X2! - 1#))

' BASTIN 1986/SCANNING
ElseIf iabs% = 10 Then
tx4! = Log(1.166 * v0! / ww!) / v1!
If tx4! < 0# Then GoTo ZAFPhiCalBadX4
x4! = 175000# / (v0! ^ 1.25 * (uu! - 1#) ^ 0.55) * (tx4!) ^ 0.5
X2! = 5# * 3.14159 * (uu! + 1#) / (Log(uu! + 1#) * uu!) * (Log(uu! + 1#) - 5# + 5# * (uu! + 1#) ^ (-0.2))
If uu! < 3# Then X2! = 1# + (uu! - 1#) / (0.3384 + 0.4742 * (uu! - 1#))
x3! = zz! / (0.4765 + 0.5473 * zz!)
x3! = x4! * (zz! ^ x3!) / aa!
       
' RIVEROS 1987/XRS
ElseIf iabs% = 11 Then
tx4! = Log(1.166 * v0! / ww!) / (v0! - v1!)
If tx4! < 0# Then GoTo ZAFPhiCalBadX4
x4! = 214000# * zz! ^ 1.16 / (aa! * v0! ^ 1.25) * (tx4!) ^ 0.5
X2! = (1# + hh!) * uu! * Log(uu!) / (uu! - 1#)
x3! = 10900# * zz! ^ 1.5 / (aa! * (v0! - v1!))
        
' PACKWOOD EPQ-1991
ElseIf iabs% = 14 Then
X2! = 10# * 3.14159 * (uu! / (uu! - 1#)) * (1# + (10# / Log(uu!)) * ((uu! ^ -0.1) - 1))

If zafinit% = 0 Then
x4! = 450000# * ((zz! - 1.3) / zz!) * (zz! / aa!) ^ 0.5 * v0! ^ -0.75
x4! = x4! * Sqr((zz! / aa!) * Log(1.166 * ((v0! + v1!) / (2# * ww!))) / (v0! ^ 2 - v1! ^ 2))
x3! = 0.4 * x4! * zz! ^ 0.6
ElseIf zafinit% = 1 Then
zz! = 0#
za! = 0#
ww! = 0#
For i1% = 1 To zaf.in0%
zz! = zz! + zaf.conc!(i1%) * zaf.Z%(i1%)
za! = za! + zaf.conc!(i1%) * zaf.Z%(i1%) / zaf.atwts!(i1%)
ww! = ww! + zaf.conc!(i1%) * (zaf.Z%(i1%) / zaf.atwts!(i1%)) * Log(1.166 * (v0! + v1!) / (2# * jm!(i1%)))
Next i1%

x4! = 450000# * ((zz! - 1.3) / zz!) * za! ^ 0.5 * v0! ^ -0.75
x4! = x4! * Sqr(ww! / (v0! ^ 2 - v1! ^ 2))
x3! = 0.4 * x4! * zz! ^ 0.6
End If
       
' BASTIN PROZA EPQ-1991
ElseIf iabs% = 15 Then
If zafinit% = 0 Then
x4! = pz!(i%, i%)
xp! = zaf.r!(i%, i%) / zaf.s!(i%, i%)
         
ElseIf zafinit% = 1 Then
x4! = 0#
zz! = 0#
za! = 0#

xp! = zaf.bks!(i%) / zaf.stp!(i%)
For i1% = 1 To zaf.in0%
zz! = zz! + zaf.conc!(i1%) * zaf.Z%(i1%)
za! = za! + zaf.conc!(i1%) * zaf.Z%(i1%) / zaf.atwts!(i1%)
x4! = x4! + zaf.conc!(i1%) * zaf.Z%(i1%) / (zaf.atwts!(i1%) * pz!(i%, i1%))
Next i1%
x4! = x4! / za!
x4! = 1# / x4!
End If

If uu! > 6# Then
X2! = 2.814333 * uu! ^ (0.262702 * zz! ^ -0.1614454)
Else
'x2! = 3.98352 * uu! ^ -0.0516861 * (1.2276233 - uu! ^ (-1.225558 * zz! ^ -0.1424549))  ' original typo
X2! = 3.98352 * uu! ^ -0.0516861 * (1.276233 - uu! ^ (-1.25558 * zz! ^ -0.1424549))     ' fixed from "green book", 7-14-2011, Carpenter
End If
If v1! < 0.7 Then X2! = X2! * v1! / (-0.041878 + 1.05975 * v1!)
End If

' End of all procedures, now calculate phi expression ("rr")
Call ZAFPhi(i%, uu!, hh!, zz!, v1!, rr!)
If ierror Then Exit Sub

x5! = X2! - rr!

' PTC modification
If UseParticleCorrectionFlag And iptc% = 1 Then
x5! = x5! / X2!
End If

' Special code for Bastin PROZA
If iabs% = 15 Then
qeO! = Log(zaf.v!(i%)) / (zaf.eC!(i%) * zaf.eC!(i%) * Exp(em!(i%) * Log(zaf.v!(i%))))
xp! = xp! / (zipi!(i%) * 66892#) * zaf.atwts!(i%)
ff! = xp! / qeO!
If Not UseBrianJoyModifications Then
rbas! = (X2! - x4! * ff! / spi!) / x5!              ' original code from CITZAF
Else
rbas! = (X2! - 2 * x4! * ff! / spi!) / x5!          ' corrected by Brian Joy, 01-2016
End If

If rbas! <= 0# Or rbas! >= 1# Then
x4! = (rr! + X2!) * spi! / (4# * ff!)   ' used to be (2# * ff!), fixed 7/14/2011, Carpenter
rbas! = 0.5
End If

' Original CITZAF code
If Not UseBrianJoyModifications Then
If rbas! >= 0.9 And rbas! < 1 Then x3! = 0.9628832 - 0.964244 * rbas!
If rbas! > 0.8 And rbas! <= 0.9 Then x3! = 1.122405 - 1.141942 * rbas!
If rbas! > 0.7 And rbas! <= 0.8 Then x3! = 13.4381 * Exp(-5.180503 * rbas!)
If rbas! > 0.57 And rbas! <= 0.7 Then x3! = 5.909606 * Exp(-4.015891 * rbas!)
If rbas! > 0.306 And rbas! <= 0.57 Then x3! = 4.852357 * Exp(-3.680818 * rbas!)
If rbas! > 0.102 And rbas! <= 0.306 Then x3! = (1 - 0.5379956 * rbas!) / (1.685638 * rbas!)
If rbas! > 0.056 And rbas! <= 0.102 Then x3! = (1 - 1.043744 * rbas!) / (1.60482 * rbas!)
If rbas! > 0.03165 And rbas! <= 0.056 Then x3! = (1 - 2.749786 * rbas!) / (1.447465 * rbas!)
If rbas! > 0# And rbas! <= 0.03165 Then x3! = (1 - 4.894396 * rbas!) / (1.341313 * rbas!)

' Corrected by Brian Joy, 01-2016
Else
If rbas! >= 0.9 And rbas! < 1 Then x3! = (0.9628832 - 0.964244 * rbas!) * 2 * x4!
If rbas! > 0.8 And rbas! <= 0.9 Then x3! = (1.122405 - 1.141942 * rbas!) * 2 * x4!
If rbas! > 0.7 And rbas! <= 0.8 Then x3! = (13.4381 * Exp(-5.180503 * rbas!)) * 2 * x4!
If rbas! > 0.57 And rbas! <= 0.7 Then x3! = (5.909606 * Exp(-4.015891 * rbas!)) * 2 * x4!
If rbas! > 0.306 And rbas! <= 0.57 Then x3! = (4.852357 * Exp(-3.680818 * rbas!)) * 2 * x4!
If rbas! > 0.102 And rbas! <= 0.306 Then x3! = ((1 - 0.5379956 * rbas!) / (1.685638 * rbas!)) * 2 * x4!
If rbas! > 0.056 And rbas! <= 0.102 Then x3! = ((1 - 1.043744 * rbas!) / (1.60482 * rbas!)) * 2 * x4!
If rbas! > 0.03165 And rbas! <= 0.056 Then x3! = ((1 - 2.749786 * rbas!) / (1.447465 * rbas!)) * 2 * x4!
If rbas! > 0# And rbas! <= 0.03165 Then x3! = ((1 - 4.894396 * rbas!) / (1.341313 * rbas!)) * 2 * x4!

' Refinement of beta(i,j), adapted from GMRFILM by R. Waldo, by Brian Joy, 02-2016
beta_iter% = 0
beta0! = x3!
y1! = x3! / (2# * x4!)
Do
    beta1! = beta0!
    Y2! = ZAFErrorFunction!(y1!) / rbas! * y1!
    beta0! = Y2! * (2# * x4!)
    If (Abs((beta1! - beta0!) / beta1!) < 0.0001) Then Exit Do
    y1! = Y2!
beta_iter% = beta_iter% + 1
If beta_iter% > 100 Then Exit Do
Loop
x3! = y1! * (2# * x4!)
End If
End If

' Calculate error function
chi! = xx! * zaf.m1!(i%)

' New call to ZAFPtc for particles and thin films
If UseParticleCorrectionFlag And iptc% = 1 Then
Call ZAFPtc(i%, aa!, v0!, zz!, er1!, er2!, er3!, a1!, a2!, xx!, X2!, x3!, x4!, x5!)
If ierror Then Exit Sub

' Normal bulk sample calculation
Else
erfx! = x3! / (2# * x4!)
er1! = ZAFErrorFunction(erfx!)

erfx! = chi! / (2# * x4!)
er2! = ZAFErrorFunction(erfx!)

erfx! = (x3! + chi!) / (2# * x4!)
er3! = ZAFErrorFunction(erfx!)

a1! = spi! * (X2! - x5! * er1!) / x4!
a2! = spi! * (X2! * er2! - x5! * er3!) / x4!
End If

' Calculate intensities
phi!(i%) = a1!
If zafinit% = 0 And a2! <> 0# Then zaf.genstd!(i%) = a1! / a2!
If zafinit% = 1 And a2! <> 0# Then zaf.gensmp!(i%) = a1! / a2!

10610:  Next i%
Exit Sub

' Errors
ZAFPhiCalError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFPhiCal"
ierror = True
Exit Sub

ZAFPhiCalBadZZ:
msg$ = "Bad (negative) zz parameter calculated for the sample analysis. This usually indicates negative concentrations so you should check that you are not analyzing epoxy." & vbCrLf & vbCrLf
msg$ = msg$ & "You should also make sure your off-peak background and interference corrections are not overcorrecting, or perhaps you have assigned a blank correction to a major or minor element and you did not enter the correct blank level in the Standard Assignments dialog."
If Not CalcImageQuantFlag Then
MiscMsgBoxTim FormMSGBOXTIME, "ZAFPhiCal", msg$, 20#
Call IOWriteLog(msg$)
Else
Call IOWriteLog(msg$)
End If
'ierror = True
Exit Sub

ZAFPhiCalBadX4:
msg$ = "Bad (negative) x4 parameter calculated for the sample analysis. This usually indicates negative concentrations so you should check that you are not analyzing epoxy." & vbCrLf & vbCrLf
msg$ = msg$ & "You should also make sure using the hydrogen by stochiometry to excess oxygen calculation appropriately (see Calculation Options button in the Analyze! window."
If Not CalcImageQuantFlag Then
MiscMsgBoxTim FormMSGBOXTIME, "ZAFPhiCal", msg$, 20#
Call IOWriteLog(msg$)
Else
Call IOWriteLog(msg$)
End If
'ierror = True
Exit Sub

End Sub

Sub ZAFSmp(row As Integer, UnkCounts() As Single, zerror As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' This routine calculates weight percents using algorithims based on John
' Armstrong's ZAF program CITZAF.  Input are the element list, standard k-ratios
' for each element in the standard for that channel, the counts on that standard
' for that element, and the counts on the unknown sample.  Output are the unknown
' weight percents, the unknown k-ratios, and the ZAF correction factors for the
' sample as explained below:

'  analysis.UnkZAFCors(1,chan%) = absorption correction
'  analysis.UnkZAFCors(2,chan%) = fluorescence correction
'  analysis.UnkZAFCors(3,chan%) = atomic number correction (Stp/s * r/Bks)
'  analysis.UnkZAFCors(4,chan%) = ZAF correction (abscor*flucor*zedcor)
'  analysis.UnkZAFCors(5,chan%) = stopping power correction
'  analysis.UnkZAFCors(6,chan%) = backscatter correction
'  analysis.UnkZAFCors(7,chan%) = primary intensity
'  analysis.UnkZAFCors(8,chan%) = sample intensity

' Note x-ray flags
'  il() = 0 = stoichiometric element (oxygen)
'  il() = 1 = Ka
'  il() = 2 = Kb
'  il() = 3 = La
'  il() = 4 = Lb
'  il() = 5 = Ma
'  il() = 6 = Mb

'  il() = 7 = Ln
'  il() = 8 = Lg
'  il() = 9 = Lv
'  il() = 10 = Ll
'  il() = 11 = Mg
'  il() = 12 = Mz

'  il() = 13 = by difference
'  il() = 14 = by specified concentration
'  il() = 15 = by stoichiometry to stoichiometric oxygen
'  il() = 16 = disabled quantification
'  il() = 17 = by stoichiometry to another element
'  il() = 18 = by hydrogen stoichiometry to oxygen (measured, specified or calculated)
'  il() = 19 = by difference (formula)

ierror = False
On Error GoTo ZAFSmpError

Dim i As Integer, j As Integer
Dim ip As Integer, ipp As Integer, ippp As Integer
Dim r0 As Integer, MaxZAFIter  As Integer
Dim ZAFMinTotal  As Single, ZAFMinToler As Single
Dim i7 As Integer, i8 As Integer
Dim temp As Single, oxygen As Single
Dim astring As String

ReDim r1(1 To MAXCHAN1%) As Single
ReDim ZAFDiff(1 To MAXCHAN1%) As Single

ReDim unkcnts(1 To MAXCHAN1%) As Single
ReDim stdcnts(1 To MAXCHAN1%) As Single

MaxZAFIter% = 100
ZAFMinTotal! = 0.001 ' in weight fraction
ZAFMinToler! = 0.0001 ' in weight fraction

Const MAXNEGATIVE_KRATIO! = -0.2
Const MAXNEGATIVE_SUMKRATIO! = -0.01

For i% = 1 To MAXCHAN1%
zaf.krat!(i%) = 0#
zaf.conc!(i%) = 0#
Next i%

If VerboseMode Then
Call IOWriteLog(vbCrLf & "Entering ZAFSmp...")
If Not sample(1).CombinedConditionsFlag% Then
Call IOWriteLog("Takeoff = " & Str$(sample(1).takeoff!) & ", Kilovolts = " & Str$(sample(1).kilovolts!))
End If

msg$ = "ELEMENT "
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "UNK WT% "
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(analysis.WtPercents!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "UNK CNT "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & MiscAutoFormat$(UnkCounts!(i%))
Next i%
Call IOWriteLog(msg$)
End If

' Check for zero standard counts
For i% = 1 To sample(1).LastElm%
If sample(1).DisableAcqFlag%(i%) = 0 Then
If sample(1).KilovoltsArray!(i%) = 0# Then GoTo ZAFSmpBadKilovolts

' Find if element is duplicated
ip% = IPOS8(i%, sample(1).Elsyms$(i%), sample(1).Xrsyms$(i%), sample())
If Not UseAggregateIntensitiesFlag Or (UseAggregateIntensitiesFlag And ip% = 0) Then
If sample(1).DisableQuantFlag%(i%) = 0 Then             ' no disabled quant flag
If analysis.StdAssignsCounts!(i%) <= 0# Then GoTo ZAFSmpBadStdCounts
If analysis.StdAssignsKfactors!(i%) <= 0# Then GoTo ZAFSmpBadStdKrat
End If
End If
End If
Next i%

' Check if standard sample (if so, load standard coating parameters to this sample)
If sample(1).Type% = 1 Then
If UseConductiveCoatingCorrectionForElectronAbsorption = True Or UseConductiveCoatingCorrectionForXrayTransmission = True Then
ippp% = IPOS2(NumberofStandards%, sample(1).number%, StandardNumbers%())
If ippp% > 0 Then
sample(1).CoatingElement% = StandardCoatingElement%(ippp%)
sample(1).CoatingDensity! = StandardCoatingDensity!(ippp%)
sample(1).CoatingThickness! = StandardCoatingThickness!(ippp%)
sample(1).CoatingSinThickness! = MathCalculateSinThickness!(StandardCoatingThickness!(ippp%), DefaultTakeOff!)
End If
End If
End If

' Conductive coating debug
If DebugMode Then
If UseConductiveCoatingCorrectionForElectronAbsorption = True Or UseConductiveCoatingCorrectionForXrayTransmission = True Then

' Print out coating parameters for this sample
If sample(1).CoatingFlag% = 1 Then
astring$ = "Sample Coating=" & Trim$(Symup$(sample(1).CoatingElement%))
astring$ = astring$ & ", Density=" & Format$(sample(1).CoatingDensity!) & " gm/cm3"
astring$ = astring$ & ", Thickness=" & Format$(sample(1).CoatingThickness!) & " angstroms"
astring$ = astring$ & ", Sin(Thickness)=" & Format$(sample(1).CoatingSinThickness!) & " angstroms"
Else
astring$ = "No Sample Coating"
End If

If UseConductiveCoatingCorrectionForElectronAbsorption = True And Not UseConductiveCoatingCorrectionForXrayTransmission = True Then
msg$ = vbCrLf & "Using Conductive Coating Correction For Electron Absorption: " & vbCrLf & astring$
End If
If Not UseConductiveCoatingCorrectionForElectronAbsorption = True And UseConductiveCoatingCorrectionForXrayTransmission = True Then
msg$ = vbCrLf & "Using Conductive Coating Correction For X-Ray Transmission: " & vbCrLf & astring$
End If
If UseConductiveCoatingCorrectionForElectronAbsorption = True And UseConductiveCoatingCorrectionForXrayTransmission = True Then
msg$ = vbCrLf & "Using Conductive Coating Correction For Electron Absorption and X-Ray Transmission: " & vbCrLf & astring$
End If
Call IOWriteLog(msg$)
End If

' Print out coating calculations (display final coating calculations in ZAFPrintSmp)
If VerboseMode Then
If UseConductiveCoatingCorrectionForElectronAbsorption = True Or UseConductiveCoatingCorrectionForXrayTransmission = True Then
msg$ = vbCrLf & "Coating Correction (X-ray transmission):"
msg$ = msg$ & vbCrLf & "UNKCOAT "
For j% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(zaf.coating_trans_smp!(j%), f85), a80$)   ' calculated in ZAFSetZAF
Next j%
Call IOWriteLog(msg$)
msg$ = "STDCOAT "
For j% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.Coating_StdAssignsTrans!(j%), f85), a80$)    ' calculated in ZAFStd
Next j%
Call IOWriteLog(msg$)
msg$ = "UNK/STD "
For j% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(zaf.coating_trans_smp!(j%) / analysis.Coating_StdAssignsTrans!(j%), f85), a80$)   ' unk/std x-ray transmission ratio
Next j%
Call IOWriteLog(msg$)

msg$ = vbCrLf & "Coating Correction (Electron absorption):"
msg$ = msg$ & vbCrLf & "UNKCOAT "
For j% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(zaf.coating_absorbs_smp!(j%), f85), a80$)   ' calculated in ZAFSetZAF
Next j%
Call IOWriteLog(msg$)
msg$ = "STDCOAT "
For j% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.Coating_StdAssignsAbsorbs!(j%), f85), a80$)    ' calculated in ZAFStd
Next j%
Call IOWriteLog(msg$)
msg$ = "UNK/STD "
For j% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(zaf.coating_absorbs_smp!(j%) / analysis.Coating_StdAssignsAbsorbs!(j%), f85), a80$)   ' unk/std electron absorption ratio
Next j%
Call IOWriteLog(msg$)
End If
End If
End If

If VerboseMode Then
msg$ = vbCrLf & "STD WT% "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & MiscAutoFormat$(analysis.StdAssignsPercents!(i%))
Next i%
Call IOWriteLog(msg$)

msg$ = "STD CNT "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & MiscAutoFormat$(analysis.StdAssignsCounts!(i%))
Next i%
Call IOWriteLog(msg$)

msg$ = "RAWKRAT "
For i% = 1 To sample(1).LastElm%
temp! = 0#
If analysis.StdAssignsCounts!(i%) <> 0# Then temp! = UnkCounts!(i%) / analysis.StdAssignsCounts!(i%)    ' input raw k-ratio
msg$ = msg$ & MiscAutoFormat$(temp!)
Next i%
Call IOWriteLog(msg$)

msg$ = "STDKFAC "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & MiscAutoFormat$(analysis.StdAssignsKfactors!(i%))
Next i%
Call IOWriteLog(msg$)
End If

' Apply coating corrections (if specified)
For i% = 1 To sample(1).LastElm%
If sample(1).DisableAcqFlag%(i%) = 0 Then
If sample(1).DisableQuantFlag%(i%) = 0 Then
unkcnts!(i%) = UnkCounts!(i%)                   ' load unknown counts
stdcnts!(i%) = analysis.StdAssignsCounts!(i%)   ' load standard counts

' Apply coating correction from x-ray transmission and electron absorption for unknown intensity (as though unknown is uncoated).
If sample(1).CoatingFlag = 1 Then
If UseConductiveCoatingCorrectionForElectronAbsorption = True And Not UseConductiveCoatingCorrectionForXrayTransmission = True Then
unkcnts!(i%) = UnkCounts!(i%) / (zaf.coating_absorbs_smp!(i%))
End If
If Not UseConductiveCoatingCorrectionForElectronAbsorption = True And UseConductiveCoatingCorrectionForXrayTransmission = True Then
unkcnts!(i%) = UnkCounts!(i%) / (zaf.coating_trans_smp!(i%))
End If
If UseConductiveCoatingCorrectionForElectronAbsorption = True And UseConductiveCoatingCorrectionForXrayTransmission = True Then
unkcnts!(i%) = UnkCounts!(i%) / (zaf.coating_absorbs_smp!(i%) * zaf.coating_trans_smp!(i%))
End If
End If

' Apply coating correction from x-ray transmission and electron absorption for standard intensity (as though standard is uncoated).
If DefaultStandardCoatingFlag% = 1 Then
If UseConductiveCoatingCorrectionForElectronAbsorption = True And Not UseConductiveCoatingCorrectionForXrayTransmission = True Then
stdcnts!(i%) = analysis.StdAssignsCounts!(i%) / (analysis.Coating_StdAssignsAbsorbs!(i%))
End If
If Not UseConductiveCoatingCorrectionForElectronAbsorption = True And UseConductiveCoatingCorrectionForXrayTransmission = True Then
stdcnts!(i%) = analysis.StdAssignsCounts!(i%) / (analysis.Coating_StdAssignsTrans!(i%))
End If
If UseConductiveCoatingCorrectionForElectronAbsorption = True And UseConductiveCoatingCorrectionForXrayTransmission = True Then
stdcnts!(i%) = analysis.StdAssignsCounts!(i%) / (analysis.Coating_StdAssignsAbsorbs!(i%) * analysis.Coating_StdAssignsTrans!(i%))
End If
End If
End If
End If
Next i%

' Calculate unknown K-ratios. Use intensities modified by coating correction (if specified)
For i% = 1 To sample(1).LastElm%
If sample(1).DisableAcqFlag%(i%) = 0 Then
If analysis.StdAssignsCounts!(i%) <> 0# Then
If sample(1).DisableQuantFlag%(i%) = 0 Then
zaf.kraw!(i%) = unkcnts!(i%) / stdcnts!(i%)
zaf.krat!(i%) = (unkcnts!(i%) / stdcnts!(i%)) * analysis.StdAssignsKfactors!(i%)
End If
End If
If zaf.krat!(i%) = 0# Then zaf.krat!(i%) = NotAnalyzedValue! / 100# ' use a non-zero value

' Check for force to zero flag
If ForceNegativeKratiosToZeroFlag = True Then
If zaf.krat!(i%) <= 0# Then zaf.krat!(i%) = NotAnalyzedValue! / 100# ' use a non-zero value
End If

' Load voltage
zaf.eO!(i%) = sample(1).KilovoltsArray!(i%)

' Check for disable quant flag
If sample(1).DisableQuantFlag%(i%) = 1 Then zaf.il%(i%) = 16    ' use for disabled quant element

' Check for extremely negative k-ratios on each measured element
If zaf.krat!(i%) <= MAXNEGATIVE_KRATIO! Then GoTo ZAFSmpVeryNegativeKratio
End If
Next i%

' Input specified elemental weight percents into ZAF array, and check for element by difference, element by stoichiometry and element to oxide
' conversion elements and element relative to another element. NOTE: specified arrays must contain elemental concentrations.
For i% = sample(1).LastElm% + 1 To sample(1).LastChan%
zaf.il%(i%) = 14
zaf.krat!(i%) = analysis.WtPercents!(i%) / 100#

' Element by difference
If sample(1).DifferenceElementFlag% And sample(1).DifferenceElement$ <> vbNullString Then
ip% = IPOS1B(sample(1).LastElm% + 1, sample(1).LastChan%, sample(1).DifferenceElement$, sample(1).Elsyms$())
If ip% = i% Then
zaf.krat!(i%) = 0#
zaf.il%(i%) = 13
End If
End If

' Formula by difference
If sample(1).DifferenceFormulaFlag% And sample(1).DifferenceFormula$ <> vbNullString Then
If ConvertIsDifferenceFormulaElement(sample(1).DifferenceFormula$, sample(1).Elsyms$(i%)) Then
zaf.krat!(i%) = 0#
zaf.il%(i%) = 19
End If
End If

' Element by stoichiometry to stoichiometric oxygen
If sample(1).StoichiometryElementFlag And sample(1).StoichiometryElement$ <> vbNullString Then
ip% = IPOS1B(sample(1).LastElm% + 1, sample(1).LastChan%, sample(1).StoichiometryElement$, sample(1).Elsyms$())
If ip% = i% Then
zaf.krat!(i%) = 0#
zaf.il%(i%) = 15
End If
End If

' Element relative to another element
If sample(1).RelativeElementFlag% And sample(1).RelativeElement$ <> vbNullString And sample(1).RelativeToElement$ <> vbNullString Then
ip% = IPOS1B(sample(1).LastElm% + 1, sample(1).LastChan%, sample(1).RelativeElement$, sample(1).Elsyms$())
If ip% = i% Then
zaf.krat!(i%) = 0#
zaf.il%(i%) = 17
End If
End If

' Determine the channel number of the relative to element
ipp% = IPOS1(sample(1).LastChan%, sample(1).RelativeToElement$, sample(1).Elsyms$())

' Hydrogen stoichiometry to excess oxygen (measured, specified or calculated)
If sample(1).OxygenChannel% > 0 Then
If sample(1).DisableQuantFlag%(sample(1).OxygenChannel%) = 0 Then
If sample(1).HydrogenStoichiometryFlag And sample(1).HydrogenStoichiometryRatio! <> 0# Then
ip% = IPOS1B(sample(1).LastElm% + 1, sample(1).LastChan%, Symlo$(ATOMIC_NUM_HYDROGEN%), sample(1).Elsyms$())
If ip% = i% Then
zaf.krat!(i%) = 0#
zaf.il%(i%) = 18
End If
End If
End If
End If

Next i%

' First estimate of composition (First sum unknown K-ratios)
zaf.ksum! = 0#
For i% = 1 To zaf.in1%
If zaf.krat!(i%) = 0# Then zaf.krat!(i%) = NotAnalyzedValue! / 100#
zaf.ksum! = zaf.ksum! + zaf.krat!(i%)
Next i%

' Add in stoichiometric element (oxygen)
If zaf.il%(zaf.in0%) = 0 Then
zaf.krat!(zaf.in0%) = 0#
For i% = 1 To zaf.in1%
zaf.krat!(zaf.in0%) = zaf.krat!(zaf.in0%) + zaf.krat!(i%) * zaf.p1!(i%)
Next i%
zaf.ksum! = zaf.ksum! + zaf.krat!(zaf.in0%)

' Add in elements calculated relative to stoichiometric element (in0%)
For i% = 1 To zaf.in1%
If zaf.il%(i%) = 15 Then
zaf.krat!(i%) = (zaf.krat!(zaf.in0%) / zaf.atwts!(zaf.in0%)) * sample(1).StoichiometryRatio! * zaf.atwts!(i%)
zaf.krat!(zaf.in0%) = zaf.krat!(zaf.in0%) + zaf.krat!(i%) * zaf.p1!(i%)
zaf.ksum! = zaf.ksum! + zaf.krat!(i%) + zaf.krat!(i%) * zaf.p1!(i%)
End If
Next i%
End If

' Add in element relative to another element
For i% = 1 To zaf.in1%
If zaf.il%(i%) = 17 Then
zaf.krat!(i%) = zaf.krat!(ipp%) / zaf.atwts!(ipp%)
zaf.krat!(i%) = zaf.krat!(i%) * sample(1).RelativeRatio! * zaf.atwts!(i%)
If zaf.il%(zaf.in0%) = 0 Then   ' if calculating oxygen by stoichiometry
zaf.krat!(zaf.in0%) = zaf.krat!(zaf.in0%) + zaf.krat!(i%) * zaf.p1!(i%)
zaf.ksum! = zaf.ksum! + zaf.krat!(i%) + zaf.krat!(i%) * zaf.p1!(i%)
Else
zaf.ksum! = zaf.ksum! + zaf.krat!(i%)
End If
End If
Next i%

' Add in hydrogen by stoichiometry to excess oxygen
For i% = 1 To zaf.in1%
If zaf.il%(i%) = 18 Then
zaf.krat!(i%) = ZAFConvertExcessOxygenToHydrogen!(zaf.krat!(), zaf, sample())
zaf.ksum! = zaf.ksum! + zaf.krat!(i%)
End If
Next i%

' Add in element by difference, set total to 100 %
For i% = 1 To zaf.in0%
If zaf.il%(i%) = 13 Then
If zaf.ksum! < 1# Then
zaf.krat!(i%) = 1# - zaf.ksum!
If zaf.il%(zaf.in0%) = 0 Then   ' if calculating oxygen by stoichiometry
zaf.krat!(i%) = zaf.krat!(i%) / (1# + zaf.p1!(i%))
zaf.krat!(zaf.in0%) = zaf.krat!(zaf.in0%) + zaf.krat!(i%) * zaf.p1!(i%)
End If
zaf.ksum! = 1#
End If
End If
Next i%

' Add in formula elements by difference
If sample(1).DifferenceFormulaFlag Then
Call FormulaFormulaToSample(sample(1).DifferenceFormula$, FormulaTmpSample())
If ierror Then Exit Sub

' Calculate sum of composition skipping formula by difference elements
zaf.ksum! = 0#
For i% = 1 To zaf.in0%
If zaf.il%(i%) <> 19 Then
zaf.ksum! = zaf.ksum! + zaf.krat!(i%)
End If
Next i%

' Determine difference from 100%
temp! = 1# - zaf.ksum!
If temp! < 0# Then temp! = 1#

' Add in formula by difference elements (search from 1 to LastChan in FormulaTmpSample())
For i% = 1 To zaf.in0%
If zaf.il%(i%) = 19 Then
If zaf.ksum! < 1# Then
ip% = IPOS1B(Int(1), FormulaTmpSample(1).LastChan%, sample(1).Elsyms$(i%), FormulaTmpSample(1).Elsyms$())
If ip% > 0 Then
zaf.krat!(i%) = FormulaTmpSample(1).ElmPercents!(ip%) / 100# * temp!
End If
End If
End If
Next i%
zaf.ksum! = 1#
End If

' Check for negative sum of k-ratios
If zaf.ksum! < MAXNEGATIVE_SUMKRATIO! Then GoTo ZAFSmpNegativeSumOfKratios

' Check for insufficient total
If zaf.ksum! < ZAFMinTotal Then GoTo ZAFSmpInsufficientTotal

' Normalize K-ratio to 1.000 to assist in convergence
For i% = 1 To zaf.in0%
r1!(i%) = 0#
zaf.conc!(i%) = zaf.krat!(i%)
zaf.conc!(i%) = zaf.conc!(i%) / zaf.ksum!
Next i%

' Perform secondary boundary fluorescence correction on measured k-ratios
If UseSecondaryBoundaryFluorescenceCorrectionFlag Then
Call SecondaryCorrection(row%, zaf.krat!(), sample())
If ierror Then Exit Sub
End If

' ZAF iteration loop
zaf.iter% = 1
zaf.n8& = sample(1).Linenumber&(row%)
2860:

' Load ZAFPtc defaults based on current model and diameter
If UseParticleCorrectionFlag And iptc% = 1 Then
Call GetPTCDefaults(Int(1), zaf)
If ierror Then Exit Sub
Else
zaf.imodel% = 1
zaf.idiam% = 1
End If

' PTC loop code (diameters and models)
For i7% = 1 To zaf.imodel%
For i8% = 1 To zaf.idiam%

' Get current model and diameter
If UseParticleCorrectionFlag And iptc% = 1 Then
Call GetPTCGetModelDiameter(i7%, i8%, zaf)
If ierror Then Exit Sub
zaf.d! = zaf.rho! * zaf.diams!(i8%) / MICRONSPERCM&    ' convert microns to cm
End If

' Calculate matrix corrections
Call ZAFMip(Int(1))
If ierror Then Exit Sub

Call ZAFBsc(Int(1))
If ierror Then Exit Sub

If istp% = 6 Then
Call ZAFAbs(Int(1))
If ierror Then Exit Sub
Call ZAFStp(Int(1))
If ierror Then Exit Sub
Call ZAFBks(Int(1))
If ierror Then Exit Sub

Else
Call ZAFStp(Int(1))
If ierror Then Exit Sub
Call ZAFBks(Int(1))
If ierror Then Exit Sub
Call ZAFAbs(Int(1))
If ierror Then Exit Sub
End If

' Fluorescence correction
If iflu% < 5 Then
Call ZAFFlu(Int(1), zaf)
If ierror Then Exit Sub
Else
Call ZAFFlu3(Int(1), zaf)
If ierror Then Exit Sub
End If

' Apply coating corrections to sample (if specified)
For i% = 1 To sample(1).LastElm%
If sample(1).DisableQuantFlag%(i%) = 0 Then
If sample(1).CoatingFlag = 1 Then
If UseConductiveCoatingCorrectionForElectronAbsorption = True And Not UseConductiveCoatingCorrectionForXrayTransmission = True Then
r1!(i%) = zaf.krat!(i%) / (zaf.coating_absorbs_smp!(i%))
End If
If Not UseConductiveCoatingCorrectionForElectronAbsorption = True And UseConductiveCoatingCorrectionForXrayTransmission = True Then
r1!(i%) = zaf.krat!(i%) / (zaf.coating_trans_smp!(i%))
End If
If UseConductiveCoatingCorrectionForElectronAbsorption = True And UseConductiveCoatingCorrectionForXrayTransmission = True Then
r1!(i%) = zaf.krat!(i%) / (zaf.coating_absorbs_smp!(i%) * zaf.coating_trans_smp!(i%))
End If
End If
End If
Next i%

' Calculate concentrations w/ ZAF correction, C = K*Z*A*F
zaf.ksum! = 0#
r0% = 0
For i% = 1 To zaf.in1%

' Calculate atomic number correction
If zaf.il%(i%) <= MAXRAY% - 1 And zaf.krat!(i%) <> 0# Then
If zaf.s!(i%, i%) = 0# Then GoTo ZAFSmpBadEnergyLoss
zaf.stp!(i%) = zaf.stp!(i%) / zaf.s!(i%, i%)
If zaf.bks!(i%) = 0# Then GoTo ZAFSmpBadElectronLoss
zaf.bks!(i%) = zaf.r!(i%, i%) / zaf.bks!(i%)
zaf.zed!(i%) = zaf.stp!(i%) * zaf.bks!(i%)
End If

' Add in specified element concentrations
If zaf.il%(i%) > MAXRAY% - 1 Then
r1!(i%) = 0#
If zaf.il%(i%) = 14 Then
r1!(i%) = zaf.krat!(i%)
zaf.ksum! = zaf.ksum! + r1!(i%)
End If
Else

' Correct intensities for matrix correction
If zaf.genstd!(i%) <> 0# And (1# + zaf.vv!(i%)) <> 0# Then
r1!(i%) = zaf.krat!(i%) * zaf.zed!(i%) * (zaf.gensmp!(i%) / zaf.genstd!(i%)) / (1# + zaf.vv!(i%))
End If
zaf.ksum! = zaf.ksum! + r1!(i%)
End If
Next i%

' Calculate element relative to stoichiometric element based on previous iteration calculation of oxygen
If zaf.il%(zaf.in0%) = 0 Then    ' if calculating oxygen by stoichiometry
For i% = 1 To zaf.in1%
If zaf.il%(i%) = 15 Then
r1!(i%) = (r1!(zaf.in0%) / zaf.atwts!(zaf.in0%)) * sample(1).StoichiometryRatio! * zaf.atwts!(i%)
zaf.ksum! = zaf.ksum! + r1!(i%)
zaf.krat!(i%) = r1!(i%)
End If
Next i%

' Calculate amount of stoichiometric element and add to total
r1!(zaf.in0%) = 0#
For i% = 1 To zaf.in1%
r1!(zaf.in0%) = r1!(zaf.in0%) + r1!(i%) * zaf.p1!(i%)
Next i%

' Calculate equivalent oxygen from halogens and subtract from calculated oxygen if flagged
If UseOxygenFromHalogensCorrectionFlag Then r1!(zaf.in0%) = r1!(zaf.in0%) - ConvertHalogensToOxygen(zaf.in1%, sample(1).Elsyms$(), sample(1).DisableQuantFlag%(), r1!())

' Add to sum
zaf.ksum! = zaf.ksum! + r1!(zaf.in0%)
End If

' Calculate element relative to another element
For i% = 1 To zaf.in1%
If zaf.il%(i%) = 17 Then
r1!(i%) = r1!(ipp%) / zaf.atwts!(ipp%)
r1!(i%) = r1!(i%) * sample(1).RelativeRatio! * zaf.atwts!(i%)
If zaf.il%(zaf.in0%) = 0 Then   ' if calculating oxygen by stoichiometry
r1!(zaf.in0%) = r1!(zaf.in0%) + r1!(i%) * zaf.p1!(i%)
zaf.ksum! = zaf.ksum! + r1!(i%) + r1!(i%) * zaf.p1!(i%)
Else
zaf.ksum! = zaf.ksum! + r1!(i%)
End If
zaf.krat!(i%) = r1!(i%)
End If
Next i%

' Calculate hydrogen by stoichiometry to excess oxygen
For i% = 1 To zaf.in1%
If zaf.il%(i%) = 18 Then
r1!(i%) = ZAFConvertExcessOxygenToHydrogen!(r1!(), zaf, sample())
zaf.ksum! = zaf.ksum! + r1!(i%)
End If
Next i%

' Calculate element by difference
For i% = 1 To zaf.in1%
If zaf.il%(i%) = 13 Then
If zaf.ksum! < 1# Then
r1!(i%) = 1# - zaf.ksum!
If zaf.il%(zaf.in0%) = 0 Then
r1!(i%) = r1!(i%) / (1# + zaf.p1!(i%))
r1!(zaf.in0%) = r1!(zaf.in0%) + r1!(i%) * zaf.p1!(i%)
End If
zaf.krat!(i%) = r1!(i%)
zaf.ksum! = 1#
End If
End If
Next i%

' Calculate formula elements by difference
If sample(1).DifferenceFormulaFlag Then

' Calculate sum of composition skipping formula by difference elements
zaf.ksum! = 0#
For i% = 1 To zaf.in1%
If zaf.il%(i%) <> 19 Then
zaf.ksum! = zaf.ksum! + r1!(i%)
End If
Next i%

' Determine difference from 100%
temp! = 1# - zaf.ksum!

' Add in formula by difference elements (search from 1 to LastChan in FormulaTmpSample())
For i% = 1 To zaf.in1%
If zaf.il%(i%) = 19 Then
If zaf.ksum! < 1# Then
ip% = IPOS1B(Int(1), FormulaTmpSample(1).LastChan%, sample(1).Elsyms$(i%), FormulaTmpSample(1).Elsyms$())
If ip% > 0 Then
r1!(i%) = FormulaTmpSample(1).ElmPercents!(ip%) / 100# * temp!
zaf.krat!(i%) = r1!(i%)
End If
End If
End If
Next i%
zaf.ksum! = 1#
End If

' Normalize and check for convergence
If zaf.ksum! < ZAFMinTotal! Then GoTo ZAFSmpInsufficientTotal
For i% = 1 To zaf.in0%
r1!(i%) = r1!(i%) / zaf.ksum!
ZAFDiff!(i%) = Abs(zaf.conc!(i%) - r1!(i%))
If zaf.conc!(i%) > ZAFMinToler! And ZAFDiff!(i%) > zaf.conc!(i%) / 1000# Then r0% = 1 ' not converged yet
zaf.conc!(i%) = r1!(i%)
Next i%

' Print out normalized concentrations for each iteration
If VerboseMode Then
msg$ = vbCrLf & "NORMELEM"
For i% = 1 To zaf.in0%
If zaf.il%(i%) > 0 And zaf.il%(i%) < MAXRAY% Then
msg$ = msg$ & Format$(Symlo$(zaf.Z%(i%)) & " " & Xraylo$(zaf.il%(i%)), a80$)
Else
msg$ = msg$ & Format$(Symlo$(zaf.Z%(i%)) & " (" & Format$(zaf.il%(i%)) & ")", a80$)
End If
Next i%
Call IOWriteLog(msg$)
msg$ = "Conc*100"
For i% = 1 To zaf.in0%
msg$ = msg$ & MiscAutoFormat$(100# * zaf.conc!(i%))
Next i%
Call IOWriteLog(msg$)
End If

' If DebugMode and VerboseMode print out intermediate calculations
If DebugMode And VerboseMode Then
Call IOWriteLog(vbCrLf & "ZAFSmp: Iteration #" & Format$(zaf.iter%))
msg$ = "ELEMENT "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%), a80$)
Next i%
Call IOWriteLog(msg$)
msg$ = "ZAFAbs: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(zaf.gensmp!(i%) / zaf.genstd!(i%), f84), a80$)
Next i%
Call IOWriteLog(msg$)
msg$ = "ZAFFlu: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(1# / (1# + zaf.vv!(i%)), f84), a80$)
Next i%
Call IOWriteLog(msg$)
msg$ = "ZAFZed: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(zaf.zed!(i%), f84), a80$)
Next i%
Call IOWriteLog(msg$)
msg$ = "ZAFCOR: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(zaf.gensmp!(i%) / zaf.genstd!(i%) * zaf.zed!(i%) / (1# + zaf.vv!(i%)), f84), a80$)
Next i%
Call IOWriteLog(msg$)
msg$ = "UNKRAT: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(zaf.krat!(i%), f84), a80$)
Next i%
Call IOWriteLog(msg$)
msg$ = "UNCONC: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(zaf.conc!(i%), f84), a80$)
Next i%
Call IOWriteLog(msg$)
End If

' Check if iteration is complete
If r0% = 0 Then GoTo 3400   ' iteration completed

If zaf.iter% > MaxZAFIter% Then
msg$ = "WARNING- ZAF not converged on line " & Str$(sample(1).Linenumber&(row%))
Call IOWriteLog(msg$)
GoTo 3400
End If

' Iterate again
zaf.iter% = zaf.iter% + 1
GoTo 2860

' Un-Normalize
3400:
For i% = 1 To zaf.in0%
zaf.conc!(i%) = zaf.conc!(i%) * zaf.ksum!
Next i%

' Load return arrays with analyzed elements
analysis.ZAFIter! = CSng(zaf.iter%)
For i% = 1 To sample(1).LastElm%
analysis.Elsyms$(i%) = sample(1).Elsyms$(i%)
analysis.Xrsyms$(i%) = sample(1).Xrsyms$(i%)
analysis.MotorNumbers%(i%) = sample(1).MotorNumbers%(i%)
analysis.CrystalNames$(i%) = sample(1).CrystalNames$(i%)
analysis.WtPercents!(i%) = zaf.conc!(i%) * 100#

' Load if not disabled
analysis.UnkKrats!(i%) = zaf.krat!(i%)
If zaf.genstd!(i%) <> 0# And (1# + zaf.vv!(i%)) <> 0# And zaf.gensmp!(i%) <> 0# And zaf.il%(i%) <> 16 Then
For j% = 1 To MAXZAFCOR%
If j% = 1 Then analysis.UnkZAFCors!(1, i%) = zaf.gensmp!(i%) / zaf.genstd!(i%)
If j% = 2 Then analysis.UnkZAFCors!(2, i%) = 1# / (1# + zaf.vv!(i%))
If j% = 3 Then analysis.UnkZAFCors!(3, i%) = zaf.zed!(i%)
If j% = 4 Then analysis.UnkZAFCors!(4, i%) = zaf.gensmp!(i%) / zaf.genstd!(i%) * zaf.zed!(i%) / (1# + zaf.vv!(i%))
If j% = 5 Then analysis.UnkZAFCors!(5, i%) = zaf.stp!(i%)
If j% = 6 Then analysis.UnkZAFCors!(6, i%) = zaf.bks!(i%)
If j% = 7 Then analysis.UnkZAFCors!(7, i%) = 1# / zaf.genstd!(i%)
If j% = 8 Then analysis.UnkZAFCors!(8, i%) = 1# / zaf.gensmp!(i%)
Next j%

analysis.UnkMACs!(i%) = ZAFMACCal(i%, zaf)   ' load average MAC for this emitter

' Load actual beam voltage and overvoltage (from coating energy loss)
analysis.ActualKilovolts!(i%) = zaf.coating_actual_kilovolts!(i%)                   ' includes beam energy loss from coating if specified
analysis.EdgeEnergies(i%) = zaf.eC!(i%)
analysis.ActualOvervoltages(i%) = zaf.coating_actual_kilovolts!(i%) / zaf.eC!(i%)   ' includes beam energy loss from coating if specified
End If
Next i%

' Load element by difference
If sample(1).DifferenceElementFlag% And sample(1).DifferenceElement$ <> vbNullString Then
ip% = IPOS1B(sample(1).LastElm% + 1, sample(1).LastChan%, sample(1).DifferenceElement$, sample(1).Elsyms$())
If ip% > sample(1).LastElm% And ip% <= sample(1).LastChan% Then
analysis.WtPercents!(ip%) = zaf.conc!(ip%) * 100#
End If
End If

' Load formula by difference (elements must already be specified)
If sample(1).DifferenceFormulaFlag% And sample(1).DifferenceFormula$ <> vbNullString Then
For i% = sample(1).LastElm% + 1 To sample(1).LastChan%
If ConvertIsDifferenceFormulaElement(sample(1).DifferenceFormula$, sample(1).Elsyms$(i%)) Then
analysis.WtPercents!(i%) = zaf.conc!(i%) * 100#
End If
Next i%
End If

' Load element by stoichiometry to oxygen
If sample(1).StoichiometryElementFlag% And sample(1).StoichiometryElement <> vbNullString Then
ip% = IPOS1B(sample(1).LastElm% + 1, sample(1).LastChan%, sample(1).StoichiometryElement$, sample(1).Elsyms$())
If ip% > sample(1).LastElm% And ip% <= sample(1).LastChan% Then
analysis.WtPercents!(ip%) = zaf.conc!(ip%) * 100#
End If
End If

' Load oxygen if analyzing oxygen or oxygen was specified. See routine ZAFCalZbar for calculation of excess oxygen.
If sample(1).OxygenChannel% > 0 Then
If sample(1).DisableQuantFlag%(sample(1).OxygenChannel%) = 0 Then
analysis.WtPercents!(sample(1).OxygenChannel%) = zaf.conc!(sample(1).OxygenChannel%) * 100#
End If
If sample(1).OxideOrElemental% = 1 Then oxygen! = zaf.conc!(zaf.in0%) * 100#
End If

' Load element relative to another element
If sample(1).RelativeElementFlag% And sample(1).RelativeElement$ <> vbNullString And sample(1).RelativeToElement$ <> vbNullString Then
ip% = IPOS1B(sample(1).LastElm% + 1, sample(1).LastChan%, sample(1).RelativeElement$, sample(1).Elsyms$())
If ip% > sample(1).LastElm% And ip% <= sample(1).LastChan% Then
analysis.WtPercents!(ip%) = zaf.conc!(ip%) * 100#
End If
End If

' Hydrogen stoichiometry to excess oxygen (measured, specified or calculated)
If sample(1).HydrogenStoichiometryFlag And sample(1).HydrogenStoichiometryRatio! <> 0# Then
ip% = IPOS1B(sample(1).LastElm% + 1, sample(1).LastChan%, Symlo$(ATOMIC_NUM_HYDROGEN%), sample(1).Elsyms$())
If ip% > sample(1).LastElm% And ip% <= sample(1).LastChan% Then
analysis.WtPercents!(ip%) = zaf.conc!(ip%) * 100#
End If
End If

' Calculate excess oxygen, total and zbar for this sample
Call ZAFCalZBar(oxygen!, analysis, sample())
If ierror Then Exit Sub
If analysis.TotalPercent! < ZAFMinTotal! * 100# Then GoTo ZAFSmpInsufficientTotal

' Type out correction factors if in DebugMode mode
If VerboseMode Then
Call IOWriteLog(vbCrLf & "ZAF correction factors for unknown line " & Str$(sample(1).Linenumber(row)))

msg$ = "ELEMENT "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%), a80$)
Next i%
Call IOWriteLog(msg$)
msg$ = "ZAFAbs: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.UnkZAFCors!(1, i%), f84), a80$)
Next i%
Call IOWriteLog(msg$)
msg$ = "ZAFFlu: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.UnkZAFCors!(2, i%), f84), a80$)
Next i%
Call IOWriteLog(msg$)
msg$ = "ZAFZed: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.UnkZAFCors!(3, i%), f84), a80$)
Next i%
Call IOWriteLog(msg$)
msg$ = "ZAFCOR: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.UnkZAFCors!(4, i%), f84), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = vbCrLf & "ZAFSTP: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.UnkZAFCors!(5, i%), f84), a80$)
Next i%
Call IOWriteLog(msg$)
msg$ = "ZAFBKS: "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.UnkZAFCors!(6, i%), f84), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = vbCrLf & "STDKFAC "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.StdAssignsKfactors!(i%), f84), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "UNKKFAC "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(analysis.UnkKrats!(i%), f84), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = vbCrLf & "UNKZBAR " & Format$(Format$(analysis.Zbar!, f84), a80$)
Call IOWriteLog(msg$)

msg$ = "ZAFITER " & Format$(Format$(analysis.ZAFIter!, f82$), a80$)
Call IOWriteLog(msg$)

msg$ = "MANITER " & Format$(Format$(analysis.MANIter!, f82$), a80$)
Call IOWriteLog(msg$)

msg$ = "HAL->O2 " & Format$(Format$(analysis.OxygenFromHalogens!, f83$), a80$)
Call IOWriteLog(msg$)

msg$ = vbCrLf & "ELEMENT "
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "UNK WT% "
For i% = 1 To sample(1).LastChan
msg$ = msg$ & Format$(Format$(analysis.WtPercents!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "ZAF%ERR "
For i% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(100# * ZAFDiff!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)
End If

' Output assigned standard coating parameters
If VerboseMode Then
msg$ = vbCrLf & "STDELEM "
For j% = 1 To sample(1).LastElm%
msg$ = msg$ & MiscAutoFormatI$(zaf.coating_std_assigns_element%(j%))
Next j%
Call IOWriteLog(msg$)

msg$ = "STDDENS "
For j% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(zaf.coating_std_assigns_density!(j%), f62$), a80$)
Next j%
Call IOWriteLog(msg$)

msg$ = "STDTHIC "
For j% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(zaf.coating_std_assigns_thickness!(j%), f61$), a80$)
Next j%
Call IOWriteLog(msg$)

msg$ = "STDSINT "
For j% = 1 To sample(1).LastElm%
msg$ = msg$ & Format$(Format$(zaf.coating_std_assigns_sinthickness!(j%), f61$), a80$)
Next j%
Call IOWriteLog(msg$)
End If

' Calculate atomic, oxide, formula, etc
If DebugMode Then
Call ZAFPrintCalculate(zaf, analysis, sample())
If ierror Then Exit Sub

' Print out analytical results for unknown
Call ZAFPrintSmp(zaf, analysis, sample(1).DisplayAsOxideFlag%)
If ierror Then Exit Sub
End If

Next i8% ' loop for particle models
Next i7% ' loop for particle diameters

Exit Sub

' Errors
ZAFSmpError:
MsgBox Error$ & " on sample " & sample(1).Name$ & ", line " & Format$(sample(1).Linenumber&(row%)), vbOKOnly + vbCritical, "ZAFSmp"
ierror = True
Exit Sub

ZAFSmpBadKilovolts:
msg$ = "Kilovolt array is not loaded properly on sample " & sample(1).Name$ & ", line " & Format$(sample(1).Linenumber&(row%))
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFSmp"
ierror = True
Exit Sub

ZAFSmpInsufficientTotal:
msg$ = "WARNING in ZAFSmp- Insufficient total (" & Format$(zaf.ksum!) & ") on line " & Format$(sample(1).Linenumber&(row%))
Call IOWriteLog(msg$)
zerror = True
Exit Sub

ZAFSmpBadStdCounts:
msg$ = "Insufficient standard counts (" & Format$(analysis.StdAssignsCounts!(i%)) & ") on " & sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & " channel " & Format$(i%)
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFSmp"
ierror = True
Exit Sub

ZAFSmpVeryNegativeKratio:
msg$ = "Very negative unknown k-ratio for " & sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & " on channel " & Format$(i%) & " on line " & Format$(sample(1).Linenumber&(row%)) & "."
msg$ = msg$ & vbCrLf & vbCrLf & "Please check for large off-peak interferences and if necessary re-acquire the data with proper background offsets."
msg$ = msg$ & vbCrLf & vbCrLf & "One can also set the Force Negative Kratios To Zero flag in the Analytical | Analysis Options dialog or disable quant for this element in the Elements/Cations dialog."
If Not CalcImageQuantFlag Then
MiscMsgBoxTim FormMSGBOXTIME, "ZAFSmp", msg$, 20#
Call IOWriteLog(msg$)
Else
Call IOWriteLog(msg$)
End If
zerror = True
Exit Sub

ZAFSmpNegativeSumOfKratios:
msg$ = "Negative sum (" & Format$(zaf.ksum!) & ") of unknown k-ratios on line " & Format$(sample(1).Linenumber&(row%)) & "."
msg$ = msg$ & vbCrLf & vbCrLf & "If analyzing only for trace elements be sure that matrix elements are specified using the Specified Concentrations dialog from the Analyze! window."
msg$ = msg$ & vbCrLf & vbCrLf & "Please also check for large off-peak interferences and if necessary re-acquire the data with proper background offsets."
msg$ = msg$ & vbCrLf & vbCrLf & "One can also set the Force Negative Kratios To Zero flag in the Analytical | Analysis Options dialog or disable quant for this element in the Elements/Cations dialog."
If Not CalcImageQuantFlag Then
MiscMsgBoxTim FormMSGBOXTIME, "ZAFSmp", msg$, 20#
Call IOWriteLog(msg$)
Else
Call IOWriteLog(msg$)
End If
zerror = True
Exit Sub

ZAFSmpBadStdKrat:
msg$ = "Bad standard k factor (" & Format$(analysis.StdAssignsKfactors!(i%)) & ") for " & sample(1).Elsyup$(i%) & " " & sample(1).Xrsyms$(i%) & ", on channel " & Format$(i%) & "."
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFSmp"
ierror = True
Exit Sub

ZAFSmpBadEnergyLoss:
msg$ = "Calculated electron energy loss for " & Format$(Symup$(zaf.Z%(i%)), a20$) & " in this matrix is zero for line " & Str$(zaf.n8) & ", and is probably a bad data point (epoxy, etc.). Delete the analysis line and try again."
If Not CalcImageQuantFlag Then
MiscMsgBoxTim FormMSGBOXTIME, "ZAFSmp", msg$, 20#
Call IOWriteLog(msg$)
Else
Call IOWriteLog(msg$)
End If
'ierror = True
Exit Sub

ZAFSmpBadElectronLoss:
msg$ = "Calculated electron backscatter loss for " & Format$(Symup$(zaf.Z%(i%)), a20$) & " in this matrix is zero for line " & Str$(zaf.n8) & ", and is probably a bad data point (epoxy, etc.). Delete the analysis line and try again."
If Not CalcImageQuantFlag Then
MiscMsgBoxTim FormMSGBOXTIME, "ZAFSmp", msg$, 20#
Call IOWriteLog(msg$)
Else
Call IOWriteLog(msg$)
End If
'ierror = True
Exit Sub

End Sub

Sub ZAFStd(row As Integer, analysis As TypeAnalysis, sample() As TypeSample, stdsample() As TypeSample)
' This routine calculates K-factors for each standard using the Tmp sample arrays. The
' algorithims based on John Armstrong's ZAF/Phi-Rho-Z program CITZAF, using
' data in the stdsample array.
'  "row" is the row number in the list of standards to load calculated values
'  analysis.StdZAFCors(1,row%,chan%) = absorption correction
'  analysis.StdZAFCors(2,row%,chan%) = fluorescence correction
'  analysis.StdZAFCors(3,row%,chan%) = atomic number correction
'  analysis.StdZAFCors(4,row%,chan%) = ZAF correction (abscor*flucor*zedcor)
'  analysis.StdZAFCors(5,row%,chan%) = stopping power correction
'  analysis.StdZAFCors(6,row%,chan%) = backscatter correction
'  analysis.StdZAFCors(7,row%,chan%) = primary intensity
'  analysis.StdZAFCors(8,row%,chan%) = sample intensity

ierror = False
On Error GoTo ZAFStdError

Dim i As Integer, j As Integer, ip As Integer, ipp As Integer
Dim tt As Single

ReDim p2(1 To MAXCHAN1%) As Single
ReDim continuum(1 To MAXCHAN1%) As Single

' Initialize arrays
For i% = 1 To MAXCHAN1%
zaf.il%(i%) = 0
zaf.Z%(i%) = 0#
p2!(i%) = 0#
zaf.atwts!(i%) = 0#
Next i%

' Load element symbols and constants
Call ElementGetData(stdsample())
If ierror Then Exit Sub

' Confirm on screen
If VerboseMode Then
Call IOWriteLog(vbCrLf & "Entering ZAFStd...")
Call IOWriteLog("Standard " & Str$(stdsample(1).number%) & ", " & stdsample(1).Name$)
Call IOWriteLog("DisplayAsOxideFlag = " & Str$(stdsample(1).DisplayAsOxideFlag%))
If Not stdsample(1).CombinedConditionsFlag% Then
Call IOWriteLog("Takeoff = " & Str$(stdsample(1).takeoff!) & ", Kilovolts = " & Str$(stdsample(1).kilovolts!))
End If
End If

' Load atomic numbers, atomic weights, x-ray line types and oxide gravimetric factors for each element in the standard.
zaf.TOA! = stdsample(1).takeoff!
For i% = 1 To stdsample(1).LastChan%

' Load default for unanalyzed elements
If i% > stdsample(1).LastElm% Then
stdsample(1).TakeoffArray!(i%) = sample(1).takeoff!
stdsample(1).KilovoltsArray!(i%) = sample(1).kilovolts!
End If

' Calculate takeoff parameter and variables
If stdsample(1).TakeoffArray!(i%) = 0# Then GoTo ZAFStdBadTakeoff
tt! = stdsample(1).TakeoffArray!(i%) * PI! / 180#
zaf.m1!(i%) = 1# / Sin(tt!)

' Check voltage
If stdsample(1).KilovoltsArray!(i%) = 0# Then GoTo ZAFStdBadKilovolts
zaf.eO!(i%) = stdsample(1).KilovoltsArray!(i%)

zaf.Z%(i%) = stdsample(1).AtomicNums%(i%)
zaf.atwts!(i%) = stdsample(1).AtomicWts!(i%)
zaf.il%(i%) = stdsample(1).XrayNums%(i%)
If stdsample(1).DisableQuantFlag%(i%) = 1 Then zaf.il%(i%) = 16    ' use for disabled element

' Standards are ALWAYS elemental!!!!! (OxideOrElemental% = 2), but do p2 calculation anyway for CalcZAF
p2!(i%) = CSng(stdsample(1).numoxd%(i%)) / CSng(stdsample(1).numcat%(i%))
Next i%

' If oxide run or element to oxide conversion and oxygen is not being
' analyzed for, load last element as oxygen by stoichiometry.
zaf.in0% = stdsample(1).LastChan%
zaf.in1% = zaf.in0%
If stdsample(1).OxideOrElemental% = 1 Then
zaf.in0% = zaf.in0% + 1
zaf.in1% = zaf.in0%
zaf.in1% = zaf.in0% - 1

zaf.eO!(zaf.in0%) = sample(1).kilovolts!   ' use sample kilovolts (no need to apply energy loss from coating to stoichiometric oxygen)
zaf.Z%(zaf.in0%) = AllAtomicNums%(ATOMIC_NUM_OXYGEN%)
zaf.il%(zaf.in0%) = 0
p2!(zaf.in0%) = 0#
zaf.atwts!(zaf.in0%) = AllAtomicWts!(ATOMIC_NUM_OXYGEN%)
End If

' Calculate oxide-elemental conversion factors
For i% = 1 To zaf.in1%
zaf.p1!(i%) = p2!(i%) * AllAtomicWts!(ATOMIC_NUM_OXYGEN%) / zaf.atwts!(i%)
Next i%

' Load xray lines, absorption edges, fluorescencent yields, MACs
Call ZAFReadLn(zaf)
If ierror Then Exit Sub
Call ZAFReadMu(zaf)
If ierror Then Exit Sub

' Check for coating correction for x-ray transmission and electron absorption for each emitting element in the sample
For i% = 1 To zaf.in1%
zaf.coating_trans_std!(row%, i%) = 1#
zaf.coating_absorbs_std!(row%, i%) = 1#
If StandardCoatingFlag%(row%) = 1 Then
If zaf.il%(i%) <= MAXRAY% - 1 Then

' Load input parameters for coating calculations
stdsample(1).CoatingDensity! = StandardCoatingDensity!(row%)
stdsample(1).CoatingElement% = StandardCoatingElement%(row%)
stdsample(1).CoatingThickness! = StandardCoatingThickness!(row%)
stdsample(1).CoatingSinThickness! = MathCalculateSinThickness(StandardCoatingThickness!(row%), DefaultTakeOff!)

' Calculate coating correction for x-ray transmission
If UseConductiveCoatingCorrectionForXrayTransmission Then
Call ConvertCalculateCoatingXrayTransmission(zaf.coating_trans_std!(row%, i%), i%, stdsample())
If ierror Then Exit Sub
End If

' Calculate coating correction for electron absorption
If UseConductiveCoatingCorrectionForElectronAbsorption Then
Call ConvertCalculateCoatingElectronAbsorption(zaf.coating_absorbs_std!(row%, i%), i%, stdsample())
If ierror Then Exit Sub

' Calculate beam energy loss for each emitting element in the sample
Call ConvertCalculateElectronEnergy2(zaf.coating_actual_kilovolts!(i%), i%, stdsample())
If ierror Then Exit Sub
End If

End If
End If
Next i%

' Load ZAFPtc defaults based on current model and diameter (normally standards are bulk)
If UseParticleCorrectionFlag And iptc% = 1 Then
Call GetPTCDefaults(Int(0), zaf)
If ierror Then Exit Sub
End If

' Calculate primary intensities
Call ZAFBsc(Int(0))
If ierror Then Exit Sub

Call ZAFMip(Int(0))
If ierror Then Exit Sub

If istp% = 6 Then
Call ZAFAbs(Int(0))
If ierror Then Exit Sub
Call ZAFStp(Int(0))
If ierror Then Exit Sub
Call ZAFBks(Int(0))
If ierror Then Exit Sub

Else
Call ZAFStp(Int(0))
If ierror Then Exit Sub
Call ZAFBks(Int(0))
If ierror Then Exit Sub
Call ZAFAbs(Int(0))
If ierror Then Exit Sub
End If

' Calculate fluorescence
If iflu% < 5 Then
Call ZAFFlu(Int(0), zaf)
If ierror Then Exit Sub
Else
Call ZAFFlu3(Int(0), zaf)
If ierror Then Exit Sub
End If

Call ZAFPrintStd(zaf)
If ierror Then Exit Sub

' Begin ZAF k-ratio calculation
zaf.iter% = 0#
zaf.n8& = stdsample(1).number%

' Load stdpcnts into k-ratio array
For i% = 1 To stdsample(1).LastChan%
analysis.WtPercents!(i%) = stdsample(1).ElmPercents!(i%)

' Check that only the first (not quant disabled) element gets loaded for std k-factor calculation (added 04/15/2014, also see changes in UpdateCalculateUpdateStandard)
If Not MiscIsElementDuplicated2(i%, stdsample(), ipp%) Then

' No duplicate element, just load std percent (only load if not disabled quant) added 12-03-2014
If stdsample(1).DisableQuantFlag%(i%) = 0 Or (stdsample(1).DisableQuantFlag%(i%) = 1 And sample(1).OxideOrElemental% = 1) Then      ' modified 05/13/2015 for Seward (this is correct!!!)
zaf.krat!(i%) = analysis.WtPercents!(i%) / 100#
Else
zaf.krat!(i%) = NotAnalyzedValue!
End If

' Duplicate element found, so just load the first valid occurance of the element, and set duplicates to a small number
Else
ip% = IPOS1DQ(stdsample(1).LastChan%, stdsample(1).Elsyms$(i%), stdsample(1).Elsyms$(), stdsample(1).DisableQuantFlag()) ' only check first non-disabled occurance of each element in standard
For j% = 1 To stdsample(1).LastChan%
If Trim$(UCase$(stdsample(1).Elsyms$(j%))) = Trim$(UCase$(stdsample(1).Elsyms$(ip%))) Then
If j% = ip% Then
zaf.krat!(j%) = analysis.WtPercents!(ip%) / 100#
Else
zaf.krat!(j%) = NotAnalyzedValue! / 100#
End If
End If
Next j%

End If
Next i%

If VerboseMode% Then
msg$ = vbCrLf & "ZAFSTD: Standard " & Format$(stdsample(1).number%) & ":"
Call IOWriteLog(msg$)
For i% = 1 To stdsample(1).LastChan%
msg$ = Format$(i%) & " " & stdsample(1).Elsyms$(i%) & " " & stdsample(1).Xrsyms$(i%) & ", " & MiscAutoFormat$(100# * zaf.krat!(i%))
Call IOWriteLog(msg$)
Next i%
End If

' Calculate K-ratios for these concentrations, (sum wt. frac.)
zaf.ksum! = 0#
For i% = 1 To zaf.in1%
If zaf.krat!(i%) < 0# Then zaf.krat!(i%) = 0#
zaf.ksum! = zaf.ksum! + zaf.krat!(i%)
Next i%

If VerboseMode% Then
msg$ = "Total sum: " & MiscAutoFormat$(100# * zaf.ksum!)
Call IOWriteLog(msg$)
End If

' Add in stoichiometric element (NOT USED FOR STD K-RATIOS!)
If zaf.il%(zaf.in0%) = 0 Then
zaf.krat!(zaf.in0%) = 0#
For i% = 1 To zaf.in1%
zaf.krat!(zaf.in0%) = zaf.krat!(zaf.in0%) + zaf.krat!(i%) * zaf.p1!(i%)
Next i%

' Calculate equivalent oxygen from halogens and subtract from calculated oxygen if flagged
analysis.OxygenFromHalogens! = ConvertHalogensToOxygen(zaf.in1%, sample(1).Elsyms$(), sample(1).DisableQuantFlag%(), zaf.krat!())
If UseOxygenFromHalogensCorrectionFlag Then zaf.krat!(zaf.in0%) = zaf.krat!(zaf.in0%) - analysis.OxygenFromHalogens!
analysis.OxygenFromHalogens! = analysis.OxygenFromHalogens! * 100#  ' convert to weight %

zaf.ksum! = zaf.ksum! + zaf.krat!(zaf.in0%)
End If

' Warn if total is under 98.0 percent (if not unknown sample for demo counts)
If sample(1).Type% = 1 Then
If zaf.ksum! <= 0# Then GoTo ZAFStdZeroSum
If zaf.ksum! * 100# < 98# Then
msg$ = "WARNING in ZAFSTD- Standard " & Str$(stdsample(1).number%) & " total is only " & Format$(zaf.ksum! * 100#, f83$) & " wt.%"
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
End If
If zaf.ksum! * 100# > 102# Then
msg$ = "WARNING in ZAFSTD- Standard " & Str$(stdsample(1).number%) & " total is " & Format$(zaf.ksum! * 100#, f83$) & " wt.%"
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
End If
End If

' Normalize concentrations to 100 % for calculating K-ratios for standards
For i% = 1 To zaf.in0%
zaf.conc!(i%) = zaf.krat!(i%)
zaf.conc!(i%) = zaf.conc!(i%) / zaf.ksum!
Next i%

' Load ZAFPtc defaults based on current model and diameter
If UseParticleCorrectionFlag And iptc% = 1 Then
Call GetPTCDefaults(Int(0), zaf)    ' force thick polished standards for now
If ierror Then Exit Sub
End If

' Calculate ZAFCORS and standard K-factors
Call ZAFMip(Int(1))
If ierror Then Exit Sub

Call ZAFBsc(Int(1))
If ierror Then Exit Sub

If istp% = 6 Then
Call ZAFAbs(Int(1))
If ierror Then Exit Sub
Call ZAFStp(Int(1))
If ierror Then Exit Sub
Call ZAFBks(Int(1))
If ierror Then Exit Sub

Else
Call ZAFStp(Int(1))
If ierror Then Exit Sub
Call ZAFBks(Int(1))
If ierror Then Exit Sub
Call ZAFAbs(Int(1))
If ierror Then Exit Sub
End If

' Calculate fluorescence correction
If iflu% < 5 Then
Call ZAFFlu(Int(1), zaf)
If ierror Then Exit Sub
Else
Call ZAFFlu3(Int(1), zaf)
If ierror Then Exit Sub
End If

' Calculate the K-ratio
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
zaf.stp!(i%) = zaf.stp!(i%) / zaf.s!(i%, i%)
zaf.bks!(i%) = zaf.r!(i%, i%) / zaf.bks!(i%)
zaf.zed!(i%) = zaf.stp!(i%) * zaf.bks!(i%)
zaf.krat!(i%) = zaf.conc!(i%) / zaf.zed!(i%) * (1# + zaf.vv!(i%)) * zaf.genstd!(i%) / zaf.gensmp!(i%)
End If
Next i%

' Denormalize the k-ratios and concentrations
For i% = 1 To zaf.in0%
zaf.conc!(i%) = zaf.conc!(i%) * zaf.ksum!
zaf.krat!(i%) = zaf.krat!(i%) * zaf.ksum!
Next i%

' Calculate total and zbar for standard sample. Oxygen is not calculated for standards!!!!
Call ZAFCalZBar(CSng(0#), analysis, stdsample())
If ierror Then Exit Sub

' Check if this standard contains any of the analyzed elements in the sample
analysis.ZAFIter! = zaf.iter%
For i% = 1 To sample(1).LastElm%

' See if this standard is the assigned standard for this element
If sample(1).StdAssigns%(i%) <> stdsample(1).number% Then GoTo 8400
ip% = IPOS14(i%, sample(), stdsample())  ' check element, xray, take-off and kilovolts
If ip% = 0 Then GoTo 8400

' Load standard assigned percents
analysis.StdAssignsPercents!(i%) = zaf.conc!(ip%) * 100#

' Load assigned standard k ratio
analysis.StdAssignsKfactors!(i%) = zaf.krat!(ip%)

' Load coating x-ray transmission for each element
analysis.Coating_StdAssignsTrans!(i%) = zaf.coating_trans_std!(row%, ip%)      ' coating x-ray transmission
zaf.coating_trans_std_assigns!(i%) = zaf.coating_trans_std!(row%, ip%)         ' coating x-ray transmission

' Load coating electron absorption for each element
analysis.Coating_StdAssignsAbsorbs!(i%) = zaf.coating_absorbs_std!(row%, ip%)  ' coating electron absorption
zaf.coating_absorbs_std_assigns!(i%) = zaf.coating_absorbs_std!(row%, ip%)     ' coating electron absorption

' Load coating parameters for each assigned standard
zaf.coating_std_assigns_element(i%) = stdsample(1).CoatingElement%                 ' coating element for assigned standard coating
zaf.coating_std_assigns_density(i%) = stdsample(1).CoatingDensity!                 ' coating density for assigned standard coating
zaf.coating_std_assigns_thickness(i%) = stdsample(1).CoatingThickness!             ' coating thickness for assigned standard coating
zaf.coating_std_assigns_sinthickness(i%) = stdsample(1).CoatingSinThickness!       ' coating thickness for assigned standard coating

' Load individual matrix factors
If zaf.genstd!(ip%) = 0# Then GoTo 8400
analysis.StdAssignsZAFCors!(1, i%) = zaf.gensmp!(ip%) / zaf.genstd!(ip%)
analysis.StdAssignsZAFCors!(2, i%) = 1# / (1# + zaf.vv!(ip%))
analysis.StdAssignsZAFCors!(3, i%) = zaf.zed!(ip%)
analysis.StdAssignsZAFCors!(4, i%) = zaf.gensmp!(ip%) / zaf.genstd!(ip%) * zaf.zed!(ip%) / (1# + zaf.vv!(ip%))
analysis.StdAssignsZAFCors!(5, i%) = zaf.stp!(ip%)
analysis.StdAssignsZAFCors!(6, i%) = zaf.bks!(ip%)
analysis.StdAssignsZAFCors!(7, i%) = 1# / zaf.genstd!(ip%)
analysis.StdAssignsZAFCors!(8, i%) = 1# / zaf.gensmp!(ip%)

analysis.StdAssignsZbars!(i%) = analysis.Zbar!
analysis.StdAssignsEdgeEnergies!(i%) = zaf.eC!(ip%)

If UseConductiveCoatingCorrectionForElectronAbsorption Then
analysis.StdAssignsActualKilovolts!(i%) = zaf.coating_actual_kilovolts!(ip%)                       ' includes beam energy loss from coating if specified
analysis.StdAssignsActualOvervoltages!(i%) = zaf.coating_actual_kilovolts!(ip%) / zaf.eC!(ip%)     ' includes beam energy loss from coating if specified
Else
analysis.StdAssignsActualKilovolts!(i%) = zaf.eO!(ip%)
analysis.StdAssignsActualOvervoltages!(i%) = zaf.v!(ip%)
End If
8400:
Next i%

' Load standard arrays for this standard into the standard list arrays
For i% = 1 To sample(1).LastChan%
ip% = IPOS14(i%, sample(), stdsample())  ' check element, xray, take-off and kilovolts
If ip% = 0 Then GoTo 8500

' Load the standard percents
analysis.StdPercents!(row%, i%) = stdsample(1).ElmPercents!(ip%)
analysis.StdZbars!(row%) = analysis.Zbar!

' If ZAF corrections are zero don't load values
If zaf.genstd!(ip%) = 0# Then GoTo 8500

' Load standard arrays
If i% <= sample(1).LastElm% Then
analysis.Elsyms$(i%) = stdsample(1).Elsyms$(ip%)
analysis.Xrsyms$(i%) = stdsample(1).Xrsyms$(ip%)
analysis.MotorNumbers%(i%) = stdsample(1).MotorNumbers%(ip%)
analysis.CrystalNames$(i%) = stdsample(1).CrystalNames$(ip%)
analysis.StdZAFCors!(1, row%, i%) = zaf.gensmp!(ip%) / zaf.genstd!(ip%)
analysis.StdZAFCors!(2, row%, i%) = 1# / (1# + zaf.vv!(ip%))
analysis.StdZAFCors!(3, row%, i%) = zaf.zed!(ip%)
analysis.StdZAFCors!(4, row%, i%) = zaf.gensmp!(ip%) / zaf.genstd!(ip%) * zaf.zed!(ip%) / (1# + zaf.vv!(ip%))
analysis.StdZAFCors!(5, row%, i%) = zaf.stp!(ip%)
analysis.StdZAFCors!(6, row%, i%) = zaf.bks!(ip%)
analysis.StdZAFCors!(7, row%, i%) = 1# / zaf.genstd!(ip%)
analysis.StdZAFCors!(8, row%, i%) = 1# / zaf.gensmp!(ip%)

analysis.StdMACs!(row%, i%) = ZAFMACCal(ip%, zaf)   ' load average MAC for this emitter
End If

8500:
Next i%

' Load continuum absorption corrections for this standard
Call ZAFGetContinuumAbsorption(continuum!(), zaf)
If ierror Then Exit Sub

For i% = 1 To sample(1).LastElm%
analysis.StdContinuums!(row%, i%) = continuum!(i%)
Next i%

' Print out the standard k-ratio calculation for this standard if Standard.exe Debugmode or CalcZAFMode = 0
If VerboseMode Or (UCase$(app.EXEName) = UCase$("Standard") And DebugMode) Or (Not VerboseMode And UCase$(app.EXEName) = UCase$("CalcZAF") And CalcZAFMode% = 0) Then
Call ZAFPrintCalculate(zaf, analysis, sample())
If ierror Then Exit Sub
Call ZAFPrintSmp(zaf, analysis, sample(1).DisplayAsOxideFlag)
If ierror Then Exit Sub
End If

Exit Sub

' Errors
ZAFStdError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFStd"
ierror = True
Exit Sub

ZAFStdBadTakeoff:
msg$ = "Takeoff array is not loaded properly"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFStd"
ierror = True
Exit Sub

ZAFStdBadKilovolts:
msg$ = "Kilovolt array is not loaded properly"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFStd"
ierror = True
Exit Sub

ZAFStdZeroSum:
msg$ = "Sum of standard " & Str$(stdsample(1).number%) & " is " & Str$(zaf.ksum! * 100#)
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFStd"
ierror = True
Exit Sub

End Sub

Sub ZAFStp(zafinit As Integer)
' When istp equals:
' 1 = "Stopping Power of Duncumb/Reed (FRAME)"
' 2 = "Stopping Power of Philibert & Tixier"
' 3 = "Stopping Power (Numerical Integration)"
' 4 = "Stopping Power of Love/Scott"
' 5 = "Stopping Power of Pouchou & Pichoir"
' 6 = "Stopping Power of Phi(pz) Integration"

ierror = False
On Error GoTo ZAFStpError

Dim i As Integer, i1 As Integer
Dim n2 As Single, n3 As Single, n4 As Single, n5 As Single
Dim m5 As Single, m6 As Single

Dim emm As Single, wght As Single, omega As Single
Dim te1 As Single, te2 As Single, te3 As Single
Dim veeO As Single, u0 As Single
Dim xpp1 As Single, xpp2 As Single, xpp3 As Single, xp As Single
Dim elinshell As Integer

If zafinit% = 1 Then GoTo 6800

' STDSTP1 / DUNCUMB/REED(FRAME) STOPPING POWER CORRECTION FOR STANDARDS
If istp% = 1 Then
For i% = 1 To zaf.in1%                  ' emitting elements
If zaf.il%(i%) <= MAXRAY% - 1 Then
For i1% = 1 To zaf.in0%                 ' matrix elements
n5! = 1000# * jm!(i1%)
zaf.s!(i1%, i%) = 2# * zaf.Z%(i1%) / (zaf.atwts!(i1%) * (zaf.eO!(i%) + zaf.eC!(i%))) * Log(583# * (zaf.eO!(i%) + zaf.eC!(i%)) / n5!)
Next i1%
End If
Next i%

' STDSTP2 / PHILIBERT and TIXIER STOPPING POWER CORRECTION FOR STANDARDS
ElseIf istp% = 2 Then
For i% = 1 To zaf.in1%
zaf.s!(i%, i%) = 0#
If zaf.il%(i%) <= MAXRAY% - 1 Then
m5! = zaf.Z%(i%) / zaf.atwts!(i%)
m6! = 1.166 * zaf.eC!(i%) / jm!(i%)
n2! = m6! * zaf.v!(i%)
n3! = ZAFFNLint(n2!)
n4! = ZAFFNLint(m6!)
zaf.s!(i%, i%) = (zaf.v!(i%) - 1 - (Log(m6!) / m6!) * (n3! - n4!)) / m5!
zaf.s!(i%, i%) = 1 / zaf.s!(i%, i%)
End If
Next i%

' STDSTP3 / NUMERICAL INTEGRATION STOPPING POWER CORRECTION FOR STANDARDS
ElseIf istp% = 3 Then
Call ZAFQsCalc(Int(0), zaf)
If ierror Then Exit Sub

' STDSTP4 / LOVE/SCOTT STOPPING POWER CORRECTION FOR STANDARDS
ElseIf istp% = 4 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
n5! = jm!(i%)
zaf.s!(i%, i%) = zaf.atwts!(i%) * (1# + 16.05 * (n5! / zaf.eC!(i%)) ^ 0.5 * ((zaf.v!(i%) ^ 0.5 - 1#) / (zaf.v!(i%) - 1#)) ^ 1.07) / zaf.Z%(i%)
zaf.s!(i%, i%) = 1# / zaf.s(i%, i%)
End If
Next i%

' STDSTP5 / POUCHOU and PICHOIR STOPPING POWER CORRECTION FOR STANDARDS
ElseIf istp% = 5 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
jbar! = jm!(i%)

If zaf.il%(i%) = 1 Or zaf.il%(i%) = 2 Then  ' K lines
emm! = 0.86 + 0.12 * Exp(-zaf.Z%(i%) * zaf.Z%(i%) / 25#)
Select Case zaf.Z%(i%)
    Case Is < 12
      wght! = 0#
    Case 12 To 33
      wght! = 2.0988 - 0.569943 * zaf.Z%(i%) + 0.05759217 * zaf.Z%(i%) ^ 2
      wght! = wght! - 0.0026994591 * zaf.Z%(i%) ^ 3 + 0.00006008776 * zaf.Z%(i%) ^ 4
      wght! = wght! - 0.000000514088 * zaf.Z%(i%) ^ 5
    Case Is > 33
      wght! = 0.14
End Select
wght! = 1# - wght!
elinshell% = 2
omega! = Exp(2.373 * Log(zaf.Z%(i%)) - 8.902)

ElseIf zaf.il%(i%) = 3 Or zaf.il%(i%) = 4 Then  ' L lines
emm! = 0.82
wght! = 0.979 - 0.00187 * zaf.Z%(i%)
elinshell% = 4
omega! = Exp(2.946 * Log(zaf.Z%(i%)) - 13.94)

ElseIf zaf.il%(i%) = 5 Or zaf.il%(i%) = 6 Then  ' M lines
emm! = 0.78
wght! = 1#
elinshell% = 6
omega! = 0.5 * Exp(2.946 * Log(zaf.Z%(i%) / 2#) - 13.94)

Else                                            ' additional x-ray lines...???
emm! = 0.78
wght! = 1#
elinshell% = 6
omega! = 0.5 * Exp(2.946 * Log(zaf.Z%(i%) / 2#) - 13.94)
End If

' Save for sample calculations
em!(i%) = emm!
zipi!(i%) = 0.38 * omega! * wght! * elinshell%
te1! = 1.78 - emm!
te2! = 1.1 - emm!
te3! = 0.5 + jbar! / 4 - emm!
veeO! = zaf.eO!(i%) / jbar!
u0! = zaf.v!(i%)
n2! = Exp(te1! * Log(u0!))
n3! = Log(u0!)
xpp1! = 0.0000066 * Exp(0.78 * Log(veeO! / u0!)) * (te1! * n2! * n3! - Exp(te1! * n3!) + 1#) / te1! / te1!
xpp2! = 0.0000112 * (1.35 - 0.45 * jbar! * jbar!) * Exp(0.1 * Log(veeO! / u0!))
xpp2! = xpp2! * (te2! * Exp(te2! * Log(u0!)) * Log(u0!) - Exp(te2! * n3!) + 1#) / te2! / te2!
xpp3! = 0.0000022 / jbar! * Exp((-0.5 + jbar! / 4#) * Log(veeO! / u0!))
xpp3! = xpp3! * (te3! * Exp(te3! * n3!) * n3! - Exp(te3! * n3!) + 1#) / te3! / te3!
xp! = 66892# * zipi!(i%) / zaf.atwts!(i%) * (u0! / veeO! / (zaf.Z%(i%) / zaf.atwts!(i%))) * (xpp1! + xpp2! + xpp3!)
zaf.s!(i%, i%) = 1# / xp!
End If
Next i%

' STDSTP6 / PHI(RZ) INTEGRATION STOPPING POWER CORRECTION FOR STANDARDS
ElseIf istp% = 6 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
zaf.r!(i%, i%) = phi!(i%) * 1000
zaf.s!(i%, i%) = 1#
End If
Next i%
End If

Exit Sub

' SMPSTP1 / DUNCUMB/REED (FRAME) STOPPING POWER CORRECTION FOR SAMPLE
6800:
If istp% = 1 Then
For i% = 1 To zaf.in1%              ' emitting elements
zaf.stp!(i%) = 0#
If zaf.il%(i%) <= MAXRAY% - 1 Then
For i1% = 1 To zaf.in0%             ' matrix elements
zaf.stp!(i%) = zaf.stp!(i%) + zaf.conc!(i1%) * zaf.s!(i1%, i%)
Next i1%
End If
Next i%

' SMPSTP2 / PHILIBERT and TIXIER STOPPING POWER CORRECTION FOR SAMPLE
ElseIf istp% = 2 Then
m5! = sumatom!
For i% = 1 To zaf.in1%
zaf.stp!(i%) = 0#
If zaf.il%(i%) <= MAXRAY% - 1 Then
m6! = 0#
For i1% = 1 To zaf.in0%
m6! = m6! + (zaf.conc!(i1%) * zaf.Z%(i1%) / zaf.atwts!(i1%)) * Log(1.166 * zaf.eC!(i%) / jm!(i1%))
Next i1%
m6! = Exp(m6! / m5!)
n2! = m6! * zaf.v!(i%)
n3! = ZAFFNLint(n2!)
n4! = ZAFFNLint(m6!)
zaf.stp!(i%) = (zaf.v!(i%) - 1 - (Log(m6!) / m6!) * (n3! - n4!)) / m5!
zaf.stp!(i%) = 1 / zaf.stp!(i%)
End If
Next i%

' SMPSTP3 / NUMERICAL INTEGRATION STOPPING POWER CORRECTION FOR SAMPLE
ElseIf istp% = 3 Then
Call ZAFQsCalc(Int(1), zaf)
If ierror Then Exit Sub

' SMPSTP4 / LOVE/SCOTT STOPPING POWER CORRECTION FOR SAMPLE
ElseIf istp% = 4 Then
m5 = sumatom!
For i% = 1 To zaf.in1%
zaf.stp!(i%) = 0
If zaf.il%(i%) <= MAXRAY% - 1 Then
m6! = jbar!
zaf.stp!(i%) = (1# + 16.05 * (m6! / zaf.eC!(i%)) ^ 0.5 * ((zaf.v!(i%) ^ 0.5 - 1) / (zaf.v!(i%) - 1)) ^ 1.07) / m5!
zaf.stp!(i%) = 1# / zaf.stp!(i%)
End If
Next i%

' SMPSTP5 / POUCHOU and PICHOIR STOPPING POWER CORRECTION FOR SAMPLE
ElseIf istp% = 5 Then
m5! = sumatom!
For i% = 1 To zaf.in1%
If zaf.il%(i) <= MAXRAY% - 1 Then
veeO! = zaf.eO!(i%) / jbar!
emm! = em!(i%)
u0! = zaf.v!(i%)
te1! = 1.78 - emm!
te2! = 1.1 - emm!
te3! = 0.5 + jbar! / 4 - emm!
n2! = Exp(te1! * Log(u0!))
n3! = Log(u0!)
xpp1! = 0.0000066 * Exp(0.78 * Log(veeO! / u0!)) * (te1! * n2! * n3! - Exp(te1! * n3!) + 1#) / te1! / te1!
xpp2! = 0.0000112 * (1.35 - 0.45 * jbar! * jbar!) * Exp(0.1 * Log(veeO! / u0!))
xpp2! = xpp2! * (te2! * Exp(te2! * Log(u0!)) * Log(u0!) - Exp(te2! * n3!) + 1#) / te2! / te2!
xpp3! = 0.0000022 / jbar! * Exp((-0.5 + jbar! / 4) * Log(veeO! / u0!))
xpp3! = xpp3! * (te3! * Exp(te3! * n3!) * n3! - Exp(te3! * n3!) + 1#) / te3! / te3!
xp! = 66892# * zipi!(i%) / zaf.atwts!(i%) * ((u0! / veeO!) / m5!) * (xpp1! + xpp2! + xpp3!)
zaf.stp!(i%) = 1# / xp!
End If
Next i%

' SMPSTP6 / PHI(RZ) INTEGRATION STOPPING POWER CORRECTION FOR SAMPLE
ElseIf istp% = 6 Then
For i% = 1 To zaf.in1%
zaf.bks!(i%) = 0#
zaf.stp!(i%) = 0#
If zaf.il%(i%) <= MAXRAY% - 1 Then
zaf.bks!(i%) = phi!(i%) * 1000#
zaf.stp!(i%) = 1#
End If
Next i%
End If

Exit Sub

' Errors
ZAFStpError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFStp"
ierror = True
Exit Sub

End Sub

Sub ZAFCalculateRange(analysis As TypeAnalysis, sample() As TypeSample)
' Calculate electron and x-ray range for CalcZAF

ierror = False
On Error GoTo ZAFCalculateRangeError

' Transmission for different thicknesses (microns)
Call ZAFCalculateRanges(zaf.mup!(), analysis, sample())
If ierror Then Exit Sub

Exit Sub

' Errors
ZAFCalculateRangeError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFCalculateRange"
ierror = True
Exit Sub

End Sub

Sub ZAFPrintMAC()
' Print current ZAF mass absorption coefficients

ierror = False
On Error GoTo ZAFPrintMACError

Call ZAFReadMu(zaf)
If ierror Then Exit Sub

Call ZAFPrintMAC2(zaf)
If ierror Then Exit Sub

Exit Sub

' Errors
ZAFPrintMACError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFPrintMAC"
ierror = True
Exit Sub

End Sub

Sub ZAFPrintAPF()
' Print current ZAF area peak factors

ierror = False
On Error GoTo ZAFPrintAPFError

Call ZAFPrintAPF2(zaf)
If ierror Then Exit Sub

Exit Sub

' Errors
ZAFPrintAPFError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFPrintAPF"
ierror = True
Exit Sub

End Sub

Sub ZAFSetZAF(sample() As TypeSample)
' This routine initializes and sets up ZAF correction parameters, and loads elemental data for the (unknown) sample.

ierror = False
On Error GoTo ZAFSetZAFError

Dim i As Integer
Dim tt As Single, m9 As Single

ReDim p2(1 To MAXCHAN1%) As Single

If VerboseMode Then Call IOWriteLog(vbCrLf & "Loading element correction parameters (Entering SetZAF)...")

' Init ZAF (do not call from ZAFStd to avoid issues with PTC calculations!)
Call ZAFInitZAF
If ierror Then Exit Sub

' Load take-off angle
If sample(1).takeoff! = 0# Then sample(1).takeoff! = DefaultTakeOff!

If VerboseMode Then
Call IOWriteLog("TakeOff = " & Str$(sample(1).takeoff!) & ", Kilovolts = " & Str$(sample(1).kilovolts!))
Call IOWriteLog("OxideOrElemental = " & Str$(sample(1).OxideOrElemental%))
End If

For i% = 1 To MAXCHAN% + 1
zaf.il%(i%) = 0
zaf.Z%(i%) = 0
p2!(i%) = 0#
zaf.atwts!(i%) = 0#
Next i%

' Calculate oxygen channel
Call ZAFGetOxygenChannel(sample())
If ierror Then Exit Sub

' Load atomic numbers, atomic weights, x-ray line types (for analyzed elements only), and oxide gravimetric factor for each element in the probe run.
zaf.in0% = sample(1).LastChan%
zaf.TOA! = sample(1).takeoff!

' Load element arrays
For i% = 1 To sample(1).LastChan%
If i% <= sample(1).LastElm% And sample(1).KilovoltsArray!(i%) = 0# Then GoTo ZAFSetZAFBadKilovolts
If sample(1).TakeoffArray!(i%) = 0# Then sample(1).TakeoffArray!(i%) = sample(1).takeoff!

' Bulk sample geometry
tt! = sample(1).TakeoffArray!(i%) * 3.14159 / 180#
zaf.m1!(i%) = 1# / Sin(tt!)

' ZAFPtc geometry
If UseParticleCorrectionFlag And iptc% = 1 Then
zaf.t1! = 1 / Sin(tt!)
zaf.t2! = Tan(tt!)
zaf.t3! = 1 / Cos(tt!)
zaf.t4! = 1 / zaf.t2!
zaf.g! = zaf.t3! / (1 + zaf.t2!)    ' geometry
End If

' Load beam energies
If sample(1).KilovoltsArray!(i%) = 0# Then sample(1).KilovoltsArray!(i%) = sample(1).kilovolts!
zaf.eO!(i%) = sample(1).KilovoltsArray!(i%)

zaf.Z%(i%) = sample(1).AtomicNums%(i%)
zaf.atwts!(i%) = sample(1).AtomicWts!(i%)
zaf.il%(i%) = sample(1).XrayNums%(i%)
'If sample(1).DisableQuantFlag%(i%) = 1 Then zaf.il%(i%) = 10    ' use for disabled element
If sample(1).DisableQuantFlag%(i%) = 1 Then zaf.il%(i%) = 15    ' use for disabled element

p2!(i%) = 0#
If sample(1).OxideOrElemental% = 1 Or sample(1).numoxd%(i%) <> 0 Then
If sample(1).numcat%(i%) < 1 Then GoTo ZAFSetZAFNoCations
p2!(i%) = CSng(sample(1).numoxd%(i%)) / CSng(sample(1).numcat%(i%))
End If
Next i%

' If oxide run or element to oxide conversion and oxygen is not being
' analyzed for, load last element as oxygen by stoichiometry.
zaf.in1% = zaf.in0%
If sample(1).OxideOrElemental% = 1 Then
zaf.in0% = zaf.in0% + 1                     ' array index for all matrix elements (including oxygen by stoichiometry)
zaf.in1% = zaf.in0%
zaf.in1% = zaf.in0% - 1                     ' array index for all emitting elements (not including oxygen by stoichiometry)

zaf.eO!(zaf.in0%) = sample(1).kilovolts!
zaf.Z%(zaf.in0%) = AllAtomicNums%(ATOMIC_NUM_OXYGEN%)
zaf.il%(zaf.in0%) = 0
p2!(zaf.in0%) = 0#
zaf.atwts!(zaf.in0%) = AllAtomicWts!(ATOMIC_NUM_OXYGEN%)
End If

' Calculate oxide-elemental conversion factors
For i% = 1 To zaf.in1%
If zaf.atwts!(i%) = 0# Then GoTo ZAFSetZAFBadAtomicWeight
zaf.p1(i%) = p2!(i%) * AllAtomicWts!(ATOMIC_NUM_OXYGEN%) / zaf.atwts!(i%)
Next i%

' Load xray lines, absorption edges, fluorescencent yields, MACs
Call ZAFReadLn(zaf)
If ierror Then Exit Sub
Call ZAFReadMu(zaf)
If ierror Then Exit Sub

' Check for coating absorption correction for each emitting element in the sample
zaf.coating_flag% = sample(1).CoatingFlag%
For i% = 1 To zaf.in1%
zaf.coating_trans_smp!(i%) = 1#
zaf.coating_absorbs_smp!(i%) = 1#
If sample(1).CoatingFlag% = 1 Then
If zaf.il%(i%) <= MAXRAY% - 1 Then

If UseConductiveCoatingCorrectionForXrayTransmission Then
Call ConvertCalculateCoatingXrayTransmission(zaf.coating_trans_smp!(i%), i%, sample())
If ierror Then Exit Sub
End If

' Check for coating electron absorption for each emitting element in the sample
If UseConductiveCoatingCorrectionForElectronAbsorption Then
Call ConvertCalculateCoatingElectronAbsorption(zaf.coating_absorbs_smp!(i%), i%, sample())
If ierror Then Exit Sub

' Calculate beam energy loss for each emitting element in the sample
Call ConvertCalculateElectronEnergy2(zaf.coating_actual_kilovolts!(i%), i%, sample())
If ierror Then Exit Sub
End If

End If
End If
Next i%

' Load ZAFPtc defaults based on current model and diameter
If UseParticleCorrectionFlag And iptc% = 1 Then
Call GetPTCDefaults(Int(0), zaf)
If ierror Then Exit Sub
End If

' Calculate correction factors
Call ZAFBsc(Int(0))
If ierror Then Exit Sub

Call ZAFMip(Int(0))
If ierror Then Exit Sub

If istp% = 6 Then
Call ZAFAbs(Int(0))
If ierror Then Exit Sub
Call ZAFStp(Int(0))
If ierror Then Exit Sub
Call ZAFBks(Int(0))
If ierror Then Exit Sub

Else
Call ZAFStp(Int(0))
If ierror Then Exit Sub
Call ZAFBks(Int(0))
If ierror Then Exit Sub
Call ZAFAbs(Int(0))
If ierror Then Exit Sub
End If

' Calculate fluorescence
If iflu% < 5 Then
Call ZAFFlu(Int(0), zaf)
If ierror Then Exit Sub
Else
Call ZAFFlu3(Int(0), zaf)
If ierror Then Exit Sub
End If

' Check for large absorption corrections for "genstd!(i%)"
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
m9 = 1# / zaf.genstd!(i%)
If m9 < 0.4 Then
msg$ = "WARNING in ZAFSetZAF- the f(x) of " & Format$(Symup$(zaf.Z%(i%)), a20$) & " " & Format$(Xraylo$(zaf.il%(i%)), a20$) & " is " & Format$(Format$(m9, f84), a80$)
Call IOWriteLog(msg$)
End If
End If
Next i%

' Print primary intensity correction factors
If VerboseMode% Then
Call ZAFPrintStd(zaf)
If ierror Then Exit Sub
End If

Exit Sub

' Errors
ZAFSetZAFError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFSetZAF"
ierror = True
Exit Sub

ZAFSetZAFBadKilovolts:
msg$ = "Kilovolt array is not loaded properly"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFSetZAF"
ierror = True
Exit Sub

ZAFSetZAFNoCations:
msg$ = "Element " & sample(1).Elsyms$(i%) & " has no cations specified"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFSetZAF"
ierror = True
Exit Sub

ZAFSetZAFBadAtomicWeight:
msg$ = "Element " & sample(1).Elsyms$(i%) & " has no atomic weight"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFSetZAF"
ierror = True
Exit Sub

End Sub

Sub ZAFInitZAF()
' Initialize the ZAF arrays

ierror = False
On Error GoTo ZAFInitZAFError

Dim i As Integer, j As Integer

' Dimension dynamic arrays
ReDim zaf.mup(1 To MAXCHAN1%, 1 To MAXCHAN1%) As Single   ' mass absorption coefficients
ReDim zaf.r(1 To MAXCHAN1%, 1 To MAXCHAN1%) As Single     ' pure element backscatter loss
ReDim zaf.s(1 To MAXCHAN1%, 1 To MAXCHAN1%) As Single     ' pure element sample stopping power

ReDim zaf.coating_trans_std(1 To MAXSTD%, 1 To MAXCHAN1%) As Single       ' x-ray transmission for standard coating
ReDim zaf.coating_absorbs_std(1 To MAXSTD%, 1 To MAXCHAN1%) As Single     ' electron absorption for standard coating

' Initialize ZED arrays
zaf.in0% = 0
zaf.in1% = 0
zaf.n8& = 0
zaf.TOA! = 0#

For i% = 1 To MAXCHAN1%
zaf.m1!(i%) = 0#
zaf.eO!(i%) = 0#
zaf.p1!(i%) = 0#
Next i%

For i% = 1 To MAXCHAN1%
For j% = 1 To MAXCHAN1%
zaf.r!(j%, i%) = 0#
zaf.s!(j%, i%) = 0#
zaf.mup!(j%, i%) = 0#
Next j%

zaf.il%(i%) = 0
zaf.atwts!(i%) = 0#
zaf.kraw!(i%) = 0#
zaf.krat!(i%) = 0#
zaf.conc!(i%) = 0#
zaf.gensmp!(i%) = 0#
zaf.genstd!(i%) = 0#
zaf.eC!(i%) = 0#

For j% = 1 To MAXRAY% - 1
zaf.eng!(j%, i%) = 0#
zaf.flu!(j%, i%) = 0#
Next j%

For j% = 1 To MAXEDG%
zaf.edg!(j%, i%) = 0#
Next j%

zaf.v!(i%) = 0#     ' over voltage
zaf.vv!(i%) = 0#    ' fluorescence correction
zaf.Z%(i%) = 0      ' atomic number
zaf.stp!(i%) = 0#
zaf.bks!(i%) = 0#
zaf.zed!(i%) = 0#
Next i%

' ZAFPtc variables
zaf.t1! = 0#
zaf.t2! = 0#
zaf.t3! = 0#
zaf.t4! = 0#

zaf.model% = 0          ' current particle model
For i% = 1 To MAXMODELS%
zaf.models%(i%) = 0
Next i%

zaf.diam! = 0#          ' current diameter in microns
For i% = 1 To MAXDIAMS%
zaf.diams!(i%) = 0#
Next i%

zaf.d! = 0#     ' diameter in cm
zaf.rho! = 0#   ' density in g/cm^3
zaf.j9! = 0#    ' particle thicknes factor
zaf.x1! = 0#    ' integration step length in g/cm^2

zaf.TotalCations! = 0#
zaf.totalatoms! = 0#
For i% = 1 To MAXCHAN1%
zaf.Formulas!(i%) = 0#
zaf.OxPercents!(i%) = 0#
zaf.AtPercents!(i%) = 0#
zaf.NormElPercents!(i%) = 0#
zaf.NormOxPercents!(i%) = 0#
Next i%

' Coating calculations (emitting elements only)
zaf.coating_flag% = 0                                  ' unknown sample coating flag
zaf.coating_sin_thickness! = 0#                        ' x-ray absorption path length

For i% = 1 To MAXCHAN%
zaf.coating_trans_smp(i%) = 1#                         ' x-ray transmission for sample coating
zaf.coating_trans_std_assigns!(i%) = 1#                ' x-ray transmission for assigned standards
zaf.coating_absorbs_smp(i%) = 1#                       ' electron absorption for sample coating
zaf.coating_absorbs_std_assigns!(i%) = 1#              ' electron absorption for assigned standards

For j% = 1 To MAXSTD%
zaf.coating_trans_std(j%, i%) = 1#                     ' x-ray transmission for standard coating
zaf.coating_absorbs_std(j%, i%) = 1#                   ' electron absorption for standard coating
Next j%

zaf.coating_actual_kilovolts(i%) = 0#                  ' actual beam energy (after electron absorption beam energy loss)
Next i%

Exit Sub

' Errors
ZAFInitZAFError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFInitZAF"
ierror = True
Exit Sub

End Sub

Sub ZAFBks(zafinit As Integer)
' When ibks equals:
' 0 = "No Backscatter (used for Packwood and Bastin original)"
' 1 = "Backscatter of Duncumb & Reed (FRAME-I)"
' 2 = "Backscatter of Duncumb & Reed (COR-II) & Heinrich"
' 3 = "Backscatter of Duncumb & Reed (COR-II) & Heinrich"
' 4 = "Backscatter of Love/Scott"
' 5 = "Backscatter of Myklebust-I"
' 6 = "Backscatter of Myklebust & Fiori"
' 7 = "Backscatter of Pouchou & Pichoir"
' 8 = "Backscatter of August, Razka & Wernisch"
' 9 = "Backscatter of Springer"

ierror = False
On Error GoTo ZAFBksError

Const MAXAUG% = 5

Dim i As Integer, i1 As Integer, i2 As Integer, ij As Integer
Dim u0 As Single
Dim ju0 As Single, gu0 As Single
Dim meanw As Single, Alpha As Single
Dim w1 As Single, w2 As Single, w3 As Single, w4 As Single, w5 As Single
Dim z1 As Single, z2 As Single, z3 As Single, z4 As Single, z5 As Single
Dim n1 As Single, n2 As Single, n3 As Single, n4 As Single, n5 As Single

Static ju(1 To MAXCHAN1%) As Single
Static gu(1 To MAXCHAN%) As Single

ReDim v2(1 To MAXCHAN%) As Single
ReDim aug(1 To MAXAUG%, 1 To MAXAUG%) As Double

If zafinit% = 1 Then GoTo 2100

' STDBKS1 / DUNCUMB and REED (FRAME) BACKSCATTER CORRECTION FOR STANDARDS
If ibks% = 1 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
v2!(i%) = zaf.v!(i%)
If v2!(i%) > 10# Then v2!(i%) = 10#
n2! = 0.00873 * v2!(i%) ^ 3 - 0.1669 * v2!(i%) ^ 2 + 0.9662 * v2!(i%) + 0.4523
n3! = 0.002703 * v2!(i%) ^ 3 - 0.05182 * v2!(i%) ^ 2 + 0.302 * v2!(i%) - 0.1836
n4! = 0.887 - 3.44 / v2!(i%) + 9.33 / v2!(i%) ^ 2 - 6.43 / v2!(i%) ^ 3

For i1% = 1 To zaf.in0%
zaf.r!(i1%, i%) = n2! - n3! * Log(n4! * zaf.Z%(i1%) + 25#)
Next i1%
End If
Next i%

' STDBKS2 and STDBKS3 / DUNCUMB and REED and HEINRICH BACKSCATTER CORRECTION FOR STANDARDS
ElseIf ibks% = 2 Or ibks% = 3 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
w1! = zaf.eC!(i%) / zaf.eO!(i%)
w2! = w1! * w1!
w3! = w2! * w1!
w4! = w3! * w1!
w5! = w4! * w1!
For i1% = 1 To zaf.in0%
z1! = zaf.Z%(i1%)
z2! = z1! * z1!
z3! = z2! * z1!
z4! = z3! * z1!
z5! = z4! * z1!

If ibks% = 2 Then
n1! = (-0.581 + 2.162 * w1! - 5.137 * w2! + 9.213 * w3! - 8.619 * w4! + 2.962 * w5!) * 0.01 * z1!
n2! = (-1.609 - 8.298 * w1! + 28.791 * w2! - 47.744 * w3! + 46.54 * w4! - 17.676 * w5!) * 0.0001 * z2!
n3! = (5.4 + 19.184 * w1! - 75.733 * w2! + 120.05 * w3! - 110.7 * w4! + 41.792 * w5!) * 0.000001 * z3!
n4! = (-5.725 - 21.645 * w1! + 88.128 * w2! - 136.06 * w3! + 117.75 * w4! - 42.445 * w5!) * 0.00000001 * z4!
n5! = (2.095 + 8.947 * w1! - 36.51 * w2! + 55.694 * w3! - 46.079 * w4! + 15.851 * w5!) * 0.0000000001 * z5!
zaf.r!(i1%, i%) = 1# + n1! + n2! + n3! + n4! + n5!
End If

If ibks% = 3 Then
n1! = zaf.atwts!(i1%) / (-0.7585 + 2.058183 * z1! + 0.005077 * z2!)
zaf.r!(i1%, i%) = 1# - ((1# - zaf.r!(i1%, i%)) * n1!)
End If
        
Next i1%
End If
Next i%

' STDBKS4 / LOVE/SCOTT BACKSCATTER CORRECTION FOR STANDARDS
ElseIf ibks% = 4 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
u0! = zaf.v!(i%)
ju!(i%) = 0.33148 * Log(u0!) + 0.05596 * Log(u0!) ^ 2 - 0.06339 * Log(u0!) ^ 3 + 0.00947 * Log(u0!) ^ 4
gu!(i%) = 2.87898 * Log(u0!) - 1.51307 * Log(u0!) ^ 2 + 0.81313 * Log(u0!) ^ 3
gu!(i%) = (gu!(i%) - 0.08241 * Log(u0!) ^ 4) / u0!
End If
Next i%

For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
zaf.r!(i%, i%) = 1# - hb!(i%) * (ju!(i%) + hb!(i%) * gu!(i%)) ^ 1.67
End If
Next i%

' STDBKS5 / MYKLEBUST-I BACKSCATTER CORRECTION FOR STANDARDS
ElseIf ibks% = 5 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
For i1% = 1 To zaf.in0%
zaf.r!(i1%, i%) = 0.0000361325 * zaf.Z%(i1%) * zaf.Z%(i1%) + 0.00958267 * zaf.Z%(i1%) * Exp(-zaf.v!(i%))
zaf.r!(i1%, i%) = zaf.r!(i1%, i%) + 1# - 0.00815168 * zaf.Z%(i1%) + 0.001141 * zaf.eO!(i%)
Next i1%
End If
Next i%

' STDBKS6 / MYKLEBUST & FIORI BACKSCATTER CORRECTION FOR STANDARDS
ElseIf ibks% = 6 Then
' Not implemented yet

' STDBKS7 / POUCHOU and PICHOIR BACKSCATTER CORRECTION FOR STANDARDS
ElseIf ibks% = 7 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
meanw! = 0.595 + hb!(i%) / 3.7 + Exp(4.55 * Log(hb!(i%)))
u0! = zaf.v!(i%)
ju0! = 1 + u0! * (Log(u0!) - 1#)
Alpha! = (2# * meanw! - 1#) / (1# - meanw!)
gu0! = (u0! - 1# - (1# - Exp((Alpha! + 1#) * Log(1# / u0!))) / (1# + Alpha!)) / (2# + Alpha!) / ju0!
zaf.r!(i%, i%) = 1# - hb!(i%) * meanw! * (1# - gu0!)
End If
Next i%

' STDBKS8 / AUGUST, RAZKA & WERNISCH BACKSCATTER CORRECTION FOR STANDARDS
ElseIf ibks% = 8 Then
For i1% = 1 To MAXAUG%
For i2% = 1 To MAXAUG%
aug#(i1%, i2%) = 0
Next i2%
Next i1%
      
aug#(1, 1) = 0.005580848699
aug#(1, 2) = 0.0002709177328
aug#(1, 3) = -0.000005531081141
aug#(1, 4) = 5.955796251E-08
aug#(1, 5) = -3.210316856E-10
aug#(2, 1) = 0.03401533559
aug#(2, 2) = -0.0001601761397
aug#(2, 3) = 0.000002473523226
aug#(2, 4) = -3.020861042E-08
aug#(3, 1) = 0.09916651666
aug#(3, 2) = -0.0004615018255
aug#(3, 3) = -4.332933627E-07
aug#(4, 1) = 0.1030099792
aug#(4, 2) = -0.0003113053618
aug#(5, 1) = 0.03630169747
      
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
For i1% = 1 To zaf.in0%
aug#(5, 5) = 1#
For ij% = 1 To MAXAUG%
For i2% = 1 To ij%
aug#(5, 5) = aug#(5, 5) + aug#(i2%, ij% - i2% + 1) * ((1 / zaf.v!(i%) - 1) ^ i2%) * (zaf.Z%(i1%) ^ (ij% - i2% + 1))
Next i2%
Next ij%
zaf.r!(i1%, i%) = aug#(5, 5)
Next i1%
End If
Next i%

' STDBKS9 / SPRINGER BACKSCATTER CORRECTION FOR STANDARDS
ElseIf ibks% = 9 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
w1! = zaf.eC!(i%) / zaf.eO!(i%)
w2! = w1! * w1!
w3! = w2! * w1!
w4! = w3! * w1!
For i1% = 1 To zaf.in0%
z1! = zaf.Z%(i1%)
z2! = z1! * z1!
z3! = z2! * z1!
z4! = z3! * z1!
n1! = (100.88 - 0.7607 * z1! - 0.0035702 * z2! + 0.00016329 * z3! - 0.00000096521 * z4!)
n2! = (-0.61134 + 0.60271 * z1! + 0.016222 * z2! - 0.00045936 * z3! + 0.0000025267 * z4!) * w1!
n3! = (-0.91447 + 2.9326 * z1! - 0.17636 * z2! + 0.0028558 * z3! - 0.000013294 * z4!) * w2!
n4! = (-0.70753 - 4.6855 * z1! + 0.29116 * z2! - 0.0046797 * z3! + 0.000021597 * z4!) * w3!
n5! = (1.3735 + 1.9015 * z1! - 0.12703 * z2! + 0.0021144 * z3! - 0.0000098423 * z4!) * w4!
zaf.r!(i1%, i%) = (n1! + n2! + n3! + n4! + n5!) / 100#
Next i1%
End If
Next i%

' STDBKS10 / DONOVAN (MODIFIED DUNCUMB/REED) BACKSCATTER CORRECTION FOR STANDARDS (same as #1)
ElseIf ibks% = 10 Then
For i% = 1 To zaf.in1%
If zaf.il%(i%) <= MAXRAY% - 1 Then
n2! = 0.00873 * zaf.v!(i%) ^ 3 - 0.1669 * zaf.v!(i%) ^ 2 + 0.9662 * zaf.v!(i%) + 0.4523
n3! = 0.002703 * zaf.v!(i%) ^ 3 - 0.05182 * zaf.v!(i%) ^ 2 + 0.302 * zaf.v!(i%) - 0.1836
n4! = 0.887 - 3.44 / zaf.v!(i%) + 9.33 / zaf.v!(i%) ^ 2 - 6.43 / zaf.v!(i%) ^ 3

For i1% = 1 To zaf.in0%
zaf.r!(i1%, i%) = n2! - n3! * Log(n4! * zaf.Z%(i1%) + 25#)
Next i1%
End If
Next i%
End If

Exit Sub

' SMPBKS1,2,3,5,8,9 / MASS FRACTION SUM BACKSCATTER CORRECTION FOR SAMPLE
2100:
If ibks% = 1 Or ibks% = 2 Or ibks% = 3 Or ibks% = 5 Or ibks% = 8 Or ibks% = 9 Then
For i% = 1 To zaf.in1%
zaf.bks!(i%) = 0#   ' must be zero for summation below (fixed 09-14-2012)
If zaf.il%(i%) <= MAXRAY% - 1 Then

' Mass fraction sum of eta
For i1% = 1 To zaf.in0%
zaf.bks!(i%) = zaf.bks!(i%) + zaf.conc!(i1%) * zaf.r!(i1%, i%)
Next i1%
End If
Next i%

' SMPBKS4 / LOVE and SCOTT BACKSCATTER CORRECTION FOR SAMPLE
ElseIf ibks% = 4 Then
For i% = 1 To zaf.in1%
zaf.bks!(i%) = 0#   ' init to zero for summation below
If zaf.il%(i%) <= MAXRAY% - 1 Then
If (ju!(i%) + eta!(i%) * gu!(i%)) < 0# Then GoTo ZAFBksNegative
zaf.bks!(i%) = 1# - eta!(i%) * (ju!(i%) + eta!(i%) * gu!(i%)) ^ 1.67
End If
Next i%

' SMPBKS6 / MYKLEBUST & FIORI BACKSCATTER CORRECTION FOR SAMPLE
ElseIf ibks% = 6 Then
' Not implemented yet

' SMPBKS7 / POUCHOU and PICHOIR BACKSCATTER CORRECTION FOR SAMPLE
ElseIf ibks% = 7 Then
For i% = 1 To zaf.in1%
zaf.bks!(i%) = 1#
If zaf.il%(i%) <= MAXRAY% - 1 Then
meanw! = 0.595 + eta!(i%) / 3.7 + Exp(4.55 * Log(eta!(i%)))
u0! = zaf.v!(i%)    ' typo fixed 5-23-2005 (was zaf.z(i%))
ju0! = 1 + u0! * (Log(u0!) - 1#)
Alpha! = (2# * meanw! - 1#) / (1# - meanw!)
gu0! = (u0! - 1# - (1# - Exp((Alpha! + 1#) * Log(1# / u0!))) / (1# + Alpha!)) / (2# + Alpha!) / ju0!
zaf.bks!(i%) = 1# - eta!(i%) * meanw! * (1# - gu0!)
End If
Next i%
End If

Exit Sub

' Errors
ZAFBksError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFBks"
ierror = True
Exit Sub

ZAFBksNegative:
msg$ = vbCrLf & "Bad (negative) bks parameter calculated for the sample analysis. This usually indicates negative concentrations so you should check that you are not analyzing epoxy." & vbCrLf
msg$ = msg$ & "You should also make sure your off-peak background and interference corrections are not overcorrecting, or perhaps you have assigned a blank correction to a major or minor element and you did not enter the correct blank level in the Standard Assignments dialog."
If Not CalcImageQuantFlag Then
MiscMsgBoxTim FormMSGBOXTIME, "ZAFBks", msg$, 20#
Call IOWriteLog(msg$)
Else
Call IOWriteLog(msg$)
End If
'ierror = True
Exit Sub

End Sub

Sub ZAFPhi(ii As Integer, uu As Single, hh As Single, zz As Single, v1 As Single, rr As Single)
' Phi-Rho-Z calculation
' When iphi equals:
' 1 = "Phi(pz) of Reuter"
' 2 = "Phi(pz) of Love/Scott"
' 3 = "Phi(pz) of Riveros"
' 4 = "Phi(pz) of Pouchou & Pichoir"
' 5 = "Phi(pz) of Karduck & Rehbach"
' 6 = "Phi(pz) of August and Wernisch"
' 7 = "Phi(pz) of Packwood"

ierror = False
On Error GoTo ZAFPhiError

Dim zm As Single, zr As Single, zra As Single, zrb As Single
Dim a1 As Single, a2 As Single, a3 As Single, ax As Single
Dim jpu As Single, gpu As Single

' RPHI1 REUTER PHI(0) EXPRESSION
If iphi% = 1 Then
rr! = 1# + 2.8 * (1# - 0.9 / uu!) * hh!

' RPHI2 LOVE/SCOTT PHI(0) EXPRESSION
ElseIf iphi% = 2 Then
jpu! = 3.43378 - 10.7872 / uu! + 10.97628 / (uu! * uu!) - 3.62286 / (uu! * uu! * uu!)
gpu! = -0.59299 + 21.55329 / uu! - 30.55248 / (uu! * uu!) + 9.59218 / (uu! * uu! * uu!)
If 1# + hh! < 0# Then GoTo ZAFPhiNegative
rr! = 1# + hh! / (1# + hh!) * (jpu! + gpu! * Log(1# + hh!))

' RPHI3 RIVEROS PHI(0) EXPRESSION
ElseIf iphi% = 3 Then
rr! = 1# + hh! * uu! * Log(uu!) / (uu! - 1#)

' RPHI4 POUCHOU and PICHOIR PHI(0) EXPRESSION
ElseIf iphi% = 4 Then
If hh! < 0# Then GoTo ZAFPhiNegative
rr! = 1# + 3.3 * (1# - Exp((2.3 * hh! - 2#) * Log(uu!))) * Exp(1.2 * Log(hh!))

' RPHI5 KARDUCK and REHBACH PHI(0) EXPRESSION
ElseIf iphi% = 5 Then
jpu! = (1# + 0.005 * zz! / v1!) * (0.68 + 3.7 / zz!)
gpu! = -0.01 + 0.04805 * zz! - 0.00051599 * zz! * zz! + 0.0000020802 * zz! * zz! * zz!
gpu! = (1# + 0.05 / v1!) * gpu!
rr! = 1# + (1# - 1# / Sqr(uu!)) ^ jpu! * gpu!

' RPHI6 AUGUST and WERNISCH PHI(0) EXPRESSION
ElseIf iphi% = 6 Then
If zaf.il%(ii%) = 1 Or zaf.il%(ii%) = 2 Then    ' K lines
  jpu! = 0.86 + 0.12 * Exp(-(zaf.Z%(ii%) / 5#) ^ 2)
ElseIf zaf.il%(ii%) = 3 Or zaf.il%(ii%) = 4 Then  ' L lines
  jpu! = 0.82
Else
  jpu! = 0.78
End If

zm! = 0.52 + 1.28 * hh! - 0.72 * hh! ^ 2
a1! = 67.2945 + 279.67 * zm! - 383.52 * zm! ^ 2 + 179.276 * zm! ^ 3
a2! = 1# / (1.3956 - 3.7819 * zm! + 4.5441 * zm! ^ 2 - 2.0704 * zm! ^ 3)
a3! = 24.3 - 78.4 * zm! + 127.44 * zm! ^ 2 - 68.733 * zm! ^ 3
ax! = hh! / (0.5 + 8.34 / (a1! + 1) - a3! / (a2! + 1))
zr! = (jpu! - a1!) - 1#
zra! = (-zr! * Log(uu!) - 1 + uu! ^ zr!) / zr! ^ 2
zr! = (jpu! - a2!) - 1#
zrb! = (-zr! * Log(uu!) - 1 + uu! ^ zr!) / zr! ^ 2
gpu! = ax! * (1# / uu!) ^ jpu! * (zrb! + 8.34 * zra! - a3! * zra!)
rr! = 1# + 2# * (gpu! * uu! ^ jpu!) / Log(uu!)

' RPHI7 PACKWOOD PHI(0) EXPRESSION
ElseIf iphi% = 7 Then
rr! = 1# + 0.75 * 3.14159 * hh! * (1# - Exp((1# - uu!) / 2))
End If

Exit Sub

' Errors
ZAFPhiError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFPhi"
ierror = True
Exit Sub

ZAFPhiNegative:
msg$ = vbCrLf & "Bad (negative) hh parameter calculated for the sample analysis. This usually indicates negative concentrations so you should check that you are not analyzing epoxy." & vbCrLf
msg$ = msg$ & "You should also make sure your off-peak background and interference corrections are not overcorrecting, or perhaps you have assigned a blank correction to a major or minor element and you did not enter the correct blank level in the Standard Assignments dialog."
If Not CalcImageQuantFlag Then
MiscMsgBoxTim FormMSGBOXTIME, "ZAFPhi", msg$, 20#
Call IOWriteLog(msg$)
Else
Call IOWriteLog(msg$)
End If
'ierror = True
Exit Sub

End Sub

Sub ZAFPap(mode As Integer, ii As Integer)
' PAP absorption correction
' mode = 1 full PAP
' mode = 2 simplified PAP

ierror = False
On Error GoTo ZAFPapError

Dim i1 As Integer

Dim g1 As Double, g2 As Double, g3 As Double
Dim r0 As Double, r00 As Double, rr0 As Double
Dim u0 As Double, eC As Double, eO As Double
Dim b As Double, emm As Double, qeO As Double
Dim dee As Double, q0 As Double, qq As Double
Dim rx As Double, rm As Double, phi0 As Double

Dim delta As Double, rc As Double, zip As Double
Dim a1 As Double, a2 As Double, b1 As Double
Dim ff As Double, gamma As Double, exx As Double
Dim rbar As Double, gg As Double, pee As Double
Dim a10 As Double, b10 As Double, etas As Double
Dim a11 As Double, b11 As Double

Dim fp1 As Double, fp2 As Double, fp3 As Double
Dim fff As Double

' Load voltage, overvoltage and critical excitation locals
eO# = CDbl(zaf.eO!(ii%))
u0# = CDbl(zaf.v!(ii%))       ' used to be zaf.y!(ii%)
eC# = CDbl(zaf.eC!(ii%))
b# = CDbl(zaf.atwts!(ii%))    ' used to be zaf.b!(ii%)
emm# = CDbl(em!(ii%))
zip# = CDbl(zipi!(ii%))

' Full PAP
If mode% = 1 Then
If zz! < 0# Then GoTo ZAFPapBadZZ
g1# = 0.11 + 0.41 * Exp(-Exp(0.75 * Log(zz! / 12.75)))
g2# = 1 - Exp(-Exp(0.35 * Log(u0# - 1)) / 1.19)
g3# = 1 - Exp(-(u0# - 0.5) * (Exp(0.4 * Log(zz!)) / 4))
dp!(2) = 0.0000112 * (1.35 - 0.45 * jbar! * jbar!)
dp!(3) = 0.0000022 / jbar!
pp!(3) = -(0.5 - 0.25 * jbar!)
r0# = 0#

For i1% = 1 To 3
r00# = Exp((1 - pp!(i1%)) * Log(jbar!)) * dp!(i1%)
rr0# = Exp((1 + pp!(i1%)) * Log(eC#))
r0# = r0# + r00# * (Exp((1# + pp!(i1%)) * Log(eO#)) - rr0#) / (1 + pp!(i1%)) / sumatom!     ' ionization range
Next i1%

dee# = 1# + Exp(-Exp(0.45 * Log(zz!)) * Log(u0#))
q0# = 1# - 0.535 * Exp(-Exp(1.2 * Log(21# / zn!))) - 0.00025 * (Exp(3.5 * Log(zn! / 20#)))
qq# = q0# + (1# - q0#) * Exp(-(u0# - 1#) * zz! / 40#)
rx# = qq# * dee# * r0#              ' depth range of ionization
rm# = g1# * g2# * g3# * rx#         ' depth of maximum phi
qeO# = Log(u0#) / (eC# * eC# * Exp(emm# * Log(u0#)))
xp! = xp! / (zip# * 66892#) * b#
ff# = xp! / qeO#

phi0# = 1# + 3.3 * (1# - Exp((2.3 * hh! - 2#) * Log(u0#))) * Exp(1.2 * Log(hh!))   ' phi(0)
delta# = (rx# - rm#) * (ff# - phi0# * rx# / 3#) * ((rx# - rm#) * ff# - phi0# * rx# * (rm# + rx# / 3#))
rc# = 1.5 * ((ff# - phi0# * rx# / 3#) / phi0# - Sqr(Abs(delta#)) / phi0# / (rx# - rm#))
a1# = phi0# / rm# / (rc# - rx# * (rc# / rm# - 1#))
b1# = phi0# - a1# * rm# * rm#
a2# = a1# * (rc# - rm#) / (rc# - rx#)

fp1# = ((rc# - rm#) * (rc# - rm# + 2 / xi!) + 2# / xi! / xi!) * Exp(-xi! * rc#)
fp1# = -a1# / xi! * (fp1# - rm# * (rm# - 2# / xi!) - 2# / xi! / xi!)
fp2# = -b1# / xi! * (Exp(-xi! * rc#) - 1#)
fp3# = ((rc# - rx#) * (rc# - rx# + 2# / xi!) + 2# / xi! / xi!) * Exp(-xi! * rc#)
fp3# = -a2# / xi! * (-fp3# + 2# / xi! / xi! * Exp(-xi! * rx#))
fff# = a1# / 3# * ((rc# - rm#) * (rc# - rm#) * (rc# - rm#) + rm# * rm# * rm#) + b1# * rc#
fff# = fff# + a2# / 3# * (rx# - rc#) * (rx# - rc#) * (rx# - rc#)
FP! = (fp1# + fp2# + fp3#) / fff#

' Code commented out by JTA
'FP4=(A1/XI)*((RC-RM)*(RX-RC-2/XI)-2/XI/XI)*EXP(-XI*RC)
'FP4=FP4-(RC-RM)*RX+RM*(RC-2/XI)+2/XI/XI
'FP5=(A2/XI)*((RX-RC)*(RX-RC-2/XI)-2/XI/XI)*EXP(-XI*RC)
'FP5=FP5-(2/XI/XI)*EXP(-XI*RX)
'FFFF=(FP4+FP5)/FFF
'Call ZAFPap2   ' numerical integration of PAP phi(pz)
'If ierror Then Exit Sub
End If

' Simplified PAP
If mode% = 2 Then
qeO# = Log(u0#) / (eC# * eC# * Exp(emm# * Log(u0#)))
xp! = xp! / (zip# * 66892#) * b#
ff# = xp! / qeO#

phi0# = 1# + 3.3 * (1# - Exp((2.3 * hh! - 2#) * Log(u0#))) * Exp(1.2 * Log(hh!))    ' phi(0)
gamma# = 0.2 + meanz! / 200#
exx = 1# + 1.3 * Log(meanz!)
rbar# = ff# / (1 + (exx# * Log(1# + gamma# * (1# - Exp(-0.42 * Log(u0#))))) / Log(1# + gamma#))

If ff# / rbar# < phi0# Then rbar# = ff# / phi0#     ' average ionization depth
gg# = 0.22 * Log(4# * meanz!) * (1# - 2# * Exp(-meanz! * (u0# - 1#) / 15#))
hh! = 1# - 10# * (1# - 1# / (1# + u0# / 10#)) / (meanz! * meanz!)
pee# = gg# * hh! * hh! * hh! * hh! * ff# / (rbar# * rbar#)

b10# = Sqr(2#) * (1# + Sqr(1# - rbar# * phi0# / ff#)) / rbar#
a10# = (pee# + b10# * (2# * phi0# - b10# * ff#)) / (b10# * ff# * (2# - b10# * rbar#) - phi0#)
etas# = (a10# - b10#) / b10#
If etas# < 0.000001 Then a10# = b10# * (1# + etas#)

b11# = (b10# * b10# * ff# * (1# + etas#) - pee# - phi0# * b10# * (2# + etas#)) / etas#
a11# = (b11# / b10# + phi0# - b10# * ff#) * (1# + etas#) / etas#
FP! = (phi0# + b11# / (b10# + xi!) - a11# * b10# * etas# / (b10# * (1# + etas#) + xi!)) / (b10# + xi!)
FP! = FP! / ff#
End If

Exit Sub

' Errors
ZAFPapError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFPap"
ierror = True
Exit Sub

ZAFPapBadZZ:
msg$ = "Bad (negative) zz parameter calculated for the sample analysis. This usually indicates negative concentrations so you should check that you are not analyzing epoxy." & vbCrLf & vbCrLf
msg$ = msg$ & "You should also make sure your off-peak background and interference corrections are not overcorrecting, or perhaps you have assigned a blank correction to a major or minor element and you did not enter the correct blank level in the Standard Assignments dialog."
If Not CalcImageQuantFlag Then
MiscMsgBoxTim FormMSGBOXTIME, "ZAFPap", msg$, 20#
Call IOWriteLog(msg$)
Else
Call IOWriteLog(msg$)
End If
'ierror = True
Exit Sub

End Sub

Function ZAFPap2(rc As Single, rm As Single, rx As Single, a1 As Single, a2 As Single, b1 As Single, xi As Single) As Single
' PAP numerical integration

ierror = False
On Error GoTo ZAFPap2Error

Dim i1 As Integer, nrsteps As Integer

Dim xint1 As Single, xint2 As Single, xint3 As Single, xint4 As Single
Dim step1 As Single, step2 As Single, depth As Single
Dim phi1 As Single, phi2 As Single

ZAFPap2! = 0#

nrsteps% = 40
xint1! = 0#
xint2! = 0#
xint3! = 0#
xint4! = 0#

step1! = rc! / nrsteps%
step2! = (rx! - rc!) / nrsteps%
For i1% = 1 To nrsteps%
depth! = step1! * (i1% - 0.5)
phi1! = a1! * (depth! - rm!) * (depth! - rm!) + b1!
xint1! = xint1! + phi1! * step1!
xint2! = xint2! + phi1! * Exp(-xi! * depth!) * step1!
Next i1%

For i1% = 1 To nrsteps%
depth! = rc! + step2! * (i1% - 0.5)
phi2! = a2! * (depth! - rx!) * (depth! - rx!)
xint3! = xint3! + phi2! * step2!
xint4! = xint4! + phi2! * Exp(-xi! * depth!) * step2!
Next i1%

ZAFPap2! = (xint2! + xint4!) / (xint1! + xint3!)
Exit Function

' Errors
ZAFPap2Error:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFPap2"
ierror = True
Exit Function

End Function

Sub ZAFPtc(ii As Integer, aa As Single, v0 As Single, zz As Single, er1 As Single, er2 As Single, er3 As Single, a1 As Single, a2 As Single, xx As Single, X2 As Single, x3 As Single, x4 As Single, x5 As Single)
' Subroutine for calculation of numerical integration of thin film and
' particle path lengths.

ierror = False
On Error GoTo ZAFPtcError

Dim r4 As Single, r5 As Single, r6 As Single
Dim j1 As Single, q1 As Single, q2 As Single
Dim d1 As Single, d2 As Single, d4 As Single
Dim p4 As Single, p5 As Single, p6 As Single
Dim b9 As Single, s1 As Single, r9 As Single

' Calculate electron range
zaf.erange!(ii%) = 0.00000276 * v0! ^ 1.67 * aa! / zz! ^ (8# / 9#)
er3! = zaf.erange!(ii%)
er1! = 0#
er2! = 0#

j1! = 0#
a1! = 0#
a2! = 0#
q1! = 0#
q2! = 0#
d4! = xx! * zaf.t3!

' BEGIN NUMERICAL INTEGRATION OF PHI(PZ) EXPRESSION
zaf.intnum&(ii%) = 0
11150:
zaf.intnum&(ii%) = zaf.intnum&(ii%) + 1
If zaf.intnum&(ii%) > 100000 Then GoTo 11200  ' check number of integrations

j1! = j1! + 1#
d1! = er1! * zaf.t4!
d2! = zaf.d! - d1!
p4! = xx! * zaf.t1! * er1!
p5! = Exp(-p4!)
p6! = Exp(-xx! * zaf.g! * er1!)
r4! = er1! * 10000!
b9! = X2! * (1# - x5! * Exp(-x3! * er1!)) * Exp(-x4! * x4! * er1! * er1!)
If b9! < 0.000001 Then
zaf.erange!(ii%) = er1!
Exit Sub
End If

s1! = 0#
If er1! <= zaf.j9! * zaf.d! Then
Select Case zaf.model%
  Case 1              ' Thin Film Model
      s1! = b9! * p5!
             
  Case 2              ' Rectangular Prism Model
    If er1! = 0 Then s1! = b9!
    If (er1! <> 0) And (d2! > 0#) Then s1! = 1# / zaf.d! * (d2! * p5! + (1# - p5!) / d4!) * b9!
    If (er1! <> 0#) And (d2! <= 0#) Then s1! = 1# / zaf.d! * (1# - Exp(-d4! * zaf.d!)) / d4! * b9!
             
  Case 3              ' Tetragonal Prism Model
    If er1! = 0# Then s1! = b9!
    If (er1! <> 0#) And (d2! > 0#) Then
      s1! = (p5! - 1#) / (2# * d4! * d4!) + d1! / (2# * d4!)
      s1! = s1! + (d2! / 2#) ^ 2 * p5! + d2! / (2# * d4!) * (1# - p5!)
      s1! = 4# / (zaf.d! * zaf.d!) * s1! * b9!
    End If
    If (er1! <> 0#) And (d2! <= 0#) Then
      s1! = (Exp(-d4! * zaf.d!) - 1#) / (2# * d4! * d4!) + zaf.d! / (2# * d4!)
      s1! = 4# / (zaf.d! * zaf.d!) * s1! * b9!
    End If
             
  Case 4              ' Triangular Prism Model
      r9! = zaf.d! / 2#
    If er1! > r9! Then
      s1! = 0#
    Else
      s1! = (p6! - Exp(-xx! * zaf.g! * (zaf.d! - er1!))) / (2# * xx! * zaf.g!)
      s1! = (s1! + (r9! - er1!) * p6!) / zaf.d! * b9!
    End If
             
  Case 5              ' Square Pyramid Model
      r9! = zaf.d! / 2#
    If er1! > r9! Then
      s1! = 0#
    Else
      s1! = (zaf.d! - 2# * er1!) * p6! / (xx! * zaf.g!)
      s1! = s1! + (Exp(-xx! * zaf.g! * zaf.d!) / p6! - p6!) / (xx! * xx! * zaf.g! * zaf.g!)
      s1! = s1! + (zaf.d! * zaf.d! - 4# * zaf.d! * er1! + 8# * er1! * er1!) * p6! / 2#
      s1! = s1! / (zaf.d! * zaf.d!) * b9!
    End If
             
  Case 6              ' Sidescatter-Mod. Rec. Prism Model
    If er1! < zaf.d! * zaf.t2! / 2# Then
      s1! = (0.75 * zaf.d! - d1! / 2# - d1! * d1! / (2# * zaf.d!) - d1! / (d4! * zaf.d!)) * p5!
      s1! = s1! + (1# - p5!) / (2# * d4!) + (1# - p5!) / (d4! * d4! * zaf.d!)
    End If
    If (er1! >= zaf.d! * zaf.t2! / 2#) And (er1! < zaf.d! * zaf.t2!) Then
      s1! = (zaf.d! - 3# * d1! / 2# + d1! * d1! / (2# * zaf.d!) + d1! / (d4! * zaf.d!) - 1# / d4!) * p5!
      s1! = s1! + (1# - p5!) / (2# * d4!) + (1# + p5!) / (d4! * d4! * zaf.d!)
      s1! = s1! - 2# * Exp(-d4! * zaf.d! / 2#) / (d4! * d4! * zaf.d!)
    End If
      If (er1! >= zaf.d! * zaf.t2!) Then
      s1! = (1# - Exp(-d4! * zaf.d!)) / (2# * d4!) + (1# + Exp(-d4! * zaf.d!)) / (d4! * d4! * zaf.d!)
      s1! = s1! - 2# * Exp(-d4! * zaf.d! / 2#) / (d4! * d4! * zaf.d!)
    End If
    s1! = s1! * b9! / zaf.d!
    
  Case Else
    MsgBox "Undefined ZAFPtc model", vbOKOnly + vbExclamation, "ZAFPtc"
    ierror = True
    Exit Sub
End Select
End If ' end of check for "If er1! <= zaf.j9! * zaf.d! Then"

11200:
If j1! > 1 Then
r6! = r4! - r5!
a1! = a1! + (q1! + b9!) / 2# * r6!
a2! = a2! + (q2! + s1!) / 2# * r6!
End If

q1! = b9!
q2! = s1!
r5! = r4!
If er1! >= er3! Then GoTo 11250

er1! = er1! + zaf.x1!
If er1! > er3! Then er1! = er3!
GoTo 11150

11250:
Exit Sub

' Errors
ZAFPtcError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFPtc"
ierror = True
Exit Sub

End Sub

Sub ZAFPrintStandards(analysis As TypeAnalysis, sample() As TypeSample)
' Print the standard parameters (after the last standard is calculated by ZAFStd)

ierror = False
On Error GoTo ZAFPrintStandardsError

' Print out results for assigned standards
Call ZAFPrintStandards2(zaf, analysis, sample())
If ierror Then Exit Sub

Exit Sub

' Errors
ZAFPrintStandardsError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFPrintStandards"
ierror = True
Exit Sub

End Sub
