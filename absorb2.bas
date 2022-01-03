Attribute VB_Name = "CodeABSORB2"
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

Dim allmacs(1 To MAXELM%, 1 To MAXELM%, 1 To MAXRAY% - 1) As Single   ' emitter by absorber by x-ray

Sub AbsorbGetMAC30(energy As Single, iz As Integer, ielm As Integer, iray As Integer, lines() As Double, edges() As Double, atotal As Single)
' This routine returns the Heinrich fit MAC from the Goldstein table.
' Code from Paul Carpenter (C_MAC.C) and modified John Donovan for VB (1998).
' energy is the energy in keV and is used if no line is specified
' iz is the absorber atomic number
' ielm is the emitter atomic number
' iray is the emitter xray line (1=ka, 2=kb, 3=la, 4=lb, 5=ma, 6=mb)
'
' Routine to compute the mass absorption coefficients using Heinrich's latest parameterization.
' Calculated from equation:
'
'   u = C Z^4 / A (12397/E)^n (1 - exp[(-E+b)/a))     eq 3 from Heinrich paper
'
'   u = mass absorption coefficient
'   C = parameter a function of Z
'   Z = atomic number of element (line)
'   A = atomic weight
'   E = photon energy in ev
'   a & b are parameters which vary with Z of absorber and
'   region between absorption edges
'
' Second equation (eq 4 from Heinrich paper)
'
'   u = 1.02 C (12397 / line)^n C Z^4 / A [(E - cutoff)/(En - cutoff))
'   cutoff = threshold energy

ierror = False
On Error GoTo AbsorbGetMAC30Error

Dim ix As Integer
Dim a As Double, b As Double, c As Double
Dim e As Double, Z As Double
Dim CC As Double, n As Double
Dim c1 As Double, c2 As Double, c3 As Double
Dim c4 As Double, c5 As Double

Dim ata As Double, cutoff As Double
Dim mu As Double
Dim tmsg As String

ReDim edge(1 To 12) As Double

atotal! = 0#
If iz% > 99 Or ielm% > 99 Then Exit Sub ' mac30 tables only go to element Es

' Convert to LINES2.DAT format
If iray% = 1 Then ix% = 1 ' Ka
If iray% = 2 Then ix% = 2 ' Kb
If iray% = 3 Then ix% = 3 ' La
If iray% = 4 Then ix% = 4 ' Lb
If iray% = 5 Then ix% = 10 ' Ma
If iray% = 6 Then ix% = 11 ' Mb

' Check for either energy or line
If energy! = 0# And (ix% = 0 Or ielm% = 0) Then GoTo AbsorbGetMAC30NoEnergyOrLine

' Check for energy and convert to eV
If ix% = 0 And ielm% = 0 Then
e# = CDbl(energy! * EVPERKEV#) ' in eV
If e# = 0# Then
tmsg$ = "Warning in AbsorbGetMAC30- no emitter energy specified"
Call IOWriteLogRichText(tmsg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
Exit Sub
End If
End If

' Check for emission line
If ix% > 0 And ielm% > 0 Then
e# = CDbl(lines#(ix%, ielm%)) ' load emitter line enery in eV
If e# = 0# Then
tmsg$ = "Warning in AbsorbGetMAC30- zero energy for emitting line specified (" & Symup$(ielm%) & " " & Xraylo$(iray%) & " in " & Symup$(iz%) & ")"
Call IOWriteLogRichText(tmsg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
Exit Sub
End If
End If

' Load absorber atomic number and atomic weight
Z# = iz%
ata# = AllAtomicWts!(iz%)

' Calculate region based on edge type of absorber
edge#(1) = edges#(1, iz%)    ' K edge
edge#(2) = edges#(2, iz%)    ' L1 edge
edge#(3) = edges#(3, iz%)    ' L2 edge
edge#(4) = edges#(4, iz%)    ' L3 edge
edge#(5) = edges#(5, iz%)    ' M1 edge
edge#(6) = edges#(6, iz%)    ' M2 edge
edge#(7) = edges#(7, iz%)    ' M3 edge
edge#(8) = edges#(8, iz%)    ' M4 edge
edge#(9) = edges#(9, iz%)    ' M5 edge
edge#(10) = edges#(10, iz%)    ' N1 edge
edge#(11) = edges#(11, iz%)    ' N2 edge
edge#(12) = edges#(12, iz%)    ' N3 edge

' Region 1: K
If (e# > edge(1)) And (edge(1) > 0#) Then
    If Z < 6# Then   ' region1: Ka   Z 1-5
        a = 24.4545 + 155.6055 * Z - 14.15422 * AbsorbPow#(Z, 2#)
        b = -103# + 18.2 * Z
        CC = -0.000287536 + 0.001808599 * Z ' jjd- was 1.808599
        n = 3.34745 + 0.02652873 * Z - 0.01273815 * AbsorbPow#(Z, 2#)
    Else    ' region1: Ka   Z > 5
        a = 47# * Z + 6.52 * AbsorbPow#(Z, 2#) - 0.152624 * AbsorbPow#(Z, 3#)
        b = 0#
        CC = 0.005253 + 0.00133257 * Z - 0.000075937 * AbsorbPow#(Z, 2#) + 0.00000169357 * AbsorbPow#(Z, 3#) - 0.000000013975 * AbsorbPow#(Z, 4#)
        n = 3.112 - 0.0121 * Z
    End If
End If

' Region 2: K - L1 = La1
If (e# > edge(2)) And (e# <= edge(1)) Then
    a = 17.8096 * Z + 0.067429 * AbsorbPow#(Z, 2#) + 0.01253775 * AbsorbPow#(Z, 3#) - 0.000116286 * AbsorbPow#(Z, 4#)
    b = 0#
    c = -0.0000924 + 0.000141478 * Z - 0.00000524999 * AbsorbPow#(Z, 2#) + 0.0000000985296 * AbsorbPow#(Z, 3#) - 0.000000000907306 * AbsorbPow#(Z, 4#) + 3.19245E-12 * AbsorbPow#(Z, 5#)
    CC = c
    n = 2.7575 + 0.001889 * Z - 0.00004982 * AbsorbPow#(Z, 2#)
End If

' Region 3: L1 - L2 = Lb1
If (e# > edge(3)) And (e# <= edge(2)) Then
    a = 17.8096 * Z + 0.067429 * AbsorbPow#(Z, 2#) + 0.01253775 * AbsorbPow#(Z, 3#) - 0.000116286 * AbsorbPow#(Z, 4#)
    b = 0#
    c = -0.0000924 + 0.000141478 * Z - 0.00000524999 * AbsorbPow#(Z, 2#) + 0.0000000985296 * AbsorbPow#(Z, 3#) - 0.000000000907306 * AbsorbPow#(Z, 4#) + 3.19245E-12 * AbsorbPow#(Z, 5#)
    n = 2.7575 + 0.001889 * Z - 0.00004982 * AbsorbPow#(Z, 2#)
    CC = c * 0.858
End If

' Region 4: L2 - L3
If (e# > edge(4)) And (e# <= edge(3)) Then
    a = 17.8096 * Z + 0.067429 * AbsorbPow#(Z, 2#) + 0.01253775 * AbsorbPow#(Z, 3#) - 0.000116286 * AbsorbPow#(Z, 4#)
    b = 0#
    c = -0.0000924 + 0.000141478 * Z - 0.00000524999 * AbsorbPow#(Z, 2#) + 0.0000000985296 * AbsorbPow#(Z, 3#) - 0.000000000907306 * AbsorbPow#(Z, 4#) + 3.19245E-12 * AbsorbPow#(Z, 5#)
    n = 2.7575 + 0.001889 * Z - 0.00004982 * AbsorbPow#(Z, 2#)
    CC = c * (0.8933 - 0.00829 * Z + 0.0000638 * AbsorbPow#(Z, 2#))
End If

' Region 5: L3 - M1 = M1
If (e# > edge(5)) And (e# <= edge(4)) Then
    ' pkc there was an error here:  2.63199611e-2 was written as 2.63199611
    a = 10.2575657 * Z - 0.822863477 * AbsorbPow#(Z, 2#) + 0.0263199611 * AbsorbPow#(Z, 3#) - 0.00018641019 * AbsorbPow#(Z, 4#)
    If Z < 61# Then
        b = 5.654 * Z - 0.536839169 * AbsorbPow#(Z, 2#) + 0.018972278 * AbsorbPow#(Z, 3#) - 0.0001683474 * AbsorbPow#(Z, 4#)
    Else
        b = -1232.4022 * Z + 51.114164 * AbsorbPow#(Z, 2#) - 0.699473097 * AbsorbPow#(Z, 3#) + 0.0031779619 * AbsorbPow#(Z, 4#)
    End If
            
    If Z < 30# Then
        c = 0.01889757 - 0.0018517159 * Z + 0.000069602789 * AbsorbPow#(Z, 2#) - 0.0000011641145 * AbsorbPow#(Z, 3#) + 7.2773258E-09 * AbsorbPow#(Z, 4#)
    Else
        c = 0.0030039 - 0.000173663566 * Z + 0.0000040424792 * AbsorbPow#(Z, 2#) - 0.000000040585911 * AbsorbPow#(Z, 3#) + 1.497763E-10 * AbsorbPow#(Z, 4#)
    End If
    
    CC = c
    n = 0.5385 + 0.084597 * Z - 0.00108246 * AbsorbPow#(Z, 2#) + 0.0000044509 * AbsorbPow#(Z, 3#)
End If

' Region 6: M1 - M2 = Ma
If (e# > edge(6)) And (e# <= edge(5)) Then
    a = 4.62 * Z - 0.04 * AbsorbPow#(Z, 2#)
    b = (2.51 - 0.052 * Z + 0.000378 * AbsorbPow#(Z, 2#)) * edge(8)
    c1 = 0.000077708 - 0.00000783544 * Z + 0.0000002209365 * AbsorbPow#(Z, 2#) - 0.00000000129086 * AbsorbPow#(Z, 3#)
    c2 = 1.406 + 0.0162 * Z - 0.0006561 * AbsorbPow#(Z, 2#) + 0.000004865 * AbsorbPow#(Z, 3#)
    c3 = 0.584 + 0.01955 * Z - 0.0001285 * AbsorbPow#(Z, 2#)
    CC = c1 * c2 * c3
    n = 3# - 0.004 * Z
End If

' Region 7: M2 - M3 = M2
If (e# > edge(7)) And (e# <= edge(6)) Then
    a = 4.62 * Z - 0.04 * AbsorbPow#(Z, 2#)
    b = (2.51 - 0.052 * Z + 0.000378 * AbsorbPow#(Z, 2#)) * edge(8)
    c1 = 0.000077708 - 0.00000783544 * Z + 0.0000002209365 * AbsorbPow#(Z, 2#) - 0.00000000129086 * AbsorbPow#(Z, 3#)
    c2 = 1.406 + 0.0162 * Z - 0.0006561 * AbsorbPow#(Z, 2#) + 0.000004865 * AbsorbPow#(Z, 3#)
    ' pkc there was an error here: 1.366e-4 instead of 1.366e-3
    c4 = 1.082 + 0.001366 * Z
    CC = c1 * c2 * c4
    n = 3# - 0.004 * Z
End If

' Region 8: M3 - M4 = Mg
If (e# > edge(8)) And (e# <= edge(7)) Then
    a = 4.62 * Z - 0.04 * AbsorbPow#(Z, 2#)
    b = (2.51 - 0.052 * Z + 0.000378 * AbsorbPow#(Z, 2#)) * edge(8)
    n = 3# - 0.004 * Z
    c1 = 0.000077708 - 0.00000783544 * Z + 0.0000002209365 * AbsorbPow#(Z, 2#) - 0.00000000129086 * AbsorbPow#(Z, 3#)
    c2 = 1.406 + 0.0162 * Z - 0.0006561 * AbsorbPow#(Z, 2#) + 0.000004865 * AbsorbPow#(Z, 3#)
    CC = 0.95 * c1 * c2
End If

'  Region 9: M4 - M5 = Mb
If (e# > edge(9)) And (e# <= edge(8)) Then
    a = 4.62 * Z - 0.04 * AbsorbPow#(Z, 2#)
    b = (2.51 - 0.052 * Z + 0.000378 * AbsorbPow#(Z, 2#)) * edge(8)
    n = 3# - 0.004 * Z
    ' if e < edge(8) then cc = 0.8 * c * ((.0005083 * z - 0.06) * z + 2.0553)   ' pkc what is this???
    c1 = 0.000077708 - 0.00000783544 * Z + 0.0000002209365 * AbsorbPow#(Z, 2#) - 0.00000000129086 * AbsorbPow#(Z, 3#)
    c2 = 1.406 + 0.0162 * Z - 0.0006561 * AbsorbPow#(Z, 2#) + 0.000004865 * AbsorbPow#(Z, 3#)
    c5 = 1.6442 - 0.048 * Z + 0.00040664 * AbsorbPow#(Z, 2#)
    CC = c1 * c2 * c5
End If

' Region 10: M5 - N1 = N1
If (e# > edge(10)) And (e# <= edge(9)) Then
    a = 19.64 * Z - 0.61239 * AbsorbPow#(Z, 2#) + 0.00539309 * AbsorbPow#(Z, 3#)
    b = -113# + 4.5 * Z
    CC = 1.08 * (0.0043156 - 0.00014653 * Z + 0.000001707073 * AbsorbPow#(Z, 2#) - 0.00000000669827 * AbsorbPow#(Z, 3#))
    n = 0.3736 + 0.02401 * Z
End If

' Region 11: below N1
' Correction in equation number 4 in Heinrich 1986 paper
If (e# <= edge(10)) Then
    a = 19.64 * Z - 0.61239 * AbsorbPow#(Z, 2#) + 0.00539309 * AbsorbPow#(Z, 3#)
    b = -113# + 4.5 * Z
    n = 0.3736 + 0.02401 * Z
    CC = 1.08 * (0.0043156 - 0.00014653 * Z + 0.000001707073 * AbsorbPow#(Z, 2#) - 0.00000000669827 * AbsorbPow#(Z, 3#))
    cutoff# = (0.252 * Z - 31.1812) * Z + 1042#
    
    mu = 1.02 * AbsorbPow#(CC, 1#) * AbsorbPow#((12397# / e#), n) * AbsorbPow#(Z, 4#) / ata * (1# - Exp((-edge(10) + b) / a)) * ((e# - cutoff#) / (edge(10) - cutoff#))
    
End If

' Calculate all oher values
If (e# > edge(10)) Then
    mu = CC * AbsorbPow#(Z, 4#) / ata * AbsorbPow#((12397# / e#), n) * (1# - Exp((-e# + b) / a))
End If

atotal! = CSng(mu#)

Exit Sub

' Errors
AbsorbGetMAC30Error:
MsgBox Error$, vbOKOnly + vbCritical, "AbsorbGetMAC30"
ierror = True
Exit Sub

AbsorbGetMAC30NoEnergyOrLine:
msg$ = "No emitting energy or line specified"
MsgBox msg$, vbOKOnly + vbExclamation, "AbsorbGetMAC30"
ierror = True
Exit Sub

End Sub

Sub AbsorbGetMACJTA(energy As Single, iz As Integer, ielm As Integer, iray As Integer, g() As Single, o() As Single, atotal As Single)
' This routine returns the MAC value calculated from Armstrong's MACCALC.BAS program.
' Modified by John Donovan for VB (1998).
' energy is the energy in keV and is used if greater than zero
' iz is the absorber atomic number
' ielm is the emitter atomic number
' iray is the emitter xray line (1=ka, 2=kb, 3=la, 4=lb, 5=ma, 6=mb)
' g() is the emission line energies from LINES.DAT from Armstrong
' o() is the absorption edge energies from LINES.DAT from Armstrong

ierror = False
On Error GoTo AbsorbGetMACJTAError

Dim i3 As Integer, i6 As Integer, ix As Integer
Dim m2 As Single, m7 As Single, m8 As Single

ReDim edge(1 To 10) As Single

ReDim d(4, 4) As Single
ReDim f(4, 10) As Single

f!(1, 1) = -0.0397931
f!(2, 1) = 2.423
f!(3, 1) = 5.5091
f!(1, 2) = -0.033916
f!(2, 2) = 2.82526
f!(3, 2) = 9.03526
f!(1, 3) = -0.0865397
f!(2, 3) = 3.32315
f!(3, 3) = 10.2505
f!(1, 4) = -0.228343
f!(2, 4) = 4.31172
f!(3, 4) = 12.0025
f!(1, 5) = 1.25179
f!(2, 5) = -7.838
f!(3, 5) = -11.5803
f!(1, 6) = 0.834903
f!(2, 6) = -4.14925
f!(3, 6) = -3.33802
f!(1, 7) = 0.442217
f!(2, 7) = -0.979241
f!(3, 7) = 3.15348
f!(1, 8) = 0.25141
f!(2, 8) = 0.931913
f!(3, 8) = 8.03561
f!(1, 9) = 0.272951
f!(2, 9) = 0.688906
f!(3, 9) = 7.4243

d!(1, 1) = -0.232229
d!(2, 1) = 4.070005
d!(3, 1) = -6.22075
d!(1, 2) = -0.254471
d!(2, 2) = 4.76924
d!(3, 2) = -10.3788
d!(1, 3) = 0.256216
d!(2, 3) = 1.15119
d!(3, 3) = -5.68485
d!(1, 4) = 1.35916
d!(2, 4) = -9.49212
d!(3, 4) = 18.6408

d!(4, 3) = 2.6
d!(4, 4) = 2.22

f!(4, 1) = 1
f!(4, 2) = 1
f!(4, 3) = 1.17
f!(4, 4) = 1.63
f!(4, 5) = 1
f!(4, 6) = 1.16
f!(4, 7) = 1.4
f!(4, 8) = 1.621
f!(4, 9) = 1.783
f!(4, 10) = 1

atotal! = 0#
If iz% > 95 Or ielm% > 95 Then Exit Sub ' jta tables only go to element Am

' Check that only alpha lines are calculated
If iray% = 2 Or iray% = 4 Or iray% = 6 Then
If DebugMode Then
'msg$ = "WARNING in AbsorbGetMACJTA- cannot calculate beta line mass absorption coefficient"
'Call IOWriteLog(msg$)
End If
Exit Sub
End If

' Check for low energy
If energy! < 0.1 Then
If DebugMode Then
msg$ = "WARNING in AbsorbGetMACJTA- energy " & Format$(Format$(energy!, f83$), a80$) & " is too low."
Call IOWriteLog(msg$)
End If
Exit Sub
End If

' Check for arbitrary energy or not
If energy! = 0# Then

' Convert to Armstrong x-ray line notation (1=ka, 2=la, 3=ma)
If iray% = 1 Then ix% = 1
If iray% = 3 Then ix% = 2
If iray% = 5 Then ix% = 3

' Check for missing emission line energy
If g!(ix%, ielm%) = 0# Then
If DebugMode Then
msg$ = "WARNING in AbsorbGetMACJTA- missing emission energy for " & Symlo$(ielm%) & " " & Xraylo$(iray%)
Call IOWriteLog(msg$)
End If
Exit Sub
End If

energy! = g!(ix%, ielm%)
End If

' Calculate region based on edge type
edge!(1) = o!(1, iz%)    ' K edge
edge!(2) = o!(2, iz%)    ' L1 edge
edge!(3) = o!(3, iz%)    ' L2 edge
edge!(4) = o!(4, iz%)    ' L3 edge
edge!(5) = o!(5, iz%)    ' M1 edge
edge!(6) = o!(6, iz%)    ' M2 edge
edge!(7) = o!(7, iz%)    ' M3 edge
edge!(8) = o!(8, iz%)    ' M4 edge
edge!(9) = o!(9, iz%)    ' M5 edge

' Perform calculation from FRAME equations
m2! = Log(iz%)
d!(4, 1) = Exp(-0.0045522 * m2! * m2! - 0.0068535 * m2! + 1.07018)
d!(4, 2) = 2.73
If iz% >= 42 Then
d!(4, 2) = Exp(-0.113159 * m2! * m2! + 0.836883 * m2! - 0.545969)
End If

For i3% = 1 To 10
If i3% < 10 Then
If energy! < edge!(i3%) Then GoTo 930    ' if emission energy is less than absorption edge energy
End If

i6% = Int(i3% + 0.01) - Int(i3% / 3 + 0.01) - Int(i3% / 4 + 0.01)
i6% = Int(i6% - Int(i3% / 7 + 0.01) + 0.01)
m7! = Exp(d(1, i6%) * m2! * m2! + d!(2, i6%) * m2! + d!(3, i6%)) / f!(4, i3%)
m8! = m7! * (12.398 / energy!) ^ d!(4, i6%)
i3% = 10
930: Next i3%

atotal! = m8!

Exit Sub

' Errors
AbsorbGetMACJTAError:
MsgBox Error$, vbOKOnly + vbCritical, "AbsorbGetMACJTA"
ierror = True
Exit Sub

End Sub

Sub AbsorbLoadLINESDataFile(g() As Single, o() As Single)
' Load the LINES.DAT data file (for self consistant calculations for MACJTA)

ierror = False
On Error GoTo AbsorbLoadLINESDataFileError

Dim i As Integer, iz As Integer
Dim dt As String, TM As String
Dim lab1 As String, lab2 As String, lab3 As String, lab4 As String

If Dir$(ApplicationCommonAppData$ & "LINES.DAT") = vbNullString Then GoTo AbsorbLoadLINESDataFileNoFile
Open ApplicationCommonAppData$ & "LINES.DAT" For Input As #Temp1FileNumber%

Input #Temp1FileNumber%, dt$, TM$
Input #Temp1FileNumber%, lab1$, lab2$

For i% = 1 To 95    ' jta tables only go to Am
Input #Temp1FileNumber%, iz%, g!(1, i%), g!(2, i%), g!(3, i%)
Next i%
Input #Temp1FileNumber%, lab3$, lab4$

For i% = 1 To 95    ' jta tables only go to Am
Input #Temp1FileNumber%, iz%, o!(1, i%), o!(2, i%), o!(3, i%), o!(4, i%), o!(5, i%), o!(6, i%), o!(7, i%), o!(8, i%), o!(9, i%)
Next i%

Close #Temp1FileNumber%

Exit Sub

' Errors
AbsorbLoadLINESDataFileError:
MsgBox Error$, vbOKOnly + vbCritical, "AbsorbLoadLINESDataFile"
Close #Temp1FileNumber%
ierror = True
Exit Sub

AbsorbLoadLINESDataFileNoFile:
msg$ = "File " & ApplicationCommonAppData$ & "LINES.DAT" & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "AbsorbLoadLINESDataFile"
ierror = True
Exit Sub

End Sub

Sub AbsorbLoadLINES2DataFile(lines() As Double, edges() As Double)
' Load the LINES2.DAT data file (for self consistant calculations for MAC30)

ierror = False
On Error GoTo AbsorbLoadLINES2DataFileError

Dim i As Integer, j As Integer
Dim astring As String, bstring As String

' Read line and edge file if first time
If Dir$(ApplicationCommonAppData$ & "LINES2.DAT") = vbNullString Then GoTo AbsorbLoadLINES2DataFileNoFile
Open ApplicationCommonAppData$ & "LINES2.DAT" For Input As #Temp1FileNumber%
Line Input #Temp1FileNumber%, astring$  ' read comment
Line Input #Temp1FileNumber%, astring$  ' read column labels

For i% = 1 To 99    ' loop on jta elements
Line Input #Temp1FileNumber%, astring$  ' read data line
Call MiscParseStringToString(astring, bstring$)    ' parse symbol
For j% = 1 To 12    ' loop on lines
Call MiscParseStringToString(astring, bstring$)    ' parse value
lines#(j%, i%) = Val(bstring$)
Next j%
Next i%

Line Input #Temp1FileNumber%, astring$  ' read comment
Line Input #Temp1FileNumber%, astring$  ' read column labels

For i% = 1 To 99    ' loop on jta elements
Line Input #Temp1FileNumber%, astring$  ' read data line
Call MiscParseStringToString(astring, bstring$)    ' parse symbol
For j% = 1 To 12    ' loop on lines
Call MiscParseStringToString(astring, bstring$)    ' parse value
edges#(j%, i%) = Val(bstring$)
Next j%
Next i%

Close Temp1FileNumber%

Exit Sub

' Errors
AbsorbLoadLINES2DataFileError:
MsgBox Error$, vbOKOnly + vbCritical, "AbsorbLoadLINES2DataFile"
Close #Temp1FileNumber%
ierror = True
Exit Sub

AbsorbLoadLINES2DataFileNoFile:
msg$ = "File " & ApplicationCommonAppData$ & "LINES2.DAT" & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "AbsorbLoadLINES2DataFile"
ierror = True
Exit Sub

End Sub

Sub AbsorbLoadCHANTLERDataFile()
' Load the Chantler .DAT data files into module level table

ierror = False
On Error GoTo AbsorbLoadCHANTLERDataFileError

Dim i As Integer, j As Integer, n As Integer
Dim astring As String, bstring As String
Dim asym As String, anum As String
Dim tfilename As String

' Loop on all six x-ray files
For n% = 1 To MAXRAY_OLD%

' Read line and edge file if first time
tfilename$ = "CHANTLER2005-" & Xraylo$(n%) & ".dat"
If Dir$(ApplicationCommonAppData$ & tfilename$) = vbNullString Then GoTo AbsorbLoadCHANTLERDataFileNoFile
Open ApplicationCommonAppData$ & tfilename$ For Input As #Temp1FileNumber%
Line Input #Temp1FileNumber%, astring$  ' read first comment line
Line Input #Temp1FileNumber%, astring$  ' read emitter labels (starts with Boron)
Line Input #Temp1FileNumber%, astring$  ' read emitter atomic numbers
Line Input #Temp1FileNumber%, astring$  ' read emitter energies

For i% = 1 To 92    ' loop on absorbing elements
Line Input #Temp1FileNumber%, astring$  ' read data line
Call MiscParseStringToStringA(astring, VbComma$, asym$)   ' parse symbol
Call MiscParseStringToStringA(astring, VbComma$, anum$)   ' parse absorbing atomic number

If Not MiscStringsAreSame(asym$, Symlo$(Val(anum$))) Then GoTo AbsorbLoadCHANTLERDataFileNoMatchAbsorber

For j% = 4 To 92    ' loop on emitting elements
Call MiscParseStringToStringA(astring, VbComma$, bstring$)   ' parse value
allmacs!(j%, i%, n%) = Val(bstring$)
Next j%
Next i%
Close #Temp1FileNumber%

Next n%
Close Temp1FileNumber%

Exit Sub

' Errors
AbsorbLoadCHANTLERDataFileError:
MsgBox Error$, vbOKOnly + vbCritical, "AbsorbLoadCHANTLERDataFile"
Close #Temp1FileNumber%
ierror = True
Exit Sub

AbsorbLoadCHANTLERDataFileNoFile:
msg$ = "File " & ApplicationCommonAppData$ & tfilename$ & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "AbsorbLoadCHANTLERDataFile"
ierror = True
Exit Sub

AbsorbLoadCHANTLERDataFileNoMatchAbsorber:
msg$ = "Absorber symbol " & asym$ & " did not match absorber atomic number " & anum$ & " in file " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "AbsorbLoadCHANTLERDataFile"
ierror = True
Exit Sub

End Sub

Sub AbsorbGetFFAST(ielm As Integer, ip As Integer, iray As Integer, atotal As Single)
' Load the Chantler .DAT data files (ka, kb, la, lb, ma, mb only)
'  ielm = absorber (1 to 92)
'  ip = emitter (5 to 92)
'  iray = x-ray line (1 to 6)

ierror = False
On Error GoTo AbsorbGetFFASTError

atotal! = allmacs!(ip%, ielm%, iray%)
Exit Sub

' Errors
AbsorbGetFFASTError:
MsgBox Error$, vbOKOnly + vbCritical, "AbsorbGetFFAST"
ierror = True
Exit Sub

End Sub

Function AbsorbPow(a As Double, b As Double) As Double
' VB replacement function for C Pow function

ierror = False
On Error GoTo AbsorbPowError

AbsorbPow# = a# ^ b#
Exit Function

' Errors
AbsorbPowError:
MsgBox Error$, vbOKOnly + vbCritical, "AbsorbPow"
ierror = True
Exit Function

End Function

