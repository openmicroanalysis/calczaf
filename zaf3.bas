Attribute VB_Name = "CodeZAF3"
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

Const MAXSHELL% = 30

Dim tmsg As String

Dim jump_ratios(1 To MAXCHAN1%) As Single                                    ' jump ratio, (r-1)/r = fraction of absorbed x-rays causing inner shell ionization (formally ri)

Dim fluor_type(1 To MAXCHAN1%, 1 To MAXCHAN1%) As Integer                       ' type of fluorescence flag (Ka by Ka, etc) (formally iq)
Dim fluor_type2(1 To MAXCHAN1%, 1 To MAXCHAN1%, 1 To MAXRAY% - 1) As Integer    ' type of fluorescence flag (emitter, absorber, absorber x-ray) (formally iq)

Dim fluor_yield(1 To MAXCHAN1%, 1 To MAXCHAN1%) As Single                       ' fluorescent yield of one element by another (formally omj)
Dim fluor_yield2(1 To MAXCHAN1%, 1 To MAXCHAN1%, 1 To MAXRAY% - 1) As Single    ' fluorescent yield of one element by another (emitter, absorber, absorber x-ray) (formally omj)

Dim rela_line_wts(1 To MAXCHAN1%, 1 To MAXCHAN1%) As Single                         ' relative line weights of fluorescing and fluorescenced lines (formally Pij)
Dim rela_line_wts2(1 To MAXCHAN1%, 1 To MAXCHAN1%, 1 To MAXRAY% - 1) As Single      ' relative line weights of fluorescing and fluorescenced lines (emitter, absorber, absorber x-ray) (formally Pij)

Dim over_volt_fluor_line(1 To MAXCHAN1%, 1 To MAXCHAN1%) As Single                          ' overvoltage for the fluorescing line (formally x)
Dim over_volt_fluor_line2(1 To MAXCHAN1%, 1 To MAXCHAN1%, 1 To MAXRAY% - 1) As Single       ' overvoltage for the fluorescing line (emitter, absorber, absorber x-ray) (formally x)

Dim fluor_MACs(1 To MAXRAY% - 1, 1 To MAXCHAN1%, 1 To MAXCHAN1%) As Single          ' MACs for fluorescencing lines (formally fp)

Dim AllJumpRatios(1 To MAXEDG%, 1 To MAXELM%) As Single                                                 ' values from NIST FFAST table

Sub ZAFFlu(zafinit As Integer, zaf As TypeZAF)
' Main stub for fluorescence correction. Comment/uncomment out calls as needed for testing.

ierror = False
On Error GoTo ZAFFluError

'UseFluorescenceByBetaLinesFlag = True         ' for testing purposes only
'UseFluorescenceByBetaLinesFlag = False        ' for testing purposes only
'Call ZAFFlu1(zafinit%, zaf)                   ' original Reed/JTA code
Call ZAFFlu2(zafinit%, zaf)                   ' new working version under development (with fluorescence by beta lines)
'Call ZAFFlu3(zafinit%, zaf)                   ' future fluorescence correction using fundamental parameters
If ierror Then Exit Sub

Exit Sub

' Errors
ZAFFluError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFFlu"
ierror = True
Exit Sub

End Sub

Sub ZAFFlu1(zafinit As Integer, zaf As TypeZAF)
' This routine calculates characteristic fluorescence and loads the MAC's for secondary fluorescing lines if necessary.
'  zafinit% = 0  calculate standard fluorescence correction factors
'  zafinit% = 1  calculate sample fluorescence correction factors
' Original JTA code (without fluorescence of or by beta lines)

ierror = False
On Error GoTo ZAFFlu1Error

Dim temp As Single, mac As Single
Dim in8 As Integer, im4 As Integer
Dim i As Integer, i1 As Integer, i2 As Integer, i4 As Integer

Dim n2 As Single, n3 As Single, n4 As Single
Dim n5 As Single, n6 As Single, m8 As Single
Dim FLUD As Single, FLUX As Single, FLUA As Single, FLUB As Single, FLUC As Single

ReDim m7(1 To MAXCHAN1%) As Single

' If sample calculation, skip fluorescence initialization
If zafinit% = 1 Then GoTo 7200

' Init variables
For i% = 1 To zaf.in1%  ' for each emitter element getting fluorescenced
For i1% = 1 To zaf.in0% ' for each absorber (matrix) element causing fluorescence
rela_line_wts!(i%, i1%) = 0#
Next i1%
Next i%

' z!(i%)  is atomic # of emitting element
' z!(i1%) is atomic # of absorbing element
' mup!(i1%,i%) is mass absorption coefficient (MACs)
For i% = 1 To zaf.in1%      ' for each emitter element being fluoresced
im4% = zaf.il%(i%)
If im4% > MAXRAY% - 1 Then GoTo 7160    ' skip non-emitters

' N6 is upper limit of fluorescence, load for each emitting line
If im4% = 1 Then n6! = 5#   ' Ka
If im4% = 2 Then n6! = 5#   ' Kb
If im4% = 3 Then n6! = 3.5  ' La
If im4% = 4 Then n6! = 3.5  ' Lb
If im4% = 5 Then n6! = 3.5  ' Ma
If im4% = 6 Then n6! = 3.5  ' Mb

' Check to see if any absorber present can be fluorescenced by the emitted lines
For i1% = 1 To zaf.in0%         ' for each absorber (matrix) element causing fluorescence
fluor_type%(i%, i1%) = 0        ' only captures a single fluorescening x-ray for each emitting line
If zaf.il%(i1%) = 2 Or zaf.il%(i1%) = 4 Or zaf.il%(i1%) = 6 Then GoTo 6970  ' skip absorber (matrix) element beta lines

' Variable fluor_type%() is code for type of fluorescence:  0 = none
' 1=Ka by Ka  3=Ka by La  5=Ka by Ma
'13=La by Ka 15=La by La 17=La by Ma
'25=Ma by Ka 27=Ma by La 29=Ma by Ma

' First test for fluorescence by K line, then L, and then M (if indicated).
' Note that "eng!(1 to MAXRAY%-1, 1 to MAXCHAN%)" is the emission line energies for all lines of  an element and "ec!(1 to MAXCHAN%)" is
' the analyzed line absorption edges. While "edg!(1 to MAXEDG%, 1 to MAXCHAN%)" is the absorption edge energies for all lines of an element.
For i2% = 1 To MAXRAY% - 1      ' for each absorber (matrix) x-ray causing fluorescence
If iflu% = 2 And (i2% = 5 Or i2% = 6) Then GoTo 6960        ' only do fluorescence of K and L lines if indicated
If i2% = 2 Or i2% = 4 Or i2% = 6 Then GoTo 6960         ' skip beta lines

' Check for energy greater than critical excitation energy
If zaf.eng!(i2%, i1%) <= zaf.eC!(i%) Then GoTo 6960

' Calculate edge index for this x-ray
If i2% = 1 Then in8% = 1   ' Ka
If i2% = 2 Then in8% = 1   ' Kb
If i2% = 3 Then in8% = 4   ' La
If i2% = 4 Then in8% = 3   ' Lb
If i2% = 5 Then in8% = 9   ' Ma
If i2% = 6 Then in8% = 8   ' Mb

' Check for missing absorption edge data
If zaf.edg!(in8%, i1%) = 0# Then
msg$ = "WARNING in ZAFFlu1- " & Format$(Symup$(zaf.Z%(i1%)), a20$) & " " & Format$(Edglo$(in8%), a20$) & " absorption edge is zero"
Call IOWriteLog(msg$)
GoTo 6960
End If

' Check if emission energy is greater than upper limit of fluorescence or edge energy is greater than electron energy (zaf.eO!(i%))
If zaf.eng!(i2%, i1%) > (zaf.eC!(i%) + n6!) Or zaf.edg!(in8%, i1%) >= zaf.eO!(i%) Then GoTo 6960

' Set fluorescence flag for type of fluorescence
fluor_type%(i%, i1%) = i2% + (MAXRAY% - 1) * (zaf.il%(i%) - 1)

' Warn if line is fluoresced
If VerboseMode Then
msg$ = "WARNING in ZAFFlu1- the " & Format$(Xraylo$(im4%), a20$) & " line of " & Format$(Symup$(zaf.Z%(i%)), a20$) & " is excited by the " & Format$(Xraylo$(i2%), a20$) & " line of " & Format$(Symup$(zaf.Z%(i1%)), a20$)
Call IOWriteLog(msg$)
End If

' Calculate overvoltage of the fluorescing line
over_volt_fluor_line!(i%, i1%) = zaf.eO!(i%) / zaf.edg!(in8%, i1%)

' Read in K, L or M shell MACs for fluorescing element. Check to see if MACs are already loaded by ZAFReadMu for primary lines (this
' includes empirical MACs for primary lines). Note "zaf.mup!()" is the mass absorption coefficient in ug/cm2 (absorbed,absorbing)
If i2% = zaf.il%(i1%) Then
For i4% = 1 To zaf.in0%
fluor_MACs!(i2%, i4%, i1%) = zaf.mup!(i4%, i1%)
Next i4%
GoTo 6960
End If

' Need additional MACs for fluorescencing lines, load MACs from disk
For i4% = 1 To zaf.in0%
Call ZAFLoadMac(i1%, i2%, i4%, mac!, zaf)
If ierror Then Exit Sub

' If MAC is zero, type error message and turn off fluorescence flag
If mac! > 0# Then
fluor_MACs!(i2%, i4%, i1%) = mac!

Else
If VerboseMode Then
msg$ = "WARNING in ZAFFlu1- MAC not loaded for " & Format$(Symup$(zaf.Z%(i1%)), a20$) & " " & Format$(Xraylo$(i2%), a20$) & " in " & Format$(Symup$(zaf.Z%(i4%)), a20$) & ", fluorescence contribution will not be calculated for this line."
Call IOWriteLog(msg$)
End If
fluor_MACs!(i2%, i4%, i1%) = 1#
fluor_type%(i%, i1%) = 0
End If

Next i4%
6960:  Next i2%     ' for each absorber (matrix) x-ray causing fluorescence
6970:  Next i1%     ' for each absorber (matrix) element causing fluorescence

' Load jump ratios
If zaf.il%(i%) = 1 Then jump_ratios!(i%) = 1.11728 - 0.07368 * Log(zaf.Z%(i%))     ' Ka line (0.88)
If zaf.il%(i%) = 3 Then jump_ratios!(i%) = 0.95478 - 0.00259 * zaf.Z%(i%)          ' La line (0.75)
If zaf.il%(i%) = 5 Then jump_ratios!(i%) = 0.5                                     ' Ma line (0.5)

' Load fluorescent yields using pointer from above
For i1% = 1 To zaf.in0% ' for each absorber element causing fluorescence
If fluor_type%(i%, i1%) > 0 Then    ' skip if no fluorescence

' Variable fluor_yield!(i%,i1%) is fluorescent yield of element i1% by element i%
If fluor_type%(i%, i1%) = 1 Or fluor_type%(i%, i1%) = 13 Or fluor_type%(i%, i1%) = 25 Then fluor_yield!(i%, i1%) = zaf.flu!(1, i1%)   ' fluorescence by Ka of absorber (matrix) element
If fluor_type%(i%, i1%) = 3 Or fluor_type%(i%, i1%) = 15 Or fluor_type%(i%, i1%) = 27 Then fluor_yield!(i%, i1%) = zaf.flu!(3, i1%)   ' fluorescence by La of absorber (matrix) element
If fluor_type%(i%, i1%) = 5 Or fluor_type%(i%, i1%) = 17 Or fluor_type%(i%, i1%) = 29 Then fluor_yield!(i%, i1%) = zaf.flu!(5, i1%)   ' fluorescence by Ma of absorber (matrix) element

' rela_line_wts!(i%,i1%) = relative line weights
If fluor_type%(i%, i1%) = 1 Then rela_line_wts!(i%, i1%) = 1#    ' Ka by Ka
If fluor_type%(i%, i1%) = 3 Then rela_line_wts!(i%, i1%) = 4.2   ' Ka by La (Reed)
If fluor_type%(i%, i1%) = 5 Then rela_line_wts!(i%, i1%) = 0#    ' Ka by Ma (?)
If fluor_type%(i%, i1%) = 13 Then rela_line_wts!(i%, i1%) = 0.24 ' La by Ka (Reed)
If fluor_type%(i%, i1%) = 15 Then rela_line_wts!(i%, i1%) = 1#   ' La by La
If fluor_type%(i%, i1%) = 17 Then rela_line_wts!(i%, i1%) = 0#   ' La by Ma (?)
If fluor_type%(i%, i1%) = 25 Then rela_line_wts!(i%, i1%) = 0.02 ' Ma by Ka (Armstrong)
If fluor_type%(i%, i1%) = 27 Then rela_line_wts!(i%, i1%) = 0.02 ' Ma by La (Armstrong)
If fluor_type%(i%, i1%) = 29 Then rela_line_wts!(i%, i1%) = 1#   ' Ma by Ma
End If
Next i1%

7160:  Next i%      ' next emitting element
Exit Sub

' Sample fluorescence correction
7200:
If DebugMode And VerboseMode% Then
msg$ = vbCrLf & Format$(Space$(11) & "FluTyp", a40$) & Format$("LinWt", a80$) & Format$("FluYld", a80$) & Format$("JmpFac", a80$) & Format$("MACSmp", a100$) & Format$("FluRay", a80$) & Format$("MACFlu", a100$)
Call IOWriteLog$(msg$)
End If

' For each emitting element
For i% = 1 To zaf.in1%
If zaf.il%(i%) > MAXRAY% - 1 Or zaf.conc!(i%) <= 0# Then GoTo 7350
zaf.vv!(i%) = 0#

' Weight fraction sum of MACs
m7!(i%) = ZAFMACCal(i%, zaf)
If m7!(i%) < 0# Then GoTo ZAFFlu1NegativeMAC

' Pointer to MAC of fluorescing line
For i1% = 1 To zaf.in0%
If zaf.il%(i1%) = 2 Or zaf.il%(i1%) = 4 Or zaf.il%(i1%) = 6 Then GoTo 7340  ' skip absorber (matrix) beta lines

If fluor_type%(i%, i1%) = 0 Then GoTo 7340
m8! = 0#
i2% = fluor_type%(i%, i1%) - (MAXRAY% - 1) * (zaf.il%(i%) - 1)  ' get index for MAC of fluorescencing line

For i4% = 1 To zaf.in0%
m8! = m8! + zaf.conc!(i4%) * fluor_MACs!(i2%, i4%, i1%)
Next i4%

' "m1!" is 1.0/SIN(takeoff*3.14159/180.0)
n2! = zaf.m1!(i%) * m7!(i%) / m8!
n3! = 333000# / (zaf.eO!(i%) ^ 1.65 - zaf.eC!(i%) ^ 1.65) / m8!
If (1# + n2!) < 0# Or (1# + n3!) < 0# Then GoTo ZAFFlu1NegativeFlu
n4! = Log(1# + n2!) / n2! + Log(1# + n3!) / n3!
n5! = zaf.conc!(i1%) * rela_line_wts!(i%, i1%) * fluor_yield!(i%, i1%) * jump_ratios!(i%) / 2#

' was n5! = n5!*((over_volt_fluor_line!(i%,i1%) - 1.0)/(zaf.v!(i%) - 1.0))^1.67
temp! = (over_volt_fluor_line!(i%, i1%) - 1#) / (zaf.v!(i%) - 1#)

' See Armstrong, "Microbeam Analysis", 1988, p. 239
If iflu% = 1 Then
If temp! < 2# / 3# Then n5! = n5! * temp! ^ 1.59
If temp! >= 2# / 3# Then n5! = n5! * 1.87 * temp! ^ 3.19
End If

If iflu% = 2 Or iflu% = 3 Then
n5! = n5! * temp! ^ 1.67
End If

' See Reed, "Microbeam Analysis", 1993, p. 109
If iflu% = 4 Then
temp! = (over_volt_fluor_line!(i%, i1%) * Log(over_volt_fluor_line!(i%, i1%)) - over_volt_fluor_line!(i%, i1%) + 1#) / (zaf.v!(i%) * Log(zaf.v!(i%)) - zaf.v!(i%) + 1#)
n5! = n5! * temp!
End If

' Return fluorescence term in "vv!(i%)"
n5! = n5! * zaf.atwts!(i%) / zaf.atwts!(i1%)

' PTC modification (THIN FILM CORRECTION ACCURATE ONLY FOR FILMS < 1 mg/cm^2 THICK)
If UseParticleCorrectionFlag And iptc% = 1 And (zaf.d! * zaf.j9! < 0.001) Then
' MODIFIED NOCKOLDS FLUORESCENCE CORRECTION FOR THIN FILMS
n4! = m8! * zaf.d! * zaf.j9! * (0.923 - Log(m8! * zaf.d! * zaf.j9!))
' MODIFIED TO USE REED RELATIVE LINE INTENSITIES
End If
    
n5! = n5! * n4! * fluor_MACs!(i2%, i%, i1%) / m8!

' PTC modification (ARMSTRONG/BUSECK FLUORESCENCE CORRECTION FOR PARTICLES)
If UseParticleCorrectionFlag And iptc% = 1 And zaf.model% <> 1 Then
FLUX! = 1 - Exp(-m8! * zaf.d! / 2#)
FLUA! = 0.026
FLUB! = 1.1409 + 0.2012 * n2!
FLUC! = -0.2471 - 0.2741 * n2! + 0.01315 * n2! * n2!
FLUD! = FLUA! + FLUB! * FLUX! + FLUC! * FLUX! * FLUX!
If FLUD! >= 0# And FLUD! <= 1! Then n5! = FLUD! * n5!
End If

zaf.vv!(i%) = zaf.vv!(i%) + n5!

If DebugMode And VerboseMode% Then
msg$ = Format$(Symup$(zaf.Z%(i%)), a20$) & " " & Format$(Xraylo$(zaf.il(i%)), a20$) & " by " & Format$(Symup$(zaf.Z%(i1%)), a20$)
msg$ = msg$ & ", " & Format$(fluor_type%(i%, i1%), a40$) & MiscAutoFormatB$(rela_line_wts!(i%, i1%)) & MiscAutoFormat$(fluor_yield!(i%, i1%)) & MiscAutoFormat$(jump_ratios!(i%)) & "  " & Format$(m7!(i%), e82$) & MiscAutoFormatI$(i2%) & "  " & Format$(m8!, e82$)
Call IOWriteLog(msg$)
End If

7340:  Next i1%
7350:  Next i%
Exit Sub

' Errors
ZAFFlu1Error:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFFlu1"
ierror = True
Exit Sub

ZAFFlu1NegativeMAC:
msg$ = "Average MAC for " & Format$(Symup$(zaf.Z%(i%)), a20$) & " in this matrix is negative for line " & Str$(zaf.n8) & ", and is probably a bad data point (epoxy, etc.). Delete the analysis line and try again."
If Not CalcImageQuantFlag Then
MiscMsgBoxTim FormMSGBOXTIME, "ZAFFlu1", msg$, 20#
Call IOWriteLog(msg$)
Else
Call IOWriteLog(msg$)
End If
'ierror = True
Exit Sub

ZAFFlu1NegativeFlu:
msg$ = "Fluorescence factor for " & Format$(Symup$(zaf.Z%(i%)), a20$) & " in this matrix is negative for line " & Str$(zaf.n8) & ", and is probably a bad data point (epoxy, etc.). Delete the analysis line and try again."
If Not CalcImageQuantFlag Then
MiscMsgBoxTim FormMSGBOXTIME, "ZAFFlu1", msg$, 20#
Call IOWriteLog(msg$)
Else
Call IOWriteLog(msg$)
End If
'ierror = True
Exit Sub

End Sub

Sub ZAFFlu2(zafinit As Integer, zaf As TypeZAF)
' This routine calculates characteristic fluorescence and loads the MAC's for secondary fluorescing lines if necessary.
'  zafinit% = 0  calculate standard fluorescence correction factors
'  zafinit% = 1  calculate sample fluorescence correction factors
' Donovan/Gainsforth modified code using N. Ritchie jump ratio table and JJD fluorescence by beta lines code.

ierror = False
On Error GoTo ZAFFlu2Error

Dim temp As Single, mac As Single
Dim in8 As Integer, i As Integer, i1 As Integer, i2 As Integer, i4 As Integer

Dim n2 As Single, n3 As Single, n4 As Single
Dim n5 As Single, n6 As Single, m8 As Single
Dim FLUD As Single, FLUX As Single, FLUA As Single, FLUB As Single, FLUC As Single
Dim flu_yield_emitter As Single, flu_yield_matrix As Single

ReDim m7(1 To MAXCHAN1%) As Single

' If sample calculation, skip fluorescence initialization
If zafinit% = 1 Then GoTo 7200

' Init variables
For i% = 1 To zaf.in1%  ' for each emitter element
    For i1% = 1 To zaf.in0% ' for each absorber (matrix) element
        For i2% = 1 To MAXRAY% - 1 ' for each absorber (matrix) x-ray
        rela_line_wts2!(i%, i1%, i2%) = 0.001       ' assume small relative line weight
        Next i2%
    Next i1%
Next i%

' z!(i%)  is atomic # of emitting element being fluoresced
' z!(i1%) is atomic # of absorbing (matrix) element causing fluorescence
' mup!(i1%,i%) is mass absorption coefficient (MACs)
For i% = 1 To zaf.in1%  ' for each emitter getting fluoresced
If zaf.il%(i%) > MAXRAY% - 1 Then GoTo 7160    ' skip non-emitters

    ' N6 is upper limit of fluorescence, load for each emitting line
    If zaf.il%(i%) = 1 Then n6! = 5#   ' Ka
    If zaf.il%(i%) = 2 Then n6! = 5#   ' Kb
    If zaf.il%(i%) = 3 Then n6! = 3.5  ' La
    If zaf.il%(i%) = 4 Then n6! = 3.5  ' Lb
    If zaf.il%(i%) = 5 Then n6! = 3.5  ' Ma
    If zaf.il%(i%) = 6 Then n6! = 3.5  ' Mb

    If zaf.il%(i%) = 7 Then n6! = 3.5   ' Ln
    If zaf.il%(i%) = 8 Then n6! = 3.5   ' Lg
    If zaf.il%(i%) = 9 Then n6! = 3.5   ' Lv
    If zaf.il%(i%) = 10 Then n6! = 3.5  ' Ll
    If zaf.il%(i%) = 11 Then n6! = 3.5  ' Mg
    If zaf.il%(i%) = 12 Then n6! = 3.5  ' Mz

    ' Check to see if any absorber (matrix) element present can fluoresce the emitted element/line
    For i1% = 1 To zaf.in0%     ' for each absorber (matrix) element causing fluorescence
        If Not UseFluorescenceByBetaLinesFlag Then
        If zaf.il%(i1%) = 2 Or zaf.il%(i1%) = 4 Or zaf.il%(i1%) = 6 Then GoTo 6970  ' skip fluorescence by beta lines
        'If zaf.il%(i1%) > MAXRAY_OLD% Then GoTo 6970                                ' skip fluorescence by additional x-ray lines
        End If

            ' First test for fluorescence by K line, then L, and then M (if indicated).
            ' Note that "eng!(1 to MAXRAY%-1, 1 to MAXCHAN%)" is the emission line energies for all lines of
            ' an element and "ec!(1 to MAXCHAN%)" is the analyzed line absorption edges. While
            ' "edg!(1 to MAXEDG%, 1 to MAXCHAN%)" is the absorption edge energies for all lines of an element.
            For i2% = 1 To MAXRAY% - 1  ' for each x-ray for each absorber (matrix) element that might cause fluorescence in the emitter element/line
            fluor_type2%(i%, i1%, i2%) = 0
            If iflu% = 2 And (i2% = 5 Or i2% = 6) Then GoTo 6960    ' only do fluorescence of K and L lines if indicated
                If Not UseFluorescenceByBetaLinesFlag Then
                If i2% = 2 Or i2% = 4 Or i2% = 6 Then GoTo 6960     ' skip fluorescence by beta lines
                'If i2% > MAXRAY_OLD% Then GoTo 6960                 ' skip fluorescence by additional x-ray lines
                End If

            ' Check for energy greater than critical excitation energy
            If zaf.eng!(i2%, i1%) <= zaf.eC!(i%) Then GoTo 6960

            ' Calculate edge index for this x-ray
            If i2% = 1 Then in8% = 1   ' Ka
            If i2% = 2 Then in8% = 1   ' Kb
            If i2% = 3 Then in8% = 4   ' La
            If i2% = 4 Then in8% = 3   ' Lb
            If i2% = 5 Then in8% = 9   ' Ma
            If i2% = 6 Then in8% = 8   ' Mb

            If i2% = 7 Then in8% = 3   ' Ln
            If i2% = 8 Then in8% = 3   ' Lg
            If i2% = 9 Then in8% = 3   ' Lv
            If i2% = 10 Then in8% = 4   ' Ll
            If i2% = 11 Then in8% = 7   ' Mg
            If i2% = 12 Then in8% = 9   ' Mz

            ' Check for missing absorption edge data
            If zaf.edg!(in8%, i1%) = 0# Then
            msg$ = "WARNING in ZAFFlu2- " & Format$(Symup$(zaf.Z%(i1%)), a20$) & " " & Format$(Edglo$(in8%), a20$) & " absorption edge is zero"
            Call IOWriteLog(msg$)
            GoTo 6960
            End If

            ' Check if emission energy is greater than upper limit of fluorescence or edge energy is greater than electron energy (zaf.eO!(i%))
            If zaf.eng!(i2%, i1%) > (zaf.eC!(i%) + n6!) Or zaf.edg!(in8%, i1%) >= zaf.eO!(i%) Then GoTo 6960

            ' Set fluorescence flag for type of fluorescence
            fluor_type2%(i%, i1%, i2%) = i2% + (MAXRAY% - 1) * (zaf.il%(i%) - 1)

            ' Warn if line is fluoresced
            If VerboseMode Then
            msg$ = "WARNING in ZAFFlu2- the " & Format$(Xraylo$(zaf.il%(i%)), a20$) & " line of " & Format$(Symup$(zaf.Z%(i%)), a20$) & " is excited by the " & Format$(Xraylo$(i2%), a20$) & " line of " & Format$(Symup$(zaf.Z%(i1%)), a20$)
            Call IOWriteLog(msg$)
            End If

            ' Calculate overvoltage of the fluorescing line
            over_volt_fluor_line2!(i%, i1%, i2%) = zaf.eO!(i%) / zaf.edg!(in8%, i1%)

            ' Read in K, L or M shell MACs for fluorescing element. Check to see if MACs are already loaded by ZAFReadMu for primary lines (this
            ' includes empirical MACs for primary lines). Note "zaf.mup!()" is the mass absorption coefficient in ug/cm2 (absorbed,absorbing)
                If i2% = zaf.il%(i1%) Then
                    For i4% = 1 To zaf.in0%
                    fluor_MACs!(i2%, i4%, i1%) = zaf.mup!(i4%, i1%)
                    Next i4%
                    GoTo 6960
                End If

                ' Need additional MACs for fluorescencing lines, load MACs from disk
                For i4% = 1 To zaf.in0%
                Call ZAFLoadMac(i1%, i2%, i4%, mac!, zaf)
                If ierror Then Exit Sub

                    ' If MAC is zero, type error message and turn off fluorescence flag
                    If mac! > 0# Then
                    fluor_MACs!(i2%, i4%, i1%) = mac!

                    Else
                    If VerboseMode Then
                    msg$ = "WARNING in ZAFFlu2- MAC not loaded for " & Format$(Symup$(zaf.Z%(i1%)), a20$) & " " & Format$(Xraylo$(i2%), a20$) & " in " & Format$(Symup$(zaf.Z%(i4%)), a20$) & ", fluorescence contribution will not be calculated for this line."
                    Call IOWriteLog(msg$)
                    End If
            
                fluor_MACs!(i2%, i4%, i1%) = 1#
                fluor_type2%(i%, i1%, i2%) = 0
                End If
                Next i4%
6960:       Next i2%        ' next absorber (matrix) x-ray causing fluorescence
6970:   Next i1%        ' next absorber (matrix) element causing fluorescence

    ' Load from Jump_ratios.dat (from Nicholas Ritchie, NIST)
    Call ZAFFLULoadJumpRatios(zaf.il%(i%), zaf.Z%(i%), jump_ratios!(i%))
    If ierror Then Exit Sub

        ' Load fluorescent yields using pointer from above
        For i1% = 1 To zaf.in0%         ' for each absorber (matrix) element causing fluorescence
            For i2% = 1 To MAXRAY% - 1      ' for each absorber (matrix) x-ray causing fluorescence
            If fluor_type2%(i%, i1%, i2%) > 0 Then    ' skip if no fluorescence (zero)

            ' Variable fluor_yield2!(i%, i1%, i2%) is fluorescent yield of emitting element i% by absorbing (matrix) element i1%, line i2%
            fluor_yield2!(i%, i1%, i2%) = ZAFFLUGetFluYield(fluor_type2%(i%, i1%, i2%), i1%, zaf)
            
            ' Load relative line weights using "tuned" values (no Z or keV dependency)
            Call ZAFFLULoadLineWeightsReed(i%, i1%, i2%, fluor_type2%(), rela_line_wts2!())
            If ierror Then Exit Sub
           
            ' Load Z correlated relative line weights to improve accuracy (i% = emitter, i1% = matrix, i2% = matrix (fluorescing) xray)
            'flu_yield_emitter! = ZAFFLUGetFluYield(fluor_type2%(i%, i1%, i2%), i1%, zaf)        ' not correct?
            'flu_yield_matrix! = ZAFFLUGetFluYield(fluor_type2%(i%, i1%, i2%), i1%, zaf)        ' not correct?
            'Call ZAFFLULoadLineWeightsPenepma(i%, i1%, zaf.TOA!, zaf.eO!(i%), zaf.Z%(i%), zaf.il%(i%), zaf.TOA!, zaf.eO!(i1%), zaf.Z%(i1%), i2%, rela_line_wts2!(i%, i1%, i2%), flu_yield_emitter!, flu_yield_matrix!, zaf)
            'If ierror Then Exit Sub
            
            End If
            Next i2%        ' next absorber (matrix) x-ray causing fluorecence
          Next i1%      ' next absorber (matrix) element causing fluorescence
7160:  Next i%      ' next emitter element getting fluoresced
Exit Sub

' Sample fluorescence correction
7200:
If DebugMode And VerboseMode% Then
msg$ = vbCrLf & Format$(Space$(14) & "FluTyp", a40$) & Format$("LinWt", a80$) & Format$("FluYld", a80$) & Format$("JmpFac", a80$) & Format$("MACSmp", a100$) & Format$("FluRay", a80$) & Format$("MACFlu", a100$)
Call IOWriteLog$(msg$)
End If

' For each emitting element line getting fluoresced
For i% = 1 To zaf.in1%
If zaf.il%(i%) > MAXRAY% - 1 Or zaf.conc!(i%) <= 0# Then GoTo 7350
zaf.vv!(i%) = 0#

    ' Weight fraction sum of MACs
    m7!(i%) = ZAFMACCal(i%, zaf)
    If m7!(i%) < 0# Then GoTo ZAFFlu2NegativeMAC

    ' For each fluorescing absorber (matrix) element
    For i1% = 1 To zaf.in0%
        
        ' For each fluorescing x-ray line of the absorber (matrix) element
        For i2% = 1 To MAXRAY% - 1
        If fluor_type2%(i%, i1%, i2%) = 0 Then GoTo 7340
        If Not UseFluorescenceByBetaLinesFlag Then
        If zaf.il%(i1%) = 2 Or zaf.il%(i1%) = 4 Or zaf.il%(i1%) = 6 Then GoTo 7340  ' skip beta lines
        'If zaf.il%(i1%) > MAXRAY_OLD% Then GoTo 7340                                ' skip fluorescence by additional x-ray lines
        End If
     
        ' Get x-ray line index for MAC of fluorescencing line
        i2% = fluor_type2%(i%, i1%, i2%) - (MAXRAY% - 1) * (zaf.il%(i%) - 1)

            m8! = 0#
            For i4% = 1 To zaf.in0%
            m8! = m8! + zaf.conc!(i4%) * fluor_MACs!(i2%, i4%, i1%)
            Next i4%

        ' "m1!" is 1.0/SIN(takeoff*3.14159/180.0)
        ' "rela_line_wts!(1 to MAXCHAN1%, 1 to MAXCHAN1%, 1 to MAXRAY% -1)" is the relative line weights loaded above
        n2! = zaf.m1!(i%) * m7!(i%) / m8!
        n3! = 333000# / (zaf.eO!(i%) ^ 1.65 - zaf.eC!(i%) ^ 1.65) / m8!
        If (1# + n2!) < 0# Or (1# + n3!) < 0# Then GoTo ZAFFlu2NegativeFlu
        n4! = Log(1# + n2!) / n2! + Log(1# + n3!) / n3!
        n5! = zaf.conc!(i1%) * rela_line_wts2!(i%, i1%, i2%) * fluor_yield2!(i%, i1%, i2%) * jump_ratios!(i%) / 2#

        ' was n5! = n5!*((over_volt_fluor_line2!(i%, i1%, i2%) - 1.0)/(zaf.v!(i%) - 1.0))^1.67
        temp! = (over_volt_fluor_line2!(i%, i1%, i2%) - 1#) / (zaf.v!(i%) - 1#)

        ' See Armstrong, "Microbeam Analysis", 1988, p. 239
        If iflu% = 1 Then
        If temp! < 2# / 3# Then n5! = n5! * temp! ^ 1.59
        If temp! >= 2# / 3# Then n5! = n5! * 1.87 * temp! ^ 3.19
        End If

        If iflu% = 2 Or iflu% = 3 Then
        n5! = n5! * temp! ^ 1.67
        End If

        ' See Reed, "Microbeam Analysis", 1993, p. 109
        If iflu% = 4 Then
        temp! = (over_volt_fluor_line2!(i%, i1%, i2%) * Log(over_volt_fluor_line2!(i%, i1%, i2%)) - over_volt_fluor_line2!(i%, i1%, i2%) + 1#) / (zaf.v!(i%) * Log(zaf.v!(i%)) - zaf.v!(i%) + 1#)
        n5! = n5! * temp!
        End If

        ' Return fluorescence term in "vv!(i%)"
        n5! = n5! * zaf.atwts!(i%) / zaf.atwts!(i1%)

        ' PTC modification (THIN FILM CORRECTION ACCURATE ONLY FOR FILMS < 1 mg/cm^2 THICK)
        If UseParticleCorrectionFlag And iptc% = 1 And (zaf.d! * zaf.j9! < 0.001) Then
        ' MODIFIED NOCKOLDS FLUORESCENCE CORRECTION FOR THIN FILMS
        n4! = m8! * zaf.d! * zaf.j9! * (0.923 - Log(m8! * zaf.d! * zaf.j9!))
        ' MODIFIED TO USE REED RELATIVE LINE INTENSITIES
        End If
    
        n5! = n5! * n4! * fluor_MACs!(i2%, i%, i1%) / m8!

        ' PTC modification (ARMSTRONG/BUSECK FLUORESCENCE CORRECTION FOR PARTICLES)
        If UseParticleCorrectionFlag And iptc% = 1 And zaf.model% <> 1 Then
        FLUX! = 1 - Exp(-m8! * zaf.d! / 2#)
        FLUA! = 0.026
        FLUB! = 1.1409 + 0.2012 * n2!
        FLUC! = -0.2471 - 0.2741 * n2! + 0.01315 * n2! * n2!
        FLUD! = FLUA! + FLUB! * FLUX! + FLUC! * FLUX! * FLUX!
        If FLUD! >= 0# And FLUD! <= 1! Then n5! = FLUD! * n5!
        End If

        zaf.vv!(i%) = zaf.vv!(i%) + n5!

        If DebugMode And VerboseMode% Then
        msg$ = Format$(Symup$(zaf.Z%(i%)), a20$) & " " & Format$(Xraylo$(zaf.il(i%)), a20$) & " by " & Format$(Symup$(zaf.Z%(i1%)), a20$) & " " & Format$(Xraylo$(i2%), a20$)
        msg$ = msg$ & ", " & Format$(fluor_type2%(i%, i1%, i2%), a40$) & MiscAutoFormatB$(rela_line_wts2!(i%, i1%, i2%)) & MiscAutoFormat$(fluor_yield2!(i%, i1%, i2%)) & MiscAutoFormat$(jump_ratios!(i%)) & "  " & Format$(m7!(i%), e82$) & MiscAutoFormatI$(i2%) & "  " & Format$(m8!, e82$)
        Call IOWriteLog(msg$)
        End If

7340:  Next i2%     ' next absorber (matrix) x-ray line
       Next i1%     ' next absorber (matrix) element
7350:  Next i%      ' next emitting element

Exit Sub

' Errors
ZAFFlu2Error:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFFlu2"
ierror = True
Exit Sub

ZAFFlu2NegativeMAC:
msg$ = "Average MAC for " & Format$(Symup$(zaf.Z%(i%)), a20$) & " in this matrix is negative for line " & Str$(zaf.n8) & ", and is probably a bad data point (epoxy, etc.). Delete the analysis line and try again."
If Not CalcImageQuantFlag Then
MiscMsgBoxTim FormMSGBOXTIM, "ZAFFlu2", msg$, 10#
Else
Call IOWriteLog(msg$)
End If
'ierror = True
Exit Sub

ZAFFlu2NegativeFlu:
msg$ = "Fluorescence factor for " & Format$(Symup$(zaf.Z%(i%)), a20$) & " in this matrix is negative for line " & Str$(zaf.n8) & ", and is probably a bad data point (epoxy, etc.). Delete the analysis line and try again."
If Not CalcImageQuantFlag Then
MiscMsgBoxTim FormMSGBOXTIM, "ZAFFlu2", msg$, 10#
Else
Call IOWriteLog(msg$)
End If
'ierror = True
Exit Sub

End Sub

Sub ZAFFlu3(zafinit As Integer, zaf As TypeZAF)
' New fundamental parameters fluorescence code to handle both characteristic and continuum fluorescence

ierror = False
On Error GoTo ZAFFlu3Error

' Not implemented
msg$ = "Feature not implemented at this time"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFFlu3"

Exit Sub

' Errors
ZAFFlu3Error:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFFlu3"
ierror = True
Exit Sub

End Sub

Sub ZAFFLULoadJumpRatios(iray As Integer, ielm As Integer, tJumpRatio As Single)
' Load the jump ratios (factors) from the Jump_Ratios.dat file from NIST

ierror = False
On Error GoTo ZAFFLULoadJumpRatiosError

Dim iz As Integer, i As Integer, in8 As Integer
Dim isym As String, astring As String, bstring As String
Dim tfilename As String

Dim tvalues(1 To MAXEDG%) As Single

Static initialized As Boolean
           
' Open the input (comma, tab or space delimited) and output files
If Not initialized Then
Close #Temp1FileNumber%
DoEvents

tfilename$ = ApplicationCommonAppData$ & "Jump_Ratios.dat"
Open tfilename$ For Input As #Temp1FileNumber%
'Call IOWritelog(vbcrlf & "Opening " & tfilename$)

' Read first line of column headings
Line Input #Temp1FileNumber%, astring   ' read comment
'Call IOWriteLog(astring$)
Line Input #Temp1FileNumber%, astring   ' read comment
'Call IOWriteLog(astring$)
Line Input #Temp1FileNumber%, astring   ' read column labels
'Call IOWriteLog(astring$)

' Loop on entries
Call IOStatusAuto(vbNullString)
icancelauto = False
Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring   ' read data line

' Parse out fluorescent yields
Call MiscParseStringToStringA(astring$, VbComma, bstring$)
If ierror Then Exit Sub
isym$ = bstring$
Call MiscParseStringToStringA(astring$, VbComma, bstring$)
If ierror Then Exit Sub
iz% = Val(bstring$)

' Check for valid values
If iz% < 1 Or iz% > MAXELM% Then GoTo ZAFFLULoadJumpRatiosBadEmitter

' Load jump ratio values
tmsg$ = "IZ=" & Format$(iz%) & ", "
For i% = 1 To MAXEDG%
Call MiscParseStringToStringA(astring$, VbComma, bstring$)
If ierror Then Exit Sub
tvalues!(i%) = Val(bstring$)
tmsg$ = tmsg$ & MiscAutoFormat$(tvalues!(i%))
Next i%
'Call IOWriteLog(tmsg$)

' Load data value (ri - 1)/ri. That is, convert from jump "factor" to jump "ratio"
For i% = 1 To MAXEDG%
If tvalues!(i%) > 0# Then AllJumpRatios!(i%, iz%) = (tvalues!(i%) - 1) / tvalues!(i%)
Next i%
Loop

Close #Temp1FileNumber%
'Call IOWriteLog(vbNullString)
initialized = True
End If

' Calculate edge index for this x-ray
If iray% = 1 Then in8% = 1   ' Ka
If iray% = 2 Then in8% = 1   ' Kb
If iray% = 3 Then in8% = 4   ' La
If iray% = 4 Then in8% = 3   ' Lb
If iray% = 5 Then in8% = 9   ' Ma
If iray% = 6 Then in8% = 8   ' Mb

If iray% = 7 Then in8% = 3    ' Ln
If iray% = 8 Then in8% = 3    ' Lg
If iray% = 9 Then in8% = 3    ' Lv
If iray% = 10 Then in8% = 4   ' Ll
If iray% = 11 Then in8% = 7   ' Mg
If iray% = 12 Then in8% = 9   ' Mz

' Return requested value
tJumpRatio! = AllJumpRatios!(in8%, ielm%)

Exit Sub

' Errors
ZAFFLULoadJumpRatiosError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFFLULoadJumpRatios"
ierror = True
Close #Temp1FileNumber%
Exit Sub

ZAFFLULoadJumpRatiosBadEmitter:
msg$ = "Invalid atomic number in " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFFLULoadJumpRatios"
ierror = True
Close #Temp1FileNumber%
Exit Sub

End Sub

Function ZAFFLUGetFluYield(tFluTyp As Integer, iabs As Integer, zaf As TypeZAF) As Single
' Return the fluorescent yield for the specified fluorescence type

ierror = False
On Error GoTo ZAFFLUGetFluYieldError

ZAFFLUGetFluYield! = 0#

' Load fluorescent yields for original and additional lines here...
'  1=Ka by Ka   2=Ka by Kb   3=Ka by La    4=Ka by Lb   5=Ka by Ma   6=Ka by Mb     7=Ka by Ln   8=Ka by Lg   9=Ka by Lv   10=Ka by Ll  11=Ka by Mg  12=Ka by Mz
' 13=Kb by Ka  14=Kb by Kb  15=Kb by La   16=Kb by Lb  17=Kb by Ma  18=Kb by Mb    19=Kb by Ln  20=Kb by Lg  21=Kb by Lv   22=Kb by Ll  23=Kb by Mg  24=Kb by Mz
' 25=La by Ka  26=La by Kb  27=La by La   28=La by Lb  29=La by Ma  30=La by Mb    31=La by Ln  32=La by Lg  33=La by Lv   34=La by Ll  35=La by Mg  36=La by Mz
' 37=Lb by Ka  38=Lb by Kb  39=Lb by La   40=Lb by Lb  41=Lb by Ma  42=Lb by Mb    43=Lb by Ln  44=Lb by Lg  45=Lb by Lv   46=Lb by Ll  47=Lb by Mg  48=Lb by Mz
' 49=Ma by Ka  50=Ma by Kb  51=Ma by La   52=Ma by Lb  53=Ma by Ma  54=Ma by Mb    55=Ma by Ln  56=Ma by Lg  57=Ma by Lv   58=Ma by Ll  59=Ma by Mg  60=Ma by Mz
' 61=Mb by Ka  62=Mb by Kb  63=Mb by La   64=Mb by Lb  65=Mb by Ma  66=Mb by Mb    67=Mb by Ln  68=Mb by Lg  69=Mb by Lv   70=Mb by Ll  71=Mb by Mg  72=Mb by Mz

' 73=Ln by Ka  74=Ln by Kb  75=Ln by La   76=Ln by Lb  77=Ln by Ma  78=Ln by Mb    79=Ln by Ln  80=Ln by Lg  81=Ln by Lv   82=Ln by Ll  83=Ln by Mg  84=Ln by Mz
' 85=Lg by Ka  86=Lg by Kb  87=Lg by La   88=Lg by Lb  89=Lg by Ma  90=Lg by Mb    91=Lg by Ln  92=Lg by Lg  93=Lg by Lv   94=Lg by Ll  95=Lg by Mg  96=Lg by Mz
' 97=Lv by Ka  98=Lv by Kb  99=Lv by La  100=Lv by Lb 101=Lv by Ma 102=Lv by Mb   103=Lv by Ln 104=Lv by Lg 105=Lv by Lv  106=Lv by Ll 107=Lv by Mg 108=Lv by Mz
'109=Ll by Ka 110=Ll by Kb 111=Ll by La  112=Ll by Lb 113=Ll by Ma 114=Ll by Mb   115=Ll by Ln 116=Ll by Lg 117=Ll by Lv  118=Ll by Ll 119=Ll by Mg 120=Ll by Mz
'121=Mg by Ka 122=Mg by Kb 123=Mg by La  124=Mg by Lb 125=Mg by Ma 126=Mg by Mb   127=Mg by Ln 128=Mg by Lg 129=Mg by Lv  130=Mg by Ll 131=Mg by Mg 132=Mg by Mz
'133=Mz by Ka 134=Mz by Kb 135=Mz by La  136=Mz by Lb 137=Mz by Ma 138=Mz by Mb   139=Mz by Ln 140=Mz by Lg 141=Mz by Lv  142=Mz by Ll 143=Mz by Mg 144=Mz by Mz
If tFluTyp% = 1 Or tFluTyp% = 13 Or tFluTyp% = 25 Or tFluTyp% = 37 Or tFluTyp% = 49 Or tFluTyp% = 61 Then ZAFFLUGetFluYield! = zaf.flu!(1, iabs%)  ' fluorescence by Ka
If tFluTyp% = 73 Or tFluTyp% = 85 Or tFluTyp% = 97 Or tFluTyp% = 109 Or tFluTyp% = 121 Or tFluTyp% = 133 Then ZAFFLUGetFluYield! = zaf.flu!(1, iabs%)  ' fluorescence by Ka

If tFluTyp% = 2 Or tFluTyp% = 14 Or tFluTyp% = 26 Or tFluTyp% = 38 Or tFluTyp% = 50 Or tFluTyp% = 62 Then ZAFFLUGetFluYield! = zaf.flu!(2, iabs%)  ' fluorescence by Kb
If tFluTyp% = 74 Or tFluTyp% = 86 Or tFluTyp% = 98 Or tFluTyp% = 110 Or tFluTyp% = 122 Or tFluTyp% = 134 Then ZAFFLUGetFluYield! = zaf.flu!(2, iabs%)  ' fluorescence by Kb

If tFluTyp% = 3 Or tFluTyp% = 15 Or tFluTyp% = 27 Or tFluTyp% = 39 Or tFluTyp% = 51 Or tFluTyp% = 63 Then ZAFFLUGetFluYield! = zaf.flu!(3, iabs%)  ' fluorescence by La
If tFluTyp% = 75 Or tFluTyp% = 87 Or tFluTyp% = 99 Or tFluTyp% = 111 Or tFluTyp% = 123 Or tFluTyp% = 135 Then ZAFFLUGetFluYield! = zaf.flu!(3, iabs%)  ' fluorescence by La

If tFluTyp% = 4 Or tFluTyp% = 16 Or tFluTyp% = 28 Or tFluTyp% = 40 Or tFluTyp% = 52 Or tFluTyp% = 64 Then ZAFFLUGetFluYield! = zaf.flu!(4, iabs%)  ' fluorescence by Lb
If tFluTyp% = 76 Or tFluTyp% = 88 Or tFluTyp% = 100 Or tFluTyp% = 112 Or tFluTyp% = 124 Or tFluTyp% = 136 Then ZAFFLUGetFluYield! = zaf.flu!(4, iabs%)  ' fluorescence by Lb

If tFluTyp% = 5 Or tFluTyp% = 17 Or tFluTyp% = 29 Or tFluTyp% = 41 Or tFluTyp% = 53 Or tFluTyp% = 65 Then ZAFFLUGetFluYield! = zaf.flu!(5, iabs%)  ' fluorescence by Ma
If tFluTyp% = 77 Or tFluTyp% = 89 Or tFluTyp% = 101 Or tFluTyp% = 113 Or tFluTyp% = 125 Or tFluTyp% = 137 Then ZAFFLUGetFluYield! = zaf.flu!(5, iabs%)  ' fluorescence by Ma

If tFluTyp% = 6 Or tFluTyp% = 18 Or tFluTyp% = 30 Or tFluTyp% = 42 Or tFluTyp% = 54 Or tFluTyp% = 66 Then ZAFFLUGetFluYield! = zaf.flu!(6, iabs%)  ' fluorescence by Mb
If tFluTyp% = 78 Or tFluTyp% = 90 Or tFluTyp% = 102 Or tFluTyp% = 114 Or tFluTyp% = 126 Or tFluTyp% = 138 Then ZAFFLUGetFluYield! = zaf.flu!(6, iabs%)  ' fluorescence by Mb


If tFluTyp% = 7 Or tFluTyp% = 19 Or tFluTyp% = 31 Or tFluTyp% = 43 Or tFluTyp% = 55 Or tFluTyp% = 67 Then ZAFFLUGetFluYield! = zaf.flu!(7, iabs%)  ' fluorescence by Ln
If tFluTyp% = 79 Or tFluTyp% = 91 Or tFluTyp% = 103 Or tFluTyp% = 115 Or tFluTyp% = 127 Or tFluTyp% = 139 Then ZAFFLUGetFluYield! = zaf.flu!(7, iabs%)  ' fluorescence by Ln

If tFluTyp% = 8 Or tFluTyp% = 20 Or tFluTyp% = 32 Or tFluTyp% = 44 Or tFluTyp% = 56 Or tFluTyp% = 68 Then ZAFFLUGetFluYield! = zaf.flu!(8, iabs%)  ' fluorescence by Lg
If tFluTyp% = 80 Or tFluTyp% = 92 Or tFluTyp% = 104 Or tFluTyp% = 116 Or tFluTyp% = 128 Or tFluTyp% = 140 Then ZAFFLUGetFluYield! = zaf.flu!(8, iabs%)  ' fluorescence by Lg

If tFluTyp% = 9 Or tFluTyp% = 21 Or tFluTyp% = 33 Or tFluTyp% = 45 Or tFluTyp% = 57 Or tFluTyp% = 69 Then ZAFFLUGetFluYield! = zaf.flu!(9, iabs%)  ' fluorescence by Lv
If tFluTyp% = 81 Or tFluTyp% = 93 Or tFluTyp% = 105 Or tFluTyp% = 117 Or tFluTyp% = 129 Or tFluTyp% = 141 Then ZAFFLUGetFluYield! = zaf.flu!(9, iabs%)  ' fluorescence by Lv

If tFluTyp% = 10 Or tFluTyp% = 22 Or tFluTyp% = 34 Or tFluTyp% = 46 Or tFluTyp% = 58 Or tFluTyp% = 70 Then ZAFFLUGetFluYield! = zaf.flu!(10, iabs%)  ' fluorescence by Ll
If tFluTyp% = 82 Or tFluTyp% = 94 Or tFluTyp% = 106 Or tFluTyp% = 118 Or tFluTyp% = 130 Or tFluTyp% = 142 Then ZAFFLUGetFluYield! = zaf.flu!(10, iabs%)  ' fluorescence by Ll

If tFluTyp% = 11 Or tFluTyp% = 23 Or tFluTyp% = 35 Or tFluTyp% = 47 Or tFluTyp% = 59 Or tFluTyp% = 71 Then ZAFFLUGetFluYield! = zaf.flu!(11, iabs%)  ' fluorescence by Mg
If tFluTyp% = 83 Or tFluTyp% = 95 Or tFluTyp% = 107 Or tFluTyp% = 119 Or tFluTyp% = 131 Or tFluTyp% = 143 Then ZAFFLUGetFluYield! = zaf.flu!(11, iabs%)  ' fluorescence by Mg

If tFluTyp% = 12 Or tFluTyp% = 24 Or tFluTyp% = 36 Or tFluTyp% = 48 Or tFluTyp% = 60 Or tFluTyp% = 72 Then ZAFFLUGetFluYield! = zaf.flu!(12, iabs%)  ' fluorescence by Mz
If tFluTyp% = 84 Or tFluTyp% = 96 Or tFluTyp% = 108 Or tFluTyp% = 120 Or tFluTyp% = 132 Or tFluTyp% = 144 Then ZAFFLUGetFluYield! = zaf.flu!(12, iabs%)  ' fluorescence by Mz

Exit Function

' Errors
ZAFFLUGetFluYieldError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFFLUGetFluYield"
ierror = True
Exit Function

End Function

Sub ZAFFLULoadLineWeightsReed(i As Integer, i1 As Integer, i2 As Integer, fluor_type2() As Integer, rela_line_wts2() As Single)
' Load the relative line weights calculated from Reed (improved by tuning to Penepma large fluorescence database)
' Note: rela_line_wts2!(i%, i1%, i2%) = relative line weights (originally Pij) without Z dependency
' Variable fluor_type2%() is code for type of fluorescence:  0 = none

ierror = False
On Error GoTo ZAFFLULoadLineWeightsReedError
           
' Load relative line weights for original and additional x-ray lines here...
'  1=Ka by Ka   2=Ka by Kb   3=Ka by La    4=Ka by Lb   5=Ka by Ma   6=Ka by Mb     7=Ka by Ln   8=Ka by Lg   9=Ka by Lv   10=Ka by Ll  11=Ka by Mg  12=Ka by Mz
' 13=Kb by Ka  14=Kb by Kb  15=Kb by La   16=Kb by Lb  17=Kb by Ma  18=Kb by Mb    19=Kb by Ln  20=Kb by Lg  21=Kb by Lv   22=Kb by Ll  23=Kb by Mg  24=Kb by Mz
' 25=La by Ka  26=La by Kb  27=La by La   28=La by Lb  29=La by Ma  30=La by Mb    31=La by Ln  32=La by Lg  33=La by Lv   34=La by Ll  35=La by Mg  36=La by Mz
' 37=Lb by Ka  38=Lb by Kb  39=Lb by La   40=Lb by Lb  41=Lb by Ma  42=Lb by Mb    43=Lb by Ln  44=Lb by Lg  45=Lb by Lv   46=Lb by Ll  47=Lb by Mg  48=Lb by Mz
' 49=Ma by Ka  50=Ma by Kb  51=Ma by La   52=Ma by Lb  53=Ma by Ma  54=Ma by Mb    55=Ma by Ln  56=Ma by Lg  57=Ma by Lv   58=Ma by Ll  59=Ma by Mg  60=Ma by Mz
' 61=Mb by Ka  62=Mb by Kb  63=Mb by La   64=Mb by Lb  65=Mb by Ma  66=Mb by Mb    67=Mb by Ln  68=Mb by Lg  69=Mb by Lv   70=Mb by Ll  71=Mb by Mg  72=Mb by Mz

' 73=Ln by Ka  74=Ln by Kb  75=Ln by La   76=Ln by Lb  77=Ln by Ma  78=Ln by Mb    79=Ln by Ln  80=Ln by Lg  81=Ln by Lv   82=Ln by Ll  83=Ln by Mg  84=Ln by Mz
' 85=Lg by Ka  86=Lg by Kb  87=Lg by La   88=Lg by Lb  89=Lg by Ma  90=Lg by Mb    91=Lg by Ln  92=Lg by Lg  93=Lg by Lv   94=Lg by Ll  95=Lg by Mg  96=Lg by Mz
' 97=Lv by Ka  98=Lv by Kb  99=Lv by La  100=Lv by Lb 101=Lv by Ma 102=Lv by Mb   103=Lv by Ln 104=Lv by Lg 105=Lv by Lv  106=Lv by Ll 107=Lv by Mg 108=Lv by Mz
'109=Ll by Ka 110=Ll by Kb 111=Ll by La  112=Ll by Lb 113=Ll by Ma 114=Ll by Mb   115=Ll by Ln 116=Ll by Lg 117=Ll by Lv  118=Ll by Ll 119=Ll by Mg 120=Ll by Mz
'121=Mg by Ka 122=Mg by Kb 123=Mg by La  124=Mg by Lb 125=Mg by Ma 126=Mg by Mb   127=Mg by Ln 128=Mg by Lg 129=Mg by Lv  130=Mg by Ll 131=Mg by Mg 132=Mg by Mz
'133=Mz by Ka 134=Mz by Kb 135=Mz by La  136=Mz by Lb 137=Mz by Ma 138=Mz by Mb   139=Mz by Ln 140=Mz by Lg 141=Mz by Lv  142=Mz by Ll 143=Mz by Mg 144=Mz by Mz
If fluor_type2%(i%, i1%, i2%) = 1 Then rela_line_wts2!(i%, i1%, i2%) = 1#     ' Ka by Ka
If fluor_type2%(i%, i1%, i2%) = 2 Then rela_line_wts2!(i%, i1%, i2%) = 0.05   ' Ka by Kb (adjusted based on Penepma12_Exper_kratios_flu.dat)
If fluor_type2%(i%, i1%, i2%) = 3 Then rela_line_wts2!(i%, i1%, i2%) = 4.2    ' Ka by La (4.2 from Reed)
If fluor_type2%(i%, i1%, i2%) = 4 Then rela_line_wts2!(i%, i1%, i2%) = 0.2    ' Ka by Lb (adjusted based on Penepma12_Exper_kratios_flu.dat)
If fluor_type2%(i%, i1%, i2%) = 5 Then rela_line_wts2!(i%, i1%, i2%) = 1.6    ' Ka by Ma (adjusted based on Penepma12_Exper_kratios Si/Pt)
If fluor_type2%(i%, i1%, i2%) = 6 Then rela_line_wts2!(i%, i1%, i2%) = 1.6    ' Ka by Mb (adjusted based on Penepma12_Exper_kratios Si/Pt)


If fluor_type2%(i%, i1%, i2%) = 13 Then rela_line_wts2!(i%, i1%, i2%) = 0.6    ' Kb by Ka (adjusted based on Penepma12_Exper_kratios_flu.dat)
If fluor_type2%(i%, i1%, i2%) = 14 Then rela_line_wts2!(i%, i1%, i2%) = 0.08   ' Kb by Kb (adjusted based on Penepma12_Exper_kratios_flu.dat)
If fluor_type2%(i%, i1%, i2%) = 15 Then rela_line_wts2!(i%, i1%, i2%) = 0.3    ' Kb by La (adjusted based on Penepma12_Exper_kratios_flu.dat)
If fluor_type2%(i%, i1%, i2%) = 16 Then rela_line_wts2!(i%, i1%, i2%) = 2.4   ' Kb by Lb (adjusted based on Penepma12_Exper_kratios_flu.dat)
If fluor_type2%(i%, i1%, i2%) = 17 Then rela_line_wts2!(i%, i1%, i2%) = 0#    ' Kb by Ma ()
If fluor_type2%(i%, i1%, i2%) = 18 Then rela_line_wts2!(i%, i1%, i2%) = 0#    ' Kb by Mb ()


If fluor_type2%(i%, i1%, i2%) = 25 Then rela_line_wts2!(i%, i1%, i2%) = 0.24  ' La by Ka (0.24 from Reed)
If fluor_type2%(i%, i1%, i2%) = 26 Then rela_line_wts2!(i%, i1%, i2%) = 0.005  ' La by Kb (adjusted based on Penepma12_Exper_kratios Cd/Ca)
If fluor_type2%(i%, i1%, i2%) = 27 Then rela_line_wts2!(i%, i1%, i2%) = 1#    ' La by La
If fluor_type2%(i%, i1%, i2%) = 28 Then rela_line_wts2!(i%, i1%, i2%) = 0.5   ' La by Lb (adjusted based on Penepma12_Exper_kratios Pb/Th)
If fluor_type2%(i%, i1%, i2%) = 29 Then rela_line_wts2!(i%, i1%, i2%) = 0#    ' La by Ma (adjusted based on Penepma12_Exper_kratios Rb/Re)
If fluor_type2%(i%, i1%, i2%) = 30 Then rela_line_wts2!(i%, i1%, i2%) = 0.05  ' La by Mb (adjusted based on Pouchou2.dat)


If fluor_type2%(i%, i1%, i2%) = 37 Then rela_line_wts2!(i%, i1%, i2%) = 0.01  ' Lb by Ka (adjusted based on Penepma12_Exper_kratios Ag/Ca)
If fluor_type2%(i%, i1%, i2%) = 38 Then rela_line_wts2!(i%, i1%, i2%) = 0.01  ' Lb by Kb (adjusted based on Penepma12_Exper_kratios Ag/Ca)
If fluor_type2%(i%, i1%, i2%) = 39 Then rela_line_wts2!(i%, i1%, i2%) = 0#  ' Lb by La ()
If fluor_type2%(i%, i1%, i2%) = 40 Then rela_line_wts2!(i%, i1%, i2%) = 1#  ' Lb by Lb ()
If fluor_type2%(i%, i1%, i2%) = 41 Then rela_line_wts2!(i%, i1%, i2%) = 0#  ' Lb by Ma ()
If fluor_type2%(i%, i1%, i2%) = 42 Then rela_line_wts2!(i%, i1%, i2%) = 0#  ' Lb by Mb ()


If fluor_type2%(i%, i1%, i2%) = 49 Then rela_line_wts2!(i%, i1%, i2%) = 0.01  ' Ma by Ka (adjusted based on Pouchou2.dat)
If fluor_type2%(i%, i1%, i2%) = 50 Then rela_line_wts2!(i%, i1%, i2%) = 0#    ' Ma by Kb (adjusted based on Penepma12_Exper_kratios U/K)
If fluor_type2%(i%, i1%, i2%) = 51 Then rela_line_wts2!(i%, i1%, i2%) = 0.02  ' Ma by La (adjusted based on Penepma12_Exper_kratios Pb/Rh)
If fluor_type2%(i%, i1%, i2%) = 52 Then rela_line_wts2!(i%, i1%, i2%) = 0.02  ' Ma by Lb (adjusted based on Pouchou2.dat)
If fluor_type2%(i%, i1%, i2%) = 53 Then rela_line_wts2!(i%, i1%, i2%) = 1#    ' Ma by Ma
If fluor_type2%(i%, i1%, i2%) = 54 Then rela_line_wts2!(i%, i1%, i2%) = 0.03  ' Ma by Mb (adjusted based on Pouchou2.dat)


If fluor_type2%(i%, i1%, i2%) = 61 Then rela_line_wts2!(i%, i1%, i2%) = 0.01  ' Mb by Ka ()
If fluor_type2%(i%, i1%, i2%) = 62 Then rela_line_wts2!(i%, i1%, i2%) = 0.4   ' Mb by Kb (adjusted based on Pouchou2.dat)
If fluor_type2%(i%, i1%, i2%) = 63 Then rela_line_wts2!(i%, i1%, i2%) = 0#   ' Mb by La ()
If fluor_type2%(i%, i1%, i2%) = 64 Then rela_line_wts2!(i%, i1%, i2%) = 0#   ' Mb by Lb ()
If fluor_type2%(i%, i1%, i2%) = 65 Then rela_line_wts2!(i%, i1%, i2%) = 2#    ' Mb by Ma (adjusted based on Penepma12_Exper_kratios Pb/Th)
If fluor_type2%(i%, i1%, i2%) = 66 Then rela_line_wts2!(i%, i1%, i2%) = 1#    ' Mb by Mb (adjusted based on Penepma12_Exper_kratios Pb/Th)

Exit Sub

' Errors
ZAFFLULoadLineWeightsReedError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFFLULoadLineWeightsReed"
ierror = True
Exit Sub

End Sub

Sub ZAFFLULoadLineWeightsPenepma(iemitter As Integer, imatrix As Integer, iemitter_takeoff As Single, iemitter_keV As Single, iemitter_elem As Integer, iemitter_xray As Integer, imatrix_takeoff As Single, imatrix_keV As Single, imatrix_elem As Integer, imatrix_xray As Integer, tLineWeightRatio As Single, flu_yield_emitter As Single, flu_yield_matrix As Single, zaf As TypeZAF)
' Load the relative line weights calculated from the ratio of the pure element generated intensities from Penfluor/Fanal
' Relative emission intensity based on the emitter element, emitter x-ray, matrix element causing fluorescence, and matrix element x-ray and takeoff and keV
'  iemitter% = array index for emitting element (fluoresced)
'  imatrix% = array index for matrix element (fluorescing)
'  iemitter_elem% = atomic number for emitting element (fluoresced)
'  imatrix_elem% = atomic number for matrix element (fluorescing)

ierror = False
On Error GoTo ZAFFLULoadLineWeightsPenepmaError

Dim notfound1 As Boolean, notfound2 As Boolean
Dim iray As Integer

Dim notfoundA As Boolean, notfoundB As Boolean

Dim emitter_generated_single As Double, emitter_emitted_single As Double
Dim matrix_generated_single As Double, matrix_emitted_single As Double

Dim emitter_corrected As Double, matrix_corrected As Double

Dim emitter_generated As Double, emitter_emitted As Double
Dim matrix_generated As Double, matrix_emitted As Double

Dim emitter_generated_sum As Double, matrix_generated_sum As Double

' First print default relative line weight
If VerboseMode Then
Call IOWriteLog(vbNullString)
Call IOWriteLog(vbNullString)
tmsg$ = "Default relative line weight for emitter " & Trim$(Symup$(iemitter_elem%)) & " " & Trim$(Xraylo$(iemitter_xray%)) & " by " & Trim$(Symup$(imatrix_elem%)) & " " & Trim$(Xraylo$(imatrix_xray%)) & " is " & MiscAutoFormat$(tLineWeightRatio!)
Call IOWriteLog(tmsg$)
End If
           
' Get pure element intensities for emitter element (the fluoresced x-ray line)
Call Penepma12PureReadMDB2(iemitter_takeoff!, iemitter_keV!, iemitter_elem%, iemitter_xray%, emitter_generated_single#, emitter_emitted_single#, notfound1)
If ierror Then Exit Sub

' Get pure element intensities for matrix element (the fluorescing x-ray line)
Call Penepma12PureReadMDB2(imatrix_takeoff!, imatrix_keV!, imatrix_elem%, imatrix_xray%, matrix_generated_single#, matrix_emitted_single#, notfound2)
If ierror Then Exit Sub

' Check if values were found
If notfound1 Then Exit Sub
If notfound2 Then Exit Sub

If VerboseMode Then
tmsg$ = "Raw (generated) intensity for emitter " & Trim$(Symup$(iemitter_elem%)) & " " & Trim$(Xraylo$(iemitter_xray%)) & " : " & MiscAutoFormat$(CSng(emitter_generated_single))
Call IOWriteLog(tmsg$)
tmsg$ = "Raw (generated) intensity for matrix " & Trim$(Symup$(imatrix_elem%)) & " " & Trim$(Xraylo$(imatrix_xray%)) & " : " & MiscAutoFormat$(CSng(matrix_generated_single))
Call IOWriteLog(tmsg$)
tmsg$ = "Uncorrected relative line weight for emitter " & Trim$(Symup$(iemitter_elem%)) & " " & Trim$(Xraylo$(iemitter_xray%)) & " by " & Trim$(Symup$(imatrix_elem%)) & " " & Trim$(Xraylo$(imatrix_xray%)) & " is " & MiscAutoFormat$(emitter_generated_single# / matrix_generated_single#)
Call IOWriteLog(tmsg$)
End If

' Get intensities for all x-ray lines for this family for the emitter element
If VerboseMode Then Call IOWriteLog(vbNullString)
For iray% = 1 To MAXRAY% - 1
If ZAFFLUCheckIfSameFamily(iray%, iemitter_xray%) Then
Call Penepma12PureReadMDB2(iemitter_takeoff!, iemitter_keV!, iemitter_elem%, iray%, emitter_generated#, emitter_emitted#, notfoundA)
If ierror Then Exit Sub

' Sum the generated intensities for this emitter family
If Not notfoundA Then
emitter_generated_sum# = emitter_generated_sum# + emitter_generated#
If VerboseMode Then
tmsg$ = "Emitter family intensity for " & Trim$(Symup$(iemitter_elem%)) & " " & Trim$(Xraylo$(iray%)) & ": " & MiscAutoFormat$(CSng(emitter_generated#))
Call IOWriteLog(tmsg$)
End If
End If
End If
Next iray%

' Get intensities for all x-ray lines for this family for the matrix element
For iray% = 1 To MAXRAY% - 1
If ZAFFLUCheckIfSameFamily(iray%, imatrix_xray%) Then
Call Penepma12PureReadMDB2(imatrix_takeoff!, imatrix_keV!, imatrix_elem%, iray%, matrix_generated#, matrix_emitted#, notfoundB)
If ierror Then Exit Sub

' Sum the generated intensities for this matrix family
If Not notfoundB Then
matrix_generated_sum# = matrix_generated_sum# + matrix_generated#
If VerboseMode Then
tmsg$ = "Matrix family intensity for " & Trim$(Symup$(imatrix_elem%)) & " " & Trim$(Xraylo$(iray%)) & ": " & MiscAutoFormat$(CSng(matrix_generated#))
Call IOWriteLog(tmsg$)
End If
End If
End If
Next iray%

' Print emission and matrix line and family intensities (only used for original Reed K by L, L by K relative line weight calculations)
If VerboseMode Then
tmsg$ = "Emitter family (sum) intensity for all " & Trim$(Left$(Xraylo$(iemitter_xray%), 1)) & " lines: " & MiscAutoFormat$(CSng(emitter_generated_sum#))
Call IOWriteLog(tmsg$)
tmsg$ = "Matrix family (sum) intensity for all " & Trim$(Left$(Xraylo$(imatrix_xray%), 1)) & " lines: " & MiscAutoFormat$(CSng(matrix_generated_sum#))
Call IOWriteLog(tmsg$)
tmsg$ = "Uncorrected relative (family) line weight for emitter " & Trim$(Symup$(iemitter_elem%)) & " " & Trim$(Xraylo$(iemitter_xray%)) & " by " & Trim$(Symup$(imatrix_elem%)) & " " & Trim$(Xraylo$(imatrix_xray%)) & " is " & MiscAutoFormat$(emitter_generated_sum# / matrix_generated_sum#)
Call IOWriteLog(tmsg$)
End If

' Now calculate improved Reed relative line weights for individual lines
If VerboseMode Then
Call IOWriteLog(vbNullString)
tmsg$ = "Corrections for emitter " & Symup$(iemitter_elem%) & " " & Xraylo$(iemitter_xray%) & ", A: " & MiscAutoFormat$(zaf.atwts!(iemitter%)) & ", U: " & MiscAutoFormat$(zaf.v!(iemitter%))
Call IOWriteLog(tmsg$)
tmsg$ = "Corrections for matrix " & Symup$(imatrix_elem%) & " " & Xraylo$(imatrix_xray%) & ", A: " & MiscAutoFormat$(zaf.atwts!(imatrix%)) & ", U: " & MiscAutoFormat$(zaf.v!(imatrix%))
Call IOWriteLog(tmsg$)
End If

' Correct for individual lines for matrix effects (Reed and Goemann)
'emitter_corrected# = emitter_generated_single# * zaf.atwts!(iemitter%) * zaf.s!(imatrix%, iemitter%) / ((zaf.v!(iemitter%) - 1#) ^ 1.67 * flu_yield_emitter! * zaf.r!(imatrix%, iemitter%))
'matrix_corrected# = matrix_generated_single# * zaf.atwts!(imatrix%) * zaf.s!(iemitter%, imatrix%) / ((zaf.v!(imatrix%) - 1#) ^ 1.67 * flu_yield_matrix! * zaf.r!(iemitter%, imatrix%))
emitter_corrected# = emitter_generated_single# * zaf.atwts!(iemitter%) / ((zaf.v!(iemitter%) - 1#) ^ 1.67 * flu_yield_emitter!)
matrix_corrected# = matrix_generated_single# * zaf.atwts!(imatrix%) / ((zaf.v!(imatrix%) - 1#) ^ 1.67 * flu_yield_matrix!)

' Check and debug output
If VerboseMode Then
tmsg$ = "Corrected intensity for emitter " & Trim$(Symup$(iemitter_elem%)) & " " & Trim$(Xraylo$(iemitter_xray%)) & " is " & MiscAutoFormat$(CSng(emitter_corrected#))
Call IOWriteLog(tmsg$)
tmsg$ = "Corrected intensity for matrix " & Trim$(Symup$(imatrix_elem%)) & " " & Trim$(Xraylo$(imatrix_xray%)) & " is " & MiscAutoFormat$(CSng(matrix_corrected#))
Call IOWriteLog(tmsg$)
End If

' Calculate relative line weight for this emitter-matrix pair (Reed uses emitter/fluorescer for relative line weight ratios, that is Fe Ka / W La)
tLineWeightRatio! = CSng(emitter_corrected# / matrix_corrected#)

If VerboseMode Then
tmsg$ = "Corrected relative line weight for emitter " & Trim$(Symup$(iemitter_elem%)) & " " & Trim$(Xraylo$(iemitter_xray%)) & " by " & Trim$(Symup$(imatrix_elem%)) & " " & Trim$(Xraylo$(imatrix_xray%)) & " is " & MiscAutoFormat$(tLineWeightRatio!)
Call IOWriteLog(tmsg$)
End If

Exit Sub

' Errors
ZAFFLULoadLineWeightsPenepmaError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFFLULoadLineWeightsPenepma"
ierror = True
Exit Sub

End Sub

Function ZAFFLUCheckIfSameFamily(iray As Integer, pray As Integer) As Boolean
' Check if the next x-ray line (iray%) is in the same family as the line in question (pray%).
'  Xraylo$(1) = "ka"
'  Xraylo$(2) = "kb"
'  Xraylo$(3) = "la"
'  Xraylo$(4) = "lb"
'  Xraylo$(5) = "ma"
'  Xraylo$(6) = "mb"

'  Xraylo$(7) = "Ln"
'  Xraylo$(8) = "Lg"
'  Xraylo$(9) = "Lv"
'  Xraylo$(10) = "Ll"
'  Xraylo$(11) = "Mg"
'  Xraylo$(12) = "Mz"

ierror = False
On Error GoTo ZAFFLUCheckIfSameFamilyError

ZAFFLUCheckIfSameFamily = False

' Check for K family
If (iray% = 1 Or iray% = 2) And (pray% = 1 Or pray% = 2) Then ZAFFLUCheckIfSameFamily = True

' Check for L family
If (iray% = 3 Or iray% = 4 Or iray% = 7 Or iray% = 8 Or iray% = 9 Or iray% = 10) And (pray% = 3 Or pray% = 4 Or pray% = 7 Or pray% = 8 Or pray% = 9 Or pray% = 10) Then ZAFFLUCheckIfSameFamily = True

' Check for M family
If (iray% = 5 Or iray% = 6 Or iray% = 11 Or iray% = 12) And (pray% = 5 Or pray% = 6 Or iray% = 11 Or iray% = 12) Then ZAFFLUCheckIfSameFamily = True

Exit Function

' Errors
ZAFFLUCheckIfSameFamilyError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFFLUCheckIfSameFamily"
ierror = True
Exit Function

End Function
