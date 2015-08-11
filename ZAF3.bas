Attribute VB_Name = "CodeZAF3"
' (c) Copyright 1995-2015 by John J. Donovan (credit to John Armstrong for original code)
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

Dim fluor_MACs(1 To MAXRAY% - 1, 1 To MAXCHAN1%, 1 To MAXCHAN1%) As Single      ' MACs for fluorescencing lines (formally fp)

Dim AllJumpRatios(1 To MAXEDG%, 1 To MAXELM%) As Single                                                 ' values from NIST FFAST table
Dim AllTransitionProbabilities(1 To MAXELM%, 1 To MAXSHELL%, 1 To MAXSHELL%) As Single                  ' values from Penepma table

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
msg$ = "WARNING in ZAFFlu1- " & Format$(Symup$(Int(zaf.Z!(i1%))), a20$) & " " & Format$(Edglo$(in8%), a20$) & " absorption edge is zero"
Call IOWriteLog(msg$)
GoTo 6960
End If

' Check if emission energy is greater than upper limit of fluorescence or edge energy is greater than electron energy (zaf.eO!(i%))
If zaf.eng!(i2%, i1%) > (zaf.eC!(i%) + n6!) Or zaf.edg!(in8%, i1%) >= zaf.eO!(i%) Then GoTo 6960

' Set fluorescence flag for type of fluorescence
fluor_type%(i%, i1%) = i2% + (MAXRAY% - 1) * (zaf.il%(i%) - 1)

' Warn if line is fluoresced
If VerboseMode Then
msg$ = "WARNING in ZAFFlu1- the " & Format$(Xraylo$(im4%), a20$) & " line of " & Format$(Symup$(Int(zaf.Z!(i%))), a20$) & " is excited by the " & Format$(Xraylo$(i2%), a20$) & " line of " & Format$(Symup$(Int(zaf.Z!(i1%))), a20$)
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
msg$ = "WARNING in ZAFFlu1- MAC not loaded for " & Format$(Symup$(Int(zaf.Z!(i1%))), a20$) & " " & Format$(Xraylo$(i2%), a20$) & " in " & Format$(Symup$(Int(zaf.Z!(i4%))), a20$) & ", fluorescence contribution will not be calculated for this line."
Call IOWriteLog(msg$)
End If
fluor_MACs!(i2%, i4%, i1%) = 1#
fluor_type%(i%, i1%) = 0
End If

Next i4%
6960:  Next i2%     ' for each absorber (matrix) x-ray causing fluorescence
6970:  Next i1%     ' for each absorber (matrix) element causing fluorescence

' Load jump ratios
If zaf.il%(i%) = 1 Then jump_ratios!(i%) = 1.11728 - 0.07368 * Log(zaf.Z!(i%))     ' Ka line (0.88)
If zaf.il%(i%) = 3 Then jump_ratios!(i%) = 0.95478 - 0.00259 * zaf.Z!(i%)          ' La line (0.75)
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
msg$ = Format$(Symup$(Int(zaf.Z!(i%))), a20$) & " " & Format$(Xraylo$(zaf.il(i%)), a20$) & " by " & Format$(Symup$(Int(zaf.Z!(i1%))), a20$)
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
msg$ = "Average MAC for " & Format$(Symup$(Int(zaf.Z!(i%))), a20$) & " in this matrix is negative for line " & Str$(zaf.n8) & ", and is probably a bad data point (epoxy, etc.). Delete the analysis line and try again."
If Not CalcImageQuantFlag Then
MiscMsgBoxTim FormMSGBOXTIME, "ZAFFlu1", msg$, 20#
Call IOWriteLog(msg$)
Else
Call IOWriteLog(msg$)
End If
'ierror = True
Exit Sub

ZAFFlu1NegativeFlu:
msg$ = "Fluorescence factor for " & Format$(Symup$(Int(zaf.Z!(i%))), a20$) & " in this matrix is negative for line " & Str$(zaf.n8) & ", and is probably a bad data point (epoxy, etc.). Delete the analysis line and try again."
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
' Note that "MAXRAY%" equals 7 so "MAXRAY% - 1" equals 6 (Ka,Kb,La,Lb,Ma,Mb)

ierror = False
On Error GoTo ZAFFlu2Error

Dim temp As Single, mac As Single
Dim in8 As Integer, i As Integer, i1 As Integer, i2 As Integer, i4 As Integer

Dim n2 As Single, n3 As Single, n4 As Single
Dim n5 As Single, n6 As Single, m8 As Single
Dim FLUD As Single, FLUX As Single, FLUA As Single, FLUB As Single, FLUC As Single

ReDim m7(1 To MAXCHAN1%) As Single

' If sample calculation, skip fluorescence initialization
If zafinit% = 1 Then GoTo 7200

' Init variables
For i% = 1 To zaf.in1%  ' for each emitter element
    For i1% = 1 To zaf.in0% ' for each absorber (matrix) element
        For i2% = 1 To MAXRAY% - 1 ' for each absorber (matrix) x-ray
        rela_line_wts2!(i%, i1%, i2%) = 0#
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

    ' Check to see if any absorber (matrix) element present can fluoresce the emitted element/line
    For i1% = 1 To zaf.in0%     ' for each absorber (matrix) element causing fluorescence
        If Not UseFluorescenceByBetaLinesFlag Then
        If zaf.il%(i1%) = 2 Or zaf.il%(i1%) = 4 Or zaf.il%(i1%) = 6 Then GoTo 6970  ' skip fluorescence by beta lines
        End If

        ' Variable fluor_type2%() is code for type of fluorescence:  0 = none
        ' 1=Ka by Ka  2=Ka by Kb  3=Ka by La  4=Ka by Lb  5=Ka by Ma  6=Ka by Mb
        ' 7=Kb by Ka  8=Kb by Kb  9=Kb by La 10=Kb by Lb 11=Kb by Ma 12=Kb by Mb
        '13=La by Ka 14=La by Kb 15=La by La 16=La by Lb 17=La by Ma 18=La by Mb
        '19=Lb by Ka 20=Lb by Kb 21=Lb by La 22=Lb by Lb 23=Lb by Ma 24=Lb by Mb
        '25=Ma by Ka 26=Ma by Kb 27=Ma by La 28=Ma by Lb 29=Ma by Ma 30=Ma by Mb
        '31=Mb by Ka 32=Mb by Kb 33=Mb by La 34=Mb by Lb 35=Mb by Ma 36=Mb by Mb

            ' First test for fluorescence by K line, then L, and then M (if indicated).
            ' Note that "eng!(1 to MAXRAY%-1, 1 to MAXCHAN%)" is the emission line energies for all lines of
            ' an element and "ec!(1 to MAXCHAN%)" is the analyzed line absorption edges. While
            ' "edg!(1 to MAXEDG%, 1 to MAXCHAN%)" is the absorption edge energies for all lines of an element.
            For i2% = 1 To MAXRAY% - 1  ' for each x-ray for each absorber (matrix) element that might cause fluorescence in the emitter element/line
            fluor_type2%(i%, i1%, i2%) = 0
            If iflu% = 2 And (i2% = 5 Or i2% = 6) Then GoTo 6960    ' only do fluorescence of K and L lines if indicated
                If Not UseFluorescenceByBetaLinesFlag Then
                If i2% = 2 Or i2% = 4 Or i2% = 6 Then GoTo 6960     ' skip fluorescence by beta lines
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

            ' Check for missing absorption edge data
            If zaf.edg!(in8%, i1%) = 0# Then
            msg$ = "WARNING in ZAFFlu2- " & Format$(Symup$(Int(zaf.Z!(i1%))), a20$) & " " & Format$(Edglo$(in8%), a20$) & " absorption edge is zero"
            Call IOWriteLog(msg$)
            GoTo 6960
            End If

            ' Check if emission energy is greater than upper limit of fluorescence or edge energy is greater than electron energy (zaf.eO!(i%))
            If zaf.eng!(i2%, i1%) > (zaf.eC!(i%) + n6!) Or zaf.edg!(in8%, i1%) >= zaf.eO!(i%) Then GoTo 6960

            ' Set fluorescence flag for type of fluorescence
            fluor_type2%(i%, i1%, i2%) = i2% + (MAXRAY% - 1) * (zaf.il%(i%) - 1)

            ' Warn if line is fluoresced
            If VerboseMode Then
            msg$ = "WARNING in ZAFFlu2- the " & Format$(Xraylo$(zaf.il%(i%)), a20$) & " line of " & Format$(Symup$(Int(zaf.Z!(i%))), a20$) & " is excited by the " & Format$(Xraylo$(i2%), a20$) & " line of " & Format$(Symup$(Int(zaf.Z!(i1%))), a20$)
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
                    msg$ = "WARNING in ZAFFlu2- MAC not loaded for " & Format$(Symup$(Int(zaf.Z!(i1%))), a20$) & " " & Format$(Xraylo$(i2%), a20$) & " in " & Format$(Symup$(Int(zaf.Z!(i4%))), a20$) & ", fluorescence contribution will not be calculated for this line."
                    Call IOWriteLog(msg$)
                    End If
            
                fluor_MACs!(i2%, i4%, i1%) = 1#
                fluor_type2%(i%, i1%, i2%) = 0
                End If
                Next i4%
6960:       Next i2%        ' next absorber (matrix) x-ray causing fluorescence
6970:   Next i1%        ' next absorber (matrix) element causing fluorescence

    ' Load from Jump_ratios.dat (from Nicholas Ritchie, NIST)
    Call ZAFFLULoadJumpRatios(zaf.il%(i%), CInt(zaf.Z!(i%)), jump_ratios!(i%))
    If ierror Then Exit Sub

        ' Load fluorescent yields using pointer from above
        For i1% = 1 To zaf.in0%         ' for each absorber (matrix) element causing fluorescence
            For i2% = 1 To MAXRAY% - 1      ' for each absorber (matrix) x-ray causing fluorescence
            If fluor_type2%(i%, i1%, i2%) > 0 Then    ' skip if no fluorescence (zero)

            ' Variable fluor_yield2!(i%, i1%, i2%) is fluorescent yield of emitting element i% by absorbing (matrix) element i1%, line i2%
            fluor_yield2!(i%, i1%, i2%) = ZAFFLUGetFluYield(fluor_type2%(i%, i1%, i2%), i1%, zaf)
        
            ' Test code for loading z correlated relative line weights to improve accuracy
            If Dir$(ProgramPath$ & "penelope_mod_transition_data.csv") <> vbNullString Then
            'Call ZAFFLULoadLineWeights(Int(zaf.z!(i%)), Int(zaf.z!(i1%)), fluor_type2%(i%, i1%, i2%), rela_line_wts2!(i%, i1%, i2%))
            'If ierror Then Exit Sub
            End If

            ' Note: rela_line_wts2!(i%, i1%, i2%) = relative line weights (originally Pij)
            If fluor_type2%(i%, i1%, i2%) = 1 Then rela_line_wts2!(i%, i1%, i2%) = 1#     ' Ka by Ka
            If fluor_type2%(i%, i1%, i2%) = 2 Then rela_line_wts2!(i%, i1%, i2%) = 0.1    ' Ka by Kb (adjusted based on Penepma12_Exper_kratios_flu.dat)
            If fluor_type2%(i%, i1%, i2%) = 3 Then rela_line_wts2!(i%, i1%, i2%) = 4.2    ' Ka by La (Reed)
            If fluor_type2%(i%, i1%, i2%) = 4 Then rela_line_wts2!(i%, i1%, i2%) = 0.1    ' Ka by Lb (adjusted based on Penepma12_Exper_kratios_flu.dat)
            If fluor_type2%(i%, i1%, i2%) = 7 Then rela_line_wts2!(i%, i1%, i2%) = 0.7    ' Kb by Ka (adjusted based on Penepma12_Exper_kratios_flu.dat)
            If fluor_type2%(i%, i1%, i2%) = 8 Then rela_line_wts2!(i%, i1%, i2%) = 0.05   ' Kb by Kb (adjusted based on Pouchou2.dat)
            If fluor_type2%(i%, i1%, i2%) = 9 Then rela_line_wts2!(i%, i1%, i2%) = 0.2    ' Kb by La (adjusted based on Penepma12_Exper_kratios_flu.dat)
            If fluor_type2%(i%, i1%, i2%) = 10 Then rela_line_wts2!(i%, i1%, i2%) = 2.4   ' Kb by Lb (adjusted based on Penepma12_Exper_kratios_flu.dat)
            If fluor_type2%(i%, i1%, i2%) = 13 Then rela_line_wts2!(i%, i1%, i2%) = 0.24  ' La by Ka (Reed)
            If fluor_type2%(i%, i1%, i2%) = 14 Then rela_line_wts2!(i%, i1%, i2%) = 0.03  ' La by Kb (adjusted based on Penepma12_Exper_kratios_flu.dat)
            If fluor_type2%(i%, i1%, i2%) = 15 Then rela_line_wts2!(i%, i1%, i2%) = 1#    ' La by La
            If fluor_type2%(i%, i1%, i2%) = 16 Then rela_line_wts2!(i%, i1%, i2%) = 0.01  ' La by Lb (adjusted based on Penepma12_Exper_kratios_flu.dat)
            If fluor_type2%(i%, i1%, i2%) = 18 Then rela_line_wts2!(i%, i1%, i2%) = 0.05  ' La by Mb (adjusted based on Pouchou2.dat)
            If fluor_type2%(i%, i1%, i2%) = 25 Then rela_line_wts2!(i%, i1%, i2%) = 0.01  ' Ma by Ka (adjusted based on Pouchou2.dat)
            If fluor_type2%(i%, i1%, i2%) = 27 Then rela_line_wts2!(i%, i1%, i2%) = 0.02  ' Ma by La (adjusted based on Penepma12_Exper_kratios_flu.dat)
            If fluor_type2%(i%, i1%, i2%) = 28 Then rela_line_wts2!(i%, i1%, i2%) = 0.02  ' Ma by Lb (adjusted based on Pouchou2.dat)
            If fluor_type2%(i%, i1%, i2%) = 29 Then rela_line_wts2!(i%, i1%, i2%) = 1#    ' Ma by Ma
            If fluor_type2%(i%, i1%, i2%) = 30 Then rela_line_wts2!(i%, i1%, i2%) = 0.03  ' Ma by Mb (adjusted based on Pouchou2.dat)
            If fluor_type2%(i%, i1%, i2%) = 32 Then rela_line_wts2!(i%, i1%, i2%) = 0.4   ' Mb by Kb (adjusted based on Pouchou2.dat)
    
            If fluor_type2%(i%, i1%, i2%) = 5 Then     ' Ka by Ma: nothing in Pouchou2.dat or Penepma12_Exper_kratios_flu.dat (try Si/Pt)
            DoEvents
            End If
            If fluor_type2%(i%, i1%, i2%) = 6 Then     ' Ka by Mb: nothing in Pouchou2.dat or Penepma12_Exper_kratios_flu.dat (try Si/Pt)
            DoEvents
            End If
            If fluor_type2%(i%, i1%, i2%) = 11 Then    ' Kb by Ma: nothing in Pouchou2.dat or Penepma12_Exper_kratios_flu.dat
            DoEvents
            End If
            If fluor_type2%(i%, i1%, i2%) = 12 Then    ' Kb by Mb: nothing in Pouchou2.dat or Penepma12_Exper_kratios_flu.dat
            DoEvents
            End If
            If fluor_type2%(i%, i1%, i2%) = 17 Then    ' La by Ma: nothing in Pouchou2.dat or Penepma12_Exper_kratios_flu.dat
            DoEvents
            End If
            If fluor_type2%(i%, i1%, i2%) = 26 Then    ' Ma by Kb: nothing in Pouchou2.dat or Penepma12_Exper_kratios_flu.dat
            DoEvents
            End If
            If fluor_type2%(i%, i1%, i2%) = 31 Then    ' Mb by Ka: nothing in Pouchou2.dat or Penepma12_Exper_kratios_flu.dat
            DoEvents
            End If
            If fluor_type2%(i%, i1%, i2%) = 33 Then    ' Mb by La: nothing in Pouchou2.dat or Penepma12_Exper_kratios_flu.dat
            DoEvents
            End If
            If fluor_type2%(i%, i1%, i2%) = 34 Then    ' Mb by Lb: nothing in Pouchou2.dat or Penepma12_Exper_kratios_flu.dat
            DoEvents
            End If
            If fluor_type2%(i%, i1%, i2%) = 35 Then    ' Mb by Ma: nothing in Pouchou2.dat or Penepma12_Exper_kratios_flu.dat
            DoEvents
            End If
            If fluor_type2%(i%, i1%, i2%) = 36 Then    ' Mb by Mb: nothing in Pouchou2.dat or Penepma12_Exper_kratios_flu.dat
            DoEvents
            End If
            
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
        End If

        m8! = 0#
        i2% = fluor_type2%(i%, i1%, i2%) - (MAXRAY% - 1) * (zaf.il%(i%) - 1)  ' get index for MAC of fluorescencing line

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
        msg$ = Format$(Symup$(Int(zaf.Z!(i%))), a20$) & " " & Format$(Xraylo$(zaf.il(i%)), a20$) & " by " & Format$(Symup$(Int(zaf.Z!(i1%))), a20$) & " " & Format$(Xraylo$(i2%), a20$)
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
msg$ = "Average MAC for " & Format$(Symup$(Int(zaf.Z!(i%))), a20$) & " in this matrix is negative for line " & Str$(zaf.n8) & ", and is probably a bad data point (epoxy, etc.). Delete the analysis line and try again."
If Not CalcImageQuantFlag Then
MiscMsgBoxTim FormMSGBOXTIME, "ZAFFlu2", msg$, 20#
Call IOWriteLog(msg$)
Else
Call IOWriteLog(msg$)
End If
'ierror = True
Exit Sub

ZAFFlu2NegativeFlu:
msg$ = "Fluorescence factor for " & Format$(Symup$(Int(zaf.Z!(i%))), a20$) & " in this matrix is negative for line " & Str$(zaf.n8) & ", and is probably a bad data point (epoxy, etc.). Delete the analysis line and try again."
If Not CalcImageQuantFlag Then
MiscMsgBoxTim FormMSGBOXTIME, "ZAFFlu2", msg$, 20#
Call IOWriteLog(msg$)
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

tfilename$ = ProgramPath$ & "Jump_Ratios.dat"
Open tfilename$ For Input As #Temp1FileNumber%

' Read first line of column headings
Line Input #Temp1FileNumber%, astring   ' read comment
Line Input #Temp1FileNumber%, astring   ' read comment
Line Input #Temp1FileNumber%, astring   ' read column labels
If VerboseMode Then Call IOWriteLog(astring$)

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
If VerboseMode Then Call IOWriteLog(tmsg$)

' Load data value (ri - 1)/ri. That is, convert from jump "factor" to jump "ratio"
For i% = 1 To MAXEDG%
If tvalues!(i%) > 0# Then AllJumpRatios!(i%, iz%) = (tvalues!(i%) - 1) / tvalues!(i%)
Next i%
Loop

Close #Temp1FileNumber%
initialized = True
End If

' Calculate edge index for this x-ray
If iray% = 1 Then in8% = 1   ' Ka
If iray% = 2 Then in8% = 1   ' Kb
If iray% = 3 Then in8% = 4   ' La
If iray% = 4 Then in8% = 3   ' Lb
If iray% = 5 Then in8% = 9   ' Ma
If iray% = 6 Then in8% = 8   ' Mb

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
' 1=Ka by Ka  2=Ka by Kb  3=Ka by La  4=Ka by Lb  5=Ka by Ma  6=Ka by Mb
' 7=Kb by Ka  8=Kb by Kb  9=Kb by La 10=Kb by Lb 11=Kb by Ma 12=Kb by Mb
'13=La by Ka 14=La by Kb 15=La by La 16=La by Lb 17=La by Ma 18=La by Mb
'19=Lb by Ka 20=Lb by Kb 21=Lb by La 22=Lb by Lb 23=Lb by Ma 24=Lb by Mb
'25=Ma by Ka 26=Ma by Kb 27=Ma by La 28=Ma by Lb 29=Ma by Ma 30=Ma by Mb
'31=Mb by Ka 32=Mb by Kb 33=Mb by La 34=Mb by Lb 35=Mb by Ma 36=Mb by Mb

ierror = False
On Error GoTo ZAFFLUGetFluYieldError

ZAFFLUGetFluYield! = 0#
If tFluTyp% = 1 Or tFluTyp% = 7 Or tFluTyp% = 13 Or tFluTyp% = 19 Or tFluTyp% = 25 Or tFluTyp% = 31 Then ZAFFLUGetFluYield! = zaf.flu!(1, iabs%)  ' fluorescence by Ka
If tFluTyp% = 2 Or tFluTyp% = 8 Or tFluTyp% = 14 Or tFluTyp% = 20 Or tFluTyp% = 26 Or tFluTyp% = 32 Then ZAFFLUGetFluYield! = zaf.flu!(2, iabs%)  ' fluorescence by Kb
If tFluTyp% = 3 Or tFluTyp% = 9 Or tFluTyp% = 15 Or tFluTyp% = 21 Or tFluTyp% = 27 Or tFluTyp% = 33 Then ZAFFLUGetFluYield! = zaf.flu!(3, iabs%)  ' fluorescence by La
If tFluTyp% = 4 Or tFluTyp% = 10 Or tFluTyp% = 16 Or tFluTyp% = 22 Or tFluTyp% = 28 Or tFluTyp% = 34 Then ZAFFLUGetFluYield! = zaf.flu!(4, iabs%)  ' fluorescence by Lb
If tFluTyp% = 5 Or tFluTyp% = 11 Or tFluTyp% = 17 Or tFluTyp% = 23 Or tFluTyp% = 29 Or tFluTyp% = 35 Then ZAFFLUGetFluYield! = zaf.flu!(5, iabs%)  ' fluorescence by Ma
If tFluTyp% = 6 Or tFluTyp% = 12 Or tFluTyp% = 18 Or tFluTyp% = 24 Or tFluTyp% = 30 Or tFluTyp% = 36 Then ZAFFLUGetFluYield! = zaf.flu!(6, iabs%)  ' fluorescence by Mb

Exit Function

' Errors
ZAFFLUGetFluYieldError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFFLUGetFluYield"
ierror = True
Exit Function

End Function

Sub ZAFFLULoadLineWeights(iemitter As Integer, iabsorber As Integer, tFluTyp As Integer, tLineWeightRatio As Single)
' Load the relative line weights from the penelope_mod_transition_data.csv file from Penelope
' and calculate the relative emission intensity based on the emitter, absorber and fluorescence type
' 1=Ka by Ka  2=Ka by Kb  3=Ka by La  4=Ka by Lb  5=Ka by Ma  6=Ka by Mb
' 7=Kb by Ka  8=Kb by Kb  9=Kb by La 10=Kb by Lb 11=Kb by Ma 12=Kb by Mb
'13=La by Ka 14=La by Kb 15=La by La 16=La by Lb 17=La by Ma 18=La by Mb
'19=Lb by Ka 20=Lb by Kb 21=Lb by La 22=Lb by Lb 23=Lb by Ma 24=Lb by Mb
'25=Ma by Ka 26=Ma by Kb 27=Ma by La 28=Ma by Lb 29=Ma by Ma 30=Ma by Mb
'31=Mb by Ka 32=Mb by Kb 33=Mb by La 34=Mb by Lb 35=Mb by Ma 36=Mb by Mb

ierror = False
On Error GoTo ZAFFLULoadLineWeightsError

Dim iz As Integer, idest As Integer, isrc As Integer
Dim idest1 As Integer, isrc1 As Integer, idest2 As Integer, isrc2 As Integer

Dim trans_prob As Single, trans_energy As Single
Dim temp1 As Single, temp2 As Single
Dim tfilename As String, astring As String, bstring As String

Static initialized As Boolean
           
' Open the input (comma, tab or space delimited) and output files
If Not initialized Then
Close #Temp1FileNumber%
DoEvents

tfilename$ = ProgramPath$ & "penelope_mod_transition_data.csv"
Open tfilename$ For Input As #Temp1FileNumber%

' Read first line of column headings
Line Input #Temp1FileNumber%, astring   ' read column labels
If VerboseMode Then Call IOWriteLog(astring$)

' Loop on entries
Call IOStatusAuto(vbNullString)
icancelauto = False
Do Until EOF(Temp1FileNumber%)
Input #Temp1FileNumber%, iz%, idest%, isrc%, astring$, bstring$

trans_prob! = Val(astring$)
trans_energy! = Val(bstring$)

' Check for valid values
If iz% < 1 Or iz% > MAXELM% Then GoTo ZAFFLULoadLineWeightsBadEmitter

' Load jump ratio values
tmsg$ = "IZ=" & Format$(iz%) & ", " & Format$(idest%) & ", " & Format$(isrc%) & ", " & Format$(trans_prob!) & ", " & Format$(trans_energy!)
If VerboseMode Then Call IOWriteLog(tmsg$)

' Load data for this line
AllTransitionProbabilities!(iz%, idest%, isrc%) = trans_prob!
Loop

Close #Temp1FileNumber%
initialized = True
End If

' Load proper dest and source shells based on fluorescence type
If tFluTyp% = 1 Then idest1% = 1: isrc1% = 4: idest2% = 1: isrc2% = 4       ' Ka by Ka

' Return requested relative line intensity ratio based on emitter z, absorber z and fluorescence type
temp1! = AllTransitionProbabilities!(iemitter%, idest1%, isrc1%)
temp2! = AllTransitionProbabilities!(iabsorber%, idest2%, isrc2%)
tLineWeightRatio! = temp1! / temp2!

Exit Sub

' Errors
ZAFFLULoadLineWeightsError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFFLULoadLineWeights"
ierror = True
Close #Temp1FileNumber%
Exit Sub

ZAFFLULoadLineWeightsBadEmitter:
msg$ = "Invalid atomic number in " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFFLULoadLineWeights"
ierror = True
Close #Temp1FileNumber%
Exit Sub

End Sub

