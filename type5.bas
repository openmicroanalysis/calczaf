Attribute VB_Name = "CodeTYPE5"
' (c) Copyright 1995-2023 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub TypeStandard(sample() As TypeSample)
' Type standard composition for CalcZAF

ierror = False
On Error GoTo TypeStandardError

Dim i As Integer
Dim temp As Single, sum As Single

' Load element strings
Call ElementLoadArrays(sample())
If ierror Then Exit Sub

' Type standard name
msg$ = StandardLoadDescription(sample())
If ierror Then Exit Sub
Call IOWriteLog(vbCrLf & vbCrLf & msg$)

' Calculate sum
sum! = 0#
For i% = 1 To sample(1).LastChan%
sum! = sum! + sample(1).ElmPercents!(i%)
Next i%

' Type elements
msg$ = vbCrLf & "ELEM: "
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(sample(1).Elsyup$(i%), a80$)
Next i%
msg$ = msg$ & Format$("   SUM  ", a80$)
Call IOWriteLog(msg$)

' Type out weight percents
msg$ = "ELWT: "
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(sample(1).ElmPercents!(i%), f83$), a80$)
Next i%
msg$ = msg$ & Format$(Format$(sum!, f83$), a80$)
Call IOWriteLog(msg$)

' Type out oxide percent
If sample(1).OxideOrElemental% = 1 Then
msg$ = "OXWT: "
For i% = 1 To sample(1).LastChan%
If Not MiscStringsAreSame(sample(1).Elsyms$(i%), Symlo$(ATOMIC_NUM_OXYGEN%)) Then
temp! = ConvertElmToOxd!(sample(1).ElmPercents!(i%), sample(1).Elsyms$(i%), sample(1).numcat%(i%), sample(1).numoxd%(i%))
Else
temp! = 0#
End If
msg$ = msg$ & Format$(Format$(temp!, f83$), a80$)
Next i%
msg$ = msg$ & Format$(Format$(sum!, f83$), a80$)
Call IOWriteLog(msg$)
End If

' Type out atomic percent
msg$ = "ATWT: "
For i% = 1 To sample(1).LastChan%
temp! = ConvertWeightToAtom(sample(1).LastChan%, i%, sample(1).ElmPercents!(), sample(1).Elsyms$())
msg$ = msg$ & Format$(Format$(temp!, f83$), a80$)
Next i%
msg$ = msg$ & Format$(Format$(100#, f83$), a80$)
Call IOWriteLog(msg$)

Exit Sub

' Errors
TypeStandardError:
MsgBox Error$, vbOKOnly + vbCritical, "TypeStandard"
ierror = True
Exit Sub

End Sub

Sub TypeStandards2(sample() As TypeSample, stdsample() As TypeSample)
' Type composition of all standards in CalcZAF run

ierror = False
On Error GoTo TypeStandards2Error

Dim stdnum As Integer, i As Integer

' Get composition of each standard
For i% = 1 To NumberofStandards%
stdnum% = StandardNumbers%(i%)

' Just load standard sample
Call StandardGetMDBStandard(stdnum%, stdsample())
If ierror Then Exit Sub

' Update takeoff and kilovolts for this run
stdsample(1).takeoff! = sample(1).takeoff!
stdsample(1).kilovolts! = sample(1).kilovolts!

' Type it out
Call TypeStandard(stdsample())
If ierror Then Exit Sub

Next i%

Exit Sub

' Errors
TypeStandards2Error:
MsgBox Error$, vbOKOnly + vbCritical, "TypeStandards2"
ierror = True
Exit Sub

End Sub
