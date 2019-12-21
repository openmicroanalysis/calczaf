Attribute VB_Name = "CodeFORMULA"
' (c) Copyright 1995-2020 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Dim FormulaString As String
Dim WeightString As String

Dim FormulaTmpSample(1 To 1) As TypeSample

Sub FormulaSaveWeight()
' Save the weight string entered by the user

ierror = False
On Error GoTo FormulaSaveWeightError

Dim astring As String

' Clear the sample
Call InitSample(FormulaTmpSample())

' Get the string
astring$ = FormWEIGHT.TextWeightPercentString.Text

' Check for decimals points
If InStr(astring$, ".") > 0 Then GoTo FormulaSaveWeightDecimal

' Convert the string
Call FormulaWeightToSample(astring$, FormulaTmpSample())
If ierror Then Exit Sub

' Save the weight percent string
WeightString$ = astring$

' Load as unknown
FormulaTmpSample(1).number% = 1
'FormulaTmpSample(1).Set% = 1       ' leave zero for CalcZAF calculations
FormulaTmpSample(1).Type% = 2
FormulaTmpSample(1).Name$ = astring$

' Unload FormWEIGHT
Unload FormWEIGHT
Exit Sub

' Errors
FormulaSaveWeightError:
MsgBox Error$, vbOKOnly + vbCritical, "FormulaSaveWeight"
ierror = True
Exit Sub

FormulaSaveWeightDecimal:
msg$ = "Please do not utilize decimal points for element subscripts. For better precision multiply all subscripts by 10 or 100, e.g. instead of Fe7.42 Mg29.80 Si19.07 O43.60, please enter Fe742 Mg2980 Si1907 O4360."
MsgBox msg$, vbOKOnly + vbExclamation, "FormulaSaveWeight"
ierror = True
Exit Sub

End Sub

Sub FormulaWeightToSample(astring As String, sample() As TypeSample)
' Routine to convert a weight percent string into a standard sample

ierror = False
On Error GoTo FormulaWeightToSampleError

Dim i As Integer, ip As Integer
Dim nels As Integer
Dim weight As Single, sum As Single

ReDim elms(1 To MAXCHAN%) As String
ReDim wtps(1 To MAXCHAN%) As Single

' Get elements and subscripts (weight percents)
Call MWCalculate(astring$, nels%, elms$(), wtps!(), weight!)
If ierror Then Exit Sub

' Calculate sum and normalize to 100% if necessary
sum! = 0#
For i% = 1 To nels%
sum! = sum! + wtps!(i%)
Next i%

For i% = 1 To nels%
wtps(i%) = 100# * wtps!(i%) / sum!
Next i%

' Save the element array strings to the sample
sample(1).LastElm% = nels%
sample(1).LastChan% = sample(1).LastElm%
sample(1).kilovolts! = DefaultKiloVolts!
sample(1).takeoff! = DefaultTakeOff!

' Load kilovolts array
For i% = 1 To sample(1).LastChan%
sample(1).TakeoffArray!(i%) = sample(1).takeoff!
sample(1).KilovoltsArray!(i%) = sample(1).kilovolts!
Next i%

For i% = 1 To sample(1).LastChan%
sample(1).Elsyms$(i%) = elms$(i%)

ip% = IPOS1(MAXELM%, sample(1).Elsyms$(i%), Symlo$())
If ip% > 0 Then
sample(1).Xrsyms$(i%) = Deflin$(ip%)
sample(1).numcat%(i%) = AllCat%(ip%)
sample(1).numoxd%(i%) = AllOxd%(ip%)
sample(1).AtomicCharges!(i%) = AllAtomicCharges!(ip%)
sample(1).CrystalNames$(i%) = Defcry$(ip%)
End If

sample(1).ElmPercents!(i%) = wtps!(i%)
Next i%

Exit Sub

' Errors
FormulaWeightToSampleError:
MsgBox Error$, vbOKOnly + vbCritical, "FormulaWeightToSample"
ierror = True
Exit Sub

End Sub

Sub FormulaFormulaToSample(astring As String, sample() As TypeSample)
' Routine to convert a formula string into a sample

ierror = False
On Error GoTo FormulaFormulaToSampleError

Dim i As Integer, ip As Integer
Dim nels As Integer
Dim weight As Single

ReDim elms(1 To MAXCHAN%) As String
ReDim subs(1 To MAXCHAN%) As Single

' Get elements and subscripts
Call MWCalculate(astring$, nels%, elms$(), subs!(), weight!)
If ierror Then Exit Sub

' Save the element array strings to the sample
sample(1).LastElm% = nels%
sample(1).LastChan% = sample(1).LastElm%
For i% = 1 To sample(1).LastChan%
sample(1).Elsyms$(i%) = elms$(i%)

ip% = IPOS1(MAXELM%, sample(1).Elsyms$(i%), Symlo$())
If ip% > 0 Then
sample(1).Xrsyms$(i%) = Deflin$(ip%)
sample(1).numcat%(i%) = AllCat%(ip%)
sample(1).numoxd%(i%) = AllOxd%(ip%)
sample(1).AtomicCharges!(i%) = AllAtomicCharges!(ip%)
End If

sample(1).ElmPercents!(i%) = ConvertAtomToWeight(sample(1).LastChan%, i%, subs!(), elms$())
If ierror Then Exit Sub

sample(1).AtomicNums%(i%) = ip%
If ierror Then Exit Sub
Next i%

Exit Sub

' Errors
FormulaFormulaToSampleError:
MsgBox Error$, vbOKOnly + vbCritical, "FormulaFormulaToSample"
ierror = True
Exit Sub

End Sub

Sub FormulaReturnSample(sample() As TypeSample)
' Return the modified sample

ierror = False
On Error GoTo FormulaReturnSampleError

' Load sample to return
sample(1) = FormulaTmpSample(1)

Exit Sub

' Errors
FormulaReturnSampleError:
MsgBox Error$, vbOKOnly + vbCritical, "FormulaReturnSample"
ierror = True
Exit Sub

End Sub

Sub FormulaSaveFormula()
' Save the formula string entered by the user

ierror = False
On Error GoTo FormulaSaveFormulaError

Dim astring As String

' Clear the sample
Call InitSample(FormulaTmpSample())

' Get the string
astring$ = FormFORMULA.TextFormulaString.Text

' Check for decimals points
If InStr(astring$, ".") > 0 Then GoTo FormulaSaveFormulaDecimal

' Convert the string
Call FormulaFormulaToSample(astring$, FormulaTmpSample())
If ierror Then Exit Sub

' Save the formula string
FormulaString$ = astring$

' Load as unknown
FormulaTmpSample(1).number% = 1
'FormulaTmpSample(1).Set% = 1       ' leave zero for CalcZAF calculations
FormulaTmpSample(1).Type% = 2
FormulaTmpSample(1).Name$ = astring$

' Unload FormFORMULA
Unload FormFORMULA
Exit Sub

' Errors
FormulaSaveFormulaError:
MsgBox Error$, vbOKOnly + vbCritical, "FormulaSaveFormula"
ierror = True
Exit Sub

FormulaSaveFormulaDecimal:
msg$ = "Please do not utilize decimal points for element subscripts. For better precision multiply all subscripts by 10 or 100, e.g. instead of Fe0.4 Mg1.6 SiO4, please enter Fe4 Mg16 Si10 O40."
MsgBox msg$, vbOKOnly + vbExclamation, "FormulaSaveFormula"
ierror = True
Exit Sub

End Sub

Sub FormulaSaveStdComp()
' Save the standard composition

ierror = False
On Error GoTo FormulaSaveStdCompError

Dim stdnum As Integer

' Save the standard
If FormSTDCOMP.ListAvailableStandards.ListCount > 0 Then
If FormSTDCOMP.ListAvailableStandards.ListIndex > -1 Then
stdnum% = FormSTDCOMP.ListAvailableStandards.ItemData(FormSTDCOMP.ListAvailableStandards.ListIndex)
End If
End If

If stdnum% = 0 Then GoTo FormulaSaveStdCompNoStandard

' Get standard composition
Call StandardGetMDBStandard(stdnum%, FormulaTmpSample())
If ierror Then Exit Sub

' Unload FormSTDCOMP
Unload FormSTDCOMP
Exit Sub

' Errors
FormulaSaveStdCompError:
MsgBox Error$, vbOKOnly + vbCritical, "FormulaSaveStdComp"
ierror = True
Exit Sub

FormulaSaveStdCompNoStandard:
msg$ = "No standard was selected"
MsgBox msg$, vbOKOnly + vbExclamation, "FormulaSaveStdComp"
ierror = True
Exit Sub

End Sub

Sub FormulaLoad(mode As Integer)
' Load the previous formula or weight string
'  mode = 0 formula
'  mode = 1 weight

ierror = False
On Error GoTo FormulaLoadError

' Load the string
If mode% = 0 Then FormFORMULA.TextFormulaString.Text = Trim$(FormulaString$)
If mode% = 1 Then FormWEIGHT.TextWeightPercentString.Text = Trim$(WeightString$)

Exit Sub

' Errors
FormulaLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "FormulaLoad"
ierror = True
Exit Sub

End Sub

