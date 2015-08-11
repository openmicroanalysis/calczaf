Attribute VB_Name = "CodeCONVERT3"
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub ConvertWtsToAtomic(sampleline As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' Convert from weight (fraction or percent) to atomic percent for a sample

On Error GoTo ConvertWtsToAtomicError

Dim i As Integer
Dim sum As Single

' Calculate sum
sum! = 0#
For i% = 1 To sample(1).LastChan%
sum! = sum! + analysis.WtsData!(sampleline%, i%) / sample(1).AtomicWts!(i%)
Next i%

For i% = 1 To sample(1).LastChan%
If sum! <> 0# Then
analysis.AtPercents!(i%) = 100# * (analysis.WtsData!(sampleline%, i%) / sample(1).AtomicWts!(i%)) / sum!
Else
analysis.AtPercents!(i%) = 0#
End If
Next i%

Exit Sub

' Errors
ConvertWtsToAtomicError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertWtsToAtomic"
ierror = True
Exit Sub

End Sub

Sub ConvertWtsToOxMol(sampleline As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' Convert from weight (fraction or percent) to oxide mole percent for a sample

On Error GoTo ConvertWtsToOxMolError

Dim i As Integer
Dim sum As Single, temp As Single

' Calculate oxide weight percents
Call ConvertWtsToOxide(sampleline%, analysis, sample())
If ierror Then Exit Sub

' Calculate sum of oxides divided by molecular weights
sum! = 0#
For i% = 1 To sample(1).LastChan%
If sample(1).AtomicNums%(i%) <> 8 Then
temp! = sample(1).AtomicWts!(i%) * sample(1).numcat%(i%) + AllAtomicWts!(8) * sample(1).numoxd%(i%)
If temp! > 0# Then
sum! = sum! + analysis.OxPercents!(i%) / temp!
End If
End If
Next i%

For i% = 1 To sample(1).LastChan%
If sample(1).AtomicNums%(i%) <> 8 Then
temp! = sample(1).AtomicWts!(i%) * sample(1).numcat%(i%) + AllAtomicWts!(8) * sample(1).numoxd%(i%)
If sum! <> 0# And temp! <> 0# Then
analysis.OxMolPercents!(i%) = 100# * (analysis.OxPercents!(i%) / temp!) / sum!
Else
analysis.OxMolPercents!(i%) = 0#
End If
End If
Next i%

Exit Sub

' Errors
ConvertWtsToOxMolError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertWtsToOxMol"
ierror = True
Exit Sub

End Sub

Sub ConvertWtsToOxide(sampleline As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' Convert from weight percent to oxide percent for a sample

On Error GoTo ConvertWtsToOxideError

Dim i As Integer
Dim temp As Single

temp! = 0#
For i% = 1 To sample(1).LastChan%
analysis.OxPercents!(i%) = 0
temp! = analysis.WtsData!(sampleline%, i%) * (sample(1).AtomicWts!(i%) * sample(1).numcat%(i%) + AllAtomicWts!(8) * sample(1).numoxd%(i%))
analysis.OxPercents!(i%) = temp! / (sample(1).AtomicWts!(i%) * sample(1).numcat%(i%))
Next i%

Exit Sub

' Errors
ConvertWtsToOxideError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertWtsToOxide"
ierror = True
Exit Sub

End Sub

Sub ConvertWtsToFormula(sampleline As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' Convert from weight (fraction or percent) to formula basis for a sample

ierror = False
On Error GoTo ConvertWtsToFormulaError

Dim i As Integer, ip As Integer
Dim totalatoms As Single, temp As Single, TotalCations As Single

ReDim atoms(1 To MAXCHAN%) As Single, basis(1 To MAXCHAN%) As Single

' Init to zero
For i% = 1 To sample(1).LastChan% + 1
analysis.CalData!(sampleline%, i%) = 0#
Next i%

' Calculate sum
totalatoms! = 0#
For i% = 1 To sample(1).LastChan%
If sample(1).AtomicWts!(i%) > 0# Then atoms!(i%) = analysis.WtsData!(sampleline%, i%) / sample(1).AtomicWts!(i%)
totalatoms! = totalatoms! + atoms!(i%)
Next i%

' Check for insufficient total atoms
If totalatoms! < 0.01 Then
msg$ = TypeLoadString$(sample())
msg$ = "There is an insufficient total number of atoms (usually caused by low totals) to calculate atomic ratios for sample " & msg$ & ". "
msg$ = msg$ & "Please delete the data point (line " & Format$(sample(1).Linenumber&(sampleline%)) & ") or remove the atomic percent or formula calculation." & vbCrLf & vbCrLf
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
Exit Sub
End If
 
' Calculate formulas
If sample(1).FormulaElement$ <> vbNullString Then     ' normal formula calculation
ip% = IPOS1(sample(1).LastChan%, sample(1).FormulaElement$, sample(1).Elsyms$())
If ip% = 0 Then GoTo ConvertWtsToFormulaInvalidFormulaElement

' Check for insufficient formula basis element
If atoms!(ip%) < 0.01 Then
msg$ = TypeLoadString$(sample())
msg$ = "There is an insufficient concentration of the formula basis element " & sample(1).Elsyup$(ip%) & " (usually caused by a very low concentration of the element), to calculate a formula for sample " & msg$ & ". "
msg$ = msg$ & "Please delete the data point (line " & Format$(sample(1).Linenumber&(sampleline%)) & ") or change the formula basis element." & vbCrLf
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
Exit Sub
End If

temp! = sample(1).FormulaRatio! / atoms!(ip%)
For i% = 1 To sample(1).LastChan%
basis!(i%) = atoms!(i%) * temp!
Next i%

' Calculate sum of cations (normalize sum of cations to formula atoms))
Else
TotalCations! = 0#
For i% = 1 To sample(1).LastChan%
If sample(1).AtomicCharges!(i%) > 0# Then TotalCations! = TotalCations! + atoms!(i%)
Next i%

If TotalCations! < 0.01 Then
msg$ = TypeLoadString$(sample())
msg$ = "There is an insufficient concentration of the cation sum (usually caused by a very low total), to calculate a formula for sample " & msg$ & ". "
msg$ = msg$ & "Please delete the data point (line " & Format$(sample(1).Linenumber&(sampleline%)) & ") or change the formula assignment." & vbCrLf & vbCrLf
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
Exit Sub
End If

' Normalize to total number of cations
temp! = sample(1).FormulaRatio! / TotalCations!
For i% = 1 To sample(1).LastChan%
basis!(i%) = atoms!(i%) * temp!
Next i%
End If

' End of calculations for this row, load into "analysis.CalData" array
For i% = 1 To sample(1).LastChan%
analysis.CalData!(sampleline%, i%) = basis!(i%)
Next i%

' Calculate the sums
temp! = 0#
For i% = 1 To sample(1).LastChan%
temp! = temp! + analysis.CalData!(sampleline%, i%)
Next i%
analysis.CalData!(sampleline%, sample(1).LastChan% + 1) = temp!

Exit Sub

' Errors
ConvertWtsToFormulaError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertWtsToFormula"
ierror = True
Exit Sub

ConvertWtsToFormulaInvalidFormulaElement:
msg$ = TypeLoadString$(sample())
msg$ = "Element " & sample(1).FormulaElement$ & " is an invalid formula element for sample " & msg$
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertWtsToFormula"
ierror = True
Exit Sub

End Sub

Function ConvertIsDifferenceFormulaElement(tformula As String, tsym As String, sample() As TypeSample) As Boolean
' Returns true if the passed element is in the passed formula by difference

ierror = False
On Error GoTo ConvertIsDifferenceFormulaElementError

Dim n As Integer

Dim numelms As Integer
Dim weight As Single

Dim elems(1 To MAXCHAN%) As String
Dim fatoms(1 To MAXCHAN%) As Single

ConvertIsDifferenceFormulaElement = False
If tformula$ = vbNullString Then Exit Function

' Parse formula
Call MWCalculate(tformula$, numelms%, elems$(), fatoms!(), weight!)
If ierror Then Exit Function

' Check for passed element in formula
For n% = 1 To numelms%
If MiscStringsAreSame(tsym$, elems$(n%)) Then ConvertIsDifferenceFormulaElement = True
Next n%

Exit Function

' Errors
ConvertIsDifferenceFormulaElementError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertIsDifferenceFormulaElement"
ierror = True
Exit Function

End Function

Function ConvertIsDifferenceFormulaElementsSpecified(tformula As String, sample() As TypeSample) As Boolean
' Returns true if all elements in the formula by difference are already specified elements

ierror = False
On Error GoTo ConvertIsDifferenceFormulaElementsSpecifiedError

Dim n As Integer, ip As Integer

Dim numelms As Integer
Dim weight As Single

Dim elems(1 To MAXCHAN%) As String
Dim fatoms(1 To MAXCHAN%) As Single

ConvertIsDifferenceFormulaElementsSpecified = True
If tformula$ = vbNullString Then Exit Function

' Parse formula
Call MWCalculate(tformula$, numelms%, elems$(), fatoms!(), weight!)
If ierror Then Exit Function

' Check if all elements in formula are already specified in sample
For n% = 1 To numelms%
ip% = IPOS1B(sample(1).LastElm + 1, sample(1).LastChan%, elems$(n%), sample(1).Elsyms$())
If ip% = 0 Then ConvertIsDifferenceFormulaElementsSpecified = False
Next n%

Exit Function

' Errors
ConvertIsDifferenceFormulaElementsSpecifiedError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertIsDifferenceFormulaElementsSpecified"
ierror = True
Exit Function

End Function



