Attribute VB_Name = "CodeCONVERT"
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

Function ConvertAtomToWeight(lchan As Integer, row As Integer, atoms() As Single, syms() As String) As Single
' Calculate weight percent from total atomic percent or formula atoms for a single element

ierror = False
On Error GoTo ConvertAtomToWeightError

Dim i As Integer, ip As Integer
Dim sum As Single, temp As Single

' Sum the atomic percents or formulas
sum = 0#
For i% = 1 To lchan%
ip% = IPOS1(MAXELM%, syms$(i%), Symlo$())
If ip% > 0 Then
sum! = sum! + atoms!(i%) * AllAtomicWts!(ip%)
End If
Next i%

' Check for bad sum
If sum! <= 0# Then
msg$ = "Bad atomic or formula sum = " & Format$(sum!)
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertAtomToWeight"
ierror = True
Exit Function
End If

' Do atomic percent or formula atoms to weight percent conversion
ConvertAtomToWeight! = 0#
ip% = IPOS1(MAXELM%, syms$(row%), Symlo$())
If ip% > 0 Then
temp! = atoms!(row%) * AllAtomicWts!(ip%)
ConvertAtomToWeight! = 100# * temp! / sum!
End If

Exit Function

' Errors
ConvertAtomToWeightError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertAtomToWeight"
ierror = True
Exit Function

End Function

Function ConvertElmToOxd(weight As Single, sym As String, cat As Integer, oxd As Integer) As Single
' Converts an elemental weight percent to an oxide weight percent

ierror = False
On Error GoTo ConvertElmToOxdError

Dim temp As Double
Dim ip As Integer

temp# = 0#
ip% = IPOS1(MAXELM%, sym$, Symlo$())
If ip% > 0 Then
temp# = weight! * (AllAtomicWts!(ip%) * cat% + AllAtomicWts!(ATOMIC_NUM_OXYGEN%) * oxd%)
ConvertElmToOxd! = temp# / (AllAtomicWts!(ip%) * cat%)
End If

Exit Function

' Errors
ConvertElmToOxdError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertElmToOxd"
ierror = True
Exit Function

End Function

Sub ConvertMinerals(analysis As TypeAnalysis, sample() As TypeSample)
' Converts atoms to mineral end members

ierror = False
On Error GoTo ConvertMineralsError

Dim j As Integer

' Check for sufficient basis of mineral end members basis elements
If ConvertMineralsLineCheck(analysis, sample()) Then

' Calculate all rows
For j% = 1 To sample(1).Datarows%
If sample(1).LineStatus(j%) Then

' Perform actual mineral end member calculation
Call ConvertMineralsLine(j%, analysis, sample())
If ierror Then Exit Sub

End If
Next j%

Else
msg$ = "Insufficient number of the basis atoms for the specified mineral end-member calculation for sample " & SampleGetString2$(sample())
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
End If

Exit Sub

' Errors
ConvertMineralsError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertMinerals"
ierror = True
Exit Sub

End Sub

Sub ConvertMineralsLine(sampleline As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' Perform the actual mineral end-member calculation

ierror = False
On Error GoTo ConvertMineralsLineError

Dim ia As Integer, ib As Integer, ic As Integer, id As Integer
Dim temp1 As Single, temp2 As Single, temp3 As Single, temp4 As Single
Dim sum As Single

sum! = 0#
ia% = 0
ib% = 0
ic% = 0
id% = 0
temp1! = 0
temp2! = 0
temp3! = 0
temp4! = 0

' Olivine (skip disabled quant elements)
If sample(1).MineralFlag% = 1 Then
ia% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_MAGNESIUM%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Mg
ib% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_IRON%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())        ' Fe

If ia% > 0 Then sum! = sum! + analysis.CalData!(sampleline%, ia%)
If ib% > 0 Then sum! = sum! + analysis.CalData!(sampleline%, ib%)
If sum! <= 0# Then GoTo ConvertMineralsLineZeroSum

If ia% > 0 Then temp1! = analysis.CalData!(sampleline%, ia%) / sum!
If ib% > 0 Then temp2! = analysis.CalData!(sampleline%, ib%) / sum!

If ia% > 0 Then analysis.CalData!(sampleline%, 1) = temp1! * 100#
If ib% > 0 Then analysis.CalData!(sampleline%, 2) = temp2! * 100#
analysis.CalData!(sampleline%, 3) = 0#
analysis.CalData!(sampleline%, 4) = 0#
End If

' Feldspar (skip disabled quant elements)
If sample(1).MineralFlag% = 2 Then
ia% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_SODIUM%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Na
ib% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_CALCIUM%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Ca
ic% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_POTASSIUM%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' K

If ia% > 0 Then sum! = sum! + analysis.CalData!(sampleline%, ia%)
If ib% > 0 Then sum! = sum! + analysis.CalData!(sampleline%, ib%)
If ic% > 0 Then sum! = sum! + analysis.CalData!(sampleline%, ic%)
If sum! <= 0# Then GoTo ConvertMineralsLineZeroSum

If ia% > 0 Then temp1! = analysis.CalData!(sampleline%, ia%) / sum!
If ib% > 0 Then temp2! = analysis.CalData!(sampleline%, ib%) / sum!
If ic% > 0 Then temp3! = analysis.CalData!(sampleline%, ic%) / sum!

If ia% > 0 Then analysis.CalData!(sampleline%, 1) = temp1! * 100#
If ib% > 0 Then analysis.CalData!(sampleline%, 2) = temp2! * 100#
If ic% > 0 Then analysis.CalData!(sampleline%, 3) = temp3! * 100#
analysis.CalData!(sampleline%, 4) = 0#
End If

' Pyroxene (skip disabled quant elements)
If sample(1).MineralFlag% = 3 Then
ia% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_CALCIUM%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Ca
ib% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_MAGNESIUM%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Mg
ic% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_IRON%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Fe

If ia% > 0 Then sum! = sum! + analysis.CalData!(sampleline%, ia%)
If ib% > 0 Then sum! = sum! + analysis.CalData!(sampleline%, ib%)
If ic% > 0 Then sum! = sum! + analysis.CalData!(sampleline%, ic%)
If sum! <= 0# Then GoTo ConvertMineralsLineZeroSum

If ia% > 0 Then temp1! = analysis.CalData!(sampleline%, ia%) / sum!
If ib% > 0 Then temp2! = analysis.CalData!(sampleline%, ib%) / sum!
If ic% > 0 Then temp3! = analysis.CalData!(sampleline%, ic%) / sum!

If ia% > 0 Then analysis.CalData!(sampleline%, 1) = temp1! * 100#
If ib% > 0 Then analysis.CalData!(sampleline%, 2) = temp2! * 100#
If ic% > 0 Then analysis.CalData!(sampleline%, 3) = temp3! * 100#
analysis.CalData!(sampleline%, 4) = 0#
End If

' Garnet (Normal) (skip disabled quant elements)
If sample(1).MineralFlag% = 4 Then
ia% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_CALCIUM%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Ca
ib% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_MAGNESIUM%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Mg
ic% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_IRON%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Fe
id% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_MANGANESE%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Mn

If ia% > 0 Then sum! = sum! + analysis.CalData!(sampleline%, ia%)
If ib% > 0 Then sum! = sum! + analysis.CalData!(sampleline%, ib%)
If ic% > 0 Then sum! = sum! + analysis.CalData!(sampleline%, ic%)
If id% > 0 Then sum! = sum! + analysis.CalData!(sampleline%, id%)
If sum! <= 0# Then GoTo ConvertMineralsLineZeroSum

If ia% > 0 Then temp1! = analysis.CalData!(sampleline%, ia%) / sum!
If ib% > 0 Then temp2! = analysis.CalData!(sampleline%, ib%) / sum!
If ic% > 0 Then temp3! = analysis.CalData!(sampleline%, ic%) / sum!
If id% > 0 Then temp4! = analysis.CalData!(sampleline%, id%) / sum!

If ia% > 0 Then analysis.CalData!(sampleline%, 1) = temp1! * 100#
If ib% > 0 Then analysis.CalData!(sampleline%, 2) = temp2! * 100#
If ic% > 0 Then analysis.CalData!(sampleline%, 3) = temp3! * 100#
If id% > 0 Then analysis.CalData!(sampleline%, 4) = temp4! * 100#
End If

' Garnet (Grossular, Andradite, Uvarovite) (skip disabled quant elements)
If sample(1).MineralFlag% = 5 Then
ia% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_ALUMINUM%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Al
ib% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_IRON%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Fe
ic% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_CHROMIUM%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Cr

If ia% > 0 Then sum! = sum! + analysis.CalData!(sampleline%, ia%)
If ib% > 0 Then sum! = sum! + analysis.CalData!(sampleline%, ib%)
If ic% > 0 Then sum! = sum! + analysis.CalData!(sampleline%, ic%)
If sum! <= 0# Then GoTo ConvertMineralsLineZeroSum

If ia% > 0 Then temp1! = analysis.CalData!(sampleline%, ia%) / sum!
If ib% > 0 Then temp2! = analysis.CalData!(sampleline%, ib%) / sum!
If ic% > 0 Then temp3! = analysis.CalData!(sampleline%, ic%) / sum!

If ia% > 0 Then analysis.CalData!(sampleline%, 1) = temp1! * 100#
If ib% > 0 Then analysis.CalData!(sampleline%, 2) = temp2! * 100#
If ic% > 0 Then analysis.CalData!(sampleline%, 3) = temp3! * 100#
End If

Exit Sub

' Errors
ConvertMineralsLineError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertMineralsLine"
ierror = True
Exit Sub

ConvertMineralsLineZeroSum:
msg$ = "Zero sum on line " & Format$(sample(1).Linenumber&(sampleline%)) & " for sample " & SampleGetString2$(sample()) & ". Most likely an insufficient number of the basis atoms for the specified mineral end-member calculation."
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertMineralsLine"
ierror = True
Exit Sub

End Sub

Function ConvertMineralsLineCheck(analysis As TypeAnalysis, sample() As TypeSample) As Boolean
' Perform only a check for sufficient concentrations for mineral end-member calculations

ierror = False
On Error GoTo ConvertMineralsLineCheckError

Dim j As Integer
Dim ia As Integer, ib As Integer, ic As Integer, id As Integer
Dim sum As Single

' Assume OK
ConvertMineralsLineCheck = True

' Check each data line
For j% = 1 To sample(1).Datarows%
If sample(1).LineStatus(j%) Then

sum! = 0#
ia% = 0
ib% = 0
ic% = 0
id% = 0

' Olivine (skip disabled quant elements)
If sample(1).MineralFlag% = 1 Then
ia% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_MAGNESIUM%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Mg
ib% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_IRON%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Fe

If ia% > 0 Then sum! = sum! + analysis.CalData!(j%, ia%)
If ib% > 0 Then sum! = sum! + analysis.CalData!(j%, ib%)
If sum! <= 0# Then ConvertMineralsLineCheck = False
End If

' Feldspar (skip disabled quant elements)
If sample(1).MineralFlag% = 2 Then
ia% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_SODIUM%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Na
ib% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_CALCIUM%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Ca
ic% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_POTASSIUM%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' K

If ia% > 0 Then sum! = sum! + analysis.CalData!(j%, ia%)
If ib% > 0 Then sum! = sum! + analysis.CalData!(j%, ib%)
If ic% > 0 Then sum! = sum! + analysis.CalData!(j%, ic%)
If sum! <= 0# Then ConvertMineralsLineCheck = False
End If

' Pyroxene (skip disabled quant elements)
If sample(1).MineralFlag% = 3 Then
ia% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_CALCIUM%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Ca
ib% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_MAGNESIUM%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Mg
ic% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_IRON%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Fe

If ia% > 0 Then sum! = sum! + analysis.CalData!(j%, ia%)
If ib% > 0 Then sum! = sum! + analysis.CalData!(j%, ib%)
If ic% > 0 Then sum! = sum! + analysis.CalData!(j%, ic%)
If sum! <= 0# Then ConvertMineralsLineCheck = False
End If

' Garnet (Normal) (skip disabled quant elements)
If sample(1).MineralFlag% = 4 Then
ia% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_CALCIUM%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Ca
ib% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_MAGNESIUM%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Mg
ic% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_IRON%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Fe
id% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_MANGANESE%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Mn

If ia% > 0 Then sum! = sum! + analysis.CalData!(j%, ia%)
If ib% > 0 Then sum! = sum! + analysis.CalData!(j%, ib%)
If ic% > 0 Then sum! = sum! + analysis.CalData!(j%, ic%)
If id% > 0 Then sum! = sum! + analysis.CalData!(j%, id%)
If sum! <= 0# Then ConvertMineralsLineCheck = False
End If

' Garnet (Grossular, Andradite, Uvarovite) (skip disabled quant elements)
If sample(1).MineralFlag% = 5 Then
ia% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_ALUMINUM%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Al
ib% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_IRON%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Fe
ic% = IPOS1DQ(sample(1).LastChan%, Symlo$(ATOMIC_NUM_CHROMIUM%), sample(1).Elsyms$(), sample(1).DisableQuantFlag%())   ' Cr

If ia% > 0 Then sum! = sum! + analysis.CalData!(j%, ia%)
If ib% > 0 Then sum! = sum! + analysis.CalData!(j%, ib%)
If ic% > 0 Then sum! = sum! + analysis.CalData!(j%, ic%)
If sum! <= 0# Then ConvertMineralsLineCheck = False
End If

End If
Next j%

Exit Function

' Errors
ConvertMineralsLineCheckError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertMineralsLineCheck"
ierror = True
Exit Function

End Function

Function ConvertOxdToElm(weight As Single, sym As String, cat As Integer, oxd As Integer) As Single
' Converts an oxide weight percent to an elemental weight percent

ierror = False
On Error GoTo ConvertOxdToElmError

Dim temp As Single
Dim ip As Integer

temp! = 0#
ip% = IPOS1(MAXELM%, sym$, Symlo$())
If ip% > 0 Then
temp! = weight! * (AllAtomicWts!(ip%) * cat%)
ConvertOxdToElm! = temp! / (AllAtomicWts!(ip%) * cat% + AllAtomicWts!(ATOMIC_NUM_OXYGEN%) * oxd%)
End If

Exit Function

' Errors
ConvertOxdToElmError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertOxdToElm"
ierror = True
Exit Function

End Function

Function ConvertTotalToExcessOxygen(mode As Integer, unksample() As TypeSample, stdsample() As TypeSample) As Single
' Function to convert the total oxygen in a standard composition (calculated and excess)
' to actual excess oxygen based on the current unknown sample cation ratios.
' mode = 1 elemental mode
' mode = 2 oxide mode

ierror = False
On Error GoTo ConvertTotalToExcessOxygenError

Dim i As Integer, ip As Integer
Dim cat As Integer, oxd As Integer
Dim total As Single, oxygen As Single
Dim temp As Single

ConvertTotalToExcessOxygen! = 0#

' Find the position of unknown specified oxygen in the standard composition array
ip% = IPOS1(stdsample(1).LastChan%, Symlo$(ATOMIC_NUM_OXYGEN%), stdsample(1).Elsyms$())
If ip% = 0 Then Exit Function

' Load total oxygen from standard composition database
total! = stdsample(1).ElmPercents!(ip%)

' If unknown is calculated as elemental, just return
ConvertTotalToExcessOxygen! = total!
If unksample(1).OxideOrElemental% = 2 And mode% = 1 Then Exit Function

' Now calculate the oxygen based on the unksample cation ratios
oxygen! = 0#
For i% = 1 To stdsample(1).LastChan%
cat% = stdsample(1).numcat%(i%)
oxd% = stdsample(1).numoxd%(i%)

' See if sample uses different cation ratios than standard, if so, use them
ip% = IPOS1(unksample(1).LastChan%, stdsample(1).Elsyms$(i%), unksample(1).Elsyms$())
If ip% > 0 Then
cat% = unksample(1).numcat%(ip%)
oxd% = unksample(1).numoxd%(ip%)
End If

' Calculated difference between oxide and elemental weight percent (calculated oxygen)
temp! = ConvertElmToOxd(stdsample(1).ElmPercents!(i%), stdsample(1).Elsyms$(i%), cat%, oxd%)
oxygen! = oxygen! + (temp! - stdsample(1).ElmPercents!(i%))
Next i%

ConvertTotalToExcessOxygen! = total! - oxygen!

Exit Function

' Errors
ConvertTotalToExcessOxygenError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertTotalToExcessOxygen"
ierror = True
Exit Function

End Function

Function ConvertWeightToAtom(lchan As Integer, chan As Integer, wts() As Single, syms() As String) As Single
' Calculate weight percent to atomic percent for a single element

ierror = False
On Error GoTo ConvertWeightToAtomError

Dim i As Integer, ip As Integer
Dim sum As Single, temp As Single

' Sum the atoms of the elemental weight percents
sum = 0#
For i% = 1 To lchan%
ip% = IPOS1(MAXELM%, syms$(i%), Symlo$())
If ip% > 0 Then
sum! = sum! + wts!(i%) / AllAtomicWts!(ip%)
End If
Next i%

' Check for bad sum
If sum! <= 0# Then
msg$ = "Bad atomic sum = " & Format$(sum!)
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertWeightToAtom"
ierror = True
Exit Function
End If

' Do weight percent to atomic percent conversion
ConvertWeightToAtom! = 0#
ip% = IPOS1(MAXELM%, syms$(chan%), Symlo$())
If ip% > 0 Then
temp! = wts!(chan%) / AllAtomicWts!(ip%)
ConvertWeightToAtom! = 100# * temp! / sum!
End If

Exit Function

' Errors
ConvertWeightToAtomError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertWeightToAtom"
ierror = True
Exit Function

End Function

Sub ConvertWeightToElectron(lchan As Integer, atnums() As Single, atwts() As Single, wts() As Single, elecs() As Single)
' Convert from weight fraction to electron fraction

ierror = False
On Error GoTo ConvertWeightToElectronError

Dim i As Integer
Dim sum As Single

' Calculate sum
sum! = 0#
For i% = 1 To lchan%
sum! = sum! + atnums!(i%) * wts!(i%) / atwts!(i%)
Next i%
If sum! = 0# Then GoTo ConvertWeightToElectronZeroSum

For i% = 1 To lchan%
If sum! <> 0# Then
elecs!(i%) = atnums!(i%) * (wts!(i%) / atwts!(i%)) / sum!
Else
elecs!(i%) = 0#
End If
Next i%

Exit Sub

' Errors
ConvertWeightToElectronError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertWeightToElectron"
ierror = True
Exit Sub

ConvertWeightToElectronZeroSum:
msg$ = "Sum of concentrations is zero"
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertWeightToElectron"
ierror = True
Exit Sub

End Sub

Sub ConvertWeightToElectron2(lchan As Integer, exponent As Single, atnums() As Single, atwts() As Single, wts() As Single, elecs() As Single)
' Convert from weight fraction to electron fraction (Z to "exponent")

ierror = False
On Error GoTo ConvertWeightToElectron2Error

Dim i As Integer
Dim sum As Single

' Calculate sum
sum! = 0#
For i% = 1 To lchan%
sum! = sum! + (atnums!(i%) ^ exponent!) * wts!(i%) / atwts!(i%)
Next i%
If sum! = 0# Then GoTo ConvertWeightToElectron2ZeroSum

For i% = 1 To lchan%
If sum! <> 0# Then
elecs!(i%) = (atnums!(i%) ^ exponent!) * (wts!(i%) / atwts!(i%)) / sum!
Else
elecs!(i%) = 0#
End If
Next i%

Exit Sub

' Errors
ConvertWeightToElectron2Error:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertWeightToElectron2"
ierror = True
Exit Sub

ConvertWeightToElectron2ZeroSum:
msg$ = "Sum of concentrations is zero"
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertWeightToElectron2"
ierror = True
Exit Sub

End Sub

Sub ConvertWeightToAtomic(lchan As Integer, atwts() As Single, wts() As Single, atoms() As Single)
' Convert from weight percent/fraction to atomic fraction for the entire array

ierror = False
On Error GoTo ConvertWeightToAtomicError

Dim i As Integer
Dim sum As Single

' Calculate sum
sum! = 0#
For i% = 1 To lchan%
sum! = sum! + wts!(i%) / atwts!(i%)
Next i%

For i% = 1 To lchan%
If sum! <> 0# Then
atoms!(i%) = (wts!(i%) / atwts!(i%)) / sum!
Else
atoms!(i%) = 0#
End If
Next i%

Exit Sub

' Errors
ConvertWeightToAtomicError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertWeightToAtomic"
ierror = True
Exit Sub

End Sub

Sub ConvertWeightToOxide(lchan As Integer, atwts() As Single, NumCats() As Integer, NumOxds() As Integer, wts() As Single, excess As Single, oxwts() As Single)
' Convert from weight percent to oxide percent for the entire array

ierror = False
On Error GoTo ConvertWeightToOxideError

Dim i As Integer
Dim temp As Single

For i% = 1 To lchan%
If atwts!(i%) <> AllAtomicWts!(ATOMIC_NUM_OXYGEN%) Then
temp! = wts!(i%) * (atwts!(i%) * NumCats%(i%) + AllAtomicWts!(ATOMIC_NUM_OXYGEN%) * NumOxds%(i%))
oxwts!(i%) = temp! / (atwts!(i%) * NumCats%(i%))
Else
oxwts!(i%) = excess!
End If
Next i%

Exit Sub

' Errors
ConvertWeightToOxideError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertWeightToOxide"
ierror = True
Exit Sub

End Sub

Sub ConvertAtomicToWeight(lchan As Integer, atwts() As Single, wts() As Single, atoms() As Single)
' Convert from atomic fraction to weight fraction for the entire array

ierror = False
On Error GoTo ConvertAtomicToWeightError

Dim i As Integer
Dim sum As Single

' Calculate sum
sum! = 0#
For i% = 1 To lchan%
sum! = sum! + atoms!(i%) * atwts!(i%)
Next i%

For i% = 1 To lchan%
If sum! <> 0# Then
wts!(i%) = (atoms!(i%) * atwts!(i%)) / sum!
Else
wts!(i%) = 0#
End If
Next i%

Exit Sub

' Errors
ConvertAtomicToWeightError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertAtomicToWeight"
ierror = True
Exit Sub

End Sub

Sub ConvertElectronToWeight(lchan As Integer, atnums() As Single, atwts() As Single, elecs() As Single, wts() As Single)
' Convert from electron fraction to weight fraction

ierror = False
On Error GoTo ConvertElectronToWeightError

Dim i As Integer
Dim sum As Single

' Calculate sum
sum! = 0#
For i% = 1 To lchan%
sum! = sum! + atwts!(i%) * elecs!(i%) / atnums!(i%)
Next i%

For i% = 1 To lchan%
If sum! <> 0# Then
wts!(i%) = atwts!(i%) * (elecs!(i%) / atnums!(i%)) / sum!
Else
wts!(i%) = 0#
End If
Next i%

Exit Sub

' Errors
ConvertElectronToWeightError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertElectronToWeight"
ierror = True
Exit Sub

End Sub

Function ConvertHalogensToOxygen(lchan As Integer, syms() As String, dqs() As Integer, conc() As Single) As Single
' Calculate the equivalent oxygen based on the halogen concentrations or weight percents
'  lchan = number of channels
'  syms() = symbol (element) array
'  dqs() = disable quant flag array
'  conc() = concentration of element array

ierror = False
On Error GoTo ConvertHalogensToOxygenError

Dim FtoO As Single, CltoO As Single
Dim BrtoO As Single, ItoO As Single

Dim ip As Integer
Dim temp As Single

FtoO! = AllAtomicWts!(ATOMIC_NUM_OXYGEN%) / (AllAtomicWts!(ATOMIC_NUM_FLUORINE%) * 2#)
CltoO! = AllAtomicWts!(ATOMIC_NUM_OXYGEN%) / (AllAtomicWts!(ATOMIC_NUM_CHLORINE%) * 2#)
BrtoO! = AllAtomicWts!(ATOMIC_NUM_OXYGEN%) / (AllAtomicWts!(ATOMIC_NUM_BROMINE%) * 2#)
ItoO! = AllAtomicWts!(ATOMIC_NUM_OXYGEN%) / (AllAtomicWts!(ATOMIC_NUM_IODINE%) * 2#)

' Check for fluorine
temp! = 0#
ip% = IPOS1DQ(lchan%, Symlo$(ATOMIC_NUM_FLUORINE%), syms$(), dqs%())
If ip% > 0 Then
temp! = conc!(ip%) * FtoO!
End If

' Check for chlorine
ip% = IPOS1DQ(lchan%, Symlo$(ATOMIC_NUM_CHLORINE%), syms$(), dqs%())
If ip% > 0 Then
temp! = temp! + conc!(ip%) * CltoO!
End If

' Check for bromine
ip% = IPOS1DQ(lchan%, Symlo$(ATOMIC_NUM_BROMINE%), syms$(), dqs%())
If ip% > 0 Then
temp! = temp! + conc!(ip%) * BrtoO!
End If

' Check for iodine
ip% = IPOS1DQ(lchan%, Symlo$(ATOMIC_NUM_IODINE%), syms$(), dqs%())
If ip% > 0 Then
temp! = temp! + conc!(ip%) * ItoO!
End If

ConvertHalogensToOxygen! = temp!

Exit Function

' Errors
ConvertHalogensToOxygenError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertHalogensToOxygen"
ierror = True
Exit Function

End Function

Function ConvertOxygenFromCations(sample() As TypeSample) As Single
' Calculate oxygen from cations for sample

ierror = False
On Error GoTo ConvertOxygenFromCationsError

Dim i As Integer
Dim temp As Single

' Calculate oxygen from cations
temp! = 0#
For i% = 1 To sample(1).LastChan%
If i% <> sample(1).OxygenChannel% Then
temp! = temp! + (ConvertElmToOxd(sample(1).ElmPercents!(i%), sample(1).Elsyms$(i%), sample(1).numcat%(i%), sample(1).numoxd%(i%)) - sample(1).ElmPercents!(i%))
End If
Next i%

ConvertOxygenFromCations! = temp!
Exit Function

' Errors
ConvertOxygenFromCationsError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertOxygenFromCations"
ierror = True
Exit Function

End Function

Function ConvertOxygenFromCations2(sampleline As Integer, analysis As TypeAnalysis, sample() As TypeSample) As Single
' Calculate oxygen from cations for sample (uses "analysis" array)

ierror = False
On Error GoTo ConvertOxygenFromCations2Error

Dim i As Integer
Dim temp As Single

' Calculate oxygen from cations
temp! = 0#
For i% = 1 To sample(1).LastChan%
If i% <> sample(1).OxygenChannel% Then
temp! = temp! + (analysis.OxPercents!(i%) - analysis.WtsData!(sampleline%, i%))
End If
Next i%

ConvertOxygenFromCations2! = temp!
Exit Function

' Errors
ConvertOxygenFromCations2Error:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertOxygenFromCations2"
ierror = True
Exit Function

End Function

Function ConvertOxygenFromCations3(analysis As TypeAnalysis, sample() As TypeSample) As Single
' Calculate oxygen from cations for sample analyzed weight percents

ierror = False
On Error GoTo ConvertOxygenFromCations3Error

Dim i As Integer
Dim temp1 As Single, temp2 As Single

' Calculate oxygen from cations
temp1! = 0#
For i% = 1 To sample(1).LastChan%
If i% <> sample(1).OxygenChannel% Then
temp2! = ConvertElmToOxd(analysis.WtPercents!(i%), sample(1).Elsyms$(i%), sample(1).numcat%(i%), sample(1).numoxd%(i%))
temp1! = temp1! + (temp2! - analysis.WtPercents!(i%))
End If
Next i%

ConvertOxygenFromCations3! = temp1!
Exit Function

' Errors
ConvertOxygenFromCations3Error:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertOxygenFromCations3"
ierror = True
Exit Function

End Function

Function ConvertChargeBalance(lchan As Integer, atwts() As Single, wts() As Single, chrgs() As Single) As Single
' Calculate charge balance

ierror = False
On Error GoTo ConvertChargeBalanceError

Dim i As Integer
Dim temp As Single

Dim atoms(1 To MAXCHAN%) As Single  ' calculate from atwts and wts

' Convert to atomic percent
Call ConvertWeightToAtomic(lchan%, atwts!(), wts!(), atoms!())
If ierror Then Exit Function

' Calculate for composition
temp! = 0#
For i% = 1 To lchan%
temp! = temp! + atoms!(i%) * chrgs!(i%)
Next i%

ConvertChargeBalance! = temp!

Exit Function

' Errors
ConvertChargeBalanceError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertChargeBalance"
ierror = True
Exit Function

End Function

Sub ConvertWeightToFormula(analysis As TypeAnalysis, sample() As TypeSample)
' Routine to convert elemental weight percents to formula coefficients (analysis.WtPercents to analysis.Formulas)

ierror = False
On Error GoTo ConvertWeightToFormulaError

Dim i As Integer, ip As Integer
Dim temp As Single
Dim totalatoms As Single, TotalCations As Single

Dim atoms(1 To MAXCHAN%) As Single

' Zero arrays
For i% = 1 To sample(1).LastChan%
analysis.Formulas!(i%) = 0#
Next i%

' Check for insufficient total atoms
For i% = 1 To sample(1).LastChan%
If sample(1).AtomicWts!(i%) > 0# Then atoms!(i%) = analysis.WtPercents!(i%) / sample(1).AtomicWts!(i%)
totalatoms! = totalatoms! + atoms!(i%)
Next i%
If totalatoms! < 0.01 Then GoTo ConvertWeightToFormulaLowTotal
 
' Calculate formulas
If sample(1).FormulaElement$ <> vbNullString Then     ' normal formula calculation
ip% = IPOS1(sample(1).LastChan%, sample(1).FormulaElement$, sample(1).Elsyms$())
If ip% = 0 Then GoTo ConvertWeightToFormulaInvalidFormulaElement

' Check for insufficient formula basis element
If atoms!(ip%) < 0.01 Then GoTo ConvertWeightToFormulaNoBasis
temp! = sample(1).FormulaRatio! / atoms!(ip%)

For i% = 1 To sample(1).LastChan%
analysis.Formulas!(i%) = atoms!(i%) * temp!
Next i%

' Calculate sum of cations (normalize sum of cations to formula atoms))
Else
TotalCations! = 0#
For i% = 1 To sample(1).LastChan%
If sample(1).AtomicCharges!(i%) > 0# Then TotalCations! = TotalCations! + atoms!(i%)
Next i%
If TotalCations! < 0.01 Then GoTo ConvertWeightToFormulaLowTotalCation

' Normalize to total number of cations
temp! = sample(1).FormulaRatio! / TotalCations!
For i% = 1 To sample(1).LastChan%
analysis.Formulas!(i%) = atoms!(i%) * temp!
Next i%
End If

Exit Sub

' Errors
ConvertWeightToFormulaError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertWeightToFormula"
ierror = True
Exit Sub

ConvertWeightToFormulaLowTotal:
msg$ = TypeLoadString$(sample())
msg$ = "There is an insufficient total number of atoms (usually caused by low totals) to calculate atomic ratios for sample " & msg$ & ". "
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertWeightToFormula"
ierror = True
Exit Sub

ConvertWeightToFormulaInvalidFormulaElement:
msg$ = TypeLoadString$(sample())
msg$ = "Element " & sample(1).FormulaElement$ & " is an invalid formula element for sample " & msg$
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertWeightToFormula"
ierror = True
Exit Sub

ConvertWeightToFormulaLowTotalCation:
msg$ = TypeLoadString$(sample())
msg$ = "There is an insufficient concentration of the cation sum (usually caused by a very low total), to calculate a formula for sample " & msg$ & ". "
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertWeightToFormula"
ierror = True
Exit Sub

ConvertWeightToFormulaNoBasis:
msg$ = TypeLoadString$(sample())
msg$ = "There is an insufficient concentration of the formula basis element " & sample(1).Elsyup$(ip%) & " (usually caused by a very low concentration of the element), to calculate a formula for sample " & msg$ & ". "
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertWeightToFormula"
ierror = True
Exit Sub

End Sub

Sub ConvertCalculatePredictedDetection(tDetect() As Single, analysis As TypeAnalysis, sample() As TypeSample)
' Calculate predicted detection limits for average of sample

ierror = False
On Error GoTo ConvertCalculatePredictedDetectionError

Dim i As Integer
Dim temp1 As Single, temp2 As Single

Dim bgaverag As TypeAverage, bmaverag As TypeAverage
Dim onaverag As TypeAverage, hiaverag As TypeAverage, loaverag As TypeAverage

Dim bgcts(1 To MAXCHAN%) As Single

' Average background
Call MathArrayAverage(bgaverag, sample(1).BgdData!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub

' Average beam
If Not sample(1).CombinedConditionsFlag Then   ' use OnBeamData in case of aggregate intensity calculation (use average aggregate beam)
Call MathArrayAverage(bmaverag, sample(1).OnBeamData!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
Else                                           ' use OnBeamDataArray in case of aggregate intensity calculation (use average aggregate beam)
Call MathArrayAverage(bmaverag, sample(1).OnBeamDataArray!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
End If

' Average on peak count time
Call MathArrayAverage(onaverag, sample(1).OnTimeData!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub

' Average hi and lo peak count time
Call MathArrayAverage(hiaverag, sample(1).HiTimeData!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub
Call MathArrayAverage(loaverag, sample(1).LoTimeData!(), sample(1).Datarows%, sample(1).LastElm%, sample())
If ierror Then Exit Sub

' All intensities in cps/nA
For i% = 1 To sample(1).LastElm%
tDetect!(i%) = 0#

' Calculate unknown background raw counts
bgcts!(i%) = bgaverag.averags!(i%) / NominalBeam!   ' background intensity in cps/nA
bgcts!(i%) = bgcts!(i%) * bmaverag.averags!(i%) * (hiaverag.averags!(i%) + loaverag.averags!(i%))      ' total counts

If bgcts!(i%) >= 0# And analysis.StdAssignsBeams!(i%) <> 0# Then
temp1! = 3# * Sqr(bgcts!(i%)) * analysis.StdAssignsPercents!(i%)
temp2! = onaverag.averags!(i%) * bmaverag.averags!(i%) * analysis.StdAssignsCounts!(i%) / NominalBeam!
If temp2! <> 0# Then tDetect!(i%) = temp1! / temp2!
End If
Next i%

Exit Sub

' Errors
ConvertCalculatePredictedDetectionError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertCalculatePredictedDetection"
Close (Temp1FileNumber%)
ierror = True
Exit Sub

End Sub

Function ConvertElmToOxdRatio(chan As Integer, sample() As TypeSample) As Single
' Calculates the elemental to oxide ratio for the channel of a sample

ierror = False
On Error GoTo ConvertElmToOxdRatioError

Dim temp1 As Single, temp2 As Single

temp1! = sample(1).AtomicWts!(chan%) * sample(1).numcat%(chan%) + AllAtomicWts!(ATOMIC_NUM_OXYGEN%) * sample(1).numoxd%(chan%)
temp2! = sample(1).AtomicWts!(chan%) * sample(1).numcat%(chan%)
If temp2! = 0# Then GoTo ConvertElmToOxdRatioBadParameter
ConvertElmToOxdRatio! = temp1! / temp2!

Exit Function

' Errors
ConvertElmToOxdRatioError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertElmToOxdRatio"
ierror = True
Exit Function

ConvertElmToOxdRatioBadParameter:
msg$ = "The atomic weight or number of cations is zero for channel " & Format$(chan%) & " in sample " & sample(1).Name$ & ". "
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertElmToOxdRatio"
ierror = True
Exit Function

End Function

Sub ConvertFerrousFerricRatioFromComposition(nelements As Integer, AtomicNumbers() As Integer, AtomicWeights() As Single, ElementalWeightFractions() As Single, NumCats() As Integer, NumOxds() As Integer, OxideProportions() As Single, DisableQuantFlags() As Integer, MineralCations As Single, MineralOxygens As Single, FerricToTotalIronRatio As Single, FerricOxygen As Single, Fe_as_FeO As Single, Fe_as_Fe2O3 As Single)
' Procedure to calculate ferrous/ferric ratio from concentrations and specified cations and oxygens (assumes only Fe is multi-valent)
'  nelements = number of elements in composition
'  AtomicNumbers() = atomic numbers of each element
'  AtomicWeights() = atomic weight of each element
'  ElementalWeightFractions() = elemental weight fractions for each element
'  OxideProportions() = oxide conversion factors for all elements (force iron stoichiometry to assume FeO stoichiometry = 0.286497)
'  DisableQuantFlags%() = diable quant flags for each element
'  MineralCations = total number of cations in mineral formula
'  MineralOxygens = total number of oxygens in mineral formula
'  NumCats() - number of cations in oxide formula
'  NumOxds() = number of oxygens in oxide formula
'  FerricToTotalIronRatio = ratio of ferric iron to total iron
'  FerricOxygen = excess oxygen in weight%
'  Fe_as_FeO = calculated FeO in weight%
'  Fe_as_Fe2O3 = calculated Fe2O3 in weight%

ierror = False
On Error GoTo ConvertFerrousFerricRatioFromCompositionError

Const FeO_to_Fe2O3! = 1.11134

Dim ip As Integer
Dim n As Integer

Dim summoles As Single
Dim sumnorms As Single
Dim ferrousmoles As Single
Dim ferricmoles As Single

ReDim atoms(1 To MAXCHAN%) As Single
ReDim oxides(1 To MAXCHAN%) As Single
ReDim moles(1 To MAXCHAN%) As Single
ReDim normcats(1 To MAXCHAN%) As Single

' Determine Fe channel
ip% = IPOS2DQ%(nelements%, ATOMIC_NUM_IRON%, AtomicNumbers%(), DisableQuantFlags%())
If ip% = 0 Then GoTo ConvertFerrousFerricRatioFromCompositionNoIron

' Check that Fe oxide stoichiometry is 1 : 1
If NumCats%(ip%) <> 1 Or NumOxds%(ip%) <> 1 Then GoTo ConvertFerrousFerricRatioFromCompositionNotFeO

' Calculate oxide weight percents (force iron to FeO)
For n% = 1 To nelements%

If AtomicNumbers%(n%) <> ATOMIC_NUM_IRON% Then
oxides!(n%) = ElementalWeightFractions!(n%) + ElementalWeightFractions!(n%) * OxideProportions!(n%)      ' convert all elements except Fe to oxides
Else
oxides!(n%) = ElementalWeightFractions!(n%) + ElementalWeightFractions!(n%) * FeO_OXIDE_PROPORTION!      ' force convert Fe to FeO to be consistent below
End If

Next n%

' Calculate moles of each oxide
summoles! = 0
For n% = 1 To nelements%

' Calculate moles of each cation
If AtomicNumbers%(n%) <> ATOMIC_NUM_OXYGEN% Then

If AtomicNumbers%(n%) <> ATOMIC_NUM_IRON% Then
moles!(n%) = (100# * oxides!(n%) * NumCats%(n%)) / (NumCats%(n%) * AtomicWeights!(n%) + NumOxds%(n%) * AllAtomicWts!(ATOMIC_NUM_OXYGEN%))        ' convert to formula moles
Else
moles!(n%) = (100# * oxides!(n%)) / (AtomicWeights!(n%) + AllAtomicWts!(ATOMIC_NUM_OXYGEN%))                                    ' convert based on Fe : O moles
End If

' Calculate the sum of the cations
summoles! = summoles + moles!(n%)

' Calculate moles of oxygen
Else
moles!(n%) = (100# * oxides!(n%)) / AtomicWeights!(n%)
End If

Next n%

' Normalize to the number of total cations
sumnorms! = 0#
For n% = 1 To nelements%
normcats!(n%) = moles!(n%) / summoles! * MineralCations!

' Calculate sum of normalized moles for cations only
If AtomicNumbers%(n%) <> ATOMIC_NUM_OXYGEN% Then
sumnorms! = sumnorms! + normcats!(n%)
End If

Next n%

If VerboseMode Then
msg$ = vbCrLf & "ELEMENT "
For n% = 1 To nelements%
msg$ = msg$ & Format$(Symup$(AtomicNumbers%(n%)), a80$)
Next n%
Call IOWriteLog(msg$)

msg$ = "OXIDES  "
For n% = 1 To nelements%
msg$ = msg$ & Format$(Format$(100# * oxides!(n%), f84), a80$)
Next n%
Call IOWriteLog(msg$)

msg$ = "MOLES   "
For n% = 1 To nelements%
msg$ = msg$ & Format$(Format$(moles!(n%), f84), a80$)
Next n%
msg$ = msg$ & Space$(8) & Format$("SUM CAT=", a80$) & Format$(Format$(summoles!, f84), a80$)

Call IOWriteLog(msg$)
msg$ = "NORMCAT "
For n% = 1 To nelements%
msg$ = msg$ & Format$(Format$(normcats!(n%), f84), a80$)
Next n%
msg$ = msg$ & Space$(8) & Format$("SUM CAT=", a80$) & Format$(Format$(sumnorms!, f84), a80$)
Call IOWriteLog(msg$)
End If

' Check charge balance
If normcats!(nelements) <= MineralOxygens! Then

If 2 * (MineralOxygens! - normcats!(nelements%)) <= normcats!(ip%) Then
ferricmoles! = 2 * (MineralOxygens! - normcats!(nelements%))
ferrousmoles! = normcats!(ip%) - ferricmoles!
Else
ferricmoles! = normcats!(ip%)
ferrousmoles! = 0
End If

Else
ferricmoles! = 0#
ferrousmoles! = moles!(ip%)
End If

' Calculate ferric to total iron ratio
FerricToTotalIronRatio! = ferricmoles! / (ferricmoles + ferrousmoles!)

' Calculate FeO and Fe2O3 in weight percent
Fe_as_FeO! = (1 - FerricToTotalIronRatio!) * 100# * oxides!(ip%)
Fe_as_Fe2O3! = (100# * oxides!(ip%) - Fe_as_FeO!) / (2 * (AllAtomicWts!(ATOMIC_NUM_IRON%) + AllAtomicWts!(ATOMIC_NUM_OXYGEN%)) / (2 * AllAtomicWts!(ATOMIC_NUM_IRON%) + 3 * AllAtomicWts!(ATOMIC_NUM_OXYGEN%)))

FerricOxygen! = Fe_as_Fe2O3! * (1 - (2 * (AllAtomicWts!(ATOMIC_NUM_IRON%) + AllAtomicWts!(ATOMIC_NUM_OXYGEN%)) / (2 * AllAtomicWts!(ATOMIC_NUM_IRON%) + 3 * AllAtomicWts!(ATOMIC_NUM_OXYGEN%))))

If DebugMode Then
Call IOWriteLog(vbNullString)
msg$ = "FerricIronToTotalIronRatio: " & Format$(Format$(FerricToTotalIronRatio!, f84), a80$)
Call IOWriteLog(msg$)
msg$ = "Ferrous Oxide (FeO):        " & Format$(Format$(Fe_as_FeO!, f84), a80$)
Call IOWriteLog(msg$)
msg$ = "Ferric Oxide (Fe2O3):       " & Format$(Format$(Fe_as_Fe2O3!, f84), a80$)
Call IOWriteLog(msg$)
msg$ = "Excess Oxygen from Fe2O3:   " & Format$(Format$(FerricOxygen!, f84), a80$)
Call IOWriteLog(msg$)
End If

Exit Sub

' Errors
ConvertFerrousFerricRatioFromCompositionError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertFerrousFerricRatioFromComposition"
ierror = True
Exit Sub

ConvertFerrousFerricRatioFromCompositionNoIron:
msg$ = "No iron is present in the element list. Cannot calculate a ferrous/ferric ratio."
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertFerrousFerricRatioFromComposition"
ierror = True
Exit Sub

ConvertFerrousFerricRatioFromCompositionNotFeO:
msg$ = "Cannot calculate a ferrous/ferric ratio. Iron stoichiometry must be specified as FeO (Cations=1, Oxygens=1) in the Elements/Cations dialog."
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertFerrousFerricRatioFromComposition"
ierror = True
Exit Sub

End Sub


