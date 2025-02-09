Attribute VB_Name = "CodeCONVERT"
' (c) Copyright 1995-2025 by John J. Donovan
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

Function ConvertWeightToAtom2(lchan As Integer, chan As Integer, wts() As Single, atwts() As Single, syms() As String) As Single
' Calculate weight percent to atomic percent for a single element (uses passed atomic weights)

ierror = False
On Error GoTo ConvertWeightToAtom2Error

Dim i As Integer, ip As Integer
Dim sum As Single, temp As Single

' Sum the atoms of the elemental weight percents
sum = 0#
For i% = 1 To lchan%
sum! = sum! + wts!(i%) / atwts!(i%)
Next i%

' Check for bad sum
If sum! <= 0# Then
msg$ = "Bad atomic sum = " & Format$(sum!)
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertWeightToAtom2"
ierror = True
Exit Function
End If

' Do weight percent to atomic percent conversion
ConvertWeightToAtom2! = 0#
temp! = wts!(chan%) / atwts!(chan%)
ConvertWeightToAtom2! = 100# * temp! / sum!

Exit Function

' Errors
ConvertWeightToAtom2Error:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertWeightToAtom2"
ierror = True
Exit Function

End Function

Sub ConvertWeightToZFraction(lchan As Integer, atnums() As Single, atwts() As Single, conc() As Single, zfracs() As Single)
' Convert from weight fraction to Z fraction (assume Z fraction exponent is 1.0) (obsolete function)

ierror = False
On Error GoTo ConvertWeightToZFractionError

Dim i As Integer
Dim sum As Single

' Calculate sum
sum! = 0#
For i% = 1 To lchan%
sum! = sum! + atnums!(i%) * conc!(i%) / atwts!(i%)
Next i%
If sum! = 0# Then GoTo ConvertWeightToZFractionZeroSum

For i% = 1 To lchan%
If sum! <> 0# Then
zfracs!(i%) = atnums!(i%) * (conc!(i%) / atwts!(i%)) / sum!
Else
zfracs!(i%) = 0#
End If
Next i%

Exit Sub

' Errors
ConvertWeightToZFractionError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertWeightToZFraction"
ierror = True
Exit Sub

ConvertWeightToZFractionZeroSum:
msg$ = "Sum of concentrations is zero"
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertWeightToZFraction"
ierror = True
Exit Sub

End Sub

Sub ConvertWeightToZFractionBSE(lchan As Integer, exponent As Single, atnums() As Integer, atwts() As Single, conc() As Single, keV() As Single, zfracs() As Single)
' Convert from weight fraction to Z fraction (based on passed exponent) for electron backscatter corrections

ierror = False
On Error GoTo ConvertWeightToZFractionBSEError

Dim i As Integer
Dim sum As Single
Dim texponent(1 To MAXCHAN%) As Single

' Load passed exponent or calculate based on electron beam energy if exponent is zero
For i% = 1 To lchan%
If exponent! = 0# Then
If keV!(i%) = 0# Then GoTo ConvertWeightToZFractionBSEZerokeV
texponent!(i%) = ConvertCalculateZFractionExponentBSE(keV!(i%))
If ierror Then Exit Sub

Else
texponent!(i%) = exponent!
End If
Next i%

' Calculate sum
sum! = 0#
For i% = 1 To lchan%
sum! = sum! + (atnums%(i%) ^ texponent!(i%)) * conc!(i%) / atwts!(i%)
Next i%
If sum! = 0# Then GoTo ConvertWeightToZFractionBSEZeroSum

' Calculate Z fractions
For i% = 1 To lchan%
If sum! <> 0# Then
zfracs!(i%) = (atnums%(i%) ^ texponent!(i%)) * (conc!(i%) / atwts!(i%)) / sum!
Else
zfracs!(i%) = 0#
End If
Next i%

Exit Sub

' Errors
ConvertWeightToZFractionBSEError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "ConvertWeightToZFractionBSE"
ierror = True
Exit Sub

ConvertWeightToZFractionBSEZerokeV:
Screen.MousePointer = vbDefault
msg$ = "Variable zbar calculation was passed a zero keV value. This error should not occur, please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertWeightToZFractionBSE"
ierror = True
Exit Sub

ConvertWeightToZFractionBSEZeroSum:
Screen.MousePointer = vbDefault
msg$ = "Sum of concentrations is zero"
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertWeightToZFractionBSE"
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

Function ConvertSulfurToOxygen(lchan As Integer, syms() As String, dqs() As Integer, conc() As Single, charges() As Single) As Single
' Calculate the equivalent oxygen based on the sulfur concentration and charge valence
'  lchan = number of channels
'  syms() = symbol (element) array
'  dqs() = disable quant flag array
'  conc() = concentration of element array
'  charges() = charge valences of element array

ierror = False
On Error GoTo ConvertSulfurToOxygenError

Dim StoO As Single

Dim ips As Integer, ipo As Integer
Dim temp As Single

ConvertSulfurToOxygen! = 0#

' Find index for sulfur and oxygen
ips% = IPOS1DQ(lchan%, Symlo$(ATOMIC_NUM_SULFUR%), syms$(), dqs%())
ipo% = IPOS1DQ(lchan%, Symlo$(ATOMIC_NUM_OXYGEN%), syms$(), dqs%())
If ips% = 0 Then Exit Function
If charges!(ips%) = 0# Then Exit Function

' Calculate stoiometric correction factor
If ipo% <> 0 Then
StoO! = AllAtomicWts!(ATOMIC_NUM_OXYGEN%) / AllAtomicWts!(ATOMIC_NUM_SULFUR%) * charges!(ipo%) / charges!(ips%)
Else
StoO! = AllAtomicWts!(ATOMIC_NUM_OXYGEN%) / AllAtomicWts!(ATOMIC_NUM_SULFUR%) * -2# / charges!(ips%)
End If

' Check for sulfur charge valence less than zero (negative charge valence replaces oxygen equivalent)
If charges!(ips%) < 0# Then
temp! = 0#
If ips% > 0 Then
temp! = conc!(ips%) * StoO!
End If

' Sulfur valence is positive
Else
temp! = 0#
End If

ConvertSulfurToOxygen! = temp!

Exit Function

' Errors
ConvertSulfurToOxygenError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertSulfurToOxygen"
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

If bgcts!(i%) >= 0# And NominalBeam! >= 0# Then
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
' Procedure to calculate ferrous/ferric ratio from concentrations and specified cations and oxygens (assumes only Fe is multi-valent) from Droop, 1987 and Locock spreadsheet
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

Sub ConvertFerrousFerricRatioFromComposition2(nelements As Integer, AtomicNumbers() As Integer, AtomicWeights() As Single, ElementalWeightFractions() As Single, NumCats() As Integer, NumOxds() As Integer, OxideProportions() As Single, DisableQuantFlags() As Integer, MineralCations As Single, MineralOxygens As Single, FerricToTotalIronRatio As Single, FerricOxygen As Single, Fe_as_FeO As Single, Fe_as_Fe2O3 As Single, Droop_option_for_amphibole As Integer)
' Procedure to calculate ferrous/ferric ratio from concentrations and specified cations and oxygens (assumes only Fe is multi-valent) from Droop, 1987 and Locock spreadsheet
' New code containing Aurelien Moy's modifications to handle amphibole compositions
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
'  Droop_option_for_amphibole
'    0 = general formula, using original Droop method
'    1 = Formula calculated on the basis of 23 oxygens anhydrous, assuming a total of 15 cations exclusive of Na and K (suitable for Fe-Mg-amphiboles and coexisting calcic amphiboles).
'    2 = Formula calculated on the basis of 23 oxygens anhydrous, assuming a total of 13 cations exclusive of Ca, Na and K (suitable for many calcic amphiboles).

ierror = False
On Error GoTo ConvertFerrousFerricRatioFromComposition2Error

Dim ip As Integer
Dim n As Integer

Dim summoles As Single
Dim ferrousmoles As Single
Dim ferricmoles As Single

ReDim oxides(1 To MAXCHAN%) As Single
ReDim moles(1 To MAXCHAN%) As Single

Dim atomic_fraction(1 To MAXCHAN%) As Single                            ' atomic fraction o the different elements
Dim sum_O_from_atomic_fraction As Single                                ' total number of O atoms in atomic fraction
Dim formula_with_mineral_oxygens_atoms_of_O(1 To MAXCHAN%) As Single    ' atomic fractions of the elements in the normalized atomic formula (normalized to a given number of O atoms given by MineralOxygens! (e.g., 23 for amphibole))
Dim total_cations As Single                                             ' total number of cations in the normalized atomic formula (normalized to MineralOxyens!)
Dim Fe3_atomic_formula As Single                                        ' number of Fe3+ atoms in the normalized formula (normalized to a given number of cations MineralCations!)
Dim Fe2_atomic_formula As Single                                        ' number of Fe2+ atoms in the normalized formula (normalized to a given number of cations MineralCations!)

' Determine Fe channel
ip% = IPOS2DQ%(nelements%, ATOMIC_NUM_IRON%, AtomicNumbers%(), DisableQuantFlags%())
If ip% = 0 Then GoTo ConvertFerrousFerricRatioFromComposition2NoIron

' Check that Fe oxide stoichiometry is 1 : 1
If NumCats%(ip%) <> 1 Or NumOxds%(ip%) <> 1 Then GoTo ConvertFerrousFerricRatioFromComposition2NotFeO

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

' Calculate the atomic raction of the different elements
For n% = 1 To nelements%
    atomic_fraction!(n%) = moles!(n%) / summoles!
Next n%

' Calculates the numner of atomic oxygen. Excludes the anions O, F and Cl as they do not contribute.
sum_O_from_atomic_fraction! = 0
For n% = 1 To nelements%
    If AtomicNumbers%(n%) <> ATOMIC_NUM_OXYGEN% And AtomicNumbers%(n%) <> ATOMIC_NUM_FLUORINE% And AtomicNumbers%(n%) <> ATOMIC_NUM_CHLORINE% Then
        sum_O_from_atomic_fraction! = sum_O_from_atomic_fraction! + atomic_fraction!(n%) * NumOxds%(n%) / NumCats%(n%)
    End If
Next n%

' Calculates the atomic formula based on MineralOxygens! atoms of O. Also calculates the total number of cations in the normalized atomic formula.
total_cations! = 0
For n% = 1 To nelements%
    formula_with_mineral_oxygens_atoms_of_O!(n%) = atomic_fraction!(n%) / sum_O_from_atomic_fraction! * MineralOxygens!
    If Droop_option_for_amphibole = 0 Then
        If AtomicNumbers%(n%) <> ATOMIC_NUM_OXYGEN% And AtomicNumbers%(n%) <> ATOMIC_NUM_FLUORINE% And AtomicNumbers%(n%) <> ATOMIC_NUM_CHLORINE% Then
            total_cations! = total_cations! + formula_with_mineral_oxygens_atoms_of_O!(n%)
        End If
    ElseIf Droop_option_for_amphibole = 1 Then
        If AtomicNumbers%(n%) <> ATOMIC_NUM_OXYGEN% And AtomicNumbers%(n%) <> ATOMIC_NUM_FLUORINE% And AtomicNumbers%(n%) <> ATOMIC_NUM_CHLORINE% And AtomicNumbers%(n%) <> ATOMIC_NUM_SODIUM% And AtomicNumbers%(n%) <> ATOMIC_NUM_POTASSIUM% Then
            total_cations! = total_cations! + formula_with_mineral_oxygens_atoms_of_O!(n%)
        End If
    ElseIf Droop_option_for_amphibole = 2 Then
        If AtomicNumbers%(n%) <> ATOMIC_NUM_OXYGEN% And AtomicNumbers%(n%) <> ATOMIC_NUM_FLUORINE% And AtomicNumbers%(n%) <> ATOMIC_NUM_CHLORINE% And AtomicNumbers%(n%) <> ATOMIC_NUM_SODIUM% And AtomicNumbers%(n%) <> ATOMIC_NUM_POTASSIUM% And AtomicNumbers%(n%) <> ATOMIC_NUM_CALCIUM% Then
            total_cations! = total_cations! + formula_with_mineral_oxygens_atoms_of_O!(n%)
        End If
    End If
    
Next n%

' Handle special cases for amhiboles. Make sure that MineralOxygens = 23 and that MineralCations is either 15 or 13.
If Droop_option_for_amphibole = 1 Then
    If MineralOxygens! <> 23 Or MineralCations! <> 15 Then
        MineralOxygens! = 23
        MineralCations! = 15
        Call IOWriteLogRichText(vbCrLf & "Warning in ConvertFerrousFerricRatioFromComposition2: Number of oxygens changed to 23 and number of cations changed to 15", vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
    End If
ElseIf Droop_option_for_amphibole = 2 Then
    If MineralOxygens! <> 23 Or MineralCations! <> 13 Then
        MineralOxygens! = 23
        MineralCations! = 13
        Call IOWriteLogRichText(vbCrLf & "Warning in ConvertFerrousFerricRatioFromComposition2: Number of oxygens changed to 23 and number of cations changed to 13", vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
    End If
End If

' Calculates Fe3+ and Fe2+ for a given number of cations (given by MineralCations!)
If MineralCations! <= total_cations! Then
    Fe3_atomic_formula! = 2 * MineralOxygens! * (1 - MineralCations! / total_cations!)      ' Droop's (1987) formula (Fe3+ = 2*X*(1-T/S))
    If Fe3_atomic_formula! < formula_with_mineral_oxygens_atoms_of_O!(ip%) * MineralCations! / total_cations! Then
        Fe2_atomic_formula! = formula_with_mineral_oxygens_atoms_of_O!(ip%) * MineralCations! / total_cations! - Fe3_atomic_formula!
    Else
        Fe3_atomic_formula! = formula_with_mineral_oxygens_atoms_of_O!(ip%)
        Fe2_atomic_formula! = 0
    End If
Else
    Fe3_atomic_formula! = 0
    Fe2_atomic_formula! = formula_with_mineral_oxygens_atoms_of_O!(ip%)
End If

' Calculates the ratio Fe3+ / (Fe3+ + Fe2+)
FerricToTotalIronRatio! = Fe3_atomic_formula! / (Fe3_atomic_formula! + Fe2_atomic_formula!)

' Convert atomic Fe3+ and Fe2+ to oxides. Also calculates FerricOxygen!
Fe_as_FeO! = (1 - FerricToTotalIronRatio!) * 100# * oxides!(ip%)
Fe_as_Fe2O3! = (100# * oxides!(ip%) - Fe_as_FeO!) / (2 * (AllAtomicWts!(ATOMIC_NUM_IRON%) + AllAtomicWts!(ATOMIC_NUM_OXYGEN%)) / (2 * AllAtomicWts!(ATOMIC_NUM_IRON%) + 3 * AllAtomicWts!(ATOMIC_NUM_OXYGEN%)))
FerricOxygen! = Fe_as_Fe2O3! * (1 - (2 * (AllAtomicWts!(ATOMIC_NUM_IRON%) + AllAtomicWts!(ATOMIC_NUM_OXYGEN%)) / (2 * AllAtomicWts!(ATOMIC_NUM_IRON%) + 3 * AllAtomicWts!(ATOMIC_NUM_OXYGEN%))))

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

msg$ = "ATFRAC  "
For n% = 1 To nelements%
msg$ = msg$ & Format$(Format$(atomic_fraction!(n%), f84), a80$)
Next n%
End If

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
ConvertFerrousFerricRatioFromComposition2Error:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertFerrousFerricRatioFromComposition2"
ierror = True
Exit Sub

ConvertFerrousFerricRatioFromComposition2NoIron:
msg$ = "No iron is present in the element list. Cannot calculate a ferrous/ferric ratio."
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertFerrousFerricRatioFromComposition2"
ierror = True
Exit Sub

ConvertFerrousFerricRatioFromComposition2NotFeO:
msg$ = "Cannot calculate a ferrous/ferric ratio. Iron stoichiometry must be specified as FeO (Cations=1, Oxygens=1) in the Elements/Cations dialog."
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertFerrousFerricRatioFromComposition2"
ierror = True
Exit Sub

End Sub

Sub ConvertFerrousFerricRatioFromComposition3(nelements As Integer, AtomicNumbers() As Integer, AtomicWeights() As Single, ElementalWeightFractions() As Single, NumCats() As Integer, NumOxds() As Integer, OxideProportions() As Single, DisableQuantFlags() As Integer, MineralCations As Single, MineralOxygens As Single, FerricToTotalIronRatio As Single, FerricOxygen As Single, Fe_as_FeO As Single, Fe_as_Fe2O3 As Single, Droop_option_for_amphibole As Integer)
' Procedure to calculate ferrous/ferric ratio from concentrations (for amphiboles) and specified cations and oxygens (assumes only Fe and Mn are multi-valent) from Locock spreadsheet
'  nelements = number of elements in composition
'  AtomicNumbers() = atomic numbers of each element
'  AtomicWeights() = atomic weight of each element
'  ElementalWeightFractions() = elemental weight fractions for each element
'  OxideProportions() = oxide conversion factors for all elements (force iron stoichiometry to assume FeO stoichiometry = 0.286497)
'  NumCats() - number of cations in oxide formula
'  NumOxds() = number of oxygens in oxide formula
'  DisableQuantFlags%() = diable quant flags for each element
'  MineralCations = total number of cations in mineral formula
'  MineralOxygens = total number of oxygens in mineral formula

'  FerricToTotalIronRatio = ratio of ferric iron to total iron
'  FerricOxygen = excess oxygen in weight%
'  Fe_as_FeO = calculated FeO in weight%
'  Fe_as_Fe2O3 = calculated Fe2O3 in weight%
'  Droop_option_for_amphibole
'    0 = general formula, using original Droop method (not called)
'    1 = sodic amphiboles
'    2 = calcic amphiboles
'    3 = Na-Ca amphiboles
'    4 = Fe-Mg-Mn amphiboles
'    5 = Oxo amphiboles
'    6 = Li amphiboles

ierror = False
On Error GoTo ConvertFerrousFerricRatioFromComposition3Error

Dim waterspecified As Boolean
Dim tipresent As Boolean
Dim ip As Integer
Dim n As Integer

' Declare the variables used by the AmphiboleCalculationLoop function to return the results
Dim Fe3overSumFe As Single
Dim Mn3overSumMn As Single
Dim finalWtPercentValues_H2O As Single

' Determine Fe channel
ip% = IPOS2DQ%(nelements%, ATOMIC_NUM_IRON%, AtomicNumbers%(), DisableQuantFlags%())
If ip% = 0 Then GoTo ConvertFerrousFerricRatioFromComposition3NoIron

' Check that Fe oxide stoichiometry is 1 : 1
If NumCats%(ip%) <> 1 Or NumOxds%(ip%) <> 1 Then GoTo ConvertFerrousFerricRatioFromComposition3NotFeO

' Determine H2O channel and set flag as suggested by Locock
ip% = IPOS2DQ%(nelements%, ATOMIC_NUM_HYDROGEN%, AtomicNumbers%(), DisableQuantFlags%())
If ip% > 0 Then
If ElementalWeightFractions!(ip%) * 100# > NOT_ANALYZED_VALUE_SINGLE! Then waterspecified = True
End If

' Check for TiO2 greater than 1 wt% (or Ti > 0.6 wt%) as suggested by Locock
ip% = IPOS2DQ%(nelements%, ATOMIC_NUM_TITANIUM%, AtomicNumbers%(), DisableQuantFlags%())
If ip% > 0 Then
If ElementalWeightFractions!(ip%) * 100# > 0.6 Then tipresent = True
End If

Fe3overSumFe! = 0#
Mn3overSumMn! = 0#
finalWtPercentValues_H2O! = 0#

' Declaration of the 8 different options. This is handled in the ferrous/ferric interface in Calculation Options.
Dim ORTHORHOMBIC As Integer
Dim USE_INITIAL_M3_OVER_SUM_M As Integer
Dim ESTIMATEOH2_2TI As Integer
Dim REQUIRE_INITIAL_H2O As Integer

Dim REQUIRE_SUM_SI_TO_CA_LE_15 As Integer
Dim REQUIRE_SUM_SI_TO_MG_GE_13 As Integer
Dim REQUIRE_SUM_SI_TO_NA_GE_15 As Integer
Dim REQUIRE_SUM_SI_TO_K_GE_15_5 As Integer

' Sodic amphibole
If Droop_option_for_amphibole% = 1 Then
ORTHORHOMBIC% = 0
USE_INITIAL_M3_OVER_SUM_M% = 0
If tipresent Then ESTIMATEOH2_2TI% = 1
If waterspecified Then REQUIRE_INITIAL_H2O% = 1

REQUIRE_SUM_SI_TO_CA_LE_15% = 0
REQUIRE_SUM_SI_TO_MG_GE_13% = 1
REQUIRE_SUM_SI_TO_NA_GE_15% = 0
REQUIRE_SUM_SI_TO_K_GE_15_5% = 1

' Calcic amphibole
ElseIf Droop_option_for_amphibole% = 2 Then
ORTHORHOMBIC% = 0
USE_INITIAL_M3_OVER_SUM_M% = 0
If tipresent Then ESTIMATEOH2_2TI% = 1
If waterspecified Then REQUIRE_INITIAL_H2O% = 1

REQUIRE_SUM_SI_TO_CA_LE_15% = 1
REQUIRE_SUM_SI_TO_MG_GE_13% = 1
REQUIRE_SUM_SI_TO_NA_GE_15% = 0
REQUIRE_SUM_SI_TO_K_GE_15_5% = 0

' Na-Ca amphibole
ElseIf Droop_option_for_amphibole% = 3 Then
ORTHORHOMBIC% = 0
USE_INITIAL_M3_OVER_SUM_M% = 0
If tipresent Then ESTIMATEOH2_2TI% = 1
If waterspecified Then REQUIRE_INITIAL_H2O% = 1

REQUIRE_SUM_SI_TO_CA_LE_15% = 1
REQUIRE_SUM_SI_TO_MG_GE_13% = 0
REQUIRE_SUM_SI_TO_NA_GE_15% = 1
REQUIRE_SUM_SI_TO_K_GE_15_5% = 1

' Fe-Mg-Mn amphibole
ElseIf Droop_option_for_amphibole% = 4 Then
ORTHORHOMBIC% = 0
USE_INITIAL_M3_OVER_SUM_M% = 0
If tipresent Then ESTIMATEOH2_2TI% = 1
If waterspecified Then REQUIRE_INITIAL_H2O% = 1

REQUIRE_SUM_SI_TO_CA_LE_15% = 1
REQUIRE_SUM_SI_TO_MG_GE_13% = 0
REQUIRE_SUM_SI_TO_NA_GE_15% = 1
REQUIRE_SUM_SI_TO_K_GE_15_5% = 0

' Oxo amphibole
ElseIf Droop_option_for_amphibole% = 5 Then
ORTHORHOMBIC% = 0
USE_INITIAL_M3_OVER_SUM_M% = 0
If tipresent Then ESTIMATEOH2_2TI% = 1
If waterspecified Then REQUIRE_INITIAL_H2O% = 1

REQUIRE_SUM_SI_TO_CA_LE_15% = 1
REQUIRE_SUM_SI_TO_MG_GE_13% = 0
REQUIRE_SUM_SI_TO_NA_GE_15% = 0
REQUIRE_SUM_SI_TO_K_GE_15_5% = 0

' Li amphibole
ElseIf Droop_option_for_amphibole% = 6 Then
ORTHORHOMBIC% = 0
USE_INITIAL_M3_OVER_SUM_M% = 0
If tipresent Then ESTIMATEOH2_2TI% = 1
If waterspecified Then REQUIRE_INITIAL_H2O% = 1

REQUIRE_SUM_SI_TO_CA_LE_15% = 0
REQUIRE_SUM_SI_TO_MG_GE_13% = 0
REQUIRE_SUM_SI_TO_NA_GE_15% = 1
REQUIRE_SUM_SI_TO_K_GE_15_5% = 0
End If

Dim options_from_Locock_spreadsheet(1 To 8) As Integer
options_from_Locock_spreadsheet%(1) = ORTHORHOMBIC%
options_from_Locock_spreadsheet%(2) = USE_INITIAL_M3_OVER_SUM_M%
options_from_Locock_spreadsheet%(3) = ESTIMATEOH2_2TI%
options_from_Locock_spreadsheet%(4) = REQUIRE_INITIAL_H2O%
options_from_Locock_spreadsheet%(5) = REQUIRE_SUM_SI_TO_CA_LE_15%
options_from_Locock_spreadsheet%(6) = REQUIRE_SUM_SI_TO_MG_GE_13%
options_from_Locock_spreadsheet%(7) = REQUIRE_SUM_SI_TO_NA_GE_15%
options_from_Locock_spreadsheet%(8) = REQUIRE_SUM_SI_TO_K_GE_15_5%

' Call the AmphiboleCalculationLoop procedure using 7 passes
For n% = 1 To 7
    Call ConvertAmphiboleCalculationLoop(nelements, AtomicNumbers(), AtomicWeights(), ElementalWeightFractions(), NumCats(), NumOxds(), OxideProportions(), DisableQuantFlags(), n%, Fe3overSumFe, Mn3overSumMn, finalWtPercentValues_H2O, options_from_Locock_spreadsheet())
    If ierror Then Exit Sub
Next n%

' Convert the ferric to total iron ratio into FeO and Fe2O3 wt%. Also does the same with Mn.
Dim finalWtPercentValues_MnO As Single
Dim finalWtPercentValues_Mn2O3 As Single
Dim finalWtPercentValues_FeO As Single
Dim finalWtPercentValues_Fe2O3 As Single
Dim finalWtPercentValues_Total As Single

' Determine Mn channel
ip% = IPOS2DQ%(nelements%, 25, AtomicNumbers%(), DisableQuantFlags%())
If ip% = 0 Then
    finalWtPercentValues_MnO! = 0#
    finalWtPercentValues_Mn2O3! = 0#
Else
    finalWtPercentValues_MnO! = Round((1 - Mn3overSumMn!) * (ElementalWeightFractions!(ip%) + ElementalWeightFractions!(ip%) * OxideProportions!(ip%) + 0# * 0.89865734954961) * 100#, 3)
    finalWtPercentValues_Mn2O3! = Round((Mn3overSumMn! / 0.89865734954961) * (ElementalWeightFractions!(ip%) + ElementalWeightFractions!(ip%) * OxideProportions!(ip%) + 0# * 0.89865734954961) * 100#, 3)
End If

' Determine Fe channel
ip% = IPOS2DQ%(nelements%, 26, AtomicNumbers%(), DisableQuantFlags%())
If ip% = 0 Then
    finalWtPercentValues_FeO! = 0#
    finalWtPercentValues_Fe2O3! = 0#
Else
    finalWtPercentValues_FeO! = Round((1 - Fe3overSumFe!) * (ElementalWeightFractions!(ip%) + ElementalWeightFractions!(ip%) * OxideProportions!(ip%) + 0# * 0.899808502) * 100#, 3)
    finalWtPercentValues_Fe2O3! = Round((Fe3overSumFe! / 0.899808502) * (ElementalWeightFractions!(ip%) + ElementalWeightFractions!(ip%) * OxideProportions!(ip%) + 0# * 0.899808502) * 100#, 3)
End If

' Calculate the H2O content in wt%
finalWtPercentValues_H2O! = finalWtPercentValues_H2O! * 100#

' Calculates FerricToTotalIronRatio, Fe_as_FeO, Fe_as_Fe2O3 and FerricOxygen
FerricToTotalIronRatio! = Fe3overSumFe!
Fe_as_FeO! = finalWtPercentValues_FeO!
Fe_as_Fe2O3! = finalWtPercentValues_Fe2O3!
FerricOxygen! = Fe_as_Fe2O3! * (1 - (2 * (AllAtomicWts!(ATOMIC_NUM_IRON%) + AllAtomicWts!(ATOMIC_NUM_OXYGEN%)) / (2 * AllAtomicWts!(ATOMIC_NUM_IRON%) + 3 * AllAtomicWts!(ATOMIC_NUM_OXYGEN%))))

If VerboseMode Then
msg$ = vbCrLf & "ELEMENT "
For n% = 1 To nelements%
msg$ = msg$ & Format$(Symup$(AtomicNumbers%(n%)), a80$)
Next n%
Call IOWriteLog(msg$)
End If

If DebugMode Then
Call IOWriteLog(vbNullString)
msg$ = "ORTHORHOMBIC: " & Format$(Format$(ORTHORHOMBIC%, f84), a80$)
Call IOWriteLog(msg$)
msg$ = "USE_INITIAL_M3_OVER_SUM_M: " & Format$(Format$(USE_INITIAL_M3_OVER_SUM_M%, f84), a80$)
Call IOWriteLog(msg$)
msg$ = "ESTIMATEOH2_2TI: " & Format$(Format$(ESTIMATEOH2_2TI%, f84), a80$)
Call IOWriteLog(msg$)
msg$ = "REQUIRE_INITIAL_H2O: " & Format$(Format$(REQUIRE_INITIAL_H2O%, f84), a80$)
Call IOWriteLog(msg$)
msg$ = "REQUIRE_SUM_SI_TO_CA_LE_15: " & Format$(Format$(REQUIRE_SUM_SI_TO_CA_LE_15%, f84), a80$)
Call IOWriteLog(msg$)
msg$ = "REQUIRE_SUM_SI_TO_MG_GE_13: " & Format$(Format$(REQUIRE_SUM_SI_TO_MG_GE_13%, f84), a80$)
Call IOWriteLog(msg$)
msg$ = "REQUIRE_SUM_SI_TO_NA_GE_15: " & Format$(Format$(REQUIRE_SUM_SI_TO_NA_GE_15%, f84), a80$)
Call IOWriteLog(msg$)
msg$ = "REQUIRE_SUM_SI_TO_K_GE_15_5: " & Format$(Format$(REQUIRE_SUM_SI_TO_K_GE_15_5%, f84), a80$)
Call IOWriteLog(msg$)

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
ConvertFerrousFerricRatioFromComposition3Error:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertFerrousFerricRatioFromComposition3"
ierror = True
Exit Sub

ConvertFerrousFerricRatioFromComposition3NoIron:
msg$ = "No iron is present in the element list. Cannot calculate a ferrous/ferric ratio."
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertFerrousFerricRatioFromComposition3"
ierror = True
Exit Sub

ConvertFerrousFerricRatioFromComposition3NotFeO:
msg$ = "Cannot calculate a ferrous/ferric ratio. Iron stoichiometry must be specified as FeO (Cations=1, Oxygens=1) in the Elements/Cations dialog."
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertFerrousFerricRatioFromComposition3"
ierror = True
Exit Sub

End Sub

Sub ConvertAmphiboleCalculationLoop(nelements As Integer, AtomicNumbers() As Integer, AtomicWeights() As Single, ElementalWeightFractions() As Single, NumCats() As Integer, NumOxds() As Integer, OxideProportions() As Single, DisableQuantFlags() As Integer, loopNumber As Integer, Fe3overSumFe As Single, Mn3overSumMn As Single, finalWtPercentValues_H2O As Single, options_from_Locock_spreadsheet() As Integer)
' Based on Locock amphibole spreadsheet

ierror = False
On Error GoTo ConvertAmphiboleCalculationLoopError

' Declaration of the 8 different options. This should be handled in the interface.
Dim ORTHORHOMBIC As Boolean
Dim USE_INITIAL_M3_OVER_SUM_M As Boolean
Dim ESTIMATEOH2_2TI As Boolean
Dim REQUIRE_INITIAL_H2O As Boolean

ORTHORHOMBIC = options_from_Locock_spreadsheet%(1)
USE_INITIAL_M3_OVER_SUM_M = options_from_Locock_spreadsheet%(2)
ESTIMATEOH2_2TI = options_from_Locock_spreadsheet%(3)
REQUIRE_INITIAL_H2O = options_from_Locock_spreadsheet%(4)

Dim REQUIRE_SUM_SI_TO_CA_LE_15 As Integer
Dim REQUIRE_SUM_SI_TO_MG_GE_13 As Integer
Dim REQUIRE_SUM_SI_TO_NA_GE_15 As Integer
Dim REQUIRE_SUM_SI_TO_K_GE_15_5 As Integer
REQUIRE_SUM_SI_TO_CA_LE_15% = options_from_Locock_spreadsheet%(5)
REQUIRE_SUM_SI_TO_MG_GE_13% = options_from_Locock_spreadsheet%(6)
REQUIRE_SUM_SI_TO_NA_GE_15% = options_from_Locock_spreadsheet%(7)
REQUIRE_SUM_SI_TO_K_GE_15_5% = options_from_Locock_spreadsheet%(8)

' Variables used in the loops (n) and when retrieving the element channel used (ip)
Dim ip As Integer
Dim n As Integer

' List of elements used in Locock's spreadsheet. In the same order as in the speadsheet.
Dim elementList(1 To 24) As Integer
elementList%(1) = ATOMIC_NUM_SILICON%
elementList%(2) = ATOMIC_NUM_PHOSPHORUS%
elementList%(3) = ATOMIC_NUM_TITANIUM%
elementList%(4) = ATOMIC_NUM_ZIRCONIUM%
elementList%(5) = ATOMIC_NUM_ALUMINUM%
elementList%(6) = ATOMIC_NUM_SCANDIUM%
elementList%(7) = ATOMIC_NUM_VANADIUM%
elementList%(8) = ATOMIC_NUM_CHROMIUM%
elementList%(9) = ATOMIC_NUM_MANGANESE%         ' MnO
elementList%(10) = ATOMIC_NUM_MANGANESE%        ' Mn2O3
elementList%(11) = ATOMIC_NUM_IRON%             ' FeO
elementList%(12) = ATOMIC_NUM_IRON%             ' Fe2O3
elementList%(13) = ATOMIC_NUM_COBALT%
elementList%(14) = ATOMIC_NUM_NICKEL%
elementList%(15) = ATOMIC_NUM_ZINC%
elementList%(16) = ATOMIC_NUM_BERYLLIUM%
elementList%(17) = ATOMIC_NUM_MAGNESIUM%
elementList%(18) = ATOMIC_NUM_CALCIUM%
elementList%(19) = ATOMIC_NUM_STRONTIUM%
elementList%(20) = ATOMIC_NUM_LITHIUM%
elementList%(21) = ATOMIC_NUM_SODIUM%
elementList%(22) = ATOMIC_NUM_LEAD%
elementList%(23) = ATOMIC_NUM_POTASSIUM%
elementList%(24) = ATOMIC_NUM_HYDROGEN%

Dim oxidesElements(1 To 24) As Single           ' array containing the oxide weight of the elements
Dim molarProportionsCations(1 To 24) As Single  ' array containing the molar proportion of the cations
Dim molarProportionsAnions(1 To 24) As Single   ' array containing the molar proportion of the anions

' Channel indexes for F and Cl
Dim ipF As Integer
Dim ipCl As Integer

' Elemental weigth fractions and molar proportions for F and Cl
Dim ElementalWeightFractionsF As Single
Dim ElementalWeightFractionsCl As Single
Dim molarProportionsAnionsF As Single
Dim molarProportionsAnionsCl As Single

Dim OFClCalc As Single
Dim initialTotal As Single

' Convert the elemental compostion to oxide. Handles the case where the element is FeO or Fe2O3 as well as MnO or Mn2O3.
' The first loop (or first pass) is handled differently than the others.
For n% = 1 To 24
    ip% = IPOS2DQ%(nelements%, elementList%(n%), AtomicNumbers%(), DisableQuantFlags%())
    If ip% = 0 Then
        oxidesElements!(n%) = 0#
        molarProportionsCations!(n%) = 0#
        molarProportionsAnions!(n%) = 0#
    Else
        If n% = 9 Then
            If loopNumber = 1 Then
                oxidesElements!(n%) = ElementalWeightFractions!(ip%) + ElementalWeightFractions!(ip%) * OxideProportions!(ip%) ' convert elements to oxides XXX need to check for H2O
            Else
                oxidesElements!(n%) = (1 - Mn3overSumMn!) * (0 * 0.89865734954961 + ElementalWeightFractions!(ip%) + ElementalWeightFractions!(ip%) * OxideProportions!(ip%))
            End If
            molarProportionsCations!(n%) = oxidesElements!(n%) / (NumCats%(ip%) * AtomicWeights!(ip%) + NumOxds%(ip%) * AllAtomicWts!(ATOMIC_NUM_OXYGEN%)) * NumCats%(ip%) * 100#
            molarProportionsAnions!(n%) = molarProportionsCations!(n%) * NumOxds%(ip%) / NumCats%(ip%)
        
        ElseIf n% = 10 Then
            If loopNumber = 1 Then
                oxidesElements!(n%) = 0#  ' initialize Oxide_Mn2O3!
            Else
                oxidesElements!(n%) = Mn3overSumMn! * (0 * 0.89865734954961 + ElementalWeightFractions!(ip%) + ElementalWeightFractions!(ip%) * OxideProportions!(ip%)) / 0.89865734954961
            End If
            molarProportionsCations!(n%) = oxidesElements!(n%) / 157.87429 * 2# * 100#
            molarProportionsAnions!(n%) = 1.5 * molarProportionsCations!(n%)
            
        ElseIf n% = 11 Then
            If loopNumber = 1 Then
                oxidesElements!(n%) = ElementalWeightFractions!(ip%) + ElementalWeightFractions!(ip%) * OxideProportions!(ip%) ' convert elements to oxides XXX need to check for H2O
            Else
                oxidesElements!(n%) = (1 - Fe3overSumFe!) * (0 * 0.899808502 + ElementalWeightFractions!(ip%) + ElementalWeightFractions!(ip%) * OxideProportions!(ip%))
            End If
            molarProportionsCations!(n%) = oxidesElements!(n%) / (NumCats%(ip%) * AtomicWeights!(ip%) + NumOxds%(ip%) * AllAtomicWts!(ATOMIC_NUM_OXYGEN%)) * NumCats%(ip%) * 100#
            molarProportionsAnions!(n%) = molarProportionsCations!(n%) * NumOxds%(ip%) / NumCats%(ip%)
        
        ElseIf n% = 12 Then
            If loopNumber = 1 Then
                oxidesElements!(n%) = 0#  ' initialize Oxide_Fe2O3!
            Else
                 oxidesElements!(n%) = Fe3overSumFe! * (0 * 0.899808502 + ElementalWeightFractions!(ip%) + ElementalWeightFractions!(ip%) * OxideProportions!(ip%)) / 0.899808502
            End If
            molarProportionsCations!(n%) = oxidesElements!(n%) / 159.6882 * 2# * 100#
            molarProportionsAnions!(n%) = 1.5 * molarProportionsCations!(n%)
            
        Else
            oxidesElements!(n%) = ElementalWeightFractions!(ip%) + ElementalWeightFractions!(ip%) * OxideProportions!(ip%) ' convert elements to oxides XXX need to check for H2O
            molarProportionsCations!(n%) = oxidesElements!(n%) / (NumCats%(ip%) * AtomicWeights!(ip%) + NumOxds%(ip%) * AllAtomicWts!(ATOMIC_NUM_OXYGEN%)) * NumCats%(ip%) * 100#
            molarProportionsAnions!(n%) = molarProportionsCations!(n%) * NumOxds%(ip%) / NumCats%(ip%)
        End If
    End If
Next n%

ipF% = IPOS2DQ%(nelements%, 9, AtomicNumbers%(), DisableQuantFlags%())
If ipF% > 0 Then
    ElementalWeightFractionsF! = ElementalWeightFractions!(ipF%)
    molarProportionsAnionsF! = ElementalWeightFractionsF! / AtomicWeights!(ipF%)
Else
    ElementalWeightFractionsF! = 0#
    molarProportionsAnionsF! = 0#
End If

ipCl% = IPOS2DQ%(nelements%, 17, AtomicNumbers%(), DisableQuantFlags%())
If ipCl% > 0 Then
    ElementalWeightFractionsCl! = ElementalWeightFractions!(ipCl%)
    molarProportionsAnionsCl! = ElementalWeightFractionsCl! / AtomicWeights!(ipCl%)
Else
    ElementalWeightFractionsCl! = 0#
    molarProportionsAnionsCl! = 0#
End If

OFClCalc! = 0#
If ElementalWeightFractionsF! + ElementalWeightFractionsCl! > 0# Then
    OFClCalc! = Round(-ElementalWeightFractionsF! * 15.9994 / (2# * 18.9984032) - ElementalWeightFractionsCl! * 15.9994 / (2# * 35.453), 2#)
End If

initialTotal = 0#
For n% = 1 To 24
    initialTotal! = initialTotal! + oxidesElements!(n%)
Next n%
initialTotal = initialTotal + ElementalWeightFractionsF! + ElementalWeightFractionsCl! + OFClCalc!

Dim sumWithH As Single
Dim sumWithoutH As Single
Dim sumSitoCa As Single
Dim sumSitoMg As Single
Dim sumSitoNa As Single

sumSitoMg! = 0#
For n% = 1 To 17
    sumSitoMg! = sumSitoMg! + molarProportionsCations!(n%)
Next n%

sumSitoMg! = sumSitoMg! + molarProportionsCations!(20)
sumSitoCa! = sumSitoMg! + molarProportionsCations!(18) + molarProportionsCations!(19)
sumSitoNa! = sumSitoCa! + molarProportionsCations!(21)
sumWithoutH! = sumSitoNa! + molarProportionsCations!(22) + molarProportionsCations!(23)
sumWithH! = sumWithoutH! + molarProportionsCations!(24)

Dim sumOfAllAnions As Single
Dim sumOfAllAnionsMinusOHFCl As Single

sumOfAllAnions! = 0#
sumOfAllAnionsMinusOHFCl! = 0#
For n% = 1 To 24
    sumOfAllAnions! = sumOfAllAnions! + molarProportionsAnions!(n%)
Next n%

sumOfAllAnions! = sumOfAllAnions! + 0.5 * molarProportionsAnionsF! + 0.5 * molarProportionsAnionsCl!
sumOfAllAnionsMinusOHFCl! = sumOfAllAnions! - 0.5 * molarProportionsAnionsF! - 0.5 * molarProportionsAnionsCl! - molarProportionsAnions!(24)

Dim FplusClper24anions As Single
FplusClper24anions! = (molarProportionsAnionsF! + molarProportionsAnionsCl!) / sumOfAllAnions! * 24#
Dim Hper24anions As Single
Hper24anions! = molarProportionsCations!(24) / sumOfAllAnions! * 24#
Dim OequivalentsMin As Single
OequivalentsMin! = 24# - 0.5 * Hper24anions! - 0.5 * FplusClper24anions!
Dim Tiper24anions As Single

If (molarProportionsCations!(3) / sumOfAllAnions! * 24# > 8# - (molarProportionsCations!(5) + molarProportionsCations!(2) + molarProportionsCations!(1)) / sumOfAllAnions! * 24#) And (8# - (molarProportionsCations!(5) + molarProportionsCations!(2) + molarProportionsCations!(1)) / sumOfAllAnions! * 24#) > 0 Then
    Tiper24anions! = (molarProportionsCations!(3) / sumOfAllAnions! * 24#) - (8# - (molarProportionsCations!(5) + molarProportionsCations!(2) + molarProportionsCations!(1)) / sumOfAllAnions! * 24#)
Else
    Tiper24anions! = (molarProportionsCations!(3) / sumOfAllAnions! * 24#)
End If

Dim OequivalentsMax As Single
If 23# + Tiper24anions! > 24# Then
    OequivalentsMax! = 24# - 0.5 * FplusClper24anions!
Else
    OequivalentsMax! = 23# + Tiper24anions!
End If

' Calculates the initial proportion of the cations
Dim initialProportionsCations(1 To 23) As Single
Dim initialProportionsCationSumSitoK As Single
initialProportionsCationSumSitoK! = 0#
For n% = 1 To 23
    If REQUIRE_INITIAL_H2O = True Then
        initialProportionsCations!(n%) = molarProportionsCations!(n%) / sumOfAllAnionsMinusOHFCl! * OequivalentsMin!
    ElseIf ESTIMATEOH2_2TI = True Then
        initialProportionsCations!(n%) = molarProportionsCations!(n%) / sumOfAllAnionsMinusOHFCl! * OequivalentsMax!
    Else
        initialProportionsCations!(n%) = molarProportionsCations!(n%) / sumOfAllAnionsMinusOHFCl! * 23#
    End If
    initialProportionsCationSumSitoK! = initialProportionsCationSumSitoK! + initialProportionsCations!(n%)
Next n%

Dim initialproportionsOH As Single
Dim initialproportionsF As Single
Dim initialproportionsCl As Single
Dim initialproportionsSumSitoCa As Single
Dim initialproportionsSumSitoMg As Single
Dim initialproportionsSumSitoNa As Single

If REQUIRE_INITIAL_H2O = True Then
        initialproportionsOH! = molarProportionsCations!(24) / sumOfAllAnionsMinusOHFCl! * OequivalentsMin!
        initialproportionsF! = molarProportionsAnionsF! / sumOfAllAnionsMinusOHFCl! * OequivalentsMin!
        initialproportionsCl! = molarProportionsAnionsCl! / sumOfAllAnionsMinusOHFCl! * OequivalentsMin!
ElseIf ESTIMATEOH2_2TI = True Then
        initialproportionsOH! = molarProportionsCations!(24) / sumOfAllAnionsMinusOHFCl! * OequivalentsMax!
        initialproportionsF! = molarProportionsAnionsF! / sumOfAllAnionsMinusOHFCl! * OequivalentsMax!
        initialproportionsCl! = molarProportionsAnionsCl! / sumOfAllAnionsMinusOHFCl! * OequivalentsMax!
Else
        initialproportionsOH! = molarProportionsCations!(24) / sumOfAllAnionsMinusOHFCl! * 23#
        initialproportionsF! = molarProportionsAnionsF! / sumOfAllAnionsMinusOHFCl! * 23#
        initialproportionsCl! = molarProportionsAnionsCl! / sumOfAllAnionsMinusOHFCl! * 23#
End If

initialproportionsSumSitoMg! = 0
For n% = 1 To 17
    initialproportionsSumSitoMg! = initialproportionsSumSitoMg! + initialProportionsCations!(n%)
Next n%
initialproportionsSumSitoMg! = initialproportionsSumSitoMg! + initialProportionsCations!(20)
initialproportionsSumSitoCa! = initialproportionsSumSitoMg! + initialProportionsCations!(18) + initialProportionsCations!(19)
initialproportionsSumSitoNa! = initialproportionsSumSitoCa! + initialProportionsCations!(21)

Dim initialproportions_Fe3overSumFe As Single ' todo XXX
Dim initialproportions_Mn3overSumMn As Single ' todo XXX

If initialProportionsCations!(12) + initialProportionsCations!(11) > 0 Then
    initialproportions_Fe3overSumFe! = initialProportionsCations!(12) / (initialProportionsCations!(12) + initialProportionsCations!(11))
    initialproportions_Mn3overSumMn! = initialProportionsCations!(10) / (initialProportionsCations!(10) + initialProportionsCations!(9))
Else
    initialproportions_Fe3overSumFe! = 0#
    initialproportions_Mn3overSumMn! = 0#
End If

' Calculates the oxygen Anions Corresponding To Cations
Dim oxygenAnionsCorrespondingToCations(1 To 23) As Single
Dim sumOequivalents As Single

sumOequivalents! = 0#
For n% = 1 To 23
    ip% = IPOS2DQ%(nelements%, elementList%(n%), AtomicNumbers%(), DisableQuantFlags%())
    If ip% = 0 Then
        oxygenAnionsCorrespondingToCations!(n%) = 0#
    ElseIf n% = 10 Or n% = 12 Then
        oxygenAnionsCorrespondingToCations!(n%) = initialProportionsCations!(n%) * 1.5
        sumOequivalents! = sumOequivalents! + oxygenAnionsCorrespondingToCations!(n%)
    Else
        oxygenAnionsCorrespondingToCations!(n%) = initialProportionsCations!(n%) * NumOxds%(ip%) / NumCats%(ip%) 'Assume MnO and FeO in PfE
        sumOequivalents! = sumOequivalents! + oxygenAnionsCorrespondingToCations!(n%)
    End If
Next n%

' All ferrous Fe and Mn2+
Dim allFerrousIronAndMn2(1 To 23) As Single
Dim cst As Single

cst! = sumOequivalents! / (sumOequivalents! - (oxygenAnionsCorrespondingToCations!(12) - initialProportionsCations!(12)) - (oxygenAnionsCorrespondingToCations!(10) - initialProportionsCations!(10)))
For n% = 1 To 23
    If n% = 9 Then
        allFerrousIronAndMn2!(n%) = (initialProportionsCations!(n%) + initialProportionsCations!(n% + 1)) * cst!
    ElseIf n% = 10 Then
        allFerrousIronAndMn2!(n%) = 0#
    ElseIf n% = 11 Then
        allFerrousIronAndMn2!(n%) = (initialProportionsCations!(n%) + initialProportionsCations!(n% + 1)) * cst!
    ElseIf n% = 12 Then
        allFerrousIronAndMn2!(n%) = 0#
    Else
        allFerrousIronAndMn2!(n%) = initialProportionsCations!(n%) * cst!
    End If
Next n%

Dim allFerrousIronAndMn2_sumSitoK As Single
Dim allFerrousIronAndMn2_sumSitoCa As Single
Dim allFerrousIronAndMn2_sumSitoMg As Single
Dim allFerrousIronAndMn2_sumSitoNa As Single

allFerrousIronAndMn2_sumSitoMg! = 0#
For n% = 1 To 17
    allFerrousIronAndMn2_sumSitoMg! = allFerrousIronAndMn2_sumSitoMg! + allFerrousIronAndMn2!(n%)
Next
allFerrousIronAndMn2_sumSitoMg! = allFerrousIronAndMn2_sumSitoMg! + allFerrousIronAndMn2!(20)
allFerrousIronAndMn2_sumSitoCa! = allFerrousIronAndMn2_sumSitoMg! + allFerrousIronAndMn2!(18) + allFerrousIronAndMn2!(19)
allFerrousIronAndMn2_sumSitoNa! = allFerrousIronAndMn2_sumSitoCa! + allFerrousIronAndMn2!(21)
allFerrousIronAndMn2_sumSitoK! = allFerrousIronAndMn2_sumSitoNa! + allFerrousIronAndMn2!(22) + allFerrousIronAndMn2!(23)

Dim allFerrousIronAndMn2_oxygenAnionsCorrespondingToCations(1 To 23) As Single
Dim allFerrousIronAndMn2_sumOequivalents As Single
allFerrousIronAndMn2_sumOequivalents! = 0#
For n% = 1 To 23
    ip% = IPOS2DQ%(nelements%, elementList%(n%), AtomicNumbers%(), DisableQuantFlags%())
    If ip% = 0 Then
        allFerrousIronAndMn2_oxygenAnionsCorrespondingToCations!(n%) = 0#
    ElseIf n% = 10 Or n% = 12 Then
        allFerrousIronAndMn2_oxygenAnionsCorrespondingToCations!(n%) = allFerrousIronAndMn2!(n%) * 1.5
        allFerrousIronAndMn2_sumOequivalents! = allFerrousIronAndMn2_sumOequivalents! + allFerrousIronAndMn2_oxygenAnionsCorrespondingToCations!(n%)
    Else
        allFerrousIronAndMn2_oxygenAnionsCorrespondingToCations!(n%) = allFerrousIronAndMn2!(n%) * NumOxds%(ip%) / NumCats%(ip%)
        allFerrousIronAndMn2_sumOequivalents! = allFerrousIronAndMn2_sumOequivalents! + allFerrousIronAndMn2_oxygenAnionsCorrespondingToCations!(n%)
    End If
Next

' All ferric Fe and Mn2+
Dim allFerricIronAndMn2(1 To 23) As Single

cst! = sumOequivalents! / (sumOequivalents! + (oxygenAnionsCorrespondingToCations!(11) * 0.5) - (oxygenAnionsCorrespondingToCations!(10) - initialProportionsCations!(10)))
For n% = 1 To 23
    If n% = 9 Then
        allFerricIronAndMn2!(n%) = (initialProportionsCations!(n%) + initialProportionsCations!(n% + 1)) * cst!
    ElseIf n% = 10 Then
        allFerricIronAndMn2!(n%) = 0#
    ElseIf n% = 12 Then
        allFerricIronAndMn2!(n%) = (initialProportionsCations!(n% - 1) + initialProportionsCations!(n%)) * cst!
    ElseIf n% = 11 Then
        allFerricIronAndMn2!(n%) = 0#
    Else
        allFerricIronAndMn2!(n%) = initialProportionsCations!(n%) * cst!
    End If
Next n%

Dim allFerricIronAndMn2_sumSitoK As Single
Dim allFerricIronAndMn2_sumSitoCa As Single
Dim allFerricIronAndMn2_sumSitoMg As Single
Dim allFerricIronAndMn2_sumSitoNa As Single
allFerricIronAndMn2_sumSitoMg! = 0#
For n% = 1 To 17
    allFerricIronAndMn2_sumSitoMg! = allFerricIronAndMn2_sumSitoMg! + allFerricIronAndMn2!(n%)
Next
allFerricIronAndMn2_sumSitoMg! = allFerricIronAndMn2_sumSitoMg! + allFerricIronAndMn2!(20)
allFerricIronAndMn2_sumSitoCa! = allFerricIronAndMn2_sumSitoMg! + allFerricIronAndMn2!(18) + allFerricIronAndMn2!(19)
allFerricIronAndMn2_sumSitoNa! = allFerricIronAndMn2_sumSitoCa! + allFerricIronAndMn2!(21)
allFerricIronAndMn2_sumSitoK! = allFerricIronAndMn2_sumSitoNa! + allFerricIronAndMn2!(22) + allFerricIronAndMn2!(23)

Dim allFerricIronAndMn2_oxygenAnionsCorrespondingToCations(1 To 23) As Single
Dim allFerricIronAndMn2_sumOequivalents As Single
allFerricIronAndMn2_sumOequivalents! = 0#
For n% = 1 To 23
    ip% = IPOS2DQ%(nelements%, elementList%(n%), AtomicNumbers%(), DisableQuantFlags%())
    If ip% = 0 Then
        allFerricIronAndMn2_oxygenAnionsCorrespondingToCations!(n%) = 0#
    ElseIf n% = 10 Or n% = 12 Then
        allFerricIronAndMn2_oxygenAnionsCorrespondingToCations!(n%) = allFerricIronAndMn2!(n%) * 1.5
        allFerricIronAndMn2_sumOequivalents! = allFerricIronAndMn2_sumOequivalents! + allFerricIronAndMn2_oxygenAnionsCorrespondingToCations!(n%)
    Else
        allFerricIronAndMn2_oxygenAnionsCorrespondingToCations!(n%) = allFerricIronAndMn2!(n%) * NumOxds%(ip%) / NumCats%(ip%)
        allFerricIronAndMn2_sumOequivalents! = allFerricIronAndMn2_sumOequivalents! + allFerricIronAndMn2_oxygenAnionsCorrespondingToCations!(n%)
    End If
Next n%

' All ferric Fe and Mn3+
Dim allFerricIronAndMn3(1 To 23) As Single

cst! = sumOequivalents! / (sumOequivalents! + (oxygenAnionsCorrespondingToCations!(11) * 0.5) + (oxygenAnionsCorrespondingToCations!(9) * 0.5))
For n% = 1 To 23
    If n% = 10 Then
        allFerricIronAndMn3!(n%) = (initialProportionsCations!(n% - 1) + initialProportionsCations!(n%)) * cst!
    ElseIf n% = 9 Then
        allFerricIronAndMn3!(n%) = 0#
    ElseIf n% = 12 Then
        allFerricIronAndMn3!(n%) = (initialProportionsCations!(n% - 1) + initialProportionsCations!(n%)) * cst!
    ElseIf n% = 11 Then
        allFerricIronAndMn3!(n%) = 0#
    Else
        allFerricIronAndMn3!(n%) = initialProportionsCations!(n%) * cst!
    End If
Next n%

Dim allFerricIronAndMn3_sumSitoK As Single
Dim allFerricIronAndMn3_sumSitoCa As Single
Dim allFerricIronAndMn3_sumSitoMg As Single
Dim allFerricIronAndMn3_sumSitoNa As Single
allFerricIronAndMn3_sumSitoMg! = 0#
For n% = 1 To 17
    allFerricIronAndMn3_sumSitoMg! = allFerricIronAndMn3_sumSitoMg! + allFerricIronAndMn3!(n%)
Next
allFerricIronAndMn3_sumSitoMg! = allFerricIronAndMn3_sumSitoMg! + allFerricIronAndMn3!(20)
allFerricIronAndMn3_sumSitoCa! = allFerricIronAndMn3_sumSitoMg! + allFerricIronAndMn3!(18) + allFerricIronAndMn3!(19)
allFerricIronAndMn3_sumSitoNa! = allFerricIronAndMn3_sumSitoCa! + allFerricIronAndMn3!(21)
allFerricIronAndMn3_sumSitoK! = allFerricIronAndMn3_sumSitoNa! + allFerricIronAndMn3!(22) + allFerricIronAndMn3!(23)

Dim allFerricIronAndMn3_oxygenAnionsCorrespondingToCations(1 To 23) As Single
Dim allFerricIronAndMn3_sumOequivalents As Single
allFerricIronAndMn3_sumOequivalents! = 0#
For n% = 1 To 23
    ip% = IPOS2DQ%(nelements%, elementList%(n%), AtomicNumbers%(), DisableQuantFlags%())
    If ip% = 0 Then
        allFerricIronAndMn3_oxygenAnionsCorrespondingToCations!(n%) = 0#
    ElseIf n% = 10 Or n% = 12 Then
        allFerricIronAndMn3_oxygenAnionsCorrespondingToCations!(n%) = allFerricIronAndMn3!(n%) * 1.5
        allFerricIronAndMn3_sumOequivalents! = allFerricIronAndMn3_sumOequivalents! + allFerricIronAndMn3_oxygenAnionsCorrespondingToCations!(n%)
    Else
        allFerricIronAndMn3_oxygenAnionsCorrespondingToCations!(n%) = allFerricIronAndMn3!(n%) * NumOxds%(ip%) / NumCats%(ip%)
        allFerricIronAndMn3_sumOequivalents! = allFerricIronAndMn3_sumOequivalents! + allFerricIronAndMn3_oxygenAnionsCorrespondingToCations!(n%)
    End If
Next n%

Dim Fe3andMn2inChargeBalance As Integer
Dim Fe3andMn3inChargeBalance As Integer

If allFerricIronAndMn2!(1) <= 8 Then
    Dim average As Single
    average! = (allFerricIronAndMn2_sumSitoK! + allFerricIronAndMn2_sumSitoCa! + allFerricIronAndMn2_sumSitoMg! + allFerricIronAndMn2_sumSitoNa!) / 4#
    If average! / sumOequivalents! <= ((16# + 15# + 13# + 15#) / 4#) / 24# Then
        Fe3andMn2inChargeBalance% = 1
    Else
        Fe3andMn2inChargeBalance% = 0
    End If
Else
    Fe3andMn2inChargeBalance% = 0
End If
If Fe3andMn2inChargeBalance% = 0 Then
    Fe3andMn3inChargeBalance% = 1
Else
    Fe3andMn3inChargeBalance% = 0
End If

Dim chargeBalancePer15CationsSitoCa(1 To 23) As Single

For n% = 1 To 23
    If Fe3andMn2inChargeBalance% = 1 Then
        If allFerrousIronAndMn2_sumSitoCa! >= 15# Then
            If allFerricIronAndMn2_sumSitoCa! <= 15# Then
                chargeBalancePer15CationsSitoCa!(n%) = (15# - allFerricIronAndMn2_sumSitoCa!) / (allFerrousIronAndMn2_sumSitoCa! - allFerricIronAndMn2_sumSitoCa!) * allFerrousIronAndMn2!(n%) + (1# - (15# - allFerricIronAndMn2_sumSitoCa!) / (allFerrousIronAndMn2_sumSitoCa! - allFerricIronAndMn2_sumSitoCa!)) * allFerricIronAndMn2!(n%)
            Else
                chargeBalancePer15CationsSitoCa!(n%) = allFerricIronAndMn2!(n%)
            End If
        Else
            chargeBalancePer15CationsSitoCa!(n%) = allFerrousIronAndMn2!(n%)
        End If
    Else
        If allFerrousIronAndMn2_sumSitoCa! >= 15# Then
            If allFerricIronAndMn3_sumSitoCa! <= 15# Then
                chargeBalancePer15CationsSitoCa!(n%) = (15# - allFerricIronAndMn3_sumSitoCa!) / (allFerrousIronAndMn2_sumSitoCa! - allFerricIronAndMn3_sumSitoCa!) * allFerrousIronAndMn2!(n%) + (1# - (15# - allFerricIronAndMn3_sumSitoCa!) / (allFerrousIronAndMn2_sumSitoCa! - allFerricIronAndMn3_sumSitoCa!)) * allFerricIronAndMn3!(n%)
            Else
                chargeBalancePer15CationsSitoCa!(n%) = allFerricIronAndMn3!(n%)
            End If
        Else
            chargeBalancePer15CationsSitoCa!(n%) = allFerrousIronAndMn2!(n%)
        End If
    End If
Next n%

Dim chargeBalancePer15CationsSitoCa_sumSitoK As Single
Dim chargeBalancePer15CationsSitoCa_sumSitoCa As Single
Dim chargeBalancePer15CationsSitoCa_sumSitoMg As Single
Dim chargeBalancePer15CationsSitoCa_sumSitoNa As Single
chargeBalancePer15CationsSitoCa_sumSitoMg! = 0
For n% = 1 To 17
    chargeBalancePer15CationsSitoCa_sumSitoMg! = chargeBalancePer15CationsSitoCa_sumSitoMg! + chargeBalancePer15CationsSitoCa!(n%)
Next
chargeBalancePer15CationsSitoCa_sumSitoMg! = chargeBalancePer15CationsSitoCa_sumSitoMg! + chargeBalancePer15CationsSitoCa!(20)
chargeBalancePer15CationsSitoCa_sumSitoCa! = chargeBalancePer15CationsSitoCa_sumSitoMg! + chargeBalancePer15CationsSitoCa!(18) + chargeBalancePer15CationsSitoCa!(19)
chargeBalancePer15CationsSitoCa_sumSitoNa! = chargeBalancePer15CationsSitoCa_sumSitoCa! + chargeBalancePer15CationsSitoCa!(21)
chargeBalancePer15CationsSitoCa_sumSitoK! = chargeBalancePer15CationsSitoCa_sumSitoNa! + chargeBalancePer15CationsSitoCa!(22) + chargeBalancePer15CationsSitoCa!(23)

Dim chargeBalancePer15CationsSitoCa_Fe3overSumFe As Single
Dim chargeBalancePer15CationsSitoCa_Mn3overSumMn As Single
If chargeBalancePer15CationsSitoCa!(11) + chargeBalancePer15CationsSitoCa!(12) > 0 Then
    chargeBalancePer15CationsSitoCa_Fe3overSumFe! = chargeBalancePer15CationsSitoCa!(12) / (chargeBalancePer15CationsSitoCa!(11) + chargeBalancePer15CationsSitoCa!(12))
Else
    chargeBalancePer15CationsSitoCa_Fe3overSumFe! = 0
End If
If chargeBalancePer15CationsSitoCa!(9) + chargeBalancePer15CationsSitoCa!(10) > 0 Then
    chargeBalancePer15CationsSitoCa_Mn3overSumMn! = chargeBalancePer15CationsSitoCa!(10) / (chargeBalancePer15CationsSitoCa!(9) + chargeBalancePer15CationsSitoCa!(10))
Else
    chargeBalancePer15CationsSitoCa_Mn3overSumMn! = 0
End If

Dim chargeBalancePer15CationsSitoCa_oxygenAnionsCorrespondingToCations(1 To 23) As Single
Dim chargeBalancePer15CationsSitoCa_sumOequivalents As Single
chargeBalancePer15CationsSitoCa_sumOequivalents! = 0
For n% = 1 To 23
    ip% = IPOS2DQ%(nelements%, elementList%(n%), AtomicNumbers%(), DisableQuantFlags%())
    If ip% = 0 Then
        chargeBalancePer15CationsSitoCa_oxygenAnionsCorrespondingToCations!(n%) = 0
    ElseIf n% = 10 Or n% = 12 Then
        chargeBalancePer15CationsSitoCa_oxygenAnionsCorrespondingToCations!(n%) = chargeBalancePer15CationsSitoCa!(n%) * 1.5
        chargeBalancePer15CationsSitoCa_sumOequivalents! = chargeBalancePer15CationsSitoCa_sumOequivalents! + chargeBalancePer15CationsSitoCa_oxygenAnionsCorrespondingToCations!(n%)
    Else
        chargeBalancePer15CationsSitoCa_oxygenAnionsCorrespondingToCations!(n%) = chargeBalancePer15CationsSitoCa!(n%) * NumOxds%(ip%) / NumCats%(ip%)
        chargeBalancePer15CationsSitoCa_sumOequivalents! = chargeBalancePer15CationsSitoCa_sumOequivalents! + chargeBalancePer15CationsSitoCa_oxygenAnionsCorrespondingToCations!(n%)
    End If
Next n%

Dim chargeBalancePer13CationsSitoMg(1 To 23) As Single

For n% = 1 To 23
    If Fe3andMn2inChargeBalance% = 1 Then
        If allFerrousIronAndMn2_sumSitoMg! >= 13# Then
            If allFerricIronAndMn2_sumSitoMg! <= 13# Then
                chargeBalancePer13CationsSitoMg!(n%) = (13# - allFerricIronAndMn2_sumSitoMg!) / (allFerrousIronAndMn2_sumSitoMg! - allFerricIronAndMn2_sumSitoMg!) * allFerrousIronAndMn2!(n%) + (1# - (13# - allFerricIronAndMn2_sumSitoMg!) / (allFerrousIronAndMn2_sumSitoMg! - allFerricIronAndMn2_sumSitoMg!)) * allFerricIronAndMn2!(n%)
            Else
                chargeBalancePer13CationsSitoMg!(n%) = allFerricIronAndMn2!(n%)
            End If
        Else
            chargeBalancePer13CationsSitoMg!(n%) = allFerrousIronAndMn2!(n%)
        End If
    Else
        If allFerrousIronAndMn2_sumSitoMg! >= 13# Then
            If allFerricIronAndMn3_sumSitoMg! <= 13# Then
                chargeBalancePer13CationsSitoMg!(n%) = (13# - allFerricIronAndMn3_sumSitoMg!) / (allFerrousIronAndMn2_sumSitoMg! - allFerricIronAndMn3_sumSitoMg!) * allFerrousIronAndMn2!(n%) + (1# - (13# - allFerricIronAndMn3_sumSitoMg!) / (allFerrousIronAndMn2_sumSitoMg! - allFerricIronAndMn3_sumSitoMg!)) * allFerricIronAndMn3!(n%)
            Else
                chargeBalancePer13CationsSitoMg!(n%) = allFerricIronAndMn3!(n%)
            End If
        Else
            chargeBalancePer13CationsSitoMg!(n%) = allFerrousIronAndMn2!(n%)
        End If
    End If
Next n%

Dim chargeBalancePer13CationsSitoMg_sumSitoK As Single
Dim chargeBalancePer13CationsSitoMg_sumSitoCa As Single
Dim chargeBalancePer13CationsSitoMg_sumSitoMg As Single
Dim chargeBalancePer13CationsSitoMg_sumSitoNa As Single
chargeBalancePer13CationsSitoMg_sumSitoMg! = 0#
For n% = 1 To 17
    chargeBalancePer13CationsSitoMg_sumSitoMg! = chargeBalancePer13CationsSitoMg_sumSitoMg! + chargeBalancePer13CationsSitoMg!(n%)
Next
chargeBalancePer13CationsSitoMg_sumSitoMg! = chargeBalancePer13CationsSitoMg_sumSitoMg! + chargeBalancePer13CationsSitoMg!(20)
chargeBalancePer13CationsSitoMg_sumSitoCa! = chargeBalancePer13CationsSitoMg_sumSitoMg! + chargeBalancePer13CationsSitoMg!(18) + chargeBalancePer13CationsSitoMg!(19)
chargeBalancePer13CationsSitoMg_sumSitoNa! = chargeBalancePer13CationsSitoMg_sumSitoCa! + chargeBalancePer13CationsSitoMg!(21)
chargeBalancePer13CationsSitoMg_sumSitoK! = chargeBalancePer13CationsSitoMg_sumSitoNa! + chargeBalancePer13CationsSitoMg!(22) + chargeBalancePer13CationsSitoMg!(23)

Dim chargeBalancePer13CationsSitoMg_Fe3overSumFe As Single
Dim chargeBalancePer13CationsSitoMg_Mn3overSumMn As Single
If chargeBalancePer13CationsSitoMg!(11) + chargeBalancePer13CationsSitoMg!(12) > 0 Then
    chargeBalancePer13CationsSitoMg_Fe3overSumFe! = chargeBalancePer13CationsSitoMg!(12) / (chargeBalancePer13CationsSitoMg!(11) + chargeBalancePer13CationsSitoMg!(12))
Else
    chargeBalancePer13CationsSitoMg_Fe3overSumFe! = 0#
End If
If chargeBalancePer13CationsSitoMg!(9) + chargeBalancePer13CationsSitoMg!(10) > 0 Then
    chargeBalancePer13CationsSitoMg_Mn3overSumMn! = chargeBalancePer13CationsSitoMg!(10) / (chargeBalancePer13CationsSitoMg!(9) + chargeBalancePer13CationsSitoMg!(10))
Else
    chargeBalancePer13CationsSitoMg_Mn3overSumMn! = 0#
End If

Dim chargeBalancePer13CationsSitoMg_oxygenAnionsCorrespondingToCations(1 To 23) As Single
Dim chargeBalancePer13CationsSitoMg_sumOequivalents As Single

chargeBalancePer13CationsSitoMg_sumOequivalents! = 0#
For n% = 1 To 23
    ip% = IPOS2DQ%(nelements%, elementList%(n%), AtomicNumbers%(), DisableQuantFlags%())
    If ip% = 0 Then
        chargeBalancePer13CationsSitoMg_oxygenAnionsCorrespondingToCations!(n%) = 0#
    ElseIf n% = 10 Or n% = 12 Then
        chargeBalancePer13CationsSitoMg_oxygenAnionsCorrespondingToCations!(n%) = chargeBalancePer13CationsSitoMg!(n%) * 1.5
        chargeBalancePer13CationsSitoMg_sumOequivalents! = chargeBalancePer13CationsSitoMg_sumOequivalents! + chargeBalancePer13CationsSitoMg_oxygenAnionsCorrespondingToCations!(n%)
    Else
        chargeBalancePer13CationsSitoMg_oxygenAnionsCorrespondingToCations!(n%) = chargeBalancePer13CationsSitoMg!(n%) * NumOxds%(ip%) / NumCats%(ip%)
        chargeBalancePer13CationsSitoMg_sumOequivalents! = chargeBalancePer13CationsSitoMg_sumOequivalents! + chargeBalancePer13CationsSitoMg_oxygenAnionsCorrespondingToCations!(n%)
    End If
Next

Dim chargeBalancePer15CationsSitoNa(1 To 23) As Single
For n% = 1 To 23
    If Fe3andMn2inChargeBalance% = 1 Then
        If allFerrousIronAndMn2_sumSitoNa! >= 15# Then
            If allFerricIronAndMn2_sumSitoNa! <= 15# Then
                chargeBalancePer15CationsSitoNa!(n%) = (15# - allFerricIronAndMn2_sumSitoNa!) / (allFerrousIronAndMn2_sumSitoNa! - allFerricIronAndMn2_sumSitoNa!) * allFerrousIronAndMn2!(n%) + (1# - (15# - allFerricIronAndMn2_sumSitoNa!) / (allFerrousIronAndMn2_sumSitoNa! - allFerricIronAndMn2_sumSitoNa!)) * allFerricIronAndMn2!(n%)
            Else
                chargeBalancePer15CationsSitoNa!(n%) = allFerricIronAndMn2!(n%)
            End If
        Else
            chargeBalancePer15CationsSitoNa!(n%) = allFerrousIronAndMn2!(n%)
        End If
    Else
        If allFerrousIronAndMn2_sumSitoNa! >= 15# Then
            If allFerricIronAndMn3_sumSitoNa! <= 150 Then
                chargeBalancePer15CationsSitoNa!(n%) = (15# - allFerricIronAndMn3_sumSitoNa!) / (allFerrousIronAndMn2_sumSitoNa! - allFerricIronAndMn3_sumSitoNa!) * allFerrousIronAndMn2!(n%) + (1# - (15# - allFerricIronAndMn3_sumSitoNa!) / (allFerrousIronAndMn2_sumSitoNa! - allFerricIronAndMn3_sumSitoNa!)) * allFerricIronAndMn3!(n%)
            Else
                chargeBalancePer15CationsSitoNa!(n%) = allFerricIronAndMn3!(n%)
            End If
        Else
            chargeBalancePer15CationsSitoNa!(n%) = allFerrousIronAndMn2!(n%)
        End If
    End If
Next n%

Dim chargeBalancePer15CationsSitoNa_sumSitoK As Single
Dim chargeBalancePer15CationsSitoNa_sumSitoCa As Single
Dim chargeBalancePer15CationsSitoNa_sumSitoMg As Single
Dim chargeBalancePer15CationsSitoNa_sumSitoNa As Single

chargeBalancePer15CationsSitoNa_sumSitoMg! = 0#
For n% = 1 To 17
    chargeBalancePer15CationsSitoNa_sumSitoMg! = chargeBalancePer15CationsSitoNa_sumSitoMg! + chargeBalancePer15CationsSitoNa!(n%)
Next
chargeBalancePer15CationsSitoNa_sumSitoMg! = chargeBalancePer15CationsSitoNa_sumSitoMg! + chargeBalancePer15CationsSitoNa!(20)
chargeBalancePer15CationsSitoNa_sumSitoCa! = chargeBalancePer15CationsSitoNa_sumSitoMg! + chargeBalancePer15CationsSitoNa!(18) + chargeBalancePer15CationsSitoNa!(19)
chargeBalancePer15CationsSitoNa_sumSitoNa! = chargeBalancePer15CationsSitoNa_sumSitoCa! + chargeBalancePer15CationsSitoNa!(21)
chargeBalancePer15CationsSitoNa_sumSitoK! = chargeBalancePer15CationsSitoNa_sumSitoNa! + chargeBalancePer15CationsSitoNa!(22) + chargeBalancePer15CationsSitoNa!(23)

Dim chargeBalancePer15CationsSitoNa_Fe3overSumFe As Single
Dim chargeBalancePer15CationsSitoNa_Mn3overSumMn As Single
If chargeBalancePer15CationsSitoNa!(11) + chargeBalancePer15CationsSitoNa!(12) > 0 Then
    chargeBalancePer15CationsSitoNa_Fe3overSumFe! = chargeBalancePer15CationsSitoNa!(12) / (chargeBalancePer15CationsSitoNa!(11) + chargeBalancePer15CationsSitoNa!(12))
Else
    chargeBalancePer15CationsSitoNa_Fe3overSumFe! = 0#
End If
If chargeBalancePer15CationsSitoNa!(9) + chargeBalancePer15CationsSitoNa!(10) > 0 Then
    chargeBalancePer15CationsSitoNa_Mn3overSumMn! = chargeBalancePer15CationsSitoNa!(10) / (chargeBalancePer15CationsSitoNa!(9) + chargeBalancePer15CationsSitoNa!(10))
Else
    chargeBalancePer15CationsSitoNa_Mn3overSumMn! = 0#
End If

Dim chargeBalancePer15CationsSitoNa_oxygenAnionsCorrespondingToCations(1 To 23) As Single
Dim chargeBalancePer15CationsSitoNa_sumOequivalents As Single
chargeBalancePer15CationsSitoNa_sumOequivalents! = 0#
For n% = 1 To 23
    ip% = IPOS2DQ%(nelements%, elementList%(n%), AtomicNumbers%(), DisableQuantFlags%())
    If ip% = 0 Then
        chargeBalancePer15CationsSitoNa_oxygenAnionsCorrespondingToCations!(n%) = 0
    ElseIf n% = 10 Or n% = 12 Then
        chargeBalancePer15CationsSitoNa_oxygenAnionsCorrespondingToCations!(n%) = chargeBalancePer15CationsSitoNa!(n%) * 1.5
        chargeBalancePer15CationsSitoNa_sumOequivalents! = chargeBalancePer15CationsSitoNa_sumOequivalents! + chargeBalancePer15CationsSitoNa_oxygenAnionsCorrespondingToCations!(n%)
    Else
        chargeBalancePer15CationsSitoNa_oxygenAnionsCorrespondingToCations!(n%) = chargeBalancePer15CationsSitoNa!(n%) * NumOxds%(ip%) / NumCats%(ip%)
        chargeBalancePer15CationsSitoNa_sumOequivalents! = chargeBalancePer15CationsSitoNa_sumOequivalents! + chargeBalancePer15CationsSitoNa_oxygenAnionsCorrespondingToCations!(n%)
    End If
Next

Dim chargeBalancePer16CationsTotalNonH(1 To 23) As Single
For n% = 1 To 23
    If Fe3andMn2inChargeBalance% = 1 Then
        If allFerrousIronAndMn2_sumSitoK! >= 16# Then
            If allFerricIronAndMn2_sumSitoK! <= 16# Then
                chargeBalancePer16CationsTotalNonH!(n%) = (16# - allFerricIronAndMn2_sumSitoK!) / (allFerrousIronAndMn2_sumSitoK! - allFerricIronAndMn2_sumSitoK!) * allFerrousIronAndMn2!(n%) + (1# - (16# - allFerricIronAndMn2_sumSitoK!) / (allFerrousIronAndMn2_sumSitoK! - allFerricIronAndMn2_sumSitoK!)) * allFerricIronAndMn2!(n%)
            Else
                chargeBalancePer16CationsTotalNonH!(n%) = allFerricIronAndMn2!(n%)
            End If
        Else
            chargeBalancePer16CationsTotalNonH!(n%) = allFerrousIronAndMn2!(n%)
        End If
    Else
        If allFerrousIronAndMn2_sumSitoK! >= 16 Then
            If allFerricIronAndMn3_sumSitoK! <= 16 Then
                chargeBalancePer16CationsTotalNonH!(n%) = (16# - allFerricIronAndMn3_sumSitoK!) / (allFerrousIronAndMn2_sumSitoK! - allFerricIronAndMn3_sumSitoK!) * allFerrousIronAndMn2!(n%) + (1# - (16# - allFerricIronAndMn3_sumSitoK!) / (allFerrousIronAndMn2_sumSitoK! - allFerricIronAndMn3_sumSitoK!)) * allFerricIronAndMn3!(n%)
            Else
                chargeBalancePer16CationsTotalNonH!(n%) = allFerricIronAndMn3!(n%)
            End If
        Else
            chargeBalancePer16CationsTotalNonH!(n%) = allFerrousIronAndMn2!(n%)
        End If
    End If
Next n%

Dim chargeBalancePer16CationsTotalNonH_sumSitoK As Single
Dim chargeBalancePer16CationsTotalNonH_sumSitoCa As Single
Dim chargeBalancePer16CationsTotalNonH_sumSitoMg As Single
Dim chargeBalancePer16CationsTotalNonH_sumSitoNa As Single
chargeBalancePer16CationsTotalNonH_sumSitoMg! = 0#
For n% = 1 To 17
    chargeBalancePer16CationsTotalNonH_sumSitoMg! = chargeBalancePer16CationsTotalNonH_sumSitoMg! + chargeBalancePer16CationsTotalNonH!(n%)
Next
chargeBalancePer16CationsTotalNonH_sumSitoMg! = chargeBalancePer16CationsTotalNonH_sumSitoMg! + chargeBalancePer16CationsTotalNonH!(20)
chargeBalancePer16CationsTotalNonH_sumSitoCa! = chargeBalancePer16CationsTotalNonH_sumSitoMg! + chargeBalancePer16CationsTotalNonH!(18) + chargeBalancePer16CationsTotalNonH!(19)
chargeBalancePer16CationsTotalNonH_sumSitoNa! = chargeBalancePer16CationsTotalNonH_sumSitoCa! + chargeBalancePer16CationsTotalNonH!(21)
chargeBalancePer16CationsTotalNonH_sumSitoK! = chargeBalancePer16CationsTotalNonH_sumSitoNa! + chargeBalancePer16CationsTotalNonH!(22) + chargeBalancePer16CationsTotalNonH!(23)

Dim chargeBalancePer16CationsTotalNonH_Fe3overSumFe As Single
Dim chargeBalancePer16CationsTotalNonH_Mn3overSumMn As Single
If chargeBalancePer16CationsTotalNonH!(11) + chargeBalancePer16CationsTotalNonH!(12) > 0 Then
    chargeBalancePer16CationsTotalNonH_Fe3overSumFe! = chargeBalancePer16CationsTotalNonH!(12) / (chargeBalancePer16CationsTotalNonH!(11) + chargeBalancePer16CationsTotalNonH!(12))
Else
    chargeBalancePer16CationsTotalNonH_Fe3overSumFe! = 0#
End If
If chargeBalancePer16CationsTotalNonH!(9) + chargeBalancePer16CationsTotalNonH!(10) > 0 Then
    chargeBalancePer16CationsTotalNonH_Mn3overSumMn! = chargeBalancePer16CationsTotalNonH!(10) / (chargeBalancePer16CationsTotalNonH!(9) + chargeBalancePer16CationsTotalNonH!(10))
Else
    chargeBalancePer16CationsTotalNonH_Mn3overSumMn! = 0#
End If

Dim chargeBalancePer16CationsTotalNonH_oxygenAnionsCorrespondingToCations(1 To 23) As Single
Dim chargeBalancePer16CationsTotalNonH_sumOequivalents As Single
chargeBalancePer16CationsTotalNonH_sumOequivalents! = 0#
For n% = 1 To 23
    ip% = IPOS2DQ%(nelements%, elementList%(n%), AtomicNumbers%(), DisableQuantFlags%())
    If ip% = 0 Then
        chargeBalancePer16CationsTotalNonH_oxygenAnionsCorrespondingToCations!(n%) = 0#
    ElseIf n% = 10 Or n% = 12 Then
        chargeBalancePer16CationsTotalNonH_oxygenAnionsCorrespondingToCations!(n%) = chargeBalancePer16CationsTotalNonH!(n%) * 1.5
        chargeBalancePer16CationsTotalNonH_sumOequivalents! = chargeBalancePer16CationsTotalNonH_sumOequivalents! + chargeBalancePer16CationsTotalNonH_oxygenAnionsCorrespondingToCations!(n%)
    Else
        chargeBalancePer16CationsTotalNonH_oxygenAnionsCorrespondingToCations!(n%) = chargeBalancePer16CationsTotalNonH!(n%) * NumOxds%(ip%) / NumCats%(ip%)
        chargeBalancePer16CationsTotalNonH_sumOequivalents! = chargeBalancePer16CationsTotalNonH_sumOequivalents! + chargeBalancePer16CationsTotalNonH_oxygenAnionsCorrespondingToCations!(n%)
    End If
Next

' Criteria tests (threshold 0.0050 apfu)
' Sum Si to Ca (+Li) <=15
Dim sumSitoCa_LE_15_Si_LE_8apfu As Single
Dim sumSitoCa_LE_15_NonHCations_LE_16apfu As Single
Dim sumSitoCa_LE_15_SumSitoCa_LE_15apfu As Single
Dim sumSitoCa_LE_15_SumSitoMg_GE_13apfu As Single
Dim sumSitoCa_LE_15_SitoNa_GE_15apfu As Single
Dim sumSitoCa_LE_15_MaxDeviation As Single
Dim sumSitoCa_LE_15_Fe3overSumFe As Single
Dim sumSitoCa_LE_15_Mn3overSumMn As Single

sumSitoCa_LE_15_MaxDeviation! = 0#
If chargeBalancePer15CationsSitoCa!(1) <= 8# Then
    sumSitoCa_LE_15_Si_LE_8apfu! = 1#
Else
    sumSitoCa_LE_15_Si_LE_8apfu! = Abs(chargeBalancePer15CationsSitoCa!(1) - 8#)
    If sumSitoCa_LE_15_Si_LE_8apfu! > sumSitoCa_LE_15_MaxDeviation! Then
        sumSitoCa_LE_15_MaxDeviation! = sumSitoCa_LE_15_Si_LE_8apfu!
    End If
End If
If chargeBalancePer15CationsSitoCa_sumSitoK! <= 16# Then
    sumSitoCa_LE_15_NonHCations_LE_16apfu! = 1#
Else
    sumSitoCa_LE_15_NonHCations_LE_16apfu! = Abs(chargeBalancePer15CationsSitoCa_sumSitoK! - 16#)
    If sumSitoCa_LE_15_NonHCations_LE_16apfu! > sumSitoCa_LE_15_MaxDeviation! Then
        sumSitoCa_LE_15_MaxDeviation! = sumSitoCa_LE_15_NonHCations_LE_16apfu!
    End If
End If
If chargeBalancePer15CationsSitoCa_sumSitoCa! <= 15# Then
    sumSitoCa_LE_15_SumSitoCa_LE_15apfu! = 1#
Else
    sumSitoCa_LE_15_SumSitoCa_LE_15apfu! = Abs(chargeBalancePer15CationsSitoCa_sumSitoCa! - 15#)
    If sumSitoCa_LE_15_SumSitoCa_LE_15apfu! > sumSitoCa_LE_15_MaxDeviation! Then
        sumSitoCa_LE_15_MaxDeviation! = sumSitoCa_LE_15_SumSitoCa_LE_15apfu!
    End If
End If
If chargeBalancePer15CationsSitoCa_sumSitoMg! >= 13# Then
    sumSitoCa_LE_15_SumSitoMg_GE_13apfu! = 1#
Else
    sumSitoCa_LE_15_SumSitoMg_GE_13apfu! = Abs(13# - chargeBalancePer15CationsSitoCa_sumSitoMg!)
    If sumSitoCa_LE_15_SumSitoMg_GE_13apfu! > sumSitoCa_LE_15_MaxDeviation! Then
        sumSitoCa_LE_15_MaxDeviation! = sumSitoCa_LE_15_SumSitoMg_GE_13apfu!
    End If
End If
If chargeBalancePer15CationsSitoCa_sumSitoNa! >= 15# Then
    sumSitoCa_LE_15_SitoNa_GE_15apfu! = 1#
Else
    sumSitoCa_LE_15_SitoNa_GE_15apfu! = Abs(15# - chargeBalancePer15CationsSitoCa_sumSitoNa!)
    If sumSitoCa_LE_15_SitoNa_GE_15apfu! > sumSitoCa_LE_15_MaxDeviation! Then
        sumSitoCa_LE_15_MaxDeviation! = sumSitoCa_LE_15_SitoNa_GE_15apfu!
    End If
End If

If sumSitoCa_LE_15_MaxDeviation! < 0.005 Then
    sumSitoCa_LE_15_MaxDeviation! = 0#
End If

sumSitoCa_LE_15_Fe3overSumFe = chargeBalancePer15CationsSitoCa_Fe3overSumFe
sumSitoCa_LE_15_Mn3overSumMn = chargeBalancePer15CationsSitoCa_Mn3overSumMn

' Sum Si to Mg (+Li) >=13
Dim sumSitoMg_GE_13_Si_LE_8apfu As Single
Dim sumSitoMg_GE_13_NonHCations_LE_16apfu As Single
Dim sumSitoMg_GE_13_SumSitoCa_LE_15apfu As Single
Dim sumSitoMg_GE_13_SumSitoMg_GE_13apfu As Single
Dim sumSitoMg_GE_13_SitoNa_GE_15apfu As Single
Dim sumSitoMg_GE_13_MaxDeviation As Single
Dim sumSitoMg_GE_13_Fe3overSumFe As Single
Dim sumSitoMg_GE_13_Mn3overSumMn As Single

sumSitoMg_GE_13_MaxDeviation! = 0#
If chargeBalancePer13CationsSitoMg!(1) <= 8# Then
    sumSitoMg_GE_13_Si_LE_8apfu! = 1#
Else
    sumSitoMg_GE_13_Si_LE_8apfu! = Abs(chargeBalancePer13CationsSitoMg!(1) - 8#)
    If sumSitoMg_GE_13_Si_LE_8apfu! > sumSitoMg_GE_13_MaxDeviation! Then
        sumSitoMg_GE_13_MaxDeviation! = sumSitoMg_GE_13_Si_LE_8apfu!
    End If
End If
If chargeBalancePer13CationsSitoMg_sumSitoK! <= 16# Then
    sumSitoMg_GE_13_NonHCations_LE_16apfu! = 1#
Else
    sumSitoMg_GE_13_NonHCations_LE_16apfu! = Abs(chargeBalancePer13CationsSitoMg_sumSitoK! - 16#)
    If sumSitoMg_GE_13_NonHCations_LE_16apfu! > sumSitoMg_GE_13_MaxDeviation! Then
        sumSitoMg_GE_13_MaxDeviation! = sumSitoMg_GE_13_NonHCations_LE_16apfu!
    End If
End If
If chargeBalancePer13CationsSitoMg_sumSitoCa! <= 15# Then
    sumSitoMg_GE_13_SumSitoCa_LE_15apfu! = 1#
Else
    sumSitoMg_GE_13_SumSitoCa_LE_15apfu! = Abs(chargeBalancePer13CationsSitoMg_sumSitoCa! - 15#)
    If sumSitoMg_GE_13_SumSitoCa_LE_15apfu! > sumSitoMg_GE_13_MaxDeviation! Then
        sumSitoMg_GE_13_MaxDeviation! = sumSitoMg_GE_13_SumSitoCa_LE_15apfu!
    End If
End If
If Round(chargeBalancePer13CationsSitoMg_sumSitoMg!, 5) >= 13# Then ' XXX has to round it otherwise the equality is not respected
    sumSitoMg_GE_13_SumSitoMg_GE_13apfu! = 1#
Else
    sumSitoMg_GE_13_SumSitoMg_GE_13apfu! = Abs(13 - chargeBalancePer13CationsSitoMg_sumSitoMg!)
    If sumSitoMg_GE_13_SumSitoMg_GE_13apfu! > sumSitoMg_GE_13_MaxDeviation! Then
        sumSitoMg_GE_13_MaxDeviation! = sumSitoMg_GE_13_SumSitoMg_GE_13apfu!
    End If
End If
If chargeBalancePer13CationsSitoMg_sumSitoNa! >= 15# Then
    sumSitoMg_GE_13_SitoNa_GE_15apfu! = 1#
Else
    sumSitoMg_GE_13_SitoNa_GE_15apfu! = Abs(15# - chargeBalancePer13CationsSitoMg_sumSitoNa!)
    If sumSitoMg_GE_13_SitoNa_GE_15apfu! > sumSitoMg_GE_13_MaxDeviation! Then
        sumSitoMg_GE_13_MaxDeviation! = sumSitoMg_GE_13_SitoNa_GE_15apfu!
    End If
End If

If sumSitoMg_GE_13_MaxDeviation! < 0.005 Then
    sumSitoMg_GE_13_MaxDeviation! = 0#
End If

sumSitoMg_GE_13_Fe3overSumFe = chargeBalancePer13CationsSitoMg_Fe3overSumFe
sumSitoMg_GE_13_Mn3overSumMn = chargeBalancePer13CationsSitoMg_Mn3overSumMn

' Sum Si to Na >=15
Dim sumSitoNa_GE_15_Si_LE_8apfu As Single
Dim sumSitoNa_GE_15_NonHCations_LE_16apfu As Single
Dim sumSitoNa_GE_15_SumSitoCa_LE_15apfu As Single
Dim sumSitoNa_GE_15_SumSitoMg_GE_13apfu As Single
Dim sumSitoNa_GE_15_SitoNa_GE_15apfu As Single
Dim sumSitoNa_GE_15_MaxDeviation As Single
Dim sumSitoNa_GE_15_Fe3overSumFe As Single
Dim sumSitoNa_GE_15_Mn3overSumMn As Single

sumSitoNa_GE_15_MaxDeviation! = 0
If chargeBalancePer15CationsSitoNa!(1) <= 8 Then
    sumSitoNa_GE_15_Si_LE_8apfu! = 1
Else
    sumSitoNa_GE_15_Si_LE_8apfu! = Abs(chargeBalancePer15CationsSitoNa!(1) - 8)
    If sumSitoNa_GE_15_Si_LE_8apfu! > sumSitoNa_GE_15_MaxDeviation! Then            ' this if is unnecessary
        sumSitoNa_GE_15_MaxDeviation! = sumSitoNa_GE_15_Si_LE_8apfu!
    End If
End If
If chargeBalancePer15CationsSitoNa_sumSitoK! <= 16 Then
    sumSitoNa_GE_15_NonHCations_LE_16apfu! = 1
Else
    sumSitoNa_GE_15_NonHCations_LE_16apfu! = Abs(chargeBalancePer15CationsSitoNa_sumSitoK! - 16)
    If sumSitoNa_GE_15_NonHCations_LE_16apfu! > sumSitoNa_GE_15_MaxDeviation! Then
        sumSitoNa_GE_15_MaxDeviation! = sumSitoNa_GE_15_NonHCations_LE_16apfu!
    End If
End If
If chargeBalancePer15CationsSitoNa_sumSitoCa! <= 15 Then
    sumSitoNa_GE_15_SumSitoCa_LE_15apfu! = 1
Else
    sumSitoNa_GE_15_SumSitoCa_LE_15apfu! = Abs(chargeBalancePer15CationsSitoNa_sumSitoCa! - 15)
    If sumSitoNa_GE_15_SumSitoCa_LE_15apfu! > sumSitoNa_GE_15_MaxDeviation! Then
        sumSitoNa_GE_15_MaxDeviation! = sumSitoNa_GE_15_SumSitoCa_LE_15apfu!
    End If
End If
If chargeBalancePer15CationsSitoNa_sumSitoMg! >= 13 Then
    sumSitoNa_GE_15_SumSitoMg_GE_13apfu! = 1
Else
    sumSitoNa_GE_15_SumSitoMg_GE_13apfu! = Abs(13 - chargeBalancePer15CationsSitoNa_sumSitoMg!)
    If sumSitoNa_GE_15_SumSitoMg_GE_13apfu! > sumSitoNa_GE_15_MaxDeviation! Then
        sumSitoNa_GE_15_MaxDeviation! = sumSitoNa_GE_15_SumSitoMg_GE_13apfu!
    End If
End If
If chargeBalancePer15CationsSitoNa_sumSitoNa! >= 15 Then
    sumSitoNa_GE_15_SitoNa_GE_15apfu! = 1
Else
    sumSitoNa_GE_15_SitoNa_GE_15apfu! = Abs(15 - chargeBalancePer15CationsSitoNa_sumSitoNa!)
    If sumSitoNa_GE_15_SitoNa_GE_15apfu! > sumSitoNa_GE_15_MaxDeviation! Then
        sumSitoNa_GE_15_MaxDeviation! = sumSitoNa_GE_15_SitoNa_GE_15apfu!
    End If
End If

If sumSitoNa_GE_15_MaxDeviation! < 0.005 Then
    sumSitoNa_GE_15_MaxDeviation! = 0
End If

sumSitoNa_GE_15_Fe3overSumFe = chargeBalancePer15CationsSitoNa_Fe3overSumFe
sumSitoNa_GE_15_Mn3overSumMn = chargeBalancePer15CationsSitoNa_Mn3overSumMn

' Sum Si to K <=16
Dim sumSitoK_LE_16_Si_LE_8apfu As Single
Dim sumSitoK_LE_16_NonHCations_LE_16apfu As Single
Dim sumSitoK_LE_16_SumSitoCa_LE_15apfu As Single
Dim sumSitoK_LE_16_SumSitoMg_GE_13apfu As Single
Dim sumSitoK_LE_16_SitoNa_GE_15apfu As Single
Dim sumSitoK_LE_16_SitoK_GE_15_5apfu As Single
Dim sumSitoK_LE_16_MaxDeviation As Single
Dim sumSitoK_LE_16_Fe3overSumFe As Single
Dim sumSitoK_LE_16_Mn3overSumMn As Single

sumSitoK_LE_16_MaxDeviation! = 0
If chargeBalancePer16CationsTotalNonH!(1) <= 8 Then
    sumSitoK_LE_16_Si_LE_8apfu! = 1
Else
    sumSitoK_LE_16_Si_LE_8apfu! = Abs(chargeBalancePer16CationsTotalNonH!(1) - 8)
    If sumSitoK_LE_16_Si_LE_8apfu! > sumSitoK_LE_16_MaxDeviation! Then              ' this if is unnecessary
        sumSitoK_LE_16_MaxDeviation! = sumSitoK_LE_16_Si_LE_8apfu!
    End If
End If
If chargeBalancePer16CationsTotalNonH_sumSitoK! <= 16 Then
    sumSitoK_LE_16_NonHCations_LE_16apfu! = 1
Else
    sumSitoK_LE_16_NonHCations_LE_16apfu! = Abs(chargeBalancePer16CationsTotalNonH_sumSitoK! - 16)
    If sumSitoK_LE_16_NonHCations_LE_16apfu! > sumSitoK_LE_16_MaxDeviation! Then
        sumSitoK_LE_16_MaxDeviation! = sumSitoK_LE_16_NonHCations_LE_16apfu!
    End If
End If
If chargeBalancePer16CationsTotalNonH_sumSitoCa! <= 15 Then
    sumSitoK_LE_16_SumSitoCa_LE_15apfu! = 1
Else
    sumSitoK_LE_16_SumSitoCa_LE_15apfu! = Abs(chargeBalancePer16CationsTotalNonH_sumSitoCa! - 15)
    If sumSitoK_LE_16_SumSitoCa_LE_15apfu! > sumSitoK_LE_16_MaxDeviation! Then
        sumSitoK_LE_16_MaxDeviation! = sumSitoK_LE_16_SumSitoCa_LE_15apfu!
    End If
End If
If chargeBalancePer16CationsTotalNonH_sumSitoMg! >= 13 Then
    sumSitoK_LE_16_SumSitoMg_GE_13apfu! = 1
Else
    sumSitoK_LE_16_SumSitoMg_GE_13apfu! = Abs(13 - chargeBalancePer16CationsTotalNonH_sumSitoMg!)
    If sumSitoK_LE_16_SumSitoMg_GE_13apfu! > sumSitoK_LE_16_MaxDeviation! Then
        sumSitoK_LE_16_MaxDeviation! = sumSitoK_LE_16_SumSitoMg_GE_13apfu!
    End If
End If
If chargeBalancePer16CationsTotalNonH_sumSitoNa! >= 15 Then
    sumSitoK_LE_16_SitoNa_GE_15apfu! = 1
Else
    sumSitoK_LE_16_SitoNa_GE_15apfu! = Abs(15 - chargeBalancePer16CationsTotalNonH_sumSitoNa!)
    If sumSitoK_LE_16_SitoNa_GE_15apfu! > sumSitoK_LE_16_MaxDeviation! Then
        sumSitoK_LE_16_MaxDeviation! = sumSitoK_LE_16_SitoNa_GE_15apfu!
    End If
End If
If chargeBalancePer16CationsTotalNonH_sumSitoK! >= 15.5 Then
    sumSitoK_LE_16_SitoK_GE_15_5apfu! = 1
Else
    sumSitoK_LE_16_SitoK_GE_15_5apfu! = Abs(15.5 - chargeBalancePer16CationsTotalNonH_sumSitoK!)
    If sumSitoK_LE_16_SitoK_GE_15_5apfu! > sumSitoK_LE_16_MaxDeviation! Then
        sumSitoK_LE_16_MaxDeviation! = sumSitoK_LE_16_SitoK_GE_15_5apfu!
    End If
End If

If sumSitoK_LE_16_MaxDeviation! < 0.005 Then
    sumSitoK_LE_16_MaxDeviation! = 0
End If

sumSitoK_LE_16_Fe3overSumFe = chargeBalancePer16CationsTotalNonH_Fe3overSumFe
sumSitoK_LE_16_Mn3overSumMn = chargeBalancePer16CationsTotalNonH_Mn3overSumMn

' Deviations
Dim deviations(1 To 4) As Single
Dim mIn As Single
Dim min2 As Single
Dim min3 As Single

deviations!(1) = sumSitoCa_LE_15_MaxDeviation!
deviations!(2) = sumSitoMg_GE_13_MaxDeviation!
deviations!(3) = sumSitoNa_GE_15_MaxDeviation!
deviations!(4) = sumSitoK_LE_16_MaxDeviation!

Call ConvertBubbleSort(deviations)
If ierror Then Exit Sub
mIn! = deviations!(1)
min2! = deviations!(2)
min3! = deviations!(3)

Dim AcceptedDeviationFromIdeal_threshold_0_0050 As Single
If min2! - mIn! <= 0.005 Then
    If min3! - min2! <= 0.005 Then
        AcceptedDeviationFromIdeal_threshold_0_0050! = min3!
    Else
        AcceptedDeviationFromIdeal_threshold_0_0050! = min2!
    End If
Else
    AcceptedDeviationFromIdeal_threshold_0_0050! = mIn!
End If

Dim preferedFormula_sumSitoCa_EQ_15 As Single
Dim preferedFormula_sumSitoMg_EQ_13 As Single
Dim preferedFormula_sumSitoNa_EQ_15 As Single
Dim preferedFormula_sumSitoK_EQ_15_5 As Single

If REQUIRE_SUM_SI_TO_CA_LE_15% = 1 Then
    preferedFormula_sumSitoCa_EQ_15! = 1
Else
    If REQUIRE_SUM_SI_TO_MG_GE_13% + REQUIRE_SUM_SI_TO_NA_GE_15% + REQUIRE_SUM_SI_TO_K_GE_15_5% > 0 Then
        preferedFormula_sumSitoCa_EQ_15! = 0
    Else
        If sumSitoCa_LE_15_MaxDeviation! <= AcceptedDeviationFromIdeal_threshold_0_0050! Then
            preferedFormula_sumSitoCa_EQ_15! = 1
        Else
            preferedFormula_sumSitoCa_EQ_15! = 0
        End If
    End If
End If

If REQUIRE_SUM_SI_TO_MG_GE_13% = 1 Then
    preferedFormula_sumSitoMg_EQ_13! = 1
Else
    If REQUIRE_SUM_SI_TO_CA_LE_15% + REQUIRE_SUM_SI_TO_NA_GE_15% + REQUIRE_SUM_SI_TO_K_GE_15_5% > 0 Then
        preferedFormula_sumSitoMg_EQ_13! = 0
    Else
        If sumSitoMg_GE_13_MaxDeviation! <= AcceptedDeviationFromIdeal_threshold_0_0050! Then
            preferedFormula_sumSitoMg_EQ_13! = 1
        Else
            preferedFormula_sumSitoMg_EQ_13! = 0
        End If
    End If
End If

If REQUIRE_SUM_SI_TO_NA_GE_15% = 1 Then
    preferedFormula_sumSitoNa_EQ_15! = 1
Else
    If REQUIRE_SUM_SI_TO_CA_LE_15% + REQUIRE_SUM_SI_TO_MG_GE_13% + REQUIRE_SUM_SI_TO_K_GE_15_5% > 0 Then
        preferedFormula_sumSitoNa_EQ_15! = 0
    Else
        If sumSitoNa_GE_15_MaxDeviation! <= AcceptedDeviationFromIdeal_threshold_0_0050! Then
            preferedFormula_sumSitoNa_EQ_15! = 1
        Else
            preferedFormula_sumSitoNa_EQ_15! = 0
        End If
    End If
End If

If REQUIRE_SUM_SI_TO_K_GE_15_5% = 1 Then
    preferedFormula_sumSitoK_EQ_15_5! = 1
Else
    If REQUIRE_SUM_SI_TO_CA_LE_15% + REQUIRE_SUM_SI_TO_MG_GE_13% + REQUIRE_SUM_SI_TO_NA_GE_15% > 0 Then
        preferedFormula_sumSitoK_EQ_15_5! = 0
    Else
        If sumSitoK_LE_16_MaxDeviation! <= AcceptedDeviationFromIdeal_threshold_0_0050! Then
            preferedFormula_sumSitoK_EQ_15_5! = 1
        Else
            preferedFormula_sumSitoK_EQ_15_5! = 0
        End If
    End If
End If

Dim preferredFormulaSubtotal As Single
preferredFormulaSubtotal! = preferedFormula_sumSitoCa_EQ_15! + preferedFormula_sumSitoMg_EQ_13! + preferedFormula_sumSitoNa_EQ_15! + preferedFormula_sumSitoK_EQ_15_5!

' Formula
Dim formula(1 To 23) As Single
For n% = 1 To 23
    formula(n%) = 0
    If USE_INITIAL_M3_OVER_SUM_M = True Then
        formula!(n%) = initialProportionsCations!(n%)
    Else
        If preferedFormula_sumSitoCa_EQ_15 > 0 Then
            formula!(n%) = formula!(n%) + chargeBalancePer15CationsSitoCa!(n%)
        End If
        If preferedFormula_sumSitoMg_EQ_13 > 0 Then
            formula!(n%) = formula!(n%) + chargeBalancePer13CationsSitoMg!(n%)
        End If
        If preferedFormula_sumSitoNa_EQ_15 > 0 Then
            formula!(n%) = formula!(n%) + chargeBalancePer15CationsSitoNa!(n%)
        End If
        If preferedFormula_sumSitoK_EQ_15_5 > 0 Then
            formula!(n%) = formula!(n%) + chargeBalancePer16CationsTotalNonH!(n%)
        End If
        formula!(n%) = formula!(n%) / preferredFormulaSubtotal!
    End If
Next

Dim formula_sumSitoK As Single
Dim formula_sumSitoCa As Single
Dim formula_sumSitoMg As Single
Dim formula_sumSitoNa As Single

formula_sumSitoMg! = 0
For n% = 1 To 17
    formula_sumSitoMg! = formula_sumSitoMg! + formula!(n%)
Next
formula_sumSitoMg! = formula_sumSitoMg! + formula!(20)
formula_sumSitoCa! = formula_sumSitoMg! + formula!(18) + formula!(19)
formula_sumSitoNa! = formula_sumSitoCa! + formula!(21)
formula_sumSitoK! = formula_sumSitoNa! + formula!(22) + formula!(23)

Dim formula_OH_preliminary_step1 As Single
Dim formula_F_preliminary As Single
Dim formula_Cl_preliminary As Single
Dim formula_OH_preliminary_step2 As Single
Dim formula_H_mesured_and_rounded As Single
Dim formula_Fe3overSumFe As Single
Dim formula_Mn3overSumMn As Single

If initialproportionsOH! > 0 Then
    formula_OH_preliminary_step1! = initialproportionsOH!
Else
    formula_OH_preliminary_step1! = molarProportionsCations!(24) / molarProportionsCations!(1) * formula!(1)
End If
If initialproportionsF! > 0 Then
    formula_F_preliminary! = initialproportionsF!
Else
    formula_F_preliminary! = molarProportionsAnionsF / molarProportionsCations!(1) * formula!(1)
End If
If initialproportionsCl! > 0 Then
    formula_Cl_preliminary! = initialproportionsCl!
Else
    formula_Cl_preliminary! = molarProportionsAnionsCl / molarProportionsCations!(1) * formula!(1)
End If

If REQUIRE_INITIAL_H2O = True Then
    formula_OH_preliminary_step2! = formula_OH_preliminary_step1!
Else
    If formula_OH_preliminary_step1! + formula_F_preliminary! + formula_Cl_preliminary! > 2 Then
        If formula_F_preliminary! + formula_Cl_preliminary! < 2 Then
            formula_OH_preliminary_step2! = 2 - (formula_F_preliminary! + formula_Cl_preliminary!)
        Else
            formula_OH_preliminary_step2! = 0
        End If
    Else
        If formula_OH_preliminary_step1! + formula_F_preliminary! + formula_Cl_preliminary! < 2 Then
            If 2 - formula_F_preliminary! - formula_Cl_preliminary! > 0 Then
                formula_OH_preliminary_step2! = 2 - formula_F_preliminary! - formula_Cl_preliminary!
            Else
                formula_OH_preliminary_step2! = 0
            End If
        Else
            formula_OH_preliminary_step2! = formula_OH_preliminary_step1!
        End If
    End If
End If


If REQUIRE_INITIAL_H2O = True Then
    formula_H_mesured_and_rounded! = formula_OH_preliminary_step2!
Else
    If ESTIMATEOH2_2TI = True Then
        If formula_OH_preliminary_step1! + formula_F_preliminary! + formula_Cl_preliminary! > 2 Then
            If 2 * formula_OH_preliminary_step1! / (formula_OH_preliminary_step1! + formula_F_preliminary! + formula_Cl_preliminary!) - 2 * formula!(3) > 0 Then
                formula_H_mesured_and_rounded = 2 * formula_OH_preliminary_step1! / (formula_OH_preliminary_step1! + formula_F_preliminary! + formula_Cl_preliminary!) - 2 * formula!(3)
            Else
                formula_H_mesured_and_rounded = 0
            End If
        Else
            If formula_OH_preliminary_step1! + formula_F_preliminary! + formula_Cl_preliminary! < 2 Then
                Dim tmp As Single
                If formula!(3) > 8 - (formula!(1) + formula!(2) + formula!(5)) And 8 - (formula!(1) + formula!(2) + formula!(5)) > 0 Then
                    tmp! = formula!(3) - (8 - (formula!(1) + formula!(2) + formula!(5)))
                Else
                    tmp! = formula!(3)
                End If
                If 2 - formula_F_preliminary! - formula_Cl_preliminary! - 2 * tmp! > 0 Then
                    formula_H_mesured_and_rounded! = 2 - formula_F_preliminary! - formula_Cl_preliminary! - 2 * tmp!
                Else
                    formula_H_mesured_and_rounded! = 0
                End If
            Else
                formula_H_mesured_and_rounded! = formula_OH_preliminary_step1!
            End If
        End If
    Else
        If formula_OH_preliminary_step1! + formula_F_preliminary! + formula_Cl_preliminary! > 2 Then
            If formula_F_preliminary! + formula_Cl_preliminary! < 2 Then
                formula_H_mesured_and_rounded! = 2 - (formula_F_preliminary! + formula_Cl_preliminary!)
            Else
                formula_H_mesured_and_rounded! = 0
            End If
        Else
            If formula_OH_preliminary_step1! + formula_F_preliminary! + formula_Cl_preliminary! < 2 Then
                If 2 - formula_F_preliminary! - formula_Cl_preliminary! > 0 Then
                    formula_H_mesured_and_rounded! = 2 - formula_F_preliminary! - formula_Cl_preliminary!
                Else
                    formula_H_mesured_and_rounded! = 0
                End If
            Else
                formula_H_mesured_and_rounded! = formula_OH_preliminary_step1!
            End If
        End If
    End If
End If


If formula!(11) + formula!(12) > 0 Then
    formula_Fe3overSumFe! = formula!(12) / (formula!(11) + formula!(12))
Else
    formula_Fe3overSumFe! = 0
End If
If formula!(9) + formula!(10) > 0 Then
    formula_Mn3overSumMn! = formula!(10) / (formula!(9) + formula!(10))
Else
    formula_Mn3overSumMn! = 0
End If

Dim formula_AnionsCorrespondingToCations(1 To 23) As Single
Dim formula_AnionsCorrespondingToCations_sum As Single
Dim formula_AnionsCorrespondingToCations_H As Single
Dim formula_AnionsCorrespondingToCations_F As Single
Dim formula_AnionsCorrespondingToCations_Cl As Single
Dim formula_AnionsCorrespondingToCations_sumTotal As Single

formula_AnionsCorrespondingToCations_sum! = 0
For n% = 1 To 23
    ip% = IPOS2DQ%(nelements%, elementList%(n%), AtomicNumbers%(), DisableQuantFlags%())
    If ip% = 0 Then
        formula_AnionsCorrespondingToCations!(n%) = 0
    ElseIf n% = 10 Or n% = 12 Then
        formula_AnionsCorrespondingToCations!(n%) = formula!(n%) * 1.5
        formula_AnionsCorrespondingToCations_sum! = formula_AnionsCorrespondingToCations_sum! + formula_AnionsCorrespondingToCations!(n%)
    Else
        formula_AnionsCorrespondingToCations!(n%) = formula!(n%) * NumOxds%(ip%) / NumCats%(ip%)
        formula_AnionsCorrespondingToCations_sum! = formula_AnionsCorrespondingToCations_sum! + formula_AnionsCorrespondingToCations!(n%)
    End If
Next

formula_AnionsCorrespondingToCations_H! = 0.5 * formula_H_mesured_and_rounded!

If formula_F_preliminary! + formula_Cl_preliminary! > 2 Then
    formula_AnionsCorrespondingToCations_F! = 2 * formula_F_preliminary! / (formula_F_preliminary! + formula_Cl_preliminary!)
Else
    formula_AnionsCorrespondingToCations_F! = formula_F_preliminary!
End If

If formula_F_preliminary! + formula_Cl_preliminary! > 2 Then
    formula_AnionsCorrespondingToCations_Cl! = 2 * formula_Cl_preliminary! / (formula_F_preliminary! + formula_Cl_preliminary!)
Else
    formula_AnionsCorrespondingToCations_Cl! = formula_Cl_preliminary!
End If

formula_AnionsCorrespondingToCations_sumTotal! = formula_AnionsCorrespondingToCations_sum! + formula_AnionsCorrespondingToCations_H! + formula_AnionsCorrespondingToCations_F! * 0.5 + formula_AnionsCorrespondingToCations_Cl! * 0.5


Dim formula_AnionsCorrespondingToCations_checkChargeBalanceCations As Single
Dim formula_AnionsCorrespondingToCations_checkChargeBalanceAnions As Single
formula_AnionsCorrespondingToCations_checkChargeBalanceCations! = 4 * (formula!(1) + formula!(3) + formula!(4)) + 3 * (formula!(5) + formula!(6) + formula!(7) + formula!(8) + formula!(10) + formula!(12)) + 2 * (formula!(9) + formula!(11) + formula!(13) + formula!(14) + formula!(15) + formula!(16) + formula!(17) + formula!(18) + formula!(19) + formula!(22)) + formula!(20) + formula!(21) + formula!(23) + formula_H_mesured_and_rounded!
formula_AnionsCorrespondingToCations_checkChargeBalanceAnions! = -2 * (formula_AnionsCorrespondingToCations_sum! - 0.5 * formula_AnionsCorrespondingToCations_F! - 0.5 * formula_AnionsCorrespondingToCations_Cl!) - 2 * formula_AnionsCorrespondingToCations_H! - formula_AnionsCorrespondingToCations_F! - formula_AnionsCorrespondingToCations_Cl!


Dim formula_AnionsCorrespondingToCations_correctedMn2 As Single
Dim formula_AnionsCorrespondingToCations_correctedMn3 As Single
Dim formula_AnionsCorrespondingToCations_correctedFe2 As Single
Dim formula_AnionsCorrespondingToCations_correctedFe3 As Single

If USE_INITIAL_M3_OVER_SUM_M = False Then
    If formula!(11) > 0 And formula!(10) > 0 Then
        If formula!(10) > formula!(11) Then
            formula_AnionsCorrespondingToCations_correctedFe2! = 0
        Else
            formula_AnionsCorrespondingToCations_correctedFe2! = 24 * (formula!(11) - formula!(10)) / formula_AnionsCorrespondingToCations_sumTotal!
        End If
    Else
        formula_AnionsCorrespondingToCations_correctedFe2! = 24 * formula!(11) / formula_AnionsCorrespondingToCations_sumTotal!
    End If
Else
    formula_AnionsCorrespondingToCations_correctedFe2! = 24 / formula_AnionsCorrespondingToCations_sumTotal! * formula!(11)
End If

If USE_INITIAL_M3_OVER_SUM_M = False Then
    If formula_AnionsCorrespondingToCations_correctedFe2! = 0 Then
        formula_AnionsCorrespondingToCations_correctedMn2! = 24 * (formula!(9) + formula!(11)) / formula_AnionsCorrespondingToCations_sumTotal!
    Else
        formula_AnionsCorrespondingToCations_correctedMn2! = 24 * (formula!(9) + formula!(10)) / formula_AnionsCorrespondingToCations_sumTotal!
    End If
Else
    formula_AnionsCorrespondingToCations_correctedMn2! = 24 / formula_AnionsCorrespondingToCations_sumTotal! * formula!(9)
End If

If USE_INITIAL_M3_OVER_SUM_M = False Then
    If formula_AnionsCorrespondingToCations_correctedFe2! = 0 Then
        formula_AnionsCorrespondingToCations_correctedMn3! = 24 * (formula!(10) - formula!(11)) / formula_AnionsCorrespondingToCations_sumTotal!
    Else
        formula_AnionsCorrespondingToCations_correctedMn3! = 0
    End If
Else
    formula_AnionsCorrespondingToCations_correctedMn3! = 24 / formula_AnionsCorrespondingToCations_sumTotal! * formula!(10)
End If

If USE_INITIAL_M3_OVER_SUM_M = False Then
    If formula_AnionsCorrespondingToCations_correctedFe2! = 0 Then
        formula_AnionsCorrespondingToCations_correctedFe3! = 24 * (formula!(12) + formula!(11)) / formula_AnionsCorrespondingToCations_sumTotal!
    Else
        formula_AnionsCorrespondingToCations_correctedFe3! = 24 * (formula!(12) + formula!(10)) / formula_AnionsCorrespondingToCations_sumTotal!
    End If
Else
    formula_AnionsCorrespondingToCations_correctedFe3! = 24 / formula_AnionsCorrespondingToCations_sumTotal! * formula!(12)
End If


' Formula normalized
Dim formulaNormalized(1 To 24) As Single
For n% = 1 To 23
    If n% = 9 Then
        formulaNormalized!(n%) = formula_AnionsCorrespondingToCations_correctedMn2!
    ElseIf n% = 10 Then
        formulaNormalized!(n%) = formula_AnionsCorrespondingToCations_correctedMn3!
    ElseIf n% = 11 Then
        formulaNormalized!(n%) = formula_AnionsCorrespondingToCations_correctedFe2!
    ElseIf n% = 12 Then
        formulaNormalized!(n%) = formula_AnionsCorrespondingToCations_correctedFe3!
    Else
        formulaNormalized!(n%) = 24 * formula!(n%) / formula_AnionsCorrespondingToCations_sumTotal!
    End If
Next

formulaNormalized!(24) = 24# * formula_H_mesured_and_rounded! / formula_AnionsCorrespondingToCations_sumTotal!

Dim formulaNormalized_sumSitoK As Single
Dim formulaNormalized_sumSitoCa As Single
Dim formulaNormalized_sumSitoMg As Single
Dim formulaNormalized_sumSitoNa As Single
formulaNormalized_sumSitoMg! = 0
For n% = 1 To 17
    formulaNormalized_sumSitoMg! = formulaNormalized_sumSitoMg! + formulaNormalized!(n%)
Next
formulaNormalized_sumSitoMg! = formulaNormalized_sumSitoMg! + formulaNormalized!(20)
formulaNormalized_sumSitoCa! = formulaNormalized_sumSitoMg! + formulaNormalized!(18) + formulaNormalized!(19)
formulaNormalized_sumSitoNa! = formulaNormalized_sumSitoCa! + formulaNormalized!(21)
formulaNormalized_sumSitoK! = formulaNormalized_sumSitoNa! + formulaNormalized!(22) + formulaNormalized!(23)

Dim formulaNormalized_Fe3overSumFe As Single
Dim formulaNormalized_Mn3overSumMn As Single
If formulaNormalized!(11) + formulaNormalized!(12) > 0 Then
    formulaNormalized_Fe3overSumFe! = formulaNormalized!(12) / (formulaNormalized!(11) + formulaNormalized!(12))
Else
    formulaNormalized_Fe3overSumFe! = 0
End If
If formulaNormalized!(9) + formulaNormalized!(10) > 0 Then
    formulaNormalized_Mn3overSumMn! = formulaNormalized!(10) / (formulaNormalized!(9) + formulaNormalized!(10))
Else
    formulaNormalized_Mn3overSumMn! = 0
End If


Dim formulaNormalized_AnionsCorrespondingToCations(1 To 24) As Single
Dim formulaNormalized_AnionsCorrespondingToCations_sum As Single
Dim formulaNormalized_AnionsCorrespondingToCations_F As Single
Dim formulaNormalized_AnionsCorrespondingToCations_Cl As Single

formulaNormalized_AnionsCorrespondingToCations_sum! = 0
For n% = 1 To 24
    ip% = IPOS2DQ%(nelements%, elementList%(n%), AtomicNumbers%(), DisableQuantFlags%())
    If ip% = 0 Then
        formulaNormalized_AnionsCorrespondingToCations!(n%) = 0
    ElseIf n% = 10 Or n% = 12 Then
        formulaNormalized_AnionsCorrespondingToCations!(n%) = formulaNormalized!(n%) * 1.5
        formulaNormalized_AnionsCorrespondingToCations_sum! = formulaNormalized_AnionsCorrespondingToCations_sum! + formulaNormalized_AnionsCorrespondingToCations!(n%)
    Else
        formulaNormalized_AnionsCorrespondingToCations!(n%) = formulaNormalized!(n%) * NumOxds%(ip%) / NumCats%(ip%)
        formulaNormalized_AnionsCorrespondingToCations_sum! = formulaNormalized_AnionsCorrespondingToCations_sum! + formulaNormalized_AnionsCorrespondingToCations!(n%)
    End If
Next

If formula_AnionsCorrespondingToCations_sumTotal! > 24 Then
    formulaNormalized_AnionsCorrespondingToCations_F! = 24 * formula_AnionsCorrespondingToCations_F! / formula_AnionsCorrespondingToCations_sumTotal!
Else
    formulaNormalized_AnionsCorrespondingToCations_F! = formula_AnionsCorrespondingToCations_F!
End If
If formula_AnionsCorrespondingToCations_sumTotal! > 24 Then
    formulaNormalized_AnionsCorrespondingToCations_Cl! = 24 * formula_AnionsCorrespondingToCations_Cl! / formula_AnionsCorrespondingToCations_sumTotal!
Else
    formulaNormalized_AnionsCorrespondingToCations_Cl! = formula_AnionsCorrespondingToCations_Cl!
End If

formulaNormalized_AnionsCorrespondingToCations_sum! = formulaNormalized_AnionsCorrespondingToCations_sum! + 0.5 * formulaNormalized_AnionsCorrespondingToCations_F! + 0.5 * formulaNormalized_AnionsCorrespondingToCations_Cl!

Dim molarMassOfEmpiricalFormula As Single
molarMassOfEmpiricalFormula! = 0
For n% = 1 To 24
    ip% = IPOS2DQ%(nelements%, elementList%(n%), AtomicNumbers%(), DisableQuantFlags%())
    If ip% <> 0 Then
        molarMassOfEmpiricalFormula! = molarMassOfEmpiricalFormula! + formulaNormalized!(n%) * (NumCats%(ip%) * AtomicWeights!(ip%) + NumOxds%(ip%) * AllAtomicWts!(ATOMIC_NUM_OXYGEN%)) / NumCats%(ip%)
    End If
Next

 ' Adding F
ip% = IPOS2DQ%(nelements%, 9, AtomicNumbers%(), DisableQuantFlags%())
If ip% <> 0 Then
    molarMassOfEmpiricalFormula! = molarMassOfEmpiricalFormula! + formulaNormalized_AnionsCorrespondingToCations_F! * AtomicWeights!(ip%)
End If

' Adding Cl
ip% = IPOS2DQ%(nelements%, 17, AtomicNumbers%(), DisableQuantFlags%())
If ip% <> 0 Then
    molarMassOfEmpiricalFormula! = molarMassOfEmpiricalFormula! + formulaNormalized_AnionsCorrespondingToCations_Cl! * AtomicWeights!(ip%)
End If

molarMassOfEmpiricalFormula! = molarMassOfEmpiricalFormula! - formulaNormalized_AnionsCorrespondingToCations_F! * 15.9994 * 0.5 - formulaNormalized_AnionsCorrespondingToCations_Cl! * 15.9994 * 0.5

' Final results to be returned by the function
If REQUIRE_INITIAL_H2O = True Then
    finalWtPercentValues_H2O! = oxidesElements!(24)
Else
    Dim molarMass_H2O As Single
    Dim molarMass_SiO2 As Single
    molarMass_H2O! = 18.01528
    molarMass_SiO2! = 60.0843
    finalWtPercentValues_H2O! = Round((formulaNormalized!(24) / 2 * molarMass_H2O! / molarMassOfEmpiricalFormula! * 100) * oxidesElements!(1) / (formulaNormalized!(1) * molarMass_SiO2! / molarMassOfEmpiricalFormula! * 100), 3)
End If

Fe3overSumFe! = formulaNormalized_Fe3overSumFe!
Mn3overSumMn! = formulaNormalized_Mn3overSumMn!

Exit Sub

' Errors
ConvertAmphiboleCalculationLoopError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertAmphibolesCalculationLoop"
ierror = True
Exit Sub

End Sub

Sub ConvertBubbleSort(ByRef pvarArray As Variant)
' Sort the elements of the array pvararray by increasing value

ierror = False
On Error GoTo ConvertBubbleSortError

Dim i As Long
Dim imin As Long
Dim imax As Long
Dim varSwap As Variant
Dim blnSwapped As Boolean

imin& = LBound(pvarArray)
imax& = UBound(pvarArray) - 1
Do
    blnSwapped = False
    For i& = imin& To imax&
        If pvarArray(i) > pvarArray(i + 1) Then
            varSwap = pvarArray(i)
            pvarArray(i) = pvarArray(i + 1)
            pvarArray(i + 1) = varSwap
            blnSwapped = True
        End If
    Next i&
    imax& = imax& - 1
Loop Until Not blnSwapped

Exit Sub

' Errors
ConvertBubbleSortError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertBubbleSort"
ierror = True
Exit Sub

End Sub

Function ConvertCalculateZFractionExponent(keV As Single) As Single
' Calculate the Z fraction exponent based on the passed emission line energy in keV (from Moy et al. 2021)

ierror = False
On Error GoTo ConvertCalculateZFractionExponentError

Dim exponent As Single

ConvertCalculateZFractionExponent! = 0.7    ' default to 0.7

' Calculate Z Fraction exponent
If keV! > 0.1 And keV! < 30# Then
exponent! = -0.001209 * keV! ^ 2 + 0.03015 * keV! + 0.5908
ConvertCalculateZFractionExponent! = exponent!
End If

Exit Function

' Errors
ConvertCalculateZFractionExponentError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertCalculateZFractionExponent"
ierror = True
Exit Function

End Function

Function ConvertCalculateZFractionExponentBSE(keV As Single) As Single
' Calculate the Z fraction exponent based on the passed electron beam energy in keV

ierror = False
On Error GoTo ConvertCalculateZFractionExponentBSEError

Dim exponent As Single

ConvertCalculateZFractionExponentBSE! = 0.7         ' default to 0.7 (e.g., zero keV for unanalyzed elements)

' Calculate Z Fraction exponent (fit Penepma BSE data for optimized exponent fit as a function of electron beam energy by Moy)
If keV! > 0.1 And keV! < 40# Then
exponent! = -0.000307143 * keV! ^ 2 + 0.0196071 * keV! + 0.475
ConvertCalculateZFractionExponentBSE! = exponent!
End If

Exit Function

' Errors
ConvertCalculateZFractionExponentBSEError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertCalculateZFractionExponentBSE"
ierror = True
Exit Function

End Function


