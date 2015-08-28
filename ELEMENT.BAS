Attribute VB_Name = "CodeELEMENT"
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

Sub ElementCheckElement(sample() As TypeSample)
' This routine will check for valid elements and xray symbols

ierror = False
On Error GoTo ElementCheckElementError

Dim i As Integer, chan As Integer, ip As Integer
Dim sym As String

' Check elements
For i% = 1 To MAXCHAN%
chan% = i%
sym$ = sample(1).Elsyms$(chan%)
If Trim$(sym) = vbNullString Then GoTo 2000
ip% = IPOS1(MAXELM%, sym, Symlo$())

If ip% = 0 Then
msg$ = TypeLoadString$(sample())
msg$ = "Error in ElementCheckElement: Element " & sym$ & " is not a legal element on channel " & Format$(chan%) & " for sample " & msg$
MsgBox msg$, vbOKOnly + vbExclamation, "ElementCheckElement"
ierror = True
Exit Sub
End If
2000:  Next i%

' Check for zero analyzed elements
If sample(1).LastElm% < 1 And Not sample(1).EDSSpectraFlag Then
msg$ = TypeLoadString$(sample())
msg$ = "Error in ElementCheckElement: No analyzed elements have been entered for sample " & msg$
MsgBox msg$, vbOKOnly + vbExclamation, "ElementCheckElement"
ierror = True
Exit Sub
End If
 
' Check for valid xray symbols for analyzed elements
For i% = 1 To sample(1).LastElm%
chan% = i%
sym$ = sample(1).Xrsyms$(chan%)
ip% = IPOS1(MAXRAY% - 1, sym$, Xraylo$())
If ip% = 0 Then
msg$ = TypeLoadString$(sample())
msg$ = "Error in ElementCheckElement: X-ray " & sym$ & " is not a legal xray line on channel " & Format$(chan%) & " for sample " & msg$
MsgBox msg$, vbOKOnly + vbExclamation, "ElementCheckElement"
ierror = True
Exit Sub
End If
Next i%

Exit Sub

' Errors
ElementCheckElementError:
MsgBox Error$, vbOKOnly + vbCritical, "ElementCheckElement"
ierror = True
Exit Sub

End Sub

Sub ElementCheckXray(mode As Integer, sample() As TypeSample)
' This routine loads the xray lines and checks for bad xray lines or overvoltages
' mode = 0 silent check (no msgbox)
' mode = 1 verbose check (msgbox)

ierror = False
On Error GoTo ElementCheckXrayError

Dim i As Integer, chan As Integer, nrec As Integer
Dim jnum As Integer, ip As Integer
Dim edge As Single, overvolt As Single
Dim sym As String
 
Dim engrow As TypeEnergy
Dim edgrow As TypeEdge

If VerboseMode Then Call IOWriteLog("Checking x-ray lines...")

' Open x-ray edge file
Open XEdgeFile$ For Random Access Read As #XEdgeFileNumber% Len = XRAY_FILE_RECORD_LENGTH%

' Open x-ray line file
Open XLineFile$ For Random Access Read As #XLineFileNumber% Len = XRAY_FILE_RECORD_LENGTH%

' Check xray line selections, load as absorber if specified or a problem is found
For i% = 1 To sample(1).LastChan%
chan% = i%
If chan% > sample(1).LastElm% Then GoTo AbsorberOnly
sym$ = sample(1).Xrsyms$(chan%)
ip% = IPOS1(MAXRAY% - 1, sym$, Xraylo$())
If ip% = 0 Then GoTo AbsorberOnly
sample(1).XrayNums%(chan%) = ip

' Check for disable quant (don't force absorber only)
If sample(1).DisableQuantFlag%(chan%) = 1 Then GoTo DisableQuant

' Check for hydrogen or helium
If sample(1).AtomicNums%(chan%) = 1 Or sample(1).AtomicNums%(chan%) = 2 Then GoTo AbsorberOnly

' Check for bad xray lines (no data)
nrec% = sample(1).AtomicNums%(chan%) + 2
Get #XLineFileNumber%, nrec%, engrow
If engrow.energy!(sample(1).XrayNums%(chan%)) <= 0# Then
msg$ = "Warning in ElementCheckXray: No x-ray emission data for " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%)
If mode% = 1 Then
MsgBox msg$, vbOKOnly + vbExclamation, "ElementCheckXray"
Else
Call IOWriteLog(msg$)
End If
ierror = True
GoTo AbsorberOnly
End If

' Load sample array (in eV)
sample(1).LineEnergy!(chan%) = engrow.energy!(sample(1).XrayNums%(chan%))
Get #XEdgeFileNumber%, nrec%, edgrow

' Calculate edge index for this x-ray
If sample(1).XrayNums%(chan%) = 1 Then jnum% = 1   ' Ka
If sample(1).XrayNums%(chan%) = 2 Then jnum% = 1   ' Kb
If sample(1).XrayNums%(chan%) = 3 Then jnum% = 4   ' La
If sample(1).XrayNums%(chan%) = 4 Then jnum% = 3   ' Lb
If sample(1).XrayNums%(chan%) = 5 Then jnum% = 9   ' Ma
If sample(1).XrayNums%(chan%) = 6 Then jnum% = 8   ' Mb

If edgrow.energy!(jnum%) <= 0# Then
msg$ = "Warning in ElementCheckXray: No x-ray absorption edge data for " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & "(edge index= " & Format$(jnum%) & ")"
If mode% = 1 Then
MsgBox msg$, vbOKOnly + vbExclamation, "ElementCheckXray"
Else
Call IOWriteLog(msg$)
End If
ierror = True
GoTo AbsorberOnly
End If

' Calculate overvoltage
sample(1).LineEdge!(chan%) = edgrow.energy!(jnum%)
edge! = edgrow.energy!(jnum%) / EVPERKEV#
overvolt! = sample(1).KilovoltsArray!(chan%) / edge!

If overvolt! <= 1# Then
msg$ = "Error in ElementCheckXray: Overvoltage of " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & ", edge energy " & Format$(edge!) & " KeV, is " & Format$(overvolt!) & ". Quantification is disabled for this element. Please increase the operating voltage and try again."
If mode% = 1 Then
MsgBox msg$, vbOKOnly + vbExclamation, "ElementCheckXray"
Else
Call IOWriteLog(msg$)
End If
GoTo DisableQuant
End If

If overvolt! <= 1.1 Then
msg$ = "Warning in ElementCheckXray: Overvoltage of " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & ", edge energy " & Format$(edge!) & " KeV, is " & Format$(overvolt!)
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
End If
GoTo 3000

' Flag as absorber only
AbsorberOnly:
sample(1).XrayNums%(chan%) = MAXRAY%  ' absorber only is MAXRAY%
sample(1).Xrsyms$(chan%) = Xraylo$(MAXRAY%)
GoTo 3000

DisableQuant:
sample(1).DisableQuantFlag%(chan%) = 1  ' 0=enable, 1=disable
3000:  Next i%

Close #XEdgeFileNumber%
Close #XLineFileNumber%

Exit Sub

' Errors
ElementCheckXrayError:
MsgBox Error$, vbOKOnly + vbCritical, "ElementCheckXray"
Close #XEdgeFileNumber%
Close #XLineFileNumber%
ierror = True
Exit Sub

End Sub

Sub ElementGetData(sample() As TypeSample)
' This routine calls ElementCheckElement to check the element symbol and then calls ElementLoadArrays
' to loads analytical arrays, followed by a call to ElementCheckXray to check the x-ray
' symbols.

ierror = False
On Error GoTo ElementGetDataError

' Check for at least one analyzed element
If sample(1).LastElm% < 1 And Not sample(1).EDSSpectraFlag Then
msg$ = TypeLoadString$(sample())
msg$ = "No analyzed elements have been entered for sample " & msg$
MsgBox msg$, vbOKOnly + vbExclamation, "ElementGetData"
ierror = True
Exit Sub
End If

' Check for valid element symbols
Call ElementCheckElement(sample())
If ierror Then Exit Sub

' Fill element arrays
Call ElementLoadArrays(sample())
If ierror Then Exit Sub

' Check for valid xray symbols
Call ElementCheckXray(Int(1), sample())
If ierror Then Exit Sub

Exit Sub

' Errors
ElementGetDataError:
MsgBox Error$, vbOKOnly + vbCritical, "ElementGetData"
ierror = True
Exit Sub

End Sub

Sub ElementGetSymbols(sym As String, cat As Integer, oxd As Integer, num As Integer, oxup As String, elup As String)
' This routine accepts an atomic symbol (sym), returns the integer
' position of that element (num) in the symlo array and the oxide
' formula name in "oxup" (the oxide formula) and the upper case
' elemental symbol in "elup".

ierror = False
On Error GoTo ElementGetSymbolsError

num% = 0
oxup$ = vbNullString
elup$ = vbNullString

' Check for a valid element symbol
num% = IPOS1(MAXELM%, sym$, Symlo$())
If num% = 0 Then
msg$ = "Element " & sym$ & " is not a legal element symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "ElementGetSymbols"
ierror = True
Exit Sub
End If

' Create the cation-oxygen formula
elup$ = Trim$(Symup$(num%))
oxup$ = Trim$(elup$)
If cat% > 1 Then
oxup$ = Trim$(elup$) + Trim$(Format$(cat%))
End If

If oxd% > 0 Then
oxup$ = Trim$(oxup$) & Trim$(Symup$(8))
If oxd% > 1 Then
oxup$ = Trim$(oxup$) + Trim$(Format$(oxd%))
End If
End If

Exit Sub

' Errors
ElementGetSymbolsError:
MsgBox Error$, vbOKOnly + vbCritical, "ElementGetSymbols"
ierror = True
Exit Sub

End Sub

Function ElementGetFormula(num As Integer) As String
' This routine accepts an atomic number and returns the default oxide formula name (the oxide formula)

ierror = False
On Error GoTo ElementGetFormulaError

Dim cat As Integer, oxd As Integer
Dim elup As String, oxup As String

ElementGetFormula$ = vbNullString

' Check for a valid element
If num% < 1 Or num% > MAXELM% Then GoTo ElementGetFormulaBadElement

' Get number of cations and oxygens
cat% = AllCat%(num%)
oxd% = AllOxd%(num%)

' Create the cation-oxygen formula
elup$ = Trim$(Symup$(num%))
oxup$ = Trim$(elup$)
If cat% > 1 Then
oxup$ = Trim$(elup$) + Trim$(Format$(cat%))
End If

If oxd% > 0 Then
oxup$ = Trim$(oxup$) & Trim$(Symup$(8))
If oxd% > 1 Then
oxup$ = Trim$(oxup$) + Trim$(Format$(oxd%))
End If
End If

ElementGetFormula$ = oxup$
Exit Function

' Errors
ElementGetFormulaError:
MsgBox Error$, vbOKOnly + vbCritical, "ElementGetFormula"
ierror = True
Exit Function

ElementGetFormulaBadElement:
msg$ = "Element " & Format$(num%) & " is not a legal element number"
MsgBox msg$, vbOKOnly + vbExclamation, "ElementGetFormula"
ierror = True
Exit Function

End Function

Sub ElementLoadArrays(sample() As TypeSample)
' This routine loads the atomic weights, atomic numbers, formula coefficients for oxide formulas, etc.

ierror = False
On Error GoTo ElementLoadArraysError

Dim chan As Integer, num As Integer
Dim oxup As String, elup As String
Dim sym As String
Dim cat As Integer, oxd As Integer

' Load analyzed element symbols, etc.
If VerboseMode Then Call IOWriteLog(vbCrLf & "Loading element arrays...")

For chan% = 1 To sample(1).LastChan%
sample(1).Oxsyup$(chan%) = vbNullString
sample(1).Elsyup$(chan%) = vbNullString
sample(1).AtomicWts!(chan%) = 0#
sample(1).AtomicNums%(chan%) = 0
sample(1).AtomicCharges!(chan%) = 0#

' Load default conditions
If sample(1).takeoff! = 0# Then sample(1).takeoff! = DefaultTakeOff!
If sample(1).kilovolts! = 0# Then sample(1).kilovolts! = DefaultKiloVolts!
If sample(1).TakeoffArray!(chan%) = 0# Then sample(1).TakeoffArray!(chan%) = sample(1).takeoff!
If sample(1).KilovoltsArray!(chan%) = 0# Then sample(1).KilovoltsArray!(chan%) = sample(1).kilovolts!

' Get the element number and formula symbols
sym$ = sample(1).Elsyms$(chan%)
cat% = sample(1).numcat%(chan%)
oxd% = sample(1).numoxd%(chan%)
Call ElementGetSymbols(sym$, cat%, oxd%, num%, oxup$, elup$)
If ierror Then Exit Sub

' Load other arrays
If num% > 0 Then
sample(1).Oxsyup$(chan%) = oxup$
sample(1).Elsyup$(chan%) = elup$
sample(1).AtomicNums%(chan%) = AllAtomicNums%(num%)
sample(1).AtomicWts!(chan%) = AllAtomicWts!(num%)
sample(1).AtomicCharges!(chan%) = AllAtomicCharges!(num%)
End If

Next chan%

Exit Sub

' Errors
ElementLoadArraysError:
MsgBox Error$, vbOKOnly + vbCritical, "ElementLoadArrays"
ierror = True
Exit Sub

End Sub
