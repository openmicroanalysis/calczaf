Attribute VB_Name = "CodeGETZAF"
' (c) Copyright 1995-2018 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub GetZAFLoad()
' Load FormGetZAF (ZAF and Phi-Rho-Z selections)

ierror = False
On Error GoTo GetZAFLoadError

Dim i As Integer

' Load the ZAF selection strings in the form
For i% = 0 To UBound(zafstring$)    ' note izaf is indexed from zero (0 = individual selections)
FormGETZAF.OptionZaf(i%).Caption = zafstring(i%)
If i% = izaf% Then FormGETZAF.OptionZaf(i%).value = True
Next i%

' Load the ZAF selection strings in the form
For i% = 1 To UBound(mipstring$)
FormGETZAF.OptionMip(i% - 1).Caption = mipstring(i%)
If i% = imip% Then FormGETZAF.OptionMip(i% - 1).value = True
Next i%

For i% = 1 To UBound(bscstring$)
FormGETZAF.OptionBsc(i% - 1).Caption = bscstring(i%)
If i% = ibsc Then FormGETZAF.OptionBsc(i% - 1).value = True
Next i%

For i% = 1 To UBound(phistring$)
FormGETZAF.OptionPhi(i% - 1).Caption = phistring(i%)
If i% = iphi% Then FormGETZAF.OptionPhi(i% - 1).value = True
Next i%

For i% = 1 To UBound(stpstring$)
FormGETZAF.OptionStp(i% - 1).Caption = stpstring(i%)
If i% = istp Then FormGETZAF.OptionStp(i% - 1).value = True
Next i%

For i% = 0 To UBound(bksstring$)    ' note ibks is indexed from zero
FormGETZAF.OptionBks(i%).Caption = bksstring(i%)
If i% = ibks Then FormGETZAF.OptionBks(i%).value = True
Next i%

For i% = 1 To UBound(absstring$)
FormGETZAF.OptionAbs(i% - 1).Caption = absstring(i%)
If i% = iabs% Then FormGETZAF.OptionAbs(i% - 1).value = True
Next i%

For i% = 1 To UBound(flustring$)
FormGETZAF.OptionFlu(i% - 1).Caption = flustring(i%)
If i% = iflu Then FormGETZAF.OptionFlu(i% - 1).value = True
Next i%

' Set beta fluorescence flag
If UseFluorescenceByBetaLinesFlag Then
FormGETZAF.CheckUseFluorescenceByBetaLines.value = vbChecked
Else
FormGETZAF.CheckUseFluorescenceByBetaLines.value = vbUnchecked
End If

Exit Sub

' Errors
GetZAFLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "GetZAFLoad"
ierror = True
Exit Sub

End Sub

Sub GetZAFSave()
' Save FormGetZAF options

ierror = False
On Error GoTo GetZAFSaveError

Dim i As Integer

' Load the ZAF selection strings in the form
For i% = 0 To UBound(zafstring$)    ' note ibks is indexed from zero
If FormGETZAF.OptionZaf(i%).value Then izaf% = i%
Next i%

' Save the ZAF selection strings from the form
For i% = 1 To UBound(mipstring$)
If FormGETZAF.OptionMip(i% - 1).value Then imip% = i%
Next i%

For i% = 1 To UBound(bscstring$)
If FormGETZAF.OptionBsc(i% - 1).value Then ibsc = i%
Next i%

For i% = 1 To UBound(phistring$)
If FormGETZAF.OptionPhi(i% - 1).value Then iphi% = i%
Next i%

For i% = 1 To UBound(stpstring)
If FormGETZAF.OptionStp(i% - 1).value Then istp = i%
Next i%

For i% = 0 To UBound(bksstring$)    ' note ibks is indexed from zero
If FormGETZAF.OptionBks(i%).value Then ibks = i%
Next i%

For i% = 1 To UBound(absstring$)
If FormGETZAF.OptionAbs(i% - 1).value Then iabs% = i%
Next i%

For i% = 1 To UBound(flustring$)
If FormGETZAF.OptionFlu(i% - 1).value Then iflu = i%
Next i%

' Set beta fluorescence flag
If FormGETZAF.CheckUseFluorescenceByBetaLines.value = vbChecked Then
UseFluorescenceByBetaLinesFlag = True
Else
UseFluorescenceByBetaLinesFlag = False
End If

' If PAP abscor is used, check that PAP stpcor is also used
If iabs% = 12 Or iabs% = 13 Then
If istp% <> 5 Then
msg$ = absstring$(12) & " or " & absstring$(13) & " requires " & stpstring$(5) & " to be selected"
MsgBox msg$, vbOKOnly + vbExclamation, "GetZAFSave"
ierror = True
End If
End If

' If Proza abscor is used, check that PAP stpcor is also used
If iabs% = 15 Then
If istp% <> 5 Then
msg$ = absstring$(15) & " requires " & stpstring$(5) & " to be selected"
MsgBox msg$, vbOKOnly + vbExclamation, "GetZAFSave"
ierror = True
End If
End If

' Check option if particle and thin film option selected
If UseParticleCorrectionFlag And iptc% = 1 Then
Call GetPTCCheck
If ierror Then Exit Sub
End If

' Indicate alpha-factor update
AllAnalysisUpdateNeeded = True
AllAFactorUpdateNeeded = True

Exit Sub

' Errors
GetZAFSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "GetZAFSave"
ierror = True
Exit Sub

End Sub

Sub GetZAFSetZAF()
' Loads a pre-defined ZAF correction

ierror = False
On Error GoTo GetZAFSetZAFError

Dim i As Integer, j As Integer

' Load predefined zaf correction
For j% = 0 To UBound(zafstring$)
If FormGETZAF.OptionZaf(j%).value Then
izaf% = j%
End If
Next j%

' Select Individual Corrections (to maintain backward compatibility with Probe for EPMA)
If izaf% = 0 Then

' Select the default (Armstrong/Love-Scott)ZAF matrix correction algorithms
ElseIf izaf% = 1 Then
i% = 2
FormGETZAF.OptionBsc(i% - 1).value = True
i% = 1
FormGETZAF.OptionMip(i% - 1).value = True
i% = 2
FormGETZAF.OptionPhi(i% - 1).value = True
i% = 9
FormGETZAF.OptionAbs(i% - 1).value = True
i% = 4
FormGETZAF.OptionStp(i% - 1).value = True
i% = 4
FormGETZAF.OptionBks(i%).value = True   ' note ibks is indexed from zero

' Conventional Philibert/Duncumb Reed
ElseIf izaf% = 2 Then
i% = 1
FormGETZAF.OptionBsc(i% - 1).value = True
i% = 2
FormGETZAF.OptionMip(i% - 1).value = True
i% = 5
FormGETZAF.OptionPhi(i% - 1).value = True
i% = 1
FormGETZAF.OptionAbs(i% - 1).value = True
i% = 1
FormGETZAF.OptionStp(i% - 1).value = True
i% = 1
FormGETZAF.OptionBks(i%).value = True   ' note ibks is indexed from zero

' Heinrich/Duncumb-Reed
ElseIf izaf% = 3 Then
i% = 1
FormGETZAF.OptionBsc(i% - 1).value = True
i% = 1
FormGETZAF.OptionMip(i% - 1).value = True
i% = 5
FormGETZAF.OptionPhi(i% - 1).value = True
i% = 1
FormGETZAF.OptionAbs(i% - 1).value = True
i% = 1
FormGETZAF.OptionStp(i% - 1).value = True
i% = 2
FormGETZAF.OptionBks(i%).value = True   ' note ibks is indexed from zero

' Love-Scott I
ElseIf izaf% = 4 Then
i% = 2
FormGETZAF.OptionBsc(i% - 1).value = True
i% = 1
FormGETZAF.OptionMip(i% - 1).value = True
i% = 2
FormGETZAF.OptionPhi(i% - 1).value = True
i% = 4
FormGETZAF.OptionAbs(i% - 1).value = True
i% = 4
FormGETZAF.OptionStp(i% - 1).value = True
i% = 4
FormGETZAF.OptionBks(i%).value = True   ' note ibks is indexed from zero

' Love-Scott II
ElseIf izaf% = 5 Then
i% = 2
FormGETZAF.OptionBsc(i% - 1).value = True
i% = 1
FormGETZAF.OptionMip(i% - 1).value = True
i% = 2
FormGETZAF.OptionPhi(i% - 1).value = True
i% = 6
FormGETZAF.OptionAbs(i% - 1).value = True
i% = 4
FormGETZAF.OptionStp(i% - 1).value = True
i% = 4
FormGETZAF.OptionBks(i%).value = True   ' note ibks is indexed from zero

' Packwood Phi(PZ) (EPQ-91)
ElseIf izaf% = 6 Then
i% = 1
FormGETZAF.OptionBsc(i% - 1).value = True
i% = 5
FormGETZAF.OptionMip(i% - 1).value = True
i% = 7
FormGETZAF.OptionPhi(i% - 1).value = True
i% = 14
FormGETZAF.OptionAbs(i% - 1).value = True
i% = 6
FormGETZAF.OptionStp(i% - 1).value = True
i% = 0
FormGETZAF.OptionBks(i%).value = True   ' note ibks is indexed from zero

' Bastin original Phi(PZ)
ElseIf izaf% = 7 Then
i% = 2
FormGETZAF.OptionBsc(i% - 1).value = True
i% = 3
FormGETZAF.OptionMip(i% - 1).value = True
i% = 2
FormGETZAF.OptionPhi(i% - 1).value = True
i% = 10
FormGETZAF.OptionAbs(i% - 1).value = True
i% = 6
FormGETZAF.OptionStp(i% - 1).value = True
i% = 0
FormGETZAF.OptionBks(i%).value = True   ' note ibks is indexed from zero

' Bastin PROZA Phi(PZ) (EPQ-91)
ElseIf izaf% = 8 Then
i% = 3
FormGETZAF.OptionBsc(i% - 1).value = True
i% = 3
FormGETZAF.OptionMip(i% - 1).value = True
i% = 4
FormGETZAF.OptionPhi(i% - 1).value = True
i% = 15
FormGETZAF.OptionAbs(i% - 1).value = True
i% = 5
FormGETZAF.OptionStp(i% - 1).value = True
i% = 7
FormGETZAF.OptionBks(i%).value = True   ' note ibks is indexed from zero

' Pouchout & Pichoir - Full
ElseIf izaf% = 9 Then
i% = 3
FormGETZAF.OptionBsc(i% - 1).value = True
i% = 3
FormGETZAF.OptionMip(i% - 1).value = True
i% = 4
FormGETZAF.OptionPhi(i% - 1).value = True
i% = 12
FormGETZAF.OptionAbs(i% - 1).value = True
i% = 5
FormGETZAF.OptionStp(i% - 1).value = True
i% = 7
FormGETZAF.OptionBks(i%).value = True   ' note ibks is indexed from zero

' Pouchout & Pichoir - Simplified
ElseIf izaf% = 10 Then
i% = 3
FormGETZAF.OptionBsc(i% - 1).value = True
i% = 3
FormGETZAF.OptionMip(i% - 1).value = True
i% = 4
FormGETZAF.OptionPhi(i% - 1).value = True
i% = 13
FormGETZAF.OptionAbs(i% - 1).value = True
i% = 5
FormGETZAF.OptionStp(i% - 1).value = True
i% = 7
FormGETZAF.OptionBks(i%).value = True   ' note ibks is indexed from zero
End If

Exit Sub

' Errors
GetZAFSetZAFError:
MsgBox Error$, vbOKOnly + vbCritical, "GetZAFSetZAF"
ierror = True
Exit Sub

End Sub

Sub GetZAFSetEnables()
' Sets enables

ierror = False
On Error GoTo GetZAFSetEnablesError

Dim i As Integer

' Enable all
If FormGETZAF.OptionZaf(0).value = True Then

For i% = 1 To UBound(mipstring$)
FormGETZAF.OptionMip(i% - 1).Enabled = True
Next i%

For i% = 1 To UBound(bscstring$)
FormGETZAF.OptionBsc(i% - 1).Enabled = True
Next i%

For i% = 1 To UBound(phistring$)
FormGETZAF.OptionPhi(i% - 1).Enabled = True
Next i%

For i% = 1 To UBound(stpstring)
FormGETZAF.OptionStp(i% - 1).Enabled = True
Next i%

For i% = 0 To UBound(bksstring$)    ' note ibks is indexed from zero
FormGETZAF.OptionBks(i%).Enabled = True
Next i%

For i% = 1 To UBound(absstring$)
FormGETZAF.OptionAbs(i% - 1).Enabled = True
Next i%

' Disable all
Else

For i% = 1 To UBound(mipstring$)
FormGETZAF.OptionMip(i% - 1).Enabled = False
Next i%

For i% = 1 To UBound(bscstring$)
FormGETZAF.OptionBsc(i% - 1).Enabled = False
Next i%

For i% = 1 To UBound(phistring$)
FormGETZAF.OptionPhi(i% - 1).Enabled = False
Next i%

For i% = 1 To UBound(stpstring)
FormGETZAF.OptionStp(i% - 1).Enabled = False
Next i%

For i% = 0 To UBound(bksstring$)    ' note ibks is indexed from zero
FormGETZAF.OptionBks(i%).Enabled = False
Next i%

For i% = 1 To UBound(absstring$)
FormGETZAF.OptionAbs(i% - 1).Enabled = False
Next i%

End If

Exit Sub

' Errors
GetZAFSetEnablesError:
MsgBox Error$, vbOKOnly + vbCritical, "GetZAFSetEnables"
ierror = True
Exit Sub

End Sub
