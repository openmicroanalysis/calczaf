Attribute VB_Name = "CodeUpdate3"
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

Function UpdateGetMaxInterfAssign(sample() As TypeSample) As Integer
' Return the maximum number of interference assignments in the sample
Dim i As Integer, j As Integer
Dim imax As Integer

imax% = 0
For i% = 1 To sample(1).LastElm%
For j% = 1 To MAXINTF%
If sample(1).StdAssignsIntfStds%(j%, i%) > 0 Then
If j% > imax% Then imax% = j%
End If
Next j%
Next i%

UpdateGetMaxInterfAssign% = imax%
Exit Function

End Function

Function UpdateGetMaxMANAssign(sample() As TypeSample) As Integer
' Return the maximum number of MAN assignments in the sample
Dim i As Integer, j As Integer '
Dim imax As Integer

imax% = 0
For i% = 1 To sample(1).LastElm%
For j% = 1 To MAXMAN%
If sample(1).MANStdAssigns(j%, i%) > 0 Then
If j% > imax% Then imax% = j%
End If
Next j%
Next i%

UpdateGetMaxMANAssign% = imax%
Exit Function

End Function

Sub UpdateAddCalculatedOxygen(elementadded As Boolean, sample() As TypeSample)
' Add calculated oxygen to the sample if not already present

ierror = False
On Error GoTo UpdateAddCalculatedOxygenError

Dim ip As Integer

If sample(1).OxideOrElemental% = 1 Then
ip% = IPOS1(sample(1).LastChan%, Symlo$(ATOMIC_NUM_OXYGEN%), sample(1).Elsyms$())
If ip% = 0 Then
If sample(1).LastChan% + 1 > MAXCHAN% Then GoTo UpdateAddCalculatedOxygenTooManyElements
elementadded = True
sample(1).LastChan% = sample(1).LastChan% + 1
sample(1).Elsyms$(sample(1).LastChan%) = Symlo$(ATOMIC_NUM_OXYGEN%)
sample(1).Xrsyms$(sample(1).LastChan%) = vbNullString
sample(1).numcat%(sample(1).LastChan%) = AllCat%(ATOMIC_NUM_OXYGEN%)
sample(1).numoxd%(sample(1).LastChan%) = AllOxd%(ATOMIC_NUM_OXYGEN%)
sample(1).AtomicCharges!(sample(1).LastChan%) = AllAtomicCharges!(ATOMIC_NUM_OXYGEN%)
End If
End If

' If elements were added to the sample, reload unknown sample element setup
If elementadded Then
If VerboseMode Then Call IOWriteLog("Elements added To element list, reloading sample arrays...")
Call ElementGetData(sample())
If ierror Then Exit Sub
End If

Exit Sub

' Errors
UpdateAddCalculatedOxygenError:
MsgBox Error$, vbOKOnly + vbCritical, "UpdateAddCalculatedOxygen"
ierror = True
Exit Sub

UpdateAddCalculatedOxygenTooManyElements:
msg$ = "Too many elements in sample number " & Str$(sample(1).number%)
MsgBox msg$, vbOKOnly + vbExclamation, "UpdateAddCalculatedOxygen"
ierror = True
Exit Sub

End Sub
