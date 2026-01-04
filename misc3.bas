Attribute VB_Name = "CodeMISC3"
' (c) Copyright 1995-2026 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Function IPOS5(ByVal mode As Integer, ByVal n As Integer, sample1() As TypeSample, sample2() As TypeSample) As Integer
' This routine returns as its value a pointer to the first occurance in "sample2()" of the element in "sample1()" specified by channel "n".
' If a match does not occur, IPOS5 = 0.
' mode = 0 check only the analyzed elements
' mode = 1 check both analyzed and specified elements

ierror = False
On Error GoTo IPOS5Error

Dim i As Integer, k As Integer

' Fail if not specified
IPOS5 = 0
If n% <= 0 Then Exit Function

If mode% = 0 Then
k% = sample2(1).LastElm%
Else
k% = sample2(1).LastChan%
End If

' Search sample for match (element, x-ray, motor, crystal)
For i% = 1 To k%

If Trim$(UCase$(sample1(1).Elsyms$(n%))) = Trim$(UCase$(sample2(1).Elsyms$(i%))) Then
If Trim$(UCase$(sample1(1).Xrsyms$(n%))) = Trim$(UCase$(sample2(1).Xrsyms$(i%))) Then
If sample1(1).MotorNumbers%(n%) = sample2(1).MotorNumbers%(i%) Then
If Trim$(UCase$(sample1(1).CrystalNames$(n%))) = Trim$(UCase$(sample2(1).CrystalNames$(i%))) Then
IPOS5 = i%
Exit Function
End If
End If
End If
End If

Next i%

Exit Function

' Errors
IPOS5Error:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS5"
ierror = True
Exit Function

End Function

Function IPOS7(ByVal n As Integer, ByVal syme As String, ByVal symx As String, sample() As TypeSample) As Integer
' This routine returns as its value a pointer to the first occurance of the
' element and x-ray in "sample1()" starting at channel "n%". Checks disable flag.
' If a match does not occur, IPOS7 = 0.

ierror = False
On Error GoTo IPOS7Error

Dim i As Integer

' Fail if not specified
IPOS7 = 0

' Search sample for match (element and x-ray only)
For i% = n% To sample(1).LastElm%
If sample(1).DisableQuantFlag%(i%) = 0 Then

If Trim$(UCase$(syme$)) = Trim$(UCase$(sample(1).Elsyms$(i%))) Then
If Trim$(UCase$(symx$)) = Trim$(UCase$(sample(1).Xrsyms$(i%))) Then
IPOS7 = i%
Exit Function
End If
End If

End If
Next i%

Exit Function

' Errors
IPOS7Error:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS7"
ierror = True
Exit Function

End Function

Function IPOS7A(ByVal n As Integer, ByVal syme As String, ByVal symx As String, sample() As TypeSample) As Integer
' This routine returns as its value a pointer to the last occurance of the
' element and x-ray in "sample1()" starting at channel "n%+1".
' If a match does not occur, IPOS7A = 0.

ierror = False
On Error GoTo IPOS7AError

Dim i As Integer

' Fail if not specified
IPOS7A = 0

' Search sample for match (element and x-ray only)
For i% = sample(1).LastElm% To n% + 1 Step -1
If sample(1).DisableQuantFlag%(i%) = 0 Then

If Trim$(UCase$(syme$)) = Trim$(UCase$(sample(1).Elsyms$(i%))) Then
If Trim$(UCase$(symx$)) = Trim$(UCase$(sample(1).Xrsyms$(i%))) Then
IPOS7A = i%
Exit Function
End If
End If

End If
Next i%

Exit Function

' Errors
IPOS7AError:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS7A"
ierror = True
Exit Function

End Function

Function IPOS8(ByVal n As Integer, ByVal syme As String, ByVal symx As String, sample() As TypeSample) As Integer
' This routine returns as its value a pointer to the first occurance of the element and x-ray
' in "sample()" up to (but not including) channel "n%". Checks disable quant flag!!! (added 10/5/07)
' If a match does not occur, IPOS8 = 0.

ierror = False
On Error GoTo IPOS8Error

Dim i As Integer

' Fail if not specified
IPOS8 = 0

' Search sample for match (element and x-ray only)
For i% = 1 To n% - 1
If sample(1).DisableQuantFlag%(i%) = 0 Then

If Trim$(UCase$(syme$)) = Trim$(UCase$(sample(1).Elsyms$(i%))) Then
If Trim$(UCase$(symx$)) = Trim$(UCase$(sample(1).Xrsyms$(i%))) Then
IPOS8 = i%
Exit Function
End If
End If

End If
Next i%

Exit Function

' Errors
IPOS8Error:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS8"
ierror = True
Exit Function

End Function

Function IPOS8A(ByVal n As Integer, ByVal syme As String, ByVal symx As String, ByVal keV As Single, sample() As TypeSample) As Integer
' This routine returns as its value a pointer to the first occurance of the element and x-ray and keV
' in "sample()" up to (but not including) channel "n%". Checks disable quant flag!!! (added 11/12/16)
' If a match does not occur, IPOS8A = 0.

ierror = False
On Error GoTo IPOS8AError

Dim i As Integer

' Fail if not specified
IPOS8A = 0

' Search sample for match (element and x-ray only)
For i% = 1 To n% - 1
If sample(1).DisableQuantFlag%(i%) = 0 Then

If Trim$(UCase$(syme$)) = Trim$(UCase$(sample(1).Elsyms$(i%))) Then
If Trim$(UCase$(symx$)) = Trim$(UCase$(sample(1).Xrsyms$(i%))) Then
If keV! = sample(1).KilovoltsArray!(i%) Then
IPOS8A = i%
Exit Function
End If
End If
End If

End If
Next i%

Exit Function

' Errors
IPOS8AError:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS8A"
ierror = True
Exit Function

End Function

Function IPOS9(ByVal syme As String, sample() As TypeSample) As Integer
' This routine returns as its value a pointer to the first occurance of the
' element in "sample1()". Checks the DisableQuant flag! If no match, IPOS9 = 0.

ierror = False
On Error GoTo IPOS9Error

Dim i As Integer

' Fail if not specified
IPOS9 = 0

' Search sample for match (element and disable quant only)
For i% = 1 To sample(1).LastChan%
If sample(1).DisableQuantFlag%(i%) = 0 Then
If Trim$(UCase$(syme$)) = Trim$(UCase$(sample(1).Elsyms$(i%))) Then
IPOS9 = i%
Exit Function
End If
End If
Next i%

Exit Function

' Errors
IPOS9Error:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS9"
ierror = True
Exit Function

End Function

Function IPOS11(ByVal syme As String, sample() As TypeSample) As Integer
' This routine returns as its value a pointer to the first occurance of the specified element in "sample1()".
' If a match does not occur, IPOS11 = 0.

ierror = False
On Error GoTo IPOS11Error

Dim i As Integer

' Fail if not specified
IPOS11 = 0

' Search sample for match (specified element only)
For i% = 1 To sample(1).LastChan%

' Standard (standards have x-rays specified from standard database)
If sample(1).Type% = 1 Then
If Trim$(UCase$(syme$)) = Trim$(UCase$(sample(1).Elsyms$(i%))) Then
IPOS11 = i%
Exit Function
End If

' Other
Else
If Trim$(UCase$(syme$)) = Trim$(UCase$(sample(1).Elsyms$(i%))) And sample(1).Xrsyms$(i%) = vbNullString Then
IPOS11 = i%
Exit Function
End If
End If

Next i%

Exit Function

' Errors
IPOS11Error:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS11"
ierror = True
Exit Function

End Function

Function IPOS13(ByVal syme As String, ByVal symx As String, ByVal imot As Integer, ByVal crys As String, sample() As TypeSample) As Integer
' Same as IPOS12 but uses a string for x-ray and crystal. Does not check disable quant flag.

ierror = False
On Error GoTo IPOS13Error

Dim i As Integer

' Fail if not specified
IPOS13 = 0

' Search sample for match (element, x-ray, motor, crystal)
For i% = 1 To sample(1).LastChan%
If Trim$(UCase$(syme$)) = Trim$(UCase$(sample(1).Elsyms$(i%))) Then
If Trim$(UCase$(symx$)) = Trim$(UCase$(sample(1).Xrsyms$(i%))) Then
If imot% = sample(1).MotorNumbers%(i%) Then
If Trim$(UCase$(crys$)) = Trim$(UCase$(sample(1).CrystalNames$(i%))) Then
IPOS13 = i%
Exit Function
End If
End If
End If
End If

Next i%

Exit Function

' Errors
IPOS13Error:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS13"
ierror = True
Exit Function

End Function

Function IPOS13A(ByVal syme As String, ByVal symx As String, ByVal imot As Integer, ByVal crys As String, sample() As TypeSample) As Integer
' Same as IPOS12 but uses a string for x-ray and crystal. Checks disable quant flag

ierror = False
On Error GoTo IPOS13AError

Dim i As Integer

' Fail if not specified
IPOS13A = 0

' Search sample for match (element, x-ray, motor, crystal)
For i% = 1 To sample(1).LastChan%
If sample(1).DisableQuantFlag%(i%) = 0 Then

If Trim$(UCase$(syme$)) = Trim$(UCase$(sample(1).Elsyms$(i%))) Then
If Trim$(UCase$(symx$)) = Trim$(UCase$(sample(1).Xrsyms$(i%))) Then
If imot% = sample(1).MotorNumbers%(i%) Then
If Trim$(UCase$(crys$)) = Trim$(UCase$(sample(1).CrystalNames$(i%))) Then
IPOS13A = i%
Exit Function
End If
End If
End If
End If

End If
Next i%

Exit Function

' Errors
IPOS13AError:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS13A"
ierror = True
Exit Function

End Function

Function IPOS13B(ByVal mode As Integer, ByVal syme As String, ByVal symx As String, ByVal imot As Integer, ByVal crys As String, ByVal keV As Single, sample() As TypeSample) As Integer
' Same as IPOS13A but also checks keV (for MAN fits). Checks disable quant flag.
' mode = 0 check only the analyzed elements
' mode = 1 check both analyzed and specified elements

ierror = False
On Error GoTo IPOS13BError

Dim i As Integer, k As Integer

' Fail if not specified
IPOS13B = 0

If mode% = 0 Then
k% = sample(1).LastElm%
Else
k% = sample(1).LastChan%
End If

' Search sample for match (element, x-ray, motor, crystal and keV)
For i% = 1 To k%
If sample(1).DisableQuantFlag%(i%) = 0 Then

If Trim$(UCase$(syme$)) = Trim$(UCase$(sample(1).Elsyms$(i%))) Then
If Trim$(UCase$(symx$)) = Trim$(UCase$(sample(1).Xrsyms$(i%))) Then
If imot% = sample(1).MotorNumbers%(i%) Then
If Trim$(UCase$(crys$)) = Trim$(UCase$(sample(1).CrystalNames$(i%))) Then
If keV! = sample(1).KilovoltsArray!(i%) Then
IPOS13B = i%
Exit Function
End If
End If
End If
End If
End If

End If
Next i%

Exit Function

' Errors
IPOS13BError:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS13B"
ierror = True
Exit Function

End Function

Function IPOS13C(ByVal mode As Integer, ByVal syme As String, ByVal symx As String, ByVal imot As Integer, ByVal crys As String, sample() As TypeSample) As Integer
' Same as IPOS13B but does not check for KilovoltsArray.
' mode = 0 check only the analyzed elements
' mode = 1 check both analyzed and specified elements

ierror = False
On Error GoTo IPOS13CError

Dim i As Integer, k As Integer

' Fail if not specified
IPOS13C = 0

If mode% = 0 Then
k% = sample(1).LastElm%
Else
k% = sample(1).LastChan%
End If

' Search sample for match (element, x-ray, motor, crystal and keV)
For i% = 1 To k%
If sample(1).DisableQuantFlag%(i%) = 0 Then

If Trim$(UCase$(syme$)) = Trim$(UCase$(sample(1).Elsyms$(i%))) Then
If Trim$(UCase$(symx$)) = Trim$(UCase$(sample(1).Xrsyms$(i%))) Then
If imot% = sample(1).MotorNumbers%(i%) Then
If Trim$(UCase$(crys$)) = Trim$(UCase$(sample(1).CrystalNames$(i%))) Then
IPOS13C = i%
Exit Function
End If
End If
End If
End If

End If
Next i%

Exit Function

' Errors
IPOS13CError:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS13C"
ierror = True
Exit Function

End Function

Function IPOS14(ByVal chan As Integer, sample1() As TypeSample, sample2() As TypeSample) As Integer
' Check for first occurance of element, x-ray, takeoff and keV in the sample2 element list (not disabled)

ierror = False
On Error GoTo IPOS14Error

Dim j As Integer

' Fail if not duplicated
IPOS14% = 0

' Search sample for match
For j% = 1 To sample2(1).LastChan%
If sample2(1).DisableQuantFlag%(j%) = 0 Then
If Trim$(UCase$(sample1(1).Elsyms$(chan%))) = Trim$(UCase$(sample2(1).Elsyms$(j%))) Then
If Trim$(UCase$(sample1(1).Xrsyms$(chan%))) = Trim$(UCase$(sample2(1).Xrsyms$(j%))) Then
If sample1(1).TakeoffArray!(chan%) = sample2(1).TakeoffArray!(j%) Then
If sample1(1).KilovoltsArray!(chan%) = sample2(1).KilovoltsArray!(j%) Then
IPOS14% = j%
Exit Function
End If
End If
End If
End If
End If

Next j%

Exit Function

' Errors
IPOS14Error:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS14"
ierror = True
Exit Function

End Function
