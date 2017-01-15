Attribute VB_Name = "CodeMISC5"
' (c) Copyright 1995-2017 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Function MiscAllNegOne(lchan As Integer, narray() As Integer) As Boolean
' Check if integer array is all negative ones

ierror = False
On Error GoTo MiscAllNegOneError

Dim i As Integer

MiscAllNegOne = True
For i% = 1 To lchan%
If narray%(i%) <> -1 Then MiscAllNegOne = False
Next i%

Exit Function

' Errors
MiscAllNegOneError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAllNegOne"
ierror = True
Exit Function

End Function

Function MiscAllZero(lchan As Integer, narray() As Integer) As Boolean
' Check if integer array is all zeros

ierror = False
On Error GoTo MiscAllZeroError

Dim i As Integer

MiscAllZero = True
For i% = 1 To lchan%
If narray%(i%) <> 0 Then MiscAllZero = False
Next i%

Exit Function

' Errors
MiscAllZeroError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAllZero"
ierror = True
Exit Function

End Function

Function MiscAllEqualToPassed(ivalue As Integer, lchan As Integer, narray() As Integer) As Boolean
' Check if integer array is all equal to the passed value (skip zero values)

ierror = False
On Error GoTo MiscAllEqualToPassedError

Dim i As Integer

MiscAllEqualToPassed = True
For i% = 1 To lchan%
If narray%(i%) <> 0 And narray%(i%) <> ivalue% Then MiscAllEqualToPassed = False
Next i%

Exit Function

' Errors
MiscAllEqualToPassedError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAllEqualToPassed"
ierror = True
Exit Function

End Function

Function MiscAllOne(lchan As Integer, narray() As Single) As Boolean
' Check if single array is all ones

ierror = False
On Error GoTo MiscAllOneError

Dim i As Integer

MiscAllOne = True
For i% = 1 To lchan%
If narray!(i%) <> 1# Then MiscAllOne = False
Next i%

Exit Function

' Errors
MiscAllOneError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAllOne"
ierror = True
Exit Function

End Function

Function MiscIsDifferent(lchan As Integer, narray() As Integer) As Boolean
' Check for differences in integer array

ierror = False
On Error GoTo MiscIsDifferentError

Dim i As Integer

MiscIsDifferent = False
For i% = 1 To lchan%
If narray%(i%) <> narray%(1) Then MiscIsDifferent = True
Next i%

' If only one element, set true
If lchan% = 1 Then MiscIsDifferent = True

Exit Function

' Errors
MiscIsDifferentError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscIsDifferent"
ierror = True
Exit Function

End Function

Function MiscIsDifferent2(lchan As Integer, narray() As Long) As Boolean
' Check for differences in long array

ierror = False
On Error GoTo MiscIsDifferent2Error

Dim i As Integer

MiscIsDifferent2 = False
For i% = 1 To lchan%
If narray&(i%) <> narray&(1) Then MiscIsDifferent2 = True
Next i%

' If only one element, set true
If lchan% = 1 Then MiscIsDifferent2 = True

Exit Function

' Errors
MiscIsDifferent2Error:
MsgBox Error$, vbOKOnly + vbCritical, "MiscIsDifferent2"
ierror = True
Exit Function

End Function

Function MiscIsDifferent3(lchan As Integer, narray() As Single) As Boolean
' Check for differences in single array

ierror = False
On Error GoTo MiscIsDifferent3Error

Dim i As Integer

MiscIsDifferent3 = False
For i% = 1 To lchan%
If narray!(i%) <> narray!(1) Then MiscIsDifferent3 = True
Next i%

' If only one element, set true
If lchan% = 1 Then MiscIsDifferent3 = True

Exit Function

' Errors
MiscIsDifferent3Error:
MsgBox Error$, vbOKOnly + vbCritical, "MiscIsDifferent3"
ierror = True
Exit Function

End Function

Function MiscIsDifferent4(lchan As Integer, sarray() As String) As Boolean
' Check for differences in string array

ierror = False
On Error GoTo MiscIsDifferent4Error

Dim i As Integer

MiscIsDifferent4 = False
For i% = 1 To lchan%
If UCase$(Trim$(sarray$(i%))) <> UCase$(Trim$(sarray$(1))) Then MiscIsDifferent4 = True
Next i%

' If only one element, set true
If lchan% = 1 Then MiscIsDifferent4 = True

Exit Function

' Errors
MiscIsDifferent4Error:
MsgBox Error$, vbOKOnly + vbCritical, "MiscIsDifferent4"
ierror = True
Exit Function

End Function

Function MiscAllZero2(n As Integer, lchan As Integer, narray() As Integer) As Integer
' Check if integer array is all zeros (two dimensional arrays)

ierror = False
On Error GoTo MiscAllZero2Error

Dim i As Integer

MiscAllZero2% = True
For i% = 1 To lchan%
If narray%(n%, i%) <> 0 Then MiscAllZero2% = False
Next i%

Exit Function

' Errors
MiscAllZero2Error:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAllZero2"
ierror = True
Exit Function

End Function

Function MiscAllZero22(n As Integer, lchan As Integer, narray() As Single) As Integer
' Check if integer array is all zeros (two dimensional float arrays)

ierror = False
On Error GoTo MiscAllZero22Error

Dim i As Integer, j As Integer

MiscAllZero22% = True
For j% = 1 To n%
For i% = 1 To lchan%
If narray!(n%, i%) <> 0# Then MiscAllZero22% = False
Next i%
Next j%

Exit Function

' Errors
MiscAllZero22Error:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAllZero22"
ierror = True
Exit Function

End Function

Function MiscAllZero3(n As Integer, lchan As Integer, narray() As Integer) As Integer
' Check if integer array is all zeros (two dimensional arrays but loop on first index instead of element channel loop)

ierror = False
On Error GoTo MiscAllZero3Error

Dim i As Integer

MiscAllZero3% = True
For i% = 1 To n%
If narray%(i%, lchan%) <> 0 Then MiscAllZero3% = False
Next i%

Exit Function

' Errors
MiscAllZero3Error:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAllZero3"
ierror = True
Exit Function

End Function

Sub MiscGetArrayMinMax(n As Long, sarray() As Single, imin As Long, imax As Long, tmin As Single, tmax As Single)
' Return the min and max array indices found in a one dimensional single precision array

ierror = False
On Error GoTo MiscGetArrayMinMaxError

Dim i As Long

imin& = 0
imax& = 0

tmin! = MAXMINIMUM!
tmax! = MAXMAXIMUM!

For i& = 1 To n&
If sarray!(i&) < tmin! Then
tmin! = sarray!(i&)
imin& = i&
End If

If sarray!(i&) > tmax! Then
tmax! = sarray!(i&)
imax& = i&
End If
Next i&

Exit Sub

' Errors
MiscGetArrayMinMaxError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetArrayMinMax"
ierror = True
Exit Sub

End Sub

Sub MiscGetArrayMinMaxZero(n As Long, sarray() As Single, imin As Long, imax As Long, tmin As Single, tmax As Single)
' Return the min and max array indices found in a one dimensional single precision array (indexed 0 to n - 1)

ierror = False
On Error GoTo MiscGetArrayMinMaxZeroError

Dim i As Long

imin& = 0
imax& = 0

tmin! = MAXMINIMUM!
tmax! = MAXMAXIMUM!

For i& = 0 To n& - 1
If sarray!(i&) < tmin! Then
tmin! = sarray!(i&)
imin& = i&
End If

If sarray!(i&) > tmax! Then
tmax! = sarray!(i&)
imax& = i&
End If
Next i&

Exit Sub

' Errors
MiscGetArrayMinMaxZeroError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetArrayMinMaxZero"
ierror = True
Exit Sub

End Sub

Sub MiscGetArrayMinMaxByte(n As Long, barray() As Byte, imin As Long, imax As Long, bmin As Byte, bmax As Byte)
' Return the min and max array indices found in a one dimensional byte array (indexed 0 to n-1)

ierror = False
On Error GoTo MiscGetArrayMinMaxByteError

Dim i As Integer

imin& = 0
imax& = 0

For i% = 0 To n& - 1
If barray(i%) < bmin Then
bmin = barray(i%)
imin& = CLng(i%)
End If

If barray(i%) > bmax Then
bmax = barray(i%)
imax& = CLng(i%)
End If
Next i%

Exit Sub

' Errors
MiscGetArrayMinMaxByteError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetArrayMinMaxByte"
ierror = True
Exit Sub

End Sub

Function MiscGetArrayMax(n As Integer, sarray() As Integer) As Integer
' Return the max value found in a one dimensional integer array

ierror = False
On Error GoTo MiscGetArrayMaxError

Dim i As Integer, tmax As Integer

tmax% = MAXMAXIMUM3%
For i% = 1 To n%
If sarray%(i%) > tmax% Then
tmax% = sarray%(i%)
End If
Next i%

MiscGetArrayMax% = tmax%
Exit Function

' Errors
MiscGetArrayMaxError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetArrayMax"
ierror = True
Exit Function

End Function

Function MiscGetArrayMin(n As Integer, sarray() As Integer) As Integer
' Return the min value found in a one dimensional integer array

ierror = False
On Error GoTo MiscGetArrayMinError

Dim i As Integer, tmin As Integer

tmin% = MAXMINIMUM3%
For i% = 1 To n%
If sarray%(i%) < tmin% Then
tmin% = sarray%(i%)
End If
Next i%

MiscGetArrayMin% = tmin%
Exit Function

' Errors
MiscGetArrayMinError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetArrayMin"
ierror = True
Exit Function

End Function

Function MiscIsEqualTo(lchan As Integer, iarray() As Integer, ivalue As Integer) As Boolean
' Check if any members of an integer array is equal to the passed value

ierror = False
On Error GoTo MiscIsEqualToError

Dim i As Integer

MiscIsEqualTo = False
For i% = 1 To lchan%
If iarray%(i%) = ivalue% Then MiscIsEqualTo = True
Next i%

Exit Function

' Errors
MiscIsEqualToError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscIsEqualTo"
ierror = True
Exit Function

End Function

Function MiscIsElementDuplicated(sample() As TypeSample) As Integer
' Check for duplicated element

ierror = False
On Error GoTo MiscIsElementDuplicatedError

Dim i As Integer, j As Integer

' Fail if not specified
MiscIsElementDuplicated% = False

' Search sample for match (element only)
For i% = 1 To sample(1).LastElm%
For j% = 1 To sample(1).LastElm%

If i% <> j% And Trim$(UCase$(sample(1).Elsyms$(j%))) = Trim$(UCase$(sample(1).Elsyms$(i%))) Then
MiscIsElementDuplicated% = True
Exit Function
End If

Next j%
Next i%

Exit Function

' Errors
MiscIsElementDuplicatedError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscIsElementDuplicated"
ierror = True
Exit Function

End Function

Function MiscIsElementDuplicatedSubsequent(chan As Integer, num As Integer, elementarray() As String, ipp As Integer) As Boolean
' Check for duplicated element in the element list after the indicated element (ipp% is the index of the duplicated element)

ierror = False
On Error GoTo MiscIsElementDuplicatedSubsequentError

Dim j As Integer

' Fail if not duplicated
MiscIsElementDuplicatedSubsequent = False

' Search sample for match (element only)
ipp% = 0
For j% = chan% + 1 To num%
If Trim$(UCase$(elementarray$(chan%))) = Trim$(UCase$(elementarray$(j%))) Then
MiscIsElementDuplicatedSubsequent = True
ipp% = j%
Exit Function
End If

Next j%

Exit Function

' Errors
MiscIsElementDuplicatedSubsequentError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscIsElementDuplicatedSubsequent"
ierror = True
Exit Function

End Function

Function MiscIsElementDuplicated2(chan As Integer, sample() As TypeSample, ipp As Integer) As Boolean
' Check for duplicated element, x-ray, takeoff and keV in the element list (skip disable quant elements)
' See also IPOS14

ierror = False
On Error GoTo MiscIsElementDuplicated2Error

Dim j As Integer

' Fail if not duplicated
MiscIsElementDuplicated2 = False

' Search sample for match
ipp% = 0
For j% = 1 To sample(1).LastChan%
If chan% <> j% And sample(1).DisableQuantFlag%(j%) = 0 Then
If Trim$(UCase$(sample(1).Elsyms$(chan%))) = Trim$(UCase$(sample(1).Elsyms$(j%))) Then
If Trim$(UCase$(sample(1).Xrsyms$(chan%))) = Trim$(UCase$(sample(1).Xrsyms$(j%))) Then
If sample(1).TakeoffArray!(chan%) = sample(1).TakeoffArray!(j%) Then
If sample(1).KilovoltsArray!(chan%) = sample(1).KilovoltsArray!(j%) Then
MiscIsElementDuplicated2 = True
ipp% = j%
Exit Function
End If
End If
End If
End If
End If

Next j%

Exit Function

' Errors
MiscIsElementDuplicated2Error:
MsgBox Error$, vbOKOnly + vbCritical, "MiscIsElementDuplicated2"
ierror = True
Exit Function

End Function

Function MiscConvertLog10(X As Double) As Double
' Calculate a Base 10 log

ierror = False
On Error GoTo MiscConvertLog10Error

If X# <= 0 Then Exit Function
MiscConvertLog10# = Log(X#) / Log(10#)
Exit Function

' Errors
MiscConvertLog10Error:
MsgBox Error$, vbOKOnly + vbCritical, "MiscConvertLog10"
ierror = True
Exit Function

End Function

Function MiscMin(X As Variant, Y As Variant) As Variant
' Finds the minimum of two values passed

ierror = False
If ierror Then GoTo MiscMinError

If X > Y Then
      MiscMin = Y
 Else
      MiscMin = X
End If

Exit Function

' Errors
MiscMinError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscMin"
ierror = True
Exit Function

End Function

Function MiscMax(X As Variant, Y As Variant) As Variant
' Finds the maximum of two values passed

ierror = False
If ierror Then GoTo MiscMaxError

If X < Y Then
    MiscMax = Y
 Else
    MiscMax = X
End If

Exit Function

' Errors
MiscMaxError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscMax"
ierror = True
Exit Function

End Function

Public Function MiscIsLongArrayInitalized(larray() As Long) As Boolean
' Return True if long integer array is initalized

On Error GoTo MiscIsLongArrayInitalizedError ' raise error if array is not initialzied

Dim temp As Long

MiscIsLongArrayInitalized = False
temp& = UBound(larray&)

' We reach this point only if arr is initalized, i.e. no error occured
If temp& > -1 Then MiscIsLongArrayInitalized = True  ' UBound is greater then -1
Exit Function

' Special error handler (if an error occurs, this function returns False. i.e. array not initialized)
MiscIsLongArrayInitalizedError:
Exit Function

End Function

