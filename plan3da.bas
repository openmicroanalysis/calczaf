Attribute VB_Name = "CodePLAN3Da"
' (c) Copyright 1995-2025 by John J. Donovan
Option Explicit

Sub Plan3dLUBKSB(a() As Double, n As Integer, np As Integer, indx() As Integer, b() As Double)
' Modified From Numerical Recipes

ierror = False
On Error GoTo Plan3dLUBKSBError

Dim sum As Double
Dim i As Integer, j As Integer
Dim ii As Integer, ll As Integer

ii% = 0
For i% = 1 To n%
ll% = indx%(i%)
sum# = b#(ll%)
b#(ll%) = b#(i%)
If ii% <> 0 Then
    For j% = ii% To i% - 1
    sum# = sum# - a#(i%, j%) * b#(j%)
    Next j%
ElseIf sum# <> 0# Then
ii% = i%
End If
b#(i%) = sum#
Next i%

For i% = n% To 1 Step -1
sum# = b#(i%)
If i% < n% Then
    For j% = i% + 1 To n%
    sum# = sum# - a#(i%, j%) * b#(j%)
    Next j%
End If
b#(i%) = sum# / a#(i%, i%)
Next i%

Exit Sub

' Errors
Plan3dLUBKSBError:
MsgBox Error$, vbOKOnly + vbCritical, "Plan3dLUBKSB"
ierror = True
Exit Sub

End Sub

Sub Plan3dLUDCMP(a() As Double, n As Integer, np As Integer, indx() As Integer, d As Double)
' Modified From Numerical Recipes

ierror = False
On Error GoTo Plan3dLUDCMPError

Const TINY1# = 1E-20

Dim i As Integer, j As Integer
Dim k As Integer, imax As Integer
Dim aamax As Double, sum As Double, dum As Double

ReDim vv(1 To n%) As Double

d# = 1#
For i% = 1 To n%
aamax# = 0#
    For j% = 1 To n%
    If Abs(a#(i%, j%)) > aamax# Then aamax# = Abs(a#(i%, j%))
    Next j%

' Check for singular matrix
If aamax# = 0# Then GoTo Plan3dLUDCMPSingularMatrix

vv#(i%) = 1# / aamax#
Next i%

For j% = 1 To n%

If j% > 1 Then
    For i% = 1 To j% - 1
    sum# = a#(i%, j%)

    If i% > 1 Then
        For k% = 1 To i% - 1
        sum# = sum# - a#(i%, k%) * a#(k%, j%)
        Next k%
    a#(i%, j%) = sum#
    End If

    Next i%
End If

aamax# = 0#
    For i% = j% To n%
    sum# = a#(i%, j%)

    If j% > 1 Then
        For k% = 1 To j% - 1
        sum# = sum# - a#(i%, k%) * a#(k%, j%)
        Next k%
    a#(i%, j%) = sum#
    End If

    dum# = vv#(i%) * Abs(sum#)
    If dum# >= aamax# Then
    imax% = i%
    aamax# = dum#
    End If
    Next i%

If j% <> imax% Then
    For k% = 1 To n%
    dum# = a#(imax%, k%)
    a#(imax%, k%) = a#(j%, k%)
    a#(j%, k%) = dum#
    Next k%
d# = -d#
vv#(imax%) = vv#(j%)
End If

indx%(j%) = imax%
If j% <> n% Then
If a#(j%, j%) = 0# Then a#(j%, j%) = TINY1#
dum = 1# / a#(j%, j%)
    For i% = j% + 1 To n%
    a#(i%, j%) = a#(i%, j%) * dum#
    Next i%
End If
Next j%

If a#(n%, n%) = 0# Then a#(n%, n%) = TINY1#

Exit Sub

' Errors
Plan3dLUDCMPError:
MsgBox Error$, vbOKOnly + vbCritical, "Plan3dLUDCMP"
ierror = True
Exit Sub

Plan3dLUDCMPSingularMatrix:
msg$ = "Singular matrix found, points are in a line."
MsgBox msg$, vbOKOnly + vbExclamation, "Plan3dLUDCMP"
ierror = True
Exit Sub

End Sub

Sub Plan3dInvertMatrix(a() As Double, n As Integer, np As Integer, Y() As Double, d As Double)
' This routine calls routines Plan3dLUDCMP and Plan3dLUBKSB. Modified From Numerical Recipes.
'  n = the size of the n x n matrix a with a physical dimension np
'  a is the matrix input (destroyed), y is the inverted output

ierror = False
On Error GoTo Plan3dInvertMatrixError

Dim i As Integer, j As Integer
        
ReDim indx(1 To np%) As Integer
ReDim temp(1 To n%) As Double   ' because VB gives type mismatch otherwise

' Initialize
For i% = 1 To n%
For j% = 1 To n%
Y#(i%, j%) = 0#
Next j%
Y#(i%, i%) = 1#
Next i%

Call Plan3dLUDCMP(a#(), n%, np%, indx%(), d#)
If ierror Then Exit Sub

For j% = 1 To n%
d# = d# * a#(j%, j%)
Next j%

' Check determinant
If Abs(d#) < 0.0000000001 Then GoTo Plan3dInvertMatrixBadDeterminant

For j% = 1 To n%
For i% = 1 To n%
temp#(i%) = Y#(i%, j%)  ' y#(1, j%)
Next i%
Call Plan3dLUBKSB(a#(), n%, np%, indx%(), temp#())
For i% = 1 To n%
Y#(i%, j%) = temp#(i%)
Next i%
Next j%

Exit Sub

' Errors
Plan3dInvertMatrixError:
MsgBox Error$, vbOKOnly + vbCritical, "Plan3dInvertMatrix"
ierror = True
Exit Sub

Plan3dInvertMatrixBadDeterminant:
msg$ = "Bad determinant = " & Str$(d#)
MsgBox msg$, vbOKOnly + vbExclamation, "Plan3dInvertMatrix"
ierror = True
Exit Sub

End Sub


