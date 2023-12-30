Attribute VB_Name = "CodeSpline"
' (c) Copyright 1995-2024 by John J. Donovan
Option Explicit

Sub SplineFit(x() As Single, y() As Single, n As Long, yp1 As Double, ypn As Double, y2() As Double)
' Calculate coefficients for a cubic spline fit (call SplineInterpolate to get interpolated values)
' Modified From Numerical Recipes
' x() and y() = tabulated x and y input array
' n = number of input array elements
' yp1 and ypn are first derivatives of the interpolating function at point 1 and n& (> 10^30 for natural spline)
' y2 is the output array of second derivatives for function SplineInterpolate

ierror = False
On Error GoTo SplineFitError

Dim i As Long

ReDim xdata(1 To n&) As Single, ydata(1 To n&) As Single
ReDim ty2(1 To n&) As Double

' Check that data is in proper order
If x!(1) < x!(n&) Then
For i& = 1 To n&
xdata!(i&) = x!(i&)
ydata!(i&) = y!(i&)
Next i&

Else
For i& = 1 To n&
xdata!(i&) = x!(n& - (i& - 1))
ydata!(i&) = y!(n& - (i& - 1))
Next i&
End If

' Call actual fit function (leave 2nd derivatives in fit order)
Call SplineFit2(xdata!(), ydata!(), n&, yp1#, ypn#, ty2#())
If ierror Then Exit Sub

' Store derivatives in fit order
For i& = 1 To n&
y2#(i&) = ty2#(i&)
Next i&

Exit Sub

' Errors
SplineFitError:
MsgBox Error$, vbOKOnly + vbCritical, "SplineFit"
ierror = True
Exit Sub

End Sub

Sub SplineFit2(x() As Single, y() As Single, n As Long, yp1 As Double, ypn As Double, y2() As Double)
' Calculate coefficients for a cubic spline fit (call SplineInterpolate to get interpolated values)
' Modified From Numerical Recipes
' x() and y() = tabulated x and y input array
' n = number of input array elements
' yp1 and ypn are first derivatives of the interpolating function at point 1 and n& (> 10^30 for natural spline)
' y2 is the output array of second derivatives for function SplineInterpolate

ierror = False
On Error GoTo SplineFit2Error

Dim k As Long, i As Long
Dim sig As Double, dum1 As Double, dum2 As Double
Dim p As Double, qn As Double, un As Double

ReDim u(1 To n&) As Double

If yp1# > 9.9E+29 Then
  y2#(1) = 0!
  u#(1) = 0!
Else
  y2#(1) = -0.5
  u#(1) = (3! / (x!(2) - x!(1))) * ((y!(2) - y!(1)) / (x!(2) - x!(1)) - yp1#)
End If
For i& = 2 To n& - 1
  sig# = (x!(i&) - x!(i& - 1)) / (x!(i& + 1) - x!(i& - 1))
  p# = sig# * y2#(i& - 1) + 2!
  y2#(i) = (sig# - 1!) / p#
  dum1# = (y!(i& + 1) - y!(i&)) / (x!(i& + 1) - x!(i&))
  dum2# = (y!(i&) - y!(i& - 1)) / (x!(i&) - x!(i& - 1))
  u#(i&) = (6! * (dum1# - dum2#) / (x!(i& + 1) - x!(i& - 1)) - sig# * u#(i& - 1)) / p#
Next i&

If ypn# > 9.9E+29 Then
  qn# = 0!
  un# = 0!
Else
  qn# = 0.5
  un# = (3! / (x!(n&) - x!(n& - 1))) * (ypn# - (y!(n&) - y!(n& - 1)) / (x!(n&) - x!(n& - 1)))
End If
y2#(n&) = (un# - qn# * u#(n& - 1)) / (qn# * y2#(n& - 1) + 1!)
For k& = n& - 1 To 1 Step -1
  y2#(k&) = y2#(k&) * y2#(k& + 1) + u#(k&)
Next k&

Exit Sub

' Errors
SplineFit2Error:
MsgBox Error$, vbOKOnly + vbCritical, "SplineFit2"
ierror = True
Exit Sub

End Sub

Sub SplineInterpolate(xa() As Single, ya() As Single, y2a() As Double, n As Long, x As Single, y As Single)
' Calculate interpolated values based on passed coefficients from SplineFit
' Modified From Numerical Recipes
' xa() and ya() = tabulated x and y input array (original data)
' y2a() = second derivative values from SplineFit
' n = number of input array elements
' x = value to interpolate y
' y = returned interpolated y value

ierror = False
On Error GoTo SplineInterpolateError

Dim i As Long

ReDim xdata(1 To n&) As Single, ydata(1 To n&) As Single

' Check that data is in proper order
If xa!(1) < xa!(n&) Then
For i& = 1 To n&
xdata!(i&) = xa!(i&)
ydata!(i&) = ya!(i&)
Next i&

Else
For i& = 1 To n&
xdata!(i&) = xa!(n& - (i& - 1))
ydata!(i&) = ya!(n& - (i& - 1))
Next i&
End If

' Call actual routine (2nd derivatives are always in proper order)
Call SplineInterpolate2(xdata!(), ydata!(), y2a#(), n&, x!, y!)
If ierror Then Exit Sub

' If x value is below or above, then set y to last value
If x! < xdata!(1) Then y! = ydata(1)
If x! > xdata!(n&) Then y! = ydata(n&)

Exit Sub

' Errors
SplineInterpolateError:
MsgBox Error$, vbOKOnly + vbCritical, "SplineInterpolate"
ierror = True
Exit Sub

End Sub

Sub SplineInterpolate2(xa() As Single, ya() As Single, y2a() As Double, n As Long, x As Single, y As Single)
' Calculate interpolated values based on passed coefficients from SplineFit
' Modified From Numerical Recipes
' xa() and ya() = tabulated x and y input array (original data)
' y2a() = second derivative values from SplineFit
' n = number of input array elements
' x = value to interpolate y
' y = returned interpolated y value

ierror = False
On Error GoTo SplineInterpolate2Error

Dim k As Long
Dim klo As Long, khi As Long
Dim h As Double, a As Double, b As Double

klo& = 1
khi& = n&
While khi& - klo& > 1
  k& = (khi& + klo&) / 2
  If xa!(k&) > x! Then
    khi = k&
  Else
    klo& = k&
  End If
Wend
h# = xa!(khi&) - xa!(klo&)
If h# = 0! Then GoTo SplineInterpolate2BadInput

a# = (xa!(khi&) - x!) / h#
b# = (x! - xa!(klo&)) / h#
y! = a# * ya!(klo&) + b# * ya!(khi&)
y! = y! + ((a# ^ 3 - a#) * y2a#(klo&) + (b# ^ 3 - b#) * y2a#(khi&)) * (h# ^ 2) / 6#

Exit Sub

' Errors
SplineInterpolate2Error:
MsgBox Error$, vbOKOnly + vbCritical, "SplineInterpolate2"
ierror = True
Exit Sub

SplineInterpolate2BadInput:
msg$ = "Bad input value, x dat first and last values are equal"
MsgBox msg$, vbOKOnly + vbExclamation, "SplineInterpolate2"
ierror = True
Exit Sub

End Sub


