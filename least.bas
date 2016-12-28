Attribute VB_Name = "CodeLEAST"
' (c) Copyright 1995-2017 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub LeastMathFit(kmax As Integer, nmax As Integer, txdata() As Single, tydata() As Single, acoeff() As Single)
' This routine does a least square polynomial fit of degree kmax-1 to the supplied data x and y.
' Nmax is the number of pairs of points (x,y).  The polynomial is of the form:
' acoeff(1) + acoeff(2)*x + ... +acoeff(kmax)*x**(kmax-1)

' This routine was originally written by Tom Dey (18-OCT-78).
' Modified by John Donovan (4-JAN-96).
' If nmax <= 0  ierror = true
' If nmax  = 1  acoeff(1) = ydata(1), acoeff(2) = 0.0, acoeff(3) = 0.0, etc.
' If nmax  > 1  and the slope is infinite (no x variation) then ierror = true; that is the matrix is singular

ierror = False
On Error GoTo LeastMathFitError

Dim i As Integer, j As Integer, k As Integer
Dim xxxold As Double, yxold As Double, xi As Double

ReDim yx(1 To MAXCOEFF9%) As Double, xxx(1 To 2 * MAXCOEFF9%) As Double
ReDim xx(1 To MAXCOEFF9%, 1 To MAXCOEFF9%) As Double
ReDim a(1 To MAXCOEFF9%) As Double

' Initialize
For i% = 1 To MAXCOEFF9%
a#(i%) = 0#
Next i%

For i% = 1 To kmax%
acoeff!(i%) = a#(i%)
Next i%

' Check for problems
If kmax% > MAXCOEFF9% Then kmax% = MAXCOEFF9%
If nmax% < kmax% Then kmax% = nmax%

' No points
If nmax% <= 0 Then GoTo LeastMathFitNoPoints

' Only one point, no error, just return constant intercept
If nmax% = 1 Then
acoeff!(1) = tydata!(1)
Exit Sub
End If

For i% = 1 To kmax%
yx#(i%) = 0#
xxx#(2 * i%) = 0#
xxx#(2 * i% - 1) = 0#
For j% = 1 To kmax%
xx#(i%, j%) = 0#
Next j%
Next i%

For i% = 1 To nmax%
yxold# = tydata!(i%)
xxxold# = 1#
yx#(1) = yx#(1) + yxold#
xxx#(2) = xxx#(2) + xxxold#
xi# = txdata!(i%)

For k% = 2 To kmax%
yxold# = yxold# * xi#
yx#(k%) = yx#(k%) + yxold#
xxxold# = xxxold# * xi#
xxx#(2 * k% - 1) = xxx#(2 * k% - 1) + xxxold#
xxxold# = xxxold# * xi#
xxx#(2 * k%) = xxx#(2 * k%) + xxxold#
Next k%
Next i%

For i% = 1 To kmax%
For j% = 1 To kmax%
xx#(i%, j%) = xxx#(i% + j%)
Next j%
Next i%

' Solve equation
Call LeastGauss(kmax%, a#(), yx#(), xx#())
If ierror Then Exit Sub

' Load returned coefficients
For i% = 1 To kmax%
acoeff!(i%) = a#(i%)
Next i%
Exit Sub

' Errors
LeastMathFitError:
MsgBox Error$, vbOKOnly + vbCritical, "LeastMathFit"
ierror = True
Exit Sub

LeastMathFitNoPoints:
msg$ = "No points to fit"
MsgBox msg$, vbOKOnly + vbExclamation, "LeastMathFit"
ierror = True
Exit Sub

End Sub

Sub LeastSquares(norder As Integer, npts As Integer, txdata() As Single, tydata() As Single, acoeff() As Single)
' Calculates a least squares fit

ierror = False
On Error GoTo LeastSquaresError

If npts% < 1 Then GoTo LeastSquaresNoPoints
If norder% < 0 Or norder% > MAXCOEFF9% Then GoTo LeastSquaresBadOrder
If norder% > npts% - 1 Then norder% = npts% - 1

' If only one point loaded, just return single y value
If npts% < 2 Then
acoeff!(1) = tydata!(1)
acoeff!(2) = 0#
acoeff!(3) = 0#
Exit Sub
End If

' Get fit using math routine
Call LeastMathFit(norder% + 1, npts%, txdata!(), tydata!(), acoeff!())
If ierror Then Exit Sub

Exit Sub

' Errors
LeastSquaresError:
MsgBox Error$, vbOKOnly + vbCritical, "LeastSquares"
ierror = True
Exit Sub

LeastSquaresNoPoints:
msg$ = "No data points to fit"
MsgBox msg$, vbOKOnly + vbExclamation, "LeastSquares"
ierror = True
Exit Sub

LeastSquaresBadOrder:
msg$ = "Polynomial fit order is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "LeastSquares"
ierror = True
Exit Sub

End Sub

Sub LeastMathFit2(kmax As Integer, nmax As Integer, txdata() As Double, tydata() As Double, acoeff() As Single)
' This routine does a least square polynomial fit of degree kmax-1 to the supplied data x and y.
' Nmax is the number of pairs of points (x,y).  The polynomial is of the form:
' acoeff(1) + acoeff(2)*x + ... +acoeff(kmax)*x**(kmax-1)

' (double precision version)

' This routine was originally written by Tom Dey (18-OCT-78).
' Modified by John Donovan (4-JAN-96).
' If nmax <= 0  ierror = true
' If nmax  = 1  acoeff(1) = ydata(1), acoeff(2) = 0.0, acoeff(3) = 0.0, etc.
' If nmax  > 1  and the slope is infinite (no x variation) then ierror = true; that is the matrix is singular

ierror = False
On Error GoTo LeastMathFit2Error

Dim i As Integer, j As Integer, k As Integer
Dim xxxold As Double, yxold As Double, xi As Double

ReDim yx(1 To MAXCOEFF9%) As Double, xxx(1 To 2 * MAXCOEFF9%) As Double
ReDim xx(1 To MAXCOEFF9%, 1 To MAXCOEFF9%) As Double
ReDim a(1 To MAXCOEFF9%) As Double

' Initialize
For i% = 1 To MAXCOEFF9%
a#(i%) = 0#
Next i%

For i% = 1 To kmax%
acoeff!(i%) = a#(i%)
Next i%

' Check for problems
If kmax% > MAXCOEFF9% Then kmax% = MAXCOEFF9%
If nmax% < kmax% Then kmax% = nmax%

' No points
If nmax% <= 0 Then GoTo LeastMathFit2NoPoints

' Only one point, no error, just return constant
If nmax% = 1 Then
acoeff!(1) = tydata#(1)
Exit Sub
End If

For i% = 1 To kmax%
yx#(i%) = 0#
xxx#(2 * i%) = 0#
xxx#(2 * i% - 1) = 0#
For j% = 1 To kmax%
xx#(i%, j%) = 0#
Next j%
Next i%

For i% = 1 To nmax%
yxold# = tydata#(i%)
xxxold# = 1#
yx#(1) = yx#(1) + yxold#
xxx#(2) = xxx#(2) + xxxold#
xi# = txdata#(i%)

For k% = 2 To kmax%
yxold# = yxold# * xi#
yx#(k%) = yx#(k%) + yxold#
xxxold# = xxxold# * xi#
xxx#(2 * k% - 1) = xxx#(2 * k% - 1) + xxxold#
xxxold# = xxxold# * xi#
xxx#(2 * k%) = xxx#(2 * k%) + xxxold#
Next k%
Next i%

For i% = 1 To kmax%
For j% = 1 To kmax%
xx#(i%, j%) = xxx#(i% + j%)
Next j%
Next i%

' Solve equation
Call LeastGauss(kmax%, a#(), yx#(), xx#())
If ierror Then Exit Sub

' Load returned coefficients
For i% = 1 To kmax%
acoeff!(i%) = a#(i%)
Next i%
Exit Sub

' Errors
LeastMathFit2Error:
MsgBox Error$, vbOKOnly + vbCritical, "LeastMathFit2"
ierror = True
Exit Sub

LeastMathFit2NoPoints:
msg$ = "No points to fit"
MsgBox msg$, vbOKOnly + vbExclamation, "LeastMathFit2"
ierror = True
Exit Sub

End Sub

Sub LeastExponential(txdata() As Single, tydata() As Single, texp As Single, bcoeff() As Double)
' Calculates the coefficients for a 2 point exponential fit of the form:
'   y = (c* e^(-ax))/x^n,   where n is user specified

ierror = False
On Error GoTo LeastExponentialError

Dim c As Double, a As Double, n As Double
Dim temp1 As Double, temp2 As Double
Dim temp As Double

' Init "bcoeff()"
bcoeff#(1) = 0#
bcoeff#(2) = 0#
bcoeff#(3) = 0#

' Check for valid input data
If txdata!(1) <= 0# Then Exit Sub
If txdata!(2) <= 0# Then Exit Sub
If tydata!(1) <= 0# Then Exit Sub
If tydata!(2) <= 0# Then Exit Sub

' Load base
If texp! < -8# Or texp! > 8# Then GoTo LeastExponentialBadExponent
n# = texp!

' Calculate c
temp1# = (txdata!(1) * Log(tydata!(2)) - txdata!(2) * Log(tydata!(1)))
temp1# = temp1# + n# * txdata!(1) * Log(txdata!(2)) - n# * txdata!(2) * Log(txdata!(1))
temp2# = txdata!(1) - txdata!(2)
If temp2# = 0# Then Exit Sub
temp# = temp1# / temp2#
c# = Exp(temp#)

' Calculate a
If c# < 0# Then Exit Sub
temp1# = (Log(tydata!(1)) + Log(tydata!(2)) + n# * (Log(txdata!(2)) + Log(txdata!(1))) - 2 * Log(c#))
temp2# = txdata!(1) + txdata!(2)
If temp2# = 0# Then Exit Sub
a# = -(temp1# / temp2#)

' Load coefficients (1=c, 2=a, 3=n)
bcoeff#(1) = c#
bcoeff#(2) = a#
bcoeff#(3) = n#
Exit Sub

' Errors
LeastExponentialError:
MsgBox Error$, vbOKOnly + vbCritical, "LeastExponential"
ierror = True
Exit Sub

LeastExponentialBadExponent:
msg$ = "Exponent must be greater than -8 and less than 8"
MsgBox msg$, vbOKOnly + vbExclamation, "LeastExponential"
ierror = True
Exit Sub

End Sub

Sub LeastDeviation(mode As Integer, avgdev As Single, npts As Integer, txdata() As Single, tydata() As Single, acoeff() As Single)
' Calculate relative average deviation for fit (relative percent)
' mode% = 1 use quadratic expression
' mode% = 2 use gaussian expression
' mode% = 3 use logarithmic expression
' mode% = 4 use logarithmic2 expression
' mode% = 5 use exponential expression (y data is in log units)

ierror = False
On Error GoTo LeastDeviationError

Dim i As Integer
Dim temp As Single, sum As Single

' No points
If npts% = 0 Then
avgdev! = 0#
Exit Sub
End If

' One point
If npts% = 1 Then
avgdev! = 0#
Exit Sub
End If

' Calculate relative standard deviation from fit
avgdev! = 0#
sum! = 0#
For i% = 1 To npts%

' Sum values
If mode% = 5 Then
sum! = sum! + NATURALE# ^ tydata!(i%)                           ' if exponential, convert log intensities to base 10
Else
sum! = sum! + tydata!(i%)
End If

' Calculate absolute deviation from fit
If mode% = 1 Then       ' quadratic
temp! = acoeff!(1) + acoeff!(2) * txdata!(i%) + acoeff!(3) * txdata!(i%) ^ 2

ElseIf mode% = 2 Then   ' gaussian
temp! = acoeff!(1) + acoeff!(2) * txdata!(i%) + acoeff!(3) * txdata!(i%) ^ 2
If temp! > MAXLOGEXPS! Then temp! = MAXLOGEXPS!
temp! = CSng(NATURALE# ^ temp!)

ElseIf mode% = 3 Then   ' logarithmic
temp! = acoeff!(1) + acoeff!(2) * Log(txdata!(i%))

ElseIf mode% = 4 Then   ' logarithmic2
temp! = acoeff!(1) + acoeff!(2) * Log(txdata!(i%)) + acoeff!(3) * Log(txdata!(i%)) ^ 2

ElseIf mode% = 5 Then   ' exponential
temp! = acoeff!(1) + acoeff!(2) * txdata!(i%) + acoeff!(3) * txdata!(i%) ^ 2
End If

' Calculate difference squared
If mode% = 5 Then
temp! = (NATURALE# ^ tydata!(i%) - NATURALE# ^ temp!) ^ 2       ' if exponential, convert log intensities to base 10

Else
temp! = (tydata!(i%) - temp!) ^ 2
End If

' Sum the relative deviations
avgdev! = avgdev! + temp!
Next i%

' Calculate average
sum! = sum! / npts%

' Calculate relative deviation percent
avgdev! = Sqr(avgdev! / (npts% - 1))

' Calculate standard deviation percent
If sum! > 0# Then
avgdev! = (avgdev! / sum!) * 100#
End If

Exit Sub

' Errors
LeastDeviationError:
MsgBox Error$, vbOKOnly + vbCritical, "LeastDeviation"
ierror = True
Exit Sub

End Sub

Sub LeastRegressGaussian(numdata As Integer, numfits As Integer, xdata() As Double, ydata() As Double, acoeff() As Double, area As Single, x0 As Single, sd As Single)
' Run the gaussian fit calculation
' numdata%                              number of data points
' numfits%                              number of coefficients (fit)
' xdata#(1 to MAXDATA%)                 x variable (peak position)
' ydata#(1 to MAXDATA%)                 y variable (count intensity)
' acoeff#(1 to MAXCOEFF%)               fit coefficients
' From Don Snyder (UCB)
'
' Calculate area under curve:
' area = Sqr(-pi/a) * e^(c-(b^2/4a)) ; where pi = 3.14159 and e = 2.718281828
'
' Calculate peak
' x0 = -1/2 * b/a
'
' Standard deviation (returned in percent deviation)
' sd = Sqr(-1/(2a))

ierror = False
On Error GoTo LeastRegressGaussianError

Const MAXDATA% = 2000    ' maximum number of data points
Const MAXFITS% = 3       ' maximum number of fit coefficients

Dim i As Integer
Dim temp As Double, determinate As Double
Dim chisq As Double, avsumsq As Double

ReDim xd(1 To MAXDATA%, 1 To MAXFITS) As Double ' design matrix
ReDim yd(1 To MAXDATA%) As Double ' log of intensities (for gaussian regression)

ReDim sig(1 To MAXDATA%) As Double
ReDim predicated(1 To MAXDATA%) As Double
ReDim residual(1 To MAXDATA%) As Double
ReDim fitparam(1 To MAXFITS%) As Double
ReDim cvm(1 To MAXFITS%, 1 To MAXFITS%) As Double
ReDim se(1 To MAXFITS%) As Double

If numdata% > MAXDATA% Then GoTo LeastRegressGaussianTooManyData
If numfits% > MAXFITS% Then GoTo LeastRegressGaussianTooManyFits

' Load x data
For i% = 1 To numdata%
xd#(i%, 1) = 1#
xd#(i%, 2) = xdata#(i%)
xd#(i%, 3) = xdata#(i%) * xdata#(i%)
yd#(i%) = Log(ydata#(i%))
sig#(i%) = 1#    ' weighting of data for fit (use 1.0 for no weighting)
Next i%

' Determine the gaussian fit coefficients: ln(y) = a + bx + cx^2
Call RegressLINREG(xd#(), yd#(), numdata%, numfits%, MAXDATA%, MAXFITS%, sig#(), determinate#, chisq#, avsumsq#, predicated#(), residual#(), fitparam#(), cvm#(), se#())
If ierror Then Exit Sub

' Return fit coefficients
For i% = 1 To numfits%
acoeff#(i%) = fitparam#(i%)
Next i%

' Display fit
If DebugMode And VerboseMode Then
Call IOWriteLog("LeastRegressGaussian- Fit parameters " & MiscAutoFormatD$(acoeff#(1)) & MiscAutoFormatD$(acoeff#(2)) & MiscAutoFormatD$(acoeff#(3)))
End If

' Calculate area, peak and standard deviation
If acoeff#(3) <> 0# Then temp# = -PI! / acoeff#(3)
area! = 0#
If temp# >= 0# And acoeff#(3) <> 0# Then
area! = Sqr(temp#) * NATURALE# ^ (acoeff#(1) - (acoeff#(2) ^ 2 / (4 * acoeff#(3))))
End If

' Calculate peak
x0! = 0#
If acoeff#(3) <> 0# Then
x0! = -1 / 2 * acoeff#(2) / acoeff#(3)
End If

' Standard deviation
sd! = 0#
If acoeff#(3) <> 0# Then
temp# = -1 / (2 * acoeff#(3))
If temp# >= 0# Then
sd! = Sqr(temp#)
End If
sd! = sd! * 100#    ' convert to percent
End If

Exit Sub

' Errors
LeastRegressGaussianError:
MsgBox Error$, vbOKOnly + vbCritical, "LeastRegressGaussian"
ierror = True
Exit Sub

LeastRegressGaussianTooManyData:
msg$ = "Too many data points"
MsgBox msg$, vbOKOnly + vbExclamation, "LeastRegressGaussian"
ierror = True
Exit Sub

LeastRegressGaussianTooManyFits:
msg$ = "Too many fit coefficients"
MsgBox msg$, vbOKOnly + vbExclamation, "LeastRegressGaussian"
ierror = True
Exit Sub

End Sub

Sub LeastSpline(npts As Integer, axdata() As Single, aydata() As Single, ycoeff() As Double)
' Fit the passed data to cubic spline

ierror = False
On Error GoTo LeastSplineError

Dim yp1 As Double, ypn As Double

' Dimension ycoeff
ReDim ycoeff(1 To npts%) As Double

' Fit data to cubic spline function
yp1# = 10# ^ 30
ypn# = 10# ^ 30
Call SplineFit(axdata!(), aydata!(), CLng(npts%), yp1#, ypn#, ycoeff#())
If ierror Then Exit Sub

Exit Sub

' Errors
LeastSplineError:
MsgBox Error$, vbOK + vbCritical, "LeastSpline"
ierror = True
Exit Sub

End Sub

Sub LeastSplineInterpolate(npts%, axdata() As Single, aydata() As Single, ycoeff() As Double, X As Single, Y As Single)
' Return the y-value for the passed x value

ierror = False
On Error GoTo LeastSplineInterpolateError

' Return the interpolated value
Call SplineInterpolate(axdata!(), aydata!(), ycoeff#(), CLng(npts%), X!, Y!)
If ierror Then Exit Sub

Exit Sub

' Errors
LeastSplineInterpolateError:
MsgBox Error$, vbOK + vbCritical, "LeastSplineInterpolate"
ierror = True
Exit Sub

End Sub

Function LeastSmoothSavitzkyGolayGetWindow(n As Integer) As Integer
' Get the optimum size of the smoothing window

ierror = False
On Error GoTo LeastSmoothSavitzkyGolayGetWindowError

Dim k As Integer

' Calculate window size
k% = n% / 10

' Check that k% is even
If k% Mod 2 <> 0 Then k% = k% + 1

' Check size
If k% < 2 Then k% = 2
If k% > 12 Then k% = 12

LeastSmoothSavitzkyGolayGetWindow% = k%
Exit Function

' Errors
LeastSmoothSavitzkyGolayGetWindowError:
MsgBox Error$, vbOK + vbCritical, "LeastSmoothSavitzkyGolayGetWindow"
ierror = True
Exit Function

End Function

Sub LeastMathLogarithmic(nmax As Integer, txdata() As Single, tydata() As Single, acoeff() As Single)
' This routine does a least square logarithmic fit Y = B*Log(X) + A to the supplied data x and y.
' nmax%                              number of data points
' txdata#(1 to MAXDATA%)             x variable
' tydata#(1 to MAXDATA%)             y variable
' acoeff#(1 to MAXCOEFF%)            fit coefficients
' If nmax <= 0  ierror = true
' If nmax  = 1  acoeff(1) = ydata(1), acoeff(2) = 0.0, acoeff(3) = 0.0, etc.
' If nmax  > 1  and the slope is infinite (no x variation) then ierror = true; that is the matrix is singular

ierror = False
On Error GoTo LeastMathLogarithmicError

Dim i As Integer
Dim chisq As Double, determinate As Double
Dim avsumsq As Double

Const MAXDATA% = 2000    ' maximum number of data points
Const MAXFITS% = 2       ' maximum number of fit coefficients

ReDim xd(1 To MAXDATA%, 1 To MAXFITS%) As Double    ' design matrix
ReDim yd(1 To MAXDATA%) As Double                   ' log of intensities (for logarithmic regression)

ReDim sig(1 To MAXDATA%) As Double
ReDim predicated(1 To MAXDATA%) As Double
ReDim residual(1 To MAXDATA%) As Double
ReDim fitparam(1 To MAXFITS%) As Double
ReDim cvm(1 To MAXFITS%, 1 To MAXFITS%) As Double
ReDim se(1 To MAXFITS%) As Double

' Initialize
For i% = 1 To MAXCOEFF%
acoeff!(i%) = 0#
Next i%

' No points
If nmax% <= 0 Then GoTo LeastMathLogarithmicNoPoints

' Too many points
If nmax% > MAXDATA% Then GoTo LeastMathLogarithmicTooManyData

' Only one point, no error, just return constant intercept
If nmax% = 1 Then
acoeff!(1) = tydata!(1)
Exit Sub
End If

' Load x data
For i% = 1 To nmax%
xd#(i%, 1) = 1#
xd#(i%, 2) = Log(txdata!(i%))
yd#(i%) = tydata!(i%)
sig#(i%) = 1#    ' weighting of data for fit (use 1.0 for no weighting)
Next i%

' Determine the fit coefficients: y = b * Log(x) + a
Call RegressLINREG(xd#(), yd#(), nmax%, MAXFITS%, MAXDATA%, MAXFITS%, sig#(), determinate#, chisq#, avsumsq#, predicated#(), residual#(), fitparam#(), cvm#(), se#())
If ierror Then Exit Sub

' Return fit coefficients
For i% = 1 To MAXFITS%
acoeff!(i%) = CSng(fitparam#(i%))
Next i%

' Display fit
If DebugMode And VerboseMode Then
Call IOWriteLog("LeastMathLogarithmic- Fit parameters " & MiscAutoFormat$(acoeff!(1)) & MiscAutoFormat$(acoeff!(2)))
End If

Exit Sub

' Errors
LeastMathLogarithmicError:
MsgBox Error$, vbOKOnly + vbCritical, "LeastMathLogarithmic"
ierror = True
Exit Sub

LeastMathLogarithmicNoPoints:
msg$ = "No points to fit"
MsgBox msg$, vbOKOnly + vbExclamation, "LeastMathLogarithmic"
ierror = True
Exit Sub

LeastMathLogarithmicTooManyData:
msg$ = "Too many data points"
MsgBox msg$, vbOKOnly + vbExclamation, "LeastMathLogarithmic"
ierror = True
Exit Sub

End Sub

Sub LeastMathLogarithmic2(nmax As Integer, txdata() As Single, tydata() As Single, acoeff() As Single)
' This routine does a logarithmic fit Y = C*Log(X)^2 + B*Log(X) + A to the supplied data x and y.
' nmax%                              number of data points
' txdata#(1 to MAXDATA%)             x variable
' tydata#(1 to MAXDATA%)             y variable
' acoeff#(1 to MAXCOEFF%)            fit coefficients
' If nmax <= 0  ierror = true
' If nmax  = 1  acoeff(1) = ydata(1), acoeff(2) = 0.0, acoeff(3) = 0.0, etc.
' If nmax  > 1  and the slope is infinite (no x variation) then ierror = true; that is the matrix is singular

ierror = False
On Error GoTo LeastMathLogarithmic2Error

Dim i As Integer
Dim chisq As Double, determinate As Double
Dim avsumsq As Double

Const MAXDATA% = 2000    ' maximum number of data points
Const MAXFITS% = 3       ' maximum number of fit coefficients

ReDim xd(1 To MAXDATA%, 1 To MAXFITS%) As Double    ' design matrix
ReDim yd(1 To MAXDATA%) As Double                   ' log of intensities (for logarithmic regression)

ReDim sig(1 To MAXDATA%) As Double
ReDim predicated(1 To MAXDATA%) As Double
ReDim residual(1 To MAXDATA%) As Double
ReDim fitparam(1 To MAXFITS%) As Double
ReDim cvm(1 To MAXFITS%, 1 To MAXFITS%) As Double
ReDim se(1 To MAXFITS%) As Double

' Initialize
For i% = 1 To MAXCOEFF%
acoeff!(i%) = 0#
Next i%

' No points
If nmax% <= 0 Then GoTo LeastMathLogarithmic2NoPoints

' Too many points
If nmax% > MAXDATA% Then GoTo LeastMathLogarithmic2TooManyData

' Only one point, no error, just return constant intercept
If nmax% = 1 Then
acoeff!(1) = tydata!(1)
Exit Sub
End If

' Load x data
For i% = 1 To nmax%
xd#(i%, 1) = 1#
xd#(i%, 2) = Log(txdata!(i%))
xd#(i%, 3) = Log(txdata!(i%)) ^ 2
yd#(i%) = tydata!(i%)
sig#(i%) = 1#    ' weighting of data for fit (use 1.0 for no weighting)
Next i%

' Determine the fit coefficients: y = b * Log(x) + a
Call RegressLINREG(xd#(), yd#(), nmax%, MAXFITS%, MAXDATA%, MAXFITS%, sig#(), determinate#, chisq#, avsumsq#, predicated#(), residual#(), fitparam#(), cvm#(), se#())
If ierror Then Exit Sub

' Return fit coefficients
For i% = 1 To MAXFITS%
acoeff!(i%) = CSng(fitparam#(i%))
Next i%

' Display fit
If DebugMode And VerboseMode Then
Call IOWriteLog("LeastMathLogarithmic2- Fit parameters " & MiscAutoFormat$(acoeff!(1)) & MiscAutoFormat$(acoeff!(2)) & MiscAutoFormat$(acoeff!(3)))
End If

Exit Sub

' Errors
LeastMathLogarithmic2Error:
MsgBox Error$, vbOKOnly + vbCritical, "LeastMathLogarithmic2"
ierror = True
Exit Sub

LeastMathLogarithmic2NoPoints:
msg$ = "No points to fit"
MsgBox msg$, vbOKOnly + vbExclamation, "LeastMathLogarithmic2"
ierror = True
Exit Sub

LeastMathLogarithmic2TooManyData:
msg$ = "Too many data points"
MsgBox msg$, vbOKOnly + vbExclamation, "LeastMathLogarithmic2"
ierror = True
Exit Sub

End Sub

Sub LeastMathNonLinear(nmax As Integer, txdata() As Single, tydata() As Single, acoeff() As Single)
' This routine does a fit to various non-linear expressions for 4 regression coefficients
' nmax%                              number of data points
' txdata#(1 to MAXDATA%)             x variable
' tydata#(1 to MAXDATA%)             y variable
' acoeff#(1 to MAXCOEFF4%)           fit coefficients
' If nmax <= 0  ierror = true
' If nmax  = 1  acoeff(1) = ydata(1), acoeff(2) = 0.0, acoeff(3) = 0.0, acoeff(4) = 0.0
' If nmax  > 1  and the slope is infinite (no x variation) then ierror = true; that is the matrix is singular

ierror = False
On Error GoTo LeastMathNonLinearError

Dim i As Integer
Dim chisq As Double, determinate As Double
Dim avsumsq As Double

Const MAXDATA% = 2000    ' maximum number of data points
Const MAXFITS% = 4       ' maximum number of fit coefficients

ReDim xd(1 To MAXDATA%, 1 To MAXFITS%) As Double    ' design matrix
ReDim yd(1 To MAXDATA%) As Double                   ' log of intensities (for logarithmic regression)

ReDim sig(1 To MAXDATA%) As Double
ReDim predicated(1 To MAXDATA%) As Double
ReDim residual(1 To MAXDATA%) As Double
ReDim fitparam(1 To MAXFITS%) As Double
ReDim cvm(1 To MAXFITS%, 1 To MAXFITS%) As Double
ReDim se(1 To MAXFITS%) As Double

' Initialize
For i% = 1 To MAXCOEFF4%
acoeff!(i%) = 0#
Next i%

' No points
If nmax% <= 0 Then GoTo LeastMathNonLinearNoPoints

' Too many points
If nmax% > MAXDATA% Then GoTo LeastMathNonLinearTooManyData

' Only one point, no error, just return constant intercept
If nmax% = 1 Then
acoeff!(1) = tydata!(1)
Exit Sub
End If

' Load x data
For i% = 1 To nmax%
xd#(i%, 1) = 1#
xd#(i%, 2) = txdata!(i%)
xd#(i%, 3) = txdata!(i%) ^ 2
xd#(i%, 4) = Exp(txdata!(i%))
yd#(i%) = tydata!(i%)
sig#(i%) = 1#    ' weighting of data for fit (use 1.0 for no weighting)
Next i%

' Determine the gaussian fit coefficients: y = b * Log(x) + a
Call RegressLINREG(xd#(), yd#(), nmax%, MAXFITS%, MAXDATA%, MAXFITS%, sig#(), determinate#, chisq#, avsumsq#, predicated#(), residual#(), fitparam#(), cvm#(), se#())
If ierror Then Exit Sub

' Return fit coefficients
For i% = 1 To MAXFITS%
acoeff!(i%) = CSng(fitparam#(i%))
Next i%

' Display fit
If DebugMode And VerboseMode Then
Call IOWriteLog("LeastMathNonLinear- Fit parameters " & MiscAutoFormat$(acoeff!(1)) & MiscAutoFormat$(acoeff!(2)) & MiscAutoFormat$(acoeff!(3)) & MiscAutoFormat$(acoeff!(4)))
End If

Exit Sub

' Errors
LeastMathNonLinearError:
MsgBox Error$, vbOKOnly + vbCritical, "LeastMathNonLinear"
ierror = True
Exit Sub

LeastMathNonLinearNoPoints:
msg$ = "No points to fit"
MsgBox msg$, vbOKOnly + vbExclamation, "LeastMathNonLinear"
ierror = True
Exit Sub

LeastMathNonLinearTooManyData:
msg$ = "Too many data points"
MsgBox msg$, vbOKOnly + vbExclamation, "LeastMathNonLinear"
ierror = True
Exit Sub

End Sub

Sub LeastMathNonLinearDeviation(avgdev As Single, npts As Integer, txdata() As Single, tydata() As Single, acoeff() As Single)
' Calculate relative average deviation for fit (relative percent) for a non linear expression

ierror = False
On Error GoTo LeastMathNonLinearDeviationError

Dim i As Integer
Dim temp As Single, sum As Single

' No points
If npts% = 0 Then
avgdev! = 0#
Exit Sub
End If

' One point
If npts% = 1 Then
avgdev! = 0#
Exit Sub
End If

' Calculate relative standard deviation from fit
avgdev! = 0#
sum! = 0#
For i% = 1 To npts%

' Sum values
sum! = sum! + tydata!(i%)

' Calculate absolute deviation from fit
temp! = acoeff!(1) + acoeff!(2) * txdata!(i%) + acoeff!(3) * txdata!(i%) ^ 2 + acoeff!(4) * Exp(txdata!(i%))

' Calculate difference squared
temp! = (tydata!(i%) - temp!) ^ 2

' Sum the relative deviations
avgdev! = avgdev! + temp!
Next i%

' Calculate average
sum! = sum! / npts%

' Calculate relative deviation percent
avgdev! = Sqr(avgdev! / (npts% - 1))

' Calculate standard deviation percent
If sum! > 0# Then
avgdev! = (avgdev! / sum!) * 100#
End If

Exit Sub

' Errors
LeastMathNonLinearDeviationError:
MsgBox Error$, vbOKOnly + vbCritical, "LeastMathNonLinearDeviation"
ierror = True
Exit Sub

End Sub

Sub LeastGauss(kmax As Integer, X() As Double, Y() As Double, a() As Double)
' This routine solves a set of kmax simultaneous algebraic equations of the form A*x = y.
' The method used is gaussian elimination with back row substitution.  A full search for
' pivot elements is done and row and column interchange is done.

ierror = False
On Error GoTo LeastGaussError

Dim i As Integer, j As Integer
Dim ii As Integer, jj As Integer
Dim ipiv As Integer, jpiv As Integer
Dim itemp As Integer

ReDim ktemp(1 To MAXCOEFF9%) As Integer

Dim alarge As Double, fac As Double
Dim temp As Double

ReDim c(1 To MAXCOEFF9%) As Double

' Zero the returned coefficients array
For i% = 1 To MAXCOEFF9%
X#(i%) = 0#
Next i%

' Variable "ktemp%(i%)" keeps track of column interchanges
For i% = 1 To kmax%
ktemp%(i%) = i%
Next i%

' Start the elimination process
For i% = 1 To kmax%

' Search for pivot element
alarge# = 0#
ipiv% = i%
jpiv% = i%

For ii% = i% To kmax%
For jj% = i% To kmax%
If Abs(a#(ii%, jj%)) > Abs(alarge#) Then
ipiv% = ii%
jpiv% = jj%
alarge# = a#(ii%, jj%)
End If
Next jj%
Next ii%

' Interchange rows and columns to move pivot element to diagonal
If ipiv% <> i% Then
For ii% = 1 To kmax%
temp# = a#(i%, ii%)
a#(i%, ii%) = a#(ipiv%, ii%)
a#(ipiv%, ii%) = temp#
Next ii%

temp# = Y#(i%)
Y#(i%) = Y#(ipiv%)
Y#(ipiv%) = temp#
End If

If jpiv% <> i% Then
For ii% = 1 To kmax%
temp# = a#(ii%, i%)
a#(ii%, i%) = a#(ii%, jpiv%)
a#(ii%, jpiv%) = temp#
Next ii%

itemp% = ktemp%(i%)
ktemp%(i%) = ktemp%(jpiv%)
ktemp%(jpiv%) = itemp%
End If

' Do a round of elimination
For ii% = i% + 1 To kmax%
If Abs(a#(i%, i%)) < 1E-100 Then GoTo LeastGaussSingularMatrix
fac# = a#(ii%, i%) / a#(i%, i%)
Y#(ii%) = Y#(ii%) - fac# * Y#(i%)
a#(ii%, i%) = 0#

For jj% = i% + 1 To kmax%
a#(ii%, jj%) = a#(ii%, jj%) - fac# * a#(i%, jj%)
Next jj%
Next ii%

For ii% = i% + 1 To kmax%
a#(i%, ii%) = a#(i%, ii%) / a#(i%, i%)
Next ii%

' Check for bad matrix
If Abs(a#(i%, i%)) < 1E-100 Then GoTo LeastGaussSingularMatrix
Y#(i%) = Y#(i%) / a#(i%, i%)
a#(i%, i%) = 1#
Next i%

' Do the back substitution
c#(kmax%) = Y#(kmax%)
For i% = 1 To kmax% - 1
j% = kmax% - i%
temp# = 0#

For ii% = j% + 1 To kmax%
temp# = temp# + a#(j%, ii%) * c#(ii%)
Next ii%

c#(j%) = Y#(j%) - temp#
Next i%

' Put everything back in the correct order
For i% = 1 To kmax%
ii% = ktemp%(i%)
X#(ii%) = c#(i%)
Next i%

Exit Sub

' Errors
LeastGaussError:
MsgBox Error$, vbOKOnly + vbCritical, "LeastGauss"
ierror = True
Exit Sub

LeastGaussSingularMatrix:
msg$ = "Singular matrix"
MsgBox msg$, vbOKOnly + vbExclamation, "LeastGauss"
ierror = True
Exit Sub

End Sub

