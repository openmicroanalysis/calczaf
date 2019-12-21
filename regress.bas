Attribute VB_Name = "CodeREGRESS"
' (c) Copyright 1995-2020 by John J. Donovan
Option Explicit

Sub RegressGetStats(tData() As Double, n As Integer, np As Integer, mean As Double, Std As Double)
' A vector of data (tdata) is received which is size np and has n data points.
' The mean (mean) and second moment (std) are computed.  These are used in scaled regression.
' Written by Don Snyder

ierror = False
On Error GoTo RegressGetStatsError

Dim i As Integer
Dim sum As Double

' Compute mean
sum# = 0#
For i% = 1 To n%
sum# = sum# + tData#(i%)
Next i%
mean# = sum# / n%
 
' Compute second moment
sum# = 0#
For i% = 1 To n%
sum# = sum# + (tData#(i%) - mean#) * (tData#(i%) - mean#)
Next i%
Std# = Sqr(sum#)

Exit Sub

' Errors
RegressGetStatsError:
MsgBox Error$, vbOKOnly + vbCritical, "RegressGetStats"
ierror = True
Exit Sub

End Sub

Sub RegressLINREG(x() As Double, y() As Double, n As Integer, p As Integer, mp As Integer, np As Integer, sig() As Double, determinate As Double, chisq As Double, avsumsq As Double, predicted() As Double, residual() As Double, fitparam() As Double, cvm() As Double, se() As Double)
' This routine accepts a user supplied design matrix and dependent variable matrix and
' performs centered and scaled SVD regresssion.
' Written by Don Snyder
'
' Passed variables:
' x#(1 to mp%,1 to np%)     design matrix
' y#(1 to mp%)              dependent variable column vector
' n%                        number of data points
' p%                        number of fit parameters (before centering)
' mp%                       physical dimension of the design matrix (mp x np)
' np%                       physical dimension of the design matrix (mp x np)
' sig#(1 to mp%)            weighting factors
'
' Returned variables:
' determinate#              determinate
' chisq#                    chi-squared statistic
' avsumsq#                  sum of squares of the residuals divided by n-p
' predicted#(1 to mp%)      predicted values of the dependent variable
' residual#(1 to mp%)       residuals
' fitparam#(1 to np%)       fit parameter matrix
' cvm#(1 to np%, 1 to np%)  covariance matrix of the fit parameters
' se#(1 to np%)             vector of std errors of the fit parameters

ierror = False
On Error GoTo RegressLINREGError

Dim i As Integer, j As Integer, k As Integer, newp As Integer
Dim df As Double, meany As Double, stdy As Double
Dim wmax As Double, thresh As Double
Dim sum As Double, sumsq As Double

ReDim b(1 To mp%) As Double
ReDim u(1 To mp%, 1 To np%) As Double
ReDim v(1 To np%, 1 To np%) As Double
ReDim w(1 To np%) As Double

ReDim dummy(1 To mp%) As Double
ReDim f(1 To np%, 1 To np%) As Double
ReDim fi(1 To np%, 1 To np%) As Double
ReDim f1(1 To np%, 1 To np%) As Double
ReDim mean(1 To np%) As Double

ReDim indx(1 To np%) As Integer
ReDim indxf(1 To np%) As Integer

ReDim X1(1 To mp%, 1 To np%) As Double
ReDim xt(1 To np%, 1 To mp%) As Double
ReDim x2t(1 To np%, 1 To mp%) As Double
ReDim Y1(1 To mp%) As Double
ReDim stddev(1 To np%) As Double

ReDim temp(1 To p%) As Double

Const TOL# = 0.00001

' Cautions concerning the number of fit parameters
If p% < 2 Then GoTo RegressLINREGTooFewFit
If p% > np% Then GoTo RegressLINREGTooManyFit
If n% > mp% Then GoTo RegressLINREGTooManyData
If n% <= p% Then GoTo RegressLINREGNotEnoughData

' Initialize variables
For i% = 1 To np%
For j% = 1 To mp%
b#(j%) = 0#
u#(j%, i%) = 0#
Next j%
fitparam#(i%) = 0#
se#(i%) = 0#
w#(i%) = 0#
Next i%

For i% = 1 To np%
For j% = 1 To np%
cvm#(i%, j%) = 0#
v#(i%, j%) = 0#
Next j%
Next i%
  
' Center and scale data for p > 2
If p% > 2 Then
For j% = 2 To p%
For i% = 1 To n%
dummy#(i) = x#(i%, j%)
Next i%

Call RegressGetStats(dummy#(), n%, mp%, mean#(j - 1), stddev#(j% - 1))
If ierror Then Exit Sub
Next j%

Call RegressGetStats(y#(), n%, mp%, meany#, stdy#)
If ierror Then Exit Sub

For i% = 1 To n%
For j% = 1 To p% - 1
X1#(i%, j%) = (x#(i%, j% + 1) - mean#(j%)) / stddev#(j%)
Next j%
Y1#(i) = (y#(i) - meany#) / stdy#
Next i%
newp% = p% - 1

Else
For i% = 1 To n%
For j% = 1 To p%
X1#(i%, j%) = x#(i%, j%)
Next j%
Y1#(i) = y#(i%)
Next i%
newp% = p%
End If

' Compute the transpose of the design matrix
For i% = 1 To n%
For j% = 1 To newp%
xt#(j%, i%) = X1#(i%, j%)
Next j%
Next i%

' Let f be the matrix xt * x1
For i% = 1 To newp%
For j% = 1 To newp%
f#(i%, j%) = 0#
For k% = 1 To n%
f#(i%, j%) = f#(i%, j%) + xt#(i%, k%) * X1#(k%, j%)
Next k%
Next j%
Next i%

' Calculate the determinate, d, of f = xt * x1
Call RegressLUDCMP(f#(), newp%, np%, indx%(), determinate#)
If ierror Then Exit Sub

For j% = 1 To newp%
determinate# = determinate# * f#(j%, j%)
Next j%

' Perform regression using SVD (Press, et al., 1986)
For i% = 1 To n%
For j% = 1 To newp%
u#(i%, j%) = X1#(i%, j%) / sig#(i%)
Next j%
b#(i%) = Y1#(i%) / sig#(i%)
Next i%

Call RegressSVDCMP(u#(), n%, newp%, w#(), v#())
If ierror Then Exit Sub

wmax# = 0#
For j% = 1 To newp%
If w#(j%) > wmax# Then wmax# = w#(j%)
Next j%
thresh# = TOL# * wmax#
For j% = 1 To newp%
If w#(j%) < thresh# Then w#(j%) = 0#
Next j%

Call RegressSVBKSB(u#(), w#(), v#(), n%, newp%, mp%, np%, b#(), fitparam#())
If ierror Then Exit Sub

' Calculate the covariance matrix of the fit parameters
Call RegressSVDVAR(v#(), newp%, w#(), cvm#())
If ierror Then Exit Sub

' Un-center and un-scale data and fit parameters for p > 2
If p% > 2 Then
For j% = p% To 2 Step -1
fitparam#(j%) = fitparam#(j% - 1) * (stdy# / stddev#(j% - 1))
Next j%
fitparam#(1) = meany#
For j% = p To 2 Step -1
fitparam#(1) = fitparam#(1) - fitparam#(j%) * mean#(j% - 1)
Next j%
End If

' Calculate predicted values, e(i) and sum of squares statistic (avsumsq)
For i% = 1 To n%
predicted#(i%) = 0#
residual#(i%) = 0#
Next i%

sum# = 0#
sumsq# = 0#
For i% = 1 To n%
For j% = 1 To p%
predicted#(i%) = predicted#(i%) + x#(i%, j%) * fitparam#(j%)
Next j%
residual#(i%) = y#(i%) - predicted#(i%)
sum# = sum# + residual#(i%)
sumsq# = sumsq# + residual#(i%) * residual#(i%)
Next i%
avsumsq# = sumsq# / (n% - p%)

' Calculate standard errors of fit parameters, se(j)
For i% = 1 To n%
For j% = 1 To p%
x2t#(j%, i%) = x#(i%, j%)
Next j%
Next i%
For i% = 1 To p%
For j% = 1 To p%
f1#(i%, j%) = 0#
For k% = 1 To n%
f1#(i%, j%) = f1#(i%, j%) + x2t#(i%, k%) * x#(k%, j%)
Next k%
Next j%
Next i%

For i% = 1 To p%
For j% = 1 To p%
fi#(i%, j%) = 0#
Next j%
fi#(i%, i%) = 1#
Next i%

Call RegressLUDCMP(f1#(), p%, np%, indxf%(), df#)
If ierror Then Exit Sub

For j% = 1 To p%
For i% = 1 To p%
temp#(i%) = fi#(i%, j%)  ' fi#(1, j%)
Next i%

Call RegressLUBKSB(f1#(), p%, np%, indxf%(), temp#())
If ierror Then Exit Sub

For i% = 1 To p%
fi#(i%, j%) = temp#(i%)
Next i%
Next j%

For j% = 1 To p%
se#(j%) = Sqr(Abs(avsumsq# * fi#(j%, j%)))
Next j%

' Evaluate chi-square statistic
chisq# = 0#
For i% = 1 To n%
sum# = 0#
For j% = 1 To p%
sum# = sum# + fitparam#(j%) * x#(i%, j%)
Next j%
chisq# = chisq# + ((y#(i%) - sum#) / sig#(i%)) ^ 2
Next i%

Exit Sub

' Errors
RegressLINREGError:
MsgBox Error$, vbOKOnly + vbCritical, "RegressLINREG"
ierror = True
Exit Sub

RegressLINREGTooFewFit:
msg$ = "Too few fit parameters"
MsgBox msg$, vbOKOnly + vbExclamation, "RegressLINREG"
ierror = True
Exit Sub

RegressLINREGTooManyFit:
msg$ = "Too many fit parameters"
MsgBox msg$, vbOKOnly + vbExclamation, "RegressLINREG"
ierror = True
Exit Sub

RegressLINREGTooManyData:
msg$ = "Too many data points"
MsgBox msg$, vbOKOnly + vbExclamation, "RegressLINREG"
ierror = True
Exit Sub

RegressLINREGNotEnoughData:
msg$ = "Not enough data for this fit"
MsgBox msg$, vbOKOnly + vbExclamation, "RegressLINREG"
ierror = True
Exit Sub

End Sub

Function RegressMAX(a As Double, b As Double) As Double
' Returns the maximum of two values

ierror = False
On Error GoTo RegressMAXError

If a# > b# Then
RegressMAX# = a#
Else
RegressMAX# = b#
End If

Exit Function

' Errors
RegressMAXError:
MsgBox Error$, vbOKOnly + vbCritical, "RegressMAX"
ierror = True
Exit Function

End Function

Function RegressSIGN(a As Double, b As Double) As Double
' Performs a sign transfer by returning the absolute value of the
' first argument multiplied by the sign of the second argument.

ierror = False
On Error GoTo RegressSIGNError

Dim n As Integer

n% = 1
If b# < 0# Then n% = -1
RegressSIGN# = Abs(a#) * n%

Exit Function

' Errors
RegressSIGNError:
MsgBox Error$, vbOKOnly + vbCritical, "RegressSIGN"
ierror = True
Exit Function

End Function

Sub RegressSVBKSB(u() As Double, w() As Double, v() As Double, m As Integer, n As Integer, mp As Integer, np As Integer, b() As Double, x() As Double)
' Back substitution. Modified From Numerical Recipes.
' Passed/returned:
' u#(1 to mp%, 1 to np%)
' w#(1 to np%)
' v#(1 to np%, 1 to np%)
' m%
' n%
' np%
' b#(1 to mp%)
' x#(1 to np%)

ierror = False
On Error GoTo RegressSVBKSBError

Dim j As Integer, i As Integer, jj As Integer
Dim s As Double

ReDim tmp(1 To n%) As Double

For j% = 1 To n%
s# = 0#
If w#(j%) <> 0# Then
For i% = 1 To m%
s# = s# + u#(i%, j%) * b#(i%)
Next i%
s# = s# / w#(j%)
End If
tmp#(j%) = s#
Next j%

For j% = 1 To n%
s# = 0#
For jj% = 1 To n%
s# = s# + v#(j%, jj%) * tmp#(jj%)
Next jj%
x#(j%) = s#
Next j%

Exit Sub

' Errors
RegressSVBKSBError:
MsgBox Error$, vbOKOnly + vbCritical, "RegressSVBKSB"
ierror = True
Exit Sub

End Sub

Sub RegressSVDCMP(a() As Double, m As Integer, n As Integer, w() As Double, v() As Double)
' Compute singular value decomposition. Modified From Numerical Recipes.
' Passed/Returned values:
' a#(1 to mp%, 1 to np%)
' m%
' n%
' mp%
' np%
' w#(1 to np%)
' v#(1 to np%, 1 to np%)

ierror = False
On Error GoTo RegressSVDCMPError

Dim i As Integer, l As Integer, k As Integer
Dim j As Integer, nm As Integer
Dim iter As Integer, maxiter As Integer

Dim g As Double, rscale As Double, anorm As Double
Dim s As Double, f As Double, c As Double, y As Double
Dim Z As Double, x As Double, h As Double

ReDim rv1(1 To n%) As Double

maxiter% = 30  ' maximum iterations

g# = 0#
rscale# = 0#
anorm# = 0#

For i% = 1 To n%
l% = i% + 1
rv1#(i%) = rscale# * g#
g# = 0#
s# = 0#
rscale# = 0#
If i% <= m% Then
For k% = i% To m%
rscale# = rscale# + Abs(a#(k%, i%))
Next k%
If rscale# <> 0# Then
For k% = i% To m%
a#(k%, i%) = a#(k%, i%) / rscale#
s# = s# + a#(k%, i%) * a#(k%, i%)
Next k%
f# = a#(i%, i%)
g# = -RegressSIGN(Sqr(s#), f#)
h# = f# * g# - s#
a#(i%, i%) = f# - g#

If i% <> n% Then
For j% = l% To n%
s# = 0#
For k% = i% To m%
s# = s# + a#(k%, i%) * a#(k%, j%)
Next k%
f# = s# / h#
For k% = i% To m%
a#(k%, j%) = a#(k%, j%) + f# * a#(k%, i%)
Next k%
Next j%
End If

For k% = i% To m%
a#(k%, i%) = rscale# * a#(k%, i%)
Next k%
End If
End If

w#(i%) = rscale# * g#
g# = 0#
s# = 0#
rscale# = 0#
If (i% <= m%) And (i% <> n%) Then
For k% = l% To n%
rscale# = rscale# + Abs(a#(i%, k%))
Next k%
If rscale# <> 0# Then
For k% = l% To n%
a#(i%, k%) = a#(i%, k%) / rscale#
s# = s# + a#(i%, k%) * a#(i%, k%)
Next k%
f# = a#(i%, l%)
g# = -RegressSIGN(Sqr(s#), f#)
h# = f# * g# - s#
a#(i%, l%) = f# - g#
For k% = l% To n%
rv1#(k) = a#(i%, k%) / h#
Next k%
If i% <> m% Then
For j% = l% To m%
s# = 0#
For k% = l% To n%
s# = s + a#(j, k) * a#(i%, k%)
Next k%
For k% = l% To n%
a#(j%, k%) = a#(j%, k%) + s * rv1#(k%)
Next k%
Next j%
End If
For k% = l% To n%
a#(i%, k%) = rscale# * a#(i%, k%)
Next k%
End If
End If
anorm# = RegressMAX(anorm#, (Abs(w#(i%)) + Abs(rv1#(i%))))
Next i%

For i% = n% To 1 Step -1
If i% < n% Then
If g# <> 0# Then
For j% = l% To n%
v#(j%, i%) = (a#(i%, j%) / a#(i, l)) / g#
Next j%
For j% = l% To n%
s# = 0#
For k% = l% To n%
s# = s# + a#(i%, k%) * v#(k%, j%)
Next k%
For k% = l% To n%
v#(k%, j%) = v#(k%, j%) + s * v#(k%, i%)
Next k%
Next j%
End If
For j% = l% To n%
v#(i%, j%) = 0#
v#(j%, i%) = 0#
Next j%
End If
v#(i%, i%) = 1#
g# = rv1#(i%)
l% = i%
Next i%

For i% = n% To 1 Step -1
l% = i% + 1
g# = w#(i%)
If i% < n% Then
For j% = l% To n%
a#(i%, j%) = 0#
Next j%
End If
If g# <> 0# Then
g# = 1# / g#

If i% <> n% Then
For j% = l% To n%

s# = 0#
For k% = l% To m%
s# = s# + a#(k%, i%) * a#(k%, j%)
Next k%
f# = (s# / a#(i%, i%)) * g#
For k% = i% To m%
a#(k%, j%) = a#(k%, j%) + f# * a#(k%, i%)
Next k%

Next j%
End If

For j% = i% To m%
a#(j%, i%) = a#(j%, i%) * g#
Next j%
Else
For j% = i% To m%
a#(j%, i%) = 0#
Next j%
End If
a#(i%, i%) = a#(i%, i%) + 1#
Next i%

For k% = n% To 1 Step -1
For iter% = 1 To maxiter%
For l% = k% To 1 Step -1
nm% = l% - 1
If (Abs(rv1#(l%)) + anorm#) = anorm# Then GoTo 2
If (Abs(w#(nm%)) + anorm#) = anorm# Then GoTo 1
Next l%

1:
c# = 0#
s# = 1#
For i% = l% To k%
f# = s# * rv1#(i%)

If (Abs(f#) + anorm#) <> anorm Then
g# = w#(i%)
h# = Sqr(f# * f# + g# * g#)
w#(i%) = h#
h# = 1# / h#
c# = (g# * h#)
s# = -(f# * h#)

For j% = 1 To m%
y# = a#(j%, nm%)
Z# = a#(j%, i%)
a#(j%, nm%) = (y# * c#) + (Z# * s#)
a#(j%, i%) = -(y# * s#) + (Z# * c#)
Next j%
End If

Next i%

2:
Z# = w#(k%)
If l% = k% Then
If Z# < 0# Then
w#(k%) = -Z#
For j% = 1 To n%
v#(j%, k%) = -v#(j%, k%)
Next j%
End If
GoTo 3
End If

If iter% = maxiter% Then GoTo RegressSVDCMPNotConverged

x# = w#(l%)
nm% = k% - 1
y# = w#(nm%)
g# = rv1#(nm%)
h# = rv1#(k%)
f# = ((y# - Z#) * (y# + Z#) + (g# - h#) * (g# + h#)) / (2# * h# * y#)
g# = Sqr(f# * f# + 1#)
f# = ((x# - Z#) * (x# + Z#) + h# * ((y# / (f# + RegressSIGN(g#, f#))) - h#)) / x#
c# = 1#
s# = 1#

For j% = l To nm%
i% = j% + 1
g# = rv1#(i%)
y# = w#(i%)
h# = s# * g#
g# = c# * g#
Z# = Sqr(f# * f# + h# * h#)
rv1#(j%) = Z#
c# = f# / Z#
s# = h# / Z#
f# = (x# * c#) + (g# * s#)
g# = -(x# * s#) + (g# * c#)
h# = y# * s#
y# = y# * c#
For nm% = 1 To n%
x# = v#(nm%, j%)
Z# = v#(nm%, i%)
v#(nm%, j%) = (x# * c#) + (Z# * s#)
v#(nm%, i%) = -(x# * s#) + (Z# * c#)
Next nm%
Z = Sqr(f# * f# + h# * h#)
w#(j%) = Z#
If Z# <> 0# Then
Z# = 1# / Z#
c# = f# * Z#
s# = h# * Z#
End If
f# = (c# * g#) + (s# * y#)
x# = -(s# * g#) + (c# * y#)
For nm% = 1 To m%
y# = a#(nm%, j%)
Z# = a#(nm%, i%)
a#(nm%, j%) = (y# * c#) + (Z# * s#)
a#(nm%, i%) = -(y# * s#) + (Z# * c#)
Next nm%
Next j%

rv1#(l%) = 0#
rv1#(k%) = f#
w#(k%) = x#
Next iter%
3:
Next k%

Exit Sub

' Errors
RegressSVDCMPError:
MsgBox Error$, vbOKOnly + vbCritical, "RegressSVDCMP"
ierror = True
Exit Sub

RegressSVDCMPNotConverged:
msg$ = "Not converged after " & Str$(maxiter%) & " iterations"
MsgBox msg$, vbOKOnly + vbExclamation, "RegressSVDCMP"
ierror = True
Exit Sub

End Sub

Sub RegressSVDVAR(v() As Double, ma As Integer, w() As Double, cvm() As Double)
' Modified From Numerical Recipes.
' Passed/returned parameters:
' v#(1 to np%, 1 to np%)
' ma%
' w#(1 to np%)
' cvm#(1 to ncvm%, 1 to ncvm%)

ierror = False
On Error GoTo RegressSVDVARError

Dim i As Integer, j As Integer, k As Integer
Dim sum As Double

ReDim wti(1 To ma%) As Double

For i% = 1 To ma%
wti#(i%) = 0#
If w#(i) <> 0# Then wti#(i%) = 1# / (w#(i%) * w#(i%))
Next i%

For i% = 1 To ma%
For j% = 1 To i%
sum# = 0#
For k% = 1 To ma%
sum# = sum# + v#(i%, k%) * v#(j%, k%) * wti#(k%)
Next k%
cvm#(i%, j%) = sum#
cvm#(j%, i%) = sum#
Next j%
Next i%

Exit Sub

' Errors
RegressSVDVARError:
MsgBox Error$, vbOKOnly + vbCritical, "RegressSVDVAR"
ierror = True
Exit Sub

End Sub

Sub RegressLUBKSB(a() As Double, n As Integer, np As Integer, indx() As Integer, b() As Double)
' Modified From Numerical Recipes.

ierror = False
On Error GoTo RegressLUBKSBError

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
RegressLUBKSBError:
MsgBox Error$, vbOKOnly + vbCritical, "RegressLUBKSB"
ierror = True
Exit Sub

End Sub

Sub RegressLUDCMP(a() As Double, n As Integer, np As Integer, indx() As Integer, d As Double)
' Modified From Numerical Recipes.

ierror = False
On Error GoTo RegressLUDCMPError

Dim i As Integer, j As Integer
Dim k As Integer, imax As Integer
Dim aamax As Double, sum As Double, dum As Double

Const TINY1# = 1E-20
ReDim vv(1 To n%) As Double

d# = 1#
For i% = 1 To n%
aamax# = 0#
    For j% = 1 To n%
    If Abs(a#(i%, j%)) > aamax# Then aamax# = Abs(a#(i%, j%))
    Next j%

' Check for singular matrix
If aamax# = 0# Then GoTo RegressLUDCMPSingularMatrix

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
RegressLUDCMPError:
MsgBox Error$, vbOKOnly + vbCritical, "RegressLUDCMP"
ierror = True
Exit Sub

RegressLUDCMPSingularMatrix:
msg$ = "Singular matrix"
MsgBox msg$, vbOKOnly + vbExclamation, "RegressLUDCMP"
ierror = True
Exit Sub

End Sub


