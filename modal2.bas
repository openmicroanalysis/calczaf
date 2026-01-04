Attribute VB_Name = "CodeMODAL2"
' (c) Copyright 1995-2026 by John J. Donovan
Option Explicit

Sub ModalFitModal(numelms As Integer, numstds As Integer, upercents() As Double, spercents() As Double, fitcoeff As Double, weightflag As Integer)
' Run the modal analysis fit calculation
' numelms%                                  number of elements in arrays
' numstds%                                  number of standards in the arrays
' upercents#(1 to MAXCHAN%)                 unknown weight percents
' spercents#(1 to MAXCHAN%, 1 to MAXSTD%)   standard weight percents
' fitcoeff#                                 measure of "closeness" to standard

ierror = False
On Error GoTo ModalFitModalError

Dim i As Integer, j As Integer
Dim determinate As Double
Dim chisq As Double
Dim avsumsq As Double

ReDim temp(1 To MAXCHAN%) As Double
ReDim sig(1 To MAXCHAN%) As Double
ReDim predicated(1 To MAXCHAN%) As Double
ReDim residual(1 To MAXCHAN%) As Double
ReDim fitparam(1 To MAXSTD%) As Double
ReDim cvm(1 To MAXSTD%, 1 To MAXSTD%) As Double
ReDim se(1 To MAXSTD%) As Double

' Set weighting factors based on average concentrations in the
' standards for each element.
For i% = 1 To numelms%
sig#(i%) = 100#
temp#(i%) = 0#

' Calculate average for this element and set weighting factor
If weightflag% Then
For j% = 1 To numstds%
temp#(i%) = temp#(i%) + spercents#(i%, j%)
Next j%
temp#(i%) = temp#(i%) / numstds%
If temp#(i%) <> 0# Then sig#(i%) = 1# / temp#(i%)
End If

Next i%

' Fit unknown to a linear set of standards
If numstds% = 1 Then
avsumsq# = 0#
For i% = 1 To numelms%
avsumsq# = avsumsq# + (upercents#(i%) - spercents#(i%, 1)) ^ 2
Next i%

' Multi standard fit
Else
Call ModalLINREG(spercents#(), upercents#(), numelms%, numstds%, MAXCHAN%, MAXSTD%, sig#(), determinate#, chisq#, avsumsq#, predicated#(), residual#(), fitparam#(), cvm#(), se#())
If ierror Then Exit Sub
End If

' Return fit
If numstds% > 2 Then avsumsq# = Sqr(Abs(avsumsq#)) / numstds%
fitcoeff# = avsumsq#

Exit Sub

' Errors
ModalFitModalError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalFitModal"
ierror = True
Exit Sub

End Sub

Sub ModalGetStats(tData() As Double, n As Integer, mean As Double, stddev As Double)
' A vector of data (tdata) is received which has n data points.  The mean (mean) and second moment (stddev)
' are computed.  These are used in scaled regression.

ierror = False
On Error GoTo ModalGetStatsError

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
stddev# = Sqr(sum#)

Exit Sub

' Errors
ModalGetStatsError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalGetStats"
ierror = True
Exit Sub

End Sub

Sub ModalLINREG(X() As Double, Y() As Double, n As Integer, p As Integer, mp As Integer, np As Integer, sig() As Double, determinate As Double, chisq As Double, avsumsq As Double, predicted() As Double, residual() As Double, fitparam() As Double, cvm() As Double, se() As Double)
' This routine accepts a user supplied design matrix and dependent variable matrix and performs centered and scaled SVD regresssion.
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
On Error GoTo ModalLINREGError

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
If p% < 2 Then GoTo ModalLINREGTooFewFit
If p% > np% Then GoTo ModalLINREGTooManyFit
If n% > mp% Then GoTo ModalLINREGTooManyData
If n% <= p% Then GoTo ModalLINREGNotEnoughData

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
dummy#(i) = X#(i%, j%)
Next i%

Call ModalGetStats(dummy#(), n%, mean#(j - 1), stddev#(j% - 1))
If ierror Then Exit Sub
Next j%

Call ModalGetStats(Y#(), n%, meany#, stdy#)
If ierror Then Exit Sub

For i% = 1 To n%
For j% = 1 To p% - 1
X1#(i%, j%) = (X#(i%, j% + 1) - mean#(j%)) / stddev#(j%)
Next j%
Y1#(i) = (Y#(i) - meany#) / stdy#
Next i%
newp% = p% - 1

Else
For i% = 1 To n%
For j% = 1 To p%
X1#(i%, j%) = X#(i%, j%)
Next j%
Y1#(i) = Y#(i%)
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
Call Plan3dLUDCMP(f#(), newp%, np%, indx%(), determinate#)
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

Call ModalSVDCMP(u#(), n%, newp%, mp%, np%, w#(), v#())
If ierror Then Exit Sub

wmax# = 0#
For j% = 1 To newp%
If w#(j%) > wmax# Then wmax# = w#(j%)
Next j%
thresh# = TOL# * wmax#
For j% = 1 To newp%
If w#(j%) < thresh# Then w#(j%) = 0#
Next j%

Call ModalSVBKSB(u#(), w#(), v#(), n%, newp%, b#(), fitparam#())
If ierror Then Exit Sub

' Calculate the covariance matrix of the fit parameters
Call ModalSVDVAR(v#(), newp%, w#(), cvm#())
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
predicted#(i%) = predicted#(i%) + X#(i%, j%) * fitparam#(j%)
Next j%
residual#(i%) = Y#(i%) - predicted#(i%)
sum# = sum# + residual#(i%)
sumsq# = sumsq# + residual#(i%) * residual#(i%)
Next i%
avsumsq# = sumsq# / (n% - p%)

' Calculate standard errors of fit parameters, se(j)
For i% = 1 To n%
For j% = 1 To p%
x2t#(j%, i%) = X#(i%, j%)
Next j%
Next i%
For i% = 1 To p%
For j% = 1 To p%
f1#(i%, j%) = 0#
For k% = 1 To n%
f1#(i%, j%) = f1#(i%, j%) + x2t#(i%, k%) * X#(k%, j%)
Next k%
Next j%
Next i%

For i% = 1 To p%
For j% = 1 To p%
fi#(i%, j%) = 0#
Next j%
fi#(i%, i%) = 1#
Next i%

Call Plan3dLUDCMP(f1#(), p%, np%, indxf%(), df#)
If ierror Then Exit Sub

For j% = 1 To p%
For i% = 1 To p%
temp#(i%) = fi#(i%, j%)  ' fi#(1, j%)
Next i%

Call Plan3dLUBKSB(f1#(), p%, np%, indxf%(), temp#())
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
sum# = sum# + fitparam#(j%) * X#(i%, j%)
Next j%
chisq# = chisq# + ((Y#(i%) - sum#) / sig#(i%)) ^ 2
Next i%

Exit Sub

' Errors
ModalLINREGError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalLINREG"
ierror = True
Exit Sub

ModalLINREGTooFewFit:
msg$ = "Too few fit parameters"
MsgBox msg$, vbOKOnly + vbExclamation, "ModalLINREG"
ierror = True
Exit Sub

ModalLINREGTooManyFit:
msg$ = "Too many fit parameters"
MsgBox msg$, vbOKOnly + vbExclamation, "ModalLINREG"
ierror = True
Exit Sub

ModalLINREGTooManyData:
msg$ = "Too many data points"
MsgBox msg$, vbOKOnly + vbExclamation, "ModalLINREG"
ierror = True
Exit Sub

ModalLINREGNotEnoughData:
msg$ = "Not enough data points for this fit order"
MsgBox msg$, vbOKOnly + vbExclamation, "ModalLINREG"
ierror = True
Exit Sub

End Sub

Function ModalMAX(a As Double, b As Double) As Double
' Returns the maximum of two values

ierror = False
On Error GoTo ModalMAXError

If a# > b# Then
ModalMAX# = a#
Else
ModalMAX# = b#
End If

Exit Function

' Errors
ModalMAXError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalMAX"
ierror = True
Exit Function

End Function

Function ModalSIGN(a As Double, b As Double) As Double
' Performs a sign transfer by returning the absolute value of the
' first argument multiplied by the sign of the second argument.

ierror = False
On Error GoTo ModalSIGNError

Dim n As Integer

n% = 1
If b# < 0# Then n% = -1
ModalSIGN# = Abs(a#) * n%

Exit Function

' Errors
ModalSIGNError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalSIGN"
ierror = True
Exit Function

End Function

Sub ModalSVBKSB(u() As Double, w() As Double, v() As Double, m As Integer, n As Integer, b() As Double, X() As Double)
' Back substitution
' Passed/returned:
' u#(1 to mp%, 1 to np%)
' w#(1 to np%)
' v#(1 to np%, 1 to np%)
' m%
' n%
' b#(1 to mp%)
' x#(1 to np%)

ierror = False
On Error GoTo ModalSVBKSBError

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
X#(j%) = s#
Next j%

Exit Sub

' Errors
ModalSVBKSBError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalSVBKSB"
ierror = True
Exit Sub

End Sub

Sub ModalSVDCMP(a() As Double, m As Integer, n As Integer, mp As Integer, np As Integer, w() As Double, v() As Double)
' Compute singular value decomposition
' Passed/Returned values:
' a#(1 to mp%, 1 to np%)
' m%
' n%
' mp%
' np%
' w#(1 to np%)
' v#(1 to np%, 1 to np%)

ierror = False
On Error GoTo ModalSVDCMPError

Dim i As Integer, l As Integer, k As Integer
Dim j As Integer, nm As Integer
Dim iter As Integer, maxiter As Integer

Dim g As Double, rscale As Double, anorm As Double
Dim s As Double, f As Double, c As Double, Y As Double
Dim Z As Double, X As Double, h As Double

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
g# = -ModalSIGN(Sqr(s#), f#)
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
g# = -ModalSIGN(Sqr(s#), f#)
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
anorm# = ModalMAX(anorm#, (Abs(w#(i%)) + Abs(rv1#(i%))))
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
Y# = a#(j%, nm%)
Z# = a#(j%, i%)
a#(j%, nm%) = (Y# * c#) + (Z# * s#)
a#(j%, i%) = -(Y# * s#) + (Z# * c#)
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

If iter% = maxiter% Then GoTo ModalSVDCMPNotConverged

X# = w#(l%)
nm% = k% - 1
Y# = w#(nm%)
g# = rv1#(nm%)
h# = rv1#(k%)
f# = ((Y# - Z#) * (Y# + Z#) + (g# - h#) * (g# + h#)) / (2# * h# * Y#)
g# = Sqr(f# * f# + 1#)
f# = ((X# - Z#) * (X# + Z#) + h# * ((Y# / (f# + ModalSIGN(g#, f#))) - h#)) / X#
c# = 1#
s# = 1#

For j% = l To nm%
i% = j% + 1
g# = rv1#(i%)
Y# = w#(i%)
h# = s# * g#
g# = c# * g#
Z# = Sqr(f# * f# + h# * h#)
rv1#(j%) = Z#
c# = f# / Z#
s# = h# / Z#
f# = (X# * c#) + (g# * s#)
g# = -(X# * s#) + (g# * c#)
h# = Y# * s#
Y# = Y# * c#
For nm% = 1 To n%
X# = v#(nm%, j%)
Z# = v#(nm%, i%)
v#(nm%, j%) = (X# * c#) + (Z# * s#)
v#(nm%, i%) = -(X# * s#) + (Z# * c#)
Next nm%
Z = Sqr(f# * f# + h# * h#)
w#(j%) = Z#
If Z# <> 0# Then
Z# = 1# / Z#
c# = f# * Z#
s# = h# * Z#
End If
f# = (c# * g#) + (s# * Y#)
X# = -(s# * g#) + (c# * Y#)
For nm% = 1 To m%
Y# = a#(nm%, j%)
Z# = a#(nm%, i%)
a#(nm%, j%) = (Y# * c#) + (Z# * s#)
a#(nm%, i%) = -(Y# * s#) + (Z# * c#)
Next nm%
Next j%

rv1#(l%) = 0#
rv1#(k%) = f#
w#(k%) = X#
Next iter%
3:
Next k%

Exit Sub

' Errors
ModalSVDCMPError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalSVDCMP"
ierror = True
Exit Sub

ModalSVDCMPNotConverged:
msg$ = "Not converged after " & Str$(maxiter%) & " iterations"
MsgBox msg$, vbOKOnly + vbExclamation, "ModalSVDCMP"
ierror = True
Exit Sub

End Sub

Sub ModalSVDVAR(v() As Double, ma As Integer, w() As Double, cvm() As Double)
' Passed/returned parameters:
' v#(1 to np%, 1 to np%)
' ma%
' w#(1 to np%)
' cvm#(1 to ncvm%, 1 to ncvm%)

ierror = False
On Error GoTo ModalSVDVARError

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
ModalSVDVARError:
MsgBox Error$, vbOKOnly + vbCritical, "ModalSVDVAR"
ierror = True
Exit Sub

End Sub

