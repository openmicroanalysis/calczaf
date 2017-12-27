Attribute VB_Name = "CodeSTUDENT"
' (c) Copyright 1995-2018 by John J. Donovan
Option Explicit

Function StudentBetacf(a As Single, b As Single, X As Single) As Single
' From Numerical Receipes

ierror = False
On Error GoTo StudentBetacfError

Const ITMAX% = 100
Const eps! = 0.0000003

Dim am As Single, bm As Single, az As Single, qab As Single, d As Single
Dim qap As Single, qam As Single, bz As Single, em As Single, tem As Single
Dim AP As Single, bp As Single, app As Single, bpp As Single, aold As Single
Dim m As Integer

am! = 1#
bm! = 1#
az! = 1#
qab! = a! + b!
qap! = a! + 1#
qam! = a! - 1#
bz! = 1# - qab! * X! / qap!

For m% = 1 To ITMAX%
em! = m%
tem! = em! + em!
d! = em! * (b! - m%) * X! / ((qam! + tem!) * (a! + tem!))
AP! = az! + d! * am!
bp! = bz! + d! * bm!
d! = -(a! + em!) * (qab! + em!) * X! / ((a! + tem!) * (qap! + tem!))
app! = AP! + d! * az!
bpp! = bp! + d! * bz!
aold! = az!
am! = AP! / bpp!
bm! = bp! / bpp!
az! = app! / bpp!
bz! = 1#

' Check for convergence
If Abs(az! - aold!) < eps! * Abs(az!) Then GoTo 1000
Next m%

' If we get here iteration failed
msg$ = "Either 'a' or 'b' large, or 'itmax' small"
MsgBox msg$, vbOKOnly + vbExclamation, "StudentBetacf"
ierror = True

1000:
StudentBetacf! = az!

Exit Function

' Errors
StudentBetacfError:
MsgBox Error$, vbOKOnly + vbCritical, "StudentBetacf"
ierror = True
Exit Function

End Function

Function StudentBetai(a As Single, b As Single, X As Single) As Single
' From Numerical Receipes

ierror = False
On Error GoTo StudentBetaiError

Dim bt As Single

If X! < 0# Or X! > 1# Then GoTo StudentBetaiBadX

If X! = 0# Or X! = 1# Then
bt! = 0#
Else
bt! = Exp(StudentGammln!(a! + b!) - StudentGammln!(a!) - StudentGammln!(b!) + a! * Log(X!) + b! * Log(1# - X!))
End If

If X < (a! + 1#) / (a! + b! + 2#) Then
StudentBetai! = bt! * StudentBetacf!(a!, b!, X!) / a!
Exit Function
Else
StudentBetai! = 1# - bt! * StudentBetacf!(b!, a!, 1# - X!) / b!
Exit Function
End If

Exit Function

' Errors
StudentBetaiError:
MsgBox Error$, vbOKOnly + vbCritical, "StudentBetai"
ierror = True
Exit Function

StudentBetaiBadX:
msg$ = "Parameter x is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "StudentBetai"
ierror = True
Exit Function

End Function

Sub StudentCalculateTable(tmsg As String)
' Table calculation (Student's "t" calculation)

ierror = False
On Error GoTo StudentCalculateTableError

Dim i As Integer
Dim Alpha As Single, df As Single, t As Single
Dim alpha60 As Single, alpha80 As Single
Dim alpha90 As Single, alpha95 As Single, alpha99 As Single

' Calculate table
tmsg$ = Format$("# pts", a80$) & Format$("d.f.", a80$) & Format$("0.60", a80$) & Format$("0.80", a80$) & Format$("0.90", a80$) & Format$("0.95", a80$) & Format$("0.99", a80$) & vbCrLf

Screen.MousePointer = vbHourglass
For i% = 1 To 49
df! = i%

Alpha! = 0.6
Call StudentGetT(df!, Alpha!, t!)
If ierror Then Exit Sub
alpha60 = t!

Alpha! = 0.8
Call StudentGetT(df!, Alpha!, t!)
If ierror Then Exit Sub
alpha80 = t!

Alpha! = 0.9
Call StudentGetT(df!, Alpha!, t!)
If ierror Then Exit Sub
alpha90 = t!

Alpha = 0.95
Call StudentGetT(df!, Alpha!, t!)
If ierror Then Exit Sub
alpha95 = t!

Alpha = 0.99
Call StudentGetT(df!, Alpha!, t!)
If ierror Then Exit Sub
alpha99 = t!

tmsg$ = tmsg$ & Format$(df! + 1, a80$) & Format$(df!, a80$) & Format$(Format$(alpha60!, f84$), a80$) & Format$(Format$(alpha80!, f84$), a80$) & Format$(Format$(alpha90!, f84$), a80$) & Format$(Format$(alpha95!, f84$), a80$) & Format$(Format$(alpha99!, f84$), a80$) & vbCrLf
Next i%

Screen.MousePointer = vbDefault
Exit Sub

' Errors
StudentCalculateTableError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StudentCalculateTable"
ierror = True
Exit Sub

End Sub

Function StudentGammln(xx As Single) As Single
' Modified from Numerical Receipes

ierror = False
On Error GoTo StudentGammlnError

Dim stp As Double, half As Double
Dim one As Double, fpf As Double, X As Double
Dim tmp As Double, ser As Double
Dim j As Integer

Static cof(6) As Double

cof#(1) = 76.18009173
cof#(2) = -86.50532033
cof#(3) = 24.01409822
cof#(4) = -1.231739516
cof#(5) = 0.00120858003
cof#(6) = -0.00000536382

stp# = 2.50662827465
half# = 0.5
one# = 1#
fpf# = 5.5

X# = xx! - one#
tmp# = X# + fpf#
tmp# = (X# + half#) * Log(tmp#) - tmp#
ser# = one#

For j% = 1 To 6
X# = X# + one#
ser# = ser# + cof#(j%) / X#
Next j%

StudentGammln! = tmp# + Log(stp# * ser#)
Exit Function

' Errors
StudentGammlnError:
MsgBox Error$, vbOKOnly + vbCritical, "StudentGammln"
ierror = True
Exit Function

End Function

Sub StudentGetT(df As Single, Alpha As Single, t As Single)
' Calculate two-sided Student's t test. Modified From Numerical Recipes.
'  df = number of degrees of freedom
'  alpha = the probability (0 < alpha < 1)
'  t = the returned critical value for a two-sided test
' BETAI which calls:
' BETACF
' GAMMLN
' BETAI
' Press et al. (1986) Numerical Recipes: The Art of Scientific
' Computing, Cambridge Univeristy Press, 818 pp.

ierror = False
On Error GoTo StudentGetTError

Const inf& = 500000
Const resolution! = 0.0001

Dim i As Long
Dim increment As Single
Dim slast As Single, studt As Single, tlast As Single

' Check for valid parameters
If df! <= 0# Or df > 1000 Then GoTo StudentGetTBadDf
If Alpha! <= 0# Or Alpha! >= 1# Then GoTo StudentGetTBadAlpha

' Set initial parameters
increment! = 10#
t! = 0#

' Perform calculations
For i& = 1 To inf&
studt! = 1# - StudentBetai(0.5 * df!, 0.5, df! / (df! + t! * t!))
If studt! >= Alpha! Then
If increment! <= resolution! Then
If studt! - slast! <> 0# Then
t! = tlast! + (t! - tlast!) * (Alpha! - slast!) / (studt! - slast!)
End If
Exit Sub

Else
t! = tlast!
increment! = 0.1 * increment!
End If
End If
slast! = studt!
tlast! = t!
t! = t! + increment!
Next i&
  
' If we get here, not enough iterations
GoTo StudentGetTIterationError
Exit Sub

' Errors
StudentGetTError:
MsgBox Error$, vbOKOnly + vbCritical, "StudentGetT"
ierror = True
Exit Sub

StudentGetTIterationError:
msg$ = "Insufficient iterations to converge"
MsgBox msg$, vbOKOnly + vbExclamation, "StudentGetT"
ierror = True
Exit Sub

StudentGetTBadDf:
msg$ = "Parameter 'df' is out of rangle"
MsgBox msg$, vbOKOnly + vbExclamation, "StudentGetT"
ierror = True
Exit Sub

StudentGetTBadAlpha:
msg$ = "Parameter 'alpha' is out of rangle"
MsgBox msg$, vbOKOnly + vbExclamation, "StudentGetT"
ierror = True
Exit Sub

End Sub

