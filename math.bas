Attribute VB_Name = "CodeMATH"
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

Sub MathArrayAverage(average As TypeAverage, ArrayPassed() As Single, rows As Integer, cols As Integer, sample() As TypeSample)
' Column and row average based on sample(1).LineStatus flags (true = use, false = skip).
' This routine takes the data in the array passed and averages the data by columns,
' and also returns the standard deviations, the square root, and the standard errors
' about the average. The results are returned in "average.Averags", "average.Stddevs",
' "average.Sqroots", and "average.Stderrs".

ierror = False
On Error GoTo MathArrayAverageError

Dim temp As Single
Dim col As Integer, row As Integer

ReDim colsums(1 To MAXCHAN1%) As Single

' First zero the return arrays and the column sums
temp! = 0#
For col% = 1 To MAXCHAN%
average.averags!(col%) = 0#
average.Stddevs!(col%) = 0#
average.Sqroots!(col%) = 0#
average.Stderrs!(col%) = 0#
average.Reldevs!(col%) = 0#
average.Minimums!(col%) = 0#
average.Maximums!(col%) = 0#
colsums!(col%) = 0#
Next col%

' Check for valid data
If sample(1).Datarows% < 1 Then Exit Sub
If sample(1).GoodDataRows% < 1 Then Exit Sub

' Load default minimums and maximums
For col% = 1 To MAXCHAN1%
average.Minimums!(col%) = MAXMINIMUM!
average.Maximums!(col%) = MAXMAXIMUM!
Next col%

' Calculate sums of all valid columns
For row% = 1 To rows%
If sample(1).LineStatus(row%) Then
For col% = 1 To cols%
If ArrayPassed!(row%, col%) < average.Minimums!(col%) Then average.Minimums!(col%) = ArrayPassed!(row%, col%)
If ArrayPassed!(row%, col%) > average.Maximums!(col%) Then average.Maximums!(col%) = ArrayPassed!(row%, col%)
colsums!(col%) = colsums!(col%) + ArrayPassed!(row%, col%)
Next col%
End If
Next row%

' Divide to determine the averages
For col% = 1 To cols%
average.averags!(col%) = colsums!(col%) / sample(1).GoodDataRows%
Next col%

' Determine standard deviations, first sum the squares
For col% = 1 To cols%
colsums!(col%) = 0#
Next col%

For row% = 1 To rows%
If sample(1).LineStatus(row%) Then
For col% = 1 To cols%
temp! = ArrayPassed!(row%, col%) - average.averags!(col%)
colsums!(col%) = colsums!(col%) + temp! * temp!
Next col%
End If
Next row%

' Calculate square root, standard deviation and standard error
For col% = 1 To cols%
average.Sqroots!(col%) = Sqr(Abs(average.averags!(col%)))
If sample(1).GoodDataRows% > 1 Then
average.Stddevs!(col%) = Sqr(Abs(colsums!(col%)) / (sample(1).GoodDataRows% - 1))
average.Stderrs!(col%) = average.Stddevs!(col%) / Sqr(sample(1).GoodDataRows%)
If average.averags!(col%) <> 0# Then
average.Reldevs!(col%) = average.Stddevs!(col%) / average.averags!(col%)
End If
End If
Next col%

Exit Sub

' Errors
MathArrayAverageError:
MsgBox Error$, vbOKOnly + vbCritical, "MathArrayAverage"
ierror = True
Exit Sub

End Sub

Sub MathArrayAverage2(average As TypeAverage, ArrayPassed() As Single, ArrayPassed2() As Single, rows As Integer, cols As Integer, sample() As TypeSample)
' Column and row average based on sample(1).LineStatus flags (true = use, false = skip).
' This routine takes the data in the array passed and averages the data by columns,
' and also returns the standard deviations, the square root, and the standard errors
' about the average. The results are returned in "average.Averags", "average.Stddevs",
' "average.Sqroots", and "average.Stderrs".

' Multiplies by count times to obtain actual standard deviations (no need to correct for count time)

ierror = False
On Error GoTo MathArrayAverage2Error

Dim temp As Single
Dim col As Integer, row As Integer

ReDim colsums(1 To MAXCHAN1%) As Single

' First zero the return arrays and the column sums
temp! = 0#
For col% = 1 To MAXCHAN%
average.averags!(col%) = 0#
average.Stddevs!(col%) = 0#
average.Sqroots!(col%) = 0#
average.Stderrs!(col%) = 0#
average.Reldevs!(col%) = 0#
average.Minimums!(col%) = 0#
average.Maximums!(col%) = 0#
colsums!(col%) = 0#
Next col%

' Check for valid data
If sample(1).Datarows% < 1 Then Exit Sub
If sample(1).GoodDataRows% < 1 Then Exit Sub

' Load default minimums and maximums
For col% = 1 To MAXCHAN1%
average.Minimums!(col%) = MAXMINIMUM!
average.Maximums!(col%) = MAXMAXIMUM!
Next col%

' Calculate sums of all valid columns
For row% = 1 To rows%
If sample(1).LineStatus(row%) Then
For col% = 1 To cols%
If col% <> cols% Then
If ArrayPassed!(row%, col%) * ArrayPassed2!(row%, col%) < average.Minimums!(col%) Then average.Minimums!(col%) = ArrayPassed!(row%, col%) * ArrayPassed2!(row%, col%)
If ArrayPassed!(row%, col%) * ArrayPassed2!(row%, col%) > average.Maximums!(col%) Then average.Maximums!(col%) = ArrayPassed!(row%, col%) * ArrayPassed2!(row%, col%)
colsums!(col%) = colsums!(col%) + ArrayPassed!(row%, col%) * ArrayPassed2!(row%, col%)
Else
If ArrayPassed!(row%, col%) < average.Minimums!(col%) Then average.Minimums!(col%) = ArrayPassed!(row%, col%)
If ArrayPassed!(row%, col%) > average.Maximums!(col%) Then average.Maximums!(col%) = ArrayPassed!(row%, col%)
colsums!(col%) = colsums!(col%) + ArrayPassed!(row%, col%)
End If
Next col%
End If
Next row%

' Divide to determine the averages
For col% = 1 To cols%
average.averags!(col%) = colsums!(col%) / sample(1).GoodDataRows%
Next col%

' Determine standard deviations, first sum the squares
For col% = 1 To cols%
colsums!(col%) = 0#
Next col%

For row% = 1 To rows%
If sample(1).LineStatus(row%) Then
For col% = 1 To cols%
If col% <> cols% Then
temp! = ArrayPassed!(row%, col%) * ArrayPassed2!(row%, col%) - average.averags!(col%)
Else
temp! = ArrayPassed!(row%, col%) - average.averags!(col%)
End If
colsums!(col%) = colsums!(col%) + temp! * temp!
Next col%
End If
Next row%

' Calculate square root, standard deviation and standard error
For col% = 1 To cols%
average.Sqroots!(col%) = Sqr(Abs(average.averags!(col%)))
If sample(1).GoodDataRows% > 1 Then
average.Stddevs!(col%) = Sqr(Abs(colsums!(col%)) / (sample(1).GoodDataRows% - 1))
average.Stderrs!(col%) = average.Stddevs!(col%) / Sqr(sample(1).GoodDataRows%)
If average.averags!(col%) <> 0# Then
average.Reldevs!(col%) = average.Stddevs!(col%) / average.averags!(col%)
End If
End If
Next col%

Exit Sub

' Errors
MathArrayAverage2Error:
MsgBox Error$, vbOKOnly + vbCritical, "MathArrayAverage2"
ierror = True
Exit Sub

End Sub

Sub MathAverage(average As TypeAverage, ArrayPassed() As Single, rows As Integer, sample() As TypeSample)
' Single Column average (typically based on TypeSample(), 1 to sample(1).Datarows%).
' This routine takes the data in the array passed and averages a single column of data,
' and returns the standard deviations, the square root, and the standard errors
' about the average. The results are returned in "average.Averags", "average.Stddevs",
' "average.Sqroots", and "average.Stderrs".

ierror = False
On Error GoTo MathAverageError

Dim colsum As Single, temp As Single
Dim row As Integer

' First zero the return arrays and the column sums
temp! = 0#
average.averags!(1) = 0#
average.Stddevs!(1) = 0#
average.Sqroots!(1) = 0#
average.Stderrs!(1) = 0#
average.Reldevs!(1) = 0#
average.Minimums!(1) = 0#
average.Maximums!(1) = 0#

colsum! = 0#

' Check for valid data
If sample(1).Datarows% < 1 Then Exit Sub
If sample(1).GoodDataRows% < 1 Then Exit Sub

' Load default maxmimums and minimums
average.Minimums!(1) = MAXMINIMUM!
average.Maximums!(1) = MAXMAXIMUM!

' Calculate sum of all valid rows
For row% = 1 To rows%
If sample(1).LineStatus(row%) Then
If ArrayPassed!(row%) < average.Minimums!(1) Then average.Minimums!(1) = ArrayPassed!(row%)
If ArrayPassed!(row%) > average.Maximums!(1) Then average.Maximums!(1) = ArrayPassed!(row%)
colsum! = colsum! + ArrayPassed!(row%)
End If
Next row%

' Divide to determine the averages
average.averags!(1) = colsum! / sample(1).GoodDataRows%

' Determine standard deviations, first sum the squares
colsum! = 0#
For row% = 1 To rows%
If sample(1).LineStatus(row%) Then
temp! = ArrayPassed!(row%) - average.averags!(1)
colsum! = colsum! + temp! * temp!
End If
Next row%

' Calculate square root, standard deviation and standard error
average.Sqroots!(1) = Sqr(Abs(average.averags!(1)))
If sample(1).GoodDataRows% > 1 Then
average.Stddevs!(1) = Sqr(Abs(colsum!) / (sample(1).GoodDataRows% - 1))
average.Stderrs!(1) = average.Stddevs!(1) / Sqr(sample(1).GoodDataRows%)
If average.averags!(1) <> 0# Then
average.Reldevs!(1) = average.Stddevs!(1) / average.averags!(1)
End If
End If

Exit Sub

' Errors
MathAverageError:
MsgBox Error$, vbOKOnly + vbCritical, "MathAverage"
ierror = True
Exit Sub

End Sub

Sub MathAverage2(average As TypeAverage, ArrayPassed() As Single, cols As Integer, sample() As TypeSample)
' Single row average (typically based on TypeSample(), 1 to sample(1).LastElm% or 1 to sample(1).LastChan%)
' This routine takes the data in the array passed and averages a single row of data,
' and returns the standard deviations, the square root, and the standard errors
' about the average. The results are returned in "average.Averags", "average.Stddevs",
' "average.Sqroots", and "average.Stderrs".

ierror = False
On Error GoTo MathAverage2Error

Dim rowsum As Single, temp As Single
Dim col As Integer, ncols As Integer

' First zero the return arrays and the column sums
temp! = 0#
average.averags!(1) = 0#
average.Stddevs!(1) = 0#
average.Sqroots!(1) = 0#
average.Stderrs!(1) = 0#
average.Reldevs!(1) = 0#
average.Minimums!(1) = 0#
average.Maximums!(1) = 0#

rowsum! = 0#

' Check for valid data
If cols% < 1 Then Exit Sub

' Load default maxmimums and minimums
average.Minimums!(1) = MAXMINIMUM!
average.Maximums!(1) = MAXMAXIMUM!

' Calculate sum of all valid rows
ncols% = 0      ' sum number "good" columns of data
For col% = 1 To cols%
If sample(1).DisableQuantFlag(col%) <> 1 Then
If ArrayPassed!(col%) < average.Minimums!(1) Then average.Minimums!(1) = ArrayPassed!(col%)
If ArrayPassed!(col%) > average.Maximums!(1) Then average.Maximums!(1) = ArrayPassed!(col%)
rowsum! = rowsum! + ArrayPassed!(col%)
ncols% = ncols% + 1
End If
Next col%

' Divide to determine the averages
If ncols% < 1 Then Exit Sub
average.averags!(1) = rowsum! / ncols%

' Determine standard deviations, first sum the squares
rowsum! = 0#
For col% = 1 To cols%
If sample(1).DisableQuantFlag(col%) <> 1 Then
temp! = ArrayPassed!(col%) - average.averags!(1)
rowsum! = rowsum! + temp! * temp!
End If
Next col%

' Calculate square root, standard deviation and standard error
average.Sqroots!(1) = Sqr(Abs(average.averags!(1)))
If ncols% > 1 Then
average.Stddevs!(1) = Sqr(Abs(rowsum!) / (ncols% - 1))
average.Stderrs!(1) = average.Stddevs!(1) / Sqr(ncols%)
If average.averags!(1) <> 0# Then
average.Reldevs!(1) = average.Stddevs!(1) / average.averags!(1)
End If
End If

Exit Sub

' Errors
MathAverage2Error:
MsgBox Error$, vbOKOnly + vbCritical, "MathAverage2"
ierror = True
Exit Sub

End Sub

Sub MathCountAverage(average As TypeAverage, sample() As TypeSample)
' This routine takes the data in the sample array and averages the count data by columns,
' and also returns the standard deviations, the square root, and the standard errors
' about the average. The results are returned in "average.Averags", "average.Stddevs", "average.Sqroots",
' and "average.Stderrs". The average "sample(1).DateTimes" is returned in "average.AverDateTime".

ierror = False
On Error GoTo MathCountAverageError

Dim temp As Single
Dim col As Integer, row As Integer

ReDim colsums(1 To MAXCHAN%) As Single

' First zero the return arrays and the column sums
temp! = 0#
average.AverDateTime# = 0#
For col% = 1 To MAXCHAN%
average.averags!(col%) = 0#
average.Stddevs!(col%) = 0#
average.Sqroots!(col%) = 0#
average.Stderrs!(col%) = 0#
average.Reldevs!(col%) = 0#
average.Minimums!(col%) = 0#
average.Maximums!(col%) = 0#
colsums!(col%) = 0#
Next col%

' Check for valid data
If sample(1).Datarows% < 1 Then Exit Sub
If sample(1).GoodDataRows% < 1 Then Exit Sub

' Load default minimums and maximums
For col% = 1 To MAXCHAN%
average.Minimums!(col%) = MAXMINIMUM!
average.Maximums!(col%) = MAXMAXIMUM!
Next col%

' Calculate sums of all columns
For row% = 1 To sample(1).Datarows%
If sample(1).LineStatus(row%) Then
For col% = 1 To sample(1).LastElm%
If sample(1).CorData!(row%, col%) < average.Minimums!(col%) Then average.Minimums!(col%) = sample(1).CorData!(row%, col%)
If sample(1).CorData!(row%, col%) > average.Maximums!(col%) Then average.Maximums!(col%) = sample(1).CorData!(row%, col%)
colsums!(col%) = colsums!(col%) + sample(1).CorData!(row%, col%)
Next col%
average.AverDateTime# = average.AverDateTime# + sample(1).DateTimes(row%)
End If
Next row%

' Divide to determine the averages
For col% = 1 To sample(1).LastElm%
average.averags!(col%) = colsums!(col%) / sample(1).GoodDataRows%
Next col%
average.AverDateTime# = average.AverDateTime# / CDbl(sample(1).GoodDataRows%)

' Determine standard deviations, first sum the squares
For col% = 1 To sample(1).LastElm%
colsums!(col%) = 0#
Next col%

For row% = 1 To sample(1).Datarows%
If sample(1).LineStatus(row%) Then
For col% = 1 To sample(1).LastElm%
temp! = sample(1).CorData!(row%, col%) - average.averags!(col%)
colsums!(col%) = colsums!(col%) + temp! * temp!
Next col%
End If
Next row%

' Calculate square root, standard deviation and standard error
For col% = 1 To sample(1).LastElm%
average.Sqroots!(col%) = Sqr(Abs(average.averags!(col%)))
If sample(1).GoodDataRows% > 1 Then
average.Stddevs!(col%) = Sqr(Abs(colsums!(col%)) / (sample(1).GoodDataRows% - 1))
average.Stderrs!(col%) = average.Stddevs!(col%) / Sqr(sample(1).GoodDataRows%)
If average.averags!(col%) <> 0# Then
average.Reldevs!(col%) = average.Stddevs!(col%) / average.averags!(col%)
End If
End If
Next col%

Exit Sub

' Errors
MathCountAverageError:
MsgBox Error$, vbOKOnly + vbCritical, "MathCountAverage"
ierror = True
Exit Sub

End Sub

Sub MathSimpleAverage(average As TypeAverage, ArrayPassed!(), rows As Integer)
' Single Column average (not based on TypeSample!)
' This routine takes the data in the array passed and averages the single column of data,
' and returns the standard deviations, the square root, and the standard errors
' about the average. The results are returned in "average.Averags", "average.Stddevs",
' "average.Sqroots", and "average.Stderrs".

ierror = False
On Error GoTo MathSimpleAverageError

Dim colsum As Single, temp As Single
Dim row As Integer

' First zero the return arrays and the column sums
temp! = 0#
average.averags!(1) = 0#
average.Stddevs!(1) = 0#
average.Sqroots!(1) = 0#
average.Stderrs!(1) = 0#
average.Reldevs!(1) = 0#
average.Minimums!(1) = 0#
average.Maximums!(1) = 0#

colsum! = 0#

' Check for valid data
If rows% < 1 Then Exit Sub

' Load default minimums and maximums
average.Minimums!(1) = MAXMINIMUM!
average.Maximums!(1) = MAXMAXIMUM!

' Calculate sum of all valid rows
For row% = 1 To rows%
If ArrayPassed!(row%) < average.Minimums!(1) Then average.Minimums!(1) = ArrayPassed!(row%)
If ArrayPassed!(row%) > average.Maximums!(1) Then average.Maximums!(1) = ArrayPassed!(row%)
colsum! = colsum! + ArrayPassed!(row%)
Next row%

' Divide to determine the averages
average.averags!(1) = colsum! / rows%

' Determine standard deviations, first sum the squares
colsum! = 0#
For row% = 1 To rows%
temp! = ArrayPassed!(row%) - average.averags!(1)
colsum! = colsum! + temp! * temp!
Next row%

' Calculate square root, standard deviation and standard error
average.Sqroots!(1) = Sqr(Abs(average.averags!(1)))
If rows% > 1 Then
average.Stddevs!(1) = Sqr(Abs(colsum!) / (rows% - 1))
average.Stderrs!(1) = average.Stddevs!(1) / Sqr(rows%)
If average.averags!(1) <> 0# Then
average.Reldevs!(1) = average.Stddevs!(1) / average.averags!(1)
End If
End If

Exit Sub

' Errors
MathSimpleAverageError:
MsgBox Error$, vbOKOnly + vbCritical, "MathSimpleAverage"
ierror = True
Exit Sub

End Sub

Sub MathSimpleAverage2(average As TypeAverage, ArrayPassed!(), col As Integer, rows As Integer)
' Single Column average (not based on TypeSample!) of a single column in a 2 dimensional array
' This routine takes the data in the array passed and averages the single column of data,
' and returns the standard deviations, the square root, and the standard errors
' about the average. The results are returned in "average.Averags", "average.Stddevs",
' "average.Sqroots", and "average.Stderrs".
'
' ArrayPassed(col, row) = 2d array passed
' col = the column to be averaged
' rows = the number of rows to be averaged

ierror = False
On Error GoTo MathSimpleAverage2Error

Dim colsum As Single, temp As Single
Dim row As Integer

' First zero the return arrays and the column sums
temp! = 0#
average.averags!(1) = 0#
average.Stddevs!(1) = 0#
average.Sqroots!(1) = 0#
average.Stderrs!(1) = 0#
average.Reldevs!(1) = 0#
average.Minimums!(1) = 0#
average.Maximums!(1) = 0#

colsum! = 0#

' Check for valid data
If rows% < 1 Then Exit Sub

' Load default minimums and maximums
average.Minimums!(1) = MAXMINIMUM!
average.Maximums!(1) = MAXMAXIMUM!

' Calculate sum of all valid rows
For row% = 1 To rows%
If ArrayPassed!(col%, row%) < average.Minimums!(1) Then average.Minimums!(1) = ArrayPassed!(col%, row%)
If ArrayPassed!(col%, row%) > average.Maximums!(1) Then average.Maximums!(1) = ArrayPassed!(col%, row%)
colsum! = colsum! + ArrayPassed!(col%, row%)
Next row%

' Divide to determine the averages
average.averags!(1) = colsum! / rows%

' Determine standard deviations, first sum the squares
colsum! = 0#
For row% = 1 To rows%
temp! = ArrayPassed!(col%, row%) - average.averags!(1)
colsum! = colsum! + temp! * temp!
Next row%

' Calculate square root, standard deviation and standard error
average.Sqroots!(1) = Sqr(Abs(average.averags!(1)))
If rows% > 1 Then
average.Stddevs!(1) = Sqr(Abs(colsum!) / (rows% - 1))
average.Stderrs!(1) = average.Stddevs!(1) / Sqr(rows%)
If average.averags!(1) <> 0# Then
average.Reldevs!(1) = average.Stddevs!(1) / average.averags!(1)
End If
End If

Exit Sub

' Errors
MathSimpleAverage2Error:
MsgBox Error$, vbOKOnly + vbCritical, "MathSimpleAverage2"
ierror = True
Exit Sub

End Sub

Sub MathSimpleAverage22(average As TypeAverage, ArrayPassed!(), cols As Integer, row As Integer)
' Single Column average (not based on TypeSample!) of a single column in a 2 dimensional array
' This routine takes the data in the array passed and averages the single row of data,
' and returns the standard deviations, the square root, and the standard errors
' about the average. The results are returned in "average.Averags", "average.Stddevs",
' "average.Sqroots", and "average.Stderrs".
'
' ArrayPassed(col, row) = 2d array passed
' col = the columns to be averaged
' rows = the row to be averaged

ierror = False
On Error GoTo MathSimpleAverage22Error

Dim rowsum As Single, temp As Single
Dim col As Integer

' First zero the return arrays and the column sums
temp! = 0#
average.averags!(1) = 0#
average.Stddevs!(1) = 0#
average.Sqroots!(1) = 0#
average.Stderrs!(1) = 0#
average.Reldevs!(1) = 0#
average.Minimums!(1) = 0#
average.Maximums!(1) = 0#

rowsum! = 0#

' Check for valid data
If cols% < 1 Then Exit Sub

' Load default minimums and maximums
average.Minimums!(1) = MAXMINIMUM!
average.Maximums!(1) = MAXMAXIMUM!

' Calculate sum of all valid rows
For col% = 1 To cols%
If ArrayPassed!(col%, row%) < average.Minimums!(1) Then average.Minimums!(1) = ArrayPassed!(col%, row%)
If ArrayPassed!(col%, row%) > average.Maximums!(1) Then average.Maximums!(1) = ArrayPassed!(col%, row%)
rowsum! = rowsum! + ArrayPassed!(col%, row%)
Next col%

' Divide to determine the averages
average.averags!(1) = rowsum! / cols%

' Determine standard deviations, first sum the squares
rowsum! = 0#
For col% = 1 To cols%
temp! = ArrayPassed!(col%, row%) - average.averags!(1)
rowsum! = rowsum! + temp! * temp!
Next col%

' Calculate square root, standard deviation and standard error
average.Sqroots!(1) = Sqr(Abs(average.averags!(1)))
If cols% > 1 Then
average.Stddevs!(1) = Sqr(Abs(rowsum!) / (cols% - 1))
average.Stderrs!(1) = average.Stddevs!(1) / Sqr(cols%)
If average.averags!(1) <> 0# Then
average.Reldevs!(1) = average.Stddevs!(1) / average.averags!(1)
End If
End If

Exit Sub

' Errors
MathSimpleAverage22Error:
MsgBox Error$, vbOKOnly + vbCritical, "MathSimpleAverage22"
ierror = True
Exit Sub

End Sub

Sub MathArrayAverage3(average As TypeAverage, ArrayPassed() As Single, rows As Long, cols As Integer)
' Column and row average (not based on TypeSample). This routine takes the data in the
' array passed and averages the data by columns, and also returns the standard deviations,
' the square root, and the standard errors about the average. The results are returned in
' "average.Averags", "average.Stddevs", "average.Sqroots", and "average.Stderrs".

ierror = False
On Error GoTo MathArrayAverage3Error

Dim temp As Single
Dim col As Integer, row As Long

ReDim colsums(1 To cols%) As Single

' Check for too few rows
If rows& < 1 Then GoTo MathArrayAverage3TooFewRows

' Just load passed values if only one row
If rows& = 1 Then
For col% = 1 To cols%
average.averags!(col%) = ArrayPassed!(col%, 1)
average.Stddevs!(col%) = 0#
average.Sqroots!(col%) = 0#
average.Stderrs!(col%) = 0#
average.Reldevs!(col%) = 0#
average.Minimums!(col%) = ArrayPassed!(col%, 1)
average.Maximums!(col%) = ArrayPassed!(col%, 1)
Next col%
Exit Sub
End If

' First zero the return arrays and the column sums
temp! = 0#
For col% = 1 To cols%
average.averags!(col%) = 0#
average.Stddevs!(col%) = 0#
average.Sqroots!(col%) = 0#
average.Stderrs!(col%) = 0#
average.Reldevs!(col%) = 0#
average.Minimums!(col%) = 0#
average.Maximums!(col%) = 0#
colsums!(col%) = 0#
Next col%

' Load default minimums and maximums
For col% = 1 To cols%
average.Minimums!(col%) = MAXMINIMUM!
average.Maximums!(col%) = MAXMAXIMUM!
Next col%

' Calculate sums of all valid columns
For row& = 1 To rows&
For col% = 1 To cols%
If ArrayPassed!(col%, row&) < average.Minimums!(col%) Then average.Minimums!(col%) = ArrayPassed!(col%, row&)
If ArrayPassed!(col%, row&) > average.Maximums!(col%) Then average.Maximums!(col%) = ArrayPassed!(col%, row&)
colsums!(col%) = colsums!(col%) + ArrayPassed!(col%, row&)
Next col%
Next row&

' Divide to determine the averages
For col% = 1 To cols%
average.averags!(col%) = colsums!(col%) / rows&
Next col%

' Determine standard deviations, first sum the squares
For col% = 1 To cols%
colsums!(col%) = 0#
Next col%

For row& = 1 To rows&
For col% = 1 To cols%
temp! = ArrayPassed!(col%, row&) - average.averags!(col%)
colsums!(col%) = colsums!(col%) + temp! * temp!
Next col%
Next row&

' Calculate square root, standard deviation and standard error
For col% = 1 To cols%
average.Sqroots!(col%) = Sqr(Abs(average.averags!(col%)))
average.Stddevs!(col%) = Sqr(Abs(colsums!(col%)) / (rows& - 1))
average.Stderrs!(col%) = average.Stddevs!(col%) / Sqr(rows&)
If average.averags!(col%) <> 0# Then
average.Reldevs!(col%) = average.Stddevs!(col%) / average.averags!(col%)
End If
Next col%

Exit Sub

' Errors
MathArrayAverage3Error:
MsgBox Error$, vbOKOnly + vbCritical, "MathArrayAverage3"
ierror = True
Exit Sub

MathArrayAverage3TooFewRows:
msg$ = "Too few data rows to average"
MsgBox msg$, vbOKOnly + vbExclamation, "MathArrayAverage3"
ierror = True
Exit Sub

End Sub

Sub MathArrayAverage4(average As TypeAverage, ArrayPassed() As Single, rows As Long, cols As Long)
' Column and row average (not based on TypeSample). This routine takes the data in the
' array passed and averages the data by columns, and also returns the standard deviations,
' the square root, and the standard errors about the average. The results are returned in
' "average.Averags", "average.Stddevs", "average.Sqroots", and "average.Stderrs".

ierror = False
On Error GoTo MathArrayAverage4Error

Dim temp As Single
Dim col As Long, row As Long

ReDim colsums(1 To cols&) As Single

' Check for too few rows
If rows& < 1 Then GoTo MathArrayAverage4TooFewRows

' Just load passed values if only one row
If rows& = 1 Then
For col& = 1 To cols&
average.averags!(col&) = ArrayPassed!(col&, 1)
average.Stddevs!(col&) = 0#
average.Sqroots!(col&) = 0#
average.Stderrs!(col&) = 0#
average.Reldevs!(col&) = 0#
average.Minimums!(col&) = ArrayPassed!(col&, 1)
average.Maximums!(col&) = ArrayPassed!(col&, 1)
Next col&
Exit Sub
End If

' First zero the return arrays and the column sums
temp! = 0#
For col& = 1 To cols&
average.averags!(col&) = 0#
average.Stddevs!(col&) = 0#
average.Sqroots!(col&) = 0#
average.Stderrs!(col&) = 0#
average.Reldevs!(col&) = 0#
average.Minimums!(col&) = 0#
average.Maximums!(col&) = 0#
colsums!(col&) = 0#
Next col&

' Load default minimums and maximums
For col& = 1 To cols&
average.Minimums!(col&) = MAXMINIMUM!
average.Maximums!(col&) = MAXMAXIMUM!
Next col&

' Calculate sums of all valid columns
For row& = 1 To rows&
For col& = 1 To cols&
If ArrayPassed!(col&, row&) < average.Minimums!(col&) Then average.Minimums!(col&) = ArrayPassed!(col&, row&)
If ArrayPassed!(col&, row&) > average.Maximums!(col&) Then average.Maximums!(col&) = ArrayPassed!(col&, row&)
colsums!(col&) = colsums!(col&) + ArrayPassed!(col&, row&)
Next col&
Next row&

' Divide to determine the averages
For col& = 1 To cols&
average.averags!(col&) = colsums!(col&) / rows&
Next col&

' Determine standard deviations, first sum the squares
For col& = 1 To cols&
colsums!(col&) = 0#
Next col&

For row& = 1 To rows&
For col& = 1 To cols&
temp! = ArrayPassed!(col&, row&) - average.averags!(col&)
colsums!(col&) = colsums!(col&) + temp! * temp!
Next col&
Next row&

' Calculate square root, standard deviation and standard error
For col& = 1 To cols&
average.Sqroots!(col&) = Sqr(Abs(average.averags!(col&)))
average.Stddevs!(col&) = Sqr(Abs(colsums!(col&)) / (rows& - 1))
average.Stderrs!(col&) = average.Stddevs!(col&) / Sqr(rows&)
If average.averags!(col&) <> 0# Then
average.Reldevs!(col&) = average.Stddevs!(col&) / average.averags!(col&)
End If
Next col&

Exit Sub

' Errors
MathArrayAverage4Error:
MsgBox Error$, vbOKOnly + vbCritical, "MathArrayAverage4"
ierror = True
Exit Sub

MathArrayAverage4TooFewRows:
Screen.MousePointer = vbDefault
msg$ = "Too few data rows to average"
MsgBox msg$, vbOKOnly + vbExclamation, "MathArrayAverage4"
ierror = True
Exit Sub

End Sub

Function MathNormalDistribution() As Double
' Returns a randomly distributed drawing from a standard normal distribution
'  i.e. one with Average = 0 and Standard Deviation = 1.0
    
ierror = False
On Error GoTo MathNormalDistibutionError
  
Dim fac As Double, rsq As Double
Dim v1 As Double, v2 As Double

Static flag As Boolean, gset As Double
    
' Each pass through the calculation of the routine produces
'  two normally-distributed deviates, so we only need to do
'  the calculations every other call. So we set the flag
'  variable (to true) if gset contains a spare NormRand value.
If flag Then
    MathNormalDistribution# = gset

' Force calculation next time
    flag = False
    
' Don't have anything saved so need to find a pair of values
' First generate a co-ordinate pair within the unit circle
Else
    Do
        v1 = 2 * Rnd - 1#
        v2 = 2 * Rnd - 1#
        rsq = v1 * v1 + v2 * v2
    Loop Until rsq <= 1#
    fac = Sqr(-2# * Log(rsq) / rsq) ' do the math
        
' Return one of the values and save the other (gset) for next time
    MathNormalDistribution# = v2 * fac
    gset = v1 * fac
    flag = True
End If
    
Exit Function
    
MathNormalDistibutionError:
MsgBox Error$, vbOKOnly + vbCritical, "MathNormalDistibution"
ierror = True
Exit Function

End Function

Function MathNormalDistribution2(m As Double, s As Double) As Double
' Returns a randomly distributed number with a mean "m" and standard deviation "s"
'  m = Mean (input)
'  s = Standard Deviation (input)

ierror = False
On Error GoTo MathNormalDistibution2Error
  
Dim r1 As Double, r2 As Double

Do Until r1# <> 0 And r2# <> 0
r1# = Rnd()
r2# = Rnd()
MathNormalDistribution2# = s# * Sqr(-2 * Log(r1#)) * Cos(2 * PI! * r2#) + m#
Loop
    
Exit Function
    
MathNormalDistibution2Error:
MsgBox Error$, vbOKOnly + vbCritical, "MathNormalDistibution2"
ierror = True
Exit Function

End Function

Sub MathCorrelationPearson(X() As Double, Y() As Double, n As Long, r As Double, prob As Double, Z As Double)
' Calculate Pearson's linear correlation coefficient
'  Given two arrays "x()" and "y()", this routine computes their correlation coefficient "r", the significance
'  level at which the null hypothesis of zero correlation is disproved ("prob" whose small value indicates
'  a significant correlation), and Fisher's "z", whose value can be used in further statistical tests.
'
'  This procedure will regularize the unusual case of complete correlation.

ierror = False
On Error GoTo MathCorrelationPearsonError

Const TINY1# = 1E-20
      
Dim j As Long
Dim ax As Double, ay As Double
Dim df As Double, sxx As Double, sxy As Double, syy As Double
Dim t As Double, xt As Double, yt As Double

ax# = 0#
ay# = 0#
For j& = 1 To n&
ax# = ax# + X#(j&)
        ay# = ay# + Y#(j&)
Next j&

ax# = ax# / n&
ay# = ay# / n&
sxx# = 0#
syy# = 0#
sxy# = 0#
      
For j& = 1 To n&
xt# = X#(j&) - ax#
        yt# = Y#(j&) - ay#
        sxx# = sxx# + xt# ^ 2
        syy# = syy# + yt# ^ 2
        sxy# = sxy# + xt# * yt#
Next j&

r# = sxy# / (Sqr(sxx# * syy#) + TINY1#)
Z# = 0.5 * Log(((1# + r#) + TINY1#) / ((1# - r#) + TINY1#))
df# = n& - 2
t# = r# * Sqr(df# / (((1# - r#) + TINY1#) * ((1# + r#) + TINY1#)))
prob# = StudentBetai(0.5 * df#, 0.5, df# / (df# + t# ^ 2))
      
' Alternative calculation for very large data sets
'prob# = MathERFCC(Abs(z# * Sqr(n& - 1#)) / 1.4142136)

Exit Sub

' Errors
MathCorrelationPearsonError:
MsgBox Error$, vbOKOnly + vbCritical, "MathCorrelationPearson"
ierror = True
Exit Sub

End Sub

Function MathERFCC(X As Double) As Double
' Calculates error function

ierror = False
On Error GoTo MathERFCCError

Dim t As Double, Z As Double

Z# = Abs(X#)
t# = 1# / (1# + 0.5 * Z#)

MathERFCC# = t# * Exp(-Z# * Z# - 1.26551223 + t# * (1.00002368 + t# * (0.37409196 + t# * _
     (0.09678418 + t# * (-0.18628806 + t# * (0.27886807 + t# * (-1.13520398 + t# * _
     (1.48851587 + t# * (-0.82215223 + t# * 0.17087277)))))))))
      
If X# < 0# Then MathERFCC# = 2# - MathERFCC#
Exit Function
    
MathERFCCError:
MsgBox Error$, vbOKOnly + vbCritical, "MathERFCC"
ierror = True
Exit Function

End Function

Sub MathAverageWeighted(average As Single, tarray() As Single, ncols As Integer, linerow As Integer, analysisarray() As Single, sample() As TypeSample)
' Calculate the weighted average of a (linerow, chan) array
' average = weighted average value to be calculated and returned
' tarray() = the data to be weight averaged
' ncols = the number of columns in the passed array
' linerow = the data row of the analysis array
' analysisarray() = the weighting data in a (linerow, chan) array
' sample() = the sample

ierror = False
On Error GoTo MathAverageWeightedError

Dim col As Integer
Dim sum As Single
Dim temp As Single

' Calculate total of analysis array to obtain weighting factors
sum! = 0#
For col% = 1 To ncols%
If sample(1).DisableQuantFlag%(col%) <> 1 Then
sum! = sum! + analysisarray!(linerow%, col%)
End If
Next col%

' Calculate weighting
temp! = 0#
For col% = 1 To ncols%
If sample(1).DisableQuantFlag%(col%) <> 1 Then
If sum! <> 0# Then temp! = temp! + tarray!(col%) * analysisarray!(linerow%, col%) / sum!
End If
Next col%

average! = temp!
Exit Sub

' Errors
MathAverageWeightedError:
MsgBox Error$, vbOKOnly + vbCritical, "MathAverageWeighted"
ierror = True
Exit Sub

End Sub

Sub MathAverageWeighted2(average As Single, tarray() As Single, ncols As Integer, analysisarray() As Single, sample() As TypeSample)
' Calculate the weighted average of an (chan) array
' average = weighted average value to be calculated and returned
' tarray() = the data to be weight averaged
' ncols = the number of columns in the passed array
' analysisarray() = the weighting data in a (chan) array
' sample() = the sample

ierror = False
On Error GoTo MathAverageWeighted2Error

Dim col As Integer
Dim sum As Single
Dim temp As Single

' Calculate total of analysis array to obtain weighting factors
sum! = 0#
For col% = 1 To ncols%
If sample(1).DisableQuantFlag%(col%) <> 1 Then
sum! = sum! + analysisarray!(col%)
End If
Next col%

' Calculate weighting
temp! = 0#
For col% = 1 To ncols%
If sample(1).DisableQuantFlag%(col%) <> 1 Then
If sum! <> 0# Then temp! = temp! + tarray!(col%) * analysisarray!(col%) / sum!
End If
Next col%

average! = temp!
Exit Sub

' Errors
MathAverageWeighted2Error:
MsgBox Error$, vbOKOnly + vbCritical, "MathAverageWeighted2"
ierror = True
Exit Sub

End Sub

Function MathRootN(X As Single, n As Integer) As Single
' Function to find the nth root of a number x

ierror = False
On Error GoTo MathRootNError

Dim temp As Single

' Find the nth root of x
If n% < 1 Then GoTo MathRootNBadRoot
temp! = X! ^ (1# / n%)
MathRootN! = temp!
Exit Function

' Errors
MathRootNError:
MsgBox Error$, vbOKOnly + vbCritical, "MathRootN"
ierror = True
Exit Function

MathRootNBadRoot:
msg$ = "Invalid root for calculation- must be greater than or equal to 1"
MsgBox msg$, vbOKOnly + vbExclamation, "MathRootN"
ierror = True
Exit Function

End Function

Function MathDVal(s As String) As Double
' Function to return proper double if "E" is missing from scientific notation

ierror = False
If ierror Then GoTo MathDValError

Dim es(1 To 2) As String
Dim mantissa As Double
Dim exponent As Integer

' Strip blanks on ends
s$ = Trim$(s$)

' See if exponent is available
If InStr(s$, "E") > 0 Then
MathDVal# = Val(s$)
Exit Function
End If

' Must be weird FORTRAN "1.0286-101", etc. So if exponent found, then extract mantissa and exponent using + or - signs for exponent
If InStr(s$, "-") > 1 Then
es$(1) = Left$(s$, InStr(2, s$, "-") - 1)
es$(2) = Mid$(s$, InStr(2, s$, "-"))

ElseIf InStr(s$, "+") > 1 Then
es$(1) = Left$(s$, InStr(2, s$, "+") - 1)
es$(2) = Mid$(s$, InStr(2, s$, "+"))

' Just regular positive or negative number
Else
MathDVal# = Val(s$)
Exit Function
End If

mantissa# = Val(es$(1))
exponent% = Val(es$(2))

MathDVal# = mantissa# * (10 ^ exponent%)
Exit Function

' Errors
MathDValError:
MsgBox Error$, vbOKOnly + vbCritical, "MathDVal"
ierror = True
Exit Function

End Function

Sub MathRandomizeArray(tarray() As Long)
' Randomize the passed array values

ierror = False
On Error GoTo MathRandomizeArrayError

Dim min_item As Long
Dim max_item As Long
Dim i As Long
Dim j As Long
Dim tmp_value As Long

min_item = LBound(tarray&)
max_item = UBound(tarray&)
    
For i& = min_item& To max_item& - 1
        
' Randomly assign item number i
j& = Int((max_item& - i& + 1) * Rnd + i&)
tmp_value& = tarray&(i&)
tarray&(i&) = tarray&(j&)
tarray&(j&) = tmp_value&

Next i&

Exit Sub

' Errors
MathRandomizeArrayError:
MsgBox Error$, vbOKOnly + vbCritical, "MathRandomizeArray"
ierror = True
Exit Sub

End Sub

Function MathArrayMaxLong(narray() As Long) As Long
' Return maximum array value for a long array

ierror = False
On Error GoTo MathArrayMaxLongError

Dim i As Long, maxp As Long

maxp& = MINLONG&
For i& = LBound(narray&) To UBound(narray&)
If narray&(i&) > maxp& Then maxp& = narray&(i&)
Next i&

MathArrayMaxLong& = maxp&
Exit Function

' Errors
MathArrayMaxLongError:
MsgBox Error$, vbOKOnly + vbCritical, "MathArrayMaxLong"
ierror = True
Exit Function

End Function

Function MathArrayMaxSingle(sarray() As Single) As Single
' Return maximum array value for a single array

ierror = False
On Error GoTo MathArrayMaxSingleError

Dim i As Long, maxp As Single

maxp! = MINSINGLE!
For i& = LBound(sarray!) To UBound(sarray!)
If sarray!(i&) > maxp! Then maxp! = sarray!(i&)
Next i&

MathArrayMaxSingle! = maxp!
Exit Function

' Errors
MathArrayMaxSingleError:
MsgBox Error$, vbOKOnly + vbCritical, "MathArrayMaxSingle"
ierror = True
Exit Function

End Function

Function MathArrayMaxDouble(darray() As Double) As Double
' Return maximum array value for a double array

ierror = False
On Error GoTo MathArrayMaxDoubleError

Dim i As Long, maxp As Double

maxp# = MINDOUBLE#
For i& = LBound(darray#) To UBound(darray#)
If darray#(i&) > maxp# Then maxp# = darray#(i&)
Next i&

MathArrayMaxDouble# = maxp#
Exit Function

' Errors
MathArrayMaxDoubleError:
MsgBox Error$, vbOKOnly + vbCritical, "MathArrayMaxDouble"
ierror = True
Exit Function

End Function

Sub MathArrayAverageSingle(average As TypeAverageMathSingle, ArrayPassed() As Single, rows As Integer, cols As Integer)
' Column and row average (not based on TypeSample). This routine takes the data in the
' array passed and averages the data by columns, and also returns the standard deviations,
' the square root, and the standard errors about the average. The results are returned in
' "average.Averags", "average.Stddevs", "average.Sqroots", and "average.Stderrs".

ierror = False
On Error GoTo MathArrayAverageSingleError

Dim temp As Single
Dim col As Integer, row As Integer

ReDim colsums(1 To cols%) As Single

ReDim average.averags(1 To cols%) As Single
ReDim average.Stddevs(1 To cols%) As Single
ReDim average.Sqroots(1 To cols%) As Single
ReDim average.Stderrs(1 To cols%) As Single
ReDim average.Reldevs(1 To cols%) As Single
ReDim average.Minimums(1 To cols%) As Single
ReDim average.Maximums(1 To cols%) As Single

' Check for too few rows
If rows% < 1 Then GoTo MathArrayAverageSingleTooFewRows

' Just load passed values if only one row
If rows% = 1 Then
For col% = 1 To cols%
average.averags!(col%) = ArrayPassed!(col%, 1)
average.Stddevs!(col%) = Sqr(Abs(ArrayPassed!(col%, 1)))
average.Sqroots!(col%) = 0#
average.Stderrs!(col%) = 0#
average.Reldevs!(col%) = 0#
average.Minimums!(col%) = ArrayPassed!(col%, 1)
average.Maximums!(col%) = ArrayPassed!(col%, 1)
Next col%
Exit Sub
End If

' First zero the return arrays and the column sums
temp! = 0#
For col% = 1 To cols%
average.averags!(col%) = 0#
average.Stddevs!(col%) = 0#
average.Sqroots!(col%) = 0#
average.Stderrs!(col%) = 0#
average.Reldevs!(col%) = 0#
average.Minimums!(col%) = 0#
average.Maximums!(col%) = 0#
colsums!(col%) = 0#
Next col%

' Load default minimums and maximums
For col% = 1 To cols%
average.Minimums!(col%) = MAXMINIMUM!
average.Maximums!(col%) = MAXMAXIMUM!
Next col%

' Calculate sums of all valid columns
For row% = 1 To rows%
For col% = 1 To cols%
If ArrayPassed!(col%, row%) < average.Minimums!(col%) Then average.Minimums!(col%) = ArrayPassed!(col%, row%)
If ArrayPassed!(col%, row%) > average.Maximums!(col%) Then average.Maximums!(col%) = ArrayPassed!(col%, row%)
colsums!(col%) = colsums!(col%) + ArrayPassed!(col%, row%)
Next col%
Next row%

' Divide to determine the averages
For col% = 1 To cols%
average.averags!(col%) = colsums!(col%) / rows%
Next col%

' Determine standard deviations, first sum the squares
For col% = 1 To cols%
colsums!(col%) = 0#
Next col%

For row% = 1 To rows%
For col% = 1 To cols%
temp! = ArrayPassed!(col%, row%) - average.averags!(col%)
colsums!(col%) = colsums!(col%) + temp! * temp!
Next col%
Next row%

' Calculate square root, standard deviation and standard error
For col% = 1 To cols%
average.Sqroots!(col%) = Sqr(Abs(average.averags!(col%)))
average.Stddevs!(col%) = Sqr(Abs(colsums!(col%)) / (rows% - 1))
average.Stderrs!(col%) = average.Stddevs!(col%) / Sqr(rows%)
If average.averags!(col%) <> 0# Then
average.Reldevs!(col%) = average.Stddevs!(col%) / average.averags!(col%)
End If
Next col%

Exit Sub

' Errors
MathArrayAverageSingleError:
MsgBox Error$, vbOKOnly + vbCritical, " MathArrayAverageSingle"
ierror = True
Exit Sub

MathArrayAverageSingleTooFewRows:
msg$ = "Too few data rows to average"
MsgBox msg$, vbOKOnly + vbExclamation, " MathArrayAverageSingle"
ierror = True
Exit Sub

End Sub

Sub MathArrayAverageDouble(average As TypeAverageMathDouble, ArrayPassed() As Double, rows As Integer, cols As Integer)
' Column and row average (not based on TypeSample). This routine takes the data in the
' array passed and averages the data by columns, and also returns the standard deviations,
' the square root, and the standard errors about the average. The results are returned in
' "average.Averags", "average.Stddevs", "average.Sqroots", and "average.Stderrs".

ierror = False
On Error GoTo MathArrayAverageDoubleError

Dim temp As Double
Dim col As Integer, row As Integer

ReDim colsums(1 To cols%) As Double

ReDim average.averags(1 To cols%) As Double
ReDim average.Stddevs(1 To cols%) As Double
ReDim average.Sqroots(1 To cols%) As Double
ReDim average.Stderrs(1 To cols%) As Double
ReDim average.Reldevs(1 To cols%) As Double
ReDim average.Minimums(1 To cols%) As Double
ReDim average.Maximums(1 To cols%) As Double

' Check for too few rows
If rows% < 1 Then GoTo MathArrayAverageDoubleTooFewRows

' Just load passed values if only one row
If rows% = 1 Then
For col% = 1 To cols%
average.averags#(col%) = ArrayPassed#(col%, 1)
average.Stddevs#(col%) = Sqr(Abs(ArrayPassed#(col%, 1)))
average.Sqroots#(col%) = 0#
average.Stderrs#(col%) = 0#
average.Reldevs#(col%) = 0#
average.Minimums#(col%) = ArrayPassed#(col%, 1)
average.Maximums#(col%) = ArrayPassed#(col%, 1)
Next col%
Exit Sub
End If

' First zero the return arrays and the column sums
temp# = 0#
For col% = 1 To cols%
average.averags#(col%) = 0#
average.Stddevs#(col%) = 0#
average.Sqroots#(col%) = 0#
average.Stderrs#(col%) = 0#
average.Reldevs#(col%) = 0#
average.Minimums#(col%) = 0#
average.Maximums#(col%) = 0#
colsums#(col%) = 0#
Next col%

' Load default minimums and maximums
For col% = 1 To cols%
average.Minimums#(col%) = MAXMINIMUM!
average.Maximums#(col%) = MAXMAXIMUM!
Next col%

' Calculate sums of all valid columns
For row% = 1 To rows%
For col% = 1 To cols%
If ArrayPassed#(col%, row%) < average.Minimums#(col%) Then average.Minimums#(col%) = ArrayPassed#(col%, row%)
If ArrayPassed#(col%, row%) > average.Maximums#(col%) Then average.Maximums#(col%) = ArrayPassed#(col%, row%)
colsums#(col%) = colsums#(col%) + ArrayPassed#(col%, row%)
Next col%
Next row%

' Divide to determine the averages
For col% = 1 To cols%
average.averags#(col%) = colsums#(col%) / rows%
Next col%

' Determine standard deviations, first sum the squares
For col% = 1 To cols%
colsums#(col%) = 0#
Next col%

For row% = 1 To rows%
For col% = 1 To cols%
temp# = ArrayPassed#(col%, row%) - average.averags#(col%)
colsums#(col%) = colsums#(col%) + temp# * temp#
Next col%
Next row%

' Calculate square root, standard deviation and standard error
For col% = 1 To cols%
average.Sqroots#(col%) = Sqr(Abs(average.averags#(col%)))
average.Stddevs#(col%) = Sqr(Abs(colsums#(col%)) / (rows% - 1))
average.Stderrs#(col%) = average.Stddevs#(col%) / Sqr(rows%)
If average.averags#(col%) <> 0# Then
average.Reldevs#(col%) = average.Stddevs#(col%) / average.averags#(col%)
End If
Next col%

Exit Sub

' Errors
MathArrayAverageDoubleError:
MsgBox Error$, vbOKOnly + vbCritical, " MathArrayAverageDouble"
ierror = True
Exit Sub

MathArrayAverageDoubleTooFewRows:
msg$ = "Too few data rows to average"
MsgBox msg$, vbOKOnly + vbExclamation, " MathArrayAverageDouble"
ierror = True
Exit Sub

End Sub

Function MathGetInterpolatedYValue(xpos As Single, nPoints As Integer, xdata() As Single, ydata() As Single) As Single
' Determines the y data value based on interpolating between nearest x data values

ierror = False
On Error GoTo MathGetInterpolatedYValueError

Dim i As Integer, defminpnt As Integer, defmaxpnt As Integer
Dim xmindif As Single, xmaxdif As Single
Dim xminpnt As Integer, xmaxpnt As Integer

Dim xmin As Single, xmax As Single
Dim ymin As Single, ymax As Single

Dim temp As Single
Dim deltaposition As Single, deltacounts As Single, shiftposition As Single

' Find closest point less and greater then xpos
xmindif! = 1E+38
xmaxdif! = 1E+38
For i% = 1 To nPoints%
If ydata!(i%) <> NOT_ANALYZED_VALUE_SINGLE! Then
If defminpnt% = 0 Then defminpnt% = i%
defmaxpnt% = i%
End If

If xpos! <= xdata!(i%) And Abs(xpos! - xdata!(i%)) <= xmindif! And ydata!(i%) <> NOT_ANALYZED_VALUE_SINGLE! Then     ' modified 07/15/2016 for cases where the values are equal
xminpnt% = i%
xmindif! = Abs(xpos! - xdata!(i%))
End If

If xpos! > xdata!(i%) And Abs(xpos! - xdata!(i%)) < xmaxdif! And ydata!(i%) <> NOT_ANALYZED_VALUE_SINGLE! Then       ' do not change this, it is correct
xmaxpnt% = i%
xmaxdif! = Abs(xpos! - xdata!(i%))
End If

Next i%

' Check
If xminpnt% = 0 And xmaxpnt% = 0 Then GoTo MathGetInterpolatedYValueNoPoints

If xdata!(1) < xdata!(nPoints%) Then
If xminpnt% = 0 Then xminpnt% = defminpnt%
If xmaxpnt% = 0 Then xmaxpnt% = defmaxpnt%

Else
If xminpnt% = 0 Then xminpnt% = defmaxpnt%
If xmaxpnt% = 0 Then xmaxpnt% = defminpnt%
End If

' Debug
If DebugMode Then
Call IOWriteLog(vbCrLf & "MathGetInterpolatedYValue: XminPnt=" & Format$(xminpnt%) & ", XmaxPnt=" & Format$(xmaxpnt%))
Call IOWriteLog("MathGetInterpolatedYValue: Xpos=" & Format$(xpos!) & ", XDataMin=" & Format$(xdata!(xminpnt%)) & ", XDataMax=" & Format$(xdata!(xmaxpnt%)))
End If

' Interpolate between points
xmin! = xdata!(xminpnt%)
xmax! = xdata!(xmaxpnt%)
ymin! = ydata!(xminpnt%)
ymax! = ydata!(xmaxpnt%)

' Calculate interpolated Y value
deltaposition! = xmax! - xmin!
If deltaposition! = 0# Then
temp! = (ymin! + ymax!) / 2#

' Do interpolation
Else
deltacounts! = ymax! - ymin!
shiftposition! = xpos! - xmin!
temp! = ymin! + deltacounts! * shiftposition! / deltaposition!
End If

' Return calculated y
MathGetInterpolatedYValue! = temp!

' Debug
If DebugMode Then
Call IOWriteLog("MathGetInterpolatedYValue: IntY=" & Format$(temp!) & ", YDataMin=" & Format$(ydata!(xminpnt%)) & ", YDataMax=" & Format$(ydata!(xmaxpnt%)))
End If

Exit Function

' Errors
MathGetInterpolatedYValueError:
MsgBox Error$, vbOKOnly + vbCritical, "MathGetInterpolatedYValue"
ierror = True
Exit Function

MathGetInterpolatedYValueNoPoints:
msg$ = "No adjacent data points were found to interpolate between"
MsgBox msg$, vbOKOnly + vbExclamation, "MathGetInterpolatedYValue"
ierror = True
Exit Function

End Function

Function MathGetInterpolatedYValue2(xpos As Double, nPoints As Long, xdata() As Double, ydata() As Double) As Double
' Determines the y data value based on interpolating between nearest x data values (double precision version)

ierror = False
On Error GoTo MathGetInterpolatedYValue2Error

Dim i As Long, defminpnt As Long, defmaxpnt As Long
Dim xmindif As Double, xmaxdif As Double
Dim xminpnt As Long, xmaxpnt As Long

Dim xmin As Double, xmax As Double
Dim ymin As Double, ymax As Double

Dim temp As Double
Dim deltaposition As Double, deltacounts As Double, shiftposition As Double

' Find closest point less and greater then xpos
xmindif# = 1E+38
xmaxdif# = 1E+38
For i& = 1 To nPoints&
If ydata#(i&) <> NOT_ANALYZED_VALUE_DOUBLE# Then
If defminpnt& = 0 Then defminpnt& = i&
defmaxpnt& = i&
End If

If xpos# <= xdata#(i&) And Abs(xpos# - xdata#(i&)) <= xmindif# And ydata#(i&) <> NOT_ANALYZED_VALUE_DOUBLE# Then     ' modified 07/15/2016 for cases where the values are equal
xminpnt& = i&
xmindif# = Abs(xpos# - xdata#(i&))
End If

If xpos# > xdata#(i&) And Abs(xpos# - xdata#(i&)) < xmaxdif# And ydata#(i&) <> NOT_ANALYZED_VALUE_DOUBLE# Then       ' do not change this, it is correct
xmaxpnt& = i&
xmaxdif# = Abs(xpos# - xdata#(i&))
End If

Next i&

' Check
If xminpnt& = 0 And xmaxpnt& = 0 Then GoTo MathGetInterpolatedYValue2NoPoints

If xdata#(1) < xdata#(nPoints&) Then
If xminpnt& = 0 Then xminpnt& = defminpnt&
If xmaxpnt& = 0 Then xmaxpnt& = defmaxpnt&

Else
If xminpnt& = 0 Then xminpnt& = defmaxpnt&
If xmaxpnt& = 0 Then xmaxpnt& = defminpnt&
End If

' Debug
If DebugMode Then
Call IOWriteLog(vbCrLf & "MathGetInterpolatedYValue2: XminPnt=" & Format$(xminpnt&) & ", XmaxPnt=" & Format$(xmaxpnt&))
Call IOWriteLog("MathGetInterpolatedYValue2: Xpos=" & Format$(xpos#) & ", XDataMin=" & Format$(xdata#(xminpnt&)) & ", XDataMax=" & Format$(xdata#(xmaxpnt&)))
End If

' Interpolate between points
xmin# = xdata#(xminpnt&)
xmax# = xdata#(xmaxpnt&)
ymin# = ydata#(xminpnt&)
ymax# = ydata#(xmaxpnt&)

' Calculate interpolated Y value
deltaposition# = xmax# - xmin#
If deltaposition# = 0# Then
temp# = (ymin# + ymax#) / 2#

' Do interpolation
Else
deltacounts# = ymax# - ymin#
shiftposition# = xpos# - xmin#
temp# = ymin# + deltacounts# * shiftposition# / deltaposition#
End If

' Return calculated y
MathGetInterpolatedYValue2# = temp#

' Debug
If DebugMode Then
Call IOWriteLog("MathGetInterpolatedYValue2: IntY=" & Format$(temp#) & ", YDataMin=" & Format$(ydata#(xminpnt&)) & ", YDataMax=" & Format$(ydata#(xmaxpnt&)))
End If

Exit Function

' Errors
MathGetInterpolatedYValue2Error:
MsgBox Error$, vbOKOnly + vbCritical, "MathGetInterpolatedYValue2"
ierror = True
Exit Function

MathGetInterpolatedYValue2NoPoints:
msg$ = "No adjacent data points were found to interpolate between"
MsgBox msg$, vbOKOnly + vbExclamation, "MathGetInterpolatedYValue2"
ierror = True
Exit Function

End Function

Function MathCalculateSinThickness(thickness As Single, takeoff As Single) As Single
' Function to calculate sin thickness based on takeoff angle

ierror = False
On Error GoTo MathCalculateSinThicknessError

Dim radians As Single
Dim sinthickness As Single

MathCalculateSinThickness! = thickness!

' Calculate thickness based on takeoff angle
radians! = takeoff! * PI! / 180#
sinthickness! = thickness! / Sin(radians!)

MathCalculateSinThickness! = sinthickness!
Exit Function

' Errors
MathCalculateSinThicknessError:
MsgBox Error$, vbOKOnly + vbCritical, "MathCalculateSinThickness"
ierror = True
Exit Function

End Function

Function MathIsValueInBetween(adata As Single, alimit1 As Single, alimit2 As Single) As Boolean
' Function to check if passed value is between data limits

ierror = False
On Error GoTo MathIsValueInBetweenError

MathIsValueInBetween = False

If alimit1! = alimit2! Then GoTo MathIsValueInBetweenNoRange

' Check for hi/lo
If alimit1! > alimit2! Then
If adata! >= alimit2! And adata! <= alimit1! Then MathIsValueInBetween = True
End If

' Check for lo/hi
If alimit1! < alimit2! Then
If adata! >= alimit1! And adata! <= alimit2! Then MathIsValueInBetween = True
End If

Exit Function

' Errors
MathIsValueInBetweenError:
MsgBox Error$, vbOKOnly + vbCritical, "MathIsValueInBetween"
ierror = True
Exit Function

MathIsValueInBetweenNoRange:
msg$ = "No data range was specified"
MsgBox msg$, vbOKOnly + vbExclamation, "MathIsValueInBetween"
ierror = True
Exit Function

End Function

Function MathIsPowerOf2(dblNum As Long) As Boolean
' Check if a number is a power of two.

ierror = False
On Error GoTo MathIsPowerOf2Error

MathIsPowerOf2 = ((dblNum& And (dblNum& - 1)) = 0)

Exit Function

' Errors
MathIsPowerOf2Error:
MsgBox Error$, vbOKOnly + vbCritical, "MathIsPowerOf2"
ierror = True
Exit Function

End Function
