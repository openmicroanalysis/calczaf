Attribute VB_Name = "CodeTRANS3D"
' (c) Copyright 1995-2018 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub Trans3dCalculateMatrixVector(method As Integer, fiducialold() As Single, fiducialnew() As Single, fiducialtranslation() As Double, fiducialmatrix() As Double)
' Routine to calculate the fiducial transformation matrix with translation vector
' method = 0 do not confirm transformation to screen
' method = 1 confirm transformation to screen

ierror = False
On Error GoTo Trans3dCalculateMatrixVectorError

Dim i As Integer, j As Integer
Dim Vol As Double

ReDim u(1 To MAXDIM%) As Double
ReDim v(1 To MAXDIM%) As Double
ReDim w(1 To MAXDIM%) As Double

ReDim up(1 To MAXDIM%) As Double
ReDim vp(1 To MAXDIM%) As Double
ReDim wp(1 To MAXDIM%) As Double

ReDim vW(1 To MAXDIM%) As Double
ReDim wu(1 To MAXDIM%) As Double

ReDim tc(1 To MAXDIM%) As Double
ReDim tt(1 To MAXDIM%) As Double

ReDim at(1 To MAXDIM%, 1 To MAXDIM%) As Double
ReDim bt(1 To MAXDIM%, 1 To MAXDIM%) As Double
ReDim ym(1 To MAXDIM%, 1 To MAXDIM%) As Double
ReDim dm(1 To MAXDIM%, 1 To MAXDIM%) As Double

' Print fiducial set
If VerboseMode And DebugMode Or method% = 1 Then
msg$ = vbCrLf & "Old Fiducial Coordinates:"
Call IOWriteLog(msg$)
For i% = 1 To MAXDIM%   ' for each fiducial point
msg$ = Str$(i%) & " " & MiscAutoFormat$(fiducialold!(1, i%)) & MiscAutoFormat$(fiducialold!(2, i%))
If NumberOfStageMotors% > 2 Then msg$ = msg$ & MiscAutoFormat$(fiducialold!(3, i%))
If NumberOfStageMotors% > 3 Then msg$ = msg$ & MiscAutoFormat$(fiducialold!(4, i%))
Call IOWriteLog(msg$)
Next i%

msg$ = vbCrLf & "New Fiducial Coordinates:"
Call IOWriteLog(msg$)
For i% = 1 To MAXDIM%
msg$ = Str$(i%) & " " & MiscAutoFormat$(fiducialnew!(1, i%)) & MiscAutoFormat$(fiducialnew!(2, i%))
If NumberOfStageMotors% > 2 Then msg$ = msg$ & MiscAutoFormat$(fiducialnew!(3, i%))
If NumberOfStageMotors% > 3 Then msg$ = msg$ & MiscAutoFormat$(fiducialnew!(4, i%))
Call IOWriteLog(msg$)
Next i%
End If

' Calculate translation vectors
For i% = 1 To MAXDIM%   ' for x, y and z
u#(i%) = fiducialold!(i%, 2) - fiducialold!(i%, 1)
v#(i%) = fiducialold!(i%, 3) - fiducialold!(i%, 2)

up#(i%) = fiducialnew!(i%, 2) - fiducialnew!(i%, 1)
vp#(i%) = fiducialnew!(i%, 3) - fiducialnew!(i%, 2)
Next i%

' Take the cross products
Call Trans3dCrossProduct(u#(), v#(), w#())
If ierror Then Exit Sub
Call Trans3dCrossProduct(up#(), vp#(), wp#())
If ierror Then Exit Sub

' Compute the sum of squares of the w vector
Vol# = 0#
For i% = 1 To MAXDIM%
Vol# = Vol# + w#(i%) ^ 2
Next i%
If Vol# = 0# Then GoTo Trans3dCalculateMatrixVectorLinear

Call Trans3dCrossProduct(v#(), w#(), vW#())
If ierror Then Exit Sub
Call Trans3dCrossProduct(w#(), u#(), wu#())
If ierror Then Exit Sub

' Load new matrix
For i% = 1 To MAXDIM%
bt#(i%, 1) = up#(i%)
bt#(i%, 2) = vp#(i%)
bt#(i%, 3) = wp#(i%)
Next i%

' Load the matrix to transpose
For i% = 1 To MAXDIM%
at#(i%, 1) = vW#(i%)
at#(i%, 2) = wu#(i%)
at#(i%, 3) = w#(i%)
Next i%

' Transpose matrix
For i% = 1 To MAXDIM%
For j% = 1 To MAXDIM%
ym#(i%, j%) = at#(j%, i%)
Next j%
Next i%

' Multiply "bt()" by transposed matrix "ym()"
Call Trans3dMultiplyMatrixMatrix(MAXDIM%, bt#(), ym#(), dm#())
If ierror Then Exit Sub

' Mutiply through to get the rotation matrix
For i% = 1 To MAXDIM%
fiducialmatrix#(1, i%) = 1# / Vol# * dm#(1, i%)
fiducialmatrix#(2, i%) = 1# / Vol# * dm#(2, i%)
fiducialmatrix#(3, i%) = 1# / Vol# * dm#(3, i%)
Next i%

' Calculate the translation matrix (sum x, y and z for each point)
For i% = 1 To MAXDIM%
tc#(i%) = fiducialold!(i%, 1) + fiducialold!(i%, 2) + fiducialold!(i%, 3)
Next i%

Call Trans3dMultiplyMatrix(MAXDIM%, fiducialmatrix#(), tc#(), tt#())
If ierror Then Exit Sub

For i% = 1 To MAXDIM%
fiducialtranslation#(i%) = 1# / 3# * (fiducialnew!(i%, 1) + fiducialnew!(i%, 2) + fiducialnew!(i%, 3) - tt#(i%))
Next i%

' Type out rotation matrix
If VerboseMode And DebugMode Or method% = 1 Then
msg$ = vbCrLf & "Fiducial Rotation Matrix:"
Call IOWriteLog(msg$)
msg$ = vbNullString
For i% = 1 To MAXDIM%
msg$ = msg$ & Format$(Format$(fiducialmatrix#(1, i%), e82$), a90$) & " "
msg$ = msg$ & Format$(Format$(fiducialmatrix#(2, i%), e82$), a90$) & " "
msg$ = msg$ & Format$(Format$(fiducialmatrix#(3, i%), e82$), a90$) & vbCrLf
Next i%
Call IOWriteLog(msg$)

' Type out translation matrix
msg$ = "Fiducial Translation Matrix:"
Call IOWriteLog(msg$)
msg$ = vbNullString
msg$ = msg$ & Format$(Format$(fiducialtranslation#(1), e82$), a90$) & " "
msg$ = msg$ & Format$(Format$(fiducialtranslation#(2), e82$), a90$) & " "
msg$ = msg$ & Format$(Format$(fiducialtranslation#(3), e82$), a90$) & vbCrLf
Call IOWriteLog(msg$)
End If

Exit Sub

' Errors
Trans3dCalculateMatrixVectorError:
MsgBox Error$, vbOKOnly + vbCritical, "Trans3dCalculateMatrixVector"
ierror = True
Exit Sub

Trans3dCalculateMatrixVectorLinear:
msg$ = "Fiducial points are in line. Cannot calculate matrix transformation."
MsgBox msg$, vbOKOnly + vbExclamation, "Trans3dCalculateMatrixVector"
ierror = True
Exit Sub

End Sub

Sub Trans3dCrossProduct(xyz1() As Double, xyz2() As Double, xyz() As Double)
' Perfrom a cross product on two 3 x 1 arrays

ierror = False
On Error GoTo Trans3dCrossProductError

xyz#(1) = (xyz1#(2) * xyz2#(3)) - (xyz2#(2) * xyz1(3))
xyz#(2) = (xyz2#(1) * xyz1#(3)) - (xyz1#(1) * xyz2(3))
xyz#(3) = (xyz1#(1) * xyz2#(2)) - (xyz2#(1) * xyz1(2))

Exit Sub

' Errors
Trans3dCrossProductError:
MsgBox Error$, vbOKOnly + vbCritical, "Trans3dCrossProduct"
ierror = True
Exit Sub

End Sub

Sub Trans3dMultiplyMatrix(n As Integer, Y() As Double, b() As Double, row() As Double)
' Multiply a matrix times a vector. That is y#() (n% x n%), by b#() (n% x 1), to get row#() (n% x 1).

ierror = False
On Error GoTo Trans3dMultiplyMatrixError

Dim i As Integer, j As Integer

' Initialize
For i% = 1 To n%
row#(i%) = 0#
Next i%

' Multiply
For i% = 1 To n%
For j% = 1 To n%
row#(i%) = row#(i%) + Y#(i%, j%) * b#(j%)
Next j%
Next i%

Exit Sub

' Errors
Trans3dMultiplyMatrixError:
MsgBox Error$, vbOKOnly + vbCritical, "Trans3dMultiplyMatrix"
ierror = True
Exit Sub

End Sub

Sub Trans3dMultiplyMatrixMatrix(n As Integer, a() As Double, b() As Double, c() As Double)
' Multiply two matrices. That is a#() (n% x n%), by b#() (n% x n), to get c#() (n% x n%).

ierror = False
On Error GoTo Trans3dMultiplyMatrixMatrixError

Dim i As Integer, j As Integer, k As Integer

' Multiply
For i% = 1 To n%
For j% = 1 To n%
c#(i%, j%) = 0#

For k% = 1 To n%
c#(i%, j%) = c#(i%, j%) + a#(i%, k%) * b#(k%, j%)
Next k%

Next j%
Next i%

Exit Sub

' Errors
Trans3dMultiplyMatrixMatrixError:
MsgBox Error$, vbOKOnly + vbCritical, "Trans3dMultiplyMatrixMatrix"
ierror = True
Exit Sub

End Sub

Sub Trans3dTransformPositionVector(fiducialtranslation() As Double, fiducialmatrix() As Double, xyz() As Single)
' Routine to transform a single coordinate using fiducial transformation matrix plus translation vector

ierror = False
On Error GoTo Trans3dTransformPositionVectorError

ReDim xyz1(1 To MAXDIM%) As Double
ReDim xyz2(1 To MAXDIM%) As Double

xyz1#(1) = xyz!(1)
xyz1#(2) = xyz!(2)
xyz1#(3) = xyz!(3)

' Multiply times fiducial matrix to transform coordinate to new position
Call Trans3dMultiplyMatrix(MAXDIM%, fiducialmatrix#(), xyz1#(), xyz2#())
If ierror Then Exit Sub

xyz!(1) = xyz2#(1) + fiducialtranslation#(1)
xyz!(2) = xyz2#(2) + fiducialtranslation#(2)
xyz!(3) = xyz2#(3) + fiducialtranslation#(3)

Exit Sub

' Errors
Trans3dTransformPositionVectorError:
MsgBox Error$, vbOKOnly + vbCritical, "Trans3dTransformPositionVector"
ierror = True
Exit Sub

End Sub

