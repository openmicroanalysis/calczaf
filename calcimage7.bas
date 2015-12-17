Attribute VB_Name = "CodeCalcImage7"
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

' Output data for GRD files
Dim PrbImgData(1 To 1) As TypeImageData

Sub CalcImageCreateGRDFromArray(tfilename As String, ix As Integer, iy As Integer, sArray() As Single, xmin As Double, xmax As Double, ymin As Double, ymax As Double, zmin As Double, zmax As Double)
' Create and save a GRD file from the passed array to the passed filename

ierror = False
On Error GoTo CalcImageCreateGRDFromArrayError

Dim i As Integer, j As Integer

' Check for no x or y extents
If xmax# = xmin# Or ymax# = ymin# Then GoTo CalcImageCreateGRDFromArrayBadExtents

PrbImgData(1).id$ = "DSBB"
PrbImgData(1).ix% = ix%
PrbImgData(1).iy% = iy%

' Load stage extents (note min and max are swapped for JEOL)
PrbImgData(1).xmin# = xmin#
PrbImgData(1).xmax# = xmax#

PrbImgData(1).ymin# = ymin#
PrbImgData(1).ymax# = ymax#

PrbImgData(1).zmin# = zmin#
PrbImgData(1).zmax# = zmax#

' Dimension array
ReDim PrbImgData(1).gData(1 To ix%, 1 To iy%) As Single

Screen.MousePointer = vbHourglass

' Load data image data
For j% = 1 To iy%
For i% = 1 To ix%
PrbImgData(1).gData!(i%, j%) = sArray!(i%, j%)
Next i%
Next j%

' Load the file based on actual file version number
If SurferOutputVersionNumber% = 6 Then
Call GridFileReadWrite(Int(2), Int(1), PrbImgData(), tfilename$)
If ierror Then
Screen.MousePointer = vbDefault
Exit Sub
End If

Else
Call GridFileReadWrite2(Int(2), Int(1), PrbImgData(), tfilename$)
If ierror Then
Screen.MousePointer = vbDefault
Exit Sub
End If
End If

Exit Sub

' Errors
CalcImageCreateGRDFromArrayError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "CalcImageCreateGRDFromArray"
ierror = True
Exit Sub

CalcImageCreateGRDFromArrayBadExtents:
Screen.MousePointer = vbDefault
msg$ = "File " & tfilename$ & " has equal x and/or y extents and is therefore not a valid GRD file (xmin= " & Format$(xmin#) & ", xmax= " & Format$(xmax#) & ", ymin= " & Format$(ymin#) & ", ymax= " & Format$(ymax#) & ") . Please try again."
MsgBox msg$, vbOKOnly + vbExclamation, "CalcImageCreateGRDFromArray"
ierror = True
Exit Sub

End Sub

Sub CalcImageCreateGRDFromArray2(tfilename As String, ix As Long, iy As Long, sArray() As Double, xmin As Double, xmax As Double, ymin As Double, ymax As Double, zmin As Double, zmax As Double)
' Create and save a GRD file from the passed double precision array to the passed filename

ierror = False
On Error GoTo CalcImageCreateGRDFromArray2Error

Dim i As Long, j As Long

' Check for no x or y extents
If xmax# = xmin# Or ymax# = ymin# Then GoTo CalcImageCreateGRDFromArray2BadExtents

PrbImgData(1).id$ = "DSBB"
PrbImgData(1).ix% = ix&
PrbImgData(1).iy% = iy&

' Load stage extents (note min and max are swapped for JEOL)
PrbImgData(1).xmin# = xmin#
PrbImgData(1).xmax# = xmax#

PrbImgData(1).ymin# = ymin#
PrbImgData(1).ymax# = ymax#

PrbImgData(1).zmin# = zmin#
PrbImgData(1).zmax# = zmax#

' Dimension array
ReDim PrbImgData(1).gData(1 To ix&, 1 To iy&) As Single

Screen.MousePointer = vbHourglass

' Load data image data
For j& = 1 To iy&
For i& = 1 To ix&
PrbImgData(1).gData!(i&, j&) = CSng(sArray#(i&, j&))
Next i&
Next j&

' Load the file based on actual file version number
If SurferOutputVersionNumber% = 6 Then
Call GridFileReadWrite(Int(2), Int(1), PrbImgData(), tfilename$)
If ierror Then
Screen.MousePointer = vbDefault
Exit Sub
End If

Else
Call GridFileReadWrite2(Int(2), Int(1), PrbImgData(), tfilename$)
If ierror Then
Screen.MousePointer = vbDefault
Exit Sub
End If
End If

Exit Sub

' Errors
CalcImageCreateGRDFromArray2Error:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "CalcImageCreateGRDFromArray2"
ierror = True
Exit Sub

CalcImageCreateGRDFromArray2BadExtents:
Screen.MousePointer = vbDefault
msg$ = "File " & tfilename$ & " has equal x and/or y extents and is therefore not a valid GRD file (xmin= " & Format$(xmin#) & ", xmax= " & Format$(xmax#) & ", ymin= " & Format$(ymin#) & ", ymax= " & Format$(ymax#) & ") . Please try again."
MsgBox msg$, vbOKOnly + vbExclamation, "CalcImageCreateGRDFromArray2"
ierror = True
Exit Sub

End Sub

