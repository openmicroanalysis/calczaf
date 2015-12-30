Attribute VB_Name = "CodeBase64Reader"
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Private Type TypeReal
realval As Single
End Type

Private Type TypeByt4
strval(1 To 4) As Byte
End Type

Sub Base64ReaderInput(lpFileName As String, keV As Single, counttime As Single, beamcurrent1 As Single, beamcurrent2 As Single, timeofacq1 As Double, timeofacq2 As Double, ix As Integer, iy As Integer, sarray() As Single, xmin As Double, xmax As Double, ymin As Double, ymax As Double, zmin As Double, zmax As Double, mag As Double, scan As Double)
' Open prbimg and read in some parameters

Dim lpDefault As String

Dim date1 As String, date2 As String
Dim var1 As Variant, var2 As Variant

Dim ImageWidth As Integer
Dim ImageHeight As Integer

Dim RegXmin As Long
Dim RegXmax As Long

Dim RegYmin As Long
Dim RegYmax As Long

Dim ImageXmin As Double
Dim ImageXmax As Double

Dim ImageYmin As Double
Dim ImageYmax As Double

Dim ImageZmin As Double
Dim ImageZmax As Double

Dim PrbImgVerStr As String
Dim PrbImgVerNum As Single
Dim tmajor As Long, tminor As Long, trevision As Long

Dim j As Integer
Dim ii As Integer, jj As Integer
Dim tWidth As Integer
Dim sRow As String, astring As String

Dim gX_Polarity As Integer, gY_Polarity As Integer

Dim barray() As Byte

ierror = False
On Error GoTo Base64ReaderInputError

' Program version number
PrbImgVerStr$ = Base64ReaderGetINIString$(lpFileName$, "ProbeImage", "Version", "")
Call MiscParseStringToStringA(PrbImgVerStr$, ".", astring$)
If ierror Then Exit Sub
tmajor& = Val(astring$)
Call MiscParseStringToStringA(PrbImgVerStr$, ".", astring$)
If ierror Then Exit Sub
tminor& = Val(astring$)
trevision& = Val(astring$)
PrbImgVerNum! = Val(Format$(tmajor) & "." & Format$(tminor) & Format$(trevision))

' Read in pixel dwell time
counttime! = Val(Base64ReaderGetINIString$(lpFileName$, "ColumnConditions", "PixelTime", "0.0"))

' Read in old style beam current (in amps) in case PrbImg file is older
beamcurrent1! = Val(Base64ReaderGetINIString$(lpFileName$, "ColumnConditions", "BeamCurrent", vbNullString$))
If ierror Then Exit Sub

lpDefault$ = Format$(beamcurrent1!) ' use first beam current as default
beamcurrent1! = Val(Base64ReaderGetINIString$(lpFileName$, "Measured/BeamCurrent", "Start", lpDefault$))
If ierror Then Exit Sub
lpDefault$ = Format$(beamcurrent1!) ' use first beam current as default
beamcurrent2! = Val(Base64ReaderGetINIString$(lpFileName$, "Measured/BeamCurrent", "End", lpDefault$))
If ierror Then Exit Sub

' Read in time of acquisition
date1$ = Base64ReaderGetINIString$(lpFileName$, "Measured/Time", "Start", FileDateTime(lpFileName$))
If ierror Then Exit Sub
date2$ = Base64ReaderGetINIString$(lpFileName$, "Measured/Time", "End", FileDateTime(lpFileName$))
If ierror Then Exit Sub

' Convert to MS date
date1$ = Replace$(date1$, "T", " ")
date2$ = Replace$(date2$, "T", " ")
var1 = CDate(date1$)
var2 = CDate(date2$)
timeofacq1# = CVDate(var1)
timeofacq2# = CVDate(var2)

' Read in keV (in keV)
keV! = Val(Base64ReaderGetINIString$(lpFileName$, "ColumnConditions", "HighVoltage", vbNullString$))
If ierror Then Exit Sub

' Read image size and zmin/zmax
ImageHeight% = Val(Base64ReaderGetINIString$(lpFileName$, "RawData", "Height", vbNullString$))
If ierror Then Exit Sub
ImageWidth% = Val(Base64ReaderGetINIString$(lpFileName$, "RawData", "Width", vbNullString$))
If ierror Then Exit Sub

' Read image orientation
RegXmin& = Val(Base64ReaderGetINIString$(lpFileName$, "Registration", "X1Pixel", vbNullString$))
If ierror Then Exit Sub
RegXmax& = Val(Base64ReaderGetINIString$(lpFileName$, "Registration", "X2Pixel", vbNullString$))
If ierror Then Exit Sub

RegYmin& = Val(Base64ReaderGetINIString$(lpFileName$, "Registration", "Y1Pixel", vbNullString$))
If ierror Then Exit Sub
RegYmax& = Val(Base64ReaderGetINIString$(lpFileName$, "Registration", "Y2Pixel", vbNullString$))
If ierror Then Exit Sub

' Read image coordinates
ImageXmin# = Val(Base64ReaderGetINIString$(lpFileName$, "Registration", "X1Real", vbNullString$))
If ierror Then Exit Sub
ImageXmax# = Val(Base64ReaderGetINIString$(lpFileName$, "Registration", "X2Real", vbNullString$))
If ierror Then Exit Sub

ImageYmin# = Val(Base64ReaderGetINIString$(lpFileName$, "Registration", "Y1Real", vbNullString$))
If ierror Then Exit Sub
ImageYmax# = Val(Base64ReaderGetINIString$(lpFileName$, "Registration", "Y2Real", vbNullString$))
If ierror Then Exit Sub

ImageZmin# = Val(Base64ReaderGetINIString$(lpFileName$, "RawData", "Min", vbNullString$))
If ierror Then Exit Sub
ImageZmax# = Val(Base64ReaderGetINIString$(lpFileName$, "RawData", "Max", vbNullString$))
If ierror Then Exit Sub

' Read in mag and scan rotation if v. 1.1 or later
If PrbImgVerNum! >= 1.1 Then
mag# = Val(Base64ReaderGetINIString$(lpFileName$, "ColumnConditions", "Magnification", Format$(mag#)))
If ierror Then Exit Sub
scan# = Val(Base64ReaderGetINIString$(lpFileName$, "ColumnConditions", "ScanRotation", Format$(scan#)))
If ierror Then Exit Sub
End If

' Re-dimension real world coordinates if only 1 pixel (that is, make the 1 pixel scan dimension equal to width of single pixel)
If RegXmin& = 0 And RegXmax& = 0 Then   ' nominally JEOL line scan
ImageXmin# = ImageXmin# - 0.5 * Abs(ImageYmax# - ImageYmin#) / ImageHeight%
ImageXmax# = ImageXmax# + 0.5 * Abs(ImageYmax# - ImageYmin#) / ImageHeight%
End If

' Re-dimension real world coordinates if only 1 pixel (that is, make the 1 pixel scan dimension equal to width of single pixel)
If RegYmin& = 0 And RegYmax& = 0 Then   ' nominally Cameca line scan
ImageYmin# = ImageYmin# - 0.5 * Abs(ImageXmax# - ImageXmin#) / ImageWidth%
ImageYmax# = ImageYmax# + 0.5 * Abs(ImageXmax# - ImageXmin#) / ImageWidth%
End If

' Add one pixel for 1/2 pixel on each side of scan
If RegXmin& = 0 And RegXmax& = 0 Then   ' nominally JEOL line scan
ImageYmin# = ImageYmin# - 0.5 * Abs(ImageYmax# - ImageYmin#) / ImageHeight%
ImageYmax# = ImageYmax# + 0.5 * Abs(ImageYmax# - ImageYmin#) / ImageHeight%
End If

' Add one pixel for 1/2 pixel on each side of scan
If RegYmin& = 0 And RegYmax& = 0 Then   ' nominally Cameca line scan
ImageXmin# = ImageXmin# - 0.5 * Abs(ImageXmax# - ImageXmin#) / ImageWidth%
ImageXmax# = ImageXmax# + 0.5 * Abs(ImageXmax# - ImageXmin#) / ImageWidth%
End If

Screen.MousePointer = vbHourglass

' Dimension float array for returned data
ix% = ImageWidth%
iy% = ImageHeight%
ReDim sarray(1 To ix%, 1 To iy%) As Single

' Return other parameters
beamcurrent1! = beamcurrent1! * NAPA#
beamcurrent2! = beamcurrent2! * NAPA#

' Convert from msec to secsonds
counttime! = counttime! / MSECPERSEC#

' Special code to determine if the PrbImg file is JEOL or Cameca orientation
gX_Polarity% = -1   ' assume JEOL PrbImg
gY_Polarity% = -1   ' assume JEOL PrbImg
If RegXmin& < RegXmax& Then ' Cameca minimum x pixels are 32 so this will work (min/max can be 0 for y axis if 1 pixel high)
gX_Polarity% = 0
gY_Polarity% = 0
End If

' Set default X polarity (this works for JEOL and Cameca)
If Default_X_Polarity = 0 And gX_Polarity% = 0 Then           ' Cameca reading or writing Cameca
xmin# = ImageXmin#
xmax# = ImageXmax#
ElseIf Default_X_Polarity% <> 0 And gX_Polarity% = 0 Then      ' JEOL reading or writing Cameca
xmin# = ImageXmin#
xmax# = ImageXmax#
ElseIf Default_X_Polarity% = 0 And gX_Polarity% <> 0 Then      ' Cameca reading or writing JEOL
xmin# = ImageXmax#  ' note flip
xmax# = ImageXmin#  ' note flip
ElseIf Default_X_Polarity% <> 0 And gX_Polarity% <> 0 Then     ' JEOL reading or writing JEOL
xmin# = ImageXmax#  ' note flip
xmax# = ImageXmin#  ' note flip
End If

' Set default polarity (this works for JEOL and Cameca)
If Default_Y_Polarity = 0 And gY_Polarity% = 0 Then            ' Cameca reading or writing Cameca
ymin# = ImageYmax#  ' note flip
ymax# = ImageYmin#  ' note flip
ElseIf Default_X_Polarity% <> 0 And gX_Polarity% = 0 Then      ' JEOL reading or writing Cameca
ymin# = ImageYmax#  ' note flip
ymax# = ImageYmin#  ' note flip
ElseIf Default_X_Polarity% = 0 And gX_Polarity% <> 0 Then      ' Cameca reading or writing JEOL
ymin# = ImageYmin#
ymax# = ImageYmax#
ElseIf Default_X_Polarity% <> 0 And gX_Polarity% <> 0 Then     ' JEOL reading or writing JEOL
ymin# = ImageYmin#
ymax# = ImageYmax#
End If

zmin# = ImageZmin#
zmax# = ImageZmax#

' Handle conversion from Probe Image PrbImg file stage units for Cameca (always mm in PrbImg)
If Default_Stage_Units$ = "um" Then
xmin# = xmin# * MICRONSPERMM&
xmax# = xmax# * MICRONSPERMM&
ymin# = ymin# * MICRONSPERMM&
ymax# = ymax# * MICRONSPERMM&
End If

' Read in data scan lines one at a time
For j% = 1 To ImageHeight%

' Get each scan line (do NOT remove the trailing "=" symbol in the returned string buffer!)
sRow$ = Base64ReaderGetINIString$(lpFileName$, "RawData", "Row" & Format$(j% - 1), vbNullString$)
If ierror Then
Screen.MousePointer = vbDefault
Exit Sub
End If

' Decode to byte array (4 bytes per pixel)
barray() = Base64Decode(sRow$)

' Check byte array dimensions
ii% = LBound(barray, 1)
jj% = UBound(barray, 1)

tWidth% = (jj% - ii%) + 1
If tWidth% <> ImageWidth% * 4 Then GoTo Base64ReaderInputBadScanLine

' Convert bytes to single precision floats (note that Y dimension needs to be inverted!)
Screen.MousePointer = vbHourglass
Call Base64ReaderConvertLine(barray(), ix%, ImageHeight% - (j% - 1), sarray!())
If ierror Then
Screen.MousePointer = vbDefault
Exit Sub
End If

Next j%

Screen.MousePointer = vbDefault
Exit Sub

' Errors
Base64ReaderInputError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "Base64ReaderInput"
ierror = True
Exit Sub

Base64ReaderInputBadScanLine:
msg$ = "Unable to read scan line " & Format$(j%) & " (expected width of " & Format$(ImageWidth%) & " but found " & Format$(tWidth%) & ") in file " & lpFileName$
MsgBox msg$, vbOKOnly + vbExclamation, "Base64ReaderInput"
ierror = True
Exit Sub

End Sub

Function Base64ReaderGetINIString(lpFileName As String, lpAppName As String, lpKeyName As String, lpDefault As String) As String
' Returns a single INI string

ierror = False
On Error GoTo Base64ReaderGetINIStringError

Dim valid As Long
Dim lpReturnString As String * MAXINTEGER%
Dim nSize As Long

' Check for existing INI file
If Dir$(lpFileName$) = vbNullString Then GoTo Base64ReaderGetINIStringMissingINI
nSize& = Len(lpReturnString$)

' Get value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
If valid& > 0 Then
Base64ReaderGetINIString$ = Left$(lpReturnString$, valid&)
Else
Base64ReaderGetINIString$ = vbNullString
End If

Exit Function

' Errors
Base64ReaderGetINIStringError:
MsgBox Error$, vbOKOnly + vbCritical, "Base64ReaderGetINIString"
ierror = True
Exit Function

Base64ReaderGetINIStringMissingINI:
msg$ = "Unable to open file " & lpFileName$
MsgBox msg$, vbOKOnly + vbExclamation, "Base64ReaderGetINIString"
ierror = True
Exit Function

End Function

Sub Base64ReaderConvertLine(barray() As Byte, ix As Integer, j As Integer, sarray() As Single)
' Converts a single scan line from 4 1 byte values into a 4 byte float

ierror = False
On Error GoTo Base64ReaderConvertLineError

Dim i As Integer

Dim tstr As TypeByt4
Dim treal As TypeReal

' Loop on scan line
For i% = 1 To ix%

' Load byte array in string array
tstr.strval(1) = barray((1 + 4 * (i% - 1)) - 1)
tstr.strval(2) = barray((2 + 4 * (i% - 1)) - 1)
tstr.strval(3) = barray((3 + 4 * (i% - 1)) - 1)
tstr.strval(4) = barray((4 + 4 * (i% - 1)) - 1)

' Copy memory location
LSet treal = tstr

' Load return array
sarray!(i%, j%) = treal.realval!
Next i%

Exit Sub

' Errors
Base64ReaderConvertLineError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "Base64ReaderConvertLine"
ierror = True
Exit Sub

End Sub
