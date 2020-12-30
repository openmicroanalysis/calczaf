Attribute VB_Name = "CodeGridFile"
' (c) Copyright 1995-2021 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

' Surfer v. 6 structure (used internally in all applications)
Type TypeImageData
id As String * 4
ix As Integer
iy As Integer
xmin As Double
xmax As Double
ymin As Double
ymax As Double
zmin As Double
zmax As Double
gData() As Single
End Type

' Surfer v. 7 structures (used for GRD export only)
Type TypeGridHeader
tag As Long
Size As Long
Version As Long
End Type

Type TypeGridGrid
tag As Long
Size As Long
nCol As Long
nRow As Long
xLL As Double
yLL As Double
xSize As Double
ySize As Double
zmin As Double
zmax As Double
rotation As Double
BlankValue As Double
End Type

Type TypeGridData
tag As Long
Size As Long
hdata() As Double
End Type

Dim gX_Polarity As Integer
Dim gY_Polarity As Integer
Dim gStage_Units As String

Dim sarray() As Single

Sub GridFileReadWrite(mode As Integer, findex As Integer, iData() As TypeImageData, tfilename As String)
' Read or write grid data from or to  disk (Surfer v. 6 and earlier)
' mode = 1 read
' mode = 2 write

ierror = False
On Error GoTo GridFileReadWriteError

Dim i As Integer, j As Integer

If Trim$(tfilename$) = vbNullString Then Exit Sub

' New code to check for GRDInfo.ini
Call GridCheckGRDInfo(tfilename$, gX_Polarity%, gY_Polarity%, gStage_Units$)
If ierror Then Exit Sub

' Read the file
If mode% = 1 Then
Screen.MousePointer = vbHourglass
Open tfilename$ For Binary Access Read As #Temp1FileNumber%
Get #Temp1FileNumber%, , iData(findex%).id$

' Check for binary vs. ASCII
If iData(findex%).id$ <> "DSBB" Then GoTo GridFileReadWriteNotBinary

Get #Temp1FileNumber%, , iData(findex%).ix%
Get #Temp1FileNumber%, , iData(findex%).iy%
Get #Temp1FileNumber%, , iData(findex%).xmin#
Get #Temp1FileNumber%, , iData(findex%).xmax#
Get #Temp1FileNumber%, , iData(findex%).ymin#
Get #Temp1FileNumber%, , iData(findex%).ymax#
Get #Temp1FileNumber%, , iData(findex%).zmin#
Get #Temp1FileNumber%, , iData(findex%).zmax#

' Dimension data array
ReDim iData(findex%).gData(1 To iData(findex%).ix%, 1 To iData(findex%).iy%) As Single
Get #Temp1FileNumber%, , iData(findex%).gData!
Close #Temp1FileNumber%

' Profile the code below
'Dim startTime As Currency
'Tanner_SupportCode.EnableHighResolutionTimers
'Tanner_SupportCode.GetHighResTime startTime

' Set out of range values to min and max
For j% = 1 To iData(findex%).iy%
For i% = 1 To iData(findex%).ix%
If iData(findex%).gData(i%, j%) < iData(findex%).zmin# Then iData(findex%).gData(i%, j%) = iData(findex%).zmin#
If iData(findex%).gData(i%, j%) > iData(findex%).zmax# Then iData(findex%).gData(i%, j%) = iData(findex%).zmax#
Next i%
Next j%

'Debug.Print "GridFileReadWrite - min/max check:"
'Tanner_SupportCode.PrintTimeTakenInMs startTime

' Check for appropriate stage polarity and units conversion of image data and min/max
Call GridCheckGRDConvert(mode%, tfilename$, gX_Polarity%, gY_Polarity%, gStage_Units$, findex%, iData())
If ierror Then Exit Sub

Screen.MousePointer = vbDefault

' Write the file
Else
Screen.MousePointer = vbHourglass

' Check for appropriate stage polarity and units conversion of image data and min/max
Call GridCheckGRDConvert(mode%, tfilename$, gX_Polarity%, gY_Polarity%, gStage_Units$, findex%, iData())
If ierror Then Exit Sub

' Write image data
Open tfilename$ For Binary Access Write As #Temp1FileNumber%
Put #Temp1FileNumber%, , iData(findex%).id$
Put #Temp1FileNumber%, , iData(findex%).ix%
Put #Temp1FileNumber%, , iData(findex%).iy%
Put #Temp1FileNumber%, , iData(findex%).xmin#
Put #Temp1FileNumber%, , iData(findex%).xmax#
Put #Temp1FileNumber%, , iData(findex%).ymin#
Put #Temp1FileNumber%, , iData(findex%).ymax#
Put #Temp1FileNumber%, , iData(findex%).zmin#
Put #Temp1FileNumber%, , iData(findex%).zmax#
Put #Temp1FileNumber%, , iData(findex%).gData!
Close #Temp1FileNumber%
Screen.MousePointer = vbDefault
End If

Exit Sub

' Errors
GridFileReadWriteError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "GridFileReadWrite"
Close #Temp1FileNumber%
ierror = True
Exit Sub

GridFileReadWriteNotBinary:
Screen.MousePointer = vbDefault
msg$ = "Grid file is not binary"
MsgBox msg$, vbOKOnly + vbExclamation, "GridFileReadWrite"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub GridFileReadWrite2(mode As Integer, findex As Integer, iData() As TypeImageData, tfilename As String)
' Read or write grid data from or to  disk (Surfer version 7 or higher)
' mode = 1 read
' mode = 2 write

ierror = False
On Error GoTo GridFileReadWrite2Error

Dim ii As Long, jj As Long

' Version 7 structures for reading and writing
Dim theader As TypeGridHeader
Dim tGrid As TypeGridGrid
Dim tData As TypeGridData

If Trim$(tfilename$) = vbNullString Then Exit Sub

' New code to check for GRDInfo.ini
Call GridCheckGRDInfo(tfilename$, gX_Polarity%, gY_Polarity%, gStage_Units$)
If ierror Then Exit Sub

' Read the file
If mode% = 1 Then
Screen.MousePointer = vbHourglass
Open tfilename$ For Binary Access Read As #Temp1FileNumber%

' Load header section
Get #Temp1FileNumber%, , theader.tag&

' Check for version 7 header section (hex 0x42525344)
If theader.tag& <> 1112691524 Then GoTo GridFileReadWrite2BadFile

Get #Temp1FileNumber%, , theader.Size&
Get #Temp1FileNumber%, , theader.Version&

' Load into image array
iData(findex%).id$ = "DSBB"

' Load grid section
Get #Temp1FileNumber%, , tGrid.tag&

' Check for version 7 grid section (hex  0x44495247)
If tGrid.tag& <> 1145655879 Then GoTo GridFileReadWrite2BadFile

Get #Temp1FileNumber%, , tGrid.Size&
Get #Temp1FileNumber%, , tGrid.nRow&
Get #Temp1FileNumber%, , tGrid.nCol&
Get #Temp1FileNumber%, , tGrid.xLL#
Get #Temp1FileNumber%, , tGrid.yLL#
Get #Temp1FileNumber%, , tGrid.xSize#
Get #Temp1FileNumber%, , tGrid.ySize#
Get #Temp1FileNumber%, , tGrid.zmin#
Get #Temp1FileNumber%, , tGrid.zmax#
Get #Temp1FileNumber%, , tGrid.rotation#
Get #Temp1FileNumber%, , tGrid.BlankValue#

' Check image size
If tGrid.nCol& > MAXINTEGER% Then GoTo GridFileReadWrite2TooLarge
If tGrid.nRow& > MAXINTEGER% Then GoTo GridFileReadWrite2TooLarge

' Load into image array
iData(findex%).ix% = tGrid.nCol&    ' number of horizontal columns
iData(findex%).iy% = tGrid.nRow&    ' number of vertical rows

' Calculate X dimensions
iData(findex%).xmin# = tGrid.xLL#
iData(findex%).xmax# = tGrid.xLL# + (tGrid.nCol& * tGrid.xSize#)

' Calculate Y dimensions
iData(findex%).ymin# = tGrid.yLL#
iData(findex%).ymax# = tGrid.yLL# + (tGrid.nRow& * tGrid.ySize#)

iData(findex%).zmin# = tGrid.zmin#
iData(findex%).zmax# = tGrid.zmax#

' Dimension data array
ReDim tData.hdata(1 To tGrid.nCol&, 1 To tGrid.nRow&) As Double

' Load data section
Get #Temp1FileNumber%, , tData.tag&

' Check for version 7 data section (hex  0x41544144) (note typo in Surfer help file)
If tData.tag& <> 1096040772 Then GoTo GridFileReadWrite2BadFile

Get #Temp1FileNumber%, , tData.Size&
Get #Temp1FileNumber%, , tData.hdata#
Close #Temp1FileNumber%

ReDim iData(findex%).gData(1 To iData(findex%).ix%, 1 To iData(findex%).iy%) As Single

' Profile the code below
'Dim startTime As Currency
'Tanner_SupportCode.EnableHighResolutionTimers
'Tanner_SupportCode.GetHighResTime startTime
    
    ' Load image array and set out of range values to min and max (to avoid "blanking values")
    Dim tmpFloat As Single
    
    ' Cache relevant array properties inside local variables for speed
    Dim iBound As Long, jBound As Long
    iBound& = iData(findex%).ix%
    jBound& = iData(findex%).iy%
    
    Dim zMinVal As Single, zMaxVal As Single
    zMinVal! = iData(findex%).zmin#
    zMaxVal! = iData(findex%).zmax#
    
    For jj& = 1 To jBound&
    For ii& = 1 To iBound&
        
        ' To minimize the number of times we need to access tData.hdata#(), cache its value up front.
        tmpFloat! = tData.hdata#(ii&, jj&)
        If (tmpFloat! <> BLANKINGVALUE!) Then
            If (tmpFloat! < zMinVal!) Then tmpFloat! = zMinVal!
            If (tmpFloat! > zMaxVal!) Then tmpFloat! = zMaxVal!
            iData(findex%).gData!(ii&, jj&) = tmpFloat!
        Else
            iData(findex%).gData!(ii&, jj&) = tmpFloat!
        End If
        
    Next ii&
    Next jj&
    
'Debug.Print "GridFileReadWrite2 - min/max check:"
'Tanner_SupportCode.PrintTimeTakenInMs startTime

' Check for appropriate stage polarity and units conversion of image data and min/max
Call GridCheckGRDConvert(mode%, tfilename$, gX_Polarity%, gY_Polarity%, gStage_Units$, findex%, iData())
If ierror Then Exit Sub

Screen.MousePointer = vbDefault

' Write the file
Else
Screen.MousePointer = vbHourglass

' Check for appropriate stage polarity and units conversion of image data and min/max
Call GridCheckGRDConvert(mode%, tfilename$, gX_Polarity%, gY_Polarity%, gStage_Units$, findex%, iData())
If ierror Then Exit Sub

' Version 7 header section (hex 0x42525344)
theader.tag& = 1112691524
theader.Size& = 4
theader.Version& = 1

' Version 7 grid section (hex  0x44495247)
tGrid.tag& = 1145655879
tGrid.Size& = 72

' Save to array
tGrid.nCol& = iData(findex%).ix%    ' number of horizontal columns!
tGrid.nRow& = iData(findex%).iy%    ' number of vertical rows!
tGrid.xLL# = iData(findex%).xmin#
tGrid.yLL# = iData(findex%).ymin#
tGrid.xSize# = (iData(findex%).xmax# - iData(findex%).xmin#) / (iData(findex%).ix%)
tGrid.ySize# = (iData(findex%).ymax# - iData(findex%).ymin#) / (iData(findex%).iy%)

tGrid.zmin# = iData(findex%).zmin#
tGrid.zmax# = iData(findex%).zmax#
tGrid.rotation# = 0#
tGrid.BlankValue# = BLANKINGVALUE!

' Dimension output data array
ReDim tData.hdata(1 To iData(findex%).ix%, 1 To iData(findex%).iy%) As Double

' Version 7 data section (hex  0x41544144) (note typo in old versions of Surfer help file)
tData.tag& = 1096040772
tData.Size& = tGrid.nRow& * tGrid.nCol& * 8

' Load data for writing
For jj& = 1 To iData(findex%).iy%
For ii& = 1 To iData(findex%).ix%
tData.hdata#(ii&, jj&) = iData(findex%).gData!(ii&, jj&)
Next ii&
Next jj&

Close #Temp1FileNumber%
DoEvents
Open tfilename$ For Binary Access Write As #Temp1FileNumber%
Put #Temp1FileNumber%, , theader.tag&
Put #Temp1FileNumber%, , theader.Size&
Put #Temp1FileNumber%, , theader.Version&

Put #Temp1FileNumber%, , tGrid.tag&
Put #Temp1FileNumber%, , tGrid.Size&
Put #Temp1FileNumber%, , tGrid.nRow&
Put #Temp1FileNumber%, , tGrid.nCol&
Put #Temp1FileNumber%, , tGrid.xLL#
Put #Temp1FileNumber%, , tGrid.yLL#
Put #Temp1FileNumber%, , tGrid.xSize#
Put #Temp1FileNumber%, , tGrid.ySize#
Put #Temp1FileNumber%, , tGrid.zmin#
Put #Temp1FileNumber%, , tGrid.zmax#
Put #Temp1FileNumber%, , tGrid.rotation#
Put #Temp1FileNumber%, , tGrid.BlankValue#

Put #Temp1FileNumber%, , tData.tag&
Put #Temp1FileNumber%, , tData.Size&

Put #Temp1FileNumber%, , tData.hdata#
Close #Temp1FileNumber%

Screen.MousePointer = vbDefault
End If

Exit Sub

' Errors
GridFileReadWrite2Error:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "GridFileReadWrite2"
Close #Temp1FileNumber%
ierror = True
Exit Sub

GridFileReadWrite2BadFile:
Screen.MousePointer = vbDefault
msg$ = "File is not a valid Surfer v. 7 grid file"
MsgBox msg$, vbOKOnly + vbExclamation, "GridFileReadWrite2"
Close #Temp1FileNumber%
ierror = True
Exit Sub

GridFileReadWrite2TooLarge:
msg$ = "GRD image is too large. Please select a smaller size and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "GridFileReadWrite2"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Function GridFileGetVersion(tfilename As String) As Single
' Return the GRID file version number

ierror = False
On Error GoTo GridFileGetVersionError

Dim tLong As Long

' Open file and read first 4 bytes
Open tfilename$ For Binary Access Read As #Temp1FileNumber%
Get #Temp1FileNumber%, , tLong&

' Check for version 6 binary ("DSBB" or 1111642948)
If tLong& = 1111642948 Then
GridFileGetVersion! = 6#

' Check for version 7 binary ("DSRB" or 1112691524)
ElseIf tLong& = 1112691524 Then
GridFileGetVersion! = 7#

' ASCII grid ("DSAA" or bad file)
Else
GoTo GridFileGetVersionBadFile
End If

Close Temp1FileNumber%
Exit Function

' Errors
GridFileGetVersionError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "GridFileGetVersion"
Close Temp1FileNumber%
ierror = True
Exit Function

GridFileGetVersionBadFile:
Screen.MousePointer = vbDefault
msg$ = "File is not a valid Surfer v. 6 or v. 7 grid file"
MsgBox msg$, vbOKOnly + vbExclamation, "GridFileGetVersion"
Close Temp1FileNumber%
ierror = True
Exit Function

End Function

Public Function GridIsArrayInitalized(ByRef arr() As TypeImageData) As Boolean
' Return True if array is initalized

On Error GoTo GridIsArrayInitalizedError ' raise error if array is not initialzied

Dim temp As Long

GridIsArrayInitalized = False
temp& = UBound(arr)

' We reach this point only if arr is initalized, i.e. no error occured
If temp& > -1 Then GridIsArrayInitalized = True  ' UBound is greater then -1
Exit Function

' Special error handler (if an error occurs, this function returns False. i.e. array not initialized)
GridIsArrayInitalizedError:
Exit Function

End Function

Public Sub GridCheckGRDInfo(tfilename As String, gX_Polarity As Integer, gY_Polarity As Integer, gStage_Units As String)
' Check for GRDinfo.ini file to read or write appropriately
'  tfilename is the path and filename of the GRD file to read or write (use for reading or creating GRDInfo.INI in the GRD file folder)
' Returned:
'  gX_Polarity is the GRD x stage axis polarity (0 = cartesian (normal GRD orientation), non-zero = anti-cartesian)
'  dY_Polarity is the GRD y stage axis polarity (0 = cartesian (normal GRD orientation), non-zero = anti-cartesian)
'  dStage_Units is the GRD units string (must be "mm" for millimeters or "um" for microns or micrometers)
'
' Globals:
'  Default_X_Polarity is the DEFAULT or EXPECTED x stage axis polarity (0 = cartesian (normal GRD orientation), non-zero = anti-cartesian)
'  Default_Y_Polarity is the DEFAULT or EXPECTED y stage axis polarity (0 = cartesian (normal GRD orientation), non-zero = anti-cartesian)
'  Default_Stage_Units is the DEFAULT or EXPECTED units string (must be "mm" for millimeters or "um" for microns or micrometers)

ierror = False
On Error GoTo GridCheckGRDInfoError

Dim valid As Long

Dim lpAppName As String
Dim lpKeyName As String
Dim lpDefault As String
Dim lpFileName As String
Dim lpReturnString As String * 255

Dim lpString As String

Dim nSize As Long
Dim nDefault As Long

' Get string buffer length
nSize& = Len(lpReturnString$)

' Check for empty file
If tfilename$ = vbNullString Then GoTo GridCheckGRDInfoNoFile

' Load expected INI file name
If Dir$(MiscGetFileNameNoExtension$(tfilename$) & ".ACQ") <> vbNullString Then
lpFileName$ = MiscGetFileNameNoExtension$(tfilename$) & ".ACQ"
Else
lpFileName$ = MiscGetPathOnly$(tfilename$) & "GRDInfo.INI"
End If

' Check for existing GRDInfo.INI file and if found, load into temporary GRD variables
If Dir$(lpFileName$) <> vbNullString Then
lpAppName$ = "Stage"
lpKeyName$ = "X_Polarity"
nDefault& = Default_X_Polarity%
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& = 0 Then
gX_Polarity% = 0
Else
gX_Polarity% = -1
End If

lpAppName$ = "Stage"
lpKeyName$ = "Y_Polarity"
nDefault& = Default_Y_Polarity%
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& = 0 Then
gY_Polarity% = 0
Else
gY_Polarity% = -1
End If

lpAppName$ = "Stage"
lpKeyName$ = "Stage_Units"
lpDefault$ = Default_Stage_Units$
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
gStage_Units$ = Left$(lpReturnString$, valid&)
If gStage_Units$ <> "mm" And gStage_Units$ <> "um" Then GoTo GridCheckGRDInfoBadUnits

' No existing GRDInfo.ini file, so create one using global default values
Else
lpAppName$ = "Stage"
lpKeyName$ = "X_Polarity"
lpString$ = Format$(Default_X_Polarity%)
valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, lpString$, lpFileName$)

lpAppName$ = "Stage"
lpKeyName$ = "Y_Polarity"
lpString$ = Format$(Default_Y_Polarity%)
valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, lpString$, lpFileName$)

lpAppName$ = "Stage"
lpKeyName$ = "Stage_Units"
lpString$ = Default_Stage_Units$
valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, lpString$, lpFileName$)

' Since new file was created, "load" variables from disk
gX_Polarity% = Default_X_Polarity%
gY_Polarity% = Default_Y_Polarity%
gStage_Units$ = Default_Stage_Units$
End If

Exit Sub

' Errors
GridCheckGRDInfoError:
MsgBox Error$, vbOKOnly + vbCritical, "GridCheckGRDInfo"
ierror = True
Exit Sub

GridCheckGRDInfoNoFile:
msg$ = "The passed filename was empty, no path can be determined. This error should not occur- please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "GridCheckGRDInfo"
ierror = True
Exit Sub

GridCheckGRDInfoBadUnits:
msg$ = "GRDInfo.INI file (" & lpFileName$ & ") does not contain mm or um stage units. Please delete or edit by hand for mm (millimeter) or um (micrometer) units and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "GridCheckGRDInfo"
ierror = True
Exit Sub

End Sub

Private Sub GridCheckGRDConvert(method As Integer, tfilename As String, gX_Polarity As Integer, gY_Polarity As Integer, gStage_Units As String, findex As Integer, iData() As TypeImageData)
' Check for GRDinfo.ini file to convert GRD data appropriately (check stage polarity and units)
'  method = 1 read GRD from disk
'  method = 2 write GRD to disk
'  tfilename  is the path and filename of the GRD file to read or write
'
'  gX_Polarity is the GRD file x stage axis polarity (0 = cartesian (normal GRD orientation), non-zero = anti-cartesian)
'  gY_Polarity is the GRD file y stage axis polarity (0 = cartesian (normal GRD orientation), non-zero = anti-cartesian)
'  gStage_Units is the GRD file units string (must be "mm" for millimeters or "um" for microns or micrometers)
'
' findex is the image index of the passed image data structure
' iData() is the array of image data structures
'
' Globals
'  Default_X_Polarity is the DEFAULT or EXPECTED x stage axis polarity (0 = cartesian (normal GRD orientation), non-zero = anti-cartesian)
'  Default_Y_Polarity is the DEFAULT or EXPECTED y stage axis polarity (0 = cartesian (normal GRD orientation), non-zero = anti-cartesian)
'  Default_Stage_Units is the DEFAULT or EXPECTED units string (must be "mm" for millimeters or "um" for microns or micrometers)

ierror = False
On Error GoTo GridCheckGRDConvertError

Dim ix As Integer, iy As Integer
Dim temp As Double
Dim dstring As String, gstring As String
Dim tX_Polarity As Integer, tY_Polarity As Integer

Dim ii As Long, jj As Long
Dim iBound As Long, jBound As Long

' Load error string for error messages
dstring$ = "DEFAULT: X_Polarity=" & Format$(Default_X_Polarity%) & ", Y_Polarity=" & Format$(Default_Y_Polarity%) & ", Stage Units=" & Default_Stage_Units$
gstring$ = ", GRIDINFO: X_Polarity=" & Format$(gX_Polarity%) & ", Y_Polarity=" & Format$(gY_Polarity%) & ", Stage Units=" & gStage_Units$

' Swap X polarity if necessary
tX_Polarity = False
If Default_X_Polarity% = 0 And gX_Polarity% = 0 Then           ' Cameca reading or writing Cameca

' Should be nothing to do, but check anyway for reading and writing (force cartesian)
If iData(findex%).xmin# > iData(findex%).xmax# Then
temp# = iData(findex%).xmin#
iData(findex%).xmin# = iData(findex%).xmax#
iData(findex%).xmax# = temp#
tX_Polarity = True
End If

ElseIf Default_X_Polarity% <> 0 And gX_Polarity% = 0 Then      ' JEOL reading or writing Cameca
If method% = 1 Then ' reading (force anti-cartesian)
If iData(findex%).xmax# > iData(findex%).xmin# Then
temp# = iData(findex%).xmin#
iData(findex%).xmin# = iData(findex%).xmax#
iData(findex%).xmax# = temp#
tX_Polarity = True
End If
Else                ' writing (force cartesian)
If iData(findex%).xmin# > iData(findex%).xmax# Then
temp# = iData(findex%).xmin#
iData(findex%).xmin# = iData(findex%).xmax#
iData(findex%).xmax# = temp#
tX_Polarity = True
End If
End If

ElseIf Default_X_Polarity% = 0 And gX_Polarity% <> 0 Then      ' Cameca reading or writing JEOL
If iData(findex%).xmax# < iData(findex%).xmin# Then            ' force cartesian reading or writing
temp# = iData(findex%).xmin#
iData(findex%).xmin# = iData(findex%).xmax#
iData(findex%).xmax# = temp#
tX_Polarity = True
End If

ElseIf Default_X_Polarity% <> 0 And gX_Polarity% <> 0 Then     ' JEOL reading or writing JEOL
If method% = 1 Then ' reading (force anti-cartesian)
If iData(findex%).xmax# > iData(findex%).xmin# Then
temp# = iData(findex%).xmin#
iData(findex%).xmin# = iData(findex%).xmax#
iData(findex%).xmax# = temp#
tX_Polarity = True
End If
Else                ' writing (force cartesian)
If iData(findex%).xmin# > iData(findex%).xmax# Then
temp# = iData(findex%).xmin#
iData(findex%).xmin# = iData(findex%).xmax#
iData(findex%).xmax# = temp#
tX_Polarity = True
End If
End If
End If

' Swap Y polarity if necessary
tY_Polarity = False
If Default_Y_Polarity% = 0 And gY_Polarity% = 0 Then           ' Cameca reading or writing Cameca

' Should be nothing to do, but check anyway for reading and writing  (force cartesian)
If iData(findex%).ymin# > iData(findex%).ymax# Then
temp# = iData(findex%).ymin#
iData(findex%).ymin# = iData(findex%).ymax#
iData(findex%).ymax# = temp#
tY_Polarity = True
End If

ElseIf Default_Y_Polarity% <> 0 And gY_Polarity% = 0 Then      ' JEOL reading or writing Cameca
If method% = 1 Then ' reading (force anti-cartesian)
If iData(findex%).ymax# > iData(findex%).ymin# Then
temp# = iData(findex%).ymin#
iData(findex%).ymin# = iData(findex%).ymax#
iData(findex%).ymax# = temp#
tY_Polarity = True
End If
Else                ' writing (force cartesian)
If iData(findex%).ymin# > iData(findex%).ymax# Then
temp# = iData(findex%).ymin#
iData(findex%).ymin# = iData(findex%).ymax#
iData(findex%).ymax# = temp#
tY_Polarity = True
End If
End If

ElseIf Default_Y_Polarity% = 0 And gY_Polarity% <> 0 Then      ' Cameca reading or writing JEOL
If iData(findex%).ymax# < iData(findex%).ymin# Then            ' force cartesian reading or writing
temp# = iData(findex%).ymin#
iData(findex%).ymin# = iData(findex%).ymax#
iData(findex%).ymax# = temp#
tY_Polarity = True
End If

ElseIf Default_Y_Polarity% <> 0 And gY_Polarity% <> 0 Then     ' JEOL reading or writing JEOL
If method% = 1 Then ' reading (force anti-cartesian)
If iData(findex%).ymax# > iData(findex%).ymin# Then
temp# = iData(findex%).ymin#
iData(findex%).ymin# = iData(findex%).ymax#
iData(findex%).ymax# = temp#
tY_Polarity = True
End If
Else                ' writing (force cartesian)
If iData(findex%).ymin# > iData(findex%).ymax# Then
temp# = iData(findex%).ymin#
iData(findex%).ymin# = iData(findex%).ymax#
iData(findex%).ymax# = temp#
tY_Polarity = True
End If
End If
End If

' Dimension temp array
ReDim sarray(1 To iData(findex%).ix%, 1 To iData(findex%).iy%) As Single

' Load temp array
sarray = iData(findex%).gData

' Check for flipping data
iy% = iData(findex%).iy%
ix% = iData(findex%).ix%
If ix% = 0 Then GoTo GridCheckGRDConvertBadIx
If iy% = 0 Then GoTo GridCheckGRDConvertBadIy

' Profile the code below
'Dim startTime As Currency
'Tanner_SupportCode.EnableHighResolutionTimers
'Tanner_SupportCode.GetHighResTime startTime

    ' Load into longs for better perormance
    iBound& = ix%
    jBound& = iy%
        
    If (tX_Polarity% = 0) And (tY_Polarity% = 0) Then
        
        iData(findex%).gData = sarray
        
    ' Invert X
    ElseIf (tX_Polarity% <> 0) And (tY_Polarity% = 0) Then
        For jj& = 1 To jBound&
            For ii& = 1 To iBound&
                iData(findex%).gData!(ii&, jj&) = sarray!(iBound& - (ii& - 1), jj&)
            Next ii&
        Next jj&
    
    ' Invert Y
    ElseIf (tX_Polarity% = 0) And (tY_Polarity% <> 0) Then
        For jj& = 1 To jBound&
            For ii& = 1 To iBound&
                iData(findex).gData!(ii&, jj&) = sarray!(ii&, jBound& - (jj& - 1))
            Next ii&
        Next jj&
    
    ' Invert X and Y
    ElseIf (tX_Polarity% <> 0) And (tY_Polarity% <> 0) Then
        For jj& = 1 To jBound&
            For ii& = 1 To iBound&
                iData(findex).gData!(ii&, jj&) = sarray!(iBound& - (ii& - 1), jBound& - (jj& - 1))
            Next ii&
        Next jj&
    
    End If
        
'Debug.Print "GridCheckGRDConvert:"
'Tanner_SupportCode.PrintTimeTakenInMs startTime
    
' Fix units if necessary
If Default_Stage_Units$ <> gStage_Units$ Then
If method% = 1 Then         ' reading GRD
If Default_Stage_Units$ = "um" And gStage_Units$ = "mm" Then
iData(findex%).xmin# = iData(findex%).xmin# * MICRONSPERMM&
iData(findex%).xmax# = iData(findex%).xmax# * MICRONSPERMM&
iData(findex%).ymin# = iData(findex%).ymin# * MICRONSPERMM&
iData(findex%).ymax# = iData(findex%).ymax# * MICRONSPERMM&
End If

If Default_Stage_Units$ = "mm" And gStage_Units$ = "um" Then
iData(findex%).xmin# = iData(findex%).xmin# / MICRONSPERMM&
iData(findex%).xmax# = iData(findex%).xmax# / MICRONSPERMM&
iData(findex%).ymin# = iData(findex%).ymin# / MICRONSPERMM&
iData(findex%).ymax# = iData(findex%).ymax# / MICRONSPERMM&
End If

' Writing GRD
Else
If Default_Stage_Units$ = "um" And gStage_Units$ = "mm" Then
iData(findex%).xmin# = iData(findex%).xmin# / MICRONSPERMM&
iData(findex%).xmax# = iData(findex%).xmax# / MICRONSPERMM&
iData(findex%).ymin# = iData(findex%).ymin# / MICRONSPERMM&
iData(findex%).ymax# = iData(findex%).ymax# / MICRONSPERMM&
End If

If Default_Stage_Units$ = "mm" And gStage_Units$ = "um" Then
iData(findex%).xmin# = iData(findex%).xmin# * MICRONSPERMM&
iData(findex%).xmax# = iData(findex%).xmax# * MICRONSPERMM&
iData(findex%).ymin# = iData(findex%).ymin# * MICRONSPERMM&
iData(findex%).ymax# = iData(findex%).ymax# * MICRONSPERMM&
End If
End If
End If

' If reading GRD file, perform sanity check
If method% = 1 Then
If iData(findex%).xmin# = iData(findex%).xmax# Then GoTo GridCheckGRDConvertNoXExtents
If iData(findex%).ymin# = iData(findex%).ymax# Then GoTo GridCheckGRDConvertNoYExtents

If Default_X_Polarity% = 0 Then
If iData(findex%).xmin# > iData(findex%).xmax# Then GoTo GridCheckGRDConvertBadXMinMax
Else
If iData(findex%).xmin# < iData(findex%).xmax# Then GoTo GridCheckGRDConvertBadXMinMax
End If

If Default_Y_Polarity% = 0 Then
If iData(findex%).ymin# > iData(findex%).ymax# Then GoTo GridCheckGRDConvertBadYMinMax
Else
If iData(findex%).ymin# < iData(findex%).ymax# Then GoTo GridCheckGRDConvertBadYMinMax
End If
End If

' If writing GRD file, perform sanity check (GRD files are always cartesian)
If method% = 2 Then
If iData(findex%).xmin# = iData(findex%).xmax# Then GoTo GridCheckGRDConvertNoXExtents2
If iData(findex%).ymin# = iData(findex%).ymax# Then GoTo GridCheckGRDConvertNoYExtents2

If iData(findex%).xmin# > iData(findex%).xmax# Then GoTo GridCheckGRDConvertBadXMinMax2
If iData(findex%).ymin# > iData(findex%).ymax# Then GoTo GridCheckGRDConvertBadYMinMax2
End If

Exit Sub

' Errors
GridCheckGRDConvertError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "GridCheckGRDConvert"
ierror = True
Exit Sub

GridCheckGRDConvertNoXExtents:
Screen.MousePointer = vbDefault
msg$ = "Error reading GRD file (" & tfilename$ & ") does not contain valid X stage extents. It cannot be loaded"
MsgBox msg$, vbOKOnly + vbExclamation, "GridCheckGRDConvert"
ierror = True
Exit Sub

GridCheckGRDConvertNoYExtents:
Screen.MousePointer = vbDefault
msg$ = "Error reading GRD file (" & tfilename$ & ") does not contain valid Y stage extents. It cannot be loaded"
MsgBox msg$, vbOKOnly + vbExclamation, "GridCheckGRDConvert"
ierror = True
Exit Sub

GridCheckGRDConvertNoXExtents2:
Screen.MousePointer = vbDefault
msg$ = "Error writing GRD file (" & tfilename$ & ") does not contain valid X stage extents. It cannot be loaded"
MsgBox msg$, vbOKOnly + vbExclamation, "GridCheckGRDConvert"
ierror = True
Exit Sub

GridCheckGRDConvertNoYExtents2:
Screen.MousePointer = vbDefault
msg$ = "Error writing GRD file (" & tfilename$ & ") does not contain valid Y stage extents. It cannot be loaded"
MsgBox msg$, vbOKOnly + vbExclamation, "GridCheckGRDConvert"
ierror = True
Exit Sub

GridCheckGRDConvertBadXMinMax:
Screen.MousePointer = vbDefault
msg$ = "Error reading GRD file (" & tfilename$ & ") is not the correct X stage polarity (" & dstring$ & gstring$ & "). This error should not occur- please try deleting the GRDInfo.INI file or contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "GridCheckGRDConvert"
ierror = True
Exit Sub

GridCheckGRDConvertBadYMinMax:
Screen.MousePointer = vbDefault
msg$ = "Error reading GRD file (" & tfilename$ & ") is not the correct Y stage polarity (" & dstring$ & gstring$ & "). This error should not occur- please try deleting the GRDInfo.INI file or contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "GridCheckGRDConvert"
ierror = True
Exit Sub

GridCheckGRDConvertBadXMinMax2:
Screen.MousePointer = vbDefault
msg$ = "Error writing GRD file (" & tfilename$ & ") is not the correct X stage polarity (" & dstring$ & gstring$ & "). This error should not occur- please try deleting the GRDInfo.INI file or contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "GridCheckGRDConvert"
ierror = True
Exit Sub

GridCheckGRDConvertBadYMinMax2:
Screen.MousePointer = vbDefault
msg$ = "Error writing GRD file (" & tfilename$ & ") is not the correct Y stage polarity (" & dstring$ & gstring$ & "). This error should not occur- please try deleting the GRDInfo.INI file or contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "GridCheckGRDConvert"
ierror = True
Exit Sub

GridCheckGRDConvertBadIx:
Screen.MousePointer = vbDefault
msg$ = "Error writing GRD file (" & tfilename$ & ") has zero ix pixel dimension. This error should not occur- please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "GridCheckGRDConvert"
ierror = True
Exit Sub

GridCheckGRDConvertBadIy:
Screen.MousePointer = vbDefault
msg$ = "Error writing GRD file (" & tfilename$ & ") has zero iy pixel dimension. This error should not occur- please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "GridCheckGRDConvert"
ierror = True
Exit Sub

End Sub
