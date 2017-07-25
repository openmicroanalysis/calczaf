Attribute VB_Name = "CodeBMP"
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

' Win32 Function Declarations
Private Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal ncount As Long, lpObject As Any) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
'Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Const MAXLEVELS& = BIT8&
Const DIB_RGB_COLORS& = 0&
Const BI_RGB& = 0&

' Type - GetObjectAPI.lpObject
Private Type TypeBITMAP
  bmType       As Long    'LONG   // Specifies the bitmap type. This member must be zero.
  bmWidth      As Long    'LONG   // Specifies the width, in pixels, of the bitmap. The width must be greater than zero.
  bmHeight     As Long    'LONG   // Specifies the height, in pixels, of the bitmap. The height must be greater than zero.
  bmWidthBytes As Long    'LONG   // Specifies the number of bytes in each scan line. This value must be divisible by 2, because Windows assumes that the bit values of a bitmap form an array that is word aligned.
  bmPlanes     As Integer 'WORD   // Specifies the count of color planes.
  bmBitsPixel  As Integer 'WORD   // Specifies the number of bits required to indicate the color of a pixel.
  bmBits       As Long    'LPVOID // Points to the location of the bit values for the bitmap. The bmBits member must be a long pointer to an array of character (1-byte) values.
End Type

Private Type TypeInt4
intval As Long
End Type

Private Type TypeByt4
strval(1 To 4) As Byte
End Type

' Structures
Private Type TypeBMPHeader
ImageFileType As Integer    ' always 4D42h ("BM")
FileSize As Long            ' physical size of file
Reserved1 As Integer        ' always zero
Reserved2 As Integer        ' always zero
ImageDataOffset As Long     ' start of image data offset in bytes
End Type

Private Type TypeBMPInfo
HeaderSize As Long          ' size of header (always 40)
ImageWidth As Long          ' in pixels
ImageHeight As Long         ' in pixels
NumberOfImagePlanes As Integer  ' always 1
BitsPerPixel As Integer     ' 1, 4, 8 or 24
CompressionMethod As Long   ' 0, 1 or 2
SizeOfBitMap As Long        ' size of bitmap in bytes
HorzResolution As Long      ' pixels per meter
VertResolution As Long      ' pixels per meter
NumColorsUsed As Long       ' number of colors in image
NumSignificantColors As Long    ' number of important colors
End Type

Private Type TypeBMPRGBQuad
Blue As Byte             ' blue intensity value
Green As Byte            ' green intensity value
Red As Byte              ' red intensity value
Reserved As Byte         ' reserved (should be zero)
End Type

Private Type BITMAPINFO
    bmiHeader As TypeBMPInfo
    bmiColors As TypeBMPRGBQuad
End Type

Dim pixels() As Byte

Dim bmparray() As Byte

Sub BMPConvertIntegerArrayToByteArray(ix As Integer, iy As Integer, iarray() As Integer, jarray() As Byte)
' Converts an integer array to a byte array (normalizes the data to 0 to 255)

ierror = False
On Error GoTo BMPConvertIntegerArrayToByteArrayError

Dim i As Integer, j As Integer
Dim imax As Integer, imin As Integer
Dim itemp As Long
Dim minmax As Single

' Find minimum and maximum of data
imax% = MININTEGER%
imin% = MAXINTEGER%
For j% = 1 To iy%
For i% = 1 To ix%
If iarray%(i%, j%) > imax% Then imax% = iarray%(i%, j%)
If iarray%(i%, j%) < imin% Then imin% = iarray%(i%, j%)
Next i%
Next j%
DoEvents

' Normalize data and load into byte array (this is time consuming!)
minmax! = (imax% - imin%)
If minmax! <> 0 Then
For j% = 1 To iy%
For i% = 1 To ix%
itemp& = MAXLEVELS& * (iarray%(i%, j%) - imin%) / minmax!
If itemp& < 0 Then itemp& = 0
If itemp& > BIT8& Then itemp& = BIT8&
jarray(i%, j%) = CByte(itemp&)
Next i%
Next j%

' No data
Else
For j% = 1 To iy%
For i% = 1 To ix%
jarray(i%, j%) = 0
Next i%
Next j%
End If

Exit Sub

' Errors
BMPConvertIntegerArrayToByteArrayError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "BMPConvertIntegerArrayToByteArray"
ierror = True
Exit Sub

End Sub

Sub BMPLoadImage(tfilename As String, tImage As Image)
' Loads the passed filename to the passed image control box

ierror = False
On Error GoTo BMPLoadImageError

' Load the passed file
Screen.MousePointer = vbHourglass
tImage.Picture = LoadPicture(tfilename$)
Screen.MousePointer = vbDefault
Exit Sub

' Errors
BMPLoadImageError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "BMPLoadImage"
ierror = True
Exit Sub

End Sub

Sub BMPSaveArrayToBMPFile(ix As Integer, iy As Integer, jarray() As Byte, tfilename As String, ImagePaletteNumber As Integer, ImagePaletteArray() As Long)
' Saves a byte array to a BMP file (8 bit images only)

ierror = False
On Error GoTo BMPSaveArrayToBMPFileError

Dim tfilenumber As Integer
Dim i As Long, j As Long
Dim BPL As Long, aBPL As Long

Dim tint4 As TypeInt4
Dim tbyt4 As TypeByt4

Dim BMPHeader As TypeBMPHeader
Dim BMPInfo As TypeBMPInfo
Dim BMPRGBQuad As TypeBMPRGBQuad

Dim lineOfBytes() As Byte
Dim lineOfBytesSize As Long

' Sanity check
If ix% = 0 Or iy% = 0 Then GoTo BMPSaveArrayToBMPFileBadIxIy

' Calculate bytes per line
BPL& = BMPBytesPerLine&(CLng(ix%), 8)   ' assume always 8 bit images

' Align the bytes per line to 4 byte boundary
aBPL& = BMPAlignToDword&(BPL&)

' Open file name (binary write only)
tfilenumber% = FreeFile()
Open tfilename$ For Binary Access Write As #tfilenumber%

' Fill in BitMap header
BMPHeader.ImageFileType% = &H4D42
BMPHeader.FileSize& = 14 + 40 + 256 * 4 + (aBPL& * iy%)
BMPHeader.Reserved1% = 0
BMPHeader.Reserved2% = 0
BMPHeader.ImageDataOffset& = 14 + 40 + 256 * 4
Put #tfilenumber%, , BMPHeader

' Fill in BitMap Info
BMPInfo.HeaderSize& = 40
BMPInfo.ImageWidth& = ix%
BMPInfo.ImageHeight& = iy%
BMPInfo.NumberOfImagePlanes% = 1
BMPInfo.BitsPerPixel% = 8
BMPInfo.CompressionMethod& = 0
BMPInfo.SizeOfBitMap& = aBPL& * iy%
BMPInfo.HorzResolution& = 1024
BMPInfo.VertResolution& = 1024
BMPInfo.NumColorsUsed& = 0
BMPInfo.NumSignificantColors& = 0
Put #tfilenumber%, , BMPInfo

' Fill RGB quad (gray palette)
For i& = 0 To BIT8&
If ImagePaletteNumber% = 0 Then
BMPRGBQuad.Blue = CByte(i&)
BMPRGBQuad.Green = CByte(i&)
BMPRGBQuad.Red = CByte(i&)
BMPRGBQuad.Reserved = 0

Put #tfilenumber%, , BMPRGBQuad
Else

' Load palette from .FC file (swap red and blue bytes)
tint4.intval& = ImagePaletteArray&(i)
LSet tbyt4 = tint4
BMPRGBQuad.Blue = tbyt4.strval(3)
BMPRGBQuad.Green = tbyt4.strval(2)
BMPRGBQuad.Red = tbyt4.strval(1)
BMPRGBQuad.Reserved = 0

Put #tfilenumber%, , BMPRGBQuad
End If
Next i&

    ' Profile the code below
    'Dim startTime As Currency
    'Tanner_SupportCode.EnableHighResolutionTimers
    'Tanner_SupportCode.GetHighResTime startTime
    
    ' Create a line of bytes that is aligned on a DWORD boundary (as required by bitmaps)
    lineOfBytesSize& = aBPL&
    ReDim lineOfBytes(0 To lineOfBytesSize& - 1) As Byte

    For j& = 1 To iy%

        ' Copy this line of pixel values into the DWORD0-aligned array
        Tanner_SupportCode.CopyMemory_Strict VarPtr(lineOfBytes(0)), VarPtr(jarray(1, j&)), ix%

        ' Dump the entire array out to file at once
        Put #tfilenumber, , lineOfBytes

    Next j&
        
    'Debug.Print "BMPSaveArrayToBMPFile:"
    'Tanner_SupportCode.PrintTimeTakenInMs startTime
        
Close #tfilenumber%
Exit Sub

' Errors
BMPSaveArrayToBMPFileError:
MsgBox Error$, vbOKOnly + vbCritical, "BMPSaveArrayToBMPFile"
Close #tfilenumber%
ierror = True
Exit Sub

BMPSaveArrayToBMPFileBadIxIy:
msg$ = "Ix or Iy pixel dimensions are zero (this error should not occur, please contact Probe Software technical support)"
MsgBox msg$, vbOKOnly + vbExclamation, "BMPSaveArrayToBMPFile"
Close #tfilenumber%
ierror = True
Exit Sub

End Sub

Sub BMPLoadArrayFromBMPFile(ix As Integer, iy As Integer, jarray() As Byte, tfilename As String)
' Loads a byte array from a BMP file (pass zero ix/iy image size for any image).

ierror = False
On Error GoTo BMPLoadArrayFromBMPFileError

Dim tfilenumber As Integer
Dim i As Integer, j As Integer
Dim ixstep As Single, iystep As Single

Dim BMPHeader As TypeBMPHeader
Dim BMPInfo As TypeBMPInfo
Dim BMPRGBQuad As TypeBMPRGBQuad

Dim BPL As Long, aBPL As Long

' Open file name (binary read only)
tfilenumber% = FreeFile()
Open tfilename$ For Binary Access Read As #tfilenumber%

' Get BitMap header
Get #tfilenumber%, , BMPHeader

' Check if image data starts at expected position for 8 bit image:  Len(BMPHeader) + Len(BMPInfo) + Len(BMPRGBQuad) * 256
If BMPHeader.ImageDataOffset& <> 1078 Then GoTo BMPLoadArrayFromBMPFileBadImageDataOffset

' Get BitMap Info
Get #tfilenumber%, , BMPInfo

' Check if 8 bit BMP
If BMPInfo.BitsPerPixel <> 8 Then GoTo BMPLoadArrayFromBMPFileNot8Bits

' Check for ix and iy
If ix% = 0 And iy% = 0 Then
ix% = BMPInfo.ImageWidth&
iy% = BMPInfo.ImageHeight&
ReDim jarray(1 To ix%, 1 To iy%) As Byte
End If

' Check that image isn't too small for requested image size
If BMPInfo.ImageWidth < ix% Then GoTo BMPLoadArrayFromBMPFileSmallWidth
If BMPInfo.ImageHeight < iy% Then GoTo BMPLoadArrayFromBMPFileSmallHeight

For i% = 0 To BIT8&
Get #tfilenumber%, , BMPRGBQuad
Next i%

' Calculate bytes per line
BPL& = BMPBytesPerLine&(BMPInfo.ImageWidth&, 8)   ' assume always 8 bit images

' Align the bytes per line to 4 byte boundary
aBPL& = BMPAlignToDword&(BPL&)

' Dimension actual file image array
ReDim bmparray(1 To aBPL&, 1 To BMPInfo.ImageHeight&) As Byte

' Read image data
Get #tfilenumber%, , bmparray()
Close #tfilenumber%

' Calculate istep
ixstep! = BMPInfo.ImageWidth& / ix%
iystep! = BMPInfo.ImageHeight& / iy%

' Transfer to image array
For j% = 1 To iy%
For i% = 1 To ix%
jarray(i%, j%) = bmparray((i% * ixstep!) - (ixstep! - 1), (j% * iystep!) - (iystep! - 1))
Next i%
Next j%

Exit Sub

' Errors
BMPLoadArrayFromBMPFileError:
MsgBox Error$, vbOKOnly + vbCritical, "BMPLoadArrayFromBMPFile"
Close #tfilenumber%
ierror = True
Exit Sub

BMPLoadArrayFromBMPFileBadImageDataOffset:
msg$ = "Image data offset " & Str$(BMPHeader.ImageDataOffset&) & ", does not equal expected offset of 1078 for 8 bit BMP image"
MsgBox msg$, vbOKOnly + vbExclamation, "BMPLoadArrayFromBMPFile"
Close #tfilenumber%
ierror = True
Exit Sub

BMPLoadArrayFromBMPFileNot8Bits:
msg$ = "Disk image file is not 8 bits (it is " & Str$(BMPInfo.BitsPerPixel%) & " bits per pixel)"
MsgBox msg$, vbOKOnly + vbExclamation, "BMPLoadArrayFromBMPFile"
Close #tfilenumber%
ierror = True
Exit Sub

BMPLoadArrayFromBMPFileSmallWidth:
msg$ = "Disk image file width " & Str$(BMPInfo.ImageWidth) & ", too small for specified pixel width, " & Str$(ix%)
MsgBox msg$, vbOKOnly + vbExclamation, "BMPLoadArrayFromBMPFile"
Close #tfilenumber%
ierror = True
Exit Sub

BMPLoadArrayFromBMPFileSmallHeight:
msg$ = "Disk image file height " & Str$(BMPInfo.ImageHeight) & ", too small for specified pixel height, " & Str$(iy%)
MsgBox msg$, vbOKOnly + vbExclamation, "BMPLoadArrayFromBMPFile"
Close #tfilenumber%
ierror = True
Exit Sub

End Sub

Sub BMPLoadArrayFromBMPStream(mode As Integer, ix As Integer, iy As Integer, jarray() As Byte, tBuffer() As Byte)
' mode = 0  Dimensions a buffer to receive the BMP stream (pass actual ix and iy of desired image size)
' mode = 1  Loads a byte array from a BMP stream (pass zero ix/iy image size for any BMP stream).

ierror = False
On Error GoTo BMPLoadArrayFromBMPStreamError

Dim tfilenumber As Integer
Dim tOffset As Long
Dim tTempBMPFileName As String

Dim BMPHeader As TypeBMPHeader
Dim BMPInfo As TypeBMPInfo
Dim BMPRGBQuad As TypeBMPRGBQuad

Dim BPL As Long, aBPL As Long, tsize As Long

' Just dimension the bitmap stream buffer and exit
If mode% = 0 Then

' Calculate image data offset
tOffset& = Len(BMPHeader) + Len(BMPInfo) + Len(BMPRGBQuad) * 256

' Calculate bytes per line
BPL& = BMPBytesPerLine&(CLng(ix%), 8)   ' assume always 8 bit images

' Align the bytes per line to 4 byte boundary
aBPL& = BMPAlignToDword&(BPL&)

' Dimension byte buffer
tsize& = tOffset& + (aBPL& * iy%) + 1
ReDim tBuffer(1 To tsize&) As Byte
Exit Sub
End If

' Mode = 1, write byte buffer to temp file and read back
tfilenumber% = FreeFile()
tTempBMPFileName$ = MiscSystemCreateTempFile$("PSI")
Open tTempBMPFileName$ For Binary Access Write As #tfilenumber%
Put #tfilenumber%, , tBuffer
Close #tfilenumber%

' Load image from file
ix% = 0
iy% = 0
Call BMPLoadArrayFromBMPFile(ix%, iy%, jarray(), tTempBMPFileName$)    ' must be 8 bit image
If ierror Then
Kill tTempBMPFileName$
Exit Sub
End If

Kill tTempBMPFileName$
Exit Sub

' Errors
BMPLoadArrayFromBMPStreamError:
MsgBox Error$, vbOKOnly + vbCritical, "BMPLoadArrayFromBMPStream"
Close #tfilenumber%
ierror = True
Exit Sub

End Sub

Sub BMPConvertSingleArrayToByteArray(ix As Integer, iy As Integer, sarray() As Single, jarray() As Byte)
' Converts an single precision array to a byte array (normalizes the data to 0 to 255)
    
    ierror = False
    On Error GoTo BMPConvertSingleArrayToByteArrayError
    
    ' Profile the code below
    'Dim startTime As Currency
    'Tanner_SupportCode.EnableHighResolutionTimers
    'Tanner_SupportCode.GetHighResTime startTime
    
    Dim i As Long, j As Long
    
    ' Instead of intermixing singles and doubles, let's stick to just one data type
    Dim minmax As Single
    Dim smax As Single, smin As Single
    Dim sTemp As Single
    Dim tmpValueF As Single
    
    ' Declare globals as local variables for speed
    Dim blankValueF As Single
    blankValueF! = BLANKINGVALUE!

    Dim maxLevelsF As Single
    maxLevelsF! = MAXLEVELS&

    Dim maxBit8 As Single
    maxBit8! = BIT8&

    ' Find minimum and maximum of data
    smax! = MINSINGLE!
    smin! = MAXSINGLE!
    For j& = 1 To iy%
    For i& = 1 To ix%
    
        ' Array comparisons are expensive (because the location of the array value in memory has to be
        ' resolved multiple times).  To speed things up, cache the array value locally.
        tmpValueF = sarray!(i&, j&)
        If (tmpValueF! <> blankValueF!) Then
            If (tmpValueF! > smax!) Then smax! = tmpValueF!
            If (tmpValueF! < smin!) Then smin! = tmpValueF!
        End If
        
    Next i&
    Next j&

    ' Normalize data and load into byte array
    minmax! = (smax! - smin!)
    If (minmax! <> 0!) Then

        ' Avoid division on the inner loop (multiplying is faster)
        minmax! = 1! / minmax!
        
        For j& = 1 To iy%
        For i& = 1 To ix%

            If (sarray!(i&, j&) <> blankValueF!) Then
                
                ' Normalize
                sTemp! = maxLevelsF! * (sarray!(i&, j&) - smin!) * minmax!
                
                ' Restructure the If/Then statements to make the most likely branch the default path
                ' (e.g. stack likely outcomes under "Then", not "Else", and avoid unnecessary comparisons).
                If (sTemp! >= 0!) Then
                    If (sTemp! <= maxBit8!) Then
                        jarray(i&, j&) = sTemp!
                    Else
                        jarray(i&, j&) = maxBit8!
                    End If
                Else
                    jarray(i&, j&) = 0!
                End If

            Else
                jarray(i&, j&) = 0!       ' if BLANKINGVALUE! then zero
            End If

        Next i&
        Next j&

    ' The "no data" case can be handled specially
    Else
        For j& = 1 To iy%
        For i& = 1 To ix%
            jarray(i&, j&) = 0!
        Next i&
        Next j&
    End If
   
    'Debug.Print "BMPConvertSingleArrayToByteArray:"
    'Tanner_SupportCode.PrintTimeTakenInMs startTime
        
    Exit Sub
    
    ' Errors
BMPConvertSingleArrayToByteArrayError:
    Screen.MousePointer = vbDefault
    MsgBox Error$, vbOKOnly + vbCritical, "BMPConvertSingleArrayToByteArray"
    ierror = True
    Exit Sub

End Sub

Sub BMPConvertLongArrayToByteArray(ix As Integer, iy As Integer, iarray() As Long, jarray() As Byte)
' Converts a long array to a byte array (normalizes the data to 0 to 255)

ierror = False
On Error GoTo BMPConvertLongArrayToByteArrayError

Dim i As Integer, j As Integer
Dim imax As Long, imin As Long, itemp As Long
Dim minmax As Single

' Find minimum and maximum of data
imax& = MINLONG&
imin& = MAXLONG&
For j% = 1 To iy%
For i% = 1 To ix%
If iarray&(i%, j%) > imax& Then imax& = iarray&(i%, j%)
If iarray&(i%, j%) < imin& Then imin& = iarray&(i%, j%)
Next i%
Next j%
DoEvents

' Normalize data and load into byte array (this is time consuming!)
minmax! = (imax& - imin&)
If minmax! <> 0 Then
For j% = 1 To iy%
For i% = 1 To ix%
itemp& = MAXLEVELS& * ((iarray&(i%, j%) - imin&) / minmax!)    ' do not use integer divide
If itemp& < 0 Then itemp& = 0
If itemp& > BIT8& Then itemp& = BIT8&
jarray(i%, j%) = CByte(itemp&)
Next i%
Next j%

' No data
Else
For j% = 1 To iy%
For i% = 1 To ix%
jarray(i%, j%) = 0
Next i%
Next j%
End If

Exit Sub

' Errors
BMPConvertLongArrayToByteArrayError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "BMPConvertLongArrayToByteArray"
ierror = True
Exit Sub

End Sub

Sub BMPConvertDoubleArrayToByteArray(ix As Integer, iy As Integer, iarray() As Double, jarray() As Byte)
' Converts a zero based double precision array to a byte array (normalizes the data to 0 to 255)

ierror = False
On Error GoTo BMPConvertDoubleArrayToByteArrayError

Dim i As Integer, j As Integer
Dim imax As Long, imin As Long, itemp As Long
Dim minmax As Single

' Find minimum and maximum of data
imax& = MINLONG&
imin& = MAXLONG&
For j% = 1 To iy%
For i% = 1 To ix%
If iarray#(i% - 1, j% - 1) > imax& Then imax& = iarray#(i% - 1, j% - 1)
If iarray#(i% - 1, j% - 1) < imin& Then imin& = iarray#(i% - 1, j% - 1)
Next i%
Next j%
DoEvents

' Normalize data and load into byte array (this is time consuming!)
minmax! = (imax& - imin&)
If minmax! <> 0 Then
For j% = 1 To iy%
For i% = 1 To ix%
itemp& = MAXLEVELS& * ((iarray#(i% - 1, j% - 1) - imin&) / minmax!)
If itemp& < 0 Then itemp& = 0
If itemp& > BIT8& Then itemp& = BIT8&
jarray(i%, j%) = CByte(itemp&)
Next i%
Next j%

' No data
Else
For j% = 1 To iy%
For i% = 1 To ix%
jarray(i%, j%) = 0
Next i%
Next j%
End If

Exit Sub

' Errors
BMPConvertDoubleArrayToByteArrayError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "BMPConvertDoubleArrayToByteArray"
ierror = True
Exit Sub

End Sub

Sub BMPGetBitmapInfo(ByVal hBitmap As Long, Return_Width As Long, Return_Height As Long, _
 Return_BitsPerPixel As Integer, Return_Size As Double, Return_PointerToBits As Long)
' This function takes a given picture and finds out all possible information about it and returns the
' results. NOTE : This function only works with BITMAPs and DIBs (Device Independant Bitmaps)
'
' hBITMAP                 Handle to the bitmap to get the information from.
' Return_Width            Optional. Returns the width (in pixels) of the picture.
' Return_Height           Optional. Returns the height (in pixels) of the picture.
' Return_BitsPerPixel     Optional. Returns the color depth of the picture in the form of "BitsPerPixel"
' Return_Size             Optional. Returns the approximate size of the picture (assuming it's RGB)
' Return_PointerToBits    Optional. Returns a memory pointer to the location of the BITMAP BITS that make
'                         up the specified image.  You can use the "CopyMemory" API to copy the BITS to a
'                         BYTE ARRAY.

ierror = False
On Error GoTo BMPGetBitmapInfoError
  
Dim tBMP As TypeBITMAP
  
' Clear the return variables
Return_Height& = 0
Return_Width& = 0
Return_BitsPerPixel% = 0
Return_Size# = 0
Return_PointerToBits& = 0
  
' Check that there's a valid input
If hBitmap = 0 Then GoTo BMPGetBitmapInfoBadHandle
  
' Get the information
If GetObjectAPI(hBitmap, Len(tBMP), tBMP) = 0 Then GoTo BMPGetBitmapInfoBadBitMap
  
' Return the information
With tBMP
    Return_Width = .bmWidth
    Return_Height = .bmHeight
    Return_BitsPerPixel = (.bmBitsPixel * .bmPlanes)
    Return_Size = ((.bmWidth * 3 + 3) And &HFFFFFFFC) * .bmHeight
    Return_PointerToBits = .bmBits
End With
  
' Copy the mask DIB bits directly into the array
'CopyMemory jarray(), ByVal Return_PointerToBits, Return_Width * Return_Height
  
Exit Sub

' Errors
BMPGetBitmapInfoError:
MsgBox Error$, vbOK + vbCritical, "BMPGetBitmapInfo"
ierror = True
Exit Sub

BMPGetBitmapInfoBadHandle:
msg$ = "Invalid handle to bitmap object"
MsgBox msg$, vbOK + vbCritical, "BMPGetBitmapInfo"
ierror = True
Exit Sub

BMPGetBitmapInfoBadBitMap:
msg$ = "Invalid bitmap object"
MsgBox msg$, vbOK + vbCritical, "BMPGetBitmapInfo"
ierror = True
Exit Sub

End Sub

Sub BMPConvertBitmapToLongArray(tPicture As PictureBox, nX As Long, nY As Long, larray() As Long)
' Procedure to save a DIB to a long array

ierror = False
On Error GoTo BMPConvertBitmapToLongArrayError

Dim i As Integer, j As Integer, bytes_per_scanLine As Integer

Dim pixels() As Byte

Dim bitmap_info As BITMAPINFO

' Prepare the bitmap description
With bitmap_info.bmiHeader
    .HeaderSize& = 40
    .ImageWidth& = nX&
    .ImageHeight& = nY&
    .NumberOfImagePlanes% = 1
    .BitsPerPixel% = 32
    .CompressionMethod = BI_RGB
    bytes_per_scanLine% = ((((.ImageWidth * .BitsPerPixel%) + 31) \ 32) * 4)
    .SizeOfBitMap = bytes_per_scanLine% * nY&
    End With

' Load the bitmap's data (by referring to the Image property, a persistent handle)
ReDim pixels(1 To 4, 1 To nX&, 1 To nY&)
GetDIBits tPicture.hdc, tPicture.Image, CLng(0), nY&, pixels(1, 1, 1), bitmap_info, DIB_RGB_COLORS

' Load the long array
For j% = 1 To nY&
For i% = 1 To nX&
larray&(i%, j%) = RGB(pixels(1, i%, j%), pixels(2, i%, j%), pixels(3, i%, j%))

Next i%
Next j%

Exit Sub

' Errors
BMPConvertBitmapToLongArrayError:
MsgBox Error$, vbOK + vbCritical, "BMPConvertBitmapToLongArray"
ierror = True
Exit Sub

End Sub

Sub BMPMakeGray(ByVal picColor As PictureBox)
' Convert a color image to gray scale

ierror = False
On Error GoTo BMPMakeGrayError

Dim bitmap_info As BITMAPINFO
Dim pixels() As Byte
Dim bytes_per_scanLine As Long
Dim pad_per_scanLine As Long
Dim X As Integer
Dim Y As Integer
Dim ave_color As Byte
Dim nBytes As Long

Const pixR& = 1
Const pixG& = 2
Const pixB& = 3

    ' Prepare the bitmap description
    With bitmap_info.bmiHeader
        .HeaderSize = 40
        .ImageWidth = picColor.ScaleWidth
        ' Use negative height to scan top-down
        .ImageHeight = -picColor.ScaleHeight
        .NumberOfImagePlanes = 1
        .BitsPerPixel = 32
        .CompressionMethod = BI_RGB
        bytes_per_scanLine = ((((.ImageWidth * .BitsPerPixel) + 31) \ 32) * 4)
        pad_per_scanLine = bytes_per_scanLine - (((.ImageWidth * .BitsPerPixel) + 7) \ 8)
        .SizeOfBitMap = bytes_per_scanLine * Abs(.ImageHeight)
    End With

' Check if too large to allocate pixel array
If CSng(picColor.ScaleWidth) * CSng(picColor.ScaleHeight) * 4# > MAXLONG& Then Exit Sub

    ' Load the bitmap's data
    nBytes& = 4
    ReDim pixels(1 To nBytes&, 1 To picColor.ScaleWidth, 1 To picColor.ScaleHeight)
    GetDIBits picColor.hdc, picColor.Image, 0, picColor.ScaleHeight, pixels(1, 1, 1), bitmap_info, DIB_RGB_COLORS

    ' Modify the pixels
    For Y = 1 To picColor.ScaleHeight
        For X = 1 To picColor.ScaleWidth
            ave_color = CByte((CInt(pixels(pixR, X, Y)) + pixels(pixG, X, Y) + pixels(pixB, X, Y)) \ 3)
            pixels(pixR, X, Y) = ave_color
            pixels(pixG, X, Y) = ave_color
            pixels(pixB, X, Y) = ave_color
        Next X
    Next Y

    ' Display the result
    SetDIBits picColor.hdc, picColor.Image, 0, picColor.ScaleHeight, pixels(1, 1, 1), bitmap_info, DIB_RGB_COLORS
    picColor.Picture = picColor.Image
    
Exit Sub

' Errors
BMPMakeGrayError:
MsgBox Error$, vbOK + vbCritical, "BMPMakeGray"
ierror = True
Exit Sub

End Sub

Sub BMPMakeColored(ByVal picColor As PictureBox, tRGB As Long)
' Convert everything except black colors to passed color (does not work on non RGB BMPs?)

ierror = False
On Error GoTo BMPMakeColoredError

Dim bitmap_info As BITMAPINFO
Dim bytes_per_scanLine As Long
Dim pad_per_scanLine As Long
Dim X As Integer
Dim Y As Integer
Dim tR As Long
Dim tG As Long
Dim tB As Long

Const pixR& = 1
Const pixG& = 2
Const pixB& = 3

    ' Prepare the bitmap description
    With bitmap_info.bmiHeader
        .HeaderSize = 40
        .ImageWidth = picColor.ScaleWidth
        ' Use negative height to scan top-down
        .ImageHeight = -picColor.ScaleHeight
        .NumberOfImagePlanes = 1
        .BitsPerPixel = 32
        .CompressionMethod = BI_RGB
        bytes_per_scanLine = ((((.ImageWidth * .BitsPerPixel) + 31) \ 32) * 4)
        pad_per_scanLine = bytes_per_scanLine - (((.ImageWidth * .BitsPerPixel) + 7) \ 8)
        .SizeOfBitMap = bytes_per_scanLine * Abs(.ImageHeight)
    End With

    ' Load the bitmap's data
    ReDim pixels(1 To 4, 1 To picColor.ScaleWidth, 1 To picColor.ScaleHeight)
    GetDIBits picColor.hdc, picColor.Image, 0, picColor.ScaleHeight, pixels(1, 1, 1), bitmap_info, DIB_RGB_COLORS

    ' Modify the non black pixels
    For Y% = 1 To picColor.ScaleHeight
        For X% = 1 To picColor.ScaleWidth
        If pixels(pixR, X, Y) <> 0 And pixels(pixG, X, Y) <> 0 And pixels(pixB, X, Y) <> 0 Then
            Call BMPUnRGB(tRGB&, tR&, tG&, tB&)
            pixels(pixR, X, Y) = CByte(tR)
            pixels(pixG, X, Y) = CByte(tG)
            pixels(pixB, X, Y) = CByte(tB)
        End If
        Next X%
    Next Y%

    ' Display the result
    SetDIBits picColor.hdc, picColor.Image, 0, picColor.ScaleHeight, pixels(1, 1, 1), bitmap_info, DIB_RGB_COLORS
    picColor.Picture = picColor.Image
    
Exit Sub

' Errors
BMPMakeColoredError:
MsgBox Error$, vbOK + vbCritical, "BMPMakeColored"
ierror = True
Exit Sub

End Sub

Sub BMPSaveArrayToBMPFile24Bit(ix As Long, iy As Long, narray() As Long, tfilename As String)
' Saves a byte array to a BMP file (24 bit images only)

ierror = False
On Error GoTo BMPSaveArrayToBMPFile24BitError

Dim tfilenumber As Integer
Dim i As Integer, j As Integer
Dim BPL As Long, aBPL As Long
Dim vred As Long, vgreen As Long, vblue As Long

Dim BMPHeader As TypeBMPHeader
Dim BMPInfo As TypeBMPInfo

' Calculate bytes per line
BPL& = BMPBytesPerLine&(ix&, 24)   ' assume always 24 bit images

' Align the bytes per line to 4 byte boundary
aBPL& = BMPAlignToDword&(BPL&)

' Open file name (binary write only)
tfilenumber% = FreeFile()
Open tfilename$ For Binary Access Write As #tfilenumber%

' Fill in BitMap header
BMPHeader.ImageFileType% = &H4D42
BMPHeader.FileSize& = 14 + 40 + (aBPL& * iy&)
BMPHeader.Reserved1% = 0
BMPHeader.Reserved2% = 0
BMPHeader.ImageDataOffset& = 14 + 40
Put #tfilenumber%, , BMPHeader

' Fill in BitMap Info
BMPInfo.HeaderSize& = 40
BMPInfo.ImageWidth& = ix&
BMPInfo.ImageHeight& = iy&
BMPInfo.NumberOfImagePlanes% = 1
BMPInfo.BitsPerPixel% = 24
BMPInfo.CompressionMethod& = 0
BMPInfo.SizeOfBitMap& = aBPL& * iy&
BMPInfo.HorzResolution& = 1024
BMPInfo.VertResolution& = 1024
BMPInfo.NumColorsUsed& = 0
BMPInfo.NumSignificantColors& = 0
Put #tfilenumber%, , BMPInfo

' Save to file
For j% = 1 To iy&
For i% = 1 To ix&

' Extract RGB values
Call BMPUnRGB(narray&(i%, j%), vred&, vgreen&, vblue&)
If ierror Then Exit Sub

' Output the RGB values.  Note the order : blue, green, red
Put #tfilenumber%, , CByte(vblue&)
Put #tfilenumber%, , CByte(vgreen&)
Put #tfilenumber%, , CByte(vred&)

Next i%

' Load extra bytes per line
If aBPL& <> BPL& Then
For i% = ix& + 1 To aBPL&
Put #tfilenumber%, , CByte(0)
Next i%
End If

Next j%

Close #tfilenumber%
Exit Sub

' Errors
BMPSaveArrayToBMPFile24BitError:
MsgBox Error$, vbOKOnly + vbCritical, "BMPSaveArrayToBMPFile24Bit"
Close #tfilenumber%
ierror = True
Exit Sub

End Sub


