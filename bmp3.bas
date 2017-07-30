Attribute VB_Name = "CodeBMP3"
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

Private Const SRCCOPY& = &HCC0020 ' (DWORD) dest = source
Private Const CF_BITMAP& = 2

' GDI functions
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

' Creates a memory DC
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long

' Creates a bitmap in memory
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

' Places a GDI Object into DC, returning the previous one
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long

' Deletes a GDI Object
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

' Clipboard functions
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long

Dim tarray() As Byte
Dim narray() As Long
Dim aarray() As Single

Function BMPBytesPerLine(Width As Long, BitsPerPixel As Integer) As Long

ierror = False
On Error GoTo BMPBytesPerLineError

Select Case BitsPerPixel%
Case 1
BMPBytesPerLine& = (Width& + 7) \ 8 ' 8 pixels per byte
Case 2
BMPBytesPerLine& = (Width& + 3) \ 4 ' 4 pixels per byte
Case 4
BMPBytesPerLine& = (Width& + 1) \ 2 ' 2 pixels per byte
Case 8
BMPBytesPerLine& = Width&           ' 1 pixel per byte
Case 24
BMPBytesPerLine& = Width& * 3       ' 3 bytes per pixel
End Select

Exit Function

' Errors
BMPBytesPerLineError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "BMPBytesPerLine"
ierror = True
Exit Function

End Function

Function BMPAlignToDword(BytesPerLine As Long) As Long

ierror = False
On Error GoTo BMPAlignToDwordError

BMPAlignToDword& = (((BytesPerLine& + 3) \ 4) * 4)  ' force to 4 byte (32 bit) boundary

Exit Function

' Errors
BMPAlignToDwordError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "BMPAlignToDword"
ierror = True
Exit Function

End Function

Sub BMPDimensionByteArray(ix As Integer, iy As Integer, jarray() As Byte)
' Dimension a byte array for bitmap data (4 byte boundary image rows)

ierror = False
On Error GoTo BMPDimensionByteArrayError

Dim BPL As Long, aBPL As Long

' Calculate bytes per line
BPL& = BMPBytesPerLine&(CLng(ix%), 8)   ' assume always 8 bit images

' Align the bytes per line to a 4 byte boundary
aBPL& = BMPAlignToDword&(BPL&)

' Dimension array
ReDim jarray(1 To aBPL&, 1 To iy%) As Byte

Exit Sub

' Errors
BMPDimensionByteArrayError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "BMPDimensionByteArray"
ierror = True
Exit Sub

End Sub

Sub BMPInvertByteArray(mode As Integer, ix As Integer, iy As Integer, iarray() As Byte)
' Invert a byte variable array in X (mode=0) or Y (mode=1)

ierror = False
On Error GoTo BMPInvertByteArrayError

Dim ii As Long, jj As Long

' Invert array in X direction
If mode% = 0 Then
ReDim tarray(1 To ix%) As Byte
For jj& = 1 To iy%
For ii& = 1 To ix%
tarray(ix% - (ii& - 1)) = iarray(ii&, jj&)
Next ii&
For ii& = 1 To ix%
iarray(ii&, jj&) = tarray(ii&)
Next ii&
Next jj&
End If

' Invert array in Y direction
If mode% = 1 Then
ReDim tarray(1 To iy%) As Byte
For ii& = 1 To ix%
For jj& = 1 To iy%
tarray(iy% - (jj& - 1)) = iarray(ii&, jj&)
Next jj&
For jj& = 1 To iy%
iarray(ii&, jj&) = tarray(jj&)
Next jj&
Next ii&
End If

Exit Sub

' Errors
BMPInvertByteArrayError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "BMPInvertByteArray"
ierror = True
Exit Sub

End Sub

Sub BMPInvertLongArray(mode As Integer, ix As Integer, iy As Integer, darray() As Long)
' Invert a long variable array in X (mode=0) or Y (mode=1)

ierror = False
On Error GoTo BMPInvertLongArrayError

Dim ii As Long, jj As Long

' Invert array in X direction
If mode% = 0 Then
ReDim narray(1 To ix%) As Long
For jj& = 1 To iy%
For ii& = 1 To ix%
narray&(ix% - (ii& - 1)) = darray&(ii&, jj&)
Next ii&
For ii& = 1 To ix%
darray&(ii&, jj&) = narray&(ii&)
Next ii&
Next jj&
End If

' Invert array in Y direction
If mode% = 1 Then
ReDim narray(1 To iy%) As Long
For ii& = 1 To ix%
For jj& = 1 To iy%
narray&(iy% - (jj& - 1)) = darray&(ii&, jj&)
Next jj&
For jj& = 1 To iy%
darray&(ii&, jj&) = narray&(jj&)
Next jj&
Next ii&
End If

Exit Sub

' Errors
BMPInvertLongArrayError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "BMPInvertLongArray"
ierror = True
Exit Sub

End Sub

Sub BMPInvertSingleArray(mode As Integer, ix As Integer, iy As Integer, sarray() As Single)
' Invert a single precision variable array in X (mode=0) or Y (mode=1)

ierror = False
On Error GoTo BMPInvertSingleArrayError

Dim ii As Long, jj As Long

' Invert array in X direction
If mode% = 0 Then
ReDim aarray(1 To ix%) As Single
For jj& = 1 To iy%
For ii& = 1 To ix%
aarray!(ix% - (ii& - 1)) = sarray!(ii&, jj&)
Next ii&
For ii& = 1 To ix%
sarray!(ii&, jj&) = aarray!(ii&)
Next ii&
Next jj&
End If

' Invert array in Y direction
If mode% = 1 Then
ReDim aarray(1 To iy%) As Single
For ii& = 1 To ix%
For jj& = 1 To iy%
aarray!(iy% - (jj& - 1)) = sarray!(ii&, jj&)
Next jj&
For jj& = 1 To iy%
sarray!(ii&, jj&) = aarray!(jj&)
Next jj&
Next ii&
End If

Exit Sub

' Errors
BMPInvertSingleArrayError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "BMPInvertSingleArray"
ierror = True
Exit Sub

End Sub

Sub BMPCopyEntirePicture(ByRef objFrom As Object)
' Copy the bitmap and results of graphics methods (line, circle, print, etc) to the clipboard

ierror = False
On Error GoTo BMPCopyEntirePictureError

Dim lhDC As Long
Dim lhBMP As Long
Dim lhBMPOld As Long
Dim lWidthPixels As Long
Dim lHeightPixels As Long

' Create a DC compatible with the object we're copying from
lhDC = CreateCompatibleDC(objFrom.hdc)
    
' Create a bitmap compatible with the object we are copying from
If (lhDC <> 0) Then
lWidthPixels = objFrom.ScaleX(objFrom.ScaleWidth, objFrom.ScaleMode, vbPixels)
lHeightPixels = objFrom.ScaleY(objFrom.ScaleHeight, objFrom.ScaleMode, vbPixels)
lhBMP = CreateCompatibleBitmap(objFrom.hdc, lWidthPixels, lHeightPixels)
    
    ' Select the bitmap into the DC we have created, and store the old bitmap that was there
    If (lhBMP <> 0) Then
    lhBMPOld = SelectObject(lhDC, lhBMP)
            
    ' Copy the contents of objFrom to the bitmap
    BitBlt lhDC, 0, 0, lWidthPixels, lHeightPixels, objFrom.hdc, 0, 0, SRCCOPY
            
    ' Remove the bitmap from the DC
    SelectObject lhDC, lhBMPOld
                        
    ' Now set the clipboard to the bitmap
    OpenClipboard 0
    Sleep 100   ' need for Win 7
    EmptyClipboard
    Sleep 100   ' need for Win 7
    SetClipboardData CF_BITMAP, lhBMP
    Sleep 100   ' need for Win 7
    CloseClipboard
    Sleep 100   ' need for Win 7
    
    ' We don't delete the Bitmap here - it is now owned by the clipboard and Windows will delete it for us
    ' when the clipboard changes or the program exits.
    
    ' If bitmap could not be created it is probably because there isn't enough video memory
    Else
    msg$ = "The system could not create a large enough bitmap compatible object. There is probably not enough video memory on the system video board. Try reducing the bit depth of the video display (Desktop | Properties | Settings) from 32 to 16 and try again."
    MsgBox msg$, vbOKOnly + vbExclamation, "BMPCopyEntirePicture"
    ierror = True
    End If
        
' Clear up the device context we created
DeleteObject lhDC
End If

Exit Sub

' Errors
BMPCopyEntirePictureError:
msg$ = Error$ & ". There is probably not enough video memory on the system video board. Try reducing the bit depth of the video display (Desktop | Properties | Settings) from 32 to 16 and try again."
MsgBox Error$ & msg$, vbOKOnly + vbCritical, "BMPCopyEntirePicture"
ierror = True
Exit Sub

End Sub

Sub BMPRGB(RGBColor As Long, redvalue As Long, greenvalue As Long, bluevalue As Long)
' Convert RGB to 24 bit color

ierror = False
On Error GoTo BMPRGBError

RGBColor& = RGB(CInt(redvalue&), CInt(greenvalue&), CInt(bluevalue&))

Exit Sub

' Errors
BMPRGBError:
MsgBox Error$, vbOKOnly + vbCritical, "BMPRGB"
ierror = True
Exit Sub

End Sub

Sub BMPUnRGB(RGBColor As Long, redvalue As Long, greenvalue As Long, bluevalue As Long)
' Convert 24 bit color to RGB

ierror = False
On Error GoTo BMPUnRGBError

' RGBcolor = Format(Hex(redvalue&) & Hex(greenvalue&) & Hex(bluevalue&), "000000")
redvalue& = (RGBColor& And &HFF&)
greenvalue& = (RGBColor& And &HFF00&) \ 256     ' this is correct appraently
bluevalue& = (RGBColor& And &HFF0000) \ 65536   ' this is correct apparently
   
Exit Sub

' Errors
BMPUnRGBError:
MsgBox Error$, vbOK + vbCritical, "BMPUnRGB"
ierror = True
Exit Sub

End Sub

Sub BMPReadImageInfo(sFileName As String, m_Width As Long, m_Height As Long, m_Depth As Long, m_ImageType As Long)
' This is the sub to call to retrieve information on an image file.
'   m_ImageType = 0 = unknown
'   m_ImageType = 1 = GIF
'   m_ImageType = 2 = JPG
'   m_ImageType = 3 = PNG
'   m_ImageType = 4 = BMP
    
    ierror = False
    On Error GoTo BMPReadImageInfoError
    
    ' Only the first X bytes of the file are read into a byte array. BUFFERSIZE is X.  A larger number will use more memory and
    ' be slower.  A smaller number may not be able to decode all JPEG files.  Feel free to play with this number.
    Const BUFFERSIZE As Long = 65535
    
    Dim lPos As Long
    
    ' Byte array buffer to store part of the file
    Dim bBuf(BUFFERSIZE&) As Byte
    Dim iFN As Integer
    
    ' Set all properties to default values
    m_Width& = 0
    m_Height& = 0
    m_Depth& = 0
    m_ImageType& = 0
    
    ' Here we will load the first part of a file into a byte array the amount of the file stored here depends on the BUFFERSIZE constant
    iFN% = FreeFile
    Open sFileName For Binary As iFN%
    Get #iFN%, 1, bBuf()
    Close iFN%
    
    ' This is a PNG file
    If bBuf(0) = 137 And bBuf(1) = 80 And bBuf(2) = 78 Then
    
        m_ImageType& = 3    ' PNG
        
        ' Get bit depth
        Select Case bBuf(25)
            Case 0
            ' Greyscale
                m_Depth& = bBuf(24)
                
            Case 2
            ' RGB encoded
                m_Depth& = bBuf(24) * 3
                
            Case 3
            ' Palette based, 8 bpp
                m_Depth& = 8
                
            Case 4
            ' Greyscale with alpha
                m_Depth& = bBuf(24) * 2
                
            Case 6
            ' RGB encoded with alpha
                m_Depth& = bBuf(24) * 4
                
            ' This value is outside of it's normal range, so we'll assume that this is not a valid file
            Case Else
                m_ImageType& = 0
                
        End Select
        
        ' If the PNG image is valid then
        If m_ImageType& Then
        
            ' get the width
            m_Width& = BMPReadImageInfoMult(bBuf(19), bBuf(18))
            
            ' get the height
            m_Height& = BMPReadImageInfoMult(bBuf(23), bBuf(22))
        End If
        
    End If
    
    ' This is a GIF file
    If bBuf(0) = 71 And bBuf(1) = 73 And bBuf(2) = 70 Then
        
        m_ImageType& = 1        ' GIF
        
        ' Get the width
        m_Width& = BMPReadImageInfoMult(bBuf(6), bBuf(7))
        
        ' Get the height
        m_Height& = BMPReadImageInfoMult(bBuf(8), bBuf(9))
        
        ' Get bit depth
        m_Depth& = (bBuf(10) And 7) + 1
    End If
    
    ' This is a BMP file
    If bBuf(0) = 66 And bBuf(1) = 77 Then
    
        m_ImageType& = 4        ' BMP
        
        ' Get the width
        m_Width& = BMPReadImageInfoMult(bBuf(18), bBuf(19))
        
        ' Get the height
        m_Height& = BMPReadImageInfoMult(bBuf(22), bBuf(23))
        
        ' Get bit depth
        m_Depth& = bBuf(28)
    End If

    ' If the file is not one of the above type then check to see if it is a JPEG file
    If m_ImageType& = 0 Then
        
        ' Loop through looking for the byte sequence FF,D8,FF which marks the begining of a JPEG file. lPos will be left at the postion of the start
        Do
            If (bBuf(lPos&) = &HFF And bBuf(lPos& + 1) = &HD8 _
                 And bBuf(lPos& + 2) = &HFF) _
                 Or (lPos& >= BUFFERSIZE& - 10) Then Exit Do
            
            ' Move our pointer up
            lPos& = lPos& + 1
        Loop
        
        lPos& = lPos& + 2
        If lPos& >= BUFFERSIZE& - 10 Then Exit Sub
                
        ' Loop through the markers until we find the one starting with FF,C0 which is the block containing the image information
        Do
            ' Loop until we find the beginning of the next marker
            Do
                If bBuf(lPos&) = &HFF And bBuf(lPos& + 1) _
               <> &HFF Then Exit Do
                lPos& = lPos& + 1
                If lPos& >= BUFFERSIZE& - 10 Then Exit Sub
            Loop
            
            ' Move pointer up
            lPos& = lPos& + 1
            
            Select Case bBuf(lPos&)
                Case &HC0 To &HC3, &HC5 To &HC7, &HC9 To &HCB, _
                &HCD To &HCF
                ' We found the right block
                    Exit Do
            End Select
            
            ' Otherwise keep looking
            lPos& = lPos& + BMPReadImageInfoMult(bBuf(lPos& + 2), bBuf(lPos& + 1))
            
            ' Check for end of buffer
            If lPos& >= BUFFERSIZE& - 10 Then Exit Sub
        Loop
        
        ' If we've gotten this far it is a JPEG and we are ready to grab the information
        m_ImageType& = 2        ' it is a JPG
        
        ' Get the height
        m_Height& = BMPReadImageInfoMult(bBuf(lPos& + 5), bBuf(lPos& + 4))
        
        ' Get the width
        m_Width& = BMPReadImageInfoMult(bBuf(lPos& + 7), bBuf(lPos& + 6))
        
        ' Get the color depth
        m_Depth& = bBuf(lPos& + 8) * 8
    End If
    
Exit Sub

' Errors
BMPReadImageInfoError:
MsgBox Error$, vbOK + vbCritical, "BMPReadImageInfo"
Close iFN%
ierror = True
Exit Sub

End Sub

Private Function BMPReadImageInfoMult(lsb As Byte, msb As Byte) As Long
' Calculate the size from the passed bytes

ierror = False
On Error GoTo BMPReadImageInfoMultError

    BMPReadImageInfoMult = lsb + (msb * CLng(256))

Exit Function

' Errors
BMPReadImageInfoMultError:
MsgBox Error$, vbOK + vbCritical, "BMPReadImageInfoMult"
ierror = True
Exit Function

End Function
