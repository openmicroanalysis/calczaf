Attribute VB_Name = "CodeIMAGE4"
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

Dim cArray(0 To BIT8&) As Long

Dim pArray(0 To MAXPALETTE%, 0 To BIT8&) As Long

Sub ImageLoadPalette(Index As Integer, narray() As Long)
' Load the palette

ierror = False
On Error GoTo ImageLoadPaletteError

Dim i As Integer, j As Integer, k As Integer
Dim n As Integer, mode As Integer, p As Integer
Dim ired As Long, igreen As Long, iblue As Long
Dim ired1 As Long, igreen1 As Long, iblue1 As Long
Dim ired2 As Long, igreen2 As Long, iblue2 As Long
Dim fraction As Single
Dim tfilename As String
Dim astring As String, bstring As String

Dim iarray(0 To BIT8&) As Integer

Static initialized As Boolean

' If initialized then just load previously loaded palette from module level array
If initialized Then
For i% = 0 To BIT8&

' Load current palette to return array
narray&(i%) = pArray&(Index%, i%)

' Store current palette selection in module level
cArray&(i%) = pArray&(Index%, i%)
Next i%

Exit Sub
End If

' First check that CUSTOM.FC file exists
If Dir$(ApplicationCommonAppData$ & "CUSTOM.FC") = vbNullString Then
Call ImageMakeCustomPaletteFile
If ierror Then Exit Sub
End If

' If not initialized, load all palettes
For p% = 0 To MAXPALETTE%

' Init interpolate array
For i% = 0 To BIT8&
iarray(i%) = False
Next i%

' Load gray palette
If p% = 0 Then
For i% = 0 To BIT8&
narray&(i%) = RGB(i%, i%, i%)
Next i%
End If

' Open each color palette file
If p% > 0 Then
k% = 0  ' re-set counter

' Load from disk
If p% = 1 Then tfilename$ = ApplicationCommonAppData$ & "THERMAL.FC"
If p% = 2 Then tfilename$ = ApplicationCommonAppData$ & "RAINBOW2.FC"
If p% = 3 Then tfilename$ = ApplicationCommonAppData$ & "BLUERED.FC"
If p% = 4 Then tfilename$ = ApplicationCommonAppData$ & "CUSTOM.FC"

If Dir$(tfilename$) = vbNullString Then GoTo ImageLoadPaletteNotFound

' Open each palette file
Open tfilename$ For Input As #Temp1FileNumber%
Do Until EOF(Temp1FileNumber%) Or Trim$(astring$) = "BEGIN Items"
Line Input #Temp1FileNumber%, astring$
Loop
If Trim$(astring$) <> "BEGIN Items" Then GoTo ImageLoadPaletteNotFCFile
Line Input #Temp1FileNumber%, astring$

' Replace "=" with space to parse correctly
astring$ = Replace$(astring$, "=", " ")

' Check for interpolation text
Call MiscParseStringToString$(astring$, bstring$)
If ierror Then Exit Sub
If bstring$ <> "Interpolate" Then GoTo ImageLoadPaletteNotFCFile

' Check for interpolation type
Call MiscParseStringToString$(astring$, bstring$)
If ierror Then Exit Sub
mode% = Val(bstring$)

' Not interpolated
If mode% = 0 Then
Do Until k% = BIT8&
Line Input #Temp1FileNumber%, astring$

' Replace "=" with space to parse correctly
astring$ = Replace$(astring$, "=", " ")

' Read "Item" string
Call MiscParseStringToString$(astring$, bstring$)
If ierror Then Exit Sub
If bstring$ <> "Item" Then GoTo ImageLoadPaletteBadFormat

' Read first and second index number
Call MiscParseStringToString$(astring$, bstring$)
If ierror Then Exit Sub
j% = Val(bstring$)
If j% < 0 Or j% > BIT8& Then GoTo ImageLoadPaletteBadIndex
Call MiscParseStringToString$(astring$, bstring$)
If ierror Then Exit Sub
k% = Val(bstring$)
If k% < 0 Or k% > BIT8& Then GoTo ImageLoadPaletteBadIndex

' Read color number (24 bit)
Call MiscParseStringToString$(astring$, bstring$)
If ierror Then Exit Sub

For n% = j% To k%
narray&(n%) = Val(bstring$)
Next n%
Loop
End If

' Interpolate palette
If mode% = 1 Then
Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$
If Trim$(astring$) = "END Items" Then Exit Do

' Replace "=" with space to parse correctly
astring$ = Replace$(astring$, "=", " ")

' Read "Item" string
Call MiscParseStringToString$(astring$, bstring$)
If ierror Then Exit Sub

If bstring$ <> "Item" Then GoTo ImageLoadPaletteBadFormat

' Read first and second index number (same)
Call MiscParseStringToString$(astring$, bstring$)
If ierror Then Exit Sub
j% = Val(bstring$)
If j% < 0 Or j% > BIT8& Then GoTo ImageLoadPaletteBadIndex
Call MiscParseStringToString$(astring$, bstring$)
If ierror Then Exit Sub
k% = Val(bstring$)
If k% < 0 Or k% > BIT8& Then GoTo ImageLoadPaletteBadIndex

' Read color number (24 bit)
Call MiscParseStringToString$(astring$, bstring$)
If ierror Then Exit Sub

narray&(j%) = Val(bstring$)
iarray%(j%) = True  ' set original value index flag
Loop

' Now go through array and interpolate missing values
For i% = 0 To BIT8&

' Find first palette value
If iarray%(i%) Then

' Break into first RGB components
Call BMPUnRGB(narray&(i%), ired1&, igreen1&, iblue1&)
If ierror Then Exit Sub

' Look for next palette value
For j% = i% + 1 To BIT8&
If iarray%(j%) Then

' Break into next RGB components
Call BMPUnRGB(narray&(j%), ired2&, igreen2&, iblue2&)
If ierror Then Exit Sub

Exit For
End If
Next j%

' Interpolate these values
For n% = i% + 1 To j% - 1
fraction! = 1#
If (i% <> j%) Then fraction! = (n% - i%) / CSng(j% - i%)
ired& = ired1& + (ired2& - ired1&) * fraction!
igreen& = igreen1& + (igreen2& - igreen1&) * fraction!
iblue& = iblue1& + (iblue2& - iblue1&) * fraction!
narray&(n%) = RGB(CInt(ired&), CInt(igreen&), CInt(iblue&))
Next n%
End If

Next i%
End If

End If

' Store each palette in module level array
For i% = 0 To BIT8&
pArray&(p%, i%) = narray&(i%)
Next i%

Close #Temp1FileNumber%
Next p%

initialized = True
Exit Sub

' Errors
ImageLoadPaletteError:
MsgBox Error$, vbOKOnly + vbCritical, "ImageLoadPalette"
Close #Temp1FileNumber%
ierror = True
Exit Sub

ImageLoadPaletteNotFound:
msg$ = tfilename$ & " was not found for loading a false color palette as defined in the " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "ImageLoadPalette"
ierror = True
Exit Sub

ImageLoadPaletteNotFCFile:
msg$ = tfilename$ & " is not an FC (false color) file"
MsgBox msg$, vbOKOnly + vbExclamation, "ImageLoadPalette"
Close #Temp1FileNumber%
ierror = True
Exit Sub

ImageLoadPaletteBadFormat:
msg$ = tfilename$ & " is not formatted correctly for an FC (false color) file"
MsgBox msg$, vbOKOnly + vbExclamation, "ImageLoadPalette"
Close #Temp1FileNumber%
ierror = True
Exit Sub

ImageLoadPaletteBadIndex:
msg$ = tfilename$ & " has a bad color index"
MsgBox msg$, vbOKOnly + vbExclamation, "ImageLoadPalette"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub ImageConvertLUTtoFC(tFileLUT As String, tFileFC As String)
' Convert the passed LUT file to a FC file with the same name

ierror = False
On Error GoTo ImageConvertLUTtoFCError

Dim i As Integer
Dim astring As String
Dim RGBColor As Long
Dim n As Long, rval As Long, gval As Long, bval As Long

' Check for proper extensions
If UCase$(MiscGetFileNameExtensionOnly$(tFileLUT$)) <> ".LUT" Then GoTo ImageConvertLUTtoFCBadLUTExtension
If UCase$(MiscGetFileNameExtensionOnly$(tFileFC$)) <> ".FC" Then GoTo ImageConvertLUTtoFCBadFCExtension

If Dir$(tFileLUT$) = vbNullString Then GoTo ImageConvertLUTtoFCLUTNotFound

Open tFileLUT$ For Input As #Temp1FileNumber%
Open tFileFC$ For Output As #Temp2FileNumber%

astring$ = "False color description for CalcImage"
Print #Temp2FileNumber%, astring$
astring$ = "BEGIN Items"
Print #Temp2FileNumber%, astring$
astring$ = " Interpolate = 0"       ' all 256 values are specified
Print #Temp2FileNumber%, astring$

' Read column labels
Line Input #Temp1FileNumber%, astring$

' Loop through all values
For i% = 0 To 255
Input #Temp1FileNumber%, n&, rval&, gval&, bval&

' Convert RGB to RGB color
Call BMPRGB(RGBColor&, rval&, gval&, bval&)
If ierror Then Exit Sub
astring$ = " Item=" & Format$(i%) & " " & Format$(i%) & " " & Format$(RGBColor&)
Print #Temp2FileNumber%, astring$
Next i%

astring$ = "END Items"
Print #Temp2FileNumber%, astring$

Close #Temp1FileNumber%
Close #Temp2FileNumber%
Exit Sub

' Errors
ImageConvertLUTtoFCError:
MsgBox Error$, vbOKOnly + vbCritical, "ImageConvertLUTtoFC"
Close #Temp1FileNumber%
Close #Temp2FileNumber%
ierror = True
Exit Sub

ImageConvertLUTtoFCBadLUTExtension:
msg$ = "File " & tFileLUT$ & " does not have the proper .LUT extension."
MsgBox msg$, vbOKOnly + vbExclamation, "ImageConvertLUTtoFC"
ierror = True
Exit Sub

ImageConvertLUTtoFCBadFCExtension:
msg$ = "File " & tFileFC$ & " does not have the proper .FC extension."
MsgBox msg$, vbOKOnly + vbExclamation, "ImageConvertLUTtoFC"
ierror = True
Exit Sub

ImageConvertLUTtoFCLUTNotFound:
msg$ = "File " & tFileLUT$ & " was not found."
MsgBox msg$, vbOKOnly + vbExclamation, "ImageConvertLUTtoFC"
ierror = True
Exit Sub

End Sub

Sub ImageConvertCLRtoFC(tFileCLR As String, tFileFC As String)
' Convert the passed CLR file to a FC file with the same name

ierror = False
On Error GoTo ImageConvertCLRtoFCError

Dim astring As String
Dim RGBColor As Long
Dim rval As Long, gval As Long, bval As Long
Dim n As Integer
Dim Percent As Single

' Check for proper extensions
If UCase$(MiscGetFileNameExtensionOnly$(tFileCLR$)) <> ".CLR" Then GoTo ImageConvertCLRtoFCBadLUTExtension
If UCase$(MiscGetFileNameExtensionOnly$(tFileFC$)) <> ".FC" Then GoTo ImageConvertCLRtoFCBadFCExtension

If Dir$(tFileCLR$) = vbNullString Then GoTo ImageConvertCLRtoFCNotFound

Open tFileCLR$ For Input As #Temp1FileNumber%
Open tFileFC$ For Output As #Temp2FileNumber%

astring$ = "False color description for CalcImage"
Print #Temp2FileNumber%, astring$
astring$ = "BEGIN Items"
Print #Temp2FileNumber%, astring$
astring$ = " Interpolate = 1"       ' perform color interpolation
Print #Temp2FileNumber%, astring$

' Read column labels in CLR file "ColorMap 1 1"
Line Input #Temp1FileNumber%, astring$
If astring$ <> "ColorMap 1 1" Then GoTo ImageConvertCLRtoFCNotCLR

' Loop through all values in .CLR file
Do Until EOF(Temp1FileNumber%)
Input #Temp1FileNumber%, Percent!, rval&, gval&, bval&

' Convert percent to color index
n% = Percent! / 100# * 255#
If n% < 0# Then n% = 0
If n% > 255# Then n% = 255

' Convert RGB to RGB color
Call BMPRGB(RGBColor&, rval&, gval&, bval&)
If ierror Then Exit Sub
astring$ = " Item=" & Format$(n%) & " " & Format$(n%) & " " & Format$(RGBColor&)
Print #Temp2FileNumber%, astring$
Loop

astring$ = "END Items"
Print #Temp2FileNumber%, astring$

Close #Temp1FileNumber%
Close #Temp2FileNumber%
Exit Sub

' Errors
ImageConvertCLRtoFCError:
MsgBox Error$, vbOKOnly + vbCritical, "ImageConvertCLRtoFC"
Close #Temp1FileNumber%
Close #Temp2FileNumber%
ierror = True
Exit Sub

ImageConvertCLRtoFCBadLUTExtension:
msg$ = "File " & tFileCLR$ & " does not have the proper .CLR extension."
MsgBox msg$, vbOKOnly + vbExclamation, "ImageConvertCLRtoFC"
ierror = True
Exit Sub

ImageConvertCLRtoFCBadFCExtension:
msg$ = "File " & tFileFC$ & " does not have the proper .FC extension."
MsgBox msg$, vbOKOnly + vbExclamation, "ImageConvertCLRtoFC"
ierror = True
Exit Sub

ImageConvertCLRtoFCNotCLR:
msg$ = "File " & tFileCLR$ & " does not have the proper first line"
MsgBox msg$, vbOKOnly + vbExclamation, "ImageConvertCLRtoFC"
ierror = True
Exit Sub

ImageConvertCLRtoFCNotFound:
msg$ = "File " & tFileCLR$ & " was not found."
MsgBox msg$, vbOKOnly + vbExclamation, "ImageConvertCLRtoFC"
ierror = True
Exit Sub

End Sub

Sub ImageMakeCustomPaletteFile()
' Make default custom palette file

ierror = False
On Error GoTo ImageMakeCustomPaletteFileError

Dim tfilename As String, astring As String

Close #Temp1FileNumber%     ' make sure temp file is closed
tfilename$ = ApplicationCommonAppData$ & "CUSTOM.FC"
If Dir$(tfilename$) <> vbNullString Then Exit Sub
Open tfilename$ For Output As #Temp1FileNumber%

astring$ = "False color description for MicroImage/Probe For EPMA/CalcImage"
Print #Temp1FileNumber%, astring$
astring$ = "BEGIN Items"
Print #Temp1FileNumber%, astring$
astring$ = " Interpolate = 1"
Print #Temp1FileNumber%, astring$
astring$ = " Item=0 0 16711808 untitled"
Print #Temp1FileNumber%, astring$
astring$ = " Item=255 255 65535 untitled"
Print #Temp1FileNumber%, astring$
astring$ = " Item=120 120 8453888 untitled"
Print #Temp1FileNumber%, astring$
astring$ = "END Items"
Print #Temp1FileNumber%, astring$

Close #Temp1FileNumber%
Exit Sub

' Errors
ImageMakeCustomPaletteFileError:
MsgBox Error$, vbOKOnly + vbCritical, "ImageMakeCustomPaletteFile"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub ImageReturnPalette(narray() As Long)
' Return the current color palette

ierror = False
On Error GoTo ImageReturnPaletteError

Dim i As Integer

' Load from module level
For i% = 0 To BIT8&
narray&(i%) = cArray&(i%)
Next i%

Exit Sub

' Errors
ImageReturnPaletteError:
MsgBox Error$, vbOKOnly + vbCritical, "ImageReturnPalette"
ierror = True
Exit Sub

End Sub

