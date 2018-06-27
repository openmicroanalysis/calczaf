Attribute VB_Name = "CodePictureSnapOpenClose"
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

Dim PictureSnapImageWidth As Single
Dim PictureSnapImageHeight As Single

Dim lastmode As Integer

Dim GridImageData(1 To 1) As TypeImageData
Dim iarray() As Byte

Dim ImagePoints As Long
Dim ImageXdata() As Single, ImageYdata() As Single, ImageZdata() As Single
Dim ImageIdata() As Integer, ImageNData() As Integer, ImageSData() As Integer    ' sample types, line (row) numbers, and sample numbers
Dim ImageSNdata() As String     ' sample names

Sub PictureSnapPrintOrClipboard(mode As Integer, tForm As Form)
' Save the current image as various
'   mode = 1 print to default printer
'   mode = 2 copy to clipboard (without graphics objects)
'   mode = 3 copy to clipboard (with graphics objects)
'   mode = 4 copy to BMP file with graphics objects

ierror = False
On Error GoTo PictureSnapPrintOrClipboardError

Dim ixiy As Single
Dim tfilename As String

' Print to default printer
If mode% = 1 Then
ixiy! = FormPICTURESNAP.Picture2.ScaleWidth / FormPICTURESNAP.Picture2.ScaleHeight
Call BMPPrintDiagram(FormPICTURESNAP.Picture2, FormPICTURESNAP.Picture3, CSng(0.5), CSng(0.5), CSng(7 * ixiy!), CSng(7#))     '  Picture3 control is in Pixel ScaleMode
If ierror Then Exit Sub

Set FormPICTURESNAP.Picture2 = LoadPicture(PictureSnapFilename$)    ' to prevent trailing
End If

' Clipboard (does not copy graphics methods)
If mode% = 2 Then
Clipboard.Clear
Sleep (200)     ' need for Win7 clipboard issues
Clipboard.SetData FormPICTURESNAP.Picture2.Picture
If ierror Then Exit Sub
End If

' Clipboard (use special function to save graphics methods)
If mode% = 3 Then
Call BMPCopyEntirePicture(FormPICTURESNAP.Picture2)    ' does not work on "hidden" bitmaps
If ierror Then Exit Sub

Set FormPICTURESNAP.Picture2 = LoadPicture(PictureSnapFilename$)    ' to prevent trailing
End If

' Save to BMP file via clipboard to save drawing objects
If mode% = 4 Then
Call BMPCopyEntirePicture(FormPICTURESNAP.Picture2)    ' does not work on "hidden" bitmaps
If ierror Then Exit Sub

' Check for a bitmap in the system clipboard
If Clipboard.GetFormat(vbCFBitmap) Then
FormPICTURESNAP.Picture3.Picture = Clipboard.GetData(vbCFBitmap)     '  Picture3 control is in Pixel ScaleMode
Else
msg$ = "There is no bitmap available in the system clipboard to save. Check that there is enough memory in the system for large bitmap compatible objects."
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapPrintOrClipboard"
ierror = True
Exit Sub
End If

' Ask user for file
tfilename$ = MiscGetFileNameNoExtension$(PictureSnapFilename$) & "_Annotations.BMP"
Call IOGetFileName(Int(1), "BMP", tfilename$, tForm)
If ierror Then Exit Sub

' Save to BMP file from Picture3 control
SavePicture FormPICTURESNAP.Picture3, tfilename$     '  Picture3 control is in Pixel ScaleMode
msg$ = "Image with drawing objects saved to " & tfilename$
MsgBox msg$, vbOKOnly + vbInformation, "PictureSnapPrintOrClipboard"
ierror = True
End If

Exit Sub

' Errors
PictureSnapPrintOrClipboardError:
msg$ = Error$
If Err = VB_OutOfMemory& Then msg$ = msg$ & ". The system could not create a large enough bitmap compatible object. There is probably not enough video memory on the system video board. Try reducing the bit depth of the video display (Desktop | Properties | Settings) from 32 to 16 and try again."
MsgBox msg$, vbOKOnly + vbCritical, "PictureSnapPrintOrClipboard"
ierror = True
Exit Sub

End Sub

Sub PictureSnapSaveAsBMP()
' Save the current image and calibration as a new BMP image and ACQ file

ierror = False
On Error GoTo PictureSnapSaveAsBMPError

Dim tfilename As String, tfilename2 As String
Dim response As Integer

' Save with ACQ file calibration and annotations
tfilename$ = MiscGetFileNameNoExtension$(PictureSnapFilename$) & ".bmp"

' Ask use for new BMP filename (allow for existing file to check for overwrite of both image and calibration file below)
Call IOGetFileName(Int(0), "BMP", tfilename$, FormPICTURESNAP)
If ierror Then Exit Sub

' Load new calibration file name
tfilename2$ = MiscGetFileNameNoExtension$(tfilename$) & ".acq"

' Check if file exists
If Dir$(tfilename$) <> vbNullString Or Dir$(tfilename2$) <> vbNullString Then
msg$ = "The specified image and/or calibration file(s) already exists. Do you want to overwrite it?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "PictureSnapSaveAsBMP")
If response% = vbCancel Then Exit Sub
If Dir$(tfilename$) <> vbNullString Then Kill tfilename$
If Dir$(tfilename2$) <> vbNullString Then Kill tfilename2$
DoEvents
End If

' Save image to new file
SavePicture FormPICTURESNAP.Picture2, tfilename$     ' does not save graphics methods

' Save ACQ file
FileCopy MiscGetFileNameNoExtension$(PictureSnapFilename$) & ".ACQ", tfilename2$

' Re-load new image file into PictureSnapApp
Call PictureSnapFileOpen(Int(0), tfilename$, FormPICTURESNAP)
If ierror Then Exit Sub

Exit Sub

' Errors
PictureSnapSaveAsBMPError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapSaveAsBMP"
ierror = True
Exit Sub

End Sub

Sub PictureSnapFileOpen(mode As Integer, tfilename As String, tForm As Form)
' Load the PictureSnap form
'  mode = 0 file is specified in tfilename$
'  mode = 1 open BMP, GIF or JPG
'  mode = 2 open GRD (special treatment)
'  tfilename$ = file to open automatically (if not blank)

ierror = False
On Error GoTo PictureSnapFileOpenError

Dim astring As String

Dim gX_Polarity As Integer, gY_Polarity As Integer
Dim gStage_Units As String

Dim m_Width As Long, m_Height As Long, m_Depth As Long, m_ImageType As Long

' Unload the calibration and full view forms in case loaded
Unload FormPICTURESNAP2
Unload FormPICTURESNAP3

' Set FormPICTURESNAP scroll bars to upper left
FormPICTURESNAP.HScroll1.value = 0
FormPICTURESNAP.VScroll1.value = 0

' Get existing filename from user
If mode% > 0 Then

' Load image file (bmp, gif, jpg)
If mode% = 1 Then
tfilename$ = vbNullString
If Trim$(PictureSnapFilename$) <> vbNullString Then tfilename$ = PictureSnapFilename$
Call IOGetFileName(Int(2), "IMG", tfilename$, tForm)
If ierror Then Exit Sub
If UCase$(MiscGetFileNameExtensionOnly$(tfilename$)) <> ".BMP" And UCase$(MiscGetFileNameExtensionOnly$(tfilename$)) <> ".GIF" And UCase$(MiscGetFileNameExtensionOnly$(tfilename$)) <> ".JPG" Then GoTo PictureSnapfileOpenWrongExtension

' Set global while image is loading to prevent mouse cursor convert coordinate errors
PictureSnapCalibrated = False
End If

' Load GRD file
If mode% = 2 Then
tfilename$ = vbNullString
If Trim$(PictureSnapFilename$) <> vbNullString Then tfilename$ = MiscGetFileNameNoExtension$(PictureSnapFilename$) & ".grd"
Call IOGetFileName(Int(2), "GRD", tfilename$, tForm)
If ierror Then Exit Sub

' Open and convert grid file
If Not MiscStringsAreSame(MiscGetFileNameExtensionOnly$(tfilename$), ".GRD") Then GoTo PictureSnapFileOpenNotGRD

' Set global while image is loading to prevent mouse cursor convert coordinate errors
PictureSnapCalibrated = False

Call PictureSnapFileOpenGrid(tfilename)
If ierror Then Exit Sub
End If

' Filename was passed, so check for file to exist
Else
If Dir$(tfilename$) = vbNullString Then GoTo PictureSnapFileOpenNotFound
End If

' Reset globals
PictureSnapFilename$ = vbNullString
PictureSnapCalibrated = False           ' reset not calibrated
PictureSnapCalibrationSaved = False     ' reset calibration not saved
Set FormPICTURESNAP.Picture2 = LoadPicture(vbNullString)
FormPICTURESNAP.Caption = "PictureSnapApp [" & PictureSnapFilename$ & "]"

' Check for existing GRD or ACQ info
Call GridCheckGRDInfo(tfilename$, gX_Polarity%, gY_Polarity%, gStage_Units$)
If ierror Then Exit Sub

' Check if stage orientation is different than current configuration
If Default_X_Polarity% = 0 And Default_Y_Polarity% = 0 Then
astring$ = "Cameca"
Else
astring$ = "JEOL"
End If

If Default_X_Polarity% <> gX_Polarity% Or Default_Y_Polarity% <> gY_Polarity% Then
msg$ = "The image file " & tfilename$ & " has a different stage orientation than the current stage configuration (" & astring$ & ")." & vbCrLf & vbCrLf
msg$ = msg$ & "Please load an image file with the correct stage orientation or change the current stage configuration from the File menu and try again."
MsgBox msg$, vbOKOnly + vbInformation, "PictureSnapFileOpen"
Screen.MousePointer = vbDefault
ierror = True
Exit Sub
End If

' Check if stage units is different than current configuration
If Default_Stage_Units$ = "mm" Then
astring$ = "mm"
Else
astring$ = "um"
End If

If Default_Stage_Units$ <> gStage_Units$ Then
msg$ = "The image file " & tfilename$ & " has a different stage units than the current stage configuration (" & astring$ & ")." & vbCrLf & vbCrLf
msg$ = msg$ & "Please load an image file with the correct stage units or change the stage configuration from the File menu and try again."
MsgBox msg$, vbOKOnly + vbInformation, "PictureSnapFileOpen"
Screen.MousePointer = vbDefault
Exit Sub
End If

' Minimize form to force resize
FormPICTURESNAP.WindowState = vbMinimized
DoEvents

' Load the file to image control
Screen.MousePointer = vbHourglass
Set FormPICTURESNAP.Picture2 = LoadPicture(tfilename$)

' Restore to re-size
FormPICTURESNAP.WindowState = vbNormal
Screen.MousePointer = vbDefault

' Set image file opened
PictureSnapFilename$ = tfilename$

' Update form caption
FormPICTURESNAP.Caption = "PictureSnap [" & tfilename$ & "]"

' Check if a calibration file already exists and load if found
Call PictureSnapLoadCalibration
If ierror Then Exit Sub

' Enable output menus if modeless mode
Call PictureSnapEnableDisable
If ierror Then Exit Sub

' Store image width and heigth after loading for setting aspect ratio in full view window
If PictureSnapFilename$ <> vbNullString Then
PictureSnapImageWidth! = FormPICTURESNAP.Picture2.ScaleWidth
PictureSnapImageHeight! = FormPICTURESNAP.Picture2.ScaleHeight
End If

' Load full view window if visible (and not called from ImageSaveAs which causes non-modal when modal loaded error)
If mode% <> 0 And FormPICTURESNAP3.Visible And Not PictureSnapWindowIsModeless Then
Call PictureSnapLoadFullWindow
If ierror Then Exit Sub
End If

' Move stage to move main PictureSnap window to current location
If RealTimeMode And PictureSnapCalibrated Then
If MiscMotorInBounds(XMotor%, RealTimeMotorPositions!(XMotor%)) And MiscMotorInBounds(YMotor%, RealTimeMotorPositions!(YMotor%)) And MiscMotorInBounds(ZMotor%, RealTimeMotorPositions!(ZMotor%)) Then
Call PictureSnapMoveToCalibrationPoint(RealTimeMotorPositions!(XMotor%), RealTimeMotorPositions!(YMotor%), RealTimeMotorPositions!(ZMotor%))
If ierror Then Exit Sub
End If
End If

Exit Sub

' Errors
PictureSnapFileOpenError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapFileOpen"
ierror = True
Exit Sub

PictureSnapfileOpenWrongExtension:
msg$ = "The image file " & tfilename$ & " was not a .BMP, .GIF or .JPG file. Please try with another image file."
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapFileOpen"
ierror = True
Exit Sub

PictureSnapFileOpenNotFound:
msg$ = "The image file " & tfilename$ & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapFileOpen"
ierror = True
Exit Sub

PictureSnapFileOpenNotGRD:
msg$ = "The specified file " & tfilename$ & " was not a GRD file."
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapFileOpen"
ierror = True
Exit Sub

PictureSnapFileOpenUnknownType:
msg$ = "The specified image file " & tfilename$ & ", is an unknown image type. Please try again with a different image."
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapFileOpen"
ierror = True
Exit Sub

End Sub

Sub PictureSnapFileOpenGrid(tfilename As String)
' Open a grid file

ierror = False
On Error GoTo PictureSnapFileOpenGridError

Dim astring As String

Dim gridversion As Single
Dim tfilename2 As String

Dim gX_Polarity As Integer, gY_Polarity As Integer
Dim gStage_Units As String

Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double

' Check grid file version
gridversion! = GridFileGetVersion!(tfilename$)
If ierror Then Exit Sub
DoEvents

' Read grid file
If gridversion! = 6 Then
Screen.MousePointer = vbHourglass
Call GridFileReadWrite(Int(1), Int(1), GridImageData(), tfilename$)
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

Else
Screen.MousePointer = vbHourglass
Call GridFileReadWrite2(Int(1), Int(1), GridImageData(), tfilename$)
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End If

' Check for existing GRD or ACQ info
Call GridCheckGRDInfo(tfilename$, gX_Polarity%, gY_Polarity%, gStage_Units$)
If ierror Then Exit Sub

' Check if stage orientation is different than current configuration
If Default_X_Polarity% = 0 And Default_Y_Polarity% = 0 Then
astring$ = "Cameca"
Else
astring$ = "JEOL"
End If

If Default_X_Polarity% <> gX_Polarity% Or Default_Y_Polarity% <> gY_Polarity% Then
msg$ = "The image file " & tfilename$ & " has a different stage orientation than the current stage configuration (" & astring$ & ")." & vbCrLf & vbCrLf
msg$ = msg$ & "Please load an image file with the correct stage orientation or change the current stage configuration from the File menu and try again."
MsgBox msg$, vbOKOnly + vbInformation, "PictureSnapFileOpenGrid"
Screen.MousePointer = vbDefault
ierror = True
Exit Sub
End If

' Check if stage units is different than current configuration
If Default_Stage_Units$ = "mm" Then
astring$ = "mm"
Else
astring$ = "um"
End If

If Default_Stage_Units$ <> gStage_Units$ Then
msg$ = "The image file " & tfilename$ & " has a different stage units than the current stage configuration (" & astring$ & ")." & vbCrLf & vbCrLf
msg$ = msg$ & "Please load an image file with the correct stage units or change the stage configuration from the File menu and try again."
MsgBox msg$, vbOKOnly + vbInformation, "PictureSnapFileOpenGrid"
Screen.MousePointer = vbDefault
Exit Sub
End If

' Save to byte array
ReDim iarray(1 To GridImageData(1).ix%, 1 To GridImageData(1).iy%)
Screen.MousePointer = vbHourglass
Call BMPConvertSingleArrayToByteArray(GridImageData(1).ix%, GridImageData(1).iy%, GridImageData(1).gData!(), iarray())
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

' Load palette
Call ImageLoadPalette(ImagePaletteNumber%, ImagePaletteArray())
If ierror Then Exit Sub

' Save to BMP file
tfilename2$ = MiscGetFileNameNoExtension$(tfilename$) & ".BMP"
Screen.MousePointer = vbHourglass
Call BMPSaveArrayToBMPFile(GridImageData(1).ix%, GridImageData(1).iy%, iarray(), tfilename2$, ImagePaletteNumber%, ImagePaletteArray&())
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

' Save BMP to original name
tfilename$ = tfilename2$

' Assume no unit conversions to begin with
xmin# = GridImageData(1).xmin#
xmax# = GridImageData(1).xmax#
ymin# = GridImageData(1).ymin#
ymax# = GridImageData(1).ymax#

' Load stage extents
FormPICTURESNAP2.TextXStage1.Text = xmin#
FormPICTURESNAP2.TextXStage2.Text = xmax#
FormPICTURESNAP2.TextYStage1.Text = ymin#
FormPICTURESNAP2.TextYStage2.Text = ymax#

' Load pixel coordinates
FormPICTURESNAP2.TextXPixel1.Text = 0
FormPICTURESNAP2.TextXPixel2.Text = GridImageData(1).ix% * Screen.TwipsPerPixelX
FormPICTURESNAP2.TextYPixel1.Text = GridImageData(1).iy% * Screen.TwipsPerPixelY    ' always flip for Y for BMP load
FormPICTURESNAP2.TextYPixel2.Text = 0

' Create .ACQ file
PictureSnapMode% = 0    ' two calibration coordinates only
PictureSnapFilename$ = tfilename$
Call PictureSnapCalibrate(Int(1))
If ierror Then
PictureSnapFilename$ = vbNullString
Exit Sub
End If

Exit Sub

' Errors
PictureSnapFileOpenGridError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapFileOpenGrid"
ierror = True
Exit Sub

End Sub

Sub PictureSnapConvertPrbImgToACQ(tfilename As String)
' Convert a PrbImg (Probe Image) file to an ACQ registration file
' Example section of PrbImag [Regsitration] section
'[Registration]
'X1Pixel = 0
'Y1Pixel = 0
'X2Pixel = 255
'Y2Pixel = 255
'X3Pixel = 0
'Y3Pixel = 255
'X1Real = -1.519
'Y1Real = 26.097
'Z1Real = 0.185
'X2Real = -1.009
'Y2Real = 25.842
'Z2Real = 0.185
'Z3Real = 0.185

ierror = False
On Error GoTo PictureSnapConvertPrbImgToACQError

Dim ixmin As Long, iymin As Long, ixmax As Long, iymax As Long
Dim xmin As Single, ymin As Single, xmax As Single, ymax As Single
Dim tvalue As Single

' Load pixel coordinates
Call InitINIReadWriteScaler(Int(1), tfilename$, "Registration", "X1Pixel", tvalue!)
If ierror Then Exit Sub
ixmin& = tvalue!
Call InitINIReadWriteScaler(Int(1), tfilename$, "Registration", "Y1Pixel", tvalue!)
If ierror Then Exit Sub
iymin& = tvalue!

Call InitINIReadWriteScaler(Int(1), tfilename$, "Registration", "X2Pixel", tvalue!)
If ierror Then Exit Sub
ixmax& = tvalue!
Call InitINIReadWriteScaler(Int(1), tfilename$, "Registration", "Y2Pixel", tvalue!)
If ierror Then Exit Sub
iymax& = tvalue!

' Load real world coordinates
Call InitINIReadWriteScaler(Int(1), tfilename$, "Registration", "X1Real", tvalue!)
If ierror Then Exit Sub
xmin! = tvalue!
Call InitINIReadWriteScaler(Int(1), tfilename$, "Registration", "Y1Real", tvalue!)
If ierror Then Exit Sub
ymin! = tvalue!

Call InitINIReadWriteScaler(Int(1), tfilename$, "Registration", "X2Real", tvalue!)
If ierror Then Exit Sub
xmax! = tvalue!
Call InitINIReadWriteScaler(Int(1), tfilename$, "Registration", "Y2Real", tvalue!)
If ierror Then Exit Sub
ymax! = tvalue!

' Check to see if valid calibration coordinates exist
If ixmin& = ixmax& Or iymin& = iymax& Then Exit Sub
If xmin! = xmax! Or ymin! = ymax! Then Exit Sub

' Load stage coordinates to the picturesnap calibration window (note min/max are inverted for JEOL)
If Not ImageInterfaceStageXPolarity Then
FormPICTURESNAP2.TextXStage1.Text = xmin! * MICRONSPERMM&
FormPICTURESNAP2.TextXStage2.Text = xmax! * MICRONSPERMM&
Else
FormPICTURESNAP2.TextXStage1.Text = xmax! * MICRONSPERMM&
FormPICTURESNAP2.TextXStage2.Text = xmin! * MICRONSPERMM&
End If

If Not ImageInterfaceStageYPolarity Then
FormPICTURESNAP2.TextYStage1.Text = ymin! * MICRONSPERMM&
FormPICTURESNAP2.TextYStage2.Text = ymax! * MICRONSPERMM&
Else
FormPICTURESNAP2.TextYStage1.Text = ymax! * MICRONSPERMM&
FormPICTURESNAP2.TextYStage2.Text = ymin! * MICRONSPERMM&
End If

' Load pixel coordinates (can't assume it is cartesian)
If Not ImageInterfaceStageXPolarity Then
FormPICTURESNAP2.TextXPixel1.Text = ixmin& * Screen.TwipsPerPixelX
FormPICTURESNAP2.TextXPixel2.Text = ixmax& * Screen.TwipsPerPixelX
Else
FormPICTURESNAP2.TextXPixel1.Text = ixmax& * Screen.TwipsPerPixelX
FormPICTURESNAP2.TextXPixel2.Text = ixmin& * Screen.TwipsPerPixelX
End If

If Not ImageInterfaceStageYPolarity Then
FormPICTURESNAP2.TextYPixel1.Text = iymin& * Screen.TwipsPerPixelY
FormPICTURESNAP2.TextYPixel2.Text = iymax& * Screen.TwipsPerPixelY
Else
FormPICTURESNAP2.TextYPixel1.Text = iymax& * Screen.TwipsPerPixelY
FormPICTURESNAP2.TextYPixel2.Text = iymin& * Screen.TwipsPerPixelY
End If

' Read in keV
PictureSnap_keV! = Val(Base64ReaderGetINIString$(tfilename$, "ColumnConditions", "HighVoltage", Format$(DefaultKiloVolts!)))
If ierror Then Exit Sub

' Read imaging magnification
PictureSnap_mag! = Val(Base64ReaderGetINIString$(tfilename$, "ColumnConditions", "Magnification", Format$(DefaultMagnificationImaging!)))
If ierror Then Exit Sub

' Read scan rotation
PictureSnap_scanrota! = Val(Base64ReaderGetINIString$(tfilename$, "ColumnConditions", "ScanRotation", Format$(DefaultScanRotation!)))
If ierror Then Exit Sub

' Create .ACQ file
PictureSnapMode% = 0    ' two calibration coordinates only
PictureSnapFilename$ = PictureSnapFilename$         ' already loaded correctly
Call PictureSnapCalibrate(Int(1))
If ierror Then
PictureSnapFilename$ = vbNullString
Exit Sub
End If

Exit Sub

' Errors
PictureSnapConvertPrbImgToACQError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapConvertPrbImgToACQ"
ierror = True
Exit Sub

End Sub

Sub PictureSnapSaveGridFile()
' Save the picture to a grid file

ierror = False
On Error GoTo PictureSnapSaveGridFileError

Dim i As Integer, j As Integer
Dim n As Long
Dim temp As Single
Dim fractionx As Single, fractiony As Single

Dim nX As Long, nY As Long
Dim bpp As Integer, nSize As Double, tptr As Long
Dim larray() As Long

Dim ixmin As Single, ixmax As Single, iymin As Single, iymax As Single, izmin As Single, izmax As Single
Dim xmin As Single, xmax As Single, ymin As Single, ymax As Single
Dim zmin As Single, zmax As Single

Dim tfilename As String, tfilename2 As String

' If no picture just exit
If PictureSnapFilename$ = vbNullString Then GoTo PictureSnapSaveGridFileNoPicture

' If not calibrated, warn user
If Not PictureSnapCalibrated Then GoTo PictureSnapSaveGridFileNotCalibrated
tfilename$ = MiscGetFileNameNoExtension$(PictureSnapFilename$) & ".GRD"

' Load current image into picture control
FormPICTURESNAP.Picture3.Picture = FormPICTURESNAP.Picture2.Picture     '  Picture3 control is in Pixel ScaleMode

' Convert BMP to byte array
Call IOStatusAuto("Loading BMP array...")
DoEvents
Screen.MousePointer = vbHourglass
Call BMPGetBitmapInfo(FormPICTURESNAP.Picture3, nX&, nY&, bpp%, nSize#, tptr&)      '  Picture3 control is in Pixel ScaleMode
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

' Convert picture to gray if not too large
Call IOStatusAuto("Converting picture to gray...")
DoEvents
Screen.MousePointer = vbHourglass
Call BMPMakeGray(FormPICTURESNAP.Picture3)     '  Picture3 control is in Pixel ScaleMode
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

' Copy bitmap to long array
Screen.MousePointer = vbHourglass
ReDim larray(1 To nX&, 1 To nY&) As Long
Call BMPConvertBitmapToLongArray(FormPICTURESNAP.Picture3, nX&, nY&, larray&())     '  Picture3 control is in Pixel ScaleMode
Screen.MousePointer = vbDefault

' Get stage extents (use Picture2 control in Twips for stage coordinate calculations)
ixmin! = 0
iymin! = 0
ixmax! = FormPICTURESNAP.Picture2.ScaleX(FormPICTURESNAP.Picture2.Picture.Width, vbHimetric, vbTwips)
iymax! = FormPICTURESNAP.Picture2.ScaleY(FormPICTURESNAP.Picture2.Picture.Height, vbHimetric, vbTwips)

' Convert to stage coordinates
Call PictureSnapConvert(Int(1), ixmin!, iymin!, izmin!, xmin!, ymin!, zmin!, fractionx!, fractiony!)
If ierror Then Exit Sub
Call PictureSnapConvert(Int(1), ixmax!, iymax!, izmax!, xmax!, ymax!, zmax!, fractionx!, fractiony!)
If ierror Then Exit Sub

' Find actual min and max (do not use ImageInterfaceStageXPolarity and ImageInterfaceStageXPolarity flags)
If xmax! < xmin! Then
temp! = xmax!
xmax! = xmin!
xmin! = temp!
End If

If ymax! < ymin! Then
temp! = ymax!
ymax! = ymin!
ymin! = temp!
End If

' Find actual zmin and zmax
zmin! = MAXMINIMUM!
zmax! = MAXMAXIMUM!
For j% = 1 To nY&
For i% = 1 To nX&
If larray&(i%, j%) < zmin! Then zmin! = larray&(i%, j%)
If larray&(i%, j%) > zmax! Then zmax! = larray&(i%, j%)
Next i%
Next j%

' Load grid structure
GridImageData(1).id$ = "DSBB"
GridImageData(1).ix% = nX&
GridImageData(1).iy% = nY&
GridImageData(1).xmin# = xmin!
GridImageData(1).xmax# = xmax!
GridImageData(1).ymin# = ymin!
GridImageData(1).ymax# = ymax!
GridImageData(1).zmin# = zmin!
GridImageData(1).zmax# = zmax!

' Dimension array
ReDim GridImageData(1).gData(1 To nX&, 1 To nY&) As Single

' Orient based on ImageInterfaceDisplay flags
Screen.MousePointer = vbHourglass
For j% = 1 To nY&
For i% = 1 To nX&
If ImageInterfaceDisplayXPolarity And ImageInterfaceDisplayYPolarity Then
GridImageData(1).gData!(i%, j%) = larray(nX& - (i% - 1), nY& - (j% - 1))

ElseIf ImageInterfaceDisplayXPolarity And Not ImageInterfaceDisplayYPolarity Then
GridImageData(1).gData!(i%, j%) = larray(nX& - (i% - 1), j%)

ElseIf Not ImageInterfaceDisplayXPolarity And ImageInterfaceDisplayYPolarity Then
GridImageData(1).gData!(i%, j%) = larray(i%, nY& - (j% - 1))

Else
GridImageData(1).gData!(i%, j%) = larray(i%, j%)
End If
Next i%
Next j%
Screen.MousePointer = vbDefault

' Load the file based on actual file version number
Call IOStatusAuto("Writing GRD array...")
If SurferOutputVersionNumber% = 6 Then
Screen.MousePointer = vbHourglass
Call GridFileReadWrite(Int(2), Int(1), GridImageData(), tfilename$)
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

Else
Screen.MousePointer = vbHourglass
Call GridFileReadWrite2(Int(2), Int(1), GridImageData(), tfilename$)
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End If

' Write post file if points are available
If ImagePoints& > 0 Then
Call IOStatusAuto("Writing image data points...")

' Write post file
Close (Temp1FileNumber%)
tfilename2$ = MiscGetFileNameNoExtension$(tfilename$) & ".DAT"
Open tfilename2$ For Output As #Temp1FileNumber%

' Output column labels
Print #Temp1FileNumber%, VbDquote$ & "xstage" & VbDquote$ & vbTab, VbDquote$ & "ystage" & VbDquote$ & vbTab, VbDquote$ & "line" & VbDquote$

' Loop on all coordinates
For n& = 1 To ImagePoints&

' Check for out of bounds
If xmin! < xmax! Then
If ImageXdata!(n&) < xmin! Or ImageXdata!(n&) > xmax! Then GoTo 3000:
Else
If ImageXdata!(n&) > xmin! Or ImageXdata!(n&) < xmax! Then GoTo 3000:
End If

If ymin! < ymax! Then
If ImageYdata!(n&) < ymin! Or ImageYdata!(n&) > ymax! Then GoTo 3000:
Else
If ImageYdata!(n&) > ymin! Or ImageYdata!(n&) < ymax! Then GoTo 3000:
End If

' Output position data
Print #Temp1FileNumber%, MiscAutoFormat$(ImageXdata!(n&)) & vbTab, MiscAutoFormat$(ImageYdata!(n&)) & vbTab, MiscAutoFormatI$(ImageSData%(n&)) & vbTab, MiscAutoFormatI$(ImageNData%(n&))

3000:
Next n&

Close (Temp1FileNumber%)
End If

Call IOStatusAuto(vbNullString)
If ImagePoints& > 0 Then
msg$ = "Grid data saved to " & tfilename$ & ", stage coordinate post data saved to " & tfilename2$
Else
msg$ = "Grid data saved to " & tfilename$
End If
MsgBox msg$, vbOKOnly + vbInformation, "PictureSnapSaveGridFile"
Exit Sub

' Errors
PictureSnapSaveGridFileError:
Close (Temp1FileNumber%)
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapSaveGridFile"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

PictureSnapSaveGridFileNoPicture:
msg$ = "No picture (*.BMP) has been loaded in the PictureSnap window. Please open a sample picture using the File | Open menu."
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapSaveGridFile"
ierror = True
Exit Sub

PictureSnapSaveGridFileNotCalibrated:
msg$ = "The picture calibration cannot be saved because the picture has not been calibrated. Use the Window | Calibrate menu to first calibrate the picture to your stage coordinate system."
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapSaveGridFile"
ierror = True
Exit Sub

End Sub

Sub PictureSnapDisplayPositions()
' Display positions from position database

ierror = False
On Error GoTo PictureSnapDisplayPositionsError

Const radius! = 80

Static calibrationsavedtodisk As Boolean

Dim tfilenumber As Integer
Dim tcolor As Long, n As Long
Dim sx As Single, sy As Single, sz As Single
Dim formx As Single, formy As Single, formz As Single
Dim fractionx As Single, fractiony As Single

Dim ixmin As Single, ixmax As Single, iymin As Single, iymax As Single
Dim xmin As Single, xmax As Single, ymin As Single, ymax As Single, zmin As Single, zmax As Single
Dim astring As String, tfilename As String

' If form not visible just exit
If Not FormPICTURESNAP.Visible Then Exit Sub

' If no picture just exit
If PictureSnapFilename$ = vbNullString Then Exit Sub

' if not calibrated, just exit
If Not PictureSnapCalibrated Then Exit Sub

' Determine image extents
If FormPICTURESNAP.Picture2.Picture.Type <> 1 Then Exit Sub     ' not bitmap
ixmax! = FormPICTURESNAP.Picture2.ScaleX(FormPICTURESNAP.Picture2.Picture.Width, vbHimetric, vbTwips)
iymax! = FormPICTURESNAP.Picture2.ScaleY(FormPICTURESNAP.Picture2.Picture.Height, vbHimetric, vbTwips)

' Convert to stage coordinates (z coordinates are not used)
'Call PictureSnapConvert(Int(1), ixmin!, iymin!, CSng(0#), xmin!, ymin!, zmin!, fractionx!, fractiony!)
'If ierror Then Exit Sub
'Call PictureSnapConvert(Int(1), ixmax!, iymax!, CSng(0#), xmax!, ymax!, zmax!, fractionx!, fractiony!)
'If ierror Then Exit Sub

' Convert to stage coordinates (z coordinates are not used) (this code flips the y min/max- changed 04/26/2016)
Call PictureSnapConvert(Int(1), ixmin!, iymax!, CSng(0#), xmin!, ymin!, zmin!, fractionx!, fractiony!)
If ierror Then Exit Sub
Call PictureSnapConvert(Int(1), ixmax!, iymin!, CSng(0#), xmax!, ymax!, zmax!, fractionx!, fractiony!)
If ierror Then Exit Sub

' Save window calibration to disk file for debugging
If Not calibrationsavedtodisk And ImagePoints& > 0 Then
tfilenumber% = FreeFile()
tfilename$ = ApplicationCommonAppData$ & "PictureSnap.txt"
Open tfilename$ For Output As #tfilenumber%

astring$ = "Image window twips: " & Format$(ixmin!) & ", " & Format$(iymin!) & ", " & Format$(ixmax!) & ", " & Format$(iymax!)
Print #tfilenumber%, astring$

astring$ = "Image window stage: " & Format$(xmin!) & ", " & Format$(ymin!) & ", " & Format$(xmax!) & ", " & Format$(ymax!)
Print #tfilenumber%, astring$
End If

' Note that x/y min/max are not correct for rotated images imported from another source, so add a "buffer" of 45 degrees rotation?
If xmin! < xmax! Then
xmin! = xmin! - (Abs(xmax! - xmin!) * Sqr(2)) / 2#
xmax! = xmax! + (Abs(xmax! - xmin!) * Sqr(2)) / 2#
Else
xmin! = xmin! + (Abs(xmax! - xmin!) * Sqr(2)) / 2#
xmax! = xmax! - (Abs(xmax! - xmin!) * Sqr(2)) / 2#
End If

If ymin! < ymax! Then
ymin! = ymin! - (Abs(ymax! - ymin!) * Sqr(2)) / 2#
ymax! = ymax! + (Abs(ymax! - ymin!) * Sqr(2)) / 2#
Else
ymin! = ymin! + (Abs(ymax! - ymin!) * Sqr(2)) / 2#
ymax! = ymax! - (Abs(ymax! - ymin!) * Sqr(2)) / 2#
End If

' Save 45 degree modified window calibration to disk file for debugging
If Not calibrationsavedtodisk And ImagePoints& > 0 Then
astring$ = "Image window stage: " & Format$(xmin!) & ", " & Format$(ymin!) & ", " & Format$(xmax!) & ", " & Format$(ymax!)
Print #tfilenumber%, astring$
End If

' Loop on all coordinates
For n& = 1 To ImagePoints&
sx! = ImageXdata!(n&)
sy! = ImageYdata!(n&)
sz! = ImageZdata!(n&)

' Check for out of bounds
If xmin! < xmax! Then
If sx! < xmin! Or sx! > xmax! Then GoTo 2000:
Else
If sx! > xmin! Or sx! < xmax! Then GoTo 2000:
End If

If ymin! < ymax! Then
If sy! < ymin! Or sy! > ymax! Then GoTo 2000:
Else
If sy! > ymin! Or sy! < ymax! Then GoTo 2000:
End If

' Calculate screen coordinate for picture control
Call PictureSnapConvert(Int(2), formx!, formy!, formz!, sx!, sy!, sz!, fractionx!, fractiony!)
If ierror Then Exit Sub

' Save first position to disk file for debugging
If Not calibrationsavedtodisk And ImagePoints& > 0 Then
astring$ = "Position stage: " & Format$(sx!) & ", " & Format$(sy!)
Print #tfilenumber%, astring$
astring$ = "Position twips: " & Format$(formx!) & ", " & Format$(formy!)
Print #tfilenumber%, astring$
Close #tfilenumber%
calibrationsavedtodisk = True
End If

' Draw positions directly on picture control
tcolor& = RGB(255, 0, 0)    ' red
FormPICTURESNAP.Picture2.DrawWidth = 2
FormPICTURESNAP.Picture2.Circle (formx!, formy!), (radius!), tcolor&

' Draw line numbers on picture control if indicated
If ImageSData%(n&) > 0 Or ImageNData%(n&) > 0 Then
If ImageSData%(n&) > 0 And ImageNData%(n&) > 0 Then
astring$ = Str$(ImageSData%(n&)) & "-" & Format$(ImageNData%(n&))
ElseIf ImageSData%(n&) = 0 And ImageNData%(n&) > 0 Then
astring$ = Str$(ImageNData%(n&))
Else
astring$ = vbNullString
End If
FormPICTURESNAP.Picture2.ForeColor = tcolor& ' set foreground color
FormPICTURESNAP.Picture2.FontSize = 10       ' set font size
FormPICTURESNAP.Picture2.FontName = LogWindowFontName$
FormPICTURESNAP.Picture2.FontSize = 10       ' set font size    (necessary for Windows)
'halfwidth! = Formpicturesnap.Picture2.TextWidth(astring$) / 2      ' calculate one-half width
'halfheight! = Formpicturesnap.Picture2.TextHeight(astring$) / 2     ' calculate one-half height
'Formpicturesnap.Picture2.CurrentX = Formpicturesnap.Picture2.CurrentX + halfwidth!   ' set X
'Formpicturesnap.Picture2.CurrentY = Formpicturesnap.Picture2.CurrentY + halfheight! ' set Y
FormPICTURESNAP.Picture2.Print astring$   ' print text string to picture
End If

2000:
Next n&

' In case no points were found to display, close debug file
If Not calibrationsavedtodisk And ImagePoints& > 0 Then
Close #tfilenumber%
calibrationsavedtodisk = True
End If

Exit Sub

' Errors
PictureSnapDisplayPositionsError:
Close #tfilenumber%
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapDisplayPositions"
ierror = True
Exit Sub

End Sub

Sub PictureSnapLoadPositions(mode As Integer)
' Load the positions for the specified samples type (called from PictureSnap Display menu for digitized positions)
'   mode = 0 just reload for labels
'   mode = 1 load standards
'   mode = 2 load unknowns
'   mode = 3 load wavescans

ierror = False
On Error GoTo PictureSnapLoadPositionsError

Dim i As Long, imode As Integer, samplerow As Integer, npts As Integer

ImagePoints& = 0

' Check if only loading data for selected sample
If FormPICTURESNAP.Visible And FormPICTURESNAP.menuDisplayDisplayDigitizedPositionsForSelectedPositionSampleOnly.Checked Then
If FormAUTOMATE.ListDigitize.ListIndex < 0 Then GoTo PictureSnapLoadPositionsNotSelected
samplerow% = FormAUTOMATE.ListDigitize.ItemData(FormAUTOMATE.ListDigitize.ListIndex)
Call PositionGetSampleDataOnly(samplerow%, npts%, ImageXdata!(), ImageYdata!(), ImageZdata!(), ImageIdata%())
If ierror Then Exit Sub

' Dimension unused arrays
ImagePoints& = npts%
If ImagePoints& > 0 Then
ReDim ImageNData(1 To ImagePoints&) As Integer
ReDim ImageSData(1 To ImagePoints&) As Integer
ReDim ImageSNdata(1 To ImagePoints&) As String
End If

' Load data for selected position sample type
ElseIf FormPICTURESNAP.menuDisplayStandards.Checked Or FormPICTURESNAP.menuDisplayUnknowns.Checked Or FormPICTURESNAP.menuDisplayWavescans.Checked Then

' Load sample types
imode% = mode%
If imode% = 0 Then imode% = lastmode%
Call PositionGetXYZ(imode%, ImagePoints&, ImageXdata!(), ImageYdata!(), ImageZdata!(), ImageIdata%(), ImageNData%(), ImageSData%(), ImageSNdata$())
If ierror Then Exit Sub
If imode% > 0 Then lastmode% = imode%  ' save for next load label call (mode% = 0)
End If

' Print out for debug
If DebugMode% Then
Call IOWriteLog(vbCrLf & "Number of position database points to plot:" & Str$(ImagePoints&))
For i& = 1 To ImagePoints&
Call IOWriteLog(Str$(i&) & ", " & Str$(ImageXdata!(i&)) & ", " & Str$(ImageYdata!(i&)) & ", " & Str$(ImageZdata!(i&)) & ", " & Str$(ImageIdata%(i&)) & ", " & Str$(ImageNData%(i&)) & ", " & Str$(ImageSData%(i&)))
Next i&
End If

' Remove label information if indicated
If Not FormPICTURESNAP.menuDisplayLongLabels.Checked And Not FormPICTURESNAP.menuDisplayShortLabels.Checked Then
For i& = 1 To ImagePoints&
ImageNData%(i&) = 0     ' line numbers
ImageSData%(i&) = 0     ' sample numbers
ImageSNdata$(i&) = vbNullString     ' sample number
Next i&

Else
If FormPICTURESNAP.menuDisplayShortLabels.Checked Then
For i& = 1 To ImagePoints&
ImageSData%(i&) = 0     ' sample numbers
Next i&
End If
End If

FormPICTURESNAP.Picture2.Refresh
Exit Sub

' Errors
PictureSnapLoadPositionsError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapLoadPositions"
ierror = True
Exit Sub

PictureSnapLoadPositionsNotSelected:
msg$ = "The Display Digitized Positions For Selected Position Sample Only menu was checked but no position sample is currently selected in the Automate! window"
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapLoadPositionsNotSelected"
ierror = True
Exit Sub

End Sub

Sub PictureSnapLoadPositions2(methodlong As Boolean, methodshort As Boolean, n As Long, sdata() As Long, ndata() As Long, xdata() As Single, ydata() As Single, zdata() As Single)
' Load the positions for the passed positions (called from Probe for EPMA Run menu for analyzed positions)
' if method1 = True then load long line labels
' if method2 = True then load short line labels
'
' n = number of stage positions
' sdata() = sample numbers
' ndata() = line numbers
' xdata() = x stage positions
' ydata() = y stage positions
' zdata() = z stage positions

ierror = False
On Error GoTo PictureSnapLoadPositions2Error

Dim i As Long

' Load number of points
ImagePoints& = n&

' Dimension arrays
If ImagePoints& > 0 Then
ReDim ImageXdata(1 To ImagePoints&) As Single
ReDim ImageYdata(1 To ImagePoints&) As Single
ReDim ImageZdata(1 To ImagePoints&) As Single
ReDim ImageNData(1 To ImagePoints&) As Integer
ReDim ImageSData(1 To ImagePoints&) As Integer

' Load arrays
For i& = 1 To n&
ImageXdata!(i&) = xdata!(i&)
ImageYdata!(i&) = ydata!(i&)
ImageZdata!(i&) = zdata!(i&)
If methodlong Then
ImageNData%(i&) = ndata&(i&)    ' line numbers
ImageSData%(i&) = sdata&(i&)    ' sample numbers
ElseIf methodshort Then
ImageNData%(i&) = ndata&(i&)    ' line numbers only
End If
Next i&
End If

FormPICTURESNAP.Picture2.Refresh
Exit Sub

' Errors
PictureSnapLoadPositions2Error:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapLoadPositions2"
ierror = True
Exit Sub

End Sub

Sub PictureSnapResizeFullView()
' Resize the full view window based on PictureSnapImageWidth and PictureSnapImageHeight to maintain original aspect ratio.

ierror = False
On Error GoTo PictureSnapResizeFullViewError

If Not FormPICTURESNAP.menuMiscMaintainAspectRatioOfFullViewWindow.Checked Then Exit Sub

' Only resize if not minimized or maximized
If FormPICTURESNAP3.WindowState = 0 Then

' Only resize if image size has been loaded (from PictureSnapFileOpen)
If PictureSnapImageHeight! <> 0# And PictureSnapImageWidth! <> 0# Then
FormPICTURESNAP3.Height = FormPICTURESNAP3.Width * PictureSnapImageHeight! / PictureSnapImageWidth! + (FormPICTURESNAP3.Height - FormPICTURESNAP3.ScaleHeight)
End If
End If

Exit Sub

' Errors
PictureSnapResizeFullViewError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapResizeFullView"
ierror = True
Exit Sub

End Sub

Sub PictureSnapImportPrbImg()
' Import a selected PrbImg and save to GRD format

ierror = False
On Error GoTo PictureSnapImportPrbImgError

Dim beamcurrent1 As Single, beamcurrent2 As Single, keV As Single, counttime As Single
Dim timeofacq1 As Double, timeofacq2 As Double
Dim mag As Double, scanrota As Double
Dim scanflag As Integer, stageflag As Integer
Dim astring As String

Dim sarray() As Single
Dim response As Integer

Dim ix As Integer, iy As Integer
Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double, zmin As Double, zmax As Double

Dim gfilename As String, bfilename$

Static tfilename As String

' Unload the calibration and full size forms in case loaded
Unload FormPICTURESNAP2
Unload FormPICTURESNAP3

' Get file from user
If tfilename$ = vbNullString Then tfilename$ = UserImagesDirectory$ & "\*.PrbImg"
Call IOGetFileName(Int(2), "PrbImg", tfilename$, FormPICTURESNAP)
If ierror Then Exit Sub

UserImagesDirectory$ = MiscGetPathOnly2$(tfilename$)
Screen.MousePointer = vbHourglass

' Extract data from PrbImg
Call IOStatusAuto("Reading PrbImg file " & tfilename$ & "...")
DoEvents
Call Base64ReaderInput(tfilename$, keV!, counttime!, beamcurrent1!, beamcurrent2!, timeofacq1#, timeofacq2#, ix%, iy%, sarray!(), xmin#, xmax#, ymin#, ymax#, zmin#, zmax#, mag#, scanrota#, scanflag%, stageflag%, astring$)
If ierror Then Exit Sub

' Create GRD file from extracted data and save to folder
gfilename$ = MiscGetFileNameNoExtension$(tfilename$) & ".grd"
Call IOStatusAuto("Writing GRD file " & MiscGetFileNameOnly$(gfilename$) & "...")
DoEvents
Call CalcImageCreateGRDFromArray(gfilename$, ix%, iy%, sarray!(), xmin#, xmax#, ymin#, ymax#, zmin#, zmax#)
If ierror Then Exit Sub

Call IOStatusAuto("Loading image grid file " & gfilename$ & "...")
DoEvents
bfilename$ = gfilename$
Call PictureSnapFileOpenGrid(bfilename$)
If ierror Then Exit Sub

PictureSnapFilename$ = bfilename$
If ierror Then Exit Sub
If bfilename$ <> vbNullString Then
Call PictureSnapFileOpen(Int(0), bfilename$, FormPICTURESNAP)    ' re-open original PictureSnap file if not blank
If ierror Then Exit Sub
End If

' Ask user if they want to save it to the current project folder
msg$ = "Do you want to save the imported PrbImg (.GRD) file to the probe database folder (" & UserDataDirectory$ & ")?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton1, "PictureSnapImportPrbImg")
If response% = vbYes Then
FileCopy gfilename$, UserDataDirectory$ & "\" & MiscGetFileNameOnly$(gfilename$)
End If

Screen.MousePointer = vbDefault
Call IOStatusAuto(vbNullString)
Exit Sub

' Errors
PictureSnapImportPrbImgError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapImportPrbImg"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub PictureSnapLoadACQ()
' Load a calibration from another ACQ file

ierror = False
On Error GoTo PictureSnapLoadACQError

Dim response As Integer
Dim tfilename As String, tfilename2 As String

Dim astring As String

Dim gX_Polarity As Integer, gY_Polarity As Integer
Dim gStage_Units As String

' Ask user to confirm
If PictureSnapCalibrated Then
msg$ = "This currently loaded image is already calibrated.  Are you sure you want to load another ACQ file for the calibration of the currently loaded image?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton1, "PictureSnapLoadACQ")
If response% = vbCancel Then Exit Sub

Else
msg$ = "Are you sure you want to load an ACQ file for the calibration of the currently loaded (uncalibrated) image?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton1, "PictureSnapLoadACQ")
If response% = vbCancel Then Exit Sub
End If

' Browse for existing ACQ
Call IOGetFileName(Int(2), "ACQ", tfilename$, FormPICTURESNAP2)
If ierror Then Exit Sub

' Check for existing GRD or ACQ info
Call GridCheckGRDInfo(tfilename$, gX_Polarity%, gY_Polarity%, gStage_Units$)
If ierror Then Exit Sub

' Check if stage orientation is different than current configuration
If Default_X_Polarity% = 0 And Default_Y_Polarity% = 0 Then
astring$ = "Cameca"
Else
astring$ = "JEOL"
End If

If Default_X_Polarity% <> gX_Polarity% Or Default_Y_Polarity% <> gY_Polarity% Then
msg$ = "The image file " & tfilename$ & " has a different stage orientation than the current stage configuration (" & astring$ & ")." & vbCrLf & vbCrLf
msg$ = msg$ & "Please load an image file with the correct stage orientation or change the current stage configuration from the File menu and try again."
MsgBox msg$, vbOKOnly + vbInformation, "PictureSnapLoadACQ"
Screen.MousePointer = vbDefault
ierror = True
Exit Sub
End If

' Check if stage units is different than current configuration
If Default_Stage_Units$ = "mm" Then
astring$ = "mm"
Else
astring$ = "um"
End If

If Default_Stage_Units$ <> gStage_Units$ Then
msg$ = "The image file " & tfilename$ & " has a different stage units than the current stage configuration (" & astring$ & ")." & vbCrLf & vbCrLf
msg$ = msg$ & "Please load an image file with the correct stage units or change the stage configuration from the File menu and try again."
MsgBox msg$, vbOKOnly + vbInformation, "PictureSnapLoadACQ"
Screen.MousePointer = vbDefault
Exit Sub
End If

' Now copy the file to the current image name
tfilename2$ = MiscGetFileNameNoExtension$(PictureSnapFilename$) & ".acq"
If tfilename$ = tfilename2$ Then GoTo PictureSnapLoadACQSameACQ
FileCopy tfilename$, tfilename2$

' Now load the calibration (new or again)
Call PictureSnapLoadCalibration
If ierror Then Exit Sub

' Check for bad stage calibration
Call PictureSnapCalibrateCheck
If ierror Then Exit Sub

' Confirm to user
msg$ = "Calibration file " & tfilename$ & " was loaded to the current image. You may want to confirm the accuracy of the calibration by using the Move To buttons."
MsgBox msg$, vbOKOnly + vbInformation, "PictureSnapLoadACQ"

Exit Sub

' Errors
PictureSnapLoadACQError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapFileLoadACQ"
ierror = True
Exit Sub

PictureSnapLoadACQSameACQ:
msg$ = "The specified ACQ file " & tfilename$ & " is the same as the current ACQ file. Please try again with a different ACQ file."
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapFileLoadACQ"
ierror = True
Exit Sub

End Sub

