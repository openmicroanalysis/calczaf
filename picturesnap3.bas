Attribute VB_Name = "CodePictureSnapCalibration"
' (c) Copyright 1995-2020 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Global PictureSnapFilename As String
Global PictureSnapMode As Integer                           ' 0 = normal two control point conversion, 1 = three control point transformation
Global PictureSnapCalibrationNumberofZPoints As Integer     ' 0 if using two control points, 3 if using three control points
Global PictureSnapCalibrationSaved As Boolean

Global PictureSnapCalibrated As Boolean
Global PictureSnapCalibratedPreviously As Boolean

Global PictureSnapRotation As Single                        ' image rotation relative to the stage calibration

Global PictureSnapWindowIsModeless As Boolean

' Form variables for keV and mag (and scan rotation)
Global PictureSnap_keV As Single, PictureSnap_mag As Single, PictureSnap_scanrota As Single

' Form variables for signal type, min and max
Global PictureSnap_signal_type As String, PictureSnap_signal_min As Single, PictureSnap_signal_max As Single

' New ACQ file version number
Const ACQ_FILE_VERSION! = 1
Dim ACQFileVersionNumber As Single

' New .ACQ calibration parameters,  where 1.0 = 100% (the default Windows setting), 1.25 = 125% and 2.0 = 200% DPI scaling
Dim ACQScreenDPI_Current As Single
Dim ACQScreenDPI_Stored As Single

' Fiducial arrays for 3 point transformation
Dim fiducialold(1 To MAXAXES%, 1 To MAXDIM%) As Single
Dim fiducialnew(1 To MAXAXES%, 1 To MAXDIM%) As Single
Dim fiducialtranslation(1 To MAXDIM%) As Double
Dim fiducialmatrix(1 To MAXDIM%, 1 To MAXDIM%) As Double

Dim cpoint1x As Single, cpoint1y As Single, cpoint1z As Single  ' reference screen coordinates
Dim cpoint2x As Single, cpoint2y As Single, cpoint2z As Single  ' (z position is dummy)

Dim apoint1x As Single, apoint1y As Single, apoint1z As Single  ' reference stage coordinates
Dim apoint2x As Single, apoint2y As Single, apoint2z As Single  ' (z position is actual)

' Additional 3rd point for fiducial transformation (including rotation)
Dim cpoint3x As Single, cpoint3y As Single, cpoint3z As Single  ' 3rd point screen coordinates
Dim apoint3x As Single, apoint3y As Single, apoint3z As Single  ' 3rd point stage coordinates

Sub PictureSnapReadCalibration(tfilename$)
' Read the screen and stage calibration from the passed ACQ file (note that CalcImage also writes a "pseudo" .ACQ file)

ierror = False
On Error GoTo PictureSnapReadCalibrationError

Dim pmode As Single
Dim points As Single
Dim zpoints As Single
Dim DPI_Ratio As Single
Dim tcomment As String

' Dimension coordinates (assume using three points and XYZ)
Dim cpoint1(1 To 3) As Single, cpoint2(1 To 3) As Single, cpoint3(1 To 3) As Single
Dim apoint1(1 To 3) As Single, apoint2(1 To 3) As Single, apoint3(1 To 3) As Single

' Read ACQ file version number
ACQFileVersionNumber! = 0                                                                               ' assume version zero of ACQ file (no entry)
Call InitINIReadWriteScaler(Int(1), tfilename$, "stage", "ACQFileVersion", ACQFileVersionNumber!)       ' 1 = read, 2 = write
If ierror Then Exit Sub

' Read signal type and signal min/max
PictureSnap_signal_type$ = vbNullString         ' default to empty string
Call InitINIReadWriteString(Int(0), tfilename$, "Signal", "SignalType", PictureSnap_signal_type$, tcomment$)         ' 0 = read, 1 = write
If ierror Then Exit Sub

PictureSnap_signal_min! = 0#                    ' default to 0
Call InitINIReadWriteScaler(Int(1), tfilename$, "Signal", "SignalMin", PictureSnap_signal_min!)                      ' 1 = read, 2 = write
If ierror Then Exit Sub
PictureSnap_signal_max! = 255#                  ' default to 255
Call InitINIReadWriteScaler(Int(1), tfilename$, "Signal", "SignalMax", PictureSnap_signal_max!)                      ' 1 = read, 2 = write
If ierror Then Exit Sub

' Load current screen DPI
ACQScreenDPI_Current! = MiscGetWindowsDPI#()
If ierror Then Exit Sub

' Read stored screen resolution parameter
ACQScreenDPI_Stored! = ACQScreenDPI_Current!        ' set default to current in case the value is not stored in the ACQ file being read
Call InitINIReadWriteScaler(Int(1), tfilename$, "stage", "ACQScreenDPI", ACQScreenDPI_Stored!)                      ' 1 = read, 2 = write
If ierror Then Exit Sub

' Read other parameters and calibration points from INI style ACQ file
Call InitINIReadWriteScaler(Int(1), tfilename$, "stage", "PictureSnap mode", pmode!)                         ' pmode!, 0 = two control points, 1 = three control points
If ierror Then Exit Sub

' Load numbers of points
Call InitINIReadWriteScaler(Int(1), tfilename$, "stage", "Number of calibration points", points!)
If ierror Then Exit Sub
Call InitINIReadWriteScaler(Int(1), tfilename$, "stage", "Number of Z calibration points", zpoints!)    ' zero or three only
If ierror Then Exit Sub

' Read screen calibration points
Call InitINIReadWriteArray(Int(1), tfilename$, "stage", "Screen reference point1 (twips)", Int(2), cpoint1!())
If ierror Then Exit Sub
Call InitINIReadWriteArray(Int(1), tfilename$, "stage", "Screen reference point2 (twips)", Int(2), cpoint2!())
If ierror Then Exit Sub

Call InitINIReadWriteScaler(Int(1), tfilename$, "stage", "Screen Z reference point1 (dummy)", cpoint1!(MAXAXES%))
If ierror Then Exit Sub
Call InitINIReadWriteScaler(Int(1), tfilename$, "stage", "Screen Z reference point2 (dummy)", cpoint2!(MAXAXES%))
If ierror Then Exit Sub

' Read stage calibration points
Call InitINIReadWriteArray(Int(1), tfilename$, "stage", "Stage reference point1", Int(2), apoint1!())
If ierror Then Exit Sub
Call InitINIReadWriteArray(Int(1), tfilename$, "stage", "Stage reference point2", Int(2), apoint2!())
If ierror Then Exit Sub

apoint1!(MAXAXES%) = RealTimeMotorPositions!(ZMotor%)        ' load current position as default for z (in case using two point calibration)
Call InitINIReadWriteScaler(Int(1), tfilename$, "stage", "Stage Z reference point1", apoint1!(MAXAXES%))
If ierror Then Exit Sub
apoint2!(MAXAXES%) = RealTimeMotorPositions!(ZMotor%)        ' load current position as default for z (in case using two point calibration)
Call InitINIReadWriteScaler(Int(1), tfilename$, "stage", "Stage Z reference point2", apoint2!(MAXAXES%))
If ierror Then Exit Sub

' Read 3rd point if indicated
Call InitINIReadWriteArray(Int(1), tfilename$, "stage", "Screen reference point3 (twips)", Int(2), cpoint3!())
If ierror Then Exit Sub
Call InitINIReadWriteArray(Int(1), tfilename$, "stage", "Stage reference point3", Int(2), apoint3!())
If ierror Then Exit Sub

Call InitINIReadWriteScaler(Int(1), tfilename$, "stage", "Screen Z reference point3 (dummy)", cpoint3!(3))
If ierror Then Exit Sub
apoint3!(MAXAXES%) = RealTimeMotorPositions!(ZMotor%)        ' load current position as default for z (in case using two point calibration)
Call InitINIReadWriteScaler(Int(1), tfilename$, "stage", "Stage Z reference point3", apoint3!(MAXAXES%))
If ierror Then Exit Sub

' Load globals
PictureSnapMode% = pmode!
PictureSnapCalibrationNumberofZPoints% = zpoints!

' Load screen coordinates to module level variables
cpoint1x! = cpoint1!(1)  ' x reference screen coordinates
cpoint1y! = cpoint1!(2)  ' y reference screen coordinates
cpoint2x! = cpoint2!(1)  ' x reference screen coordinates
cpoint2y! = cpoint2!(2)  ' y reference screen coordinates

cpoint1z! = cpoint1!(3)  ' Z reference screen coordinates
cpoint2z! = cpoint2!(3)  ' Z reference screen coordinates

' Load stage coordinates to module level variables
apoint1x! = apoint1!(1)  ' x reference stage coordinates
apoint2x! = apoint2!(1)  ' x reference stage coordinates

apoint1y! = apoint1!(2)  ' y reference stage coordinates
apoint2y! = apoint2!(2)  ' y reference stage coordinates

apoint1z! = apoint1!(3)  ' Z reference stage coordinates
apoint2z! = apoint2!(3)  ' Z reference stage coordinates

' Load third point screen and stage coordinates to module level variables
cpoint3x! = cpoint3!(1)  ' x reference screen coordinates
cpoint3y! = cpoint3!(2)  ' y reference screen coordinates
cpoint3z! = cpoint3!(3)  ' Z reference screen coordinates

apoint3x! = apoint3!(1)  ' x reference stage coordinates
apoint3y! = apoint3!(2)  ' y reference stage coordinates
apoint3z! = apoint3!(3)  ' Z reference stage coordinates

' Load keV and mag from .ACQ file
PictureSnap_keV! = DefaultKiloVolts!    ' load default in case parameter is missing
Call InitINIReadWriteScaler(Int(1), tfilename$, "ColumnConditions", "kilovolts", PictureSnap_keV!)
If ierror Then Exit Sub

PictureSnap_mag! = DefaultMagnificationImaging!    ' load default in case parameter is missing
Call InitINIReadWriteScaler(Int(1), tfilename$, "ColumnConditions", "magnification", PictureSnap_mag!)
If ierror Then Exit Sub

PictureSnap_scanrota! = DefaultScanRotation!    ' load default in case parameter is missing
Call InitINIReadWriteScaler(Int(1), tfilename$, "ColumnConditions", "scanrotation", PictureSnap_scanrota!)
If ierror Then Exit Sub

' Modify based on screen DPI ratio
If ACQScreenDPI_Stored! <> ACQScreenDPI_Current! Then
DPI_Ratio! = ACQScreenDPI_Stored! / ACQScreenDPI_Current!
cpoint1x! = cpoint1x! * DPI_Ratio!   ' x reference screen coordinates
cpoint1y! = cpoint1y! * DPI_Ratio!   ' y reference screen coordinates
cpoint2x! = cpoint2x! * DPI_Ratio!   ' x reference screen coordinates
cpoint2y! = cpoint2y! * DPI_Ratio!   ' y reference screen coordinates
cpoint1z! = cpoint1z! * DPI_Ratio!   ' Z reference screen coordinates
cpoint2z! = cpoint2z! * DPI_Ratio!   ' Z reference screen coordinates

' Load third point
cpoint3x! = cpoint3x! * DPI_Ratio!   ' x reference screen coordinates
cpoint3y! = cpoint3y! * DPI_Ratio!   ' y reference screen coordinates
cpoint3z! = cpoint3z! * DPI_Ratio!   ' Z reference screen coordinates
End If

Exit Sub

' Errors
PictureSnapReadCalibrationError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapReadCalibration"
ierror = True
Exit Sub

End Sub

Sub PictureSnapCalibrateLoad2(tForm As Form)
' Loads the form calibration points

ierror = False
On Error GoTo PictureSnapCalibrateLoad2Error

tForm.TextXStage1.Text = apoint1x!
tForm.TextYStage1.Text = apoint1y!

tForm.TextXStage2.Text = apoint2x!
tForm.TextYStage2.Text = apoint2y!

tForm.TextXStage3.Text = apoint3x!
tForm.TextYStage3.Text = apoint3y!

' Load existing z calibrations
tForm.TextZStage1.Text = apoint1z!
tForm.TextZStage2.Text = apoint2z!
tForm.TextZStage3.Text = apoint3z!

tForm.TextXPixel1.Text = cpoint1x!
tForm.TextYPixel1.Text = cpoint1y!
tForm.TextXPixel2.Text = cpoint2x!
tForm.TextYPixel2.Text = cpoint2y!

tForm.TextXPixel3.Text = cpoint3x!
tForm.TextYPixel3.Text = cpoint3y!

Exit Sub

' Errors
PictureSnapCalibrateLoad2Error:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapCalibrateLoad2"
ierror = True
Exit Sub

End Sub

Sub PictureSnapCalibrateSave(tForm As Form)
' Save the form calibration points

ierror = False
On Error GoTo PictureSnapCalibrateSaveError

' Load stage coordinates
apoint1x! = Val(tForm.TextXStage1.Text)
apoint1y! = Val(tForm.TextYStage1.Text)

apoint2x! = Val(tForm.TextXStage2.Text)
apoint2y! = Val(tForm.TextYStage2.Text)

apoint3x! = Val(tForm.TextXStage3.Text)
apoint3y! = Val(tForm.TextYStage3.Text)

' Load z stage coordinates
apoint1z! = Val(tForm.TextZStage1.Text)
apoint2z! = Val(tForm.TextZStage2.Text)
apoint3z! = Val(tForm.TextZStage3.Text)

' Load pixel coordinates
cpoint1x! = Val(tForm.TextXPixel1.Text)
cpoint1y! = Val(tForm.TextYPixel1.Text)

cpoint2x! = Val(tForm.TextXPixel2.Text)
cpoint2y! = Val(tForm.TextYPixel2.Text)

cpoint3x! = Val(tForm.TextXPixel3.Text)
cpoint3y! = Val(tForm.TextYPixel3.Text)
cpoint1z! = 0#
cpoint2z! = 0#
cpoint3z! = 0#

' Check that stage coordinates are not the same
If apoint1x! = apoint2x! Then GoTo PictureSnapCalibrateSaveStageCoordinatesSame
If apoint1y! = apoint2y! Then GoTo PictureSnapCalibrateSaveStageCoordinatesSame
If PictureSnapMode% = 1 Then
If apoint1x! = apoint3x! Or apoint2x! = apoint3x! Then GoTo PictureSnapCalibrateSaveStageCoordinatesSame
If apoint1y! = apoint3y! Or apoint2y! = apoint3y! Then GoTo PictureSnapCalibrateSaveStageCoordinatesSame
End If

' Check that pixel coordinates are not the same
If cpoint1x! = cpoint2x! Then GoTo PictureSnapCalibrateSavePixelCoordinatesSame
If cpoint1y! = cpoint2y! Then GoTo PictureSnapCalibrateSavePixelCoordinatesSame
If PictureSnapMode% = 1 Then
If cpoint1x! = cpoint3x! Or cpoint2x! = cpoint3x! Then GoTo PictureSnapCalibrateSavePixelCoordinatesSame
If cpoint1y! = cpoint3y! Or cpoint2y! = cpoint3y! Then GoTo PictureSnapCalibrateSavePixelCoordinatesSame
End If

Exit Sub

' Errors
PictureSnapCalibrateSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapCalibrateSave"
ierror = True
Exit Sub

PictureSnapCalibrateSaveStageCoordinatesSame:
msg$ = "The X or Y coordinate is the same for two (or three) points of the stage positions"
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapCalibrateSave"
ierror = True
Exit Sub

PictureSnapCalibrateSavePixelCoordinatesSame:
msg$ = "The X or Y coordinate is the same for two (or three) points of the pixel positions"
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapCalibrateSave"
ierror = True
Exit Sub

End Sub

Sub PictureSnapConvert(mode As Integer, formx As Single, formy As Single, formz As Single, stagex As Single, stagey As Single, stagez As Single, fractionx As Single, fractiony As Single)
' Convert from form coordinates to stage coordinates
' mode = 1 convert form to stage
' mode = 2 convert stage to form
' fractionx! and fractiony! are the fractional distance from the image upper left corner for the calculated stage or form position

ierror = False
On Error GoTo PictureSnapConvertError

Dim csx As Single, csy As Single, cox As Single, coy As Single
Dim smallamount As Single

' Check for open PictureSnap file
If Trim$(PictureSnapFilename$) = vbNullString Then Exit Sub

' Convert using two calibration points (no Z stage interpolation)
If PictureSnapMode% = 0 Then

' Check for bad data
'If cpoint1x! - cpoint2x! = 0# Then GoTo PictureSnapConvertBadConvert
'If cpoint1y! - cpoint2y! = 0# Then GoTo PictureSnapConvertBadConvert
If cpoint1x! - cpoint2x! = 0# Then Exit Sub     ' in case calibration hasn't finished loading yet for mouse cursor event
If cpoint1y! - cpoint2y! = 0# Then Exit Sub     ' in case calibration hasn't finished loading yet for mouse cursor event

' Calculate offset and conversion factors
csx! = (apoint1x! - apoint2x!) / (cpoint1x! - cpoint2x!)
csy! = (apoint1y! - apoint2y!) / (cpoint1y! - cpoint2y!)

cox! = apoint1x! - csx! * cpoint1x!
coy! = apoint1y! - csy! * cpoint1y!

If csx! = 0# Then GoTo PictureSnapConvertBadConvert
If csy! = 0# Then GoTo PictureSnapConvertBadConvert

' Convert form to stage
If mode% = 1 Then
stagex! = csx! * formx! + cox!
stagey! = csy! * formy! + coy!

' Convert stage to form
Else
formx! = (stagex! - cox!) / csx!
formy! = (stagey! - coy!) / csy!
End If
End If

' Transform using three calibration points (uses Z stage calibration)
If PictureSnapMode% = 1 Then
Call PictureSnapConvertFiducialsCalculate(mode%, formx!, formy!, formz!, stagex!, stagey!, stagez!)
If ierror Then Exit Sub

' Check to see that calculated z position is in range (if converting from form to stage)
If mode% = 1 Then
smallamount! = Abs(MotHiLimits!(ZMotor%) - MotLoLimits!(ZMotor%)) * SMALLAMOUNTFRACTION!     ' to place it inside the stage limits
If stagez! > MotHiLimits!(ZMotor%) Then stagez! = MotHiLimits!(ZMotor) - smallamount!
If stagez! < MotLoLimits!(ZMotor%) Then stagez! = MotLoLimits!(ZMotor) + smallamount!
End If

' Use current z if using two point calibration
Else
If mode% = 1 Then stagez! = RealTimeMotorPositions!(ZMotor%)
End If

' Calculate fractional distance using form units
If FormPICTURESNAP.Picture2.ScaleWidth <> 0 And FormPICTURESNAP.Picture2.ScaleHeight <> 0 Then
If mode% = 2 Then
fractionx! = formx! / FormPICTURESNAP.Picture2.ScaleWidth
fractiony! = formy! / FormPICTURESNAP.Picture2.ScaleHeight
End If
End If

Exit Sub

' Errors
PictureSnapConvertError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapConvert"
ierror = True
Exit Sub

PictureSnapConvertBadConvert:
msg$ = "One or both of the calibration points are not valid. Try re-loading the image or the calibration again with different points"
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapConvert"
PictureSnapCalibrated = False   ' disable calibration (02-28-2012 for Carpenter)
ierror = True
Exit Sub

End Sub

Sub PictureSnapSaveCalibration(mode As Integer, pFileName As String, pcalibrationsaved As Boolean)
' Save the picture calibration to a text file of INI format (*.ACQ)
'  mode = 0 confirm save
'  mode = 1 do not confirm save

ierror = False
On Error GoTo PictureSnapSaveCalibrationError

Dim tfilename As String, tcomment As String

' Dimension coordinates (assume using three points and XYZ)
Dim cpoint1(1 To 3) As Single, cpoint2(1 To 3) As Single, cpoint3(1 To 3) As Single
Dim apoint1(1 To 3) As Single, apoint2(1 To 3) As Single, apoint3(1 To 3) As Single

' Check if picture file is loaded
If Trim$(pFileName$) = vbNullString Then GoTo PictureSnapSaveCalibrationNoPicture

' Check if picture is calibrated
If Not PictureSnapCalibrated Then GoTo PictureSnapSaveCalibrationNotCalibrated

' Screen coordinates
cpoint1!(1) = cpoint1x!  ' reference screen coordinates
cpoint1!(2) = cpoint1y!  ' reference screen coordinates

cpoint2!(1) = cpoint2x!  ' reference screen coordinates
cpoint2!(2) = cpoint2y!  ' reference screen coordinates

cpoint3!(1) = cpoint3x!  ' reference screen coordinates
cpoint3!(2) = cpoint3y!  ' reference screen coordinates
cpoint1!(3) = cpoint1z!  ' dummy z reference
cpoint2!(3) = cpoint2z!  ' dummy z reference
cpoint3!(3) = cpoint3z!  ' dummy z reference

' Stage coodinates
apoint1!(1) = apoint1x!  ' reference stage coordinates
apoint1!(2) = apoint1y!  ' reference stage coordinates

apoint2!(1) = apoint2x!  ' reference stage coordinates
apoint2!(2) = apoint2y!  ' reference stage coordinates

apoint3!(1) = apoint3x!  ' reference stage coordinates
apoint3!(2) = apoint3y!  ' reference stage coordinates
apoint1!(3) = apoint1z!  ' actual z reference
apoint2!(3) = apoint2z!  ' actual z reference
apoint3!(3) = apoint3z!  ' actual z reference

' Write calibration points to INI style ACQ file
tfilename$ = MiscGetFileNameNoExtension$(pFileName$) & ".ACQ"

' Save parameters
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "PictureSnap mode", CSng(PictureSnapMode%))
If ierror Then Exit Sub

' Save PictureSnap mode
If PictureSnapMode% = 0 Then
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "Number of calibration points", CSng(2))       ' using two control points
If ierror Then Exit Sub
Else
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "Number of calibration points", CSng(3))       ' using three control points
If ierror Then Exit Sub
End If

' Save number of Z positions
If PictureSnapMode% = 1 And NumberOfStageMotors% > 2 Then
PictureSnapCalibrationNumberofZPoints% = 3
Else
PictureSnapCalibrationNumberofZPoints% = 0
End If

Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "Number of Z calibration points", CSng(PictureSnapCalibrationNumberofZPoints%))
If ierror Then Exit Sub

' Write screen calibration points
Call InitINIReadWriteArray(Int(2), tfilename$, "stage", "Screen reference point1 (twips)", Int(2), cpoint1!())
If ierror Then Exit Sub
Call InitINIReadWriteArray(Int(2), tfilename$, "stage", "Screen reference point2 (twips)", Int(2), cpoint2!())
If ierror Then Exit Sub

Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "Screen Z reference point1 (dummy)", cpoint1!(3))
If ierror Then Exit Sub
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "Screen Z reference point2 (dummy)", cpoint2!(3))
If ierror Then Exit Sub

' Write stage calibration points
Call InitINIReadWriteArray(Int(2), tfilename$, "stage", "Stage reference point1", Int(2), apoint1!())
If ierror Then Exit Sub
Call InitINIReadWriteArray(Int(2), tfilename$, "stage", "Stage reference point2", Int(2), apoint2!())
If ierror Then Exit Sub

Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "Stage Z reference point1", apoint1!(3))
If ierror Then Exit Sub
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "Stage Z reference point2", apoint2!(3))
If ierror Then Exit Sub

' Save 3rd point
Call InitINIReadWriteArray(Int(2), tfilename$, "stage", "Screen reference point3 (twips)", Int(2), cpoint3!())
If ierror Then Exit Sub
Call InitINIReadWriteArray(Int(2), tfilename$, "stage", "Stage reference point3", Int(2), apoint3!())
If ierror Then Exit Sub

Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "Screen Z reference point3 (dummy)", cpoint3!(3))
If ierror Then Exit Sub
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "Stage Z reference point3", apoint3!(3))
If ierror Then Exit Sub

' Now save coordinate system of stage orientation and units
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "X_Polarity", CSng(Default_X_Polarity%))              ' 1 = read, 2 = write
If ierror Then Exit Sub
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "Y_Polarity", CSng(Default_Y_Polarity%))              ' 1 = read, 2 = write
If ierror Then Exit Sub
Call InitINIReadWriteString(Int(1), tfilename$, "stage", "Stage_Units", Default_Stage_Units$, vbNullString)    ' 0 = read, 1 = write
If ierror Then Exit Sub

' Save keV, mag and scan rotation
Call InitINIReadWriteScaler(Int(2), tfilename$, "ColumnConditions", "kilovolts", PictureSnap_keV!)
If ierror Then Exit Sub

Call InitINIReadWriteScaler(Int(2), tfilename$, "ColumnConditions", "Magnification", PictureSnap_mag!)
If ierror Then Exit Sub

Call InitINIReadWriteScaler(Int(2), tfilename$, "ColumnConditions", "ScanRotation", PictureSnap_scanrota!)
If ierror Then Exit Sub

' Save current screen DPI for saving new images
ACQScreenDPI_Current! = MiscGetWindowsDPI#()
If ierror Then Exit Sub

' Write screen resolution DPI
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "ACQScreenDPI", ACQScreenDPI_Current!)                      ' 1 = read, 2 = write
If ierror Then Exit Sub

' Write signal type and signal min/max
tcomment$ = "signal type"
Call InitINIReadWriteString(Int(1), tfilename$, "Signal", "SignalType", PictureSnap_signal_type$, tcomment$)         ' 0 = read, 1 = write
If ierror Then Exit Sub

Call InitINIReadWriteScaler(Int(2), tfilename$, "Signal", "SignalMin", PictureSnap_signal_min!)                      ' 1 = read, 2 = write
If ierror Then Exit Sub
Call InitINIReadWriteScaler(Int(2), tfilename$, "Signal", "SignalMax", PictureSnap_signal_max!)                      ' 1 = read, 2 = write
If ierror Then Exit Sub

' Write ACQ file version number
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "ACQFileVersion", ACQ_FILE_VERSION!)                          ' 1 = read, 2 = write
If ierror Then Exit Sub

pcalibrationsaved = True

If mode% = 0 Then
msg$ = "Picture calibration saved to " & tfilename$
MsgBox msg$, vbOKOnly + vbInformation, "PictureSnapSaveCalibration"
End If

Exit Sub

' Errors
PictureSnapSaveCalibrationError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapSaveCalibration"
ierror = True
Exit Sub

PictureSnapSaveCalibrationNoPicture:
msg$ = "No picture (*.BMP) has been loaded in the PictureSnap window. Please open a sample picture using the File | Open menu."
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapSaveCalibration"
ierror = True
Exit Sub

PictureSnapSaveCalibrationNotCalibrated:
msg$ = "The picture calibration cannot be saved because the picture has not been calibrated. Use the Window | Calibrate menu to first calibrate the picture to your stage coordinate system."
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapSaveCalibration"
ierror = True
Exit Sub

End Sub

Sub PictureSnapDisplayCalibrationPoints(tForm As Form, tForm3 As Form)
' Display the calibration points on the PictureSnap form and the full view window form PictureSnap3.

ierror = False
On Error GoTo PictureSnapDisplayCalibrationPointsError

Dim radius As Single, tWidth As Single
Dim fractionx As Single, fractiony As Single
Dim formx As Single, formy As Single, formz As Single
Dim tcolor As Long

' Check if image loaded
If PictureSnapFilename$ = vbNullString Then Exit Sub

' Check if calibrated
If Not PictureSnapCalibrated Then Exit Sub

' Calculate a radius
tWidth! = tForm.Width
If tWidth! = 0# Then Exit Sub
radius! = tWidth! / 100#

tcolor& = RGB(0, 255, 0)

' Display two calibration points (this works since FormPICTURESNAP.Picture2 is 1:1 twips, that is unstretched, but doesn't deal with orientation)
tForm.Picture2.DrawWidth = 2
tForm.Picture2.Circle (cpoint1x!, cpoint1y!), radius!, tcolor&

tForm.Picture2.ForeColor = tcolor&                          ' set foreground color
tForm.Picture2.FontSize = 13                                ' set font size
tForm.Picture2.FontName = LogWindowFontName$
tForm.Picture2.FontSize = 13                                ' set font size    (necessary for Windows)
tForm.Picture2.FontBold = False

tForm.Picture2.CurrentX = cpoint1x!
tForm.Picture2.CurrentY = cpoint1y!
tForm.Picture2.CurrentX = tForm.Picture2.CurrentX + 150     ' set X
tForm.Picture2.CurrentY = tForm.Picture2.CurrentY + 150     ' set Y
tForm.Picture2.Print "1"                                    ' print text string to form

tForm.Picture2.DrawWidth = 2
tForm.Picture2.Circle (cpoint2x!, cpoint2y!), radius!, tcolor&

tForm.Picture2.CurrentX = cpoint2x!
tForm.Picture2.CurrentY = cpoint2y!
tForm.Picture2.CurrentX = tForm.Picture2.CurrentX + 150     ' set X
tForm.Picture2.CurrentY = tForm.Picture2.CurrentY + 150     ' set Y
tForm.Picture2.Print "2"                                    ' print text string to form

' Display third point if indicated
If PictureSnapMode% = 1 Then
tForm.Picture2.DrawWidth = 2
tForm.Picture2.Circle (cpoint3x!, cpoint3y!), radius!, tcolor&

tForm.Picture2.CurrentX = cpoint3x!
tForm.Picture2.CurrentY = cpoint3y!
tForm.Picture2.CurrentX = tForm.Picture2.CurrentX + 150     ' set X
tForm.Picture2.CurrentY = tForm.Picture2.CurrentY + 150     ' set Y
tForm.Picture2.Print "3"                                    ' print text string to form
End If

' Update full window
If tForm3.Visible Then
tWidth! = tForm3.ScaleWidth                                 ' calculate a radius
If tWidth! <> 0# Then
radius! = tWidth! / 150#

' Draw first calibration point for full view window (need to scale to full view form)
Call PictureSnapConvert(Int(2), formx!, formy!, formz!, apoint1x!, apoint1y!, apoint1z!, fractionx!, fractiony!)
tForm3.DrawWidth = 2
tForm3.Circle (tForm3.ScaleWidth * fractionx!, tForm3.ScaleHeight * fractiony!), radius!, tcolor&

tForm3.ForeColor = tcolor&                                  ' set foreground color
tForm3.FontSize = 13                                        ' set font size
tForm3.FontName = LogWindowFontName$
tForm3.FontSize = 13                                        ' set font size    (necessary for Windows)
tForm3.FontBold = False

tForm3.CurrentX = tForm3.ScaleWidth * fractionx!
tForm3.CurrentY = tForm3.ScaleHeight * fractiony!
tForm3.CurrentX = tForm3.CurrentX + 150                     ' set X
tForm3.CurrentY = tForm3.CurrentY + 150                     ' set Y
tForm3.Print "1"                                            ' print text string to form

' Draw second calibration point for full view window (need to scale to full view form)
Call PictureSnapConvert(Int(2), formx!, formy!, formz!, apoint2x!, apoint2y!, apoint2z!, fractionx!, fractiony!)
tForm3.DrawWidth = 2
tForm3.Circle (tForm3.ScaleWidth * fractionx!, tForm3.ScaleHeight * fractiony!), radius!, tcolor&

tForm3.CurrentX = tForm3.ScaleWidth * fractionx!
tForm3.CurrentY = tForm3.ScaleHeight * fractiony!
tForm3.CurrentX = tForm3.CurrentX + 150                     ' set X
tForm3.CurrentY = tForm3.CurrentY + 150                     ' set Y
tForm3.Print "2"                                            ' print text string to form

' Display third point if indicated
If PictureSnapMode% = 1 Then
Call PictureSnapConvert(Int(2), formx!, formy!, formz!, apoint3x!, apoint3y!, apoint3z!, fractionx!, fractiony!)
tForm3.DrawWidth = 2
tForm3.Circle (tForm3.ScaleWidth * fractionx!, tForm3.ScaleHeight * fractiony!), radius!, tcolor&

tForm3.CurrentX = tForm3.ScaleWidth * fractionx!
tForm3.CurrentY = tForm3.ScaleHeight * fractiony!
tForm3.CurrentX = tForm3.CurrentX + 150                     ' set X
tForm3.CurrentY = tForm3.CurrentY + 150                     ' set Y
tForm3.Print "3"                                            ' print text string to form
End If
End If
End If

Exit Sub

' Errors
PictureSnapDisplayCalibrationPointsError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapDisplayCalibrationPoints"
ierror = True
Exit Sub

End Sub

Sub PictureSnapSaveCalibration3(tImage As TypeImage)
' Save calibration for BMP file

ierror = False
On Error GoTo PictureSnapSaveCalibration3Error

' Assume unflipped coordinates
cpoint1x! = tImage.ImageIx% * Screen.TwipsPerPixelX ' reference screen coordinates
cpoint2x! = CSng(0) ' reference screen coordinates

cpoint1y! = tImage.ImageIy% * Screen.TwipsPerPixelY ' reference screen coordinates
cpoint2y! = CSng(0)  ' reference screen coordinates

apoint1x! = tImage.ImageXmax!  ' reference stage coordinates
apoint2x! = tImage.ImageXmin!  ' reference stage coordinates

apoint1y! = tImage.ImageYmin!  ' reference stage coordinates (flipped for BMP)
apoint2y! = tImage.ImageYmax!  ' reference stage coordinates (flipped for BMP)

' Zero out the third point since it is not used for simple images
cpoint3x! = 0#
cpoint3y! = 0#
cpoint3z! = 0#

apoint3x! = 0#
apoint3y! = 0#
apoint3z! = 0#

PictureSnap_keV! = tImage.ImageKilovolts!
PictureSnap_mag! = tImage.ImageMag!
PictureSnap_scanrota! = tImage.ImageScanRotation!

ACQScreenDPI_Stored! = tImage.ImageDisplayDPI!          ' export of BMP image will utilize the current screen DPI value, not the stored screen DPI value

' Store image signal info
PictureSnap_signal_type$ = tImage.ImageChannelName$
PictureSnap_signal_min! = tImage.ImageZmin&
PictureSnap_signal_max! = tImage.ImageZmax&

Exit Sub

' Errors
PictureSnapSaveCalibration3Error:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapSaveCalibration3"
ierror = True
Exit Sub

End Sub

Sub PictureSnapConvertFiducialsCalculate(mode As Integer, formx As Single, formy As Single, formz As Single, stagex As Single, stagey As Single, stagez As Single)
' Calculate the three coordinate fiducial transformation
'  mode = 1 calculate transform from screen coordinates to stage coordinates
'  mode = 2 calculate transform from stage coordinates to screen coordinates

ierror = False
On Error GoTo PictureSnapConvertFiducialsCalculateError

Dim i As Integer

' Load "from" fiducials
If mode% = 1 Then       ' screen to stage
For i% = 1 To MAXDIM%
If i% = 1 Then fiducialold!(1, i%) = cpoint1x!
If i% = 1 Then fiducialold!(2, i%) = cpoint1y!
If i% = 2 Then fiducialold!(1, i%) = cpoint2x!
If i% = 2 Then fiducialold!(2, i%) = cpoint2y!
If i% = 3 Then fiducialold!(1, i%) = cpoint3x!
If i% = 3 Then fiducialold!(2, i%) = cpoint3y!
fiducialold!(3, i%) = 0#
Next i%

Else                    ' stage to screen
For i% = 1 To MAXDIM%
If i% = 1 Then fiducialold!(1, i%) = apoint1x!
If i% = 1 Then fiducialold!(2, i%) = apoint1y!
If i% = 2 Then fiducialold!(1, i%) = apoint2x!
If i% = 2 Then fiducialold!(2, i%) = apoint2y!
If i% = 3 Then fiducialold!(1, i%) = apoint3x!
If i% = 3 Then fiducialold!(2, i%) = apoint3y!

' Load Z stage positions
If PictureSnapMode% = 1 And PictureSnapCalibrationNumberofZPoints% > 0 Then
If i% = 1 Then fiducialold!(3, i%) = apoint1z!
If i% = 2 Then fiducialold!(3, i%) = apoint2z!
If i% = 3 Then fiducialold!(3, i%) = apoint3z!
Else
If i% = 1 Then fiducialold!(3, i%) = RealTimeMotorPositions!(ZMotor%)
If i% = 2 Then fiducialold!(3, i%) = RealTimeMotorPositions!(ZMotor%)
If i% = 3 Then fiducialold!(3, i%) = RealTimeMotorPositions!(ZMotor%)
End If
Next i%
End If

' Load "to" fiducials
If mode% = 1 Then       ' screen to stage
For i% = 1 To MAXDIM%
If i% = 1 Then fiducialnew!(1, i%) = apoint1x!
If i% = 1 Then fiducialnew!(2, i%) = apoint1y!
If i% = 2 Then fiducialnew!(1, i%) = apoint2x!
If i% = 2 Then fiducialnew!(2, i%) = apoint2y!
If i% = 3 Then fiducialnew!(1, i%) = apoint3x!
If i% = 3 Then fiducialnew!(2, i%) = apoint3y!

' Load Z stage positions
If PictureSnapMode% = 1 And PictureSnapCalibrationNumberofZPoints% > 0 Then
If i% = 1 Then fiducialnew!(3, i%) = apoint1z!
If i% = 2 Then fiducialnew!(3, i%) = apoint2z!
If i% = 3 Then fiducialnew!(3, i%) = apoint3z!
Else
If i% = 1 Then fiducialnew!(3, i%) = RealTimeMotorPositions!(ZMotor%)
If i% = 2 Then fiducialnew!(3, i%) = RealTimeMotorPositions!(ZMotor%)
If i% = 3 Then fiducialnew!(3, i%) = RealTimeMotorPositions!(ZMotor%)
End If
Next i%

Else                    ' stage to screen
For i% = 1 To MAXDIM%
If i% = 1 Then fiducialnew!(1, i%) = cpoint1x!
If i% = 1 Then fiducialnew!(2, i%) = cpoint1y!
If i% = 2 Then fiducialnew!(1, i%) = cpoint2x!
If i% = 2 Then fiducialnew!(2, i%) = cpoint2y!
If i% = 3 Then fiducialnew!(1, i%) = cpoint3x!
If i% = 3 Then fiducialnew!(2, i%) = cpoint3y!
fiducialnew!(3, i%) = 0
Next i%
End If

' Calculate transformation matrix
If DebugMode And VerboseMode Then
Call Trans3dCalculateMatrixVector(Int(1), fiducialold!(), fiducialnew!(), fiducialtranslation#(), fiducialmatrix#())
If ierror Then Exit Sub
Else
Call Trans3dCalculateMatrixVector(Int(0), fiducialold!(), fiducialnew!(), fiducialtranslation#(), fiducialmatrix#())
If ierror Then Exit Sub
End If

' Transform coordinates
If mode% = 1 Then   ' screen to stage
stagex! = formx!
stagey! = formy!
stagez! = formz!
Call PictureSnapConvertFiducialsConvert(stagex!, stagey!, stagez!)
If ierror Then Exit Sub

Else                ' stage to screen
formx! = stagex!
formy! = stagey!
formz! = stagez!
Call PictureSnapConvertFiducialsConvert(formx!, formy!, formz!)
If ierror Then Exit Sub
End If

Exit Sub

' Errors
PictureSnapConvertFiducialsCalculateError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapConvertFiducialsCalculate"
ierror = True
Exit Sub

End Sub

Sub PictureSnapConvertFiducialsConvert(convx As Single, convy As Single, convz As Single)
' Do a three coordinate fiducial transformation for passed coordinate

ierror = False
On Error GoTo PictureSnapConvertFiducialsConvertError

ReDim xyz(1 To MAXAXES%) As Single

' Load coordinate
xyz!(1) = convx!
xyz!(2) = convy!
xyz!(3) = convz!

' Transform coordinate
Call Trans3dTransformPositionVector(fiducialtranslation#(), fiducialmatrix#(), xyz!())
If ierror Then GoTo PictureSnapConvertFiducialsConvertBadTransform

convx! = xyz!(1)
convy! = xyz!(2)
convz! = xyz!(3)

Exit Sub

' Errors
PictureSnapConvertFiducialsConvertError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapConvertFiducialsConvert"
ierror = True
Exit Sub

PictureSnapConvertFiducialsConvertBadTransform:
msg$ = "Bad coordinate transformation, the position will not be transformed"
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapConvertFiducialsConvert"
ierror = True
Exit Sub

End Sub

Sub PictureSnapUnStretch(mode As Integer, tX!, tY!, tImage As Image)
' Convert the stretched image mouse coordinates to unstretched units for subsequent stage conversion
'  mode = 0 from image control (stretch = true) to unstretched FormPICTURESNAP.Picture2 twips (.ACQ file units)
'  mode = 1 from unstretched FormPICTURESNAP.Picture2 twips (.ACQ file units) to image control (stretch = true)

ierror = False
On Error GoTo PictureSnapUnStretchError

' Convert X and Y (image width and height are always zero based)
If mode% = 0 Then
tX! = tX! * FormPICTURESNAP.Picture2.Width / tImage.Width
tY! = tY! * FormPICTURESNAP.Picture2.Height / tImage.Height

Else
tX! = tX! / (FormPICTURESNAP.Picture2.Width / tImage.Width)
tY! = tY! / (FormPICTURESNAP.Picture2.Height / tImage.Height)
End If

Exit Sub

' Errors
PictureSnapUnStretchError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapUnStretch"
ierror = True
Exit Sub

End Sub

Sub PictureSnapSendCalibration(pmode As Integer, cpoint1() As Single, cpoint2() As Single, cpoint3() As Single, apoint1() As Single, apoint2() As Single, apoint3() As Single, keV As Single, mag As Single, scanrota As Single)
' Load the passed calibration (used by Secondary.bas in CalcZAF and BeamCalibrate.bas (BeamCalibrateSaveACQ) in CalcImage)

ierror = False
On Error GoTo PictureSnapSendCalibrationError

' Load mode (0 = 2 point, 1 = 3 point)
PictureSnapMode% = pmode%

' Load to calibration variables
cpoint1x! = cpoint1!(1)  ' reference screen coordinates
cpoint1y! = cpoint1!(2)  ' reference screen coordinates
cpoint2x! = cpoint2!(1)  ' reference screen coordinates
cpoint2y! = cpoint2!(2)  ' reference screen coordinates

cpoint1z! = cpoint1!(3)  ' reference screen coordinates
cpoint2z! = cpoint2!(3)  ' reference screen coordinates

apoint1x! = apoint1!(1)  ' reference stage coordinates
apoint1y! = apoint1!(2)  ' reference stage coordinates
apoint2x! = apoint2!(1)  ' reference stage coordinates
apoint2y! = apoint2!(2)  ' reference stage coordinates

apoint1z! = apoint1!(3)  ' reference stage coordinates
apoint2z! = apoint2!(3)  ' reference stage coordinates

' Load third point
cpoint3x! = cpoint3!(1)  ' reference screen coordinates
cpoint3y! = cpoint3!(2)  ' reference screen coordinates
cpoint3z! = cpoint3!(3)  ' reference screen coordinates

apoint3x! = apoint3!(1)  ' reference stage coordinates
apoint3y! = apoint3!(2)  ' reference stage coordinates
apoint3z! = apoint3!(3)  ' reference stage coordinates

' Load conditions to form variables
PictureSnap_keV! = keV!
PictureSnap_mag! = mag!
PictureSnap_scanrota! = scanrota!

' Set global flag
PictureSnapCalibrated = True

' Load fake filename for Secondary.bas and BeamCalibrate.bas
PictureSnapFilename$ = ApplicationCommonAppData$ & "GRDInfo.ini"

Exit Sub

' Errors
PictureSnapSendCalibrationError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapSendCalibration"
ierror = True
Exit Sub

End Sub

Sub PictureSnapCalculateHFW(hfw As Single, mag As Single)
' Calculates the HFW for the current ACQ file (handles all types of calibrations)

ierror = False
On Error GoTo PictureSnapCalculateHFWError

Dim xmin As Single, xmax As Single
Dim formy As Single, formz As Single
Dim stagey As Single, stagez As Single
Dim fractionx As Single, fractiony As Single

' First calculate stage coordinates of xmin and xmax
Call PictureSnapConvert(Int(1), 0#, formy!, formz!, xmin!, stagey!, stagez!, fractionx!, fractiony!)
If ierror Then Exit Sub

Call PictureSnapConvert(Int(1), FormPICTURESNAP.Picture2.Width, formy!, formz!, xmax!, stagey!, stagez!, fractionx!, fractiony!)
If ierror Then Exit Sub

hfw! = Abs(xmax! - xmin!) * MotUnitsToAngstromMicrons!(XMotor%)

' Convert from um field of view
mag! = (ImageDisplaySizeInCentimeters! * MICRONSPERCM&) / hfw!

Exit Sub

' Errors
PictureSnapCalculateHFWError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapCalculateHFW"
ierror = True
Exit Sub

End Sub

Sub PictureSnapReturnXYMinMax(txmin As Single, txmax As Single, tymin As Single, tymax As Single, tz1 As Single, tz2 As Single, tz3 As Single, tz4 As Single)
' Returns the stage X/Y min/max for the currently loaded ACQ file (NOTE: assumes that loaded calibration represents the full image
' width- not always true for manually calibrated images).
'  txmin = x stage min
'  txmax = x stage min
'  tymin = y stage min
'  tymax = y stage max
'  tz1, tz2, tz3, tz4 = z stage coordinates

ierror = False
On Error GoTo PictureSnapReturnXYMinMaxError

txmin! = apoint1x!
txmax! = apoint2x!
tymin! = apoint1y!
tymax! = apoint2y!

tz1! = apoint1z!
tz2! = apoint2z!
tz3! = apoint3z!
tz4! = 0#            ' not used by PictureSnap

Exit Sub

' Errors
PictureSnapReturnXYMinMaxError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapReturnXYMinMax"
ierror = True
Exit Sub

End Sub

Sub PictureSnapSaveCalibration2(tfilename As String, tImage As TypeImage)
' Save the picture calibration for an image structure (called from ImageSaveAs) for BMP file export

ierror = False
On Error GoTo PictureSnapSaveCalibration2Error

Dim tPictureSnapMode As Integer
Dim tPictureSnapCalibrated As Boolean
Dim tPictureSnapCalibrationSaved As Boolean
Dim tPictureSnapFilename As String

Dim afilename As String

' Save the passed image calibration for saving a BMP file
Call PictureSnapSaveCalibration3(tImage)
If ierror Then Exit Sub

' Save current picture filename (if any)
tPictureSnapMode% = PictureSnapMode%
tPictureSnapCalibrated = PictureSnapCalibrated
tPictureSnapCalibrationSaved = PictureSnapCalibrationSaved
tPictureSnapFilename$ = PictureSnapFilename$

PictureSnapMode% = 0                    ' only two point calibration supported for export of BMP file
PictureSnapCalibrated = True
PictureSnapCalibrationSaved = False
PictureSnapFilename$ = MiscGetFileNameNoExtension$(tfilename$) & ".BMP" ' just a dummy extension
Call PictureSnapSaveCalibration(Int(1), PictureSnapFilename$, PictureSnapCalibrationSaved)

PictureSnapMode% = tPictureSnapMode%    ' restore original PictureSnap mode
PictureSnapCalibrated = tPictureSnapCalibrated    ' restore original PictureSnap calibration flag
PictureSnapCalibrationSaved = tPictureSnapCalibrationSaved
PictureSnapFilename$ = tPictureSnapFilename$
If ierror Then Exit Sub

' Reload original picture calibration (if any)
If PictureSnapFilename$ <> vbNullString Then
afilename$ = MiscGetFileNameNoExtension$(PictureSnapFilename$) & ".acq"
If Dir$(afilename$) <> vbNullString Then
Call PictureSnapReadCalibration(afilename$)
If ierror Then Exit Sub
End If
End If

Exit Sub

' Errors
PictureSnapSaveCalibration2Error:
PictureSnapFilename$ = vbNullString
PictureSnapCalibrated = False
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapSaveCalibration2"
ierror = True
Exit Sub

End Sub

Sub PictureSnapDisplayImageFOVs(tForm As Form, tForm3 As Form)
' Display the acquire image FOVs on the PictureSnap form and the full view window form.

ierror = False
On Error GoTo PictureSnapDisplayImageFOVsError

Dim n As Integer

Dim fractionx1 As Single, fractiony1 As Single
Dim fractionx2 As Single, fractiony2 As Single
Dim formx1 As Single, formy1 As Single, formz1 As Single
Dim formx2 As Single, formy2 As Single, formz2 As Single
Dim tcolor As Long

' Check if image loaded
If PictureSnapFilename$ = vbNullString Then Exit Sub

' Check if calibrated
If Not PictureSnapCalibrated Then Exit Sub

If NumberOfImages% < 1 Then Exit Sub

' Loop on all acquired images
For n% = 1 To NumberOfImages%

' Convert image stage extents to screen coordinates
Call PictureSnapConvert(Int(2), formx1!, formy1!, formz1!, ImageXMins!(n%), ImageYMins!(n%), CSng(0#), fractionx1!, fractiony1!)
If ierror Then Exit Sub

Call PictureSnapConvert(Int(2), formx2!, formy2!, formz2!, ImageXMaxs!(n%), ImageYMaxs!(n%), CSng(0#), fractionx2!, fractiony2!)
If ierror Then Exit Sub

' Check screen extents
If formx1! > 0 And formy1! > 0 And formx2! > 0 And formy1! > 0 Then
If formx1! < tForm.Picture2.ScaleWidth And formy1! < tForm.Picture2.ScaleHeight And formx2! < tForm.Picture2.ScaleWidth And formy1! < tForm.Picture2.ScaleHeight Then

' Calculate a radius
If tForm.Width <> 0 Then
tcolor& = RGB(255, 0, 255)

' Draw image extents on FormPICTURESNAP
tForm.Picture2.DrawWidth = 2
tForm.Picture2.Line (formx1!, formy1!)-(formx2!, formy2!), tcolor&, B

tForm.Picture2.ForeColor = tcolor&                          ' set foreground color
tForm.Picture2.FontSize = 13                                ' set font size
tForm.Picture2.FontName = LogWindowFontName$
tForm.Picture2.FontSize = 13                                ' set font size    (necessary for Windows)
tForm.Picture2.FontBold = False

tForm.Picture2.CurrentX = formx2!
tForm.Picture2.CurrentY = formy2!
tForm.Picture2.CurrentX = tForm.Picture2.CurrentX + 150     ' set X
tForm.Picture2.CurrentY = tForm.Picture2.CurrentY + 150     ' set Y
tForm.Picture2.Print Format$(n%)                            ' print text string to form

End If

' Draw image extents on full view form
If tForm3.Visible Then
If tForm3.ScaleWidth <> 0 Then

tForm3.DrawWidth = 2
tForm3.Line (tForm3.ScaleWidth * fractionx1!, tForm3.ScaleHeight * fractiony1!)-(tForm3.ScaleWidth * fractionx2!, tForm3.ScaleHeight * fractiony2!), tcolor&, B

tForm3.ForeColor = tcolor&                          ' set foreground color
tForm3.FontSize = 13                                ' set font size
tForm3.FontName = LogWindowFontName$
tForm3.FontSize = 13                                ' set font size    (necessary for Windows)
tForm3.FontBold = False

tForm3.CurrentX = tForm3.ScaleWidth * fractionx2!
tForm3.CurrentY = tForm3.ScaleHeight * fractiony2!
tForm3.CurrentX = tForm3.CurrentX + 150             ' set X
tForm3.CurrentY = tForm3.CurrentY + 150             ' set Y
tForm3.Print Format$(n%)                            ' print text string to form
End If
End If

End If
End If

Next n%

Exit Sub

' Errors
PictureSnapDisplayImageFOVsError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapDisplayImageFOVs"
ierror = True
Exit Sub

End Sub

Sub PictureSnapSaveMode(Index As Integer)
' Save the PictureSnap mode (0 = two points, 1 = three points)

ierror = False
On Error GoTo PictureSnapSaveModeError

' If changing mode and image was calibrated, re-set
If FormPICTURESNAP2.Visible Then
If PictureSnapCalibrated Then PictureSnapCalibrated = False
End If

' Save PictureSnapMode
PictureSnapMode% = Index%

' Resize calibration form
If Index% = 0 Then
FormPICTURESNAP2.Height = 8445
FormPICTURESNAP2.LabelZStage1.Visible = False
FormPICTURESNAP2.LabelZStage2.Visible = False
FormPICTURESNAP2.LabelZStage3.Visible = False
FormPICTURESNAP2.TextZStage1.Visible = False
FormPICTURESNAP2.TextZStage2.Visible = False
FormPICTURESNAP2.TextZStage3.Visible = False
End If

If Index% = 1 Then
FormPICTURESNAP2.Height = 12495
FormPICTURESNAP2.LabelZStage1.Visible = True
FormPICTURESNAP2.LabelZStage2.Visible = True
FormPICTURESNAP2.LabelZStage3.Visible = True
FormPICTURESNAP2.TextZStage1.Visible = True
FormPICTURESNAP2.TextZStage2.Visible = True
FormPICTURESNAP2.TextZStage3.Visible = True
End If

If PictureSnapCalibrated Then
FormPICTURESNAP2.LabelCalibration.Caption = "Image Is Calibrated"
Else
FormPICTURESNAP2.LabelCalibration.Caption = "Image Is NOT Calibrated"
FormPICTURESNAP2.LabelCalibrationAccuracy.Caption = vbNullString
End If

Exit Sub

' Errors
PictureSnapSaveModeError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapSaveMode"
ierror = True
Exit Sub

End Sub

Sub PictureSnapCalculateRotation()
' Calculate the image rotation (relative to the stage) based on the current calibration

ierror = False
On Error GoTo PictureSnapCalculateRotationError

Dim n As Long

Dim aslope1 As Single, aslope2 As Single, aslope3 As Single                 ' stage coordinate slopes
Dim cslope1 As Single, cslope2 As Single, cslope3 As Single                 ' screen coordinate slopes

Dim arotation1 As Single, arotation2 As Single, arotation3 As Single
Dim crotation1 As Single, crotation2 As Single, crotation3 As Single

Dim trotation As Single, trotation1 As Single, trotation2 As Single, trotation3 As Single
Dim tradians As Single

Const MAX_SLOPE! = 80#

' Check for current image
If PictureSnapFilename$ = vbNullString Then Exit Sub

' Check if calibrated
If Not PictureSnapCalibrated Then Exit Sub

' Assume zero rotation
PictureSnapRotation! = 0#

' Calculate slopes for two or three screen calibration points (cartesian stage)
If Default_X_Polarity% = 0 And Default_Y_Polarity% = 0 Then
If cpoint1x! - cpoint2x! = 0 Then Exit Sub
cslope1! = (cpoint2y! - cpoint1y!) / (cpoint1x! - cpoint2x!)

If PictureSnapMode% = 1 Then
If cpoint2x! - cpoint3x! = 0 Then Exit Sub
If cpoint3x! - cpoint1x! = 0 Then Exit Sub
cslope2! = (cpoint3y! - cpoint2y!) / (cpoint2x! - cpoint3x!)
cslope3! = (cpoint1y! - cpoint3y!) / (cpoint3x! - cpoint1x!)
End If

' Calculate slopes for two or three screen calibration points (anti-cartesian stage)
ElseIf Default_X_Polarity% = -1 And Default_Y_Polarity% = -1 Then
If cpoint2x! - cpoint1x! = 0 Then Exit Sub
cslope1! = (cpoint1y! - cpoint2y!) / (cpoint2x! - cpoint1x!)

If PictureSnapMode% = 1 Then
If cpoint3x! - cpoint2x! = 0 Then Exit Sub
If cpoint1x! - cpoint3x! = 0 Then Exit Sub
cslope2! = (cpoint2y! - cpoint3y!) / (cpoint3x! - cpoint2x!)
cslope3! = (cpoint3y! - cpoint1y!) / (cpoint1x! - cpoint3x!)
End If

' Calculate slopes for two or three screen calibration points (half-cartesian stage)
ElseIf Default_X_Polarity% = 0 And Default_Y_Polarity% = -1 Then
If cpoint1x! - cpoint2x! = 0 Then Exit Sub
cslope1! = (cpoint1y! - cpoint2y!) / (cpoint1x! - cpoint2x!)

If PictureSnapMode% = 1 Then
If cpoint2x! - cpoint3x! = 0 Then Exit Sub
If cpoint3x! - cpoint1x! = 0 Then Exit Sub
cslope2! = (cpoint2y! - cpoint3y!) / (cpoint2x! - cpoint3x!)
cslope3! = (cpoint3y! - cpoint1y!) / (cpoint3x! - cpoint1x!)
End If

End If

' Now calculate angles for each set of screen calibration points
tradians! = Atn(cslope1!)
crotation1! = tradians! * 180 / PI!

If PictureSnapMode% = 1 Then
tradians! = Atn(cslope2!)
crotation2! = tradians! * 180 / PI!
tradians! = Atn(cslope3!)
crotation3! = tradians! * 180 / PI!
End If

' Calculate slopes for two or three stage calibration points
If apoint1x! - apoint2x! = 0 Then Exit Sub
aslope1! = (apoint1y! - apoint2y!) / (apoint1x! - apoint2x!)

If PictureSnapMode% = 1 Then
If apoint2x! - apoint3x! = 0 Then Exit Sub
If apoint3x! - apoint1x! = 0 Then Exit Sub
aslope2! = (apoint2y! - apoint3y!) / (apoint2x! - apoint3x!)
aslope3! = (apoint3y! - apoint1y!) / (apoint3x! - apoint1x!)
End If

' Now calculate angles for each set of stage calibration points
tradians! = Atn(aslope1!)
arotation1! = tradians! * 180 / PI!

If PictureSnapMode% = 1 Then
tradians! = Atn(aslope2!)
arotation2! = tradians! * 180 / PI!
tradians! = Atn(aslope3!)
arotation3! = tradians! * 180 / PI!
End If

' Now check for bad rotations if two point mode
If PictureSnapMode% = 0 And (Abs(crotation1!) > MAX_SLOPE! Or Abs(arotation1!) > MAX_SLOPE!) Then GoTo PictureSnapCalculateRotationBadSlope

' Calculate the change in rotation between screen and stage calibrations
trotation1! = arotation1! - crotation1!
trotation! = trotation1!

If PictureSnapMode% = 1 Then
trotation2! = arotation2! - crotation2!
trotation3! = arotation3! - crotation3!

' Average the rotations of the calibration points
n& = 0
trotation! = 0
If Abs(crotation1!) < MAX_SLOPE! And Abs(arotation1!) < MAX_SLOPE! Then
trotation! = trotation! + trotation1!
n& = n& + 1
End If
If Abs(crotation2!) < MAX_SLOPE! And Abs(arotation2!) < MAX_SLOPE! Then
trotation! = trotation! + trotation2!
n& = n& + 1
End If
If Abs(crotation3!) < MAX_SLOPE! And Abs(arotation3!) < MAX_SLOPE! Then
trotation! = trotation! + trotation3!
n& = n& + 1
End If

' Calculate average rotation
If n& = 0 Then Exit Sub
trotation! = trotation! / n&
End If

' Load the rotation to a global
PictureSnapRotation! = trotation!

' Output calculations
If DebugMode Then
Call IOWriteLog(vbNullString$)
msg$ = "Stage to Image rotation #1: " & Format$(arotation1!) & " - " & Format$(crotation1!) & " = " & Format$(trotation1!)
Call IOWriteLog(msg$)

If PictureSnapMode% = 1 Then
msg$ = "Stage to Image rotation #2: " & Format$(arotation2!) & " - " & Format$(crotation2!) & " = " & Format$(trotation2!)
Call IOWriteLog(msg$)
msg$ = "Stage to Image rotation #3: " & Format$(arotation3!) & " - " & Format$(crotation3!) & " = " & Format$(trotation3!)
Call IOWriteLog(msg$)

Call IOWriteLog(vbNullString$)
msg$ = "Stage to Image rotation average: " & Format$(trotation!)
Call IOWriteLog(msg$)
End If
End If

Exit Sub

' Errors
PictureSnapCalculateRotationError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapCalculateRotation"
ierror = True
Exit Sub

PictureSnapCalculateRotationBadSlope:
msg$ = "The angle between the two calibration points is insufficiently diagonal. Please pick calibration points with slopes closer to 45 degrees for best results."
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapCalculateRotation"
ierror = True
Exit Sub

End Sub

