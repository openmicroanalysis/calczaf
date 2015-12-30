Attribute VB_Name = "CodePictureSnap3"
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

Global PictureSnapFilename As String
Global PictureSnapMode As Integer       ' 0 = normal two point conversion, 1 = three point transformation
Global PictureSnapCalibrationNumberofZPoints As Integer
Global PictureSnapCalibrated As Boolean
Global PictureSnapCalibrationSaved As Boolean

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

' Form variables for keV and mag (and scan rotation)
Dim keV As Single, mag As Single, scan As Single

Sub PictureSnapReadCalibration(tfilename$)
' Read the screen and stage calibration from the passed ACQ file

ierror = False
On Error GoTo PictureSnapReadCalibrationError

Dim pmode As Single
Dim points As Single
Dim zpoints As Single

Dim gX_Polarity As Integer, gY_Polarity As Integer
Dim gStage_Units As String

' Dimension coordinates (assume using three points and XYZ)
Dim cpoint1(1 To 3) As Single, cpoint2(1 To 3) As Single, cpoint3(1 To 3) As Single
Dim apoint1(1 To 3) As Single, apoint2(1 To 3) As Single, apoint3(1 To 3) As Single

' Read parameters and calibration points from INI style ACQ file
Call InitINIReadWriteScaler(Int(1), tfilename$, "stage", "PictureSnap mode", pmode!)
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

If zpoints! > 0 Then
Call InitINIReadWriteScaler(Int(1), tfilename$, "stage", "Screen Z reference point1 (dummy)", cpoint1!(3))
If ierror Then Exit Sub
Call InitINIReadWriteScaler(Int(1), tfilename$, "stage", "Screen Z reference point2 (dummy)", cpoint2!(3))
If ierror Then Exit Sub
End If

' Read stage calibration points
Call InitINIReadWriteArray(Int(1), tfilename$, "stage", "Stage reference point1", Int(2), apoint1!())
If ierror Then Exit Sub
Call InitINIReadWriteArray(Int(1), tfilename$, "stage", "Stage reference point2", Int(2), apoint2!())
If ierror Then Exit Sub

If zpoints! > 0 Then
Call InitINIReadWriteScaler(Int(1), tfilename$, "stage", "Stage Z reference point1", apoint1!(3))
If ierror Then Exit Sub
Call InitINIReadWriteScaler(Int(1), tfilename$, "stage", "Stage Z reference point2", apoint2!(3))
If ierror Then Exit Sub
End If

' Read 3rd point if indicated
If pmode! = 1 Then
Call InitINIReadWriteArray(Int(1), tfilename$, "stage", "Screen reference point3 (twips)", Int(2), cpoint3!())
If ierror Then Exit Sub
Call InitINIReadWriteArray(Int(1), tfilename$, "stage", "Stage reference point3", Int(2), apoint3!())
If ierror Then Exit Sub

If zpoints! > 0 Then
Call InitINIReadWriteScaler(Int(1), tfilename$, "stage", "Screen Z reference point3 (dummy)", cpoint3!(3))
If ierror Then Exit Sub
Call InitINIReadWriteScaler(Int(1), tfilename$, "stage", "Stage Z reference point3", apoint3!(3))
If ierror Then Exit Sub
End If
End If

' Load global ("points" is not used)
PictureSnapMode% = CInt(pmode!)
PictureSnapCalibrationNumberofZPoints% = CInt(zpoints!)

' Check for existing GRD info
Call GridCheckGRDInfo(tfilename$, gX_Polarity%, gY_Polarity%, gStage_Units$)
If ierror Then Exit Sub

' Assume no unit conversions
apoint1!(1) = apoint1!(1)      ' X stage (point 1)
apoint1!(2) = apoint1!(2)      ' Y stage (point 1)
apoint2!(1) = apoint2!(1)      ' X stage (point 2)
apoint2!(2) = apoint2!(2)      ' Y stage (point 2)
apoint3!(1) = apoint3!(1)      ' X stage (point 3)
apoint3!(2) = apoint3!(2)      ' Y stage (point 3)

' Load third point
If PictureSnapMode% = 1 Then
apoint1!(3) = apoint1!(3)      ' Z stage (point 1)
apoint2!(3) = apoint2!(3)      ' Z stage (point 2)
apoint3!(3) = apoint3!(3)      ' Z stage (point 3)
End If

' Fix units if necessary
If Default_Stage_Units$ <> gStage_Units$ Then
If Default_Stage_Units$ = "um" And gStage_Units$ = "mm" Then
apoint1!(1) = apoint1!(1) * MICRONSPERMM&     ' X stage
apoint1!(2) = apoint1!(2) * MICRONSPERMM&     ' Y stage
apoint2!(1) = apoint2!(1) * MICRONSPERMM&     ' X stage
apoint2!(2) = apoint2!(2) * MICRONSPERMM&     ' Y stage
apoint3!(1) = apoint3!(1) * MICRONSPERMM&     ' X stage
apoint3!(2) = apoint3!(2) * MICRONSPERMM&     ' Y stage

If PictureSnapMode% = 1 Then
apoint1!(3) = apoint1!(3) * MICRONSPERMM&       ' Z stage (point 1)
apoint2!(3) = apoint2!(3) * MICRONSPERMM&       ' Z stage (point 2)
apoint3!(3) = apoint3!(3) * MICRONSPERMM&       ' Z stage (point 3)
End If
End If

If Default_Stage_Units$ = "mm" And gStage_Units$ = "um" Then
apoint1!(1) = apoint1!(1) / MICRONSPERMM&     ' X stage
apoint1!(2) = apoint1!(2) / MICRONSPERMM&     ' Y stage
apoint2!(1) = apoint2!(1) / MICRONSPERMM&     ' X stage
apoint2!(2) = apoint2!(2) / MICRONSPERMM&     ' Y stage
apoint3!(1) = apoint3!(1) / MICRONSPERMM&     ' X stage
apoint3!(2) = apoint3!(2) / MICRONSPERMM&     ' Y stage

If PictureSnapMode% = 1 Then
apoint1!(3) = apoint1!(3) / MICRONSPERMM&       ' Z stage (point 1)
apoint2!(3) = apoint2!(3) / MICRONSPERMM&       ' Z stage (point 2)
apoint3!(3) = apoint3!(3) / MICRONSPERMM&       ' Z stage (point 3)
End If
End If

End If

' Load to calibration variables
cpoint1x! = cpoint1!(1)  ' x reference screen coordinates
cpoint1y! = cpoint1!(2)  ' y reference screen coordinates
cpoint2x! = cpoint2!(1)  ' x reference screen coordinates
cpoint2y! = cpoint2!(2)  ' y reference screen coordinates
If PictureSnapMode% = 1 Then
cpoint1z! = cpoint1!(3)  ' Z reference screen coordinates
cpoint2z! = cpoint2!(3)  ' Z reference screen coordinates
End If

apoint1x! = apoint1!(1)  ' x reference stage coordinates
apoint2x! = apoint2!(1)  ' x reference stage coordinates

apoint1y! = apoint1!(2)  ' y reference stage coordinates
apoint2y! = apoint2!(2)  ' y reference stage coordinates

' Load proper image orientation
If Default_X_Polarity% <> gX_Polarity% Then
If Default_Y_Polarity = 0 And gY_Polarity = -1 Then
apoint1x! = apoint2!(1)  ' reference stage coordinates
apoint2x! = apoint1!(1)  ' reference stage coordinates
End If
If Default_X_Polarity = -1 And gX_Polarity = 0 Then
apoint1x! = apoint1!(1)  ' reference stage coordinates
apoint2x! = apoint2!(1)  ' reference stage coordinates
End If
End If

If Default_Y_Polarity% <> gY_Polarity% Then
If Default_Y_Polarity = 0 And gY_Polarity = -1 Then
apoint1y! = apoint2!(2)  ' reference stage coordinates
apoint2y! = apoint1!(2)  ' reference stage coordinates
End If
If Default_Y_Polarity = -1 And gY_Polarity = 0 Then
apoint1y! = apoint1!(2)  ' reference stage coordinates
apoint2y! = apoint2!(2)  ' reference stage coordinates
End If
End If

If PictureSnapMode% = 1 Then
apoint1z! = apoint1!(3)  ' Z reference stage coordinates
apoint2z! = apoint2!(3)  ' Z reference stage coordinates
End If

' Load third point
If PictureSnapMode% = 1 Then
cpoint3x! = cpoint3!(1)  ' x reference screen coordinates
cpoint3y! = cpoint3!(2)  ' y reference screen coordinates
cpoint3z! = cpoint3!(3)  ' Z reference screen coordinates

apoint3x! = apoint3!(1)  ' x reference stage coordinates
apoint3y! = apoint3!(2)  ' y reference stage coordinates
apoint3z! = apoint3!(3)  ' Z reference stage coordinates
End If

' Load keV and mag from hidden text fields
keV! = DefaultKiloVolts!
Call InitINIReadWriteScaler(Int(1), tfilename$, "ColumnConditions", "kilovolts", keV!)
If ierror Then Exit Sub

mag! = DefaultMagnification!
Call InitINIReadWriteScaler(Int(1), tfilename$, "ColumnConditions", "magnification", mag!)
If ierror Then Exit Sub

scan! = DefaultScanRotation!
Call InitINIReadWriteScaler(Int(1), tfilename$, "ColumnConditions", "scanrotation", scan!)
If ierror Then Exit Sub

FormPICTURESNAP2.TextkeV.Text = keV!
FormPICTURESNAP2.TextMag.Text = mag!
FormPICTURESNAP2.TextScan.Text = scan!      ' scan rotation

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

If PictureSnapMode% = 1 Then
tForm.TextXStage3.Text = apoint3x!
tForm.TextYStage3.Text = apoint3y!
End If

' Load existing z calibrations (if not zero)
If PictureSnapMode% = 1 And PictureSnapCalibrationNumberofZPoints% > 0 Then
tForm.TextZStage1.Text = apoint1z!
tForm.TextZStage2.Text = apoint2z!
tForm.TextZStage3.Text = apoint3z!
End If

tForm.TextXPixel1.Text = cpoint1x!
tForm.TextYPixel1.Text = cpoint1y!
tForm.TextXPixel2.Text = cpoint2x!
tForm.TextYPixel2.Text = cpoint2y!

If PictureSnapMode% = 1 Then
tForm.TextXPixel3.Text = cpoint3x!
tForm.TextYPixel3.Text = cpoint3y!
End If

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

If PictureSnapMode% = 1 Then
apoint3x! = Val(tForm.TextXStage3.Text)
apoint3y! = Val(tForm.TextYStage3.Text)
End If

' Load z stage coordinates
If PictureSnapMode% = 1 Then
apoint1z! = Val(tForm.TextZStage1.Text)
apoint2z! = Val(tForm.TextZStage2.Text)
apoint3z! = Val(tForm.TextZStage3.Text)
End If

' Load pixel coordinates
cpoint1x! = Val(tForm.TextXPixel1.Text)
cpoint1y! = Val(tForm.TextYPixel1.Text)

cpoint2x! = Val(tForm.TextXPixel2.Text)
cpoint2y! = Val(tForm.TextYPixel2.Text)

If PictureSnapMode% = 1 Then
cpoint3x! = Val(tForm.TextXPixel3.Text)
cpoint3y! = Val(tForm.TextYPixel3.Text)
cpoint1z! = 0#
cpoint2z! = 0#
cpoint3z! = 0#
End If

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

' Save keV and mag from hidden text fields
keV! = Val(FormPICTURESNAP2.TextkeV.Text)
mag! = Val(FormPICTURESNAP2.TextMag.Text)
scan! = Val(FormPICTURESNAP2.TextScan.Text)

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

Dim gX_Polarity As Integer, gY_Polarity As Integer
Dim gStage_Units As String

' Check for open picture snap file
If Trim$(PictureSnapFilename$) = vbNullString Then Exit Sub

' Check for existing GRD info
Call GridCheckGRDInfo(PictureSnapFilename$, gX_Polarity%, gY_Polarity%, gStage_Units$)
If ierror Then Exit Sub

' Convert using two calibration points (no Z stage interpolation)
If PictureSnapMode% = 0 Then

' Check for bad data
If cpoint1x! - cpoint2x! = 0# Then GoTo PictureSnapConvertBadConvert
If cpoint1y! - cpoint2y! = 0# Then GoTo PictureSnapConvertBadConvert

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
If Default_X_Polarity% <> 0 And gX_Polarity% = 0 Then
stagex! = apoint2x! + (apoint1x! - stagex!)
End If

stagey! = csy! * formy! + coy!
If Default_X_Polarity% <> 0 And gX_Polarity% = 0 Then
stagey! = apoint2y! + (apoint1y! - stagey!)
End If

' Convert stage to form
Else
formx! = (stagex! - cox!) / csx!
If Default_X_Polarity% <> 0 And gX_Polarity% = 0 Then
formx! = cpoint2x! + (cpoint1x! - formx!)
End If

formy! = (stagey! - coy!) / csy!
If Default_X_Polarity% <> 0 And gX_Polarity% = 0 Then
formy! = cpoint2y! + (cpoint1y! - formy!)
End If
End If
End If

' Transform using three calibration points (uses Z stage calibration)
If PictureSnapMode% = 1 Then
Call PictureSnapConvertFiducialsCalculate(mode%, formx!, formy!, formz!, stagex!, stagey!, stagez!)
If ierror Then Exit Sub

' Check to see that calculated z position is in range
smallamount! = Abs(MotHiLimits!(ZMotor%) - MotLoLimits!(ZMotor%)) * SMALLAMOUNTFRACTION!     ' to place it inside the stage limits
If stagez! > MotHiLimits!(ZMotor%) Then stagez! = MotHiLimits!(ZMotor) - smallamount!
If stagez! < MotLoLimits!(ZMotor%) Then stagez! = MotLoLimits!(ZMotor) + smallamount!

' Use current z if using two point calibration
Else
stagez! = RealTimeMotorPositions!(ZMotor%)
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
msg$ = "One or both of the calibration points are not valid. Try the calibration again with different points"
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

Dim tfilename As String

Dim gX_Polarity As Integer, gY_Polarity As Integer
Dim gStage_Units As String

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

If PictureSnapMode% = 1 Then
cpoint3!(1) = cpoint3x!  ' reference screen coordinates
cpoint3!(2) = cpoint3y!  ' reference screen coordinates
cpoint1!(3) = cpoint1z!  ' dummy z reference
cpoint2!(3) = cpoint2z!  ' dummy z reference
cpoint3!(3) = cpoint3z!  ' dummy z reference
End If

' Stage coodinates
apoint1!(1) = apoint1x!  ' reference stage coordinates
apoint1!(2) = apoint1y!  ' reference stage coordinates

apoint2!(1) = apoint2x!  ' reference stage coordinates
apoint2!(2) = apoint2y!  ' reference stage coordinates

If PictureSnapMode% = 1 Then
apoint3!(1) = apoint3x!  ' reference stage coordinates
apoint3!(2) = apoint3y!  ' reference stage coordinates
apoint1!(3) = apoint1z!  ' actual z reference
apoint2!(3) = apoint2z!  ' actual z reference
apoint3!(3) = apoint3z!  ' actual z reference
End If

' Write calibration points to INI style ACQ file
tfilename$ = MiscGetFileNameNoExtension$(pFileName$) & ".ACQ"

' Check for existing GRD info
Call GridCheckGRDInfo(pFileName$, gX_Polarity%, gY_Polarity%, gStage_Units$)
If ierror Then Exit Sub

' Save parameters
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "PictureSnap mode", CSng(PictureSnapMode%))
If ierror Then Exit Sub

' Save picture snap mode
If PictureSnapMode% = 0 Then
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "Number of calibration points", CSng(2))
If ierror Then Exit Sub
Else
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "Number of calibration points", CSng(3))
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

If PictureSnapCalibrationNumberofZPoints% > 0 Then
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "Screen Z reference point1 (dummy)", cpoint1!(3))
If ierror Then Exit Sub
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "Screen Z reference point2 (dummy)", cpoint2!(3))
If ierror Then Exit Sub
End If

' Write stage calibration points
Call InitINIReadWriteArray(Int(2), tfilename$, "stage", "Stage reference point1", Int(2), apoint1!())
If ierror Then Exit Sub
Call InitINIReadWriteArray(Int(2), tfilename$, "stage", "Stage reference point2", Int(2), apoint2!())
If ierror Then Exit Sub

If PictureSnapCalibrationNumberofZPoints% > 0 Then
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "Stage Z reference point1", apoint1!(3))
If ierror Then Exit Sub
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "Stage Z reference point2", apoint2!(3))
If ierror Then Exit Sub
End If

' Save 3rd point if indicated
If PictureSnapMode% = 1 Then
Call InitINIReadWriteArray(Int(2), tfilename$, "stage", "Screen reference point3 (twips)", Int(2), cpoint3!())
If ierror Then Exit Sub
Call InitINIReadWriteArray(Int(2), tfilename$, "stage", "Stage reference point3", Int(2), apoint3!())
If ierror Then Exit Sub

If PictureSnapCalibrationNumberofZPoints% > 0 Then
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "Screen Z reference point3 (dummy)", cpoint3!(3))
If ierror Then Exit Sub
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "Stage Z reference point3", apoint3!(3))
If ierror Then Exit Sub
End If
End If

' Now save coordinate system of stage orientation and units
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "X_Polarity", CSng(gX_Polarity%))              ' 1 = read, 2 = write
If ierror Then Exit Sub
Call InitINIReadWriteScaler(Int(2), tfilename$, "stage", "Y_Polarity", CSng(gY_Polarity%))              ' 1 = read, 2 = write
If ierror Then Exit Sub
Call InitINIReadWriteString(Int(1), tfilename$, "stage", "Stage_Units", gStage_Units$, vbNullString)    ' 0 = read, 1 = write
If ierror Then Exit Sub

pcalibrationsaved = True

' Save keV and mag also
Call InitINIReadWriteScaler(Int(2), tfilename$, "ColumnConditions", "kilovolts", keV!)
If ierror Then Exit Sub

Call InitINIReadWriteScaler(Int(2), tfilename$, "ColumnConditions", "Magnification", mag!)
If ierror Then Exit Sub

Call InitINIReadWriteScaler(Int(2), tfilename$, "ColumnConditions", "ScanRotation", scan!)
If ierror Then Exit Sub

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

Dim gX_Polarity As Integer, gY_Polarity As Integer
Dim gStage_Units As String

' Check for existing GRD info
Call GridCheckGRDInfo(PictureSnapFilename$, gX_Polarity%, gY_Polarity%, gStage_Units$)
If ierror Then Exit Sub

' Calculate a radius
tWidth! = tForm.Width
If tWidth! = 0# Then Exit Sub
radius! = tWidth! / 100#

' Convert to form coordinates (and return fractional distance for drawing on full view window)
tForm.Picture2.DrawWidth = 2

' Draw first calibration point
Call PictureSnapConvert(Int(2), formx!, formy!, formz!, apoint1x!, apoint1y!, apoint1z!, fractionx!, fractiony!)
If Default_X_Polarity% <> gX_Polarity% Then
fractionx! = 1# - fractionx!
End If
If Default_Y_Polarity% <> gY_Polarity% Then
fractiony! = 1# - fractiony!
End If
tForm.Picture2.Circle (tForm.Picture2.ScaleWidth * fractionx!, tForm.Picture2.ScaleHeight * fractiony!), radius!, RGB(0, 255, 0)

' Draw second calibration point
Call PictureSnapConvert(Int(2), formx!, formy!, formz!, apoint2x!, apoint2y!, apoint2z!, fractionx!, fractiony!)
If Default_X_Polarity% <> gX_Polarity% Then
fractionx! = 1# - fractionx!
End If
If Default_Y_Polarity% <> gY_Polarity% Then
fractiony! = 1# - fractiony!
End If
tForm.Picture2.Circle (tForm.Picture2.ScaleWidth * fractionx!, tForm.Picture2.ScaleHeight * fractiony!), radius!, RGB(0, 255, 0)

' Display two calibration points (this works since FormPICTURESNAP.Picture2 is 1:1 twips, that is unstretched, but doesn't deal with orientation)
'tForm.Picture2.Circle (cpoint1x!, cpoint1y!), radius!, RGB(0, 255, 0)
'tForm.Picture2.Circle (cpoint2x!, cpoint2y!), radius!, RGB(0, 255, 0)

' Display third point if indicated
If PictureSnapMode% = 1 Then
Call PictureSnapConvert(Int(2), formx!, formy!, formz!, apoint3x!, apoint3y!, apoint3z!, fractionx!, fractiony!)
If Default_X_Polarity% <> gX_Polarity% Then
fractionx! = 1# - fractionx!
End If
If Default_Y_Polarity% <> gY_Polarity% Then
fractiony! = 1# - fractiony!
End If
tForm.Picture2.Circle (tForm.Picture2.ScaleWidth * fractionx!, tForm.Picture2.ScaleHeight * fractiony!), radius!, RGB(0, 255, 0)
'tForm.Picture2.Circle (cpoint3x!, cpoint3y!), radius!, RGB(0, 255, 0)
End If

' Update full window
If tForm3.Visible Then
tWidth! = tForm3.ScaleWidth   ' calculate a radius
If tWidth! <> 0# Then
radius! = (tWidth! / 50#) ^ 0.8

' Convert to form coordinates (and return fractional distance for drawing on full view window)

' Draw first calibration point
Call PictureSnapConvert(Int(2), formx!, formy!, formz!, apoint1x!, apoint1y!, apoint1z!, fractionx!, fractiony!)
If Default_X_Polarity% <> gX_Polarity% Then
fractionx! = 1# - fractionx!
End If
If Default_Y_Polarity% <> gY_Polarity% Then
fractiony! = 1# - fractiony!
End If
tForm3.Circle (tForm3.ScaleWidth * fractionx!, tForm3.ScaleHeight * fractiony!), radius!, RGB(0, 255, 0)

' Draw second calibration point
Call PictureSnapConvert(Int(2), formx!, formy!, formz!, apoint2x!, apoint2y!, apoint2z!, fractionx!, fractiony!)
If Default_X_Polarity% <> gX_Polarity% Then
fractionx! = 1# - fractionx!
End If
If Default_Y_Polarity% <> gY_Polarity% Then
fractiony! = 1# - fractiony!
End If
tForm3.Circle (tForm3.ScaleWidth * fractionx!, tForm3.ScaleHeight * fractiony!), radius!, RGB(0, 255, 0)

' Display third point if indicated
If PictureSnapMode% = 1 Then

' Convert to form coordinates (and return fractional distance for drawing on full view window)
Call PictureSnapConvert(Int(2), formx!, formy!, formz!, apoint3x!, apoint3y!, apoint3z!, fractionx!, fractiony!)
If Default_X_Polarity% <> gX_Polarity% Then
fractionx! = 1# - fractionx!
End If
If Default_Y_Polarity% <> gY_Polarity% Then
fractiony! = 1# - fractiony!
End If
tForm3.Circle (tForm3.ScaleWidth * fractionx!, tForm3.ScaleHeight * fractiony!), radius!, RGB(0, 255, 0)
End If
End If
End If

tForm3.DrawWidth = 1
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
cpoint1x! = CSng(tImage.ImageIx% * Screen.TwipsPerPixelX) ' reference screen coordinates
cpoint2x! = CSng(0) ' reference screen coordinates

cpoint1y! = CSng(tImage.ImageIy% * Screen.TwipsPerPixelY) ' reference screen coordinates
cpoint2y! = CSng(0)  ' reference screen coordinates

apoint1x! = tImage.ImageXmax!  ' reference stage coordinates
apoint2x! = tImage.ImageXmin!  ' reference stage coordinates

apoint1y! = tImage.ImageYmin!  ' reference stage coordinates (flipped for BMP)
apoint2y! = tImage.ImageYmax!  ' reference stage coordinates (flipped for BMP)

keV! = tImage.ImageKilovolts!
mag! = tImage.ImageMag!
scan! = tImage.ImageScanRotation!

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
fiducialold!(4, i%) = 0#      ' W motor position (not used)
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
fiducialold!(4, i%) = RealTimeMotorPositions!(WMotor%)      ' W motor position (not used)
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

fiducialnew!(4, i%) = RealTimeMotorPositions!(WMotor%)      ' W motor position (not used)
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
fiducialnew!(4, i%) = 0      ' W motor position (not used)
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

ReDim xyzw(1 To MAXAXES%) As Single

' Load coordinate
xyzw!(1) = convx!
xyzw!(2) = convy!
xyzw!(3) = convz!
xyzw!(4) = RealTimeMotorPositions!(WMotor%)

' Transform coordinate
Call Trans3dTransformPositionVector(fiducialtranslation#(), fiducialmatrix#(), xyzw!())
If ierror Then GoTo PictureSnapConvertFiducialsConvertBadTransform

convx! = xyzw!(1)
convy! = xyzw!(2)
convz! = xyzw!(3)

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

Sub PictureSnapSendCalibration(pmode As Integer, cpoint1() As Single, cpoint2() As Single, cpoint3() As Single, apoint1() As Single, apoint2() As Single, apoint3() As Single)
' Load the passed calibration (used only by Secondary.bas in CalcZAF)

ierror = False
On Error GoTo PictureSnapSendCalibrationError

' Load mode (0 = 2 point, 1 = 3 point)
PictureSnapMode% = pmode%

' Load to calibration variables
cpoint1x! = cpoint1!(1)  ' reference screen coordinates
cpoint1y! = cpoint1!(2)  ' reference screen coordinates
cpoint2x! = cpoint2!(1)  ' reference screen coordinates
cpoint2y! = cpoint2!(2)  ' reference screen coordinates
If PictureSnapMode% = 1 Then
cpoint1z! = cpoint1!(3)  ' reference screen coordinates
cpoint2z! = cpoint2!(3)  ' reference screen coordinates
End If

apoint1x! = apoint1!(1)  ' reference stage coordinates
apoint1y! = apoint1!(2)  ' reference stage coordinates
apoint2x! = apoint2!(1)  ' reference stage coordinates
apoint2y! = apoint2!(2)  ' reference stage coordinates
If PictureSnapMode% = 1 Then
apoint1z! = apoint1!(3)  ' reference stage coordinates
apoint2z! = apoint2!(3)  ' reference stage coordinates
End If

' Load third point
If PictureSnapMode% = 1 Then
cpoint3x! = cpoint3!(1)  ' reference screen coordinates
cpoint3y! = cpoint3!(2)  ' reference screen coordinates
cpoint3z! = cpoint3!(3)  ' reference screen coordinates

apoint3x! = apoint3!(1)  ' reference stage coordinates
apoint3y! = apoint3!(2)  ' reference stage coordinates
apoint3z! = apoint3!(3)  ' reference stage coordinates
End If

' Set global flag
PictureSnapCalibrated = True

' Load fake filename for Secondary.bas
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
mag! = CSng((ImageDisplaySizeInCentimeters! * MICRONSPERCM&) / hfw!)

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
Dim tPictureSnapFileName As String

Dim afilename As String

' Save the passed image calibration for saving a BMP file
Call PictureSnapSaveCalibration3(tImage)
If ierror Then Exit Sub

' Save current picture filename (if any)
tPictureSnapMode% = PictureSnapMode%
tPictureSnapCalibrated = PictureSnapCalibrated
tPictureSnapCalibrationSaved = PictureSnapCalibrationSaved
tPictureSnapFileName$ = PictureSnapFilename$

PictureSnapMode% = 0                    ' only two point calibration supported for export of BMP file
PictureSnapCalibrated = True
PictureSnapCalibrationSaved = False
PictureSnapFilename$ = MiscGetFileNameNoExtension$(tfilename$) & ".BMP" ' just a dummy extension
Call PictureSnapSaveCalibration(Int(1), PictureSnapFilename$, PictureSnapCalibrationSaved)

PictureSnapMode% = tPictureSnapMode%    ' restore original PictureSnap mode
PictureSnapCalibrated = tPictureSnapCalibrated    ' restore original PictureSnap calibration flag
PictureSnapCalibrationSaved = tPictureSnapCalibrationSaved
PictureSnapFilename$ = tPictureSnapFileName$
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

Sub PictureSnapInvert(mode As Integer)
' Invert the specified x or y min/max
'  mode = 0 invert x
'  mode = 1 invert y

ierror = False
On Error GoTo PictureSnapInvertError

Dim temp As Single

' Invert x
If mode% = 0 Then
temp! = apoint2x!
apoint2x! = apoint1x!
apoint1x! = temp!

Else
temp! = apoint2y!
apoint2y! = apoint1y!
apoint1y! = temp!
End If

Exit Sub

' Errors
PictureSnapInvertError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapInvert"
ierror = True
Exit Sub

End Sub
