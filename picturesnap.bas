Attribute VB_Name = "CodePictureSnap"
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

Global WaitingForCalibrationClick As Integer
Global PictureSnapDisplayCalibrationPointsFlag As Boolean

Dim CurrentPointX As Single
Dim CurrentPointY As Single

' Scale bar variables for scroll event
Dim oldx As Single, oldy As Single

Sub PictureSnapLoad(bmode As Boolean)
' Load the PictureSnap form (automatically load picture if already specified)
'  bmode = True load modeless
'  bmode = False load modal

ierror = False
On Error GoTo PictureSnapLoadError

' Check for stage motors
If NumberOfStageMotors% < 1 Then GoTo PictureSnapLoadNoStage

' Save window load mode to global
PictureSnapWindowIsModeless = bmode

' If picture file is already specified and found then load it
If PictureSnapFilename$ <> vbNullString And Dir$(PictureSnapFilename$) <> vbNullString Then
Screen.MousePointer = vbHourglass
Set FormPICTURESNAP.Picture2 = LoadPicture(PictureSnapFilename$)

' Minimize and restore to re-size
FormPICTURESNAP.WindowState = vbMinimized
DoEvents
FormPICTURESNAP.WindowState = vbNormal
Screen.MousePointer = vbDefault

' Update form caption
If RealTimeMode And PictureSnapCalibrated Then
FormPICTURESNAP.Caption = "PictureSnap [" & PictureSnapFilename$ & "] (double-click to move)"
Else
FormPICTURESNAP.Caption = "PictureSnap [" & PictureSnapFilename$ & "]"
End If

Else
PictureSnapFilename$ = vbNullString
PictureSnapCalibrated = False
End If

' If not realtime then disable some menus
If Not RealTimeMode Then
FormPICTURESNAP.menuDisplayStandards.Enabled = False
FormPICTURESNAP.menuDisplayUnknowns.Enabled = False
FormPICTURESNAP.menuDisplayWavescans.Enabled = False
FormPICTURESNAP.menuDisplayLongLabels.Enabled = False
FormPICTURESNAP.menuDisplayShortLabels.Enabled = False

FormPICTURESNAP.menuMiscUseBeamBlankForStageMotion.Enabled = False
FormPICTURESNAP.menuMiscUseRightMouseClickToDigitize.Enabled = False
FormPICTURESNAP.menuMiscUseLineDrawingMode.Enabled = False
FormPICTURESNAP.menuMiscUseRectangleDrawingMode.Enabled = False
End If

' Check if a calibration file already exists and load if found
If PictureSnapFilename$ <> vbNullString And Dir$(PictureSnapFilename$) <> vbNullString Then
Call PictureSnapLoadCalibration
If ierror Then Exit Sub
End If

' Enable output menus if modeless (and image file is loaded)
Call PictureSnapEnableDisable
If ierror Then Exit Sub

' Check mode parameter and load depending on value
If PictureSnapWindowIsModeless Then
FormPICTURESNAP.Show vbModeless
Else
FormPICTURESNAP.Show vbModal
End If

' Load full view window if visible
If FormPICTURESNAP3.Visible Then
Call PictureSnapLoadFullWindow
If ierror Then Exit Sub
End If

Exit Sub

' Errors
PictureSnapLoadError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapLoad"
ierror = True
Exit Sub

PictureSnapLoadNoStage:
msg$ = "No stage motors are specified for this instrument configuration- this feature is not available"
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapLoad"
ierror = True
Exit Sub

End Sub

Sub PictureSnapSave()
' Save the PictureSnap form parameters

ierror = False
On Error GoTo PictureSnapSaveError

Dim response As Integer

' Check for calibration not saved
If PictureSnapFilename$ <> vbNullString And PictureSnapCalibrated And Not PictureSnapCalibrationSaved Then
msg$ = "Do you want to save the picture calibration for the current picture?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton1, "PictureSnapSave")
If response = vbYes Then
Call PictureSnapSaveCalibration(Int(0), PictureSnapFilename$, PictureSnapCalibrationSaved)
If ierror Then Exit Sub
End If
End If

Exit Sub

' Errors
PictureSnapSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapSave"
ierror = True
Exit Sub

End Sub

Sub PictureSnapCalibrateLoad(method As Integer)
' Load the PictureSnap2 (calibrate) form
'  method = 0 do not show calibration form
'  method = 1 show calibration form

ierror = False
On Error GoTo PictureSnapCalibrateLoadError

If PictureSnapFilename$ = vbNullString Then GoTo PictureSnapCalibrateLoadNoPicture

' Load PictureSnap mode
FormPICTURESNAP2.OptionPictureSnapMode(PictureSnapMode%).value = True

' Load current stage positions
FormPICTURESNAP2.TextXStage1.Text = RealTimeMotorPositions!(XMotor%)
FormPICTURESNAP2.TextYStage1.Text = RealTimeMotorPositions!(YMotor%)

FormPICTURESNAP2.TextXStage2.Text = RealTimeMotorPositions!(XMotor%)
FormPICTURESNAP2.TextYStage2.Text = RealTimeMotorPositions!(YMotor%)

FormPICTURESNAP2.TextXStage3.Text = RealTimeMotorPositions!(XMotor%)
FormPICTURESNAP2.TextYStage3.Text = RealTimeMotorPositions!(YMotor%)

' Load current z stage positions
'If PictureSnapMode% = 1 Then
FormPICTURESNAP2.TextZStage1.Text = RealTimeMotorPositions!(ZMotor%)
FormPICTURESNAP2.TextZStage2.Text = RealTimeMotorPositions!(ZMotor%)
FormPICTURESNAP2.TextZStage3.Text = RealTimeMotorPositions!(ZMotor%)
'End If

' Load default calibration conditions for new calibrations
If Not PictureSnapCalibrated Then
PictureSnap_keV! = DefaultKiloVolts!
PictureSnap_mag! = DefaultMagnificationImaging!
PictureSnap_scanrota! = DefaultScanRotation!
End If

' If the picture is loaded and calibrated, load the existing calibration to the calibration form
If PictureSnapFilename$ <> vbNullString And PictureSnapCalibrated Then
Call PictureSnapCalibrateLoad2(FormPICTURESNAP2)
If ierror Then Exit Sub
End If

FormPICTURESNAP2.TextXPixel1.ForeColor = vbBlack
FormPICTURESNAP2.TextYPixel1.ForeColor = vbBlack
FormPICTURESNAP2.TextXPixel2.ForeColor = vbBlack
FormPICTURESNAP2.TextYPixel2.ForeColor = vbBlack
FormPICTURESNAP2.TextXPixel3.ForeColor = vbBlack
FormPICTURESNAP2.TextYPixel3.ForeColor = vbBlack

If PictureSnapDisplayCalibrationPointsFlag Then
FormPICTURESNAP2.CommandDisplayCalibrationPoints.Caption = "Do Not Display Calibration Points"
Else
FormPICTURESNAP2.CommandDisplayCalibrationPoints.Caption = "Display Calibration Points"
End If

If PictureSnapCalibrated Then
FormPICTURESNAP2.LabelCalibration.Caption = "Image Is Calibrated"
Else
FormPICTURESNAP2.LabelCalibration.Caption = "Image Is NOT Calibrated"
FormPICTURESNAP2.LabelCalibrationAccuracy.Caption = vbNullString
End If

' Disable stage buttons if not realtime mode
If Not RealTimeMode Then
FormPICTURESNAP2.CommandMoveTo1.Enabled = False
FormPICTURESNAP2.CommandMoveTo2.Enabled = False
FormPICTURESNAP2.CommandMoveTo3.Enabled = False
FormPICTURESNAP2.CommandReadCurrentStageCoordinate1.Enabled = False
FormPICTURESNAP2.CommandReadCurrentStageCoordinate2.Enabled = False
FormPICTURESNAP2.CommandReadCurrentStageCoordinate3.Enabled = False
FormPICTURESNAP2.CommandLightModeOff.Enabled = False
FormPICTURESNAP2.CommandLightModeOn.Enabled = False
FormPICTURESNAP2.CommandLightModeReflected.Enabled = False
FormPICTURESNAP2.CommandLightModeTransmitted.Enabled = False
Else
FormPICTURESNAP2.CommandMoveTo1.Enabled = True
FormPICTURESNAP2.CommandMoveTo2.Enabled = True
FormPICTURESNAP2.CommandMoveTo3.Enabled = True
FormPICTURESNAP2.CommandReadCurrentStageCoordinate1.Enabled = True
FormPICTURESNAP2.CommandReadCurrentStageCoordinate2.Enabled = True
FormPICTURESNAP2.CommandReadCurrentStageCoordinate3.Enabled = True
FormPICTURESNAP2.CommandLightModeOff.Enabled = True
FormPICTURESNAP2.CommandLightModeOn.Enabled = True
FormPICTURESNAP2.CommandLightModeReflected.Enabled = True
FormPICTURESNAP2.CommandLightModeTransmitted.Enabled = True
End If

' Load the form
If method% = 1 Then
FormPICTURESNAP2.Show vbModeless
End If

Exit Sub

' Errors
PictureSnapCalibrateLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapCalibrateLoad"
ierror = True
Exit Sub

PictureSnapCalibrateLoadNoPicture:
msg$ = "No image has been loaded in the PictureSnap window. Please open a sample image using the File | Open menu."
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapCalibrateLOad"
ierror = True
Exit Sub

End Sub

Sub PictureSnapCalibrate(mode As Integer)
' Calculate the picture calibration (stage registration)
'  mode = 0 confirm with user
'  mode = 1 do not confirm with user

ierror = False
On Error GoTo PictureSnapCalibrateError

Dim formx As Single, formy As Single, formz As Single
Dim d As Double

Dim tilt As Single
Dim astring As String

ReDim acoeff(1 To MAXCOEFF%) As Single
ReDim dxdata(1 To MAXDIM%) As Single
ReDim dydata(1 To MAXDIM%) As Single
ReDim dzdata(1 To MAXDIM%) As Single

' Check for picture loaded
If Trim$(PictureSnapFilename$) = vbNullString Then GoTo PictureSnapCalibrateNoPicture

' Check for in bounds
If RealTimeMode Then
If Not MiscMotorInBounds(XMotor%, Val(FormPICTURESNAP2.TextXStage1.Text)) Then GoTo PictureSnapCalibrateOutofBoundsX
If Not MiscMotorInBounds(YMotor%, Val(FormPICTURESNAP2.TextYStage1.Text)) Then GoTo PictureSnapCalibrateOutofBoundsY

If Not MiscMotorInBounds(XMotor%, Val(FormPICTURESNAP2.TextXStage2.Text)) Then GoTo PictureSnapCalibrateOutofBoundsX
If Not MiscMotorInBounds(YMotor%, Val(FormPICTURESNAP2.TextYStage2.Text)) Then GoTo PictureSnapCalibrateOutofBoundsY

If FormPICTURESNAP2.OptionPictureSnapMode(1).value = True Then
If Not MiscMotorInBounds(XMotor%, Val(FormPICTURESNAP2.TextXStage3.Text)) Then GoTo PictureSnapCalibrateOutofBoundsX
If Not MiscMotorInBounds(YMotor%, Val(FormPICTURESNAP2.TextYStage3.Text)) Then GoTo PictureSnapCalibrateOutofBoundsY
End If

If FormPICTURESNAP2.OptionPictureSnapMode(1).value = True And NumberOfStageMotors% > 2 Then
If Not MiscMotorInBounds(ZMotor%, Val(FormPICTURESNAP2.TextZStage1.Text)) Then GoTo PictureSnapCalibrateOutofBoundsZ
If Not MiscMotorInBounds(ZMotor%, Val(FormPICTURESNAP2.TextZStage2.Text)) Then GoTo PictureSnapCalibrateOutofBoundsZ
If Not MiscMotorInBounds(ZMotor%, Val(FormPICTURESNAP2.TextZStage3.Text)) Then GoTo PictureSnapCalibrateOutofBoundsZ
End If
End If

' Check for excessive tilt if 3 point calibration
If FormPICTURESNAP2.OptionPictureSnapMode(1).value = True Then
dxdata!(1) = Val(FormPICTURESNAP2.TextXStage1.Text)
dydata!(1) = Val(FormPICTURESNAP2.TextYStage1.Text)
dzdata!(1) = Val(FormPICTURESNAP2.TextZStage1.Text)

dxdata!(2) = Val(FormPICTURESNAP2.TextXStage2.Text)
dydata!(2) = Val(FormPICTURESNAP2.TextYStage2.Text)
dzdata!(2) = Val(FormPICTURESNAP2.TextZStage2.Text)

dxdata!(3) = Val(FormPICTURESNAP2.TextXStage3.Text)
dydata!(3) = Val(FormPICTURESNAP2.TextYStage3.Text)
dzdata!(3) = Val(FormPICTURESNAP2.TextZStage3.Text)

' Fit data
Call Plan3dCalculate(Int(3), dxdata!(), dydata!(), dzdata!(), acoeff!(), d#)
If ierror Then Exit Sub

' Calculate sample tilt
Call Plan3dCalculateTilt(acoeff!(), tilt!, astring$)
If ierror Then Exit Sub

' Inform user of sample tilt if greater than 0.5 degrees
If tilt! > 0.5 Then
MsgBox astring$, vbOKOnly + vbInformation, "PictureSnapCalibrate"
End If
End If

' Save the form variables
Call PictureSnapCalibrateSave(FormPICTURESNAP2)
If ierror Then Exit Sub

' Check stage calibration is orthogonal
Call PictureSnapCalibrateCheck
If ierror Then Exit Sub

FormPICTURESNAP.Caption = "PictureSnap [" & PictureSnapFilename$ & "] (double-click to move)"
PictureSnapCalibrated = True

If PictureSnapCalibrated Then
FormPICTURESNAP2.LabelCalibration.Caption = "Image Is Calibrated"
Else
FormPICTURESNAP2.LabelCalibration.Caption = "Image Is NOT Calibrated"
FormPICTURESNAP2.LabelCalibrationAccuracy.Caption = vbNullString
End If

' Save calibration
Call PictureSnapSaveCalibration(mode%, PictureSnapFilename$, PictureSnapCalibrationSaved)
If ierror Then Exit Sub

' If 3 point mode, then call a calibration just to print out the transformation matrix
If PictureSnapCalibrated And PictureSnapMode% = 1 Then
Call IOWriteLog(vbCrLf & "Picturesnap Calibration Fiducial Matrix Transformation (3 point):")
DebugMode = True
VerboseMode = True
Call PictureSnapConvertFiducialsCalculate(Int(2), formx!, formy!, formz!, RealTimeMotorPositions!(XMotor%), RealTimeMotorPositions!(YMotor%), RealTimeMotorPositions!(ZMotor%))
VerboseMode = False
DebugMode = False
If ierror Then Exit Sub
End If

FormPICTURESNAP.Picture2.Refresh
Exit Sub

' Errors
PictureSnapCalibrateError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapCalibrate"
ierror = True
Exit Sub

PictureSnapCalibrateNoPicture:
msg$ = "No image has been loaded in the PictureSnap window. Please open a sample image using the File | Open menu."
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapCalibrate"
ierror = True
Exit Sub

PictureSnapCalibrateOutofBoundsX:
msg$ = "The stage X position for one of the entered control coordinates is out of the stage limits (must be between " & Str$(MotLoLimits!(XMotor%)) & " and " & Str$(MotHiLimits!(XMotor%)) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapCalibrate"
ierror = True
Exit Sub

PictureSnapCalibrateOutofBoundsY:
msg$ = "The stage Y position for one of the entered control coordinates is out of the stage limits (must be between " & Str$(MotLoLimits!(YMotor%)) & " and " & Str$(MotHiLimits!(YMotor%)) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapCalibrate"
ierror = True
Exit Sub

PictureSnapCalibrateOutofBoundsZ:
msg$ = "The stage Z position for one of the entered control coordinates is out of the stage limits (must be between " & Str$(MotLoLimits!(ZMotor%)) & " and " & Str$(MotHiLimits!(ZMotor%)) & ") "
msg$ = msg$ & "You might need to move the stage back to the first and second calibration points and re-read them so the Z stage positions are loaded properly."
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapCalibrate"
ierror = True
Exit Sub

End Sub

Sub PictureSnapCalibratePoint(mode As Integer)
' Load the pixel coordinates for the selected point
'  mode = 1 load point #1 pixel coordinates
'  mode = 2 load point #2 pixel coordinates
'  mode = 3 load point #3 pixel coordinates     ' only if PictureSnapMode% = 1

ierror = False
On Error GoTo PictureSnapCalibratePointError

' Check for picture loaded
If Trim$(PictureSnapFilename$) = vbNullString Then GoTo PictureSnapCalibratePointNoPicture

' Point 1 pixel
If mode% = 1 Then
FormPICTURESELECT.Show vbModeless
DoEvents

WaitingForCalibrationClick = True
FormPICTURESNAP2.TextXPixel1.ForeColor = vbBlack
FormPICTURESNAP2.TextYPixel1.ForeColor = vbBlack
Do Until Not WaitingForCalibrationClick Or icancel
Call MiscDelay5(CDbl(0.2), Now) ' delay a little
If ierror Then
WaitingForCalibrationClick = False
FormPICTURESNAP.Picture2.MousePointer = vbDefault
FormPICTURESNAP3.MousePointer = vbDefault
Exit Sub
End If
Loop
FormPICTURESNAP2.TextXPixel1.Text = CurrentPointX!  ' save screen coordinates
FormPICTURESNAP2.TextYPixel1.Text = CurrentPointY!
FormPICTURESNAP2.TextXPixel1.ForeColor = vbRed
FormPICTURESNAP2.TextYPixel1.ForeColor = vbRed
End If

' Point 2 pixel
If mode% = 2 Then
FormPICTURESELECT.Show vbModeless
DoEvents

WaitingForCalibrationClick = True
FormPICTURESNAP2.TextXPixel2.ForeColor = vbBlack
FormPICTURESNAP2.TextYPixel2.ForeColor = vbBlack
Do Until Not WaitingForCalibrationClick Or icancel
Call MiscDelay5(CDbl(0.2), Now) ' delay a little
If ierror Then
WaitingForCalibrationClick = False
FormPICTURESNAP.Picture2.MousePointer = vbDefault
FormPICTURESNAP3.MousePointer = vbDefault
Exit Sub
End If
Loop
FormPICTURESNAP2.TextXPixel2.Text = CurrentPointX!   ' save screen coordinates
FormPICTURESNAP2.TextYPixel2.Text = CurrentPointY!
FormPICTURESNAP2.TextXPixel2.ForeColor = vbRed
FormPICTURESNAP2.TextYPixel2.ForeColor = vbRed
End If

' Point 3 pixel
If mode% = 3 Then
FormPICTURESELECT.Show vbModeless
DoEvents

WaitingForCalibrationClick = True
FormPICTURESNAP2.TextXPixel3.ForeColor = vbBlack
FormPICTURESNAP2.TextYPixel3.ForeColor = vbBlack
Do Until Not WaitingForCalibrationClick Or icancel
Call MiscDelay5(CDbl(0.2), Now) ' delay a little
If ierror Then
WaitingForCalibrationClick = False
FormPICTURESNAP.Picture2.MousePointer = vbDefault
FormPICTURESNAP3.MousePointer = vbDefault
Exit Sub
End If
Loop
FormPICTURESNAP2.TextXPixel3.Text = CurrentPointX!  ' save screen coordinates
FormPICTURESNAP2.TextYPixel3.Text = CurrentPointY!
FormPICTURESNAP2.TextXPixel3.ForeColor = vbRed
FormPICTURESNAP2.TextYPixel3.ForeColor = vbRed
End If

WaitingForCalibrationClick = False
FormPICTURESNAP.Picture2.MousePointer = vbDefault
FormPICTURESNAP3.MousePointer = vbDefault

Exit Sub

' Errors
PictureSnapCalibratePointError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapCalibratePoint"
ierror = True
Exit Sub

PictureSnapCalibratePointNoPicture:
msg$ = "No picture (*.BMP) has been loaded in the PictureSnap window. Please open a sample picture using the File | Open menu."
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapCalibratePoint"
ierror = True
Exit Sub

End Sub

Sub PictureSnapSelectUpdate(cpointx As Single, cpointy As Single)
' Update calibration variables

ierror = False
On Error GoTo PictureSnapSelectUpdateError

CurrentPointX! = cpointx!
CurrentPointY! = cpointy!
DoEvents

Unload FormPICTURESELECT

Exit Sub

' Errors
PictureSnapSelectUpdateError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapSelectUpdate"
ierror = True
Exit Sub

End Sub

Sub PictureSnapUpdateCursor(mode As Integer, xpos As Single, ypos As Single)
' Update the cursor display
' mode = 0  passed position is mouse coordinates
' mode = 1  passed position is stage coordinates

ierror = False
On Error GoTo PictureSnapUpdateCursorError

Dim stagex As Single, stagey As Single, stagez As Single
Dim fractionx As Single, fractiony As Single
Dim xpix As Long, ypix As Long

' If picture file is not loaded just exit
If PictureSnapFilename$ = vbNullString Then Exit Sub

' If not calibrated just calculate pixels and exit
If Not PictureSnapCalibrated Then
xpix& = xpos! / Screen.TwipsPerPixelX
ypix& = ypos! / Screen.TwipsPerPixelY
FormPICTURESNAP.Caption = "PictureSnap [" & PictureSnapFilename$ & "], Pixel X=" & Format$(xpix&) & ", Y=" & Format$(ypix&)
Exit Sub
End If

' Convert to stage coordinates
If mode% = 0 Then
Call PictureSnapConvert(Int(1), xpos!, ypos!, CSng(0#), stagex!, stagey!, stagez!, fractionx!, fractiony!)
If ierror Then Exit Sub

' Already in stage coordinates
Else
stagex! = xpos!
stagey! = ypos!
End If

' Update form
If Not RealTimeMode Then
FormPICTURESNAP.Caption = "PictureSnap [" & PictureSnapFilename$ & "], Stage X=" & MiscAutoFormat$(stagex!) & ", Y=" & MiscAutoFormat$(stagey!)
Else
If PictureSnapCalibrationNumberofZPoints% = 0 Then
FormPICTURESNAP.Caption = "PictureSnap [" & PictureSnapFilename$ & "] (double-click to move), Stage X=" & MiscAutoFormat$(stagex!) & ", Y=" & MiscAutoFormat$(stagey!)
Else
FormPICTURESNAP.Caption = "PictureSnap [" & PictureSnapFilename$ & "] (double-click to move), Stage X=" & MiscAutoFormat$(stagex!) & ", Y=" & MiscAutoFormat$(stagey!) & ", Z=" & MiscAutoFormat$(stagez!)
End If
End If

Exit Sub

' Errors
PictureSnapUpdateCursorError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapUpdateCursor"
ierror = True
Exit Sub

End Sub

Sub PictureSnapUpdateCursor2(mode As Integer, xpos As Single, ypos As Single)
' Update the cursor display for FormPICTURESNAP2
' mode = 0  passed position is mouse coordinates
' mode = 1  passed position is stage coordinates

ierror = False
On Error GoTo PictureSnapUpdateCursor2Error

Dim stagex As Single, stagey As Single, stagez As Single
Dim fractionx As Single, fractiony As Single
Dim xpix As Long, ypix As Long

' If picture file is not loaded just exit
If PictureSnapFilename$ = vbNullString Then Exit Sub

' If not calibrated just calculate pixels and exit
If Not PictureSnapCalibrated Then
xpix& = xpos! / Screen.TwipsPerPixelX
ypix& = ypos! / Screen.TwipsPerPixelY
FormPICTURESNAP3.Caption = "PictureSnap Full View (double-click to move), Pixel X=" & Format$(xpix&) & ", Y=" & Format$(ypix&)
Exit Sub
End If

' Convert FormPICTURESNAP3 form coordinates to FormPICTURESNAP.Picture2 coordinates
xpos! = FormPICTURESNAP.Picture2.ScaleWidth * xpos! / FormPICTURESNAP3.ScaleWidth
ypos! = FormPICTURESNAP.Picture2.ScaleHeight * ypos! / FormPICTURESNAP3.ScaleHeight

' Convert to stage coordinates
If mode% = 0 Then
Call PictureSnapConvert(Int(1), xpos!, ypos!, CSng(0#), stagex!, stagey!, stagez!, fractionx!, fractiony!)
If ierror Then Exit Sub

' Already in stage coordinates
Else
stagex! = xpos!
stagey! = ypos!
End If

' Update form
If PictureSnapMode% = 0 Then
If RealTimeMode Then
FormPICTURESNAP3.Caption = "PictureSnap Full View, Stage X=" & MiscAutoFormat$(stagex!) & ", Y=" & MiscAutoFormat$(stagey!) & " (double-click to move stage)"
Else
FormPICTURESNAP3.Caption = "PictureSnap Full View, Stage X=" & MiscAutoFormat$(stagex!) & ", Y=" & MiscAutoFormat$(stagey!)
End If

Else
If RealTimeMode Then
FormPICTURESNAP3.Caption = "PictureSnap Full View, Stage X=" & MiscAutoFormat$(stagex!) & ", Y=" & MiscAutoFormat$(stagey!) & ", Z=" & MiscAutoFormat$(stagez!) & " (double-click to move stage)"
Else
FormPICTURESNAP3.Caption = "PictureSnap Full View, Stage X=" & MiscAutoFormat$(stagex!) & ", Y=" & MiscAutoFormat$(stagey!) & ", Z=" & MiscAutoFormat$(stagez!)
End If
End If

Exit Sub

' Errors
PictureSnapUpdateCursor2Error:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapUpdateCursor2"
ierror = True
Exit Sub

End Sub

Sub PictureSnapLoadCalibration()
' Load the picture calibration from a text file of INI format (*.ACQ) (called when loading a BMP file in case it exists)

ierror = False
On Error GoTo PictureSnapLoadCalibrationError

Dim tfilename As String, tfilename2 As String

' Read calibration points to INI style ACQ file
tfilename$ = MiscGetFileNameNoExtension$(PictureSnapFilename$) & ".ACQ"
If Dir$(tfilename$) = vbNullString Then

' Check if a PrbImg file already exists and if so convert it to an ACQ
tfilename2$ = MiscGetFileNameNoExtension$(PictureSnapFilename$) & ".PrbImg"
If Dir$(tfilename2$) <> vbNullString Then
Call PictureSnapConvertPrbImgToACQ(tfilename2$)
If ierror Then Exit Sub
End If

' If ACQ file still does not exist, then just exit
If Dir$(tfilename$) = vbNullString Then Exit Sub
End If

' Read the .ACQ file calibration
Call PictureSnapReadCalibration(tfilename$)
If ierror Then Exit Sub

' Reload caption
If Not RealTimeMode Then
FormPICTURESNAP.Caption = "PictureSnap [" & PictureSnapFilename$ & "]"
Else
FormPICTURESNAP.Caption = "PictureSnap [" & PictureSnapFilename$ & "] (double-click to move)"
End If

PictureSnapCalibrated = True
PictureSnapCalibrationSaved = True  ' since it was read, the calibration is already saved

' Load calibration dialog if open
Call PictureSnapCalibrateLoad(Int(0))
If ierror Then Exit Sub

Exit Sub

' Errors
PictureSnapLoadCalibrationError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapLoadCalibration"
ierror = True
Exit Sub

End Sub

Sub PictureSnapDrawCurrentPosition()
' Draw current position on pic2 if real time mode

ierror = False
On Error GoTo PictureSnapDrawCurrentPositionError

Dim formx As Single, formy As Single, formz As Single
Dim radius As Single, tWidth As Single
Dim a1 As Single, a2 As Single
Dim fractionx As Single, fractiony As Single

Static oldx As Single, oldy As Single

' Skip if not in real time mode (or not CalcImage)
If Not MiscStringsAreSame(app.EXEName, "CalcImage") And Not RealTimeMode Then Exit Sub

' Skip if interface is busy
If RealTimeInterfaceBusy Then Exit Sub

' Skip if pausing automation
If RealTimePauseAutomation Then Exit Sub

' If form not visible just exit
If Not FormPICTURESNAP.Visible Then Exit Sub

' If no picture just exit
If PictureSnapFilename$ = vbNullString Then Exit Sub

' If not calibrated, just exit
If Not PictureSnapCalibrated Then Exit Sub

' If not realtime and no coordinates, just exit
If Not RealTimeMode And RealTimeMotorPositions!(XMotor%) = 0# And RealTimeMotorPositions!(YMotor%) = 0# Then Exit Sub

' Update caption position (causes mouse move cursor coordinates to be overwritten)
'Call PictureSnapUpdateCursor(Int(1), RealTimeMotorPositions!(XMotor%), RealTimeMotorPositions!(YMotor%))
'If ierror Then Exit Sub

' Convert to form coordinates
Call PictureSnapConvert(Int(2), formx!, formy!, formz!, RealTimeMotorPositions!(XMotor%), RealTimeMotorPositions!(YMotor%), RealTimeMotorPositions!(ZMotor%), fractionx!, fractiony!)
If ierror Then Exit Sub

' Calculate a radius
tWidth! = Screen.Width
If tWidth! = 0# Then Exit Sub
radius! = tWidth! / 100#

' Erase the old circle
If CLng(oldx!) <> CLng(formx!) Or CLng(oldy!) <> CLng(formy!) Then
FormPICTURESNAP.Picture2.Refresh
End If

' Draw current position of normal PictureSnap image
FormPICTURESNAP.Picture2.DrawWidth = 2
FormPICTURESNAP.Picture2.Circle (formx!, formy!), radius!, RGB(150, 0, 150)
a1! = formx! + radius! * 2
a2! = formx! - radius! * 2
FormPICTURESNAP.Picture2.Line (a1!, formy!)-(formx! + radius! / 2, formy!), RGB(150, 0, 150)
FormPICTURESNAP.Picture2.Line (a2!, formy!)-(formx! - radius! / 2, formy!), RGB(150, 0, 150)
a1! = formy! + radius! * 2
a2! = formy! - radius! * 2
FormPICTURESNAP.Picture2.Line (formx!, a1!)-(formx!, formy! + radius! / 2), RGB(150, 0, 150)
FormPICTURESNAP.Picture2.Line (formx!, a2!)-(formx!, formy! - radius! / 2), RGB(150, 0, 150)

' Update full window
If FormPICTURESNAP3.Visible Then
tWidth! = FormPICTURESNAP3.ScaleWidth   ' calculate a radius
If tWidth! <> 0# Then
radius! = (tWidth! / 50#) ^ 0.8

' Erase the old circle
If CLng(oldx!) <> CLng(formx!) Or CLng(oldy!) <> CLng(formy!) Then
FormPICTURESNAP3.Image1.Refresh
End If

' Draw current position on full view window
FormPICTURESNAP3.DrawWidth = 2
FormPICTURESNAP3.Circle (FormPICTURESNAP3.ScaleWidth * fractionx!, FormPICTURESNAP3.ScaleHeight * fractiony!), radius!, RGB(150, 0, 150)
End If
End If

' Display current mag box
Call PictureSnapDisplayCurrentMagBox
If ierror Then Exit Sub

' Save this position
oldx! = formx!
oldy! = formy!
Exit Sub

' Errors
PictureSnapDrawCurrentPositionError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapDrawCurrentPosition"
ierror = True
Exit Sub

End Sub

Sub PictureSnapLoadFullWindow()
' Open the full window view

ierror = False
On Error GoTo PictureSnapLoadFullWindowError

' If picture file is already specified load it
If PictureSnapFilename$ = vbNullString Then GoTo PictureSnapLoadFullWindowNoPicture

' Load into picturebox control to perform flipping
Screen.MousePointer = vbHourglass
Set FormPICTURESNAP3.Image1.Picture = LoadPicture(PictureSnapFilename$)

' Rescale form to image aspect
If FormPICTURESNAP3.Image1.Picture.Type > 0 Then   ' bitmap
If FormPICTURESNAP3.Image1.Picture.Height <> 0# Then
FormPICTURESNAP3.Width = FormPICTURESNAP3.ScaleHeight * FormPICTURESNAP3.Image1.Picture.Width / FormPICTURESNAP3.Image1.Picture.Height
End If
End If

Screen.MousePointer = vbDefault
FormPICTURESNAP3.Show vbModeless

Exit Sub

' Errors
PictureSnapLoadFullWindowError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapLoadFullWindow"
ierror = True
Exit Sub

PictureSnapLoadFullWindowNoPicture:
msg$ = "No image file has been opened yet. Use the File | Open menu in the PictureSnap window to open a scanned image of your sample."
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapLoadFullWindow"
ierror = True
Exit Sub

End Sub

Sub PictureSnapDrawScaleBar()
' Draw a scale bar for the PictureSnap window

ierror = False
On Error GoTo PictureSnapDrawScaleBarError

Dim tcolor As Long
Dim xrange As Single, xrange2 As Single
Dim tuleftx As Single, tulefty As Single
Dim tlrightx As Single, tlrighty As Single
Dim astring As String

Dim sx1 As Single, sy1 As Single, sz1 As Single
Dim sx2 As Single, sy2 As Single, sz2 As Single

Dim fractionx As Single, fractiony As Single
Dim x1 As Single, y1 As Single
Dim xmin As Single, ymin As Single, zmin As Single
Dim xmax As Single, ymax As Single, zmax As Single

Dim halfwidth As Single, halfheight As Single
Dim tcurrentx As Single, tcurrenty As Single

Dim tStageConversion As Single

' If form not visible just exit
If Not FormPICTURESNAP.Visible Then Exit Sub

' If no picture just exit
If PictureSnapFilename$ = vbNullString Then Exit Sub

' if not calibrated, just exit
If Not PictureSnapCalibrated Then Exit Sub

' Determine image extents
If FormPICTURESNAP.Picture2.Picture.Type <> 1 Then Exit Sub     ' not bitmap

' Erase the old scale bar
If CLng(oldx!) <> CLng(FormPICTURESNAP.ScaleWidth) Or CLng(oldy!) <> CLng(FormPICTURESNAP.ScaleHeight) Then
FormPICTURESNAP.Picture2.Refresh
oldx! = FormPICTURESNAP.ScaleWidth
oldy! = FormPICTURESNAP.ScaleHeight
End If

' Calculate 1/8 of visible image
xrange! = FormPICTURESNAP.Picture1.Width / 8#

' Convert to stage coordinates (z coordinates are not used)
Call PictureSnapConvert(Int(1), CSng(0#), ymin!, zmin!, sx1!, sy1!, sz1!, fractionx!, fractiony!)
If ierror Then Exit Sub
Call PictureSnapConvert(Int(1), xrange!, ymax!, zmax!, sx2!, sy2!, sz2!, fractionx!, fractiony!)
If ierror Then Exit Sub

' Update micron scale bar conversion
If Default_Stage_Units$ = "um" Then tStageConversion! = 1#
If Default_Stage_Units$ = "mm" Then tStageConversion! = 1000#

' Calculate stage distance in microns
xrange2! = Abs(sx2! - sx1!) * tStageConversion!

' Round to nice number
xrange2! = MiscAutoFormatZ!(100#, xrange2!)
If ierror Then Exit Sub

' Convert back to stage units
xrange2! = xrange2! / tStageConversion!

' Recalculate scale bar in screen units
sx1! = RealTimeMotorPositions!(XMotor%)
sy1! = RealTimeMotorPositions!(YMotor%)
sx2! = RealTimeMotorPositions!(XMotor%) + xrange2!
sy2! = RealTimeMotorPositions!(YMotor%)

Call PictureSnapConvert(Int(2), xmin!, ymin!, zmin!, sx1!, sy1!, sz1!, fractionx!, fractiony!)
If ierror Then Exit Sub
Call PictureSnapConvert(Int(2), xmax!, ymax!, zmax!, sx2!, sy2!, sz2!, fractionx!, fractiony!)
If ierror Then Exit Sub

' Calculate scale bar in screen units
xrange! = Abs(xmax! - xmin!)

' Convert stage units back to microns
xrange2! = xrange2! * tStageConversion!

' Fix position of scale bar to lower left
tuleftx! = xrange! * 0.6
tulefty! = FormPICTURESNAP.Picture1.ScaleHeight * 0.9 - 500     ' (the 500 offset is for the actual bar and text height)

' Add scroll offset to keep scale bar in view
If FormPICTURESNAP.HScroll1.Max > 0 Then
x1! = FormPICTURESNAP.HScroll1.value / FormPICTURESNAP.HScroll1.Max
End If
If FormPICTURESNAP.VScroll1.Max > 0 Then
y1! = FormPICTURESNAP.VScroll1.value / FormPICTURESNAP.VScroll1.Max
End If

tuleftx! = tuleftx! + x1! * (FormPICTURESNAP.Picture2.ScaleWidth - FormPICTURESNAP.Picture1.ScaleWidth)
tulefty! = tulefty! + y1! * (FormPICTURESNAP.Picture2.ScaleHeight - FormPICTURESNAP.Picture1.ScaleHeight)

' Make scale bar rectangle
tlrightx! = tuleftx! + xrange!
tlrighty! = tulefty! + 100

' Draw on form
If Not FormPICTURESNAP.menuDisplayUseBlackScaleBar.Checked Then
tcolor& = RGB(255, 255, 255)
Else
tcolor& = RGB(0, 0, 0)
End If
FormPICTURESNAP.Picture2.DrawWidth = 1      ' use DrawWidth = 1 for filled box accuracy
FormPICTURESNAP.Picture2.Line (tuleftx!, tulefty!)-(tlrightx!, tlrighty!), tcolor&, BF
   
' Print text of microns
FormPICTURESNAP.Picture2.CurrentX = FormPICTURESNAP.Picture2.CurrentX - xrange! / 2#
tcurrentx = FormPICTURESNAP.Picture2.CurrentX
tcurrenty = FormPICTURESNAP.Picture2.CurrentY
astring$ = Format$(xrange2!) & " um"
FormPICTURESNAP.Picture2.ForeColor = tcolor& ' set foreground color
FormPICTURESNAP.Picture2.FontSize = 13       ' set font size
FormPICTURESNAP.Picture2.FontName = LogWindowFontName$
FormPICTURESNAP.Picture2.FontSize = 13       ' set font size    (necessary for Windows)
FormPICTURESNAP.Picture2.FontBold = False
halfwidth! = FormPICTURESNAP.Picture2.TextWidth(astring$) / 2      ' calculate one-half width
'halfheight! = FormPICTURESNAP.Picture2.TextHeight(astring$) / 2     ' calculate one-half height
FormPICTURESNAP.Picture2.CurrentX = FormPICTURESNAP.Picture2.CurrentX - halfwidth!   ' set X
'FormPICTURESNAP.Picture2.CurrentY = FormPICTURESNAP.Picture2.CurrentY + halfheight! ' set Y
FormPICTURESNAP.Picture2.Print astring$   ' print text string to form

Exit Sub

' Errors
PictureSnapDrawScaleBarError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapDrawScaleBar"
ierror = True
Exit Sub

End Sub

Sub PictureSnapResetScaleBar()
' Reset the ScaleBar

ierror = False
On Error GoTo PictureSnapResetScaleBarError

oldx! = 0#
oldy! = 0#

Exit Sub

' Errors
PictureSnapResetScaleBarError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapResetScaleBar"
ierror = True
Exit Sub

End Sub

Sub PictureSnapDisplayCurrentMagBox()
' Draw the current magnification scan box

ierror = False
On Error GoTo PictureSnapDisplayCurrentMagBoxError

Dim tWidth As Single

Dim formx1 As Single, formy1 As Single, formz1 As Single
Dim formx2 As Single, formy2 As Single, formz2 As Single
Dim xdata1 As Single, ydata1 As Single, zdata1 As Single
Dim xdata2 As Single, ydata2 As Single, zdata2 As Single
Dim fraction1x As Single, fraction1y As Single
Dim fraction2x As Single, fraction2y As Single
Dim xoffset As Single, yoffset As Single

Dim tmagnification As Single
Dim tbeammode As Integer

Static oldx As Single, oldy As Single

' Check for pathological conditions
If NumberOfStageMotors% < 1 Then Exit Sub

' Get beam mode
If Not UseSharedMonitorDataFlag% Then
Call RealTimeGetBeamMode(tbeammode%)
If ierror Then Exit Sub
Else
tbeammode% = MonitorStateBeamMode%  ' 0 = spot, 1  = scan, 2 = digital
End If

' Read magnification (only read magnification if scan mode)
If tbeammode% = 1 Then
If Not UseSharedMonitorDataFlag% Then
Call RealTimeGetMagnification(tmagnification!)
If ierror Then Exit Sub
Else
tmagnification! = MonitorStateMagnification!
End If
End If

' Calculate mag box corners in +/- microns
If tmagnification! <> 0# Then
xoffset! = (RealTimeGetBeamScanCalibration!(XMotor%, DefaultKiloVolts!, tmagnification!, DefaultScanRotation!)) / 2#
'yoffset! = (RealTimeGetBeamScanCalibration!(YMotor%, DefaultKiloVolts!, tmagnification!, DefaultScanRotation!)) / 2#
yoffset! = xoffset! / ImageInterfaceImageIxIy!          ' use this instead for aspect ratio?

' Calculate absolute stage positions of mag box
xdata1! = RealTimeMotorPositions!(XMotor%) - xoffset! / MotUnitsToAngstromMicrons!(XMotor%)
ydata1! = RealTimeMotorPositions!(YMotor%) + yoffset! / MotUnitsToAngstromMicrons!(YMotor%)

xdata2! = RealTimeMotorPositions!(XMotor%) + xoffset! / MotUnitsToAngstromMicrons!(XMotor%)
ydata2! = RealTimeMotorPositions!(YMotor%) - yoffset! / MotUnitsToAngstromMicrons!(YMotor%)

' Convert to form coordinates
Call PictureSnapConvert(Int(2), formx1!, formy1!, formz1!, xdata1!, ydata1!, zdata1!, fraction1x!, fraction1y!)
If ierror Then Exit Sub

Call PictureSnapConvert(Int(2), formx2!, formy2!, formz2!, xdata2!, ydata2!, zdata2!, fraction2x!, fraction2y!)
If ierror Then Exit Sub
End If

If CLng(oldx!) <> CLng(formx1!) Or CLng(oldy!) <> CLng(formy1!) Then
FormPICTURESNAP.Picture2.Refresh
End If

' Update mag box if scan mode
If tbeammode% = 1 Then
FormPICTURESNAP.Picture2.DrawWidth = 2
FormPICTURESNAP.Picture2.Line (formx1!, formy1!)-(formx2!, formy1!), RGB(0, 0, 150)
FormPICTURESNAP.Picture2.Line (formx1!, formy2!)-(formx2!, formy2!), RGB(0, 0, 150)

FormPICTURESNAP.Picture2.Line (formx1!, formy1!)-(formx1!, formy2!), RGB(0, 0, 150)
FormPICTURESNAP.Picture2.Line (formx2!, formy1!)-(formx2!, formy2!), RGB(0, 0, 150)
End If

' Save this position
oldx! = formx1!
oldy! = formy1!
Exit Sub

' Errors
PictureSnapDisplayCurrentMagBoxError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapDisplayCurrentMagBox"
ierror = True
Exit Sub

End Sub

Sub PictureSnapCalibrateUnLoad()
' Check if image is calibrated before unloading calibration form

ierror = False
On Error GoTo PictureSnapCalibrateUnLoadError

If PictureSnapCalibratedPreviously And Not PictureSnapCalibrated Then
MsgBox "The current image is no longer calibrated to the stage coordinates. If the image was previously calibrated, simply re-load the image from the file on disk.", vbOKOnly + vbInformation, "PictureSnapCalibrateUnLoad"
End If

Exit Sub

' Errors
PictureSnapCalibrateUnLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapCalibrateUnLoad"
ierror = True
Exit Sub

End Sub

Sub PictureSnapEnableDisable()
' Enable/disable menus in the PictureSnap form

ierror = False
On Error GoTo PictureSnapEnableDisableError

' Enable output and modeless window menus if modeless
If PictureSnapWindowIsModeless And PictureSnapFilename$ <> vbNullString Then
FormPICTURESNAP.menuFileSaveAsGRD.Enabled = True
FormPICTURESNAP.menuFileClipboard1.Enabled = True
FormPICTURESNAP.menuFileClipboard2.Enabled = True
FormPICTURESNAP.menuFileSaveAsBMPOnly.Enabled = True
FormPICTURESNAP.menuFileSaveAsBMP.Enabled = True
FormPICTURESNAP.menuFilePrintSetup.Enabled = True
FormPICTURESNAP.menuFilePrint.Enabled = True

FormPICTURESNAP.menuWindowCalibrate.Enabled = True
FormPICTURESNAP.menuWindowFullPicture.Enabled = True

' Disable output menus and modeless window menus if not modeless (if modal)
Else
FormPICTURESNAP.menuFileSaveAsGRD.Enabled = False
FormPICTURESNAP.menuFileClipboard1.Enabled = False
FormPICTURESNAP.menuFileClipboard2.Enabled = False
FormPICTURESNAP.menuFileSaveAsBMPOnly.Enabled = False
FormPICTURESNAP.menuFileSaveAsBMP.Enabled = False
FormPICTURESNAP.menuFilePrintSetup.Enabled = False
FormPICTURESNAP.menuFilePrint.Enabled = False

FormPICTURESNAP.menuWindowCalibrate.Enabled = False
FormPICTURESNAP.menuWindowFullPicture.Enabled = False
End If

Exit Sub

' Errors
PictureSnapEnableDisableError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapEnableDisable"
ierror = True
Exit Sub

End Sub

Sub PictureSnapCalibrateCheck()
' Check the new stage calibration and check that X and Y are orthogonal

ierror = False
On Error GoTo PictureSnapCalibrateCheckError

Const tolerance! = 0.1        ' 10 %

Dim tmsg As String
Dim tdist As Single, xdist As Single, ydist As Single

Dim sx1 As Single, sy1 As Single, sz1 As Single
Dim sx2 As Single, sy2 As Single, sz2 As Single

Dim fractionx As Single, fractiony As Single
Dim xmin As Single, ymin As Single, zmin As Single
Dim xmax As Single, ymax As Single, zmax As Single

' Take an arbitrary screen distance and check that X and Y are equal within a tolerance
tdist! = 1000       ' try 1000 twips

' Convert screen to stage coordinates
Call PictureSnapConvert(Int(1), CSng(0#), CSng(0#), zmin!, sx1!, sy1!, sz1!, fractionx!, fractiony!)
If ierror Then Exit Sub

Call PictureSnapConvert(Int(1), tdist!, tdist!, zmax!, sx2!, sy2!, sz2!, fractionx!, fractiony!)
If ierror Then Exit Sub

' Calculate x and y distances for given screen distance
xdist! = Abs(sx2! - sx1!)
ydist! = Abs(sy2! - sy1!)

' Update calibration window for accuracy
tmsg$ = "X=" & Format$(xdist!) & vbCrLf & "Y=" & Format$(ydist!) & vbCrLf & "(X-Y)/X=" & MiscAutoFormat4$(Abs((xdist! - ydist!) / xdist!) * 100#) & "%"
FormPICTURESNAP2.LabelCalibrationAccuracy.Caption = tmsg$

' Warn user if not equal in X and Y within tolerance
If Not MiscDifferenceIsSmall(xdist!, ydist!, tolerance!) Then
msg$ = "The stage calibration will be saved, but the nominal X (" & Format$(xdist!) & ") and nominal Y (" & Format$(ydist!) & ") calibration distances are different by more than " & Format$(CInt(tolerance! * 100#)) & "%." & vbCrLf & vbCrLf
msg$ = msg$ & "Please check your stage calibration and image pixel positions and make sure that they are correctly located and specified! "
msg$ = msg$ & "Please note that this accuracy error can also occur if the sample itself is significantly rotated with respect to the loaded image."
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapCalibrateCheck"
End If

Exit Sub

' Errors
PictureSnapCalibrateCheckError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapCalibrateCheck"
ierror = True
Exit Sub

End Sub

