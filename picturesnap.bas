Attribute VB_Name = "CodePictureSnap"
' (c) Copyright 1995-2022 by John J. Donovan
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

Private Type VertexType
    X As Single
    Y As Single
End Type

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
FormPICTURESNAP2.TextZStage1.Text = RealTimeMotorPositions!(ZMotor%)
FormPICTURESNAP2.TextZStage2.Text = RealTimeMotorPositions!(ZMotor%)
FormPICTURESNAP2.TextZStage3.Text = RealTimeMotorPositions!(ZMotor%)

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

' Calculate the image rotation relative to the stage
Call PictureSnapCalculateRotation
If ierror Then Exit Sub

' Check stage calibration is orthogonal
Call PictureSnapCalibrateCheck
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
' Update the cursor display for FormPICTURESNAP3
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
' Draw current position on Picturesnap.frm and PictureSnap3.frm.

ierror = False
On Error GoTo PictureSnapDrawCurrentPositionError

Dim formx As Single, formy As Single, formz As Single
Dim radius As Single, tWidth As Single
Dim a1 As Single, a2 As Single
Dim fractionx As Single, fractiony As Single

Dim i As Integer
Dim lineColor As Long
Dim lineWidth As Single
Dim cX As Single, cY As Single, cRadius As Single
Dim twipsToPixelX As Single, twipsToPixelY As Single
    
Dim lineVertices() As VertexType
    
ReDim lineVertices(0 To 7) As VertexType

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
    twipsToPixelX = Screen.TwipsPerPixelX               ' all twip measurements must be converted to pixels; calculate the conversion now
    If (twipsToPixelX = 0!) Then twipsToPixelX = 15!
    twipsToPixelY = Screen.TwipsPerPixelY               ' all twip measurements must be converted to pixels; calculate the conversion now
    If (twipsToPixelY = 0!) Then twipsToPixelY = 15!
    
    lineColor = RGB(150, 0, 150)
    lineWidth = 2
    
    ' Convert the circle coords into pixels
    cX = formx! / twipsToPixelX
    cY = formy! / twipsToPixelY
    cRadius = radius! / twipsToPixelX
    
    GDIPlus_Interface.GDIPlus_DrawCircle FormPICTURESNAP.Picture2.hDC, cX, cY, cRadius, lineColor, lineWidth
    
    ' Because we have to convert all line coordinates into twips, it's easier to stuff all points into an array
    a1! = formx! + radius! * 2
    a2! = formx! - radius! * 2
    
    lineVertices(0).X = a1!
    lineVertices(0).Y = formy!
    lineVertices(1).X = formx! + radius! / 2
    lineVertices(1).Y = formy!
    lineVertices(2).X = a2!
    lineVertices(2).Y = formy!
    lineVertices(3).X = formx! - radius! / 2
    lineVertices(3).Y = formy!
    
    a1! = formy! + radius! * 2
    a2! = formy! - radius! * 2
    
    lineVertices(4).X = formx!
    lineVertices(4).Y = a1!
    lineVertices(5).X = formx!
    lineVertices(5).Y = formy! + radius! / 2
    lineVertices(6).X = formx!
    lineVertices(6).Y = a2!
    lineVertices(7).X = formx!
    lineVertices(7).Y = formy! - radius! / 2
    
    For i% = 0 To 7
        lineVertices(i%).X = lineVertices(i).X / twipsToPixelX
        lineVertices(i%).Y = lineVertices(i).Y / twipsToPixelY
    Next i%
    
    ' Render all lines in turn
    For i% = 0 To 3
        GDIPlus_Interface.GDIPlus_DrawLine FormPICTURESNAP.Picture2.hDC, lineVertices(i% * 2).X, lineVertices(i% * 2).Y, lineVertices(i% * 2 + 1).X, lineVertices(i% * 2 + 1).Y, lineColor, lineWidth
    Next i%

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

' Minimize and restore to re-size
FormPICTURESNAP3.WindowState = vbMinimized
DoEvents
FormPICTURESNAP3.WindowState = vbNormal

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
Dim X1 As Single, Y1 As Single
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
X1! = FormPICTURESNAP.HScroll1.value / FormPICTURESNAP.HScroll1.Max
End If
If FormPICTURESNAP.VScroll1.Max > 0 Then
Y1! = FormPICTURESNAP.VScroll1.value / FormPICTURESNAP.VScroll1.Max
End If

tuleftx! = tuleftx! + X1! * (FormPICTURESNAP.Picture2.ScaleWidth - FormPICTURESNAP.Picture1.ScaleWidth)
tulefty! = tulefty! + Y1! * (FormPICTURESNAP.Picture2.ScaleHeight - FormPICTURESNAP.Picture1.ScaleHeight)

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
Dim tdistx As Single, tdisty As Single
Dim xdist As Single, ydist As Single

Dim sx1 As Single, sy1 As Single, sz1 As Single
Dim sx2 As Single, sy2 As Single, sz2 As Single

Dim fractionx As Single, fractiony As Single
Dim xmin As Single, ymin As Single, zmin As Single
Dim xmax As Single, ymax As Single, zmax As Single

Dim vCorner(1 To 1) As VertexType
Dim vOrigin As VertexType

Dim rotatedx As Single, rotatedy As Single

' Take an arbitrary screen distance and check that X and Y are equal within a tolerance
tdistx! = 1000       ' try 1000 twips
tdisty! = 1000       ' try 1000 twips

' Use rotate vertex code to calculate stage orthogonal distances
vOrigin.X! = 0#
vOrigin.Y! = 0#

vCorner(1).X! = tdistx!
vCorner(1).Y! = tdisty!

' Rotate the screen distance by the current rotation angle before converting to stage units
vCorner(1) = PictureSnapRotateVertex(vCorner(1), vOrigin, PictureSnapRotation!)

' Load the rotated corners
rotatedx! = vCorner(1).X
rotatedy! = vCorner(1).Y

' Convert screen to stage coordinates
Call PictureSnapConvert(Int(1), CSng(0#), CSng(0#), zmin!, sx1!, sy1!, sz1!, fractionx!, fractiony!)
If ierror Then Exit Sub

Call PictureSnapConvert(Int(1), rotatedx!, rotatedy!, zmax!, sx2!, sy2!, sz2!, fractionx!, fractiony!)
If ierror Then Exit Sub

' Calculate x and y distances for given screen distance
xdist! = Abs(sx2! - sx1!)
ydist! = Abs(sy2! - sy1!)

' Update calibration window for accuracy
tmsg$ = "X=" & MiscAutoFormat6$(xdist!) & ", Y=" & MiscAutoFormat6$(ydist!) & vbCrLf & "(X-Y)/X=" & MiscAutoFormat4$(Abs((xdist! - ydist!) / xdist!) * 100#) & "%" & vbCrLf
tmsg$ = tmsg$ & "Rotation=" & Format$(PictureSnapRotation!, "0.00") & " degrees"
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

Sub PictureSnapDisplayCurrentMagBox()
' Draw the current magnification scan box

ierror = False
On Error GoTo PictureSnapDisplayCurrentMagBoxError

Dim tcolor As Long
Dim tWidth As Integer

Dim formx As Single, formy As Single, formz As Single
Dim fractionx As Single, fractiony As Single

Dim xmicrons As Single, ymicrons As Single           ' x and y FOV in microns
Dim xdistance As Single, ydistance As Single         ' x and y FOV in stage units
Dim xwidth As Single, ywidth As Single               ' x and y FOV in form units

Dim tmagnification As Single
Dim tbeammode As Integer

Static oldx As Single, oldy As Single

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
xmicrons! = RealTimeGetBeamScanCalibration!(XMotor%, DefaultKiloVolts!, tmagnification!, DefaultScanRotation!)
If ierror Then Exit Sub
ymicrons! = xmicrons! / ImageInterfaceImageIxIy!

' Convert micron FOV to stage units
xdistance! = xmicrons! / MotUnitsToAngstromMicrons!(XMotor%)
ydistance! = xdistance! / ImageInterfaceImageIxIy!

' Convert FOV distance from stage units to form units (using image rotation)
xwidth! = PictureSnapConvertStageDistancetoImageDistance(Int(0), xdistance!, PictureSnapRotation!)
If ierror Then Exit Sub
ywidth! = PictureSnapConvertStageDistancetoImageDistance(Int(1), ydistance!, PictureSnapRotation!)
If ierror Then Exit Sub
ywidth! = xwidth! / ImageInterfaceImageIxIy!                              ' use actual aspect ratio for best accuracy

' Convert current stage position to form coordinates
Call PictureSnapConvert(Int(2), formx!, formy!, formz!, RealTimeMotorPositions!(XMotor%), RealTimeMotorPositions!(YMotor%), RealTimeMotorPositions!(ZMotor%), fractionx!, fractiony!)
If ierror Then Exit Sub

' Refresh image if current position changed from last time
If CLng(oldx!) <> CLng(formx!) Or CLng(oldy!) <> CLng(formy!) Then
FormPICTURESNAP.Picture2.Refresh
End If

' Update mag box if scan mode
If tbeammode% = 1 Then
tcolor& = RGB(0, 0, 150)
tWidth% = 2

' New code to draw magbox corners using rectangle rotation
Call PictureSnapDrawRectangle(formx!, formy!, xwidth!, ywidth!, PictureSnapRotation!, tcolor&, tWidth%)
If ierror Then Exit Sub
End If

' Save this position
oldx! = formx!
oldy! = formy!
End If

Exit Sub

' Errors
PictureSnapDisplayCurrentMagBoxError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapDisplayCurrentMagBox"
ierror = True
Exit Sub

End Sub

Function PictureSnapRotateVertex(vCorner As VertexType, vOrigin As VertexType, AngleDegrees As Single) As VertexType
' Calculate the rotated corners of a rectangle

ierror = False
On Error GoTo PictureSnapRotateVertexError

Dim arad As Single
    
arad! = AngleDegrees! * PI! / 180
    
PictureSnapRotateVertex.X = ((vCorner.X - vOrigin.X) * Cos(arad!) - (vCorner.Y - vOrigin.Y) * Sin(arad)) + vOrigin.X
PictureSnapRotateVertex.Y = ((vCorner.Y - vOrigin.Y) * Cos(arad!) + (vCorner.X - vOrigin.X) * Sin(arad)) + vOrigin.Y

Exit Function

' Errors
PictureSnapRotateVertexError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapRotateVertex"
ierror = True
Exit Function

End Function

Function PictureSnapConvertStageDistancetoImageDistance(mode As Integer, sdistance As Single, rotation As Single) As Single
' Convert the passed stage distance to form (image) units
'  mode = 0 calculate x distance
'  mode = 1 calculate y distance

ierror = False
On Error GoTo PictureSnapConvertStageDistancetoImageDistanceError

Dim temp As Single
Dim arad As Single

Dim formx1 As Single, formy1 As Single, formz1 As Single
Dim formx2 As Single, formy2 As Single, formz2 As Single

Dim fractionx1 As Single, fractiony1 As Single
Dim fractionx2 As Single, fractiony2 As Single

' Utilize image to stage rotation value for rigorous form distance calculation
arad! = rotation! * PI! / 180
temp! = sdistance! / Cos(arad!)

' Convert the stage distance to form units
If mode% = 0 Then
Call PictureSnapConvert(Int(2), formx1!, formy1!, formz1!, CSng(0#), CSng(0#), CSng(0#), fractionx1!, fractiony1!)
If ierror Then Exit Function

Call PictureSnapConvert(Int(2), formx2!, formy2!, formz2!, temp!, CSng(0#), CSng(0#), fractionx2!, fractiony2!)
If ierror Then Exit Function

temp! = Abs(formx2! - formx1!)

Else
Call PictureSnapConvert(Int(2), formx1!, formy1!, formz1!, CSng(0#), CSng(0#), CSng(0#), fractionx1!, fractiony1!)
If ierror Then Exit Function

Call PictureSnapConvert(Int(2), formx2!, formy2!, formz2!, CSng(0#), temp!, CSng(0#), fractionx2!, fractiony2!)
If ierror Then Exit Function

temp! = Abs(formy2! - formy1!)
End If

PictureSnapConvertStageDistancetoImageDistance! = temp!

Exit Function

' Errors
PictureSnapConvertStageDistancetoImageDistanceError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapConvertStageDistancetoImageDistance"
ierror = True
Exit Function

End Function

Function PictureSnapConvertImageDistancetoStageDistance(mode As Integer, fdistance As Single, rotation As Single) As Single
' Convert the passed form (image) distance to stage units
'  mode = 0 calculate x distance
'  mode = 1 calculate y distance

ierror = False
On Error GoTo PictureSnapConvertImageDistancetoStageDistanceError

Dim temp As Single
Dim arad As Single

Dim stagex1 As Single, stagey1 As Single, stagez1 As Single
Dim stagex2 As Single, stagey2 As Single, stagez2 As Single

Dim fractionx1 As Single, fractiony1 As Single
Dim fractionx2 As Single, fractiony2 As Single

' Utilize image to stage rotation value for rigorous stage distance calculation
arad! = rotation! * PI! / 180
temp! = fdistance! / Cos(arad!)

' Calculate the form distance to stage units
If mode% = 0 Then
Call PictureSnapConvert(Int(1), CSng(0#), CSng(0#), CSng(0#), stagex1!, stagey1!, stagez1!, fractionx1!, fractiony1!)
If ierror Then Exit Function

Call PictureSnapConvert(Int(1), temp!, CSng(0#), CSng(0#), stagex2!, stagey2!, stagez2!, fractionx2!, fractiony2!)
If ierror Then Exit Function

temp! = Abs(stagex2! - stagex1!)

Else
Call PictureSnapConvert(Int(1), CSng(0#), CSng(0#), CSng(0#), stagex1!, stagey1!, stagez1!, fractionx1!, fractiony1!)
If ierror Then Exit Function

Call PictureSnapConvert(Int(1), CSng(0#), temp!, CSng(0#), stagex2!, stagey2!, stagez2!, fractionx2!, fractiony2!)
If ierror Then Exit Function

temp! = Abs(stagey2! - stagey1!)
End If

PictureSnapConvertImageDistancetoStageDistance! = temp!

Exit Function

' Errors
PictureSnapConvertImageDistancetoStageDistanceError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapConvertImageDistancetoStageDistance"
ierror = True
Exit Function

End Function

Sub PictureSnapDrawRectangle(xcenter As Single, ycenter As Single, xwidth As Single, ywidth As Single, rotation As Single, tcolor As Long, tWidth As Integer)
' Draws a rectangle of the specified width and height at the specified screen location (all values in screen units)
'  xcenter = rectangle center x form position
'  ycenter = rectangle center y form position
'  xwidth = rectangle width in form units
'  ywidth = rectangle height in form units
'  tcolor = line color
'  twidth = line width

ierror = False
On Error GoTo PictureSnapDrawRectangleError

Dim formx1 As Single, formy1 As Single
Dim formx2 As Single, formy2 As Single

Dim formx3 As Single, formy3 As Single
Dim formx4 As Single, formy4 As Single

Dim vCorner(1 To 4) As VertexType
Dim vOrigin As VertexType

Dim listOfPoints() As VertexType
    
Dim twipsToPixelX As Single, twipsToPixelY As Single
Dim i As Integer
    
' Calculate the vertices of the rectangle
formx1! = xcenter! - xwidth! / 2#
formy1! = ycenter! - ywidth! / 2#

formx2! = xcenter! + xwidth! / 2#
formy2! = ycenter! - ywidth! / 2#

formx3! = xcenter! + xwidth! / 2#
formy3! = ycenter! + ywidth! / 2#

formx4! = xcenter! - xwidth! / 2#
formy4! = ycenter! + ywidth! / 2#

' Rotate the rectangle based on stage to image rotation
vCorner(1).X! = formx1!
vCorner(1).Y! = formy1!

vCorner(2).X! = formx2!
vCorner(2).Y! = formy2!

vCorner(3).X! = formx3!
vCorner(3).Y! = formy3!

vCorner(4).X! = formx4!
vCorner(4).Y! = formy4!

vOrigin.X! = xcenter!
vOrigin.Y! = ycenter!

' Rotate the rectangle for each corner
vCorner(1) = PictureSnapRotateVertex(vCorner(1), vOrigin, rotation!)
vCorner(2) = PictureSnapRotateVertex(vCorner(2), vOrigin, rotation!)
vCorner(3) = PictureSnapRotateVertex(vCorner(3), vOrigin, rotation!)
vCorner(4) = PictureSnapRotateVertex(vCorner(4), vOrigin, rotation!)

' Load the rotated corners
formx1! = vCorner(1).X
formy1! = vCorner(1).Y

formx2! = vCorner(2).X
formy2! = vCorner(2).Y

formx3! = vCorner(3).X
formy3! = vCorner(3).Y

formx4! = vCorner(4).X
formy4! = vCorner(4).Y

' Draw the rectangle lines (old native VB6 code)
'FormPICTURESNAP.Picture2.DrawWidth = twidth%
'FormPICTURESNAP.Picture2.Line (formx1!, formy1!)-(formx2!, formy2!), tcolor&
'FormPICTURESNAP.Picture2.Line (formx2!, formy2!)-(formx3!, formy3!), tcolor&
'
'FormPICTURESNAP.Picture2.Line (formx3!, formy3!)-(formx4!, formy4!), tcolor&
'FormPICTURESNAP.Picture2.Line (formx4!, formy4!)-(formx1!, formy1!), tcolor&
    
    ' GDI+, like most graphics libraries, operates in pixel measurements.  Convert all twips measurements to pixels.
    ReDim listOfPoints(0 To 3) As VertexType
    listOfPoints(0).X = formx1!
    listOfPoints(0).Y = formy1!
    listOfPoints(1).X = formx2!
    listOfPoints(1).Y = formy2!
    listOfPoints(2).X = formx3!
    listOfPoints(2).Y = formy3!
    listOfPoints(3).X = formx4!
    listOfPoints(3).Y = formy4!
    
    twipsToPixelX = Screen.TwipsPerPixelX
    If (twipsToPixelX = 0!) Then twipsToPixelX = 15!
    twipsToPixelY = Screen.TwipsPerPixelY
    If (twipsToPixelY = 0!) Then twipsToPixelY = 15!
    
    For i% = 0 To 3
        listOfPoints(i%).X = listOfPoints(i%).X / twipsToPixelX
        listOfPoints(i%).Y = listOfPoints(i%).Y / twipsToPixelY
    Next i%
    
    ' Render each line in turn
    For i% = 0 To 3
        If (i% < 3) Then
            GDIPlus_Interface.GDIPlus_DrawLine FormPICTURESNAP.Picture2.hDC, listOfPoints(i%).X, listOfPoints(i%).Y, listOfPoints(i% + 1).X, listOfPoints(i% + 1).Y, tcolor&, tWidth%
        Else
            GDIPlus_Interface.GDIPlus_DrawLine FormPICTURESNAP.Picture2.hDC, listOfPoints(i%).X, listOfPoints(i%).Y, listOfPoints(0).X, listOfPoints(0).Y, tcolor&, tWidth%
        End If
    Next i%
    
Exit Sub

' Errors
PictureSnapDrawRectangleError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapDrawRectangle"
ierror = True
Exit Sub

End Sub

Sub PictureSnapDrawRectangle2(xcenter As Single, ycenter As Single, xwidth As Single, ywidth As Single, rotation As Single, tcolor As Long, tWidth As Integer)
' Draws a rectangle of the specified width and height at the specified screen location (all values in screen units) for FormPICTURESNAP3.Image1
'  xcenter = rectangle center x form position
'  ycenter = rectangle center y form position
'  xwidth = rectangle width in form units
'  ywidth = rectangle height in form units
'  tcolor = line color
'  twidth = line width

ierror = False
On Error GoTo PictureSnapDrawRectangle2Error

Dim formx1 As Single, formy1 As Single
Dim formx2 As Single, formy2 As Single

Dim formx3 As Single, formy3 As Single
Dim formx4 As Single, formy4 As Single

Dim vCorner(1 To 4) As VertexType
Dim vOrigin As VertexType

Dim listOfPoints() As VertexType
    
Dim twipsToPixelX As Single, twipsToPixelY As Single
Dim i As Integer
    
' Calculate the vertices of the rectangle
formx1! = xcenter! - xwidth! / 2#
formy1! = ycenter! - ywidth! / 2#

formx2! = xcenter! + xwidth! / 2#
formy2! = ycenter! - ywidth! / 2#

formx3! = xcenter! + xwidth! / 2#
formy3! = ycenter! + ywidth! / 2#

formx4! = xcenter! - xwidth! / 2#
formy4! = ycenter! + ywidth! / 2#

' Rotate the rectangle based on stage to image rotation
vCorner(1).X! = formx1!
vCorner(1).Y! = formy1!

vCorner(2).X! = formx2!
vCorner(2).Y! = formy2!

vCorner(3).X! = formx3!
vCorner(3).Y! = formy3!

vCorner(4).X! = formx4!
vCorner(4).Y! = formy4!

vOrigin.X! = xcenter!
vOrigin.Y! = ycenter!

' Rotate the rectangle for each corner
vCorner(1) = PictureSnapRotateVertex(vCorner(1), vOrigin, rotation!)
vCorner(2) = PictureSnapRotateVertex(vCorner(2), vOrigin, rotation!)
vCorner(3) = PictureSnapRotateVertex(vCorner(3), vOrigin, rotation!)
vCorner(4) = PictureSnapRotateVertex(vCorner(4), vOrigin, rotation!)

' Load the rotated corners
formx1! = vCorner(1).X
formy1! = vCorner(1).Y

formx2! = vCorner(2).X
formy2! = vCorner(2).Y

formx3! = vCorner(3).X
formy3! = vCorner(3).Y

formx4! = vCorner(4).X
formy4! = vCorner(4).Y
    
    ' GDI+, like most graphics libraries, operates in pixel measurements.  Convert all twips measurements to pixels.
    ReDim listOfPoints(0 To 3) As VertexType
    listOfPoints(0).X = formx1!
    listOfPoints(0).Y = formy1!
    listOfPoints(1).X = formx2!
    listOfPoints(1).Y = formy2!
    listOfPoints(2).X = formx3!
    listOfPoints(2).Y = formy3!
    listOfPoints(3).X = formx4!
    listOfPoints(3).Y = formy4!
    
    ' Draw using VB graphics
    'FormPICTURESNAP3.DrawWidth = 2
    'FormPICTURESNAP3.Line (listOfPoints(0).x, listOfPoints(0).Y)-(listOfPoints(1).x, listOfPoints(1).Y), tcolor&
    'FormPICTURESNAP3.Line (listOfPoints(1).x, listOfPoints(1).Y)-(listOfPoints(2).x, listOfPoints(2).Y), tcolor&
    'FormPICTURESNAP3.Line (listOfPoints(2).x, listOfPoints(2).Y)-(listOfPoints(3).x, listOfPoints(3).Y), tcolor&
    'FormPICTURESNAP3.Line (listOfPoints(3).x, listOfPoints(3).Y)-(listOfPoints(0).x, listOfPoints(0).Y), tcolor&
        
    ' Draw using GDI code
    twipsToPixelX = Screen.TwipsPerPixelX
    If (twipsToPixelX = 0!) Then twipsToPixelX = 15!
    twipsToPixelY = Screen.TwipsPerPixelY
    If (twipsToPixelY = 0!) Then twipsToPixelY = 15!
    
    For i% = 0 To 3
        listOfPoints(i%).X = listOfPoints(i%).X / twipsToPixelX
        listOfPoints(i%).Y = listOfPoints(i%).Y / twipsToPixelY
    Next i%
    
    ' Render each line in turn on FormPICTURESNAP3 main form because image controls do not have an hDC
    For i% = 0 To 3
        If (i% < 3) Then
            GDIPlus_Interface.GDIPlus_DrawLine FormPICTURESNAP3.hDC, listOfPoints(i%).X, listOfPoints(i%).Y, listOfPoints(i% + 1).X, listOfPoints(i% + 1).Y, tcolor&, tWidth%
        Else
            GDIPlus_Interface.GDIPlus_DrawLine FormPICTURESNAP3.hDC, listOfPoints(i%).X, listOfPoints(i%).Y, listOfPoints(0).X, listOfPoints(0).Y, tcolor&, tWidth%
        End If
    Next i%
    
Exit Sub

' Errors
PictureSnapDrawRectangle2Error:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapDrawRectangle2"
ierror = True
Exit Sub

End Sub

Sub PictureSnapDrawStageLimits2()
' Draw stage limits for FormPICTURESNAP and FormPICTURESNAP3 (using rotation)

ierror = False
On Error GoTo PictureSnapDrawStageLimits2Error

Dim tcolor As Long
Dim tWidth As Integer

Dim fractionx As Single, fractiony As Single

Dim fractionx1 As Single, fractiony1 As Single
Dim fractionx2 As Single, fractiony2 As Single
Dim lowx As Single, lowy As Single
Dim highx As Single, highy As Single
Dim zwidth As Single, zdistance As Single                   ' just leave zero

Dim stagex As Single, stagey As Single, stagez As Single        ' stage limit center in stage coordinates
Dim centerx As Single, centery As Single, centerz As Single     ' stage limit center in form coordinates

Dim xdistance As Single, ydistance As Single          ' x and y stage limit distances in stage units
Dim xwidth As Single, ywidth As Single                ' x and y stage limit distances in form (screen) units

' If no picture just exit
If PictureSnapFilename$ = vbNullString Then Exit Sub

' If not calibrated, just exit
If Not PictureSnapCalibrated Then Exit Sub

' Skip if interface is busy
If RealTimeInterfaceBusy Then Exit Sub

' Skip if pausing automation
If RealTimePauseAutomation Then Exit Sub

' If form not visible just exit
If Not FormPICTURESNAP.Visible Then Exit Sub

' Calculate horizontal and vertical distances for stage limits
xdistance! = MotHiLimits!(XMotor%) - MotLoLimits(XMotor%)
ydistance! = MotHiLimits!(YMotor%) - MotLoLimits(YMotor%)

' Convert stage limit distance from stage units to form units (using image rotation)
xwidth! = PictureSnapConvertStageDistancetoImageDistance(Int(0), xdistance!, PictureSnapRotation!)
If ierror Then Exit Sub
ywidth! = PictureSnapConvertStageDistancetoImageDistance(Int(1), ydistance!, PictureSnapRotation!)
If ierror Then Exit Sub

' Calculate stage limit centers
stagex! = MotLoLimits(XMotor%) + (MotHiLimits!(XMotor%) - MotLoLimits(XMotor%)) / 2#
stagey! = MotLoLimits(YMotor%) + (MotHiLimits!(YMotor%) - MotLoLimits(YMotor%)) / 2#
stagez! = MotLoLimits(ZMotor%) + (MotHiLimits!(ZMotor%) - MotLoLimits(ZMotor%)) / 2#

' Convert center of stage limits to form coordinates
Call PictureSnapConvert(Int(2), centerx!, centery!, centerz!, stagex!, stagey!, stagez!, fractionx!, fractiony!)
If ierror Then Exit Sub

' Set line color and width
tcolor& = vbYellow
tWidth% = 2

' New code to draw stage limit rectangle using rectangle rotation to FormPICTURESNAP.Picture2 control
Call PictureSnapDrawRectangle(centerx!, centery!, xwidth!, ywidth!, PictureSnapRotation!, tcolor&, tWidth%)
If ierror Then Exit Sub

' Calculate stage limit distances in FormPICTURESNAP.Picture2 form units to FormPICTURESNAP3 form units
xwidth! = xwidth! * FormPICTURESNAP3.ScaleWidth / FormPICTURESNAP.Picture2.ScaleWidth
ywidth! = ywidth! * FormPICTURESNAP3.ScaleHeight / FormPICTURESNAP.Picture2.ScaleHeight

' Now draw stage limit rectangle to FormPICTURESNAP3.Image1 control
Call PictureSnapDrawRectangle2(FormPICTURESNAP3.ScaleWidth * fractionx!, FormPICTURESNAP3.ScaleHeight * fractiony!, xwidth!, ywidth!, PictureSnapRotation!, tcolor&, tWidth%)
If ierror Then Exit Sub

Exit Sub

' Errors
PictureSnapDrawStageLimits2Error:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapDrawStageLimits2"
ierror = True
Exit Sub

End Sub


