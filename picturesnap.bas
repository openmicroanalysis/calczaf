Attribute VB_Name = "CodePictureSnap"
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Global PictureSnapClicked As Integer
Global PictureSnapDisplayCalibrationPointsFlag As Boolean

Dim CurrentPointX As Single
Dim CurrentPointY As Single

' Scale bar variables for scroll event
Dim oldx As Single, oldy As Single

Sub PictureSnapLoad()
' Load the PictureSnap form (automatically load picture if already specified)

ierror = False
On Error GoTo PictureSnapLoadError

' Check for stage motors
If NumberOfStageMotors% < 1 Then GoTo PictureSnapLoadNoStage

' If picture file is already specified and found then load it
If PictureSnapFilename$ <> vbNullString And Dir$(PictureSnapFilename$) <> vbNullString Then
Screen.MousePointer = vbHourglass
Set FormPICTURESNAP.Picture2 = LoadPicture(PictureSnapFilename$)

' Minimize and restore to re-size
FormPICTURESNAP.WindowState = vbMinimized
FormPICTURESNAP.WindowState = vbNormal
Screen.MousePointer = vbDefault

' Update form caption
If RealTimeMode And PictureSnapCalibrated Then
FormPICTURESNAP.Caption = "Picture Snap [" & PictureSnapFilename$ & "] (double-click to move)"
Else
FormPICTURESNAP.Caption = "Picture Snap [" & PictureSnapFilename$ & "]"
End If

Else
PictureSnapFilename$ = vbNullString
PictureSnapCalibrated = False
End If

' If not realtime then disable stagemap and calibrate windows
If Not RealTimeMode Then
If MiscStringsAreSame(app.EXEName, "CalcImage") Then
FormPICTURESNAP.menuWindowCalibrate.Enabled = True
FormPICTURESNAP.menuWindowFullPicture.Enabled = True
Else
FormPICTURESNAP.menuWindowCalibrate.Enabled = False
FormPICTURESNAP.menuWindowFullPicture.Enabled = False
End If

FormPICTURESNAP.menuDisplayStandards.Enabled = False
FormPICTURESNAP.menuDisplayUnknowns.Enabled = False
FormPICTURESNAP.menuDisplayWavescans.Enabled = False
FormPICTURESNAP.menuDisplayLongLabels.Enabled = False
FormPICTURESNAP.menuDisplayShortLabels.Enabled = False

FormPICTURESNAP.menuMiscUseBeamBlankForStageMotion.Enabled = False
FormPICTURESNAP.menuMiscUseRightMouseClickToDigitize.Enabled = False
FormPICTURESNAP.menuMiscMaintainAspectRatioOfFullViewWindow.Enabled = False
End If

FormPICTURESNAP.Show vbModeless

' Check if a calibration file already exists and load if found
If PictureSnapFilename$ <> vbNullString And Dir$(PictureSnapFilename$) <> vbNullString Then
Call PictureSnapLoadCalibration
If ierror Then Exit Sub

' Enable output menus
FormPICTURESNAP.menuFileClipboard1.Enabled = True
FormPICTURESNAP.menuFileClipboard2.Enabled = True
FormPICTURESNAP.menuFileSaveAsBMPOnly.Enabled = True
FormPICTURESNAP.menuFileSaveAsBMP.Enabled = True
FormPICTURESNAP.menuFilePrintSetup.Enabled = True
FormPICTURESNAP.menuFilePrint.Enabled = True
FormPICTURESNAP.menuFileSaveAsGRD.Enabled = True
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

If Not MiscStringsAreSame(app.EXEName, "CalcImage") And Not RealTimeMode Then Exit Sub

' Load PictureSnap mode
FormPICTURESNAP2.OptionPictureSnapMode(PictureSnapMode%).Value = True

' Load current stage positions
FormPICTURESNAP2.TextXStage1.Text = RealTimeMotorPositions!(XMotor%)
FormPICTURESNAP2.TextYStage1.Text = RealTimeMotorPositions!(YMotor%)

FormPICTURESNAP2.TextXStage2.Text = RealTimeMotorPositions!(XMotor%)
FormPICTURESNAP2.TextYStage2.Text = RealTimeMotorPositions!(YMotor%)

FormPICTURESNAP2.TextXStage3.Text = RealTimeMotorPositions!(XMotor%)
FormPICTURESNAP2.TextYStage3.Text = RealTimeMotorPositions!(YMotor%)

' Load current z stage positions
If PictureSnapMode% = 1 Then
FormPICTURESNAP2.TextZStage1.Text = RealTimeMotorPositions!(ZMotor%)
FormPICTURESNAP2.TextZStage2.Text = RealTimeMotorPositions!(ZMotor%)
FormPICTURESNAP2.TextZStage3.Text = RealTimeMotorPositions!(ZMotor%)
End If

' If the picture is loaded and calibrated, load the existing calibration
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
FormPICTURESNAP2.Command9.Caption = "Do Not Display Calibration Points"
Else
FormPICTURESNAP2.Command9.Caption = "Display Calibration Points"
End If

If PictureSnapCalibrated Then
FormPICTURESNAP2.LabelCalibration.Caption = "Image Is Calibrated"
Else
FormPICTURESNAP2.LabelCalibration.Caption = "Image Is NOT Calibrated"
End If

' Load column conditions to hidden text fields
FormPICTURESNAP2.TextkeV.Text = DefaultKiloVolts!
FormPICTURESNAP2.TextMag.Text = DefaultMagnification!
FormPICTURESNAP2.TextScan.Text = DefaultScanRotation!

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

End Sub

Sub PictureSnapCalibrate(mode As Integer)
' Calculate the picture calibration (stage registration)
'  mode = 0 confirm with user
'  mode = 1 do not confirm with user

ierror = False
On Error GoTo PictureSnapCalibrateError

Dim formx As Single, formy As Single, formz As Single

' Check for picture loaded
If Trim$(PictureSnapFilename$) = vbNullString Then GoTo PictureSnapCalibrateNoPicture

' Check for in bounds
If RealTimeMode Then
If Not MiscMotorInBounds(XMotor%, Val(FormPICTURESNAP2.TextXStage1.Text)) Then GoTo PictureSnapCalibrateOutofBoundsX
If Not MiscMotorInBounds(YMotor%, Val(FormPICTURESNAP2.TextYStage1.Text)) Then GoTo PictureSnapCalibrateOutofBoundsY

If Not MiscMotorInBounds(XMotor%, Val(FormPICTURESNAP2.TextXStage2.Text)) Then GoTo PictureSnapCalibrateOutofBoundsX
If Not MiscMotorInBounds(YMotor%, Val(FormPICTURESNAP2.TextYStage2.Text)) Then GoTo PictureSnapCalibrateOutofBoundsY

If PictureSnapMode% = 1 Then
If Not MiscMotorInBounds(XMotor%, Val(FormPICTURESNAP2.TextXStage3.Text)) Then GoTo PictureSnapCalibrateOutofBoundsX
If Not MiscMotorInBounds(YMotor%, Val(FormPICTURESNAP2.TextYStage3.Text)) Then GoTo PictureSnapCalibrateOutofBoundsY
End If

If PictureSnapMode% = 1 And NumberOfStageMotors% > 2 Then
If Not MiscMotorInBounds(ZMotor%, Val(FormPICTURESNAP2.TextZStage1.Text)) Then GoTo PictureSnapCalibrateOutofBoundsZ
If Not MiscMotorInBounds(ZMotor%, Val(FormPICTURESNAP2.TextZStage2.Text)) Then GoTo PictureSnapCalibrateOutofBoundsZ
If Not MiscMotorInBounds(ZMotor%, Val(FormPICTURESNAP2.TextZStage3.Text)) Then GoTo PictureSnapCalibrateOutofBoundsZ
End If
End If

' Save the form variables
Call PictureSnapCalibrateSave(FormPICTURESNAP2)
If ierror Then Exit Sub

FormPICTURESNAP.Caption = "Picture Snap [" & PictureSnapFilename$ & "] (double-click to move)"
PictureSnapCalibrated = True

If PictureSnapCalibrated Then
FormPICTURESNAP2.LabelCalibration.Caption = "Image Is Calibrated"
Else
FormPICTURESNAP2.LabelCalibration.Caption = "Image Is NOT Calibrated"
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

Exit Sub

' Errors
PictureSnapCalibrateError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapCalibrate"
ierror = True
Exit Sub

PictureSnapCalibrateNoPicture:
msg$ = "No picture (*.BMP) has been loaded in the PictureSnap window. Please open a sample picture using the File | Open menu."
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

PictureSnapClicked = False
Screen.MousePointer = vbArrowQuestion
FormPICTURESNAP2.TextXPixel1.ForeColor = vbBlack
FormPICTURESNAP2.TextYPixel1.ForeColor = vbBlack
Do Until PictureSnapClicked Or icancel
Call MiscDelay5(CDbl(0.2), Now) ' delay a little
If ierror Then Exit Sub
Loop
Screen.MousePointer = vbDefault
DoEvents
FormPICTURESNAP2.TextXPixel1.Text = CurrentPointX!  ' save screen coordinates
FormPICTURESNAP2.TextYPixel1.Text = CurrentPointY!
FormPICTURESNAP2.TextXPixel1.ForeColor = vbRed
FormPICTURESNAP2.TextYPixel1.ForeColor = vbRed
End If

' Point 2 pixel
If mode% = 2 Then
FormPICTURESELECT.Show vbModeless
DoEvents

PictureSnapClicked = False
Screen.MousePointer = vbArrowQuestion
FormPICTURESNAP2.TextXPixel2.ForeColor = vbBlack
FormPICTURESNAP2.TextYPixel2.ForeColor = vbBlack
Do Until PictureSnapClicked Or icancel
Call MiscDelay5(CDbl(0.2), Now) ' delay a little
If ierror Then Exit Sub
Loop
Screen.MousePointer = vbDefault
DoEvents
FormPICTURESNAP2.TextXPixel2.Text = CurrentPointX!   ' save screen coordinates
FormPICTURESNAP2.TextYPixel2.Text = CurrentPointY!
FormPICTURESNAP2.TextXPixel2.ForeColor = vbRed
FormPICTURESNAP2.TextYPixel2.ForeColor = vbRed
End If

' Point 3 pixel
If mode% = 3 Then
FormPICTURESELECT.Show vbModeless
DoEvents

PictureSnapClicked = False
Screen.MousePointer = vbArrowQuestion
FormPICTURESNAP2.TextXPixel3.ForeColor = vbBlack
FormPICTURESNAP2.TextYPixel3.ForeColor = vbBlack
Do Until PictureSnapClicked Or icancel
Call MiscDelay5(CDbl(0.2), Now) ' delay a little
If ierror Then Exit Sub
Loop
Screen.MousePointer = vbDefault
DoEvents
FormPICTURESNAP2.TextXPixel3.Text = CurrentPointX!  ' save screen coordinates
FormPICTURESNAP2.TextYPixel3.Text = CurrentPointY!
FormPICTURESNAP2.TextXPixel3.ForeColor = vbRed
FormPICTURESNAP2.TextYPixel3.ForeColor = vbRed
End If

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
FormPICTURESNAP.Caption = "Picture Snap [" & PictureSnapFilename$ & "], Pixel X=" & Format$(xpix&) & ", Y=" & Format$(ypix&)
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
FormPICTURESNAP.Caption = "Picture Snap [" & PictureSnapFilename$ & "], Stage X=" & MiscAutoFormat$(stagex!) & ", Y=" & MiscAutoFormat$(stagey!)
Else
If PictureSnapCalibrationNumberofZPoints% = 0 Then
FormPICTURESNAP.Caption = "Picture Snap [" & PictureSnapFilename$ & "] (double-click to move), Stage X=" & MiscAutoFormat$(stagex!) & ", Y=" & MiscAutoFormat$(stagey!)
Else
FormPICTURESNAP.Caption = "Picture Snap [" & PictureSnapFilename$ & "] (double-click to move), Stage X=" & MiscAutoFormat$(stagex!) & ", Y=" & MiscAutoFormat$(stagey!) & ", Z=" & MiscAutoFormat$(stagez!)
End If
End If

Exit Sub

' Errors
PictureSnapUpdateCursorError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapUpdateCursor"
ierror = True
Exit Sub

End Sub

Sub PictureSnapLoadCalibration()
' Load the picture calibration from a text file of INI format (*.ACQ) (called when loading a BMP file in case it exists)

ierror = False
On Error GoTo PictureSnapLoadCalibrationError

Dim tfilename As String, tfilename2 As String

' Load PictureSnapMode
FormPICTURESNAP2.OptionPictureSnapMode(PictureSnapMode%).Value = True

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
FormPICTURESNAP.Caption = "Picture Snap [" & PictureSnapFilename$ & "]"
Else
FormPICTURESNAP.Caption = "Picture Snap [" & PictureSnapFilename$ & "] (double-click to move)"
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
' Draw current position on pic2

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
If oldx! <> formx! Or oldy! <> formy! Then
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
FormPICTURESNAP.Picture2.DrawWidth = 1

' Update full window
If FormPICTURESNAP3.Visible Then
tWidth! = FormPICTURESNAP3.ScaleWidth   ' calculate a radius
If tWidth! <> 0# Then
radius! = (tWidth! / 50#) ^ 0.8

' Erase the old circle
If oldx! <> formx! Or oldy! <> formy! Then
FormPICTURESNAP3.Image1.Refresh
End If

' Draw current position on full view window
FormPICTURESNAP3.DrawWidth = 2
FormPICTURESNAP3.Circle (FormPICTURESNAP3.ScaleWidth * fractionx!, FormPICTURESNAP3.ScaleHeight * fractiony!), radius!, RGB(150, 0, 150)
FormPICTURESNAP3.DrawWidth = 1
End If
End If

' Display calibration points if indicated
If PictureSnapDisplayCalibrationPointsFlag Then
Call PictureSnapDisplayCalibrationPoints(FormPICTURESNAP, FormPICTURESNAP3)
If ierror Then Exit Sub
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

Dim gX_Polarity As Integer, gY_Polarity As Integer
Dim gStage_Units As String

' If picture file is already specified load it
If PictureSnapFilename$ = vbNullString Then
msg$ = "No picture file has been opened yet. Use the File | Open menu in the Picture Snap window to open a JPG, BMP or TIF optically scanned image of your sample."
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapLoadFullWindow"
ierror = True
Exit Sub
End If

' Check for existing GRD info
Call GridCheckGRDInfo(PictureSnapFilename$, gX_Polarity%, gY_Polarity%, gStage_Units$)
If ierror Then Exit Sub

' Load into picturebox control to perform flipping
Screen.MousePointer = vbHourglass
Set FormPICTURESNAP3.Image1.Picture = LoadPicture(PictureSnapFilename$)

' Invert X (if JEOL config loading Cameca or Cameca config loading JEOL files)
If Default_X_Polarity% <> gX_Polarity% Then
Set FormPICTURESNAP3.Picture1.Picture = FormPICTURESNAP3.Image1.Picture
FormPICTURESNAP3.Picture1.AutoSize = True
FormPICTURESNAP3.Picture1.AutoRedraw = True
FormPICTURESNAP3.Picture1.PaintPicture FormPICTURESNAP3.Picture1.Picture, FormPICTURESNAP3.Picture1.ScaleWidth, 0, -FormPICTURESNAP3.Picture1.ScaleWidth, FormPICTURESNAP3.Picture1.ScaleHeight
FormPICTURESNAP3.Image1.Picture = FormPICTURESNAP3.Picture1.Image
End If

' Invert Y (if JEOL config loading Cameca or  Cameca config loading JEOL files)
If Default_Y_Polarity% <> gY_Polarity% Then
Set FormPICTURESNAP3.Picture1.Picture = FormPICTURESNAP3.Image1.Picture
FormPICTURESNAP3.Picture1.AutoSize = True
FormPICTURESNAP3.Picture1.AutoRedraw = True
FormPICTURESNAP3.Picture1.PaintPicture FormPICTURESNAP3.Picture1.Picture, 0, FormPICTURESNAP3.Picture1.ScaleHeight, FormPICTURESNAP3.Picture1.ScaleWidth, -FormPICTURESNAP3.Picture1.ScaleHeight
FormPICTURESNAP3.Image1.Picture = FormPICTURESNAP3.Picture1.Image
End If

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

End Sub

Sub PictureSnapSaveMode(Index As Integer)
' Save the PictureSnap mode (0 = two points, 1 = three points)

ierror = False
On Error GoTo PictureSnapSaveModeError

' If going from two points to three points and image was calibrated, re-set
If FormPICTURESNAP2.Visible Then
If PictureSnapCalibrated And Index% = 1 Then PictureSnapCalibrated = False
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
End If

Exit Sub

' Errors
PictureSnapSaveModeError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapSaveMode"
ierror = True
Exit Sub

End Sub

Sub PictureSnapDrawScaleBar()
' Load the palette display for the passed form

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

Dim gX_Polarity As Integer, gY_Polarity As Integer
Dim gStage_Units As String

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
If oldx! <> FormPICTURESNAP.ScaleWidth Or oldy! <> FormPICTURESNAP.ScaleHeight Then
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

' Check for existing GRD info
Call GridCheckGRDInfo(PictureSnapFilename$, gX_Polarity%, gY_Polarity%, gStage_Units$)
If ierror Then Exit Sub

' Update micron scale bar conversion
If Default_Stage_Units$ = "um" Then tStageConversion! = 1#
If Default_Stage_Units$ = "hm" Then tStageConversion! = 100#
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
tulefty! = FormPICTURESNAP.ScaleHeight - 1000

' Add scroll offset to keep scale bar in view
If FormPICTURESNAP.HScroll1.Max > 0 Then
x1! = FormPICTURESNAP.HScroll1.Value / FormPICTURESNAP.HScroll1.Max
End If
If FormPICTURESNAP.VScroll1.Max > 0 Then
y1! = FormPICTURESNAP.VScroll1.Value / FormPICTURESNAP.VScroll1.Max
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
FormPICTURESNAP.Picture2.DrawWidth = 1
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
