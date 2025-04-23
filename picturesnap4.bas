Attribute VB_Name = "CodePictureSnapStageMove"
' (c) Copyright 1995-2025 by John J. Donovan
Option Explicit

Dim DrawLineRectanglePositions1(1 To 2) As Single
Dim DrawLineRectanglePositions2(1 To 2) As Single

Dim position As TypePosition
Dim positiondata As TypePositionData

Sub PictureSnapCalibratePointStage(mode As Integer)
' Load the stage coordinates for the selected point
'  mode = 4 load point #1 stage coordinates
'  mode = 5 load point #2 stage coordinates
'  mode = 6 load point #3 stage coordinates     ' only if PictureSnapMode% = 1

ierror = False
On Error GoTo PictureSnapCalibratePointStageError

' Check for picture loaded
If Trim$(PictureSnapFilename$) = vbNullString Then GoTo PictureSnapCalibratePointStageNoPicture

' Point 1 stage
If mode% = 4 Then
Call RealTimeGetAllPositions(Int(1))
FormPICTURESNAP2.TextXStage1.Text = RealTimeMotorPositions!(XMotor%)
FormPICTURESNAP2.TextYStage1.Text = RealTimeMotorPositions!(YMotor%)
FormPICTURESNAP2.TextZStage1.Text = RealTimeMotorPositions!(ZMotor%)
End If

' Point 2 stage
If mode% = 5 Then
Call RealTimeGetAllPositions(Int(1))
FormPICTURESNAP2.TextXStage2.Text = RealTimeMotorPositions!(XMotor%)
FormPICTURESNAP2.TextYStage2.Text = RealTimeMotorPositions!(YMotor%)
FormPICTURESNAP2.TextZStage2.Text = RealTimeMotorPositions!(ZMotor%)
End If

' Point 3 stage
If mode% = 6 Then
Call RealTimeGetAllPositions(Int(1))
FormPICTURESNAP2.TextXStage3.Text = RealTimeMotorPositions!(XMotor%)
FormPICTURESNAP2.TextYStage3.Text = RealTimeMotorPositions!(YMotor%)
FormPICTURESNAP2.TextZStage3.Text = RealTimeMotorPositions!(ZMotor%)
End If

Exit Sub

' Errors
PictureSnapCalibratePointStageError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapCalibratePointStage"
ierror = True
Exit Sub

PictureSnapCalibratePointStageNoPicture:
msg$ = "No picture (*.BMP) has been loaded in the PictureSnap window. Please open a sample picture using the File | Open menu."
MsgBox msg$, vbOKOnly + vbExclamation, "PictureSnapCalibratePointStage"
ierror = True
Exit Sub

End Sub

Sub PictureSnapStageMove(BitMapX As Single, BitMapY As Single)
' Convert and move to the stage position double clicked

ierror = False
On Error GoTo PictureSnapStageMoveError

Dim blankbeam As Boolean
Dim stagex As Single, stagey As Single, stagez As Single
Dim fractionx As Single, fractiony As Single

' If not real time then exit sub
If Not RealTimeMode Then Exit Sub

' Check that image is loaded and calibrated
If PictureSnapFilename$ = vbNullString Then Exit Sub
If Not PictureSnapCalibrated Then Exit Sub

' Convert to stage coordinates
Call PictureSnapConvert(Int(1), BitMapX!, BitMapY!, CSng(0#), stagex!, stagey!, stagez!, fractionx!, fractiony!)
If ierror Then Exit Sub

' Check if Z stage axis disabled
If FormPICTURESNAP.menuMiscDisableZStageMove.Checked Then
stagez! = RealTimeMotorPositions!(ZMotor%)
End If

' Move to position clicked by user
blankbeam = FormPICTURESNAP.menuMiscUseBeamBlankForStageMotion.Checked
Call MoveStageMoveXYZ(blankbeam, stagex!, stagey!, stagez!)
If ierror Then Exit Sub

' Store these stage positions in case drawing a rectangle
DrawLineRectanglePositions1!(1) = DrawLineRectanglePositions2!(1)
DrawLineRectanglePositions1!(2) = DrawLineRectanglePositions2!(2)

DrawLineRectanglePositions2!(1) = BitMapX!
DrawLineRectanglePositions2!(2) = BitMapY!

Exit Sub

' Errors
PictureSnapStageMoveError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapStageMove"
ierror = True
Exit Sub

End Sub

Sub PictureSnapStageMove2(BitMapX As Single, BitMapY As Single)
' Convert full view form coordinates and move to the stage position double clicked

ierror = False
On Error GoTo PictureSnapStageMove2Error

Dim blankbeam As Boolean
Dim formx As Single, formy As Single, formz As Single
Dim stagex As Single, stagey As Single, stagez As Single
Dim fractionx As Single, fractiony As Single
Dim temp As Single

' If not real time then exit sub
If Not RealTimeMode Then Exit Sub

' Check that image is loaded and calibrated
If PictureSnapFilename$ = vbNullString Then Exit Sub
If Not PictureSnapCalibrated Then Exit Sub

' Convert FormPICTURESNAP3 form coordinates to FormPICTURESNAP.Picture2 coordinates
formx! = FormPICTURESNAP.Picture2.ScaleWidth * BitMapX! / FormPICTURESNAP3.ScaleWidth
formy! = FormPICTURESNAP.Picture2.ScaleHeight * BitMapY! / FormPICTURESNAP3.ScaleHeight
formz! = 0#

' Convert to stage coordinates
Call PictureSnapConvert(Int(1), formx!, formy!, formz!, stagex!, stagey!, stagez!, fractionx!, fractiony!)
If ierror Then Exit Sub

' Convert back to form coordinates and get fraction parameters
Call PictureSnapConvert(Int(2), BitMapX!, BitMapY!, CSng(0#), stagex!, stagey!, stagez!, fractionx!, fractiony!)
If ierror Then Exit Sub

' Check if Z stage axis disabled
If FormPICTURESNAP.menuMiscDisableZStageMove.Checked Then
stagez! = RealTimeMotorPositions!(ZMotor%)
End If

' Move scroll bars on main window to this position
If fractionx! >= 0# And fractionx! <= 1# Then
temp! = FormPICTURESNAP.HScroll1.Max - FormPICTURESNAP.HScroll1.Min
FormPICTURESNAP.HScroll1.value = CInt(temp! * fractionx!)
DoEvents
End If

If fractionx! >= 0# And fractionx! <= 1# Then
temp! = FormPICTURESNAP.VScroll1.Max - FormPICTURESNAP.VScroll1.Min
FormPICTURESNAP.VScroll1.value = CInt(temp! * fractiony!)
DoEvents
End If

' Move to position clicked by user
blankbeam = FormPICTURESNAP.menuMiscUseBeamBlankForStageMotion.Checked
Call MoveStageMoveXYZ(blankbeam, stagex!, stagey!, stagez!)
If ierror Then Exit Sub

Exit Sub

' Errors
PictureSnapStageMove2Error:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapStageMove2"
ierror = True
Exit Sub

End Sub

Sub PictureSnapMoveToCalibrationPoint(stagex As Single, stagey As Single, stagez As Single)
' Move to the stage calibration point

ierror = False
On Error GoTo PictureSnapMoveToCalibrationPointError

Dim blankbeam As Boolean
Dim formx As Single, formy As Single, formz As Single
Dim fractionx As Single, fractiony As Single
Dim temp As Single

' If not real time then exit sub
If Not RealTimeMode Then Exit Sub

' Check that image is loaded and calibrated
If PictureSnapFilename$ = vbNullString Then Exit Sub
If Not PictureSnapCalibrated Then Exit Sub

' Convert back to form coordinates and get fraction parameters
Call PictureSnapConvert(Int(2), formx!, formy!, CSng(0#), stagex!, stagey!, stagez!, fractionx!, fractiony!)
If ierror Then Exit Sub

' Move scroll bars on main window to this position
temp! = FormPICTURESNAP.HScroll1.Max - FormPICTURESNAP.HScroll1.Min
If fractionx! >= 0# And fractionx! <= 1# Then
FormPICTURESNAP.HScroll1.value = CInt(temp! * fractionx!)
End If
temp! = FormPICTURESNAP.VScroll1.Max - FormPICTURESNAP.VScroll1.Min
If fractiony! >= 0# And fractiony! <= 1# Then
FormPICTURESNAP.VScroll1.value = CInt(temp! * fractiony!)
End If

If stagez! = 0# Then stagez! = RealTimeMotorPositions!(ZMotor%)

' Move to stage position
blankbeam = FormPICTURESNAP.menuMiscUseBeamBlankForStageMotion.Checked
Call MoveStageMoveXYZ(blankbeam, stagex!, stagey!, stagez!)
If ierror Then Exit Sub

Exit Sub

' Errors
PictureSnapMoveToCalibrationPointError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapMoveToCalibrationPoint"
ierror = True
Exit Sub

End Sub

Sub PictureSnapDigitizePoint(mode As Integer, BitMapX As Single, BitMapY As Single)
' Convert and digitize the stage position on the PictureSnap image
'   mode = 0 digitize from FormPICTURESNAP
'   mode = 1 digitize from FormPICTURESNAP3

ierror = False
On Error GoTo PictureSnapDigitizePointError

Dim samplerow As Integer
Dim stagex As Single, stagey As Single, stagez As Single
Dim fractionx As Single, fractiony As Single

' If not real time then exit sub
If Not RealTimeMode Then Exit Sub

' Check that image is loaded and calibrated
If PictureSnapFilename$ = vbNullString Then Exit Sub
If Not PictureSnapCalibrated Then Exit Sub

' Convert FormPICTURESNAP3 form coordinates to FormPICTURESNAP.Picture2 coordinates
If mode% = 1 Then
BitMapX! = FormPICTURESNAP.Picture2.ScaleWidth * BitMapX! / FormPICTURESNAP3.ScaleWidth
BitMapY! = FormPICTURESNAP.Picture2.ScaleHeight * BitMapY! / FormPICTURESNAP3.ScaleHeight
End If

' Convert to stage coordinates
Call PictureSnapConvert(Int(1), BitMapX!, BitMapY!, CSng(0#), stagex!, stagey!, stagez!, fractionx!, fractiony!)
If ierror Then Exit Sub

' Check for a currently selected sample
samplerow% = DigitizeCheckSample%()
If ierror Then Exit Sub

' Load current grain number and autofocus
positiondata.grainnumber% = Val(FormDIGITIZE.LabelGrainNumber.Caption)
If FormDIGITIZE.CheckAutoFocusOn.value Then
positiondata.autofocus% = True
Else
positiondata.autofocus% = False
End If

' Add point to current position sample
Call DigitizeUpdateAutomate(position)
If ierror Then Exit Sub

' Save the x and y coordinates
positiondata.xyz!(1) = stagex!
positiondata.xyz!(2) = stagey!
positiondata.xyz!(3) = stagez!

' Add the position to the database
Call DigitizeAddPosition(positiondata, FormAUTOMATE.ListDigitize, FormAUTOMATE.GridDigitize)
If ierror Then
Call IOStatusAuto(vbNullString)
Exit Sub
End If

' Update parameters in case using beam deflection
Call DigitizeUpdateForBeamDeflection(Int(0), samplerow%)
If ierror Then Exit Sub

Exit Sub

' Errors
PictureSnapDigitizePointError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapDigitizePoint"
ierror = True
Exit Sub

End Sub

Sub PictureSnapDrawLineRectangle()
' Draws a line or rectangle based on the two last positions that were double clicked

ierror = False
On Error GoTo PictureSnapDrawLineRectangleError

Dim tWidth As Integer
Dim tcolor As Long

Dim centerx As Single, centery As Single
Dim xwidth As Single, ywidth As Single

Dim widthx As Single, widthy As Single

Dim stagex1 As Single, stagey1 As Single, stagez1 As Single
Dim stagex2 As Single, stagey2 As Single, stagez2 As Single

Dim fractionx1 As Single, fractiony1 As Single
Dim fractionx2 As Single, fractiony2 As Single

If Not UseLineDrawingModeFlag And Not UseRectangleDrawingModeFlag Then Exit Sub

' Check if two positions loaded
If DrawLineRectanglePositions1!(1) = 0# And DrawLineRectanglePositions1!(2) = 0# Then Exit Sub

' Load color and thickness
tcolor& = RGB(255, 0, 0)
FormPICTURESNAP.Picture2.DrawWidth = 2

' Draw line (need to convert to GDI+ code)
If UseLineDrawingModeFlag Then
FormPICTURESNAP.Picture2.Line (DrawLineRectanglePositions1!(1), DrawLineRectanglePositions1!(2))-(DrawLineRectanglePositions2!(1), DrawLineRectanglePositions2!(2)), tcolor&
End If

' Draw rectangle
If UseRectangleDrawingModeFlag Then
'FormPICTURESNAP.Picture2.Line (DrawLineRectanglePositions1!(1), DrawLineRectanglePositions1!(2))-(DrawLineRectanglePositions2!(1), DrawLineRectanglePositions2!(2)), tcolor&, B

' Calculate screen center position
centerx! = DrawLineRectanglePositions1!(1) + (DrawLineRectanglePositions2!(1) - DrawLineRectanglePositions1!(1)) / 2#
centery! = DrawLineRectanglePositions1!(2) + (DrawLineRectanglePositions2!(2) - DrawLineRectanglePositions1!(2)) / 2#

' Convert screen corners to stage corners
Call PictureSnapConvert(Int(1), DrawLineRectanglePositions1!(1), DrawLineRectanglePositions1!(2), CSng(0#), stagex1!, stagey1!, stagez1!, fractionx1!, fractiony1!)
If ierror Then Exit Sub

Call PictureSnapConvert(Int(1), DrawLineRectanglePositions2!(1), DrawLineRectanglePositions2!(2), CSng(0#), stagex2!, stagey2!, stagez2!, fractionx2!, fractiony2!)
If ierror Then Exit Sub

' Calculate stage widths
widthx! = Abs(stagex2! - stagex1!)
widthy! = Abs(stagey2! - stagey1!)

' Convert stage distance to screen distance (to correct for rotation)
xwidth! = PictureSnapConvertStageDistancetoImageDistance(Int(0), widthx!, PictureSnapRotation!)
If ierror Then Exit Sub
ywidth! = PictureSnapConvertStageDistancetoImageDistance(Int(1), widthy!, PictureSnapRotation!)
If ierror Then Exit Sub

' New code to draw magbox corners using rectangle rotation
tWidth% = 2
Call PictureSnapDrawRectangle(centerx!, centery!, xwidth!, ywidth!, PictureSnapRotation!, tcolor&, tWidth%)
If ierror Then Exit Sub
End If

Exit Sub

PictureSnapDrawLineRectangleError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapDrawLineRectangle"
ierror = True
Exit Sub

End Sub

Sub PictureSnapMoveScrollBarsToCurrentStagePosition()
' Move the scroll bars of the main PictureSnap window to show the current stage position

ierror = False
On Error GoTo PictureSnapMoveScrollBarsToCurrentStagePositionError

Dim fractionx As Single, fractiony As Single
Dim formx As Single, formy As Single, formz As Single
Dim temp As Single

' Do not enter if not connected to stage
If Not RealTimeMode Then Exit Sub

' Check if image loaded
If PictureSnapFilename$ = vbNullString Then Exit Sub

' Check if calibrated
If Not PictureSnapCalibrated Then Exit Sub

' Convert stage coordinates to screen coordinates
Call PictureSnapConvert(Int(2), formx!, formy!, formz!, RealTimeMotorPositions!(XMotor%), RealTimeMotorPositions!(YMotor%), RealTimeMotorPositions!(ZMotor%), fractionx!, fractiony!)
If ierror Then Exit Sub

' Move scroll bars on main window to this position
If fractionx! >= 0# And fractionx! <= 1# Then
temp! = FormPICTURESNAP.HScroll1.Max - FormPICTURESNAP.HScroll1.Min
FormPICTURESNAP.HScroll1.value = CInt(temp! * fractionx!)
Else
Call IOWriteLog("PictureSnapMoveScrollBarsToCurrentStagePosition: X stage position is outside of image extents")
End If

If fractiony! >= 0# And fractiony! <= 1# Then
temp! = FormPICTURESNAP.VScroll1.Max - FormPICTURESNAP.VScroll1.Min
FormPICTURESNAP.VScroll1.value = CInt(temp! * fractiony!)
Else
Call IOWriteLog("PictureSnapMoveScrollBarsToCurrentStagePosition: Y stage position is outside of image extents")
End If

Exit Sub

PictureSnapMoveScrollBarsToCurrentStagePositionError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapMoveScrollBarsToCurrentStagePosition"
ierror = True
Exit Sub

End Sub

Sub PictureSnapMoveScrollBarsToCurrentImagePosition(BitMapX As Single, BitMapY As Single)
' Move the scroll bars of the main PictureSnap window to the full view image location (without moving stage)
' basd on double click using right mouse button

ierror = False
On Error GoTo PictureSnapMoveScrollBarsToCurrentImagePositionError

Dim stagex As Single, stagey As Single, stagez As Single
Dim tBitMapX As Single, tBitmapY As Single
Dim fractionx As Single, fractiony As Single
Dim xpix As Long, ypix As Long
Dim temp As Single

' If picture file is not loaded just exit
If PictureSnapFilename$ = vbNullString Then Exit Sub

' If not calibrated just calculate pixels and exit
If Not PictureSnapCalibrated Then
xpix& = BitMapX! / Screen.TwipsPerPixelX
ypix& = BitMapY! / Screen.TwipsPerPixelY
Exit Sub
End If

' Convert FormPICTURESNAP3 form coordinates to FormPICTURESNAP.Picture2 coordinates
tBitMapX! = FormPICTURESNAP.Picture2.ScaleWidth * BitMapX! / FormPICTURESNAP3.ScaleWidth
tBitmapY! = FormPICTURESNAP.Picture2.ScaleHeight * BitMapY! / FormPICTURESNAP3.ScaleHeight

' Convert to stage coordinates
Call PictureSnapConvert(Int(1), tBitMapX!, tBitmapY!, CSng(0#), stagex!, stagey!, stagez!, fractionx!, fractiony!)
If ierror Then Exit Sub

' Convert back to form coordinates and get fraction parameters
Call PictureSnapConvert(Int(2), tBitMapX!, tBitmapY!, CSng(0#), stagex!, stagey!, stagez!, fractionx!, fractiony!)
If ierror Then Exit Sub

' Move scroll bars on main window to this position
If fractionx! >= 0# And fractionx! <= 1# Then
temp! = FormPICTURESNAP.HScroll1.Max - FormPICTURESNAP.HScroll1.Min
FormPICTURESNAP.HScroll1.value = CInt(temp! * fractionx!)
DoEvents
End If

If fractiony! >= 0# And fractiony! <= 1# Then
temp! = FormPICTURESNAP.VScroll1.Max - FormPICTURESNAP.VScroll1.Min
FormPICTURESNAP.VScroll1.value = CInt(temp! * fractiony!)
DoEvents
End If

Exit Sub

PictureSnapMoveScrollBarsToCurrentImagePositionError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapMoveScrollBarsToCurrentImagePosition"
ierror = True
Exit Sub

End Sub


