Attribute VB_Name = "CodePictureSnap5"
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

Sub PictureSnapDigitizePoint(mode As Integer, BitMapX As Single, BitMapY As Single)
' Dummy routine for CalcImage

ierror = False
On Error GoTo PictureSnapDigitizePointError


Exit Sub

' Errors
PictureSnapDigitizePointError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapDigitizePoint"
ierror = True
Exit Sub

End Sub

Sub PictureSnapStageMove(BitMapX As Single, BitMapY As Single)
' Dummy routine for CalcImage

ierror = False
On Error GoTo PictureSnapStageMoveError

Dim stagex As Single, stagey As Single, stagez As Single
Dim fractionx As Single, fractiony As Single

' Check that image is loaded and calibrated
If PictureSnapFilename$ = vbNullString Then Exit Sub
If Not PictureSnapCalibrated Then Exit Sub

' Convert to stage coordinates
Call PictureSnapConvert(Int(1), BitMapX!, BitMapY!, CSng(0#), stagex!, stagey!, stagez!, fractionx!, fractiony!)
If ierror Then Exit Sub

' Just force Realtime positions to stage positions
RealTimeMotorPositions!(XMotor%) = stagex!
RealTimeMotorPositions!(YMotor%) = stagey!
Exit Sub

' Errors
PictureSnapStageMoveError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapStageMove"
ierror = True
Exit Sub

End Sub

Sub PictureSnapStageMove2(BitMapX As Single, BitMapY As Single)
' Dummy routine for CalcImage

ierror = False
On Error GoTo PictureSnapStageMove2Error


Exit Sub

' Errors
PictureSnapStageMove2Error:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapStageMove2"
ierror = True
Exit Sub

End Sub

Sub RealTimeGetBeamMode(tbeammode As Integer)
' Dummy routine for CalcImage

ierror = False
On Error GoTo RealTimeGetBeamModeError


Exit Sub

' Errors
RealTimeGetBeamModeError:
MsgBox Error$, vbOKOnly + vbCritical, "RealTimeGetBeamMode"
ierror = True
Exit Sub

End Sub

Sub RealTimeGetMagnification(mag As Single)
' Dummy routine for CalcImage

ierror = False
On Error GoTo RealTimeGetMagnificationError


Exit Sub

' Errors
RealTimeGetMagnificationError:
MsgBox Error$, vbOKOnly + vbCritical, "RealTimeGetMagnification"
ierror = True
Exit Sub

End Sub

Sub PictureSnapDrawLineRectangle()
' Dummy routine for CalcImage

ierror = False
On Error GoTo PictureSnapDrawLineRectangleError


Exit Sub

' Errors
PictureSnapDrawLineRectangleError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapDrawLineRectangle"
ierror = True
Exit Sub

End Sub

Sub PictureSnapMoveToCalibrationPoint(stagex As Single, stagey As Single, stagez As Single)
' Dummy routine for CalcImage

ierror = False
On Error GoTo PictureSnapMoveToCalibrationPointError


Exit Sub

' Errors
PictureSnapMoveToCalibrationPointError:
MsgBox Error$, vbOKOnly + vbCritical, "PictureSnapMoveToCalibrationPoint"
ierror = True
Exit Sub

End Sub

