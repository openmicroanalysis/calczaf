Attribute VB_Name = "CodeMiscReal"
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

Const MAXMEAS% = 20

Dim tmsg As String
Dim maxmeasure As Long

Function MiscGetSlitSizeIndex(motor As Integer, ssize As String) As Integer
' Return the slit size index for the given slit size on the given spectrometer

ierror = False
On Error GoTo MiscGetSlitSizeIndexError

Dim j As Integer, ip As Integer

' Assume no match
MiscGetSlitSizeIndex = 0

' Check for valid spectro
If motor% < 1 Or motor% > NumberOfTunableSpecs% Then GoTo MiscGetSlitSizeIndexBadScaler

' Loop on all Detector positions and try to match
For j% = 1 To DetSlitSizesNumber%(motor%)
If Trim$(LCase$(ssize$)) = Trim$(LCase$(DetSlitSizes$(j%, motor%))) Then ip% = j%
Next j%

MiscGetSlitSizeIndex = ip%

Exit Function

' Errors
MiscGetSlitSizeIndexError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetSlitSizeIndex"
ierror = True
Exit Function

MiscGetSlitSizeIndexBadScaler:
msg$ = "Invalid spectro number"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscGetSlitSizeIndex"
ierror = True
Exit Function

End Function

Function MiscGetCrystalIndex(motor As Integer, crystal As String) As Integer
' Return the crystal index for the given crystal on the given spectrometer

ierror = False
On Error GoTo MiscGetCrystalIndexError

Dim j As Integer, ip As Integer

' Assume no match
MiscGetCrystalIndex = 0

' Check for valid spectro
If motor% < 1 Or motor% > NumberOfTunableSpecs% Then GoTo MiscGetCrystalIndexBadScaler

' Loop on all crystal positions and try to match
For j% = 1 To ScalNumberOfCrystals%(motor%)
If Trim$(LCase$(crystal$)) = Trim$(LCase$(ScalCrystalNames$(j%, motor%))) Then ip% = j%
Next j%

'If ip% = 0 Then GoTo MiscGetCrystalIndexNoMatch ' disabled for GetElmSetElmCrystalUpdate
MiscGetCrystalIndex = ip%
Exit Function

' Errors
MiscGetCrystalIndexError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetCrystalIndex"
ierror = True
Exit Function

MiscGetCrystalIndexBadScaler:
msg$ = "Spectro" & Str$(motor%) & " is an invalid spectro number"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscGetCrystalIndex"
ierror = True
Exit Function

'MiscGetCrystalIndexNoMatch:
'msg$ = "Could not match crystal " & crystal$ & " on spectro " & Str$(motor%)
'MsgBox msg$, vbOKOnly + vbExclamation, "MiscGetCrystalIndex"
'ierror = True
'Exit Function

End Function

Sub MiscGetCrystalParameters(xtal As String, x2d As Single, k As Single, syme As String, symx As String)
' Return crystal parameters based on crystal name

ierror = False
On Error GoTo MiscGetCrystalParametersError

Dim ip As Integer

' Find position in crystal list
If Trim$(xtal$) = vbNullString Then Exit Sub            ' new for 8/7/2006 (to prevent problem in XrayCalculatePositions)
ip% = IPOS1(MAXCRYSTYPE%, xtal$, AllCrystalNames$())
If ip% = 0 Then GoTo MiscGetCrystalParametersBadXtal

' Load parameters
x2d! = AllCrystal2ds!(ip%)
k! = AllCrystalKs!(ip%)
syme$ = AllCrystalElements$(ip%)
symx$ = AllCrystalXrays$(ip%)

Exit Sub

' Errors
MiscGetCrystalParametersError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetCrystalParameters"
ierror = True
Exit Sub

MiscGetCrystalParametersBadXtal:
msg$ = "Crystal name not found in crystal list from " & CrystalsFile$
MsgBox msg$, vbOKOnly + vbExclamation, "MiscGetCrystalParameters"
ierror = True
Exit Sub

End Sub

Function MiscMotorInBounds(motor As Integer, pos As Single) As Integer
' Checks for a valid motor position

ierror = False
On Error GoTo MiscMotorInBoundsError

MiscMotorInBounds = True
If NoMotorPositionBoundsChecking(motor%) Then Exit Function

' Check for out of range motor
If motor% < 1 Or motor% > NumberOfTunableSpecs% + NumberOfStageMotors% Then GoTo MiscMotorInBoundsBadMotor

' Check high and low limits
If pos! < MotLoLimits!(motor%) Or pos! > MotHiLimits!(motor%) Then MiscMotorInBounds = False

Exit Function

' Errors
MiscMotorInBoundsError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscMotorInBounds"
ierror = True
Exit Function

MiscMotorInBoundsBadMotor:
msg$ = "Invalid motor number"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscMotorInBounds"
ierror = True
Exit Function

End Function

Function MiscGetSlitPositionIndex(motor As Integer, sposition As String) As Integer
' Return the slit position index for the given slit position on the given spectrometer

ierror = False
On Error GoTo MiscGetSlitPositionIndexError

Dim j As Integer, ip As Integer

' Assume no match
MiscGetSlitPositionIndex = 0

' Check for valid spectro
If motor% < 1 Or motor% > NumberOfTunableSpecs% Then GoTo MiscGetSlitPositionIndexBadScaler

' Loop on all Detector positions and try to match
For j% = 1 To DetSlitPositionsNumber%(motor%)
If Trim$(LCase$(sposition$)) = Trim$(LCase$(DetSlitPositions$(j%, motor%))) Then ip% = j%
Next j%

MiscGetSlitPositionIndex = ip%

Exit Function

' Errors
MiscGetSlitPositionIndexError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetSlitPositionIndex"
ierror = True
Exit Function

MiscGetSlitPositionIndexBadScaler:
msg$ = "Invalid spectro number"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscGetSlitPositionIndex"
ierror = True
Exit Function

End Function

Function MiscGetDetectorModeIndex(motor As Integer, dmode As String) As Integer
' Return the detector mode index for the given detector mode on the given spectrometer

ierror = False
On Error GoTo MiscGetDetectorModeIndexError

Dim j As Integer, ip As Integer

' Assume no match
MiscGetDetectorModeIndex = 0

' Check for valid spectro
If motor% < 1 Or motor% > NumberOfTunableSpecs% Then GoTo MiscGetDetectorModeIndexBadScaler

' Loop on all Detector positions and try to match
For j% = 1 To DetDetectorModesNumber%(motor%)
If Trim$(LCase$(dmode$)) = Trim$(LCase$(DetDetectorModes$(j%, motor%))) Then ip% = j%
Next j%

MiscGetDetectorModeIndex = ip%

Exit Function

' Errors
MiscGetDetectorModeIndexError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetDetectorModeIndex"
ierror = True
Exit Function

MiscGetDetectorModeIndexBadScaler:
msg$ = "Invalid spectro number"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscGetDetectorModeIndex"
ierror = True
Exit Function

End Function

Sub MiscSetConditionCheckAperture(mode As Integer, tbeamcurrent As Single)
' Check the aperture by comparing the number of times the beam current was set using the same condition versus how many times the current need to be set
' mode = 1 check if beam current changed and increment number of times called)
' mode = 2 increment number of times called and warn user if setting beam current too often
' tbeamcurrent = requested beam current

ierror = False
On Error GoTo MiscSetConditionCheckApertureError

Dim temp As Single

' Check for change in requested beam current
If mode% = 1 Then
If tbeamcurrent! <> LastBeamCurrentMeasured! Then
maxmeasure& = MAXMEAS%
RealTimeBeamCurrentNumberofTimesCalled& = 0
RealTimeBeamCurrentNumberofTimesSet& = 0

' Save for next call
LastBeamCurrentMeasured! = tbeamcurrent!
End If

RealTimeBeamCurrentNumberofTimesCalled& = RealTimeBeamCurrentNumberofTimesCalled& + 1
End If

' Check for dirty aperture
If mode% = 2 Then
RealTimeBeamCurrentNumberofTimesSet& = RealTimeBeamCurrentNumberofTimesSet& + 1

' Calculate rate needing to set beam current
If RealTimeBeamCurrentNumberofTimesCalled& > maxmeasure& Then    ' do not call if not enough data
temp! = 100# * RealTimeBeamCurrentNumberofTimesSet& / RealTimeBeamCurrentNumberofTimesCalled&

If VerboseMode Then
tmsg$ = "MiscSetConditionCheckAperture: " & Format$(RealTimeBeamCurrentNumberofTimesSet&) & " / " & Format$(RealTimeBeamCurrentNumberofTimesCalled&) & " = " & Format$(temp!) & "%"
Call IOWriteLog(tmsg$)
End If

' If greater than 20%, warn user
If temp! > 20# Then
tmsg$ = "The application was forced to set the beam current " & Format$(RealTimeBeamCurrentNumberofTimesSet&)
tmsg$ = tmsg$ & " times  out of the last " & Format$(RealTimeBeamCurrentNumberofTimesCalled&) & " (" & MiscAutoFormat4$(temp!) & "%)."
tmsg$ = tmsg$ & vbCrLf & vbCrLf & "This usually indicates either that the BeamCurrentTolerance parameter in the "
tmsg$ = tmsg$ & ProbeWinINIFile$ & " file is too small or that the beam current regulation aperture is dirty "
tmsg$ = tmsg$ & "and should be changed or cleaned. If not, the acquisition will require some additional time "
tmsg$ = tmsg$ & "for setting the beam current unnecessarily."
MiscMsgBoxTim FormMSGBOXTIME, "MiscSetConditionCheckAperture", tmsg$, 30#
Call IOWriteLogRichText(tmsg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))

' Increment measurement criteria for next warning
maxmeasure& = 3# * maxmeasure&
If maxmeasure& > 10# * MAXINTEGER% Then maxmeasure& = MAXMEAS%
End If
End If
End If

Exit Sub

' Errors
MiscSetConditionCheckApertureError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSetConditionCheckAperture"
ierror = True
Exit Sub

End Sub

Function MiscIsInstrumentStage(InstType As String) As Boolean
' Return true if instrument stage is passed type basd on "MotUnitsToAngstromMicrons!(XMotor%)"
' InstType$ = "JEOL"
' InstType$ = "CAMECA"

ierror = False
On Error GoTo MiscIsInstrumentStageError

' Check for valid instrument types
If InstType$ <> "JEOL" And InstType$ <> "CAMECA" Then GoTo MiscIsInstrumentStageInvalidInstType

' Check for valid motor
If XMotor% = 0 Then GoTo MiscIsInstrumentStageBadXMotor

' Check that MOTORS.DAT file has been read
If MotUnitsToAngstromMicrons!(XMotor%) = 0 Then GoTo MiscIsInstrumentStageNotInitialized

' Assume no match
MiscIsInstrumentStage = False

' If no stage motors, assume mm
If NumberOfStageMotors% = 0 Then
If InstType$ = "JEOL" Then MiscIsInstrumentStage = True
Exit Function
End If

' Check stage units
If InstType$ = "JEOL" And MotUnitsToAngstromMicrons!(XMotor%) = 1000 Then MiscIsInstrumentStage = True
If InstType$ = "CAMECA" And MotUnitsToAngstromMicrons!(XMotor%) = 1 Then MiscIsInstrumentStage = True

Exit Function

' Errors
MiscIsInstrumentStageError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscIsInstrumentStage"
ierror = True
Exit Function

MiscIsInstrumentStageInvalidInstType:
msg$ = "Invalid instrument stage type was passed (" & InstType$ & "). Please contact Probe Software Technical Support"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscIsInstrumentStage"
ierror = True
Exit Function

MiscIsInstrumentStageBadXMotor:
msg$ = "The XMotor variable is not initialized. Please contact Probe Software Technical Support"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscIsInstrumentStage"
ierror = True
Exit Function

MiscIsInstrumentStageNotInitialized:
msg$ = "The MotUnitsToAngstromMicrons paremeters have not been initialized. Please contact Probe Software Technical Support"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscIsInstrumentStage"
ierror = True
Exit Function

End Function

Function MiscIsInstrumentStageStored(InstType As String) As Boolean
' Return true if instrument stage is passed type based on "InterfaceTypeStored"
' InstType$ = "JEOL"
' InstType$ = "ARL"
' InstType$ = "CAMECA"

ierror = False
On Error GoTo MiscIsInstrumentStageStoredError

' Check for valid instrument types
If InstType$ <> "JEOL" And InstType$ <> "ARL" And InstType$ <> "CAMECA" Then GoTo MiscIsInstrumentStageStoredInvalidInstType

' Assume no match
MiscIsInstrumentStageStored = False

' For demo mode, base on stage units
If InterfaceTypeStored% = 0 Or InterfaceTypeStored% = 1 Then
MiscIsInstrumentStageStored = MiscIsInstrumentStage(InstType$)
Exit Function
End If

If InterfaceTypeStored% = 2 And InstType$ = "JEOL" Then MiscIsInstrumentStageStored = True      ' JEOL 8900/8200/8500/8x30
If InterfaceTypeStored% = 3 And InstType$ = "JEOL" Then MiscIsInstrumentStageStored = True      ' not used
If InterfaceTypeStored% = 4 And InstType$ = "JEOL" Then MiscIsInstrumentStageStored = True      ' not used
If InterfaceTypeStored% = 5 And InstType$ = "CAMECA" Then MiscIsInstrumentStageStored = True    ' Cameca SX100/SXFive
If InterfaceTypeStored% = 6 And InstType$ = "CAMECA" Then MiscIsInstrumentStageStored = True    ' not used
If InterfaceTypeStored% = 7 And InstType$ = "JEOL" Then MiscIsInstrumentStageStored = True      ' not used
If InterfaceTypeStored% = 8 And InstType$ = "JEOL" Then MiscIsInstrumentStageStored = True      ' not used
If InterfaceTypeStored% = 9 And InstType$ = "JEOL" Then MiscIsInstrumentStageStored = True      ' not used
If InterfaceTypeStored% = 10 And InstType$ = "JEOL" Then MiscIsInstrumentStageStored = True     ' assume JEOL (mm) for Axioscope
If InterfaceTypeStored% = 11 And InstType$ = "JEOL" Then MiscIsInstrumentStageStored = True     ' not used

Exit Function

' Errors
MiscIsInstrumentStageStoredError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscIsInstrumentStageStored"
ierror = True
Exit Function

MiscIsInstrumentStageStoredInvalidInstType:
msg$ = "Invalid instrument stage type was passed (" & InstType$ & "). Please contact Probe Software Technical Support"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscIsInstrumentStageStored"
ierror = True
Exit Function

End Function

Function MiscRealGetDotVersion(tVersion As Single) As String
' Returns the "." version struing for the passed version number

ierror = False
On Error GoTo MiscRealGetDotVersionError

Dim tstring As String, astring As String, bstring As String, cstring As String
Dim achar As String

' Find decimal point or comma (language issues)
If InStr(tVersion!, ".") > 0 Then achar$ = "."
If InStr(tVersion!, ",") > 0 Then achar$ = ","

' If "dots" found, modify for language
If achar$ <> vbNullString Then
tstring$ = Format$(tVersion!)
Call MiscParseStringToStringA(tstring$, achar$, bstring$)
If ierror Then Exit Function

' No dot or comma, might be a whole number version
Else
MiscRealGetDotVersion$ = Format$(tVersion!)
Exit Function
End If

astring$ = bstring$

' Get next digit
If Len(tstring$) > 0 Then
bstring$ = Left$(tstring$, 1)
End If

' Get last digit
If Len(tstring$) > 1 Then
cstring$ = Mid$(tstring$, 2)
End If

' Make string
MiscRealGetDotVersion$ = astring$ & "." & bstring$ & "." & cstring$
Exit Function

' Errors
MiscRealGetDotVersionError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscRealGetDotVersion"
ierror = True
Exit Function

End Function

