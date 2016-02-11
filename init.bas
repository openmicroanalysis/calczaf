Attribute VB_Name = "CodeINIT"
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

Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetModuleFileName Lib "Kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Global ScalersFile As String
Global MotorsFile As String
Global ChargesFile As String

Global Const OldInstallationFolder$ = "C:\Probe Software\Probe for EPMA"
Global Const PFE_SourceCodeFolder$ = "C:\Source\Probewin32-E"

' Modules level only
Dim DensityFile As String
Dim DensityFile2 As String

Dim tScalBaseLines(1 To MAXSPEC%) As Single    ' for backward compatibility
Dim tScalWindows(1 To MAXSPEC%) As Single
Dim tScalGains(1 To MAXSPEC%) As Single
Dim tScalBiases(1 To MAXSPEC%) As Single

Dim tScalInteDiffModes(1 To MAXSPEC%) As Integer
Dim tScalDeadtimes(1 To MAXSPEC%) As Single

Sub InitElements()
' Reads the ELEMENTS.DAT file (element defaults)

ierror = False
On Error GoTo InitElementsError

Dim i As Integer, linecount As Integer
Dim ip As Integer

linecount% = 1
For i% = 1 To MAXELM%
Input #Temp1FileNumber%, AllAtomicNums%(i%), Symlo$(i%), Symup$(i%), Deflin$(i%), Defcry$(i%), AllAtomicWts!(i%), AllCat%(i%), AllOxd%(i%)

' Check for valid atomic number
If AllAtomicNums%(i%) < 1 Or AllAtomicNums%(i%) > MAXELM% Then GoTo InitElementsInvalidData

' Check that "Symlo" and "SymUp" match
If Not MiscStringsAreSame(Symlo$(i%), Symup$(i%)) Then GoTo InitElementsInvalidData

' Check for valid default xray
ip% = IPOS1(MAXRAY% - 1, Deflin$(i%), Xraylo$())
If ip% = 0 Then GoTo InitElementsInvalidData

' Check for valid default crystal
ip% = IPOS1(MAXCRYSTYPE%, Defcry$(i%), AllCrystalNames$())
If ip% = 0 Then GoTo InitElementsInvalidDataCrystal

If AllAtomicWts!(i%) < 1# Or AllAtomicWts!(i%) > 254# Then GoTo InitElementsInvalidData
If AllCat%(i%) < 1 Or AllCat%(i%) > 9 Then GoTo InitElementsInvalidData
If AllOxd%(i%) < 0 Or AllOxd%(i%) > 9 Then GoTo InitElementsInvalidData

linecount% = linecount% + 1
Next i%

Exit Sub

' Errors
InitElementsError:
MsgBox Error$, vbOKOnly + vbCritical, "InitElements"
ierror = True
Exit Sub

InitElementsInvalidData:
msg$ = "Invalid element data in " & ElementsFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitElements"
ierror = True
Exit Sub

InitElementsInvalidDataCrystal:
msg$ = "Invalid default crystal data in " & ElementsFile$ & " on line " & Str$(linecount%) & ". Please check the " & CrystalsFile$ & " file for the specified crystal."
MsgBox msg$, vbOKOnly + vbExclamation, "InitElements"
ierror = True
Exit Sub

End Sub

Sub InitMotors()
' Reads the MOTORS.DAT file for microprobe motor configuration

ierror = False
On Error GoTo InitMotorsError

Dim comment As String
Dim i As Integer, linecount As Integer
Dim tmsg As String

If DebugMode Then
Call IOWriteLog(vbCrLf & vbCrLf & "Motors Configuration Information:")
End If

' Load motor labels
linecount% = 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs% + NumberOfStageMotors%
Input #Temp1FileNumber%, MotLabels$(i%)
msg$ = msg$ & Format$(MotLabels$(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then
Call IOWriteLog(msg$ & Space$(2) & comment$)
End If

' Load low and high limits
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs% + NumberOfStageMotors%
Input #Temp1FileNumber%, MotLoLimits!(i%)
msg$ = msg$ & Format$(MotLoLimits!(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs% + NumberOfStageMotors%
Input #Temp1FileNumber%, MotHiLimits!(i%)
msg$ = msg$ & Format$(MotHiLimits!(i%), a80$)
If MotLoLimits!(i%) >= MotHiLimits!(i%) Then GoTo InitMotorsLowGreaterThanOrEqualHigh
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, msg$
If DebugMode Then Call IOWriteLog(msg$)

linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs% + NumberOfStageMotors%
Input #Temp1FileNumber%, MotUnitsToAngstromMicrons!(i%)
msg$ = msg$ & MiscAutoFormat$(MotUnitsToAngstromMicrons!(i%))
If MotUnitsToAngstromMicrons!(i%) = 0# Then GoTo InitMotorsInvalidData
If i% > NumberOfTunableSpecs% Then
If MotUnitsToAngstromMicrons!(i%) <> 1# And MotUnitsToAngstromMicrons!(i%) <> 100# And MotUnitsToAngstromMicrons!(i%) <> 1000# Then GoTo InitMotorsInvalidDataUnitMicrons
End If
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, msg$
If DebugMode Then Call IOWriteLog(msg$)

' Load backlash factors
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs% + NumberOfStageMotors%
Input #Temp1FileNumber%, MotBacklashFactors!(i%)
msg$ = msg$ & Format$(MotBacklashFactors!(i%), a80$)
If i% <> WMotor% And Abs(MotBacklashFactors!(i%)) < 10# Then GoTo InitMotorsInvalidData
If InterfaceType% = 2 And MotBacklashFactors!(i%) > 0# Then GoTo InitMotorsJeolPositiveBacklash
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, msg$
If DebugMode Then Call IOWriteLog(msg$)

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, msg$
If DebugMode Then Call IOWriteLog(msg$)

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, msg$
If DebugMode Then Call IOWriteLog(msg$)

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, msg$
If DebugMode Then Call IOWriteLog(msg$)

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, msg$
If DebugMode Then Call IOWriteLog(msg$)

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, msg$
If DebugMode Then Call IOWriteLog(msg$)

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, msg$
If DebugMode Then Call IOWriteLog(msg$)

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, msg$
If DebugMode Then Call IOWriteLog(msg$)

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, msg$
If DebugMode Then Call IOWriteLog(msg$)

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, msg$
If DebugMode Then Call IOWriteLog(msg$)

' Load backlash tolerances (for smart backlash)
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs% + NumberOfStageMotors%
Input #Temp1FileNumber%, MotBacklashTolerances!(i%)
If MotBacklashTolerances!(i%) = 0# Then MotBacklashTolerances!(i%) = 0.002
msg$ = msg$ & Format$(MotBacklashTolerances!(i%), a80$)
If i% <> WMotor% And Abs(MotBacklashTolerances!(i%)) > 0.1 Then GoTo InitMotorsInvalidData
If i% <> WMotor% And Abs(MotBacklashTolerances!(i%)) < 0.00001 Then GoTo InitMotorsInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load motor park positions
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs% + NumberOfStageMotors%
Input #Temp1FileNumber%, MotParkPositions!(i%)
If MotParkPositions!(i%) = 0# And i% <= NumberOfTunableSpecs% Then MotParkPositions!(i%) = MotHiLimits!(i%)
If MotParkPositions!(i%) = 0# And i% > NumberOfTunableSpecs% Then
If InterfaceType% = 5 Then       ' SX100/SXFive
MotParkPositions!(i%) = MotParkPositions!(i%)       ' zero OK for Cameca stage park positions
Else
MotParkPositions!(i%) = MotLoLimits!(i%) + (MotHiLimits!(i%) - MotLoLimits!(i%)) / 2#
End If
End If
msg$ = msg$ & Format$(MotParkPositions!(i%), a80$)
If i% <> WMotor% And MotParkPositions!(i%) > MotHiLimits!(i%) Then GoTo InitMotorsInvalidData
If i% <> WMotor% And MotParkPositions!(i%) < MotLoLimits!(i%) Then GoTo InitMotorsInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load JEOL motor velocities (1/100th micrometers per second)
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs% + NumberOfStageMotors%
Input #Temp1FileNumber%, JEOLVelocity&(i%)
If JEOLVelocity&(i%) = 0 And i% <= NumberOfTunableSpecs% Then JEOLVelocity&(i%) = 500000
If JEOLVelocity&(i%) = 0 And i% > NumberOfTunableSpecs% Then JEOLVelocity&(i%) = 400000
msg$ = msg$ & Format$(JEOLVelocity&(i%), a80$)
If i% <> WMotor% And JEOLVelocity&(i%) < 2000 Then GoTo InitMotorsInvalidSpeed
If i% <> WMotor% And JEOLVelocity&(i%) > 500000 Then GoTo InitMotorsInvalidSpeed
If InterfaceType% = 2 And JeolEOSInterfaceType& = 2 Then        ' 8900 is limited to 4mm/sec
If i% <> WMotor% And JEOLVelocity&(i%) > 400000 Then GoTo InitMotorsInvalidSpeed
End If
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load JEOL backlash (+/- 1/100th micrometers)
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs% + NumberOfStageMotors%
Input #Temp1FileNumber%, JEOLBacklash&(i%)      ' allow zero JEOLBacklash values (P. Carpenter)
'If JEOLBacklash&(i%) = 0 And i% <= NumberOfTunableSpecs% Then JEOLBacklash&(i%) = 50000     ' 500 um
'If JEOLBacklash&(i%) = 0 And i% = XMotor% Then JEOLBacklash&(i%) = -5000   ' -50 um
'If JEOLBacklash&(i%) = 0 And i% = YMotor% Then JEOLBacklash&(i%) = -5000   ' -50 um
'If JEOLBacklash&(i%) = 0 And i% = ZMotor% Then JEOLBacklash&(i%) = 5000    '  50 um
msg$ = msg$ & Format$(JEOLBacklash&(i%), a80$)
If i% <= NumberOfTunableSpecs% And JEOLBacklash&(i%) < 0 Then GoTo InitMotorsInvalidData
If i% = XMotor% And JEOLBacklash&(i%) > 0 Then GoTo InitMotorsInvalidData
If i% = YMotor% And JEOLBacklash&(i%) > 0 Then GoTo InitMotorsInvalidData
If i% = ZMotor% And JEOLBacklash&(i%) < 0 Then GoTo InitMotorsInvalidData
If i% <> WMotor% And Abs(JEOLBacklash&(i%)) > 100000 Then GoTo InitMotorsInvalidData
If i% <> WMotor% And InterfaceType% = 2 And JEOLBacklash&(i%) <> 0 And Abs(JEOLBacklash&(i%)) < 100 Then
tmsg$ = "Warning: JEOL backlash value for motor " & Format$(i%) & " (line " & Format$(linecount%) & ") is too small in " & MotorsFile$
Call IOWriteLogRichText(tmsg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
End If
If i% <> WMotor% And InterfaceType% = 2 And JEOLBacklash&(i%) = 0 Then
tmsg$ = "Warning: JEOL backlash value for motor " & Format$(i%) & " (line " & Format$(linecount%) & ") is zero in " & MotorsFile$
Call IOWriteLogRichText(tmsg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
End If
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load SX100 motor velocities (SX100 Z speed range is between 10 and 200, SXFive Z speed range is between 1 and 10).
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs% + NumberOfStageMotors%
Input #Temp1FileNumber%, SX100Velocity&(i%)
If SX100Velocity&(i%) = 0 And i% <= NumberOfTunableSpecs% Then SX100Velocity&(i%) = 3000    ' in steps/sec spectrometers
If SX100Velocity&(i%) = 0 And i% = XMotor% Then SX100Velocity&(i%) = 10000     ' in steps/sec stage
If SX100Velocity&(i%) = 0 And i% = YMotor% Then SX100Velocity&(i%) = 10000     ' in steps/sec stage
If SX100Velocity&(i%) = 0 And i% = ZMotor% Then SX100Velocity&(i%) = 50     ' in steps/sec stage
msg$ = msg$ & Format$(SX100Velocity&(i%), a80$)
If i% <> WMotor% And i% <= NumberOfTunableSpecs% And SX100Velocity&(i%) < 1000 Then GoTo InitMotorsInvalidSpeed
If i% <> WMotor% And i% <= NumberOfTunableSpecs% And SX100Velocity&(i%) > 4000 Then GoTo InitMotorsInvalidSpeed
If i% <> WMotor% And i% = XMotor% And SX100Velocity&(i%) < 500 Then GoTo InitMotorsInvalidSpeed    ' change to 500 on 07/04/2010
If i% <> WMotor% And i% = XMotor% And SX100Velocity&(i%) > 15000 Then GoTo InitMotorsInvalidSpeed
If i% <> WMotor% And i% = YMotor% And SX100Velocity&(i%) < 500 Then GoTo InitMotorsInvalidSpeed    ' change to 1000 on 07/04/2010
If i% <> WMotor% And i% = YMotor% And SX100Velocity&(i%) > 15000 Then GoTo InitMotorsInvalidSpeed
If i% <> WMotor% And i% = ZMotor% And SX100Velocity&(i%) < 1 Then GoTo InitMotorsInvalidSpeed
If i% <> WMotor% And i% = ZMotor% And SX100Velocity&(i%) > 200 Then GoTo InitMotorsInvalidSpeed
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load SX100 minimum speeds (only used by ROM spectro scanning)
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs% + NumberOfStageMotors%
Input #Temp1FileNumber%, SX100MinimumSpeeds&(i%)
If SX100MinimumSpeeds&(i%) = 0 And i% <= NumberOfTunableSpecs% Then SX100MinimumSpeeds&(i%) = 10    ' in steps/sec spectrometers
If SX100MinimumSpeeds&(i%) = 0 And i% = XMotor% Then SX100MinimumSpeeds&(i%) = 100     ' in steps/sec stage
If SX100MinimumSpeeds&(i%) = 0 And i% = YMotor% Then SX100MinimumSpeeds&(i%) = 100     ' in steps/sec stage
If SX100MinimumSpeeds&(i%) = 0 And i% = ZMotor% Then SX100MinimumSpeeds&(i%) = 5       ' in steps/sec stage
msg$ = msg$ & Format$(SX100MinimumSpeeds&(i%), a80$)
If i% <> WMotor% And i% <= NumberOfTunableSpecs% And SX100MinimumSpeeds&(i%) < 2 Then GoTo InitMotorsInvalidSpeed
If i% <> WMotor% And i% <= NumberOfTunableSpecs% And SX100MinimumSpeeds&(i%) > 1000 Then GoTo InitMotorsInvalidSpeed
If i% <> WMotor% And i% = XMotor% And SX100MinimumSpeeds&(i%) < 10 Then GoTo InitMotorsInvalidSpeed
If i% <> WMotor% And i% = XMotor% And SX100MinimumSpeeds&(i%) > 1500 Then GoTo InitMotorsInvalidSpeed
If i% <> WMotor% And i% = YMotor% And SX100MinimumSpeeds&(i%) < 10 Then GoTo InitMotorsInvalidSpeed
If i% <> WMotor% And i% = YMotor% And SX100MinimumSpeeds&(i%) > 1500 Then GoTo InitMotorsInvalidSpeed
If i% <> WMotor% And i% = ZMotor% And SX100MinimumSpeeds&(i%) < 2 Then GoTo InitMotorsInvalidSpeed
If i% <> WMotor% And i% = ZMotor% And SX100MinimumSpeeds&(i%) > 100 Then GoTo InitMotorsInvalidSpeed
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

Exit Sub

' Errors
InitMotorsError:
MsgBox Error$, vbOKOnly + vbCritical, "InitMotors"
ierror = True
Exit Sub

InitMotorsInvalidData:
msg$ = "Invalid motor data in " & MotorsFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitMotors"
ierror = True
Exit Sub

InitMotorsLowGreaterThanOrEqualHigh:
msg$ = "Low is greater than or equal to high in " & MotorsFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitMotors"
ierror = True
Exit Sub

InitMotorsInvalidDataUnitMicrons:
msg$ = "Stage Units To Microns parameters in " & MotorsFile$ & " on line " & Str$(linecount%) & " must be 1, 100 or 1000"
MsgBox msg$, vbOKOnly + vbExclamation, "InitMotors"
ierror = True
Exit Sub

InitMotorsJeolPositiveBacklash:
msg$ = "Only negative backlash factors are allowed for JEOL 8900/8200/8500/8230/8530 in " & MotorsFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitMotors"
ierror = True
Exit Sub

InitMotorsInvalidSpeed:
msg$ = "Invalid velocity data in " & MotorsFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitMotors"
ierror = True
Exit Sub

End Sub

Sub InitMotors2(mode As Integer)
' Reads the MOTORS.DAT file for specific parameters
' mode = line number
' e.g., mode% = 5 read stage to micron conversion factors

ierror = False
On Error GoTo InitMotors2Error

Dim comment As String, astring As String
Dim i As Integer, linecount As Integer
Dim tmsg As String

' Check for valid line#
If mode% = 0 Then Exit Sub

' Load motor labels
linecount% = 1
Line Input #Temp1FileNumber%, astring$
If DebugMode Then Call IOWriteLog(astring$)
If mode% = 1 Then Exit Sub

' Load low and high limits
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, astring$
If DebugMode Then Call IOWriteLog(astring$)
If mode% = 2 Then Exit Sub

linecount% = linecount% + 1
Line Input #Temp1FileNumber%, astring$
If DebugMode Then Call IOWriteLog(astring$)
If mode% = 3 Then Exit Sub

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, astring$
If DebugMode Then Call IOWriteLog(astring$)
If mode% = 4 Then Exit Sub

' Stage to micron conversion
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs% + NumberOfStageMotors%
Input #Temp1FileNumber%, MotUnitsToAngstromMicrons!(i%)
msg$ = msg$ & MiscAutoFormat$(MotUnitsToAngstromMicrons!(i%))
If MotUnitsToAngstromMicrons!(i%) = 0# Then GoTo InitMotors2InvalidData
If i% > NumberOfTunableSpecs% Then
If MotUnitsToAngstromMicrons!(i%) <> 1# And MotUnitsToAngstromMicrons!(i%) <> 1000# Then GoTo InitMotors2InvalidDataUnitMicrons
End If
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)
If mode% = 5 Then Exit Sub

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, astring$
If DebugMode Then Call IOWriteLog(astring$)
If mode% = 6 Then Exit Sub

' Load backlash factors
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, astring$
If DebugMode Then Call IOWriteLog(astring$)
If mode% = 7 Then Exit Sub

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, astring$
If DebugMode Then Call IOWriteLog(astring$)
If mode% = 8 Then Exit Sub

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, astring$
If DebugMode Then Call IOWriteLog(astring$)
If mode% = 9 Then Exit Sub

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, astring$
If DebugMode Then Call IOWriteLog(astring$)
If mode% = 10 Then Exit Sub

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, astring$
If DebugMode Then Call IOWriteLog(astring$)
If mode% = 11 Then Exit Sub

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, astring$
If DebugMode Then Call IOWriteLog(astring$)
If mode% = 12 Then Exit Sub

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, astring$
If DebugMode Then Call IOWriteLog(astring$)
If mode% = 13 Then Exit Sub

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, astring$
If DebugMode Then Call IOWriteLog(astring$)
If mode% = 14 Then Exit Sub

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, astring$
If DebugMode Then Call IOWriteLog(astring$)
If mode% = 15 Then Exit Sub

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, astring$
If DebugMode Then Call IOWriteLog(astring$)
If mode% = 16 Then Exit Sub

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, astring$
If DebugMode Then Call IOWriteLog(astring$)
If mode% = 17 Then Exit Sub

' Load backlash tolerances (for smart backlash)
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs% + NumberOfStageMotors%
Input #Temp1FileNumber%, MotBacklashTolerances!(i%)
If MotBacklashTolerances!(i%) = 0# Then MotBacklashTolerances!(i%) = 0.002
msg$ = msg$ & Format$(MotBacklashTolerances!(i%), a80$)
If i% <> WMotor% And Abs(MotBacklashTolerances!(i%)) > 0.1 Then GoTo InitMotors2InvalidData
If i% <> WMotor% And Abs(MotBacklashTolerances!(i%)) < 0.00001 Then GoTo InitMotors2InvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)
If mode% = 18 Then Exit Sub

' Load motor park positions
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs% + NumberOfStageMotors%
Input #Temp1FileNumber%, MotParkPositions!(i%)
If MotParkPositions!(i%) = 0# And i% <= NumberOfTunableSpecs% Then MotParkPositions!(i%) = MotHiLimits!(i%)
If MotParkPositions!(i%) = 0# And i% > NumberOfTunableSpecs% Then
If InterfaceType% = 5 Then       ' SX100/SXFive
MotParkPositions!(i%) = MotParkPositions!(i%)       ' zero OK for Cameca stage park positions
Else
MotParkPositions!(i%) = MotLoLimits!(i%) + (MotHiLimits!(i%) - MotLoLimits!(i%)) / 2#
End If
End If
msg$ = msg$ & Format$(MotParkPositions!(i%), a80$)
If i% <> WMotor% And MotParkPositions!(i%) > MotHiLimits!(i%) Then GoTo InitMotors2InvalidData
If i% <> WMotor% And MotParkPositions!(i%) < MotLoLimits!(i%) Then GoTo InitMotors2InvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)
If mode% = 19 Then Exit Sub

' Load JEOL motor velocities (1/100th micrometers per second)
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs% + NumberOfStageMotors%
Input #Temp1FileNumber%, JEOLVelocity&(i%)
If JEOLVelocity&(i%) = 0 And i% <= NumberOfTunableSpecs% Then JEOLVelocity&(i%) = 500000
If JEOLVelocity&(i%) = 0 And i% > NumberOfTunableSpecs% Then JEOLVelocity&(i%) = 400000
msg$ = msg$ & Format$(JEOLVelocity&(i%), a80$)
If i% <> WMotor% And JEOLVelocity&(i%) < 2000 Then GoTo InitMotors2InvalidSpeed
If i% <> WMotor% And JEOLVelocity&(i%) > 500000 Then GoTo InitMotors2InvalidSpeed
If InterfaceType% = 2 And JeolEOSInterfaceType& = 2 Then        ' 8900 is limited to 4mm/sec
If i% <> WMotor% And JEOLVelocity&(i%) > 400000 Then GoTo InitMotors2InvalidSpeed
End If
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)
If mode% = 20 Then Exit Sub

' Load JEOL backlash (+/- 1/100th micrometers)
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs% + NumberOfStageMotors%
Input #Temp1FileNumber%, JEOLBacklash&(i%)      ' allow zero JEOLBacklash values (P. Carpenter)
'If JEOLBacklash&(i%) = 0 And i% <= NumberOfTunableSpecs% Then JEOLBacklash&(i%) = 50000     ' 500 um
'If JEOLBacklash&(i%) = 0 And i% = XMotor% Then JEOLBacklash&(i%) = -5000   ' -50 um
'If JEOLBacklash&(i%) = 0 And i% = YMotor% Then JEOLBacklash&(i%) = -5000   ' -50 um
'If JEOLBacklash&(i%) = 0 And i% = ZMotor% Then JEOLBacklash&(i%) = 5000    '  50 um
msg$ = msg$ & Format$(JEOLBacklash&(i%), a80$)
If i% <= NumberOfTunableSpecs% And JEOLBacklash&(i%) < 0 Then GoTo InitMotors2InvalidData
If i% = XMotor% And JEOLBacklash&(i%) > 0 Then GoTo InitMotors2InvalidData
If i% = YMotor% And JEOLBacklash&(i%) > 0 Then GoTo InitMotors2InvalidData
If i% = ZMotor% And JEOLBacklash&(i%) < 0 Then GoTo InitMotors2InvalidData
If i% <> WMotor% And Abs(JEOLBacklash&(i%)) > 100000 Then GoTo InitMotors2InvalidData
If i% <> WMotor% And InterfaceType% = 2 And JEOLBacklash&(i%) <> 0 And Abs(JEOLBacklash&(i%)) < 100 Then
tmsg$ = "Warning: JEOL backlash value for motor " & Format$(i%) & " (line " & Format$(linecount%) & ") is too small in " & MotorsFile$
Call IOWriteLogRichText(tmsg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
End If
If i% <> WMotor% And InterfaceType% = 2 And JEOLBacklash&(i%) = 0 Then
tmsg$ = "Warning: JEOL backlash value for motor " & Format$(i%) & " (line " & Format$(linecount%) & ") is zero in " & MotorsFile$
Call IOWriteLogRichText(tmsg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
End If
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)
If mode% = 21 Then Exit Sub

' Load SX100 motor velocities (SX100 Z speed range is between 10 and 200, SXFive Z speed range is between 1 and 10).
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs% + NumberOfStageMotors%
Input #Temp1FileNumber%, SX100Velocity&(i%)
If SX100Velocity&(i%) = 0 And i% <= NumberOfTunableSpecs% Then SX100Velocity&(i%) = 3000    ' in steps/sec spectrometers
If SX100Velocity&(i%) = 0 And i% = XMotor% Then SX100Velocity&(i%) = 10000     ' in steps/sec stage
If SX100Velocity&(i%) = 0 And i% = YMotor% Then SX100Velocity&(i%) = 10000     ' in steps/sec stage
If SX100Velocity&(i%) = 0 And i% = ZMotor% Then SX100Velocity&(i%) = 50     ' in steps/sec stage
msg$ = msg$ & Format$(SX100Velocity&(i%), a80$)
If i% <> WMotor% And i% <= NumberOfTunableSpecs% And SX100Velocity&(i%) < 1000 Then GoTo InitMotors2InvalidSpeed
If i% <> WMotor% And i% <= NumberOfTunableSpecs% And SX100Velocity&(i%) > 4000 Then GoTo InitMotors2InvalidSpeed
If i% <> WMotor% And i% = XMotor% And SX100Velocity&(i%) < 500 Then GoTo InitMotors2InvalidSpeed    ' change to 500 on 07/04/2010
If i% <> WMotor% And i% = XMotor% And SX100Velocity&(i%) > 15000 Then GoTo InitMotors2InvalidSpeed
If i% <> WMotor% And i% = YMotor% And SX100Velocity&(i%) < 500 Then GoTo InitMotors2InvalidSpeed    ' change to 1000 on 07/04/2010
If i% <> WMotor% And i% = YMotor% And SX100Velocity&(i%) > 15000 Then GoTo InitMotors2InvalidSpeed
If i% <> WMotor% And i% = ZMotor% And SX100Velocity&(i%) < 1 Then GoTo InitMotors2InvalidSpeed
If i% <> WMotor% And i% = ZMotor% And SX100Velocity&(i%) > 200 Then GoTo InitMotors2InvalidSpeed
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)
If mode% = 22 Then Exit Sub

' Load SX100 minimum speeds (only used by ROM spectro scanning)
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs% + NumberOfStageMotors%
Input #Temp1FileNumber%, SX100MinimumSpeeds&(i%)
If SX100MinimumSpeeds&(i%) = 0 And i% <= NumberOfTunableSpecs% Then SX100MinimumSpeeds&(i%) = 10    ' in steps/sec spectrometers
If SX100MinimumSpeeds&(i%) = 0 And i% = XMotor% Then SX100MinimumSpeeds&(i%) = 100     ' in steps/sec stage
If SX100MinimumSpeeds&(i%) = 0 And i% = YMotor% Then SX100MinimumSpeeds&(i%) = 100     ' in steps/sec stage
If SX100MinimumSpeeds&(i%) = 0 And i% = ZMotor% Then SX100MinimumSpeeds&(i%) = 5       ' in steps/sec stage
msg$ = msg$ & Format$(SX100MinimumSpeeds&(i%), a80$)
If i% <> WMotor% And i% <= NumberOfTunableSpecs% And SX100MinimumSpeeds&(i%) < 2 Then GoTo InitMotors2InvalidSpeed
If i% <> WMotor% And i% <= NumberOfTunableSpecs% And SX100MinimumSpeeds&(i%) > 1000 Then GoTo InitMotors2InvalidSpeed
If i% <> WMotor% And i% = XMotor% And SX100MinimumSpeeds&(i%) < 10 Then GoTo InitMotors2InvalidSpeed
If i% <> WMotor% And i% = XMotor% And SX100MinimumSpeeds&(i%) > 1500 Then GoTo InitMotors2InvalidSpeed
If i% <> WMotor% And i% = YMotor% And SX100MinimumSpeeds&(i%) < 10 Then GoTo InitMotors2InvalidSpeed
If i% <> WMotor% And i% = YMotor% And SX100MinimumSpeeds&(i%) > 1500 Then GoTo InitMotors2InvalidSpeed
If i% <> WMotor% And i% = ZMotor% And SX100MinimumSpeeds&(i%) < 2 Then GoTo InitMotors2InvalidSpeed
If i% <> WMotor% And i% = ZMotor% And SX100MinimumSpeeds&(i%) > 100 Then GoTo InitMotors2InvalidSpeed
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)
If mode% = 23 Then Exit Sub

Exit Sub

' Errors
InitMotors2Error:
MsgBox Error$, vbOKOnly + vbCritical, "InitMotors2"
ierror = True
Exit Sub

InitMotors2InvalidData:
msg$ = "Invalid motor data in " & MotorsFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitMotors2"
ierror = True
Exit Sub

InitMotors2LowGreaterThanOrEqualHigh:
msg$ = "Low is greater than or equal to high in " & MotorsFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitMotors2"
ierror = True
Exit Sub

InitMotors2InvalidDataUnitMicrons:
msg$ = "Stage Units To Microns parameters in " & MotorsFile$ & " on line " & Str$(linecount%) & " must be 1 (Cameca um) or 1000 (JEOL mm)"
MsgBox msg$, vbOKOnly + vbExclamation, "InitMotors2"
ierror = True
Exit Sub

InitMotors2JeolPositiveBacklash:
msg$ = "Only negative backlash factors are allowed for JEOL 8900/8200/8500/8230/8530 in " & MotorsFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitMotors2"
ierror = True
Exit Sub

InitMotors2InvalidSpeed:
msg$ = "Invalid velocity data in " & MotorsFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitMotors2"
ierror = True
Exit Sub

End Sub

Sub InitScalers()
' Reads the SCALERS.DAT file for microprobe scaler configuration

ierror = False
On Error GoTo InitScalersError

Dim comment As String
Dim i As Integer, j As Integer, ip As Integer
Dim linecount As Integer
Dim temp As Single

If DebugMode Then
Call IOWriteLog(vbCrLf & vbCrLf & "Scalers Configuration Information:")
End If

' Load channel labels
linecount% = 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalLabels$(i%)
msg$ = msg$ & Format$(ScalLabels$(i%), a80$)
RealTimeScalLabels$(i%) = ScalLabels$(i%)  ' set as default
If Not IsNumeric(RealTimeScalLabels$(i%)) Then GoTo InitScalersBadLabel ' need for interpreting motor number control fields in Elements/Cations dialog
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, msg$
If DebugMode Then Call IOWriteLog(msg$)

' Unused
linecount% = linecount% + 1
Line Input #Temp1FileNumber%, msg$
If DebugMode Then Call IOWriteLog(msg$)

' Load crystal flipping flags (0=none,1=any,2=at,3=over,4=below)
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalCrystalFlipFlags%(i%)
msg$ = msg$ & Format$(ScalCrystalFlipFlags%(i%), a80$)

If ScalCrystalFlipFlags%(i%) < 0 Or ScalCrystalFlipFlags%(i%) > 4 Then
msg$ = "Invalid crystal flip flag on line " & Str$(linecount%) & " in " & ScalersFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitScalers"
ierror = True
End
End If

Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load crystal flipping positions
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalCrystalFlipPositions!(i%)
msg$ = msg$ & Format$(ScalCrystalFlipPositions!(i%), a80$)
If InterfaceType% > 0 And ScalCrystalFlipFlags%(i%) > 1 Then  ' only check if real time mode and not (no flipping or flip in any position)
If Not NoMotorPositionLimitsCheckingFlag And Not MiscMotorInBounds(i%, ScalCrystalFlipPositions!(i%)) Then GoTo InitScalersBadFlipPosition
End If
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load number of crystals
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalNumberOfCrystals%(i%)
msg$ = msg$ & Format$(ScalNumberOfCrystals%(i%) & " ", a80$)
If ScalNumberOfCrystals%(i%) < 1 Or ScalNumberOfCrystals%(i%) > MAXCRYS% Then
msg$ = "Invalid number of crystals on spectro " & Str$(i%) & " on line " & Str$(linecount%) & " in " & ScalersFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitScalers"
ierror = True
End
End If

Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load crystal names
For j% = 1 To MAXCRYS%
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalCrystalNames$(j%, i%)
If j% <= ScalNumberOfCrystals%(i%) Then
If Trim$(ScalCrystalNames$(j%, i%)) = vbNullString Then
msg$ = "Crystal name for spectro " & Format$(i%) & " on line " & Str$(linecount%) & " in " & ScalersFile$ & " is blank"
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
ScalCrystalNames$(j%, i%) = " "
End If
End If

msg$ = msg$ & Format$(ScalCrystalNames$(j%, i%), a80$)

' Check crystal name against CRYSTALS.DAT file
If j% <= ScalNumberOfCrystals%(i%) Then
ip% = IPOS1(MAXCRYSTYPE%, ScalCrystalNames$(j%, i%), AllCrystalNames$())
If ip% = 0 Then
msg$ = "Crystal name " & ScalCrystalNames$(j%, i%) & ", on line " & Str$(linecount%) & " in " & ScalersFile$ & " does not match any crystal from " & CrystalsFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitScalers"
ierror = True
End
End If
End If

Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)
Next j%

' Load spectrometer default deadtimes (in Microseconds)  (load for backward compatibility defaults)
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, tScalDeadtimes!(i%)
msg$ = msg$ & Format$(tScalDeadtimes!(i%), a80$)
If tScalDeadtimes!(i%) < 0.1 Then GoTo InitScalersInvalidData
If tScalDeadtimes!(i%) > 10# Then GoTo InitScalersInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load spectrometer default off-peak factors
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalOffPeakFactors!(i%)
msg$ = msg$ & Format$(ScalOffPeakFactors!(i%), a80$)
If ScalOffPeakFactors!(i%) < 20# Then GoTo InitScalersInvalidData
If ScalOffPeakFactors!(i%) > 500# Then GoTo InitScalersInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load spectrometer default wave-scan factors
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalWaveScanSizeFactors!(i%)
msg$ = msg$ & Format$(ScalWaveScanSizeFactors!(i%), a80$)
If ScalWaveScanSizeFactors!(i%) < 20# Then GoTo InitScalersInvalidData
If ScalWaveScanSizeFactors!(i%) > 500# Then GoTo InitScalersInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load spectrometer default peak-scan factors
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalPeakScanSizeFactors!(i%)
msg$ = msg$ & Format$(ScalPeakScanSizeFactors!(i%), a80$)
If ScalPeakScanSizeFactors!(i%) < 20# Then GoTo InitScalersInvalidData
If ScalPeakScanSizeFactors!(i%) > 10000# Then GoTo InitScalersInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load spectrometer default wave-scan steps
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalWaveScanPoints%(i%)
msg$ = msg$ & Format$(ScalWaveScanPoints%(i%), a80$)
If ScalWaveScanPoints%(i%) < 5# Then GoTo InitScalersInvalidData
If ScalWaveScanPoints%(i%) > 500# Then GoTo InitScalersInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load spectrometer default peak-scan steps
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalPeakScanPoints%(i%)
msg$ = msg$ & Format$(ScalPeakScanPoints%(i%), a80$)
If ScalPeakScanPoints%(i%) < 5# Then GoTo InitScalersInvalidData
If ScalPeakScanPoints%(i%) > MAXROMSCAN% Then GoTo InitScalersInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load spectrometer default peaking start sizes
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalLiFPeakingStartSizes!(i%)
msg$ = msg$ & Format$(ScalLiFPeakingStartSizes!(i%), a80$)
temp! = Abs(MotHiLimits!(i%) - MotLoLimits!(i%)) / 6000#
If ScalLiFPeakingStartSizes!(i%) < temp! Then GoTo InitScalersStartSizeTooSmall
If ScalLiFPeakingStartSizes!(i%) > temp! * 100 Then GoTo InitScalersStartSizeTooLarge
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load spectrometer default peaking stop sizes
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalLiFPeakingStopSizes!(i%)
msg$ = msg$ & Format$(ScalLiFPeakingStopSizes!(i%), a80$)
If ScalLiFPeakingStopSizes!(i%) <= 0# Then GoTo InitScalersInvalidData
If ScalLiFPeakingStopSizes!(i%) >= ScalLiFPeakingStartSizes!(i%) Then GoTo InitScalersStopGreaterOrEqualThanStart
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load spectrometer default peaking maximum cycles
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalMaximumPeakAttempts%(i%)
msg$ = msg$ & Format$(ScalMaximumPeakAttempts%(i%), a80$)
If ScalMaximumPeakAttempts%(i%) < 5 Then GoTo InitScalersInvalidData
If ScalMaximumPeakAttempts%(i%) > 50 Then GoTo InitScalersInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load spectrometer default minimum peak to backgrounds
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalMinimumPeakToBackgrounds!(i%)
msg$ = msg$ & Format$(ScalMinimumPeakToBackgrounds!(i%), a80$)
If ScalMinimumPeakToBackgrounds!(i%) < 2 Then GoTo InitScalersInvalidData
If ScalMinimumPeakToBackgrounds!(i%) > 200 Then GoTo InitScalersInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load minimum peaking counts
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalMinimumPeakCounts!(i%)
msg$ = msg$ & Format$(ScalMinimumPeakCounts!(i%), a80$)
If ScalMinimumPeakCounts!(i%) < 10# Then GoTo InitScalersInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load PHA baselines (load for backward compatibility defaults)
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, tScalBaseLines!(i%)
msg$ = msg$ & Format$(tScalBaseLines!(i%), a80$)
If tScalBaseLines!(i%) <= 0# Then GoTo InitScalersInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load PHA windows
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, tScalWindows!(i%)
msg$ = msg$ & Format$(tScalWindows!(i%), a80$)
If tScalWindows!(i%) <= 0# Then GoTo InitScalersInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load PHA gains
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, tScalGains!(i%)
msg$ = msg$ & Format$(tScalGains!(i%), a80$)
If tScalGains!(i%) < MinPHAGainWindow! Then GoTo InitScalersInvalidData
If tScalGains!(i%) > MaxPHAGainWindow! Then GoTo InitScalersInvalidData2
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load PHA biases
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, tScalBiases!(i%)
msg$ = msg$ & Format$(tScalBiases!(i%), a80$)
If tScalBiases!(i%) <= 0# Then GoTo InitScalersInvalidData
If tScalBiases!(i%) <= 900# Then GoTo InitScalersLowBias
If tScalBiases!(i%) > MaxPHABiasWindow! Then GoTo InitScalersInvalidData2
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load PHA scale factors
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalBaseLineScaleFactors!(i%)
msg$ = msg$ & Format$(ScalBaseLineScaleFactors!(i%), a80$)
If ScalBaseLineScaleFactors!(i%) <= 0# Then GoTo InitScalersInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalWindowScaleFactors!(i%)
msg$ = msg$ & Format$(ScalWindowScaleFactors!(i%), a80$)
If ScalWindowScaleFactors!(i%) <= 0# Then GoTo InitScalersInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalGainScaleFactors!(i%)
msg$ = msg$ & Format$(ScalGainScaleFactors!(i%), a80$)
If ScalGainScaleFactors!(i%) <= 0# Then GoTo InitScalersInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalBiasScaleFactors!(i%)
msg$ = msg$ & Format$(ScalBiasScaleFactors!(i%), a80$)
If ScalBiasScaleFactors!(i%) <= 0# Then GoTo InitScalersInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Roland Circle (spectrometer focal circle)
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalRolandCircleMMs!(i%)
msg$ = msg$ & Format$(ScalRolandCircleMMs!(i%), a80$)
If ScalRolandCircleMMs!(i%) <= 0# Then GoTo InitScalersInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Crystal flip delay
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalCrystalFlipDelays!(i%)
msg$ = msg$ & Format$(ScalCrystalFlipDelays!(i%), a80$)
If ScalCrystalFlipDelays!(i%) <> -1 And ScalCrystalFlipDelays!(i%) < 0# Then GoTo InitScalersInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Spectrometer offset warning factor
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalSpecOffsetFactors!(i%)
msg$ = msg$ & Format$(ScalSpecOffsetFactors!(i%), a80$)
If ScalSpecOffsetFactors!(i%) <= 0# Then ScalSpecOffsetFactors!(i%) = 400#
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load spectrometer integer deadtimes for Cameca (in Microseconds)
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalInteDeadTimes%(i%)
msg$ = msg$ & Format$(ScalInteDeadTimes%(i%), a80$)
If ScalInteDeadTimes%(i%) = 0 Then ScalInteDeadTimes%(i%) = CInt(tScalDeadtimes!(i%))   ' just truncate for default if not specified in SCALERS.DAT
If ScalInteDeadTimes%(i%) < 1 Then GoTo InitScalersInvalidData  ' SX100/SXFive requires non-zero integer deadtime
If ScalInteDeadTimes%(i%) > 10 Then GoTo InitScalersInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load bias scan ranges
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalBiasScanLows!(i%)
msg$ = msg$ & Format$(ScalBiasScanLows!(i%), a80$)
If ScalBiasScanLows!(i%) <= 0# Then ScalBiasScanLows!(i%) = 1500#
If ScalBiasScanLows!(i%) > MaxPHABiasWindow! Then GoTo InitScalersInvalidData2
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalBiasScanHighs!(i%)
msg$ = msg$ & Format$(ScalBiasScanHighs!(i%), a80$)
If ScalBiasScanHighs!(i%) <= 0# Then ScalBiasScanHighs!(i%) = 1800#
If ScalBiasScanHighs!(i%) > MaxPHABiasWindow! Then GoTo InitScalersInvalidData2
If ScalBiasScanLows!(i%) = ScalBiasScanHighs!(i%) Then GoTo InitScalersSameLowAndHigh
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Gain scan ranges
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalGainScanLows!(i%)
msg$ = msg$ & Format$(ScalGainScanLows!(i%), a80$)
If (InterfaceType% = 0 And MiscIsInstrumentStage("JEOL")) Or InterfaceType% = 2 Then
If ScalGainScanLows!(i%) <= 0# Then ScalGainScanLows!(i%) = MinPHAGainWindow!
Else
If ScalGainScanLows!(i%) <= 0# Then ScalGainScanLows!(i%) = MaxPHAGainWindow! / 4#
End If
If ScalGainScanLows!(i%) > MaxPHAGainWindow! Then GoTo InitScalersInvalidData2
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalGainScanHighs!(i%)
msg$ = msg$ & Format$(ScalGainScanHighs!(i%), a80$)
If (InterfaceType% = 0 And MiscIsInstrumentStage("JEOL")) Or InterfaceType% = 2 Then
If ScalGainScanHighs!(i%) <= 0# Then ScalGainScanHighs!(i%) = MaxPHAGainWindow!
Else
If ScalGainScanHighs!(i%) <= 0# Then ScalGainScanHighs!(i%) = MaxPHAGainWindow!
End If
If ScalGainScanHighs!(i%) > MaxPHAGainWindow! Then GoTo InitScalersInvalidData2
If ScalGainScanLows!(i%) = ScalGainScanHighs!(i%) Then GoTo InitScalersSameLowAndHigh
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Scan baseline and windows
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalScanBaselines!(i%)
msg$ = msg$ & Format$(ScalScanBaselines!(i%), a80$)
If ScalScanBaselines!(i%) <= 0# Then
If MiscIsInstrumentStage("CAMECA") Then
ScalScanBaselines!(i%) = 2.5    ' Cameca
Else
ScalScanBaselines!(i%) = 4#     ' JEOL
End If
End If
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalScanWindows!(i%)
msg$ = msg$ & Format$(ScalScanWindows!(i%), a80$)
If ScalScanWindows!(i%) <= 0# Then ScalScanWindows!(i%) = 0.1
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' New scaler parameters for up to 6 crystals (baselines)
For j% = 1 To MAXCRYS%
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalBaseLines!(j%, i%)
If j% <= ScalNumberOfCrystals%(i%) And ScalBaseLines!(j%, i%) <= 0# Then ScalBaseLines!(j%, i%) = tScalBaseLines!(i%)      ' for backward compatibility
msg$ = msg$ & Format$(ScalBaseLines!(j%, i%), a80$)
If j% <= ScalNumberOfCrystals%(i%) And ScalBaseLines!(j%, i%) <= 0# Then GoTo InitScalersInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)
Next j%

' New scaler parameters for up to 6 crystals (windows)
For j% = 1 To MAXCRYS%
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalWindows!(j%, i%)
If j% <= ScalNumberOfCrystals%(i%) And ScalWindows!(j%, i%) <= 0# Then ScalWindows!(j%, i%) = tScalWindows!(i%)      ' for backward compatibility
msg$ = msg$ & Format$(ScalWindows!(j%, i%), a80$)
If j% <= ScalNumberOfCrystals%(i%) And ScalWindows!(j%, i%) <= 0# Then GoTo InitScalersInvalidData
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)
Next j%

' New scaler parameters for up to 6 crystals (gains)
For j% = 1 To MAXCRYS%
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalGains!(j%, i%)
If j% <= ScalNumberOfCrystals%(i%) And ScalGains!(j%, i%) <= 0# Then ScalGains!(j%, i%) = tScalGains!(i%)      ' for backward compatibility
msg$ = msg$ & Format$(ScalGains!(j%, i%), a80$)
If j% <= ScalNumberOfCrystals%(i%) And ScalGains!(j%, i%) < MinPHAGainWindow! Then GoTo InitScalersInvalidData
If j% <= ScalNumberOfCrystals%(i%) And ScalGains!(j%, i%) > MaxPHAGainWindow! Then GoTo InitScalersInvalidData2
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)
Next j%

' New scaler parameters for up to 6 crystals (biases)
For j% = 1 To MAXCRYS%
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalBiases!(j%, i%)
If j% <= ScalNumberOfCrystals%(i%) And ScalBiases!(j%, i%) <= 0# Then ScalBiases!(j%, i%) = tScalBiases!(i%)      ' for backward compatibility
msg$ = msg$ & Format$(ScalBiases!(j%, i%), a80$)
If j% <= ScalNumberOfCrystals%(i%) And ScalBiases!(j%, i%) <= 0# Then GoTo InitScalersInvalidData
If j% <= ScalNumberOfCrystals%(i%) And ScalBiases!(j%, i%) <= 900# Then GoTo InitScalersLowBias
If j% <= ScalNumberOfCrystals%(i%) And ScalBiases!(j%, i%) > MaxPHABiasWindow! Then GoTo InitScalersInvalidData2
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)
Next j%

' New intediff modes for up to 6 crystals (no previous values, just load as True or False)
For j% = 1 To MAXCRYS%
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, tScalInteDiffModes%(i%)    ' load temp variable
If tScalInteDiffModes%(i%) = 0 Then
ScalInteDiffModes%(j%, i%) = False
Else
ScalInteDiffModes%(j%, i%) = True
End If
msg$ = msg$ & Format$(ScalInteDiffModes%(j%, i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)
Next j%

' New deadtimes for up to 6 crystals
For j% = 1 To MAXCRYS%
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalDeadTimes!(j%, i%)
If j% <= ScalNumberOfCrystals%(i%) Then
If ScalDeadTimes!(j%, i%) <= 0# Then ScalDeadTimes!(j%, i%) = tScalDeadtimes!(i%)      ' for backward compatibility
If ScalDeadTimes!(j%, i%) < 0.1 Then GoTo InitScalersInvalidData
If ScalDeadTimes!(j%, i%) > 10# Then GoTo InitScalersInvalidData
End If
msg$ = msg$ & Format$(ScalDeadTimes!(j%, i%), a80$)
ScalDeadTimes!(j%, i%) = ScalDeadTimes!(j%, i%) / MSPS! ' convert to seconds
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)
Next j%

' New large area crystal flags for up to 6 crystals
For j% = 1 To MAXCRYS%
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, ScalLargeArea%(j%, i%)
If j% <= ScalNumberOfCrystals%(i%) Then
If ScalLargeArea%(j%, i%) <> 0 And ScalLargeArea%(j%, i%) <> 1 Then GoTo InitScalersInvalidData
End If
msg$ = msg$ & Format$(ScalLargeArea%(j%, i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)
Next j%

Exit Sub

' Errors
InitScalersError:
MsgBox Error$, vbOKOnly + vbCritical, "InitScalers"
ierror = True
Exit Sub

InitScalersBadLabel:
msg$ = "Invalid spectro label for spectro " & Format$(i%) & ", (tunable spectrometers must have numeric labels) in " & ScalersFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitScalers"
ierror = True
Exit Sub

InitScalersBadFlipPosition:
msg$ = "Crystal flip position for spectro " & Format$(i%) & " is out of range in " & ScalersFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitScalers"
ierror = True
Exit Sub

InitScalersInvalidData:
msg$ = "Invalid data value for spectro " & Format$(i%) & " in " & ScalersFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitScalers"
ierror = True
Exit Sub

InitScalersInvalidData2:
msg$ = "Data value for spectro " & Format$(i%) & " in " & ScalersFile$ & " on line " & Str$(linecount%) & " is higher than allowable range"
MsgBox msg$, vbOKOnly + vbExclamation, "InitScalers"
ierror = True
Exit Sub

InitScalersStartSizeTooSmall:
msg$ = "Peaking start size for spectro " & Format$(i%) & " is too small in " & ScalersFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitScalers"
ierror = True
Exit Sub

InitScalersStartSizeTooLarge:
msg$ = "Peaking start size for spectro " & Format$(i%) & " is too large in " & ScalersFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitScalers"
ierror = True
Exit Sub

InitScalersStopGreaterOrEqualThanStart:
msg$ = "Peaking stop size for spectro " & Format$(i%) & " is greater than or equal to peaking start size in " & ScalersFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitScalers"
ierror = True
Exit Sub

InitScalersSameLowAndHigh:
msg$ = "Scan low and high values for spectro " & Format$(i%) & " are equal in " & ScalersFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitScalers"
ierror = True
Exit Sub

InitScalersLowBias:
msg$ = "Low detector bias value for spectro " & Format$(i%) & " in " & ScalersFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitScalers"
ierror = True
Exit Sub

End Sub

Sub InitFiles()
' Routine to load file names (everytime) and to create an Userdata directory (first time installation only).

ierror = False
On Error GoTo InitFilesError

Dim amsg As String
Dim sLanguage As String
Dim sLocaleD As String, sLocaleT As String

Dim tProgramPath As String

' Check language
sLanguage$ = MiscSystemGetLanguage()
If ierror Then Exit Sub

'If InStr(sLanguage$, "English") = 0 Then
'msg$ = "The Windows language is not English therefore program will not function properly unless the numerical format is changed in the Control Panel | Region and Language | Format section."
'MsgBox msg$, vbOKOnly + vbInformation, "InitFiles"
'End
'End If

' Check decimal and thousands separator
Call MiscSystemGetRegionalSettings(LOCALE_SDECIMAL&, sLocaleD$)
If ierror Then Exit Sub
Call MiscSystemGetRegionalSettings(LOCALE_STHOUSAND&, sLocaleT$)
If ierror Then Exit Sub

If InStr(sLocaleD$, ".") = 0 Or InStr(sLocaleT$, ",") = 0 Then
msg$ = "The Windows Regional settings do not specify a period (.) for the decimal point and/or comma (,) for the thousands separator, therefore program will not function properly unless the numerical format is changed in the Control Panel | Region and Language | Formats section." & vbCrLf & vbCrLf
msg$ = msg$ & "Please change the Region and Language | Formats to English (United States, United Kingdom, Australia, etc.) format and try again."
MsgBox msg$, vbOKOnly + vbInformation, "InitFiles"
End
End If

' Load character constants first
amsg$ = "InitFiles started..."
If DebugMode Then Call IOWriteLog(amsg$)
VbSpace$ = ChrW$(32)     ' space character
VbDquote$ = ChrW$(34)    ' double quote character
VbSquote$ = ChrW$(39)    ' single quote character
VbComma$ = ChrW$(44)     ' comma character
VbForwardSlash$ = ChrW$(47)     ' foward slash character

' Load program path (works under VB6 IDE!!!)
ProgramPath$ = app.Path & "\"

' Get the Windows system path
If MiscSystemIsHost64Bit() = False Then
SystemPath$ = InitGetWindowsDirectory$() & "\System32"
If ierror Then End
Else
SystemPath$ = InitGetWindowsDirectory$() & "\SysWOW64"
If ierror Then End
End If

' Get application path (does not work under VB IDE!!!) (returns VB6 IDE path)
ApplicationPath$ = InitGetApplicationPath$()
If ierror Then End

' Get INI file locations
ApplicationCommonAppData$ = IOBrowseGetAppDataFolder$(SpecialFolder_CommonAppData) & "\Probe Software\Probe for EPMA\"      ' all users
If ierror Then End
ApplicationAppData$ = IOBrowseGetAppDataFolder$(SpecialFolder_AppData) & "\Probe Software\Probe for EPMA\"                  ' local user (roaming)
If ierror Then End

' Load special program path for Remote app
If MiscStringsAreSame(app.EXEName, "Remote") Then
If Dir$(SystemPath$ & "\Remote.ini") <> vbNullString Then
ProgramPath$ = InitGetINIData(SystemPath$ & "\Remote.ini", "Software", "ProgramPath", ProgramPath$)
End If
End If

' Load special program path for Matrix app
If MiscStringsAreSame(app.EXEName, "Matrix") Then
If Dir$(SystemPath$ & "\Matrix.ini") <> vbNullString Then
ProgramPath$ = InitGetINIData(SystemPath$ & "\Matrix.ini", "Software", "ProgramPath", ProgramPath$)
End If
End If

' Change working directory to the directory where the application was executed
amsg$ = "Changing path to application folder..."
If DebugMode Then Call IOWriteLog(amsg$)
Call MiscChangePath(app.Path)
If ierror Then Exit Sub
    
' Move existing config/data files from original ProgramPath folder to new ProgramData folder (for existing V10 installations)
Call InitFilesMove(Int(0))
If ierror Then Exit Sub
    
' Load file names (ProgramData folder)
amsg$ = "Loading file names..."
If DebugMode Then Call IOWriteLog(amsg$)
StandardDataFile$ = ApplicationCommonAppData$ & "STANDARD.MDB"
XEdgeFile$ = ApplicationCommonAppData$ & "XEDGE.DAT"

XLineFile$ = ApplicationCommonAppData$ & "XLINE.DAT"
XFlurFile$ = ApplicationCommonAppData$ & "XFLUR.DAT"

XLineFile2$ = ApplicationCommonAppData$ & "XLINE2.DAT"      ' for additional x-ray lines
XFlurFile2$ = ApplicationCommonAppData$ & "XFLUR2.DAT"      ' for additional x-ray lines

AbsorbFile$ = ApplicationCommonAppData$ & "ABSORB.DAT"
XrayDataFile$ = ApplicationCommonAppData$ & "XRAY.MDB"

EmpMACFile$ = ApplicationCommonAppData$ & "EMPMAC.DAT"
EmpAPFFile$ = ApplicationCommonAppData$ & "EMPAPF.DAT"
EmpFACFile$ = ApplicationCommonAppData$ & "EMPFAC.DAT"
EmpPHAFile$ = ApplicationCommonAppData$ & "EMPPHA.DAT"

ElementsFile$ = ApplicationCommonAppData$ & "ELEMENTS.DAT"
CrystalsFile$ = ApplicationCommonAppData$ & "CRYSTALS.DAT"
MotorsFile$ = ApplicationCommonAppData$ & "MOTORS.DAT"
ScalersFile$ = ApplicationCommonAppData$ & "SCALERS.DAT"
ChargesFile$ = ApplicationCommonAppData$ & "CHARGES.DAT"
DensityFile$ = ApplicationCommonAppData$ & "DENSITY.DAT"
DensityFile2$ = ApplicationCommonAppData$ & "DENSITY2.DAT"
DetectorsFile$ = ApplicationCommonAppData$ & "DETECTORS.DAT"

ProbeWinINIFile$ = ApplicationCommonAppData$ & "PROBEWIN.INI"     ' Probe for EPMA INI file (recently opened file menu array is updated)
WindowINIFile$ = ApplicationCommonAppData$ & "WINDOW.INI"         ' user window size/position preferences
UserDataFile$ = ApplicationCommonAppData$ & "USER.MDB"            ' user database
PositionDataFile$ = ApplicationCommonAppData$ & "POSITION.MDB"       ' position sample database

SetupDataFile$ = ApplicationCommonAppData$ & "SETUP.MDB"             ' element setup database
SetupDataFile2$ = ApplicationCommonAppData$ & "SETUP2.MDB"             ' MAN element setup database
SetupDataFile3$ = ApplicationCommonAppData$ & "SETUP3.MDB"             ' interference element setup database
CurrentSetupDataFile$ = SetupDataFile$

ProbeErrorLogFile$ = ApplicationCommonAppData$ & app.EXEName & ".ERR"      ' error log
ProbeElmFile$ = ApplicationCommonAppData$ & "PROBEWIN.ELM"    ' wds element/parameter file
ProbePHAFile$ = ApplicationCommonAppData$ & "PROBEWIN.PHA"    ' wds PHA parameter file
ProbeTextLogFile$ = ApplicationCommonAppData$ & app.EXEName & ".TXT"      ' text log

AFactorDataFile$ = ApplicationCommonAppData$ & "UNTITLED.AFA"
CalibratePeakCenterFiles$(0) = ApplicationCommonAppData$ & "PROBEWIN-KA.CAL"
CalibratePeakCenterFiles$(1) = ApplicationCommonAppData$ & "PROBEWIN-KB.CAL"
CalibratePeakCenterFiles$(2) = ApplicationCommonAppData$ & "PROBEWIN-LA.CAL"
CalibratePeakCenterFiles$(3) = ApplicationCommonAppData$ & "PROBEWIN-LB.CAL"
CalibratePeakCenterFiles$(4) = ApplicationCommonAppData$ & "PROBEWIN-MA.CAL"
CalibratePeakCenterFiles$(5) = ApplicationCommonAppData$ & "PROBEWIN-MB.CAL"

MDB_Template$ = ApplicationCommonAppData$ & "MDB_Database.mdb"      ' MDB template for new MDB databases
MatrixMDBFile$ = ApplicationCommonAppData$ & "Matrix.mdb"           ' penepma matrix correction k-ratio database
PureMDBFile$ = ApplicationCommonAppData$ & "Pure.mdb"               ' penepma pure element k-ratio database
BoundaryMDBFile$ = ApplicationCommonAppData$ & "Boundary.mdb"       ' Penepma boundary fluorescence k-ratio database
CustomMDBFile$ = ApplicationCommonAppData$ & "Custom.mdb"           ' custom composition standard database

OutputDataFile$ = ApplicationCommonAppData$ & app.EXEName$ & ".OUT"

' Load model (base) script files (read only but still...)
GRIDBB_BAS_File$ = ApplicationCommonAppData$ & "GRIDBB.BAS"
GRIDCC_BAS_File$ = ApplicationCommonAppData$ & "GRIDCC.BAS"

SLICEXY_BAS_File$ = ApplicationCommonAppData$ & "SLICEXY.BAS"
POLYXY_BAS_File$ = ApplicationCommonAppData$ & "POLYXY.BAS"
MODALXY_BAS_File$ = ApplicationCommonAppData$ & "MODALXY.BAS"

STRIPXY1_BAS_File$ = ApplicationCommonAppData$ & "STRIPXY1.BAS"     ' equal aspect images
STRIPXY2_BAS_File$ = ApplicationCommonAppData$ & "STRIPXY2.BAS"     ' tall aspect images
STRIPXY3_BAS_File$ = ApplicationCommonAppData$ & "STRIPXY3.BAS"     ' wide aspect images

' Load Tip of the Day file
TipOfTheDayFile$ = ApplicationCommonAppData$ & "TipOfTheDay.txt"

' Load PDF file names (in application folder)
ProbeforEPMAQuickStartGuide$ = ProgramPath$ & "Probe for EPMA_Quick Start.pdf"
ProbeforEPMAFAQ$ = ProgramPath$ & "Probe for EPMA Frequently Asked Questions.pdf"

GettingStartedManual$ = ProgramPath$ & "GettingStarted_ProbeSoftware.pdf"
AdvancedTopicsManual$ = ProgramPath$ & "AdvancedTopics_ProbeSoftware.pdf"
UserReferenceManual$ = ProgramPath$ & "probewin.pdf"

' Load help files
ProbewinHelpFile$ = ProgramPath$ & "Probewin.chm"
CalcImageHelpFile$ = ProgramPath$ & "CalcImage.chm"
RemoteHelpFile$ = SystemPath$ & "\Remote.chm"
MatrixHelpFile$ = SystemPath$ & "\Matrix.chm"

' Set the Help Files depending on executable
amsg$ = "Loading help files..."
If DebugMode Then Call IOWriteLog(amsg$)
If MiscStringsAreSame(app.EXEName, "CalcImage") Then
If Dir$(CalcImageHelpFile$) = vbNullString Then GoTo InitFilesHelpFileNotFoundCalcImage
app.HelpFile = CalcImageHelpFile$

ElseIf MiscStringsAreSame(app.EXEName, "Remote") Then
If Dir$(RemoteHelpFile$) = vbNullString Then GoTo InitFilesHelpFileNotFoundRemote
app.HelpFile = RemoteHelpFile$

ElseIf MiscStringsAreSame(app.EXEName, "Matrix") Then
If Dir$(MatrixHelpFile$) = vbNullString Then GoTo InitFilesHelpFileNotFoundMatrix
app.HelpFile = MatrixHelpFile$

Else
If Dir$(ProbewinHelpFile$) = vbNullString Then GoTo InitFilesHelpFileNotFoundProbewin
app.HelpFile = ProbewinHelpFile$
End If

       
amsg$ = "InitFiles completed"
If DebugMode Then Call IOWriteLog(amsg$)
Exit Sub

' Errors
InitFilesError:
msg$ = Error$ & vbCrLf & vbCrLf
msg$ = msg$ & "Current Process: " & amsg$ & vbCrLf
msg$ = msg$ & "Program Path: " & ProgramPath$ & vbCrLf
msg$ = msg$ & "System Path: " & SystemPath$ & vbCrLf
msg$ = msg$ & "Application Path: " & ApplicationPath$ & vbCrLf
msg$ = msg$ & "Program Data Path: " & ApplicationCommonAppData$ & vbCrLf
msg$ = msg$ & "UserData Directory: " & UserDataDirectory$ & vbCrLf
msg$ = msg$ & "UserImages Directory: " & UserImagesDirectory$ & vbCrLf
msg$ = msg$ & "Standard POS File Directory: " & StandardPOSFileDirectory$ & vbCrLf
msg$ = msg$ & "Column PCC File Directory: " & ColumnPCCFileDirectory$ & vbCrLf
msg$ = msg$ & "CalcZAF Data Directory: " & CalcZAFDATFileDirectory$ & vbCrLf
msg$ = msg$ & "Surfer Data Directory: " & SurferDataDirectory$ & vbCrLf
msg$ = msg$ & "Grapher Data Directory: " & GrapherDataDirectory$ & vbCrLf
msg$ = msg$ & "Demo Images Directory: " & DemoImagesDirectory$ & vbCrLf
MsgBox msg$, vbOKOnly + vbCritical, "InitFiles"
Close (Temp1FileNumber%)
ierror = True
Exit Sub

InitFilesHelpFileNotFoundProbewin:
msg$ = "The specified Help file (" & ProbewinHelpFile$ & ") was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitFiles"
ierror = True
Exit Sub

InitFilesHelpFileNotFoundCalcImage:
msg$ = "The specified Help file (" & CalcImageHelpFile$ & ") was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitFiles"
ierror = True
Exit Sub

InitFilesHelpFileNotFoundRemote:
msg$ = "The specified Help file (" & RemoteHelpFile$ & ") was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitFiles"
ierror = True
Exit Sub

InitFilesHelpFileNotFoundMatrix:
msg$ = "The specified Help file (" & MatrixHelpFile$ & ") was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitFiles"
ierror = True
Exit Sub

End Sub

Sub InitFilesMove(mode As Integer)
' Copy files from previous (v10) installation to ProgramData folder (v11)
' mode = 0  move old files except stage bitmaps
' mode = 1  move only stage bitmaps

ierror = False
On Error GoTo InitFilesMoveError

Dim tProgramPath As String
Dim i As Integer

' Check if app is running from PFE source folder (if so just exit)
If Trim$(UCase$(app.Path)) = Trim$(UCase$(PFE_SourceCodeFolder$)) Then Exit Sub

' Set program path to old installation folder to update current installation
tProgramPath$ = OldInstallationFolder$ & "\"

' Copy all files except stage bit maps
If mode% = 0 Then

' Check for old installation config and data files (copy and delete old config and data files)
If Dir$(tProgramPath$ & "Probewin.ini") <> vbNullString Then
Call InitFilesMove2("PROBEWIN.INI", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("MOTORS.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("SCALERS.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("ELEMENTS.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("CRYSTALS.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("CHARGES.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("DENSITY.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("DENSITY2.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("DETECTORS.DAT", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("STANDARD.MDB", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("XLINE.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("XEDGE.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("XFLUR.DAT", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("LINEMU.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("CITZMU.DAT", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("ABSORB.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("XRAY.MDB", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("EMPMAC.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("EMPAPF.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("EMPFAC.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("EMPPHA.DAT", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("Jump_Ratios.dat", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("LINES.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("LINES2.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("CHANTLER2005-" & Xraylo$(1) & ".dat", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("CHANTLER2005-" & Xraylo$(2) & ".dat", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("CHANTLER2005-" & Xraylo$(3) & ".dat", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("CHANTLER2005-" & Xraylo$(4) & ".dat", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("CHANTLER2005-" & Xraylo$(5) & ".dat", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("CHANTLER2005-" & Xraylo$(6) & ".dat", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("all_weights-modified.txt", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("WINDOW.INI", tProgramPath$, ApplicationCommonAppData$)              ' user window size/position preferences
Call InitFilesMove2("USER.MDB", tProgramPath$, ApplicationCommonAppData$)                ' user database
Call InitFilesMove2("POSITION.MDB", tProgramPath$, ApplicationCommonAppData$)            ' position sample database
  
Call InitFilesMove2("DHZ.MDB", tProgramPath$, ApplicationCommonAppData$)                ' composition match database
Call InitFilesMove2("ORE.MDB", tProgramPath$, ApplicationCommonAppData$)                ' dana composition database
Call InitFilesMove2("SRM.MDB", tProgramPath$, ApplicationCommonAppData$)                ' dana composition database
Call InitFilesMove2("AMCSD.MDB", tProgramPath$, ApplicationCommonAppData$)              ' crystal diffraction (ideal formula) composition database

Call InitFilesMove2("SETUP.MDB", tProgramPath$, ApplicationCommonAppData$)                ' element setup database
Call InitFilesMove2("SETUP2.MDB", tProgramPath$, ApplicationCommonAppData$)               ' MAN element setup database
Call InitFilesMove2("SETUP3.MDB", tProgramPath$, ApplicationCommonAppData$)               ' interference element setup database
   
Call InitFilesMove2("PROBEWIN.ELM", tProgramPath$, ApplicationCommonAppData$)               ' wds element/parameter file
Call InitFilesMove2("PROBEWIN.PHA", tProgramPath$, ApplicationCommonAppData$)               ' wds PHA parameter file
Call InitFilesMove2("Version.txt", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("TipOfTheDay.txt", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2(app.EXEName & ".OUT", tProgramPath$, ApplicationCommonAppData$)     ' output log
Call InitFilesMove2(app.EXEName & ".TXT", tProgramPath$, ApplicationCommonAppData$)     ' text log
Call InitFilesMove2(app.EXEName & ".ERR", tProgramPath$, ApplicationCommonAppData$)     ' error log

Call InitFilesMove2("UNTITLED.AFA", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("PROBEWIN-KA.CAL", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("PROBEWIN-KB.CAL", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("PROBEWIN-LA.CAL", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("PROBEWIN-LB.CAL", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("PROBEWIN-MA.CAL", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("PROBEWIN-MB.CAL", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("Matrix.mdb", tProgramPath$, ApplicationCommonAppData$)            ' penepma matrix correction k-ratio database
Call InitFilesMove2("Boundary.mdb", tProgramPath$, ApplicationCommonAppData$)          ' Penepma boundary fluorescence k-ratio database
Call InitFilesMove2("Custom.mdb", tProgramPath$, ApplicationCommonAppData$)            ' custom composition standard database

Call InitFilesMove2("THERMAL.FC", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("RAINBOW2.FC", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("BLUERED.FC", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("CUSTOM.FC", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("GRIDBB.BAS", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("GRIDCC.BAS", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("SLICEXY.BAS", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("POLYXY.BAS", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("MODALXY.BAS", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("STRIPXY1.BAS", tProgramPath$, ApplicationCommonAppData$)     ' equal aspect images
Call InitFilesMove2("STRIPXY2.BAS", tProgramPath$, ApplicationCommonAppData$)     ' tall aspect images
Call InitFilesMove2("STRIPXY3.BAS", tProgramPath$, ApplicationCommonAppData$)     ' wide aspect images

Call InitFilesMove2("GRIDXY_Custom1.BAS", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("GRIDXY_Custom2.BAS", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("GRIDXY_Custom3.BAS", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("GRIDXY_Custom4.BAS", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("SLICEXY_Custom1.BAS", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("SLICEXY_Custom2.BAS", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("POLYXY_Custom1.BAS", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("POLYXY_Custom2.BAS", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("STRIPXY1_Custom1.BAS", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("STRIPXY2_Custom1.BAS", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("STRIPXY3_Custom1.BAS", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("STRIPXY1_Custom2.BAS", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("STRIPXY2_Custom2.BAS", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("STRIPXY3_Custom2.BAS", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("TRAVXY.BAS", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("TRAVXY_Custom1.BAS", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("TRAVXY_Custom2.BAS", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("TRAVXY_Custom3.BAS", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("TRAVXY_Custom4.BAS", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("COLUMN2.DAT", tProgramPath$, ApplicationCommonAppData$)

' Files that will get moved to UserData folders
Call InitFilesMove2("MODAL.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("CALCZAF.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("CALCZAF2.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("CALCBIN.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("NISTBIN.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("NISTBIN2.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("NISTBIN3.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("POUCHOU.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("POUCHOU2.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("AUAGCU2.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("NISTBINA20.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("NISTBINZ10.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("POUCHOUA20.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("POUCHOUZ10.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("AuCu_NBS-K-ratios.DAT", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("Olivine particle-JTA-1.0um.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("Olivine particle-JTA-0.5um.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("Pouchou2_Au,Cu,Ag_only.dat", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("CaF2_Particles_20 keV.DAT", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("Wark-Watson Exper. Data (CalcZAF format)_JEOL.dat", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("Wark-Watson Exper. Data (CalcZAF format)_Cameca.dat", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("SiO2-TiO2_400um_JEOL.BMP", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("SiO2-TiO2_400um_JEOL.ACQ", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("SiO2-TiO2_400um_Cameca.BMP", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("SiO2-TiO2_400um_Cameca.ACQ", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("XYSCAN2.BAS", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("XYSCAN2.BLN", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("XYSCAN2.DAT", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("XYSLICE2.BAS", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("Montel-1_Quant_Point_Classify.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("Silicates-2_Quant_Image_Classify.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("Silicates-2_Quant_Image_Classify.INI", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("Silicates-2_Quant_Image_Classify.TXT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("Silicates-2_00485_VS1.grd", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("XYTRAV.BAS", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("XYTRAV.DAT", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("XYTRAV2.DAT", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("DEMO2_JEOL.BMP", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("DEMO2_JEOL.JPG", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("DEMO2_JEOL.GIF", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("DEMO2_JEOL.ACQ", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("DEMO2_Cameca.BMP", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("DEMO2_Cameca.JPG", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("DEMO2_Cameca.GIF", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("DEMO2_Cameca.ACQ", tProgramPath$, ApplicationCommonAppData$)

Call InitFilesMove2("DEMO3_JEOL.BMP", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("DEMO3_JEOL.ACQ", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("DEMO3_Cameca.BMP", tProgramPath$, ApplicationCommonAppData$)
Call InitFilesMove2("DEMO3_Cameca.ACQ", tProgramPath$, ApplicationCommonAppData$)

' See special treatment in InitFilesMove2 (until all apps support new JEOL/Cameca drivers)
Call InitFilesMove2("Xline.dll", tProgramPath$, ApplicationCommonAppData$)

' Warn user about moving EIKS files
If InterfaceType% = 2 And JeolEOSInterfaceType& = 3 Then
msg$ = "JEOL EIKS files must be moved from the old Probe for EPMA application folder (usually C:\Probe Software\Probe for EPMA), to the new Probe for EPMA application folder (usually C:\Program Files (x86)\Probe Software\probe for EPMA) manually." & vbCrLf & vbCrLf
msg$ = msg$ & "This warning will only be given once!" & vbCrLf & vbCrLf
msg$ = msg$ & "Please contact Probe Software technical support if you need assistance."
MsgBox msg$, vbOKOnly + vbExclamation, "InitFilesMove"
ierror = True
End If

End If
End If

' Note that stage bitmaps are moved from InitINIStageBitmaps procedure
If mode% = 1 Then

' Check for old installation stage bitmap files (copy and delete old files)
For i% = 1 To StageBitMapCount%
If Dir$(tProgramPath$ & StageBitMapFile$(i%)) <> vbNullString Then
Call InitFilesMove2(StageBitMapFile$(i%), tProgramPath$, ApplicationCommonAppData$)         ' move stage bit maps to ProgramData folder
End If
Next i%

End If

Exit Sub

' Errors
InitFilesMoveError:
MsgBox Error$, vbOKOnly + vbCritical, "InitFilesMove"
ierror = True
Exit Sub

End Sub

Sub InitFilesMove2(tfilename As String, tProgramPath As String, tApplicationCommonAppData As String)
' Move the passed filename from the application folder to the program data folder (for existing installations)

ierror = False
On Error GoTo InitFilesMove2Error
 
' Check if not PFE Source folder
If Trim$(UCase$(ProgramPath$)) = Trim$(UCase$(PFE_SourceCodeFolder$ & "\")) Then Exit Sub

' Check if old file location exists
If Dir$(tProgramPath$ & tfilename$) = vbNullString Then Exit Sub

' Check if target location file already exists (default V11 installation being overwritten by V10 site specific config and data files)
If Dir$(tApplicationCommonAppData & tfilename$) <> vbNullString Then
If DebugMode Then Call IOWriteLog("Deleting default config or data file " & tfilename$ & " from ProgramData folder...")
Kill tApplicationCommonAppData & tfilename$
DoEvents
If Dir$(tApplicationCommonAppData & tfilename$) <> vbNullString Then GoTo InitFilesMove2NotDeletedProgramData
End If

' Now copy old location to new location (if it exists)
If Dir$(tProgramPath$ & tfilename$) <> vbNullString Then
If DebugMode Then Call IOWriteLog("Copying custom config or data file " & tfilename$ & " to ProgramData folder...")
FileCopy tProgramPath$ & tfilename$, tApplicationCommonAppData & tfilename$
DoEvents

' Now double check it was copied
If Dir$(tApplicationCommonAppData & tfilename$) = vbNullString Then GoTo InitFilesMove2NotCopiedProgramData

' File was not found in ProgramPath$ (e.g., Probewin.elm, Probewin.pha)
Else
Call IOWriteLog("Custom config or data file " & tfilename$ & " was not found and therefore not copied to the ProgramData folder...")
End If

' Now delete the original V10 file from the ProgramPath folder (note special exception for Xline.dll until new driver is available)
If Trim$(UCase$(tfilename$)) <> Trim$(UCase$("Xline.dll")) Then
Kill tProgramPath$ & tfilename$
DoEvents

' Now double check it was deleted
If Dir$(tProgramPath$ & tfilename$) <> vbNullString Then GoTo InitFilesMove2NotDeletedProgramPath
End If

Exit Sub

' Errors
InitFilesMove2Error:
MsgBox Error$, vbOKOnly + vbCritical, "InitFilesMove2"
ierror = True
Exit Sub

InitFilesMove2NotDeletedProgramData:
msg$ = "The default config or data file (" & tfilename$ & ") was not properly deleted from the ProgramData folder properly. Please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "InitFilesMove2"
ierror = True
End

InitFilesMove2NotCopiedProgramData:
msg$ = "The custom config or data file (" & tfilename$ & ") was not properly copied to the ProgramData folder properly. Please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "InitFilesMove2"
ierror = True
End

InitFilesMove2NotDeletedProgramPath:
msg$ = "The custom config or data file (" & tfilename$ & ") was not properly deleted from the ProgramPath (app) folder properly. Please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "InitFilesMove2"
ierror = True
End

End Sub

Sub InitCheckFiles()
' Routine to check for files in program directory

ierror = False
On Error GoTo InitCheckFilesError

Dim gridxyfile As String

' Do not check if TestFid.exe
If UCase$(app.EXEName) = UCase$("TestFid") Then Exit Sub

' Check for standard database file
If Dir$(StandardDataFile$) = vbNullString Then
msg$ = "The " & StandardDataFile$ & " file was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
End
End If

' Check for each file
If Dir$(XLineFile$) = vbNullString Then
msg$ = "The " & XLineFile$ & " file was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
End
End If

'If Dir$(XLineFile2$) = vbNullString Then
'msg$ = "The " & XLineFile2$ & " file was not found"
'MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
'End
'End If

If Dir$(XEdgeFile$) = vbNullString Then
msg$ = "The " & XEdgeFile$ & " file was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
End
End If

If Dir$(XFlurFile$) = vbNullString Then
msg$ = "The " & XFlurFile$ & " file was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
End
End If

'If Dir$(XFlurFile2$) = vbNullString Then
'msg$ = "The " & XFlurFile2$ & " file was not found"
'MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
'End
'End If

If Dir$(ElementsFile$) = vbNullString Then
msg$ = "The " & ElementsFile$ & " file was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
End
End If

If Dir$(AbsorbFile$) = vbNullString Then
msg$ = "The " & AbsorbFile$ & " file was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
End
End If

If Dir$(EmpMACFile$) = vbNullString Then
msg$ = "The " & EmpMACFile$ & " file was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
End
End If

If Dir$(EmpAPFFile$) = vbNullString Then
msg$ = "The " & EmpAPFFile$ & " file was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
End
End If

' Make sure both default MAC files are present
If Dir$(ApplicationCommonAppData$ & "LINEMU.DAT") = vbNullString Then
msg$ = "The " & ApplicationCommonAppData$ & "LINEMU.DAT" & " file was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
End
End If

If Dir$(ApplicationCommonAppData$ & "CITZMU.DAT") = vbNullString Then
msg$ = "The " & ApplicationCommonAppData$ & "CITZMU.DAT" & " file was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
End
End If

' Check for EmpFACFile (empirical alpha-factor file)
If Dir$(EmpFACFile$) = vbNullString Then
msg$ = "The " & EmpFACFile$ & " file was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
End
End If

' Check for EmpPHAFile (empirical PHA file)
If Dir$(EmpPHAFile$) = vbNullString Then
msg$ = "The " & EmpPHAFile$ & " file was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
End
End If

' Check for Surfer model scripts
If UCase$(app.EXEName) = UCase$("Probewin") Then
If Dir$(GRIDBB_BAS_File$) = vbNullString Then
msg$ = "The " & GRIDBB_BAS_File$ & " file was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
End
End If
If Dir$(GRIDCC_BAS_File$) = vbNullString Then
msg$ = "The " & GRIDCC_BAS_File$ & " file was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
End
End If
End If

If UCase$(app.EXEName) = UCase$("CalcImage") Then
If Dir$(SLICEXY_BAS_File$) = vbNullString Then
msg$ = "The " & SLICEXY_BAS_File$ & " file was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
End
End If
End If

If UCase$(app.EXEName) = UCase$("CalcImage") Then
If Dir$(POLYXY_BAS_File$) = vbNullString Then
msg$ = "The " & POLYXY_BAS_File$ & " file was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
End
End If
End If

If UCase$(app.EXEName) = UCase$("CalcImage") Then
If Dir$(MODALXY_BAS_File$) = vbNullString Then
msg$ = "The " & MODALXY_BAS_File$ & " file was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
End
End If
End If

If UCase$(app.EXEName) = UCase$("CalcImage") Then
If Dir$(STRIPXY1_BAS_File$) = vbNullString Then
msg$ = "The " & STRIPXY1_BAS_File$ & " file was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
End
End If
End If

If UCase$(app.EXEName) = UCase$("CalcImage") Then
If Dir$(STRIPXY2_BAS_File$) = vbNullString Then
msg$ = "The " & STRIPXY2_BAS_File$ & " file was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
End
End If
End If

If UCase$(app.EXEName) = UCase$("CalcImage") Then
If Dir$(STRIPXY3_BAS_File$) = vbNullString Then
msg$ = "The " & STRIPXY3_BAS_File$ & " file was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitCheckFiles"
End
End If
End If

If UCase$(app.EXEName) = UCase$("CalcImage") Then
gridxyfile$ = ApplicationCommonAppData$ & "GRIDXY_Custom1.BAS"
If Dir$(gridxyfile$) = vbNullString Then
msg$ = "The " & gridxyfile$ & " file was not found in the application data folder. Please contact Probe Software for an updated copy"
Call IOWriteLog(msg$)
End If
gridxyfile$ = ApplicationCommonAppData$ & "GRIDXY_Custom2.BAS"
If Dir$(gridxyfile$) = vbNullString Then
msg$ = "The " & gridxyfile$ & " file was not found in the application data folder. Please contact Probe Software for an updated copy"
Call IOWriteLog(msg$)
End If
gridxyfile$ = ApplicationCommonAppData$ & "GRIDXY_Custom3.BAS"
If Dir$(gridxyfile$) = vbNullString Then
msg$ = "The " & gridxyfile$ & " file was not found in the application data folder. Please contact Probe Software for an updated copy"
Call IOWriteLog(msg$)
End If
gridxyfile$ = ApplicationCommonAppData$ & "GRIDXY_Custom4.BAS"
If Dir$(gridxyfile$) = vbNullString Then
msg$ = "The " & gridxyfile$ & " file was not found in the application data folder. Please contact Probe Software for an updated copy"
Call IOWriteLog(msg$)
End If
gridxyfile$ = ApplicationCommonAppData$ & "SLICEXY_Custom1.BAS"
If Dir$(gridxyfile$) = vbNullString Then
msg$ = "The " & gridxyfile$ & " file was not found in the application data folder. Please contact Probe Software for an updated copy"
Call IOWriteLog(msg$)
End If
gridxyfile$ = ApplicationCommonAppData$ & "SLICEXY_Custom2.BAS"
If Dir$(gridxyfile$) = vbNullString Then
msg$ = "The " & gridxyfile$ & " file was not found in the application data folder. Please contact Probe Software for an updated copy"
Call IOWriteLog(msg$)
End If
gridxyfile$ = ApplicationCommonAppData$ & "POLYXY_Custom1.BAS"
If Dir$(gridxyfile$) = vbNullString Then
msg$ = "The " & gridxyfile$ & " file was not found in the application data folder. Please contact Probe Software for an updated copy"
Call IOWriteLog(msg$)
End If
gridxyfile$ = ApplicationCommonAppData$ & "POLYXY_Custom2.BAS"
If Dir$(gridxyfile$) = vbNullString Then
msg$ = "The " & gridxyfile$ & " file was not found in the application data folder. Please contact Probe Software for an updated copy"
Call IOWriteLog(msg$)
End If
gridxyfile$ = ApplicationCommonAppData$ & "STRIPXY1_Custom1.BAS"
If Dir$(gridxyfile$) = vbNullString Then
msg$ = "The " & gridxyfile$ & " file was not found in the application data folder. Please contact Probe Software for an updated copy"
Call IOWriteLog(msg$)
End If
gridxyfile$ = ApplicationCommonAppData$ & "STRIPXY2_Custom1.BAS"
If Dir$(gridxyfile$) = vbNullString Then
msg$ = "The " & gridxyfile$ & " file was not found in the application data folder. Please contact Probe Software for an updated copy"
Call IOWriteLog(msg$)
End If
gridxyfile$ = ApplicationCommonAppData$ & "STRIPXY3_Custom1.BAS"
If Dir$(gridxyfile$) = vbNullString Then
msg$ = "The " & gridxyfile$ & " file was not found in the application data folder. Please contact Probe Software for an updated copy"
Call IOWriteLog(msg$)
End If
gridxyfile$ = ApplicationCommonAppData$ & "STRIPXY1_Custom2.BAS"
If Dir$(gridxyfile$) = vbNullString Then
msg$ = "The " & gridxyfile$ & " file was not found in the application data folder. Please contact Probe Software for an updated copy"
Call IOWriteLog(msg$)
End If
gridxyfile$ = ApplicationCommonAppData$ & "STRIPXY2_Custom2.BAS"
If Dir$(gridxyfile$) = vbNullString Then
msg$ = "The " & gridxyfile$ & " file was not found in the application data folder. Please contact Probe Software for an updated copy"
Call IOWriteLog(msg$)
End If
gridxyfile$ = ApplicationCommonAppData$ & "STRIPXY3_Custom2.BAS"
If Dir$(gridxyfile$) = vbNullString Then
msg$ = "The " & gridxyfile$ & " file was not found in the application data folder. Please contact Probe Software for an updated copy"
Call IOWriteLog(msg$)
End If
End If

' Check for Tip of the day file
If UCase$(app.EXEName) = UCase$("Probewin") Then
If Dir$(TipOfTheDayFile$) = vbNullString Then
msg$ = "The " & TipOfTheDayFile$ & " file was not found"
Call IOWriteLog(msg$)
End If
End If

' Check for private DLLs




Exit Sub

' Errors
InitCheckFilesError:
MsgBox Error$, vbOKOnly + vbCritical, "InitCheckFiles"
ierror = True
Exit Sub

End Sub

Sub InitCrystals()
' Reads the CRYSTALS.DAT file for defined crystal types

ierror = False
On Error GoTo InitCrystalsError

Dim i As Integer, ip As Integer, linecount As Integer

' Load crystal names, 2ds, k and element and line for peaking
linecount% = 1
i% = 1
Do Until EOF(Temp1FileNumber%) Or i% > MAXCRYSTYPE%

Input #Temp1FileNumber%, AllCrystalNames$(i%), AllCrystal2ds!(i%), AllCrystalKs!(i%), AllCrystalElements$(i%), AllCrystalXrays$(i%)

' If crystal name is not blank, check for valid 2d and k
If Trim$(AllCrystalNames$(i%)) <> vbNullString Then
If AllCrystal2ds!(i%) <= 1# Or AllCrystal2ds!(i%) > 300# Then GoTo InitCrystalsInvalidData
If AllCrystalKs!(i%) < 0# Then GoTo InitCrystalsInvalidData

' Check element symbols
'ip% = IPOS1(MAXELM%, AllCrystalElements$(i%), Symlo$())    ' symlo$() not loaded yet so do not check
'If ip% = 0 Then GoTo InitCrystalsInvalidData

' Check xray symbols
ip% = IPOS1(MAXRAY% - 1, AllCrystalXrays$(i%), Xraylo$())
If ip% = 0 Then GoTo InitCrystalsInvalidData
End If

linecount% = linecount% + 1
i% = i% + 1
Loop

Exit Sub

' Errors
InitCrystalsError:
MsgBox Error$, vbOKOnly + vbCritical, "InitCrystals"
ierror = True
Exit Sub

InitCrystalsInvalidData:
msg$ = "Invalid crystal data in " & CrystalsFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitCrystals"
ierror = True
Exit Sub

End Sub

Sub InitData()
' Initialize global data and load configuration files

ierror = False
On Error GoTo InitDataError

Dim i As Integer, linecount As Integer
Dim tmsg As String, astring As String

Static initialized As Boolean

' Program version number
ProgramVersionString$ = Format$(app.major) & "." & Format$(app.minor) & "." & Format$(app.Revision)
ProgramVersionNumber! = Val(Format$(app.major) & "." & Format$(app.minor) & Format$(app.Revision))

RealTimeInterfaceBusy = False

' Initialize open data file names
ProbeDataFile$ = vbNullString
ProbeImageFile$ = vbNullString

' Make sure temp files are closed
Close #Temp1FileNumber%
Close #Temp2FileNumber%

' Randomize random number generator
Randomize

' Re-set default folders (in case error opening file)
CalcZAFDATFileDirectory$ = OriginalCalcZAFDATDirectory$
ColumnPCCFileDirectory = OriginalColumnPCCFileDirectory$
SurferDataDirectory$ = OriginalSurferDataDirectory$
GrapherDataDirectory$ = OriginalGrapherDataDirectory$

UserDataDirectory$ = OriginalUserDataDirectory$
UserImagesDirectory = OriginalUserImagesDirectory$
DemoImagesDirectory$ = OriginalDemoImagesDirectory$

UserEDSDirectory$ = OriginalUserEDSDirectory$
UserCLDirectory$ = OriginalUserCLDirectory$
UserEBSDDirectory$ = OriginalUserEBSDDirectory$

' Set exclusive (write) access permissions
DatabaseExclusiveAccess% = True
StandardDatabaseExclusiveAccess% = True
ProbeDatabaseExclusiveAccess% = True
SetupDatabaseExclusiveAccess% = True
UserDatabaseExclusiveAccess% = True
PositionDatabaseExclusiveAccess% = True
XrayDatabaseExclusiveAccess% = True
MatrixDatabaseExclusiveAccess% = True
BoundaryDatabaseExclusiveAccess% = True
PureDatabaseExclusiveAccess% = True

' Set non exclusive (read only) access permissions
DatabaseNonExclusiveAccess% = False
StandardDatabaseNonExclusiveAccess% = False
ProbeDatabaseNonExclusiveAccess% = False
SetupDatabaseNonExclusiveAccess% = False
UserDatabaseNonExclusiveAccess% = False
PositionDatabaseNonExclusiveAccess% = False
XrayDatabaseNonExclusiveAccess% = False
MatrixDatabaseNonExclusiveAccess% = False
BoundaryDatabaseNonExclusiveAccess% = False
PureDatabaseNonExclusiveAccess% = False

' Set defaults
InterfaceString$(0) = "EPMA Demonstration"
InterfaceString$(1) = "Unused"

If InterfaceType% = 2 Then
If JeolEOSInterfaceType& = 1 Then InterfaceString$(2) = "JEOL 8200/8500 (TCP/IP Socket)"
If JeolEOSInterfaceType& = 2 Then InterfaceString$(2) = "JEOL 8900 (TCP/IP Socket)"
If JeolEOSInterfaceType& = 3 Then InterfaceString$(2) = "JEOL 8230/8530 (TCP/IP Socket and EIKS)"
End If

InterfaceString$(3) = "Unused"
InterfaceString$(4) = "Unused"

InterfaceString$(5) = "Cameca SX100 (TCP/IP Socket)"
InterfaceString$(6) = "Axioscope (Zeiss Serial ASCII)"

' EDS interface type, 0 = Demo, 1 = Unused, 2 = Bruker, 3 = Oxford, 4 = Unused, 5 = Thermo, 6 = JEOL OEM
InterfaceStringEDS$(0) = "EDS Demonstration"
InterfaceStringEDS$(1) = "Unused"
InterfaceStringEDS$(2) = "Bruker Quantax"
InterfaceStringEDS$(3) = "Oxford INCA"
InterfaceStringEDS$(4) = "Unused"
InterfaceStringEDS$(5) = "Thermo NSS"
InterfaceStringEDS$(6) = "JEOL OEM"

' Set defaults
InterfaceStringCL$(0) = "CL Demonstration"
InterfaceStringCL$(1) = "Ocean Optics"
InterfaceStringCL$(2) = "Gatan"
InterfaceStringCL$(3) = "Newport"
InterfaceStringCL$(4) = "Unused"

If CLSpectraInterfaceType% = 0 Then InterfaceStringCLUnitsX$(0) = "nm"
If CLSpectraInterfaceType% = 1 Then InterfaceStringCLUnitsX$(1) = "nm"
If CLSpectraInterfaceType% = 2 Then InterfaceStringCLUnitsX$(2) = "nm"
If CLSpectraInterfaceType% = 3 Then InterfaceStringCLUnitsX$(3) = "nm"
If CLSpectraInterfaceType% = 4 Then InterfaceStringCLUnitsX$(4) = "nm"

' Demo (Bruker)
MaxEnergyArraySize% = 4
ReDim MaxEnergyArrayValue(1 To MaxEnergyArraySize%) As Single
MaxEnergyArrayValue!(1) = 10#
MaxEnergyArrayValue!(2) = 20#
MaxEnergyArrayValue!(3) = 40#
MaxEnergyArrayValue!(4) = 80#

MaxThroughputArraySize% = 4
ReDim MaxThroughputArrayValue(1 To MaxThroughputArraySize%) As Single
MaxThroughputArrayValue!(1) = 60#
MaxThroughputArrayValue!(2) = 90#
MaxThroughputArrayValue!(3) = 130#
MaxThroughputArrayValue!(4) = 275#

' Bruker
If EDSSpectraInterfaceType% = 2 Then
MaxEnergyArraySize% = 4
ReDim MaxEnergyArrayValue(1 To MaxEnergyArraySize%) As Single
MaxEnergyArrayValue!(1) = 10#
MaxEnergyArrayValue!(2) = 20#
MaxEnergyArrayValue!(3) = 40#
MaxEnergyArrayValue!(4) = 80#

MaxThroughputArraySize% = 4
ReDim MaxThroughputArrayValue(1 To MaxThroughputArraySize%) As Single
MaxThroughputArrayValue!(1) = 60#
MaxThroughputArrayValue!(2) = 90#
MaxThroughputArrayValue!(3) = 130#
MaxThroughputArrayValue!(4) = 275#

' Thermo
ElseIf EDSSpectraInterfaceType% = 5 Then
MaxEnergyArraySize% = 5
ReDim MaxEnergyArrayValue(1 To MaxEnergyArraySize%) As Single
MaxEnergyArrayValue!(1) = 5#
MaxEnergyArrayValue!(2) = 10#
MaxEnergyArrayValue!(3) = 20#
MaxEnergyArrayValue!(4) = 40#
MaxEnergyArrayValue!(4) = 80#

MaxThroughputArraySize% = 11    ' (0 (AUTO), 6400, 4000, 3200, 2000, 1600, 1000, 800, 600, 400, 200 nano-secs)
ReDim MaxThroughputArrayValue(1 To MaxThroughputArraySize%) As Single
MaxThroughputArrayValue!(1) = 200#
MaxThroughputArrayValue!(2) = 400#
MaxThroughputArrayValue!(3) = 600#
MaxThroughputArrayValue!(4) = 800#
MaxThroughputArrayValue!(5) = 1000#
MaxThroughputArrayValue!(6) = 1600#
MaxThroughputArrayValue!(7) = 2000#
MaxThroughputArrayValue!(8) = 3200#
MaxThroughputArrayValue!(9) = 4000#
MaxThroughputArrayValue!(10) = 6400#
MaxThroughputArrayValue!(11) = 0#
End If

' Image interface type, 0=Demo, 1=Unused, 2=Unused, 3=Unused, 4=8900/8200/8500/8x30, 5=SX100/SXFive mapping, 6=SX100/SXFive Video, 7=Unused, 8=Unused, 9=Bruker, 10=Thermo
InterfaceStringImage(0) = "Demonstration (Imaging)"
InterfaceStringImage(1) = "Unused"
InterfaceStringImage(2) = "Unused"
InterfaceStringImage(3) = "Unused"
InterfaceStringImage(4) = "JEOL 8900/8200/8500"
InterfaceStringImage(5) = "Cameca SX100 Mapping"
InterfaceStringImage(6) = "Cameca SX100 Video"
InterfaceStringImage(7) = "Unused"
InterfaceStringImage(8) = "Unused"
InterfaceStringImage(9) = "Bruker RTIfcCLIENT"
InterfaceStringImage(10) = "Thermo TEPortal"

' Update Output menu items
DebugMode = DefaultDebugMode
If DebugMode Then
FormMAIN.menuOutputDebugMode.Checked = True
Else
FormMAIN.menuOutputDebugMode.Checked = False
End If

If ExtendedFormat Then
FormMAIN.menuOutputExtendedFormat.Checked = True
Else
FormMAIN.menuOutputExtendedFormat.Checked = False
End If

' See if user wants a log of init messages
If DebugMode Then
Close #OutputDataFileNumber%        ' in case another app is opening Matrix or Remote both in debug mode
SaveToDisk = True
FormMAIN.menuOutputSaveToDiskLog.Checked = True
Open OutputDataFile$ For Output As #OutputDataFileNumber%
Else
FormMAIN.menuOutputSaveToDiskLog.Checked = False
End If

' Standard database flags
MinSpecifiedValue! = 0.05 ' minimum value (wt%) to force specified load for standard analysis
NotAnalyzedValue! = 0.00000001
StdMinimumValue! = 0.01  ' minimum value (wt%) for use as assigned standard
MANMaximumValue! = 0.01  ' maximum value (wt%) for use as MAN background standard
MinTotalValue! = 0.0001  ' minimum valid total

' Standard arrays
Call InitStandard
If ierror Then Exit Sub

' Sample counters
NumberofSamples% = 0    ' 1 to MaxSample
NumberofUnknowns% = 0   ' 1 to MaxSample
NumberofWavescans% = 0    ' 1 to MaxSample
NumberofLines& = 0

' Dimension large global arrays
ReDim MANAssignsDriftCounts(1 To MAXSET%, 1 To MAXMAN%, 1 To MAXCHAN%) As Single
ReDim MANAssignsDateTimes(1 To MAXSET%, 1 To MAXMAN%, 1 To MAXCHAN%) As Double
ReDim MANAssignsSets(1 To MAXMAN%, 1 To MAXCHAN%) As Integer
ReDim MANAssignsSampleRows(1 To MAXSET%, 1 To MAXMAN%, 1 To MAXCHAN%) As Integer

ReDim MANAssignsCountTimes(1 To MAXSET%, 1 To MAXMAN%, 1 To MAXCHAN%) As Single
ReDim MANAssignsBeamCurrents(1 To MAXSET%, 1 To MAXMAN%, 1 To MAXCHAN%) As Single

ReDim CurrentMultiPointAcquirePositionsHi(1 To MAXCHAN%, 1 To MAXMULTI%) As Single
ReDim CurrentMultiPointAcquirePositionsLo(1 To MAXCHAN%, 1 To MAXMULTI%) As Single
ReDim CurrentMultiPointAcquireLastCountTimesHi(1 To MAXCHAN%, 1 To MAXMULTI%) As Single
ReDim CurrentMultiPointAcquireLastCountTimesLo(1 To MAXCHAN%, 1 To MAXMULTI%) As Single

' Load standard names and numbers
Call InitStandardIndex
If ierror Then Exit Sub

' Init sample arrays
For i% = 1 To MAXSAMPLE%
SampleNums%(i%) = 0
SampleTyps%(i%) = 0
SampleSets%(i%) = 0
SampleNams$(i%) = vbNullString
SampleDess$(i%) = vbNullString
SampleDels%(i%) = True
SampleMags!(i%) = 0#
Next i%

' Special string formats (need to be variables)
a08$ = String$(8, 64) & "!"
a10$ = String$(10, 64) & "!"
a12$ = String$(12, 64) & "!"
a14$ = String$(14, 64) & "!"
a16$ = String$(16, 64) & "!"
a18$ = String$(18, 64) & "!"
a22$ = String$(22, 64) & "!"
a24$ = String$(24, 64) & "!"
a32$ = String$(32, 64) & "!"
a64$ = String$(64, 64) & "!"

EmpiricalAlphaFlag% = 1 ' do not use empirical alpha factors
'CorrectionFlag% = 0  ' ZAF/Phi-Rho-Z default (set in PROBEWIN.INI) (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
'izaf% = 1   ' zaf correction (pre-selected options) (set in PROBEWIN.INI)

' Select the default ZAF matrix correction algorithms (set in PROBEWIN.INI)
'ibsc% = 2
'imip% = 1
'iphi% = 2
'iabs% = 9
'istp% = 4
'ibks% = 4

iflu% = 1

zafstring$(0) = "Select Individual Corrections"
zafstring$(1) = "Armstrong/Love Scott (default)"
zafstring$(2) = "Conventional Philibert/Duncumb-Reed"
zafstring$(3) = "Heinrich/Duncumb-Reed"
zafstring$(4) = "Love-Scott I"
zafstring$(5) = "Love-Scott II"
zafstring$(6) = "Packwood Phi(pz) (EPQ-91)"
zafstring$(7) = "Bastin (original) Phi(pz)"
zafstring$(8) = "Bastin PROZA Phi(pz) (EPQ-91)"
zafstring$(9) = "Pouchou and Pichoir-Full (Original)"
zafstring$(10) = "Pouchou and Pichoir-Simplified (XPP)"

zafstring2$(0) = "Individual"
zafstring2$(1) = "Armstrong"
zafstring2$(2) = "Philibert"
zafstring2$(3) = "Heinrich"
zafstring2$(4) = "Love-Scott I"
zafstring2$(5) = "Love-Scott II"
zafstring2$(6) = "Packwood"
zafstring2$(7) = "Bastin"
zafstring2$(8) = "Bastin PROZA"
zafstring2$(9) = "PAP Full"
zafstring2$(10) = "PAP XPP"

' Define ZAF selection strings
mipstring$(1) = "Mean Ionization of Berger-Seltzer"
mipstring$(2) = "Mean Ionization of Duncumb-Da Casa"
mipstring$(3) = "Mean Ionization of Ruste and Zeller"
mipstring$(4) = "Mean Ionization of Springer"
mipstring$(5) = "Mean Ionization of Wilson"
mipstring$(6) = "Mean Ionization of Heinrich"
mipstring$(7) = "Mean Ionization of Bloch (Love-Scott)"
mipstring$(8) = "Mean Ionization of Armstrong (Springer-Berger)"
mipstring$(9) = "Mean Ionization of Joy (Wilson-Berger)"

bscstring$(1) = "Backscatter Coefficient of Heinrich"
bscstring$(2) = "Backscatter Coefficient of Love-Scott"
bscstring$(3) = "Backscatter Coefficient of Pouchou and Pichoir"
bscstring$(4) = "Backscatter Coefficient of Hungler-Kuchler (A-W Mod.)"

phistring$(1) = "Phi(pz) Equation of Reuter"
phistring$(2) = "Phi(pz) Equation of Love-Scott"
phistring$(3) = "Phi(pz) Equation of Riveros"
phistring$(4) = "Phi(pz) Equation of Pouchou and Pichoir"
phistring$(5) = "Phi(pz) Equation of Karduck and Rehbach"
phistring$(6) = "Phi(pz) Equation of August and Wernisch"
phistring$(7) = "Phi(pz) Equation of Packwood"

stpstring$(1) = "Stopping Power of Duncumb-Reed (FRAME)"
stpstring$(2) = "Stopping Power of Philibert and Tixier"
stpstring$(3) = "Stopping Power (Numerical Integration)"
stpstring$(4) = "Stopping Power of Love-Scott"
stpstring$(5) = "Stopping Power of Pouchou and Pichoir"
stpstring$(6) = "Stopping Power of Phi(pz) Integration"

bksstring$(0) = "No Backscatter (used for Packwood and Bastin original)"
bksstring$(1) = "Backscatter of Duncumb-Reed (FRAME-I)"
bksstring$(2) = "Backscatter of Duncumb-Reed (COR-II) and Heinrich"
bksstring$(3) = "Backscatter of Duncumb-Reed (COR-II) and Heinrich"
bksstring$(4) = "Backscatter of Love-Scott"
bksstring$(5) = "Backscatter of Myklebust-I"
bksstring$(6) = "Backscatter of Myklebust and Fiori (not implemented)"
bksstring$(7) = "Backscatter of Pouchou and Pichoir"
bksstring$(8) = "Backscatter of August, Razka and Wernisch"
bksstring$(9) = "Backscatter of Springer"

absstring$(1) = "Absorption of Philibert (FRAME)"
absstring$(2) = "Absorption of Heinrich (Quadratic Anal. Chem.)"
absstring$(3) = "Absorption of Heinrich (Duplex 1989 MAS)"
absstring$(4) = "Absorption of Love-Scott (1983 J. Phys. D.)"
absstring$(5) = "Absorption of Sewell-Love-Scott (1985-I J. Phys. D.)"
absstring$(6) = "Absorption of Sewell-Love-Scott (1985-II J. Phys. D.)"
absstring$(7) = "Phi(pz) Absorption of Packwood-Brown 1982/XRS Alpha"
absstring$(8) = "Phi(pz) Absorption of Bastin 1984/XRS Alpha"
absstring$(9) = "Phi(pz) Absorption of Armstrong/Packwood-Brown 1981 MAS"
absstring$(10) = "Phi(pz) Absorption of Bastin 1986/Scanning"
absstring$(11) = "Phi(pz) Absorption of Riveros 1987/XRS"
absstring$(12) = "Phi(pz) Absorption of Pouchou and Pichoir-Full (Original)"
absstring$(13) = "Phi(pz) Absorption of Pouchou and Pichoir-Simplified (XPP)"
absstring$(14) = "Phi(pz) Absorption of Packwood (New)"
absstring$(15) = "Phi(pz) Absorption of Bastin Proza (EPQ-91)"

macstring$(1) = "LINEMU   Henke (LBL, 1985) < 10KeV / CITZMU > 10KeV"
macstring$(2) = "CITZMU   Heinrich (1966) and Henke and Ebisu (1974)"
macstring$(3) = "MCMASTER McMaster (LLL, 1969) (modified by Rivers)"
macstring$(4) = "MAC30    Heinrich (Fit to Goldstein tables, 1987)"
macstring$(5) = "MACJTA   Armstrong (FRAME equations, 1992)"
macstring$(6) = "FFAST    Chantler (NIST v 2.1, 2005)"
macstring$(7) = "USERMAC  User Defined MAC Table"

macstring2$(1) = "LINEMU"
macstring2$(2) = "CITZMU"
macstring2$(3) = "MCMASTER"
macstring2$(4) = "MAC30"
macstring2$(5) = "MACJTA"
macstring2$(6) = "FFAST"
macstring2$(7) = "USERMAC"

flustring$(1) = "Reed/JTA w/ M-Line Correction and JTA Intensity Mod."
flustring$(2) = "Reed/JTA (CITZAF.BAS- original with no M-Line Correction)"
flustring$(3) = "Reed/JTA w/ M-Line Correction Only"
flustring$(4) = "Reed/JTA w/ M-Line Correction and Reed Intensity Mod."
flustring$(5) = "Improved Armstrong Fluorescence Correction"

corstring$(0) = "ZAF or Phi-Rho-Z Calculations"
corstring$(1) = "Constant Alpha Factors (single coefficient)"
corstring$(2) = "Linear Alpha Factors (two coefficients)"
corstring$(3) = "Polynomial Alpha Factors (three coefficients)"
corstring$(4) = "Non-Linear Alpha Factors (four coefficients)"
corstring$(5) = "Calibration Curve (multi-standard)"
corstring$(6) = "Fundamental Parameters (Llovet/Pinard)"

empstring$(1) = "Empirical Alpha Factors Not Used"
empstring$(2) = "Empirical Alpha Factors Used"

' Thin film strings
ptcstring$(1) = "Thin Film or Thick Polished Section (unsupported film or no substrate)"
ptcstring$(2) = "Rectangular Prism (flat top and flat sides or cube)"
ptcstring$(3) = "Tetragonal Prism (flat top and curved sides)"
ptcstring$(4) = "Trigonal Prism (curved top and flat sides or fiber)"
ptcstring$(5) = "Square Pyramid (curved top and curved sides or sphere)"
ptcstring$(6) = "Backscatter-Modified Rectangular Prism (modified rectangular prism)"

' Load volatile fit type strings
vstring$(0) = "LOG-LIN"
vstring$(1) = "HYP-EXP"
vstring$(2) = "LOG-LOG"

' Initialize xray strings
Xraylo$(1) = "ka"
Xraylo$(2) = "kb"
Xraylo$(3) = "la"
Xraylo$(4) = "lb"
Xraylo$(5) = "ma"
Xraylo$(6) = "mb"
Xraylo$(7) = vbNullString     ' unanalyzed element

' Initialize additional xray strings
If MAXRAY% - 1 > MAXRAY_OLD% Then
Xraylo$(7) = "Ln"
Xraylo$(8) = "Lg"
Xraylo$(9) = "Lv"
Xraylo$(10) = "Ll"
Xraylo$(11) = "Mg"
Xraylo$(12) = "Mz"
Xraylo$(13) = vbNullString     ' unanalyzed element
End If

Edglo$(1) = "k"
Edglo$(2) = "l1"
Edglo$(3) = "l2"
Edglo$(4) = "l3"
Edglo$(5) = "m1"
Edglo$(6) = "m2"
Edglo$(7) = "m3"
Edglo$(8) = "m4"
Edglo$(9) = "m5"

' Load maximum order
RomanNum$(1) = "I"
RomanNum$(2) = "II"
RomanNum$(3) = "III"
RomanNum$(4) = "IV"
RomanNum$(5) = "V"
RomanNum$(6) = "VI"
RomanNum$(7) = "VII"
RomanNum$(8) = "VIII"
RomanNum$(9) = "IX"
RomanNum$(10) = "X"
RomanNum$(11) = "XI"
RomanNum$(12) = "XII"
RomanNum$(13) = "XIII"
RomanNum$(14) = "XIV"
RomanNum$(15) = "XV"
RomanNum$(16) = "XVI"
RomanNum$(17) = "XVII"
RomanNum$(18) = "XVIII"
RomanNum$(19) = "XIX"
RomanNum$(20) = "XX"

SampleSyms$(1) = "St"
SampleSyms$(2) = "Un"
SampleSyms$(3) = "Wa"

beammodestrings$(0) = "Spot"
beammodestrings$(1) = "Scan"
beammodestrings$(2) = "Digital"

bgdtypestrings$(0) = "OFF"
bgdtypestrings$(1) = "MAN"
bgdtypestrings$(2) = "MULT"

bgdstrings$(0) = "Linear"
bgdstrings$(1) = "Average"
bgdstrings$(2) = "High Only"
bgdstrings$(3) = "Low Only"
bgdstrings$(4) = "Exponential"
bgdstrings$(5) = "Slope (Hi)"
bgdstrings$(6) = "Slope (Lo)"
bgdstrings$(7) = "Polynomial"
bgdstrings$(8) = "Multi-Point"

bgstrings$(0) = "LIN"
bgstrings$(1) = "AVG"
bgstrings$(2) = "HIGH"
bgstrings$(3) = "LOW"
bgstrings$(4) = "EXP"
bgstrings$(5) = "S-Hi"
bgstrings$(6) = "S-Lo"
bgstrings$(7) = "POLY"
bgstrings$(8) = "MULT"

bglstrings$(0) = "LINEAR"
bglstrings$(1) = "AVERAGE"
bglstrings$(2) = "HIGH"
bglstrings$(3) = "LOW"
bglstrings$(4) = "EXPONEN"
bglstrings$(5) = "SLOPEHI"
bglstrings$(6) = "SLOPELO"
bglstrings$(7) = "POLYNOM"
bglstrings$(8) = "MULTI"

monthsyms$(1) = "JAN"
monthsyms$(2) = "FEB"
monthsyms$(3) = "MAR"
monthsyms$(4) = "APR"
monthsyms$(5) = "MAY"
monthsyms$(6) = "JUN"
monthsyms$(7) = "JUL"
monthsyms$(8) = "AUG"
monthsyms$(9) = "SEP"
monthsyms$(10) = "OCT"
monthsyms$(11) = "NOV"
monthsyms$(12) = "DEC"

MineralStrings$(0) = "       --------       "
MineralStrings$(1) = "Olivine(Fo,Fa)        "
MineralStrings$(2) = "Feldspar(Ab,An,Or)    "
MineralStrings$(3) = "Pyroxene(Wo,En,Fs)    "
MineralStrings$(4) = "Garnet(Gro,Pyr,Alm,Sp)"
MineralStrings$(5) = "Garnet(Gro,And,Uva)   "

' DefaultImageAnalogUnits short strings
ImageAnalogUnitsShortStrings$(0) = "A/D Averages/Pixel"      ' demo
ImageAnalogUnitsShortStrings$(1) = "Not Implemented"         ' Unused
ImageAnalogUnitsShortStrings$(2) = "Not Implemented"         ' Unused
ImageAnalogUnitsShortStrings$(3) = "Not Implemented"         ' Unused

If JeolEOSInterfaceType& = 3 Then
ImageAnalogUnitsShortStrings$(4) = "Microsecs/Pixel"         ' JEOL 8230/8530
Else
ImageAnalogUnitsShortStrings$(4) = "A/D Averages/Pixel"      ' JEOL 8900/8200/8500
End If

ImageAnalogUnitsShortStrings$(5) = "Millisecs/Pixel"         ' SX100 mapping
ImageAnalogUnitsShortStrings$(6) = "Scan Rate"               ' SX100 video
ImageAnalogUnitsShortStrings$(7) = "Not Implemented"         ' Unused
ImageAnalogUnitsShortStrings$(8) = "Not Implemented"         ' Unused
ImageAnalogUnitsShortStrings$(9) = "A/D Averages/Pixel"      ' Bruker RTIfcClient
ImageAnalogUnitsShortStrings$(10) = "Frame Time in Secs"     ' Thermo TEPortal

' DefaultImageAnalogUnits long strings
ImageAnalogUnitsLongStrings(0) = "A/D Averages/Pixel (1-1000)"      ' demo
ImageAnalogUnitsLongStrings(1) = "Not Implemented"                  ' Unused
ImageAnalogUnitsLongStrings(2) = "Not Implemented"                  ' Unused
ImageAnalogUnitsLongStrings(3) = "Not Implemented"                  ' Unused

If JeolEOSInterfaceType& = 3 Then
ImageAnalogUnitsLongStrings(4) = "Microsecs/Pixel (100-100000000)"  ' JEOL 8230/8530
Else
ImageAnalogUnitsLongStrings(4) = "A/D Averages/Pixel (1-1000)"      ' JEOL 8900/8200/8500
End If

ImageAnalogUnitsLongStrings(5) = "Millisecs/Pixel (1-1000)"         ' SX100 mapping
ImageAnalogUnitsLongStrings(6) = "Scan Rate (1-7)"                  ' SX100 video
ImageAnalogUnitsLongStrings(7) = "Not Implemented"                  ' Unused
ImageAnalogUnitsLongStrings(8) = "Not Implemented"                  ' Unused
ImageAnalogUnitsLongStrings(9) = "A/D Averages/Pixel (2 to 1000)"   ' Bruker RTIfcClient
ImageAnalogUnitsLongStrings(10) = "Frame Time in Secs (1-100)"      ' Thermo TEPortal

' DefaultImageAnalogUnits ToolTip strings
ImageAnalogUnitsToolTipStrings(0) = "Specify the number of A-D conversions to average per pixel (range 1 - 1000)"       ' demo
ImageAnalogUnitsToolTipStrings(1) = "Not Implemented"                                                                   ' Unused
ImageAnalogUnitsToolTipStrings(2) = "Not Implemented"                                                                   ' Unused
ImageAnalogUnitsToolTipStrings(3) = "Not Implemented"                                                                   ' Unused

If JeolEOSInterfaceType& = 3 Then
ImageAnalogUnitsToolTipStrings(4) = "Specify the number micro-seconds per pixel (range 100 to 100000000)"               ' JEOL 8230/8530
Else
ImageAnalogUnitsToolTipStrings(4) = "Specify the number of A-D conversions to average per pixel (range 1 - 1000)"       ' JEOL 8900/8200/8500
End If

ImageAnalogUnitsToolTipStrings(5) = "Specify the image dwell time in milli-seconds per pixel (range 1 - 1000)"          ' SX100 mapping
ImageAnalogUnitsToolTipStrings(6) = "Specify the image scan speed (range = 1 - 7)"                                      ' SX100 Video
ImageAnalogUnitsToolTipStrings(7) = "Not Implemented"                                                                   ' Unused
ImageAnalogUnitsToolTipStrings(8) = "Not implemented"                                                                   ' Unused
ImageAnalogUnitsToolTipStrings(9) = "Specify the number of A-D conversions to average per pixel (range 2 - 1000)"       ' Bruker RTIfcClient
ImageAnalogUnitsToolTipStrings(10) = "Enter the imaging frame time in seconds (1 to 100)"                               ' Thermo TEPortal

' Load interference assignment orders
For i% = 1 To MAXINTF%
If i% = 1 Then
InterfSyms$(i%) = "1st"
ElseIf i% = 2 Then
InterfSyms$(i%) = "2nd"
ElseIf i% = 3 Then
InterfSyms$(i%) = "3rd"
Else
InterfSyms$(i%) = Str$(i%) & "th"
End If
Next i%

' Check for application files
Call InitCheckFiles
If ierror Then Exit Sub

' Read the CRYSTALS.DAT file (read first)
Close #Temp1FileNumber%     ' close first in case it is already open
If Dir$(CrystalsFile$) = vbNullString Then GoTo InitDataNotFoundCrystalsFile
Open CrystalsFile$ For Input As #Temp1FileNumber%
Call InitCrystals
Close #Temp1FileNumber%
If ierror Then Exit Sub

' Read element data from ELEMENTS.DAT file (read second)
If Dir$(ElementsFile$) = vbNullString Then GoTo InitDataNotFoundElementsFile
Open ElementsFile$ For Input As #Temp1FileNumber%
Call InitElements
Close #Temp1FileNumber%
If ierror Then Exit Sub

' Read the MOTORS file (read third)
If Dir$(MotorsFile$) = vbNullString Then GoTo InitDataNotFoundMotorsFile
Open MotorsFile$ For Input As #Temp1FileNumber%
Call InitMotors
Close #Temp1FileNumber%
If ierror Then Exit Sub

' Now load PHA parameters (must be called before reading SCALERS.DAT and after MOTORS.DAT)
Call InitMinMax
If ierror Then Exit Sub

' Check the SCALERS.DAT file for enough lines and update if necessary
If Dir$(ScalersFile$) = vbNullString Then GoTo InitDataNotFoundScalersFile
Open ScalersFile$ For Input As #Temp1FileNumber%
Call InitScalersCheck(Int(0), linecount%, astring$)
Close #Temp1FileNumber%
If ierror Then Exit Sub

If linecount% < SCALERSDATLINESNEW% Then
Open ScalersFile$ For Append As #Temp1FileNumber%
Call InitScalersCheck(Int(1), linecount%, astring$)
Close #Temp1FileNumber%
If ierror Then Exit Sub
End If

' Read the SCALERS.DAT file (read fourth)
Open ScalersFile$ For Input As #Temp1FileNumber%
Call InitScalers
Close #Temp1FileNumber%
If ierror Then Exit Sub

' Create the CHARGES.DAT if not present
If Dir$(ChargesFile$) = vbNullString Then
Open ChargesFile$ For Output As #Temp1FileNumber%
Call InitChargesCreate
Close #Temp1FileNumber%
If ierror Then Exit Sub
End If

' Load CHARGES.DAT
If Dir$(ChargesFile$) = vbNullString Then GoTo InitDataNotFoundChargesFile
Open ChargesFile$ For Input As #Temp1FileNumber%
Call InitCharges
Close #Temp1FileNumber%
If ierror Then Exit Sub

' Create the DENSITY.DAT if not present
If Dir$(DensityFile$) = vbNullString Then
Open DensityFile$ For Output As #Temp1FileNumber%
Call InitDensityCreate
Close #Temp1FileNumber%
If ierror Then Exit Sub
End If

' Load DENSITY.DAT
If Dir$(DensityFile$) = vbNullString Then GoTo InitDataNotFoundDensityFile
Open DensityFile$ For Input As #Temp1FileNumber%
Call InitDensity
Close #Temp1FileNumber%
If ierror Then Exit Sub

' Load DENSITY2.DAT
If Dir$(DensityFile2$) = vbNullString Then GoTo InitDataNotFoundDensityFile2
Open DensityFile2$ For Input As #Temp1FileNumber%
Call InitDensity2
Close #Temp1FileNumber%
If ierror Then Exit Sub

' Create the DETECTORS.DAT if not present
If Dir$(DetectorsFile$) = vbNullString Then
Open DetectorsFile$ For Output As #Temp1FileNumber%
Call InitDetectorsCreate
Close #Temp1FileNumber%
If ierror Then Exit Sub
End If

' Load DETECTORS.DAT
If Dir$(DetectorsFile$) = vbNullString Then GoTo InitDataNotFoundDetectorsFile
Open DetectorsFile$ For Input As #Temp1FileNumber%
Call InitDetectors
Close #Temp1FileNumber%
If ierror Then Exit Sub

' Load multiple peak calibration files
For i% = 0 To 5 ' Ka, Kb, La, Lb, Ma, Mb
Call InitCoefficients(i%)
If ierror Then Exit Sub
Next i%

' Load RGB values from Probewin.ini (after other config files)
Call InitINIImage2
If ierror Then Exit Sub

' Check FaradayStagePositions from INI file
If FaradayStagePresent Then
If FaradayStagePositions!(1) < MotLoLimits!(XMotor%) Or FaradayStagePositions!(1) > MotHiLimits!(XMotor%) Then
msg$ = "Faraday Cup X Stage Position out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitData"
ierror = True
Exit Sub
End If
If FaradayStagePositions!(2) < MotLoLimits!(YMotor%) Or FaradayStagePositions!(2) > MotHiLimits!(YMotor%) Then
msg$ = "Faraday Cup X Stage Position out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitData"
ierror = True
Exit Sub
End If
If FaradayStagePositions!(3) < MotLoLimits!(ZMotor%) Or FaradayStagePositions!(3) > MotHiLimits!(ZMotor%) Then
msg$ = "Faraday Cup X Stage Position out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitData"
ierror = True
Exit Sub
End If
If FaradayStagePositions!(4) < MotLoLimits!(WMotor%) Or FaradayStagePositions!(4) > MotHiLimits!(WMotor%) Then
msg$ = "Faraday Cup X Stage Position out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitData"
ierror = True
Exit Sub
End If
End If

' Misc flags
UseMACFlag = False      ' don't use empirical MACs as default
UseAPFFlag = False      ' don't use empirical APFs as default

UseInterfFlag = False    ' don't use interference correction as default
UseVolElFlag = False    ' don't use volatile element correction as default
UseVolElType% = 0    ' use linear as default
UseMANAbsFlag = True    ' use MAN continuum absorption as default
UseDetailedFlag = True    ' use extra printout as default
UseDriftFlag = True    ' use standard drift correction as default
UseBlankCorFlag = False  ' don't use blank trace correction as default

UseOffPeakElementsForMANFlag = False    ' don't use off-peak samples in MAN fits as default
UseMANForOffPeakElementsFlag = False     ' don't use MAN correction on off-peak samples as default
UseBeamDriftCorrectionFlag = True       ' use beam drift correction
If UCase$(app.EXEName) = UCase$("Probewin") Then UseDeadtimeCorrectionFlag = True        ' use deadtime correction for ProbeWin only by default

AcquisitionOrderFlag = 0
AcquisitionMotionFlag = 0
ReturnToOnPeakFlag = True
UseQuickStandardsFlag = False

PeakCenterMethodFlag = DefaultPeakCenterMethod% ' from INI file
PeakCenterPreScanFlag = False   ' default is no peaking prescan
PeakCenterPostScanFlag = False   ' default is no peaking postscan

SyncSpecMotionBeamBlankFlag = False
StdUnkMeasureFaradayFlag = True
WaveScanMeasureFaradayFlag = False
AbsorbedCurrentMeasureFlag = False
AcquireEDSSpectraFlag = False

' Use PHA automation based on PHA hardware flag
UseAutomatedPHAControlFlag = PHAHardware
LoadStdDataFromFileSetupFlag = True

AcquisitionOnSample = False
AcquisitionOnWavescan = False
AcquisitionOnAutomate = False
AcquisitionOnVolatile = False
AcquisitionOnQuickscan = False
AcquisitionOnAutoFocus = False

AcquisitionOnEDS = False
AcquisitionOnCL = False

' Base new automated samples on last unknown (or standard)
AutomateNewSampleBasisFlag% = 0
'AutomateConfirmFlag = True          ' do not initialize here (only in INI)

' Load defaults
If DefaultTakeOff! = 0# Then DefaultTakeOff! = 40#
If DefaultKiloVolts! = 0# Then DefaultKiloVolts! = 15#
If DefaultBeamCurrent! = 0# Then DefaultBeamCurrent! = 30#
If DefaultBeamSize! = 0# Then DefaultBeamSize! = 0#

If DefaultLIFPeakWidth! = 0# Then DefaultLIFPeakWidth! = 0.08
If DefaultMinimumOverlap! = 0# Then DefaultMinimumOverlap! = 0.1
If DefaultPHADiscrimination! = 0# Then DefaultPHADiscrimination! = 4#
If DefaultRangeFraction! = 0# Then DefaultRangeFraction! = 0.05

DefaultFiducialSetNumber% = 0

DefaultSampleSetupNumber% = 0
DefaultFileSetupName$ = vbNullString
DefaultFileSetupNumber% = 0
DefaultMultipleSetupNumber% = 0

' Force reload of standard arrays in case user opens another run with sample sample setup
AllAnalysisUpdateNeeded = True

' Force reload of afactor arrays
AllAFactorUpdateNeeded = True

' Load the default stage bit map index
If StageBitMapCount% > 0 Then
If StageBitMapIndex% < 1 Then StageBitMapIndex% = 1
End If

' Initialize x-axis increment for wavescans
WavescanXIncrement! = CSng(IncrementXForAdditionalPoints%)
WavescanYIncrement! = 0#    ' make zero the default
WavescanXIncrementInterval! = 10#

' Autofocus
UseAutoFocusFlag = False

FaradayCurrentUnits$ = "nA"
FaradayCurrentFormat$ = f83$

' Default halogen correction flag
UseOxygenFromHalogensCorrectionFlag = False

' Process synchronization interval
ProcessInterval! = 1#   ' seconds
RealTimePauseAutomation = False

' Set zero point calibration curve default
UseZeroPointCalibrationCurveFlag% = False

' Set calculation flags
CalculateElectronandXrayRangesFlag% = False

' Set default detector indexes
For i% = 1 To MAXSPEC%
RealTimeDetectorSlitSizes%(i%) = 1
RealTimeDetectorSlitPositions%(i%) = 1
RealTimeDetectorModes%(i%) = 1
Next i%

' Set auto format flags
UseAutomaticFormatForResultsFlag% = False
UseAutomaticFormatForResultsType% = 0

UseAPFOption% = 0
ExcelMethodOption% = 0

DefaultReplicates% = 1
ReplicatesStep% = 1

UseAggregateIntensitiesFlag = 0
UseForceSizeFlag = 0
UseForceColumnConditionFlag = 0

DisplayCountIntensitiesUnnormalizedFlag = False

' PTC defaults
iptc% = 0
PTCModel% = 1
PTCDiameter! = 10000#       ' in microns
PTCDensity! = 3#
PTCThicknessFactor! = 1#
PTCNumericalIntegrationStep! = 0.00001
PTCDoNotNormalizeSpecifiedFlag = False

' Initialize preset count time flag
UseEDSPresetCountTimeFlag = False   ' use passed count time for EDS
EDSPresetCountTime! = DefaultOnCountTime! * 4
EDSSpecifiedCountTime! = DefaultOnCountTime! * 4

CLSpecifiedCountTime! = DefaultOnCountTime! * 4
CLDarkSpectraCountTimeFraction! = 0.1

UseParticleCorrectionFlag% = False

' Custom color constants
VbDarkBlue& = RGB(0, 0, 230)

' Load element colors
XrayColor&(1) = RGB(255, 80, 80) ' Ka
XrayColor&(2) = RGB(220, 80, 80) ' Kb
XrayColor&(3) = RGB(80, 255, 80) ' La
XrayColor&(4) = RGB(80, 220, 80) ' Lb
XrayColor&(5) = RGB(80, 80, 255) ' Ma
XrayColor&(6) = RGB(80, 80, 220) ' Mb

If MAXRAY% - 1 > MAXRAY_OLD% Then
XrayColor&(7) = RGB(200, 100, 100)  ' Ln
XrayColor&(8) = RGB(128, 64, 64)    ' Lg
XrayColor&(9) = RGB(100, 200, 100)  ' Lv
XrayColor&(10) = RGB(64, 128, 64)   ' Ll
XrayColor&(11) = RGB(100, 100, 200) ' Mg
XrayColor&(12) = RGB(64, 64, 100)   ' Mz
End If

' Alternative array for graph control
XrayColor2%(1) = 12             ' Ka
XrayColor2%(2) = 4              ' Kb
XrayColor2%(3) = 10             ' La
XrayColor2%(4) = 2              ' Lb
XrayColor2%(5) = 9              ' Ma
XrayColor2%(6) = 1              ' Mb

If MAXRAY% - 1 > MAXRAY_OLD% Then
XrayColor2%(7) = 11             ' Ln
XrayColor2%(8) = 3              ' Lg
XrayColor2%(9) = 8              ' Lv
XrayColor2%(10) = 3             ' Ll
XrayColor2%(11) = 13            ' Mg
XrayColor2%(12) = 14            ' Mz
End If

' Load sample exchange positions
Call InitINI3

' Load default quick standard mode
UseQuickStandardsMode% = 0  ' normal
UseQuickStandardsMinimum! = 10  ' 10 wt. %

If ImageInterfaceType% = 0 Then         ' Demo
    DefaultImageAnalogAverages% = 8
    DefaultImageAnalogUnits$ = ImageAnalogUnitsLongStrings$(ImageInterfaceType%)

ElseIf ImageInterfaceType% = 1 Then     ' Unused
    DefaultImageAnalogAverages% = 1
    DefaultImageAnalogUnits$ = ImageAnalogUnitsLongStrings$(ImageInterfaceType%)

ElseIf ImageInterfaceType% = 2 Then     ' Unused
    DefaultImageAnalogAverages% = 1
    DefaultImageAnalogUnits$ = ImageAnalogUnitsLongStrings$(ImageInterfaceType%)

ElseIf ImageInterfaceType% = 3 Then     ' Unused
    DefaultImageAnalogAverages% = 1
    DefaultImageAnalogUnits$ = ImageAnalogUnitsLongStrings$(ImageInterfaceType%)

ElseIf ImageInterfaceType% = 4 Then     ' JEOL
    If JeolEOSInterfaceType& = 3 Then
    DefaultImageAnalogAverages% = 200   ' JEOL micro-secs (range 100 - 100000000)
    DefaultImageAnalogUnits$ = ImageAnalogUnitsLongStrings$(ImageInterfaceType%)
    Else
    DefaultImageAnalogAverages% = 8     ' JEOL A/D averages (range 1 - 1000)
    DefaultImageAnalogUnits$ = ImageAnalogUnitsLongStrings$(ImageInterfaceType%)
    End If

ElseIf ImageInterfaceType% = 5 Then     ' SX100 mapping rate in msec/pixel (range 1 - 1000)
    DefaultImageAnalogAverages% = 2
    DefaultImageAnalogUnits$ = ImageAnalogUnitsLongStrings$(ImageInterfaceType%)

ElseIf ImageInterfaceType% = 6 Then     ' SX100 video scan speed (range 1 - 7)
    DefaultImageAnalogAverages% = 5
    DefaultImageAnalogUnits$ = ImageAnalogUnitsLongStrings$(ImageInterfaceType%)

ElseIf ImageInterfaceType% = 7 Then     ' Unused
    DefaultImageAnalogAverages% = 1
    DefaultImageAnalogUnits$ = ImageAnalogUnitsLongStrings$(ImageInterfaceType%)

ElseIf ImageInterfaceType% = 8 Then     ' Unused
    DefaultImageAnalogAverages% = 1
    DefaultImageAnalogUnits$ = ImageAnalogUnitsLongStrings$(ImageInterfaceType%)

ElseIf ImageInterfaceType% = 9 Then     ' Bruker RTIFClient DCOM
    DefaultImageAnalogAverages% = 16
    DefaultImageAnalogUnits$ = ImageAnalogUnitsLongStrings$(ImageInterfaceType%)

ElseIf ImageInterfaceType% = 10 Then    ' Thermo TE_Portal
    DefaultImageAnalogAverages% = 20
    DefaultImageAnalogUnits$ = ImageAnalogUnitsLongStrings$(ImageInterfaceType%)
End If

DefaultImageChannelNumber% = 1
DefaultImageIx% = 128
DefaultImageIy% = 128

NumberOfImages% = 0
UseImageAutomateModeOnStds% = False
UseImageAutomateModeOnUnks% = False
UseImageAutomateModeOnWavs% = False
UseImageAutomateModes% = 2  ' default to after sample acquisition

ImagePaletteNumber% = DefaultImagePaletteNumber%

' Load palette first time to initialize all palettes
If Not MiscStringsAreSame(app.EXEName, "CalcZAF") And Not MiscStringsAreSame(app.EXEName, "Standard") And Not MiscStringsAreSame(app.EXEName, "Calmac") Then
Call ImageLoadPalette(ImagePaletteNumber%, ImagePaletteArray())
If ierror Then Exit Sub
End If

UseROMBasedSpectrometerScanFlag = False
DisplayPHAParameterDialogPriorFlag = False
DisplayPHAParameterDialogAfterFlag = False

AutomatedPHAParameterDialogPriorFlag = False
AutomatedPHAParameterDialogAfterFlag = True     ' changed 12/3/2014
AutomatedPHAParameterDialogTypeFlag% = 0

DoNotDisplayStandardImagesDuringDigitizationFlag% = False
NumberOfScans& = 0

BeamModeString$(0) = "Analog  Spot"
BeamModeString$(1) = "Analog  Scan"
BeamModeString$(2) = "Digital Spot"

If DefaultMaximumOrder% = 0 Then DefaultMaximumOrder% = 5
If DefaultKLMSpecificElement% = 0 Then DefaultKLMSpecificElement% = 26      ' default to iron

ROMPeakingString$(0) = "Internal"
ROMPeakingString$(1) = "Parabolic"
ROMPeakingString$(2) = "Maxima"
ROMPeakingString$(3) = "Gaussian"
ROMPeakingString$(4) = "Dual Maxima/Parabolic"
ROMPeakingString$(5) = "Dual Maxima/Gaussian"
ROMPeakingString$(6) = "Highest Intensity"

ROMPeakingString2$(0) = "Fine Scan"     ' JEOL 8900/8200/8500 and Cameca SX100/SXFive
ROMPeakingString2$(1) = "Coarse Scan"   ' JEOL 8900/8200/8500 and Cameca SX100/SXFive
ROMPeakingString2$(2) = "2nd Fine Scan" ' JEOL 8900/8200/8500 and Cameca SX100/SXFive

PHAHardwareTypeString(0) = "Trad. PHA"  ' traditional PHA scanning
PHAHardwareTypeString(1) = "MCA PHA"    ' MCA PHA scanning

' Turn off shared monitor packet for single stepping in debugger!!!!!
If InterfaceType% = 0 And MiscIsInstrumentStage("JEOL") And JeolEOSInterfaceType& < 3 Then
UseSharedMonitorDataFlag% = True    ' use shared monitor packets for DEMO JEOL 8900/8200/8500 instrument status
ElseIf InterfaceType% = 2 And JeolEOSInterfaceType& < 3 Then     ' only JEOL 8900/8200/8500 at this time
UseSharedMonitorDataFlag% = True    ' use shared monitor packets for instrument status
Else
UseSharedMonitorDataFlag% = False   ' use normal calls for instrument status
End If

' If TestType, turn off monitor packet messaging for testing purposes anyway
If (InterfaceType% = 0 And MiscIsInstrumentStage("JEOL") And JeolEOSInterfaceType& < 3) Or (InterfaceType% = 2 And JeolEOSInterfaceType& < 3) Then
If MiscStringsAreSame(app.EXEName, "TestType") Then
UseSharedMonitorDataFlag% = False   ' use normal calls for instrument status
tmsg$ = "WARNING in InitData- disabling shared monitor packet messaging for TestType application"
Call IOWriteLogRichText(tmsg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
End If
End If

AnalysisCheckForSamePeakPositions% = False
AnalysisCheckForSamePHASettings% = False

' Set limit to limit time for tunable specs travel
If InterfaceType% = 0 Then LimitToLimit! = 14   ' Demo
If InterfaceType% = 1 Then LimitToLimit! = 20   ' Unused
If InterfaceType% = 2 Then LimitToLimit! = 36   ' JEOL spectrometers are slower (use JEOLVelocity&(1) in 1/100th um/sec?)
If InterfaceType% = 3 Then LimitToLimit! = 50   ' Unused
If InterfaceType% = 4 Then LimitToLimit! = 50   ' Unused
If InterfaceType% = 5 Then LimitToLimit! = 20   ' SX100 (use SX100Velocity&(1) in steps/sec?, assume 60000 steps full range)
If InterfaceType% = 6 Then LimitToLimit! = 50   ' Axioscope

' Set limit to limit time for stage travel
If InterfaceType% = 0 Then LimitToLimit2! = 10   ' Demo
If InterfaceType% = 1 Then LimitToLimit2! = 20   ' Unused
If InterfaceType% = 2 Then LimitToLimit2! = 6    ' JEOL  (use JEOLVelocity&(XMotor%) in 1/100th um/sec?)
If InterfaceType% = 3 Then LimitToLimit2! = 50   ' Unused
If InterfaceType% = 4 Then LimitToLimit2! = 50   ' Unused
If InterfaceType% = 5 Then LimitToLimit2! = 10   ' SX100 (use SX100Velocity&(XMotor%) in steps/sec?, assume 48000 steps full range in X)
If InterfaceType% = 6 Then LimitToLimit2! = 50   ' Axioscope

'AutomationReStandardizationInterval = 0.0020833     ' in days (3 minutes for testing)
AutomationReStandardizationInterval = 0.25     ' in days (6 hours)

' Force TDI acquisition and quickscan flags to off
VolatileSelfCalibrationAcquisitionFlag = False
VolatileAssignedCalibrationAcquisitionFlag = False
VolatileCountIntervals% = DEFAULTVOLATILEINTERVALS%
QuickWaveScanAcquisitionFlag = False
If QuickscanSpeed! = 0# Then QuickscanSpeed! = 10#          ' assume 10% default

DoAnalysisOutputFlag = True         ' output analyses to log window and Analyze! form

' Always force the SMTP password blank
SMTPUserPassword$ = vbNullString

WavescanXIncrementFlag = False
PeakingXIncrementFlag = False
WaveScanMeasureFaradayNthPoint% = 1

DefaultMultiPointNumberofPointsAcquireHi% = 4
DefaultMultiPointNumberofPointsAcquireLo% = 4
DefaultMultiPointNumberofPointsIterateHi% = 2
DefaultMultiPointNumberofPointsIterateLo% = 2

'UseFluorescenceByBetaLinesFlag = False        ' do not initialize here (only in INI)
GeologicalSortOrderFlag% = 0

UseUnknownCountTimeForInterferenceStandardFlag% = 0
VolElTimeWeightingFactor% = 2

UseDoNotSetConditionsFlag = False

'UseCurrentBeamBlankStateOnStartUpAndTerminationFlag = False        ' do not initialize here (only in INI)
'ShowAllPeakingOptionsFlag = False                                  ' do not initialize here (only in INI)
UseRightMouseClickToDigitizeFlag = False
UseChemicalAgeCalculationFlag = False

IntegratedIntensitySmoothingPointsPerSide% = 2      ' use 2 for default to prevent over-fitting
IntegratedIntensityUseSmoothingFlag = False

AnalysisInSilentModeFlag = False

' Dimension alpha factor arrays
Call InitKratios
If ierror Then Exit Sub

' Load binary compositional ranges
BinaryRanges!(1) = 99#
BinaryRanges!(2) = 95#
BinaryRanges!(3) = 90#
BinaryRanges!(4) = 80#
BinaryRanges!(5) = 60#
BinaryRanges!(6) = 50#
BinaryRanges!(7) = 40#
BinaryRanges!(8) = 20#
BinaryRanges!(9) = 10#
BinaryRanges!(10) = 5#
BinaryRanges!(11) = 1#

DecontaminationTime! = 0#
DecontaminationTimeFlag = False

UsePenepmaKratiosFlag% = 1 ' do not use PENEPMA based k-ratios for alpha factor calculations (1 = no, 2 = yes)
UseFilamentStandbyFlag = False  ' re-set flag when opening new run

OnPeakTimeFractionFlag = False
OnPeakTimeFractionValue! = 1#

TimeStampMode = False

'SampleImportExportFlag = False  ' undocumented feature loaded in INI file!!!

If MiscIsInstrumentStage("CAMECA") Then
Default_X_Polarity% = 0                     ' Cameca stage
Default_Y_Polarity% = 0
Default_Stage_Units$ = "um"
Else
Default_X_Polarity% = -1                    ' JEOL stage
Default_Y_Polarity% = -1
Default_Stage_Units$ = "mm"
End If

EDSUnknownCountFactor! = 1#
CLUnknownCountFactor! = 1#

UseDefaultFocusFlag = True

MACFile$ = ApplicationCommonAppData$ & macstring2$(MACTypeFlag%) & ".DAT"

MinimumOverVoltageType% = 1   ' use 10 percent minimum overvoltage (0 = 2%, 1 = 10%, 2 = 20%)
DefaultEDSDeadtimePercent! = 50#    ' assume 50% deadtime
SurferPageSecondsDelay% = 8     ' used by CalcImage

' Set EDS and CL intensity defaults
If UCase$(Trim$(app.EXEName$)) = UCase$(Trim$("Standard")) Then
EDSIntensityOption% = 1         ' cps for Standard
CLIntensityOption% = 1          ' cps for Standard
Else
EDSIntensityOption% = 0         ' raw intensity for PFE/TestEDS
CLIntensityOption% = 0          ' raw intensity for PFE/TestEDS
End If

RunProbeImageFlag = False
ProbeImageAcquisitionFile$ = vbNullString
ProbeImageSampleSetupNumber% = 0

UseStageReproducibilityCorrectionFlag = False
ImageSizeIndex% = -1            ' to force loading of default image size

If CLSpectraInterfaceType% = 0 Then
CLSpectrumAcquisitionOverhead = 1#          ' demo mode
ElseIf CLSpectraInterfaceType% = 1 Then
CLSpectrumAcquisitionOverhead = 2.7         ' CL acquisition overhead (Ocean Optics driver using RealTimeInterval! acquisition intervals)
ElseIf CLSpectraInterfaceType% = 2 Then
CLSpectrumAcquisitionOverhead = 1#          ' Gatan
ElseIf CLSpectraInterfaceType% = 3 Then
CLSpectrumAcquisitionOverhead = 1#          ' Newport
ElseIf CLSpectraInterfaceType% = 4 Then
CLSpectrumAcquisitionOverhead = 1#          ' unused
End If

' Make sure sample data files are up to date (use root path as of 3-20-2007)
Call InitFilesUserData
If ierror Then Exit Sub

' Get registration information (do not check Matrix or Remote Matrix)
If Not MiscStringsAreSame(app.EXEName, "Matrix") And Not MiscStringsAreSame(app.EXEName, "Remote") Then
Call RegisterLoad2
If ierror Then Exit Sub
End If

' Type welcome message
If Not initialized Then
If DebugMode Then Call IOWriteLog(vbNullString)   ' to clear output from TestType
If MiscStringsAreSame(app.EXEName, "Probewin") Then
tmsg$ = "Welcome to Probe for EPMA (Xtreme Edition) v. " & ProgramVersionString$
Else
tmsg$ = "Welcome to " & app.EXEName & ", Probe for EPMA (Xtreme Edition) v. " & ProgramVersionString$
End If
Call IOWriteLogRichText(tmsg$, vbNullString, Int(LogWindowFontSize% + 2), vbBlue, Int(FONT_BOLD% Or FONT_UNDERLINE%), Int(0))
tmsg$ = "Copyright (c) 1995-2016 John J. Donovan"
Call IOWriteLogRichText(tmsg$, vbNullString, Int(LogWindowFontSize% + 2), vbBlue, Int(FONT_BOLD%), Int(0))
tmsg$ = vbCrLf & "This software is registered to :"
Call IOWriteLog(tmsg$)
tmsg$ = RegistrationName$ & vbCrLf & RegistrationInstitution$
Call IOWriteLog(tmsg$)
tmsg$ = vbCrLf & "Press the F1 key in any window for context sensitive help. To get help on a menu item simply highlight with the mouse and hit the F1 key."
Call IOWriteLogRichText(tmsg$ & vbCrLf, vbNullString, Int(LogWindowFontSize% + 2), vbRed, Int(FONT_BOLD%), Int(0))

' Load nominal beam
NominalBeam! = DefaultBeamCurrent!
initialized = True
End If

Exit Sub

' Errors
InitDataError:
If Err = VB_FileNotFound& Or Err = VB_FileAlreadyOpen Then
MsgBox Error$ & ": " & OutputDataFile$, vbOKOnly + vbCritical, "InitData"
Else
MsgBox Error$, vbOKOnly + vbCritical, "InitData"
End If
Close #Temp1FileNumber%
ierror = True
Exit Sub

InitDataNotFoundCrystalsFile:
msg$ = "File " & CrystalsFile$ & " was not found in the application data folder " & ApplicationCommonAppData$
MsgBox msg$, vbOKOnly + vbExclamation, "InitData"
ierror = True
Exit Sub

InitDataNotFoundElementsFile:
msg$ = "File " & ElementsFile$ & " was not found in the application data folder " & ApplicationCommonAppData$
MsgBox msg$, vbOKOnly + vbExclamation, "InitData"
ierror = True
Exit Sub

InitDataNotFoundMotorsFile:
msg$ = "File " & MotorsFile$ & " was not found in the application data folder " & ApplicationCommonAppData$
MsgBox msg$, vbOKOnly + vbExclamation, "InitData"
ierror = True
Exit Sub

InitDataNotFoundScalersFile:
msg$ = "File " & ScalersFile$ & " was not found in the application data folder " & ApplicationCommonAppData$
MsgBox msg$, vbOKOnly + vbExclamation, "InitData"
ierror = True
Exit Sub

InitDataNotFoundChargesFile:
msg$ = "File " & ChargesFile$ & " was not found in the application data folder " & ApplicationCommonAppData$
MsgBox msg$, vbOKOnly + vbExclamation, "InitData"
ierror = True
Exit Sub

InitDataNotFoundDensityFile:
msg$ = "File " & DensityFile$ & " was not found in the application data folder " & ApplicationCommonAppData$
MsgBox msg$, vbOKOnly + vbExclamation, "InitData"
ierror = True
Exit Sub

InitDataNotFoundDensityFile2:
msg$ = "File " & DensityFile2$ & " was not found in the application data folder " & ApplicationCommonAppData$
MsgBox msg$, vbOKOnly + vbExclamation, "InitData"
ierror = True
Exit Sub

InitDataNotFoundDetectorsFile:
msg$ = "File " & DetectorsFile$ & " was not found in the application data folder " & ApplicationCommonAppData$
MsgBox msg$, vbOKOnly + vbExclamation, "InitData"
ierror = True
Exit Sub

End Sub

Sub InitCharges()
' Reads the CHARGES.DAT file (charge defaults)

ierror = False
On Error GoTo InitChargesError

Dim inum As Integer
Dim isym As String
Dim i As Integer, linecount As Integer

linecount% = 1
For i% = 1 To MAXELM%
Input #Temp1FileNumber%, inum%, isym$, AllAtomicCharges!(i%)
If AllAtomicCharges!(i%) < -10# Or AllAtomicCharges!(i%) > 10# Then GoTo InitChargesInvalidData
linecount% = linecount% + 1
Next i%

Exit Sub

' Errors
InitChargesError:
MsgBox Error$, vbOKOnly + vbCritical, "InitCharges"
ierror = True
Exit Sub

InitChargesInvalidData:
msg$ = "Invalid charge data in " & ChargesFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitCharges"
ierror = True
Exit Sub

End Sub

Sub InitChargesCreate()
' Create a default CHARGES.DAT file

ierror = False
On Error GoTo InitChargesCreateError

Dim i As Integer, itemp As Integer

For i% = 1 To MAXELM%
If AllOxd%(i%) <> 0 Then
itemp% = -1 * AllOxd%(i%) * -2 / AllCat%(i%)     ' if oxide
Else
itemp% = -1     ' if halogen
End If
If i% = 2 Then itemp% = 0    ' if noble
If i% = 8 Then itemp% = -2   ' if oxygen
If i% = 10 Then itemp% = 0   ' if noble
If i% = 18 Then itemp% = 0   ' if noble
If i% = 36 Then itemp% = 0   ' if noble
If i% = 54 Then itemp% = 0   ' if noble
If i% = 86 Then itemp% = 0   ' if noble
Print #Temp1FileNumber%, AllAtomicNums%(i%), VbDquote$ & Symup$(i%) & VbDquote$, itemp%
Next i%

Exit Sub

' Errors
InitChargesCreateError:
MsgBox Error$, vbOKOnly + vbCritical, "InitChargesCreate"
ierror = True
Exit Sub

End Sub

Sub InitDetectors()
' Reads the DETECTORS.DAT file for microprobe detector configuration

ierror = False
On Error GoTo InitDetectorsError

Dim comment As String
Dim i As Integer, j As Integer
Dim linecount As Integer

Dim itemp(1 To MAXSPEC%) As Integer
Dim atemp(1 To MAXSPEC%) As Single
Dim astring(1 To MAXSPEC%) As String
Dim bstring(1 To MAXDET%, 1 To MAXSPEC%) As String

If DebugMode Then
Call IOWriteLog(vbCrLf & vbCrLf & "Detectors Configuration Information:")
End If

' Load channel labels (temporary variable)
linecount% = 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, astring$(i%)
msg$ = msg$ & Format$(astring$(i%), a80$)
If Not MiscStringsAreSame(astring$(i%), ScalLabels$(i%)) Then GoTo InitDetectorsBadScalerLabels
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Load number of detector slit sizes
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, DetSlitSizesNumber%(i%)
msg$ = msg$ & Format$(DetSlitSizesNumber%(i%), a80$)
If DetSlitSizesNumber%(i%) > MAXDET% Then GoTo InitDetectorsTooManySlits
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

For j% = 1 To MAXDET%
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, DetSlitSizes$(j%, i%)
If DetSlitSizes$(j%, i%) = vbNullString Then DetSlitSizes$(j%, i%) = " "
msg$ = msg$ & Format$(DetSlitSizes$(j%, i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)
Next j%

' Load number of detector slit positions
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, DetSlitPositionsNumber%(i%)
If DetSlitPositionsNumber%(i%) > MAXDET% Then GoTo InitDetectorsTooManyPositions
msg$ = msg$ & Format$(DetSlitPositionsNumber%(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

For j% = 1 To MAXDET%
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, DetSlitPositions$(j%, i%)
If DetSlitPositions$(j%, i%) = vbNullString Then DetSlitPositions$(j%, i%) = " "
msg$ = msg$ & Format$(DetSlitPositions$(j%, i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)
Next j%

' Load number of detector modes
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, DetDetectorModesNumber%(i%)
If DetDetectorModesNumber%(i%) > MAXDET% Then GoTo InitDetectorsTooManyModes
msg$ = msg$ & Format$(DetDetectorModesNumber%(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

For j% = 1 To MAXDET%
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, DetDetectorModes$(j%, i%)
If DetDetectorModes$(j%, i%) = vbNullString Then DetDetectorModes$(j%, i%) = " "
msg$ = msg$ & Format$(DetDetectorModes$(j%, i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)
Next j%

' Unused parameters
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, itemp%(i%)
If itemp%(i%) > MAXDET% Then GoTo InitDetectorsTooManyUnused
msg$ = msg$ & Format$(itemp%(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

For j% = 1 To MAXDET%
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, bstring$(j%, i%)
If bstring$(j%, i%) = vbNullString Then bstring$(j%, i%) = " "
msg$ = msg$ & Format$(bstring$(j%, i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)
Next j%

' Exchange flags
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, DetSlitSizeExchangeFlags%(i%)
If DetSlitSizeExchangeFlags%(i%) < 0 Or DetSlitSizeExchangeFlags%(i%) > 2 Then GoTo InitDetectorsInvalidData
msg$ = msg$ & Format$(DetSlitSizeExchangeFlags%(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, DetSlitPositionExchangeFlags%(i%)
If DetSlitPositionExchangeFlags%(i%) < 0 Or DetSlitPositionExchangeFlags%(i%) > 2 Then GoTo InitDetectorsInvalidData
msg$ = msg$ & Format$(DetSlitPositionExchangeFlags%(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, DetDetectorModeExchangeFlags%(i%)
If DetDetectorModeExchangeFlags%(i%) < 0 Or DetDetectorModeExchangeFlags%(i%) > 2 Then GoTo InitDetectorsInvalidData
msg$ = msg$ & Format$(DetDetectorModeExchangeFlags%(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, itemp%(i%) ' unused
If itemp%(i%) < 0 Or itemp%(i%) > 2 Then GoTo InitDetectorsInvalidData
msg$ = msg$ & Format$(itemp%(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Exchange positions
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, DetSlitSizeExchangePositions!(i%)
If DetSlitSizeExchangeFlags%(i%) = 2 Then
If Not MiscMotorInBounds(i%, DetSlitSizeExchangePositions!(i%)) Then GoTo InitDetectorsBadPosition
End If
msg$ = msg$ & Format$(DetSlitSizeExchangePositions!(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, DetSlitPositionExchangePositions!(i%)
If DetSlitPositionExchangeFlags%(i%) = 2 Then
If Not MiscMotorInBounds(i%, DetSlitPositionExchangePositions!(i%)) Then GoTo InitDetectorsBadPosition
End If
msg$ = msg$ & Format$(DetSlitPositionExchangePositions!(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, DetDetectorModeExchangePositions!(i%)
If DetDetectorModeExchangeFlags%(i%) = 2 Then
If Not MiscMotorInBounds(i%, DetDetectorModeExchangePositions!(i%)) Then GoTo InitDetectorsBadPosition
End If
msg$ = msg$ & Format$(DetDetectorModeExchangePositions!(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, atemp!(i%) ' unused
'If atemp!(i%) = 2 Then
'If Not MiscMotorInBounds(i% , atemp!(i%)) Then GoTo InitDetectorsBadPosition
'End If
msg$ = msg$ & Format$(atemp!(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Exchange rowland
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, DetSlitSizeExchangeRowlands!(i%)
msg$ = msg$ & Format$(DetSlitSizeExchangeRowlands!(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, DetSlitPositionExchangeRowlands!(i%)
msg$ = msg$ & Format$(DetSlitPositionExchangeRowlands!(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, DetDetectorModeExchangeRowlands!(i%)
msg$ = msg$ & Format$(DetDetectorModeExchangeRowlands!(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, atemp!(i%) ' unused
msg$ = msg$ & Format$(atemp!(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

' Default indexes
linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, DetSlitSizeDefaultIndexes%(i%)
If DetSlitSizeDefaultIndexes%(i%) <= 0 Then GoTo InitDetectorsInvalidData
If DetSlitSizesNumber%(i%) > 0 And (DetSlitSizeDefaultIndexes%(i%) < 1 Or DetSlitSizeDefaultIndexes%(i%) > DetSlitSizesNumber%(i%)) Then GoTo InitDetectorsBadIndex
msg$ = msg$ & Format$(DetSlitSizeDefaultIndexes%(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, DetSlitPositionDefaultIndexes%(i%)
If DetSlitPositionDefaultIndexes%(i%) <= 0 Then GoTo InitDetectorsInvalidData
If DetSlitPositionsNumber%(i%) > 0 And (DetSlitPositionDefaultIndexes%(i%) < 1 Or DetSlitPositionDefaultIndexes%(i%) > DetSlitPositionsNumber%(i%)) Then GoTo InitDetectorsBadIndex
msg$ = msg$ & Format$(DetSlitPositionDefaultIndexes%(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, DetDetectorModeDefaultIndexes%(i%)
If DetDetectorModeDefaultIndexes%(i%) <= 0 Then GoTo InitDetectorsInvalidData
If DetDetectorModesNumber%(i%) > 0 And (DetDetectorModeDefaultIndexes%(i%) < 1 Or DetDetectorModeDefaultIndexes%(i%) > DetDetectorModesNumber%(i%)) Then GoTo InitDetectorsBadIndex
msg$ = msg$ & Format$(DetDetectorModeDefaultIndexes%(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

linecount% = linecount% + 1
msg$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
Input #Temp1FileNumber%, itemp%(i%) ' unused
'If itemp%(i%) < 0 Then GoTo InitDetectorsInvalidData
'If itemp%(i%) > 0 And (itemp%(i%) < 1 Or itemp%(i%) > itemp%(i%)) Then GoTo InitDetectorsBadIndex
msg$ = msg$ & Format$(itemp%(i%), a80$)
Next i%
Input #Temp1FileNumber%, comment$
If DebugMode Then Call IOWriteLog(msg$ & Space$(2) & comment$)

Exit Sub

' Errors
InitDetectorsError:
MsgBox Error$, vbOKOnly + vbCritical, "InitDetectors"
ierror = True
Exit Sub

InitDetectorsBadScalerLabels:
msg$ = "Invalid spectro labels in " & DetectorsFile$ & " (must be the same as those in " & ScalersFile$ & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "InitDetectors"
ierror = True
Exit Sub

InitDetectorsTooManySlits:
msg$ = "Too many slit sizes defined in " & DetectorsFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitDetectors"
ierror = True
Exit Sub

InitDetectorsTooManyPositions:
msg$ = "Too many slit positions defined in " & DetectorsFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitDetectors"
ierror = True
Exit Sub

InitDetectorsTooManyModes:
msg$ = "Too many detector modes defined in " & DetectorsFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitDetectors"
ierror = True
Exit Sub

InitDetectorsTooManyUnused:
msg$ = "Too many unused defined in " & DetectorsFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitDetectors"
ierror = True
Exit Sub

InitDetectorsInvalidData:
msg$ = "Invalid detector data values in " & DetectorsFile$ & " on line " & Str$(linecount%)
msg$ = msg$ & ". If you are not using this file to specify your detector configuration you can "
msg$ = msg$ & "just delete it and the program will make a new default DETECTORS.DAT file for "
msg$ = msg$ & "you the next time the program is started. Otherwise edit the file and fix the errors."
MsgBox msg$, vbOKOnly + vbExclamation, "InitDetectors"
ierror = True
Exit Sub

InitDetectorsBadPosition:
msg$ = "Position out of range in " & DetectorsFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitDetectors"
ierror = True
Exit Sub

InitDetectorsBadIndex:
msg$ = "Detector default index out of range in " & DetectorsFile$ & " on line " & Str$(linecount%)
MsgBox msg$, vbOKOnly + vbExclamation, "InitDetectors"
ierror = True
Exit Sub

End Sub

Sub InitDetectorsCreate()
' Creates a default DETECTORS.DAT file

ierror = False
On Error GoTo InitDetectorsCreateError

Dim astring As String
Dim i As Integer, j As Integer

' Create default file
astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$(VbDquote$ & ScalLabels$(i%) & VbDquote$, a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Spectro Labels" & VbDquote$

' Slit sizes
astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$("0", a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Number of Slit Sizes" & VbDquote$

For j% = 1 To MAXDET%
astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$(VbDquote$ & vbNullString & VbDquote$, a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Slit Size Strings [" & Format$(j%) & "]" & VbDquote$
Next j%

' Slit positions
astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$("0", a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Number of Slit Positions" & VbDquote$

For j% = 1 To MAXDET%
astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$(VbDquote$ & vbNullString & VbDquote$, a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Slit Position Strings [" & Format$(j%) & "]" & VbDquote$
Next j%

' Detector modes
astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$("0", a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Number of Detector Modes" & VbDquote$

For j% = 1 To MAXDET%
astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$(VbDquote$ & vbNullString & VbDquote$, a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Detector Mode Strings [" & Format$(j%) & "]" & VbDquote$
Next j%

' Unused
astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$("0", a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Number of Unused" & VbDquote$

For j% = 1 To MAXDET%
astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$(VbDquote$ & vbNullString & VbDquote$, a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Unused Strings [" & Format$(j%) & "]" & VbDquote$
Next j%

' Exchange flags
astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$("0", a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Slit Size Exchange Flags [0=none, 1=any position, 2=at position]" & VbDquote$

astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$("0", a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Slit Position Exchange Flags [0=none, 1=any position, 2=at position]" & VbDquote$

astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$("0", a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Detector Mode Exchange Flags [0=none, 1=any position, 2=at position]" & VbDquote$

astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$("0", a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Unused Exchange Flags [0=none, 1=any position, 2=at position]" & VbDquote$

' Exchange positions
astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$("0.0", a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Slit Size Exchange Positions" & VbDquote$

astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$("0.0", a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Slit Position Exchange Positions" & VbDquote$

astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$("0.0", a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Detector Mode Exchange Positions" & VbDquote$

astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$("0.0", a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Unused Exchange Positions" & VbDquote$

' Exchange Rowland circle
astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$("0.0", a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Slit Size Exchange Rowlands" & VbDquote$

astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$("0.0", a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Slit Position Exchange Rowlands" & VbDquote$

astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$("0.0", a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Detector Mode Exchange Rowlands" & VbDquote$

astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$("0.0", a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Unused Exchange Rowlands" & VbDquote$

' Default detector parameter index
astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$("1", a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Default Slit Size Indexes [1 to Number of Slit Sizes]" & VbDquote$

astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$("1", a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Default Slit Position Indexes [1 to Number of Slit Positions]" & VbDquote$

astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$("1", a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Default Detector Mode Indexes [1 to Number of Detector Modes]" & VbDquote$

astring$ = vbNullString
For i% = 1 To NumberOfTunableSpecs%
astring$ = astring$ & Format$("1", a80$)
Next i%
Print #Temp1FileNumber%, astring$ & Space$(8) & VbDquote$ & "Default Unused Indexes [1 to Number of Unused]" & VbDquote$

Exit Sub

' Errors
InitDetectorsCreateError:
MsgBox Error$, vbOKOnly + vbCritical, "InitDetectorsCreate"
ierror = True
Exit Sub

End Sub

Sub InitCoefficients(mode As Integer)
' Loads the K, L or M multiple peak coefficients (if present)
'  mode = 0  Ka-lines
'  mode = 1  Kb-lines
'  mode = 2  La-lines
'  mode = 3  Lb-lines
'  mode = 4  Ma-lines
'  mode = 5  Mb-lines

ierror = False
On Error GoTo InitCoefficientsError

Dim imot As Integer, xtal As Integer
Dim cpcMot As Integer
Dim cpcXtal As String
Dim cpcNumOf As Integer

Dim cpcElm(1 To MAXCALIBRATE%) As String, cpcXray(1 To MAXCALIBRATE%) As String
Dim cpcPosT(1 To MAXCALIBRATE%) As Single, cpcPosA(1 To MAXCALIBRATE%) As Single
Dim cpcStd(1 To MAXCALIBRATE%) As Integer
Dim cpcCoeff(1 To MAXCOEFF%) As Single

' Check if using multiple peak calibration data
If Not UseMultiplePeakCalibrationOffsetFlag Then Exit Sub

' Read Probewin.cal file for all coefficients
If Dir$(CalibratePeakCenterFiles$(mode%)) = vbNullString Then Exit Sub

' Read current calibration file
Open CalibratePeakCenterFiles$(mode%) For Input As #Temp1FileNumber%

For imot% = 1 To NumberOfTunableSpecs%
For xtal% = 1 To ScalNumberOfCrystals%(imot%)

Call WaveCalibrateReadWrite2(cpcMot%, cpcXtal$, cpcNumOf%, cpcElm$(), cpcXray$(), cpcPosT!(), cpcPosA!(), cpcStd%(), cpcCoeff!())
If ierror Then
Close (Temp1FileNumber%)
Exit Sub
End If

' Check for motor and crystal Match
If imot% = cpcMot% And MiscStringsAreSame(ScalCrystalNames$(xtal%, imot%), cpcXtal$) Then

' Load measured coefficients
MultiplePeakCoefficient1!(mode%, xtal%, imot%) = cpcCoeff!(1)
MultiplePeakCoefficient2!(mode%, xtal%, imot%) = cpcCoeff!(2)
MultiplePeakCoefficient3!(mode%, xtal%, imot%) = cpcCoeff!(3)

' Load zero for default
Else
MultiplePeakCoefficient1!(mode%, xtal%, imot%) = 0#
MultiplePeakCoefficient2!(mode%, xtal%, imot%) = 0#
MultiplePeakCoefficient3!(mode%, xtal%, imot%) = 0#
End If

Next xtal%
Next imot%

Close (Temp1FileNumber%)
Exit Sub

' Errors
InitCoefficientsError:
MsgBox Error$, vbOKOnly + vbCritical, "InitCoefficients"
Close (Temp1FileNumber%)
ierror = True
Exit Sub

End Sub

Sub InitDensity()
' Reads the DENSITY.DAT file (density defaults)

ierror = False
On Error GoTo InitDensityError

Dim inum As Integer, ip As Integer
Dim isym As String
Dim i As Integer, linecount As Integer

linecount% = 1
For i% = 1 To MAXELM%
Input #Temp1FileNumber%, inum%, isym$, AllAtomicDensities!(i%)
If inum% <> i% Then GoTo InitDensityInvalidNumber
ip% = IPOS1%(MAXELM%, isym$, Symup$())
If ip% = 0 Then GoTo InitDensityInvalidSymbol
If AllAtomicDensities!(i%) < 0.00001 Or AllAtomicDensities!(i%) > 25# Then GoTo InitDensityInvalidData
linecount% = linecount% + 1
Next i%

Exit Sub

' Errors
InitDensityError:
MsgBox Error$, vbOKOnly + vbCritical, "InitDensity"
ierror = True
Exit Sub

InitDensityInvalidNumber:
msg$ = "Invalid density number (elements must be listed in atomic number order) in " & DensityFile$ & " on line " & Str$(linecount%) & "."
msg$ = msg$ & " This error might also be caused by using non US-like operating system settings where the comma vs decimal point characters"
msg$ = msg$ & " in numerical values are not being interpreted properly."
msg$ = msg$ & " Try changing your operating system settings to US numerical values and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "InitDensity"
ierror = True
Exit Sub

InitDensityInvalidSymbol:
msg$ = "Invalid density element (" & isym$ & ") in " & DensityFile$ & " on line " & Str$(linecount%) & "."
msg$ = msg$ & " This error might also be caused by using non US-like operating system settings where the comma vs decimal point characters"
msg$ = msg$ & " in numerical values are not being interpreted properly."
msg$ = msg$ & " Try changing your operating system settings to US numerical values and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "InitDensity"
ierror = True
Exit Sub

InitDensityInvalidData:
msg$ = "Invalid density data (must be between 0.00001 and 25) in " & DensityFile$ & " on line " & Str$(linecount%) & "."
msg$ = msg$ & " This error might also be caused by using non US-like operating system settings where the comma vs decimal point characters"
msg$ = msg$ & " in numerical values are not being interpreted properly."
msg$ = msg$ & " Try changing your operating system settings to US numerical values and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "InitDensity"
ierror = True
Exit Sub

End Sub

Sub InitDensity2()
' Reads the DENSITY2.DAT file (liquid, solid densities and atomic volumes for calculation of compound densities)

ierror = False
On Error GoTo InitDensity2Error

Dim inum As Integer, ip As Integer
Dim isym As String
Dim i As Integer, linecount As Integer

linecount% = 1
For i% = 1 To MAXELM%
Input #Temp1FileNumber%, inum%, isym$, AllAtomicDensities2!(i%), AllAtomicDensities3!(i%), AllAtomicVolumes!(i%)
If inum% <> i% Then GoTo InitDensity2InvalidNumber
ip% = IPOS1%(MAXELM%, isym$, Symup$())
If ip% = 0 Then GoTo InitDensity2InvalidSymbol
If AllAtomicDensities2!(i%) < 0.00001 Or AllAtomicDensities2!(i%) > 25# Then GoTo InitDensity2InvalidData
If AllAtomicDensities3!(i%) < 0.00001 Or AllAtomicDensities3!(i%) > 25# Then GoTo InitDensity2InvalidData
If AllAtomicVolumes!(i%) < 20 Or AllAtomicVolumes!(i%) > 300# Then GoTo InitDensity2InvalidData
linecount% = linecount% + 1
Next i%

Exit Sub

' Errors
InitDensity2Error:
MsgBox Error$, vbOKOnly + vbCritical, "InitDensity2"
ierror = True
Exit Sub

InitDensity2InvalidNumber:
msg$ = "Invalid density number (elements must be listed in atomic number order) in " & DensityFile2$ & " on line " & Str$(linecount%) & "."
msg$ = msg$ & " This error might also be caused by using non US-like operating system settings where the comma vs decimal point characters in numerical values are not being interpreted properly."
msg$ = msg$ & " Try changing your operating system settings to US numerical values and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "InitDensity2"
ierror = True
Exit Sub

InitDensity2InvalidSymbol:
msg$ = "Invalid density element (" & isym$ & ") in " & DensityFile2$ & " on line " & Str$(linecount%) & "."
msg$ = msg$ & " This error might also be caused by using non US-like operating system settings where the comma vs decimal point characters in numerical values are not being interpreted properly."
msg$ = msg$ & " Try changing your operating system settings to US numerical values and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "InitDensity2"
ierror = True
Exit Sub

InitDensity2InvalidData:
msg$ = "Invalid density data (must be between 0.00001 and 25) in " & DensityFile2$ & " on line " & Str$(linecount%) & "."
msg$ = msg$ & " This error might also be caused by using non US-like operating system settings where the comma vs decimal point characters in numerical values are not being interpreted properly."
msg$ = msg$ & " Try changing your operating system settings to US numerical values and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "InitDensity2"
ierror = True
Exit Sub

End Sub

Sub InitDensityCreate()
' Create a default DENSITY.DAT file

ierror = False
On Error GoTo InitDensityCreateError

Dim i As Integer, tDensity As Single

For i% = 1 To MAXELM%
If i% = 1 Then tDensity! = 0.0000899
If i% = 2 Then tDensity! = 0.0001787
If i% = 3 Then tDensity! = 0.53
If i% = 4 Then tDensity! = 1.848
If i% = 5 Then tDensity! = 2.34
If i% = 6 Then tDensity! = 2.26
If i% = 7 Then tDensity! = 0.0012506
If i% = 8 Then tDensity! = 0.001429
If i% = 9 Then tDensity! = 0.001696
If i% = 10 Then tDensity! = 0.0009
If i% = 11 Then tDensity! = 0.971
If i% = 12 Then tDensity! = 1.738
If i% = 13 Then tDensity! = 2.702
If i% = 14 Then tDensity! = 2.33
If i% = 15 Then tDensity! = 1.82
If i% = 16 Then tDensity! = 2.07
If i% = 17 Then tDensity! = 0.003214
If i% = 18 Then tDensity! = 0.0017824
If i% = 19 Then tDensity! = 0.862
If i% = 20 Then tDensity! = 1.55
If i% = 21 Then tDensity! = 3#
If i% = 22 Then tDensity! = 4.5
If i% = 23 Then tDensity! = 5.8
If i% = 24 Then tDensity! = 7.19
If i% = 25 Then tDensity! = 7.43
If i% = 26 Then tDensity! = 7.86
If i% = 27 Then tDensity! = 8.9
If i% = 28 Then tDensity! = 8.9
If i% = 29 Then tDensity! = 8.96
If i% = 30 Then tDensity! = 7.14
If i% = 31 Then tDensity! = 5.907
If i% = 32 Then tDensity! = 5.323
If i% = 33 Then tDensity! = 5.72
If i% = 34 Then tDensity! = 4.79
If i% = 35 Then tDensity! = 3.119
If i% = 36 Then tDensity! = 0.003708
If i% = 37 Then tDensity! = 1.53
If i% = 38 Then tDensity! = 2.6
If i% = 39 Then tDensity! = 4.47
If i% = 40 Then tDensity! = 6.4
If i% = 41 Then tDensity! = 8.57
If i% = 42 Then tDensity! = 10.2
If i% = 43 Then tDensity! = 11.5
If i% = 44 Then tDensity! = 12.2
If i% = 45 Then tDensity! = 12.4
If i% = 46 Then tDensity! = 12.02
If i% = 47 Then tDensity! = 10.5
If i% = 48 Then tDensity! = 8.65
If i% = 49 Then tDensity! = 7.31
If i% = 50 Then tDensity! = 7.3
If i% = 51 Then tDensity! = 6.684
If i% = 52 Then tDensity! = 6.24
If i% = 53 Then tDensity! = 4.93
If i% = 54 Then tDensity! = 0.00588
If i% = 55 Then tDensity! = 1.837
If i% = 56 Then tDensity! = 3.51
If i% = 57 Then tDensity! = 6.7
If i% = 58 Then tDensity! = 6.78
If i% = 59 Then tDensity! = 6.77
If i% = 60 Then tDensity! = 7#
If i% = 61 Then tDensity! = 6.475
If i% = 62 Then tDensity! = 7.54
If i% = 63 Then tDensity! = 5.259
If i% = 64 Then tDensity! = 7.895
If i% = 65 Then tDensity! = 8.27
If i% = 66 Then tDensity! = 8.536
If i% = 67 Then tDensity! = 8.8
If i% = 68 Then tDensity! = 9.05
If i% = 69 Then tDensity! = 9.33
If i% = 70 Then tDensity! = 6.98
If i% = 71 Then tDensity! = 9.85
If i% = 72 Then tDensity! = 13.2
If i% = 73 Then tDensity! = 16.6
If i% = 74 Then tDensity! = 19.3
If i% = 75 Then tDensity! = 21#
If i% = 76 Then tDensity! = 22.4
If i% = 77 Then tDensity! = 22.42
If i% = 78 Then tDensity! = 21.45
If i% = 79 Then tDensity! = 19.32
If i% = 80 Then tDensity! = 13.546
If i% = 81 Then tDensity! = 11.85
If i% = 82 Then tDensity! = 11.34
If i% = 83 Then tDensity! = 9.8
If i% = 84 Then tDensity! = 9.4
If i% = 85 Then tDensity! = 10#
If i% = 86 Then tDensity! = 0.00973
If i% = 87 Then tDensity! = 10#
If i% = 88 Then tDensity! = 5#
If i% = 89 Then tDensity! = 10.07
If i% = 90 Then tDensity! = 11.7
If i% = 91 Then tDensity! = 15.4
If i% = 92 Then tDensity! = 18.9
If i% = 93 Then tDensity! = 20.45
If i% = 94 Then tDensity! = 19.8
If i% = 95 Then tDensity! = 13.6
If i% = 96 Then tDensity! = 13.5
If i% = 97 Then tDensity! = 10#
If i% = 98 Then tDensity! = 10#
If i% = 99 Then tDensity! = 10#
If i% = 100 Then tDensity! = 10#
Print #Temp1FileNumber%, AllAtomicNums%(i%), VbDquote$ & Symup$(i%) & VbDquote$, tDensity!
Next i%

Exit Sub

' Errors
InitDensityCreateError:
MsgBox Error$, vbOKOnly + vbCritical, "InitDensityCreate"
ierror = True
Exit Sub

End Sub

Sub InitScalersCheck(mode As Integer, linecount As Integer, astring As String)
' Reads the SCALERS.DAT file and checks to see if it has enough lines
' mode = 0 read and determine number of lines in file
' mode = 1 append extra lines to file (use last line saved in astring$)

ierror = False
On Error GoTo InitScalersCheckError

Dim j As Integer
Dim bstring As String

If DebugMode Then
If mode% = 0 Then Call IOWriteLog(vbCrLf & "Checking SCALERS.DAT...")
If mode% = 1 Then Call IOWriteLog(vbCrLf & "Updating SCALERS.DAT...")
End If

' Loop until EOF
If mode% = 0 Then
linecount% = 0
Do Until EOF(Temp1FileNumber%)
msg$ = vbNullString
Line Input #Temp1FileNumber%, bstring$
If Trim$(bstring$) <> vbNullString Then astring$ = bstring$
linecount% = linecount% + 1
Loop
Exit Sub
End If

' Append additional lines
If mode% = 1 Then
For j% = linecount% + 1 To SCALERSDATLINESNEW%
Print #Temp1FileNumber%, astring$
Next j%

If DebugMode Then
If mode% = 1 Then Call IOWriteLog(vbCrLf & "SCALERS.DAT updated")
End If

Exit Sub
End If

Exit Sub

' Errors
InitScalersCheckError:
MsgBox Error$, vbOKOnly + vbCritical, "InitScalersCheck"
ierror = True
Exit Sub

End Sub

Sub InitWindow(mode As Integer, userstring As String, tForm As Form)
' Open the WINDOW.INI file and write or read window positions for the user
' mode = 1 then write settings
' mode = 2 read settings

ierror = False
On Error GoTo InitWindowError

Dim valid As Long
Dim twindow(1 To 4) As Single

Dim lpAppName As String
Dim lpKeyName As String
Dim lpDefault As String
Dim lpFileName As String
Dim lpString As String
Dim lpReturnString As String * MAXPATHLENGTH%

Dim nSize As Long
Dim astring As String, tcomment As String

Dim nMonitors As Long
Dim vWidth As Long, vHeight As Long
Dim tWidth() As Long, tHeight() As Long

' Check for blank username (not Probewin.exe)
If Trim$(userstring$) = vbNullString Then
userstring$ = app.EXEName
If UCase$(app.EXEName$) = UCase$("Probewin") Then userstring$ = "Probe for EPMA"
End If

' Check for blank INI file name
If Trim$(WindowINIFile$) = vbNullString Then WindowINIFile$ = ApplicationCommonAppData$ & "WINDOW.INI"

' Load comma character for non Probewin.exe applications
VbComma$ = ChrW$(44)    ' comma character

' Use Windows API function to read WINDOW.INI
lpFileName$ = WindowINIFile$
nSize& = Len(lpReturnString$)

' Write [Window] section
If mode% = 1 Then
If tForm.WindowState = vbMinimized Then Exit Sub
If tForm.Visible = False Then Exit Sub

' Save window position
astring$ = Str$(tForm.Left) & ", " & Str$(tForm.Top) & ", " & Str$(tForm.Width) & ", " & Str$(tForm.Height)
lpAppName$ = tForm.Name
lpKeyName$ = userstring$
lpString$ = astring$
valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, lpString$, lpFileName$)
End If

' Read [Window] section
If mode% = 2 Then
lpAppName$ = tForm.Name

' Read window positions
lpKeyName$ = userstring$
lpDefault$ = vbNullString
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then astring$ = Left$(lpReturnString$, valid&)

' Parse window positions
If astring$ <> vbNullString Then
Call InitParseStringToReal(astring$, Int(4), twindow!())
If ierror Then Exit Sub

If tForm.WindowState = vbMaximized Then Exit Sub

' Check if saved window position is outside visible area (different monitor) (Screen object in VB6 does not handle dual monitors)
Call MonitorsGetVirtualExtents(nMonitors&, tWidth&(), tHeight&(), vWidth&, vHeight&)
If ierror Then Exit Sub
If (twindow!(1) < 0 Or twindow!(1) > vWidth& * Screen.TwipsPerPixelX) Or (twindow!(2) < 0 Or twindow!(2) > vHeight& * Screen.TwipsPerPixelY) Then
Call MiscCenterForm(tForm)
Exit Sub
End If

tForm.Left = twindow!(1)
tForm.Top = twindow!(2)
If tForm.Name = "FormCALCIMAGE" Then    ' MDIForm does not have borderstyle property (but is sizable)
tForm.Width = twindow!(3)
tForm.Height = twindow!(4)
Else
If tForm.BorderStyle = vbSizable Then
tForm.Width = twindow!(3)
tForm.Height = twindow!(4)
End If
End If

' Just center
Else
Call MiscCenterForm(tForm)
End If
End If

Exit Sub

' Errors
InitWindowError:
MsgBox Error$, vbOKOnly + vbCritical, "InitWindow"
ierror = True
Exit Sub

End Sub

Function InitWindowIsSpecified(userstring As String, tForm As Form) As Boolean
' Open the WINDOW.INI file and just check if there is size information

ierror = False
On Error GoTo InitWindowIsSpecifiedError

Dim valid As Long

Dim lpAppName As String
Dim lpKeyName As String
Dim lpDefault As String
Dim lpFileName As String
Dim lpReturnString As String * MAXPATHLENGTH%

Dim nSize As Long
Dim astring As String, tcomment As String

' Use Windows API function to read WINDOW.INI
lpFileName$ = WindowINIFile$
nSize& = Len(lpReturnString$)

' Check for entry
lpAppName$ = tForm.Name
lpKeyName$ = userstring$
lpDefault$ = vbNullString
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then astring$ = Left$(lpReturnString$, valid&)

' Set return if window positions are there
If astring$ <> vbNullString Then
InitWindowIsSpecified = True
Else
InitWindowIsSpecified = False
End If

Exit Function

' Errors
InitWindowIsSpecifiedError:
MsgBox Error$, vbOKOnly + vbCritical, "InitWindowIsSpecified"
ierror = True
Exit Function

End Function

Public Function InitGetWindowsDirectory() As String
' Returns the Window folder

ierror = False
On Error GoTo InitGetWindowsDirectoryError

Dim strWindowsDir As String * MAXPATHLENGTH%      ' variable to return the path of Windows Directory
Dim lngWindowsDirLength As Long    ' variable to return the the length of the path
    
lngWindowsDirLength& = GetWindowsDirectory(strWindowsDir$, Len(strWindowsDir$)) ' read the path of the windows directory
InitGetWindowsDirectory$ = Left$(strWindowsDir$, lngWindowsDirLength&) ' extract the windows path from the buffer
Exit Function

' Errors
InitGetWindowsDirectoryError:
MsgBox Error$, vbOKOnly + vbCritical, "InitGetWindowsDirectory"
ierror = True
Exit Function

End Function

Function InitGetApplicationPath() As String
' Returns the application path

ierror = False
On Error GoTo InitGetApplicationPathError

Dim lngApplicationDirLength As Long
Dim astring As String
Dim tappPath As String * 256

' Get app path
lngApplicationDirLength& = GetModuleFileName(0, tappPath$, Len(tappPath$))
If lngApplicationDirLength& = 0 Then GoTo InitGetApplicationPathFail
astring$ = Left$(tappPath$, lngApplicationDirLength&)
InitGetApplicationPath$ = MiscGetPathOnly$(astring$)
Exit Function

' Errors
InitGetApplicationPathError:
MsgBox Error$, vbOKOnly + vbCritical, "InitGetApplicationPath"
ierror = True
Exit Function

InitGetApplicationPathFail:
msg$ = "Zero length application path length (should not occur)"
MsgBox msg$, vbOKOnly + vbExclamation, "InitGetApplicationPath"
ierror = True
Exit Function

End Function

Function InitIsDriveMediaPresent(tdrive As String) As Boolean
' Returns whether the drive is present and has media present (if removable)

ierror = False
On Error GoTo InitIsDriveMediaPresentError

Dim Buffer As String

' Check if string contains a drive letter (no colon means not a drive)
InitIsDriveMediaPresent = True
If Mid$(tdrive$, 2, 1) <> ":" Then Exit Function

' Allow hard errors (to indicate missing disk)
On Error Resume Next

' Check for current directory (failure means drive is not present)
Buffer$ = CurDir$(tdrive$)
If Buffer$ = vbNullString Then
InitIsDriveMediaPresent = False
Exit Function
End If

' Try root directory (failure means media not present)
If Err.number = 0 Then
Buffer$ = Dir(Left$(tdrive$, 1) & ":\")
InitIsDriveMediaPresent = (Err.number = 0)
End If

Exit Function

' Errors
InitIsDriveMediaPresentError:
MsgBox Error$, vbOKOnly + vbCritical, "InitIsDriveMediaPresent"
ierror = True
Exit Function

End Function

Sub InitMinMax()
' Load the minimum and maximum allowed values for a number of parameters

ierror = False
On Error GoTo InitMinMaxError

' Load maximum baseline + window levels and spectro count time based on "InterfaceType"
If InterfaceType% = 0 Then      ' demo
MinPHABaselineWindow! = 0.05
MaxPHABaselineWindow! = 10#
If MiscIsInstrumentStage("CAMECA") Then MaxPHABaselineWindow! = 5.8               ' SX100 demo
MinScalerCountTime! = 0.01
MaxScalerCountTime! = 1000000#
MinPHAGainWindow! = 4#
MaxPHAGainWindow! = 128#                ' JEOL demo
If MiscIsInstrumentStage("CAMECA") Then MaxPHAGainWindow! = 4095#               ' SX100 demo
MaxPHABiasWindow! = 2000#

ElseIf InterfaceType% = 1 Then  ' Unused
MinPHABaselineWindow! = 0.05
MaxPHABaselineWindow! = 10#
MinScalerCountTime! = 0.01
MaxScalerCountTime! = 1000000#
MinPHAGainWindow! = 4#
MaxPHAGainWindow! = 128#
MaxPHABiasWindow! = 2000#

ElseIf InterfaceType% = 2 Then  ' JEOL 8200/8500, 8900, 8230/8530 (direct socket)
MinPHABaselineWindow! = 0.05
MaxPHABaselineWindow! = 10#
MinScalerCountTime! = 0.01
If JeolEOSInterfaceType& < 3 Then   ' 8900/8200/8500
MaxScalerCountTime! = 2147#
Else
MaxScalerCountTime! = 10000#        ' 8230/8530
End If
MinPHAGainWindow! = 4#
MaxPHAGainWindow! = 128#
MaxPHABiasWindow! = 2000#

ElseIf InterfaceType% = 3 Then  ' Unused
MinPHABaselineWindow! = 0.05
MaxPHABaselineWindow! = 10#
MinScalerCountTime! = 0.01
MaxScalerCountTime! = 4294#
MinPHAGainWindow! = 0#
MaxPHAGainWindow! = 64#
MaxPHABiasWindow! = 2000#

ElseIf InterfaceType% = 4 Then  ' Unused
MinPHABaselineWindow! = 0.05
MaxPHABaselineWindow! = 10#
MinScalerCountTime! = 0.1
MaxScalerCountTime! = 1000000#
MinPHAGainWindow! = 4#
MaxPHAGainWindow! = 128#
MaxPHABiasWindow! = 2000#

ElseIf InterfaceType% = 5 Then  ' SX100/SXFive
MinPHABaselineWindow! = 0.05
MaxPHABaselineWindow! = 5.6
MinScalerCountTime! = 0.01
MaxScalerCountTime! = 1000000#
MinPHAGainWindow! = 1#
MaxPHAGainWindow! = 4095#
MaxPHABiasWindow! = 2000#

ElseIf InterfaceType% = 6 Then  ' Axioscope
MinPHABaselineWindow! = 0.05
MaxPHABaselineWindow! = 10#
MinScalerCountTime! = 0.1
MaxScalerCountTime! = 600#
MinPHAGainWindow! = 1#
MaxPHAGainWindow! = 4095#
MaxPHABiasWindow! = 2000#
End If

Exit Sub

' Errors
InitMinMaxError:
MsgBox Error$, vbOKOnly + vbCritical, "InitMinMax"
ierror = True
Exit Sub

End Sub

Sub InitUserData()
' Initialize the UserData folder (only the first time the program is run)

ierror = False
On Error GoTo InitUserDataError

Dim firsttime As Boolean
Dim amsg As String

' Special call to make sure Userdata folder exists
amsg$ = "Checking UserData directory folder..."
If DebugMode Then Call IOWriteLog(amsg$)
Call InitUserDataDirectory(firsttime)
If ierror Then Exit Sub
If Not firsttime Then Exit Sub

If UCase$(app.EXEName$) <> UCase$("TestFid") Then
amsg$ = "Copying standard data files..."
If DebugMode Then Call IOWriteLog(amsg$)
FileCopy ApplicationCommonAppData$ & "DHZ.DAT", CalcZAFDATFileDirectory$ & "\DHZ.DAT"
FileCopy ApplicationCommonAppData$ & "SRM.DAT", CalcZAFDATFileDirectory$ & "\SRM.DAT"
FileCopy ApplicationCommonAppData$ & "ORE.DAT", CalcZAFDATFileDirectory$ & "\ORE.DAT"
End If

' If TestFid or demo mode copy sample data to userdata folder
If UCase$(app.EXEName$) = UCase$("Testfid") Or InterfaceType% = 0 Then
amsg$ = "Copying TestFid or demo sample position files..."
If DebugMode Then Call IOWriteLog(amsg$)
If Dir$(ApplicationCommonAppData$ & "Alkali-glass.pos") <> vbNullString Then FileCopy ApplicationCommonAppData$ & "Alkali-glass.pos", StandardPOSFileDirectory$ & "\Alkali-glass.pos"
If Dir$(ApplicationCommonAppData$ & "Johnson-metal.pos") <> vbNullString Then FileCopy ApplicationCommonAppData$ & "Johnson-metal.pos", StandardPOSFileDirectory$ & "\Johnson-metal.pos"
End If

' If demo mode then copy sample standard bitmaps to Standard POS directory
If InterfaceType% = 0 Then
amsg$ = "Copying sample standard bitmap files..."
If DebugMode Then Call IOWriteLog(amsg$)
If Dir$(ApplicationCommonAppData$ & "0395_magnetite.bmp") <> vbNullString Then FileCopy ApplicationCommonAppData$ & "0395_magnetite.bmp", StandardPOSFileDirectory$ & "\0395_magnetite.bmp"
If Dir$(ApplicationCommonAppData$ & "0012_MgO.bmp") <> vbNullString Then FileCopy ApplicationCommonAppData$ & "0012_MgO.bmp", StandardPOSFileDirectory$ & "\0012_MgO.bmp"
If Dir$(ApplicationCommonAppData$ & "0023_V2O3.bmp") <> vbNullString Then FileCopy ApplicationCommonAppData$ & "0023_V2O3.bmp", StandardPOSFileDirectory$ & "\0023_V2O3.bmp"
If Dir$(ApplicationCommonAppData$ & "0140_calcite.bmp") <> vbNullString Then FileCopy ApplicationCommonAppData$ & "0140_calcite.bmp", StandardPOSFileDirectory$ & "\0140_calcite.bmp"
If Dir$(ApplicationCommonAppData$ & "0298_VG2 glass.bmp") <> vbNullString Then FileCopy ApplicationCommonAppData$ & "0298_VG2 glass.bmp", StandardPOSFileDirectory$ & "\0298_VG2_glass.bmp"
End If

Exit Sub

' Errors
InitUserDataError:
MsgBox Error$ & ", " & amsg$, vbOKOnly + vbCritical, "InitUserData"
ierror = True
Exit Sub

End Sub

Sub InitFilesUserData()
' Update the files in user data directory (root). Only called when application first runs.

ierror = False
On Error GoTo InitFilesUserDataError

Dim amsg As String, astring As String
Dim tfilename As String, tfolder As String
Dim taskID As Long

' Check for UserData (and CalcZAFDATData and ColumnPCCData) folder and create if necessary
Call InitUserData
If ierror Then Exit Sub

' Copy updated data files to default user data directory (if not TestFid)
If UCase$(app.EXEName$) <> UCase$("TestFid") Then
amsg$ = "Copying application data files..."
If DebugMode Then Call IOWriteLog(amsg$)

amsg$ = "Copying CalcZAF (and Standard) demo data files..."
If DebugMode Then Call IOWriteLog(amsg$)

' Sample files for Standard
Call InitFilesUserData2(Int(0), "MODAL.DAT", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub

' Samples files for CalcZAF
Call InitFilesUserData2(Int(0), "CALCZAF.DAT", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "CALCZAF2.DAT", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "CALCBIN.DAT", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "NISTBIN.DAT", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "NISTBIN2.DAT", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "NISTBIN3.DAT", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "POUCHOU.DAT", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "POUCHOU2.DAT", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "AUAGCU2.DAT", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "NISTBINA20.DAT", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "NISTBINZ10.DAT", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "POUCHOUA20.DAT", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "POUCHOUZ10.DAT", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "AuCu_NBS-K-ratios.DAT", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub

' Only copy if CalcZAF or Standard
If UCase$(app.EXEName$) = UCase$("CalcZAF") Or UCase$(app.EXEName$) = UCase$("Standard") Then
Call InitFilesUserData2(Int(0), "Olivine particle-JTA-1.0um.DAT", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "Olivine particle-JTA-0.5um.DAT", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "Pouchou2_Au,Cu,Ag_only.dat", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "CaF2_Particles_20 keV.DAT", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub
End If

' Files for SF calculations
Call InitFilesUserData2(Int(0), "Wark-Watson Exper. Data (CalcZAF format)_JEOL.dat", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "Wark-Watson Exper. Data (CalcZAF format)_Cameca.dat", CalcZAFDATFileDirectory$)
If ierror Then Exit Sub

' Copy demo image files to proper folders
Call InitFilesUserData2(Int(0), "SiO2-TiO2_400um_JEOL.BMP", DemoImagesDirectoryJEOL$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "SiO2-TiO2_400um_JEOL.ACQ", DemoImagesDirectoryJEOL$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "SiO2-TiO2_400um_Cameca.BMP", DemoImagesDirectoryCameca$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "SiO2-TiO2_400um_Cameca.ACQ", DemoImagesDirectoryCameca$)
If ierror Then Exit Sub

amsg$ = "CalcZAF (and Standard) sample data files copied to " & CalcZAFDATFileDirectory$
If DebugMode Then Call IOWriteLog(amsg$)

' Copy Surfer data files to Surfer data folder
If UCase$(app.EXEName$) = UCase$("Probewin") Or UCase$(app.EXEName$) = UCase$("CalcImage") Then
amsg$ = "Copying demonstration Surfer data files..."
If DebugMode Then Call IOWriteLog(amsg$)

Call InitFilesUserData2(Int(0), "XYSCAN2.BAS", SurferDataDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "XYSCAN2.BLN", SurferDataDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "XYSCAN2.DAT", SurferDataDirectory$)
If ierror Then Exit Sub

Call InitFilesUserData2(Int(0), "XYSLICE2.BAS", SurferDataDirectory$)
If ierror Then Exit Sub

Call InitFilesUserData2(Int(0), "Montel-1_Quant_Point_Classify.DAT", SurferDataDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "Silicates-2_Quant_Image_Classify.DAT", SurferDataDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "Silicates-2_Quant_Image_Classify.INI", SurferDataDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "Silicates-2_Quant_Image_Classify.TXT", SurferDataDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "Silicates-2_00485_VS1.grd", SurferDataDirectory$)
If ierror Then Exit Sub

amsg$ = "Surfer demonstration data files copied to " & SurferDataDirectory$
If DebugMode Then Call IOWriteLog(amsg$)
End If

' Copy Grapher data files to Grapher data folder
If UCase$(app.EXEName$) = UCase$("Probewin") Then
amsg$ = "Copying demonstration Grapher data files..."
If DebugMode Then Call IOWriteLog(amsg$)
Call InitFilesUserData2(Int(0), "XYTRAV.BAS", GrapherDataDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "XYTRAV.DAT", GrapherDataDirectory$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "XYTRAV2.DAT", GrapherDataDirectory$)
If ierror Then Exit Sub

amsg$ = "Grapher demonstration data files copied to " & GrapherDataDirectory$
If DebugMode Then Call IOWriteLog(amsg$)

' Copy sample image files to demo images data folder
amsg$ = "Copying demonstration image files..."
If DebugMode Then Call IOWriteLog(amsg$)

Call InitFilesUserData2(Int(0), "DEMO2_JEOL.BMP", DemoImagesDirectoryJEOL$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "DEMO2_JEOL.JPG", DemoImagesDirectoryJEOL$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "DEMO2_JEOL.GIF", DemoImagesDirectoryJEOL$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "DEMO2_JEOL.ACQ", DemoImagesDirectoryJEOL$)
If ierror Then Exit Sub

Call InitFilesUserData2(Int(0), "DEMO2_Cameca.BMP", DemoImagesDirectoryCameca$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "DEMO2_Cameca.JPG", DemoImagesDirectoryCameca$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "DEMO2_Cameca.GIF", DemoImagesDirectoryCameca$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "DEMO2_Cameca.ACQ", DemoImagesDirectoryCameca$)
If ierror Then Exit Sub
End If

' Files used by CalcZAF
Call InitFilesUserData2(Int(0), "DEMO3_JEOL.BMP", DemoImagesDirectoryJEOL$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "DEMO3_JEOL.ACQ", DemoImagesDirectoryJEOL$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "DEMO3_Cameca.BMP", DemoImagesDirectoryCameca$)
If ierror Then Exit Sub
Call InitFilesUserData2(Int(0), "DEMO3_Cameca.ACQ", DemoImagesDirectoryCameca$)
If ierror Then Exit Sub

' Delete existing files in OriginalDemoImages (if demo images are stored)
amsg$ = "Deleting image files in " & OriginalDemoImagesDirectory$
astring$ = "CMD /C DEL /Q " & OriginalDemoImagesDirectory$ & "\*.*"
taskID& = Shell(astring$, vbMinimizedFocus)
If DebugMode Then Call IOWriteLog(amsg$)
Call MiscDelay(CDbl(RealTimeInterval!), Now)

' Copy JEOL or Cameca Demo Images to DemoImagesDirectory
If MiscIsInstrumentStage("JEOL") Then
amsg$ = "JEOL Demonstration image files being copied from " & DemoImagesDirectoryJEOL$ & " to " & OriginalDemoImagesDirectory$
astring$ = "CMD /C XCOPY /Y " & DemoImagesDirectoryJEOL$ & " " & OriginalDemoImagesDirectory$
taskID& = Shell(astring$, vbMinimizedFocus)
If DebugMode Then Call IOWriteLog(amsg$)
Else
amsg$ = "Cameca Demonstration image files being copied from " & DemoImagesDirectoryCameca$ & " to " & OriginalDemoImagesDirectory$
astring$ = "CMD /C XCOPY /Y " & DemoImagesDirectoryCameca$ & " " & OriginalDemoImagesDirectory$
taskID& = Shell(astring$, vbMinimizedFocus)
If DebugMode Then Call IOWriteLog(amsg$)
End If

' Do not check for Penepma12 files if unzip DLL is missing from Windows system folder
amsg$ = "Checking for Windows system folder..."
astring$ = MiscSystemGetWindowsSystemDirectory()
If Dir$(astring$ & "\vbuzip10.dll") <> vbNullString Then

' Always check that Penepma12 directories exists
amsg$ = "Checking for Penepma folders..."
If Not InitIsDriveMediaPresent(PENEPMA_Root$) Then  ' check if drive exists
msg$ = "The specified drive letter " & Left$(PENEPMA_Root$, 2) & " does not exist, either insert the drive and/or media (if removable) or edit the " & ProbeWinINIFile$ & " file to indicate the correct Penepma12 Root directory and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "InitFilesUserData"
End
End If
If Dir$(PENEPMA_Root$, vbDirectory) = vbNullString Then
amsg$ = "Creating Penepma folder " & PENEPMA_Root$ & "..." & vbCrLf & vbCrLf & "Please make sure the Penepma path is properly defined in " & ProbeWinINIFile$ & " and consistent with the User Data Directory Path."
MkDir PENEPMA_Root$
amsg$ = "Creating Penepma folder " & PENEPMA_Root$ & "\Penfluor" & "..." & vbCrLf & vbCrLf & "Please make sure the Penepma path is properly defined in " & ProbeWinINIFile$ & " and consistent with the User Data Directory Path."
MkDir PENEPMA_Root$ & "\Penfluor"
amsg$ = "Creating Penepma folder " & PENDBASE_Path$ & "..." & vbCrLf & vbCrLf & "Please make sure the Penepma path is properly defined in" & ProbeWinINIFile$ & " and consistent with the User Data Directory Path."
MkDir PENDBASE_Path$
amsg$ = "Creating Penepma folder " & PENEPMA_Path$ & "..." & vbCrLf & vbCrLf & "Please make sure the Penepma path is properly defined in " & ProbeWinINIFile$ & " and consistent with the User Data Directory Path."
MkDir PENEPMA_Path$
msg$ = "The Penepma12 folder " & PENEPMA_Root$ & " (and various Penepma12 sub folders) was created." & vbCrLf & vbCrLf
msg$ = msg$ & "This is utilized for Monte-Carlo modeling and correction of secondary fluorescence boundary effects." & vbCrLf & vbCrLf
msg$ = msg$ & "If you would prefer to use a different directory name for this purpose, please edit the Penepma_Root keyword in the [software] section of the " & ProbeWinINIFile$ & " file."
MsgBox msg$, vbOKOnly + vbInformation, "IniFilesUserData"

' Extract Penepma12 folders and files
tfilename$ = ApplicationCommonAppData$ & "Penepma12.zip"
tfolder$ = PENEPMA_Root$
FormUnzip.TextUnzipFile.Text = tfilename$
FormUnzip.TextUnzipFolder.Text = tfolder$

' Unzip files
'DebugMode = True    ' to see Unzip dialog
Call IOWriteLog(vbCrLf & "Extracting Penepma12 folder and files. This may take a few minutes, please wait...")
Call FormUnzip.CommandExtract_Click
If ierror Then Exit Sub

Call MiscDelay(CDbl(1), Now)
Unload FormUnzip

Call IOWriteLog(vbCrLf & "Penepma12 folders and files extracted." & vbCrLf)
DoEvents
End If
End If

' Check if Standard.exe or CalcZAF.exe or Probewin.exe for updating of Penepma files
If UCase$(app.EXEName$) = UCase$("Standard") Or UCase$(app.EXEName$) = UCase$("CalcZAF") Or UCase$(app.EXEName$) = UCase$("Probewin") Then

' Check for newer pdfiles (needed for new executables)
If Dir$(PENEPMA_Path$, vbDirectory) <> vbNullString Then
If Dir$(PENDBASE_Path$ & "\pdfiles\pdgph29.p12") <> vbNullString Then
If FileLen(PENDBASE_Path$ & "\pdfiles\pdgph29.p12") > 0 Then

' Check for Penepma folder present
If Dir$(PENEPMA_Root$ & "\Penepma", vbDirectory) <> vbNullString Then
amsg$ = "Copying updated Penepma files..."
If DebugMode Then Call IOWriteLog(amsg$)

Call InitFilesUserData2(Int(1), "penepma.exe", PENEPMA_Root$ & "\Penepma")
If ierror Then Exit Sub
Call InitFilesUserData2(Int(1), "convolg.exe", PENEPMA_Root$ & "\Penepma")
If ierror Then Exit Sub

' Copy Material files
If Dir$(PENDBASE_Path$, vbDirectory) <> vbNullString Then
amsg$ = "Copying updated Penepma Pendbase files..."
If DebugMode Then Call IOWriteLog(amsg$)

Call InitFilesUserData2(Int(1), "material.exe", PENDBASE_Path$)
If ierror Then Exit Sub
End If

' Copy Penfluor and Fitall files
If Dir$(PENEPMA_Root$ & "\Fanal", vbDirectory) <> vbNullString Then
amsg$ = "Copying updated Penepma Penfluor and Fitall files..."
If DebugMode Then Call IOWriteLog(amsg$)

Call InitFilesUserData2(Int(1), "penfluor.exe", PENEPMA_Root$ & "\Penfluor")
If ierror Then Exit Sub
Call InitFilesUserData2(Int(1), "fitall.exe", PENEPMA_Root$ & "\Penfluor")
If ierror Then Exit Sub
End If

' Copy Fanal files
If Dir$(PENEPMA_Root$ & "\Fanal", vbDirectory) <> vbNullString Then
amsg$ = "Copying updated Penepma Fanal files..."
If DebugMode Then Call IOWriteLog(amsg$)

Call InitFilesUserData2(Int(1), "fanal.exe", PENEPMA_Root$ & "\Fanal")
If ierror Then Exit Sub
End If

End If

' Newer pdfiles are not present- warn user to update
Else
msg$ = "In order to utilize the newest features for secondary fluorescence calculations you should update your Penepma12 files as soon as possible. "
msg$ = msg$ & "The Penepma, Material, Penfluor, Fitall and Fanal executables will not be updated until the new data files have been downloaded properly." & vbCrLf & vbCrLf
msg$ = msg$ & "You may download the latest Penepma12 files at this location: http://probesoftware.com/download/PENEPMA12.ZIP. "
msg$ = msg$ & "Be sure to extract the files to your Penepma12 folder, usually " & UserDataDirectory$ & "\Penepma12."
MsgBox msg$, vbOKOnly + vbInformation, "InitFilesUserData"
End If

' Newer pdfiles are not present- warn user to update
Else
msg$ = "In order to utilize the newest features for secondary fluorescence calculations you should update your Penepma12 files as soon as possible. "
msg$ = msg$ & "The Penepma, Material, Penfluor, Fitall and Fanal executables will not be updated until the new data files have been downloaded properly. " & vbCrLf & vbCrLf
msg$ = msg$ & "You may download the latest Penepma12 files at this location: http://probesoftware.com/download/PENEPMA12.ZIP. "
msg$ = msg$ & "Be sure to extract the files to your Penepma12 folder, usually " & UserDataDirectory$ & "\Penepma12."
MsgBox msg$, vbOKOnly + vbInformation, "InitFilesUserData"
End If
End If
End If

End If

Exit Sub

' Errors
InitFilesUserDataError:
MsgBox Error$ & ", " & amsg$, vbOKOnly + vbCritical, "InitFilesUserData"
ierror = True
Exit Sub

End Sub

Sub InitFilesUserData2(mode As Integer, tfilename As String, tfolder As String)
' Procedure to copy files (if needing to be updated)
'   mode = 0 warn if source file not found
'   mode = 1 do not warn if source file not found

ierror = False
On Error GoTo InitFilesUserData2Error

Dim dt1 As Variant, dt2 As Variant

' Source file is not available
If Dir$(ApplicationCommonAppData$ & tfilename$) = vbNullString Then
If mode% = 0 Then Call IOWriteLog("InitFilesUserData2: " & ApplicationCommonAppData$ & tfilename$ & " cannot be located for updating, please contact Probe Software for an updated file")

' Source file is available
Else

' First check if it needs to be updated based on file date/time
dt1 = FileDateTime(ApplicationCommonAppData$ & tfilename$)
If Dir$(tfolder$ & "\" & tfilename$) <> vbNullString Then dt2 = FileDateTime(tfolder$ & "\" & tfilename$)

' Target is older than source, so update it
If dt1 > dt2 Then

' Target is not found
If Dir$(tfolder$ & "\" & tfilename$) = vbNullString Then
FileCopy ApplicationCommonAppData$ & tfilename$, tfolder$ & "\" & tfilename$

' Target is found
Else
Kill tfolder$ & "\" & tfilename$
FileCopy ApplicationCommonAppData$ & tfilename$, tfolder$ & "\" & tfilename$
End If

' Target not copied, warn user if debug
Else
If DebugMode Then Call IOWriteLog("InitFilesUserData2: " & ApplicationCommonAppData$ & tfilename$ & " is same date or older than target file. Source file will not be copied to target for updating.")
End If
End If

Exit Sub

' Errors
InitFilesUserData2Error:
MsgBox Error$ & ", Source: " & ApplicationCommonAppData$ & tfilename$ & ", Target: " & tfolder$ & "\" & tfilename$, vbOKOnly + vbCritical, "InitFilesUserData2"
ierror = True
Exit Sub

End Sub

Sub InitMotorsUpdate(tlinecount As Integer, tmotor As Integer, tvalue As Single)
' Reads the MOTORS.DAT file and updates the specified line number and motor with the passed data value

ierror = False
On Error GoTo InitMotorsUpdateError

Dim motor As Integer, linecount As Integer, n As Integer
Dim astring As String, cstring As String
Dim bstring() As String

Dim atemp() As Single
ReDim atemp(1 To NumberOfTunableSpecs% + NumberOfStageMotors%) As Single

' Open file for input
Open MotorsFile$ For Input As #Temp1FileNumber%

' Loop until EOF
linecount% = 0
Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$
linecount% = linecount% + 1
ReDim Preserve bstring(1 To linecount%) As String
bstring$(linecount%) = astring$
Loop

' Close file and load desired line
Close #Temp1FileNumber%
astring$ = bstring$(tlinecount%)

' Remove leading and training spaces
astring$ = Trim$(astring$)

' Remove duplicate delimiters (spaces)
astring$ = Replace$(astring$, "  ", " ")
astring$ = Replace$(astring$, "  ", " ")
astring$ = Replace$(astring$, "  ", " ")
astring$ = Replace$(astring$, "  ", " ")

' Parse line into values (remainder of bstring contains comment string)
Call InitParseStringToRealDelimit(astring$, NumberOfTunableSpecs% + NumberOfStageMotors%, atemp!(), VbSpace$)
If ierror Then Exit Sub

' Store comment line in separate string (starting at first double quote)
cstring$ = Mid$(astring$, InStr(astring$, VbDquote$))

' Update specified value
atemp!(tmotor%) = tvalue!

' Convert values back to string
astring$ = vbNullString
For motor% = 1 To NumberOfTunableSpecs% + NumberOfStageMotors%
If motor% <= NumberOfTunableSpecs% Then
astring$ = astring$ & MiscAutoFormat$(atemp!(motor%)) & " "
Else
astring$ = astring$ & MiscAutoFormat$(atemp!(motor%)) & "  "
End If
Next motor%

' Open file for output
Open MotorsFile$ For Output As #Temp1FileNumber%

' Write line back to file
For n% = 1 To linecount%
If n% = tlinecount% Then
Print #Temp1FileNumber%, astring$ & cstring$
Else
Print #Temp1FileNumber%, bstring$(n%)
End If
Next n%

Close #Temp1FileNumber%

If DebugMode Then
Call IOWriteLog("InitMotorsUpdate, MOTORS.DAT updated for passed value " & Str$(tvalue!) & ", line " & Str$(tlinecount%) & ", motor " & Str$(motor%))
End If

' Confirm update
If tlinecount% = 2 Then
If tmotor% <= NumberOfTunableSpecs% Then Call IOWriteLog("Spectrometer " & Str$(tmotor%) & " low limits were updated in " & MotorsFile$)
If tmotor% = XMotor% Then Call IOWriteLog("Stage X low limits were updated in " & MotorsFile$)
If tmotor% = YMotor% Then Call IOWriteLog("Stage Y low limits were updated in " & MotorsFile$)
If tmotor% = ZMotor% Then Call IOWriteLog("Stage Z low limits were updated in " & MotorsFile$)
If tmotor% = WMotor% Then Call IOWriteLog("Stage W low limits were updated in " & MotorsFile$)
End If

If tlinecount% = 3 Then
If tmotor% <= NumberOfTunableSpecs% Then Call IOWriteLog("Spectrometer " & Str$(tmotor%) & " high limits were updated in " & MotorsFile$)
If tmotor% = XMotor% Then Call IOWriteLog("Stage X high limits were updated in " & MotorsFile$)
If tmotor% = YMotor% Then Call IOWriteLog("Stage Y high limits were updated in " & MotorsFile$)
If tmotor% = ZMotor% Then Call IOWriteLog("Stage Z high limits were updated in " & MotorsFile$)
If tmotor% = WMotor% Then Call IOWriteLog("Stage W high limits were updated in " & MotorsFile$)
End If

Exit Sub

' Errors
InitMotorsUpdateError:
MsgBox Error$, vbOKOnly + vbCritical, "InitMotorsUpdate"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub InitUserDataDirectory(firsttime As Boolean)
' Create the UserData, StandardPOSData, CalcZAFDATData and ColumnPCCData directories

ierror = False
On Error GoTo InitUserDataDirectoryError

Dim response As Integer

' Check that user data directory drive already exists
firsttime = False
If Not InitIsDriveMediaPresent(UserDataDirectory$) Then  ' check if drive exists
msg$ = "The specified drive letter " & Left$(UserDataDirectory$, 2) & " does not exist, either insert the drive and/or media (if removable) or edit the " & ProbeWinINIFile$ & " file to indicate the correct user data directory and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "InitUserDataDirectory"
End
End If

' Check that user data directory already exists
If Dir$(UserDataDirectory$, vbDirectory) = vbNullString Then
msg$ = "User Data Directory " & UserDataDirectory$ & " as specified in " & ProbeWinINIFile$ & " was not found." & vbCrLf & vbCrLf
msg$ = msg$ & "Would you like Probe for EPMA to create the folder for you?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton1, "InitUserDataDirectory")
If response% = vbYes Then
MkDir UserDataDirectory$
firsttime = True

Else
msg$ = "Please create the folder " & UserDataDirectory$ & " manually or edit the " & ProbeWinINIFile$ & " file to indicate the correct User Data Directory and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "InitUserDataDirectory"
End
End If
End If

' Check that standard position folder exists
If Not InitIsDriveMediaPresent(StandardPOSFileDirectory$) Then  ' check if drive exists
msg$ = "The specified drive letter " & Left$(StandardPOSFileDirectory$, 2) & " does not exist, either insert the drive and/or media (if removable) or edit the " & ProbeWinINIFile$ & " file to indicate the correct standard POS file directory and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "InitUserDataDirectory"
End
End If
If Dir$(StandardPOSFileDirectory$, vbDirectory) = vbNullString Then
msg$ = "Standard POS File Directory " & StandardPOSFileDirectory$ & " as specified in " & ProbeWinINIFile$ & " was not found." & vbCrLf & vbCrLf
msg$ = msg$ & "Would you like Probe for EPMA to create the folder for you?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton1, "InitUserDataDirectory")
If response% = vbYes Then
MkDir StandardPOSFileDirectory$
Else
msg$ = "Please create the folder " & StandardPOSFileDirectory$ & " manually or edit the " & ProbeWinINIFile$ & " file to indicate the correct Standard POS Data Directory and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "InitUserDataDirectory"
End
End If
End If

' Check that CalcZAFData directory already exists
If Not InitIsDriveMediaPresent(CalcZAFDATFileDirectory$) Then  ' check if drive exists
msg$ = "The specified drive letter " & Left$(CalcZAFDATFileDirectory$, 2) & " does not exist, either insert the drive and/or media (if removable) or edit the " & ProbeWinINIFile$ & " file to indicate the correct CalcZAF DAT file directory and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "InitUserDataDirectory"
End
End If
If Dir$(CalcZAFDATFileDirectory$, vbDirectory) = vbNullString Then
MkDir CalcZAFDATFileDirectory$
msg$ = "A CalcZAF DAT File Directory " & CalcZAFDATFileDirectory$ & " was created." & vbCrLf & vbCrLf
msg$ = msg$ & "If you would prefer to use a different directory name for this purpose, please edit the CalcZAFDATFileDirectory keyword in the [software] section of the " & ProbeWinINIFile$ & " file."
MsgBox msg$, vbOKOnly + vbInformation, "InitUserDataDirectory"
End If

' Check that ColumnData directory already exists
If Not InitIsDriveMediaPresent(ColumnPCCFileDirectory$) Then  ' check if drive exists
msg$ = "The specified drive letter " & Left$(ColumnPCCFileDirectory$, 2) & " does not exist, either insert the drive and/or media (if removable) or edit the " & ProbeWinINIFile$ & " file to indicate the correct Column Condition PCC file directory and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "InitUserDataDirectory"
End
End If
If Dir$(ColumnPCCFileDirectory$, vbDirectory) = vbNullString Then
MkDir ColumnPCCFileDirectory$
msg$ = "A Column PCC File Directory " & ColumnPCCFileDirectory$ & " was created." & vbCrLf & vbCrLf
msg$ = msg$ & "You will want to move your .PCC files to this folder to access them easily from now on. If you would prefer to use a different directory name for this purpose, please edit the ColumnPCCFileDirectory keyword in the [software] section of the " & ProbeWinINIFile$ & " file."
MsgBox msg$, vbOKOnly + vbInformation, "InitUserDataDirectory"
End If

' Check that SurferData directory already exists
If Not InitIsDriveMediaPresent(SurferDataDirectory$) Then  ' check if drive exists
msg$ = "The specified drive letter " & Left$(SurferDataDirectory$, 2) & " does not exist, either insert the drive and/or media (if removable) or edit the " & ProbeWinINIFile$ & " file to indicate the correct SurferData directory and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "InitUserDataDirectory"
End
End If
If Dir$(SurferDataDirectory$, vbDirectory) = vbNullString Then
MkDir SurferDataDirectory$
msg$ = "A Surfer Data Directory " & SurferDataDirectory$ & " was created." & vbCrLf & vbCrLf
msg$ = msg$ & "You will want to move your XYSCAN*.* files to this folder to access them easily from now on. If you would prefer to use a different directory name for this purpose, please edit the SurferDataDirectory keyword in the [software] section of the " & ProbeWinINIFile$ & " file."
MsgBox msg$, vbOKOnly + vbInformation, "InitUserDataDirectory"
End If

' Check that GrapherData directory already exists
If Not InitIsDriveMediaPresent(GrapherDataDirectory$) Then  ' check if drive exists
msg$ = "The specified drive letter " & Left$(GrapherDataDirectory$, 2) & " does not exist, either insert the drive and/or media (if removable) or edit the " & ProbeWinINIFile$ & " file to indicate the correct GrapherData directory and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "InitUserDataDirectory"
End
End If
If Dir$(GrapherDataDirectory$, vbDirectory) = vbNullString Then
MkDir GrapherDataDirectory$
msg$ = "A Grapher Data Directory " & GrapherDataDirectory$ & " was created." & vbCrLf & vbCrLf
msg$ = msg$ & "If you would prefer to use a different directory name for this purpose, please edit the GrapherDataDirectory keyword in the [software] section of the " & ProbeWinINIFile$ & " file."
MsgBox msg$, vbOKOnly + vbInformation, "InitUserDataDirectory"
End If

' Check that DemoImage directory already exists
If Not InitIsDriveMediaPresent(OriginalDemoImagesDirectory$) Then  ' check if drive exists
msg$ = "The specified drive letter " & Left$(OriginalDemoImagesDirectory$, 2) & " does not exist, either insert the drive and/or media (if removable) or edit the " & ProbeWinINIFile$ & " file to indicate the correct DemoImages directory and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "InitUserDataDirectory"
End
End If
If Dir$(OriginalDemoImagesDirectory$, vbDirectory) = vbNullString Then
MkDir OriginalDemoImagesDirectory$
msg$ = "A Demo Images Directory " & OriginalDemoImagesDirectory$ & " was created." & vbCrLf & vbCrLf
msg$ = msg$ & "If you would prefer to use a different folder name for this purpose, please edit the DemoImagesDirectory keyword in the [software] section of the " & ProbeWinINIFile$ & " file."
MsgBox msg$, vbOKOnly + vbInformation, "InitUserDataDirectory"
End If

' Load folder for JEOL and Cameca demo images
DemoImagesDirectoryJEOL$ = OriginalDemoImagesDirectory$ & "JEOL"
DemoImagesDirectoryCameca$ = OriginalDemoImagesDirectory$ & "Cameca"

If Dir$(DemoImagesDirectoryJEOL$, vbDirectory) = vbNullString Then MkDir DemoImagesDirectoryJEOL$
If Dir$(DemoImagesDirectoryCameca$, vbDirectory) = vbNullString Then MkDir DemoImagesDirectoryCameca$

' Check that folders were made ok
If Dir$(DemoImagesDirectoryJEOL$, vbDirectory) = vbNullString Then
msg$ = "The folder " & DemoImagesDirectoryJEOL$ & " could not be created." & vbCrLf & vbCrLf
msg$ = msg$ & "Please check that you have the necessary permissions to create a new folder on that drive."
MsgBox msg$, vbOKOnly + vbExclamation, "InitUserDataDirectory"
End
End If

If Dir$(DemoImagesDirectoryCameca$, vbDirectory) = vbNullString Then
msg$ = "The folder " & DemoImagesDirectoryCameca$ & " could not be created." & vbCrLf & vbCrLf
msg$ = msg$ & "Please check that you have the necessary permissions to create a new folder on that drive."
MsgBox msg$, vbOKOnly + vbExclamation, "InitUserDataDirectory"
End
End If

' Check that UserImages directory already exists
If Not InitIsDriveMediaPresent(UserImagesDirectory$) Then  ' check if drive exists
msg$ = "The specified drive letter " & Left$(UserImagesDirectory$, 2) & " does not exist, either insert the drive and/or media (if removable) or edit the " & ProbeWinINIFile$ & " file to indicate the correct User Images directory and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "InitUserDataDirectory"
End
End If
If Dir$(UserImagesDirectory$, vbDirectory) = vbNullString Then
MkDir UserImagesDirectory$
msg$ = "A User Images Directory " & UserImagesDirectory$ & " was created." & vbCrLf & vbCrLf
msg$ = msg$ & "You will want to store your Probe Image image and map files here. If you would prefer to use a different directory name for this purpose, please edit the UserImagesDirectory keyword in the [software] section of the " & ProbeWinINIFile$ & " file."
MsgBox msg$, vbOKOnly + vbInformation, "InitUserDataDirectory"
End If

Exit Sub

' Errors
InitUserDataDirectoryError:
msg$ = "UserDataDirectory$= " & UserDataDirectory$ & vbCrLf
msg$ = msg$ & "StandardPOSFileDirectory$= " & StandardPOSFileDirectory$ & vbCrLf
msg$ = msg$ & "CalcZAFDATFileDirectory$= " & CalcZAFDATFileDirectory$ & vbCrLf
msg$ = msg$ & "ColumnPCCFileDirectory$= " & ColumnPCCFileDirectory$ & vbCrLf
msg$ = msg$ & "SurferDataDirectory$= " & SurferDataDirectory$ & vbCrLf
msg$ = msg$ & "GrapherDataDirectory$= " & GrapherDataDirectory$ & vbCrLf
msg$ = msg$ & "DemoImagesDirectory$= " & OriginalDemoImagesDirectory$ & vbCrLf
msg$ = msg$ & "UserImagesDirectory$= " & UserImagesDirectory$ & vbCrLf
MsgBox Error$ & vbCrLf & vbCrLf & msg$, vbOKOnly + vbCritical, "InitUserDataDirectory"
ierror = True
Exit Sub

End Sub

Sub InitKratios()
' Initialize the K-ratio and alpha factor arrays for Penepma12 calculations

ierror = False
On Error GoTo InitKratiosError

' Dimension global alpha factor arrays for matrix calculations
ReDim CalcZAF_ZAF_Kratios(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Double
ReDim CalcZAF_ZA_Kratios(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Double
ReDim CalcZAF_F_Kratios(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Double

ReDim Binary_ZAF_Kratios(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Double
ReDim Binary_ZA_Kratios(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Double
ReDim Binary_F_Kratios(1 To MAXRAY% - 1, 1 To MAXBINARY%) As Double

' Dimension alpha factor arrays for boundary calculations (last index is 1 to npoints&)
ReDim Boundary_ZAF_Kratios(1 To MAXBINARY%, 1 To MAXBINARY%, 1 To 1) As Double
ReDim Boundary_ZAF_Factors(1 To MAXBINARY%, 1 To MAXBINARY%, 1 To 1) As Single

ReDim Boundary_Linear_Distances(1 To 1) As Single
ReDim Boundary_Mass_Distances(1 To MAXBINARY%, 1 To 1) As Single

ReDim Boundary_Material_A_Densities(1 To MAXBINARY%) As Single
ReDim Boundary_Material_B_Densities(1 To MAXBINARY%) As Single

ReDim Boundary_ZAF_Betas(1 To MAXBINARY%, 1 To MAXBINARY%, 1 To 1) As Single

ReDim PureGenerated_Intensities(1 To MAXRAY% - 1) As Double
ReDim PureEmitted_Intensities(1 To MAXRAY% - 1) As Double

Exit Sub

' Errors
InitKratiosError:
MsgBox Error$, vbOKOnly + vbCritical, "InitKratios"
ierror = True
Exit Sub

End Sub
