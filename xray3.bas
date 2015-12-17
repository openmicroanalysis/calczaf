Attribute VB_Name = "CodeXRAY3"
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

Function XrayCalculatePositions(mode As Integer, motor As Integer, order As Integer, x2d As Single, k As Single, onpos As Single) As Single
' Routine to return various default spectrometer positions based on motor, 2d and on-peak position
'  mode = 0 return on peak position (based on lambda, order, x2d and crystal k)
'  mode = 1 return hi-off peak position
'  mode = 2 return lo-off peak position
'  mode = 3 return hi-wavescan position
'  mode = 4 return lo-wavescan position
'  mode = 5 return hi-peakscan position
'  mode = 6 return lo-peakscan position
'  mode = 7 return hi-quickscan position
'  mode = 8 return lo-quickscan position
'  mode = 9 return peaking start size
'  mode = 10 return peaking stop size

ierror = False
On Error GoTo XrayCalculatePositionsError

Dim pos As Single, temp As Single, factor As Single
Dim smallamount As Single

XrayCalculatePositions! = 0#
If motor% = 0 Then GoTo XrayCalculatePositionsNoMotor
If onpos! = 0# Then GoTo XrayCalculatePositionsZeroPos
If order% = 0 Then GoTo XrayCalculatePositionsZeroOrder
If x2d! = 0# Then GoTo XrayCalculatePositionsZero2d

' Load default (theoritical) on peak (onpos is lambda)
If mode% = 0 Then

' Calculation based on: sin0 = N * lambda/(2d * (1.0 - k/N^2))
pos! = order% * onpos! / MotUnitsToAngstromMicrons!(motor%) * LIF2D! / (x2d! * (1# - k! / order% ^ 2))
XrayCalculatePositions! = pos!
Exit Function
End If

' Calculate crystal/spectrometer factor for offsets and start/stop peaking sizes
factor! = (x2d! / LIF2D!) ^ 0.3333 ' cube root
If x2d! > MAXCRYSTAL2D_NOT_LDE! Then factor! = factor! * 3#  ' double for LDE analyzers (changed to triple 9-28-2006)
If x2d! > MAXCRYSTAL2D_LARGE_LDE! Then factor! = factor! * 2#  ' increase again for large 2d LDEs
temp! = MotLoLimits!(motor%) + Abs(MotHiLimits!(motor%) - MotLoLimits!(motor%))
If temp! / onpos! < 0# Then GoTo XrayCalculatePositionsNegative
factor! = factor! * Sqr(temp! / onpos!)

' Make sure range checking is enabled
NoMotorPositionBoundsChecking(motor%) = False

' Hi and Lo off-peaks offset
If mode% = 1 Or mode% = 2 Then
temp! = Abs(MotHiLimits!(motor%) - MotLoLimits!(motor%)) / ScalOffPeakFactors!(motor%)
temp! = temp! * factor! ' apply crystal/spectrometer factor

' Load hi off-peak default
If mode% = 1 Then
pos! = onpos! + temp!
If Not MiscMotorInBounds(motor%, pos!) Then
smallamount! = Abs(MotHiLimits!(motor%) - MotLoLimits!(motor%)) * SMALLAMOUNTFRACTION!     ' to place it inside the limits
pos! = MotHiLimits!(motor%) - smallamount!
End If
XrayCalculatePositions! = pos!
End If

' Load lo off-peak default
If mode% = 2 Then
pos! = onpos! - temp!
If Not MiscMotorInBounds(motor%, pos!) Then
smallamount! = Abs(MotHiLimits!(motor%) - MotLoLimits!(motor%)) * SMALLAMOUNTFRACTION!     ' to place it inside the limits
If InterfaceType% = 2 Then smallamount! = smallamount! + JEOL_SPECTRO_JOG_SIZE#        ' add for JEOL low limit spectro jog
pos! = MotLoLimits!(motor%) + smallamount!
End If
XrayCalculatePositions! = pos!
End If
Exit Function
End If


' Hi and Lo wavescan positions
If mode% = 3 Or mode% = 4 Then
temp! = Abs(MotHiLimits!(motor%) - MotLoLimits!(motor%)) / ScalWaveScanSizeFactors!(motor%)
temp! = temp! * factor! ' apply crystal/spectrometer factor

' Load hi wavescan default
If mode% = 3 Then
pos! = onpos! + temp!
If Not MiscMotorInBounds(motor%, pos!) Then
smallamount! = Abs(MotHiLimits!(motor%) - MotLoLimits!(motor%)) * SMALLAMOUNTFRACTION!     ' to place it inside the limits
pos! = MotHiLimits!(motor%) - smallamount!
End If
XrayCalculatePositions! = pos!
End If

' Load lo wavescan default
If mode% = 4 Then
pos! = onpos! - temp!
If Not MiscMotorInBounds(motor%, pos!) Then
smallamount! = Abs(MotHiLimits!(motor%) - MotLoLimits!(motor%)) * SMALLAMOUNTFRACTION!     ' to place it inside the limits
If InterfaceType% = 2 Then smallamount! = smallamount! + JEOL_SPECTRO_JOG_SIZE#        ' add for JEOL low limit spectro jog
pos! = MotLoLimits!(motor%) + smallamount!
End If
XrayCalculatePositions! = pos!
End If
Exit Function
End If


' Hi and Lo peakscan positions
If mode% = 5 Or mode% = 6 Then
temp! = Abs(MotHiLimits!(motor%) - MotLoLimits!(motor%)) / ScalPeakScanSizeFactors!(motor%)
temp! = temp! * factor! ' apply crystal/spectrometer factor

If PeakCenterPostScanFlag Then
temp! = temp! / 20#         ' as documented in help file
End If

' Load hi peakscan default
If mode% = 5 Then
pos! = onpos! + temp!
If Not MiscMotorInBounds(motor%, pos!) Then
smallamount! = Abs(MotHiLimits!(motor%) - MotLoLimits!(motor%)) * SMALLAMOUNTFRACTION!
pos! = MotHiLimits!(motor%) - smallamount!
End If
XrayCalculatePositions! = pos!
End If

' Load lo peakscan default
If mode% = 6 Then
pos! = onpos! - temp!
If Not MiscMotorInBounds(motor%, pos!) Then
smallamount! = Abs(MotHiLimits!(motor%) - MotLoLimits!(motor%)) * SMALLAMOUNTFRACTION!
If InterfaceType% = 2 Then smallamount! = smallamount! + JEOL_SPECTRO_JOG_SIZE#        ' add for JEOL low limit spectro jog
pos! = MotLoLimits!(motor%) + smallamount!
End If
XrayCalculatePositions! = pos!
End If
Exit Function
End If


' Hi and Lo quickscan positions (onpos! is ignored, base on actual limits)
If mode% = 7 Or mode% = 8 Then
temp! = Abs(MotHiLimits!(motor%) - MotLoLimits!(motor%)) * SMALLAMOUNTFRACTION!

' Load hi quickscan default
If mode% = 7 Then
pos! = MotHiLimits!(motor%) - temp!
If (InterfaceType% = 0 And MiscIsInstrumentStage("JEOL")) Or InterfaceType% = 2 Then pos! = MotHiLimits!(motor%) - 2#     ' add 2mm slack for JEOL
XrayCalculatePositions! = pos!
End If

' Load lo quickscan default
If mode% = 8 Then
pos! = MotLoLimits!(motor%) + temp!
If (InterfaceType% = 0 And MiscIsInstrumentStage("JEOL")) Or InterfaceType% = 2 Then pos! = MotLoLimits!(motor%) + 2#     ' add 2mm slack for JEOL
XrayCalculatePositions! = pos!
End If
Exit Function
End If


' Peaking start and stop size
If mode% = 9 Or mode% = 10 Then

' Load start size
If mode% = 9 Then
pos! = ScalLiFPeakingStartSizes!(motor%) * factor!
XrayCalculatePositions! = pos!
End If

' Load stop size
If mode% = 10 Then
pos! = ScalLiFPeakingStopSizes!(motor%) * factor!
XrayCalculatePositions! = pos!
End If
Exit Function
End If

Exit Function

' Errors
XrayCalculatePositionsError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayCalculatePositions"
ierror = True
Exit Function

XrayCalculatePositionsNoMotor:
msg$ = "No spectrometer number specified"
MsgBox msg$, vbOKOnly + vbExclamation, "XrayCalculatePositions"
ierror = True
Exit Function

XrayCalculatePositionsZeroPos:
msg$ = "Lambda position is zero"
If mode% > 0 Then msg$ = "OnPos position is zero on spectro " & Format$(motor%)
MsgBox msg$, vbOKOnly + vbExclamation, "XrayCalculatePositions"
ierror = True
Exit Function

XrayCalculatePositionsZeroOrder:
msg$ = "Crystal order is zero on spectro " & Format$(motor%)
MsgBox msg$, vbOKOnly + vbExclamation, "XrayCalculatePositions"
ierror = True
Exit Function

XrayCalculatePositionsZero2d:
msg$ = "Crystal 2d is zero on spectro " & Format$(motor%) & "(mode= " & Format$(mode%) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "XrayCalculatePositions"
ierror = True
Exit Function

XrayCalculatePositionsNegative:
msg$ = "Negative result prior to square root on spectro " & Format$(motor%) & ", mode " & Format$(mode%) & ", onpos " & Format$(onpos!)
MsgBox msg$, vbOKOnly + vbExclamation, "XrayCalculatePositions"
ierror = True
Exit Function

End Function

Function XrayConvert(motor As Integer, crystal As String, pos As Single, offset As Single) As Single
' Routine to convert spectrometer position of specified motor based on current crystal position

ierror = False
On Error GoTo XrayConvertError

Dim ip As Integer
Dim x2d As Single, k As Single

If motor% < 1 Or motor > NumberOfTunableSpecs% Then GoTo XrayConvertBadMotor
If pos! <= 0# Then GoTo XrayConvertBadPosition

' Check for crystal
If Trim$(crystal$) = vbNullString Then Exit Function
ip% = IPOS1(MAXCRYSTYPE%, crystal$, AllCrystalNames$())
If ip% = 0 Then Exit Function

' Convert position to angstroms: sin0 = N * lambda/(2d * (1.0 - k/N^2))
x2d! = AllCrystal2ds!(ip%)
k! = AllCrystalKs!(ip%)
XrayConvert! = (pos! + offset!) * MotUnitsToAngstromMicrons!(motor%) * (x2d! * (1# - k!)) / LIF2D!

Exit Function

' Errors
XrayConvertError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayConvert"
ierror = True
Exit Function

XrayConvertBadMotor:
msg$ = "Bad spectrometer number"
MsgBox msg$, vbOKOnly + vbExclamation, "XrayConvert"
ierror = True
Exit Function

XrayConvertBadPosition:
msg$ = "Bad spectrometer position"
MsgBox msg$, vbOKOnly + vbExclamation, "XrayConvert"
ierror = True
Exit Function

End Function

Function XrayConvertSpecAng(mode As Integer, chan As Integer, pos As Single, order As Integer, sample() As TypeSample) As Single
' Converts spectrometer to angstroms or angstroms to spectrometer
'  mode = 1  convert spectrometer to angstroms
'  mode = 2  convert angstroms to spectrometer
'  mode = 3  convert spectrometer to angstroms (w/o offset)
'  mode = 4  convert angstroms to spectrometer (w/o offset)
'  mode = 5  convert angstroms to angstroms (refraction correction only)
'  mode = 6  convert angstroms to kilovolts (with refraction correction)

ierror = False
On Error GoTo XrayConvertSpecAngError

Dim motor As Integer, ip As Integer, ipp As Integer
Dim x2d As Single, k As Single, temp As Single, onpos As Single
Dim pos1 As Single, coeff1 As Single, coeff2 As Single, coeff3 As Single
Dim offset As Single, voffset As Single, temp1 As Single, temp2 As Single

' If no valid mode or channel, just exit
If mode% = 0 Or chan% = 0 Then
If mode% = 0 Then msg$ = "Warning in XrayConvertSpecAng: mode parameter is zero"
If chan% = 0 Then msg$ = "Warning in XrayConvertSpecAng: chan parameter is zero"
Call IOWriteLog(msg$)
Exit Function
End If

' Check parameters
If pos! = 0# Then GoTo XrayConvertSpecAngZeroPos
If order% = 0 Then GoTo XrayConvertSpecAngZeroOrder
If sample(1).MotorNumbers%(chan%) = 0 Then GoTo XrayConvertSpecAngNoMotor
If sample(1).Crystal2ds!(chan%) = 0# Then GoTo XrayConvertSpecAngZero2d

' Calculate conversion factor
motor% = sample(1).MotorNumbers%(chan%)
x2d! = sample(1).Crystal2ds!(chan%)
k! = sample(1).CrystalKs!(chan%)

' Calculate conversion factor for sin0 = N * lambda/(2d * (1.0 - k/N^2))
temp! = MotUnitsToAngstromMicrons!(motor%) * (x2d! * (1# - (k! / order% ^ 2))) / LIF2D!

' If using multiple peak calibration, modify coefficients for current on-peak
If UseMultiplePeakCalibrationOffsetFlag And order% = 1 Then

' Calculate the theoritical on-peak position
Call XrayGetTheoritical(chan%, onpos!, sample())
If ierror Then Exit Function
pos1! = sample(1).OnPeaks!(chan%)   ' load default in case unable to calculate actual

' Calculate the actual position based on the theoritical peak position
ip% = IPOS1(MAXRAY% - 1, sample(1).Xrsyms$(chan%), Xraylo$())
If ip% > 0 Then
pos1! = XrayCalculateActualPosition(ip% - 1, sample(1).MotorNumbers%(chan%), sample(1).CrystalNames$(chan%), onpos!)
If ierror Then Exit Function
End If

' Calculate variable offset at the on-peak (calibrated on-peak minus polynomial calculated on-peak)
voffset! = sample(1).OnPeaks!(chan%) - pos1!

' Adjust intercept coefficient so that on-peak position gives value equal to constant offset
ipp% = MiscGetCrystalIndex%(sample(1).MotorNumbers%(chan%), sample(1).CrystalNames$(chan%))
If ipp% > 0 Then
coeff1! = MultiplePeakCoefficient1!(ip% - 1, ipp%, sample(1).MotorNumbers%(chan%)) - voffset!
coeff2! = MultiplePeakCoefficient2!(ip% - 1, ipp%, sample(1).MotorNumbers%(chan%))
coeff3! = MultiplePeakCoefficient3!(ip% - 1, ipp%, sample(1).MotorNumbers%(chan%))

' Calculate offset based on passed spectrometer position
If mode% = 1 Then
offset! = coeff1! + coeff2! * pos! + coeff3! * pos! ^ 2
Else
temp1! = pos! / temp!  ' convert angstrom to spectrometer position first
offset! = coeff1! + coeff2! * temp1! + coeff3! * temp1! ^ 2
End If
End If

If DebugMode And VerboseMode Then
msg$ = vbCrLf & "XrayConvert (variable) Calculations for: " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & ", on spec " & Str$(sample(1).MotorNumbers%(chan%)) & " " & sample(1).CrystalNames$(chan%)
Call IOWriteLog(msg$)
msg$ = "Position (on-peak): " & Str$(sample(1).OnPeaks!(chan%)) & ", position (theoritical): " & Str$(onpos!) & ", position (predicted): " & Str$(pos1!)
Call IOWriteLog(msg$)
msg$ = "Offset (constant): " & Str$(sample(1).Offsets!(chan%)) & ", offset (variable): " & Str$(voffset!)
Call IOWriteLog(msg$)
msg$ = "Intercept coefficient (original): " & Str$(MultiplePeakCoefficient1!(ip% - 1, ipp%, sample(1).MotorNumbers%(chan%))) & ", intercept (modified): " & Str$(coeff1!)
Call IOWriteLog(msg$)
msg$ = "Passed position: " & Str$(pos!) & ", calculated variable offset at passed position: " & Str$(offset!)
Call IOWriteLog(msg$)
End If
End If

' Spectrometer to angstroms
If mode% = 1 Then
If UseMultiplePeakCalibrationOffsetFlag And order% = 1 Then
temp! = (pos! + offset!) * temp!
Else
temp! = (pos! + sample(1).Offsets!(chan%)) * temp!
End If
If DebugMode And VerboseMode Then
msg$ = "Passed position (spectrometer): " & Str$(pos!) & ", returned position (angstroms): " & Str$(temp!)
Call IOWriteLog(msg$)
End If
End If

' Angstroms to spectrometer
If mode% = 2 Then
If UseMultiplePeakCalibrationOffsetFlag And order% = 1 Then
temp! = pos! / temp! - offset!
Else
temp! = pos! / temp! - sample(1).Offsets!(chan%)
End If
If DebugMode And VerboseMode Then
msg$ = "Passed position (angstrom): " & Str$(pos!) & ", returned position (spectrometer): " & Str$(temp!)
Call IOWriteLog(msg$)
End If
End If

' Spectrometer to angstroms (w/o offset)
If mode% = 3 Then
temp! = pos * temp!
End If

' Angstroms to spectrometer (w/o offset)
If mode% = 4 Then
temp! = pos! / temp!
End If

' Angstroms to angstroms (refractive order correction only)
If mode% = 5 Then
temp2! = 1# - (k - (k! / order% ^ 2))
temp! = pos! * temp2!
If DebugMode And VerboseMode Then
msg$ = "Passed position (angstroms): " & Str$(pos!) & ", returned position (angstroms): " & Str$(temp!)
Call IOWriteLog(msg$)
End If
End If

' Angstroms to kilovolts (with refractive order correction)
If mode% = 6 Then
temp2! = 1# - (k - (k! / order% ^ 2))
temp! = pos! * temp2!
temp! = order% * ANGKEV! / temp! ' convert to actual keV
If DebugMode And VerboseMode Then
msg$ = "Passed position (angstroms): " & Str$(pos!) & ", order: " & Str$(order%) & ", returned position (keV): " & Str$(temp!)
Call IOWriteLog(msg$)
End If
End If

XrayConvertSpecAng! = temp!

Exit Function

' Errors
XrayConvertSpecAngError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayConvertSpecAng"
ierror = True
Exit Function

XrayConvertSpecAngNoMotor:
msg$ = "No spectrometer number passed for channel " & Format$(chan%)
MsgBox msg$, vbOKOnly + vbExclamation, "XrayConvertSpecAng"
ierror = True
Exit Function

XrayConvertSpecAngZeroPos:
msg$ = "Lambda position is zero"
If mode% > 0 Then msg$ = "OnPos position is zero"
MsgBox msg$, vbOKOnly + vbExclamation, "XrayConvertSpecAng"
ierror = True
Exit Function

XrayConvertSpecAngZero2d:
msg$ = "Crystal 2d is zero"
MsgBox msg$, vbOKOnly + vbExclamation, "XrayConvertSpecAng"
ierror = True
Exit Function

XrayConvertSpecAngZeroOrder:
msg$ = "Bragg order is zero"
MsgBox msg$, vbOKOnly + vbExclamation, "XrayConvertSpecAng"
ierror = True
Exit Function

End Function

Sub XrayGetKevLambda(syme As String, symx As String, keV As Single, lam As Single)
' Subroutine to return xray energy and wavelength for a given xray

ierror = False
On Error GoTo XrayGetKevLambdaError

Dim ielm As Integer, iray As Integer
Dim nrec As Integer
Dim tenergy As Single

Dim engrow As TypeEnergy

' Determine the element number
ielm% = IPOS1(MAXELM%, syme$, Symlo$())
If ielm% = 0 Then GoTo XrayGetKevLambdaInvalidElement

' Check for blank x-ray (overvoltage problem)
If Trim$(symx$) = vbNullString Then symx$ = Deflin$(ielm%)

' Determine xray number (ka, kb, la, lb, ma or mb)
iray% = IPOS1(MAXRAY% - 1, symx$, Xraylo$())
If iray% = 0 Then GoTo XrayGetKevLambdaInvalidXray

' Check if original or additional x-ray line
If iray% <= MAXRAY_OLD% Then

' Read original x-ray line file
nrec% = ielm% + 2
Open XLineFile$ For Random Access Read As #XLineFileNumber% Len = XRAY_FILE_RECORD_LENGTH%
Get #XLineFileNumber%, nrec%, engrow
Close #XLineFileNumber%
tenergy! = engrow.energy!(iray%)

' Read additional x-ray line file
Else
nrec% = ielm% + 2
If Dir$(XLineFile2$) = vbNullString Then GoTo XrayGetKevLambdaNotFoundXLINE2DAT
If FileLen(XLineFile2$) = 0 Then GoTo XrayGetKevLambdaZeroSizeXLINE2DAT
Open XLineFile2$ For Random Access Read As #XLineFileNumber2% Len = XRAY_FILE_RECORD_LENGTH%
Get #XLineFileNumber2%, nrec%, engrow
Close #XLineFileNumber2%
tenergy! = engrow.energy!(iray% - MAXRAY_OLD%)
End If

' Check invalid value
If tenergy! <= 0# Then
If UCase$(app.EXEName) <> UCase$("Matrix") Then
If iray% <= MAXRAY_OLD% Then
msg$ = "No x-ray data found in " & XLineFile$ & " for " & syme$ & " " & symx$
Else
msg$ = "No x-ray data found in " & XLineFile2$ & " for " & syme$ & " " & symx$
End If
MsgBox msg$, vbOKOnly + vbExclamation, "XrayGetKevLambda"
ierror = True
Exit Sub

' Return negative number for no x-ray line data if Matrix application
Else
keV! = True
lam! = True
Exit Sub
End If
End If

' Calculate kev and angstroms and return
keV! = tenergy! / EVPERKEV#
lam! = ANGEV! / tenergy!

Exit Sub

' Errors
XrayGetKevLambdaError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayGetKevLambda"
Close #XLineFileNumber%
Close #XLineFileNumber2%
ierror = True
Exit Sub

XrayGetKevLambdaInvalidElement:
msg$ = "Element " & syme$ & " is an invalid element symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "XrayGetKevLambda"
Close #XLineFileNumber%
Close #XLineFileNumber2%
ierror = True
Exit Sub

XrayGetKevLambdaInvalidXray:
msg$ = "Xray " & symx$ & " is an invalid x-ray symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "XrayGetKevLambda"
Close #XLineFileNumber%
Close #XLineFileNumber2%
ierror = True
Exit Sub

XrayGetKevLambdaNotFoundXLINE2DAT:
msg$ = "The " & XLineFile2$ & " was not found." & vbCrLf & vbCrLf
msg$ = msg$ & "Please run the latest CalcZAF.msi installer to obtain this additional x-ray line file."
MsgBox msg$, vbOKOnly + vbExclamation, "XrayGetKevLambda"
Close #XLineFileNumber%
Close #XLineFileNumber2%
ierror = True
Exit Sub

XrayGetKevLambdaZeroSizeXLINE2DAT:
Kill XLineFile2$
msg$ = "The " & XLineFile2$ & " was not found." & vbCrLf & vbCrLf
msg$ = msg$ & "Please run the latest CalcZAF.msi installer to obtain this additional x-ray line file."
MsgBox msg$, vbOKOnly + vbExclamation, "XrayGetKevLambda"
Close #XLineFileNumber%
Close #XLineFileNumber2%
ierror = True
Exit Sub

End Sub

Sub XrayGetOffsets(mode As Integer, sample() As TypeSample)
' Calculates spectrometer offset values for each tunable element
' mode = 0 type warnings
' mode = 1 do not type warnings

ierror = False
On Error GoTo XrayGetOffsetsError

Dim i As Integer
Dim onpos As Single, temp As Single

' Loop on each element
For i% = 1 To sample(1).LastElm%
If sample(1).MotorNumbers%(i%) > 0 Then     ' skip if EDS (or old fixed spectro probe run)

' Get theoritical on-peak peak position
Call XrayGetTheoritical(i%, onpos!, sample())
If ierror Then Exit Sub

' Calculate offset
sample(1).Offsets!(i%) = onpos! - sample(1).OnPeaks!(i%)
End If
Next i%

If mode% = 1 Then Exit Sub

' Check if offset if too large and type warning to log window
For i% = 1 To sample(1).LastElm%
If sample(1).MotorNumbers%(i%) > 0 Then     ' skip if EDS (or old fixed spectro probe run)
If ScalSpecOffsetFactors!(sample(1).MotorNumbers%(i%)) > 0# Then
temp! = (MotHiLimits!(sample(1).MotorNumbers%(i%)) - MotLoLimits!(sample(1).MotorNumbers%(i%))) / ScalSpecOffsetFactors!(sample(1).MotorNumbers%(i%))
If sample(1).Crystal2ds!(i%) > MAXCRYSTAL2D_NOT_LDE! Then temp! = temp! * 3#  ' double if LDE crystal (changed to triple 02-13-2010)
If sample(1).Crystal2ds!(i%) > MAXCRYSTAL2D_LARGE_LDE! Then temp! = temp! * 2#  ' increase again for large LDEs
If Abs(sample(1).Offsets!(i%)) > Abs(temp!) Then
msg$ = "Warning: On-peak position offset for " & sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & " on spectrometer " & Str$(sample(1).MotorNumbers%(i%)) & " is " & MiscAutoFormat$(sample(1).Offsets!(i%))
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
End If
End If
End If
Next i%

' Check for bad off-peaks (if not MAN)
For i% = 1 To sample(1).LastElm%
If sample(1).BackgroundTypes%(i%) <> 1 Then  ' 0=off-peak, 1=MAN, 2=multipoint
If sample(1).MotorNumbers%(i%) > 0 Then
If ScalSpecOffsetFactors!(sample(1).MotorNumbers%(i%)) > 0# Then
temp! = (MotHiLimits!(sample(1).MotorNumbers%(i%)) - MotLoLimits!(sample(1).MotorNumbers%(i%))) / ScalSpecOffsetFactors!(sample(1).MotorNumbers%(i%))
If Abs(sample(1).HiPeaks!(i%) - sample(1).OnPeaks!(i%)) < Abs(temp!) Then
msg$ = "Warning: Hi off-peak (" & Format$(sample(1).HiPeaks!(i%)) & ") is close to on-peak position (" & Format$(sample(1).OnPeaks!(i%)) & ") for " & sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & " on spectrometer " & Str$(sample(1).MotorNumbers%(i%))
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
End If
If Abs(sample(1).LoPeaks!(i%) - sample(1).OnPeaks!(i%)) < Abs(temp!) Then
msg$ = "Warning: Lo off-peak (" & Format$(sample(1).LoPeaks!(i%)) & ") is close to on-peak position (" & Format$(sample(1).OnPeaks!(i%)) & ") for " & sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & " on spectrometer " & Str$(sample(1).MotorNumbers%(i%))
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
End If
End If
End If
End If
Next i%

' Check for large same side off-peak extrapolation (only check if linear, exponential or polynomial)
For i% = 1 To sample(1).LastElm%
If sample(1).BackgroundTypes%(i%) <> 1 Then  ' 0=off-peak, 1=MAN, 2=multipoint
If sample(1).MotorNumbers%(i%) > 0 Then
If sample(1).OffPeakCorrectionTypes%(i%) = 0 And sample(1).OffPeakCorrectionTypes%(i%) = 4 And sample(1).OffPeakCorrectionTypes%(i%) = 7 Then
If (sample(1).HiPeaks!(i%) < sample(1).OnPeaks!(i%) And sample(1).LoPeaks!(i%) < sample(1).OnPeaks!(i%)) Or (sample(1).HiPeaks!(i%) > sample(1).OnPeaks!(i%) And sample(1).LoPeaks!(i%) > sample(1).OnPeaks!(i%)) Then
temp! = sample(1).HiPeaks!(i%) - sample(1).LoPeaks!(i%)
If Abs(temp!) / Abs(sample(1).OnPeaks!(i%) - (sample(1).LoPeaks!(i%) + temp! / 2)) < 0.5 Then
msg$ = "Warning: Large same side off-peak extrapolation for " & sample(1).Elsyms$(i%) & " " & sample(1).Xrsyms$(i%) & " on spectrometer " & Str$(sample(1).MotorNumbers%(i%))
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
End If
End If
End If
End If
End If
Next i%

Exit Sub

' Errors
XrayGetOffsetsError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayGetOffsets"
ierror = True
Exit Sub

End Sub

Sub XrayCheckArXeAbsorptionEdge(chan As Integer, sample() As TypeSample)
' Checks for the Ar or Xe absorption edge and if it lies between the on-peak and either off-peak position

ierror = False
On Error GoTo XrayCheckArXeAbsorptionEdgeError

Dim ArEnergy As Single, XeEnergy As Single
Dim ArEdge As Single, XeEdge As Single
Dim LineEnergy As Single, LineLambda As Single
Dim HighEnergy As Single, LowEnergy As Single

' Check if element is using off-peak measurements (not MAN)
If sample(1).BackgroundTypes%(chan%) = 1 Then Exit Sub  ' 0=off-peak, 1=MAN, 2=multipoint

' Check Ar and Xe absorption edge energy
Call XrayGetEnergy(Int(18), Int(1), ArEnergy!, ArEdge!)     ' use K edge
If ierror Then Exit Sub
Call XrayGetEnergy(Int(54), Int(3), XeEnergy!, XeEdge!)     ' use L edge
If ierror Then Exit Sub

' Check x-ray energy
Call XrayGetKevLambda(sample(1).Elsyms$(chan%), sample(1).Xrsyms$(chan%), LineEnergy!, LineLambda!)
If ierror Then Exit Sub

' Check if line energy is below excitation energy of Ar and Xe
If LineEnergy! < ArEdge! And LineEnergy! < XeEdge! Then Exit Sub

' Check high and low peak energies
HighEnergy! = XrayConvert!(sample(1).MotorNumbers%(chan%), sample(1).CrystalNames$(chan%), sample(1).HiPeaks!(chan%), sample(1).Offsets!(chan%))
If ierror Then Exit Sub
If HighEnergy! <> 0# Then HighEnergy! = ANGKEV! / HighEnergy!

LowEnergy! = XrayConvert!(sample(1).MotorNumbers%(chan%), sample(1).CrystalNames$(chan%), sample(1).LoPeaks!(chan%), sample(1).Offsets!(chan%))
If LowEnergy! <> 0# Then LowEnergy! = ANGKEV! / LowEnergy!
If ierror Then Exit Sub

' Check for high off-peak and Ar edge (only do if not Xe detector for both Cameca and JEOL)
If Not MiscStringsAreSimilar("XPC", DetDetectorModes$(1, sample(1).MotorNumbers%(chan%))) Then
If LineEnergy! > ArEdge! And HighEnergy! < ArEdge! Then
msg$ = "Warning: Ar (K) Absorption edge (" & Format$(ArEdge!) & " keV) lies between on peak (" & Format$(LineEnergy!) & " keV) and high off-peak (" & Format$(HighEnergy!) & " keV) position for " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " on spectrometer " & Str$(sample(1).MotorNumbers%(chan%))
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
End If

If LineEnergy! < ArEdge! And HighEnergy! > ArEdge! Then
msg$ = "Warning: Ar (K) Absorption edge (" & Format$(ArEdge!) & " keV) lies between on peak (" & Format$(LineEnergy!) & " keV) and high off-peak (" & Format$(HighEnergy!) & " keV) position for " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " on spectrometer " & Str$(sample(1).MotorNumbers%(chan%))
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
End If

' Check for low off-peak and Ar edge
If LineEnergy! > ArEdge! And LowEnergy! < ArEdge! Then
msg$ = "Warning: Ar (K) Absorption edge (" & Format$(ArEdge!) & " keV) lies between on peak (" & Format$(LineEnergy!) & " keV) and low off-peak (" & Format$(LowEnergy!) & " keV) position for " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " on spectrometer " & Str$(sample(1).MotorNumbers%(chan%))
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
End If

If LineEnergy! < ArEdge! And LowEnergy! > ArEdge! Then
msg$ = "Warning: Ar (K) Absorption edge (" & Format$(ArEdge!) & " keV) lies between on peak (" & Format$(LineEnergy!) & " keV) and low off-peak (" & Format$(LowEnergy!) & " keV) position for " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " on spectrometer " & Str$(sample(1).MotorNumbers%(chan%))
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
End If
End If

' Check for high off-peak and Xe edge (only do if JEOL and Xe detector)
If MiscIsInstrumentStage("JEOL") And MiscStringsAreSimilar("XPC", DetDetectorModes$(1, sample(1).MotorNumbers%(chan%))) Then
If LineEnergy! > XeEdge! And HighEnergy! < XeEdge! Then
msg$ = "Warning: Xe (L) Absorption edge (" & Format$(XeEdge!) & " keV) lies between on peak (" & Format$(LineEnergy!) & " keV) and high off-peak (" & Format$(HighEnergy!) & " keV) position for " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " on spectrometer " & Str$(sample(1).MotorNumbers%(chan%))
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
End If

If LineEnergy! < XeEdge! And HighEnergy! > XeEdge! Then
msg$ = "Warning: Xe (L) Absorption edge (" & Format$(XeEdge!) & " keV) lies between on peak (" & Format$(LineEnergy!) & " keV) and high off-peak (" & Format$(HighEnergy!) & " keV) position for " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " on spectrometer " & Str$(sample(1).MotorNumbers%(chan%))
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
End If

' Check for low off-peak and Xe edge
If LineEnergy! > XeEdge! And LowEnergy! < XeEdge! Then
msg$ = "Warning: Xe (L) Absorption edge (" & Format$(XeEdge!) & " keV) lies between on peak (" & Format$(LineEnergy!) & " keV) and low off-peak (" & Format$(LowEnergy!) & " keV) position for " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " on spectrometer " & Str$(sample(1).MotorNumbers%(chan%))
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
End If

If LineEnergy! < XeEdge! And LowEnergy! > XeEdge! Then
msg$ = "Warning: Xe (L) Absorption edge (" & Format$(XeEdge!) & " keV) lies between on peak (" & Format$(LineEnergy!) & " keV) and low off-peak (" & Format$(LowEnergy!) & " keV) position for " & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & " on spectrometer " & Str$(sample(1).MotorNumbers%(chan%))
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
End If
End If

Exit Sub

' Errors
XrayCheckArXeAbsorptionEdgeError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayCheckArXeAbsorptionEdge"
ierror = True
Exit Sub

End Sub

Sub XrayGetEnergy(ielm As Integer, iray As Integer, energy As Single, edge As Single)
' Subroutine to return xray emission energy and edge for a given element and xray (in keV)

ierror = False
On Error GoTo XrayGetEnergyError

Dim nrec As Integer, jnum As Integer

Dim engrow As TypeEnergy, edgrow As TypeEdge

' Check
If ielm% < 1 Or ielm% > MAXELM% Then GoTo XrayGetEnergyInvalidElement
If iray% < 1 Or iray% > MAXRAY% - 1 Then GoTo XrayGetEnergyInvalidXray

' Add record location
nrec% = ielm% + 2

' Check if original or additional x-ray line
If iray% <= MAXRAY_OLD% Then

' Read original x-ray line file
Open XLineFile$ For Random Access Read As #XLineFileNumber% Len = XRAY_FILE_RECORD_LENGTH%
Get #XLineFileNumber%, nrec%, engrow
Close #XLineFileNumber%

' Calculate emission energy in kev and return
energy! = engrow.energy!(iray%) / EVPERKEV#

' Read original x-ray line file
Else
If Dir$(XLineFile2$) = vbNullString Then GoTo XrayGetEnergyNotFoundXLINE2DAT
If FileLen(XLineFile2$) = 0 Then GoTo XrayGetEnergyZeroSizeXLINE2DAT
Open XLineFile2$ For Random Access Read As #XLineFileNumber2% Len = XRAY_FILE_RECORD_LENGTH%
Get #XLineFileNumber2%, nrec%, engrow
Close #XLineFileNumber2%

' Calculate emission energy in kev and return
energy! = engrow.energy!(iray% - MAXRAY_OLD%) / EVPERKEV#
End If

' Read x-ray edge file
Open XEdgeFile$ For Random Access Read As #XEdgeFileNumber% Len = XRAY_FILE_RECORD_LENGTH%
Get #XEdgeFileNumber%, nrec%, edgrow
Close #XEdgeFileNumber%

If iray% = 1 Then jnum% = 1   ' Ka
If iray% = 2 Then jnum% = 1   ' Kb
If iray% = 3 Then jnum% = 4   ' La
If iray% = 4 Then jnum% = 3   ' Lb
If iray% = 5 Then jnum% = 9   ' Ma
If iray% = 6 Then jnum% = 8   ' Mb

If iray% = 7 Then jnum% = 3    ' Ln
If iray% = 8 Then jnum% = 3    ' Lg
If iray% = 9 Then jnum% = 3    ' Lv
If iray% = 10 Then jnum% = 4   ' Ll
If iray% = 11 Then jnum% = 7   ' Mg
If iray% = 12 Then jnum% = 9   ' Mz

' Load edge energy (in keV)
edge! = edgrow.energy!(jnum%) / EVPERKEV#
Exit Sub

' Errors
XrayGetEnergyError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayGetEnergy"
Close #XLineFileNumber%
Close #XEdgeFileNumber%
Close #XLineFileNumber2%
ierror = True
Exit Sub

XrayGetEnergyInvalidElement:
msg$ = "Invalid element number"
MsgBox msg$, vbOKOnly + vbExclamation, "XrayGetEnergy"
Close #XLineFileNumber%
Close #XEdgeFileNumber%
Close #XLineFileNumber2%
ierror = True
Exit Sub

XrayGetEnergyInvalidXray:
msg$ = "Invalid xray number"
MsgBox msg$, vbOKOnly + vbExclamation, "XrayGetEnergy"
Close #XLineFileNumber%
Close #XEdgeFileNumber%
Close #XLineFileNumber2%
ierror = True
Exit Sub

XrayGetEnergyNotFoundXLINE2DAT:
msg$ = "The " & XLineFile2$ & " was not found." & vbCrLf & vbCrLf
msg$ = msg$ & "Please run the latest CalcZAF.msi installer to obtain this additional x-ray line file."
MsgBox msg$, vbOKOnly + vbExclamation, "XrayGetEnergy"
Close #XEdgeFileNumber%
Close #XLineFileNumber%
Close #XLineFileNumber2%
ierror = True
Exit Sub

XrayGetEnergyZeroSizeXLINE2DAT:
Kill XLineFile2$
msg$ = "The " & XLineFile2$ & " was not found." & vbCrLf & vbCrLf
msg$ = msg$ & "Please run the latest CalcZAF.msi installer to obtain this additional x-ray line file."
MsgBox msg$, vbOKOnly + vbExclamation, "XrayGetEnergy"
Close #XEdgeFileNumber%
Close #XLineFileNumber%
Close #XLineFileNumber2%
ierror = True
Exit Sub

End Sub

Function XrayCalculateActualPosition(mode As Integer, motor As Integer, crystal As String, pos As Single) As Single
' Calculates spectrometer offset at this position for this spectrometer/crystal
'  mode = 0  Ka-lines
'  mode = 1  Kb-lines
'  mode = 2  La-lines
'  mode = 3  Lb-lines
'  mode = 4  Ma-lines
'  mode = 5  Mb-lines

ierror = False
On Error GoTo XrayCalculateActualPositionError

Dim ip As Integer
Dim offset As Single

XrayCalculateActualPosition! = pos! ' default to current position

' Check if using multiple peak calibration data
If Not UseMultiplePeakCalibrationOffsetFlag Then Exit Function

' Check for valid motor
If motor% < 1 Or motor% > NumberOfTunableSpecs% Then Exit Function

' Find crystal row number
ip% = MiscGetCrystalIndex%(motor%, crystal$)
If ip% = 0 Then Exit Function

' Check for valid coefficients, exit if all are zero
If MultiplePeakCoefficient1!(mode%, ip%, motor%) = 0# And MultiplePeakCoefficient2!(mode%, ip%, motor%) = 0# And MultiplePeakCoefficient3!(mode%, ip%, motor%) = 0# Then Exit Function

' Calculate spectrometer position based on variable offset
offset! = MultiplePeakCoefficient1!(mode%, ip%, motor%) + MultiplePeakCoefficient2!(mode%, ip%, motor%) * pos! + MultiplePeakCoefficient3!(mode%, ip%, motor%) * pos! ^ 2
XrayCalculateActualPosition! = pos! - offset!

Exit Function

' Errors
XrayCalculateActualPositionError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayCalculateActualPosition"
ierror = True
Exit Function

End Function

Sub XrayGetTheoritical(chan As Integer, onpos As Single, sample() As TypeSample)
' Returns the theoritical on-peak position for the given channel

ierror = False
On Error GoTo XrayGetTheoriticalError

Dim motor As Integer, order As Integer
Dim syme As String, symx As String
Dim x2d As Single, k As Single
Dim keV As Single, lam As Single

' Determine energy and wavelength based on element and xray
syme$ = sample(1).Elsyms$(chan%)
symx$ = sample(1).Xrsyms$(chan%)
Call XrayGetKevLambda(syme$, symx$, keV!, lam!)
If ierror Then Exit Sub

' First convert angstroms to theoritical spectrometer position
motor% = sample(1).MotorNumbers%(chan%)
x2d! = sample(1).Crystal2ds!(chan%)
k! = sample(1).CrystalKs!(chan%)
order% = sample(1).BraggOrders%(chan%)

' Load default
onpos! = sample(1).OnPeaks!(chan%)

' Check for division by zero (only possible when transferring data files to another system)
If MotUnitsToAngstromMicrons!(motor%) <> 0# And x2d! <> 0# And order% <> 0 Then

' Using Bragg expression: sin0 = N * lambda/(2d * (1.0 - k/N^2))
onpos! = order% * lam! / MotUnitsToAngstromMicrons!(motor%) * LIF2D! / (x2d! * (1# - k! / order% ^ 2))
End If

Exit Sub

' Errors
XrayGetTheoriticalError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayGetTheoritical"
ierror = True
Exit Sub

End Sub

Function XrayCheckOffset(chan As Integer, sample() As TypeSample) As Integer
' Calculates spectrometer offset values for and check if it exceeds the tolerance
' returns false if offset is smaller than tolerance returns true if offset exceeds tolerance

ierror = False
On Error GoTo XrayCheckOffsetError

Dim onpos As Single, temp As Single

' Check for spectrometer
XrayCheckOffset = False
If chan% = 0 Then Exit Function

' Get theoritical on-peak peak position
Call XrayGetTheoritical(chan%, onpos!, sample())
If ierror Then Exit Function

' Calculate offset
sample(1).Offsets!(chan%) = onpos! - sample(1).OnPeaks!(chan%)

' Check if offset if too large
If ScalSpecOffsetFactors!(sample(1).MotorNumbers%(chan%)) > 0# Then
temp! = (MotHiLimits!(sample(1).MotorNumbers%(chan%)) - MotLoLimits!(sample(1).MotorNumbers%(chan%))) / ScalSpecOffsetFactors!(sample(1).MotorNumbers%(chan%))
If sample(1).Crystal2ds!(chan%) > MAXCRYSTAL2D_NOT_LDE! Then temp! = temp! * 3#  ' double if LDE crystal (changed to triple 02-13-2010)
If sample(1).Crystal2ds!(chan%) > MAXCRYSTAL2D_LARGE_LDE! Then temp! = temp! * 2#  ' increase again for large LDEs
If Abs(sample(1).Offsets!(chan%)) > Abs(temp!) Then XrayCheckOffset = True
End If

Exit Function

' Errors
XrayCheckOffsetError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayCheckOffset"
ierror = True
Exit Function

End Function

Function XrayCheckOffset2(motor As Integer, oldpos As Single, newpos As Single) As Integer
' Calculates new spectrometer exceeds the tolerance (relative to old spectometer position)
' returns false if offset is smaller than tolerance
' returns true if offset exceeds tolerance

ierror = False
On Error GoTo XrayCheckOffset2Error

Dim offset As Single, temp As Single
Dim x2d As Single, k As Single
Dim telm As String, tray As String

XrayCheckOffset2 = False
If motor% < 1 Or motor% > MAXSPEC% Then Exit Function

' Get 2d for the current crystal
Call MiscGetCrystalParameters(RealTimeCrystalPositions$(motor%), x2d!, k!, telm$, tray$)
If ierror Then Exit Function

' Calculate offset
offset! = oldpos! - newpos!

' Check if offset if too large (use twice normal warning offset)
If ScalSpecOffsetFactors!(motor%) > 0# Then
temp! = 2# * (MotHiLimits!(motor%) - MotLoLimits!(motor%)) / ScalSpecOffsetFactors!(motor%)
If x2d! > MAXCRYSTAL2D_NOT_LDE! Then temp! = temp! * 3#  ' double if LDE crystal (changed to triple 02-13-2010)
If x2d! > MAXCRYSTAL2D_LARGE_LDE! Then temp! = temp! * 2#  ' increase again for large LDEs
If Abs(offset!) > Abs(temp!) Then XrayCheckOffset2 = True
End If

Exit Function

' Errors
XrayCheckOffset2Error:
MsgBox Error$, vbOKOnly + vbCritical, "XrayCheckOffset2"
ierror = True
Exit Function

End Function

Function XrayCheckOffset3(chan As Integer, factor As Single, sample() As TypeSample) As Integer
' Calculates spectrometer offset values for and check if it exceeds the tolerance times the factor.
' returns false if offset is smaller than tolerance
' returns true if offset exceeds tolerance

ierror = False
On Error GoTo XrayCheckOffset3Error

Dim onpos As Single, temp As Single

' Check for spectrometer
XrayCheckOffset3 = False
If chan% = 0 Then Exit Function

' Get theoritical on-peak peak position
Call XrayGetTheoritical(chan%, onpos!, sample())
If ierror Then Exit Function

' Calculate offset
sample(1).Offsets!(chan%) = onpos! - sample(1).OnPeaks!(chan%)

' Check if offset if too large
If ScalSpecOffsetFactors!(sample(1).MotorNumbers%(chan%)) > 0# Then
temp! = (MotHiLimits!(sample(1).MotorNumbers%(chan%)) - MotLoLimits!(sample(1).MotorNumbers%(chan%))) / ScalSpecOffsetFactors!(sample(1).MotorNumbers%(chan%))
If sample(1).Crystal2ds!(chan%) > MAXCRYSTAL2D_NOT_LDE! Then temp! = temp! * 3#  ' double if LDE crystal (changed to triple 02-13-2010)
If sample(1).Crystal2ds!(chan%) > MAXCRYSTAL2D_LARGE_LDE! Then temp! = temp! * 2#  ' increase again for large LDEs
If Abs(sample(1).Offsets!(chan%)) > Abs(temp! * factor!) Then XrayCheckOffset3 = True
End If

Exit Function

' Errors
XrayCheckOffset3Error:
MsgBox Error$, vbOKOnly + vbCritical, "XrayCheckOffset3"
ierror = True
Exit Function

End Function

Function XrayGetOffset(imot As Integer, esym As String, xsym As String, csym As String, currentpos As Single) As Integer
' Check the spectrometer offset for specified element, xray, crystal and motor (returns true if offset is greater than allowed)

ierror = False
On Error GoTo XrayGetOffsetError

Dim onpos As Single, temp As Single, offset As Single
Dim x2d As Single, k As Single
Dim keV As Single, lam As Single
Dim telm As String, tray As String

' Check for spectrometer
XrayGetOffset = False

' Determine energy and wavelength based on element and xray
Call XrayGetKevLambda(esym$, xsym$, keV!, lam!)
If ierror Then Exit Function

' Get crystal info
Call MiscGetCrystalParameters(csym$, x2d!, k!, telm$, tray$)
If ierror Then Exit Function

' Get ideal on-peak position
onpos! = XrayCalculatePositions!(Int(0), imot%, Int(1), x2d!, k!, lam!)
If ierror Then Exit Function

' Calculate offset from theoretical
offset! = onpos! - currentpos!

' Check if offset is too large
If ScalSpecOffsetFactors!(imot%) > 0# Then
temp! = (MotHiLimits!(imot%) - MotLoLimits!(imot%)) / ScalSpecOffsetFactors!(imot%)
If x2d! > MAXCRYSTAL2D_NOT_LDE! Then temp! = temp! * 3#  ' double if LDE crystal (changed to triple 02-13-2010)
If x2d! > MAXCRYSTAL2D_LARGE_LDE! Then temp! = temp! * 2#  ' increase again for large LDEs
If Abs(offset!) > Abs(temp!) Then XrayGetOffset = True
End If

Exit Function

' Errors
XrayGetOffsetError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayGetOffset"
ierror = True
Exit Function

End Function
