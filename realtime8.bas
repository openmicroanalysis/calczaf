Attribute VB_Name = "CodeRealTime8"
' (c) Copyright 1995-2020 by John J. Donovan
Option Explicit

Function RealTimeGetBeamScanCalibration(XorYMotor As Integer, keV As Single, mag As Single, srota As Single) As Single
' Returns the beam scan calibration in microns (and scan rotation) for the specified keV and mag
' XorYMotor = XMotor% or YMotor%

ierror = False
On Error GoTo RealTimeGetBeamScanCalibrationError

Dim i As Integer
Dim tfov As Single

Dim norder As Integer
Dim npts As Integer

ReDim acoeff(1 To MAXCOEFF%) As Single

ReDim txdata(1 To MAXBEAMCALIBRATIONS%) As Single
ReDim tydata(1 To MAXBEAMCALIBRATIONS%) As Single
ReDim tsdata(1 To MAXBEAMCALIBRATIONS%) As Single

Dim txdata2(1 To 2) As Single   ' magnification data
Dim tydata2(1 To 2) As Single   ' X or Y micron data
Dim tsdata2(1 To 2) As Single   ' scan rotation data

Static old_mag As Single

' Check for valid number of beam scan calibrations
If ImageInterfaceCalNumberOfBeamCalibrations% < 1 Then GoTo RealTimeGetBeamScanCalibrationBadNumberOfCalibrations
If ImageInterfaceCalNumberOfBeamCalibrations% > MAXBEAMCALIBRATIONS% Then GoTo RealTimeGetBeamScanCalibrationBadNumberOfCalibrations

' Check for valid X or Y motor
If XorYMotor% <> XMotor% And XorYMotor% <> YMotor% Then GoTo RealTimeGetBeamScanCalibrationBadXorYMotor

' Check for zero mag
If mag! = 0# Then GoTo RealTimeGetBeamScanCalibrationZeroMag

' Check if only one beam scan calibration
If ImageInterfaceCalNumberOfBeamCalibrations% = 1 Then
If XorYMotor% = XMotor% Then tfov! = ImageInterfaceCalXMicronsArray!(1) * ImageInterfaceCalMagArray!(1) / mag!
If XorYMotor% = YMotor% Then tfov! = ImageInterfaceCalYMicronsArray!(1) * ImageInterfaceCalMagArray!(1) / mag!
srota! = DefaultScanRotation!  ' load global as default
RealTimeGetBeamScanCalibration! = tfov!
If VerboseMode And mag! <> old_mag! Then
msg$ = "Beam scan calibration (single), passed mag= " & Format$(mag!) & ", microns= " & Format$(tfov!) & ", rota= " & Format$(srota!)
Call IOWriteLog(msg$)
End If
old_mag! = mag!
Exit Function
End If

' Check for more than one beam scan calibration (might need to do least squares fit for interpolation)
If ImageInterfaceCalNumberOfBeamCalibrations% > 1 Then

' Load values based on number of matching keV values
npts% = 0
For i% = 1 To ImageInterfaceCalNumberOfBeamCalibrations%
If keV! = ImageInterfaceCalKeVArray!(i%) Then
npts% = npts% + 1
ReDim Preserve txdata(1 To npts%) As Single     ' magnification data
ReDim Preserve tydata(1 To npts%) As Single     ' X or Y micron data
txdata!(npts%) = ImageInterfaceCalMagArray!(i%)
If XorYMotor% = XMotor% Then tydata!(npts%) = ImageInterfaceCalXMicronsArray!(i%)
If XorYMotor% = YMotor% Then tydata!(npts%) = ImageInterfaceCalYMicronsArray!(i%)
tsdata!(npts%) = ImageInterfaceCalScanRotationArray!(i%)
End If
Next i%

' Check for no matching calibrations (use default)
If npts% = 0 Then
If XorYMotor% = XMotor% Then tfov! = ImageInterfaceCalXMicronsArray!(1) * ImageInterfaceCalMagArray!(1) / mag!
If XorYMotor% = YMotor% Then tfov! = ImageInterfaceCalYMicronsArray!(1) * ImageInterfaceCalMagArray!(1) / mag!
srota! = DefaultScanRotation!  ' load global as default
RealTimeGetBeamScanCalibration! = tfov!
If VerboseMode And mag! <> old_mag! Then
msg$ = "Beam scan calibration (default), passed mag= " & Format$(mag!) & ", microns= " & Format$(tfov!) & ", rota= " & Format$(srota!)
Call IOWriteLog(msg$)
End If
old_mag! = mag!
Exit Function

' Check for one calibration (use one matching the current keV)
ElseIf npts% = 1 Then
tfov! = tydata!(1) * txdata!(1) / mag!
srota! = tsdata!(1)
RealTimeGetBeamScanCalibration! = tfov!
If VerboseMode And mag! <> old_mag! Then
msg$ = "Beam scan calibration (single keV match), passed mag= " & Format$(mag!) & ", microns= " & Format$(tfov!) & ", rota= " & Format$(srota!)
Call IOWriteLog(msg$)
End If
old_mag! = mag!
Exit Function

' Check for more than one calibration (only load closest two for linear interpolations of micron and scan rotation data)
Else
If VerboseMode And mag! <> old_mag! Then
msg$ = vbCrLf & "Number of beam scan calibrations that match passed keV (" & Format$(keV!) & "), stored in PROBEWIN.INI = " & Format$(npts%)
Call IOWriteLog(msg$)
End If

' Load mag into temp variables for fit (we only want to use the two calibrations on each side of the specified mag)
txdata2!(1) = MAXMINIMUM!
txdata2!(2) = MAXMAXIMUM!
For i% = 1 To npts%
If VerboseMode And mag! <> old_mag! Then
If XorYMotor% = XMotor% Then msg$ = "Matching X beam scan calibration " & Format$(i%) & ", mag()= " & Format$(txdata!(i%)) & ", micron()= " & Format$(tydata!(i%)) & ", rota()= " & Format$(tsdata!(i%))
If XorYMotor% = YMotor% Then msg$ = "Matching Y beam scan calibration " & Format$(i%) & ", mag()= " & Format$(txdata!(i%)) & ", micron()= " & Format$(tydata!(i%)) & ", rota()= " & Format$(tsdata!(i%))
Call IOWriteLog(msg$)
End If

If txdata!(i%) < txdata2!(1) And txdata!(i%) >= mag! Then    ' smallest value that is greater than the specified mag
txdata2!(1) = txdata!(i%)
tydata2!(1) = tydata!(i%)
tsdata2!(1) = tsdata!(i%)
End If
If txdata!(i%) > txdata2!(2) And txdata!(i%) <= mag! Then    ' largest value that is smaller than the specified mag
txdata2!(2) = txdata!(i%)
tydata2!(2) = tydata!(i%)
tsdata2!(2) = tsdata!(i%)
End If
Next i%

' Check if extrapolating on high side and if so, just use first parameter that matches the keV (tydata loaded for X or Y micron calibrations)
If txdata2!(1) = MAXMINIMUM! Or txdata2!(1) = MAXMINIMUM! Then
tfov! = tydata2!(2) * txdata2!(2) / mag!
srota! = tsdata2!(2)
If VerboseMode And mag! <> old_mag! Then
msg$ = "Extrapolating from high side, passed mag= " & Format$(mag!) & ", Extrapolated: ROTA= " & Format$(srota!) & ", FOV= " & Format$(tfov!)
Call IOWriteLog(msg$)
End If
RealTimeGetBeamScanCalibration! = tfov!
old_mag! = mag!
Exit Function
End If
End If

' Check if extrapolating on low side and if so, just use first parameter that matches the keV (tydata loaded for X or Y micron calibrations)
If txdata2!(2) = MAXMINIMUM! Or txdata2!(2) = MAXMAXIMUM! Then
tfov! = tydata2!(1) * txdata2!(1) / mag!
srota! = tsdata2!(1)
If VerboseMode And mag! <> old_mag! Then
msg$ = "Extrapolating from low side, passed mag= " & Format$(mag!) & ", Extrapolated: ROTA= " & Format$(srota!) & ", FOV= " & Format$(tfov!)
Call IOWriteLog(msg$)
End If
RealTimeGetBeamScanCalibration! = tfov!
old_mag! = mag!
Exit Function
End If
End If

' Convert to log-log fit
txdata2!(1) = Log(txdata2!(1))    ' mag
txdata2!(2) = Log(txdata2!(2))    ' mag

tydata2!(1) = Log(tydata2!(1))    ' microns
tydata2!(2) = Log(tydata2!(2))    ' microns

' Do fit of X or Y microns as a function of magnification if interpolating
norder% = 1
Call LeastSquares(norder%, Int(2), txdata2!(), tydata2!(), acoeff!())
If ierror Then Exit Function

' Calculate X or Y scan size for this mag
tfov! = acoeff!(1) + acoeff!(2) * Log(mag!)

' Take exponent of
tfov! = NATURALE# ^ tfov!

' Convert from log-log fit
txdata2!(1) = NATURALE# ^ txdata2!(1)   ' mag
txdata2!(2) = NATURALE# ^ txdata2!(2)   ' mag

tydata2!(1) = NATURALE# ^ tydata2!(1)   ' microns
tydata2!(2) = NATURALE# ^ tydata2!(2)   ' microns

' Do fit of scan rotation as a function of magnification
norder% = 1
Call LeastSquares(norder%, Int(2), txdata2!(), tsdata2!(), acoeff!())
If ierror Then Exit Function

' Calculate scan rotation for this mag
srota! = acoeff!(1) + acoeff!(2) * mag!

If VerboseMode And mag! <> old_mag! Then
msg$ = "High side matching beam scan calibration, passed mag= " & Format$(mag!) & ", mag()= " & Format$(txdata2!(1)) & ", micron()= " & Format$(tydata2!(1)) & ", rota()= " & Format$(tsdata2!(1))
Call IOWriteLog(msg$)
msg$ = "Low side matching beam scan calibration, passed mag= " & Format$(mag!) & ", mag()= " & Format$(txdata2!(2)) & ", micron()= " & Format$(tydata2!(2)) & ", rota()= " & Format$(tsdata2!(2))
Call IOWriteLog(msg$)
msg$ = "Linear fit to beam scan calibrations, ROTA= " & Format$(srota!) & ", FOV= " & Format$(tfov!)
Call IOWriteLog(msg$)
End If

RealTimeGetBeamScanCalibration! = tfov!
old_mag! = mag!
Exit Function

' Errors
RealTimeGetBeamScanCalibrationError:
MsgBox Error$, vbOKOnly + vbCritical, "RealTimeGetBeamScanCalibration"
ierror = True
Exit Function

RealTimeGetBeamScanCalibrationBadNumberOfCalibrations:
msg$ = "Invalid number of beam scan calibrations (this error should not occur, please contact Probe Software technical support with details)."
MsgBox msg$, vbOKOnly + vbExclamation, "RealTimeGetBeamScanCalibration"
ierror = True
Exit Function

RealTimeGetBeamScanCalibrationBadXorYMotor:
msg$ = "Motor " & Format$(XorYMotor%) & " is not valid stage axis motor (this error should not occur, please contact Probe Software technical support with details)."
MsgBox msg$, vbOKOnly + vbExclamation, "RealTimeGetBeamScanCalibration"
ierror = True
Exit Function

RealTimeGetBeamScanCalibrationZeroMag:
msg$ = "Passed magnification value is zero (motor= " & Format$(XorYMotor%) & ", keV= " & Format$(keV!) & ", srota= " & Format$(srota!) & "), (this error should not occur, please contact Probe Software technical support with details)."
MsgBox msg$, vbOKOnly + vbExclamation, "RealTimeGetBeamScanCalibration"
ierror = True
Exit Function

End Function


