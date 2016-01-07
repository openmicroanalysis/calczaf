Attribute VB_Name = "CodeINIT1"
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

Sub InitINI()
' Reads the INI file

ierror = False
On Error GoTo InitINIError

Call InitINIGeneral
If ierror Then Exit Sub
Call InitINISoftware
If ierror Then Exit Sub
Call InitINIHardware
If ierror Then Exit Sub
Call InitINIHardware2
If ierror Then Exit Sub
Call InitINIImage
If ierror Then Exit Sub
Call InitINICounting
If ierror Then Exit Sub
Call InitINIFaraday
If ierror Then Exit Sub
Call InitINIPHA
If ierror Then Exit Sub
Call InitINIPlot
If ierror Then Exit Sub
Call InitINIStandards
If ierror Then Exit Sub
Call InitINISerial
If ierror Then Exit Sub

Exit Sub

' Errors
InitINIError:
MsgBox Error$, vbOKOnly + vbCritical, "InitINI"
ierror = True
Exit Sub

End Sub

Function InitGetINIData(lpFileName As String, lpAppName As String, lpKeyName As String, lpDefault As String) As String
' Returns a single INI data string

ierror = False
On Error GoTo InitGetINIDataError

Dim valid As Long
Dim lpReturnString As String * 255

Dim tcomment As String
Dim nSize As Long

' Check for existing INI file
If Dir$(lpFileName$) = vbNullString Then GoTo InitGetINIDataMissingINI
nSize& = Len(lpReturnString$)

' Get value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
InitGetINIData$ = Left$(lpReturnString$, valid&)
Exit Function

' Errors
InitGetINIDataError:
MsgBox Error$, vbOKOnly + vbCritical, "InitGetINIData"
ierror = True
Exit Function

InitGetINIDataMissingINI:
msg$ = "Unable to open file " & lpFileName$
MsgBox msg$, vbOKOnly + vbExclamation, "InitGetINIData"
ierror = True
Exit Function

End Function

Sub InitINI1()
' Load defaults for the Monitor list boxes (8200/8500 only)

ierror = False
On Error GoTo InitINI1Error

Dim n As Integer
Dim astring As String

For n% = 1 To MAXMONITOR%
If n% = 1 Then ScanComboLabels$(n%) = "Image Source"
If n% = 2 Then ScanComboLabels$(n%) = "Scan Mode"
If n% = 3 Then ScanComboLabels$(n%) = "Scan Speed"
If n% = 4 Then ScanComboLabels$(n%) = "EOS Mode"

If n% = 1 Then ScanComboCommands$(n%) = "IMS"   ' not used by 8200
If n% = 2 Then ScanComboCommands$(n%) = "SM"   ' not used by 8200
If n% = 3 Then ScanComboCommands$(n%) = "SS"   ' not used by 8200
If n% = 4 Then ScanComboCommands$(n%) = "EM"   ' not used by 8200

If n% = 1 Then ScanComboNumberOf%(n%) = 7
If n% = 2 Then ScanComboNumberOf%(n%) = 5
If n% = 3 Then ScanComboNumberOf%(n%) = 12
If n% = 4 Then ScanComboNumberOf%(n%) = 5

If n% = 1 Then astring$ = "SEI,BSE,TOPO,EDS,AUX1,AUX2,CL"
If n% = 2 Then astring$ = "PICT,BUP,LSP,SPOT,AREA"
If n% = 3 Then astring$ = "SR1,SR2,SR3,SR4,SR5,SR6,SR7,SR8,SR9,SR10,SR11,SR12"
If n% = 4 Then astring$ = "NOR,LDF,MDF,ECP,EMP"
Call InitParseStringToString2(astring$, n%, ScanComboNumberOf%(n%), ScanComboNames$())
If ierror Then Exit Sub

If n% = 1 Then astring$ = "2,3,4,5,6,7,8"   ' actual-parameters
If n% = 2 Then astring$ = "1,2,3,4,5"   ' pseudo-parameters
If n% = 3 Then astring$ = "0,2,6,14,30,62,173,350,735,1125,2254,3385"   ' actual parameters
If n% = 4 Then astring$ = "0,1,2,3,7"   ' actual parameters
Call InitParseStringToString2(astring$, n%, ScanComboNumberOf%(n%), ScanComboParameters$())
If ierror Then Exit Sub
Next n%

Exit Sub

' Errors
InitINI1Error:
MsgBox Error$, vbOKOnly + vbCritical, "InitINI1"
ierror = True
Exit Sub

End Sub

Sub InitINI2()
' Open the PROBEWIN.INI file and read defaults for the Monitor list boxes (8900 only)

ierror = False
On Error GoTo InitINI2Error

Dim n As Integer
Dim valid As Long

Dim lpAppName As String
Dim lpKeyName As String
Dim lpDefault As String
Dim lpFileName As String
Dim lpReturnString As String * 255

Dim nSize As Long
Dim nDefault As Long
Dim astring As String, tcomment As String

' Check for existing PROBEWIN.INI
If Dir$(ProbeWinINIFile$) = vbNullString Then
msg$ = "Unable to open file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINI2"
End
End If

' Use Windows API function to read PROBEWIN.INI
lpFileName$ = ProbeWinINIFile$
nSize& = Len(lpReturnString$)

' Read [Monitor] section, first number of labels
For n% = 1 To MAXMONITOR% ' number of drop-down menus

lpAppName$ = "Monitor"
lpKeyName$ = "ScanComboLabel" & Format$(n%)
If n% = 1 Then lpDefault$ = "Image Source"
If n% = 2 Then lpDefault$ = "Scan Mode"
If n% = 3 Then lpDefault$ = "Scan Speed"
If n% = 4 Then lpDefault$ = "EOS Mode"
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then ScanComboLabels$(n%) = Left$(lpReturnString$, valid&)

lpAppName$ = "Monitor"
lpKeyName$ = "ScanComboCommand" & Format$(n%)
If n% = 1 Then lpDefault$ = "IMS"
If n% = 2 Then lpDefault$ = "SM"
If n% = 3 Then lpDefault$ = "SS"
If n% = 4 Then lpDefault$ = "EM"
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then ScanComboCommands$(n%) = Left$(lpReturnString$, valid&)

lpAppName$ = "Monitor"
lpKeyName$ = "ScanComboNumberOf" & Format$(n%)
If n% = 1 Then nDefault& = 10
If n% = 2 Then nDefault& = 4
If n% = 3 Then nDefault& = 5
If n% = 4 Then nDefault& = 4
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
ScanComboNumberOf%(n%) = valid&
If ScanComboNumberOf%(n%) < 1 Or ScanComboNumberOf%(n%) > 20 Then
msg$ = "ScanComboNumberOf keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINI2"
ScanComboNumberOf%(n%) = nDefault&
End If

' Only load if at least one value is indicated
If ScanComboNumberOf%(n%) > 0 Then
lpAppName$ = "Monitor"
lpKeyName$ = "ScanComboNames" & Format$(n%)
If n% = 1 Then lpDefault$ = "SEI,COMPO,TOPO,AUX,XR1,XR2,XR3,XR4,XR5,EDS"
If n% = 2 Then lpDefault$ = "PIC,CROSS,LSP,SPOT"
If n% = 3 Then lpDefault$ = "TV,SR,S1,S2,S3"
If n% = 4 Then lpDefault$ = "NOR,EMP,LDF,ECP"
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then astring$ = Left$(lpReturnString$, valid&)
If Trim$(astring$) = vbNullString Then
msg$ = "ScanComboNames keyword string is empty in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINI2"
End
End If
Call InitParseStringToString2(astring$, n%, ScanComboNumberOf%(n%), ScanComboNames$())
If ierror Then End

lpAppName$ = "Monitor"
lpKeyName$ = "ScanComboParameters" & Format$(n%)
If n% = 1 Then lpDefault$ = "SEI,COM,TOP,AUX,XR1,XR2,XR3,XR4,XR5,EDS"
If n% = 2 Then lpDefault$ = "PIC,BUP,LSP,SPT"
If n% = 3 Then lpDefault$ = "TV,SR,S1,S2,S3"
If n% = 4 Then lpDefault$ = "NOR,EMP,LDF,ECP"
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then astring$ = Left$(lpReturnString$, valid&)
If Trim$(astring$) = vbNullString Then
msg$ = "ScanComboParameters keyword string is empty in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINI2"
End
End If
Call InitParseStringToString2(astring$, n%, ScanComboNumberOf%(n%), ScanComboParameters$())
If ierror Then End
End If

Next n%

Exit Sub

' Errors
InitINI2Error:
MsgBox Error$, vbOKOnly + vbCritical, "InitINI2"
ierror = True
Exit Sub

End Sub

Sub InitINI3()
' Open the PROBEWIN.INI file and read defaults for the sample exchange positions

ierror = False
On Error GoTo InitINI3Error

Dim valid As Long, nSize As Long, tValid As Long
Dim tcomment As String

Dim lpAppName As String
Dim lpKeyName As String
Dim lpDefault As String
Dim lpFileName As String
Dim lpReturnString As String * 255
Dim lpReturnString2 As String * 255

' Check for pathological conditions
If NumberOfStageMotors% < 1 Then Exit Sub

' Check for existing PROBEWIN.INI
If Dir$(ProbeWinINIFile$) = vbNullString Then
msg$ = "Unable to open file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINI3"
End
End If

' Use Windows API function to read PROBEWIN.INI
lpFileName$ = ProbeWinINIFile$
nSize& = Len(lpReturnString$)

lpAppName$ = "Hardware"
lpKeyName$ = "SampleExchangePositionX"
lpDefault$ = Str$(MotParkPositions!(XMotor%))
If InterfaceType% = 2 Then lpDefault$ = "44.5"   ' 8200 or 8900 or 8500 or 8x30
If InterfaceType% = 5 Then lpDefault$ = "0"      ' SX100/SXFive
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then SampleExchangePositions!(XMotor% - NumberOfTunableSpecs%) = Val(Left$(lpReturnString$, valid&))
If SampleExchangePositions!(XMotor% - NumberOfTunableSpecs%) < MotLoLimits!(XMotor%) Or SampleExchangePositions!(XMotor% - NumberOfTunableSpecs%) > MotHiLimits!(XMotor%) Then
msg$ = "SampleExchangePositionX keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINI3"
SampleExchangePositions!(XMotor% - NumberOfTunableSpecs%) = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "SampleExchangePositionY"
lpDefault$ = Str$(MotParkPositions!(YMotor%))
If InterfaceType% = 2 Then lpDefault$ = "1"                             ' 8200 or 8900 or 8500 or 8x30
If InterfaceType% = 5 Then lpDefault$ = Str$(MotHiLimits!(YMotor%))     ' SX100/SXFive
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then SampleExchangePositions!(YMotor% - NumberOfTunableSpecs%) = Val(Left$(lpReturnString$, valid&))
If SampleExchangePositions!(YMotor% - NumberOfTunableSpecs%) < MotLoLimits!(YMotor%) Or SampleExchangePositions!(YMotor% - NumberOfTunableSpecs%) > MotHiLimits!(YMotor%) Then
msg$ = "SampleExchangePositionY keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINI3"
SampleExchangePositions!(YMotor% - NumberOfTunableSpecs%) = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Z axis
If NumberOfStageMotors% > 2 Then
lpAppName$ = "Hardware"
lpKeyName$ = "SampleExchangePositionZ"
lpDefault$ = Str$(MotParkPositions!(ZMotor%))
If InterfaceType% = 2 Then lpDefault$ = "11"      ' 8200 or 8900 or 8500 or 8x30
If InterfaceType% = 5 Then lpDefault$ = "0"       ' SX100/SXFive
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then SampleExchangePositions!(ZMotor% - NumberOfTunableSpecs%) = Val(Left$(lpReturnString$, valid&))
If SampleExchangePositions!(ZMotor% - NumberOfTunableSpecs%) < MotLoLimits!(ZMotor%) Or SampleExchangePositions!(ZMotor% - NumberOfTunableSpecs%) > MotHiLimits!(ZMotor%) Then
msg$ = "SampleExchangePositionZ keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINI3"
SampleExchangePositions!(ZMotor% - NumberOfTunableSpecs%) = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)
End If

' W axis
'If NumberOfStageMotors% > 3 Then
'lpAppName$ = "Hardware"
'lpKeyName$ = "SampleExchangePositionW"
'lpDefault$ = Str$(MotParkPositions!(WMotor%))
'tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
'valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
'Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
'If Left$(lpReturnString$, valid&) <> vbNullString Then SampleExchangePositions!(WMotor% - NumberOfTunableSpecs%) = Val(Left$(lpReturnString$, valid&))
'If SampleExchangePositions!(WMotor% - NumberOfTunableSpecs%) < MotLoLimits!(WMotor%) Or SampleExchangePositions!(WMotor% - NumberOfTunableSpecs%) > MotHiLimits!(WMotor%) Then
'msg$ = "SampleExchangePositionW keyword value out of range in " & ProbeWinINIFile$
'MsgBox msg$, vbOKOnly + vbExclamation, "InitINI3"
'SampleExchangePositions!(WMotor% - NumberOfTunableSpecs%) = Val(lpDefault$)
'End If
'If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)
'End If

Exit Sub

' Errors
InitINI3Error:
MsgBox Error$, vbOKOnly + vbCritical, "InitINI3"
ierror = True
Exit Sub

End Sub

Sub InitINI4(tBoolean As Boolean, tKeyword As String, tSection As String)
' Open the PROBEWIN.INI file and read defaults for the specified boolean keyword in the specified section

ierror = False
On Error GoTo InitINI4Error

Dim valid As Long, nSize As Long, tValid As Long

Dim lpAppName As String
Dim lpKeyName As String
Dim nDefault As Long, tDefault As Long
Dim lpFileName As String
Dim lpReturnString As String * 255

' Check for existing PROBEWIN.INI
If Dir$(ProbeWinINIFile$) = vbNullString Then
msg$ = "Unable to open file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINI4"
End
End If

' Use Windows API function to read PROBEWIN.INI
lpFileName$ = ProbeWinINIFile$
nSize& = Len(lpReturnString$)

' Get defaults for [software] section
If tSection$ = "Software" Then
If tKeyword$ = "UserSpecifiedOutputSampleName" Then tDefault& = 1
If tKeyword$ = "UserSpecifiedOutputLineNumber" Then tDefault& = 1
If tKeyword$ = "UserSpecifiedOutputWeightPercent" Then tDefault& = 1
If tKeyword$ = "UserSpecifiedOutputOxidePercent" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputAtomicPercent" Then tDefault& = 1
If tKeyword$ = "UserSpecifiedOutputTotal" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputDetectionLimits" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputPercentError" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputStageX" Then tDefault& = 1
If tKeyword$ = "UserSpecifiedOutputStageY" Then tDefault& = 1
If tKeyword$ = "UserSpecifiedOutputStageZ" Then tDefault& = 1
If tKeyword$ = "UserSpecifiedOutputRelativeDistance" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputBeamCurrent" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputOnPeakTime" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputHiPeakTime" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputLoPeakTime" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputOnPeakCounts" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputOffPeakCounts" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputNetPeakCounts" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputKraw" Then tDefault& = 1
If tKeyword$ = "UserSpecifiedOutputDateTime" Then tDefault& = 1

If tKeyword$ = "UserSpecifiedOutputKratio" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputZAF" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputMAC" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputStdAssigns" Then tDefault& = 0

If tKeyword$ = "UserSpecifiedOutputSampleNumber" Then tDefault& = 1
If tKeyword$ = "UserSpecifiedOutputSampleConditions" Then tDefault& = 0

If tKeyword$ = "UserSpecifiedOutputFormulaFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputTotalPercentFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputTotalOxygenFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputTotalCationsFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputCalculatedOxygenFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputExcessOxygenFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputZbarFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputAtomicWeightFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputOxygenFromHalogensFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputHalogenCorrectedOxygenFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputChargeBalanceFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputFeChargeFlag" Then tDefault& = 0

If tKeyword$ = "UserSpecifiedOutputSpaceBeforeFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputAverageFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputStandardDeviationFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputStandardErrorFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputMinimumFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputMaximumFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputSpaceAfterFlag" Then tDefault& = 0

If tKeyword$ = "UserSpecifiedOutputUnkIntfCorsFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputUnkMANAbsCorsFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputUnkAPFCorsFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputUnkVolElCorsFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputUnkVolElDevsFlag" Then tDefault& = 0

If tKeyword$ = "UserSpecifiedOutputSampleDescription" Then tDefault& = 0

If tKeyword$ = "UserSpecifiedOutputEndMembers" Then tDefault& = 0

If tKeyword$ = "UserSpecifiedOutputOxideMolePercentFlag" Then tDefault& = 0

If tKeyword$ = "UserSpecifiedOutputStandardPublishedValuesFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputStandardPercentVariancesFlag" Then tDefault& = 0
If tKeyword$ = "UserSpecifiedOutputStandardAgebraicDifferencesFlag" Then tDefault& = 0
End If

' Load passed boolean keyword
lpAppName$ = tSection$
lpKeyName$ = tKeyword$
nDefault& = tDefault&
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
tBoolean = True
Else
tBoolean = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

Exit Sub

' Errors
InitINI4Error:
MsgBox Error$, vbOKOnly + vbCritical, "InitINI4"
ierror = True
Exit Sub

End Sub

Sub InitINIGeneral()
' Open the PROBEWIN.INI file and read defaults

ierror = False
On Error GoTo InitINIGeneralError

Dim valid As Long, tValid As Long

Dim lpAppName As String
Dim lpKeyName As String
Dim lpDefault As String
Dim lpFileName As String
Dim lpReturnString As String * 255
Dim lpReturnString2 As String * 255

Dim nSize As Long
Dim nDefault As Long
Dim tcomment As String

' Check for existing PROBEWIN.INI
If Dir$(ProbeWinINIFile$) = vbNullString Then
msg$ = "Unable to open file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIGeneral"
End
End If

' Use Windows API function to read PROBEWIN.INI
lpFileName$ = ProbeWinINIFile$
nSize& = Len(lpReturnString$)

' Get default KiloVolts
lpAppName$ = "General"
lpKeyName$ = "KiloVolts"
lpDefault$ = "15"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)      ' get default value
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultKiloVolts! = Val(Left$(lpReturnString$, valid&))
If DefaultKiloVolts! < MINKILOVOLTS! Or DefaultKiloVolts! > MAXKILOVOLTS! Then
msg$ = "Kilovolts keyword value out of range in " & ProbeWinINIFile$ & " (must be between " & Format$(MINKILOVOLTS!) & " and " & Format$(MAXKILOVOLTS!) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIGeneral"
DefaultKiloVolts! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Get default Takeoff angle
lpAppName$ = "General"
lpKeyName$ = "Takeoff"
lpDefault$ = "40"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultTakeOff! = Val(Left$(lpReturnString$, valid&))
If DefaultTakeOff! < MINTAKEOFF! Or DefaultTakeOff! > MAXTAKEOFF! Then
msg$ = "TakeOff keyword value out of range in " & ProbeWinINIFile$ & " (must be between " & Format$(MINTAKEOFF!) & " and " & Format$(MAXTAKEOFF!) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIGeneral"
DefaultTakeOff! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Get default Beam Current (nanoamps)
lpAppName$ = "General"
lpKeyName$ = "BeamCurrent"
lpDefault$ = "30"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultBeamCurrent! = Val(Left$(lpReturnString$, valid&))
If DefaultBeamCurrent! < MINBEAMCURRENT! * 10# Or DefaultBeamCurrent! > MAXBEAMSIZE! Then
msg$ = "BeamCurrent keyword value out of range in " & ProbeWinINIFile$ & " (must be between " & Format$(MINBEAMCURRENT! * 10) & " and " & Format$(MAXBEAMCURRENT!) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIGeneral"
DefaultBeamCurrent! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Get default Beam Size (microns)
lpAppName$ = "General"
lpKeyName$ = "BeamSize"
lpDefault$ = "0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultBeamSize! = Val(Left$(lpReturnString$, valid&))
If DefaultBeamSize! < 0 Or DefaultBeamSize! > MAXBEAMSIZE! Then
msg$ = "BeamSize keyword value out of range in " & ProbeWinINIFile$ & " (must be between " & Format$(MINBEAMSIZE!) & " and " & Format$(MAXBEAMSIZE!) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIGeneral"
DefaultBeamSize! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "General"
lpKeyName$ = "BeamMode"
nDefault& = 0       ' analog spot
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DefaultBeamMode% = valid&
If DefaultBeamMode% < 0 Or DefaultBeamMode% > 1 Then
msg$ = "BeamMode keyword value out of range in " & ProbeWinINIFile$ & ". Must be 0 for analog spot or 1 for analog scan."
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIGeneral"
DefaultBeamMode% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Get default magnification
lpAppName$ = "General"
lpKeyName$ = "Magnification"
lpDefault$ = "1000"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultMagnification! = Val(Left$(lpReturnString$, valid&))
If DefaultMagnification! < 10# Or DefaultMagnification! > 100000# Then
msg$ = "Magnification keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIGeneral"
DefaultMagnification! = Val(lpDefault$)
End If
DefaultMagnificationDefault! = DefaultMagnification!    ' actual default magnification
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "General"
lpKeyName$ = "Aperture"
nDefault& = 1       ' aperture 1 is default
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DefaultAperture% = valid&
If DefaultAperture% < 1 Or DefaultAperture% > 4 Then
msg$ = "Aperture keyword value out of range in " & ProbeWinINIFile$ & " (Must be between 1 and 4)."
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIGeneral"
DefaultAperture% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Get default Oxide or Elemental mode
lpAppName$ = "General"
lpKeyName$ = "OxideOrElemental"
nDefault& = 2
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DefaultOxideOrElemental% = valid&
If DefaultOxideOrElemental% < 1 Or DefaultOxideOrElemental% > 2 Then
msg$ = "OxideOrElemental keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIGeneral"
DefaultOxideOrElemental% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Get default peak center method
lpAppName$ = "General"
lpKeyName$ = "PeakCenterMethod"
nDefault& = 1   ' default is parabolic
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DefaultPeakCenterMethod% = valid&
If DefaultPeakCenterMethod% < 0 Or DefaultPeakCenterMethod% > 2 Then
msg$ = "PeakCenterMethod keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIGeneral"
DefaultPeakCenterMethod% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "General"
lpKeyName$ = "DebugMode"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
DefaultDebugMode% = True
Else
DefaultDebugMode% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "General"
lpKeyName$ = "UserName"
lpDefault$ = "User Name"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If (Trim$(MDBUserName$) = vbNullString And MiscStringsAreSame(app.EXEName, "Probewin")) Then
If Left$(lpReturnString$, valid&) <> vbNullString Then MDBUserName$ = Left$(lpReturnString$, valid&)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "General"
lpKeyName$ = "Title"
lpDefault$ = "File Title"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then MDBFileTitle$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "General"
lpKeyName$ = "Description"
lpDefault$ = "File Description"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then MDBFileDescription$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Get default text file viewer
lpAppName$ = "General"
lpKeyName$ = "FileViewer"
lpDefault$ = "NOTEPAD.EXE"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then FileViewer$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Get custom fields for FileInfo and UserInfo forms
lpAppName$ = "General"
lpKeyName$ = "CustomLabel1"
lpDefault$ = "Department"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then CustomLabel1$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "General"
lpKeyName$ = "CustomLabel2"
lpDefault$ = "Account #"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then CustomLabel2$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "General"
lpKeyName$ = "CustomLabel3"
lpDefault$ = "Group"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then CustomLabel3$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "General"
lpKeyName$ = "CustomText1"
lpDefault$ = vbNullString
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then CustomText1$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "General"
lpKeyName$ = "CustomText2"
lpDefault$ = vbNullString
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then CustomText2$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "General"
lpKeyName$ = "CustomText3"
lpDefault$ = vbNullString
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then CustomText3$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "General"
lpKeyName$ = "SMTPServerAddress"
lpDefault$ = vbNullString
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then SMTPServerAddress$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "General"
lpKeyName$ = "SMTPAddressFrom"
lpDefault$ = vbNullString
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then SMTPAddressFrom$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "General"
lpKeyName$ = "SMTPAddressTo"
lpDefault$ = vbNullString
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then SMTPAddressTo$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "General"
lpKeyName$ = "SMTPUserName"
lpDefault$ = vbNullString
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then SMTPUserName$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "General"
lpKeyName$ = "UseWavFileAfterAutomationString"
lpDefault$ = vbNullString
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then UseWavFileAfterAutomationString$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Check that sound file exists
If Dir$(ApplicationCommonAppData$ & UseWavFileAfterAutomationString$) = vbNullString Then
msg$ = "Automation sound file " & UseWavFileAfterAutomationString$ & " as defined in " & ProbeWinINIFile$ & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIGeneral"
End If

lpAppName$ = "General"
lpKeyName$ = "PeakCenterSkipPBCheck"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
DefaultPeakCenterSkipPBCheck% = True
Else
DefaultPeakCenterSkipPBCheck% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Get default magnification (analytical)
lpAppName$ = "General"
lpKeyName$ = "MagnificationAnalytical"
lpDefault$ = "20000"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultMagnificationAnalytical! = Val(Left$(lpReturnString$, valid&))
If DefaultMagnificationAnalytical! < 100# Or DefaultMagnificationAnalytical! > 800000# Then
msg$ = "MagnificationAnalytical keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIGeneral"
DefaultMagnificationAnalytical! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Get default magnification (imaging)
lpAppName$ = "General"
lpKeyName$ = "MagnificationImaging"
lpDefault$ = "100"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultMagnificationImaging! = Val(Left$(lpReturnString$, valid&))
If DefaultMagnificationImaging! < 10# Or DefaultMagnificationImaging! > 100000# Then
msg$ = "MagnificationImaging keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIGeneral"
DefaultMagnificationImaging! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Get default nominal beam current (nanoamps)
lpAppName$ = "General"
lpKeyName$ = "NominalBeam"
lpDefault$ = DefaultBeamCurrent!
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then NominalBeam! = Val(Left$(lpReturnString$, valid&))
If NominalBeam! < MINBEAMCURRENT! * 10# Or NominalBeam! > MAXBEAMSIZE! Then
msg$ = "NominalBeam keyword value out of range in " & ProbeWinINIFile$ & " (must be between " & Format$(MINBEAMCURRENT! * 10) & " and " & Format$(MAXBEAMCURRENT!) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIGeneral"
NominalBeam! = Val(lpDefault$)
End If
OriginalNominalBeam! = NominalBeam!
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "General"
lpKeyName$ = "InstrumentAcknowledgementString"
lpDefault$ = vbNullString
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then InstrumentAcknowledgementString$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

Exit Sub

' Errors
InitINIGeneralError:
MsgBox Error$, vbOKOnly + vbCritical, "InitINIGeneral"
ierror = True
Exit Sub

End Sub

Sub InitINISoftware()
' Open the PROBEWIN.INI file and read defaults

ierror = False
On Error GoTo InitINISoftwareError

Dim i As Integer
Dim valid As Long, tValid As Long

Dim lpAppName As String
Dim lpKeyName As String
Dim lpDefault As String
Dim lpFileName As String
Dim lpReturnString As String * 255
Dim lpReturnString2 As String * 255

Dim nSize As Long
Dim nDefault As Long
Dim tcomment As String

' Check for existing PROBEWIN.INI
If Dir$(ProbeWinINIFile$) = vbNullString Then
msg$ = "Unable to open file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
End
End If

' Use Windows API function to read PROBEWIN.INI
lpFileName$ = ProbeWinINIFile$
nSize& = Len(lpReturnString$)

lpAppName$ = "Software"
lpKeyName$ = "UserDataDirectory"
lpDefault$ = "C:\UserData"          ' changed to root default 03-20-2007
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
UserDataDirectory$ = lpDefault$      ' set to default in case keyword in INI file is a null string
If Left$(lpReturnString$, valid&) <> vbNullString Then UserDataDirectory$ = Left$(lpReturnString$, valid&)
If Right$(UserDataDirectory$, 1) = "\" Then UserDataDirectory$ = Left$(UserDataDirectory$, Len(UserDataDirectory$) - 1) ' remove trailing backslash
OriginalUserDataDirectory$ = UserDataDirectory$     ' save for special uses
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "CalcZAFDATFileDirectory"
lpDefault$ = UserDataDirectory$ & "\CalcZAFDATData"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
CalcZAFDATFileDirectory$ = lpDefault$      ' set to default in case keyword in INI file is a null string
If Left$(lpReturnString$, valid&) <> vbNullString Then CalcZAFDATFileDirectory$ = Left$(lpReturnString$, valid&)
If Right$(CalcZAFDATFileDirectory$, 1) = "\" Then CalcZAFDATFileDirectory$ = Left$(CalcZAFDATFileDirectory$, Len(CalcZAFDATFileDirectory$) - 1) ' remove trailing backslash
OriginalCalcZAFDATDirectory$ = CalcZAFDATFileDirectory$
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "ColumnPCCFileDirectory"
lpDefault$ = UserDataDirectory$ & "\ColumnPCCData"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
ColumnPCCFileDirectory$ = lpDefault$      ' set to default in case keyword in INI file is a null string
If Left$(lpReturnString$, valid&) <> vbNullString Then ColumnPCCFileDirectory$ = Left$(lpReturnString$, valid&)
If Right$(ColumnPCCFileDirectory$, 1) = "\" Then ColumnPCCFileDirectory$ = Left$(ColumnPCCFileDirectory$, Len(ColumnPCCFileDirectory$) - 1) ' remove trailing backslash
OriginalColumnPCCFileDirectory$ = ColumnPCCFileDirectory$
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "SurferDataDirectory"
lpDefault$ = UserDataDirectory$ & "\SurferData"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
SurferDataDirectory$ = lpDefault$      ' set to default in case keyword in INI file is a null string
If Left$(lpReturnString$, valid&) <> vbNullString Then SurferDataDirectory$ = Left$(lpReturnString$, valid&)
If Right$(SurferDataDirectory$, 1) = "\" Then SurferDataDirectory$ = Left$(SurferDataDirectory$, Len(SurferDataDirectory$) - 1) ' remove trailing backslash
OriginalSurferDataDirectory$ = SurferDataDirectory$
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "GrapherDataDirectory"
lpDefault$ = UserDataDirectory$ & "\GrapherData"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
GrapherDataDirectory$ = lpDefault$      ' set to default in case keyword in INI file is a null string
If Left$(lpReturnString$, valid&) <> vbNullString Then GrapherDataDirectory$ = Left$(lpReturnString$, valid&)
If Right$(GrapherDataDirectory$, 1) = "\" Then GrapherDataDirectory$ = Left$(GrapherDataDirectory$, Len(GrapherDataDirectory$) - 1) ' remove trailing backslash
OriginalGrapherDataDirectory$ = GrapherDataDirectory$
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "DemoImagesDirectory"
lpDefault$ = UserDataDirectory$ & "\DemoImages"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
DemoImagesDirectory$ = lpDefault$      ' set to default in case keyword in INI file is a null string
If Left$(lpReturnString$, valid&) <> vbNullString Then DemoImagesDirectory$ = Left$(lpReturnString$, valid&)
If Right$(DemoImagesDirectory$, 1) = "\" Then DemoImagesDirectory$ = Left$(DemoImagesDirectory$, Len(DemoImagesDirectory$) - 1) ' remove trailing backslash
OriginalDemoImagesDirectory$ = DemoImagesDirectory$
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Software section, now get default Log Window font
lpAppName$ = "Software"
lpKeyName$ = "LogWindowFontName"
lpDefault$ = "Courier New"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then LogWindowFontName$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Get default Log Window Font Size
lpAppName$ = "Software"
lpKeyName$ = "LogWindowFontSize"
nDefault& = 10
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
LogWindowFontSize% = valid&
If LogWindowFontSize% < 6 Or LogWindowFontSize% > 32 Then
msg$ = "LogWindowFontSize keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
LogWindowFontSize% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Get default Log Window Font Bold
lpAppName$ = "Software"
lpKeyName$ = "LogWindowFontBold"
nDefault& = True
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
LogWindowFontBold = True
Else
LogWindowFontBold = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Get Motor Font Size
lpAppName$ = "Software"
lpKeyName$ = "AcquirePositionFontSize"
nDefault& = 12
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
AcquirePositionFontSize% = valid&
If AcquirePositionFontSize% < 6 Or AcquirePositionFontSize% > 32 Then
msg$ = "AcquirePositionFontSize keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
AcquirePositionFontSize% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Get Count Font Size
lpAppName$ = "Software"
lpKeyName$ = "AcquireCountFontSize"
nDefault& = 12
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
AcquireCountFontSize% = valid&
If AcquireCountFontSize% < 6 Or AcquireCountFontSize% > 32 Then
msg$ = "AcquireCountFontSize keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
AcquireCountFontSize% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Change the Log Window font and size if the user specified it in the PROBEWIN.INI file and it is present
For i% = 0 To Screen.FontCount - 1
If Screen.Fonts(i%) = LogWindowFontName$ Then
FormMAIN.TextLog.SelFontName = LogWindowFontName$
Exit For
End If
Next i%
FormMAIN.TextLog.SelFontSize = LogWindowFontSize%
FormMAIN.TextLog.SelBold = LogWindowFontBold

' Get timer intervals
lpAppName$ = "Software"
lpKeyName$ = "LogWindowInterval"
lpDefault$ = "0.5"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then LogWindowInterval! = Val(Left$(lpReturnString$, valid&))
If LogWindowInterval! < 0.1 Or LogWindowInterval! > 10# Then
msg$ = "LogWindowInterval keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
End
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "RealTimeInterval"
lpDefault$ = "0.2"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then RealTimeInterval! = Val(Left$(lpReturnString$, valid&))
If RealTimeInterval! < 0.1 Or RealTimeInterval! > 10# Then
msg$ = "RealTimeInterval keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
RealTimeInterval! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "AutomateConfirmDelay"
lpDefault$ = "10.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then AutomateConfirmDelay! = Val(Left$(lpReturnString$, valid&))
If AutomateConfirmDelay! < 0# Or AutomateConfirmDelay! > 100# Then
msg$ = "AutomateConfirmDelay keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
AutomateConfirmDelay! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Get enter off peak positions flag
lpAppName$ = "Software"
lpKeyName$ = "EnterPositionsRelative"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
EnterPositionsRelativeFlag = True
Else
EnterPositionsRelativeFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Get update peak/wavescan positions flag
lpAppName$ = "Software"
lpKeyName$ = "UpdatePeakWaveScanPositions"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
UpdatePeakWaveScanPositionsFlag = True
Else
UpdatePeakWaveScanPositionsFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Get maximum file array menu items
lpAppName$ = "Software"
lpKeyName$ = "MaxMenuFileArray"
nDefault& = 12
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
MaxMenuFileArray% = CInt(valid&)
If MaxMenuFileArray% < 0 Or MaxMenuFileArray% > 12 Then
msg$ = "MaxMenuFileArray keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
MaxMenuFileArray% = CInt(nDefault&)
End If
MaxMenuFileArray% = MaxMenuFileArray% + 1   ' add one for separator bar
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Get extended format flag
lpAppName$ = "Software"
lpKeyName$ = "ExtendedFormat"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ExtendedFormat = True
Else
ExtendedFormat = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "MACTypeFlag"
nDefault& = 1   ' LINEMU.DAT
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
MACTypeFlag% = CInt(valid&)
If MACTypeFlag% < 1 Or MACTypeFlag% > MAXMACTYPE% Then
msg$ = "MACTypeFlag keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
MACTypeFlag% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "PositionImportExportFileType"
nDefault& = 1
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
PositionImportExportFileType% = valid&
If PositionImportExportFileType% < 1 Or PositionImportExportFileType% > 2 Then
msg$ = "PositionImportExportFileType keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
PositionImportExportFileType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "DeadtimeCorrectionType"
nDefault& = 1
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DeadTimeCorrectionType% = valid&
If DeadTimeCorrectionType% < 1 Or DeadTimeCorrectionType% > 2 Then
msg$ = "DeadtimeCorrectionType keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
DeadTimeCorrectionType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "AutoFocusStyle"
nDefault& = 1
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
AutoFocusStyle% = CInt(valid&)
If AutoFocusStyle% < 1 Or AutoFocusStyle% > 4 Then
msg$ = "AutoFocusStyle keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
AutoFocusStyle% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "AutoFocusInterval"
nDefault& = 5
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
AutoFocusInterval% = CInt(valid&)
If AutoFocusInterval% < 1 Or AutoFocusInterval% > 1000 Then
msg$ = "AutoFocusInterval keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
AutoFocusInterval% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "BiasChangeDelay"
lpDefault$ = "2.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then BiasChangeDelay! = Val(Left$(lpReturnString$, valid&))
If BiasChangeDelay! < 0# Or BiasChangeDelay! > 100# Then
msg$ = "BiasChangeDelay keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
BiasChangeDelay! = Val(lpDefault$)
End If
DefaultBiasChangeDelay! = BiasChangeDelay!  ' for restoring
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "UseEmpiricalPHADefaults"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
UseEmpiricalPHADefaults = True
Else
UseEmpiricalPHADefaults = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "KilovoltChangeDelay"
lpDefault$ = "1.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then KilovoltChangeDelay! = Val(Left$(lpReturnString$, valid&))
If KilovoltChangeDelay! < 0# Or KilovoltChangeDelay! > 100# Then
msg$ = "KilovoltChangeDelay keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
KilovoltChangeDelay! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "BeamCurrentChangeDelay"
lpDefault$ = "1.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then BeamCurrentChangeDelay! = Val(Left$(lpReturnString$, valid&))
If BeamCurrentChangeDelay! < 0# Or BeamCurrentChangeDelay! > 100# Then
msg$ = "BeamCurrentChangeDelay keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
BeamCurrentChangeDelay! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "BeamSizeChangeDelay"
lpDefault$ = "0.5"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then BeamSizeChangeDelay! = Val(Left$(lpReturnString$, valid&))
If BeamSizeChangeDelay! < 0# Or BeamSizeChangeDelay! > 100# Then
msg$ = "BeamSizeChangeDelay keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
BeamSizeChangeDelay! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Log window size (new for 32 bit NT)
lpAppName$ = "Software"
lpKeyName$ = "LogWindowBufferSize"
If MiscSystemGetOSVersionNumber&() < OS_VERSION_NT4& Then    ' 0 = Win32s, 1 = Win95, 2 for NT, 3 = NT351, 4 = NT4, 5 = XP, 6 = Vista, 7 = Win7
nDefault& = MAXINTEGER%   ' works for Win95
Else
nDefault& = MAXINTEGER% * 16&  ' for NT/2000/XP
End If
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
LogWindowBufferSize& = valid&
If LogWindowBufferSize& < 1000 Or LogWindowBufferSize& > 10000000 Then
msg$ = "LogWindowBufferSize keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
LogWindowBufferSize& = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "CommandPacingInterval"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
CommandPacingInterval% = CInt(valid&)
If CommandPacingInterval% < 0 Or CommandPacingInterval% > 100 Then
msg$ = "CommandPacingInterval keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
CommandPacingInterval% = CInt(nDefault&)
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "PeakOnAssignedStandards"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
PeakOnAssignedStandardsFlag% = True
Else
PeakOnAssignedStandardsFlag% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "PrintAnalyzedAndSpecifiedOnSameLine"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
PrintAnalyzedAndSpecifiedOnSameLineFlag% = True
Else
PrintAnalyzedAndSpecifiedOnSameLineFlag% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "NoMotorPositionLimitsChecking"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
NoMotorPositionLimitsCheckingFlag% = True
Else
NoMotorPositionLimitsCheckingFlag% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "ExtendedMenu"     ' used in CalcZAF.exe and Help Update only
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ExtendedMenuFlag% = True
Else
ExtendedMenuFlag% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "AutoAnalyze"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
AutoAnalyzeFlag% = True
Else
AutoAnalyzeFlag% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "FaradayAlwaysOnTop"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
FaradayAlwaysOnTopFlag = True
Else
FaradayAlwaysOnTopFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "ColumnConditionChangeDelay"
lpDefault$ = "2.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then ColumnConditionChangeDelay! = Val(Left$(lpReturnString$, valid&))
If ColumnConditionChangeDelay! < 0# Or ColumnConditionChangeDelay! > 100# Then
msg$ = "ColumnConditionChangeDelay keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
ColumnConditionChangeDelay! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "SurferOutputVersionNumber"
nDefault& = 7       ' new default to use VBA code for versions 7 and up
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
SurferOutputVersionNumber% = CInt(valid&)
If SurferOutputVersionNumber% < 6 Then
msg$ = "SurferOutputVersionNumber keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
SurferOutputVersionNumber% = CInt(nDefault&)
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "SelPrintStartDoc"     ' only for VB6
nDefault& = True
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
SelPrintStartDocFlag = True
Else
SelPrintStartDocFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "UseSimpleRegistration"     ' (for bypassing "hash" registration, but use by default for Unicode languages to avoid byte issues)
nDefault& = True
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
UseSimpleRegistrationFlag = True
Else
UseSimpleRegistrationFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "UseMultiplePeakCalibrationOffset"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
UseMultiplePeakCalibrationOffsetFlag = True
Else
UseMultiplePeakCalibrationOffsetFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Flag to force "wide" ROM scan
lpAppName$ = "Software"
lpKeyName$ = "UseWideROMPeakScanAlways"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
UseWideROMPeakScanAlwaysFlag = True
Else
UseWideROMPeakScanAlwaysFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Flag to read current isnrtument conditions when starting application
lpAppName$ = "Software"
lpKeyName$ = "UseCurrentConditionsOnStartUp"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
UseCurrentConditionsOnStartUpFlag = True
Else
UseCurrentConditionsOnStartUpFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Flag to read current instrument conditions when starting application and starting acquisitions
lpAppName$ = "Software"
lpKeyName$ = "UseCurrentConditionsAlways"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
UseCurrentConditionsAlwaysFlag = True
Else
UseCurrentConditionsAlwaysFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Default vacuum units display
lpAppName$ = "Software"
lpKeyName$ = "DefaultVacuumUnitsType"
nDefault& = 0       ' assume Pascals (1 = Torr, 2 = mBar)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DefaultVacuumUnitsType% = CInt(valid&)
If DefaultVacuumUnitsType% < 0 Or DefaultVacuumUnitsType% > 2 Then
msg$ = "DefaultVacuumUnitsType keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
DefaultVacuumUnitsType% = CInt(nDefault&)
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Correction flag
lpAppName$ = "Software"
lpKeyName$ = "DefaultCorrectionType"
nDefault& = 0       ' assume ZAF/Phi-Rho-Z (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
CorrectionFlag% = CInt(valid&)
If CorrectionFlag% < 0 Or CorrectionFlag% > MAXCORRECTION% Then
msg$ = "DefaultCorrectionType keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
CorrectionFlag% = CInt(nDefault&)
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' ZAF flag
lpAppName$ = "Software"
lpKeyName$ = "DefaultZAFType"
nDefault& = 1       ' assume Armstrong Phi-Rho-Z
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
izaf% = CInt(valid&)
If izaf% < 1 Or izaf% > 10 Then
msg$ = "DefaultZAFType keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
izaf% = CInt(nDefault&)
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)
Call InitGetZAFSetZAF2(izaf%)   ' load default ZAF selections

lpAppName$ = "Software"
lpKeyName$ = "PENDBASE_Path"
'lpDefault$ = UserDataDirectory$ & "\Penepma06\Pendbase"     ' "C:\Userdata\Penepma06\Pendbase"
'lpDefault$ = UserDataDirectory$ & "\Penepma08\Pendbase"     ' "C:\Userdata\Penepma08\Pendbase"
lpDefault$ = UserDataDirectory$ & "\Penepma12\Pendbase"     ' "C:\Userdata\Penepma12\Pendbase"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then PENDBASE_Path$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "PENEPMA_Path"
'lpDefault$ = UserDataDirectory$ & "\Penepma06\Penepma"        ' "C:\Userdata\Penepma06\Penepma"
'lpDefault$ = UserDataDirectory$ & "\Penepma08\Penepma"        ' "C:\Userdata\Penepma08\Penepma"
lpDefault$ = UserDataDirectory$ & "\Penepma12\Penepma"        ' "C:\Userdata\Penepma12\Penepma"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then PENEPMA_Path$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Load PENEPMA root
lpAppName$ = "Software"
lpKeyName$ = "PENEPMA_Root"
'lpDefault$ = UserDataDirectory$ & "\Penepma06"                ' "C:\Userdata\Penepma06"
'lpDefault$ = UserDataDirectory$ & "\Penepma08"                ' "C:\Userdata\Penepma08"
lpDefault$ = UserDataDirectory$ & "\Penepma12"                ' "C:\Userdata\Penepma12"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then PENEPMA_Root$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Load PENEPMA PAR path (for network shared folder)
lpAppName$ = "Software"
lpKeyName$ = "PENEPMA_PAR_Path"
lpDefault$ = UserDataDirectory$ & "\Penepma12\Penfluor"                ' "C:\Userdata\Penepma12\Penfluor"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then PENEPMA_PAR_Path$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Load ForceNegativeKratiosToZeroFlag
lpAppName$ = "Software"
lpKeyName$ = "ForceNegativeKratiosToZero"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ForceNegativeKratiosToZeroFlag = True
Else
ForceNegativeKratiosToZeroFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "AutoIncrementDelimiterString"
lpDefault$ = "_"    ' underscore is default character
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then AutoIncrementDelimiterString$ = Left$(lpReturnString$, valid&)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Load UseLastUnknownAsWavescanSetupFlag
lpAppName$ = "Software"
lpKeyName$ = "UseLastUnknownAsWavescanSetup"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
UseLastUnknownAsWavescanSetupFlag = True
Else
UseLastUnknownAsWavescanSetupFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' User specified custom output
Call InitINI4(UserSpecifiedOutputSampleNameFlag, "UserSpecifiedOutputSampleName", "Software")
Call InitINI4(UserSpecifiedOutputLineNumberFlag, "UserSpecifiedOutputLineNumber", "Software")
Call InitINI4(UserSpecifiedOutputWeightPercentFlag, "UserSpecifiedOutputWeightPercent", "Software")
Call InitINI4(UserSpecifiedOutputOxidePercentFlag, "UserSpecifiedOutputOxidePercent", "Software")
Call InitINI4(UserSpecifiedOutputAtomicPercentFlag, "UserSpecifiedOutputAtomicPercent", "Software")
Call InitINI4(UserSpecifiedOutputTotalFlag, "UserSpecifiedOutputTotal", "Software")
Call InitINI4(UserSpecifiedOutputDetectionLimitsFlag, "UserSpecifiedOutputDetectionLimits", "Software")
Call InitINI4(UserSpecifiedOutputPercentErrorFlag, "UserSpecifiedOutputPercentError", "Software")
Call InitINI4(UserSpecifiedOutputStageXFlag, "UserSpecifiedOutputStageX", "Software")
Call InitINI4(UserSpecifiedOutputStageYFlag, "UserSpecifiedOutputStageY", "Software")
Call InitINI4(UserSpecifiedOutputStageZFlag, "UserSpecifiedOutputStageZ", "Software")
Call InitINI4(UserSpecifiedOutputRelativeDistanceFlag, "UserSpecifiedOutputRelativeDistance", "Software")
Call InitINI4(UserSpecifiedOutputBeamCurrentFlag, "UserSpecifiedOutputBeamCurrent", "Software")
Call InitINI4(UserSpecifiedOutputOnPeakTimeFlag, "UserSpecifiedOutputOnPeakTime", "Software")
Call InitINI4(UserSpecifiedOutputHiPeakTimeFlag, "UserSpecifiedOutputHiPeakTime", "Software")
Call InitINI4(UserSpecifiedOutputLoPeakTimeFlag, "UserSpecifiedOutputLoPeakTime", "Software")
Call InitINI4(UserSpecifiedOutputOnPeakCountsFlag, "UserSpecifiedOutputOnPeakCounts", "Software")
Call InitINI4(UserSpecifiedOutputOffPeakCountsFlag, "UserSpecifiedOutputOffPeakCounts", "Software")
Call InitINI4(UserSpecifiedOutputNetPeakCountsFlag, "UserSpecifiedOutputNetPeakCounts", "Software")
Call InitINI4(UserSpecifiedOutputKrawFlag, "UserSpecifiedOutputKraw", "Software")
Call InitINI4(UserSpecifiedOutputDateTimeFlag, "UserSpecifiedOutputDateTime", "Software")

Call InitINI4(UserSpecifiedOutputKratioFlag, "UserSpecifiedOutputKratio", "Software")
Call InitINI4(UserSpecifiedOutputZAFFlag, "UserSpecifiedOutputZAF", "Software")
Call InitINI4(UserSpecifiedOutputMACFlag, "UserSpecifiedOutputMAC", "Software")
Call InitINI4(UserSpecifiedOutputStdAssignsFlag, "UserSpecifiedOutputStdAssigns", "Software")

Call InitINI4(UserSpecifiedOutputSampleNumberFlag, "UserSpecifiedOutputSampleNumber", "Software")
Call InitINI4(UserSpecifiedOutputSampleConditionsFlag, "UserSpecifiedOutputSampleConditions", "Software")

Call InitINI4(UserSpecifiedOutputFormulaFlag, "UserSpecifiedOutputFormula", "Software")

Call InitINI4(UserSpecifiedOutputTotalPercentFlag, "UserSpecifiedOutputTotalPercent", "Software")
Call InitINI4(UserSpecifiedOutputTotalOxygenFlag, "UserSpecifiedOutputTotalOxygen", "Software")
Call InitINI4(UserSpecifiedOutputTotalCationsFlag, "UserSpecifiedOutputTotalCations", "Software")
Call InitINI4(UserSpecifiedOutputCalculatedOxygenFlag, "UserSpecifiedOutputCalculatedOxygen", "Software")
Call InitINI4(UserSpecifiedOutputExcessOxygenFlag, "UserSpecifiedOutputExcessOxygen", "Software")
Call InitINI4(UserSpecifiedOutputZbarFlag, "UserSpecifiedOutputZbar", "Software")
Call InitINI4(UserSpecifiedOutputAtomicWeightFlag, "UserSpecifiedOutputAtomicWeight", "Software")
Call InitINI4(UserSpecifiedOutputOxygenFromHalogensFlag, "UserSpecifiedOutputOxygenFromHalogens", "Software")
Call InitINI4(UserSpecifiedOutputHalogenCorrectedOxygenFlag, "UserSpecifiedOutputHalogenCorrectedOxygen", "Software")
Call InitINI4(UserSpecifiedOutputChargeBalanceFlag, "UserSpecifiedOutputChargeBalance", "Software")
Call InitINI4(UserSpecifiedOutputFeChargeFlag, "UserSpecifiedOutputFeCharge", "Software")

Call InitINI4(UserSpecifiedOutputSpaceBeforeFlag, "UserSpecifiedOutputSpaceBefore", "Software")
Call InitINI4(UserSpecifiedOutputAverageFlag, "UserSpecifiedOutputAverage", "Software")
Call InitINI4(UserSpecifiedOutputStandardDeviationFlag, "UserSpecifiedOutputStandardDeviation", "Software")
Call InitINI4(UserSpecifiedOutputStandardErrorFlag, "UserSpecifiedOutputStandardError", "Software")
Call InitINI4(UserSpecifiedOutputMinimumFlag, "UserSpecifiedOutputMinimum", "Software")
Call InitINI4(UserSpecifiedOutputMaximumFlag, "UserSpecifiedOutputMaximum", "Software")
Call InitINI4(UserSpecifiedOutputSpaceAfterFlag, "UserSpecifiedOutputSpaceAfter", "Software")

Call InitINI4(UserSpecifiedOutputUnkIntfCorsFlag, "UserSpecifiedOutputUnkIntfCors", "Software")
Call InitINI4(UserSpecifiedOutputUnkMANAbsCorsFlag, "UserSpecifiedOutputUnkMANAbsCors", "Software")
Call InitINI4(UserSpecifiedOutputUnkAPFCorsFlag, "UserSpecifiedOutputUnkAPFCors", "Software")
Call InitINI4(UserSpecifiedOutputUnkVolElCorsFlag, "UserSpecifiedOutputUnkVolElCors", "Software")
Call InitINI4(UserSpecifiedOutputUnkVolElDevsFlag, "UserSpecifiedOutputUnkVolElDevs", "Software")

Call InitINI4(UserSpecifiedOutputSampleDescriptionFlag, "UserSpecifiedOutputSampleDescription", "Software")

Call InitINI4(UserSpecifiedOutputEndMembersFlag, "UserSpecifiedOutputEndmembersFlag", "Software")

Call InitINI4(UserSpecifiedOutputOxideMolePercentFlag, "UserSpecifiedOutputOxideMolePercentFlag", "Software")

Call InitINI4(UserSpecifiedOutputStandardPublishedValuesFlag, "UserSpecifiedOutputStandardPublishedValuesFlag", "Software")
Call InitINI4(UserSpecifiedOutputStandardPercentVariancesFlag, "UserSpecifiedOutputStandardPercentVariancesFlag", "Software")
Call InitINI4(UserSpecifiedOutputStandardAlgebraicDifferencesFlag, "UserSpecifiedOutputStandardAlgebraicDifferencesFlag", "Software")

Call InitINI4(UserSpecifiedOutputTotalAtomsFlag, "UserSpecifiedOutputTotalAtomsFlag", "Software")

Call InitINI4(UserSpecifiedOutputRelativeLineNumberFlag, "UserSpecifiedOutputRelativeLineNumberFlag", "Software")

Call InitINI4(UserSpecifiedOutputAbsorbedCurrentFlag, "UserSpecifiedOutputAbsorbedCurrentFlag", "Software")

Call InitINI4(UserSpecifiedOutputBeamCurrent2Flag, "UserSpecifiedOutputBeamCurrent2Flag", "Software")
Call InitINI4(UserSpecifiedOutputAbsorbedCurrent2Flag, "UserSpecifiedOutputAbsorbedCurrent2Flag", "Software")

' Load DefaultNthPointAcquisitionInterval
lpAppName$ = "Software"
lpKeyName$ = "DefaultNthPointAcquisitionInterval"
nDefault& = 10       ' assume 10 points per interval
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DefaultNthPointAcquisitionInterval% = CInt(valid&)
If DefaultNthPointAcquisitionInterval% < 1 Or DefaultNthPointAcquisitionInterval% > 100 Then
msg$ = "DefaultNthPointAcquisitionInterval keyword value out of range (must be between 1 and 100) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
DefaultNthPointAcquisitionInterval% = CInt(nDefault&)
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Geological sort order flag
lpAppName$ = "Software"
lpKeyName$ = "GeologicalSortOrder"
nDefault& = 0       ' 0 = no sorting, 1 = traditional, 2 = low to high Z, 3 = high to low Z
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
GeologicalSortOrderFlag% = CInt(valid&)
If GeologicalSortOrderFlag% < 0 Or GeologicalSortOrderFlag% > MAXELEMENTSORTMETHODS% Then
msg$ = "GeolgicalSortOrder keyword value out of range (must be between 0 and " & Format$(MAXELEMENTSORTMETHODS%) & ") in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
GeologicalSortOrderFlag% = CInt(nDefault&)
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "DefaultLIFPeakWidth"
lpDefault$ = "0.08"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultLIFPeakWidth! = Val(Left$(lpReturnString$, valid&))
If DefaultLIFPeakWidth! < 0.01 Or DefaultLIFPeakWidth! > 0.1 Then
msg$ = "DefaultLIFPeakWidth keyword value (for nominal interference calculations) is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
DefaultLIFPeakWidth! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "ThermoNSSLocalRemoteMode"
nDefault& = 0       ' assume local mode (Thermo NSS and Probe for EPMA running on the same computer)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
ThermoNSSLocalRemoteMode% = CInt(valid&)
If ThermoNSSLocalRemoteMode% < 0 Or ThermoNSSLocalRemoteMode% > 1 Then
msg$ = "ThermoNSSLocalRemoteMode keyword value out of range (must be 0 or 1) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
ThermoNSSLocalRemoteMode% = CInt(nDefault&)
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Get default Monitor app Font Size
lpAppName$ = "Software"
lpKeyName$ = "MonitorFontSize"
nDefault& = 10
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
MonitorFontSize% = valid&
If MonitorFontSize% < 6 Or MonitorFontSize% > 32 Then
msg$ = "MonitorFontSize keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
MonitorFontSize% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Get JEOL Security Number (do not check for interfacetype, just read it and check in JeolInitInterface)
lpAppName$ = "Software"
lpKeyName$ = "JEOLSecurityNumber"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
JEOLSecurityNumber& = valid&
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "UseCurrentBeamBlankStateOnStartUpAndTermination"
nDefault& = 0       ' default is to blank beam on startup and terminate
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
UseCurrentBeamBlankStateOnStartUpAndTerminationFlag = True
Else
UseCurrentBeamBlankStateOnStartUpAndTerminationFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "ShowAllPeakingOptions"
nDefault& = 0       ' default is to not show all peaking options (if JEOL 8900/8200/8500 or SX100 interface)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ShowAllPeakingOptionsFlag = True
Else
ShowAllPeakingOptionsFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "ForceSetPHAParameters"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ForceSetPHAParametersFlag = True
Else
ForceSetPHAParametersFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "DoNotRescaleKLM"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
DefaultDoNotRescaleKLMFlag = True
Else
DefaultDoNotRescaleKLMFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "UsePenepmaKratiosLimit"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
UsePenepmaKratiosLimitFlag = True
Else
UsePenepmaKratiosLimitFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "PenepmaKratiosLimitValue"
lpDefault$ = "90"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then PenepmaKratiosLimitValue! = Val(Left$(lpReturnString$, valid&))
If PenepmaKratiosLimitValue! < 50 Or PenepmaKratiosLimitValue! > 99 Then
msg$ = "PenepmaKratiosLimitValue keyword value (for alpha factor calculations) is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
PenepmaKratiosLimitValue! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "PenepmaMinimumElectronEnergy"
lpDefault$ = "1.0"      ' default is 1 keV
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then PenepmaMinimumElectronEnergy! = Val(Left$(lpReturnString$, valid&))
If PenepmaMinimumElectronEnergy! < 0.001 Or PenepmaMinimumElectronEnergy! > 10# Then
msg$ = "PenepmaMinimumElectronEnergy keyword value (for monte-carlo simulations) is out of range in " & ProbeWinINIFile$ & ", (must be between 0.001 and 10 keV)"
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
PenepmaMinimumElectronEnergy! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "UserImagesDirectory"
lpDefault$ = Left(UserDataDirectory$, 1) & ":\UserImages"       ' use same drive as UserDataDirectory$
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
UserImagesDirectory$ = lpDefault$      ' set to default in case keyword in INI file is a null string
If Left$(lpReturnString$, valid&) <> vbNullString Then UserImagesDirectory$ = Left$(lpReturnString$, valid&)
If Right$(UserImagesDirectory$, 1) = "\" Then UserImagesDirectory$ = Left$(UserImagesDirectory$, Len(UserImagesDirectory$) - 1) ' remove trailing backslash
OriginalUserImagesDirectory$ = UserImagesDirectory$     ' save for special uses
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "UserEDSDirectory"
lpDefault$ = Left(UserDataDirectory$, 1) & ":\UserEDS"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
UserEDSDirectory$ = lpDefault$      ' set to default in case keyword in INI file is a null string
If Left$(lpReturnString$, valid&) <> vbNullString Then UserEDSDirectory$ = Left$(lpReturnString$, valid&)
If Right$(UserEDSDirectory$, 1) = "\" Then UserEDSDirectory$ = Left$(UserEDSDirectory$, Len(UserEDSDirectory$) - 1) ' remove trailing backslash
OriginalUserEDSDirectory$ = UserEDSDirectory$     ' save for special uses
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "UserCLDirectory"
lpDefault$ = Left(UserDataDirectory$, 1) & ":\UserCL"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
UserCLDirectory$ = lpDefault$      ' set to default in case keyword in INI file is a null string
If Left$(lpReturnString$, valid&) <> vbNullString Then UserCLDirectory$ = Left$(lpReturnString$, valid&)
If Right$(UserCLDirectory$, 1) = "\" Then UserCLDirectory$ = Left$(UserCLDirectory$, Len(UserCLDirectory$) - 1) ' remove trailing backslash
OriginalUserCLDirectory$ = UserCLDirectory$     ' save for special uses
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "UserEBSDDirectory"
lpDefault$ = Left(UserDataDirectory$, 1) & ":\UserEBSD"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
UserEBSDDirectory$ = lpDefault$      ' set to default in case keyword in INI file is a null string
If Left$(lpReturnString$, valid&) <> vbNullString Then UserEBSDDirectory$ = Left$(lpReturnString$, valid&)
If Right$(UserEBSDDirectory$, 1) = "\" Then UserEBSDDirectory$ = Left$(UserEBSDDirectory$, Len(UserEBSDDirectory$) - 1) ' remove trailing backslash
OriginalUserEBSDDirectory$ = UserEBSDDirectory$     ' save for special uses
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "SurferPlotsPerPage"
nDefault& = 4       ' must be 1, 4 or 9
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
SurferPlotsPerPage% = valid&
If SurferPlotsPerPage% <> 1 And SurferPlotsPerPage% <> 4 And SurferPlotsPerPage% <> 9 Then
msg$ = "SurferPlotsPerPage keyword value is invalid in " & ProbeWinINIFile$ & ", (must be 1, 4 or 9)"
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
SurferPlotsPerPage% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "GrapherAppDirectory"
lpDefault$ = vbNullString       ' no default!
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then
GrapherAppDirectory$ = Left$(lpReturnString$, valid&)
Else
Call InitDetermineScripterPath(Int(1), Int(0))
If ierror Then End
End If
If Right$(GrapherAppDirectory$, 1) = "\" Then GrapherAppDirectory$ = Left$(GrapherAppDirectory$, Len(GrapherAppDirectory$) - 1) ' remove trailing backslash
'If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "SurferAppDirectory"
lpDefault$ = vbNullString       ' no default!
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then
SurferAppDirectory$ = Left$(lpReturnString$, valid&)
Else
Call InitDetermineScripterPath(Int(2), Int(0))
If ierror Then End
End If
If Right$(SurferAppDirectory$, 1) = "\" Then SurferAppDirectory$ = Left$(SurferAppDirectory$, Len(SurferAppDirectory$) - 1) ' remove trailing backslash
'If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "SurferPlotsPerPagePolygon"
nDefault& = 3        ' must be 1, 3 or 8
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
SurferPlotsPerPagePolygon% = valid&
If SurferPlotsPerPagePolygon% <> 1 And SurferPlotsPerPagePolygon% <> 3 And SurferPlotsPerPagePolygon% <> 8 Then
msg$ = "SurferPlotsPerPagePolygon keyword value is invalid in " & ProbeWinINIFile$ & ", (must be 1, 3 or 8)"
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
SurferPlotsPerPagePolygon% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "ProbeSoftwareInternetBrowseMethod"
nDefault& = 0        ' 0 = WWW, 1 = DVD
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
ProbeSoftwareInternetBrowseMethod% = valid&
If ProbeSoftwareInternetBrowseMethod% < 0 Or ProbeSoftwareInternetBrowseMethod% > 1 Then
msg$ = "ProbeSoftwareInternetBrowseMethod keyword value is invalid in " & ProbeWinINIFile$ & ", (must be 0 or 1)"
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
ProbeSoftwareInternetBrowseMethod% = CInt(nDefault&)
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "UseFluorescenceByBetaLinesFlag"
nDefault& = True    ' changed 04-20-2015
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
UseFluorescenceByBetaLinesFlag = True
Else
UseFluorescenceByBetaLinesFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

StrataGEMVersion = 6
lpAppName$ = "Software"
lpKeyName$ = "StrataGEMVersion"
lpDefault$ = "6.0"      ' default is version 6
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then StrataGEMVersion! = Val(Left$(lpReturnString$, valid&))
If StrataGEMVersion! < 4# Then
msg$ = "StrataGEMVersion keyword value is out of range in " & ProbeWinINIFile$ & ", (must be greater than or equal to 4)"
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
StrataGEMVersion! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Software"
lpKeyName$ = "UseConfirmDuringAcquisitionFlag"
nDefault& = True
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
AutomateConfirmFlag = True
Else
AutomateConfirmFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

Exit Sub

' Errors
InitINISoftwareError:
MsgBox Error$, vbOKOnly + vbCritical, "InitINISoftware"
ierror = True
Exit Sub

End Sub

Sub InitINIHardware()
' Open the PROBEWIN.INI file and read defaults

ierror = False
On Error GoTo InitINIHardwareError

Dim valid As Long, tValid As Long

Dim lpAppName As String
Dim lpKeyName As String
Dim lpDefault As String
Dim lpFileName As String
Dim lpReturnString As String * 255
Dim lpReturnString2 As String * 255

Dim nSize As Long
Dim nDefault As Long
Dim tcomment As String

' Check for existing PROBEWIN.INI
If Dir$(ProbeWinINIFile$) = vbNullString Then
msg$ = "Unable to open file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If

' Use Windows API function to read PROBEWIN.INI
lpFileName$ = ProbeWinINIFile$
nSize& = Len(lpReturnString$)

' Hardware section, first get interface type
lpAppName$ = "Hardware"
lpKeyName$ = "InterfaceType"
nDefault& = 0   ' 0=Demo, 1=Unused, 2=JEOL 8900/8200/8500/8x30, 3=Unused, 4=Unused, 5=SX100, 6=Axioscope
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
InterfaceType% = valid&
If InterfaceType% < 0 Or InterfaceType% > MAXINTERFACE% Then
msg$ = "InterfaceType keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
If InterfaceType% = 1 Or InterfaceType% = 3 Or InterfaceType% = 4 Then
msg$ = "InterfaceType keyword value is no longer supported in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "NumberOfFixedSpecs"   ' obsolete but keep for now
nDefault& = 0
'tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
NumberOfFixedSpecs% = valid&
If NumberOfFixedSpecs% > 0 Then
msg$ = "NumberOfFixedSpecs keyword is no longer supported in Probe for EPMA (must be set to zero) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
'If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "NumberOfTunableSpecs"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
NumberOfTunableSpecs% = valid&
If NumberOfTunableSpecs% < 0 Or NumberOfTunableSpecs% > MAXSPEC% Then
msg$ = "NumberOfTunableSpecs keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Read number of stage motors (x, y and z but w is disabled)
lpAppName$ = "Hardware"
lpKeyName$ = "NumberOfStageMotors"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
NumberOfStageMotors% = valid&
If NumberOfStageMotors% < 0 Or NumberOfStageMotors% > MAXAXES% - 1 Or NumberOfStageMotors% = 1 Then
msg$ = "NumberOfStageMotors keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Load motor pointers
XMotor% = NumberOfTunableSpecs% + 1
YMotor% = NumberOfTunableSpecs% + 2
ZMotor% = NumberOfTunableSpecs% + 3
WMotor% = NumberOfTunableSpecs% + 4

' Check for total number of spectrometers
If NumberOfTunableSpecs% < 1 Then
msg$ = "Warning- No spectrometers defined in " & ProbeWinINIFile$
Call IOWriteLog(msg$)
End If

' Check for total number of stage motors
If NumberOfStageMotors% < 1 Then
msg$ = "Warning- No stage motors defined in " & ProbeWinINIFile$
Call IOWriteLog(msg$)
End If

If NumberOfTunableSpecs% > MAXSPEC% Then
msg$ = "Too many total spectrometers defined in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If

' Joystick parameters used for FormMOVE increment spin buttons.
lpAppName$ = "Hardware"
lpKeyName$ = "JoyStickXPolarity"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
JoyStickXPolarity% = valid&
If JoyStickXPolarity% <> 0 Then
JoyStickXPolarity% = True
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "JoyStickYPolarity"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
JoyStickYPolarity% = valid&
If JoyStickYPolarity% <> 0 Then
JoyStickYPolarity% = True
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "JoyStickZPolarity"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
JoyStickZPolarity% = valid&
If JoyStickZPolarity% <> 0 Then
JoyStickZPolarity% = True
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' JEOL specific parameters
lpAppName$ = "Hardware"
lpKeyName$ = "JeolEOSInterfaceType"
nDefault& = 0  ' 1=8200,8500, 2=8900, 3=8230,8530
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
JeolEOSInterfaceType& = valid&
If InterfaceType% = 2 Then  ' only check if JEOL 8200/8900 tcp/ip direct socket
If JeolEOSInterfaceType& < 1 Or JeolEOSInterfaceType& > 3 Then
msg$ = "JEOLEOSInterfaceType keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "JEOLEIKSVersionNumber"
lpDefault$ = "4"        ' 3 = 2009, 4 = 2011, 5 = 2012
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then JEOLEIKSVersionNumber! = Val(Left$(lpReturnString$, valid&))
If JEOLEIKSVersionNumber! < 3# Or JEOLEIKSVersionNumber! > 6# Then
msg$ = "JEOLEIKSVersionNumber keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
JEOLEIKSVersionNumber! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "ThermalFieldEmissionPresent"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ThermalFieldEmissionPresentFlag% = True     ' 8500, 8530
Else
ThermalFieldEmissionPresentFlag% = False    ' 8200, 8230
End If
If InterfaceType% = 2 Then  ' only check if JEOL 8900 tcp/ip direct socket
If JeolEOSInterfaceType& = 2 And ThermalFieldEmissionPresentFlag Then
msg$ = "JEOLEOSInterfaceType = 2 (8900) and ThermalFieldEmissionPresent flag <> 0 are incompatible in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "Jeol8900PreAcquireString"
lpDefault$ = vbNullString
If InterfaceType% = 0 Or JeolEOSInterfaceType& = 2 Then lpDefault$ = "PB OFF"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then Jeol8900PreAcquireString$ = Left$(lpReturnString$, valid&)
If Len(Jeol8900PreAcquireString$) > 64 Then
msg$ = "JEOL8900PreAcquireString is too long in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
lpDefault$ = Jeol8900PreAcquireString$
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "Jeol8900PostAcquireString"
lpDefault$ = vbNullString
If InterfaceType% = 0 Or JeolEOSInterfaceType& = 2 Then lpDefault$ = "PB ON"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then Jeol8900PostAcquireString$ = Left$(lpReturnString$, valid&)
If Len(Jeol8900PostAcquireString$) > 64 Then
msg$ = "JEOL8900PostAcquireString is too long in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
lpDefault$ = Jeol8900PostAcquireString$
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Backlash defaults
lpAppName$ = "Hardware"
lpKeyName$ = "SpecBacklashFlag"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
SpecBackLashFlag = True
Else
SpecBackLashFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "StageBacklashFlag"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
StageBacklashFlag = True
StageStdBacklashFlag = True
StageUnkBacklashFlag = True
StageWavBacklashFlag = True
Else
StageBacklashFlag = False
StageStdBacklashFlag = False
StageUnkBacklashFlag = False
StageWavBacklashFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Backlash types
lpAppName$ = "Hardware"
lpKeyName$ = "SpecBacklashType"
nDefault& = 1
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
SpecBackLashType% = valid&
If SpecBackLashType% < 1 Or SpecBackLashType% > 2 Then
msg$ = "SpectrometerBacklashType keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
SpecBackLashType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "StageBacklashType"
nDefault& = 1
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
StageBacklashType% = valid&
If StageBacklashType% < 1 Or StageBacklashType% > 2 Then
msg$ = "StageBacklashType keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
StageBacklashType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Filament standby
lpAppName$ = "Hardware"
lpKeyName$ = "FilamentStandbyPresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
FilamentStandbyPresent% = True
Else
FilamentStandbyPresent% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "FilamentStandbyType"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
FilamentStandbyType% = valid&
If FilamentStandbyType% < 0 Or FilamentStandbyType% > 2 Then
msg$ = "FilamentStandbyType keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
FilamentStandbyType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' EDS Spectra interface parameters
lpAppName$ = "Hardware"
lpKeyName$ = "EDSSpectraInterfacePresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
EDSSpectraInterfacePresent% = True
Else
EDSSpectraInterfacePresent% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "EDSSpectraInterfaceType"
nDefault& = 0   ' 0 = Demo, 1 = Unused, 2 = Bruker, 3 = Oxford, 4 = Unused, 5 = Thermo, 6 = JEOL OEM
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
EDSSpectraInterfaceType% = valid&
If EDSSpectraInterfaceType% < 0 Or EDSSpectraInterfaceType% > MAXINTERFACE_EDS% Then
msg$ = "EDSSpectraInterfaceType keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
EDSSpectraInterfaceType% = nDefault&
End If
If EDSSpectraInterfaceType% = 1 Or EDSSpectraInterfaceType% = 3 Or EDSSpectraInterfaceType% = 4 Then
msg$ = "EDSSpectraInterfaceType (" & Format$(EDSSpectraInterfaceType%) & ") is not supported by Probe for EPMA in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "EDSSpectraNetIntensityInterfaceType"
nDefault& = EDSSpectraInterfaceType%   ' 0 = Demo, 1 = Unused, 2 = Bruker, 3 = Oxford, 4 = Unused, 5 = Thermo, 6 = JEOL OEM
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
EDSSpectraNetIntensityInterfaceType% = valid&
If EDSSpectraNetIntensityInterfaceType% < 0 Or EDSSpectraNetIntensityInterfaceType% > MAXINTERFACE_EDS% Then
msg$ = "EDSSpectraNetIntensityInterfaceType keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
EDSSpectraNetIntensityInterfaceType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "EDSThinWindowPresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
EDSThinWindowPresent% = True
Else
EDSThinWindowPresent% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' EDS TCP/IP socket parameters
If EDSSpectraInterfaceType% = 5 Then    ' Thermo
lpAppName$ = "Hardware"
lpKeyName$ = "EDS_IPAddress"
lpDefault$ = vbNullString
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then EDS_IPAddress$ = Left$(lpReturnString$, valid&)
If Trim$(EDS_IPAddress$) = vbNullString Then
If EDSSpectraInterfaceType% = 5 Then msg$ = "EDS_IPAddress keyword value (Thermo NSS interface) is blank in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
lpDefault$ = Trim$(EDS_IPAddress$)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "EDS_ServicePort"         ' not used by any EDS vendor currently
lpDefault$ = vbNullString
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then EDS_ServicePort$ = Left$(lpReturnString$, valid&)
End If
lpDefault$ = Trim$(EDS_ServicePort$)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

If EDSSpectraInterfaceType% = 2 Then   ' used by Bruker
lpAppName$ = "Hardware"
lpKeyName$ = "EDS_ServerName"
lpDefault$ = "Local Server"            ' for remote clients use server name defined in Bruker app Configuration menu
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then EDS_ServerName$ = Left$(lpReturnString$, valid&)
lpDefault$ = Trim$(EDS_ServerName$)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "EDS_LoginName"
lpDefault$ = "edx"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then EDS_LoginName$ = Left$(lpReturnString$, valid&)
lpDefault$ = Trim$(EDS_LoginName$)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "EDS_LoginPassword"
lpDefault$ = "edx"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then EDS_LoginPassword$ = Left$(lpReturnString$, valid&)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' WDS TCP/IP (direct socket) parameters for JEOL 8900/8200/8500/8230/8530 and Cameca SX100
If InterfaceType = 2 Or InterfaceType = 5 Then    ' JEOL 8900/8200/8500/8230/8530 (direct socket) and SX100
lpAppName$ = "Hardware"
lpKeyName$ = "WDS_IPAddress"    ' for device EPMA (JEOL system controller or Cameca SX100)
lpDefault$ = vbNullString
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then WDS_IPAddress$ = Left$(lpReturnString$, valid&)
If Trim$(WDS_IPAddress$) = vbNullString Then
msg$ = "WDS_IPAddress keyword value (TCP/IP interface) is blank in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
lpDefault$ = Trim$(WDS_IPAddress$)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

If InterfaceType% = 2 Then      ' JEOL 8900/8200/8500/8230/8530 only (SX100 uses single port)
lpAppName$ = "Hardware"
lpKeyName$ = "WDS_IPAddress2"   ' for EOS/Notify JEOL instruments (not used by Cameca SX100)
lpDefault$ = vbNullString
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then WDS_IPAddress2$ = Left$(lpReturnString$, valid&)
If Trim$(WDS_IPAddress2$) = vbNullString Then
msg$ = "WDS_IPAddress2 keyword value (TCP/IP interface) is blank in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
lpDefault$ = WDS_IPAddress2$
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)
End If

' JEOL direct socket interface (SC socket for 8900 and 8200, port for SX100)
lpAppName$ = "Hardware"
lpKeyName$ = "WDS_ServicePort"    ' for JEOL SC (8900/8200/8500/8230/8530) or SX100
nDefault& = 0                     ' should be 2785 for JEOL 8900/8200/8500, 22200 for 8230/8530, 4000 for SX100
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
WDS_ServicePort% = CInt(valid&)
If WDS_ServicePort% = 0 Then
msg$ = "WDS_ServicePort keyword value (TCP/IP interface) is blank in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
If InterfaceType% = 2 And JeolEOSInterfaceType& < 3 And WDS_ServicePort% <> 2785 Then   ' 8900/8200/8500
msg$ = "WDS_ServicePort keyword value (TCP/IP interface) for JEOL 8900/8200/8500 is not equal to 2785 in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
If InterfaceType% = 2 And JeolEOSInterfaceType& = 3 And WDS_ServicePort% <> 22200 Then  ' 8230/8530
msg$ = "WDS_ServicePort keyword value (TCP/IP interface) for JEOL 8230/8530 is not equal to 22200 in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
If InterfaceType% = 5 And WDS_ServicePort% <> 4000 Then ' SX100
msg$ = "WDS_ServicePort keyword value (TCP/IP interface) for SX100 is not equal to 4000 in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' JEOL direct socket interface (Driver socket for 8900, EOS socket for 8200/8500, Notify socket for 8230/8530)
If InterfaceType% = 2 Then      ' JEOL only
lpAppName$ = "Hardware"
lpKeyName$ = "WDS_ServicePort2"     ' EOS or Notify socket
nDefault& = 0                       ' should be 2785 for 8900, 22200 for 8200/8500 and 22210 for 8230/8530
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
WDS_ServicePort2% = CInt(valid&)
If WDS_ServicePort2% = 0 Then
msg$ = "WDS_ServicePort2 keyword value (TCP/IP interface) is blank in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
If JeolEOSInterfaceType& = 1 And WDS_ServicePort2% <> 22200 Then  ' 8200/8500
msg$ = "WDS_ServicePort2 keyword value (TCP/IP interface) is not equal to 22200 in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
If JeolEOSInterfaceType& = 2 And WDS_ServicePort2% <> 2785 Then  ' 8900
msg$ = "WDS_ServicePort2 keyword value (TCP/IP interface) is not equal to 2785 in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
If JeolEOSInterfaceType& = 3 And WDS_ServicePort2% <> 22210 Then  ' 8230/8530
msg$ = "WDS_ServicePort2 keyword value (TCP/IP interface) is not equal to 22210 in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End
End If
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)
End If

' Operating voltage
lpAppName$ = "Hardware"
lpKeyName$ = "OperatingVoltagePresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
OperatingVoltagePresent% = True
Else
OperatingVoltagePresent% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "OperatingVoltageType"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
OperatingVoltageType% = valid&
If OperatingVoltageType% < 0 Or OperatingVoltageType% > 2 Then
msg$ = "OperatingVoltageType keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
OperatingVoltageType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Get operating voltage tolerance
lpAppName$ = "Hardware"
lpKeyName$ = "OperatingVoltageTolerance"
lpDefault$ = "0.002"    ' .2%
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then OperatingVoltageTolerance! = Val(Left$(lpReturnString$, valid&))
If OperatingVoltageTolerance! < 0.0001 Or OperatingVoltageTolerance! > 0.1 Then
msg$ = "OperatingVoltageTolerance keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
OperatingVoltageTolerance! = Val(lpDefault$)
End If
tValid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Beam current control
lpAppName$ = "Hardware"
lpKeyName$ = "BeamCurrentPresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
BeamCurrentPresent% = True
Else
BeamCurrentPresent% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "BeamCurrentType"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
BeamCurrentType% = valid&
If BeamCurrentType% < 0 Or BeamCurrentType% > 2 Then
msg$ = "BeamCurrentType keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
BeamCurrentType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Get beam current tolerance
lpAppName$ = "Hardware"
lpKeyName$ = "BeamCurrentTolerance"
lpDefault$ = "0.02"    ' 2%
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then BeamCurrentTolerance! = Val(Left$(lpReturnString$, valid&))
If BeamCurrentTolerance! < 0.00001 Or BeamCurrentTolerance! > 0.2 Then
msg$ = "BeamCurrentTolerance keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
BeamCurrentTolerance! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "BeamCurrentToleranceSet"
lpDefault$ = "0.01"    ' 1% (only for 8200 and 8900)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then BeamCurrentToleranceSet! = Val(Left$(lpReturnString$, valid&))
If BeamCurrentToleranceSet! < 0.0001 Or BeamCurrentToleranceSet! > 0.1 Then
msg$ = "BeamCurrentToleranceSet keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
BeamCurrentToleranceSet! = Val(lpDefault$)
If BeamCurrentToleranceSet! >= BeamCurrentTolerance! Then
msg$ = "BeamCurrentToleranceSet keyword value must be less than BeamCurrentTolerance keyword value in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
End If
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Beam size control
lpAppName$ = "Hardware"
lpKeyName$ = "BeamSizePresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
BeamSizePresent% = True
Else
BeamSizePresent% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "BeamSizeType"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
BeamSizeType% = valid&
If BeamSizeType% < 0 Or BeamSizeType% > 2 Then
msg$ = "BeamSizeType keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
BeamSizeType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Magnification control
lpAppName$ = "Hardware"
lpKeyName$ = "MagnificationPresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
MagnificationPresent% = True
Else
MagnificationPresent% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "MagnificationType"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
BeamSizeType% = valid&
If MagnificationType% < 0 Or MagnificationType% > 2 Then
msg$ = "MagnificationType keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
MagnificationType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' AutoFocus control
lpAppName$ = "Hardware"
lpKeyName$ = "AutoFocusPresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
AutoFocusPresent% = True
Else
AutoFocusPresent% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "AutoFocusType"
nDefault& = 2       ' 0 = parabolic, 1 = Gaussian, 2 = maximum value
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
AutoFocusType% = valid&
If AutoFocusType% < 0 Or AutoFocusType% > 2 Then
msg$ = "AutoFocusType keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
AutoFocusType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "AutoFocusOffset"
lpDefault$ = "0.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then AutoFocusOffset! = Val(Left$(lpReturnString$, valid&))
If Abs(AutoFocusOffset!) > 100# Then  ' must be less than 100 um
msg$ = "AutoFocusOffset keyword value (in microns) out of range (must be less than 100 um) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
AutoFocusOffset! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "AutoFocusMaxDeviation"
lpDefault$ = "30.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then AutoFocusMaxDeviation! = Val(Left$(lpReturnString$, valid&))
If AutoFocusMaxDeviation! < 1# Or AutoFocusMaxDeviation! > 80# Then
msg$ = "AutoFocusMaxDeviation keyword value (in percent) out of range (must be between 1 and 80) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
AutoFocusMaxDeviation! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "AutoFocusThresholdFraction"
lpDefault$ = "0.33"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then AutoFocusThresholdFraction! = Val(Left$(lpReturnString$, valid&))
If AutoFocusThresholdFraction! < 0.1 Or AutoFocusThresholdFraction > 0.9 Then
msg$ = "AutoFocusThresholdFraction keyword value out of range (must be between 0.1 and 0.9) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
AutoFocusThresholdFraction! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "AutoFocusMinimumPtoB"
lpDefault$ = "1.4"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then AutoFocusMinimumPtoB! = Val(Left$(lpReturnString$, valid&))
If AutoFocusMinimumPtoB! < 1.05 Or AutoFocusMinimumPtoB! > 10# Then
msg$ = "AutoFocusMinimumPtoB keyword value out of range (must be between 1.05 and 10.0) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
AutoFocusMinimumPtoB! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "AutoFocusRangeFineScan"
lpDefault$ = "100"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then AutoFocusRangeFineScan! = Val(Left$(lpReturnString$, valid&))
If AutoFocusRangeFineScan! < 10# Or AutoFocusRangeFineScan! > 1000# Then
msg$ = "AutoFocusRangeFineScan keyword value out of range (must be between 10 and 1000 um) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
AutoFocusRangeFineScan! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "AutoFocusRangeCoarseScan"
lpDefault$ = "600"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then AutoFocusRangeCoarseScan! = Val(Left$(lpReturnString$, valid&))
If AutoFocusRangeCoarseScan! < 60# Or AutoFocusRangeCoarseScan! > 6000# Then
msg$ = "AutoFocusRangeCoarseScan keyword value out of range (must be between 60 and 6000 um) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
AutoFocusRangeCoarseScan! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "AutoFocusPointsFineScan"
lpDefault$ = "200"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then AutoFocusPointsFineScan& = Val(Left$(lpReturnString$, valid&))
If AutoFocusPointsFineScan& < 50 Or AutoFocusPointsFineScan& > MAXAUTOFOCUSPOINTS& Then
msg$ = "AutoFocusPointsFineScan keyword value out of range (must be between 50 and " & Str$(MAXAUTOFOCUSPOINTS&) & " points) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
AutoFocusPointsFineScan& = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "AutoFocusPointsCoarseScan"
lpDefault$ = "100"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then AutoFocusPointsCoarseScan& = Val(Left$(lpReturnString$, valid&))
If AutoFocusPointsCoarseScan& < 50 Or AutoFocusPointsCoarseScan& > MAXAUTOFOCUSPOINTS& Then
msg$ = "AutoFocusPointsCoarseScan keyword value out of range (must be between 50 and " & Str$(MAXAUTOFOCUSPOINTS&) & " points) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
AutoFocusPointsCoarseScan& = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "AutoFocusTimeFineScan"
lpDefault$ = "20"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then AutoFocusTimeFineScan% = Val(Left$(lpReturnString$, valid&))
If AutoFocusTimeFineScan% < 1 Or AutoFocusTimeFineScan% > 500 Then
msg$ = "AutoFocusTimeFineScan keyword value out of range (must be between 1 and 500 msec) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
AutoFocusTimeFineScan% = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "AutoFocusTimeCoarseScan"
lpDefault$ = "20"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then AutoFocusTimeCoarseScan% = Val(Left$(lpReturnString$, valid&))
If AutoFocusTimeCoarseScan% < 1 Or AutoFocusTimeCoarseScan% > 500 Then
msg$ = "AutoFocusTimeCoarseScan keyword value out of range (must be between 1 and 500 msec) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
AutoFocusTimeCoarseScan% = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' ROM based peaking control
lpAppName$ = "Hardware"
lpKeyName$ = "ROMPeakingPresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ROMPeakingPresent% = True
Else
ROMPeakingPresent% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Check for correct default peak method
If Not ROMPeakingPresent And DefaultPeakCenterMethod% > 1 Then
msg$ = "ROM Peaking was specified as the default peak center method, however ROM Peaking is specified as not available in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
DefaultPeakCenterMethod% = 1
End If

lpAppName$ = "Hardware"
lpKeyName$ = "ROMPeakingType"
nDefault& = 2   ' maxima
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DefaultROMPeakingType% = valid&
If InterfaceType% = 0 Or InterfaceType% = 2 Then
If DefaultROMPeakingType% = 0 Then DefaultROMPeakingType% = 2    ' maxima is default for JEOL 8200/8900
End If
If DefaultROMPeakingType% < 0 Or DefaultROMPeakingType% > MAXROMPEAKTYPES% Then    ' 0=internal, 1=parabolic, 2=maxima, 3=gaussian, 4 = smart1, 5 = smart2, 6 = highest
msg$ = "ROMPeakingType keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
DefaultROMPeakingType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "ROMPeakingParabolicThresholdFraction"
lpDefault$ = "0.33"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then ROMPeakingParabolicThresholdFraction! = Val(Left$(lpReturnString$, valid&))
If ROMPeakingParabolicThresholdFraction! < 0.1 Or ROMPeakingParabolicThresholdFraction > 0.9 Then
msg$ = "ROMPeakingParabolicThresholdFraction keyword value out of range (must be between 0.1 and 0.9) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
ROMPeakingParabolicThresholdFraction! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "ROMPeakingMaximaThresholdFraction"
lpDefault$ = "0.33"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then ROMPeakingMaximaThresholdFraction! = Val(Left$(lpReturnString$, valid&))
If ROMPeakingMaximaThresholdFraction! < 0.1 Or ROMPeakingMaximaThresholdFraction > 0.9 Then
msg$ = "ROMPeakingMaximaThresholdFraction keyword value out of range (must be between 0.1 and 0.9) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
ROMPeakingMaximaThresholdFraction! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "ROMPeakingGaussianThresholdFraction"
lpDefault$ = "0.33"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then ROMPeakingGaussianThresholdFraction! = Val(Left$(lpReturnString$, valid&))
If ROMPeakingGaussianThresholdFraction! < 0.1 Or ROMPeakingGaussianThresholdFraction > 0.9 Then
msg$ = "ROMPeakingGaussianThresholdFraction keyword value out of range (must be between 0.1 and 0.9) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
ROMPeakingGaussianThresholdFraction! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "ROMPeakingMaxDeviation"
lpDefault$ = "20"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then ROMPeakingMaxDeviation! = Val(Left$(lpReturnString$, valid&))
If ROMPeakingMaxDeviation! < 5 Or ROMPeakingMaxDeviation > 80# Then
msg$ = "ROMPeakingMaxDeviation keyword value (in percent) out of range (must be between 5 and 80) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware"
ROMPeakingMaxDeviation! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

Exit Sub

' Errors
InitINIHardwareError:
MsgBox Error$, vbOKOnly + vbCritical, "InitINIHardware"
ierror = True
Exit Sub

End Sub

Sub InitINIHardware2()
' Open the PROBEWIN.INI file and read defaults (2nd half of InitINIHardware)

ierror = False
On Error GoTo InitINIHardware2Error

Dim i As Integer
Dim valid As Long, tValid As Long

Dim lpAppName As String
Dim lpKeyName As String
Dim lpDefault As String
Dim lpFileName As String
Dim lpReturnString As String * 255
Dim lpReturnString2 As String * 255

Dim nSize As Long
Dim nDefault As Long
Dim astring As String, tcomment As String

' Check for existing PROBEWIN.INI
If Dir$(ProbeWinINIFile$) = vbNullString Then
msg$ = "Unable to open file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End
End If

' Use Windows API function to read PROBEWIN.INI
lpFileName$ = ProbeWinINIFile$
nSize& = Len(lpReturnString$)

lpAppName$ = "Hardware"
lpKeyName$ = "ReadOnlySpecPositions"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ReadOnlySpecPositions% = True
Else
ReadOnlySpecPositions% = False
End If
If InterfaceType% <> 6 Then ReadOnlySpecPositions% = False     ' only applies to axioscope interface
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "ReadOnlyStagePositions"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ReadOnlyStagePositions% = True
Else
ReadOnlyStagePositions% = False
End If
If InterfaceType% <> 0 And InterfaceType% <> 6 Then ReadOnlyStagePositions% = False     ' only applies to axioscope interface
If InterfaceType% = 6 And Not ReadOnlyStagePositions% Then                              ' check if axioscope
msg$ = "ReadOnlyStagePositions keyword value must be non-zero for the specified InterfaceType in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "ColumnConditionPresent"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ColumnConditionPresent% = True
Else
ColumnConditionPresent% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

If ColumnConditionPresent% Then
lpAppName$ = "Hardware"
lpKeyName$ = "ColumnConditionType"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
ColumnConditionType% = valid&
If ColumnConditionType% < 0 Or ColumnConditionType% > 2 Then
msg$ = "ColumnConditionType keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
ColumnConditionType% = nDefault&
End
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)
End If

If ColumnConditionPresent% Then
lpAppName$ = "Hardware"
lpKeyName$ = "ColumnConditionMethod"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DefaultColumnConditionMethod% = valid&
If DefaultColumnConditionMethod% < 0 Or DefaultColumnConditionMethod% > 1 Then
msg$ = "ColumnConditionMethod keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
DefaultColumnConditionMethod% = nDefault&
End
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)
End If

If ColumnConditionPresent% Then
lpAppName$ = "Hardware"
lpKeyName$ = "ColumnConditionString"
lpDefault$ = vbNullString
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultColumnConditionString$ = Left$(lpReturnString$, valid&)
If Len(DefaultColumnConditionString$) > DbTextFilenameLength% Then
msg$ = "ColumnConditionString keyword value is too long in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End
End If
lpDefault$ = DefaultColumnConditionString$
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)
End If

' Scan rotation control
lpAppName$ = "Hardware"
lpKeyName$ = "ScanRotationPresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ScanRotationPresent% = True
Else
ScanRotationPresent% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "ScanRotation"
lpDefault$ = "0.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultScanRotation! = Val(Left$(lpReturnString$, valid&))
If DefaultScanRotation! < 0# Or DefaultScanRotation! >= 360# Then
msg$ = "DefaultScanRotation keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
DefaultScanRotation! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Detector parameters
lpAppName$ = "Hardware"
lpKeyName$ = "DetectorSlitSizePresent"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
DetectorSlitSizePresent% = True
Else
DetectorSlitSizePresent% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

If DetectorSlitSizePresent% Then
lpAppName$ = "Hardware"
lpKeyName$ = "DetectorSlitSizeType"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DetectorSlitSizeType% = valid&
If DetectorSlitSizeType% < 0 Or DetectorSlitSizeType% > 0 Then
msg$ = "DetectorSlitSizeType keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
DetectorSlitSizeType% = nDefault&
End
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)
End If

lpAppName$ = "Hardware"
lpKeyName$ = "DetectorSlitPositionPresent"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
DetectorSlitPositionPresent% = True
Else
DetectorSlitPositionPresent% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

If DetectorSlitPositionPresent% Then
lpAppName$ = "Hardware"
lpKeyName$ = "DetectorSlitPositionType"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DetectorSlitPositionType% = valid&
If DetectorSlitPositionType% < 0 Or DetectorSlitPositionType% > 0 Then
msg$ = "DetectorSlitPositionType keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
DetectorSlitPositionType% = nDefault&
End
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)
End If

lpAppName$ = "Hardware"
lpKeyName$ = "DetectorModePresent"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
DetectorModePresent% = True
Else
DetectorModePresent% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

If DetectorModePresent% Then
lpAppName$ = "Hardware"
lpKeyName$ = "DetectorModeType"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DetectorModeType% = valid&
If DetectorModeType% < 0 Or DetectorModeType% > 0 Then
msg$ = "DetectorModeType keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
DetectorModeType% = nDefault&
End
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)
End If

' Tilt/Rotation
lpAppName$ = "Hardware"
lpKeyName$ = "TiltRotationPresent"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
TiltRotationPresent% = True
Else
TiltRotationPresent% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

If TiltRotationPresent% Then
lpAppName$ = "Hardware"
lpKeyName$ = "TiltRotationType"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
TiltRotationType% = valid&
If TiltRotationType% < 0 Or TiltRotationType% > 0 Then
msg$ = "TiltRotationType keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
TiltRotationType% = nDefault&
End
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)
End If

' Move all stage motors flag (JEOL 8900 only)
lpAppName$ = "Hardware"
lpKeyName$ = "MoveAllStageMotorsHardwarePresent"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
MoveAllStageMotorsHardwarePresent% = True
Else
MoveAllStageMotorsHardwarePresent% = False
End If
If Not MoveAllStageMotorsHardwarePresent% And InterfaceType% = 2 Then
msg$ = "JEOL 8900/8200/8500/8230/8530 interface requires the MoveAllStageMotorsHardwarePresent  keyword value to be non-zero in file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "AlwaysPollFaradayCupState"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
AlwaysPollFaradayCupStateFlag% = True
Else
AlwaysPollFaradayCupStateFlag% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "DriverLoggingLevel"
nDefault& = 0  ' 0 - disabled, 1 - basic logging, 2 - detailed
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DriverLoggingLevel& = valid&
If DriverLoggingLevel& < 0 Or DriverLoggingLevel& > 2 Then
msg$ = "DriverLoggingLevel keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "JeolCondenserNumberOfApertures"
nDefault& = 1     ' default is 1 aperture (aperture #1)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
JeolCondenserNumberOfApertures% = valid&
If JeolCondenserNumberOfApertures% < 1 Or JeolCondenserNumberOfApertures% > 4 Then
msg$ = "JEOLCondenserNumberOfApertures keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Check that default aperature is within number of JEOL aperatures if demo or JEOL interface
If InterfaceType% = 0 Or InterfaceType% = 2 Then
If DefaultAperture% > JeolCondenserNumberOfApertures% Then
msg$ = "DefaultAperture keyword value (" & Format$(DefaultAperture%) & ") in [general] section is greater than the number of JEOLCondenserNumberOfApertures keyword value (" & Format$(JeolCondenserNumberOfApertures%) & ") defined in the [hardware] section in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End
End If
End If

lpAppName$ = "Hardware"
lpKeyName$ = "JeolCondenserCoarseCalibrationSettingLow"
lpDefault$ = "20"     ' for 8900 and 8200 (note that JEOL 8200 display of condenser values are inverted)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
astring$ = Left$(lpReturnString$, valid&)
Call InitParseStringToInteger(astring$, JeolCondenserNumberOfApertures%, JeolCondenserCoarseCalibrationSettingLow%())
For i% = 1 To JeolCondenserNumberOfApertures%
If JeolCondenserCoarseCalibrationSettingLow%(i%) < 1 Or JeolCondenserCoarseCalibrationSettingLow%(i%) > 90 Then
msg$ = "JEOLCondenserCoarseCalibrationSettingLow keyword value (aperture " & Format$(i%) & ") out of range (must be between 1 and 90) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End
End If
Next i%
lpDefault$ = Trim$(lpReturnString$)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "JeolCondenserCoarseCalibrationSettingMedium"
lpDefault$ = "35"     ' for 8900 and 8200 (note that JEOL 8200 display of condenser values are inverted)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
astring$ = Left$(lpReturnString$, valid&)
Call InitParseStringToInteger(astring$, JeolCondenserNumberOfApertures%, JeolCondenserCoarseCalibrationSettingMedium%())
For i% = 1 To JeolCondenserNumberOfApertures%
If JeolCondenserCoarseCalibrationSettingMedium%(i%) < 1 Or JeolCondenserCoarseCalibrationSettingMedium%(i%) > 90 Then
msg$ = "JEOLCondenserCoarseCalibrationSettingMedium keyword value (aperture " & Format$(i%) & ") out of range (must be between 1 and 90) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End
End If
Next i%
lpDefault$ = Trim$(lpReturnString$)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "JeolCondenserCoarseCalibrationSettingHigh"
lpDefault$ = "45"     ' for 8900 and 8200 (note that JEOL 8200 display of condenser values are inverted)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
astring$ = Left$(lpReturnString$, valid&)
Call InitParseStringToInteger(astring$, JeolCondenserNumberOfApertures%, JeolCondenserCoarseCalibrationSettingHigh%())
For i% = 1 To JeolCondenserNumberOfApertures%
If JeolCondenserCoarseCalibrationSettingHigh%(i%) < 1 Or JeolCondenserCoarseCalibrationSettingHigh%(i%) > 90 Then
msg$ = "JEOLCondenserCoarseCalibrationSettingHigh keyword value (aperture " & Format$(i%) & ") out of range (must be between 1 and 90) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End
End If
Next i%
lpDefault$ = Trim$(lpReturnString$)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "JeolCondenserFineCalibrationSetting"
nDefault& = 128     ' for 8900 (0-255)
If JeolEOSInterfaceType& = 1 Then   ' for 8200/8500
If Not ThermalFieldEmissionPresentFlag Then nDefault& = 255    ' JEOL 8200 W/LBG: 0-511 (actually 511-0)
If ThermalFieldEmissionPresentFlag Then nDefault& = 255        ' JEOL 8500 TFE: 0-1023 (actually 511-0)
End If
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
JeolCondenserFineCalibrationSetting% = valid&

If JeolEOSInterfaceType& = 1 Then   ' for 8200/8500
If ThermalFieldEmissionPresentFlag Then                         ' 8500
If JeolCondenserFineCalibrationSetting% < 50 Or JeolCondenserFineCalibrationSetting% > 500 Then
msg$ = "JEOLCondenserFineCalibrationSetting keyword value (8500) out of range (must be between 50 and 500) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End
End If

Else                                                            ' 8200
If JeolCondenserFineCalibrationSetting% < 50 Or JeolCondenserFineCalibrationSetting% > 500 Then
msg$ = "JEOLCondenserFineCalibrationSetting keyword value (8200) out of range (must be between 50 and 500) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End
End If
End If

ElseIf JeolEOSInterfaceType& = 2 Then                           ' 8900
If JeolCondenserFineCalibrationSetting% < 20 Or JeolCondenserFineCalibrationSetting% > 200 Then
msg$ = "JEOL Condenser Fine CalibrationSetting keyword value (8900) out of range (must be between 20 and 200) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End
End If

ElseIf JeolEOSInterfaceType& = 3 Then                           ' 8230/8530
If ThermalFieldEmissionPresentFlag Then                         ' 8530
If JeolCondenserFineCalibrationSetting% < 100 Or JeolCondenserFineCalibrationSetting% > 1000 Then
msg$ = "JEOLCondenserFineCalibrationSetting keyword value (8530) out of range (must be between 100 and 1000) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End
End If

Else
If JeolCondenserFineCalibrationSetting% < 100 Or JeolCondenserFineCalibrationSetting% > 1000 Then
msg$ = "JEOLCondenserFineCalibrationSetting keyword value (8230) out of range (must be between 100 and 1000) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End
End If
End If
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "JeolCondenserCoarseCalibrationMode"
nDefault& = 0   ' 0 = use internal calibration, 1 = use INI file calibration
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
JeolCondenserCoarseCalibrationMode% = valid&
If JeolCondenserCoarseCalibrationMode% < 0 Or JeolCondenserCoarseCalibrationMode% > 1 Then
msg$ = "JEOLCondenserCoarseCalibrationMode keyword value is out of range (must be 0 or 1) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
JeolCondenserCoarseCalibrationMode% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Load measured beam currents for the specified coarse condenser values
lpAppName$ = "Hardware"
lpKeyName$ = "JeolCondenserCoarseCalibrationBeamLow"
If JeolEOSInterfaceType& = 1 Then lpDefault$ = "199"     ' 8200
If JeolEOSInterfaceType& = 2 Then lpDefault$ = "199"     ' 8900
If JeolEOSInterfaceType& = 3 Then lpDefault$ = "199"     ' 8230/8530
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
astring$ = Left$(lpReturnString$, valid&)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
Call InitParseStringToReal(astring$, JeolCondenserNumberOfApertures%, JeolCondenserCoarseCalibrationBeamLow!())
For i% = 1 To JeolCondenserNumberOfApertures%
If JeolCondenserCoarseCalibrationBeamLow!(i%) < 0.001 Or JeolCondenserCoarseCalibrationBeamLow!(i%) > 1000# Then
msg$ = "JEOLCondenserCoarseCalibrationBeamLow keyword value (aperture " & Format$(i%) & ") out of range (must be between 0.001 and 1000) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End If
Next i%
lpDefault$ = Trim$(lpReturnString$)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "JeolCondenserCoarseCalibrationBeamMedium"
If JeolEOSInterfaceType& = 1 Then lpDefault$ = "24"   ' 8200/8500
If JeolEOSInterfaceType& = 2 Then lpDefault$ = "24"   ' 8900
If JeolEOSInterfaceType& = 3 Then lpDefault$ = "24"   ' 8230/8530
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
astring$ = Left$(lpReturnString$, valid&)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
Call InitParseStringToReal(astring$, JeolCondenserNumberOfApertures%, JeolCondenserCoarseCalibrationBeamMedium!())
For i% = 1 To JeolCondenserNumberOfApertures%
If JeolCondenserCoarseCalibrationBeamMedium!(i%) < 0.001 Or JeolCondenserCoarseCalibrationBeamMedium!(i%) > 1000# Then
msg$ = "JEOLCondenserCoarseCalibrationBeamMedium keyword value (aperture " & Format$(i%) & ") out of range (must be between 0.001 and 1000) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End If
Next i%
lpDefault$ = Trim$(lpReturnString$)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "JeolCondenserCoarseCalibrationBeamHigh"
If JeolEOSInterfaceType& = 1 Then lpDefault$ = "1.9"    ' 8200/8500
If JeolEOSInterfaceType& = 2 Then lpDefault$ = "1.9"    ' 8900
If JeolEOSInterfaceType& = 3 Then lpDefault$ = "1.9"    ' 8230/8530
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
astring$ = Left$(lpReturnString$, valid&)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
Call InitParseStringToReal(astring$, JeolCondenserNumberOfApertures%, JeolCondenserCoarseCalibrationBeamHigh!())
For i% = 1 To JeolCondenserNumberOfApertures%
If JeolCondenserCoarseCalibrationBeamHigh!(i%) < 0.001 Or JeolCondenserCoarseCalibrationBeamHigh!(i%) > 1000# Then
msg$ = "JEOLCondenserCoarseCalibrationBeamHigh keyword value (aperture " & Format$(i%) & ") out of range (must be between 0.001 and 1000) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End If
Next i%
lpDefault$ = Trim$(lpReturnString$)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "JeolCoarseCondenserCalibrationDelay"
lpDefault$ = "0.1"     ' 100 msec
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then JeolCoarseCondenserCalibrationDelay! = Val(Left$(lpReturnString$, valid&))
If JeolCoarseCondenserCalibrationDelay! < 0# Or JeolCoarseCondenserCalibrationDelay! > 4# Then
msg$ = "JEOLCoarseCondenserCalibrationDelay keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End
End If
lpDefault$ = Trim$(lpReturnString$)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "JeolMonitorInterval"
nDefault& = 400     ' 400 msec
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
JeolMonitorInterval& = valid&
If JeolMonitorInterval& < 100 Or JeolMonitorInterval& > 10000 Then
msg$ = "JEOLMonitorInterval keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Beammode control
lpAppName$ = "Hardware"
lpKeyName$ = "BeamModePresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
BeamModePresent% = True
Else
BeamModePresent% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "BeamModeType"         ' not used at this time
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
BeamSizeType% = valid&
If BeamModeType% < 0 Or BeamModeType% > 1 Then
msg$ = "BeamModeType keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
BeamModeType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Optical light intensity
lpAppName$ = "Hardware"
lpKeyName$ = "ReflectedLightIntensity"
If InterfaceType% = 5 Then
nDefault& = 32      ' SX100 (0-64)
Else
nDefault& = 63      ' JEOL (0-127)
End If
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DefaultReflectedLightIntensity% = valid&
If DefaultReflectedLightIntensity% < 0 Or DefaultReflectedLightIntensity% > 127 Or (InterfaceType% = 5 And DefaultReflectedLightIntensity% > 64) Then
msg$ = "ReflectedLightIntensity keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
DefaultReflectedLightIntensity% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "TransmittedLightIntensity"
If InterfaceType% = 5 Then
nDefault& = 32      ' SX100 (0-64)
Else
nDefault& = 63      ' JEOL (0-127)
End If
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DefaultTransmittedLightIntensity% = valid&
If DefaultTransmittedLightIntensity% < 0 Or DefaultTransmittedLightIntensity% > 127 Or (InterfaceType% = 5 And DefaultTransmittedLightIntensity% > 64) Then
msg$ = "TransmittedLightIntensity keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
DefaultTransmittedLightIntensity% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "DisableSpectrometerNumber"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DisableSpectrometerNumber% = valid&
If DisableSpectrometerNumber% < 0 Or DisableSpectrometerNumber% > NumberOfTunableSpecs% Then
msg$ = "DisableSpectrometerNumber keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
DisableSpectrometerNumber% = nDefault&
End If
If DisableSpectrometerNumber% > 0 Then
msg$ = "WARNING: spectrometer " & Str$(DisableSpectrometerNumber%) & " is disabled for all real time operations!"
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "SpectrometerROMScanMode"
nDefault& = 0           ' SX100 only, 0 = use absolute scan, 1 = use relative scan
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
SpectrometerROMScanMode% = valid&
If SpectrometerROMScanMode% < 0 Or SpectrometerROMScanMode% > 1 Then
msg$ = "SpectrometerROMScanMode keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
SpectrometerROMScanMode% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "FilamentWarmUpInterval"
lpDefault$ = "2.0"     ' 2 second default
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then FilamentWarmUpInterval! = Val(Left$(lpReturnString$, valid&))
If FilamentWarmUpInterval! < 0.1 Or FilamentWarmUpInterval! > 1000# Then
msg$ = "FilamentWarmUpInterval keyword value out of range (must be between 0.1 and 1000 seconds) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
End
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "TurnOffSEDetectorBeforeAcquisition"
nDefault& = 0           ' 0 = do not turn off SE detector before analysis, 1 = turn off SE detector before analysis
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
TurnOffSEDetectorBeforeAcquisitionFlag% = valid&
If TurnOffSEDetectorBeforeAcquisitionFlag% < 0 Or TurnOffSEDetectorBeforeAcquisitionFlag% > 1 Then
msg$ = "TurnOffSEDetectorBeforeAcquisition keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
TurnOffSEDetectorBeforeAcquisitionFlag% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "AutomationOverheadPerAnalysis"
If InterfaceType% = 0 Then lpDefault$ = "5.0"             ' Demo
If InterfaceType% = 1 Then lpDefault$ = "10.0"            ' Unused
If InterfaceType% = 2 Then lpDefault$ = "10.0"            ' JEOL 8900/8200/8500/8230/8530
If InterfaceType% = 3 Then lpDefault$ = "10.0"            ' Unused
If InterfaceType% = 4 Then lpDefault$ = "10.0"            ' Unused
If InterfaceType% = 5 Then lpDefault$ = "10.0"            ' SX100/SXFive
If InterfaceType% = 6 Then lpDefault$ = "10.0"            ' Axioscope
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then AutomationOverheadPerAnalysis! = Val(Left$(lpReturnString$, valid&))
If AutomationOverheadPerAnalysis! < 0.1 Or AutomationOverheadPerAnalysis! > 100# Then
msg$ = "AutomationOverheadPerAnalysis keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISoftware"
AutomationOverheadPerAnalysis! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Reflected optical light present
lpAppName$ = "Hardware"
lpKeyName$ = "ReflectedLightPresent"
nDefault& = True
If InterfaceType% = 2 And JeolEOSInterfaceType& = 1 Then nDefault& = True   ' 8200/8500
If InterfaceType% = 2 And JeolEOSInterfaceType& = 2 Then nDefault& = True   ' 8900
If InterfaceType% = 2 And JeolEOSInterfaceType& = 3 Then nDefault& = False  ' 8230/8530
If InterfaceType% = 5 Then nDefault& = True                                 ' SX100
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ReflectedLightPresent = True
Else
ReflectedLightPresent = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Transmitted optical light present
lpAppName$ = "Hardware"
lpKeyName$ = "TransmittedLightPresent"
nDefault& = False
If InterfaceType% = 2 And JeolEOSInterfaceType& = 1 Then nDefault& = True   ' 8200/8500
If InterfaceType% = 2 And JeolEOSInterfaceType& = 2 Then nDefault& = True   ' 8900
If InterfaceType% = 2 And JeolEOSInterfaceType& = 3 Then nDefault& = True  ' 8230/8530
If InterfaceType% = 5 Then nDefault& = True                                 ' SX100
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
TransmittedLightPresent = True
Else
TransmittedLightPresent = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "HysteresisPresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
HysteresisPresentFlag = True
Else
HysteresisPresentFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "SX100MoveSpectroMilliSecDelayBefore"
nDefault& = 100           ' SX100 only (seems to be required when using software backlash!)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
SX100MoveSpectroMilliSecDelayBefore& = valid&
If SX100MoveSpectroMilliSecDelayBefore& < 0 Or SX100MoveSpectroMilliSecDelayBefore& > 10000 Then
msg$ = "SX100MoveSpectroMilliSecDelayBefore keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
SX100MoveSpectroMilliSecDelayBefore& = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "SX100MoveSpectroMilliSecDelayAfter"
nDefault& = 10           ' SX100 only (seems to be required when using software backlash!)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
SX100MoveSpectroMilliSecDelayAfter& = valid&
If SX100MoveSpectroMilliSecDelayAfter& < 0 Or SX100MoveSpectroMilliSecDelayAfter& > 10000 Then
msg$ = "SX100MoveSpectroMilliSecDelayAfter keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
SX100MoveSpectroMilliSecDelayAfter& = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "SX100MoveStageMilliSecDelayBefore"
nDefault& = 100           ' SX100 only (seems to be required when using software backlash!)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
SX100MoveStageMilliSecDelayBefore& = valid&
If SX100MoveStageMilliSecDelayBefore& < 0 Or SX100MoveStageMilliSecDelayBefore& > 10000 Then
msg$ = "SX100MoveStageMilliSecDelayBefore keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
SX100MoveStageMilliSecDelayBefore& = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "SX100MoveStageMilliSecDelayAfter"
nDefault& = 10           ' SX100 only (seems to be required when using software backlash!)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
SX100MoveStageMilliSecDelayAfter& = valid&
If SX100MoveStageMilliSecDelayAfter& < 0 Or SX100MoveStageMilliSecDelayAfter& > 10000 Then
msg$ = "SX100MoveStageMilliSecDelayAfter keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
SX100MoveStageMilliSecDelayAfter& = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "SX100ScanSpectroMilliSecDelayBefore"
nDefault& = 200           ' SX100/SXFive only (seems to be required for SXFive only)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
SX100ScanSpectroMilliSecDelayBefore = valid&
If SX100ScanSpectroMilliSecDelayBefore& < 0 Or SX100ScanSpectroMilliSecDelayBefore& > 10000 Then
msg$ = "SX100ScanSpectroMilliSecDelayBefore keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
SX100ScanSpectroMilliSecDelayBefore& = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "SX100ScanSpectroMilliSecDelayAfter"
nDefault& = 200           ' SX100/SXFive only (seems to be required for SX100 and SXFive)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
SX100ScanSpectroMilliSecDelayAfter& = valid&
If SX100ScanSpectroMilliSecDelayAfter& < 0 Or SX100ScanSpectroMilliSecDelayAfter& > 10000 Then
msg$ = "SX100ScanSpectroMilliSecDelayAfter keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
SX100ScanSpectroMilliSecDelayAfter& = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "SX100FlipCrystalMilliSecDelayBefore"
nDefault& = 200           ' SX100/SXFive only (seems to be required for SXFive only)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
SX100FlipCrystalMilliSecDelayBefore = valid&
If SX100FlipCrystalMilliSecDelayBefore& < 0 Or SX100FlipCrystalMilliSecDelayBefore& > 10000 Then
msg$ = "SX100FlipCrystalMilliSecDelayBefore keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
SX100FlipCrystalMilliSecDelayBefore& = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "SX100FlipCrystalMilliSecDelayAfter"
nDefault& = 200           ' SX100/SXFive only (seems to be required for SX100 and SXFive)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
SX100FlipCrystalMilliSecDelayAfter& = valid&
If SX100FlipCrystalMilliSecDelayAfter& < 0 Or SX100FlipCrystalMilliSecDelayAfter& > 10000 Then
msg$ = "SX100FlipCrystalMilliSecDelayAfter keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
SX100FlipCrystalMilliSecDelayAfter& = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "MinMagWindow"
If InterfaceType% = 0 Then nDefault& = 40   ' demo
If InterfaceType% = 1 Then nDefault& = 40   ' unused
If InterfaceType% = 2 Then nDefault& = 12   ' JEOL  (40x at 11 mm WD)
If InterfaceType% = 3 Then nDefault& = 40   ' unused
If InterfaceType% = 4 Then nDefault& = 40   ' unused
If InterfaceType% = 5 Then nDefault& = 63   ' SX100/SXFive
If InterfaceType% = 6 Then nDefault& = 40   ' Axioscope
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
MinMagWindow! = CSng(valid&)
If MinMagWindow! <= 1 Or MinMagWindow! > 10000 Then
msg$ = "MinMagWindow keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
MinMagWindow! = CSng(nDefault&)
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "MaxMagWindow"
If InterfaceType% = 0 Then nDefault& = 12000000   ' demo
If InterfaceType% = 1 Then nDefault& = 900000     ' unused
If InterfaceType% = 2 Then nDefault& = 12000000   ' JEOL
If InterfaceType% = 3 Then nDefault& = 900000     ' unused
If InterfaceType% = 4 Then nDefault& = 900000     ' unused
If InterfaceType% = 5 Then nDefault& = 12000000   ' SX100/SXFive
If InterfaceType% = 6 Then nDefault& = 900000     ' Axioscope
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
MaxMagWindow! = CSng(valid&)
If MaxMagWindow! <= 1 Or MaxMagWindow! > 50000000 Then
msg$ = "MaxMagWindow keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
MaxMagWindow! = CSng(nDefault&)
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Image shift control
lpAppName$ = "Hardware"
lpKeyName$ = "ImageShiftPresent"
nDefault& = False
If InterfaceType = 0 Then nDefault& = True    ' demo mode
If InterfaceType = 2 Then nDefault& = True    ' JEOL 8900/8200/8500
If InterfaceType = 5 Then nDefault& = True    ' SX100/SXFive
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ImageShiftPresent% = True
Else
ImageShiftPresent% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "ImageShiftType"         ' not used at this time
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
ImageShiftType% = valid&
If ImageShiftType% < 0 Or ImageShiftType% > 1 Then
msg$ = "ImageShift keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
ImageShiftType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "EDSInsertRetractPresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
EDSInterfaceInsertRetractPresent = True
Else
EDSInterfaceInsertRetractPresent = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "EDSMaxEnergyThroughputPresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
EDSInterfaceMaxEnergyThroughputPresent = True
Else
EDSInterfaceMaxEnergyThroughputPresent = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "EDSMCSInputsPresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
EDSInterfaceMCSInputsPresent = True
Else
EDSInterfaceMCSInputsPresent = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' CL spectrometer
lpAppName$ = "Hardware"
lpKeyName$ = "CLSpectraInterfacePresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
CLSpectraInterfacePresent = True
Else
CLSpectraInterfacePresent = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "CLSpectraInterfaceType"         ' 0 = demo, 1 = Ocean Optics, 2 = Gatan, 3 = Newport, 4 = not used yet
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
CLSpectraInterfaceType% = valid&
If CLSpectraInterfaceType% < 0 Or CLSpectraInterfaceType% > 1 Then
msg$ = "CLSpectraInterfaceType keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIHardware2"
CLSpectraInterfaceType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Hardware"
lpKeyName$ = "CLInterfaceInsertRetractPresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
CLInterfaceInsertRetractPresent = True
Else
CLInterfaceInsertRetractPresent = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

Exit Sub

' Errors
InitINIHardware2Error:
MsgBox Error$, vbOKOnly + vbCritical, "InitINIHardware2"
ierror = True
Exit Sub

End Sub

Sub InitINIImage()
' Open the PROBEWIN.INI file and read defaults

ierror = False
On Error GoTo InitINIImageError

Dim i As Integer
Dim valid As Long, tValid As Long
Dim astring As String, tcomment As String

Dim lpAppName As String
Dim lpKeyName As String
Dim lpDefault As String
Dim lpFileName As String
Dim lpReturnString As String * 255
Dim lpReturnString2 As String * 255

Dim nSize As Long
Dim nDefault As Long

' Check for existing PROBEWIN.INI
If Dir$(ProbeWinINIFile$) = vbNullString Then
msg$ = "Unable to open file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
End
End If

' Use Windows API function to read PROBEWIN.INI
lpFileName$ = ProbeWinINIFile$
nSize& = Len(lpReturnString$)

' Imaging parameters
lpAppName$ = "Image"
lpKeyName$ = "ImageInterfacePresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ImageInterfacePresent% = True
Else
ImageInterfacePresent% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Check for unsupported image interfaces
If InterfaceType = 10 And ImageInterfacePresent = True Then
msg$ = "ImageInterface is not supported for this InterfaceType keyword value in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
End
End If

lpAppName$ = "Image"
lpKeyName$ = "ImageInterfaceType"
nDefault& = 0   ' 0=Demo, 1=Unused, 2=Unused, 3=Unused, 4=8900/8200/8500/8x30, 5=SX100/SXFive Mapping, 6=SX100/SXFive Video, 7=Unused, 8=Unused, 9=Bruker, 10=Thermo
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
ImageInterfaceType% = valid&
If ImageInterfaceType% < 0 Or ImageInterfaceType% > MAXINTERFACE_IMAGE% Then
msg$ = "ImageInterfaceType keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
ImageInterfaceType% = nDefault&
End If
If ImageInterfaceType% = 1 Or ImageInterfaceType% = 2 Or ImageInterfaceType% = 3 Or ImageInterfaceType% = 7 Or ImageInterfaceType% = 8 Then
msg$ = "ImageInterfaceType keyword value (" & Format$(ImageInterfaceType%) & ") is no longer supported by Probe for EPMA in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
ImageInterfaceType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageInterfaceNameChan1"
If InterfaceType% = 2 Then  ' JEOL 8900/8200
lpDefault$ = "SEI"
Else
lpDefault$ = vbNullString
End If
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then ImageInterfaceNameChan1$ = Left$(lpReturnString$, valid&)
If Len(ImageInterfaceNameChan1$) > 12 Then
msg$ = "ImageInterfaceNameChan1 string is too long in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
End
End If
lpDefault$ = Trim$(ImageInterfaceNameChan1$)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageInterfaceNameChan2"
If InterfaceType% = 2 Then  ' JEOL 8900/8200
lpDefault$ = "COMPO"
Else
lpDefault$ = vbNullString
End If
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then ImageInterfaceNameChan2$ = Left$(lpReturnString$, valid&)
If Len(ImageInterfaceNameChan2$) > 12 Then
msg$ = "ImageInterfaceNameChan2 string is too long in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
End
End If
lpDefault$ = Trim$(ImageInterfaceNameChan2$)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageInterfaceNameChan3"
If InterfaceType% = 2 Then  ' JEOL 8900/8200
If JeolEOSInterfaceType& = 1 Then lpDefault$ = "AUX"    ' 8200/8500
If JeolEOSInterfaceType& = 2 Then lpDefault$ = "AUX"    ' 8900
If JeolEOSInterfaceType& = 3 Then lpDefault$ = vbNullString       ' not available for 8230/8530
Else
lpDefault$ = vbNullString
End If
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then ImageInterfaceNameChan3$ = Left$(lpReturnString$, valid&)
If Len(ImageInterfaceNameChan3$) > 12 Then
msg$ = "ImageInterfaceNameChan3 string is too long in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
End
End If
lpDefault$ = Trim$(ImageInterfaceNameChan3$)
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageInterfacePolarityChan1"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ImageInterfacePolarityChan1% = True
Else
ImageInterfacePolarityChan1% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageInterfacePolarityChan2"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ImageInterfacePolarityChan2% = True
Else
ImageInterfacePolarityChan2% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageInterfacePolarityChan3"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ImageInterfacePolarityChan3% = True
Else
ImageInterfacePolarityChan3% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageInterfaceImageIxIy"
lpDefault$ = "1.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then ImageInterfaceImageIxIy! = Val(Left$(lpReturnString$, valid&))
If ImageInterfaceImageIxIy! < 0.5 Or ImageInterfaceImageIxIy! > 2# Then
msg$ = "ImageInterfaceImageIxIy keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
ImageInterfaceImageIxIy! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Newly modified beam calibration variables for calibration array based on keV and magnification
lpAppName$ = "Image"
lpKeyName$ = "ImageInterfaceCalNumberOfBeamCalibrations"
nDefault& = 1
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
ImageInterfaceCalNumberOfBeamCalibrations% = valid&
If ImageInterfaceCalNumberOfBeamCalibrations% < 1 Or ImageInterfaceCalNumberOfBeamCalibrations% > MAXBEAMCALIBRATIONS% Then
msg$ = "ImageInterfaceCalNumberOfBeamCalibrations keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIStageBitmaps"
ImageInterfaceCalNumberOfBeamCalibrations% = nDefault&
End
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageInterfaceCalKeV"
lpDefault$ = "15.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
astring$ = Left$(lpReturnString$, valid&)
Call InitParseStringToReal(astring$, ImageInterfaceCalNumberOfBeamCalibrations%, ImageInterfaceCalKeVArray!())
If ierror Then End
For i% = 1 To ImageInterfaceCalNumberOfBeamCalibrations%
If ImageInterfaceCalKeVArray!(i%) < MINKILOVOLTS! Or ImageInterfaceCalKeVArray!(i%) > MAXKILOVOLTS! Then
msg$ = "ImageInterfaceCalKeV keyword value (" & Format$(ImageInterfaceCalKeVArray!(i%)) & ", array index " & Format$(i%) & ") is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
End
End If
Next i%
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageInterfaceCalMag"
lpDefault$ = "400.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
astring$ = Left$(lpReturnString$, valid&)
Call InitParseStringToReal(astring$, ImageInterfaceCalNumberOfBeamCalibrations%, ImageInterfaceCalMagArray!())
If ierror Then End
For i% = 1 To ImageInterfaceCalNumberOfBeamCalibrations%
If ImageInterfaceCalMagArray!(i%) < 10# Or ImageInterfaceCalMagArray!(i%) > 100000! Then
msg$ = "ImageInterfaceCalMag keyword value (" & Format$(ImageInterfaceCalMagArray!(i%)) & ", array index " & Format$(i%) & ") is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
End
End If
Next i%
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageInterfaceCalXMicrons"
lpDefault$ = "800.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
astring$ = Left$(lpReturnString$, valid&)
Call InitParseStringToReal(astring$, ImageInterfaceCalNumberOfBeamCalibrations%, ImageInterfaceCalXMicronsArray!())
If ierror Then End
For i% = 1 To ImageInterfaceCalNumberOfBeamCalibrations%
If ImageInterfaceCalXMicronsArray!(i%) < 1# Or ImageInterfaceCalXMicronsArray!(i%) > 10000# Then
msg$ = "ImageInterfaceCalXMicrons keyword value (" & Format$(ImageInterfaceCalXMicronsArray!(i%)) & ", array index " & Format$(i%) & ") is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
End
End If
Next i%
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageInterfaceCalYMicrons"
lpDefault$ = "800.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
astring$ = Left$(lpReturnString$, valid&)
Call InitParseStringToReal(astring$, ImageInterfaceCalNumberOfBeamCalibrations%, ImageInterfaceCalYMicronsArray!())
If ierror Then End
For i% = 1 To ImageInterfaceCalNumberOfBeamCalibrations%
If ImageInterfaceCalYMicronsArray!(i%) < 1# Or ImageInterfaceCalYMicronsArray!(i%) > 10000# Then
msg$ = "ImageInterfaceCalYMicrons keyword value (" & Format$(ImageInterfaceCalYMicronsArray!(i%)) & ", array index " & Format$(i%) & ") is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
End
End If
Next i%
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageInterfaceCalScanRotation"
lpDefault$ = DefaultScanRotation!  ' use default rotation
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
astring$ = Left$(lpReturnString$, valid&)
Call InitParseStringToReal(astring$, ImageInterfaceCalNumberOfBeamCalibrations%, ImageInterfaceCalScanRotationArray!())
If ierror Then End
For i% = 1 To ImageInterfaceCalNumberOfBeamCalibrations%
If ImageInterfaceCalScanRotationArray!(i%) < 0# Or ImageInterfaceCalScanRotationArray!(i%) >= 360# Then
msg$ = "ImageInterfaceCalScanRotation keyword value (" & Format$(ImageInterfaceCalScanRotationArray!(i%)) & ", array index " & Format$(i%) & ") is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
End
End If
Next i%
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageInterfaceBeamXPolarity"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ImageInterfaceBeamXPolarity% = True
Else
ImageInterfaceBeamXPolarity% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageInterfaceBeamYPolarity"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ImageInterfaceBeamYPolarity% = True
Else
ImageInterfaceBeamYPolarity% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageInterfaceStageXPolarity"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ImageInterfaceStageXPolarity% = True
Else
ImageInterfaceStageXPolarity% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageInterfaceStageYPolarity"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ImageInterfaceStageYPolarity% = True
Else
ImageInterfaceStageYPolarity% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageInterfaceDisplayXPolarity"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ImageInterfaceDisplayXPolarity% = True
Else
ImageInterfaceDisplayXPolarity% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageInterfaceDisplayYPolarity"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
ImageInterfaceDisplayYPolarity% = True
Else
ImageInterfaceDisplayYPolarity% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImagePaletteNumber"
nDefault& = 1
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DefaultImagePaletteNumber% = valid&
If DefaultImagePaletteNumber% < 0 Or DefaultImagePaletteNumber% > 4 Then
msg$ = "ImagePaletteNumber keyword value is out of range (must be between 0 and 4) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
DefaultImagePaletteNumber% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageDisplaySizeInCentimeters"
If InterfaceType% = 5 Then
lpDefault$ = "38.0"     ' SX100/SXFive = 38 cm
Else
lpDefault$ = "12.5"     ' 4" x 5" display
End If
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then ImageDisplaySizeInCentimeters! = Val(Left$(lpReturnString$, valid&))
If ImageDisplaySizeInCentimeters! < 1# Or ImageDisplaySizeInCentimeters! > 100# Then
msg$ = "ImageDisplaySizeInCentimeters keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
ImageDisplaySizeInCentimeters! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageAutoBrightnessContrastSEGain"
nDefault& = 350
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
ImageAutoBrightnessContrastSEGain% = valid&
If ImageAutoBrightnessContrastSEGain% < 1 Or ImageAutoBrightnessContrastSEGain% > 1000 Then
msg$ = "ImageAutoBrightnessContrastSEGain keyword value is out of range (must be between 1 and 1000) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
ImageAutoBrightnessContrastSEGain% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageAutoBrightnessContrastSEOffset"
nDefault& = 480
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
ImageAutoBrightnessContrastSEOffset% = valid&
If ImageAutoBrightnessContrastSEOffset% < 1 Or ImageAutoBrightnessContrastSEOffset% > 1000 Then
msg$ = "ImageAutoBrightnessContrastSEOffset keyword value is out of range (must be between 1 and 1000) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
ImageAutoBrightnessContrastSEOffset% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageAutoBrightnessContrastBSEGain"
nDefault& = 350
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
ImageAutoBrightnessContrastBSEGain% = valid&
If ImageAutoBrightnessContrastBSEGain% < 1 Or ImageAutoBrightnessContrastBSEGain% > 1000 Then
msg$ = "ImageAutoBrightnessContrastBSEGain keyword value is out of range (must be between 1 and 1000) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
ImageAutoBrightnessContrastBSEGain% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageAutoBrightnessContrastBSEOffset"
nDefault& = 480
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
ImageAutoBrightnessContrastBSEOffset% = valid&
If ImageAutoBrightnessContrastBSEOffset% < 1 Or ImageAutoBrightnessContrastBSEOffset% > 1000 Then
msg$ = "ImageAutoBrightnessContrastBSEOffset keyword value is out of range (must be between 1 and 1000) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
ImageAutoBrightnessContrastBSEOffset% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageAlternateScaleBarUnits"
nDefault& = 0   ' no alternate units (1 = nm, 2 = um, 3 = mm, 4 = cm, 5 = meters, 6 = microinches, 7 = milliinches, 8 = inches)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
ImageAlternateScaleBarUnits% = valid&
If ImageAlternateScaleBarUnits% < 0 Or ImageAlternateScaleBarUnits% > 8 Then
msg$ = "ImageAlternateScaleBarUnits keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
ImageAlternateScaleBarUnits% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Read stage to micron conversion factors here for image shift minimum mag and mosaic limits
If Dir$(MotorsFile$) = vbNullString Then GoTo InitINIImageNotFoundMotorsFile
Close #Temp1FileNumber%
DoEvents
Open MotorsFile$ For Input As #Temp1FileNumber%
Call InitMotors2(Int(5))
Close #Temp1FileNumber%
If ierror Then Exit Sub

lpAppName$ = "Image"
lpKeyName$ = "ImageShiftMinimumMag"
lpDefault$ = "3200.0"                                                                                            ' assume JEOL (mm)
If (InterfaceType% = 0 And MiscIsInstrumentStage("JEOL")) Or InterfaceType% = 2 Then lpDefault$ = "3200.0"       ' JEOL (3200x or 100 um FOV for image shift)
If (InterfaceType% = 0 And MiscIsInstrumentStage("CAMECA")) Or InterfaceType% = 5 Then lpDefault$ = "1267.0"     ' SX100/SXFive = 1267x for first mag coil switch
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then ImageShiftMinimumMag! = Val(Left$(lpReturnString$, valid&))
If ImageShiftMinimumMag! < 1000# Or ImageShiftMinimumMag! > 5000# Then
msg$ = "ImageShiftMinimumMag keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
ImageShiftMinimumMag! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Load wide area mosaic parameters
lpAppName$ = "Image"
lpKeyName$ = "ImageMosaicSizeX"
lpDefault$ = "10.0"
If MiscIsInstrumentStage("JEOL") Then lpDefault$ = "10.0"     ' 10 mm for mosiac area
If MiscIsInstrumentStage("CAMECA") Then lpDefault$ = "10000.0"     ' 10,000 um or 10 mm for mosiac area
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then MosaicSizeX! = Val(Left$(lpReturnString$, valid&))
If MosaicSizeX! <= 0# Then
msg$ = "ImageMosaicSizeX keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
MosaicSizeX! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageMosaicSizeY"
If MiscIsInstrumentStage("JEOL") Then lpDefault$ = "10.0"     ' 10 mm for mosiac area
If MiscIsInstrumentStage("CAMECA") Then lpDefault$ = "10000.0"     ' 10,000 um or 10 mm for mosiac area
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then MosaicSizeY! = Val(Left$(lpReturnString$, valid&))
If MosaicSizeY! <= 0# Then
msg$ = "ImageMosaicSizeY keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
MosaicSizeY! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageMosaicMag"
lpDefault$ = "400.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then MosaicMagnification! = Val(Left$(lpReturnString$, valid&))
If MosaicMagnification! <= 10# Or MosaicMagnification! > 100000# Then
msg$ = "ImageMosaicMag keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
MosaicMagnification! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

Exit Sub

' Errors
InitINIImageError:
MsgBox Error$, vbOKOnly + vbCritical, "InitINIImage"
ierror = True
Exit Sub

InitINIImageNotFoundMotorsFile:
msg$ = "File " & MotorsFile$ & " was not found in the application data folder."
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage"
ierror = True
Exit Sub

End Sub

Sub InitINIImage2()
' Open the PROBEWIN.INI file and read defaults for parameters that must be read *after* all config files have loaded

ierror = False
On Error GoTo InitINIImage2Error

Dim ip As Integer
Dim valid As Long, tValid As Long
Dim tcomment As String

Dim lpAppName As String
Dim lpKeyName As String
Dim lpDefault As String
Dim lpFileName As String
Dim lpReturnString As String * 255
Dim lpReturnString2 As String * 255

Dim nSize As Long

' Check for existing PROBEWIN.INI
If Dir$(ProbeWinINIFile$) = vbNullString Then
msg$ = "Unable to open file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage2"
End
End If

' Use Windows API function to read PROBEWIN.INI
lpFileName$ = ProbeWinINIFile$
nSize& = Len(lpReturnString$)

' Imaging parameters that must be read *after* all config files have loaded
lpAppName$ = "Image"
lpKeyName$ = "ImageRGB1_R"
lpDefault$ = "Fe"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then ImageRGB1_R$ = Left$(lpReturnString$, valid&)
ip% = IPOS1%(MAXELM%, ImageRGB1_R$, Symlo$())
If ip% = 0 Then
msg$ = "ImageImageRGB1_R is not a valid element symbol in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage2"
End
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageRGB1_G"
lpDefault$ = "Mg"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then ImageRGB1_G$ = Left$(lpReturnString$, valid&)
ip% = IPOS1%(MAXELM%, ImageRGB1_G$, Symlo$())
If ip% = 0 Then
msg$ = "ImageImageRGB1_G is not a valid element symbol in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage2"
End
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Image"
lpKeyName$ = "ImageRGB1_B"
lpDefault$ = "Ca"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then ImageRGB1_B$ = Left$(lpReturnString$, valid&)
ip% = IPOS1%(MAXELM%, ImageRGB1_B$, Symlo$())
If ip% = 0 Then
msg$ = "ImageImageRGB1_B is not a valid element symbol in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIImage2"
End
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

Exit Sub

' Errors
InitINIImage2Error:
MsgBox Error$, vbOKOnly + vbCritical, "InitINIImage2"
ierror = True
Exit Sub

End Sub

Sub InitINICounting()
' Open the PROBEWIN.INI file and read defaults

ierror = False
On Error GoTo InitINICountingError

Dim tValid As Long, valid As Long, nSize As Long
Dim tcomment As String

Dim lpAppName As String
Dim lpKeyName As String
Dim lpDefault As String
Dim lpFileName As String
Dim lpReturnString As String * 255
Dim lpReturnString2 As String * 255

' Check for existing PROBEWIN.INI
If Dir$(ProbeWinINIFile$) = vbNullString Then
msg$ = "Unable to open file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINICounting"
End
End If

' Use Windows API function to read PROBEWIN.INI
lpFileName$ = ProbeWinINIFile$
nSize& = Len(lpReturnString$)

' Counting section, first get default count times
lpAppName$ = "Counting"
lpKeyName$ = "OnPeakCountTime"
lpDefault$ = "10.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultOnCountTime! = Val(Left$(lpReturnString$, valid&))
If DefaultOnCountTime! < 0.01 Or DefaultOnCountTime! > 1000# Then
msg$ = "OnPeakCountTime keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINICounting"
DefaultOnCountTime! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Counting"
lpKeyName$ = "OffPeakCountTime"
lpDefault$ = "5.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultOffCountTime! = Val(Left$(lpReturnString$, valid&))
If DefaultOffCountTime! < 0.01 Or DefaultOffCountTime! > 1000# Then
msg$ = "OffPeakCountTime keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINICounting"
DefaultOffCountTime! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Peaking count time
lpAppName$ = "Counting"
lpKeyName$ = "PeakingCountTime"
lpDefault$ = "8.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultPeakingCountTime! = Val(Left$(lpReturnString$, valid&))
If DefaultPeakingCountTime! < 0.01 Or DefaultPeakingCountTime! > 1000# Then
msg$ = "PeakingCountTime keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINICounting"
DefaultPeakingCountTime! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Wavescan count time
lpAppName$ = "Counting"
lpKeyName$ = "WavescanCountTime"
lpDefault$ = "6.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultWavescanCountTime! = Val(Left$(lpReturnString$, valid&))
If DefaultWavescanCountTime! < 0.01 Or DefaultWavescanCountTime! > 1000# Then
msg$ = "WavescanCountTime keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINICounting"
DefaultWavescanCountTime! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Quickscan count time
lpAppName$ = "Counting"
lpKeyName$ = "QuickscanCountTime"
lpDefault$ = "0.5"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultQuickscanCountTime! = Val(Left$(lpReturnString$, valid&))
If DefaultQuickscanCountTime! < 0.01 Or DefaultQuickscanCountTime! > 1000# Then
msg$ = "QuickscanCountTime keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINICounting"
DefaultQuickscanCountTime! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Specify default unknown max counts
lpAppName$ = "Counting"
lpKeyName$ = "UnknownMaxCounts"
lpDefault$ = Str$(MAXCOUNT&)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultUnknownMaxCounts& = Val(Left$(lpReturnString$, valid&))
If DefaultUnknownMaxCounts& < 1000 Or DefaultUnknownMaxCounts& > MAXCOUNT& Then
msg$ = "UnknownMaxCounts keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINICounting"
DefaultUnknownMaxCounts& = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

Exit Sub

' Errors
InitINICountingError:
MsgBox Error$, vbOKOnly + vbCritical, "InitINICounting"
ierror = True
Exit Sub

End Sub

Sub InitINIFaraday()
' Open the PROBEWIN.INI file and read defaults

ierror = False
On Error GoTo InitINIFaradayError

Dim valid As Long, tValid As Long

Dim lpAppName As String
Dim lpKeyName As String
Dim lpDefault As String
Dim lpFileName As String
Dim lpReturnString As String * 255
Dim lpReturnString2 As String * 255

Dim nSize As Long
Dim nDefault As Long
Dim tcomment As String

' Check for existing PROBEWIN.INI
If Dir$(ProbeWinINIFile$) = vbNullString Then
msg$ = "Unable to open file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIFaraday"
End
End If

' Use Windows API function to read PROBEWIN.INI
lpFileName$ = ProbeWinINIFile$
nSize& = Len(lpReturnString$)

lpAppName$ = "Faraday"
lpKeyName$ = "FaradayCupType"
lpDefault$ = "0"   ' automatic
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then FaradayCupType% = Val(Left$(lpReturnString$, valid&))
If FaradayCupType% < 0 Or FaradayCupType% > 1 Then
msg$ = "FaradayCupType keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIFaraday"
FaradayCupType% = nDefault&
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Faraday"
lpKeyName$ = "FaradayAverages"
nDefault& = 1   ' default is one current measurement
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DefaultBeamAverages% = valid&
If DefaultBeamAverages% < 1 Or DefaultBeamAverages% > 100# Then
msg$ = "FaradayAverages keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIFaraday"
DefaultBeamAverages% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Faraday"
lpKeyName$ = "FaradayWaitInTime"
lpDefault$ = "0.5"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then FaradayWaitInTime! = Val(Left$(lpReturnString$, valid&))
If FaradayWaitInTime! < 0# Or FaradayWaitInTime! > 100# Then
msg$ = "FaradayWaitInTime keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIFaraday"
FaradayWaitInTime! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Faraday"
lpKeyName$ = "FaradayWaitOutTime"
lpDefault$ = "0.5"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then FaradayWaitOutTime! = Val(Left$(lpReturnString$, valid&))
If FaradayWaitOutTime! < 0# Or FaradayWaitOutTime! > 100# Then
msg$ = "FaradayWaitOutTime keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIFaraday"
FaradayWaitOutTime! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Faraday"
lpKeyName$ = "DefaultBlankBeamFlag"
nDefault& = True
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
DefaultBlankBeamFlag = True
Else
DefaultBlankBeamFlag = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Faraday"
lpKeyName$ = "AbsorbedCurrentPresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
AbsorbedCurrentPresent = True
Else
AbsorbedCurrentPresent = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Faraday"
lpKeyName$ = "AbsorbedCurrentType"
nDefault& = 0   ' automatic
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
AbsorbedCurrentType% = valid&
If AbsorbedCurrentType% < 0 Or AbsorbedCurrentType% > 1 Then
msg$ = "AbsorbedCurrentType keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIFaraday"
AbsorbedCurrentType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Faraday"
lpKeyName$ = "MinimumFaradayCurrent"
lpDefault$ = "0.1"  ' 0.1 nA default
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then MinimumFaradayCurrent! = Val(Left$(lpReturnString$, valid&))
If MinimumFaradayCurrent! < 0.001 Or MinimumFaradayCurrent! > 1000# Then
msg$ = "MinimumFaradayCurrent keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIFaraday"
MinimumFaradayCurrent! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Faraday"
lpKeyName$ = "FaradayBeamCurrentSafeThreshold"
lpDefault$ = "500.0"  ' in nA
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then FaradayBeamCurrentSafeThreshold! = Val(Left$(lpReturnString$, valid&))
If FaradayBeamCurrentSafeThreshold! < 1# Or FaradayBeamCurrentSafeThreshold! > 10000# Then
msg$ = "FaradayBeamCurrentSafeThreshold keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIFaraday"
FaradayBeamCurrentSafeThreshold! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Specimen (stage) mounted faraday cup flag
lpAppName$ = "Faraday"
lpKeyName$ = "FaradayStagePresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
FaradayStagePresent = True
Else
FaradayStagePresent = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Specimen (stage) mounted faraday cup positions, X, Y, Z, W, plus T (tilt) and R (rotation)
lpAppName$ = "Faraday"
lpKeyName$ = "FaradayStagePositionsX"
lpDefault$ = "0.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then FaradayStagePositions!(1) = Val(Left$(lpReturnString$, valid&))
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Faraday"
lpKeyName$ = "FaradayStagePositionsY"
lpDefault$ = "0.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then FaradayStagePositions!(2) = Val(Left$(lpReturnString$, valid&))
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Faraday"
lpKeyName$ = "FaradayStagePositionsZ"
lpDefault$ = "0.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then FaradayStagePositions!(3) = Val(Left$(lpReturnString$, valid&))
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Faraday"
lpKeyName$ = "FaradayStagePositionsW"
lpDefault$ = "0.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then FaradayStagePositions!(4) = Val(Left$(lpReturnString$, valid&))
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Faraday"
lpKeyName$ = "FaradayStagePositionsT"
lpDefault$ = "0.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then FaradayStagePositions!(5) = Val(Left$(lpReturnString$, valid&))
If FaradayStagePositions!(5) < 0# Or FaradayStagePositions!(5) > 90# Then
msg$ = "FaradayStagePositionT keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIFaraday"
End
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Faraday"
lpKeyName$ = "FaradayStagePositionsR"
lpDefault$ = "0.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then FaradayStagePositions!(6) = Val(Left$(lpReturnString$, valid&))
If FaradayStagePositions!(6) < 0# Or FaradayStagePositions!(6) > 360# Then
msg$ = "FaradayStagePositionsR keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIFaraday"
End
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

Exit Sub

' Errors
InitINIFaradayError:
MsgBox Error$, vbOKOnly + vbCritical, "InitINIFaraday"
ierror = True
Exit Sub

End Sub

Sub InitINIPHA()
' Open the PROBEWIN.INI file and read defaults

ierror = False
On Error GoTo InitINIPHAError

Dim valid As Long, nSize As Long, nDefault As Long, tValid As Long
Dim tcomment As String

Dim lpAppName As String
Dim lpKeyName As String
Dim lpDefault As String
Dim lpFileName As String
Dim lpReturnString As String * 255
Dim lpReturnString2 As String * 255

' Check for existing PROBEWIN.INI
If Dir$(ProbeWinINIFile$) = vbNullString Then
msg$ = "Unable to open file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIPHA"
End
End If

' Use Windows API function to read PROBEWIN.INI
lpFileName$ = ProbeWinINIFile$
nSize& = Len(lpReturnString$)

' PHA section
lpAppName$ = "PHA"
lpKeyName$ = "PHAHardwarePresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
PHAHardware = True
Else
PHAHardware = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "PHA"
lpKeyName$ = "PHAHardwareType"
nDefault& = 0       ' 0 = traditional PHA acquisition, 1 = MCA PHA acquisition
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
PHAHardwareType% = valid&
If PHAHardwareType% < 0 Or PHAHardwareType% > 1 Then
msg$ = "PHAHardwareType keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIPHA"
PHAHardwareType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "PHA"
lpKeyName$ = "PHAGainBiasPresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
PHAGainBias% = True
Else
PHAGainBias% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "PHA"
lpKeyName$ = "PHAGainBiasType"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
PHAGainBiasType% = valid&
If PHAGainBiasType% < 0 Or PHAGainBiasType% > 1 Then
msg$ = "PHAGainBiasType keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIPHA"
PHAGainBiasType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Integral/differential mode parameters
lpAppName$ = "PHA"
lpKeyName$ = "PHAInteDiffPresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
PHAInteDiff% = True
Else
PHAInteDiff% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "PHA"
lpKeyName$ = "PHAInteDiffType"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
PHAInteDiffType% = valid&
If PHAInteDiffType% < 0 Or PHAInteDiffType% > 1 Then
msg$ = "PHAInteDiffType keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIPHA"
PHAInteDiffType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Deadtime parameters
lpAppName$ = "PHA"
lpKeyName$ = "PHADeadtimePresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
PHADeadTime% = True
Else
PHADeadTime% = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "PHA"
lpKeyName$ = "PHADeadtimeType"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
PHADeadTimeType% = valid&
If PHADeadTimeType% < 0 Or PHAInteDiffType% > 1 Then
msg$ = "PHADeadtimeType keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIPHA"
PHAInteDiffType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' PHA distribution count time and intervals defaults
lpAppName$ = "PHA"
lpKeyName$ = "PHACountTime"
lpDefault$ = "0.5"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultPHACountTime! = Val(Left$(lpReturnString$, valid&))
If DefaultPHACountTime! < 0.1 Or DefaultPHACountTime! > 1000# Then
msg$ = "PHACountTime keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIPHA"
DefaultPHACountTime! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "PHA"
lpKeyName$ = "PHAIntervals"
nDefault& = 20
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DefaultPHAIntervals% = valid&
If DefaultPHAIntervals% < 5 Or DefaultPHAIntervals% > 500 Then
msg$ = "PHAIntervals keyword value is out of range (must be between 5 and 500) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIPHA"
DefaultPHAIntervals% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "PHA"
lpKeyName$ = "PHAAdjustPresent"
nDefault& = False
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
If valid& <> 0 Then
PHAAdjustPresent = True
Else
PHAAdjustPresent = False
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Bias scan distribution count time and intervals defaults
lpAppName$ = "PHA"
lpKeyName$ = "BiasScanCountTime"
lpDefault$ = "0.5"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultBiasScanCountTime! = Val(Left$(lpReturnString$, valid&))
If DefaultBiasScanCountTime! < 0.1 Or DefaultBiasScanCountTime! > 1000# Then
msg$ = "BiasScanCountTime keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIPHA"
DefaultBiasScanCountTime! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "PHA"
lpKeyName$ = "BiasScanIntervals"
nDefault& = 40
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DefaultBiasScanIntervals% = valid&
If DefaultBiasScanIntervals% < 5 Or DefaultBiasScanIntervals% > 1000 Then
msg$ = "BiasScanIntervals keyword value is out of range (must be between 5 and 1000) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIPHA"
DefaultBiasScanIntervals% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Bias scan distribution count time and intervals defaults
lpAppName$ = "PHA"
lpKeyName$ = "GainScanCountTime"
lpDefault$ = "0.5"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultGainScanCountTime! = Val(Left$(lpReturnString$, valid&))
If DefaultGainScanCountTime! < 0.1 Or DefaultGainScanCountTime! > 1000# Then
msg$ = "GainScanCountTime keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIPHA"
DefaultGainScanCountTime! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "PHA"
lpKeyName$ = "GainScanIntervals"
nDefault& = 30
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DefaultGainScanIntervals% = valid&
If DefaultGainScanIntervals% < 5 Or DefaultGainScanIntervals% > 1000 Then
msg$ = "GainScanIntervals keyword value is out of range (must be between 5 and 1000) in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIPHA"
DefaultGainScanIntervals% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "PHA"
lpKeyName$ = "PHAFirstTimeDelay"
lpDefault$ = "0.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then PHAFirstTimeDelay! = Val(Left$(lpReturnString$, valid&))
If PHAFirstTimeDelay! < 0# Or PHAFirstTimeDelay! > 100# Then
msg$ = "PHAFirstTimeDelay keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIPHA"
PHAFirstTimeDelay! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "PHA"
lpKeyName$ = "PHAMultiChannelMin"
lpDefault$ = "0.805"    ' 805 mV
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then PHAMultiChannelMin! = Val(Left$(lpReturnString$, valid&))
If PHAMultiChannelMin! < 0# Or PHAMultiChannelMin! > 6# Then
msg$ = "PHAMultiChannelMin keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIPHA"
PHAMultiChannelMin! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "PHA"
lpKeyName$ = "PHAMultiChannelMax"
lpDefault$ = "5.637"     ' 5637 mV
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then PHAMultiChannelMax! = Val(Left$(lpReturnString$, valid&))
If PHAMultiChannelMax! < 0# Or PHAMultiChannelMax! > 6# Then
msg$ = "PHAMultiChannelMax keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIPHA"
PHAMultiChannelMax! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

Exit Sub

' Errors
InitINIPHAError:
MsgBox Error$, vbOKOnly + vbCritical, "InitINIPHA"
ierror = True
Exit Sub

End Sub

Sub InitINIPlot()
' Open the PROBEWIN.INI file and read defaults

ierror = False
On Error GoTo InitINIPlotError

Dim valid As Long, tValid As Long
Dim tcomment As String

Dim astring As String
Dim lpAppName As String
Dim lpKeyName As String
Dim lpDefault As String
Dim lpFileName As String
Dim lpReturnString As String * 255
Dim lpReturnString2 As String * 255

Dim nSize As Long
Dim nDefault As Long

' Check for existing PROBEWIN.INI
If Dir$(ProbeWinINIFile$) = vbNullString Then
msg$ = "Unable to open file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIPlot"
End
End If

' Use Windows API function to read PROBEWIN.INI
lpFileName$ = ProbeWinINIFile$
nSize& = Len(lpReturnString$)

' Plot section
lpAppName$ = "Plot"
lpKeyName$ = "MinimumKLMDisplay"
lpDefault$ = Str$(0.5)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultMinimumKLMDisplay! = Val(Left$(lpReturnString$, valid&))
If DefaultMinimumKLMDisplay! < 0.01 Or DefaultMinimumKLMDisplay! > 10# Then
msg$ = "MinimumKLMDisplay keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIPlot"
DefaultMinimumKLMDisplay! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Plot"
lpKeyName$ = "GraphType"
nDefault& = 1
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DefaultGraphType% = valid&
If DefaultGraphType% < 0 Or DefaultGraphType% > 3 Then
msg$ = "GraphType keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIPlot"
DefaultGraphType% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Plot"
lpKeyName$ = "GraphTypeWave"
nDefault& = 1
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DefaultGraphTypeWav% = valid&
If DefaultGraphTypeWav% < 0 Or DefaultGraphTypeWav% > 2 Then
msg$ = "GraphTypeWave keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIPlot"
DefaultGraphTypeWav% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Load forbidden elements
lpAppName$ = "Plot"
lpKeyName$ = "NumberofForbiddenElements"
nDefault& = 10
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DefaultNumberofForbiddenElements% = valid&
If DefaultNumberofForbiddenElements% < 0 Or DefaultNumberofForbiddenElements% > MAXFORBIDDEN% Then
msg$ = "NumberofForbiddenElements keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIPlot"
DefaultNumberofForbiddenElements% = nDefault&
End If
NumberofForbiddenElements% = DefaultNumberofForbiddenElements%      ' load temp variable
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

If NumberofForbiddenElements% > 0 Then
lpAppName$ = "Plot"
lpKeyName$ = "ForbiddenElements"
lpDefault$ = "1,2,43,61,84,85,86,87,88,89"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
astring$ = Left$(lpReturnString$, valid&)
Call InitParseStringToInteger(astring$, NumberofForbiddenElements%, ForbiddenElements%())
If ierror Then End
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

Exit Sub

' Errors
InitINIPlotError:
MsgBox Error$, vbOKOnly + vbCritical, "InitINIPlot"
ierror = True
Exit Sub

End Sub

Sub InitINIStandards()
' Open the PROBEWIN.INI file and read defaults

ierror = False
On Error GoTo InitINIStandardsError

Dim valid As Long, tValid As Long
Dim tcomment As String

Dim lpAppName As String
Dim lpKeyName As String
Dim lpDefault As String
Dim lpFileName As String
Dim lpReturnString As String * 255
Dim lpReturnString2 As String * 255

Dim nSize As Long
Dim nDefault As Long

' Check for existing PROBEWIN.INI
If Dir$(ProbeWinINIFile$) = vbNullString Then
msg$ = "Unable to open file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIStandards"
End
End If

' Use Windows API function to read PROBEWIN.INI
lpFileName$ = ProbeWinINIFile$
nSize& = Len(lpReturnString$)

' Standards section
lpAppName$ = "Standards"
lpKeyName$ = "IncrementXForAdditionalPoints"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
IncrementXForAdditionalPoints% = valid&
If Abs(IncrementXForAdditionalPoints%) < 0 Or Abs(IncrementXForAdditionalPoints%) > 100 Then
msg$ = "IncrementXForAdditionalPoints keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIStandards"
IncrementXForAdditionalPoints% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Standards"
lpKeyName$ = "IncrementYForReStandardizations"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
IncrementYForReStandardizations% = valid&
If Abs(IncrementYForReStandardizations%) < 0 Or Abs(IncrementYForReStandardizations%) > 100 Then
msg$ = "IncrementYForReStandardizations keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIStandards"
IncrementYForReStandardizations% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Standards"
lpKeyName$ = "StandardPointsToAcquire"
nDefault& = 5
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
StandardPointsToAcquire% = valid&
If StandardPointsToAcquire% < 0 Or StandardPointsToAcquire% > 50 Then
msg$ = "StandardPointsToAcquire keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIStandards"
StandardPointsToAcquire% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Stage bitmap files
Call InitINIStageBitmaps(Int(0), ProbeWinINIFile$)
If ierror Then Exit Sub

' Stage standard mount folder location
lpAppName$ = "Standards"
lpKeyName$ = "StandardPOSFileDirectory"
lpDefault$ = UserDataDirectory$ & "\StandardPOSData"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
StandardPOSFileDirectory$ = lpDefault$      ' set to default in case keyword in INI file is a null string
If Left$(lpReturnString$, valid&) <> vbNullString Then StandardPOSFileDirectory$ = Left$(lpReturnString$, valid&)
If Right$(StandardPOSFileDirectory$, 1) = "\" Then StandardPOSFileDirectory$ = Left$(StandardPOSFileDirectory$, Len(StandardPOSFileDirectory$) - 1) ' remove trailing backslash
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Standards"
lpKeyName$ = "MatchStandardDatabase"
lpDefault$ = "DHZ.MDB"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultMatchStandardDatabase$ = Left$(lpReturnString$, valid&)
If Dir$(ApplicationCommonAppData$ & DefaultMatchStandardDatabase) = vbNullString Then
msg$ = "MatchStandardDatabase keyword value " & DefaultMatchStandardDatabase$ & " in " & ProbeWinINIFile$ & " was not found in the Probe for EPMA folder"
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIStandards"
End
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Standards"
lpKeyName$ = "StandardCoatingFlag"
nDefault& = 1    ' 0 = not coated, 1 = coated
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DefaultStandardCoatingFlag% = valid&
If DefaultStandardCoatingFlag% < 0 Or DefaultStandardCoatingFlag% > 1 Then
msg$ = "StandardCoatingFlag keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIStandards"
DefaultStandardCoatingFlag% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Standards"
lpKeyName$ = "StandardCoatingElement"
nDefault& = 6    ' assume carbon
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
DefaultStandardCoatingElement% = valid&
If DefaultStandardCoatingElement% < 1 Or DefaultStandardCoatingElement% > MAXELM% Then
msg$ = "StandardCoatingElement keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIStandards"
DefaultStandardCoatingElement% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Standards"
lpKeyName$ = "StandardCoatingDensity"
lpDefault$ = Format$(2.1)
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultStandardCoatingDensity! = Val(Left$(lpReturnString$, valid&))
If DefaultStandardCoatingDensity! < 0.3 Or DefaultStandardCoatingDensity! > 30# Then
msg$ = "StandardCoatingDensity keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIStandards"
DefaultStandardCoatingDensity! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Standards"
lpKeyName$ = "StandardCoatingThickness"
lpDefault$ = Format$(200)      ' in angstroms
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then DefaultStandardCoatingThickness! = Val(Left$(lpReturnString$, valid&))
If DefaultStandardCoatingThickness! < 1# Or DefaultStandardCoatingThickness! > 10000# Then
msg$ = "StandardCoatingThickness keyword value (in angstroms) is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIStandards"
DefaultStandardCoatingThickness! = Val(lpDefault$)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

Exit Sub

' Errors
InitINIStandardsError:
MsgBox Error$, vbOKOnly + vbCritical, "InitINIStandards"
ierror = True
Exit Sub

End Sub

Sub InitINISerial()
' Open the PROBEWIN.INI file and read defaults

ierror = False
On Error GoTo InitINISerialError

Dim valid As Long, tValid As Long
Dim tcomment As String

Dim lpAppName As String
Dim lpKeyName As String
Dim lpDefault As String
Dim lpFileName As String
Dim lpReturnString As String * 255
Dim lpReturnString2 As String * 255

Dim nSize As Long
Dim nDefault As Long

' Check for existing PROBEWIN.INI
If Dir$(ProbeWinINIFile$) = vbNullString Then
msg$ = "Unable to open file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISerial"
End
End If

' Use Windows API function to read PROBEWIN.INI
lpFileName$ = ProbeWinINIFile$
nSize& = Len(lpReturnString$)

' Serial port parameters
lpAppName$ = "Serial"
lpKeyName$ = "Port"
nDefault& = 1
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
SerialPort% = valid&
If SerialPort% < 1 Or SerialPort% > 4 Then
msg$ = "Serial Port keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISerial"
SerialPort% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Serial"
lpKeyName$ = "HandShaking"
nDefault& = 1
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
SerialHandShaking% = valid&
If SerialHandShaking% < 0 Or SerialHandShaking% > 3 Then
msg$ = "Serial HandShaking keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISerial"
SerialHandShaking% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Serial"
lpKeyName$ = "Baud"
nDefault& = 9600
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
SerialBaud% = valid&
If SerialBaud% < 110 Or SerialBaud% > 19200 Then
msg$ = "Serial Baud keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISerial"
SerialBaud% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Serial"
lpKeyName$ = "Parity"
lpDefault$ = "N"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then SerialParity$ = Left$(lpReturnString$, valid&)
If SerialParity$ <> "E" And SerialParity$ <> "M" And SerialParity$ <> "N" And SerialParity$ <> "O" And SerialParity$ <> "S" Then
msg$ = "Serial Parity keyword value out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISerial"
SerialParity$ = Left$(lpReturnString$, valid&)
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Serial"
lpKeyName$ = "DataBits"
nDefault& = 8
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
SerialDataBits% = valid&
If SerialDataBits% < 4 Or SerialDataBits% > 8 Then
msg$ = "Serial DataBits keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISerial"
SerialDataBits% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

lpAppName$ = "Serial"
lpKeyName$ = "StopBits"
nDefault& = 1
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
SerialStopBits% = valid&
If SerialStopBits% < 1 Or SerialStopBits% > 2 Then
msg$ = "Serial StopBits keyword value is out of range in " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINISerial"
SerialStopBits% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

Exit Sub

' Errors
InitINISerialError:
MsgBox Error$, vbOKOnly + vbCritical, "InitINISerial"
ierror = True
Exit Sub

End Sub

Sub InitINIStageBitmaps(mode As Integer, tfilename As String)
' Read the stage bit maps values only from the passed (usually PROBEWIN.INI) file
'  0 = normal INI read
'  1 = read for export (ignore missing files, just warn user)

ierror = False
On Error GoTo InitINIStageBitmapsError

Dim i As Integer
Dim valid As Long, tValid As Long

Dim lpAppName As String
Dim lpKeyName As String
Dim lpDefault As String
Dim lpFileName As String
Dim lpReturnString As String * 255
Dim lpReturnString2 As String * 255

Dim nSize As Long
Dim nDefault As Long
Dim astring As String, tcomment As String

' Check for existing PROBEWIN.INI
If Dir$(tfilename$) = vbNullString Then
msg$ = "Unable to open file " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIStageBitmaps"
ierror = True
Exit Sub
End If

' Use Windows API function to read PROBEWIN.INI
lpFileName$ = tfilename$
nSize& = Len(lpReturnString$)

lpAppName$ = "Standards"
lpKeyName$ = "StageBitMapCount"
nDefault& = 0
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)
valid& = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault&, lpFileName$)
StageBitMapCount% = valid&
If StageBitMapCount% < 0 Or StageBitMapCount% > MAXBITMAP% Then
msg$ = "StageBitMapCount keyword value is out of range in " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIStageBitmaps"
StageBitMapCount% = nDefault&
End If
If Left$(lpReturnString$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, Format$(nDefault&), lpFileName$)

' Only load if at least one stage bit map is indicated
If StageBitMapCount% > 0 Then
lpAppName$ = "Standards"
lpKeyName$ = "StageBitMapFile"
lpDefault$ = vbNullString
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
If Left$(lpReturnString$, valid&) <> vbNullString Then astring$ = Left$(lpReturnString$, valid&)
If Trim$(astring$) = vbNullString Then
msg$ = "StageBitMapFile string is empty in " & tfilename$ & vbCrLf & vbCrLf
msg$ = msg$ & "Program will now end."
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIStageBitmaps"
End
End If
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

' Parse into array
Call InitParseStringToString(astring$, StageBitMapCount%, StageBitMapFile$())
If ierror Then End

' Move stage bitmaps specified from probewin.ini file
Call InitFilesMove(Int(1))
If ierror Then Exit Sub

' Check that files are only WMF (WindowsMetaFiles) and exist
For i% = 1 To StageBitMapCount%

' Check that file is .WMF (WMF files will re-size automatically when loaded to the picture property of a form, bitmap files will not automatically re-size to the form)
If InStr(1, StageBitMapFile$(i%), ".WMF", 1) = 0 Then
msg$ = "StageBitMapFile " & StageBitMapFile$(i%) & " in " & tfilename$ & " is not a Windows MetaFile (.WMF)." & vbCrLf & vbCrLf
msg$ = msg$ & "Program will now end."
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIStageBitmaps"
End
End If

' Check that no path is specified
If InStr(1, StageBitMapFile$(i%), "\", 1) <> 0 Then
msg$ = "StageBitMapFile string " & StageBitMapFile$(i%) & " in " & tfilename$ & ", specifies a full path. Please move the file to the " & ApplicationCommonAppData$ & " directory and edit the INI entry so that only the actual file name is specified." & vbCrLf & vbCrLf
msg$ = msg$ & "Program will now end."
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIStageBitmaps"
End
End If

' Check that file exists (skip if exporting INI file)
If mode% = 0 Then
If Dir$(ApplicationCommonAppData$ & StageBitMapFile$(i%)) = vbNullString Then
msg$ = "StageBitMapFile " & StageBitMapFile$(i%) & " in " & tfilename$ & " was not found" & vbCrLf & vbCrLf
'msg$ = msg$ & "Program will now end."
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIStageBitmaps"
'End
End If
End If
Next i%

lpAppName$ = "Standards"
lpKeyName$ = "StageBitMapXmin"
lpDefault$ = "0.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
astring$ = Left$(lpReturnString$, valid&)
Call InitParseStringToReal(astring$, StageBitMapCount%, StageBitMapXmin!())
If ierror Then End
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Standards"
lpKeyName$ = "StageBitMapXmax"
lpDefault$ = "0.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
astring$ = Left$(lpReturnString$, valid&)
Call InitParseStringToReal(astring$, StageBitMapCount%, StageBitMapXmax!())
If ierror Then End
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Standards"
lpKeyName$ = "StageBitMapYmin"
lpDefault$ = "0.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
astring$ = Left$(lpReturnString$, valid&)
Call InitParseStringToReal(astring$, StageBitMapCount%, StageBitMapYmin!())
If ierror Then End
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)

lpAppName$ = "Standards"
lpKeyName$ = "StageBitMapYmax"
lpDefault$ = "0.0"
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString2$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
astring$ = Left$(lpReturnString$, valid&)
Call InitParseStringToReal(astring$, StageBitMapCount%, StageBitMapYmax!())
If ierror Then End
If Left$(lpReturnString2$, tValid&) = vbNullString Then valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)
End If

' Check valid X and Y min and max
For i% = 1 To StageBitMapCount%
If StageBitMapXmin!(i%) = StageBitMapXmax!(i%) Or StageBitMapYmin!(i%) = StageBitMapYmax!(i%) Then
msg$ = "StageBitMapFile " & Str$(i%) & " X or Y min and max values are equal for " & StageBitMapFile$(i%) & " in " & tfilename$ & vbCrLf & vbCrLf
msg$ = msg$ & "Program will now end."
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIStageBitmaps"
End
End If
Next i%

Exit Sub

' Errors
InitINIStageBitmapsError:
MsgBox Error$, vbOKOnly + vbCritical, "InitINIStageBitmaps"
ierror = True
Exit Sub

End Sub
