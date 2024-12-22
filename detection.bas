Attribute VB_Name = "CodeDETECTION"
' (c) Copyright 1995-2025 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Dim DetectionUnknownBackgroundIntensity As Single
Dim DetectionUnknownBeamCurrent As Single

Dim DetectionStandardIntensity As Single
Dim DetectionStandardWeightPercent As Single

Dim DetectionUnknownOnPeakTime As Single
Dim DetectionUnknownWeightPercent As Single

Sub DetectionLoad()
' Load the form with module level variables

ierror = False
On Error GoTo DetectionLoadError

' Load variables
FormDETECTION.TextUnknownBackgroundIntensity.Text = Str$(DetectionUnknownBackgroundIntensity!)
FormDETECTION.TextUnknownBeamCurrent.Text = Str$(DetectionUnknownBeamCurrent!)

FormDETECTION.TextStandardIntensity.Text = Str$(DetectionStandardIntensity!)
FormDETECTION.TextStandardWeightPercent.Text = Str$(DetectionStandardWeightPercent!)

FormDETECTION.TextUnknownOnPeakTime.Text = Str$(DetectionUnknownOnPeakTime!)
FormDETECTION.TextUnknownWeightPercent.Text = Str$(DetectionUnknownWeightPercent!)

Exit Sub

' Errors
DetectionLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "DetectionLoad"
ierror = True
Exit Sub

End Sub

Sub DetectionSave()
' Save the form module level variables

ierror = False
On Error GoTo DetectionSaveError

' Load variables
DetectionUnknownBackgroundIntensity! = Val(FormDETECTION.TextUnknownBackgroundIntensity.Text)
DetectionUnknownBeamCurrent! = Val(FormDETECTION.TextUnknownBeamCurrent.Text)

DetectionStandardIntensity! = Val(FormDETECTION.TextStandardIntensity.Text)
DetectionStandardWeightPercent! = Val(FormDETECTION.TextStandardWeightPercent.Text)

DetectionUnknownOnPeakTime! = Val(FormDETECTION.TextUnknownOnPeakTime.Text)
DetectionUnknownWeightPercent! = Val(FormDETECTION.TextUnknownWeightPercent.Text)

Exit Sub

' Errors
DetectionSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "DetectionSave"
ierror = True
Exit Sub

End Sub

Sub DetectionCalculateConcentration()
' Calculate detection limit with a specified count time

ierror = False
On Error GoTo DetectionCalculateConcentrationError

Dim ucts As Single
Dim temp1 As Single, temp2 As Single

' Check basic parameters
Call DetectionCheckParameters
If ierror Then Exit Sub

' Check input
If DetectionUnknownOnPeakTime! <= 0# Then GoTo DetectionCalculateConcentrationBadUnknownOnPeakTime
FormDETECTION.LabelWeightPercentDetected.Caption = vbNullString

' Calculate total unknown and standard counts
ucts! = DetectionUnknownBackgroundIntensity! * DetectionUnknownOnPeakTime! * DetectionUnknownBeamCurrent!

temp1! = 3# * Sqr(ucts!) * DetectionStandardWeightPercent!
temp2! = DetectionUnknownOnPeakTime! * DetectionUnknownBeamCurrent! * DetectionStandardIntensity!

If temp2! <= 0# Then Exit Sub
DetectionUnknownWeightPercent! = temp1! / temp2!

FormDETECTION.LabelWeightPercentDetected.Caption = MiscAutoFormat$(DetectionUnknownWeightPercent!)
Exit Sub

' Errors
DetectionCalculateConcentrationError:
MsgBox Error$, vbOKOnly + vbCritical, "DetectionCalculateConcentration"
ierror = True
Exit Sub

DetectionCalculateConcentrationBadUnknownOnPeakTime:
msg$ = "Invalid Unknown On Peak Time"
MsgBox msg$, vbOKOnly + vbExclamation, "DetectionCalculateConcentration"
ierror = True
Exit Sub

End Sub

Sub DetectionCalculateCountTime()
' Calculate count time with a specified concentration

ierror = False
On Error GoTo DetectionCalculateCountTimeError

Dim temp1 As Single, temp2 As Single

' Check basic parameters
Call DetectionCheckParameters
If ierror Then Exit Sub

' Check input
If DetectionUnknownWeightPercent! <= 0# Then GoTo DetectionCalculateCountTimeBadUnknownWeightPercent
FormDETECTION.LabelTimePredicted.Caption = vbNullString

' Calculate total unknown and standard counts

temp1! = 9# * DetectionUnknownBackgroundIntensity! * DetectionStandardWeightPercent! ^ 2
temp2! = DetectionUnknownWeightPercent! ^ 2 * DetectionStandardIntensity! ^ 2 * DetectionUnknownBeamCurrent!

If temp2! <= 0# Then Exit Sub
DetectionUnknownOnPeakTime! = temp1! / temp2!

FormDETECTION.LabelTimePredicted.Caption = MiscAutoFormat$(DetectionUnknownOnPeakTime!)
Exit Sub

' Errors
DetectionCalculateCountTimeError:
MsgBox Error$, vbOKOnly + vbCritical, "DetectionCalculateCountTime"
ierror = True
Exit Sub

DetectionCalculateCountTimeBadUnknownWeightPercent:
msg$ = "Invalid Unknown Weight Percent"
MsgBox msg$, vbOKOnly + vbExclamation, "DetectionCalculateCountTime"
ierror = True
Exit Sub
End Sub

Sub DetectionCheckParameters()
' Check common parameters

ierror = False
On Error GoTo DetectionCheckParametersError

' Check basic parameters
If DetectionUnknownBackgroundIntensity! <= 0# Then GoTo DetectionCheckParametersBadUnknownBackgroundIntensity
If DetectionUnknownBeamCurrent! <= 0# Then GoTo DetectionCheckParametersBadUnknownBeamCurrent

If DetectionStandardIntensity! <= 0# Then GoTo DetectionCheckParametersBadStandardIntensity
If DetectionStandardWeightPercent! <= 0# Or DetectionStandardWeightPercent! > 100# Then GoTo DetectionCheckParametersBadStandardWeightPercent

Exit Sub

' Errors
DetectionCheckParametersError:
MsgBox Error$, vbOKOnly + vbCritical, "DetectionCheckParameters"
ierror = True
Exit Sub

DetectionCheckParametersBadUnknownBackgroundIntensity:
msg$ = "Invalid Unknown Background Intensity"
MsgBox msg$, vbOKOnly + vbExclamation, "DetectionCheckParameters"
ierror = True
Exit Sub

DetectionCheckParametersBadUnknownBeamCurrent:
msg$ = "Invalid Unknown Beam Current"
MsgBox msg$, vbOKOnly + vbExclamation, "DetectionCheckParameters"
ierror = True
Exit Sub

DetectionCheckParametersBadStandardIntensity:
msg$ = "Invalid Standard Intensity"
MsgBox msg$, vbOKOnly + vbExclamation, "DetectionCheckParameters"
ierror = True
Exit Sub

DetectionCheckParametersBadStandardWeightPercent:
msg$ = "Invalid Standard Weight Percent"
MsgBox msg$, vbOKOnly + vbExclamation, "DetectionCheckParameters"
ierror = True
Exit Sub

End Sub

Sub DetectionPrint(mode As Integer)
' Print the results
' mode = 1 print detection limit
' mode = 2 print predicted count time

ierror = False
On Error GoTo DetectionPrintError

' Print variables
Call IOWriteLog(vbCrLf)
If mode% = 1 Then Call IOWriteLog("Detection Limit Results Calculated at: " & Now)
If mode% = 2 Then Call IOWriteLog("Predicted Unknown Count Time Calculated at: " & Now)
Call IOWriteLog("Unknown Background Intensity (cps/nA) = " & Str$(DetectionUnknownBackgroundIntensity!))
Call IOWriteLog("Unknown Beam Current (nA) = " & Str$(DetectionUnknownBeamCurrent!))
If mode% = 1 Then Call IOWriteLog("Unknown On Peak Time (sec.) = " & Str$(DetectionUnknownOnPeakTime!))
If mode% = 2 Then Call IOWriteLog("Unknown Weight Percent (wt. %) = " & Str$(DetectionUnknownWeightPercent!))

Call IOWriteLog("Standard Intensity (cps/nA) = " & Str$(DetectionStandardIntensity!))
Call IOWriteLog("Standard Weight Percent (wt. %) = " & Str$(DetectionStandardWeightPercent!))

' Results
If mode% = 1 Then Call IOWriteLog("Unknown Weight Percent (wt. %) = " & Str$(DetectionUnknownWeightPercent!))
If mode% = 2 Then Call IOWriteLog("Unknown On Peak Time (sec.) = " & Str$(DetectionUnknownOnPeakTime!))

Exit Sub

' Errors
DetectionPrintError:
MsgBox Error$, vbOKOnly + vbCritical, "DetectionPrint"
ierror = True
Exit Sub

End Sub

