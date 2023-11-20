Attribute VB_Name = "CodeEffective"
' (c) Copyright 1995-2023 by John J. Donovan
Option Explicit

Dim PrimaryStandardNum As Integer, SecondaryStandardNum As Integer

Dim ElementSym As String, XraySym As String

Dim TakeoffLow As Single, TakeoffHigh As Single, TakeoffIncrement As Single
Dim BeamEnergy As Single

Dim CalcZAFAnalysis As TypeAnalysis
Dim CalcZAFOldSample(1 To 1) As TypeSample
Dim CalcZAFTmpSample(1 To 1) As TypeSample

Sub EffectiveTakeoffAngleKRatiosLoad()
' Load the form with module level variables

ierror = False
On Error GoTo EffectiveTakeoffAngleKRatiosLoadError

' Get available standard names and numbers from database
Call StandardGetMDBIndex
If ierror Then Exit Sub

' Load standard one dropdown list
Call StandardLoadList(FormEffective.ListStandardPrimary)
If ierror Then Exit Sub

' Select first standard in list
If FormEffective.ListStandardPrimary.ListCount > 0 Then FormEffective.ListStandardPrimary.ListIndex = 0

' Load standard two dropdown list
Call StandardLoadList(FormEffective.ListStandardSecondary)
If ierror Then Exit Sub

' Select second standard in list
If FormEffective.ListStandardSecondary.ListCount > 1 Then FormEffective.ListStandardSecondary.ListIndex = 1

' Load takeoff angle parameters
FormEffective.TextTakeoffLow.Text = Format$(35)
FormEffective.TextTakeoffHigh.Text = Format$(45)
FormEffective.TextTakeoffIncrement.Text = Format$(0.5)
FormEffective.TextBeamEnergy.Text = Format$(DefaultKiloVolts!)

Exit Sub

' Errors
EffectiveTakeoffAngleKRatiosLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "EffectiveTakeoffAngleKRatiosLoad"
ierror = True
Exit Sub

End Sub

Sub EffectiveTakeoffAngleKRatiosSave()
' Save the form module level variables

ierror = False
On Error GoTo EffectiveTakeoffAngleKRatiosSaveError

Dim sym As String
Dim i As Integer
Dim ip  As Integer, ipp As Integer
Dim keV As Single, lam As Single

' Save to module level
If FormEffective.ListStandardPrimary.ListIndex < 0 Then Exit Sub
If FormEffective.ListStandardPrimary.ListCount < 1 Then Exit Sub
PrimaryStandardNum% = FormEffective.ListStandardPrimary.ItemData(FormEffective.ListStandardPrimary.ListIndex)

If FormEffective.ListStandardSecondary.ListIndex < 0 Then Exit Sub
If FormEffective.ListStandardSecondary.ListCount < 1 Then Exit Sub
SecondaryStandardNum% = FormEffective.ListStandardSecondary.ItemData(FormEffective.ListStandardSecondary.ListIndex)

' Save element parameters
sym$ = FormEffective.ComboElement.Text
ip% = IPOS1(MAXELM%, sym$, Symlo$())
If ip% = 0 Then GoTo EffectiveTakeoffAngleKRatiosSaveBadElement
ElementSym$ = sym$
    
' Get the xray symbol
sym$ = FormEffective.ComboXRay.Text
ipp% = IPOS1(MAXRAY% - 1, sym$, Xraylo$()) ' must be an analyzed element
If ipp% = 0 Then GoTo EffectiveTakeoffAngleKRatiosSaveBadXray
XraySym$ = sym$

' Save to default x-ray list
Deflin$(ipp%) = CalcZAFOldSample(1).Xrsyms$(1)

' Check for a valid xray line
Call XrayGetKevLambda(Symlo$(ip%), Xraylo$(ipp%), keV!, lam!)
If ierror Then Exit Sub

' Save take off range and beam energy
If Val(FormEffective.TextTakeoffLow) < 10# Or Val(FormEffective.TextTakeoffLow) > 90# Then
msg$ = FormEffective.TextTakeoffLow.Text & " low take off angle is out of range! (must be between " & Format$(10#) & " and " & Format$(90#) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "EffectiveTakeoffAngleKRatiosSave"
ierror = True
Exit Sub
Else
TakeoffLow! = Val(FormEffective.TextTakeoffLow)
End If

If Val(FormEffective.TextTakeoffHigh) < 10# Or Val(FormEffective.TextTakeoffHigh) > 90# Then
msg$ = FormEffective.TextTakeoffHigh.Text & " high take off angle is out of range! (must be between " & Format$(10#) & " and " & Format$(90#) & " degrees)"
MsgBox msg$, vbOKOnly + vbExclamation, "EffectiveTakeoffAngleKRatiosSave"
ierror = True
Exit Sub
Else
TakeoffHigh! = Val(FormEffective.TextTakeoffHigh)
End If

' Check low less than high
If TakeoffLow! > TakeoffHigh Then
msg$ = "Takeoff angle low (" & FormEffective.TextTakeoffLow.Text & ") is greater than takeoff angle high (" & FormEffective.TextTakeoffHigh.Text & " degrees)"
MsgBox msg$, vbOKOnly + vbExclamation, "EffectiveTakeoffAngleKRatiosSave"
ierror = True
Exit Sub
End If

If Val(FormEffective.TextTakeoffIncrement) < 0.05 Or Val(FormEffective.TextTakeoffIncrement) > 1# Then
msg$ = FormEffective.TextTakeoffIncrement.Text & " take off angle increment is out of range! (must be between " & Format$(0.05) & " and " & Format$(1#) & " degrees)"
MsgBox msg$, vbOKOnly + vbExclamation, "EffectiveTakeoffAngleKRatiosSave"
ierror = True
Exit Sub
Else
TakeoffIncrement! = Val(FormEffective.TextTakeoffIncrement)
End If

If Val(FormEffective.TextBeamEnergy) < 0.1 Or Val(FormEffective.TextBeamEnergy) > 50# Then
msg$ = FormEffective.TextBeamEnergy.Text & " beam energy is out of range! (must be between " & Format$(0.1) & " and " & Format$(50#) & " keV)"
MsgBox msg$, vbOKOnly + vbExclamation, "EffectiveTakeoffAngleKRatiosSave"
ierror = True
Exit Sub
Else
BeamEnergy! = Val(FormEffective.TextBeamEnergy)
End If

' Make sure specified element is in both standards and greater than zero percent
Call StandardGetMDBStandard(PrimaryStandardNum%, CalcZAFOldSample())
If ierror Then Exit Sub

ip% = IPOS1(CalcZAFOldSample(1).LastChan%, ElementSym$, CalcZAFOldSample(1).Elsyms$())
If ip% = 0 Then GoTo EffectiveTakeoffAngleKRatiosSavePrimaryNotFound
If CalcZAFOldSample(1).ElmPercents!(ip%) <= 0# Then GoTo EffectiveTakeoffAngleKRatiosSavePrimaryZero

Call StandardGetMDBStandard(SecondaryStandardNum%, CalcZAFOldSample())
If ierror Then Exit Sub

ipp% = IPOS1(CalcZAFOldSample(1).LastChan%, ElementSym$, CalcZAFOldSample(1).Elsyms$())
If ipp% = 0 Then GoTo EffectiveTakeoffAngleKRatiosSaveSecondaryNotFound
If CalcZAFOldSample(1).ElmPercents!(ipp%) <= 0# Then GoTo EffectiveTakeoffAngleKRatiosSaveSecondaryZero

Exit Sub

' Errors
EffectiveTakeoffAngleKRatiosSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "EffectiveTakeoffAngleKRatiosSave"
ierror = True
Exit Sub

EffectiveTakeoffAngleKRatiosSaveBadElement:
msg$ = "Element " & sym$ & " is an invalid element symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "EffectiveTakeoffAngleKRatiosSave"
ierror = True
Exit Sub

EffectiveTakeoffAngleKRatiosSaveBadXray:
msg$ = "Xray " & sym$ & " is an invalid xray symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "EffectiveTakeoffAngleKRatiosSave"
ierror = True
Exit Sub

EffectiveTakeoffAngleKRatiosSavePrimaryNotFound:
msg$ = "Element " & ElementSym$ & " was not found in standard " & Format$(PrimaryStandardNum%) & " " & CalcZAFOldSample(1).Name$
MsgBox msg$, vbOKOnly + vbExclamation, "EffectiveTakeoffAngleKRatiosSave"
ierror = True
Exit Sub

EffectiveTakeoffAngleKRatiosSaveSecondaryNotFound:
msg$ = "Element " & ElementSym$ & " was not found in standard " & Format$(SecondaryStandardNum%) & " " & CalcZAFOldSample(1).Name$
MsgBox msg$, vbOKOnly + vbExclamation, "EffectiveTakeoffAngleKRatiosSave"
ierror = True
Exit Sub

EffectiveTakeoffAngleKRatiosSavePrimaryZero:
msg$ = "Element " & ElementSym$ & " has a zero concentration in standard " & Format$(PrimaryStandardNum%) & " " & CalcZAFOldSample(1).Name$
MsgBox msg$, vbOKOnly + vbExclamation, "EffectiveTakeoffAngleKRatiosSave"
ierror = True
Exit Sub

EffectiveTakeoffAngleKRatiosSaveSecondaryZero:
msg$ = "Element " & ElementSym$ & " has a zero concentration in standard " & Format$(SecondaryStandardNum%) & " " & CalcZAFOldSample(1).Name$
MsgBox msg$, vbOKOnly + vbExclamation, "EffectiveTakeoffAngleKRatiosSave"
ierror = True
Exit Sub

End Sub

Sub EffectiveTakeoffAngleLoadElements()
' Load the element list based on the primary standard composition

ierror = False
On Error GoTo EffectiveTakeoffAngleLoadElementsError

Dim number As Integer, i As Integer

' Get the selected standard
If FormEffective.ListStandardPrimary.ListIndex < 0 Then Exit Sub
If FormEffective.ListStandardPrimary.ListCount < 1 Then Exit Sub

number% = FormEffective.ListStandardPrimary.ItemData(FormEffective.ListStandardPrimary.ListIndex)

' Get this composition
Call StandardGetMDBStandard(number%, CalcZAFOldSample())
If ierror Then Exit Sub

' Load element list
FormEffective.ComboElement.Clear
For i% = 1 To CalcZAFOldSample(1).LastChan
FormEffective.ComboElement.AddItem CalcZAFOldSample(1).Elsyms$(i%)
Next i%

' Load form with elements from primary standard (on listbox click events)
If FormEffective.ComboElement.ListCount > 0 Then FormEffective.ComboElement.ListIndex = 0            ' select first element

' Load x-ray on element combo event

Exit Sub

' Errors
EffectiveTakeoffAngleLoadElementsError:
MsgBox Error$, vbOKOnly + vbCritical, "EffectiveTakeoffAngleLoadElements"
ierror = True
Exit Sub

End Sub


Sub EffectiveTakeoffAngleElementUpdate()
' Updates the xray combo if the element changes

ierror = False
On Error GoTo EffectiveTakeoffAngleElementUpdateError

Dim ip As Integer
Dim sym As String

sym$ = FormEffective.ComboElement.Text
ip% = IPOS1(MAXELM%, sym$, Symlo$())

' Update xray if element matches
If ip% > 0 Then
FormEffective.ComboXRay.Text = Deflin$(ip%)
End If

Exit Sub

' Errors
EffectiveTakeoffAngleElementUpdateError:
MsgBox Error$, vbOKOnly + vbCritical, "EffectiveTakeoffAngleElementUpdate"
ierror = True
Exit Sub

End Sub

Sub EffectiveTakeoffAngleKRatiosCalculate()
' Calculate k-ratios for a range of effective takeoff angles

ierror = False
On Error GoTo EffectiveTakeoffAngleKRatiosCalculateError

Dim i As Integer, ntaks As Integer
Dim tTakeoff As Single
Dim temp As Single, temp1 As Single, temp2 As Single

Dim EffectiveTakeOffStandardName1 As String
Dim EffectiveTakeOffStandardName2 As String
Dim EffectiveTakeoffAngles() As Single
Dim EffectiveTakeoffKratios1() As Single
Dim EffectiveTakeoffKratios2() As Single

' Check if particle corrections are selected with alpha factors or calibration curve (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag <> 0 Then
msg$ = "Only ZAF or Phi-rho-z corrections are supported for effective takeoff angle k-ratio calculations."
MsgBox msg$, vbOKOnly + vbExclamation, "EffectiveTakeoffAngleKRatiosCalculate"
ierror = True
Exit Sub
End If

' Init ZAF arrays
Call ZAFInitZAF
If ierror Then Exit Sub

' Initialize arrays
Call InitStandards(CalcZAFAnalysis)
If ierror Then Exit Sub

Call InitLine(CalcZAFAnalysis)
If ierror Then Exit Sub

' CALCZAF Calculate intensity from weight
CalcZAFMode% = 0

' Force standard sample type
CalcZAFOldSample(1).Type% = 1
CalcZAFOldSample(1).Datarows% = 1   ' always a single data point
CalcZAFOldSample(1).GoodDataRows% = 1
CalcZAFOldSample(1).LineStatus(1) = True      ' force status flag always true (good data point)

' Loop on take off angles
icancelauto = False
ntaks% = 1
tTakeoff! = TakeoffLow!
Do Until tTakeoff! > TakeoffHigh!

' Dimension arrays
ReDim Preserve EffectiveTakeoffKratios1(1 To ntaks%) As Single
ReDim Preserve EffectiveTakeoffKratios2(1 To ntaks%) As Single
ReDim Preserve EffectiveTakeoffAngles(1 To ntaks%) As Single

' Load takeoff angle for output
EffectiveTakeoffAngles!(ntaks%) = tTakeoff!

' First calculate primary standard
Call StandardGetMDBStandard(PrimaryStandardNum%, CalcZAFOldSample())
If ierror Then Exit Sub

' Make all elements unanalyzed except the k-ratio element
For i% = 1 To CalcZAFOldSample(1).LastChan%
If UCase$(ElementSym$) <> UCase$(CalcZAFOldSample(1).Elsyms$(i%)) Then
CalcZAFOldSample(1).Xrsyms$(i%) = vbNullString
Else
CalcZAFOldSample(1).Xrsyms$(i%) = XraySym$
End If
Next i%

' Re-sort standard composition for one analyzed element
Call CalcZAFSave2(CalcZAFOldSample())
If ierror Then Exit Sub

CalcZAFOldSample(1).takeoff! = tTakeoff!            ' load next takeoff angle
CalcZAFOldSample(1).TakeoffArray!(1) = tTakeoff!            ' load next takeoff angle
CalcZAFOldSample(1).kilovolts! = BeamEnergy!         ' load beam energy
CalcZAFOldSample(1).KilovoltsArray!(1) = BeamEnergy!         ' load beam energy

CalcZAFOldSample(1).number% = PrimaryStandardNum%
CalcZAFOldSample(1).StdAssigns%(1) = PrimaryStandardNum%
NumberofStandards% = 1
StandardNumbers%(1) = PrimaryStandardNum%

' Call k-ratio calculation code
Call EffectiveTakeoffAngleKRatiosCalculate2(CalcZAFOldSample(), CalcZAFTmpSample(), CalcZAFAnalysis)
If ierror Then Exit Sub

' Save k-ratios for primary standard
EffectiveTakeOffStandardName1$ = CalcZAFOldSample(1).Name$
EffectiveTakeoffKratios1!(ntaks%) = CalcZAFAnalysis.StdAssignsKfactors!(1)

' Next calculate secondary standard
Call StandardGetMDBStandard(SecondaryStandardNum%, CalcZAFOldSample())
If ierror Then Exit Sub

' Make all elements unanalyzed except the k-ratio element
For i% = 1 To CalcZAFOldSample(1).LastChan%
If UCase$(ElementSym$) <> UCase$(CalcZAFOldSample(1).Elsyms$(i%)) Then
CalcZAFOldSample(1).Xrsyms$(i%) = vbNullString
Else
CalcZAFOldSample(1).Xrsyms$(i%) = XraySym$
End If
Next i%

' Re-sort standard composition for one analyzed element
Call CalcZAFSave2(CalcZAFOldSample())
If ierror Then Exit Sub

CalcZAFOldSample(1).takeoff! = tTakeoff!            ' load next takeoff angle
CalcZAFOldSample(1).TakeoffArray!(1) = tTakeoff!            ' load next takeoff angle
CalcZAFOldSample(1).kilovolts! = BeamEnergy!         ' load beam energy
CalcZAFOldSample(1).KilovoltsArray!(1) = BeamEnergy!         ' load beam energy

CalcZAFOldSample(1).number% = SecondaryStandardNum%
CalcZAFOldSample(1).StdAssigns%(1) = SecondaryStandardNum%
NumberofStandards% = 1
StandardNumbers%(1) = SecondaryStandardNum%

' Call k-ratio calculation code
Call EffectiveTakeoffAngleKRatiosCalculate2(CalcZAFOldSample(), CalcZAFTmpSample(), CalcZAFAnalysis)
If ierror Then Exit Sub

' Save k-ratios for secondary standard
EffectiveTakeOffStandardName2$ = CalcZAFOldSample(1).Name$
EffectiveTakeoffKratios2!(ntaks%) = CalcZAFAnalysis.StdAssignsKfactors!(1)

ntaks% = ntaks% + 1
tTakeoff! = tTakeoff! + TakeoffIncrement!
DoEvents
If icancelauto Then Exit Sub
Loop

' Output results
msg$ = vbCrLf & vbCrLf & "Effective K-Ratios for Primary Standard: " & Format$(PrimaryStandardNum%) & " " & EffectiveTakeOffStandardName1$
Call IOWriteLog(msg$)
msg$ = "Secondary Standard: " & Format$(SecondaryStandardNum%) & " " & EffectiveTakeOffStandardName2$
Call IOWriteLog(msg$)
msg$ = "Emission line: " & ElementSym$ & " " & XraySym$ & " at " & Format$(BeamEnergy!) & " keV"
Call IOWriteLog(msg$)
msg$ = "Absorption Correction Method: " & absstring$(iabs%)
Call IOWriteLog(msg$)
msg$ = "MAC File: " & macstring$(MACTypeFlag%) & vbCrLf
Call IOWriteLog(msg$)

' Calculate percent change from 39 to 41 degrees
For i% = 1 To ntaks% - 1
If EffectiveTakeoffAngles!(i%) = 39# Then
temp1! = EffectiveTakeoffKratios2!(i%) / EffectiveTakeoffKratios1!(i%)
End If
If EffectiveTakeoffAngles!(i%) = 41# Then
temp2! = EffectiveTakeoffKratios2!(i%) / EffectiveTakeoffKratios1!(i%)
End If
Next i%

' Calculate change per degree
If temp1! <> 0# And temp2! <> 0# Then
temp! = (temp2! - temp1!) / 2#
msg$ = "Absolute k-ratio change per degree at 40 degrees: " & MiscAutoFormat$(temp!)
Call IOWriteLog(msg$)
temp! = (temp2! - temp1!) / (2# * temp2!) * 100#
msg$ = "Percent (relative) k-ratio change per degree at 40 degrees: " & MiscAutoFormat$(temp!)
Call IOWriteLog(msg$)
End If

Call IOWriteLog$(vbNullString)

For i% = 1 To ntaks% - 1
msg$ = "Takeoff Angle: " & MiscAutoFormat$(EffectiveTakeoffAngles!(i%))
msg$ = msg$ & ", K-Ratio: " & MiscAutoFormat$(EffectiveTakeoffKratios2!(i%) / EffectiveTakeoffKratios1!(i%))
Call IOWriteLog(msg$)
Next i%

Exit Sub

' Errors
EffectiveTakeoffAngleKRatiosCalculateError:
MsgBox Error$, vbOKOnly + vbCritical, "EffectiveTakeoffAngleKRatiosCalculate"
ierror = True
Exit Sub

End Sub

Sub EffectiveTakeoffAngleKRatiosCalculate2(sample() As TypeSample, tmpsample() As TypeSample, analysis As TypeAnalysis)
' Perform actual k-ratio calculation

ierror = False
On Error GoTo EffectiveTakeoffAngleKRatiosCalculate2Error

Dim i As Integer

' Reload the element arrays
Call ElementGetData(sample())
If ierror Then Exit Sub

' Initialize calculations (needed for ZAFPTC and coating calculations) (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% <> MAXCORRECTION% Then
Call ZAFSetZAF(sample())
If ierror Then Exit Sub
Else
'Call ZAFSetZAF3(sample())
'If ierror Then Exit Sub
End If

' Set TmpSample equal to OldSample so k factors and ZAF corrections get loaded in ZAFStd
tmpsample(1) = sample(1)

' Fake sample coating for ZAFStd calculation
If UseConductiveCoatingCorrectionForElectronAbsorption Then                   ' fake standard coating
StandardCoatingFlag%(1) = sample(1).CoatingFlag%
StandardCoatingDensity!(1) = sample(1).CoatingDensity!
StandardCoatingThickness!(1) = sample(1).CoatingThickness!
StandardCoatingElement%(1) = sample(1).CoatingElement%
End If

' Run the intensity from concentration calculations on the "standard"
If CorrectionFlag% = 0 Then
Call ZAFStd2(Int(1), analysis, sample(), tmpsample())
If ierror Then Exit Sub
ElseIf CorrectionFlag% = MAXCORRECTION% Then
'Call ZAFStd3(Int(1), analysis, sample(), tmpsample())
'If ierror Then Exit Sub

' Calculate the standard beta factors for this standard
Else
AllAFactorUpdateNeeded = True   ' indicate alpha-factor update
Call AFactorStd(Int(1), analysis, sample(), tmpsample())
If ierror Then Exit Sub

Call AFactorTypeStandard(analysis, sample())
If ierror Then Exit Sub
End If

Exit Sub

' Errors
EffectiveTakeoffAngleKRatiosCalculate2Error:
MsgBox Error$, vbOKOnly + vbCritical, "EffectiveTakeoffAngleKRatiosCalculate2"
ierror = True
Exit Sub

End Sub
