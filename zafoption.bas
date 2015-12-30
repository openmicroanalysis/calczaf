Attribute VB_Name = "CodeZAFOPTION"
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

Dim ZAFOptionSample(1 To 1) As TypeSample
Dim FormulaTmpSample(1 To 1) As TypeSample

Sub ZAFOptionLoad(sample() As TypeSample)
' Loads the CalcZAF calculation options form

ierror = False
On Error GoTo ZAFOptionLoadError

Dim i As Integer, ip As Integer, ipp As Integer

' Load the sample
ZAFOptionSample(1) = sample(1)

' Load form with calculation options
If ZAFOptionSample(1).OxideOrElemental% = 1 Then
FormZAFOPT.OptionOxide.Value = True
Else
FormZAFOPT.OptionElemental.Value = True
End If

' Save DisplayAsOxideFlag option
If ZAFOptionSample(1).DisplayAsOxideFlag = True Then
FormZAFOPT.CheckDisplayAsOxide.Value = vbChecked
Else
FormZAFOPT.CheckDisplayAsOxide.Value = vbUnchecked
End If

If ZAFOptionSample(1).AtomicPercentFlag = True Then
FormZAFOPT.CheckAtomicPercents.Value = vbChecked
Else
FormZAFOPT.CheckAtomicPercents.Value = vbUnchecked
End If

' Clear and load combo boxes
FormZAFOPT.ComboDifference.Clear
FormZAFOPT.ComboStoichiometry.Clear
FormZAFOPT.ComboRelative.Clear
FormZAFOPT.ComboRelativeTo.Clear
FormZAFOPT.ComboFormula.Clear

' Elements by difference, stoiciometry or relative to must be specified elements
FormZAFOPT.ComboFormula.AddItem "Sum"  ' zero index indicates sum all cations
For i% = 1 To ZAFOptionSample(1).LastChan%
If i% > ZAFOptionSample(1).LastElm% Then
FormZAFOPT.ComboDifference.AddItem ZAFOptionSample(1).Elsyms$(i%)
FormZAFOPT.ComboStoichiometry.AddItem ZAFOptionSample(1).Elsyms$(i%)
FormZAFOPT.ComboRelative.AddItem ZAFOptionSample(1).Elsyms$(i%)
End If
FormZAFOPT.ComboRelativeTo.AddItem ZAFOptionSample(1).Elsyms$(i%)
FormZAFOPT.ComboFormula.AddItem ZAFOptionSample(1).Elsyms$(i%)
Next i%

' Load formula by difference
FormZAFOPT.CheckDifference.Value = vbUnchecked
If ZAFOptionSample(1).DifferenceElement$ <> vbNullString Then
ip% = IPOS1B(ZAFOptionSample(1).LastElm + 1, ZAFOptionSample(1).LastChan%, ZAFOptionSample(1).DifferenceElement$, ZAFOptionSample(1).Elsyms$())
ip% = ip% - ZAFOptionSample(1).LastElm%
If ip% > 0 Then
FormZAFOPT.CheckDifference.Value = vbChecked
FormZAFOPT.ComboDifference.ListIndex = ip% - 1
End If
End If

' Load defaults based on sample setup
If ZAFOptionSample(1).DifferenceFormulaFlag Then
FormZAFOPT.CheckDifferenceFormula.Value = vbChecked
Else
FormZAFOPT.CheckDifferenceFormula.Value = vbUnchecked
End If
FormZAFOPT.TextDifferenceFormula.Text = ZAFOptionSample(1).DifferenceFormula$

FormZAFOPT.CheckStoichiometry.Value = vbUnchecked
If ZAFOptionSample(1).StoichiometryElement$ <> vbNullString Then
ip% = IPOS1B(ZAFOptionSample(1).LastElm + 1, ZAFOptionSample(1).LastChan%, ZAFOptionSample(1).StoichiometryElement$, ZAFOptionSample(1).Elsyms$())
ip% = ip% - ZAFOptionSample(1).LastElm%
If ip% > 0 Then
FormZAFOPT.CheckStoichiometry.Value = vbChecked
FormZAFOPT.ComboStoichiometry.ListIndex = ip% - 1
FormZAFOPT.TextStoichiometry.Text = Str$(ZAFOptionSample(1).StoichiometryRatio!)
End If
End If

FormZAFOPT.CheckRelative.Value = vbUnchecked
If ZAFOptionSample(1).RelativeElement$ <> vbNullString And ZAFOptionSample(1).RelativeToElement$ <> vbNullString Then
ip% = IPOS1B(ZAFOptionSample(1).LastElm + 1, ZAFOptionSample(1).LastChan%, ZAFOptionSample(1).RelativeElement$, ZAFOptionSample(1).Elsyms$())
ipp% = IPOS1(ZAFOptionSample(1).LastChan, ZAFOptionSample(1).RelativeToElement$, ZAFOptionSample(1).Elsyms$())
ip% = ip% - ZAFOptionSample(1).LastElm%     ' the RelativeElement must be specified but the RelativeToElement can be an analyzed or specified element
If ip% > 0 And ipp% > 0 Then
FormZAFOPT.CheckRelative.Value = vbChecked
FormZAFOPT.ComboRelative.ListIndex = ip% - 1
FormZAFOPT.ComboRelativeTo.ListIndex = ipp% - 1
FormZAFOPT.TextRelative.Text = Str$(ZAFOptionSample(1).RelativeRatio!)
End If
End If

If CalculateElectronandXrayRangesFlag Then
FormZAFOPT.CheckCalculateElectronandXrayRanges = vbChecked
Else
FormZAFOPT.CheckCalculateElectronandXrayRanges = vbUnchecked
End If

If UseOxygenFromHalogensCorrectionFlag Then
FormZAFOPT.CheckUseOxygenFromHalogensCorrection.Value = vbChecked
Else
FormZAFOPT.CheckUseOxygenFromHalogensCorrection.Value = vbUnchecked
End If

' Load formula calculations
If ZAFOptionSample(1).FormulaElementFlag% Then
FormZAFOPT.CheckFormula.Value = vbChecked
Else
FormZAFOPT.CheckFormula.Value = vbUnchecked
End If

FormZAFOPT.ComboFormula.ListIndex = 0  ' default to sum of cations
If ZAFOptionSample(1).FormulaRatio! > 0# Then FormZAFOPT.TextFormula.Text = Str$(ZAFOptionSample(1).FormulaRatio!)
If ZAFOptionSample(1).FormulaElement$ <> vbNullString Then
ip% = IPOS1(ZAFOptionSample(1).LastChan%, ZAFOptionSample(1).FormulaElement$, ZAFOptionSample(1).Elsyms$())
If ip% > 0 Then
FormZAFOPT.ComboFormula.ListIndex = ip%
End If
End If

' Load unknown coating controls
FormZAFOPT.ComboCoatingElement.Clear
For i% = 0 To MAXELM% - 1
FormZAFOPT.ComboCoatingElement.AddItem Symlo$(i% + 1)
Next i%

'If ZAFOptionSample(1).CoatingFlag% = 0 Then ZAFOptionSample(1).CoatingFlag% = DefaultStandardCoatingFlag%
If ZAFOptionSample(1).CoatingElement% = 0 Then ZAFOptionSample(1).CoatingElement% = DefaultStandardCoatingElement%
If ZAFOptionSample(1).CoatingDensity! = 0# Then ZAFOptionSample(1).CoatingDensity! = DefaultStandardCoatingDensity!
If ZAFOptionSample(1).CoatingThickness! = 0# Then ZAFOptionSample(1).CoatingThickness! = DefaultStandardCoatingThickness!

If ZAFOptionSample(1).CoatingFlag% = 1 Then
FormZAFOPT.CheckCoatingFlag.Value = vbChecked
Else
FormZAFOPT.CheckCoatingFlag.Value = vbUnchecked
End If
FormZAFOPT.ComboCoatingElement.Text = Symlo$(ZAFOptionSample(1).CoatingElement%)
FormZAFOPT.TextCoatingDensity.Text = Format$(ZAFOptionSample(1).CoatingDensity!)
FormZAFOPT.TextCoatingThickness.Text = Format$(ZAFOptionSample(1).CoatingThickness!)

' Sample density
If ZAFOptionSample(1).SampleDensity! = 0# Then ZAFOptionSample(1).SampleDensity! = 5#
FormZAFOPT.TextDensity.Text = Format$(ZAFOptionSample(1).SampleDensity!)

' Load EDS options
FormZAFOPT.OptionEDSUse(0).Enabled = False
FormZAFOPT.OptionEDSUse(2).Enabled = False
FormZAFOPT.CommandBrowseHyperspectralEDSData.Enabled = False
If MiscStringsAreSame(app.EXEName, "CalcImage") Then
If ZAFOptionSample(1).EDSSpectraFlag Then
FormZAFOPT.OptionEDSUse(0).Enabled = True
FormZAFOPT.OptionEDSUse(2).Enabled = True
FormZAFOPT.CommandBrowseHyperspectralEDSData.Enabled = True
End If
End If

FormZAFOPT.OptionEDSUse(0).Value = True
If ZAFOptionSample(1).EDSSpectraUseFlag Then
FormZAFOPT.OptionEDSUse(2).Value = True
End If

Exit Sub

' Errors
ZAFOptionLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFOptionLoad"
ierror = True
Exit Sub

End Sub

Sub ZAFOptionSave()
' Saves the option form

ierror = False
On Error GoTo ZAFOptionSaveError

Dim i As Integer, j As Integer, ip As Integer
Dim sym As String

' Check for proper oxygen flags
If FormZAFOPT.CheckDisplayAsOxide.Value = vbChecked Then
Call ZAFOptionCheckForOxygen
If ierror Then Exit Sub
End If

' Init calculation flags
ZAFOptionSample(1).DifferenceElementFlag% = False
ZAFOptionSample(1).StoichiometryElementFlag% = False
ZAFOptionSample(1).RelativeElementFlag% = False

' Check if oxygen is analyzed for, if changing to oxide calculation
ip% = IPOS1(ZAFOptionSample(1).LastChan%, Symlo$(8), ZAFOptionSample(1).Elsyms$())
If ip% > 0 And ip% <= ZAFOptionSample(1).LastElm% And FormZAFOPT.OptionOxide.Value Then
msg$ = "You cannot calculate oxygen by stoichiometry because Oxygen is already an Analyzed Element. "
msg$ = msg$ & "If you want to display the results as oxides, select Display As Oxides. "
MsgBox msg$
FormZAFOPT.OptionElemental.Value = True
End If

' Save oxide or elemental mode flag
If FormZAFOPT.OptionOxide.Value Then
ZAFOptionSample(1).OxideOrElemental% = 1
Else
ZAFOptionSample(1).OxideOrElemental% = 2
End If

' Save DisplayAsOxideFlag options
If FormZAFOPT.CheckDisplayAsOxide.Value = vbChecked Then
ZAFOptionSample(1).DisplayAsOxideFlag = True
Else
ZAFOptionSample(1).DisplayAsOxideFlag = False
End If

If FormZAFOPT.CheckAtomicPercents.Value = vbChecked Then
ZAFOptionSample(1).AtomicPercentFlag = True
Else
ZAFOptionSample(1).AtomicPercentFlag = False
End If

' Save other calculation options
ZAFOptionSample(1).DifferenceElement$ = vbNullString
If FormZAFOPT.ComboDifference.ListCount > 0 Then
If FormZAFOPT.CheckDifference.Value = vbChecked And FormZAFOPT.ComboDifference.ListIndex > -1 Then
i% = ZAFOptionSample(1).LastElm% + FormZAFOPT.ComboDifference.ListIndex + 1
If i% > ZAFOptionSample(1).LastElm% And i% <= ZAFOptionSample(1).LastChan% Then
ZAFOptionSample(1).DifferenceElement$ = ZAFOptionSample(1).Elsyms$(i%)
End If
End If
End If

' Save formula by difference
If FormZAFOPT.CheckDifferenceFormula.Value = vbChecked Then
ZAFOptionSample(1).DifferenceFormulaFlag = True
Else
ZAFOptionSample(1).DifferenceFormulaFlag = False
End If
ZAFOptionSample(1).DifferenceFormula$ = Trim$(FormZAFOPT.TextDifferenceFormula.Text)

' Convert formula by difference string to temp sample and check for errors
If Trim$(ZAFOptionSample(1).DifferenceFormula$) <> vbNullString Then
Call FormulaFormulaToSample(ZAFOptionSample(1).DifferenceFormula$, FormulaTmpSample())
If ierror Then Exit Sub
End If

' Check for both element by difference and formula by difference
If ZAFOptionSample(1).DifferenceElementFlag And ZAFOptionSample(1).DifferenceFormulaFlag Then
msg$ = "You cannot specify both the element by difference and the formula by difference flags at the same time."
MsgBox msg$, vbOKOnly + vbInformation, "ZAFOptionSave"
ierror = True
Exit Sub
End If

ZAFOptionSample(1).StoichiometryElement$ = vbNullString
ZAFOptionSample(1).StoichiometryRatio! = 0#
If FormZAFOPT.ComboStoichiometry.ListCount > 0 Then
If FormZAFOPT.CheckStoichiometry.Value = vbChecked And FormZAFOPT.ComboStoichiometry.ListIndex > -1 Then
If Val(FormZAFOPT.TextStoichiometry.Text) > 0# Then
i% = ZAFOptionSample(1).LastElm% + FormZAFOPT.ComboStoichiometry.ListIndex + 1
If i% > ZAFOptionSample(1).LastElm% And i% <= ZAFOptionSample(1).LastChan% Then
ZAFOptionSample(1).StoichiometryElement$ = ZAFOptionSample(1).Elsyms$(i%)
ZAFOptionSample(1).StoichiometryRatio! = Val(FormZAFOPT.TextStoichiometry.Text)
End If
End If
End If
End If

ZAFOptionSample(1).RelativeElement$ = vbNullString
ZAFOptionSample(1).RelativeToElement$ = vbNullString
ZAFOptionSample(1).RelativeRatio! = 0#
If FormZAFOPT.ComboRelative.ListCount > 0 And FormZAFOPT.ComboRelativeTo.ListCount > 0 Then
If FormZAFOPT.CheckRelative.Value = vbChecked And FormZAFOPT.ComboRelative.ListIndex > -1 And FormZAFOPT.ComboRelativeTo.ListIndex > -1 Then
If Val(FormZAFOPT.TextRelative.Text) > 0# Then
i% = ZAFOptionSample(1).LastElm% + FormZAFOPT.ComboRelative.ListIndex + 1
j% = FormZAFOPT.ComboRelativeTo.ListIndex + 1
If i% > ZAFOptionSample(1).LastElm% And i% <= ZAFOptionSample(1).LastChan% Then
If j% > 0 And j% <= ZAFOptionSample(1).LastChan% Then
ZAFOptionSample(1).RelativeElement$ = ZAFOptionSample(1).Elsyms$(i%)
ZAFOptionSample(1).RelativeToElement$ = ZAFOptionSample(1).Elsyms$(j%)
ZAFOptionSample(1).RelativeRatio! = Val(FormZAFOPT.TextRelative.Text)
End If
End If
End If
End If
End If

If FormZAFOPT.CheckCalculateElectronandXrayRanges = vbChecked Then
CalculateElectronandXrayRangesFlag = True
Else
CalculateElectronandXrayRangesFlag = False
End If

If FormZAFOPT.CheckUseOxygenFromHalogensCorrection.Value = vbChecked Then
UseOxygenFromHalogensCorrectionFlag = True
Else
UseOxygenFromHalogensCorrectionFlag = False
End If

' Set calculation flags
If ZAFOptionSample(1).DifferenceElement$ <> vbNullString Then ZAFOptionSample(1).DifferenceElementFlag% = True
If ZAFOptionSample(1).StoichiometryElement$ <> vbNullString Then ZAFOptionSample(1).StoichiometryElementFlag% = True
If ZAFOptionSample(1).RelativeElement$ <> vbNullString Then ZAFOptionSample(1).RelativeElementFlag% = True

' Save formula and mineral end member calculation option
ZAFOptionSample(1).FormulaElement$ = vbNullString
ZAFOptionSample(1).FormulaRatio! = 0#
ZAFOptionSample(1).MineralFlag% = 0
If FormZAFOPT.ComboFormula.ListCount > 0 Then
If FormZAFOPT.ComboFormula.ListIndex > -1 Then
If Val(FormZAFOPT.TextFormula.Text) > 0# Then
i% = FormZAFOPT.ComboFormula.ListIndex     ' zero index indicates sum all cations
ZAFOptionSample(1).FormulaRatio! = Val(FormZAFOPT.TextFormula.Text)

' If element is greater than zero then it is a specific cation (no element indicates sum all cations)
If i% > 0 And i% <= ZAFOptionSample(1).LastChan% Then
ZAFOptionSample(1).FormulaElement$ = ZAFOptionSample(1).Elsyms$(i%)
End If

End If
End If
End If

If FormZAFOPT.CheckFormula.Value = vbChecked Then
ZAFOptionSample(1).FormulaElementFlag% = True
Else
ZAFOptionSample(1).FormulaElementFlag% = False
End If

' Check for no formula atoms
If ZAFOptionSample(1).FormulaElementFlag% And ZAFOptionSample(1).FormulaRatio! = 0# Then GoTo ZAFOptionSaveNoFormulaAtoms

' Warn user if formula option is checked but no atoms is specified
'  (no element is ok since that indicates sum all cations)
If FormZAFOPT.CheckFormula.Value = vbChecked And ZAFOptionSample(1).FormulaRatio! = 0# Then
msg$ = "Formula option was selected, but no formula atoms were specified"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFOptionSave"
ierror = True
Exit Sub
End If

' Sample conductive coating
If FormZAFOPT.CheckCoatingFlag.Value = vbChecked Then
ZAFOptionSample(1).CoatingFlag% = 1
Else
ZAFOptionSample(1).CoatingFlag% = 0
End If

sym$ = FormZAFOPT.ComboCoatingElement.Text
ip% = IPOS1(MAXELM%, sym$, Symlo$())
If ip% = 0 Then
msg$ = "Not a valid element for the Sample Conductive Coating"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFOptionSave"
ierror = True
Exit Sub
End If
ZAFOptionSample(1).CoatingElement% = ip%

If Val(FormZAFOPT.TextCoatingDensity.Text) < 0.1 Or Val(FormZAFOPT.TextCoatingDensity.Text) > 50# Then
msg$ = "Density out of range for the Sample Conductive Coating (must be between 0.1 and 50 gm/cm^3)"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFOptionSave"
Else
ZAFOptionSample(1).CoatingDensity! = Val(FormZAFOPT.TextCoatingDensity.Text)
End If

If Val(FormZAFOPT.TextCoatingThickness.Text) < 1 Or Val(FormZAFOPT.TextCoatingThickness.Text) > 10000# Then
msg$ = "Thickness out of range for the Sample Conductive Coating (must be between 1 and 10,000 angstroms)"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFOptionSave"
Else
ZAFOptionSample(1).CoatingThickness! = Val(FormZAFOPT.TextCoatingThickness.Text)
End If

' Save coating globals (leave commented out so user has to explicitly turn on in Analytical menu)
If ZAFOptionSample(1).CoatingFlag% = 1 Then
'UseConductiveCoatingCorrectionForElectronAbsorption = True
'UseConductiveCoatingCorrectionForXrayTransmission = True
End If

' Sample density
If Val(FormZAFOPT.TextDensity.Text) < 0.1 Or Val(FormZAFOPT.TextDensity.Text) > 50# Then
msg$ = "Sample Density out of range (must be between 0.1 and 50 gm/cm^3)"
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFOptionSave"
Else
ZAFOptionSample(1).SampleDensity! = Val(FormZAFOPT.TextDensity.Text)
End If

' Save EDS options
ZAFOptionSample(1).EDSSpectraUseFlag = False
If FormZAFOPT.OptionEDSUse(2).Value Then
ZAFOptionSample(1).EDSSpectraUseFlag = True
End If

Exit Sub

' Errors
ZAFOptionSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFOptionSave"
ierror = True
Exit Sub

ZAFOptionSaveNoFormulaAtoms:
msg$ = "No formula atoms were specified. Either uncheck the Formula Element checkbox or specify the formula atoms."
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFOptionSave"
ierror = True
Exit Sub

End Sub

Sub ZAFOptionReturnSample(sample() As TypeSample)
' Returns the modified sample

ierror = False
On Error GoTo ZAFOptionReturnSampleError

sample(1) = ZAFOptionSample(1)

Exit Sub

' Errors
ZAFOptionReturnSampleError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFOptionReturnSample"
ierror = True
Exit Sub

End Sub

Sub ZAFOptionOxygen()
' Add oxygen an a specified element if Oxide calculation is selected

ierror = False
On Error GoTo ZAFOptionOxygenError

Dim ip As Integer

' Check if oxygen is already present in ZAFOptionsample
ip% = IPOS1(ZAFOptionSample(1).LastChan%, Symlo$(8), ZAFOptionSample(1).Elsyms$())
If ip% > 0 Then Exit Sub

' Check if too many elements
If ZAFOptionSample(1).LastChan% > MAXCHAN% Then Exit Sub

ZAFOptionSample(1).LastChan% = ZAFOptionSample(1).LastChan% + 1
ZAFOptionSample(1).Elsyms$(ZAFOptionSample(1).LastChan%) = Symlo$(8)
ZAFOptionSample(1).Xrsyms$(ZAFOptionSample(1).LastChan%) = vbNullString

ZAFOptionSample(1).numcat%(ZAFOptionSample(1).LastChan%) = AllCat%(8)
ZAFOptionSample(1).numoxd%(ZAFOptionSample(1).LastChan%) = AllOxd%(8)
ZAFOptionSample(1).ElmPercents!(ZAFOptionSample(1).LastChan%) = 0#

Exit Sub

' Errors
ZAFOptionOxygenError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFOptionOxygen"
ierror = True
Exit Sub

End Sub

Sub ZAFOptionCheckForOxygen()
' Routine to check for analyzed or non-zero specified oxygen, if user selects "DisplayAsOxides".

ierror = False
On Error GoTo ZAFOptionCheckForOxygenError

Dim ip As Integer

' Check for calculated oxygen
If ZAFOptionSample(1).LastElm% = 0 Or ZAFOptionSample(1).LastChan% = 0 Then Exit Sub
If FormZAFOPT.OptionOxide.Value = True Then Exit Sub

' Check for analyzed oxygen
ip% = IPOS1(ZAFOptionSample(1).LastElm%, Symlo$(8), ZAFOptionSample(1).Elsyms$())
If ip% > 0 Then Exit Sub

' Check if non-zero specified oxygen value
ip% = IPOS1(ZAFOptionSample(1).LastChan%, Symlo$(8), ZAFOptionSample(1).Elsyms$())
If ip% > 0 Then
If ZAFOptionSample(1).ElmPercents!(ip%) > 0# Then Exit Sub
End If

' Check if oxygen is element by difference
If ZAFOptionSample(1).DifferenceElementFlag% And MiscStringsAreSame(ZAFOptionSample(1).DifferenceElement$, Symlo$(8)) Then Exit Sub

' Check if oxygen is element by relative stoichiometry
If ZAFOptionSample(1).RelativeElementFlag% And MiscStringsAreSame(ZAFOptionSample(1).RelativeElement$, Symlo$(8)) Then Exit Sub

' Check if sample is a standard, since it will be automatically specified
If ZAFOptionSample(1).Type% = 1 Then Exit Sub

' Oxygen will not be calculated correctly
msg$ = "WARNING: Display As Oxides was selected, but oxygen is not either an "
msg$ = msg$ & "analyzed element or an unanalyzed element specified with a "
msg$ = msg$ & "concentration greater than zero. Therefore, the "
msg$ = msg$ & "analytical calculations will not be correct."
MsgBox msg$, vbOKOnly + vbExclamation, "ZAFOptionCheckForOxygen"

Exit Sub

' Errors
ZAFOptionCheckForOxygenError:
MsgBox Error$, vbOKOnly + vbCritical, "ZAFOptionCheckForOxygen"
ierror = True
Exit Sub

End Sub


