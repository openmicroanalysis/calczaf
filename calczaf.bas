Attribute VB_Name = "CodeCALCZAF"
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

Global Const MAXELMOLD% = 94            ' maximum elements (old format)

Type TypeMuOld
    mac(1 To 564) As Single
End Type

' Histogram variables
Global HistogramOutputOption As Integer

' Histogram data
Dim ConcError() As Single
Dim KratioExpr() As Single
Dim KratioCalc() As Single
Dim KratioError() As Single

Dim KratioLine() As Long
Dim KratioEsym() As String
Dim KratioXsym() As String
Dim KratioConc() As Single
Dim KratioTOAeO() As Single
Dim KratioOver() As Single

' CalcZAF sample structures
Dim CalcZAFTmpSample(1 To 1) As TypeSample
Dim CalcZAFOldSample(1 To 1) As TypeSample
Dim CalcZAFNewSample(1 To 1) As TypeSample  ' declare here to prevent access violation in CalcZAFSave

Dim CalcZAFAnalysis As TypeAnalysis

Dim CalcZAFRow As Integer
Dim CalcZAFLineCount As Long
Dim CalcZAFSampleCount As Integer
Dim CalcZAFOutputCount As Long

Dim KratioAlpha(1 To 2) As Single     ' calculated k-ratios from beta factors and concentrations (Beta/Conc)
Dim StdKFactors(1 To MAXCHAN%) As Single

' Input data
Dim UnkCounts(1 To MAXCHAN%) As Single
Dim StdCounts(1 To MAXCHAN%) As Single

' K-factor and Alpha factor output data
Dim kfactor(1 To MAXELM%, 1 To MAXELM%, 1 To MAXBINARY%) As Single     ' k-ratios at MAXBIN compositions
Dim oxfactor(1 To MAXELM%) As Single                                ' oxide end-member k-ratios

Dim alpha11(1 To MAXELM%, 1 To MAXELM%) As Single

Dim alpha21(1 To MAXELM%, 1 To MAXELM%) As Single
Dim alpha22(1 To MAXELM%, 1 To MAXELM%) As Single

Dim alpha31(1 To MAXELM%, 1 To MAXELM%) As Single
Dim alpha32(1 To MAXELM%, 1 To MAXELM%) As Single
Dim alpha33(1 To MAXELM%, 1 To MAXELM%) As Single

Dim alpha41(1 To MAXELM%, 1 To MAXELM%) As Single
Dim alpha42(1 To MAXELM%, 1 To MAXELM%) As Single
Dim alpha43(1 To MAXELM%, 1 To MAXELM%) As Single

Dim tzaftype As Integer, tmactype As Integer

Private Sub Main()
On Error Resume Next
FormSPLASH.Show vbModal
FormMAIN.Show
End Sub

Sub CalcZAFCalculate()
' Calculate ZAF parameters based on the current ZAF element setup

ierror = False
On Error GoTo CalcZAFCalculateError

Dim i As Integer, zerror As Integer
Dim excess As Single

' Check if particle corrections are selected with alpha factors or calibration curve (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If iptc% > 0 And CorrectionFlag% > 0 Then
msg$ = "Only ZAF or Phi-rho-z corrections are supported with particle/thin film corrections."
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFCalculate"
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

' Check for valid data
If CalcZAFOldSample(1).LastElm% < 1 Then GoTo CalcZAFCalculateNoElements

' Perform secondary fluorescence init (CalcZAF only calculates the first data line)
If UseSecondaryBoundaryFluorescenceCorrectionFlag Then
Call SecondaryInit(CalcZAFOldSample())
If ierror Then Exit Sub
Call SecondaryInitLine(Int(1), CalcZAFOldSample())
If ierror Then Exit Sub

For i% = 1 To CalcZAFOldSample(1).LastElm%
Call SecondaryInitChan(i%, CalcZAFOldSample())
If ierror Then Exit Sub
Next i%
End If

' Load default name and print
If Trim$(CalcZAFOldSample(1).Name$) = vbNullString Then
msg$ = "CalcZAF Sample (" & CalcZAFOldSample(1).Description$ & ") at " & Format$(CalcZAFOldSample(1).takeoff!) & " degrees and " & Format$(CalcZAFOldSample(1).kilovolts!) & " keV"
Call IOWriteLogRichText(vbCrLf & vbCrLf & msg$, vbNullString, Int(LogWindowFontSize% + 2), vbBlue, Int(FONT_BOLD% Or FONT_UNDERLINE%), Int(0))
Else
Call IOWriteLogRichText(vbCrLf & vbCrLf & CalcZAFOldSample(1).Name$, vbNullString, Int(LogWindowFontSize% + 2), vbBlue, Int(FONT_BOLD% Or FONT_UNDERLINE%), Int(0))
End If

' Make sure that new condition arrays are loaded
If Not CalcZAFOldSample(1).CombinedConditionsFlag Then
For i% = 1 To CalcZAFOldSample(1).LastChan%
CalcZAFOldSample(1).TakeoffArray!(i%) = CalcZAFOldSample(1).takeoff!
CalcZAFOldSample(1).KilovoltsArray!(i%) = CalcZAFOldSample(1).kilovolts!
CalcZAFOldSample(1).BeamCurrentArray!(i%) = DefaultBeamCurrent!
CalcZAFOldSample(1).BeamSizeArray!(i%) = DefaultBeamSize!
Next i%
End If

' Force unknown if type not specified
If CalcZAFOldSample(1).Type% = 0 Then CalcZAFOldSample(1).Type% = 2
CalcZAFOldSample(1).Datarows% = 1   ' always a single data point
CalcZAFOldSample(1).GoodDataRows% = 1
CalcZAFOldSample(1).LineStatus(1) = True      ' force status flag always true (good data point)
CalcZAFOldSample(1).AtomicPercentFlag% = True

' CALCZAF Calculate intensity from weight
If CalcZAFMode% = 0 Then

' Reload the element arrays
Call ElementGetData(CalcZAFOldSample())
If ierror Then Exit Sub

' Initialize calculations (needed for ZAFPTC and coating calculations) (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% <> MAXCORRECTION% Then
Call ZAFSetZAF(CalcZAFOldSample())
If ierror Then Exit Sub
Else
'Call ZAFSetZAF3(CalcZAFOldSample())
'If ierror Then Exit Sub
End If

' Force standard assignment for intensity calculation
For i% = 1 To CalcZAFOldSample(1).LastElm%
CalcZAFOldSample(1).StdAssigns%(i%) = MAXINTEGER%     ' fake standard assignment
Next i%

' Set TmpSample equal to OldSample so k factors and ZAF corrections get loaded in ZAFStd
CalcZAFTmpSample(1) = CalcZAFOldSample(1)
CalcZAFTmpSample(1).number% = MAXINTEGER%             ' fake standard number

' Fake sample coating for ZAFStd calculation
If UseConductiveCoatingCorrectionForElectronAbsorption Then                   ' fake standard coating
StandardCoatingFlag%(1) = CalcZAFOldSample(1).CoatingFlag%
StandardCoatingDensity!(1) = CalcZAFOldSample(1).CoatingDensity!
StandardCoatingThickness!(1) = CalcZAFOldSample(1).CoatingThickness!
StandardCoatingElement%(1) = CalcZAFOldSample(1).CoatingElement%
End If

' Run the intensity from concentration calculations on the "standard"
If CorrectionFlag% = 0 Then
Call ZAFStd2(Int(1), CalcZAFAnalysis, CalcZAFOldSample(), CalcZAFTmpSample())
If ierror Then Exit Sub
ElseIf CorrectionFlag% = MAXCORRECTION% Then
'Call ZAFStd3(Int(1), CalcZAFAnalysis, CalcZAFOldSample(), CalcZAFTmpSample())
'If ierror Then Exit Sub

' Calculate the standard beta factors for this standard
Else
AllAFactorUpdateNeeded = True   ' indicate alpha-factor update
Call AFactorStd(Int(1), CalcZAFAnalysis, CalcZAFOldSample(), CalcZAFTmpSample())
If ierror Then Exit Sub

Call AFactorTypeStandard(CalcZAFAnalysis, CalcZAFOldSample())
If ierror Then Exit Sub
End If

' CALCZAF Calculate weight from intensity
Else
AllAFactorUpdateNeeded = True   ' indicate alpha-factor update
CalcZAFOldSample(1).Type% = 2   ' assume unknown type for all samples

' Reload the element arrays based on the unknown sample setup
Call ElementGetData(CalcZAFOldSample())
If ierror Then Exit Sub

' Initialize calculations (needed for ZAFPTC calculations) (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% <> MAXCORRECTION% Then
Call ZAFSetZAF(CalcZAFOldSample())
If ierror Then Exit Sub
Else
'Call ZAFSetZAF3(CalcZAFOldSample())
'If ierror Then Exit Sub
End If

' Check for standard assignments
If CalcZAFMode% < 3 Then
For i% = 1 To CalcZAFOldSample(1).LastElm%
If CalcZAFOldSample(1).StdAssigns%(i%) = 0 Then GoTo CalcZAFCalculateNoStdAssigned
Next i%

' Calculate and load assigned standard k-factors
Call UpdateStdKfacs(CalcZAFAnalysis, CalcZAFOldSample(), CalcZAFTmpSample())
If ierror Then Exit Sub

' No assigned standard used in k-ratio calculation
Else
For i% = 1 To CalcZAFOldSample(1).LastElm%
CalcZAFOldSample(1).StdAssigns%(i%) = MAXINTEGER%     ' fake standard assignment
Next i%
End If

' Load alpha-factor arrays for this sample (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% > 0 And CorrectionFlag% < 5 Then
Call AFactorLoadFactors(CalcZAFAnalysis, CalcZAFOldSample())
If ierror Then Exit Sub
End If

' Calculate primary intensity arrays for this sample (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% = 0 Or CorrectionFlag% = MAXCORRECTION% Then

' Reload the element arrays
Call ElementGetData(CalcZAFOldSample())
If ierror Then Exit Sub

' Initialize sample calculations
If CorrectionFlag% <> MAXCORRECTION% Then
Call ZAFSetZAF(CalcZAFOldSample())
If ierror Then Exit Sub
Else
'Call ZAFSetZAF3(CalcZAFOldSample())
'If ierror Then Exit Sub
End If

' Calculate oxygen channel
Else
Call ZAFGetOxygenChannel(CalcZAFOldSample())
If ierror Then Exit Sub
End If

' Initialize the analysis
Call InitLine(CalcZAFAnalysis)
If ierror Then Exit Sub

' Init intensities for unknown and standard
For i% = 1 To CalcZAFOldSample(1).LastChan%
CalcZAFAnalysis.WtPercents!(i%) = 0#

' Counts
If CalcZAFMode% = 1 Then
CalcZAFAnalysis.StdAssignsCounts!(i%) = StdCounts!(i%)

' K-raws
ElseIf CalcZAFMode% = 2 Then
CalcZAFAnalysis.StdAssignsCounts!(i%) = 1#

' K-ratios
ElseIf CalcZAFMode% = 3 Then
CalcZAFAnalysis.StdAssignsCounts!(i%) = 1#
CalcZAFAnalysis.StdAssignsKfactors!(i%) = 1#
CalcZAFAnalysis.StdAssignsBetas!(i%) = 1#
CalcZAFAnalysis.StdAssignsPercents!(i%) = 100#
End If

Next i%

' Load specified weight percents for this sample
excess! = 0#
For i% = CalcZAFOldSample(1).LastElm% + 1 To CalcZAFOldSample(1).LastChan%
CalcZAFAnalysis.WtPercents!(i%) = CalcZAFOldSample(1).ElmPercents!(i%)

' Store excess oxygen for iteration
If i% = CalcZAFOldSample(1).OxygenChannel% Then
excess! = CalcZAFAnalysis.WtPercents!(i%)
End If
Next i%

' Save analysis and alpha-factor arrays if using alpha-factors (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calibration curve, 6 = fundamental parameters)
If CorrectionFlag% > 0 And CorrectionFlag% < 5 Then
Call AFactorAFASaveFactors(CalcZAFAnalysis, CalcZAFOldSample())
If ierror Then Exit Sub
End If

' Calculate ZAF weights for first data line only in CalcZAF
If CorrectionFlag% = 0 Then
Call ZAFSmp(Int(1), UnkCounts!(), zerror%, CalcZAFAnalysis, CalcZAFOldSample())
If ierror Then Exit Sub

' Calculate alpha-factor weights
ElseIf CorrectionFlag% > 0 And CorrectionFlag% < 5 Then
Call AFactorSmp(Int(1), excess!, UnkCounts!(), zerror%, CalcZAFAnalysis, CalcZAFOldSample())
If ierror Then Exit Sub

' Fundamental parameter correction
ElseIf CorrectionFlag% = MAXCORRECTION% Then
'Call ZAFSmp3(Int(1), UnkCounts!(), zerror%, CalcZAFAnalysis, CalcZAFOldSample())
'If ierror Then Exit Sub
End If
End If

' Re-load grid
Call CalcZAFLoadList
If ierror Then Exit Sub

' Calculate electron and x-ray ranges
If CalculateElectronandXrayRangesFlag Then
Call ZAFCalculateRange(CalcZAFAnalysis, CalcZAFOldSample)
If ierror Then Exit Sub
End If

' Plot phi-rho-z curves if specified
If CalculatePhiRhoZPlotCurves Then
If CorrectionFlag% = 0 And iabs% >= 7 Then
FormPlotPhiRhoZ.Show vbModeless
Call PlotPhiRhoZCurves(CalcZAFOldSample())
If ierror Then Exit Sub

Else
Unload FormPlotPhiRhoZ
Call MiscPlotInit(FormPlotPhiRhoZ.Pesgo1, True)
If ierror Then Exit Sub
End If
End If

Exit Sub

' Errors
CalcZAFCalculateError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFCalculate"
ierror = True
Exit Sub

CalcZAFCalculateNoElements:
msg$ = "No elements were specified for the current sample " & CalcZAFOldSample(1).Name$ & "."
msg$ = msg$ & vbCrLf & vbCrLf & "Please specify the elements "
msg$ = msg$ & "and x-ray emitters and composition or intensities first by clicking the File | Open menu or by entering a "
msg$ = msg$ & "composition by formula or from the standard database or by clicking an element row to enter data manually."
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFCalculate"
ierror = True
Exit Sub

CalcZAFCalculateNoStdAssigned:
msg$ = "No standard assigned for " & CalcZAFOldSample(1).Elsyms$(i%) & " " & CalcZAFOldSample(1).Xrsyms$(i%)
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFCalculate"
ierror = True
Exit Sub

End Sub

Sub CalcZAFElementLoad(elementrow As Integer)
' Load element setup for a single element

ierror = False
On Error GoTo CalcZAFElementLoadError

Dim i As Integer, ip As Integer

CalcZAFRow% = elementrow%

' Add the list box items
FormZAFELM.ComboElement.Clear
For i% = 0 To MAXELM% - 1
FormZAFELM.ComboElement.AddItem Symlo$(i% + 1)
Next i%

FormZAFELM.ComboXRay.Clear
For i% = 0 To MAXRAY% - 1
FormZAFELM.ComboXRay.AddItem Xraylo$(i% + 1)
Next i%

FormZAFELM.ComboCations.Clear
For i% = 1 To MAXCATION% - 1    ' 1 to 99
FormZAFELM.ComboCations.AddItem Format$(i%)
Next i%

FormZAFELM.ComboOxygens.Clear
For i% = 0 To MAXCATION% - 1    ' 0 to 99
FormZAFELM.ComboOxygens.AddItem Format$(i%)
Next i%

' Load the primary assigned standard combo selections
FormZAFELM.ComboStandard.Clear
For i% = 1 To NumberofStandards%
msg$ = Format$(StandardNumbers(i%), a40) & " " & StandardNames$(i%)
FormZAFELM.ComboStandard.AddItem msg$
Next i%

' Load the element properties. Note, to avoid overloading by
' text change events, these text fields MUST be loading in order
' by element, xray, etc.
If CalcZAFRow% > 0 Then
FormZAFELM.ComboElement.Text = CalcZAFOldSample(1).Elsyms$(CalcZAFRow%)
FormZAFELM.ComboXRay.Text = CalcZAFOldSample(1).Xrsyms$(CalcZAFRow%)
FormZAFELM.ComboCations.Text = Format$(CalcZAFOldSample(1).numcat%(CalcZAFRow%))
FormZAFELM.ComboOxygens.Text = Format$(CalcZAFOldSample(1).numoxd%(CalcZAFRow%))
End If

' Load the default assigned (primary) standard
If FormZAFELM.ComboStandard.ListCount > 0 Then
ip% = IPOS2(NumberofStandards%, CalcZAFOldSample(1).StdAssigns%(CalcZAFRow%), StandardNumbers%())
If ip% > 0 Then FormZAFELM.ComboStandard.ListIndex = ip% - 1
End If

' Load weight percent and intensity
FormZAFELM.TextWeight.Text = MiscAutoFormat$(CalcZAFOldSample(1).ElmPercents!(CalcZAFRow%))
FormZAFELM.TextIntensity.Text = MiscAutoFormat$(UnkCounts!(CalcZAFRow%))

' Load standard intensity
FormZAFELM.TextIntensityStd.Text = MiscAutoFormat$(StdCounts!(CalcZAFRow%))

' Set enables for calculate intensities
If FormZAF.OptionCalculate(0).Value Then
FormZAFELM.TextIntensity.Enabled = False
FormZAFELM.ComboStandard.Enabled = False
FormZAFELM.TextIntensityStd.Enabled = False
FormZAFELM.OptionAnalyzed.Enabled = False
FormZAFELM.OptionSpecified.Enabled = False
FormZAFELM.TextWeight.Enabled = True
FormZAFELM.CommandAddStandardsToRun.Enabled = False

' Counts
ElseIf FormZAF.OptionCalculate(1).Value Then
FormZAFELM.TextWeight.Enabled = False
FormZAFELM.TextIntensity.Enabled = True
FormZAFELM.ComboStandard.Enabled = True
FormZAFELM.TextIntensityStd.Enabled = True
FormZAFELM.CommandAddStandardsToRun.Enabled = True

' K-raws
ElseIf FormZAF.OptionCalculate(2).Value Then
FormZAFELM.TextWeight.Enabled = False
FormZAFELM.TextIntensity.Enabled = True
FormZAFELM.ComboStandard.Enabled = True
FormZAFELM.TextIntensityStd.Enabled = False
FormZAFELM.TextIntensityStd.Text = MiscAutoFormat$(CSng(1#))
FormZAFELM.CommandAddStandardsToRun.Enabled = True

' K-ratios
ElseIf FormZAF.OptionCalculate(3).Value Then
FormZAFELM.TextWeight.Enabled = False
FormZAFELM.TextIntensity.Enabled = True
FormZAFELM.ComboStandard.Enabled = False
FormZAFELM.TextIntensityStd.Enabled = False
FormZAFELM.TextIntensityStd.Text = MiscAutoFormat$(CSng(1#))
FormZAFELM.CommandAddStandardsToRun.Enabled = False
End If

' Set enables for specified element
If CalcZAFOldSample(1).Elsyms$(CalcZAFRow%) <> vbNullString And CalcZAFOldSample(1).Xrsyms$(CalcZAFRow%) = vbNullString Then
FormZAFELM.OptionSpecified.Value = True

FormZAFELM.TextWeight.Enabled = True
FormZAFELM.TextIntensity.Enabled = False
FormZAFELM.ComboStandard.Enabled = False
FormZAFELM.TextIntensityStd.Enabled = False
End If

FormZAFELM.Show vbModal

Exit Sub

' Errors
CalcZAFElementLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFElementLoad"
ierror = True
Exit Sub

End Sub

Sub CalcZAFElementSave()
' Save element setup for a single element

ierror = False
On Error GoTo CalcZAFElementSaveError

Dim sym As String
Dim i As Integer
Dim ip  As Integer, ipp As Integer
Dim keV As Single, lam As Single

ReDim numbers(1 To MAXCATION%) As Integer

For i% = 1 To MAXCATION%
numbers(i%) = MAXCATION% - i%       ' load in 0 to MAXCATION% - 1 in reverse value order
Next i%

' Check if adding analyzed oxygen to an oxide calculated sample, if so change to elemental calculation
sym$ = FormZAFELM.ComboElement.Text
ip% = IPOS1(MAXELM%, sym$, Symlo$())
sym$ = FormZAFELM.ComboXRay.Text
ipp% = IPOS1(MAXRAY%, sym$, Xraylo$())  ' including unanalyzed elements
If ip% = AllAtomicNums%(ATOMIC_NUM_OXYGEN%) And ipp% <= MAXRAY% - 1 Then CalcZAFOldSample(1).OxideOrElemental% = 2

' Get the element symbol
CalcZAFOldSample(1).Elsyms$(CalcZAFRow%) = vbNullString
sym$ = FormZAFELM.ComboElement.Text
If sym$ = vbNullString Then           ' deleted element
CalcZAFOldSample(1).Elsyms$(CalcZAFRow%) = vbNullString
CalcZAFOldSample(1).Xrsyms$(CalcZAFRow%) = vbNullString
Exit Sub
End If

' Get element symbol
ipp% = IPOS1(MAXELM%, sym$, Symlo$())
If ipp% = 0 Then GoTo CalcZAFSaveSaveBadElement
CalcZAFOldSample(1).Elsyms$(CalcZAFRow%) = sym$
    
' Get the xray symbol
CalcZAFOldSample(1).Xrsyms$(CalcZAFRow%) = vbNullString
sym$ = FormZAFELM.ComboXRay.Text

' Only save the xray symbol as default if it is analyzed
ip% = IPOS1(MAXRAY%, sym$, Xraylo$())
If ip% = 0 Then GoTo CalcZAFSaveSaveBadXray
CalcZAFOldSample(1).Xrsyms$(CalcZAFRow%) = sym$
If ip% <= MAXRAY% - 1 Then Deflin$(ipp%) = CalcZAFOldSample(1).Xrsyms$(CalcZAFRow%)

' Check for a valid xray line, if analyzed element
If ip% <= MAXRAY% - 1 Then
Call XrayGetKevLambda(Symlo$(ipp%), Xraylo$(ip%), keV!, lam!)
If ierror Then Exit Sub
End If

' Save the cation and oxygen subscripts, note that "NumCat" must be at least one
CalcZAFOldSample(1).numcat%(CalcZAFRow%) = 0
i% = Val(FormZAFELM.ComboCations.Text)
ip% = IPOS2(MAXCATION% - 1, i%, numbers())      ' zero is not valid
If ip% < 1 Then GoTo CalcZAFSaveSaveBadCation
CalcZAFOldSample(1).numcat%(CalcZAFRow%) = i%
AllCat%(ipp%) = CalcZAFOldSample(1).numcat%(CalcZAFRow%)

CalcZAFOldSample(1).numoxd%(CalcZAFRow%) = 0
i% = Val(FormZAFELM.ComboOxygens.Text)
ip% = IPOS2(MAXCATION%, i%, numbers())          ' zero is valid
If ip% = 0 Then GoTo CalcZAFSaveSaveBadOxygen
CalcZAFOldSample(1).numoxd%(CalcZAFRow%) = i%
AllOxd%(ipp%) = CalcZAFOldSample(1).numoxd%(CalcZAFRow%)

' Save weight/intensities
CalcZAFOldSample(1).ElmPercents!(CalcZAFRow%) = Val(FormZAFELM.TextWeight.Text)
UnkCounts!(CalcZAFRow%) = Val(FormZAFELM.TextIntensity.Text)

' Save the standard assignments for this element
CalcZAFOldSample(1).StdAssigns%(CalcZAFRow%) = MAXINTEGER%
If FormZAFELM.ComboStandard.ListCount > 0 Then
If FormZAFELM.ComboStandard.ListIndex > -1 Then
ip% = FormZAFELM.ComboStandard.ListIndex + 1
If ip% >= 1 And ip% <= NumberofStandards% Then
CalcZAFOldSample(1).StdAssigns%(CalcZAFRow%) = StandardNumbers%(ip%)
End If
End If
End If

' Save standard intensity for this element
StdCounts!(CalcZAFRow%) = Val(FormZAFELM.TextIntensityStd.Text)

' Save analyzed type
If FormZAFELM.OptionAnalyzed.Value = True Then
If CalcZAFOldSample(1).Xrsyms$(CalcZAFRow%) = vbNullString Then GoTo CalcZAFElementSaveNoXray
End If

If FormZAFELM.OptionSpecified.Value = True Then
CalcZAFOldSample(1).Xrsyms$(CalcZAFRow%) = vbNullString
End If

Exit Sub

' Errors
CalcZAFElementSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFElementSave"
ierror = True
Exit Sub

CalcZAFSaveSaveBadElement:
msg$ = "Element " & sym$ & " is an invalid element symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFSaveSave"
ierror = True
Exit Sub

CalcZAFSaveSaveBadXray:
msg$ = "Xray " & sym$ & " is an invalid xray symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFSaveSave"
ierror = True
Exit Sub

CalcZAFSaveSaveBadCation:
msg$ = "Invalid number of cations"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFSaveSave"
ierror = True
Exit Sub

CalcZAFSaveSaveBadOxygen:
msg$ = "Invalid number of oxygens"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFSaveSave"
ierror = True
Exit Sub

CalcZAFElementSaveNoXray:
msg$ = "Analyzed element requires an x-ray line to be specified"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFSaveSave"
ierror = True
Exit Sub

End Sub

Sub CalcZAFElementUpdate()
' Updates the xray, cation and oxygen combos if the element changes

ierror = False
On Error GoTo CalcZAFElementUpdateError

Dim ip As Integer
Dim sym As String

sym$ = FormZAFELM.ComboElement.Text
ip% = IPOS1(MAXELM%, sym$, Symlo$())

' Only update xray if no data
If ip% > 0 Then
If FormZAFELM.ComboXRay.Text = vbNullString Then FormZAFELM.ComboXRay.Text = Deflin$(ip%)
If sym$ <> CalcZAFOldSample(1).Elsyms$(CalcZAFRow%) Then FormZAFELM.ComboXRay.Text = Deflin$(ip%)

If FormZAFELM.ComboCations.Text = vbNullString Then FormZAFELM.ComboCations.Text = AllCat%(ip%)
If sym$ <> CalcZAFOldSample(1).Elsyms$(CalcZAFRow%) Then FormZAFELM.ComboCations.Text = AllCat%(ip%)

If FormZAFELM.ComboOxygens.Text = vbNullString Then FormZAFELM.ComboOxygens.Text = AllOxd%(ip%)
If sym$ <> CalcZAFOldSample(1).Elsyms$(CalcZAFRow%) Then FormZAFELM.ComboOxygens.Text = AllOxd%(ip%)
End If

Exit Sub

' Errors
CalcZAFElementUpdateError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFElementUpdate"
ierror = True
Exit Sub

End Sub

Sub CalcZAFGetMode(Index As Integer)
' Save the current ZAF calculation mode

ierror = False
On Error GoTo CalcZAFGetModeError

' Save default correction type
CalcZAFMode% = Index%

' Disable All Matrix Corrections if k-ratio mode
If CalcZAFMode% = 0 Then
FormZAF.CheckUseAllMatrixCorrections.Enabled = False
Else
FormZAF.CheckUseAllMatrixCorrections.Enabled = True
End If

Exit Sub

' Errors
CalcZAFGetModeError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFGetMode"
ierror = True
Exit Sub

End Sub

Sub CalcZAFImportClose()
' Close the import file

ierror = False
On Error GoTo CalcZAFImportCloseError

' Close file
Close #ImportDataFileNumber%

' Initialize
Call CalcZAFInit
If ierror Then Exit Sub

' Re-load grid
ImportDataFile$ = vbNullString
ImportDataFile2$ = vbNullString
Call CalcZAFLoadList
If ierror Then Exit Sub

FormMAIN.Caption = "CalcZAF (Calculate ZAF and Phi-Rho-Z Corrections)"
FormZAF.Caption = "Calculate ZAF Corrections"

' Set enables
Call CalcZAFSetEnables
If ierror Then Exit Sub

Exit Sub

' Errors
CalcZAFImportCloseError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFImportClose"
ierror = True
Exit Sub

End Sub

Sub CalcZAFImportNext()
' Load the next data set line from the import file

ierror = False
On Error GoTo CalcZAFImportNextError

Dim i As Integer, ip As Integer
Dim astring As String, bstring As String

' Check for end of file
If EOF(ImportDataFileNumber%) Then GoTo CalcZAFImportNextEOF

' Initialize
Call CalcZAFInit
If ierror Then Exit Sub

' Init sample
Call InitSample(CalcZAFOldSample())
If ierror Then Exit Sub
Call InitSample(CalcZAFTmpSample())
If ierror Then Exit Sub
Call InitSample(CalcZAFNewSample())
If ierror Then Exit Sub

' Load particle parameters
CalcZAFOldSample(1).iptc% = iptc%
CalcZAFOldSample(1).PTCModel% = PTCModel%
CalcZAFOldSample(1).PTCDiameter! = PTCDiameter!
CalcZAFOldSample(1).PTCDensity! = PTCDensity!
CalcZAFOldSample(1).PTCThicknessFactor! = PTCThicknessFactor!
CalcZAFOldSample(1).PTCNumericalIntegrationStep! = PTCNumericalIntegrationStep!

' Increment sample count
CalcZAFSampleCount% = CalcZAFSampleCount% + 1

' Read first line and parse
CalcZAFLineCount& = CalcZAFLineCount& + 1
Line Input #ImportDataFileNumber%, astring$

' Check for wrong format in file (wrong .dat file!)
If Left$(astring$, 1) <> "0" And Left$(astring$, 1) <> "1" And Left$(astring$, 1) <> "2" And Left$(astring$, 1) <> "3" Then GoTo CalcZAFImportNextWrongFile
If InStr(astring$, ",") = 0 Then GoTo CalcZAFImportNextWrongFile

' Parse mode using comma delimiter
Call MiscParseStringToStringA(astring$, ",", bstring$)
If ierror Then Exit Sub
CalcZAFMode% = Val(bstring$)

' Parse lastchan
Call MiscParseStringToStringA(astring$, ",", bstring$)
If ierror Then Exit Sub
CalcZAFOldSample(1).LastChan% = Val(bstring$)

' Parse kilovolts
Call MiscParseStringToStringA(astring$, ",", bstring$)
If ierror Then Exit Sub
CalcZAFOldSample(1).kilovolts! = Val(bstring$)

' Parse takeoff
Call MiscParseStringToStringA(astring$, ",", bstring$)
If ierror Then Exit Sub
CalcZAFOldSample(1).takeoff! = Val(bstring$)

' Check for sample name
If Trim$(astring$) = vbNullString Then
CalcZAFOldSample(1).Name$ = UCase$(MiscGetFileNameOnly$(ImportDataFile$)) & ", Sample" & Str$(CalcZAFSampleCount%)
CalcZAFOldSample(1).StagePositions!(1, 1) = 0#      ' X
CalcZAFOldSample(1).StagePositions!(1, 2) = 0#      ' Y
CalcZAFOldSample(1).StagePositions!(1, 3) = 0#      ' Z

' Sample name string found
Else

' Remove commas between double quotes enclosing sample name
ip% = InStr(astring$, VbDquote & VbComma)   ' find end of sample name string
If ip% = 0 Then ip% = Len(astring$)     ' no sample coordinates
For i% = 1 To ip%
If Mid$(astring$, i%, 1) = VbComma Then
Mid$(astring$, i%, 1) = Space(1)
End If
Next i%

' Parse the sameple name string (now without enclosed commas)
Call MiscParseStringToStringA(astring$, ",", bstring$)
If ierror Then Exit Sub
CalcZAFOldSample(1).Name$ = Trim$(bstring$)

' Check for sample coordinates
If Trim$(astring$) = vbNullString Then
CalcZAFOldSample(1).StagePositions!(1, 1) = 0#      ' X
CalcZAFOldSample(1).StagePositions!(1, 2) = 0#      ' Y
CalcZAFOldSample(1).StagePositions!(1, 3) = 0#      ' Z

' Sample coordinates string found
Else
Call MiscParseStringToStringA(astring$, ",", bstring$)
If ierror Then Exit Sub
CalcZAFOldSample(1).StagePositions!(1, 1) = Val(bstring$)      ' X

Call MiscParseStringToStringA(astring$, ",", bstring$)
If ierror Then Exit Sub
CalcZAFOldSample(1).StagePositions!(1, 2) = Val(bstring$)      ' Y

Call MiscParseStringToStringA(astring$, ",", bstring$)
If ierror Then Exit Sub
CalcZAFOldSample(1).StagePositions!(1, 3) = Val(bstring$)      ' Z

End If
End If

' Check for valid mode and number of elements, kilovolts and takeoff
If CalcZAFMode% < 0 Or CalcZAFMode% > 3 Then GoTo CalcZAFImportNextBadMode
If CalcZAFOldSample(1).kilovolts! < 1# Or CalcZAFOldSample(1).kilovolts! > 100# Then GoTo CalcZAFImportNextBadKeV
If CalcZAFOldSample(1).takeoff! < 1# Or CalcZAFOldSample(1).takeoff! > 90# Then GoTo CalcZAFImportNextBadTakeoff
If CalcZAFOldSample(1).LastChan% < 1 Or CalcZAFOldSample(1).LastChan% > MAXCHAN% Then GoTo CalcZAFImportNextTooMany

' Update defaults
DefaultTakeOff! = CalcZAFOldSample(1).takeoff!
DefaultKiloVolts! = CalcZAFOldSample(1).kilovolts!

' Read oxide, difference, stoichiometry, relative parameters
CalcZAFLineCount& = CalcZAFLineCount& + 1
Input #ImportDataFileNumber%, CalcZAFOldSample(1).OxideOrElemental%, CalcZAFOldSample(1).DifferenceElement$, CalcZAFOldSample(1).StoichiometryElement$, CalcZAFOldSample(1).StoichiometryRatio!, CalcZAFOldSample(1).RelativeElement$, CalcZAFOldSample(1).RelativeToElement$, CalcZAFOldSample(1).RelativeRatio!

' Set calculation flags
If CalcZAFOldSample(1).DifferenceElement$ <> vbNullString Then CalcZAFOldSample(1).DifferenceElementFlag% = True
If CalcZAFOldSample(1).StoichiometryElement$ <> vbNullString Then CalcZAFOldSample(1).StoichiometryElementFlag% = True
If CalcZAFOldSample(1).RelativeElement$ <> vbNullString Then CalcZAFOldSample(1).RelativeElementFlag% = True
CalcZAFLineCount& = CalcZAFLineCount& + 1

' Loop on each element
NumberofStandards% = 0
For i% = 1 To CalcZAFOldSample(1).LastChan%
CalcZAFLineCount& = CalcZAFLineCount& + 1
Input #ImportDataFileNumber%, CalcZAFOldSample(1).Elsyms$(i%), CalcZAFOldSample(1).Xrsyms$(i%)
Input #ImportDataFileNumber%, CalcZAFOldSample(1).numcat%(i%), CalcZAFOldSample(1).numoxd%(i%)
Input #ImportDataFileNumber%, CalcZAFOldSample(1).StdAssigns%(i%), CalcZAFOldSample(1).ElmPercents!(i%), UnkCounts!(i%), StdCounts!(i%)

' Add standard to run
If CalcZAFOldSample(1).StdAssigns%(i%) > 0 And CalcZAFOldSample(1).StdAssigns%(i%) <> MAXINTEGER% Then
ip% = IPOS2(NumberofStandards%, CalcZAFOldSample(1).StdAssigns%(i%), StandardNumbers%())
If ip% = 0 Then
Call AddStdSaveStd(CalcZAFOldSample(1).StdAssigns%(i%))
If ierror Then Exit Sub
End If
End If
Next i%

' Update sample number
CalcZAFOldSample(1).number% = CalcZAFSampleCount%           ' for standard (k-ratio calculation)
CalcZAFOldSample(1).Linenumber&(1) = CalcZAFSampleCount%    ' for unknown (other calculations)

' Always a single line
CalcZAFOldSample(1).Datarows = 1
CalcZAFOldSample(1).GoodDataRows = 1
CalcZAFOldSample(1).LineStatus(1) = True

' Sort elements
Call CalcZAFSave
If ierror Then Exit Sub

' Load combined conditions for all elements
For i% = 1 To CalcZAFOldSample(1).LastElm%
CalcZAFOldSample(1).TakeoffArray!(i%) = CalcZAFOldSample(1).takeoff!
CalcZAFOldSample(1).KilovoltsArray!(i%) = CalcZAFOldSample(1).kilovolts!
CalcZAFOldSample(1).BeamCurrentArray!(i%) = DefaultBeamCurrent!
CalcZAFOldSample(1).BeamSizeArray!(i%) = DefaultBeamSize!
Next i%

' Set the mode
FormZAF.OptionCalculate(CalcZAFMode%).Value = True

' Re-load grid
Call CalcZAFLoadList
If ierror Then Exit Sub

' Display sample name
FormZAF.Caption = "Calculate ZAF Corrections    [" & CalcZAFOldSample(1).Name$ & "]"
If CalcZAFOldSample(1).StagePositions!(1, 1) <> 0# Or CalcZAFOldSample(1).StagePositions!(1, 2) <> 0# Then
FormZAF.Caption = FormZAF.Caption & " [" & Format$(CalcZAFOldSample(1).StagePositions!(1, 1)) & ", " & Format$(CalcZAFOldSample(1).StagePositions!(1, 2)) & "]"
End If

Exit Sub

' Errors
CalcZAFImportNextError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFImportNext"
Call CalcZAFImportClose
ierror = True
Exit Sub

CalcZAFImportNextWrongFile:
CalcZAFMode% = 0
msg$ = "The CalcZAF calculation mode parameter is out of range (the data file may have the wrong format) in " & ImportDataFile$ & " on line " & Str$(CalcZAFLineCount&)
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFImportNext"
ierror = True
Exit Sub

CalcZAFImportNextBadMode:
CalcZAFMode% = 0
msg$ = "The CalcZAF calculation mode parameter is out of range (the data file may have the wrong format) in " & ImportDataFile$ & " on line " & Str$(CalcZAFLineCount&)
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFImportNext"
ierror = True
Exit Sub

CalcZAFImportNextBadKeV:
msg$ = "Kilovolts out of range in " & ImportDataFile$ & " on line " & Str$(CalcZAFLineCount&)
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFImportNext"
ierror = True
Exit Sub

CalcZAFImportNextBadTakeoff:
msg$ = "Takeoff out of range in " & ImportDataFile$ & " on line " & Str$(CalcZAFLineCount&)
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFImportNext"
ierror = True
Exit Sub

CalcZAFImportNextTooMany:
msg$ = "Too many elements in " & ImportDataFile$ & " on line " & Str$(CalcZAFLineCount&)
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFImportNext"
ierror = True
Exit Sub

CalcZAFImportNextEOF:
msg$ = "No more data set lines in " & ImportDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFImportNext"
ierror = True
Exit Sub

End Sub

Sub CalcZAFImportOpen(tForm As Form)
' Open the import file

ierror = False
On Error GoTo CalcZAFImportOpenError

Dim tfilename As String

' Get available standard names and numbers from database
Call StandardGetMDBIndex
If ierror Then Exit Sub

' Make sure no file is open
If ImportDataFile$ <> vbNullString Then
Call CalcZAFImportClose
If ierror Then Exit Sub
End If

' Get filename from user
If ImportDataFile$ <> vbNullString Then
tfilename$ = ImportDataFile$
Else
tfilename$ = "Calczaf.dat"
End If
Call IOGetFileName(Int(2), "DAT", tfilename$, tForm)
If ierror Then Exit Sub

' Save current path
CalcZAFDATFileDirectory$ = MiscGetPathOnly$(tfilename$)

' No errors, save file name
ImportDataFile$ = tfilename$
Open ImportDataFile$ For Input As #ImportDataFileNumber%

' Read first line of data
CalcZAFLineCount& = 0
CalcZAFSampleCount% = 0
Call CalcZAFImportNext
If ierror Then Exit Sub

' Update filename on display
FormMAIN.Caption = "CalcZAF (Calculate ZAF and Phi-Rho-Z Corrections)    [" & ImportDataFile$ & "]"
FormMAIN.menuFileUpdateCalcZAFSampleDataFiles.Enabled = False   ' in case a sample data file is open

' Set enables
Call CalcZAFSetEnables
If ierror Then Exit Sub
Exit Sub

' Errors
CalcZAFImportOpenError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFImportOpen"
ierror = True
Exit Sub

End Sub

Sub CalcZAFInit()
' Init module level variables for the sample calculation

ierror = False
On Error GoTo CalcZAFInitError

Dim i As Integer

For i% = 1 To MAXCHAN%
UnkCounts!(i%) = 0#
StdCounts!(i%) = 0#
Next i%

' Initialize sample
Call InitSample(CalcZAFOldSample())
If ierror Then Exit Sub

Call InitSample(CalcZAFTmpSample())
If ierror Then Exit Sub

Call InitSample(CalcZAFNewSample())
If ierror Then Exit Sub

' Initialize standards
Call InitStandards(CalcZAFAnalysis)
If ierror Then Exit Sub

' Initialize analysis
Call InitLine(CalcZAFAnalysis)
If ierror Then Exit Sub

Exit Sub

' Errors
CalcZAFInitError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFInit"
ierror = True
Exit Sub

End Sub

Sub CalcZAFLoad()
' Loads the current ZAF element setup

ierror = False
On Error GoTo CalcZAFLoadError

Dim chan As Integer

' Load default "TakeOff" and "KiloVolts"
CalcZAFOldSample(1).takeoff! = DefaultTakeOff!
CalcZAFOldSample(1).kilovolts! = DefaultKiloVolts!

' Load default "TakeOff" and "KiloVolts" to all elements
For chan% = 1 To CalcZAFOldSample(1).LastElm%
CalcZAFOldSample(1).TakeoffArray!(chan%) = DefaultTakeOff!
CalcZAFOldSample(1).KilovoltsArray!(chan%) = DefaultKiloVolts!
Next chan%

' Load element list
Call CalcZAFLoadList
If ierror Then Exit Sub

' Set enables
Call CalcZAFSetEnables
If ierror Then Exit Sub

' Make extended menus visible if flag is set
If ExtendedMenuFlag Then
FormMAIN.menuXraySeparator10.Visible = True
FormMAIN.menuXraySeparator7.Visible = True
FormMAIN.menuXrayUpdateXLineTable.Visible = True
FormMAIN.menuXrayUpdateXEdgeTable.Visible = True
FormMAIN.menuXrayUpdateXFlurTable.Visible = True

FormMAIN.menuAnalyticalSeparator4.Visible = True
FormMAIN.menuAnalyticalSeparator6.Visible = True
FormMAIN.menuAnalyticalSeparator5.Visible = True

FormMAIN.menuAnalyticalCalculateBinaryIntensities1.Visible = True
FormMAIN.menuAnalyticalCalculateBinaryIntensities2.Visible = True
FormMAIN.menuAnalyticalCalculateBinaryIntensities3.Visible = True
FormMAIN.menuAnalyticalCalculateFirstApproximations1.Visible = True
FormMAIN.menuAnalyticalCalculateFirstApproximations2.Visible = True
FormMAIN.menuAnalyticalCalculateFirstApproximations3.Visible = True

FormMAIN.menuXrayUpdateEdgeLineFlurFiles.Visible = True
FormMAIN.menuXrayConvertTextToData.Visible = True
FormMAIN.menuXrayConvertDataToText.Visible = True

FormBINARY.CheckFirstApproximationApplyAbsorption.Enabled = True
FormBINARY.CheckFirstApproximationApplyFluorescence.Enabled = True
FormBINARY.CheckFirstApproximationApplyAtomicNumber.Enabled = True
End If

NumberofSamples% = 1    ' fake
FormZAF.CommandCalculate.BackColor = vbYellow

' Load default correction type (load last because of click event weirdness)
FormZAF.OptionCalculate(CalcZAFMode%).Value = True

Exit Sub

' Errors
CalcZAFLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFLoad"
ierror = True
Exit Sub

End Sub

Sub CalcZAFLoadList()
' Loads the current ZAF element setup

ierror = False
On Error GoTo CalcZAFLoadListError

Dim i As Integer
Dim tmsg As String

' Load the frame caption
tmsg$ = "Element List (click element row to edit)"
If Not CalcZAFOldSample(1).CombinedConditionsFlag Then
tmsg$ = tmsg$ & ", calculations based on " & Str$(CalcZAFOldSample(1).takeoff!) & " degrees (TO)"
tmsg$ = tmsg$ & " and " & Str$(CalcZAFOldSample(1).kilovolts!) & " KeV"
End If
FormZAF.FrameElementList.Caption = tmsg$

' Blank the element grid
FormZAF.GridElementList.Clear

' Initialize the Element List Grid Width
For i% = 0 To FormZAF.GridElementList.cols - 1
FormZAF.GridElementList.ColWidth(i%) = (FormZAF.GridElementList.Width - SCROLLBARWIDTH%) / FormZAF.GridElementList.cols - 1
Next i%

' Load the Grid Column labels
FormZAF.GridElementList.row = 0
FormZAF.GridElementList.col = 0
FormZAF.GridElementList.Text = "Element"
FormZAF.GridElementList.col = 1
FormZAF.GridElementList.Text = "Analyzed"

' Cations assignments
FormZAF.GridElementList.col = 2
FormZAF.GridElementList.Text = "Cations"

' Standard assignments
FormZAF.GridElementList.col = 3
FormZAF.GridElementList.Text = "Standard"
FormZAF.GridElementList.col = 4
FormZAF.GridElementList.Text = "Std K-fac."
FormZAF.GridElementList.col = 5
FormZAF.GridElementList.Text = "Std Inten."

' ZAF (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
FormZAF.GridElementList.col = 6
FormZAF.GridElementList.Text = "Unk Wt. %"
FormZAF.GridElementList.col = 7
If CalcZAFMode% > 0 Then
FormZAF.GridElementList.Text = "Unk Inten."
Else
If CorrectionFlag% = 0 Or CorrectionFlag% = MAXCORRECTION% Then
FormZAF.GridElementList.Text = "Unk K-fac."
Else
FormZAF.GridElementList.Text = "Unk B-fac."
End If
End If

' Load element grid
For i% = 1 To CalcZAFOldSample(1).LastChan%
FormZAF.GridElementList.row = i%
FormZAF.GridElementList.col = 0
FormZAF.GridElementList.Text = CalcZAFOldSample(1).Elsyms$(i%) & " " & CalcZAFOldSample(1).Xrsyms$(i%)

' Load analyzed/specified string
FormZAF.GridElementList.col = 1
FormZAF.GridElementList.Text = vbNullString
If Trim$(CalcZAFOldSample(1).Elsyms$(i%)) <> vbNullString Then
FormZAF.GridElementList.Text = "No"
If CalcZAFOldSample(1).Xrsyms$(i%) = "ka" Or CalcZAFOldSample(1).Xrsyms$(i%) = "la" Or CalcZAFOldSample(1).Xrsyms$(i%) = "ma" Then
FormZAF.GridElementList.Text = "Yes"
End If
End If

' Cation assignments
FormZAF.GridElementList.col = 2
FormZAF.GridElementList.Text = Format$(CalcZAFOldSample(1).numcat%(i%)) & "/" & Format$(CalcZAFOldSample(1).numoxd%(i%))

' Standard assignments
FormZAF.GridElementList.col = 3
If CalcZAFOldSample(1).StdAssigns%(i%) <> 0 And CalcZAFOldSample(1).StdAssigns%(i%) <> MAXINTEGER% Then
FormZAF.GridElementList.Text = Format$(CalcZAFOldSample(1).StdAssigns%(i%))
Else
FormZAF.GridElementList.Text = Format$("-----", a80$)
End If

' Standard data
FormZAF.GridElementList.col = 4
FormZAF.GridElementList.Text = MiscAutoFormat$(CalcZAFAnalysis.StdAssignsKfactors!(i%))
FormZAF.GridElementList.col = 5
FormZAF.GridElementList.Text = MiscAutoFormat$(StdCounts!(i%))

' Weight input or results
FormZAF.GridElementList.col = 6
If CalcZAFMode% = 0 Then
FormZAF.GridElementList.Text = Format$(Format$(CalcZAFOldSample(1).ElmPercents!(i%), f83$), a80$)
Else
FormZAF.GridElementList.Text = Format$(Format$(CalcZAFAnalysis.WtPercents!(i%), f83$), a80$)
End If

' Unknown data
FormZAF.GridElementList.col = 7
If CalcZAFMode% = 0 Then
FormZAF.GridElementList.Text = MiscAutoFormat$(CalcZAFAnalysis.UnkKrats!(i%))
Else
FormZAF.GridElementList.Text = MiscAutoFormat$(UnkCounts!(i%))
End If
Next i%

Exit Sub

' Errors
CalcZAFLoadListError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFLoadList"
ierror = True
Exit Sub

End Sub

Sub CalcZAFSave()
' Sort the modified sample

ierror = False
On Error GoTo CalcZAFSaveError

Dim i As Integer, ip As Integer
Dim ipp As Integer, ippp As Integer
Dim sym As String

ReDim tunkcounts(1 To MAXCHAN%) As Single
ReDim tstdcounts(1 To MAXCHAN%) As Single

' Load name, description, etc
CalcZAFNewSample(1) = CalcZAFOldSample(1)

' Load counts
For i% = 1 To MAXCHAN%
tunkcounts!(i%) = UnkCounts!(i%)
tstdcounts!(i%) = StdCounts!(i%)
UnkCounts!(i%) = 0#
StdCounts!(i%) = 0#
Next i%

' Zero elements
CalcZAFNewSample(1).LastElm% = 0  ' analyzed
CalcZAFNewSample(1).LastChan% = 0 ' analyzed and specified

' Load the analyzed elements first
For i% = 1 To MAXCHAN%

' Initialize the element channel
Call InitElement(i%, CalcZAFNewSample())
If ierror Then Exit Sub

' Find element and xray symbol
sym$ = CalcZAFOldSample(1).Elsyms$(i%)
ip% = IPOS1(MAXELM%, sym$, Symlo$())

sym$ = CalcZAFOldSample(1).Xrsyms$(i%)
ipp% = IPOS1(MAXRAY%, sym, Xraylo$())

' Skip if element if not analyzed
If ip% = 0 Or ipp% = 0 Or ipp% > MAXRAY% - 1 Then GoTo 2000

' Check for element already loaded as an analyzed element
ippp% = IPOS5(Int(1), i%, CalcZAFOldSample(), CalcZAFNewSample())
If ippp% > 0 Then
msg$ = "Element " & CalcZAFOldSample(1).Elsyms$(i%) & " " & CalcZAFOldSample(1).Xrsyms$(i%) & " is already present as an analyzed element, it will be skipped"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFSave"
GoTo 2000
End If

' Increment number of analyzed elements
CalcZAFNewSample(1).LastElm% = CalcZAFNewSample(1).LastElm% + 1
CalcZAFNewSample(1).Elsyms$(CalcZAFNewSample(1).LastElm%) = CalcZAFOldSample(1).Elsyms$(i%)
CalcZAFNewSample(1).Xrsyms$(CalcZAFNewSample(1).LastElm%) = CalcZAFOldSample(1).Xrsyms$(i%)
    
' Make sure cations and oxygens are loaded
If CalcZAFOldSample(1).numcat%(i%) = 0 Or (CalcZAFOldSample(1).numcat%(i%) = 0 And CalcZAFOldSample(1).numoxd%(i%) = 0) Then
CalcZAFOldSample(1).numcat%(i%) = AllCat%(ip%)
CalcZAFOldSample(1).numoxd%(i%) = AllOxd%(ip%)
End If
CalcZAFNewSample(1).numcat%(CalcZAFNewSample(1).LastElm%) = CalcZAFOldSample(1).numcat%(i%)
CalcZAFNewSample(1).numoxd%(CalcZAFNewSample(1).LastElm%) = CalcZAFOldSample(1).numoxd%(i%)

CalcZAFNewSample(1).ElmPercents!(CalcZAFNewSample(1).LastElm%) = CalcZAFOldSample(1).ElmPercents!(i%)
CalcZAFNewSample(1).StdAssigns%(CalcZAFNewSample(1).LastElm%) = CalcZAFOldSample(1).StdAssigns%(i%)

CalcZAFNewSample(1).Elsyup$(CalcZAFNewSample(1).LastElm%) = CalcZAFOldSample(1).Elsyup$(i%)
CalcZAFNewSample(1).Oxsyup$(CalcZAFNewSample(1).LastElm%) = CalcZAFOldSample(1).Oxsyup$(i%)

' Save combined conditions
CalcZAFNewSample(1).TakeoffArray!(CalcZAFNewSample(1).LastElm%) = CalcZAFOldSample(1).TakeoffArray!(i%)
CalcZAFNewSample(1).KilovoltsArray!(CalcZAFNewSample(1).LastElm%) = CalcZAFOldSample(1).KilovoltsArray!(i%)
CalcZAFNewSample(1).BeamCurrentArray!(CalcZAFNewSample(1).LastElm%) = CalcZAFOldSample(1).BeamCurrentArray!(i%)
CalcZAFNewSample(1).BeamSizeArray!(CalcZAFNewSample(1).LastElm%) = CalcZAFOldSample(1).BeamSizeArray!(i%)

If CalcZAFNewSample(1).TakeoffArray!(CalcZAFNewSample(1).LastElm%) = 0# Then CalcZAFNewSample(1).TakeoffArray!(CalcZAFNewSample(1).LastElm%) = CalcZAFNewSample(1).takeoff!
If CalcZAFNewSample(1).KilovoltsArray!(CalcZAFNewSample(1).LastElm%) = 0# Then CalcZAFNewSample(1).KilovoltsArray!(CalcZAFNewSample(1).LastElm%) = CalcZAFNewSample(1).kilovolts!
If CalcZAFNewSample(1).BeamCurrentArray!(CalcZAFNewSample(1).LastElm%) = 0# Then CalcZAFNewSample(1).BeamCurrentArray!(CalcZAFNewSample(1).LastElm%) = DefaultBeamCurrent!
If CalcZAFNewSample(1).BeamSizeArray!(CalcZAFNewSample(1).LastElm%) = 0# Then CalcZAFNewSample(1).BeamSizeArray!(CalcZAFNewSample(1).LastElm%) = DefaultBeamSize!

' Save counts also
UnkCounts!(CalcZAFNewSample(1).LastElm%) = tunkcounts!(i%)
StdCounts!(CalcZAFNewSample(1).LastElm%) = tstdcounts!(i%)

2000:  Next i%

' Update number of analyzed elements
CalcZAFNewSample(1).LastChan% = CalcZAFNewSample(1).LastElm%

' Load the specified elements next, set x-ray, etc. to blank
For i% = 1 To MAXCHAN%
sym$ = CalcZAFOldSample(1).Elsyms$(i%)
ip% = IPOS1(MAXELM%, sym$, Symlo$())

sym$ = CalcZAFOldSample(1).Xrsyms$(i%)
ipp% = IPOS1(MAXRAY%, sym$, Xraylo$())

' Skip if element is analyzed
If ip% = 0 Or ipp% = 0 Or ipp% <= MAXRAY% - 1 Then GoTo 3000

' Check for element already analyzed or specified and skip if found
ippp% = IPOS1(CalcZAFNewSample(1).LastChan%, CalcZAFOldSample(1).Elsyms$(i%), CalcZAFNewSample(1).Elsyms$())
If ippp% > 0 Then GoTo 3000

' Increment number of specified elements
CalcZAFNewSample(1).LastChan% = CalcZAFNewSample(1).LastChan% + 1

' Load specified element parameters
CalcZAFNewSample(1).Elsyms$(CalcZAFNewSample(1).LastChan%) = CalcZAFOldSample(1).Elsyms$(i%)
CalcZAFNewSample(1).Xrsyms$(CalcZAFNewSample(1).LastChan%) = vbNullString
    
' Make sure cations are loaded
If CalcZAFOldSample(1).numcat%(i%) = 0 Then CalcZAFOldSample(1).numcat%(i%) = AllCat%(ip%)
If CalcZAFOldSample(1).numoxd%(i%) = 0 Then CalcZAFOldSample(1).numoxd%(i%) = AllOxd%(ip%)
CalcZAFNewSample(1).numcat%(CalcZAFNewSample(1).LastChan%) = CalcZAFOldSample(1).numcat%(i%)
CalcZAFNewSample(1).numoxd%(CalcZAFNewSample(1).LastChan%) = CalcZAFOldSample(1).numoxd%(i%)

CalcZAFNewSample(1).ElmPercents!(CalcZAFNewSample(1).LastChan%) = CalcZAFOldSample(1).ElmPercents!(i%)

CalcZAFNewSample(1).Elsyup$(CalcZAFNewSample(1).LastChan%) = CalcZAFOldSample(1).Elsyup$(i%)
CalcZAFNewSample(1).Oxsyup$(CalcZAFNewSample(1).LastChan%) = CalcZAFOldSample(1).Oxsyup$(i%)

3000:  Next i%

' Check for analyzed oxygen if oxide sample
If CalcZAFNewSample(1).OxideOrElemental% = 1 Then
ip% = IPOS1(CalcZAFNewSample(1).LastElm%, Symlo$(ATOMIC_NUM_OXYGEN%), CalcZAFNewSample(1).Elsyms$())
If ip% > 0 Then GoTo CalcZAFSaveOxygenOnOxide
End If

' Add specified oxygen if necessary
Call UpdateStdElements(CalcZAFAnalysis, CalcZAFNewSample(), CalcZAFTmpSample())
If ierror Then Exit Sub

' No errors, re-load modified sample
CalcZAFOldSample(1) = CalcZAFNewSample(1)

Exit Sub

' Errors
CalcZAFSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFSave"
ierror = True
Exit Sub

CalcZAFSaveOxygenOnOxide:
msg$ = "Cannot calculate stoichiometric oxygen if oxygen if an analyzed element"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFSave"
ierror = True
Exit Sub

End Sub

Sub CalcZAFTypeAnalysis()
' Type last analysis parameters

ierror = False
On Error GoTo CalcZAFTypeAnalysisError

' Type last analysis parameters
Call CalcZAFTypeAnalysis2(CalcZAFAnalysis, CalcZAFOldSample())
If ierror Then Exit Sub

Exit Sub

' Errors
CalcZAFTypeAnalysisError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFTypeAnalysis"
ierror = True
Exit Sub

End Sub

Sub CalcZAFTypeStandards()
' Type composition of all standards

ierror = False
On Error GoTo CalcZAFTypeStandardsError

' Type them out
Call TypeStandards2(CalcZAFOldSample(), CalcZAFTmpSample())
If ierror Then Exit Sub

Exit Sub

' Errors
CalcZAFTypeStandardsError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFTypeStandards"
ierror = True
Exit Sub

End Sub

Sub CalcZAFUpdateAllStdKfacs()
' Update all standard k-factors

ierror = False
On Error GoTo CalcZAFUpdateAllStdKfacsError

' Type out ZAF selections
Call TypeZAFSelections
If ierror Then Exit Sub

' Load element arrays
If CalcZAFOldSample(1).LastElm% > 0 Then
Call ElementGetData(CalcZAFOldSample())
If ierror Then Exit Sub

' Initialize calculations (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% <> MAXCORRECTION% Then
Call ZAFSetZAF(CalcZAFOldSample())
If ierror Then Exit Sub
Else
'Call ZAFSetZAF3(CalcZAFOldSample())
'If ierror Then Exit Sub
End If

' Type out PTC and coating selections
Call TypeSampleFlags(CalcZAFAnalysis, CalcZAFOldSample())
If ierror Then Exit Sub
End If

' Calculate standard k-factors
Call UpdateAllStdKfacs(CalcZAFAnalysis, CalcZAFOldSample(), CalcZAFTmpSample())
If ierror Then Exit Sub

Exit Sub

' Errors
CalcZAFUpdateAllStdKfacsError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFUpdateAllStdKfacs"
ierror = True
Exit Sub

End Sub

Sub CalcZAFBinary(mode As Integer, tForm As Form)
' Calculate binary data from .DAT file
' mode=0 normal k-ratio output (the only mode that includes support for alpha factor calculations)
' mode=1 first approximation atomic fraction output
' mode=2 first approximation mass fraction output
' mode=3 first approximation electron fraction output
' mode=4 loop on all correction options (izaf% = 1 to 10, MACFileType% = 1 to MAXMACTYPE%, single line, single file)
' mode=5 loop on all correction options (izaf% = 1 to 10, MACFileType% = 1 to MAXMACTYPE%, all lines, multiple files)
'
' Data file format assumes one line for each binary. The first two
' columns are the atomic numbers of the two binary components
' to be calculated. The second two columns are the xray lines to use.
' ( 1 = Ka, 2 = Kb, 3 = La, 4 = Lb, 5 = Ma, 6 = Mb, 7 = Ln, 8 = Lg,
' 9 = Lv, 10 = Ll, 11 = Mg, 12 = Mz, 13 = by difference). The next
' two columns are the operating voltage and take-off angle. The next
' two columns are the wt. fractions of the binary components. The
' last two columns contains the k-exp values for calculation of k-calc/k-exp.
'
'       79     29     5    13    15.     52.5    .8015   .1983   .7400   .0
'       79     29     5    13    15.     52.5    .6036   .3964   .5110   .0
'       79     29     5    13    15.     52.5    .4010   .5992   .3120   .0
'       79     29     5    13    15.     52.5    .2012   .7985   .1450   .0

ierror = False
On Error GoTo CalcZAFBinaryError

Dim tfilename As String, astring As String
Dim eO As Single, TOA As Single
Dim eng1 As Single, eng2 As Single
Dim edg1 As Single, edg2 As Single
Dim masszbar As Single, zedzbar As Single
Dim ii As Long
Dim n As Integer, nn As Integer
Dim m As Integer, mm As Integer
Dim tfilenumber As Integer

Dim average As TypeAverage

Static lastfilename As String

ReDim isym(1 To 2) As Integer
ReDim iray(1 To 2) As Integer
ReDim conc(1 To 2) As Single
ReDim conc1(1 To 2) As Single
ReDim kexp(1 To 2) As Single
ReDim temp(1 To 2) As Single

' Show form
FormZAF.Show vbModeless
icancelauto = False

' Check for ZAF/Phi-Rho-Z corrections with modes 1, 2, or 3 (first approximation calculations) or 4 or 5 (multiple ZAFs and MACs)
If mode% <> 0 And (CorrectionFlag% <> 0 Or CorrectionFlag% <> MAXCORRECTION%) Then
msg$ = "ZAF or Phi-Rho-Z corrections are not selected. Changing matrix correction type to default ZAF or Phi-Rho-Z for multiple ZAF/MAC calculations."
MsgBox msg$, vbOKOnly + vbInformation, "CalcZAFBinary"
CorrectionFlag% = 0
End If

' Get import filename from user
If lastfilename$ = vbNullString Then lastfilename$ = "pouchouz10.dat"
tfilename$ = lastfilename$
Call IOGetFileName(Int(2), "DAT", tfilename$, tForm)
If ierror Then Exit Sub

' Save current ZAF and MAC selection
tzaftype% = izaf%
tmactype% = MACTypeFlag%

' Save current path
CalcZAFDATFileDirectory$ = CurDir$

' No errors, save file name
lastfilename$ = tfilename$
ImportDataFile2$ = lastfilename$

' Get export filename from user
tfilename$ = MiscGetFileNameNoExtension(tfilename$) & ".out"
If mode% <> 5 Then
Call IOGetFileName(Int(1), "OUT", tfilename$, tForm)
If ierror Then Exit Sub
End If

' No errors, save file name
ExportDataFile$ = tfilename$
HistogramDataFile$ = MiscGetFileNameNoExtension(tfilename$) & ".txt"

Call CalcZAFSetEnables
If ierror Then Exit Sub

' Delete all previous files if mode% = 5 and write column headings
If mode% = 5 Then
nn% = MAXZAF%  ' loop on all correction options
mm% = MAXMACTYPE%   ' loop on all MAC files
For n% = 1 To nn%
For m% = 1 To mm%

' Delete file
tfilenumber% = CalcZAFBinaryFile%(Int(1), n%, m%)
If ierror Then Exit Sub

' Open each file
tfilenumber% = CalcZAFBinaryFile%(Int(2), n%, m%)  ' open for APPEND
If ierror Then Exit Sub

' Write column headings for each (multiple) file
Print #tfilenumber%, " ", VbDquote$ & "Line" & VbDquote$, vbTab, VbDquote$ & "SymA" & VbDquote$, vbTab, VbDquote$ & "SymB" & VbDquote$, vbTab, VbDquote$ & "RayA" & VbDquote$, vbTab, VbDquote$ & "RayB" & VbDquote$, vbTab, _
    VbDquote$ & "KeV" & VbDquote$, vbTab, VbDquote$ & "Takeoff" & VbDquote$, vbTab, VbDquote$ & "ConcA" & VbDquote$, vbTab, VbDquote$ & "ConcB" & VbDquote$, vbTab, _
    VbDquote$ & "KexpA" & VbDquote$, vbTab, VbDquote$ & "KexpB" & VbDquote$, vbTab, VbDquote$ & "KratA" & VbDquote$, vbTab, VbDquote$ & "KratB" & VbDquote$, vbTab, _
    VbDquote$ & "Pri F(Chi)A" & VbDquote$, vbTab, VbDquote$ & "Sec F(Chi)A" & VbDquote$, vbTab, VbDquote$ & "AbsA" & VbDquote$, vbTab, VbDquote$ & "FluA" & VbDquote$, vbTab, _
    VbDquote$ & "ZedA" & VbDquote$, vbTab, VbDquote$ & "StpA" & VbDquote$, vbTab, VbDquote$ & "BksA" & VbDquote$, vbTab, VbDquote$ & "ZAFA" & VbDquote$, vbTab, _
    VbDquote$ & "Pri F(Chi)B" & VbDquote$, vbTab, VbDquote$ & "Sec F(Chi)B" & VbDquote$, vbTab, VbDquote$ & "AbsB" & VbDquote$, vbTab, VbDquote$ & "FluB" & VbDquote$, vbTab, _
    VbDquote$ & "ZedB" & VbDquote$, vbTab, VbDquote$ & "StpB" & VbDquote$, vbTab, VbDquote$ & "BksB" & VbDquote$, vbTab, VbDquote$ & "ZAFB" & VbDquote$, vbTab, _
    VbDquote$ & "KerrA" & VbDquote$, vbTab, VbDquote$ & "KerrB" & VbDquote$

' Close each file
tfilenumber% = CalcZAFBinaryFile%(Int(3), n%, m%)
If ierror Then Exit Sub

Next m%
Next n%
End If

' Open normal files
Open ImportDataFile2$ For Input As #ImportDataFileNumber2%
Open ExportDataFile$ For Output As #ExportDataFileNumber2%
CalcZAFLineCount& = 0
CalcZAFOutputCount& = 0
CalcZAFLineCount& = 0
CalcZAFOutputCount& = 0
Call IOStatusAuto(vbNullString)

' Output column labels (all except mode% = 5)
If mode% = 0 Or mode = 4 Then
    If CorrectionFlag% = 0 Or CorrectionFlag% = MAXCORRECTION% Then     ' ZAF/Phi-rho-z calculations (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
    Print #ExportDataFileNumber2%, " ", VbDquote$ & "Line" & VbDquote$, vbTab, VbDquote$ & "SymA" & VbDquote$, vbTab, VbDquote$ & "SymB" & VbDquote$, vbTab, VbDquote$ & "RayA" & VbDquote$, vbTab, VbDquote$ & "RayB" & VbDquote$, vbTab, _
    VbDquote$ & "KeV" & VbDquote$, vbTab, VbDquote$ & "Takeoff" & VbDquote$, vbTab, VbDquote$ & "ConcA" & VbDquote$, vbTab, VbDquote$ & "ConcB" & VbDquote$, vbTab, _
    VbDquote$ & "KexpA" & VbDquote$, vbTab, VbDquote$ & "KexpB" & VbDquote$, vbTab, VbDquote$ & "KratA" & VbDquote$, vbTab, VbDquote$ & "KratB" & VbDquote$, vbTab, _
    VbDquote$ & "Pri F(Chi)A" & VbDquote$, vbTab, VbDquote$ & "Sec F(Chi)A" & VbDquote$, vbTab, VbDquote$ & "AbsA" & VbDquote$, vbTab, VbDquote$ & "FluA" & VbDquote$, vbTab, _
    VbDquote$ & "ZedA" & VbDquote$, vbTab, VbDquote$ & "StpA" & VbDquote$, vbTab, VbDquote$ & "BksA" & VbDquote$, vbTab, VbDquote$ & "ZAFA" & VbDquote$, vbTab, _
    VbDquote$ & "Pri F(Chi)B" & VbDquote$, vbTab, VbDquote$ & "Sec F(Chi)B" & VbDquote$, vbTab, VbDquote$ & "AbsB" & VbDquote$, vbTab, VbDquote$ & "FluB" & VbDquote$, vbTab, _
    VbDquote$ & "ZedB" & VbDquote$, vbTab, VbDquote$ & "StpB" & VbDquote$, vbTab, VbDquote$ & "BksB" & VbDquote$, vbTab, VbDquote$ & "ZAFB" & VbDquote$, vbTab, _
    VbDquote$ & "KerrA" & VbDquote$, vbTab, VbDquote$ & "KerrB" & VbDquote$
    Else                            ' alpha factor calculations (mode% = 0 only)
    Print #ExportDataFileNumber2%, " ", VbDquote$ & "Line" & VbDquote$, vbTab, VbDquote$ & "SymA" & VbDquote$, vbTab, VbDquote$ & "SymB" & VbDquote$, vbTab, VbDquote$ & "RayA" & VbDquote$, vbTab, VbDquote$ & "RayB" & VbDquote$, vbTab, _
    VbDquote$ & "KeV" & VbDquote$, vbTab, VbDquote$ & "Takeoff" & VbDquote$, vbTab, VbDquote$ & "ConcA" & VbDquote$, vbTab, VbDquote$ & "ConcB" & VbDquote$, vbTab, _
    VbDquote$ & "KexpA" & VbDquote$, vbTab, VbDquote$ & "KexpB" & VbDquote$, vbTab, VbDquote$ & "KratA" & VbDquote$, vbTab, VbDquote$ & "KratB" & VbDquote$, vbTab, _
    VbDquote$ & "KerrA" & VbDquote$, vbTab, VbDquote$ & "KerrB" & VbDquote$
    End If
    
ElseIf mode% = 1 Or mode% = 2 Or mode% = 3 Then
Print #ExportDataFileNumber2%, " ", VbDquote$ & "Line" & VbDquote$, vbTab, VbDquote$ & "SymA" & VbDquote$, vbTab, VbDquote$ & "SymB" & VbDquote$, vbTab, VbDquote$ & "RayA" & VbDquote$, vbTab, VbDquote$ & "RayB" & VbDquote$, vbTab, _
    VbDquote$ & "KeV" & VbDquote$, vbTab, VbDquote$ & "Takeoff" & VbDquote$, vbTab, VbDquote$ & "ConcA" & VbDquote$, vbTab, VbDquote$ & "ConcB" & VbDquote$, vbTab, _
    VbDquote$ & "KexpA" & VbDquote$, vbTab, VbDquote$ & "KexpB" & VbDquote$, vbTab, VbDquote$ & "Conc1A" & VbDquote$, vbTab, VbDquote$ & "Conc1B" & VbDquote$, vbTab, _
    VbDquote$ & "CerrA" & VbDquote$, vbTab, VbDquote$ & "CerrB" & VbDquote$
End If

' Check for end of file
Do While Not EOF(ImportDataFileNumber2%)
CalcZAFLineCount& = CalcZAFLineCount& + 1

' Check for Pause button
Do Until Not RealTimePauseAutomation
DoEvents
Sleep 200
Loop

msg$ = "Calculating binary " & Str$(CalcZAFLineCount&) & "..."
Call IOStatusAuto(msg$)
If icancelauto Then
Call IOStatusAuto(vbNullString)
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
ierror = True
Exit Sub
End If

' Initialize
Call CalcZAFInit
If ierror Then Exit Sub

CalcZAFMode% = 0    ' calculate intensities from concentrations
CalcZAFOldSample(1).number% = MAXINTEGER%   ' fake standard assignment
CalcZAFOldSample(1).Description$ = "Line " & Format$(CalcZAFLineCount&)
CalcZAFOldSample(1).OxideOrElemental% = 2
CalcZAFOldSample(1).LastElm% = 2
CalcZAFOldSample(1).LastChan% = 2

CalcZAFOldSample(1).numcat%(1) = 1
CalcZAFOldSample(1).numoxd%(1) = 0
CalcZAFOldSample(1).numcat%(2) = 1
CalcZAFOldSample(1).numoxd%(2) = 0

CalcZAFOldSample(1).StdAssigns%(1) = MAXINTEGER%     '  for proper loading of parameters
CalcZAFOldSample(1).StdAssigns%(2) = MAXINTEGER%     '

' Read binary elements, kilovolts and takeoff
Input #ImportDataFileNumber2%, isym%(1), isym%(2), iray%(1), iray%(2), eO!, TOA!, conc!(1), conc!(2), kexp!(1), kexp!(2)

' Check limits
If isym%(1) < 1 Or isym%(1) > MAXELM% Then GoTo CalcZAFBinaryOutofLimits
If isym%(2) < 1 Or isym%(2) > MAXELM% Then GoTo CalcZAFBinaryOutofLimits
If iray%(1) < 1 Or iray%(1) > MAXRAY% Then GoTo CalcZAFBinaryOutofLimits
If iray%(2) < 1 Or iray%(2) > MAXRAY% Then GoTo CalcZAFBinaryOutofLimits
If eO! < 1# Or eO! > 100# Then GoTo CalcZAFBinaryOutofLimits
If TOA! < 1# Or TOA! > 90# Then GoTo CalcZAFBinaryOutofLimits
If conc!(1) < 0# Or conc!(1) > 1# Then GoTo CalcZAFBinaryOutofLimits
If conc!(2) < 0# Or conc!(2) > 1# Then GoTo CalcZAFBinaryOutofLimits
If kexp!(1) < 0# Or kexp!(1) > 2# Then GoTo CalcZAFBinaryOutofLimits    ' use 2.0 for extreme fluorescence cases
If kexp!(2) < 0# Or kexp!(2) > 2# Then GoTo CalcZAFBinaryOutofLimits    ' use 2.0 for extreme fluorescence cases

' Check that both elements are not by difference
If iray%(1) = MAXRAY% And iray%(2) = MAXRAY% Then GoTo CalcZAFBinaryBothByDifference

' Check that at least one concentration is entered
If conc!(1) = 0# And conc!(2) = 0# Then GoTo CalcZAFBinaryNoConcData

' Check for valid kexp data if x-ray used
If iray%(1) <= MAXRAY% - 1 And kexp!(1) = 0# Then GoTo CalcZAFBinaryNoKexpData
If iray%(2) <= MAXRAY% - 1 And kexp!(2) = 0# Then GoTo CalcZAFBinaryNoKexpData

' Load sample
CalcZAFOldSample(1).Elsyms$(1) = Symlo$(isym%(1))
CalcZAFOldSample(1).Elsyms$(2) = Symlo$(isym%(2))

If iray%(1) > MAXRAY% - 1 Then
CalcZAFOldSample(1).Xrsyms$(1) = vbNullString
Else
CalcZAFOldSample(1).Xrsyms$(1) = Xraylo$(iray%(1))
End If

If iray%(2) > MAXRAY% - 1 Then
CalcZAFOldSample(1).Xrsyms$(2) = vbNullString
Else
CalcZAFOldSample(1).Xrsyms$(2) = Xraylo$(iray%(2))
End If

CalcZAFOldSample(1).kilovolts! = eO!
CalcZAFOldSample(1).takeoff! = TOA!

' Calculate elements by difference
If iray%(1) = MAXRAY% Then
conc!(1) = 1# - conc!(2)
End If

If iray%(2) = MAXRAY% Then
conc!(2) = 1# - conc!(1)
End If

' Load concentrations
CalcZAFOldSample(1).ElmPercents!(1) = conc!(1) * 100#
CalcZAFOldSample(1).ElmPercents!(2) = conc!(2) * 100#

' Update defaults
DefaultTakeOff! = CalcZAFOldSample(1).takeoff!
DefaultKiloVolts! = CalcZAFOldSample(1).kilovolts!

' Sort elements
Call CalcZAFSave
If ierror Then
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
Exit Sub
End If

' Load form
Call CalcZAFLoad
If ierror Then
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
Exit Sub
End If

' Load next matrix correction
nn% = 1
mm% = 1
If mode% = 4 Or mode% = 5 Then
nn% = MAXZAF%  ' loop on all correction options
mm% = MAXMACTYPE%   ' loop on all MAC files
End If
For n% = 1 To nn%
For m% = 1 To mm%

' Set ZAF and MAC if looping on all
If mode% = 4 Or mode% = 5 Then
izaf% = n%

' Check for MAC file
Call GetZAFAllSaveMAC2(m%)
If ierror Then
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
Exit Sub
End If
MACTypeFlag% = m%   ' set after check for exist

' Set ZAF corrections
Call InitGetZAFSetZAF2(izaf%)
If ierror Then
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
Exit Sub
End If

' Update k-factors and parameters
Call CalcZAFUpdateAllStdKfacs
If ierror Then
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
Exit Sub
End If
End If

If mode% = 4 Or mode% = 5 Then
If mode% = 4 Then msg$ = "Calculating binary with " & zafstring$(izaf%) & ", " & macstring$(MACTypeFlag%)
If mode% = 5 Then msg$ = "Calculating binary " & Str$(CalcZAFLineCount&) & " with " & zafstring$(izaf%) & ", " & macstring$(MACTypeFlag%)
Call IOWriteLog(vbCrLf & vbCrLf & msg$)
Call IOStatusAuto(msg$ & "...")
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
ierror = True
Exit Sub
End If
End If

' Calculate actual binary intensities
Call CalcZAFCalculate
If ierror Then
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
Exit Sub
End If

' Check for minimum output parameters
If mode% = 0 Then
If BinaryOutputRangeMinAbs And Abs(CalcZAFAnalysis.StdZAFCors!(1, 1, 1) - 1#) < BinaryOutputRangeAbsMin! Then GoTo CalcZAFBinarySkip
If BinaryOutputRangeMinFlu And Abs(CalcZAFAnalysis.StdZAFCors!(2, 1, 1) - 1#) < BinaryOutputRangeFluMin! Then GoTo CalcZAFBinarySkip
If BinaryOutputRangeMinZed And Abs(CalcZAFAnalysis.StdZAFCors!(3, 1, 1) - 1#) < BinaryOutputRangeZedMin! Then GoTo CalcZAFBinarySkip

' Check for maximum output parameters
If BinaryOutputRangeMaxAbs And Abs(CalcZAFAnalysis.StdZAFCors!(1, 1, 1) - 1#) > BinaryOutputRangeAbsMax! Then GoTo CalcZAFBinarySkip
If BinaryOutputRangeMaxFlu And Abs(CalcZAFAnalysis.StdZAFCors!(2, 1, 1) - 1#) > BinaryOutputRangeFluMax! Then GoTo CalcZAFBinarySkip
If BinaryOutputRangeMaxZed And Abs(CalcZAFAnalysis.StdZAFCors!(3, 1, 1) - 1#) > BinaryOutputRangeZedMax! Then GoTo CalcZAFBinarySkip
End If

' Calculate electron fractions
If mode% = 1 Or mode% = 2 Or mode% = 3 Then
Call ConvertWeightToElectron(Int(2), CalcZAFAnalysis.AtomicNumbers!(), CalcZAFAnalysis.AtomicWeights!(), conc!(), temp!())
If ierror Then
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
Exit Sub
End If

' Calculate zbars
masszbar! = conc!(1) * CalcZAFAnalysis.AtomicNumbers!(1) + conc!(2) * CalcZAFAnalysis.AtomicNumbers!(2)
zedzbar! = temp!(1) * CalcZAFAnalysis.AtomicNumbers!(1) + temp!(2) * CalcZAFAnalysis.AtomicNumbers!(2)

If DebugMode Then
Call IOWriteLog(vbNullString)
msg$ = "Mass Zbar= " & MiscAutoFormat$(masszbar!)
Call IOWriteLog(msg$)
msg$ = "Electron Zbar= " & MiscAutoFormat$(zedzbar!)
Call IOWriteLog(msg$)
msg$ = "Zbar Percent Difference= " & MiscAutoFormat$((masszbar! - zedzbar!) / masszbar! * 100#)
Call IOWriteLog(msg$)
End If

' Check for minimum or maximum mass/electron fraction difference
If BinaryOutputMinimumZbar And Abs((masszbar! - zedzbar!) / masszbar! * 100#) < BinaryOutputMinimumZbarDiff! Then GoTo CalcZAFBinarySkip
If BinaryOutputMaximumZbar And Abs((masszbar! - zedzbar!) / masszbar! * 100#) > BinaryOutputMaximumZbarDiff! Then GoTo CalcZAFBinarySkip
End If

' Save data
CalcZAFOutputCount& = CalcZAFOutputCount& + 1

ReDim Preserve KratioExpr!(1 To 2, 1 To CalcZAFOutputCount&)
ReDim Preserve KratioCalc!(1 To 2, 1 To CalcZAFOutputCount&)
ReDim Preserve KratioError!(1 To 2, 1 To CalcZAFOutputCount&)

KratioError!(1, CalcZAFOutputCount&) = 0#
KratioError!(2, CalcZAFOutputCount&) = 0#

' Calculate error
If mode% = 0 Or mode% = 4 Or mode% = 5 Then
If CorrectionFlag% = 0 Or CorrectionFlag% = MAXCORRECTION% Then              ' ZAF/Phi-rho-z calculations (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
    If kexp!(1) <> 0# Then KratioError!(1, CalcZAFOutputCount&) = CalcZAFAnalysis.StdAssignsKfactors!(1) / kexp!(1)
    If kexp!(2) <> 0# Then KratioError!(2, CalcZAFOutputCount&) = CalcZAFAnalysis.StdAssignsKfactors!(2) / kexp!(2)
Else                                     ' alpha factor calculations (mode% = 0 only)
    KratioAlpha!(1) = 0#
    KratioAlpha!(2) = 0#
    If CalcZAFAnalysis.StdAssignsBetas!(1) <> 0# Then KratioAlpha!(1) = conc!(1) / CalcZAFAnalysis.StdAssignsBetas!(1)
    If CalcZAFAnalysis.StdAssignsBetas!(2) <> 0# Then KratioAlpha!(1) = conc!(2) / CalcZAFAnalysis.StdAssignsBetas!(2)
    If kexp!(1) <> 0# Then KratioError!(1, CalcZAFOutputCount&) = KratioAlpha!(1) / kexp!(1)
    If kexp!(2) <> 0# Then KratioError!(2, CalcZAFOutputCount&) = KratioAlpha!(2) / kexp!(2)
End If

KratioExpr!(1, CalcZAFOutputCount&) = kexp!(1)
KratioExpr!(2, CalcZAFOutputCount&) = kexp!(2)

KratioCalc!(1, CalcZAFOutputCount&) = KratioError!(1, CalcZAFOutputCount&) * KratioExpr!(1, CalcZAFOutputCount&)
KratioCalc!(2, CalcZAFOutputCount&) = KratioError!(2, CalcZAFOutputCount&) * KratioExpr!(2, CalcZAFOutputCount&)

' Save parameters for problematic output below
ReDim Preserve KratioLine&(1 To 2, 1 To CalcZAFOutputCount&)
ReDim Preserve KratioEsym$(1 To 2, 1 To CalcZAFOutputCount&)
ReDim Preserve KratioXsym$(1 To 2, 1 To CalcZAFOutputCount&)
ReDim Preserve KratioConc!(1 To 2, 1 To CalcZAFOutputCount&)
ReDim Preserve KratioTOAeO!(1 To 2, 1 To CalcZAFOutputCount&)
ReDim Preserve KratioOver!(1 To 2, 1 To CalcZAFOutputCount&)

KratioLine&(1, CalcZAFOutputCount&) = CalcZAFLineCount&
KratioLine&(2, CalcZAFOutputCount&) = CalcZAFLineCount&

KratioEsym$(1, CalcZAFOutputCount&) = Symup$(isym%(1))
KratioEsym$(2, CalcZAFOutputCount&) = Symup$(isym%(2))
KratioXsym$(1, CalcZAFOutputCount&) = Xraylo$(iray%(1))
KratioXsym$(2, CalcZAFOutputCount&) = Xraylo$(iray%(2))

KratioConc!(1, CalcZAFOutputCount&) = conc!(1)
KratioConc!(2, CalcZAFOutputCount&) = conc!(2)

KratioTOAeO!(1, CalcZAFOutputCount&) = TOA!
KratioTOAeO!(2, CalcZAFOutputCount&) = eO!

' Store the overvoltages
KratioOver!(1, CalcZAFOutputCount&) = 0#
If iray%(1) < MAXRAY% Then
Call XrayGetEnergy(isym%(1), iray%(1), eng1!, edg1!)
If ierror Then Exit Sub
If edg1! <> 0# Then KratioOver!(1, CalcZAFOutputCount&) = eO! / edg1!
End If
KratioOver!(2, CalcZAFOutputCount&) = 0#
If iray%(2) < MAXRAY% Then
Call XrayGetEnergy(isym%(2), iray%(2), eng2!, edg2!)
If ierror Then Exit Sub
If edg2! <> 0# Then KratioOver!(2, CalcZAFOutputCount&) = eO! / edg2!
End If

' Output normal binary k-ratio results
If mode% = 0 Then
    If CorrectionFlag% = 0 Or CorrectionFlag% = MAXCORRECTION% Then              ' ZAF/Phi-rho-z calculations (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
    Print #ExportDataFileNumber2%, CalcZAFLineCount&, vbTab, isym%(1), vbTab, isym%(2), vbTab, iray%(1), vbTab, iray%(2), vbTab, _
    eO!, vbTab, TOA!, vbTab, conc!(1), vbTab, conc!(2), vbTab, kexp!(1), vbTab, kexp!(2), vbTab, _
    CalcZAFAnalysis.StdAssignsKfactors!(1), vbTab, CalcZAFAnalysis.StdAssignsKfactors!(2), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(7, 1, 1), vbTab, CalcZAFAnalysis.StdZAFCors!(8, 1, 1), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(1, 1, 1), vbTab, CalcZAFAnalysis.StdZAFCors!(2, 1, 1), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(3, 1, 1), vbTab, CalcZAFAnalysis.StdZAFCors!(5, 1, 1), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(6, 1, 1), vbTab, CalcZAFAnalysis.StdZAFCors!(4, 1, 1), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(7, 1, 2), vbTab, CalcZAFAnalysis.StdZAFCors!(8, 1, 2), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(1, 1, 2), vbTab, CalcZAFAnalysis.StdZAFCors!(2, 1, 2), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(3, 1, 2), vbTab, CalcZAFAnalysis.StdZAFCors!(5, 1, 2), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(6, 1, 2), vbTab, CalcZAFAnalysis.StdZAFCors!(4, 1, 2), vbTab, _
    KratioError!(1, CalcZAFOutputCount&), vbTab, KratioError!(2, CalcZAFOutputCount&)
    Else                                     ' alpha factor calculations (mode% = 0 only)
    Print #ExportDataFileNumber2%, CalcZAFLineCount&, vbTab, isym%(1), vbTab, isym%(2), vbTab, iray%(1), vbTab, iray%(2), vbTab, _
    eO!, vbTab, TOA!, vbTab, conc!(1), vbTab, conc!(2), vbTab, kexp!(1), vbTab, kexp!(2), vbTab, _
    KratioAlpha!(1), vbTab, KratioAlpha!(2), vbTab, _
    KratioError!(1, CalcZAFOutputCount&), vbTab, KratioError!(2, CalcZAFOutputCount&)
    End If
ElseIf mode% = 4 Then
Print #ExportDataFileNumber2%, CalcZAFLineCount&, vbTab, isym%(1), vbTab, isym%(2), vbTab, iray%(1), vbTab, iray%(2), vbTab, _
    eO!, vbTab, TOA!, vbTab, conc!(1), vbTab, conc!(2), vbTab, kexp!(1), vbTab, kexp!(2), vbTab, _
    CalcZAFAnalysis.StdAssignsKfactors!(1), vbTab, CalcZAFAnalysis.StdAssignsKfactors!(2), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(7, 1, 1), vbTab, CalcZAFAnalysis.StdZAFCors!(8, 1, 1), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(1, 1, 1), vbTab, CalcZAFAnalysis.StdZAFCors!(2, 1, 1), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(3, 1, 1), vbTab, CalcZAFAnalysis.StdZAFCors!(5, 1, 1), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(6, 1, 1), vbTab, CalcZAFAnalysis.StdZAFCors!(4, 1, 1), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(7, 1, 2), vbTab, CalcZAFAnalysis.StdZAFCors!(8, 1, 2), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(1, 1, 2), vbTab, CalcZAFAnalysis.StdZAFCors!(2, 1, 2), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(3, 1, 2), vbTab, CalcZAFAnalysis.StdZAFCors!(5, 1, 2), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(6, 1, 2), vbTab, CalcZAFAnalysis.StdZAFCors!(4, 1, 2), vbTab, _
    KratioError!(1, CalcZAFOutputCount&), vbTab, KratioError!(2, CalcZAFOutputCount&), vbTab, _
    VbDquote$ & zafstring$(izaf%) & ", " & macstring$(MACTypeFlag%) & VbDquote$
ElseIf mode% = 5 Then
tfilenumber% = CalcZAFBinaryFile%(Int(2), izaf%, MACTypeFlag%)  ' open for APPEND
Print #tfilenumber%, CalcZAFLineCount&, vbTab, isym%(1), vbTab, isym%(2), vbTab, iray%(1), vbTab, iray%(2), vbTab, _
    eO!, vbTab, TOA!, vbTab, conc!(1), vbTab, conc!(2), vbTab, kexp!(1), vbTab, kexp!(2), vbTab, _
    CalcZAFAnalysis.StdAssignsKfactors!(1), vbTab, CalcZAFAnalysis.StdAssignsKfactors!(2), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(7, 1, 1), vbTab, CalcZAFAnalysis.StdZAFCors!(8, 1, 1), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(1, 1, 1), vbTab, CalcZAFAnalysis.StdZAFCors!(2, 1, 1), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(3, 1, 1), vbTab, CalcZAFAnalysis.StdZAFCors!(5, 1, 1), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(6, 1, 1), vbTab, CalcZAFAnalysis.StdZAFCors!(4, 1, 1), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(7, 1, 2), vbTab, CalcZAFAnalysis.StdZAFCors!(8, 1, 2), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(1, 1, 2), vbTab, CalcZAFAnalysis.StdZAFCors!(2, 1, 2), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(3, 1, 2), vbTab, CalcZAFAnalysis.StdZAFCors!(5, 1, 2), vbTab, _
    CalcZAFAnalysis.StdZAFCors!(6, 1, 2), vbTab, CalcZAFAnalysis.StdZAFCors!(4, 1, 2), vbTab, _
    KratioError!(1, CalcZAFOutputCount&), vbTab, KratioError!(2, CalcZAFOutputCount&)
tfilenumber% = CalcZAFBinaryFile%(Int(3), izaf%, MACTypeFlag%)   ' close
End If

If DebugMode Then
Call IOWriteLog(vbNullString)
msg$ = Space$(4) & Format$("Conc", a80$) & Format$("K-Exp", a80$) & Format$("K-Rat", a80$) & Format$("K-Err", a80$)
Call IOWriteLog(msg$)
If CorrectionFlag% = 0 Or CorrectionFlag% = MAXCORRECTION% Then              ' ZAF/Phi-rho-z calculations (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
msg$ = Symup$(isym%(1)) & " " & Xraylo$(iray%(1)) & MiscAutoFormat$(conc!(1)) & MiscAutoFormat$(kexp!(1)) & MiscAutoFormat$(CalcZAFAnalysis.StdAssignsKfactors!(1)) & MiscAutoFormat$(KratioError!(1, CalcZAFOutputCount&))
Else                                     ' alpha factor calculations (mode% = 0 only)
msg$ = Symup$(isym%(1)) & " " & Xraylo$(iray%(1)) & MiscAutoFormat$(conc!(1)) & MiscAutoFormat$(kexp!(1)) & MiscAutoFormat$(KratioAlpha!(1)) & MiscAutoFormat$(KratioError!(1, CalcZAFOutputCount&))
End If
Call IOWriteLog(msg$)

If iray%(2) <> MAXRAY% Then
If CorrectionFlag% = 0 Or CorrectionFlag% = MAXCORRECTION% Then              ' ZAF/Phi-rho-z calculations
msg$ = Symup$(isym%(2)) & " " & Xraylo$(iray%(2)) & MiscAutoFormat$(conc!(2)) & MiscAutoFormat$(kexp!(2)) & MiscAutoFormat$(CalcZAFAnalysis.StdAssignsKfactors!(2)) & MiscAutoFormat$(KratioError!(2, CalcZAFOutputCount&))
Else                                     ' alpha factor calculations (mode% = 0 only)
msg$ = Symup$(isym%(2)) & " " & Xraylo$(iray%(2)) & MiscAutoFormat$(conc!(2)) & MiscAutoFormat$(kexp!(2)) & MiscAutoFormat$(KratioAlpha!(2)) & MiscAutoFormat$(KratioError!(2, CalcZAFOutputCount&))
End If
Call IOWriteLog(msg$)
End If
End If

' Output first approximation results only (hidden menu, see "ExtendedMenu" INI file parameter)
Else

' Convert mass concentration to appropriate fraction
Call CalcZAFFirstApproximationConvert(mode%, Int(2), isym%(), conc!())
If ierror Then Exit Sub

conc1!(1) = conc!(1)
conc1!(2) = conc!(2)

' Modify calculated fraction for intensity correction (conversion from weight to k-ratio)
If FirstApproximationApplyAbsorption Then
conc1!(1) = conc1!(1) / CalcZAFAnalysis.StdZAFCors!(1, 1, 1)
conc1!(2) = conc1!(2) / CalcZAFAnalysis.StdZAFCors!(1, 1, 2)
End If
If FirstApproximationApplyFluorescence Then
conc1!(1) = conc1!(1) / CalcZAFAnalysis.StdZAFCors!(2, 1, 1)
conc1!(2) = conc1!(2) / CalcZAFAnalysis.StdZAFCors!(2, 1, 2)
End If
If FirstApproximationApplyAtomicNumber Then
conc1!(1) = conc1!(1) / CalcZAFAnalysis.StdZAFCors!(3, 1, 1)
conc1!(2) = conc1!(2) / CalcZAFAnalysis.StdZAFCors!(3, 1, 2)
End If

If kexp!(1) <> 0# Then KratioError!(1, CalcZAFOutputCount&) = conc1!(1) / kexp!(1)
If kexp!(2) <> 0# Then KratioError!(2, CalcZAFOutputCount&) = conc1!(2) / kexp!(2)

' Output results
Print #ExportDataFileNumber2%, CalcZAFLineCount&, vbTab, isym%(1), vbTab, isym%(2), vbTab, iray%(1), vbTab, iray%(2), vbTab, eO!, vbTab, TOA!, vbTab, conc!(1), vbTab, conc!(2), vbTab, kexp!(1), vbTab, kexp!(2), vbTab, conc1!(1), vbTab, conc1!(2), vbTab, KratioError!(1, CalcZAFOutputCount&), vbTab, KratioError!(2, CalcZAFOutputCount&)

If DebugMode Then
Call IOWriteLog(vbNullString)
msg$ = Space$(4) & Format$("Frac", a80$) & Format$("Frac'", a80$) & Format$("K-Exp", a80$) & Format$("K-Err", a80$)
Call IOWriteLog(msg$)
msg$ = Symup$(isym%(1)) & " " & Xraylo$(iray%(1)) & MiscAutoFormat$(conc!(1)) & MiscAutoFormat$(conc1!(1)) & MiscAutoFormat$(kexp!(1)) & MiscAutoFormat$(KratioError!(1, CalcZAFOutputCount&))
Call IOWriteLog(msg$)
If iray%(2) = MAXRAY% Then
msg$ = Symup$(isym%(2)) & " " & Xraylo$(iray%(2)) & MiscAutoFormat$(conc!(2)) & MiscAutoFormat$(conc1!(2)) & MiscAutoFormat$(kexp!(2)) & MiscAutoFormat$(KratioError!(2, CalcZAFOutputCount&))
Call IOWriteLog(msg$)
End If
End If

End If

CalcZAFBinarySkip:
Next m% ' next MAC file
Next n% ' next matrix correction
If mode% = 4 Then GoTo CalcZAFBinarySingle
Loop

' Calculate average and standard deviation
CalcZAFBinarySingle:
If mode% <> 5 Then
Call MathArrayAverage3(average, KratioError!(), CalcZAFOutputCount&, 2)
If ierror Then Exit Sub

' Write to file
Print #ExportDataFileNumber2%, " "
Print #ExportDataFileNumber2%, VbDquote$ & "AverageA" & VbDquote$, vbTab, MiscAutoFormat$(average.averags!(1))
Print #ExportDataFileNumber2%, VbDquote$ & "StdDevA" & VbDquote$, vbTab, MiscAutoFormat$(average.Stddevs!(1))
Print #ExportDataFileNumber2%, VbDquote$ & "MinimumA" & VbDquote$, vbTab, MiscAutoFormat$(average.Minimums!(1))
Print #ExportDataFileNumber2%, VbDquote$ & "MaximumA" & VbDquote$, vbTab, MiscAutoFormat$(average.Maximums!(1))
Print #ExportDataFileNumber2%, VbDquote$ & "AverageB" & VbDquote$, vbTab, MiscAutoFormat$(average.averags!(2))
Print #ExportDataFileNumber2%, VbDquote$ & "StdDevB" & VbDquote$, vbTab, MiscAutoFormat$(average.Stddevs!(2))
Print #ExportDataFileNumber2%, VbDquote$ & "MinimumB" & VbDquote$, vbTab, MiscAutoFormat$(average.Minimums!(2))
Print #ExportDataFileNumber2%, VbDquote$ & "MaximumB" & VbDquote$, vbTab, MiscAutoFormat$(average.Maximums!(2))

Call IOWriteLog(vbNullString)
Call IOWriteLog("AverageA" & MiscAutoFormat$(average.averags!(1)))
Call IOWriteLog("StdDevA" & MiscAutoFormat$(average.Stddevs!(1)))
Call IOWriteLog("MinimumA" & MiscAutoFormat$(average.Minimums!(1)))
Call IOWriteLog("MaximumA" & MiscAutoFormat$(average.Maximums!(1)))
Call IOWriteLog("AverageB" & MiscAutoFormat$(average.averags!(2)))
Call IOWriteLog("StdDevB" & MiscAutoFormat$(average.Stddevs!(2)))
Call IOWriteLog("MinimumB" & MiscAutoFormat$(average.Minimums!(2)))
Call IOWriteLog("MaximumB" & MiscAutoFormat$(average.Maximums!(2)))
End If

' Close file
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%

' Read and calculate averages for mode% = 5 files
If mode% = 5 Then
nn% = MAXZAF%  ' loop on all correction options
mm% = MAXMACTYPE%   ' loop on all MAC files
For n% = 1 To nn%
For m% = 1 To mm%

' Open each file
tfilenumber% = CalcZAFBinaryFile%(Int(4), n%, m%)  ' open for INPUT
Line Input #tfilenumber%, astring$  ' read first line of symbols

' Read file
CalcZAFOutputCount& = 0
Do While Not EOF(tfilenumber%)
CalcZAFOutputCount& = CalcZAFOutputCount& + 1
ReDim Preserve KratioError!(1 To 2, 1 To CalcZAFOutputCount&)
Input #tfilenumber%, CalcZAFLineCount&, isym%(1), isym%(2), iray%(1), iray%(2), _
eO!, TOA!, conc!(1), conc!(2), kexp!(1), kexp!(2), _
CalcZAFAnalysis.StdAssignsKfactors!(1), CalcZAFAnalysis.StdAssignsKfactors!(2), _
CalcZAFAnalysis.StdZAFCors!(7, 1, 1), CalcZAFAnalysis.StdZAFCors!(8, 1, 1), _
CalcZAFAnalysis.StdZAFCors!(1, 1, 1), CalcZAFAnalysis.StdZAFCors!(2, 1, 1), _
CalcZAFAnalysis.StdZAFCors!(3, 1, 1), CalcZAFAnalysis.StdZAFCors!(5, 1, 1), _
CalcZAFAnalysis.StdZAFCors!(6, 1, 1), CalcZAFAnalysis.StdZAFCors!(4, 1, 1), _
CalcZAFAnalysis.StdZAFCors!(7, 1, 2), CalcZAFAnalysis.StdZAFCors!(8, 1, 2), _
CalcZAFAnalysis.StdZAFCors!(1, 1, 2), CalcZAFAnalysis.StdZAFCors!(2, 1, 2), _
CalcZAFAnalysis.StdZAFCors!(3, 1, 2), CalcZAFAnalysis.StdZAFCors!(5, 1, 2), _
CalcZAFAnalysis.StdZAFCors!(6, 1, 2), CalcZAFAnalysis.StdZAFCors!(4, 1, 2), _
KratioError!(1, CalcZAFOutputCount&), KratioError!(2, CalcZAFOutputCount&)
Loop

' Close each file
tfilenumber% = CalcZAFBinaryFile%(Int(3), n%, m%)

' Average
Call MathArrayAverage3(average, KratioError!(), CalcZAFOutputCount&, 2)
If ierror Then Exit Sub

' Open for APPEND
tfilenumber% = CalcZAFBinaryFile%(Int(2), n%, m%)

' Write results
Print #tfilenumber%, " "
Print #tfilenumber%, VbDquote$ & "AverageA" & VbDquote$, vbTab, MiscAutoFormat$(average.averags!(1))
Print #tfilenumber%, VbDquote$ & "StdDevA" & VbDquote$, vbTab, MiscAutoFormat$(average.Stddevs!(1))
Print #tfilenumber%, VbDquote$ & "MinimumA" & VbDquote$, vbTab, MiscAutoFormat$(average.Minimums!(1))
Print #tfilenumber%, VbDquote$ & "MaximumA" & VbDquote$, vbTab, MiscAutoFormat$(average.Maximums!(1))
Print #tfilenumber%, VbDquote$ & "AverageB" & VbDquote$, vbTab, MiscAutoFormat$(average.averags!(2))
Print #tfilenumber%, VbDquote$ & "StdDevB" & VbDquote$, vbTab, MiscAutoFormat$(average.Stddevs!(2))
Print #tfilenumber%, VbDquote$ & "MinimumB" & VbDquote$, vbTab, MiscAutoFormat$(average.Minimums!(2))
Print #tfilenumber%, VbDquote$ & "MaximumB" & VbDquote$, vbTab, MiscAutoFormat$(average.Maximums!(2))

' Close each file
tfilenumber% = CalcZAFBinaryFile%(Int(3), n%, m%)

Next m%
Next n%
End If

Call IOStatusAuto(vbNullString)
msg$ = "Binary calculations completed on file " & ImportDataFile2$ & vbCrLf
If mode% <> 5 Then msg$ = msg$ & "Data output saved to " & ExportDataFile$ & vbCrLf
If mode% < 4 Then
msg$ = msg$ & "Histogram output saved to " & HistogramDataFile$ & vbCrLf
ElseIf mode% = 4 Then
msg$ = msg$ & "For the binary " & MiscAutoUcase$(Symlo$(isym%(1))) & " " & Xraylo$(iray%(1)) & " in " & MiscAutoUcase$(Symlo$(isym%(2))) & " using all matrix correction options (1 to " & Str$(MAXZAF%) & ") and MAC files (1 to " & Str$(MAXMACTYPE%) & ")."
Else
msg$ = msg$ & "Multiple output files were created based on the input file " & ImportDataFile2$ & " and the ZAF and MAC file strings (1 to " & Str$(MAXZAF%) & ") and MAC files (1 to " & Str$(MAXMACTYPE%) & ")."
End If
Call IOWriteLog(vbCrLf & vbCrLf & msg$)

If mode% = 4 Or mode% = 5 Then
MsgBox msg$, vbOKOnly + vbInformation, "CalcZAFBinary"
End If

' Print out problematic lines
If mode% = 0 Then
Call IOWriteLog(vbNullString)
msg$ = "Problematic k-ratio errors (< 0.8 or > 1.2)"
Call IOWriteLog(msg$)
msg$ = Format$("Line", a60$) & Space$(10) & Format$("ConcA", a80$) & Format$("ConcB", a80$) & Format$("TOA", a80$) & Format$("eO", a80$) & Format$("Uo", a80$) & Format$("K-Exp", a80$) & Format$("K-Cal", a80$) & Format$("K-Err", a80$)
Call IOWriteLog(msg$)

For ii& = 1 To CalcZAFOutputCount&
If KratioError!(1, ii&) <> 0# Then
If VerboseMode Or (KratioError!(1, ii&) < 0.8 Or KratioError!(1, ii&) > 1.2) Then
msg$ = Format$(KratioLine&(1, ii&), "!@@@@@@") & KratioEsym$(1, ii&) & " " & KratioXsym$(1, ii&) & " in " & KratioEsym$(2, ii&) & MiscAutoFormat$(KratioConc!(1, ii&)) & MiscAutoFormat$(KratioConc!(2, ii&)) & MiscAutoFormat$(KratioTOAeO!(1, ii&)) & MiscAutoFormat$(KratioTOAeO!(2, ii&)) & MiscAutoFormat$(KratioOver!(1, ii&)) & MiscAutoFormat$(KratioExpr!(1, ii&)) & MiscAutoFormat$(KratioCalc!(1, ii&)) & MiscAutoFormat$(KratioError!(1, ii&))
Call IOWriteLog(msg$)
End If
End If

If KratioError!(2, ii&) <> 0# Then
If VerboseMode Or (KratioError!(2, ii&) < 0.8 Or KratioError!(2, ii&) > 1.2) Then
msg$ = Format$(KratioLine&(2, ii&), "!@@@@@@") & KratioEsym$(2, ii&) & " " & KratioXsym$(2, ii&) & " in " & KratioEsym$(1, ii&) & MiscAutoFormat$(KratioConc!(1, ii&)) & MiscAutoFormat$(KratioConc!(2, ii&)) & MiscAutoFormat$(KratioTOAeO!(1, ii&)) & MiscAutoFormat$(KratioTOAeO!(2, ii&)) & MiscAutoFormat$(KratioOver!(2, ii&)) & MiscAutoFormat$(KratioExpr!(2, ii&)) & MiscAutoFormat$(KratioCalc!(2, ii&)) & MiscAutoFormat$(KratioError!(2, ii&))
Call IOWriteLog(msg$)
End If
End If

Next ii&
End If

' Show histogram
If mode% < 4 Then
HistogramOutputOption% = mode%
Call CalcZAFPlotHistogram(Int(1))
If ierror Then Exit Sub
End If

' Restore current ZAF and MAC selection
izaf% = tzaftype%
Call InitGetZAFSetZAF2(izaf%)
If ierror Then Exit Sub
MACTypeFlag% = tmactype%
Call GetZAFAllSaveMAC2(MACTypeFlag%)
If ierror Then Exit Sub

' To set form enables
ImportDataFile2$ = vbNullString
Call CalcZAFSetEnables
If ierror Then Exit Sub

Exit Sub

' Errors
CalcZAFBinaryError:
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
MsgBox Error$ & ", on line " & Format$(CalcZAFLineCount&) & ", for binary " & Symup$(isym%(1)) & " " & Xraylo$(iray%(1)) & ", TOA= " & Format$(TOA!) & ", eO= " & Format$(eO!) & ", Conc= " & MiscAutoFormat$(conc!(1)) & ", Kexp= " & MiscAutoFormat$(kexp!(1)), vbOKOnly + vbCritical, "CalcZAFBinary"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

CalcZAFBinaryOutofLimits:
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
msg$ = "Bad data on line " & Str$(CalcZAFLineCount&) & " in " & ImportDataFile2$ & " (file format may be wrong)."
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFBinary"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

CalcZAFBinaryBothByDifference:
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
msg$ = "Both elements are by difference on line " & Str$(CalcZAFLineCount&) & " in " & ImportDataFile2$
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFBinary"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

CalcZAFBinaryNoConcData:
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
msg$ = "No Conc data on line " & Str$(CalcZAFLineCount&) & " in " & ImportDataFile2$
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFBinary"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

CalcZAFBinaryNoKexpData:
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
msg$ = "No K-exp data on line " & Str$(CalcZAFLineCount&) & " in " & ImportDataFile2$
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFBinary"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub CalcZAFFirstApproximation(mode As Integer, tForm As Form)
' Calculate binary data from .DAT file (first approximations only). See
' routine CalcZAFBinary also.
' mode=1 atomic fraction
' mode=2 mass fraction
' mode=3 electron fraction

ierror = False
On Error GoTo CalcZAFFirstApproximationError

Dim tfilename As String
Dim eO As Single, TOA As Single

ReDim isym(1 To 2) As Integer
ReDim iray(1 To 2) As Integer

ReDim conc(1 To 2) As Single
ReDim kexp(1 To 2) As Single

icancelauto = False

' Get import filename from user
If Trim$(ImportDataFile$) = vbNullString Then ImportDataFile$ = "pouchou.dat"
tfilename$ = ImportDataFile$
Call IOGetFileName(Int(2), "DAT", tfilename$, tForm)
If ierror Then Exit Sub

' Save current path
CalcZAFDATFileDirectory$ = CurDir$

' No errors, save file name
ImportDataFile$ = tfilename$

' Get export filename from user
tfilename$ = MiscGetFileNameNoExtension(tfilename$) & ".out"
Call IOGetFileName(Int(1), "OUT", tfilename$, tForm)
If ierror Then Exit Sub

' No errors, save file name
ExportDataFile$ = tfilename$
HistogramDataFile$ = MiscGetFileNameNoExtension(tfilename$) & ".txt"

Open ImportDataFile$ For Input As #ImportDataFileNumber%
Open ExportDataFile$ For Output As #ExportDataFileNumber%
CalcZAFLineCount& = 0
CalcZAFOutputCount& = 0
Call IOStatusAuto(vbNullString)

' Output colum labels
Print #ExportDataFileNumber%, "SymA", "SymB", "RayA", "RayB", "KeV", "Takeoff", "CalcA", "CalcB", "KexpA", "KexpB", "KerrA", "KerrB"

' Check for end of file
Do While Not EOF(ImportDataFileNumber%)
CalcZAFLineCount& = CalcZAFLineCount& + 1

msg$ = "Calculating first approximation " & Str$(CalcZAFLineCount&) & "..."
Call IOStatusAuto(msg$)
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
ierror = True
Exit Sub
End If

' Read mode and number of elements, kilovlts and takeoff
Input #ImportDataFileNumber%, isym%(1), isym%(2), iray%(1), iray%(2), eO!, TOA!, conc!(1), conc!(2), kexp!(1), kexp!(2)

' Check limits
If isym%(1) < 1 Or isym%(1) > MAXELM% Then GoTo CalcZAFFirstApproximationOutofLimits
If isym%(2) < 1 Or isym%(2) > MAXELM% Then GoTo CalcZAFFirstApproximationOutofLimits
If iray%(1) < 1 Or iray%(1) > MAXRAY% Then GoTo CalcZAFFirstApproximationOutofLimits
If iray%(2) < 1 Or iray%(2) > MAXRAY% Then GoTo CalcZAFFirstApproximationOutofLimits
If eO! < 1# Or eO! > 100# Then GoTo CalcZAFFirstApproximationOutofLimits
If TOA! < 1# Or TOA! > 90# Then GoTo CalcZAFFirstApproximationOutofLimits
If conc!(1) < 0# Or conc!(1) > 1# Then GoTo CalcZAFFirstApproximationOutofLimits
If conc!(2) < 0# Or conc!(2) > 1# Then GoTo CalcZAFFirstApproximationOutofLimits
If kexp!(1) < 0# Or kexp!(1) > 1# Then GoTo CalcZAFFirstApproximationOutofLimits
If kexp!(2) < 0# Or kexp!(2) > 1# Then GoTo CalcZAFFirstApproximationOutofLimits

' Check that both elements are not by difference
If iray%(1) = MAXRAY% And iray%(2) = MAXRAY% Then GoTo CalcZAFFirstApproximationBothByDifference

' Check that at least one concentration is entered
If conc!(1) = 0# And conc!(2) = 0# Then GoTo CalcZAFFirstApproximationNoConcData

' Check for valid kexp data if x-ray used
If iray%(1) <= MAXRAY% - 1 And kexp!(1) = 0# Then GoTo CalcZAFFirstApproximationNoKexpData
If iray%(2) <= MAXRAY% - 1 And kexp!(2) = 0# Then GoTo CalcZAFFirstApproximationNoKexpData

' Calculate elements by difference
If iray%(1) = MAXRAY% Then
conc!(1) = 1# - conc!(2)
End If

If iray%(2) = MAXRAY% Then
conc!(2) = 1# - conc!(1)
End If

' Convert mass concentration to appropriate fraction
Call CalcZAFFirstApproximationConvert(mode%, Int(2), isym%(), conc!())
If ierror Then Exit Sub

' Save data
CalcZAFOutputCount& = CalcZAFOutputCount& + 1
ReDim Preserve KratioError!(1 To 2, 1 To CalcZAFOutputCount&)

' Calculate error for first approximation
KratioError!(1, CalcZAFOutputCount&) = 0#
KratioError!(2, CalcZAFOutputCount&) = 0#

If kexp!(1) <> 0# Then KratioError!(1, CalcZAFOutputCount&) = conc!(1) / kexp!(1)
If kexp!(2) <> 0# Then KratioError!(2, CalcZAFOutputCount&) = conc!(2) / kexp!(2)

' Output results
Print #ExportDataFileNumber%, isym%(1), isym%(2), iray%(1), iray%(2), eO!, TOA!, conc!(1), conc!(2), kexp!(1), kexp!(2), KratioError!(1, CalcZAFOutputCount&), KratioError!(2, CalcZAFOutputCount&)
Loop

' Close file
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%

Call IOStatusAuto(vbNullString)
msg$ = vbCrLf & vbCrLf & "Calculations on completed on file " & ImportDataFile$ & vbCrLf
msg$ = msg$ & "Data output saved to " & ExportDataFile$ & vbCrLf
msg$ = msg$ & "Histogram output saved to " & HistogramDataFile$ & vbCrLf
Call IOWriteLog(msg$)

' Calculate histogram
HistogramOutputOption% = -mode%
Call CalcZAFPlotHistogram(Int(1))
If ierror Then Exit Sub

Exit Sub

' Errors
CalcZAFFirstApproximationError:
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFFirstApproximation"
ierror = True
Exit Sub

CalcZAFFirstApproximationOutofLimits:
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
msg$ = "Bad data on line " & Str$(CalcZAFLineCount&) & " in " & ImportDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFFirstApproximation"
ierror = True
Exit Sub

CalcZAFFirstApproximationBothByDifference:
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
msg$ = "Both elements are by difference on line " & Str$(CalcZAFLineCount&) & " in " & ImportDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFFirstApproximation"
ierror = True
Exit Sub

CalcZAFFirstApproximationNoConcData:
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
msg$ = "No Conc data on line " & Str$(CalcZAFLineCount&) & " in " & ImportDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFFirstApproximation"
ierror = True
Exit Sub

CalcZAFFirstApproximationNoKexpData:
Close #ImportDataFileNumber%
Close #ExportDataFileNumber%
msg$ = "No K-exp data on line " & Str$(CalcZAFLineCount&) & " in " & ImportDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFFirstApproximation"
ierror = True
Exit Sub

End Sub

Sub CalcZAFFirstApproximationConvert(mode As Integer, num As Integer, isym() As Integer, conc() As Single)
' Convert mass concentration to atomic or electron fractions

ierror = False
On Error GoTo CalcZAFFirstApproximationConvertError

Dim i As Integer

ReDim temp(1 To num%) As Single
ReDim atnm(1 To num%) As Single
ReDim atwt(1 To num%) As Single

' Load atomic numbers
For i% = 1 To num%
atnm!(i%) = CSng(AllAtomicNums%(isym%(i%)))
Next i%

' Load atomic weights
For i% = 1 To num%
atwt!(i%) = AllAtomicWts!(isym%(i%))
Next i%

' Load temp array
If mode% = 1 Or mode% = 3 Then
For i% = 1 To num%
temp!(i%) = conc!(i%)
Next i%

' Calculate weight to atomic fraction
If mode% = 1 Then
Call ConvertWeightToAtomic(num%, atwt!(), temp!(), conc!())
If ierror Then Exit Sub
End If

' Calculate weight to electron fraction
If mode% = 3 Then
Call ConvertWeightToElectron(num%, atnm!(), atwt!(), temp!(), conc!())
If ierror Then Exit Sub
End If
End If

Exit Sub

' Errors
CalcZAFFirstApproximationConvertError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFFirstApproximationConvert"
Call CalcZAFImportClose
ierror = True
Exit Sub

End Sub

Sub CalcZAFBinaryLoad()
' Load binary options

ierror = False
On Error GoTo CalcZAFBinaryLoadError

Static initialized As Integer

If Not initialized Then
BinaryOutputRangeAbsMin! = 0.2  ' 20% correction
BinaryOutputRangeFluMin! = 0.1  ' 10% correction
BinaryOutputRangeZedMin! = 0.1  ' 10% correction
BinaryOutputRangeAbsMax! = 0.02  ' 2% correction
BinaryOutputRangeFluMax! = 0.02  ' 2% correction
BinaryOutputRangeZedMax! = 0.02  ' 2% correction
BinaryOutputMinimumZbarDiff! = 2#
BinaryOutputMaximumZbarDiff! = 0.5
initialized = True
End If

' Minimum absorption
If BinaryOutputRangeMinAbs Then
FormBINARY.CheckUseMinimumAbsorptionCorrectionOutput.Value = vbChecked
Else
FormBINARY.CheckUseMinimumAbsorptionCorrectionOutput.Value = vbUnchecked
End If

FormBINARY.TextMinimumAbsorptionCorrection.Text = Str$(BinaryOutputRangeAbsMin!)

' Minimum fluorescence
If BinaryOutputRangeMinFlu Then
FormBINARY.CheckUseMinimumFluorescenceCorrectionOutput.Value = vbChecked
Else
FormBINARY.CheckUseMinimumFluorescenceCorrectionOutput.Value = vbUnchecked
End If

FormBINARY.TextMinimumFluorescenceCorrection.Text = Str$(BinaryOutputRangeFluMin!)

' Minimum atomic number
If BinaryOutputRangeMinZed Then
FormBINARY.CheckUseMinimumAtomicNumberCorrectionOutput.Value = vbChecked
Else
FormBINARY.CheckUseMinimumAtomicNumberCorrectionOutput.Value = vbUnchecked
End If

FormBINARY.TextMinimumAtomicNumberCorrection.Text = Str$(BinaryOutputRangeZedMin!)

' Maximum absorption
If BinaryOutputRangeMaxAbs Then
FormBINARY.CheckUseMaximumAbsorptionCorrectionOutput.Value = vbChecked
Else
FormBINARY.CheckUseMaximumAbsorptionCorrectionOutput.Value = vbUnchecked
End If

FormBINARY.TextMaximumAbsorptionCorrection.Text = Str$(BinaryOutputRangeAbsMax!)

' Maximum fluorescence
If BinaryOutputRangeMaxFlu Then
FormBINARY.CheckUseMaximumFluorescenceCorrectionOutput.Value = vbChecked
Else
FormBINARY.CheckUseMaximumFluorescenceCorrectionOutput.Value = vbUnchecked
End If

FormBINARY.TextMaximumFluorescenceCorrection.Text = Str$(BinaryOutputRangeFluMax!)

' Maximum atomic number
If BinaryOutputRangeMaxZed Then
FormBINARY.CheckUseMaximumAtomicNumberCorrectionOutput.Value = vbChecked
Else
FormBINARY.CheckUseMaximumAtomicNumberCorrectionOutput.Value = vbUnchecked
End If

FormBINARY.TextMaximumAtomicNumberCorrection.Text = Str$(BinaryOutputRangeZedMax!)

' Zbar filters
If BinaryOutputMinimumZbar Then
FormBINARY.CheckBinaryOutputMinimumZbar.Value = vbChecked
Else
FormBINARY.CheckBinaryOutputMinimumZbar.Value = vbUnchecked
End If

FormBINARY.TextBinaryOutputMinimumZbarDiff.Text = Str$(BinaryOutputMinimumZbarDiff!)

If BinaryOutputMaximumZbar Then
FormBINARY.CheckBinaryOutputMaximumZbar.Value = vbChecked
Else
FormBINARY.CheckBinaryOutputMaximumZbar.Value = vbUnchecked
End If

FormBINARY.TextBinaryOutputMaximumZbarDiff.Text = Str$(BinaryOutputMaximumZbarDiff!)

' First approximation options
If FirstApproximationApplyAbsorption Then
FormBINARY.CheckFirstApproximationApplyAbsorption.Value = vbChecked
Else
FormBINARY.CheckFirstApproximationApplyAbsorption.Value = vbUnchecked
End If

If FirstApproximationApplyFluorescence Then
FormBINARY.CheckFirstApproximationApplyFluorescence.Value = vbChecked
Else
FormBINARY.CheckFirstApproximationApplyFluorescence.Value = vbUnchecked
End If

If FirstApproximationApplyAtomicNumber Then
FormBINARY.CheckFirstApproximationApplyAtomicNumber.Value = vbChecked
Else
FormBINARY.CheckFirstApproximationApplyAtomicNumber.Value = vbUnchecked
End If

Exit Sub

' Errors
CalcZAFBinaryLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFBinaryLoad"
ierror = True
Exit Sub

End Sub

Sub CalcZAFBinarySave()
' Save binary options

ierror = False
On Error GoTo CalcZAFBinarySaveError

' Output ranges
If FormBINARY.CheckUseMinimumAbsorptionCorrectionOutput.Value = vbChecked Then
BinaryOutputRangeMinAbs = True
Else
BinaryOutputRangeMinAbs = False
End If

If Val(FormBINARY.TextMinimumAbsorptionCorrection.Text) > 0# Then
BinaryOutputRangeAbsMin! = Val(FormBINARY.TextMinimumAbsorptionCorrection.Text)
Else
msg$ = "Minimum absorption correction is out of range (must be greater than 0.0)"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFHistogramSave"
ierror = True
Exit Sub
End If

If FormBINARY.CheckUseMinimumFluorescenceCorrectionOutput.Value = vbChecked Then
BinaryOutputRangeMinFlu = True
Else
BinaryOutputRangeMinFlu = False
End If

If Val(FormBINARY.TextMinimumFluorescenceCorrection.Text) > 0# Then
BinaryOutputRangeFluMin! = Val(FormBINARY.TextMinimumFluorescenceCorrection.Text)
Else
msg$ = "Minimum fluorescence correction is out of range (must be greater than 0.0)"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFHistogramSave"
ierror = True
Exit Sub
End If

If FormBINARY.CheckUseMinimumAtomicNumberCorrectionOutput.Value = vbChecked Then
BinaryOutputRangeMinZed = True
Else
BinaryOutputRangeMinZed = False
End If

If Val(FormBINARY.TextMinimumAtomicNumberCorrection.Text) > 0# Then
BinaryOutputRangeZedMin! = Val(FormBINARY.TextMinimumAtomicNumberCorrection.Text)
Else
msg$ = "Minimum atomic number correction is out of range (must be greater than 0.0)"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFHistogramSave"
ierror = True
Exit Sub
End If

' Maximum absorption
If FormBINARY.CheckUseMaximumAbsorptionCorrectionOutput.Value = vbChecked Then
BinaryOutputRangeMaxAbs = True
Else
BinaryOutputRangeMaxAbs = False
End If

If Val(FormBINARY.TextMaximumAbsorptionCorrection.Text) > 0# Then
BinaryOutputRangeAbsMax! = Val(FormBINARY.TextMaximumAbsorptionCorrection.Text)
Else
msg$ = "Maximum absorption correction is out of range (must be greater than 0.0)"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFHistogramSave"
ierror = True
Exit Sub
End If

' Maximum fluorescence
If FormBINARY.CheckUseMaximumFluorescenceCorrectionOutput.Value = vbChecked Then
BinaryOutputRangeMaxFlu = True
Else
BinaryOutputRangeMaxFlu = False
End If

If Val(FormBINARY.TextMaximumFluorescenceCorrection.Text) > 0# Then
BinaryOutputRangeFluMax! = Val(FormBINARY.TextMaximumFluorescenceCorrection.Text)
Else
msg$ = "Maximum fluorescence correction is out of range (must be greater than 0.0)"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFHistogramSave"
ierror = True
Exit Sub
End If

' Maximum atomic number
If FormBINARY.CheckUseMaximumAtomicNumberCorrectionOutput.Value = vbChecked Then
BinaryOutputRangeMaxZed = True
Else
BinaryOutputRangeMaxZed = False
End If

If Val(FormBINARY.TextMaximumAtomicNumberCorrection.Text) > 0# Then
BinaryOutputRangeZedMax! = Val(FormBINARY.TextMaximumAtomicNumberCorrection.Text)
Else
msg$ = "Maximum atomic number correction is out of range (must be greater than 0.0)"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFHistogramSave"
ierror = True
Exit Sub
End If

' Zbar output filters
If FormBINARY.CheckBinaryOutputMinimumZbar.Value = vbChecked Then
BinaryOutputMinimumZbar = True
Else
BinaryOutputMinimumZbar = False
End If

If Val(FormBINARY.TextBinaryOutputMinimumZbarDiff.Text) > 0# Then
BinaryOutputMinimumZbarDiff! = Val(FormBINARY.TextBinaryOutputMinimumZbarDiff.Text)
Else
msg$ = "Minimum Mass-Electron zbar percent difference is out of range (must be greater than 0.0)"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFHistogramSave"
ierror = True
Exit Sub
End If

If FormBINARY.CheckBinaryOutputMaximumZbar.Value = vbChecked Then
BinaryOutputMaximumZbar = True
Else
BinaryOutputMaximumZbar = False
End If

If Val(FormBINARY.TextBinaryOutputMaximumZbarDiff.Text) > 0# Then
BinaryOutputMaximumZbarDiff! = Val(FormBINARY.TextBinaryOutputMaximumZbarDiff.Text)
Else
msg$ = "Maximum Mass-Electron zbar percent difference is out of range (must be greater than 0.0)"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFHistogramSave"
ierror = True
Exit Sub
End If

' First approximation options
If FormBINARY.CheckFirstApproximationApplyAbsorption.Value = vbChecked Then
FirstApproximationApplyAbsorption = True
Else
FirstApproximationApplyAbsorption = False
End If

If FormBINARY.CheckFirstApproximationApplyFluorescence.Value = vbChecked Then
FirstApproximationApplyFluorescence = True
Else
FirstApproximationApplyFluorescence = False
End If

If FormBINARY.CheckFirstApproximationApplyAtomicNumber.Value = vbChecked Then
FirstApproximationApplyAtomicNumber = True
Else
FirstApproximationApplyAtomicNumber = False
End If

Exit Sub

' Errors
CalcZAFBinarySaveError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFBinarySave"
ierror = True
Exit Sub

End Sub

Sub CalcZAFOption()
' Get ZAF options

ierror = False
On Error GoTo CalcZAFOptionError

' Load options
Call ZAFOptionLoad(CalcZAFOldSample())
If ierror Then Exit Sub

' Show the form
FormZAFOPT.Show vbModal
If ierror Then Exit Sub

' Save options
Call ZAFOptionReturnSample(CalcZAFOldSample())
If ierror Then Exit Sub

Exit Sub

' Errors
CalcZAFOptionError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFOption"
ierror = True
Exit Sub

End Sub

Function CalcZAFBinaryFile(mode As Integer, iz As Integer, im As Integer) As Integer
' Handles multiple file open for CalcZAFBinary based on file and izaf and MAC file type
' mode% = 1 delete file
' mode% = 2 open file for APPEND
' mode% = 3 close file
' mode% = 4 open file for INPUT
' returns the open filenumber handle (or zero for delete file)

ierror = False
On Error GoTo CalcZAFBinaryFileError

Dim tfilename As String
Dim tfilenumber As Integer

' Create file name based on file, and izaf and MAC file type
CalcZAFBinaryFile% = 0
tfilename$ = MiscGetFileNameOnly$(ImportDataFile2$)
tfilename$ = MiscGetFileNameNoExtension$(tfilename$) & " " & zafstring$(iz%) & ", " & macstring$(im%) & " .DAT"

' Check for invalid characters for filenames
Call MiscModifyStringToFilename(tfilename$)
If ierror Then Exit Function

' Add user data path
tfilename$ = MiscGetPathOnly2$(ImportDataFile2$) & "\" & tfilename$

' Close output file in case already open no matter what
tfilenumber% = 200 + (iz% - 1) * (iz% - 1) + im%
Close (tfilenumber%)

' Delete file
If mode% = 1 Then
If Dir$(tfilename$) <> vbNullString Then Kill tfilename$

' Open file for APPEND
ElseIf mode% = 2 Then
Open tfilename$ For Append As #tfilenumber%

' Close file
ElseIf mode% = 3 Then
Close (tfilenumber%)

' Open file for INPUT (calculating averages only)
ElseIf mode% = 4 Then
Open tfilename$ For Input As #tfilenumber%
End If

CalcZAFBinaryFile% = tfilenumber%
Exit Function

' Errors
CalcZAFBinaryFileError:
Close (tfilenumber%)
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFBinaryFile"
ierror = True
Exit Function

End Function

Sub CalcZAFGetPTC()
' Load PTC options

ierror = False
On Error GoTo CalcZAFGetPTCError

Call GetPTCLoad
If ierror Then Exit Sub

Call TypeZAFSelections
If ierror Then Exit Sub

' Load element arrays
If CalcZAFOldSample(1).LastElm% > 0 Then
Call ElementGetData(CalcZAFOldSample())
If ierror Then Exit Sub

' Initialize calculations (needed for ZAFPTC and coating calculations) (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% <> MAXCORRECTION% Then
Call ZAFSetZAF(CalcZAFOldSample())
If ierror Then Exit Sub
Else
'Call ZAFSetZAF3(CalcZAFOldSample())
'If ierror Then Exit Sub
End If
End If

' Calculate standard k-factors
Call UpdateAllStdKfacs(CalcZAFAnalysis, CalcZAFOldSample(), CalcZAFTmpSample())
If ierror Then Exit Sub

Exit Sub

' Errors
CalcZAFGetPTCError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFGetPTC"
ierror = True
Exit Sub

End Sub

Sub CalcZAFExportOpen(tForm As Form)
' Open export file

ierror = False
On Error GoTo CalcZAFExportOpenError

Dim response As Integer
Dim tfilename As String

If CalcZAFOldSample(1).LastChan% = 0 Then GoTo CalcZAFExportOpenNoElements

' Get filename from user
tfilename$ = ExportDataFile$
If Trim$(tfilename$) = vbNullString Then tfilename$ = "CalcZAF-Export.dat"
Call IOGetFileName(Int(0), "DAT", tfilename$, tForm)
If ierror Then Exit Sub

' Since user wants to open file make sure it is closed
Close #ExportDataFileNumber%
DoEvents

If Dir$(tfilename$) <> vbNullString Then
msg$ = "Output File: " & vbCrLf
msg$ = msg$ & tfilename$ & vbCrLf
msg$ = msg$ & " already exists, are you sure you want to overwrite it (click No to append)?"
response% = MsgBox(msg$, vbYesNoCancel + vbQuestion + vbDefaultButton2, "CalcZAFExportOpen")

If response% = vbCancel Then
ierror = True
Exit Sub
End If

' If user selects overwrite, erase it
If response% = vbYes Then
Kill tfilename$
End If
End If

' No errors, save file name
ExportDataFile$ = tfilename$
Open ExportDataFile$ For Append As #ExportDataFileNumber%

Exit Sub

' Errors
CalcZAFExportOpenError:
Close #ExportDataFileNumber%
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFExportOpen"
ierror = True
Exit Sub

CalcZAFExportOpenNoElements:
msg$ = "No elements to export"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFExportOpen"
ierror = True
Exit Sub

End Sub

Sub CalcZAFExportSend()
' Export current data to file

ierror = False
On Error GoTo CalcZAFExportSendError

Dim i As Integer

' Write mode and number of elements, kilovolts and takeoff (and sample name and stage coordinates)
'Print #ExportDataFileNumber%, Str$(CalcZAFMode%) & VbComma, Str$(CalcZAFOldSample(1).LastChan%) & VbComma, Str$(CalcZAFOldSample(1).kilovolts!) & VbComma, Str$(CalcZAFOldSample(1).takeoff!) & VbComma, VbDquote & CalcZAFOldSample(1).Name$ & VbDquote & VbComma, Str$(CalcZAFOldSample(1).StagePositions!(1, 1)) & VbComma, , Str$(CalcZAFOldSample(1).StagePositions!(1, 2)) & VbComma, Str$(CalcZAFOldSample(1).StagePositions!(1, 3))
Print #ExportDataFileNumber%, Format$(CalcZAFMode%) & VbComma$ & Format$(CalcZAFOldSample(1).LastChan%) & VbComma$ & Format$(CalcZAFOldSample(1).kilovolts!) & VbComma$ & Format$(CalcZAFOldSample(1).takeoff!) & VbComma$ & VbDquote & CalcZAFOldSample(1).Name$ & VbDquote & VbComma & Format$(CalcZAFOldSample(1).StagePositions!(1, 1)) & VbComma$ & Format$(CalcZAFOldSample(1).StagePositions!(1, 2)) & VbComma$ & Format$(CalcZAFOldSample(1).StagePositions!(1, 3))

' Write oxide, difference, stoichiometry, relative parameters
Print #ExportDataFileNumber%, CalcZAFOldSample(1).OxideOrElemental%, VbDquote$ & CalcZAFOldSample(1).DifferenceElement$ & VbDquote$, VbDquote$ & CalcZAFOldSample(1).StoichiometryElement$ & VbDquote$, CalcZAFOldSample(1).StoichiometryRatio!, VbDquote$ & CalcZAFOldSample(1).RelativeElement$ & VbDquote$, VbDquote$ & CalcZAFOldSample(1).RelativeToElement$ & VbDquote$, CalcZAFOldSample(1).RelativeRatio!

' Loop on each element
For i% = 1 To CalcZAFOldSample(1).LastChan%
Print #ExportDataFileNumber%, VbDquote$ & CalcZAFOldSample(1).Elsyms$(i%) & VbDquote$, VbDquote$ & CalcZAFOldSample(1).Xrsyms$(i%) & VbDquote$, CalcZAFOldSample(1).numcat%(i%), CalcZAFOldSample(1).numoxd%(i%), CalcZAFOldSample(1).StdAssigns%(i%), CalcZAFOldSample(1).ElmPercents!(i%), UnkCounts!(i%), StdCounts!(i%)
Next i%

Exit Sub

' Errors
CalcZAFExportSendError:
Close #ExportDataFileNumber%
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFExportSend"
ierror = True
Exit Sub

End Sub

Sub CalcZAFExportSend2(firstline As Boolean, xydist As Single, laststring As String)
' Export current data to file (uses single line output)

ierror = False
On Error GoTo CalcZAFExportSend2Error

Call IOStatusAuto("CalcZAFExportSend2: exporting sample " & CalcZAFOldSample(1).Name$ & ", line " & Format$(CalcZAFOldSample(1).Linenumber&(1)) & "...")

' Create column label string
Call CalcZAFExportColumnString(laststring$, CalcZAFOldSample())
If ierror Then Exit Sub

' Get relative distance (CalcZAF always used just the first data line)
Call PlotGetRelativeMicrons(firstline, Int(1), xydist!, CalcZAFOldSample())
If ierror Then Exit Sub

' Create data string
Call CalcZAFExportDataString(xydist!, CalcZAFAnalysis, CalcZAFOldSample())
If ierror Then Exit Sub

Exit Sub

' Errors
CalcZAFExportSend2Error:
Close #ExportDataFileNumber%
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFExportSend2"
ierror = True
Exit Sub

End Sub

Sub CalcZAFExportClose(mode As Integer)
' Close export file
'   mode = 0 report current sample name
'   mode = 1 report all samples

ierror = False
On Error GoTo CalcZAFExportCloseError

Close #ExportDataFileNumber%

If mode% = 0 Then
msg$ = "Sample " & CalcZAFOldSample(1).Name$ & " data was exported to " & ExportDataFile$
Else
msg$ = "All sample data was exported to " & ExportDataFile$
End If
MsgBox msg$, vbOKOnly + vbInformation, "CalcZAFExportClose"

Exit Sub

' Errors
CalcZAFExportCloseError:
Close #ExportDataFileNumber%
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFExportClose"
ierror = True
Exit Sub

End Sub

Sub CalcZAFStandard(mode As Integer, tForm As Form)
' Calculate composition of standard from raw k-ratio input data from .DAT file, Data file format assumes one line
' for each composition.
' mode = 0 process all data lines in file using current matrix corrections
' mode = 1 process only first line in file, loop on all correction options (izaf% = 1 to 10, MACFileType% = 1 to 6)
'
' The first column is "standard name" in double quotes.
' The second column is the takeoff angle.
' The third column is the kilovolts.
' The fourth column is the oxide or elemental flag (1 = elemental, 2 = oxide).
' The fifth column is the number of elements.
' The next n columns are the element symbols (use double quotes).
' The next n columns are the element x-ray lines (use double quotes) (leave empty for specified concentration).
' The next n columns are the measured k-ratios for each element (blank for specified concentration).
' The next n columns are the "published" weight percent for each element.
' The next n columns are the primary standard number assignment for each element.
' The next n columns are the number of cations per element.
' The next n columns are the number of oxygens per element.
'
' Example (first line contains all analyzed elements, 2nd line contains oxygen as a specified concentration)
' "Fayalite", 40.0, 15.0, 2, 3, "fe", "si", "o", "ka", "ka", "ka", .7372, .2523, 1.0086, 54.809, 13.785, 31.407, 895, 14, 895, 1, 1, 1, 1, 2, 0
' "Fayalite", 40.0, 15.0, 2, 3, "fe", "si", "o", "ka", "ka", "", .7372, .2523, 0.0, 54.809, 13.785, 31.407, 895, 14, 0, 1, 1, 1, 1, 2, 0

ierror = False
On Error GoTo CalcZAFStandardError

Dim i As Integer, ip As Integer
Dim n As Integer, nn As Integer
Dim m As Integer, mm As Integer
Dim temp As Single
Dim tfilename As String
Dim originalstring As String, astring As String

Dim average As TypeAverage

' Show form
FormZAF.Show vbModeless
icancelauto = False

' Get import filename from user
If ImportDataFile2$ = vbNullString Then ImportDataFile2$ = "calczaf2.dat"
tfilename$ = ImportDataFile2$
Call IOGetFileName(Int(2), "DAT", tfilename$, tForm)
If ierror Then Exit Sub

' Save current ZAF and MAC selection
tzaftype% = izaf%
tmactype% = MACTypeFlag%

' Save current path
CalcZAFDATFileDirectory$ = CurDir$

' No errors, save file name
ImportDataFile2$ = tfilename$

' Get export filename from user
tfilename$ = MiscGetFileNameNoExtension(tfilename$) & ".out"
Call IOGetFileName(Int(1), "OUT", tfilename$, tForm)
If ierror Then Exit Sub

' No errors, save file name
ExportDataFile$ = tfilename$
HistogramDataFile$ = MiscGetFileNameNoExtension(tfilename$) & ".txt"

' Get available standard names and numbers from database
Call StandardGetMDBIndex
If ierror Then Exit Sub

' Open normal files
Open ImportDataFile2$ For Input As #ImportDataFileNumber2%
Open ExportDataFile$ For Output As #ExportDataFileNumber2%
CalcZAFLineCount& = 0
CalcZAFOutputCount& = 0
Call IOStatusAuto(vbNullString)

' Check for end of file
Do While Not EOF(ImportDataFileNumber2%)
CalcZAFLineCount& = CalcZAFLineCount& + 1

' Initialize
Call CalcZAFInit
If ierror Then Exit Sub

' Check for Pause button
Do Until Not RealTimePauseAutomation
DoEvents
Sleep 200
Loop

CalcZAFMode% = 2    ' calculate composition from raw k-ratios
CalcZAFOldSample(1).number% = CalcZAFLineCount&

' Read binary elements, kilovlts and takeoff
Input #ImportDataFileNumber2%, CalcZAFOldSample(1).Name$, CalcZAFOldSample(1).takeoff!, CalcZAFOldSample(1).kilovolts!
Input #ImportDataFileNumber2%, CalcZAFOldSample(1).OxideOrElemental%, CalcZAFOldSample(1).LastChan%

' Update status
If mode% = 0 Then
msg$ = "Calculating standard " & CalcZAFOldSample(1).Name$ & ", line " & Str$(CalcZAFLineCount&) & "..."
Else
msg$ = "Calculating standard " & CalcZAFOldSample(1).Name$ & "..."
End If
Call IOStatusAuto(msg$)
If icancelauto Then
Call IOStatusAuto(vbNullString)
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
ierror = True
Exit Sub
End If

' Loop on element columns
For i% = 1 To CalcZAFOldSample(1).LastChan%
Input #ImportDataFileNumber2%, CalcZAFOldSample(1).Elsyms$(i%)
Next i%

For i% = 1 To CalcZAFOldSample(1).LastChan%
Input #ImportDataFileNumber2%, CalcZAFOldSample(1).Xrsyms$(i%)
Next i%

' Measured raw k-ratio
For i% = 1 To CalcZAFOldSample(1).LastChan%
Input #ImportDataFileNumber2%, UnkCounts!(i%)
Next i%

' Published percents
For i% = 1 To CalcZAFOldSample(1).LastChan%
Input #ImportDataFileNumber2%, CalcZAFOldSample(1).ElmPercents!(i%)
Next i%

' Add standard to run
For i% = 1 To CalcZAFOldSample(1).LastChan%
Input #ImportDataFileNumber2%, CalcZAFOldSample(1).StdAssigns%(i%)

ip% = IPOS2(NumberofStandards%, CalcZAFOldSample(1).StdAssigns%(i%), StandardNumbers%())
If ip% = 0 And CalcZAFOldSample(1).StdAssigns%(i%) > 0 Then
Call AddStdSaveStd(CalcZAFOldSample(1).StdAssigns%(i%))
If ierror Then
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
Exit Sub
End If
End If
Next i%

' Cations/oxygens
For i% = 1 To CalcZAFOldSample(1).LastChan%
Input #ImportDataFileNumber2%, CalcZAFOldSample(1).numcat%(i%)
Next i%

For i% = 1 To CalcZAFOldSample(1).LastChan%
Input #ImportDataFileNumber2%, CalcZAFOldSample(1).numoxd%(i%)
Next i%

' Check limits
If CalcZAFOldSample(1).kilovolts! < 1# Or CalcZAFOldSample(1).kilovolts! > 100# Then GoTo CalcZAFStandardOutofLimits
If CalcZAFOldSample(1).takeoff! < 1# Or CalcZAFOldSample(1).takeoff! > 90# Then GoTo CalcZAFStandardOutofLimits
If CalcZAFOldSample(1).LastChan% < 1 Or CalcZAFOldSample(1).LastChan% > MAXCHAN% Then GoTo CalcZAFStandardOutofLimits

' Check for valid kexp data if x-ray used (k-raw data can be zero!)
'For i% = 1 To CalcZAFOldSample(1).LastChan%
'If CalcZAFOldSample(1).Xrsyms$(i%) <> vbNullString And UnkCounts!(i%) = 0# Then GoTo CalcZAFStandardNoKexpData
'Next i%

' Update defaults
DefaultTakeOff! = CalcZAFOldSample(1).takeoff!
DefaultKiloVolts! = CalcZAFOldSample(1).kilovolts!

' Make sure that new condition arrays are loaded
For i% = 1 To CalcZAFOldSample(1).LastChan%
CalcZAFOldSample(1).TakeoffArray!(i%) = CalcZAFOldSample(1).takeoff!
CalcZAFOldSample(1).KilovoltsArray!(i%) = CalcZAFOldSample(1).kilovolts!
Next i%

' Load sample number
CalcZAFOldSample(1).Linenumber&(1) = CalcZAFLineCount&

' Check for stoichiometric oxygen and subtract if so
If CalcZAFOldSample(1).OxideOrElemental% = 1 Then
Call ZAFGetOxygenChannel(CalcZAFOldSample())
If ierror Then Exit Sub
temp! = ConvertOxygenFromCations(CalcZAFOldSample())
If ierror Then Exit Sub
CalcZAFOldSample(1).ElmPercents!(CalcZAFOldSample(1).OxygenChannel%) = CalcZAFOldSample(1).ElmPercents!(CalcZAFOldSample(1).OxygenChannel%) - temp!
End If

' Sort elements
Call CalcZAFSave
If ierror Then
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
Exit Sub
End If

' Load form
Call CalcZAFLoad
If ierror Then
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
Exit Sub
End If

' Load element strings
Call ElementLoadArrays(CalcZAFOldSample())
If ierror Then Exit Sub

' Output column headings (if different)
Call CalcZAFStandardOutputColumns(astring$, CalcZAFOldSample())
If ierror Then Exit Sub

' Output column labels
If originalstring$ <> astring$ Then
Print #ExportDataFileNumber2%, astring$
originalstring$ = astring$
End If

' Load next matrix correction
nn% = 1
mm% = 1
If mode% = 1 Then
nn% = MAXZAF%  ' loop on all correction options
mm% = MAXMACTYPE%   ' loop on all MAC files
End If
For n% = 1 To nn%
For m% = 1 To mm%

' Set ZAF and MAC if looping on all
If mode% = 1 Then
izaf% = n%

' Check for MAC file
Call GetZAFAllSaveMAC2(m%)
If ierror Then
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
Exit Sub
End If
MACTypeFlag% = m%   ' set after check for exist

' Set ZAF corrections
Call InitGetZAFSetZAF2(izaf%)
If ierror Then
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
Exit Sub
End If

' Update k-factors and parameters
Call CalcZAFUpdateAllStdKfacs
If ierror Then
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
Exit Sub
End If
End If

If mode% = 1 Then
msg$ = "Calculating standard with " & zafstring$(izaf%) & ", " & macstring$(MACTypeFlag%)
Call IOWriteLog(vbCrLf & vbCrLf & msg$)
Call IOStatusAuto(msg$ & "...")
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
ierror = True
Exit Sub
End If
End If

' Calculate k-factors based on published wt% (for output)
Call CalcZAFStandardCalculateStdKfactors
If ierror Then
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
Exit Sub
End If

' Calculate actual binary intensities
Call CalcZAFCalculate
If ierror Then
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
Exit Sub
End If

' Check for stoichiometric oxygen and add back in to published wt%
If CalcZAFOldSample(1).OxideOrElemental% = 1 Then
Call ZAFGetOxygenChannel(CalcZAFOldSample())
If ierror Then Exit Sub
temp! = ConvertOxygenFromCations(CalcZAFOldSample())
If ierror Then Exit Sub
CalcZAFOldSample(1).ElmPercents!(CalcZAFOldSample(1).OxygenChannel%) = CalcZAFOldSample(1).ElmPercents!(CalcZAFOldSample(1).OxygenChannel%) + temp!
End If

' Save data
CalcZAFOutputCount& = CalcZAFOutputCount& + 1
ReDim Preserve ConcError!(1 To MAXCHAN%, 1 To CalcZAFOutputCount&)

' Calculate error
For i% = 1 To CalcZAFOldSample(1).LastChan%
ConcError!(i%, CalcZAFOutputCount&) = 0#
If CalcZAFOldSample(1).ElmPercents!(i%) <> 0# Then ConcError!(i%, CalcZAFOutputCount&) = CalcZAFAnalysis.WtPercents!(i%) / CalcZAFOldSample(1).ElmPercents!(i%)
Next i%

' Output data
Call CalcZAFStandardOutputData(mode%, CalcZAFAnalysis, CalcZAFOldSample())
If ierror Then Exit Sub

Next m% ' next MAC file
Next n% ' next matrix correction
If mode% = 1 Then Exit Do
Loop

' Calculate average and standard deviation
Call MathArrayAverage3(average, ConcError!(), CalcZAFOutputCount&, CalcZAFOldSample(1).LastChan%)
If ierror Then
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
Exit Sub
End If

' Write to file
Print #ExportDataFileNumber2%, " "
For i% = 1 To CalcZAFOldSample(1).LastChan%
Print #ExportDataFileNumber2%, VbDquote$ & CalcZAFOldSample(1).Elsyup$(i%) & " Average" & VbDquote$, vbTab, MiscAutoFormat$(average.averags!(i%))
Print #ExportDataFileNumber2%, VbDquote$ & CalcZAFOldSample(1).Elsyup$(i%) & " StdDev" & VbDquote$, vbTab, MiscAutoFormat$(average.Stddevs!(i%))
Print #ExportDataFileNumber2%, VbDquote$ & CalcZAFOldSample(1).Elsyup$(i%) & " Minimum" & VbDquote$, vbTab, MiscAutoFormat$(average.Minimums!(i%))
Print #ExportDataFileNumber2%, VbDquote$ & CalcZAFOldSample(1).Elsyup$(i%) & " Maximum" & VbDquote$, vbTab, MiscAutoFormat$(average.Maximums!(i%))
Next i%

Call IOWriteLog(vbNullString)
Call IOWriteLog("Calculated Compositional Accuracy Errors (1.000 = no error):")
For i% = 1 To CalcZAFOldSample(1).LastChan%
Call IOWriteLog(CalcZAFOldSample(1).Elsyms$(i%) & " Average" & MiscAutoFormat$(average.averags!(i%)))
Call IOWriteLog(CalcZAFOldSample(1).Elsyms$(i%) & " StdDev" & MiscAutoFormat$(average.Stddevs!(i%)))
Call IOWriteLog(CalcZAFOldSample(1).Elsyms$(i%) & " Minimum" & MiscAutoFormat$(average.Minimums!(i%)))
Call IOWriteLog(CalcZAFOldSample(1).Elsyms$(i%) & " Maximum" & MiscAutoFormat$(average.Maximums!(i%)))
Call IOWriteLog(vbNullString)
Next i%

' Close file
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%

Call IOStatusAuto(vbNullString)
msg$ = "Standard calculations completed on file " & ImportDataFile2$ & vbCrLf
msg$ = msg$ & "Data output saved to " & ExportDataFile$ & vbCrLf
If mode% = 1 Then
msg$ = msg$ & "For standard " & CalcZAFOldSample(1).Name$ & " using all matrix correction options (1 to " & Str$(MAXZAF%) & ") and MAC files (1 to " & Str$(MAXMACTYPE%) & ")."
End If
Call IOWriteLog(vbCrLf & vbCrLf & msg$)

MsgBox msg$, vbOKOnly + vbInformation, "CalcZAFStandard"

' Restore current ZAF and MAC selection
izaf% = tzaftype%
Call InitGetZAFSetZAF2(izaf%)
If ierror Then Exit Sub
MACTypeFlag% = tmactype%
Call GetZAFAllSaveMAC2(MACTypeFlag%)
If ierror Then Exit Sub

Exit Sub

' Errors
CalcZAFStandardError:
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFStandard"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

CalcZAFStandardOutofLimits:
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
msg$ = "Bad data on line " & Str$(CalcZAFLineCount&) & " in " & ImportDataFile2$ & " (file format may be wrong)."
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFStandard"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

CalcZAFStandardBothByDifference:
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
msg$ = "All elements are by difference on line " & Str$(CalcZAFLineCount&) & " in " & ImportDataFile2$
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFStandard"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

CalcZAFStandardNoConcData:
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
msg$ = "No Conc data on line " & Str$(CalcZAFLineCount&) & " in " & ImportDataFile2$
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFStandard"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

CalcZAFStandardNoKexpData:
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
msg$ = "No K-exp data on line " & Str$(CalcZAFLineCount&) & " in " & ImportDataFile2$
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFStandard"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub CalcZAFStandardOutputColumns(astring As String, sample() As TypeSample)
' Output the column headings

ierror = False
On Error GoTo CalcZAFStandardOutputColumnsError

Dim i As Integer

' Create output string
astring$ = " " & VbDquote$ & "Line" & VbDquote$ & vbTab & VbDquote$ & "Name" & VbDquote$ & vbTab & _
    VbDquote$ & "Takeoff" & VbDquote$ & vbTab & VbDquote$ & "keV" & VbDquote$ & vbTab
    
' K-raws
For i% = 1 To CalcZAFOldSample(1).LastChan%
astring$ = astring$ & VbDquote$ & MiscAutoUcase$(CalcZAFOldSample(1).Elsyms$(i%)) & " Kraw" & VbDquote$ & vbTab
Next i%

' Elemental kratio
For i% = 1 To CalcZAFOldSample(1).LastChan%
astring$ = astring$ & VbDquote$ & MiscAutoUcase$(CalcZAFOldSample(1).Elsyms$(i%)) & " Krat" & VbDquote$ & vbTab
Next i%

' Published Elemental k-ratio
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & VbDquote$ & MiscAutoUcase$(CalcZAFOldSample(1).Elsyms$(i%)) & " Publ-Krat" & VbDquote$ & vbTab
Next i%

' Elemental wt%
For i% = 1 To CalcZAFOldSample(1).LastChan%
astring$ = astring$ & VbDquote$ & MiscAutoUcase$(CalcZAFOldSample(1).Elsyms$(i%)) & " Wt%" & VbDquote$ & vbTab
Next i%

' Oxide wt%
If CalcZAFOldSample(1).OxideOrElemental% = 1 Then
For i% = 1 To CalcZAFOldSample(1).LastChan%
astring$ = astring$ & VbDquote$ & CalcZAFOldSample(1).Oxsyup$(i%) & " Wt%" & VbDquote$ & vbTab
Next i%
End If

astring$ = astring$ & VbDquote$ & "Wt% Total" & VbDquote$ & vbTab

For i% = 1 To CalcZAFOldSample(1).LastChan%
astring$ = astring$ & VbDquote$ & MiscAutoUcase$(CalcZAFOldSample(1).Elsyms$(i%)) & " Publ" & VbDquote$ & vbTab
Next i%

astring$ = astring$ & VbDquote$ & "Publ Total" & VbDquote$ & vbTab

' Calculation errors
For i% = 1 To CalcZAFOldSample(1).LastChan%
astring$ = astring$ & VbDquote$ & MiscAutoUcase$(CalcZAFOldSample(1).Elsyms$(i%)) & " Conc Error" & VbDquote$ & vbTab
Next i%

' ZAF factors
For i% = 1 To CalcZAFOldSample(1).LastChan%
astring$ = astring$ & VbDquote$ & MiscAutoUcase$(CalcZAFOldSample(1).Elsyms$(i%)) & " Pri F(Chi)" & VbDquote$ & vbTab
Next i%

For i% = 1 To CalcZAFOldSample(1).LastChan%
astring$ = astring$ & VbDquote$ & MiscAutoUcase$(CalcZAFOldSample(1).Elsyms$(i%)) & " Sec F(Chi)" & VbDquote$ & vbTab
Next i%

For i% = 1 To CalcZAFOldSample(1).LastChan%
astring$ = astring$ & VbDquote$ & MiscAutoUcase$(CalcZAFOldSample(1).Elsyms$(i%)) & " Absorp" & VbDquote$ & vbTab
Next i%

For i% = 1 To CalcZAFOldSample(1).LastChan%
astring$ = astring$ & VbDquote$ & MiscAutoUcase$(CalcZAFOldSample(1).Elsyms$(i%)) & " Fluor" & VbDquote$ & vbTab
Next i%

For i% = 1 To CalcZAFOldSample(1).LastChan%
astring$ = astring$ & VbDquote$ & MiscAutoUcase$(CalcZAFOldSample(1).Elsyms$(i%)) & " Zed" & VbDquote$ & vbTab
Next i%

For i% = 1 To CalcZAFOldSample(1).LastChan%
astring$ = astring$ & VbDquote$ & MiscAutoUcase$(CalcZAFOldSample(1).Elsyms$(i%)) & " Stp" & VbDquote$ & vbTab
Next i%

For i% = 1 To CalcZAFOldSample(1).LastChan%
astring$ = astring$ & VbDquote$ & MiscAutoUcase$(CalcZAFOldSample(1).Elsyms$(i%)) & " Bks" & VbDquote$ & vbTab
Next i%

For i% = 1 To CalcZAFOldSample(1).LastChan%
astring$ = astring$ & VbDquote$ & MiscAutoUcase$(CalcZAFOldSample(1).Elsyms$(i%)) & " ZAF" & VbDquote$ & vbTab
Next i%

Exit Sub

' Errors
CalcZAFStandardOutputColumnsError:
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFStandardOutputColumns"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub CalcZAFStandardOutputData(mode As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' Output the data

ierror = False
On Error GoTo CalcZAFStandardOutputDataError

Dim i As Integer
Dim sum As Single
Dim astring As String

' Create output string
astring$ = Str$(CalcZAFLineCount&) & vbTab & VbDquote & sample(1).Name$ & VbDquote & vbTab
astring$ = astring$ & Str$(sample(1).takeoff!) & vbTab & Str$(sample(1).kilovolts!) & vbTab

' Kraw value
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & Format$(UnkCounts!(i%), f85$) & vbTab
Next i%

' Elemental k-ratio
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & Format$(analysis.UnkKrats!(i%), f85$) & vbTab
Next i%

' Published Elemental k-ratio
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & Format$(StdKFactors!(i%), f85$) & vbTab
Next i%

' Calculated wt%
sum! = 0#
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & Format$(analysis.WtPercents!(i%), f83$) & vbTab
sum! = sum! + analysis.WtPercents!(i%)
Next i%

' Oxide wt%
If CalcZAFOldSample(1).OxideOrElemental% = 1 Then
For i% = 1 To sample(1).LastChan%
analysis.OxPercents!(i%) = ConvertElmToOxd!(analysis.WtPercents!(i%), sample(1).Elsyms$(i%), sample(1).numcat%(i%), sample(1).numoxd%(i%))
astring$ = astring$ & Format$(analysis.OxPercents!(i%), f83$) & vbTab
Next i%
End If

' Calculated total
astring$ = astring$ & Format$(sum!, f83$) & vbTab

' Published wt%
sum! = 0#
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & Format$(sample(1).ElmPercents!(i%), f83$) & vbTab
sum! = sum! + sample(1).ElmPercents!(i%)
Next i%

' Published total
astring$ = astring$ & Format$(sum!, f83$) & vbTab

' Calculation errors
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & Format$(ConcError!(i%, CalcZAFOutputCount&), f85$) & vbTab
Next i%
    
' Emitted/generated intensities
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & Format$(analysis.UnkZAFCors!(7, i%), f85$) & vbTab
Next i%
    
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & Format$(analysis.UnkZAFCors!(8, i%), f85$) & vbTab
Next i%
    
' ZAF factors
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & Format$(analysis.UnkZAFCors!(1, i%), f85$) & vbTab
Next i%
    
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & Format$(analysis.UnkZAFCors!(2, i%), f85$) & vbTab
Next i%
    
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & Format$(analysis.UnkZAFCors!(3, i%), f85$) & vbTab
Next i%
    
' Stopping power/backscatter
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & Format$(analysis.UnkZAFCors!(5, i%), f85$) & vbTab
Next i%

For i% = 1 To sample(1).LastChan%
astring$ = astring$ & Format$(analysis.UnkZAFCors!(6, i%), f85$) & vbTab
Next i%
    
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & Format$(analysis.UnkZAFCors!(4, i%), f85$) & vbTab
Next i%
    
' If all matrix corrections, output correction string
If mode% = 1 Then
astring$ = astring$ & VbDquote$ & zafstring$(izaf%) & ", " & macstring$(MACTypeFlag%) & VbDquote$
End If

' Output normal k-ratio results
Print #ExportDataFileNumber2%, astring$
Exit Sub

' Errors
CalcZAFStandardOutputDataError:
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFStandardOutputData"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub CalcZAFStandardCalculateStdKfactors()
' Calculate standard k-factors from published wt% data (called by CalcZAFStandard only)
 
ierror = False
On Error GoTo CalcZAFStandardCalculateStdKfactorsError

Dim i As Integer
Dim tStdAssigns(1 To MAXCHAN%) As Integer

' Save original standard assignments
For i% = 1 To CalcZAFOldSample(1).LastChan%
tStdAssigns%(i%) = CalcZAFOldSample(1).StdAssigns%(i%)
CalcZAFOldSample(1).StdAssigns%(i%) = CalcZAFOldSample(1).number%    ' fake standard assignment
Next i%

' Fake sample coating for ZAFStd calculation
If UseConductiveCoatingCorrectionForElectronAbsorption Then                  ' fake standard coating
StandardCoatingFlag%(1) = CalcZAFOldSample(1).CoatingFlag%
StandardCoatingDensity!(1) = CalcZAFOldSample(1).CoatingDensity!
StandardCoatingThickness!(1) = CalcZAFOldSample(1).CoatingThickness!
StandardCoatingElement%(1) = CalcZAFOldSample(1).CoatingElement%
End If

' Set TmpSample equal to OldSample so k factors and ZAF corrections get loaded in ZAFStd
CalcZAFTmpSample(1) = CalcZAFOldSample(1)

' Reload the element arrays
Call ElementGetData(CalcZAFOldSample())
If ierror Then Exit Sub

' Initialize calculations (needed for ZAFPTC and coating calculations) (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% <> MAXCORRECTION% Then
Call ZAFSetZAF(CalcZAFOldSample())
If ierror Then Exit Sub
Else
'Call ZAFSetZAF3(CalcZAFOldSample())
'If ierror Then Exit Sub
End If

' Run the calculations on the "standard"
If CorrectionFlag% = 0 Then
VerboseMode% = True
Call ZAFStd2(Int(1), CalcZAFAnalysis, CalcZAFOldSample(), CalcZAFTmpSample())
VerboseMode% = False
If ierror Then Exit Sub
ElseIf CorrectionFlag% = MAXCORRECTION% Then
'VerboseMode% = True
'Call ZAFStd3(Int(1), CalcZAFAnalysis, CalcZAFOldSample(), CalcZAFTmpSample())
'VerboseMode% = False
'If ierror Then Exit Sub
End If

' Load calculated k-ratios from published wt percents and restore original standard assignments
For i% = 1 To CalcZAFOldSample(1).LastChan%
StdKFactors!(i%) = CalcZAFAnalysis.StdAssignsKfactors!(i%)
CalcZAFOldSample(1).StdAssigns%(i%) = tStdAssigns%(i%)
Next i%

Exit Sub

' Errors
CalcZAFStandardCalculateStdKfactorsError:
Close #ImportDataFileNumber2%
Close #ExportDataFileNumber2%
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFStandardCalculateStdKfactors"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub CalcZAFSelectStandardDatabase(tForm As Form)
' Change the default standard database

ierror = False
On Error GoTo CalcZAFSelectStandardDatabaseError

Dim tfilename As String
Dim versionnumber As Single

' Get probe file name
Call IOGetMDBFileName(Int(4), tfilename$, tForm)
If ierror Then Exit Sub

' Check file and version
versionnumber! = FileInfoGetVersion(tfilename$, "STANDARD")
If ierror Then Exit Sub

' Load new name
StandardDataFile$ = tfilename$

' Calculate ZAF again
If ProbeDataFile$ <> vbNullString Then
Call UpdateAllStdKfacs(CalcZAFAnalysis, CalcZAFOldSample(), CalcZAFTmpSample())
If ierror Then Exit Sub
End If

Exit Sub

' Errors
CalcZAFSelectStandardDatabaseError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFSelectStandardDatabase"
ierror = True
Exit Sub

End Sub

Sub CalcZAFCalculateKRatiosAlphaFactors(tForm As Form)
' Calculate and output k-Ratios and alpha factors for the periodic table

ierror = False
On Error GoTo CalcZAFCalculateKRatiosAlphaFactorsError

Dim response As Integer
Dim i As Integer, j As Integer
Dim k As Integer, m As Integer
Dim tfilename As String
Dim stddev1 As Single, stddev2 As Single

ReDim wout(1 To MAXBINARY% * 2) As Single, rout(1 To MAXBINARY% * 2) As Single
ReDim eout(1 To MAXBINARY% * 2) As String, xout(1 To MAXBINARY% * 2) As String
ReDim zout(1 To MAXBINARY% * 2) As Integer

ReDim wout2(1 To 2) As Single, rout2(1 To 2) As Single
ReDim eout2(1 To 2) As String, xout2(1 To 2) As String
ReDim zout2(1 To 2) As Integer

Dim npts1 As Integer, npts2 As Integer
ReDim xdata1(1 To MAXBINARY%) As Single, xdata2(1 To MAXBINARY%) As Single
ReDim ydata1(1 To MAXBINARY%) As Single, ydata2(1 To MAXBINARY%) As Single
ReDim acoeff1(1 To MAXCOEFF%) As Single, acoeff2(1 To MAXCOEFF%) As Single

ReDim filenamearray(1 To MAXBINARY% + 1) As String     ' filearray for k-ratio Excel import
ReDim filenamearray2(1 To 3) As String               ' filearray for a-factor Excel import

For i% = 1 To MAXELM%
oxfactor!(i%) = 0#      ' oxide end-member k-ratios

For j% = 1 To MAXELM%
For k% = 1 To MAXBINARY%
kfactor!(i%, j%, k%) = 0#   ' elemental k-ratio
Next k%

alpha11!(i%, j%) = 1#       ' constant alpha factors

alpha21!(i%, j%) = 1#       ' linear (intercept) alpha factors
alpha22!(i%, j%) = 0#       ' linear (slope) alpha factors

alpha31!(i%, j%) = 1#       ' polynomial alpha factors
alpha32!(i%, j%) = 0#       ' polynomial alpha factors
alpha33!(i%, j%) = 0#       ' polynomial alpha factors

alpha41!(i%, j%) = 1#       ' non-linear polynomial alpha factors
alpha42!(i%, j%) = 0#       ' non-linear polynomial alpha factors
alpha43!(i%, j%) = 0#       ' non-linear polynomial alpha factors
Next j%
Next i%

' Ask user if they want to calculate the entire table
msg$ = "Are you sure you want to calculate k-ratios and alpha factors for the entire periodic table at " & Format$(DefaultKiloVolts!) & " keV?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "CalcZAFCalculateKRatiosAlphaFactors")
If response% = vbCancel Then Exit Sub

' Check for Bence-Albee corrections
If CorrectionFlag% < 1 Or CorrectionFlag% > 4 Then
msg$ = "Alpha Factor corrections are not currently selected. Changing matrix correction type to alpha-factors for k-ratio a-factor calculations."
MsgBox msg$, vbOKOnly + vbInformation, "CalcZAFCalculateKRatiosAlphaFactors"
CorrectionFlag% = 3     ' for polynomial alpha factors
End If

' Indicate alpha-factor update
AllAFactorUpdateNeeded = True

' Determine which binary to calculate
icancelauto = False
For i% = 1 To MAXELM%
'If i% < 8 Or i% > 26 Then GoTo 4000       ' testing purposes only
For j% = i% + 1 To MAXELM%
'If j% < 8 Or j% > 26 Then GoTo 4000       ' testing purposes only

' Load the binary
Call CalcZAFLoadBinary(i%, j%)
If ierror Then Exit Sub

msg$ = Format$(CalcZAFOldSample(1).Elsyms$(1), a20$) & " " & CalcZAFOldSample(1).Xrsyms$(1) & " and " & Format$(CalcZAFOldSample(1).Elsyms$(2), a20$) & " " & CalcZAFOldSample(1).Xrsyms$(2)
msg$ = "Calculating alpha-factor binary " & msg$
Call IOWriteLog(msg$)
Call IOStatusAuto(msg$)

' Initialize
For k% = 1 To MAXBINARY% * 2
wout!(k%) = 0#
rout!(k%) = 1#
Next k%

' Fill element arrays
Call ElementLoadArrays(CalcZAFOldSample())
If ierror Then Exit Sub

Call ElementCheckXray(Int(0), CalcZAFOldSample())
If ierror Then
msg$ = "Skipping binary " & Symlo$(i%) & "-" & Symlo$(j%) & "..."
Call IOWriteLog(msg$)
GoTo 4000
End If

' Initialize calculations (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% <> MAXCORRECTION% Then
Call ZAFSetZAF(CalcZAFOldSample())
If ierror Then Exit Sub
Else
'Call ZAFSetZAF3(CalcZAFOldSample())
'If ierror Then Exit Sub
End If

If icancelauto Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

' Calculate array of intensities using ZAF or Phi-Rho-Z
Call ZAFAFactor(wout!(), rout!(), eout$(), xout$(), zout%(), CalcZAFAnalysis, CalcZAFOldSample())
If ierror Then Exit Sub

If icancelauto Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

' Output k-ratios to element arrays for both emitters and all MAXBIN concentrations
For k% = 1 To MAXBINARY%
kfactor!(i%, j%, k%) = rout!(2 * k% - 1)
kfactor!(i%, j%, (MAXBINARY% - (k% - 1))) = rout!(MAXBINARY% * 2 - (2 * k% - 1))

kfactor!(j%, i%, (MAXBINARY% - (k% - 1))) = rout!(2 * k%)
kfactor!(j%, i%, k%) = rout!(MAXBINARY% * 2 - (2 * (k% - 1)))
Next k%

' Fit the alpha factors and load into alpha-factor look up tables
For m% = 1 To 4
CorrectionFlag% = m%
Call AFactorCalculateFitFactors(i%, j%, wout!(), rout!(), eout$(), xout$(), zout%())
If ierror Then Exit Sub

' Return the a-factor fit data
Call AFactorReturnAFactors(Int(1), npts1%, xdata1!(), ydata1!(), acoeff1!(), stddev1!)
If ierror Then Exit Sub

Call AFactorReturnAFactors(Int(2), npts2%, xdata2!(), ydata2!(), acoeff2!(), stddev2!)
If ierror Then Exit Sub

If icancelauto Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

' Output alpha factors to element arrays
If CorrectionFlag% = 1 Then
If npts1% > 0 Then alpha11!(i%, j%) = acoeff1!(1)
If npts2% > 0 Then alpha11!(j%, i%) = acoeff2!(1)
End If

If CorrectionFlag% = 2 Then
If npts1% > 0 Then alpha21!(i%, j%) = acoeff1!(1)
If npts1% > 0 Then alpha22!(i%, j%) = acoeff1!(2)
If npts2% > 0 Then alpha21!(j%, i%) = acoeff2!(1)
If npts2% > 0 Then alpha22!(j%, i%) = acoeff2!(2)
End If

If CorrectionFlag% = 3 Then
If npts1% > 0 Then alpha31!(i%, j%) = acoeff1!(1)
If npts1% > 0 Then alpha32!(i%, j%) = acoeff1!(2)
If npts1% > 0 Then alpha33!(i%, j%) = acoeff1!(3)
If npts2% > 0 Then alpha31!(j%, i%) = acoeff2!(1)
If npts2% > 0 Then alpha32!(j%, i%) = acoeff2!(2)
If npts2% > 0 Then alpha33!(j%, i%) = acoeff2!(3)
End If

If CorrectionFlag% = 4 Then
If npts1% > 0 Then alpha41!(i%, j%) = acoeff1!(1)
If npts1% > 0 Then alpha42!(i%, j%) = acoeff1!(2)
If npts1% > 0 Then alpha43!(i%, j%) = acoeff1!(3)
If npts2% > 0 Then alpha41!(j%, i%) = acoeff2!(1)
If npts2% > 0 Then alpha42!(j%, i%) = acoeff2!(2)
If npts2% > 0 Then alpha43!(j%, i%) = acoeff2!(3)
End If
Next m%

4000:
Next j%
Next i%
Call IOStatusAuto(vbNullString)

' Calculate array of end-member oxide intensities using ZAF or Phi-Rho-Z
For i% = 1 To MAXELM%

' Load the binary oxide
Call CalcZAFLoadBinary(i%, Int(ATOMIC_NUM_OXYGEN%))
If ierror Then Exit Sub

msg$ = Format$(CalcZAFOldSample(1).Elsyms$(1), a20$) & " " & CalcZAFOldSample(1).Xrsyms$(1) & " and " & Format$(CalcZAFOldSample(1).Elsyms$(2), a20$) & " " & CalcZAFOldSample(1).Xrsyms$(2)
msg$ = "Calculating oxide k-ratio binary " & msg$
Call IOWriteLog(msg$)
Call IOStatusAuto(msg$)

' Initialize
For k% = 1 To 2
wout2!(k%) = 0#
rout2!(k%) = 1#
Next k%

' Fill element arrays
Call ElementLoadArrays(CalcZAFOldSample())
If ierror Then Exit Sub

Call ElementCheckXray(Int(0), CalcZAFOldSample())
If ierror Then
msg$ = "Skipping binary " & Symlo$(i%) & "-" & Symlo$(j%) & "..."
Call IOWriteLog(msg$)
GoTo 5000
End If

' Initialize calculations (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% <> MAXCORRECTION% Then
Call ZAFSetZAF(CalcZAFOldSample())
If ierror Then Exit Sub
Else
'Call ZAFSetZAF3(CalcZAFOldSample())
'If ierror Then Exit Sub
End If

If icancelauto Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

' Calculate array of oxide intensities using ZAF or Phi-Rho-Z
Call ZAFAFactorOxide(wout2!(), rout2!(), eout2$(), xout2$(), zout2%(), CalcZAFAnalysis, CalcZAFOldSample())
If ierror Then Exit Sub

If icancelauto Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

' Output oxide composition k-ratios to element arrays for all emitters
oxfactor!(i%) = rout2!(1)
5000:
Next i%
Call IOStatusAuto(vbNullString)

' Get base filename from user
tfilename$ = "CalcZAF_Output.dat"
Call IOGetFileName(Int(1), "DAT", tfilename$, tForm)
If ierror Then Exit Sub

msg$ = vbCrLf & "Saving K-Ratios to ASCII files..."
Call IOWriteLog(msg$)
Call IOStatusAuto(msg$)

' Output k-ratios
Call CalcZAFSaveRatios(tfilename$, wout!(), Int(0), filenamearray$())      ' output k-ratios for all MAXBIN concentrations
Call CalcZAFSaveRatios(tfilename$, wout!(), Int(4), filenamearray$())      ' output k-ratios for oxide end-members

msg$ = "Saving alpha factors to ASCII files..."
Call IOWriteLog(msg$)
Call IOStatusAuto(msg$)

' Output alpha-factors
Call CalcZAFSaveFactors(tfilename$, Int(1), filenamearray2$())    ' constant
Call CalcZAFSaveFactors(tfilename$, Int(2), filenamearray2$())    ' linear
Call CalcZAFSaveFactors(tfilename$, Int(3), filenamearray2$())    ' polynomial
Call CalcZAFSaveFactors(tfilename$, Int(4), filenamearray2$())    ' non-linear

msg$ = "All K-ratios and alpha factors were output to ASCII files based on filename " & MiscGetFileNameNoExtension$(tfilename$)
Call IOWriteLog(msg$)
Call IOStatusAuto(vbNullString)

' Check if user wants to send k-ratio data files to Excel
msg$ = "Do you want to send the k-ratio elemental and oxide end-member data files to Excel?"
response% = MsgBox(msg$, vbYesNoCancel + vbQuestion + vbDefaultButton1, "CalcZAFCalculateKRatiosAlphaFactors")

' Send k-ratio files to excel
If response% = vbYes Then
Call ExcelSendFileListToExcel(MAXBINARY% + 1, filenamearray$(), tForm)
If ierror Then Exit Sub
End If

' Check if user wants to send alpha-factor data files to Excel
msg$ = "Do you want to send the alpha factor data files to Excel?"
response% = MsgBox(msg$, vbYesNoCancel + vbQuestion + vbDefaultButton1, "CalcZAFCalculateKRatiosAlphaFactors")

' Send alpha factor files to Excel
If response% = vbYes Then
Call ExcelSendFileListToExcel(Int(3), filenamearray2$(), tForm)
If ierror Then Exit Sub
End If

Call IOStatusAuto(vbNullString)
Exit Sub

' Errors
CalcZAFCalculateKRatiosAlphaFactorsError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFCalculateKRatiosAlphaFactors"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub CalcZAFLoadBinary(i As Integer, j As Integer)
' Load a binary sample based on the periodic table

ierror = False
On Error GoTo CalcZAFLoadBinaryError

' Initialize
Call CalcZAFInit
If ierror Then Exit Sub

CalcZAFMode% = 0    ' calculate intensities from concentrations
CalcZAFOldSample(1).number% = i% + (j% - 1) * MAXELM%
CalcZAFOldSample(1).OxideOrElemental% = 2
CalcZAFOldSample(1).LastElm% = 2
CalcZAFOldSample(1).LastChan% = 2

CalcZAFOldSample(1).numcat%(1) = 1
CalcZAFOldSample(1).numoxd%(1) = 0
CalcZAFOldSample(1).numcat%(2) = 1
CalcZAFOldSample(1).numoxd%(2) = 0

CalcZAFOldSample(1).StdAssigns%(1) = CalcZAFOldSample(1).number%    ' for proper loading of parameters
CalcZAFOldSample(1).StdAssigns%(2) = CalcZAFOldSample(1).number%

' Load sample
CalcZAFOldSample(1).Elsyms$(1) = Symlo$(i%)
CalcZAFOldSample(1).Elsyms$(2) = Symlo$(j%)

CalcZAFOldSample(1).Xrsyms$(1) = Deflin$(i%)
CalcZAFOldSample(1).Xrsyms$(2) = Deflin$(j%)

CalcZAFOldSample(1).takeoff! = DefaultTakeOff!
CalcZAFOldSample(1).kilovolts! = DefaultKiloVolts!

Exit Sub

' Errors
CalcZAFLoadBinaryError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFLoadBinary"
ierror = True
Exit Sub

End Sub

Sub CalcZAFSaveFactors(tfilename As String, mode As Integer, filenamearray2() As String)
' Save the alpha-factors to an ASCII file
'  mode = 1  output constant alpha factors
'  mode = 2  output linear alpha factors
'  mode = 3  output polynomial alpha factors
'  mode = 4  output non-linear alpha factors

ierror = False
On Error GoTo CalcZAFSaveFactorsError

Dim i As Integer, j As Integer
Dim astring As String

' Open the alpha-factor ASCII file
If mode% = 1 Then AFactorDataFile$ = MiscGetFileNameNoExtension$(tfilename$) & "_Afactor1" & ".dat"
If mode% = 2 Then AFactorDataFile$ = MiscGetFileNameNoExtension$(tfilename$) & "_Afactor2" & ".dat"
If mode% = 3 Then AFactorDataFile$ = MiscGetFileNameNoExtension$(tfilename$) & "_Afactor3" & ".dat"
If mode% = 4 Then AFactorDataFile$ = MiscGetFileNameNoExtension$(tfilename$) & "_Afactor4" & ".dat"
Open AFactorDataFile$ For Output As #AFactorDataFileNumber%
filenamearray2$(mode%) = AFactorDataFile$

' Write run info
If mode% = 1 Then astring$ = "CONSTANT Alpha Factors"
If mode% = 2 Then astring$ = "LINEAR Alpha Factors"
If mode% = 3 Then astring$ = "POLYNOMIAL Alpha Factors"
If mode% = 4 Then astring$ = "NON-LINEAR Alpha Factors"
astring$ = astring$ & " derived from elemental k-ratios"
Print #AFactorDataFileNumber%, VbDquote$ & astring$ & VbDquote$

astring$ = "ZAF Corr: " & zafstring$(izaf%) & ", MAC Table: " & macstring$(MACTypeFlag%)
Print #AFactorDataFileNumber%, VbDquote$ & astring$ & VbDquote$

Print #AFactorDataFileNumber%, VbDquote$ & Now & VbDquote$
Print #AFactorDataFileNumber%, CalcZAFOldSample(1).takeoff!, vbTab, CalcZAFOldSample(1).kilovolts!

' Write absorbers column labels
msg$ = Space$(8) & vbTab
For i% = 1 To MAXELM%
msg$ = msg$ & Format$(VbDquote$ & Trim$(Symlo$(i%)) & VbDquote$, a80$) & vbTab
Next i%
Print #AFactorDataFileNumber%, msg$

' Loop on emitters
For j% = 1 To MAXELM%

' Loop on absorbers
msg$ = Format$(VbDquote$ & Trim$(Symlo$(j%)) & " " & Deflin$(j%) & VbDquote$, a80$) & vbTab
For i% = 1 To MAXELM%
If mode% = 1 Then msg$ = msg$ & Format$(Format$(alpha11(j%, i%), f85$), a80$) & vbTab
If mode% = 2 Then msg$ = msg$ & Format$(Format$(alpha21(j%, i%), f85$), a80$) & vbTab
If mode% = 3 Then msg$ = msg$ & Format$(Format$(alpha31(j%, i%), f85$), a80$) & vbTab
If mode% = 4 Then msg$ = msg$ & Format$(Format$(alpha41(j%, i%), f85$), a80$) & vbTab
Next i%
Print #AFactorDataFileNumber%, msg$

msg$ = Format$(VbDquote$ & Trim$(Symlo$(j%)) & " " & Deflin$(j%) & VbDquote$, a80$) & vbTab
For i% = 1 To MAXELM%
If mode% = 1 Then msg$ = msg$ & Format$(Format$(0#, f85$), a80$) & vbTab
If mode% = 2 Then msg$ = msg$ & Format$(Format$(alpha22(j%, i%), f85$), a80$) & vbTab
If mode% = 3 Then msg$ = msg$ & Format$(Format$(alpha32(j%, i%), f85$), a80$) & vbTab
If mode% = 4 Then msg$ = msg$ & Format$(Format$(alpha42(j%, i%), f85$), a80$) & vbTab
Next i%
Print #AFactorDataFileNumber%, msg$

msg$ = Format$(VbDquote$ & Trim$(Symlo$(j%)) & " " & Deflin$(j%) & VbDquote$, a80$) & vbTab
For i% = 1 To MAXELM%
If mode% = 1 Then msg$ = msg$ & Format$(Format$(0#, f85$), a80$) & vbTab
If mode% = 2 Then msg$ = msg$ & Format$(Format$(0#, f85$), a80$) & vbTab
If mode% = 3 Then msg$ = msg$ & Format$(Format$(alpha33(j%, i%), f85$), a80$) & vbTab
If mode% = 4 Then msg$ = msg$ & Format$(Format$(alpha43(j%, i%), f85$), a80$) & vbTab
Next i%
Print #AFactorDataFileNumber%, msg$

Next j%

' Close file
Close #AFactorDataFileNumber%

Exit Sub

' Errors
CalcZAFSaveFactorsError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFSaveFactors"
Close #AFactorDataFileNumber%
ierror = True
Exit Sub

End Sub

Sub CalcZAFSaveRatios(tfilename As String, wout() As Single, mode As Integer, filenamearray() As String)
' Save the calculated k-Ratios for MAXBIN compositions to ASCII files
'  mode = 0  output k-Ratio files (1 through MAXBINARY%)
'  mode = 4  output oxide end-member k-Ratios (single file)

ierror = False
On Error GoTo CalcZAFSaveRatiosError

Dim i As Integer, j As Integer, k As Integer
Dim astring As String

' Output k-Ratios for all MAXBIN concentrations for both emitting elements (MAXBINARY% *2)
If mode% = 0 Then
For k% = 1 To MAXBINARY%

' Open the k-ratio ASCII file
AFactorDataFile$ = MiscGetFileNameNoExtension$(tfilename$) & "_Kratio" & Format$(Int(wout!(2 * k% - 1))) & ".dat"
filenamearray$(k%) = AFactorDataFile$
Open AFactorDataFile$ For Output As #AFactorDataFileNumber%

' Write run info
astring$ = "Elemental K Ratios based on " & Format$(wout!(2 * k% - 1)) & " wt % of the emitting element"
Print #AFactorDataFileNumber%, VbDquote$ & astring$ & VbDquote$

astring$ = "ZAF Corr: " & zafstring$(izaf%) & ", MAC Table: " & macstring$(MACTypeFlag%)
Print #AFactorDataFileNumber%, VbDquote$ & astring$ & VbDquote$

Print #AFactorDataFileNumber%, VbDquote$ & Now & VbDquote$
Print #AFactorDataFileNumber%, CalcZAFOldSample(1).takeoff!, vbTab, CalcZAFOldSample(1).kilovolts!

' Write absorbers column labels
msg$ = Space$(8) & vbTab
For i% = 1 To MAXELM%
msg$ = msg$ & Format$(VbDquote$ & Trim$(Symlo$(i%)) & VbDquote$, a80$) & vbTab
Next i%
Print #AFactorDataFileNumber%, msg$

' Loop on emitters
For j% = 1 To MAXELM%

' Loop on absorbers
msg$ = Format$(VbDquote$ & Trim$(Symlo$(j%)) & " " & Deflin$(j%) & VbDquote$, a80$) & vbTab
For i% = 1 To MAXELM%
msg$ = msg$ & Format$(Format$(kfactor!(j%, i%, k%), f86$), a80$) & vbTab
Next i%
Print #AFactorDataFileNumber%, msg$

Next j%

' Close file
Close #AFactorDataFileNumber%
Next k%
End If

' Output oxide end-member k-Ratios for both emitting elements
If mode% = 4 Then
AFactorDataFile$ = MiscGetFileNameNoExtension$(tfilename$) & "_Oxide-Kratio" & ".dat"
filenamearray$(MAXBINARY% + 1) = AFactorDataFile$

' Open the alpha-factor ASCII file
Open AFactorDataFile$ For Output As #AFactorDataFileNumber%

' Write run info
astring$ = "Oxide end-member K Ratios (for calculation of oxide alpha-factors)"
Print #AFactorDataFileNumber%, VbDquote$ & astring$ & VbDquote$

astring$ = "ZAF Corr: " & zafstring$(izaf%) & vbCrLf & "MAC Table: " & macstring$(MACTypeFlag%)
Print #AFactorDataFileNumber%, VbDquote$ & astring$ & VbDquote$

Print #AFactorDataFileNumber%, VbDquote$ & Now & VbDquote$
Print #AFactorDataFileNumber%, CalcZAFOldSample(1).takeoff!, vbTab, CalcZAFOldSample(1).kilovolts!

' Write absorbers column labels
msg$ = Space$(8) & vbTab & Format$(VbDquote$ & "KRatio" & VbDquote$, a80$) & vbTab & Format$(VbDquote$ & "Oxide" & VbDquote$, a80$)
Print #AFactorDataFileNumber%, msg$

' Loop on emitters
For j% = 1 To MAXELM%
msg$ = Format$(VbDquote$ & Trim$(Symlo$(j%)) & " " & Deflin$(j%) & VbDquote$, a80$) & vbTab
msg$ = msg$ & Format$(Format$(oxfactor!(j%), f86$), a80$) & vbTab & Format$(VbDquote$ & ElementGetFormula$(j%) & VbDquote$, a80$)
Print #AFactorDataFileNumber%, msg$
Next j%

' Close file
Close #AFactorDataFileNumber%
End If

Exit Sub

' Errors
CalcZAFSaveRatiosError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFSaveRatios"
Close #AFactorDataFileNumber%
ierror = True
Exit Sub

End Sub

Sub CalcZAFGetComposition(mode As Integer)
' Get a composition
' mode = 1 get formula
' mode = 2 get weight
' mode = 3 get standard composition

ierror = False
On Error GoTo CalcZAFGetCompositionError

Dim i As Integer, ip As Integer
Dim astring As String

' Write space to log window for new composition
icancel = False
Call IOWriteLog(vbNullString)

' Get formula or weight from user
If mode% = 1 Then FormFORMULA.Show vbModal
If mode% = 2 Then FormWEIGHT.Show vbModal
If mode% = 3 Then FormSTDCOMP.Show vbModal

' If error, just clear and exit
If ierror Or icancel Then
Call InitSample(CalcZAFOldSample())
Exit Sub
End If

' Initialize
Call CalcZAFInit
If ierror Then Exit Sub

' Init sample
Call InitSample(CalcZAFOldSample())
If ierror Then Exit Sub
Call InitSample(CalcZAFTmpSample())
If ierror Then Exit Sub
Call InitSample(CalcZAFNewSample())
If ierror Then Exit Sub

' Return modified sample
Call FormulaReturnSample(CalcZAFOldSample())
If ierror Then Exit Sub

' Check elements returned
Call ElementGetData(CalcZAFOldSample())
If ierror Then Exit Sub

' Load string
For i% = 1 To CalcZAFOldSample(1).LastChan%
astring$ = astring$ & CalcZAFOldSample(1).Elsyms$(i%) & MiscAutoFormat$(CalcZAFOldSample(1).ElmPercents!(i%)) & " "
Next i%

Call IOWriteLog(astring$)

' Load particle parameters
CalcZAFOldSample(1).iptc% = iptc%
CalcZAFOldSample(1).PTCModel% = PTCModel%
CalcZAFOldSample(1).PTCDiameter! = PTCDiameter!
CalcZAFOldSample(1).PTCDensity! = PTCDensity!
CalcZAFOldSample(1).PTCThicknessFactor! = PTCThicknessFactor!
CalcZAFOldSample(1).PTCNumericalIntegrationStep! = PTCNumericalIntegrationStep!

' Load sample name
If mode% = 3 Then
CalcZAFOldSample(1).Type% = 1       ' standard
CalcZAFOldSample(1).Name$ = SampleGetString2$(CalcZAFOldSample())
Else
CalcZAFOldSample(1).Type% = 2       ' unknown
If CalcZAFSampleCount% = 0 Then CalcZAFSampleCount% = 1
CalcZAFOldSample(1).Name$ = CalcZAFOldSample(1).Name$ & ", sample" & Str$(CalcZAFSampleCount%)
End If

' Add standard to run
If mode% = 3 Then
ip% = IPOS2(NumberofStandards%, CalcZAFOldSample(1).number%, StandardNumbers%())
If ip% = 0 Then
Call AddStdSaveStd(CalcZAFOldSample(1).number%)
If ierror Then Exit Sub
End If
End If

' Update sample number if not standard
If CalcZAFOldSample(1).Type% <> 1 Then
CalcZAFOldSample(1).number% = CalcZAFSampleCount%
End If
CalcZAFOldSample(1).Linenumber&(1) = CalcZAFSampleCount%

' Re-set CalcZAF mode back to intensities from composition
CalcZAFMode% = 0
FormZAF.OptionCalculate(CalcZAFMode%).Value = True

' Sort elements
Call CalcZAFSave
If ierror Then Exit Sub

' Load form
FormZAF.Show vbModeless
Call CalcZAFLoad
If ierror Then Exit Sub

CalcZAFSampleCount% = CalcZAFSampleCount% + 1
Exit Sub

' Errors
CalcZAFGetCompositionError:
msg$ = ", (mode= " & Format$(mode%) & ")"
MsgBox Error$ & msg$, vbOKOnly + vbCritical, "CalcZAFGetComposition"
ierror = True
Exit Sub

End Sub

Sub CalcZAFCombinedConditions()
' Change conditions for a single element of the sample

ierror = False
On Error GoTo CalcZAFCombinedConditionsError

' Check if sample contains elements
If CalcZAFOldSample(1).LastElm% = 0 Then
msg$ = "The current sample contains no analyzed elements. Please add some analyzed elements and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFCombinedConditions"
ierror = True
Exit Sub
End If

' Load the form
Call Cond2Load(CalcZAFOldSample())
If ierror Then Exit Sub

' Load COND form
FormCOND2.Show vbModal
If icancelload Then Exit Sub

' Get the modified sample back
Call Cond2Return(CalcZAFOldSample())
If ierror Then Exit Sub

' Re-load sample element setup
If CalcZAFOldSample(1).LastElm% > 0 Then
Call ElementGetData(CalcZAFOldSample())
If ierror Then Exit Sub
End If

Exit Sub

' Errors
CalcZAFCombinedConditionsError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFCombinedConditions"
ierror = True
Exit Sub

End Sub

Sub CalcZAFGetExcel()
' Routine to call CalcZAFGetExcel2

ierror = False
On Error GoTo CalcZAFGetExcelError

Call CalcZAFGetExcel2(CalcZAFAnalysis, CalcZAFOldSample())
If ierror Then Exit Sub

Exit Sub

' Errors
CalcZAFGetExcelError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFGetExcel"
ierror = True
Exit Sub

End Sub

Sub CalcZAFReturnSample(sample() As TypeSample)
' Returns the current CalcZAF sample to the calling routine

ierror = False
On Error GoTo CalcZAFReturnSampleError

' Return the current sample
sample(1) = CalcZAFOldSample(1)

Exit Sub

' Errors
CalcZAFReturnSampleError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFReturnSample"
ierror = True
Exit Sub

End Sub

Sub CalcZAFElementLoad2()
' Re-loads the standard list for FormZAFELM

ierror = False
On Error GoTo CalcZAFElementLoad2Error

Dim i As Integer

' Load the primary assigned standard combo selections
FormZAFELM.ComboStandard.Clear
For i% = 1 To NumberofStandards%
msg$ = Format$(StandardNumbers(i%), a40) & " " & StandardNames$(i%)
FormZAFELM.ComboStandard.AddItem msg$
Next i%

Exit Sub

' Errors
CalcZAFElementLoad2Error:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFElementLoad2"
ierror = True
Exit Sub

End Sub

Sub CalcZAFCalculateAll(tForm As Form)
' Routine to perform all matrix corrections on a sample

ierror = False
On Error GoTo CalcZAFCalculateAllError

Dim nstring As String
Dim i As Integer, tzaf As Integer, j As Integer
Dim sum1 As Single, sum2 As Single

icancelanal = False

' Check if calculating all matrix corrections
tzaf% = izaf%
For j% = 1 To MAXZAF%   ' do not use 0 (individual selections)

' Increment matrix correction (if calculate all matrix corrections)
If CalculateAllMatrixCorrections Then
izaf% = j%
Call AnalyzeChangeZAF(CalcZAFAnalysis, CalcZAFOldSample, CalcZAFTmpSample)
If ierror Then
izaf% = tzaf%
Call InitGetZAFSetZAF2(izaf%)
ierror = True
Exit Sub
End If
End If

' Analyze the sample (check for calculating all matrix corrections if using ZAF or Phi-Rho-Z)
Call AnalyzeStatusAnal(vbNullString)
Call CalcZAFCalculate
If ierror Then
Call AnalyzeStatusAnal(vbNullString)
izaf% = tzaf%
Call InitGetZAFSetZAF2(izaf%)
ierror = True
Exit Sub
End If
Call AnalyzeStatusAnal(vbNullString)

' Load WtsData arrays with analysis data (assume single data row)
sum1! = 0#
sum2! = 0#
For i% = 1 To CalcZAFOldSample(1).LastChan%
CalcZAFAnalysis.WtsData!(1, i%) = CalcZAFAnalysis.WtPercents!(i%)
CalcZAFAnalysis.CalData!(1, i%) = CalcZAFAnalysis.AtPercents!(i%)
sum1! = sum1! + CalcZAFAnalysis.WtPercents!(i%)
sum2! = sum2! + CalcZAFAnalysis.AtPercents!(i%)
Next i%

' Add totals
CalcZAFAnalysis.WtPercents!(CalcZAFOldSample(1).LastChan% + 1) = sum1!
CalcZAFAnalysis.AtPercents!(CalcZAFOldSample(1).LastChan% + 1) = sum2!
CalcZAFAnalysis.WtsData!(1, CalcZAFOldSample(1).LastChan% + 1) = CalcZAFAnalysis.WtPercents!(CalcZAFOldSample(1).LastChan% + 1)   ' add total
CalcZAFAnalysis.CalData!(1, CalcZAFOldSample(1).LastChan% + 1) = CalcZAFAnalysis.AtPercents!(CalcZAFOldSample(1).LastChan% + 1)   ' add total

' Save all matrix correction results for export
nstring$ = CalcZAFOldSample(1).Name$
Call AnalyzeCalculateAllSave(ImportDataFile$, Int(1), nstring$, CalcZAFAnalysis, CalcZAFOldSample(), tForm)
If ierror Then
izaf% = tzaf%
Call InitGetZAFSetZAF2(izaf%)
ierror = True
Exit Sub
End If

Next j%     ' calculate all matrix corrections loop

' Restore original ZAF
izaf% = tzaf%
Call InitGetZAFSetZAF2(izaf%)
If ierror Then Exit Sub
AllAFactorUpdateNeeded = True

' Enable Excel button
If ExcelSheetIsOpen() Then
FormZAF.CommandExcel.Enabled = True
Else
FormZAF.CommandExcel.Enabled = False
End If

Call AnalyzeStatusAnal(vbNullString)
DoEvents
Exit Sub

' Errors
CalcZAFCalculateAllError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFCalculateAll"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

CalcZAFCalculateAllAnalysisInProgress:
msg$ = "Analysis calculation is already in progress, please try again later"
MsgBox msg$, vbOKOnly + vbExclamation, "CalcZAFCalculateAll"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

End Sub

Sub CalcZAFChangeZAF(analysis As TypeAnalysis, sample() As TypeSample, stdsample() As TypeSample)
' Change ZAF selections for calculating all matrix corrections

ierror = False
On Error GoTo CalcZAFChangeZAFError

' Load individual selections
Call InitGetZAFSetZAF2(izaf%)
If ierror Then Exit Sub

' Print current ZAF selections
Call TypeZAFSelections
If ierror Then Exit Sub

' Load element arrays
Call ElementGetData(sample())
If ierror Then Exit Sub

' Load primary intensities (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% <> MAXCORRECTION% Then
Call ZAFSetZAF(sample())
If ierror Then Exit Sub
Else
'Call ZAFSetZAF3(sample())
'If ierror Then Exit Sub
End If

' Update the standard kfacs based on changed conditions
Call UpdateAllStdKfacs(analysis, sample(), stdsample())
If ierror Then Exit Sub

' Force re-load of standard counts
'AllAnalysisUpdateNeeded = True

Exit Sub

' Errors
CalcZAFChangeZAFError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFChangeZAF"
ierror = True
Exit Sub

End Sub

Sub CalcZAFListCurrentAlphas()
' List alpha factors for sample

ierror = False
On Error GoTo CalcZAFListCurrentAlphasError

' Get the alphas and print them out
Call AFactorTypeAlphas(CalcZAFAnalysis, CalcZAFOldSample())
If ierror Then Exit Sub

Exit Sub

' Errors
CalcZAFListCurrentAlphasError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFListCurrentAlphas"
ierror = True
Exit Sub

End Sub

Sub CalcZAFSetEnables()
' Set form enables

ierror = False
On Error GoTo CalcZAFSetEnablesError

' Enable FormZAF buttons
If ImportDataFile$ <> vbNullString Or ImportDataFile2$ <> vbNullString Then
FormMAIN.menuFileOpen.Enabled = False
FormMAIN.menuFileClose.Enabled = True
'FormMAIN.menuFileExport.Enabled = True
FormMAIN.menuFileOpenAndProcess.Enabled = False
FormMAIN.menuFileUpdateCalcZAFSampleDataFiles.Enabled = False

If ImportDataFile$ <> vbNullString Then
FormZAF.CommandNext.Enabled = True
Else
FormZAF.CommandNext.Enabled = False
End If

FormZAF.OptionCalculate(0).Enabled = False
FormZAF.OptionCalculate(1).Enabled = False
FormZAF.OptionCalculate(2).Enabled = False
FormZAF.OptionCalculate(3).Enabled = False
FormZAF.CommandCompositionAtom.Enabled = False
FormZAF.CommandCompositionWeight.Enabled = False
FormZAF.CommandCompositionStandard.Enabled = False

Else
FormMAIN.menuFileOpen.Enabled = True
FormMAIN.menuFileClose.Enabled = False
'FormMAIN.menuFileExport.Enabled = False
FormMAIN.menuFileOpenAndProcess.Enabled = True
FormMAIN.menuFileUpdateCalcZAFSampleDataFiles.Enabled = True

FormZAF.CommandNext.Enabled = False

FormZAF.OptionCalculate(0).Enabled = True
FormZAF.OptionCalculate(1).Enabled = True
FormZAF.OptionCalculate(2).Enabled = True
FormZAF.OptionCalculate(3).Enabled = True
FormZAF.CommandCompositionAtom.Enabled = True
FormZAF.CommandCompositionWeight.Enabled = True
FormZAF.CommandCompositionStandard.Enabled = True
End If

Exit Sub

' Errors
CalcZAFSetEnablesError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFSetEnables"
ierror = True
Exit Sub

End Sub

Sub CalcZAFPlotAlphas()
' Load and plot alpha factors for the current CalcZAF sample

ierror = False
On Error GoTo CalcZAFPlotAlphasError

Call CalcZAFLoadAlphas_PE(CalcZAFOldSample())
If ierror Then Exit Sub

Exit Sub

' Errors
CalcZAFPlotAlphasError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFPlotAlphas"
ierror = True
Exit Sub

End Sub

Sub CalcZAFPlotHistogram(mode As Integer)
' Load and plot histogram for the current binary data set

ierror = False
On Error GoTo CalcZAFPlotHistogramError

Call CalcZAFPlotHistogram_PE(CalcZAFOutputCount&, KratioError!())
If ierror Then Exit Sub

If mode% > 0 Then
FormPLOTHISTO_PE.Show vbModal
End If

Exit Sub

' Errors
CalcZAFPlotHistogramError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFPlotHistogram"
ierror = True
Exit Sub

End Sub

Sub CalcZAFPlotHistogramConcentration()
' Load and plot concentration histogram for the current binary data set

ierror = False
On Error GoTo CalcZAFPlotHistogramConcentrationError

Call CalcZAFPlotHistogramConcentration_PE(CalcZAFOutputCount&, KratioConc!(), KratioError!(), KratioLine&())
If ierror Then Exit Sub

Exit Sub

' Errors
CalcZAFPlotHistogramConcentrationError:
MsgBox Error$, vbOKOnly + vbCritical, "CalcZAFPlotHistogramConcentration"
ierror = True
Exit Sub

End Sub


