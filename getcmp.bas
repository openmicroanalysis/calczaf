Attribute VB_Name = "CodeGETCMP"
' (c) Copyright 1995-2018 by John J. Donovan
Option Explicit

Dim GetCmpOldSample(1 To 1) As TypeSample
Dim GetCmpTmpSample(1 To 1) As TypeSample

Dim GetCmpAnalysis As TypeAnalysis

Dim GetCmpRow As Integer

Sub GetCmpLoadGrid()
' Load the Grid Column labels

ierror = False
On Error GoTo GetCmpLoadGridError

FormGETCMP.GridElementList.Clear

' Load grid
FormGETCMP.GridElementList.row = 0
FormGETCMP.GridElementList.col = 0
FormGETCMP.GridElementList.Text = "Channel"
FormGETCMP.GridElementList.col = 1
FormGETCMP.GridElementList.Text = "Element"
FormGETCMP.GridElementList.col = 2
FormGETCMP.GridElementList.Text = "X-Ray"

FormGETCMP.GridElementList.col = 3
FormGETCMP.GridElementList.Text = "Cations"
FormGETCMP.GridElementList.col = 4
FormGETCMP.GridElementList.Text = "Oxygens"
FormGETCMP.GridElementList.col = 5
FormGETCMP.GridElementList.Text = "Elemental"
FormGETCMP.GridElementList.col = 6
FormGETCMP.GridElementList.Text = "Oxide"
FormGETCMP.GridElementList.col = 7
FormGETCMP.GridElementList.Text = "Atomic"

' Update the oxide and atomic arrays
Call GetCmpUpdate
If ierror Then Exit Sub

' Update the totals columns
Call GetCmpUpdateTotals
If ierror Then Exit Sub

Exit Sub

' Errors
GetCmpLoadGridError:
MsgBox Error$, vbOKOnly + vbCritical, "GetCmpLoadGrid"
ierror = True
Exit Sub

End Sub

Sub GetCmpReturn(sample() As TypeSample)
' Return the saved sample

ierror = False
On Error GoTo GetCmpReturnError

sample(1) = GetCmpOldSample(1)
Exit Sub

' Errors
GetCmpReturnError:
MsgBox Error$, vbOKOnly + vbCritical, "GetCmpReturn"
ierror = True
Exit Sub

End Sub

Sub GetCmpSave()
' Save element list and composition changes back to GetCmpTmpSample arrays

ierror = False
On Error GoTo GetCmpSaveError

Dim i As Integer, ip As Integer, ipp As Integer
Dim sym As String
Dim sum As Single

' Check for valid standard number
If Val(FormGETCMP.TextNumber.Text) <= 0 Then GoTo GetCmpSaveNoNumber

' Check for non-blank standard name
If FormGETCMP.TextName.Text = vbNullString Then GoTo GetCmpSaveNoName

' Save name, number and description fields
GetCmpTmpSample(1).number% = Val(FormGETCMP.TextNumber.Text)
GetCmpTmpSample(1).Name$ = FormGETCMP.TextName.Text
GetCmpTmpSample(1).Description$ = FormGETCMP.TextDescription.Text

' Check for valid standard number
If GetCmpTmpSample(1).number% = MAXINTEGER% Then GoTo GetCmpSaveInvalidNumber

' Save the DisplayAsOxide option button
If FormGETCMP.OptionDisplayAsOxide.Value Then
GetCmpTmpSample(1).DisplayAsOxideFlag = True
Else
GetCmpTmpSample(1).DisplayAsOxideFlag = False
End If

' Save density
If Val(FormGETCMP.TextDensity.Text) <= 0# Or Val(FormGETCMP.TextDensity.Text) > MAXDENSITY# Then GoTo GetCmpSaveBadDensity
GetCmpTmpSample(1).SampleDensity! = Val(FormGETCMP.TextDensity.Text)

' Save material type (text string)
GetCmpTmpSample(1).MaterialType$ = Trim$(FormGETCMP.TextMaterialType.Text)

' New code to save formula options (06/15/2017)
If FormGETCMP.ComboFormula.ListCount > 0 Then
If FormGETCMP.ComboFormula.ListIndex > -1 Then
If Val(FormGETCMP.TextFormula.Text) > 0# Then
i% = FormGETCMP.ComboFormula.ListIndex     ' zero index indicates sum all cations
GetCmpTmpSample(1).FormulaRatio! = Val(FormGETCMP.TextFormula.Text)

' If element is not empty then it is a specific cation (no element indicates sum all cations)
If i% > 0 And i% <= GetCmpTmpSample(1).LastChan% Then
GetCmpTmpSample(1).FormulaElement$ = GetCmpTmpSample(1).Elsyms$(i%)
End If

End If
End If
End If

If FormGETCMP.CheckFormula.Value = vbChecked Then
GetCmpTmpSample(1).FormulaElementFlag% = True
Else
GetCmpTmpSample(1).FormulaElementFlag% = False
End If

' Check for no formula atoms
If GetCmpTmpSample(1).FormulaElementFlag% And GetCmpTmpSample(1).FormulaRatio! = 0# Then GoTo GetCmpSaveNoFormulaAtoms

' Check if formula concentration is too low
ip% = IPOS1(GetCmpTmpSample(1).LastChan%, GetCmpTmpSample(1).FormulaElement$, GetCmpTmpSample(1).Elsyms$())
If GetCmpTmpSample(1).ElmPercents!(ip%) < MinSpecifiedValue! Then GoTo GetCmpSaveInsufficientBasis

' Warn user if formula option is checked but no atoms is specified (blank element is ok since that indicates sum all cations)
If FormGETCMP.CheckFormula.Value = vbChecked And GetCmpTmpSample(1).FormulaRatio! = 0# Then
msg$ = "Formula option was selected, but no formula atoms or cation sum were specified"
MsgBox msg$, vbOKOnly + vbExclamation, "GetCmpSave"
ierror = True
Exit Sub
End If

' Load the element, xray, cation, oxygen and wtpercents lists into the GetCmpOldSample array to remove possible blank spaces
Call InitSample(GetCmpOldSample())
If ierror Then Exit Sub

For i% = 1 To MAXCHAN%
sym$ = GetCmpTmpSample(1).Elsyms$(i%)
ip% = IPOS1(MAXELM%, sym$, Symlo$())

sym$ = GetCmpTmpSample(1).Xrsyms$(i%)
ipp% = IPOS1(MAXRAY% - 1, sym$, Xraylo$())

If ip% > 0 And ipp% > 0 Then
    GetCmpOldSample(1).LastElm% = GetCmpOldSample(1).LastElm% + 1
    GetCmpOldSample(1).Elsyms$(GetCmpOldSample(1).LastElm%) = GetCmpTmpSample(1).Elsyms$(i%)
    GetCmpOldSample(1).Xrsyms$(GetCmpOldSample(1).LastElm%) = GetCmpTmpSample(1).Xrsyms$(i%)

    ' Make sure cations are loaded
    If GetCmpTmpSample(1).numcat%(i%) = 0 Or (GetCmpTmpSample(1).numcat%(i%) = 0 And GetCmpTmpSample(1).numoxd%(i%) = 0) Then
    GetCmpTmpSample(1).numcat%(i%) = AllCat%(ip%)
    GetCmpTmpSample(1).numoxd%(i%) = AllOxd%(ip%)
    End If
    GetCmpOldSample(1).numcat%(GetCmpOldSample(1).LastElm%) = GetCmpTmpSample(1).numcat%(i%)
    GetCmpOldSample(1).numoxd%(GetCmpOldSample(1).LastElm%) = GetCmpTmpSample(1).numoxd%(i%)
    GetCmpOldSample(1).AtomicCharges!(GetCmpOldSample(1).LastElm%) = GetCmpTmpSample(1).AtomicCharges!(i%)

    ' Load elemental compositions
    GetCmpOldSample(1).ElmPercents!(GetCmpOldSample(1).LastElm%) = GetCmpTmpSample(1).ElmPercents!(i%)
End If
Next i%

' Load other parameters
GetCmpOldSample(1).LastChan% = GetCmpOldSample(1).LastElm%

GetCmpOldSample(1).Type% = GetCmpTmpSample(1).Type%
GetCmpOldSample(1).Set% = GetCmpTmpSample(1).Set%
GetCmpOldSample(1).number% = GetCmpTmpSample(1).number%
GetCmpOldSample(1).Name$ = GetCmpTmpSample(1).Name$
GetCmpOldSample(1).Description$ = GetCmpTmpSample(1).Description$
GetCmpOldSample(1).DisplayAsOxideFlag% = GetCmpTmpSample(1).DisplayAsOxideFlag%
GetCmpOldSample(1).OxideOrElemental% = GetCmpTmpSample(1).OxideOrElemental%

GetCmpOldSample(1).SampleDensity! = GetCmpTmpSample(1).SampleDensity!

GetCmpOldSample(1).MaterialType$ = GetCmpTmpSample(1).MaterialType$

GetCmpOldSample(1).FormulaElementFlag = GetCmpTmpSample(1).FormulaElementFlag
GetCmpOldSample(1).FormulaRatio! = GetCmpTmpSample(1).FormulaRatio!
GetCmpOldSample(1).FormulaElement$ = GetCmpTmpSample(1).FormulaElement$

GetCmpOldSample(1).takeoff! = GetCmpTmpSample(1).takeoff!
GetCmpOldSample(1).kilovolts! = GetCmpTmpSample(1).kilovolts!

' Now reload GetCmpTmpSample
GetCmpTmpSample(1) = GetCmpOldSample(1)

' Check for valid sum if sample contains elements
If GetCmpTmpSample(1).LastChan% > 0 Then
sum! = 0#
For i% = 1 To GetCmpTmpSample(1).LastChan%
sum! = sum! + GetCmpTmpSample(1).ElmPercents!(i%)
Next i%
If sum! <= 0# Then GoTo GetCmpSaveBadSum
End If

Exit Sub

' Errors
GetCmpSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "GetCmpSave"
ierror = True
Exit Sub

GetCmpSaveNoNumber:
msg$ = "Standard number " & FormGETCMP.TextNumber.Text & " is invalid. Please enter an unused standard number from the standard list before entering the composition."
MsgBox msg$, vbOKOnly + vbExclamation, "GetCmpSave"
ierror = True
Exit Sub

GetCmpSaveBadSum:
msg$ = "Standard number " & Str$(GetCmpTmpSample(1).number%) & " sum is " & Str$(sum!) & ". Please modify the composition."
MsgBox msg$, vbOKOnly + vbExclamation, "GetCmpSave"
ierror = True
Exit Sub

GetCmpSaveNoName:
msg$ = "Standard number " & Str$(GetCmpTmpSample(1).number%) & " has a blank standard name"
MsgBox msg$, vbOKOnly + vbExclamation, "GetCmpSave"
ierror = True
Exit Sub

GetCmpSaveInvalidNumber:
msg$ = "Standard number " & Str$(GetCmpTmpSample(1).number%) & " is a reserved standard number and cannot be used for a standard."
MsgBox msg$, vbOKOnly + vbExclamation, "GetCmpSave"
ierror = True
Exit Sub

GetCmpSaveBadDensity:
msg$ = "Standard number " & Str$(GetCmpTmpSample(1).number%) & " has an invalid density value (must be between 0 and " & Format$(MAXDENSITY#) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "GetCmpSave"
ierror = True
Exit Sub

GetCmpSaveNoFormulaAtoms:
msg$ = "No formula atoms were specified. Either uncheck the Formula Element checkbox or specify the formula atoms."
MsgBox msg$, vbOKOnly + vbExclamation, "GetCmpSave"
ierror = True
Exit Sub

GetCmpSaveInsufficientBasis:
msg$ = GetCmpTmpSample(1).FormulaElement$ & " is not present in a sufficient concentration for the formula basis calculation."
MsgBox msg$, vbOKOnly + vbExclamation, "GetCmpSave"
ierror = True
Exit Sub

End Sub

Sub GetCmpSaveAll()
' Routine to handle new or modified standard composition

ierror = False
On Error GoTo GetCmpSaveAllError

Dim ip As Integer
Dim sym As String

' Save the changes to element list and or composition from Form GETCMP
Call GetCmpSave
If ierror Then Exit Sub

' Check for at least one element in standard
If GetCmpTmpSample(1).LastChan% < 1 Then GoTo GetCmpSaveAllNoElements

' Save excess oxygen if user entered any in text box
If GetCmpTmpSample(1).OxygenChannel% > 0 Then
GetCmpTmpSample(1).ElmPercents!(GetCmpTmpSample(1).OxygenChannel%) = GetCmpAnalysis.CalculatedOxygen! + Val(FormGETCMP.TextExcessOxygen.Text)
End If

' Check if oxygen is entered for oxide standards
If FormGETCMP.OptionDisplayAsOxide.Value Or FormGETCMP.OptionEnterOxide.Value Then
sym$ = "o"
ip% = IPOS1(GetCmpTmpSample(1).LastChan%, sym$, GetCmpTmpSample(1).Elsyms$())
If ip% = 0 Then GoTo GetCmpSaveAllNoOxygen
End If

' If adding a new or duplicate standard, update standard database
If GetCmpFlag% = 1 Or GetCmpFlag% = 3 Then

' First check for existing standard
ip% = StandardGetRow%(GetCmpTmpSample(1).number%)

' No existing standard
If ip% = 0 Then
Call StandardAddRecord(GetCmpTmpSample())
If ierror Then Exit Sub

' Found existing standard, ask for confirm to replace
Else
Call StandardReplaceRecord(GetCmpTmpSample())
If ierror Then Exit Sub
End If
End If

' If modifying standard, update standard database
If GetCmpFlag% = 2 Then
Call StandardReplaceRecord(GetCmpTmpSample())
If ierror Then Exit Sub
End If

Unload FormGETCMP
DoEvents

Exit Sub

' Errors
GetCmpSaveAllError:
MsgBox Error$, vbOKOnly + vbCritical, "GetCmpSaveAll"
ierror = True
Exit Sub

GetCmpSaveAllNoElements:
msg$ = "No elements entered for standard " & GetCmpTmpSample(1).Name$
MsgBox msg$, vbOKOnly + vbExclamation, "GetCmpSaveAll"
ierror = True
Exit Sub

GetCmpSaveAllNoOxygen:
msg$ = "Standard number " & Str$(GetCmpTmpSample(1).number%) & " was entered or displayed as an oxide standard but does not include oxygen as a compositional element. Please also enter the elemental oxygen composition (using the value displayed in the Total Oxygen From Cations field if necessary)"
MsgBox msg$, vbOKOnly + vbExclamation, "GetCmpSave"
ierror = True
Exit Sub

End Sub

Sub GetCmpSetCmpClear()
' Clear elements

ierror = False
On Error GoTo GetCmpSetCmpClearError

FormSETCMP.ComboElement.Text = vbNullString
FormSETCMP.ComboXray.Text = vbNullString
FormSETCMP.ComboCations.Text = vbNullString
FormSETCMP.ComboOxygens.Text = vbNullString
FormSETCMP.TextComposition.Text = vbNullString

FormSETCMP.TextCharge.Text = vbNullString
FormSETCMP.ComboCrystal.Text = vbNullString

Exit Sub

' Errors
GetCmpSetCmpClearError:
MsgBox Error$, vbOKOnly + vbCritical, "GetCmpSetCmpClear"
ierror = True
Exit Sub

End Sub

Sub GetCmpSetCmpLoad()
' Load the SETCMP form for the specified element and disables the element and xray fields if
' the sample already has data.

ierror = False
On Error GoTo GetCmpSetCmpLoadError

Dim i As Integer
Dim sym As String
Dim cat As Integer, oxd As Integer, num As Integer
Dim oxup As String, elup As String
Dim temp As Single

' Title the frame control
FormSETCMP.Frame1.Caption = "Enter Element Properties and Weight Percent For : " & MiscAutoUcase$(GetCmpTmpSample(1).Elsyms$(GetCmpRow%)) & " " & GetCmpTmpSample(1).Xrsyms$(GetCmpRow%)

' Label the composition label field
oxup$ = vbNullString
elup$ = vbNullString
If GetCmpTmpSample(1).Elsyms$(GetCmpRow%) <> vbNullString Then
sym$ = GetCmpTmpSample(1).Elsyms$(GetCmpRow%)
cat% = GetCmpTmpSample(1).numcat%(GetCmpRow%)
oxd% = GetCmpTmpSample(1).numoxd%(GetCmpRow%)
Call ElementGetSymbols(sym$, cat%, oxd%, num%, oxup$, elup$)
If ierror Then Exit Sub
End If

If FormGETCMP.OptionEnterOxide.Value Then msg$ = " Oxide Weight Percent " & oxup$
If FormGETCMP.OptionEnterElemental.Value Then msg$ = " Elemental Weight Percent " & elup$
FormSETCMP.LabelComposition.Caption = "Enter Composition in" & msg$

' Add the list box items
FormSETCMP.ComboElement.Clear
For i% = 0 To MAXELM% - 1
FormSETCMP.ComboElement.AddItem Symlo$(i% + 1)
Next i%

FormSETCMP.ComboXray.Clear
For i% = 0 To MAXRAY% - 2
FormSETCMP.ComboXray.AddItem Xraylo$(i% + 1)
Next i%

FormSETCMP.ComboCations.Clear
For i% = 1 To MAXCATION% - 1    ' 1 to 99
FormSETCMP.ComboCations.AddItem Format$(i%)
Next i%

FormSETCMP.ComboOxygens.Clear
For i% = 0 To MAXCATION% - 1    ' 0 to 99
FormSETCMP.ComboOxygens.AddItem Format$(i%)
Next i%

FormSETCMP.ComboCrystal.Clear
For i% = 1 To MAXCRYSTYPE%
If Trim$(AllCrystalNames$(i%)) <> vbNullString Then
FormSETCMP.ComboCrystal.AddItem AllCrystalNames$(i%)
End If
Next i%

' Load the current element properties
If GetCmpRow% > 0 Then
FormSETCMP.ComboElement.Text = GetCmpTmpSample(1).Elsyms$(GetCmpRow%)
FormSETCMP.ComboXray.Text = GetCmpTmpSample(1).Xrsyms$(GetCmpRow%)
FormSETCMP.ComboCations.Text = Format$(GetCmpTmpSample(1).numcat%(GetCmpRow%))
FormSETCMP.ComboOxygens.Text = Format$(GetCmpTmpSample(1).numoxd%(GetCmpRow%))
End If

' Disable cation combos if not oxide entry or save
FormSETCMP.ComboCations.Enabled = False
FormSETCMP.ComboOxygens.Enabled = False
If FormGETCMP.OptionEnterOxide.Value = True Or FormGETCMP.OptionDisplayAsOxide.Value Then
FormSETCMP.ComboCations.Enabled = True
FormSETCMP.ComboOxygens.Enabled = True
End If

FormSETCMP.ComboElement.Enabled = True
FormSETCMP.ComboXray.Enabled = True

' Load the composition field, convert to oxide if indicated
temp! = GetCmpTmpSample(1).ElmPercents!(GetCmpRow%)
If FormGETCMP.OptionEnterOxide.Value Then
temp! = ConvertElmToOxd(temp!, GetCmpTmpSample(1).Elsyms$(GetCmpRow%), GetCmpTmpSample(1).numcat%(GetCmpRow%), GetCmpTmpSample(1).numoxd%(GetCmpRow%))
End If

FormSETCMP.TextComposition.Text = Format$(temp!, f83$)
FormSETCMP.TextCharge.Text = Format$(GetCmpTmpSample(1).AtomicCharges!(GetCmpRow%), f83$)

Exit Sub

' Errors
GetCmpSetCmpLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "GetCmpSetCmpLoad"
ierror = True
Exit Sub

End Sub

Sub GetCmpSetCmpLoadElement(elementrow As Integer)
' Loads "GetCmpRow and load SETCMP form

ierror = False
On Error GoTo GetCmpSetCmpLoadElementError

' Load passed element row
GetCmpRow% = elementrow%

' Ask for element, xray and cations for all elements
If GetCmpRow% > 0 And GetCmpRow% <= MAXCHAN% Then
Call GetCmpSetCmpLoad
If ierror Then Exit Sub

FormSETCMP.Show vbModal
End If

Exit Sub

' Errors
GetCmpSetCmpLoadElementError:
MsgBox Error$, vbOKOnly + vbCritical, "GetCmpSetCmpLoadElement"
ierror = True
Exit Sub

End Sub

Sub GetCmpSetCmpSave()
' This routine saves the element and xray symbols entered from the combo boxes

ierror = False
On Error GoTo GetCmpSetCmpSaveError

Dim sym As String
Dim i As Integer
Dim ip As Integer, ipp As Integer
Dim temp As Single
Dim keV As Single, lam As Single

ReDim numbers(1 To MAXCATION%) As Integer

For i% = 1 To MAXCATION%
numbers(i%) = MAXCATION% - i%   ' load 0 to MAXCATION% - 1 in decreasing value order
Next i%

' Get the element symbol
GetCmpTmpSample(1).Elsyms$(GetCmpRow%) = vbNullString
sym$ = FormSETCMP.ComboElement.Text

' Check for blank (deleted) element
If sym$ = vbNullString Then
GetCmpTmpSample(1).Elsyms$(GetCmpRow%) = vbNullString
GetCmpTmpSample(1).Xrsyms$(GetCmpRow%) = vbNullString
GetCmpTmpSample(1).numcat%(GetCmpRow%) = 0
GetCmpTmpSample(1).numoxd%(GetCmpRow%) = 0
GetCmpTmpSample(1).ElmPercents!(GetCmpRow%) = 0#
Exit Sub
End If

' Warn user if invalid symbol
ipp% = IPOS1(MAXELM%, sym$, Symlo$())
If ipp% = 0 Then GoTo GetCmpSetCmpBadElement

' Check that the element is not already entered
ip% = IPOS1(MAXCHAN%, sym$, GetCmpTmpSample(1).Elsyms$())
If ip% <> 0 And ip% <> GetCmpRow% Then GoTo GetCmpSetCmpDuplicateElement
GetCmpTmpSample(1).Elsyms$(GetCmpRow%) = sym$

' Check for a valid x-ray symbol
sym$ = FormSETCMP.ComboXray.Text
ip% = IPOS1(MAXRAY% - 1, sym$, Xraylo$())
If ip% = 0 Then GoTo GetCmpSetCmpBadXray
GetCmpTmpSample(1).Xrsyms$(GetCmpRow%) = sym$

' Save the xray line as a default for this element
Deflin$(ipp%) = GetCmpTmpSample(1).Xrsyms$(GetCmpRow%)

' Determine energy and wavelength based on element and xray (skip hydrogen and helium)
If Symlo$(ipp%) <> Symlo$(ATOMIC_NUM_HYDROGEN%) And Symlo$(ipp%) <> Symlo$(ATOMIC_NUM_HELIUM%) Then
Call XrayGetKevLambda(Symlo$(ipp%), Xraylo$(ip%), keV!, lam!)
If ierror Then Exit Sub
End If

' Save the cation subscripts
i% = Val(FormSETCMP.ComboCations.Text)
ip% = IPOS2(MAXCATION% - 1, i%, numbers%())     ' zero is invalid
If ip% = 0 Then GoTo GetCmpSetCmpSaveBadCation
GetCmpTmpSample(1).numcat%(GetCmpRow%) = i%

' Save cation as default if oxide standard
If GetCmpTmpSample(1).OxideOrElemental% = 1 Or GetCmpTmpSample(1).DisplayAsOxideFlag Then
AllCat%(ipp%) = GetCmpTmpSample(1).numcat%(GetCmpRow%)
End If

' Save number of oxygens (must be zero to nine)
i% = Val(FormSETCMP.ComboOxygens.Text)
ip% = IPOS2(MAXCATION%, i%, numbers%())         ' zero is valid
If ip% = 0 Then GoTo GetCmpSetCmpSaveBadOxygen
GetCmpTmpSample(1).numoxd%(GetCmpRow%) = i%

' Save oxygens as default if oxide standard
If GetCmpTmpSample(1).OxideOrElemental% = 1 Or GetCmpTmpSample(1).DisplayAsOxideFlag Then
AllOxd%(ipp%) = GetCmpTmpSample(1).numoxd%(GetCmpRow%)
End If

' Save crystal as default for this session
sym$ = Trim$(FormSETCMP.ComboCrystal.Text)
ip% = IPOS1%(MAXCRYSTYPE%, sym$, AllCrystalNames$())
If ip% = 0 Then GoTo GetCmpSetCmpBadCrystal
GetCmpTmpSample(1).CrystalNames$(GetCmpRow%) = AllCrystalNames$(ip%)
Defcry$(ipp%) = GetCmpTmpSample(1).CrystalNames$(GetCmpRow%)

' Load weight percent values
temp! = Val(FormSETCMP.TextComposition.Text)

' Convert if oxide
If FormGETCMP.OptionEnterOxide.Value Then
temp! = ConvertOxdToElm(temp!, GetCmpTmpSample(1).Elsyms$(GetCmpRow%), GetCmpTmpSample(1).numcat%(GetCmpRow%), GetCmpTmpSample(1).numoxd%(GetCmpRow%))
End If

GetCmpTmpSample(1).ElmPercents!(GetCmpRow%) = temp!

' Save charge and save as default
If Val(FormSETCMP.TextCharge.Text) < -10 Or Val(FormSETCMP.TextCharge.Text) > 10 Then GoTo GetCmpSetCmpSaveBadCharge
GetCmpTmpSample(1).AtomicCharges!(GetCmpRow%) = Val(FormSETCMP.TextCharge.Text)
AllAtomicCharges!(ipp%) = GetCmpTmpSample(1).AtomicCharges!(GetCmpRow%)

Exit Sub

' Errors
GetCmpSetCmpSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "GetCmpSetCmpSave"
ierror = True
Exit Sub

GetCmpSetCmpBadElement:
msg$ = "Element " & sym$ & " is not a valid element symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "GetCmpSetCmpSave"
ierror = True
Exit Sub

GetCmpSetCmpDuplicateElement:
msg$ = "Element " & sym$ & " is already entered in the standard composition"
MsgBox msg$, vbOKOnly + vbExclamation, "GetCmpSetCmpSave"
ierror = True
Exit Sub

GetCmpSetCmpBadXray:
msg$ = "Xray " & sym$ & " is not a valid xray symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "GetCmpSetCmpSave"
ierror = True
Exit Sub

GetCmpSetCmpSaveBadCation:
msg$ = "Invalid number of cations"
MsgBox msg$, vbOKOnly + vbExclamation, "GetCmpSetCmpSave"
ierror = True
Exit Sub

GetCmpSetCmpSaveBadOxygen:
msg$ = "Invalid number of oxygens"
MsgBox msg$, vbOKOnly + vbExclamation, "GetCmpSetCmpSave"
ierror = True
Exit Sub

GetCmpSetCmpBadCrystal:
msg$ = "Invalid crystal"
MsgBox msg$, vbOKOnly + vbExclamation, "GetCmpSetCmpSave"
ierror = True
Exit Sub

GetCmpSetCmpSaveBadCharge:
msg$ = "Invalid charge"
MsgBox msg$, vbOKOnly + vbExclamation, "GetCmpSetCmpSave"
ierror = True
Exit Sub

End Sub

Sub GetCmpSetCmpUpdateCombo()
' Update the xray and cation fields based on element

ierror = False
On Error GoTo GetCmpSetCmpUpdateComboError

Dim ip As Integer, ipp As Integer
Dim sym As String

sym$ = FormSETCMP.ComboElement.Text
ip% = IPOS1(MAXELM%, sym$, Symlo$())

If ip% > 0 Then
If FormSETCMP.ComboXray.Text = vbNullString Then FormSETCMP.ComboXray.Text = Deflin$(ip%)
If sym$ <> GetCmpTmpSample(1).Elsyms$(GetCmpRow%) Then FormSETCMP.ComboXray.Text = Deflin$(ip%)

If FormSETCMP.ComboCations.Text = vbNullString Then FormSETCMP.ComboCations.Text = AllCat%(ip%)
If sym$ <> GetCmpTmpSample(1).Elsyms$(GetCmpRow%) Then FormSETCMP.ComboCations.Text = AllCat%(ip%)

If FormSETCMP.ComboOxygens.Text = vbNullString Then FormSETCMP.ComboOxygens.Text = AllOxd%(ip%)
If sym$ <> GetCmpTmpSample(1).Elsyms$(GetCmpRow%) Then FormSETCMP.ComboOxygens.Text = AllOxd%(ip%)

If Val(FormSETCMP.TextCharge.Text) = 0 Then FormSETCMP.TextCharge.Text = Str$(AllAtomicCharges!(ip%))
If sym$ <> GetCmpTmpSample(1).Elsyms$(GetCmpRow%) Then FormSETCMP.TextCharge.Text = Str$(AllAtomicCharges!(ip%))

' Select the default crystal
ipp% = IPOS1(MAXCRYSTYPE%, Defcry$(ip%), AllCrystalNames$())
If ipp% > 0 Then FormSETCMP.ComboCrystal.Text = AllCrystalNames$(ipp%)

End If

Exit Sub

' Errors
GetCmpSetCmpUpdateComboError:
MsgBox Error$, vbOKOnly + vbCritical, "GetCmpSetCmpUpdateCombo"
ierror = True
Exit Sub

End Sub

Sub GetCmpUpdate()
' This routine updates the current element list grid "GetCmpRow%" based on the GetCmpTmpSample arrays

ierror = False
On Error GoTo GetCmpUpdateError

Dim ip As Integer, chan As Integer
Dim sym As String
Dim temp1 As Single, temp2 As Single

' Calculate "GetCmpTmpSample(1).OxygenChannel%" for elemental to oxide conversions
GetCmpTmpSample(1).OxygenChannel% = 0
sym$ = "o"
ip% = IPOS1(GetCmpTmpSample(1).LastChan%, sym$, GetCmpTmpSample(1).Elsyms$())
GetCmpTmpSample(1).OxygenChannel% = ip%

' Calculate oxides and atomic percents
Call StanFormCalculateOxideAtomic(GetCmpTmpSample())
If ierror Then Exit Sub

' Subtract calculated from total oxygen to calculate excess
GetCmpOldSample(1) = GetCmpTmpSample(1)
Call StanFormCalculateExcessOxygen(GetCmpAnalysis, GetCmpOldSample(), GetCmpTmpSample())
If ierror Then Exit Sub

' Update excess oxygen text box
FormGETCMP.TextExcessOxygen.Text = vbNullString
If FormGETCMP.OptionDisplayAsOxide.Value Then
FormGETCMP.TextExcessOxygen.Text = Format$(Format$(GetCmpAnalysis.ExcessOxygen!, f83$), a80$)
End If

' Update grid
For chan% = 1 To GetCmpTmpSample(1).LastChan%
FormGETCMP.GridElementList.row = chan%
FormGETCMP.GridElementList.col = 0
FormGETCMP.GridElementList.Text = Format$(chan%)
FormGETCMP.GridElementList.col = 1
FormGETCMP.GridElementList.Text = GetCmpTmpSample(1).Elsyms$(chan%)
FormGETCMP.GridElementList.col = 2
FormGETCMP.GridElementList.Text = GetCmpTmpSample(1).Xrsyms$(chan%)

' Show cations and oxygens if oxide entry or display
If FormGETCMP.OptionDisplayAsOxide.Value = True Or FormGETCMP.OptionEnterOxide.Value Then
FormGETCMP.GridElementList.col = 3
FormGETCMP.GridElementList.Text = Format$(GetCmpTmpSample(1).numcat%(chan%))
FormGETCMP.GridElementList.col = 4
FormGETCMP.GridElementList.Text = Format$(GetCmpTmpSample(1).numoxd%(chan%))
Else
FormGETCMP.GridElementList.col = 3
FormGETCMP.GridElementList.Text = vbNullString
FormGETCMP.GridElementList.col = 4
FormGETCMP.GridElementList.Text = vbNullString
End If
FormGETCMP.GridElementList.col = 5
FormGETCMP.GridElementList.Text = Format$(Format$(GetCmpTmpSample(1).ElmPercents!(chan%), f83$), a80$)

' Calculate oxide composition if sample is display as oxide or enter as oxide
FormGETCMP.GridElementList.col = 6
FormGETCMP.GridElementList.Text = vbNullString

If GetCmpTmpSample(1).DisplayAsOxideFlag Or FormGETCMP.OptionEnterOxide.Value = True Or FormGETCMP.OptionDisplayAsOxide.Value Then
FormGETCMP.GridElementList.col = 6
FormGETCMP.GridElementList.Text = Format$(Format$(OxPercents!(chan%), f83$), a80$)
End If

' Calculate atomic percents for all elements
FormGETCMP.GridElementList.col = 7
FormGETCMP.GridElementList.Text = Format$(Format$(AtPercents!(chan%), f83$), a80$)

Next chan%

' Update oxygen from cations for display (even if no oxygen is entered yet)
FormGETCMP.LabelOxygenFromCations.Caption = vbNullString
If FormGETCMP.OptionDisplayAsOxide.Value Then
temp1! = ConvertOxygenFromCations(GetCmpTmpSample())
FormGETCMP.LabelOxygenFromCations.Caption = Format$(Format$(temp1!, f83$), a80$)
End If

' Calculate oxygen equivalent of halogens
FormGETCMP.LabelOxygenFromHalogens.Caption = vbNullString
If FormGETCMP.OptionDisplayAsOxide.Value Then
temp2! = ConvertHalogensToOxygen(GetCmpTmpSample(1).LastChan%, GetCmpTmpSample(1).Elsyms$(), GetCmpTmpSample(1).DisableQuantFlag%(), GetCmpTmpSample(1).ElmPercents!())
FormGETCMP.LabelOxygenFromHalogens.Caption = Format$(Format$(temp2!, f83$), a80$)
End If

FormGETCMP.LabelHalogenCorrectedOxygen.Caption = vbNullString
If FormGETCMP.OptionDisplayAsOxide.Value Then
FormGETCMP.LabelHalogenCorrectedOxygen.Caption = Format$(Format$(temp1! - temp2!, f83$), a80$)
End If

Exit Sub

' Errors
GetCmpUpdateError:
MsgBox Error$, vbOKOnly + vbCritical, "GetCmpUpdate"
ierror = True
Exit Sub

End Sub

Sub GetCmpUpdateTotals()
' Update the total text fields in FormGETCMP

ierror = False
On Error GoTo GetCmpTotalsError

Dim i As Integer
Dim sum1 As Single, sum2 As Single, sum3 As Single

sum1! = 0#
sum2! = 0#
sum3! = 0#

For i% = 1 To GetCmpTmpSample(1).LastChan%
sum1! = sum1! + GetCmpTmpSample(1).ElmPercents!(i%)
sum2! = sum2! + OxPercents!(i%)
sum3! = sum3! + AtPercents!(i%)
Next i%

FormGETCMP.LabelElemental.Caption = Format$(Format$(sum1!, f83$), a80$)
FormGETCMP.LabelOxide.Caption = vbNullString
If FormGETCMP.OptionDisplayAsOxide.Value Then
FormGETCMP.LabelOxide.Caption = Format$(Format$(sum2!, f83$), a80$)
End If
FormGETCMP.LabelAtomic.Caption = Format$(Format$(sum3!, f83$), a80$)

Exit Sub

' Errors
GetCmpTotalsError:
MsgBox Error$, vbOKOnly + vbCritical, "GetCmpTotals"
ierror = True
Exit Sub

End Sub

Sub GetCmpChangedExcess()
' Called when user (or program) changes the excess oxygen text box

ierror = False
On Error GoTo GetCmpChangedExcessError

' Save excess oxygen if user entered any
If GetCmpTmpSample(1).OxygenChannel% > 0 Then
GetCmpTmpSample(1).ElmPercents!(GetCmpTmpSample(1).OxygenChannel%) = GetCmpAnalysis.CalculatedOxygen! + Val(FormGETCMP.TextExcessOxygen.Text)
End If

' Update grid
Call GetCmpUpdate
If ierror Then Exit Sub

' Update totals
Call GetCmpUpdateTotals
If ierror Then Exit Sub

Exit Sub

' Errors
GetCmpChangedExcessError:
MsgBox Error$, vbOKOnly + vbCritical, "GetCmpChangedExcess"
ierror = True
Exit Sub

End Sub

Sub GetCmpLoad(sample() As TypeSample)
' Loads the GetCmp form based on "GetCmpFlag%"
' GetCmpFlag = 1 loads a NEW standard composition (allow user to change standard number)
' GetCmpFlag = 2 loads a MODIFIED standard composition (no change of standard number)
' GetCmpFlag = 3 loads a DUPLICATE standard composition (allow user to change standard number)

ierror = False
On Error GoTo GetCmpLoadError

Dim i As Integer, ip As Integer

' Initialize standard calculation arrays
Call InitStandards(GetCmpAnalysis)
If ierror Then Exit Sub
Call InitLine(GetCmpAnalysis)
If ierror Then Exit Sub

' Load the passed sample
GetCmpTmpSample(1) = sample(1)

' If adding a standard, create next standard number
If GetCmpFlag% = 1 Then
GetCmpTmpSample(1).number% = StandardGetNumber%()
End If

' Disable Atomic and Formula entry if modifying or duplicating a composition
'If GetCmpFlag% = 2 Or GetCmpFlag% = 3 Then
'FormGETCMP.CommandEnterAtomFormula.Enabled = False
'Else
'FormGETCMP.CommandEnterAtomFormula.Enabled = True
'End If

' Disable number field if modifying a standard
FormGETCMP.TextNumber.Enabled = True
FormGETCMP.TextName.Enabled = True
FormGETCMP.TextDescription.Enabled = True

If GetCmpFlag% = 2 Then
FormGETCMP.TextNumber.Enabled = False
End If

' Initialize the Element List Grid Width
For i% = 0 To FormGETCMP.GridElementList.cols - 1
FormGETCMP.GridElementList.ColWidth(i%) = (FormGETCMP.GridElementList.Width - SCROLLBARWIDTH%) / FormGETCMP.GridElementList.cols
Next i%

Call GetCmpLoadGrid
If ierror Then Exit Sub

' Initialize sample number, name and description fields
FormGETCMP.TextNumber.Text = Str$(GetCmpTmpSample(1).number%)
FormGETCMP.TextDescription.Text = GetCmpTmpSample(1).Description$
FormGETCMP.TextName.Text = GetCmpTmpSample(1).Name$

FormGETCMP.TextDensity.Text = MiscAutoFormat$(GetCmpTmpSample(1).SampleDensity!)

FormGETCMP.TextMaterialType.Text = Trim$(GetCmpTmpSample(1).MaterialType$)

' Load formula calculation controls (new code 06/15/2017)
FormGETCMP.ComboFormula.Clear
FormGETCMP.ComboFormula.AddItem "Sum"  ' zero index indicates sum all cations
For i% = 1 To GetCmpTmpSample(1).LastChan%
FormGETCMP.ComboFormula.AddItem GetCmpTmpSample(1).Elsyms$(i%)
Next i%

If GetCmpTmpSample(1).FormulaElementFlag% Then
FormGETCMP.CheckFormula.Value = vbChecked
FormGETCMP.TextFormula.Enabled = True
FormGETCMP.ComboFormula.Enabled = True
Else
FormGETCMP.CheckFormula.Value = vbUnchecked
FormGETCMP.TextFormula.Enabled = False
FormGETCMP.ComboFormula.Enabled = False
End If

FormGETCMP.ComboFormula.ListIndex = 0  ' default to sum of cations
If GetCmpTmpSample(1).FormulaRatio! > 0# Then FormGETCMP.TextFormula.Text = Format$(GetCmpTmpSample(1).FormulaRatio!)
If GetCmpTmpSample(1).FormulaElement$ <> vbNullString Then
ip% = IPOS1(GetCmpTmpSample(1).LastChan%, GetCmpTmpSample(1).FormulaElement$, GetCmpTmpSample(1).Elsyms$())
If ip% > 0 Then
FormGETCMP.ComboFormula.ListIndex = ip%
End If
End If

' Initialize the Enter As option buttons (if sample has elements)
If GetCmpTmpSample(1).LastElm > 0 Then
If GetCmpTmpSample(1).DisplayAsOxideFlag Then
FormGETCMP.OptionEnterOxide.Value = True
Else
FormGETCMP.OptionEnterElemental.Value = True
End If

' Otherwise use INI file default
Else
If DefaultOxideOrElemental% = 1 Then
FormGETCMP.OptionEnterOxide.Value = True
Else
FormGETCMP.OptionEnterElemental.Value = True
End If
End If

' Initialize the Display As Oxide option buttons (last) if sample has elements
If GetCmpTmpSample(1).LastElm% > 0 Then
If GetCmpTmpSample(1).DisplayAsOxideFlag Then
FormGETCMP.OptionDisplayAsOxide.Value = True
Else
FormGETCMP.OptionNotDisplayAsOxide.Value = True
End If

' Otherwise use INI file default
Else
If DefaultOxideOrElemental% = 1 Then
FormGETCMP.OptionDisplayAsOxide.Value = True
Else
FormGETCMP.OptionNotDisplayAsOxide.Value = True
End If
End If

' Reload list of spectra
Call GetCmpLoadSpectra(Int(0))
If ierror Then Exit Sub
Call GetCmpLoadSpectra(Int(1))
If ierror Then Exit Sub

Exit Sub

' Errors
GetCmpLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "GetCmpLoad"
ierror = True
Exit Sub

End Sub

Sub GetCmpLoadFormula()
' Loads a formula composition for a new standard

ierror = False
On Error GoTo GetCmpLoadFormulaError

Dim response As Integer

' Check for existing element data
If GetCmpFlag% = 2 Or GetCmpFlag% = 3 Then
msg$ = "Using atomic formula data entry will cause the current standard composition to be overwritten. Are you sure that you want to replace the standard composition?"
response% = MsgBox(msg$, vbOKCancel + vbInformation + vbDefaultButton2, "GetCmpLoadFormula")
If response% = vbCancel Then Exit Sub
End If

' Load FormFORMULA
FormFORMULA.Frame1.Caption = "Enter Formula String For : " & GetCmpTmpSample(1).Name$

' Get formula from user
FormFORMULA.Show vbModal
If icancel Then Exit Sub

' Return modified sample
Call FormulaReturnSample(GetCmpTmpSample())
If ierror Then Exit Sub

' Remove blank rows
Call GetCmpSave
If ierror Then Exit Sub

' Reload the entire grid
Call GetCmpLoadGrid
If ierror Then Exit Sub

Exit Sub

' Errors
GetCmpLoadFormulaError:
MsgBox Error$, vbOKOnly + vbCritical, "GetCmpLoadFormula"
ierror = True
Exit Sub

End Sub

Sub GetCmpCalculateDensity()
' Perform a crude density calculation

ierror = False
On Error GoTo GetCmpCalculateDensityError

Dim i As Integer

Dim atoms(1 To MAXCHAN%) As Single
Dim tvol As Single, vols(1 To MAXCHAN%) As Single

' Load element data
Call ElementGetData(GetCmpTmpSample())
If ierror Then Exit Sub

' Calculate density based on composition
Call ConvertWeightToAtomic(GetCmpTmpSample(1).LastChan%, GetCmpTmpSample(1).AtomicWts!(), GetCmpTmpSample(1).ElmPercents!(), atoms!())
If ierror Then Exit Sub

' Calculate atomic volume total
tvol! = 0#
For i% = 1 To GetCmpTmpSample(1).LastChan%
tvol! = tvol! + atoms!(i%) * AllAtomicVolumes!(GetCmpTmpSample(1).AtomicNums%(i%))
Next i%

' Calculate atomic volume fractions
For i% = 1 To GetCmpTmpSample(1).LastChan%
vols!(i%) = (atoms!(i%) * AllAtomicVolumes!(GetCmpTmpSample(1).AtomicNums%(i%))) / tvol!
Next i%

' Calculate density
GetCmpTmpSample(1).SampleDensity! = 0#
For i% = 1 To GetCmpTmpSample(1).LastChan%
GetCmpTmpSample(1).SampleDensity! = GetCmpTmpSample(1).SampleDensity! + atoms!(i%) * AllAtomicDensities3!(GetCmpTmpSample(1).AtomicNums%(i%))
Next i%

FormGETCMP.TextDensity.Text = MiscAutoFormat$(GetCmpTmpSample(1).SampleDensity!)
Exit Sub

' Errors
GetCmpCalculateDensityError:
MsgBox Error$, vbOKOnly + vbCritical, "GetCmpCalculateDensity"
ierror = True
Exit Sub

End Sub

Sub GetCmpImportSpectrum(mode As Integer, tForm As Form)
' Import an EDS or CL spectrum into the standard database

ierror = False
On Error GoTo GetCmpImportEDSSpectrumError

' Call the import EDS spectrum
If mode% = 0 Then
Call StandardImportEDSSpectrum(GetCmpTmpSample(1).number%, tForm, GetCmpTmpSample())
If ierror Then Exit Sub
End If

' Call the import CL spectrum
If mode% = 1 Then
Call StandardImportCLSpectrum(GetCmpTmpSample(1).number%, tForm, GetCmpTmpSample())
If ierror Then Exit Sub
End If

' Reload list of spectra
Call GetCmpLoadSpectra(mode%)
If ierror Then Exit Sub

Exit Sub

' Errors
GetCmpImportEDSSpectrumError:
MsgBox Error$, vbOKOnly + vbCritical, "GetCmpImportEDSSpectrum"
ierror = True
Exit Sub

End Sub

Sub GetCmpExportSpectrum(mode As Integer, tForm As Form)
' Export an EDS or CL spectrum into the standard database

ierror = False
On Error GoTo GetCmpExportEDSSpectrumError

Dim specnum As Integer

' Call the export EDS spectrum
If mode% = 0 Then
If FormGETCMP.ListEDSSpectra.ListCount < 1 Then Exit Sub
If FormGETCMP.ListEDSSpectra.ListIndex < 0 Then Exit Sub
specnum% = FormGETCMP.ListEDSSpectra.ItemData(FormGETCMP.ListEDSSpectra.ListIndex)

'Call StandardExportEDSSpectrum(GetCmpTmpSample(1).number%, specnum%, tForm, GetCmpTmpSample())
If ierror Then Exit Sub
End If

' Call the export CL spectrum
If mode% = 1 Then
If FormGETCMP.ListCLSpectra.ListCount < 1 Then Exit Sub
If FormGETCMP.ListCLSpectra.ListIndex < 0 Then Exit Sub
specnum% = FormGETCMP.ListCLSpectra.ItemData(FormGETCMP.ListCLSpectra.ListIndex)

'Call StandardExportCLSpectrum(GetCmpTmpSample(1).number%, specnum%, tForm, GetCmpTmpSample())
If ierror Then Exit Sub
End If

Exit Sub

' Errors
GetCmpExportEDSSpectrumError:
MsgBox Error$, vbOKOnly + vbCritical, "GetCmpExportEDSSpectrum"
ierror = True
Exit Sub

End Sub

Sub GetCmpDeleteSpectrum(mode As Integer)
' Delete an EDS or CL spectrum into the standard database

ierror = False
On Error GoTo GetCmpDeleteEDSSpectrumError

Dim specnum As Integer

' Call the delete EDS spectrum
If mode% = 0 Then
If FormGETCMP.ListEDSSpectra.ListCount < 1 Then Exit Sub
If FormGETCMP.ListEDSSpectra.ListIndex < 0 Then Exit Sub
specnum% = FormGETCMP.ListEDSSpectra.ItemData(FormGETCMP.ListEDSSpectra.ListIndex)

Call StandardDeleteEDSSpectrum(GetCmpTmpSample(1).number%, specnum%)
If ierror Then Exit Sub
End If

' Call the delete CL spectrum
If mode% = 1 Then
If FormGETCMP.ListCLSpectra.ListCount < 1 Then Exit Sub
If FormGETCMP.ListCLSpectra.ListIndex < 0 Then Exit Sub
specnum% = FormGETCMP.ListCLSpectra.ItemData(FormGETCMP.ListCLSpectra.ListIndex)

Call StandardDeleteCLSpectrum(GetCmpTmpSample(1).number%, specnum%)
If ierror Then Exit Sub
End If

' Reload list of spectra
Call GetCmpLoadSpectra(mode%)
If ierror Then Exit Sub

Exit Sub

' Errors
GetCmpDeleteEDSSpectrumError:
MsgBox Error$, vbOKOnly + vbCritical, "GetCmpDeleteEDSSpectrum"
ierror = True
Exit Sub

End Sub

Sub GetCmpDisplaySpectrum(mode As Integer)
' Display an EDS or CL spectrum into the standard database

ierror = False
On Error GoTo GetCmpDisplaySpectrumError

Dim specnum As Integer

' Display EDS spectrum
If mode% = 0 Then
If FormGETCMP.ListEDSSpectra.ListCount < 1 Then Exit Sub
If FormGETCMP.ListEDSSpectra.ListIndex < 0 Then Exit Sub
specnum% = FormGETCMP.ListEDSSpectra.ItemData(FormGETCMP.ListEDSSpectra.ListIndex)

Call StandardDisplayEDSSpectrum(GetCmpTmpSample(1).number%, specnum%, GetCmpTmpSample())
If ierror Then Exit Sub
End If

' Display CL spectrum
If mode% = 1 Then
If FormGETCMP.ListCLSpectra.ListCount < 1 Then Exit Sub
If FormGETCMP.ListCLSpectra.ListIndex < 0 Then Exit Sub
specnum% = FormGETCMP.ListCLSpectra.ItemData(FormGETCMP.ListCLSpectra.ListIndex)

Call StandardDisplayCLSpectrum(GetCmpTmpSample(1).number%, specnum%, GetCmpTmpSample())
If ierror Then Exit Sub
End If

Exit Sub

' Errors
GetCmpDisplaySpectrumError:
MsgBox Error$, vbOKOnly + vbCritical, "GetCmpDisplaySpectrum"
ierror = True
Exit Sub

End Sub

Sub GetCmpLoadSpectra(mode As Integer)
' Load all EDS or CL spectrum into the list boxes for this standard

ierror = False
On Error GoTo GetCmpLoadSpectraError

' Load all EDS spectra
If mode% = 0 Then
Call StandardLoadSpectra(mode%, GetCmpTmpSample(1).number%, FormGETCMP.ListEDSSpectra)
If ierror Then Exit Sub
If FormGETCMP.ListEDSSpectra.ListCount > 0 Then FormGETCMP.ListEDSSpectra.ListIndex = FormGETCMP.ListEDSSpectra.ListCount - 1
End If

' Load all CL spectrum
If mode% = 1 Then
Call StandardLoadSpectra(mode%, GetCmpTmpSample(1).number%, FormGETCMP.ListCLSpectra)
If ierror Then Exit Sub
If FormGETCMP.ListCLSpectra.ListCount > 0 Then FormGETCMP.ListCLSpectra.ListIndex = FormGETCMP.ListCLSpectra.ListCount - 1
End If

Exit Sub

' Errors
GetCmpLoadSpectraError:
MsgBox Error$, vbOKOnly + vbCritical, "GetCmpLoadSpectra"
ierror = True
Exit Sub

End Sub

