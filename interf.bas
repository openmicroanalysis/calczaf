Attribute VB_Name = "CodeINTERF"
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit

Dim InterfTmpSample(1 To 1) As TypeSample

Sub InterfLoadWeight()
' Get user specified composition for interference calculation

ierror = False
On Error GoTo InterfLoadWeightError

Dim i As Integer
Dim tmsg As String, astring As String

' Create default string
astring$ = MatchGetWeightString(InterfTmpSample())
If ierror Then Exit Sub

' Load WEIGHT form and get unknown weight percents
FormWEIGHT.TextWeightPercentString.Text = astring$
FormWEIGHT.Show vbModal
If icancel Then Exit Sub

' Return modified sample
Call FormulaReturnSample(InterfTmpSample())

' Get sample string to display
tmsg$ = TypeWeight$(Int(2), InterfTmpSample())
FormINTERF.TextComposition.Text = tmsg$

' Load sample elements in list
FormINTERF.ComboElement.Clear
If InterfTmpSample(1).LastChan% > 0 Then
For i% = 1 To InterfTmpSample(1).LastChan%
FormINTERF.ComboElement.AddItem InterfTmpSample(1).Elsyms$(i%)
Next i%
FormINTERF.ComboElement.ListIndex = 0
End If

Exit Sub

' Errors
InterfLoadWeightError:
MsgBox Error$, vbOKOnly + vbCritical, "InterfLoadWeight"
ierror = True
Exit Sub

End Sub

Sub InterfLoadXrayDatabase()
' Loads an element xray range

ierror = False
On Error GoTo InterfLoadXrayDatabaseError

Dim i As Integer
Dim klm As Single, keV As Single, lam As Single
Dim xstart As Single, xstop As Single
Dim syme As String, symx As String
Dim rangefraction As Single

' Check range fraction
If Val(FormINTERF.TextRangeFraction.Text) < 0.01 Or Val(FormINTERF.TextRangeFraction.Text) > 0.5 Then
msg$ = "Xray database search range fraction is out of range. Must be between 0.01 and 0.50."
MsgBox msg$, vbOKOnly + vbExclamation, "InterfLoadXrayDatabase"
ierror = True
Exit Sub
Else
rangefraction! = Val(FormINTERF.TextRangeFraction.Text)
End If

' Save default
DefaultRangeFraction! = rangefraction!

' Load element and xray symbols
i% = FormINTERF.ComboElement.ListIndex + 1
If i% < 1 Or i% > InterfTmpSample(1).LastChan% Then Exit Sub
syme$ = InterfTmpSample(1).Elsyms$(i%)
symx$ = InterfTmpSample(1).Xrsyms$(i%)

' Get angstroms for this x-ray
Call XrayGetKevLambda(syme$, symx$, keV!, lam!)
If ierror Then Exit Sub

' Calculate x-ray range to load from XRAY.MDB
xstart! = lam! - lam! * rangefraction!
xstop! = lam! + lam! * rangefraction!

' Load form
keV! = DefaultKiloVolts!
klm! = DefaultMinimumKLMDisplay!
Call XrayLoad(Int(2), Int(0), klm!, keV!, xstart!, xstop!)
If ierror Then Exit Sub

FormXRAY.Show vbModal

Exit Sub

' Errors
InterfLoadXrayDatabaseError:
MsgBox Error$, vbOKOnly + vbCritical, "InterfLoadXrayDatabase"
ierror = True
Exit Sub

End Sub

Sub InterfSave()
' Save FormINTERF options and call for interference calculations

ierror = False
On Error GoTo InterfSaveError

Dim ip As Integer, i As Integer
Dim sym As String
Dim mode As Integer, chan As Integer
Dim lifwidth As Single, overlap As Single, discrimination As Single
Dim x2d As Single, k As Single, pos As Single, angs As Single
Dim elm As String, ray As String
Dim tmsg As String, astring As String
Dim rangefraction As Single

' Check range fraction
If Val(FormINTERF.TextRangeFraction.Text) < 0.01 Or Val(FormINTERF.TextRangeFraction.Text) > 0.5 Then
msg$ = "Xray database search range fraction is out of range. Must be between 0.01 and 0.50."
MsgBox msg$, vbOKOnly + vbExclamation, "InterfSave"
ierror = True
Exit Sub
Else
rangefraction! = Val(FormINTERF.TextRangeFraction.Text)
End If

' Get user specified options
If Val(FormINTERF.TextLiFPeakWidth.Text) < 0.01 Or Val(FormINTERF.TextLiFPeakWidth.Text) > 0.5 Then
msg$ = "LiF peak width is out of range. Must be between 0.01 and 0.50."
MsgBox msg$, vbOKOnly + vbExclamation, "InterfSave"
ierror = True
Exit Sub
Else
lifwidth! = Val(FormINTERF.TextLiFPeakWidth.Text)
End If

If Val(FormINTERF.TextMinimumOverlap.Text) < 0.01 Or Val(FormINTERF.TextMinimumOverlap.Text) > 50# Then
msg$ = "Minimum overlap is out of range. Must be between 0.01 and 50."
MsgBox msg$, vbOKOnly + vbExclamation, "InterfSave"
ierror = True
Exit Sub
Else
overlap! = Val(FormINTERF.TextMinimumOverlap.Text)
End If

If Val(FormINTERF.TextPHADiscrimination.Text) < 1# Or Val(FormINTERF.TextPHADiscrimination.Text) > 100# Then
msg$ = "PHA discrimination is out of range. Must be between 1 and 100."
MsgBox msg$, vbOKOnly + vbExclamation, "InterfSave"
ierror = True
Exit Sub
Else
discrimination! = Val(FormINTERF.TextPHADiscrimination.Text)
End If

' Get interfered peak
For i% = 0 To 2
If FormINTERF.OptionInterferencePeak(i%).Value Then mode% = i%
Next i%

' Get element
If FormINTERF.OptionAll.Value Then
chan% = 0

Else
sym$ = FormINTERF.ComboElement.Text
ip% = IPOS1(InterfTmpSample(1).LastElm%, sym$, InterfTmpSample(1).Elsyms$())
chan% = ip%

sym$ = FormINTERF.ComboXray.Text
InterfTmpSample(1).Xrsyms$(chan%) = sym$
End If

' Save defaults
DefaultRangeFraction! = rangefraction!
DefaultLIFPeakWidth! = lifwidth!
DefaultMinimumOverlap! = overlap!
DefaultPHADiscrimination! = discrimination!

' Re-load hydrogen and helium dummy x-ray lines (for subsequent times)
For i% = 1 To InterfTmpSample(1).LastChan%
If InterfTmpSample(1).AtomicNums%(i%) = 1 Or InterfTmpSample(1).AtomicNums%(i%) = 2 Then
InterfTmpSample(1).Xrsyms$(i%) = Xraylo$(1)
End If
Next i%

' Load element data
Call ElementGetData(InterfTmpSample())
If ierror Then Exit Sub

' Calculate off-peak positions for this sample
For i% = 1 To InterfTmpSample(1).LastElm%
If InterfTmpSample(1).Xrsyms$(i%) <> vbNullString Then
InterfTmpSample(1).MotorNumbers%(i%) = 1    ' assume spectrometer 1 for all elements

' Get the 2d and k of this crystal for position calculation below
Call MiscGetCrystalParameters(InterfTmpSample(1).CrystalNames$(i%), x2d!, k!, elm$, ray$)
If ierror Then Exit Sub

' Get on peak in angstroms
angs! = ANGEV! / InterfTmpSample(1).LineEnergy!(i%)
pos! = XrayCalculatePositions(Int(0), InterfTmpSample(1).MotorNumbers%(i%), Int(1), x2d!, k!, angs!)
InterfTmpSample(1).OnPeaks!(i%) = pos!

' Get hi off peak
pos! = XrayCalculatePositions(Int(1), InterfTmpSample(1).MotorNumbers%(i%), Int(1), x2d!, k!, InterfTmpSample(1).OnPeaks!(i%))
InterfTmpSample(1).HiPeaks!(i%) = pos!

' Get lo off peak
pos! = XrayCalculatePositions(Int(2), InterfTmpSample(1).MotorNumbers%(i%), Int(1), x2d!, k!, InterfTmpSample(1).OnPeaks!(i%))
InterfTmpSample(1).LoPeaks!(i%) = pos!

' Save crystal 2d
InterfTmpSample(1).Crystal2ds!(i%) = x2d!
End If
Next i%

' Get interferences for this sample
astring$ = Interf2Calculate(Int(1), mode%, chan%, rangefraction!, lifwidth!, overlap!, discrimination!, InterfTmpSample())
If ierror Then Exit Sub

' Write interferences to Log Window
tmsg$ = vbCrLf
If mode% = 0 Then tmsg$ = tmsg$ & "On Peak Interferences for : " & TypeLoadString$(InterfTmpSample())
If mode% = 1 Then tmsg$ = tmsg$ & "Hi Peak Interferences for : " & TypeLoadString$(InterfTmpSample())
If mode% = 2 Then tmsg$ = tmsg$ & "Lo Peak Interferences for : " & TypeLoadString$(InterfTmpSample())
tmsg$ = tmsg$ & vbCrLf

' Write to log window
Call IOWriteLog(tmsg$ & astring$)

Exit Sub

' Errors
InterfSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "InterfSave"
ierror = True
Exit Sub

End Sub

Sub InterfLoad()
' Load FormINTERF for nominal interference calculations

ierror = False
On Error GoTo InterfLoadError

Dim tmsg As String
Dim number As Integer, i As Integer

' Get standard from listbox
If FormMAIN.ListAvailableStandards.ListIndex < 0 Then Exit Sub
If FormMAIN.ListAvailableStandards.ListCount < 1 Then GoTo InterfLoadNoStandards
number% = FormMAIN.ListAvailableStandards.ItemData(FormMAIN.ListAvailableStandards.ListIndex)

' Get standard from database
Call StandardGetMDBStandard(number%, InterfTmpSample())
If ierror Then Exit Sub

' Convert standard to string
tmsg$ = TypeWeight(Int(2), InterfTmpSample())
FormINTERF.TextComposition.Text = tmsg$

' Default is calculate interferences for all elements in sample
FormINTERF.OptionAll.Value = True

' Default is On peak interferences
FormINTERF.OptionInterferencePeak(0).Value = True

' Load default values
FormINTERF.TextRangeFraction.Text = Str$(DefaultRangeFraction!)
FormINTERF.TextLiFPeakWidth.Text = Str$(DefaultLIFPeakWidth!)
FormINTERF.TextMinimumOverlap.Text = Str$(DefaultMinimumOverlap!)
FormINTERF.TextPHADiscrimination.Text = Str$(DefaultPHADiscrimination!)

' Load sample elements in list
FormINTERF.ComboElement.Clear
If InterfTmpSample(1).LastChan% > 0 Then
For i% = 1 To InterfTmpSample(1).LastChan%
FormINTERF.ComboElement.AddItem InterfTmpSample(1).Elsyms$(i%)
Next i%
End If

' Load sample elements in list
FormINTERF.ComboXray.Clear
For i% = 1 To MAXRAY% - 1
FormINTERF.ComboXray.AddItem Xraylo$(i%)
Next i%

' Load first element
FormINTERF.ComboElement.ListIndex = 0
Exit Sub

' Errors
InterfLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "InterfLoad"
ierror = True
Exit Sub

InterfLoadNoStandards:
msg$ = "No standards entered in standard database yet"
MsgBox msg$, vbOKOnly + vbExclamation, "InterfLoad"
ierror = True
Exit Sub

End Sub

Sub InterfUpdateElement()
' Update the x-ray line based on the element

ierror = False
On Error GoTo InterfUpdateElementError

Dim ip As Integer

ip% = IPOS1%(MAXELM%, FormINTERF.ComboElement.Text, Symlo$())
If ip% > 0 Then
FormINTERF.ComboXray.Text = Deflin$(ip%)
End If

Exit Sub

' Errors
InterfUpdateElementError:
MsgBox Error$, vbOKOnly + vbCritical, "InterfUpdateElement"
ierror = True
Exit Sub

End Sub
