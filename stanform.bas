Attribute VB_Name = "CodeSTANFORM"
' (c) Copyright 1995-2026 by John J. Donovan
Option Explicit

' Used by Stanform and GetCmp routines
Global OxPercents(1 To MAXCHAN%) As Single
Global AtPercents(1 To MAXCHAN%) As Single

Dim StanFormOldSample(1 To 1) As TypeSample
Dim StanFormTmpSample(1 To 1) As TypeSample

Dim StanFormAnalysis As TypeAnalysis

Sub StanFormCalculate(stdnum As Integer, mode As Integer)
' Routine to recalculate the new standard values and update the form
' mode = 0 normal output
' mode = 1 supress output

ierror = False
On Error GoTo StanFormCalculateError

Dim i As Integer, stdrow As Integer
Dim bfilename As String
Dim maszbar As Single, zedzbar As Single

' Get standard from database
Call StandardGetMDBStandard(stdnum%, StanFormTmpSample())
If ierror Then Exit Sub

' Check the element xrays for hydrogen and helium
Call ElementCheckXray(Int(0), StanFormTmpSample())
If ierror Then Exit Sub

' Sort for hydrogen and helium
Call GetElmSaveSampleOnly(Int(0), StanFormTmpSample(), Int(0), Int(0))
If ierror Then Exit Sub

' Set StanformOldSample equal to Tmp sample so k factors and ZAF corrections get loaded in ZAFStd
StanFormOldSample(1) = StanFormTmpSample(1)

' Initialize arrays
Call InitStandards(StanFormAnalysis)
If ierror Then Exit Sub

' Set sample standard assignments so that ZAF and K-factors get loaded
For i% = 1 To StanFormTmpSample(1).LastChan%
StanFormOldSample(1).StdAssigns%(i%) = stdnum%
Next i%

' Run the calculations on the standard
stdrow% = MAXSTD%
'Call ZAFStd(stdrow%, StanFormAnalysis, StanFormOldSample(), StanFormTmpSample())
Call ZAFStd2(stdrow%, StanFormAnalysis, StanFormOldSample(), StanFormTmpSample())
If ierror Then Exit Sub

If CorrectionFlag% > 0 And CorrectionFlag% < 4 Then
Call AFactorStd(stdrow%, StanFormAnalysis, StanFormOldSample(), StanFormTmpSample())
If ierror Then Exit Sub
End If

' Get oxygen channel
Call ZAFGetOxygenChannel(StanFormTmpSample())
If ierror Then Exit Sub

' Calculate oxides and atomic percents
Call StanFormCalculateOxideAtomic(StanFormTmpSample())
If ierror Then Exit Sub

' Calculate formula (new code 06/15/2017)
If StanFormTmpSample(1).FormulaElementFlag Then
Call ConvertWeightToFormula(StanFormAnalysis, StanFormTmpSample())
If ierror Then Exit Sub
End If

' Subtract calculated from total oxygen
StanFormOldSample(1) = StanFormTmpSample(1)
Call StanFormCalculateExcessOxygen(StanFormAnalysis, StanFormOldSample(), StanFormTmpSample())
If ierror Then Exit Sub

' If supressing output then exit
If mode% > 0 Then Exit Sub

' Update the Standard sample description text box
msg$ = StandardLoadDescription2(StanFormOldSample())
If ierror Then Exit Sub
FormMAIN.LabelStandard.Caption = msg$

' Type out the standard composition to the log window
Call StanformTypeStandard(stdrow%, StanFormAnalysis, StanFormOldSample())
If ierror Then Exit Sub

' Calculate alternative zbars
If CalculateAlternativeZbarsFlag Then
Screen.MousePointer = vbHourglass
Call StanFormCalculateZbars(maszbar!, zedzbar!, StanFormOldSample())
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End If

' Calculate continuum absorption
If CalculateContinuumAbsorptionFlag Then
Screen.MousePointer = vbHourglass
Call StanFormCalculateContinuum(StanFormOldSample())
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End If

' Calculate electron and x-ray ranges
If CalculateElectronandXrayRangesFlag Then
Screen.MousePointer = vbHourglass
Call ZAFCalculateRange(StanFormAnalysis, StanFormOldSample())
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End If

' Send composition to MQ input file
If FormMQOPTIONS.Visible Then
bfilename$ = UserDataDirectory$ & "\MQData\" & Trim$(Format$(StanFormOldSample(1).kilovolts!)) & "-" & "STANDARD.BAT"
Call MqOptionsSendCompToFile(maszbar!, zedzbar!, bfilename$, StanFormOldSample())
If ierror Then Exit Sub
End If

' Amphibole and biotite calculations
If DisplayAmphiboleCalculationFlag Then
If Not StanFormOldSample(1).DisplayAsOxideFlag Then GoTo StanFormCalculateNotOxide
Call ConvertMinerals2(Int(6), OxPercents!(), StanFormOldSample())
If ierror Then Exit Sub
End If

If DisplayBiotiteCalculationFlag Then
If Not StanFormOldSample(1).DisplayAsOxideFlag Then GoTo StanFormCalculateNotOxide
Call ConvertMinerals2(Int(7), OxPercents!(), StanFormOldSample())
If ierror Then Exit Sub
End If

Exit Sub

' Errors
StanFormCalculateError:
MsgBox Error$, vbOKOnly + vbCritical, "StanFormCalculate"
ierror = True
Exit Sub

StanFormCalculateNotOxide:
msg$ = "Must have Display As Oxide flag selected for mineral formula calculations. See Standard | Modify menu."
MsgBox msg$, vbOKOnly + vbExclamation, "StanFormCalculate"
ierror = True
Exit Sub

End Sub

Sub StanFormDeleteSingleStandard()
' Delete the currently selected standard

ierror = False
On Error GoTo StanFormDeleteSingleStandardError

Dim stdnum As Integer, numberselected As Integer
Dim i As Integer, response As Integer

' Get standard from listbox
If FormMAIN.ListAvailableStandards.ListIndex < 0 Then Exit Sub
If FormMAIN.ListAvailableStandards.ListCount < 1 Then GoTo StanFormDeleteSingleStandardNoStandards

' Check if more than one standard is selected
numberselected = 0
For i% = 0 To FormMAIN.ListAvailableStandards.ListCount - 1
If FormMAIN.ListAvailableStandards.Selected(i%) Then numberselected% = numberselected% + 1
Next i%
If numberselected% > 1 Then GoTo StanFormDeleteSingleStandardTooMany

' Warn user about deleting standards
If UCase$(Dir$(StandardDataFile$)) = "STANDARD.MDB" Then
msg$ = "Warning- The user should be aware that deleting a standard composition from the "
msg$ = msg$ & "default Standard Database that is also referenced by a Probe "
msg$ = msg$ & "Database file, will result in the Probe Database file being unusable. Be "
msg$ = msg$ & "sure that the standard to be deleted is not used by any Probe Database files."
response% = MsgBox(msg$, vbOKCancel + vbInformation + vbDefaultButton2, "StanFormDeleteSingleStandards")
If response% = vbCancel Then
ierror = True
Exit Sub
End If
End If

' Get number of standard to delete
stdnum% = FormMAIN.ListAvailableStandards.ItemData(FormMAIN.ListAvailableStandards.ListIndex)

' Confirm with user
msg$ = "Are you sure that you want to delete standard :"
msg$ = msg$ & vbCrLf & StandardGetString2$(stdnum%)
msg$ = msg$ & vbCrLf & " from database :"
msg$ = msg$ & vbCrLf & StandardDataFile$ & "?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton2, "StanFormDeleteSingleStandard")
If response% = vbNo Then Exit Sub

' Delete the selected standard
Screen.MousePointer = vbHourglass
Call StandardDeleteRecord(stdnum%)
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

' Update the screen
Call StanFormClear

' Reopen the database and re-load available standards
Call StandardGetMDBIndex
If ierror Then Exit Sub

' Update the screen
Call StandardLoadList(FormMAIN.ListAvailableStandards)
If ierror Then Exit Sub

' Update form
Call StanFormUpdate
If ierror Then Exit Sub

Exit Sub

' Errors
StanFormDeleteSingleStandardError:
MsgBox Error$, vbOKOnly + vbCritical, "StanFormDeleteSingleStandard"
ierror = True
Exit Sub

StanFormDeleteSingleStandardNoStandards:
msg$ = "No standards entered in standard database yet"
MsgBox msg$, vbOKOnly + vbExclamation, "StanFormDeleteSingleStandard"
ierror = True
Exit Sub

StanFormDeleteSingleStandardTooMany:
msg$ = "More than one standard is selected in the standard list. Select a single standard and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "StanFormDeleteSingleStandard"
ierror = True
Exit Sub

End Sub

Sub StanFormListNames(mode As Integer, method As Integer)
' List standard names only
' mode = 0 all standards
' mode = 1 elemental standards
' mode = 2 oxide standards
' method = 0 no additional output
' method = 1 output average Z

ierror = False
On Error GoTo StanFormListNamesError

Dim i As Integer, stdnum As Integer
Dim astring As String

icancelauto = False

If method% = 0 Then
If mode% = 0 Then Call IOWriteLog(vbCrLf & "All Standards:")
If mode% = 1 Then Call IOWriteLog(vbCrLf & "Elemental Standards:")
If mode% = 2 Then Call IOWriteLog(vbCrLf & "Oxide Standards:")
Else
If method% = 1 Then
If UseZFractionZbarCalculationsFlag Then
If mode% = 0 Then Call IOWriteLog(vbCrLf & "All Standards (Yukawa Z fraction Zbars):")
If mode% = 1 Then Call IOWriteLog(vbCrLf & "Elemental Standards (Yukawa Z fraction Zbars):")
If mode% = 2 Then Call IOWriteLog(vbCrLf & "Oxide Standards (Yukawa Z fraction Zbars):")
Else
If mode% = 0 Then Call IOWriteLog(vbCrLf & "All Standards (Mass fraction Zbars):")
If mode% = 1 Then Call IOWriteLog(vbCrLf & "Elemental Standards (Mass fraction Zbars):")
If mode% = 2 Then Call IOWriteLog(vbCrLf & "Oxide Standards (Mass fraction Zbars):")
End If
End If
End If

' Write to Log Window
For i% = 1 To FormMAIN.ListAvailableStandards.ListCount
msg$ = vbNullString

' Check for method
If method% = 1 Then
stdnum% = FormMAIN.ListAvailableStandards.ItemData(i% - 1)
Call StanFormCalculate(stdnum%, Int(1))     ' suppress output
If ierror Then Exit Sub

' Calculated values are in StanFormOldSample() and StanFormAnalysis
astring$ = vbTab & "Zbar = " & vbTab & Format$(StanFormAnalysis.zbar!)
End If

' All
If mode% = 0 Then
msg$ = FormMAIN.ListAvailableStandards.List(i% - 1)
If method% = 1 Then msg$ = msg$ & astring$
Call IOWriteLog(msg$)
End If

' Elemental
If mode% = 1 Then
If Not StandardIsStandardOxide(FormMAIN.ListAvailableStandards.ItemData(i% - 1)) Then
msg$ = FormMAIN.ListAvailableStandards.List(i% - 1)
If method% = 1 Then msg$ = msg$ & astring$
Call IOWriteLog(msg$)
End If
End If

' Oxide
If mode% = 2 Then
If StandardIsStandardOxide(FormMAIN.ListAvailableStandards.ItemData(i% - 1)) Then
msg$ = FormMAIN.ListAvailableStandards.List(i% - 1)
If method% = 1 Then msg$ = msg$ & astring$
Call IOWriteLog(msg$)
End If
End If

DoEvents
If icancelauto Then
ierror = True
Exit Sub
End If
Next i%

Exit Sub

' Errors
StanFormListNamesError:
MsgBox Error$, vbOKOnly + vbCritical, "StanFormListNames"
ierror = True
Exit Sub

End Sub

Sub StanFormListStandards(mode As Integer)
' mode = 1 list selected standards
' mode = 2 list all standards
' mode = 3 list elemental standards
' mode = 4 list oxide standards

ierror = False
On Error GoTo StanFormListStandardsError

Dim i As Integer
Dim stdnum As Integer, numselstds As Integer, n As Integer

icancelauto = False

' Get standards from listbox
If FormMAIN.ListAvailableStandards.ListIndex < 0 Then Exit Sub

If mode% = 1 Then Call IOWriteLog(vbCrLf & "Selected Standards:")
If mode% = 2 Then Call IOWriteLog(vbCrLf & "All Standards:")
If mode% = 3 Then Call IOWriteLog(vbCrLf & "Elemental Standards:")
If mode% = 4 Then Call IOWriteLog(vbCrLf & "Oxide Standards:")

' If doing selected standards, calculate number of selected standards
If mode% = 1 Then
numselstds% = 0
For i% = 0 To FormMAIN.ListAvailableStandards.ListCount - 1
If FormMAIN.ListAvailableStandards.Selected(i%) Then numselstds% = numselstds% + 1
Next i%
End If

' If doing elemental or oxide standards calculate number
If mode% = 3 Then
numselstds% = StandardGetNumberofStandards(Int(1))
If ierror Then Exit Sub
End If
If mode% = 4 Then
numselstds% = StandardGetNumberofStandards(Int(2))
If ierror Then Exit Sub
End If

' Loop on all standards in list
n% = 0
For i% = 0 To FormMAIN.ListAvailableStandards.ListCount - 1
stdnum% = FormMAIN.ListAvailableStandards.ItemData(i%)

' Recalculate selected
If mode% = 1 And FormMAIN.ListAvailableStandards.Selected(i%) Then
n% = n% + 1

Call IOStatusAuto("Calculating Selected Standard " & Str$(stdnum%) & " (" & Str$(n%) & " of " & Str$(numselstds%) & ")")
DoEvents
If icancelauto Then
ierror = True
Exit Sub
End If

Call StanFormCalculate(stdnum%, Int(0))
If ierror Then Exit Sub
End If

' Recalculate all
If mode% = 2 Then
Call IOStatusAuto("Calculating Standard " & Str$(stdnum%) & " (" & Str$(i% + 1) & " of " & Str$(NumberOfAvailableStandards%) & ")")
DoEvents
If icancelauto Then
ierror = True
Exit Sub
End If

Call StanFormCalculate(stdnum%, Int(0))
If ierror Then Exit Sub
End If

' Recalculate elemental
If mode% = 3 Then
If Not StandardIsStandardOxide(stdnum%) Then
If ierror Then Exit Sub
n% = n% + 1
Call IOStatusAuto("Calculating Elemental Standard " & Str$(stdnum%) & " (" & Str$(n%) & " of " & Str$(numselstds%) & ")")
DoEvents
If icancelauto Then
ierror = True
Exit Sub
End If

Call StanFormCalculate(stdnum%, Int(0))
If ierror Then Exit Sub
End If
End If

' Recalculate elemental
If mode% = 4 Then
If StandardIsStandardOxide(stdnum%) Then
If ierror Then Exit Sub
n% = n% + 1
Call IOStatusAuto("Calculating Oxide Standard " & Str$(stdnum%) & " (" & Str$(n%) & " of " & Str$(numselstds%) & ")")
DoEvents
If icancelauto Then
ierror = True
Exit Sub
End If

Call StanFormCalculate(stdnum%, Int(0))
If ierror Then Exit Sub
End If
End If

DoEvents
Next i%

Exit Sub

' Errors
StanFormListStandardsError:
MsgBox Error$, vbOKOnly + vbCritical, "StanFormListStandards"
ierror = True
Exit Sub

End Sub

Sub StanFormModify()
' Get a modified standard element list and composition from user

ierror = False
On Error GoTo StanFormModifyError

Dim samplenumber As Integer

' Get standard from listbox
If FormMAIN.ListAvailableStandards.ListIndex < 0 Then Exit Sub
If FormMAIN.ListAvailableStandards.ListCount < 1 Then GoTo StanFormModifyNoStandards
samplenumber% = FormMAIN.ListAvailableStandards.ItemData(FormMAIN.ListAvailableStandards.ListIndex)

' Init the sample (for EDS and CL arrays)
Call InitSample(StanFormTmpSample())
If ierror Then Exit Sub

' Get standard from database
Call StandardGetMDBStandard(samplenumber%, StanFormTmpSample())
If ierror Then Exit Sub

' Get modified element list and/or composition
GetCmpFlag% = 2
Call GetCmpLoad(StanFormTmpSample())
If ierror Then Exit Sub

' Standard Database updated from OK button in FormGETCMP (GetCmpSaveAll)
FormGETCMP.Show vbModal

' Get modified sample
Call GetCmpReturn(StanFormTmpSample())
If ierror Then Exit Sub

' Recalculate k-factors
If samplenumber% > 0 Then
Call StanFormCalculate(samplenumber%, Int(0))
If ierror Then Exit Sub
End If

' Reopen the database and re-load available standards
Call StandardGetMDBIndex
If ierror Then Exit Sub

' Update the screen
Call StandardLoadList(FormMAIN.ListAvailableStandards)
If ierror Then Exit Sub

' Re-select current standard
Call StandardSelectList(samplenumber%, FormMAIN.ListAvailableStandards)
If ierror Then Exit Sub

Exit Sub

' Errors
StanFormModifyError:
MsgBox Error$, vbOKOnly + vbCritical, "StanFormModify"
ierror = True
Exit Sub

StanFormModifyNoStandards:
msg$ = "No standards entered in standard database yet"
MsgBox msg$, vbOKOnly + vbExclamation, "StanFormModify"
ierror = True
Exit Sub

End Sub

Sub StanFormNew()
' Get a new standard element list and composition from user

ierror = False
On Error GoTo StanFormNewError

Dim samplenumber As Integer

' Check for too many standards
If NumberOfAvailableStandards% >= MAXINDEX% Then GoTo StanFormNewTooMany

' Initialize the Tmp sample arrays
Call InitSample(StanFormTmpSample())
If ierror Then Exit Sub

' Load default name and description
StanFormTmpSample(1).Name$ = "Standard Name"
StanFormTmpSample(1).Description$ = "Standard Description"

GetCmpFlag% = 1
Call GetCmpLoad(StanFormTmpSample())
If ierror Then Exit Sub

' Standard Database updated from OK button in FormGETCMP
FormGETCMP.Show vbModal
If ierror Then Exit Sub

' Get changes from FormGETCMP
Call GetCmpReturn(StanFormTmpSample())
If ierror Then Exit Sub

' Re-calculate the new standard
samplenumber% = StanFormTmpSample(1).number%
If samplenumber% > 0 Then
Call StanFormCalculate(samplenumber%, Int(0))
If ierror Then Exit Sub
End If

' Reopen the database and re-load available standards
Call StandardGetMDBIndex
If ierror Then Exit Sub

' Update the screen
Call StandardLoadList(FormMAIN.ListAvailableStandards)
If ierror Then Exit Sub

' Re-select current standard
Call StandardSelectList(samplenumber%, FormMAIN.ListAvailableStandards)
If ierror Then Exit Sub

Exit Sub

' Errors
StanFormNewError:
MsgBox Error$, vbOKOnly + vbCritical, "StanFormNew"
ierror = True
Exit Sub

StanFormNewTooMany:
msg$ = "There are too many standards already in the standard database"
MsgBox msg$, vbOKOnly + vbExclamation, "StanFormNew"
ierror = True
Exit Sub

End Sub

Sub StanFormSaveAsFile(tForm As Form)
' Saves the file to another name

ierror = False
On Error GoTo StanFormSaveAsFileError

Dim Filename As String, tempname As String

' Copy standard file to temp file name
tempname$ = ApplicationCommonAppData$ & "temp.mdb"
FileCopy StandardDataFile$, tempname$

' Get new file
Call IOGetMDBFileName(Int(3), Filename$, tForm)
If ierror Then Exit Sub

' Copy standard file to new file name
FileCopy tempname$, Filename$

' Save the new standard database file name
StandardDataFile$ = Filename$

' No errors, so delete temp file
Kill tempname$

' Load database file info
Call FileInfoLoad(Int(1), StandardDataFile$)
If ierror Then Exit Sub

' Get changed file information, (do not exit on error)
FormFILEINFO.Show vbModal
'If ierror Then Exit Sub

Call StanFormUpdate
If ierror Then Exit Sub

Exit Sub

StanFormSaveAsFileError:
MsgBox Error$, vbOKOnly + vbCritical, "StanFormSaveAsFile"
ierror = True
Exit Sub

End Sub

Sub StanformTypeStandard(stdrow As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' Type the standard composition and calculations for a standard
' "stdrow" is the position in the standard list

ierror = False
On Error GoTo StanformTypeStandardError

Dim n As Integer, i As Integer
Dim ii As Integer, jj As Integer
Dim temp As Single, TotalCations As Single

' Load the Total, Zbar, etc. text fields
Call ZAFCalZbarLoadText(FormMAIN, analysis)
If ierror Then Exit Sub

' Type the sample name and description (skip a space)
msg$ = TypeLoadString(sample())
Call IOWriteLogRichText(vbCrLf & msg$, vbNullString, Int(LogWindowFontSize% + 3), vbBlue, Int(FONT_ITALIC% Or FONT_UNDERLINE%), Int(0))
If ierror Then Exit Sub
msg$ = StandardLoadDescription(sample())
Call IOWriteLog(msg$)

' Note all standards are elemental composition, but can be displayed as both
If sample(1).DisplayAsOxideFlag Then
msg$ = "Oxide and Elemental Composition"
Call IOWriteLog(msg$)
Else
msg$ = "Elemental Composition"
Call IOWriteLog(msg$)
End If

' Write the Zbar, etc to Log window
Call TypeZbar(Int(2), analysis)
If ierror Then Exit Sub

' Type out analyzed data for the sample
n = 0
Do Until False
n% = n% + 1
Call TypeGetRange(Int(2), n%, ii%, jj%, sample())
If ierror Then Exit Sub
If ii% > sample(1).LastChan% Then Exit Do

' Type out standard composition to Log window
msg$ = vbCrLf & "ELEM: "
For i% = ii% To jj%
If sample(1).DisplayAsOxideFlag Then
msg$ = msg$ & Format$(sample(1).Oxsyup$(i%), a80$)
Else
msg$ = msg$ & Format$(sample(1).Elsyup$(i%), a80$)
End If
Next i%
Call IOWriteLog(msg$)

' Type out the default x-ray line used in the calculations
msg$ = "XRAY: "
For i% = ii% To jj%
msg$ = msg$ & Format$(sample(1).Xrsyms$(i%) & " ", a80$)
Next i%
Call IOWriteLog(msg$)

' Type out elemental and oxide weight percents
If sample(1).DisplayAsOxideFlag Then
msg$ = "OXWT: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(OxPercents!(i%), f83$), a80$)
Next i%
Call IOWriteLog(msg$)
End If

msg$ = "ELWT: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).ElmPercents!(i%), f83$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "ATWT: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(sample(1).AtomicWts!(i%), f83$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "KFAC: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(analysis.StdAssignsKfactors!(i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

If DisplayZAFCalculationFlag Then
msg$ = "MACS: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(analysis.StdMACs!(stdrow%, i%), e71$), a80$)
Next i%
Call IOWriteLog(msg$)
End If

msg$ = "ZCOR: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(analysis.StdZAFCors!(4, stdrow%, i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

' Display detailed ZAF calculations
If DisplayZAFCalculationFlag Then
msg$ = "ZABS: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(analysis.StdZAFCors!(1, stdrow%, i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)
msg$ = "ZFLU: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(analysis.StdZAFCors!(2, stdrow%, i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)
msg$ = "ZZED: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(analysis.StdZAFCors!(3, stdrow%, i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)
msg$ = "ZSTP: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(analysis.StdZAFCors!(5, stdrow%, i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)
msg$ = "ZBKS: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(analysis.StdZAFCors!(6, stdrow%, i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)

msg$ = "FCHI: "
For i% = ii% To jj%
If analysis.StdZAFCors!(8, stdrow%, i%) <> 0# Then
temp! = analysis.StdZAFCors!(8, stdrow%, i%)
msg$ = msg$ & Format$(Format$(temp!, f84$), a80$)
Else
msg$ = msg$ & Format$(Format$(1#, f84$), a80$)
End If
Next i%
Call IOWriteLog(msg$)

If CalculateContinuumAbsorptionFlag Then
msg$ = "ABSC: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(analysis.StdContinuumCorrections!(stdrow%, i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)
End If
End If

If CorrectionFlag% > 0 And CorrectionFlag% < 4 Then
msg$ = "BETA: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(analysis.StdBetas!(stdrow%, i%), f84$), a80$)
Next i%
Call IOWriteLog(msg$)
End If

msg$ = "AT% : "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(AtPercents!(i%), f83$), a80$)
Next i%
Call IOWriteLog(msg$)

' Calculate formula if oxide
If sample(1).DisplayAsOxideFlag% And sample(1).OxygenChannel% > 0 Then
If AtPercents!(sample(1).OxygenChannel%) <> 0# Then
msg$ = "24 O: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(24# * AtPercents!(i%) / AtPercents!(sample(1).OxygenChannel%), f83$), a80$)
Next i%
Call IOWriteLog(msg$)
End If
End If

' Display formula if specified (new code 06/15/2017)
If sample(1).FormulaElementFlag Then
msg$ = "FORM: "
For i% = ii% To jj%
msg$ = msg$ & Format$(Format$(analysis.Formulas!(i%), f83$), a80$)
Next i%
Call IOWriteLog(msg$)
End If

Loop

' Calculate sum of cations
If UseTotalCationsCalculationFlag Then
If sample(1).DisplayAsOxideFlag% And sample(1).OxygenChannel% > 0 Then
If AtPercents!(sample(1).OxygenChannel%) <> 0# Then
TotalCations! = 0#
For i% = 1 To sample(1).LastChan%
If sample(1).AtomicCharges!(i%) > 0# Then TotalCations! = TotalCations! + 24# * AtPercents!(i%) / AtPercents!(sample(1).OxygenChannel%)
Next i%
msg$ = Format$(Format$(TotalCations!, f83$), a80$)
Call IOWriteLog(vbCrLf & "Total Cations (Based on 24 O) = " & msg$)
End If
End If
End If

Exit Sub

' Errors
StanformTypeStandardError:
MsgBox Error$, vbOKOnly + vbCritical, "StanformTypeStandard"
ierror = True
Exit Sub

End Sub

Sub StanFormUpdate()
' Update FormMAIN for Standard

ierror = False
On Error GoTo StanFormUpdateError

' Enable or disable the File menu items
If StandardDataFile$ = vbNullString Then
FormMAIN.menuFileNew.Enabled = True
FormMAIN.menuFileOpen.Enabled = True
FormMAIN.menuFileSaveAs.Enabled = False
FormMAIN.menuFileClose.Enabled = False
FormMAIN.menuFileImport.Enabled = False
FormMAIN.menuFileExport.Enabled = False
FormMAIN.menuFileImportSingleRowFormat.Enabled = False
FormMAIN.menuFileExportSingleRowFormat.Enabled = False
FormMAIN.menuFileImportStandardsFromCamecaPeakSight.Enabled = False
'FormMAIN.menuFileImportStandardsFromJEOLTextFile.Enabled = False
FormMAIN.menuFileImportStandardsFromJEOL8x30TextFile.Enabled = False

FormMAIN.menuFileInputCalcZAFStandardFormatKratios.Enabled = False
FormMAIN.menuFileAMCSD.Enabled = True                    ' always enabled
FormMAIN.menuFileFileInformation.Enabled = False

Else
FormMAIN.menuFileNew.Enabled = False
FormMAIN.menuFileOpen.Enabled = False
FormMAIN.menuFileSaveAs.Enabled = True
FormMAIN.menuFileClose.Enabled = True
FormMAIN.menuFileImport.Enabled = True
FormMAIN.menuFileExport.Enabled = True
FormMAIN.menuFileImportSingleRowFormat.Enabled = True
FormMAIN.menuFileExportSingleRowFormat.Enabled = True
FormMAIN.menuFileImportStandardsFromCamecaPeakSight.Enabled = True
'FormMAIN.menuFileImportStandardsFromJEOLTextFile.Enabled = True
FormMAIN.menuFileImportStandardsFromJEOL8x30TextFile.Enabled = True

'FormMAIN.menuFileInputCalcZAFStandardFormatKratios.Enabled = True
FormMAIN.menuFileAMCSD.Enabled = True
FormMAIN.menuFileFileInformation.Enabled = True
End If

' Enable/Disable other menu items
If StandardDataFile$ = vbNullString Then
FormMAIN.menuStandardNew.Enabled = False
FormMAIN.menuStandardModify.Enabled = False
FormMAIN.menuStandardDuplicate.Enabled = False
FormMAIN.menuStandardDelete.Enabled = False
FormMAIN.menuStandardDeleteSelected.Enabled = False
FormMAIN.menuStandardListStandardNames.Enabled = False
FormMAIN.menuStandardListSelectedStandards.Enabled = False
FormMAIN.menuStandardListAllStandards.Enabled = False
FormMAIN.menuStandardListStandardNamesZbar.Enabled = False
FormMAIN.menuStandardListStandardNamesZbar2.Enabled = False
FormMAIN.menuStandardListElementalStandardNames.Enabled = False
FormMAIN.menuStandardListOxideStandardNames.Enabled = False
FormMAIN.menuStandardListElementalStandards.Enabled = False
FormMAIN.menuStandardListOxideStandards.Enabled = False

FormMAIN.menuOptionsSearch.Enabled = False
FormMAIN.menuOptionsFind.Enabled = False
FormMAIN.menuOptionsMatch.Enabled = False
FormMAIN.menuOptionsModalAnalysis.Enabled = False
FormMAIN.menuOptionsInterferences.Enabled = False

FormMAIN.menuAnalyticalMQOptions.Enabled = False
FormMAIN.menuAnalyticalPENEPMA.Enabled = False
FormMAIN.menuAnalyticalPENFLUOR.Enabled = False
FormMAIN.menuAnalyticalMQOptions.Enabled = False

Else
FormMAIN.menuStandardNew.Enabled = True
FormMAIN.menuStandardModify.Enabled = True
FormMAIN.menuStandardDuplicate.Enabled = True
FormMAIN.menuStandardDelete.Enabled = True
FormMAIN.menuStandardDeleteSelected.Enabled = True
FormMAIN.menuStandardListStandardNames.Enabled = True
FormMAIN.menuStandardListSelectedStandards.Enabled = True
FormMAIN.menuStandardListAllStandards.Enabled = True
FormMAIN.menuStandardListStandardNamesZbar.Enabled = True
FormMAIN.menuStandardListStandardNamesZbar2.Enabled = True
FormMAIN.menuStandardListElementalStandardNames.Enabled = True
FormMAIN.menuStandardListOxideStandardNames.Enabled = True
FormMAIN.menuStandardListElementalStandards.Enabled = True
FormMAIN.menuStandardListOxideStandards.Enabled = True

FormMAIN.menuOptionsSearch.Enabled = True
FormMAIN.menuOptionsFind.Enabled = True
FormMAIN.menuOptionsMatch.Enabled = True
FormMAIN.menuOptionsModalAnalysis.Enabled = True
FormMAIN.menuOptionsInterferences.Enabled = True

FormMAIN.menuAnalyticalMQOptions.Enabled = True
FormMAIN.menuAnalyticalPENEPMA.Enabled = True
FormMAIN.menuAnalyticalPENFLUOR.Enabled = True
FormMAIN.menuAnalyticalMQOptions.Enabled = True
End If

' Set MAIN window title
FormMAIN.Caption = "Standard (Compositional Database)"
If StandardDataFile$ <> vbNullString Then
FormMAIN.Caption = "Standard [" & StandardDataFile$ & "]"

CalculateAlternativeZbarsFlag% = False
CalculateContinuumAbsorptionFlag% = False
End If

' Set angstrom conversion flag to just use "nominal" spectrometer to angstrom calculations
UseMultiplePeakCalibrationOffsetFlag = False    ' for interference calculations to prevent errors

Exit Sub

' Errors
StanFormUpdateError:
MsgBox Error$, vbOKOnly + vbCritical, "StanFormUpdate"
ierror = True
Exit Sub

End Sub

Sub StanFormCalculateExcessOxygen(analysis As TypeAnalysis, unksample() As TypeSample, stdsample() As TypeSample)
' Calculate excess oxygen for standards

ierror = False
On Error GoTo StanFormCalculateExcessOxygenError

' Calculate excess oxygen if DisplayAsOxideFlag flag is set
If stdsample(1).DisplayAsOxideFlag And stdsample(1).OxygenChannel% > 0 Then

unksample(1).OxideOrElemental% = 1
analysis.ExcessOxygen! = ConvertTotalToExcessOxygen!(Int(2), unksample(), stdsample())
unksample(1).OxideOrElemental% = 2  ' standards are always elemental

analysis.totaloxygen! = stdsample(1).ElmPercents!(stdsample(1).OxygenChannel%)
analysis.CalculatedOxygen! = analysis.totaloxygen! - analysis.ExcessOxygen!
analysis.HalogenCorrectedOxygen! = 0#   ' does not apply to standard compositions (it may or may not be subtracted by user)
OxPercents!(stdsample(1).OxygenChannel%) = analysis.ExcessOxygen!
End If

Exit Sub

' Errors
StanFormCalculateExcessOxygenError:
MsgBox Error$, vbOKOnly + vbCritical, "StanFormCalculateExcessOxygen"
ierror = True
Exit Sub

End Sub

Sub StanFormCalculateOxideAtomic(sample() As TypeSample)
' Calculates oxide and atomic percents for the passed standard sample

ierror = False
On Error GoTo StanFormCalculateOxideAtomicError

Dim i As Integer

' Calculate the totals, oxide percents, atoms, etc.
For i% = 1 To sample(1).LastChan%
If sample(1).DisplayAsOxideFlag Then
OxPercents!(i%) = ConvertElmToOxd(sample(1).ElmPercents!(i%), sample(1).Elsyms$(i%), sample(1).numcat%(i%), sample(1).numoxd%(i%))
End If

'AtPercents!(i%) = ConvertWeightToAtom(sample(1).LastChan%, i%, sample(1).ElmPercents!(), sample(1).Elsyms$())
AtPercents!(i%) = ConvertWeightToAtom2(sample(1).LastChan%, i%, sample(1).ElmPercents!(), sample(1).AtomicWts!(), sample(1).Elsyms$())
Next i%

Exit Sub

' Errors
StanFormCalculateOxideAtomicError:
MsgBox Error$, vbOKOnly + vbCritical, "StanFormCalculateOxideAtomic"
ierror = True
Exit Sub

End Sub

Sub StanFormClear()
' Clear the form

ierror = False
On Error GoTo StanFormClearError

FormMAIN.Caption = vbNullString
FormMAIN.ListAvailableStandards.Clear

FormMAIN.LabelStandard.Caption = vbNullString
FormMAIN.TextLog.Text = vbNullString

FormMAIN.LabelTotal.Caption = vbNullString
FormMAIN.LabelTotalOxygen.Caption = vbNullString
FormMAIN.LabelCalculated.Caption = vbNullString
FormMAIN.LabelExcess.Caption = vbNullString
FormMAIN.LabelAtomic.Caption = vbNullString
FormMAIN.LabelZbar.Caption = vbNullString

Exit Sub

' Errors
StanFormClearError:
MsgBox Error$, vbOKOnly + vbCritical, "StanFormClear"
ierror = True
Exit Sub

End Sub

Sub StanFormDeleteSelectedStandards()
' Delete selected standards

ierror = False
On Error GoTo StanFormDeleteSelectedStandardsError

Dim stdnum As Integer, i As Integer
Dim response As Integer, response2 As Integer

ReDim standardstodelete(1 To MAXINDEX%) As Integer

' Get selected standards from listbox
If FormMAIN.ListAvailableStandards.ListIndex < 0 Then Exit Sub

If FormMAIN.ListAvailableStandards.ListCount < 1 Then GoTo StanFormDeleteSelectedNoStandards

' Warn user about deleting standards
If UCase$(Dir$(StandardDataFile$)) = "STANDARD.MDB" Then
msg$ = "Warning- The user should be aware that deleting a standard composition from the "
msg$ = msg$ & "default Standard Database that is also referenced by a Probe "
msg$ = msg$ & "Database file, will result in the Probe Database file being unusable. Be "
msg$ = msg$ & "sure that the standard to be deleted is not used by any Probe Database files."
response% = MsgBox(msg$, vbOKCancel + vbInformation + vbDefaultButton2, "StanFormDeleteSelectedStandards")
If response% = vbCancel Then
ierror = True
Exit Sub
End If
End If

' Loop on all selected and save standard numbers to delete
For i% = 0 To FormMAIN.ListAvailableStandards.ListCount - 1
If FormMAIN.ListAvailableStandards.Selected(i%) Then
stdnum% = FormMAIN.ListAvailableStandards.ItemData(i%)
standardstodelete%(i% + 1) = stdnum%    ' save standard numbers to delete
End If
Next i%

' Loop and delete standards
For i% = 1 To MAXINDEX%

' Save standard number
If standardstodelete%(i%) > 0 Then

If response% <> vbYesToAll Then
msg$ = "Are you sure that you want to delete standard :"
msg$ = msg$ & vbCrLf & StandardGetString2$(standardstodelete%(i%))
msg$ = msg$ & vbCrLf & " from database :"
msg$ = msg$ & vbCrLf & StandardDataFile$ & "?"
response% = MiscMsgBoxAll(FormMSGBOXALL, "StanFormDeleteSelectedStandards", msg$, CSng(0#))
If response% = vbNo Then GoTo 1000
If response% = vbCancel Then Exit For

' Double confirm with user if yes to all
If response% = vbYesToAll Then
msg$ = "Are you really, really sure that you want to delete all of the selected standards? They really will be lost forever if you haven't backed them up!"
response2% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "StanFormDeleteSelectedStandards")
If response2% = vbCancel Then Exit For
End If
End If

' Delete and update database
If response% = vbYes Or response% = vbYesToAll Then

' Delete the selected standard
Screen.MousePointer = vbHourglass
Call StandardDeleteRecord(standardstodelete%(i%))
Screen.MousePointer = vbDefault
If ierror Then Exit For
End If

End If
1000:
Next i%

' Update the screen
Call StanFormClear

' Reopen the database and re-load available standards
Call StandardGetMDBIndex

' Update the screen
Call StandardLoadList(FormMAIN.ListAvailableStandards)

' Update form
Call StanFormUpdate
If ierror Then Exit Sub

Exit Sub

' Errors
StanFormDeleteSelectedStandardsError:
MsgBox Error$, vbOKOnly + vbCritical, "StanFormDeleteSelectedStandards"
ierror = True
Exit Sub

StanFormDeleteSelectedNoStandards:
msg$ = "No standards entered in standard database yet"
MsgBox msg$, vbOKOnly + vbExclamation, "StanFormDeleteSelectedStandards"
ierror = True
Exit Sub

End Sub

Sub StanFormCalculateBinary()
' Calculate zbars for 50:50 atomic binary compounds

ierror = False
On Error GoTo StanFormCalculateBinaryError

Dim numA As Integer, numB As Integer
Dim numberofbinaries As Integer, inum As Integer
Dim bfilename As String
Dim maszbar As Single, zedzbar As Single

ReDim atoms(1 To 2) As Single
ReDim syms(1 To 2) As String

icancelauto = False

' Init sample
Call InitSample(StanFormOldSample())
If ierror Then Exit Sub

' Calculate for all elements
numberofbinaries% = 0
For numA% = 1 To MAXELM%
For numB% = numA% + 1 To MAXELM%
numberofbinaries% = numberofbinaries% + 1
Next numB%
Next numA%

Call IOWriteLog(vbNullString)
Call IOWriteLog("Number of binaries to be calculated = " & Str$(numberofbinaries%))
Call IOStatusAuto(vbNullString)

' Calculate each binary
inum% = 0
For numA% = 1 To MAXELM%
For numB% = numA% + 1 To MAXELM%
inum% = inum% + 1

msg$ = "Calculating binary " & Format$(inum%, a50$) & " of " & Format$(numberofbinaries%, a50$) & "..."
Call IOStatusAuto(msg$)
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub
End If

' Load conditions
StanFormOldSample(1).Name$ = Trim$(Symup$(numA%)) & "-" & Trim$(Symup$(numB%))
StanFormOldSample(1).LastElm% = 2
StanFormOldSample(1).LastChan% = 2
StanFormOldSample(1).kilovolts = DefaultKiloVolts!
StanFormOldSample(1).takeoff = DefaultTakeOff!

' Load element and default line
StanFormOldSample(1).Elsyms$(1) = Symlo$(numA%)
StanFormOldSample(1).Elsyms$(2) = Symlo$(numB%)
StanFormOldSample(1).Xrsyms$(1) = Deflin$(numA%)
StanFormOldSample(1).Xrsyms$(2) = Deflin$(numB%)

' Fill element arrays
Call ElementLoadArrays(StanFormOldSample())
If ierror Then Exit Sub

Call ElementCheckXray(Int(0), StanFormOldSample())
If ierror Then
msg$ = "Skipping binary " & Format$(inum%, a50$) & "(" & Symlo$(numA%) & "-" & Symlo$(numB%) & ")..."
Call IOWriteLog(msg$)
GoTo 2000
End If

' Load composition
atoms!(1) = 1
atoms!(2) = 1
syms$(1) = StanFormOldSample(1).Elsyms$(1)
syms$(2) = StanFormOldSample(1).Elsyms$(2)
StanFormOldSample(1).ElmPercents!(1) = ConvertAtomToWeight!(2, 1, atoms!(), syms$())
StanFormOldSample(1).ElmPercents!(2) = ConvertAtomToWeight!(2, 2, atoms!(), syms$())

' Load global atomic weights
AtPercents!(1) = 50#
AtPercents!(2) = 50#

' Calculate zbars
Call StanFormCalculateZbars(maszbar!, zedzbar!, StanFormOldSample())
If ierror Then
Call IOStatusAuto(vbNullString)
Exit Sub
End If

' Send composition to MQ input file
If FormMQOPTIONS.Visible Then
bfilename$ = UserDataDirectory$ & "\MQData\" & Trim$(Format$(StanFormOldSample(1).kilovolts!)) & "-" & "BINARY.BAT"
Call MqOptionsSendCompToFile(maszbar!, zedzbar!, bfilename$, StanFormOldSample())
If ierror Then
Call IOStatusAuto(vbNullString)
Exit Sub
End If

End If
2000:
Next numB%
Next numA%

Call IOStatusAuto(vbNullString)
msg$ = "Binary (1:1 atom) composition files created and saved to " & UserDataDirectory$ & "\MQData folder. " & vbCrLf
msg$ = msg$ & "Batch file saved to " & bfilename$ & ". Run this batch file to calculate all "
msg$ = msg$ & "binary composition calculations (this will take a while depending on the number of trajectories specified)."
MsgBox msg$, vbOKOnly + vbInformation, "StanFormCalculateBinary"

Exit Sub

' Errors
StanFormCalculateBinaryError:
MsgBox Error$, vbOKOnly + vbCritical, "StanFormCalculateBinary"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub StanFormDuplicate()
' Duplicate the selected standard and allow user to modify

ierror = False
On Error GoTo StanFormDuplicateError

Dim samplenumber As Integer

' Get standard from listbox
If FormMAIN.ListAvailableStandards.ListIndex < 0 Then Exit Sub
If FormMAIN.ListAvailableStandards.ListCount < 1 Then GoTo StanFormDuplicateNoStandards
samplenumber% = FormMAIN.ListAvailableStandards.ItemData(FormMAIN.ListAvailableStandards.ListIndex)

' Init the sample (for EDS and CL arrays)
Call InitSample(StanFormTmpSample())
If ierror Then Exit Sub

' Get standard from database
Call StandardGetMDBStandard(samplenumber%, StanFormTmpSample())
If ierror Then Exit Sub

' Load as a modified standard
GetCmpFlag% = 3
Call GetCmpLoad(StanFormTmpSample())
If ierror Then Exit Sub

' Standard Database updated from OK button in FormGETCMP (GetCmpSaveAll)
FormGETCMP.Show vbModal
If ierror Then Exit Sub

' Get modified sample
Call GetCmpReturn(StanFormTmpSample())
If ierror Then Exit Sub

' Re-calculate the new standard
samplenumber% = StanFormTmpSample(1).number%
If samplenumber% > 0 Then
Call StanFormCalculate(samplenumber%, Int(0))
If ierror Then Exit Sub
End If

' Reopen the database and re-load available standards
Call StandardGetMDBIndex
If ierror Then Exit Sub

' Update the screen
Call StandardLoadList(FormMAIN.ListAvailableStandards)
If ierror Then Exit Sub

' Re-select current standard
Call StandardSelectList(samplenumber%, FormMAIN.ListAvailableStandards)
If ierror Then Exit Sub

Exit Sub

' Errors
StanFormDuplicateError:
MsgBox Error$, vbOKOnly + vbCritical, "StanFormDuplicate"
ierror = True
Exit Sub

StanFormDuplicateNoStandards:
msg$ = "No standards entered in standard database yet"
MsgBox msg$, vbOKOnly + vbExclamation, "StanFormDuplicate"
ierror = True
Exit Sub

End Sub

Sub StanFormCommandLine(tForm As Form)
' Check for file on command line

ierror = False
On Error GoTo StanFormCommandLineError

Dim tfilename As String
Dim StandardEXEHasRunFlag As Single

' Check for command line string
tfilename$ = Command$()

' No command line, just open default database
If tfilename$ = vbNullString Then

' Check if first time Standard.exe has run
Call InitINIReadWriteScaler(Int(1), ProbeWinINIFile$, "Software", "StandardEXEHasRun", StandardEXEHasRunFlag!)
If ierror Then Exit Sub

' If first time run, clear standard database name and just exit
If StandardEXEHasRunFlag = 0 Then
StandardDataFile$ = vbNullString

' Specify that Standard.exe has run
StandardEXEHasRunFlag = -1
Call InitINIReadWriteScaler(Int(2), ProbeWinINIFile$, "Software", "StandardEXEHasRun", StandardEXEHasRunFlag!)
If ierror Then Exit Sub

Exit Sub
End If

' Open the default standard database
If StandardDataFile$ <> vbNullString Then tfilename$ = StandardDataFile$
Call StandardOpenMDBFile(tfilename$, tForm)
If ierror Then Exit Sub

Exit Sub
End If

' Confirm command line argument
If DebugMode Then Call IOWriteLog(vbCrLf & "Raw Command line argument (MDB file): " & tfilename$)

' Remove double quotes from filename (put there by Command function) which causes Dir$ function to fail
If Left$(tfilename$, 1) = VbDquote And Right$(tfilename$, 1) = VbDquote Then    ' no double quotes if drag and drop
tfilename$ = Mid$(tfilename$, 2)
tfilename$ = Left$(tfilename$, Len(tfilename$) - 1)
End If

' Check for valid file
If Dir$(tfilename$) = vbNullString Then GoTo StanFormCommandLineNotFound
If UCase$(MiscGetFileNameExtensionOnly$(tfilename$)) <> ".MDB" Then GoTo StanFormCommandLineBadExtension

' Open an existing file
Call StandardOpenMDBFile(tfilename$, tForm)

' If error from Open routine, re-initialize all arrays
If ierror Then
Call InitData
If ierror Then Exit Sub
Exit Sub
End If

If DebugMode Then Call IOWriteLog("Command line filename: " & tfilename$)
Exit Sub

' Errors
StanFormCommandLineError:
MsgBox Error$ & " (command line argument: " & tfilename$ & ")", vbOKOnly + vbCritical, "StanFormCommandLine"
ierror = True
Exit Sub

StanFormCommandLineNotFound:
msg$ = "Command line argument string, " & tfilename$ & ", is not a valid file"
MsgBox msg$, vbOKOnly + vbExclamation, "StanFormCommandLine"
ierror = True
Exit Sub

StanFormCommandLineBadExtension:
msg$ = "Command line argument filename, " & tfilename$ & ", does not have an MDB extension"
MsgBox msg$, vbOKOnly + vbExclamation, "StanFormCommandLine"
ierror = True
Exit Sub

End Sub

