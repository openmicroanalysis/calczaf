Attribute VB_Name = "CodeFIND"
Option Explicit

Global ElementListNumber As Integer

Dim StandardAnalysis As TypeAnalysis

Dim StandardTmpSample(1 To 1) As TypeSample
Dim StandardOldSample(1 To 1) As TypeSample

Sub FindLoad()
' Load the FormFIND
Dim i As Integer

ierror = False
On Error GoTo FindLoadError

' Add the list box items
FormFIND.ComboElement.Clear
For i% = 0 To MAXELM% - 1
FormFIND.ComboElement.AddItem Symlo$(i% + 1)
Next i%

FormFIND.ComboElement.Text = vbNullString
FormFIND.TextLow.Text = Str$(0#)
FormFIND.TextHigh.Text = Str$(100#)

' Load default element if not loaded and set list element
If ElementListNumber% = 0 Then ElementListNumber% = 26
FormFIND.ComboElement.ListIndex = ElementListNumber% - 1

Exit Sub

' Errors
FindLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "FindLoad"
ierror = True
Exit Sub

End Sub

Sub FindStandards(tList As ListBox)
' Routine to find all standards within a range and select them
' in the passed list box

ierror = False
On Error GoTo FindStandardsError

Dim ip As Integer
Dim sym As String
Dim low As Single, high As Single

' Get the element to find and range
sym$ = FormFIND.ComboElement.Text

ip% = IPOS1(MAXELM%, sym$, Symlo$())
If ip% = 0 Then
msg$ = "Element " & sym$ & " is not a valid element symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "FindStandards"
ierror = True
Exit Sub
End If

' Save as default to global
ElementListNumber% = ip%

low! = Val(FormFIND.TextLow.Text)
high! = Val(FormFIND.TextHigh.Text)

If low! < 0# Or low! > 100# Then
msg$ = "Low weight limit of " & Str$(low!) & " is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "FindStandards"
ierror = True
Exit Sub
End If

If high! < 0# Or high! > 100# Then
msg$ = "High weight limit of " & Str$(low!) & " is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "FindStandards"
ierror = True
Exit Sub
End If

If low! = high! Then
msg$ = "Low weight limit equals high weight limit"
MsgBox msg$, vbOKOnly + vbExclamation, "FindStandards"
ierror = True
Exit Sub
End If

' Query the standard database
msg$ = vbCrLf & "Searching standard database for element " & sym$ & "..."
Call IOWriteLog(msg$)
Call StandardFindStandard(sym$, low!, high!, tList)
If ierror Then Exit Sub

Exit Sub

' Errors
FindStandardsError:
MsgBox Error$, vbOKOnly + vbCritical, "FindStandards"
ierror = True
Exit Sub

End Sub

Sub FindStandardsFilter()
' Routine to filter standards in list based on k-ratio to concentration ratio

ierror = False
On Error GoTo FindStandardsFilterError

Dim i As Integer, j As Integer
Dim stdnum As Integer, stdrow As Integer

icancelauto = False

Call IOWriteLog("FindStandardsFilter: Starting standard filter operation...")
For i% = 0 To FormFIND.ListStandards.ListCount% - 1
stdnum% = FormFIND.ListStandards.ItemData(i%)

Screen.MousePointer = vbHourglass
Call StandardGetMDBStandard(stdnum%, StandardTmpSample())
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

' Check the element xrays for hydrogen and helium
Call ElementCheckXray(Int(0), StandardTmpSample())
If ierror Then Exit Sub

' Sort for hydrogen and helium
Screen.MousePointer = vbHourglass
Call GetElmSaveSampleOnly(StandardTmpSample(), Int(0), Int(0))
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

' Set StandardOldSample equal to Tmp sample so k factors and ZAF corrections get loaded in ZAFStd
StandardOldSample(1) = StandardTmpSample(1)

' Initialize arrays
Call InitStandards(StandardAnalysis)
If ierror Then Exit Sub

' Set sample standard assignments so that ZAF and K-factors get loaded
For j% = 1 To StandardTmpSample(1).LastChan%
StandardOldSample(1).StdAssigns%(j%) = stdnum%
Next j%

' Run the calculations on the standard
stdrow% = MAXSTD%
Screen.MousePointer = vbHourglass
Call ZAFStd(stdrow%, StandardAnalysis, StandardOldSample(), StandardTmpSample())
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

If CorrectionFlag% > 0 And CorrectionFlag% < 4 Then
Screen.MousePointer = vbHourglass
Call AFactorStd(stdrow%, StandardAnalysis, StandardOldSample(), StandardTmpSample())
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End If

' Check size of matrix correction for each element in standard
For j% = 1 To StandardTmpSample(1).LastChan%

' Large fluorescence
If FormFIND.OptionGreaterOrLess(0).Value And StandardAnalysis.StdAssignsZAFCors!(4, j%) <> 0# Then
If StandardAnalysis.StdAssignsZAFCors!(4, j%) < 0.95 Then
msg$ = "FindStandardsFilter: large fluorescence correction for " & StandardTmpSample(1).Name$ & " (" & Format$(stdnum%) & "), " & StandardTmpSample(1).Elsyms$(j%) & " " & StandardTmpSample(1).Xrsyms$(j%) & " (" & Format$(StandardAnalysis.StdAssignsPercents!(j%)) & " wt.%), = " & Format$(StandardAnalysis.StdAssignsZAFCors!(4, j%))
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
End If

' Large absorption
Else
If StandardAnalysis.StdAssignsZAFCors!(4, j%) > 1.5 Then
msg$ = "FindStandardsFilter: large absorption correction for " & StandardTmpSample(1).Name$ & " (" & Format$(stdnum%) & "), " & StandardTmpSample(1).Elsyms$(j%) & " " & StandardTmpSample(1).Xrsyms$(j%) & " (" & Format$(StandardAnalysis.StdAssignsPercents!(j%)) & " wt.%), = " & Format$(StandardAnalysis.StdAssignsZAFCors!(4, j%))
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
End If
End If
Next j%

DoEvents
If icancelauto Then
ierror = True
Exit Sub
End If
Next i%

Call IOWriteLog("FindStandardsFilter: Standard filter operation complete!")
Exit Sub

' Errors
FindStandardsFilterError:
MsgBox Error$, vbOKOnly + vbCritical, "FindStandardsFilter"
ierror = True
Exit Sub

End Sub


