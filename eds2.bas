Attribute VB_Name = "CodeEDS2"
' (c) Copyright 1995-2017 by John J. Donovan
Option Explicit

Dim EDSKLMCounter As Long

Dim num As Long
Dim sarray() As String, xarray() As String
Dim iarray() As Single, earray() As Single

Dim total_num As Long
Dim total_sarray() As String, total_xarray() As String

Dim EDSKLMIndex As Integer, EDSKLMElement As Integer
Dim EDSDataRow As Integer

Dim EDSOldSample(1 To 1) As TypeSample

Sub EDSDisplaySpectra(tForm As Form, datarow As Integer, sample() As TypeSample)
' Display current spectrum from the interface

ierror = False
On Error GoTo EDSDisplaySpectraError

' Save elements for KLM markers
EDSOldSample(1) = sample(1)
EDSDataRow% = datarow%

' Check for data
If sample(1).EDSSpectraNumberofChannels%(datarow%) < 1 Then Exit Sub

' Call graphics routines
Call EDSDisplaySpectra_PE(tForm, datarow%, sample())
If ierror Then Exit Sub

Exit Sub

' Errors
EDSDisplaySpectraError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSDisplaySpectra"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub EDSInitDisplay(tForm As Form, tCaption As String, datarow As Integer, sample() As TypeSample)
' Init spectrum display

ierror = False
On Error GoTo EDSInitDisplayError

Dim astring As String

' Load form caption
If tCaption$ <> vbNullString Then
tForm.Caption = "EDS Spectrum Display" & " [" & tCaption$ & "]"
End If

' Load sample label
astring$ = sample(1).Name$
If datarow% > 0 Then
astring$ = sample(1).Name$ & ", Line " & Format$(sample(1).Linenumber&(datarow%))                          ' PFE
Else
astring$ = sample(1).EDSSpectraAcceleratingVoltage!(datarow%) & " keV , " & sample(1).Name$                ' Standard
End If
tForm.LabelSpectrumName.Caption = astring$

' Load KLM controls
Call EDSLoadKLM(tForm)
If ierror Then Exit Sub

' Call graphics routines
Call EDSInitDisplay_PE(tForm, tCaption$)
If ierror Then Exit Sub

Exit Sub

' Errors
EDSInitDisplayError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSInitDisplay"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub EDSPlotKLM1(tForm As Form)
' Plot specified KLM lines

ierror = False
On Error GoTo EDSPlotKLM1Error

Dim i As Integer, ip As Integer, chan As Integer
Dim n As Long
Dim sym As String

Dim xstart As Single, xstop As Single

EDSKLMCounter& = 0                              ' re-set KLM annotation counter
total_num& = 0                                  ' re-set KLM symbol and xray counter

' No KLM markers
If tForm.OptionKLM(0).Value = True Then Exit Sub

' Get keV start and stop
xstart! = Val(tForm.TextStartkeV.Text)
xstop! = Val(tForm.TextStopkeV.Text)

' Analyzed elements
If tForm.OptionKLM(1).Value Then

' Plot analyzed and not analyzed elements (skip disable quant elements)
For chan% = 1 To EDSOldSample(1).LastChan%
If EDSOldSample(1).DisableQuantFlag%(chan%) = 0 Then
ip% = IPOS1(MAXELM%, EDSOldSample(1).Elsyms$(chan%), Symlo$())

If ip% > 0 Then
Call XrayLoadDatabase2(Int(1), ip%, xstart!, xstop!, num&, sarray$(), xarray$(), iarray!(), earray!())
If ierror Then Exit Sub

If VerboseMode Then
For n& = 1 To num&
msg$ = "EDSPlotKLM1: " & Str$(n&) & " " & sarray$(n&) & " " & xarray$(n&) & " " & Str$(iarray!(n&)) & " " & MiscAutoFormat$(earray!(n&))
Call IOWriteLog(msg$)
Next n&
End If

Call EDSPlotKLM2(tForm, num&, sarray$(), xarray$(), iarray!(), earray!(), EDSKLMCounter&)
If ierror Then Exit Sub

' Load into total arrays for zoom
For n& = 1 To num&
total_num& = total_num& + 1
ReDim Preserve total_sarray(1 To total_num&) As String
ReDim Preserve total_xarray(1 To total_num&) As String
total_sarray$(total_num&) = sarray$(n&)
total_xarray$(total_num&) = xarray$(n&)
Next n&
End If

End If
Next chan%
End If

' All elements
If tForm.OptionKLM(2).Value Then
For i% = 1 To MAXELM%
ip% = i%

Call XrayLoadDatabase2(Int(1), ip%, xstart!, xstop!, num&, sarray$(), xarray$(), iarray!(), earray!())
If ierror Then Exit Sub

Call EDSPlotKLM2(tForm, num&, sarray$(), xarray$(), iarray!(), earray!(), EDSKLMCounter&)
If ierror Then Exit Sub

' Load into total arrays for zoom
For n& = 1 To num&
total_num& = total_num& + 1
ReDim Preserve total_sarray(1 To total_num&) As String
ReDim Preserve total_xarray(1 To total_num&) As String
total_sarray$(total_num&) = sarray$(n&)
total_xarray$(total_num&) = xarray$(n&)
Next n&

Next i%
End If

' Specific element
If tForm.OptionKLM(3).Value Then
sym$ = tForm.ComboSpecificElement.Text  ' specific element atomic number
ip% = IPOS1(MAXELM%, sym$, Symlo$())
If ip% > 0 Then
Call XrayLoadDatabase2(Int(1), ip%, xstart!, xstop!, num&, sarray$(), xarray$(), iarray!(), earray!())
If ierror Then Exit Sub

If VerboseMode Then
For n& = 1 To num&
msg$ = "EDSPlotKLM1: " & Str$(n&) & " " & sarray$(n&) & " " & xarray$(n&) & " " & Str$(iarray!(n&)) & " " & MiscAutoFormat$(earray!(n&))
Call IOWriteLog(msg$)
Next n&
End If

Call EDSPlotKLM2(tForm, num&, sarray$(), xarray$(), iarray!(), earray!(), EDSKLMCounter&)
If ierror Then Exit Sub

' Load into total arrays for zoom
For n& = 1 To num&
total_num& = total_num& + 1
ReDim Preserve total_sarray(1 To total_num&) As String
ReDim Preserve total_xarray(1 To total_num&) As String
total_sarray$(total_num&) = sarray$(n&)
total_xarray$(total_num&) = xarray$(n&)
Next n&

End If
End If

Exit Sub

' Errors
EDSPlotKLM1Error:
MsgBox Error$, vbOKOnly + vbCritical, "EDSPlotKLM1"
ierror = True
Exit Sub

End Sub

Sub EDSLoadKLM(tForm As Form)
' Load the KLM controls

ierror = False
On Error GoTo EDSLoadKLMError

Dim i As Integer

' Load combo list for KLM element
tForm.ComboSpecificElement.Clear
For i% = 1 To MAXELM%
tForm.ComboSpecificElement.AddItem Symup$(i%)
Next i%

' Load default KLM element and index
If EDSKLMElement% = 0 Then EDSKLMElement% = 14   ' default to silicon
tForm.ComboSpecificElement.ListIndex = EDSKLMElement% - 1
tForm.OptionKLM(EDSKLMIndex%).Value = True

Exit Sub

' Errors
EDSLoadKLMError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSLoadKLM"
ierror = True
Exit Sub

End Sub

Sub EDSSaveKLM(tForm As Form)
' Save the KLM controls

ierror = False
On Error GoTo EDSSaveKLMError

Dim i As Integer

For i% = 0 To 3
If tForm.OptionKLM(i%).Value Then EDSKLMIndex% = i%
Next i%

' Save the EDS specific element
EDSKLMElement% = tForm.ComboSpecificElement.ListIndex% + 1

Exit Sub

' Errors
EDSSaveKLMError:
IOMsgBox Error$, vbOKOnly + vbCritical, "EDSSaveKLM"
ierror = True
Exit Sub

End Sub

Sub EDSDisplayRedraw(tForm As Form)
' Redraw the display, if graph has data

ierror = False
On Error GoTo EDSDisplayRedrawError

If EDSDataRow% < 1 Then Exit Sub
If EDSOldSample(1).EDSSpectraNumberofChannels%(EDSDataRow%) < 1 Then Exit Sub

' Call graphics routines
Call EDSDisplayRedraw_PE(tForm)
If ierror Then Exit Sub

Exit Sub

' Errors
EDSDisplayRedrawError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSDisplayRedraw"
ierror = True
Exit Sub

End Sub

Sub EDSWriteDiskEMSA(method As Integer, datarow As Integer, sample() As TypeSample, tfilename As String, tForm As Form)
' Write an EMSA format spectrum file based on sample and datarow
'  method = 0 do not ask user to confirm filename
'  method = 1 ask user to confirm file name

ierror = False
On Error GoTo EDSWriteDiskEMSAError

' Add extension
If tfilename$ = vbNullString Then tfilename$ = UserDataDirectory$ & "\untitled"

' Confirm filename
If method% = 1 Then
Call IOGetFileName(Int(1), "EMSA", tfilename$, tForm)
If ierror Then Exit Sub

Else
tfilename$ = tfilename$ & "_EDS.emsa"
End If

' Export spectrum
Call EMSAWriteSpectrum(Int(0), datarow%, sample(), tfilename$)          ' mode = 0 is EDS spectra
If ierror Then Exit Sub

' If no error, save UserData directory
UserDataDirectory$ = MiscGetPathOnly$(tfilename$)

Exit Sub

' Errors
EDSWriteDiskEMSAError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSWriteDiskEMSA"
ierror = True
Exit Sub

End Sub

Sub EDSReadDiskEMSA(datarow As Integer, sample() As TypeSample, tfilename As String, tForm As Form)
' Read an EMSA format spectrum file based on sample and datarow

ierror = False
On Error GoTo EDSReadDiskEMSAError

' Add extension
If tfilename$ = vbNullString Then tfilename$ = UserDataDirectory$ & "\untitled"

' Confirm filename
Call IOGetFileName(Int(2), "EMSA", tfilename$, tForm)
If ierror Then Exit Sub

' Export spectrum
Call EMSAReadSpectrum(Int(0), datarow%, sample(), tfilename$)
If ierror Then Exit Sub

' If no error, save UserData directory
UserDataDirectory$ = MiscGetPathOnly$(tfilename$)

Exit Sub

' Errors
EDSReadDiskEMSAError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSReadDiskEMSA"
ierror = True
Exit Sub

End Sub

Sub EDSPlotKLM2(tForm As Form, num As Long, sarray() As String, xarray() As String, iarray() As Single, earray() As Single, EDSKLMCounter As Long)
' Plot the actual KLM markers

ierror = False
On Error GoTo EDSPlotKLM2Error

' Call graphics routines
Call EDSPlotKLM2_PE(tForm, num&, sarray$(), xarray$(), iarray!(), earray!(), EDSKLMCounter&)
If ierror Then Exit Sub

Exit Sub

' Errors
EDSPlotKLM2Error:
MsgBox Error$, vbOKOnly + vbCritical, "EDSPlotKLM2"
ierror = True
Exit Sub

End Sub

Sub EDSRePlotKLM(tForm As Form)
' Re plot specified KLM lines and modify x-ray text based on current zoom

ierror = False
On Error GoTo EDSRePlotKLMError

' Re-load KLM annotations with x-ray text if KLM annotation height is larger then 1/10 current Y zoom
Call EDSRePlotKLM_PE(tForm, total_num&, total_sarray$(), total_xarray$(), EDSKLMCounter&)
If ierror Then Exit Sub

Exit Sub

' Errors
EDSRePlotKLMError:
MsgBox Error$, vbOKOnly + vbCritical, "EDSRePlotKLM"
ierror = True
Exit Sub

End Sub

