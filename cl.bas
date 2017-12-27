Attribute VB_Name = "CodeCL"
' (c) Copyright 1995-2018 by John J. Donovan
Option Explicit

Dim CLDataRow As Integer

Dim CLOldSample(1 To 1) As TypeSample

Sub CLDisplaySpectra(tCLDarkSpectra As Boolean, tForm As Form, datarow As Integer, sample() As TypeSample)
' Display current spectrum from the interface
'  tCLDarkSpectra = false normal CL spectrum
'  tCLDarkSpectra = true dark CL spectrum

ierror = False
On Error GoTo CLDisplaySpectraError

' Check for data
If datarow% = 0 Then Exit Sub

' Save
CLOldSample(1) = sample(1)
CLDataRow% = datarow%

' Check for data
If sample(1).CLSpectraNumberofChannels%(datarow%) < 1 Then Exit Sub

' Call graphics routines
Call CLDisplaySpectra_PE(tCLDarkSpectra, tForm, datarow%, sample())
If ierror Then Exit Sub

Exit Sub

' Errors
CLDisplaySpectraError:
MsgBox Error$, vbOKOnly + vbCritical, "CLDisplaySpectra"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

CLDisplaySpectraZeroAcqTime:
msg$ = "CL acquisition time is zero for datarow " & Format$(datarow%)
IOMsgBox msg$, vbOKOnly + vbExclamation, "CLDisplaySpectra"
ierror = True
Exit Sub

CLDisplaySpectraZeroFraction:
msg$ = "CL dark spectra fraction time is zero for datarow " & Format$(datarow%)
IOMsgBox msg$, vbOKOnly + vbExclamation, "CLDisplaySpectra"
ierror = True
Exit Sub

End Sub

Sub CLInitDisplay(tForm As Form, tCaption As String, datarow As Integer, sample() As TypeSample)
' Init spectrum display

ierror = False
On Error GoTo CLInitDisplayError

Dim astring As String

' Load form caption
If tCaption$ <> vbNullString Then
tForm.Caption = "CL Spectrum Display" & " [" & tCaption$ & "]"
End If

' Load label
astring$ = sample(1).Name$
If sample(1).Linenumber&(datarow%) > 0 Then
astring$ = sample(1).Name$ & ", Line " & Format$(sample(1).Linenumber&(datarow%))                         ' PFE
Else
astring$ = sample(1).CLSpectraKilovolts!(datarow%) & " keV, " & sample(1).Name$                           ' Standard
End If
tForm.LabelSpectrumName.Caption = astring$

' Call graphics routines
Call CLInitDisplay_PE(tForm)
If ierror Then Exit Sub

Exit Sub

' Errors
CLInitDisplayError:
MsgBox Error$, vbOKOnly + vbCritical, "CLInitDisplay"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub CLWriteDiskEMSA(method As Integer, datarow As Integer, sample() As TypeSample, tfilename As String, tForm As Form)
' Write an EMSA format spectrum file based on sample and datarow
'  method = 0 do not ask user to confirm filename
'  method = 1 ask user to confirm file name

ierror = False
On Error GoTo CLWriteDiskEMSAError

' Add extension
If tfilename$ = vbNullString Then tfilename$ = UserDataDirectory$ & "\untitled"

' Confirm filename
If method% = 1 Then
Call IOGetFileName(Int(1), "EMSA", tfilename$, tForm)
If ierror Then Exit Sub

Else
tfilename$ = tfilename$ & "_CL.emsa"
End If

' Export spectrum
Call EMSAWriteSpectrum(Int(1), datarow%, sample(), tfilename$)          ' mode = 1 for CL spectra
If ierror Then Exit Sub

' If no error, save UserData directory
UserDataDirectory$ = MiscGetPathOnly$(tfilename$)

Exit Sub

' Errors
CLWriteDiskEMSAError:
MsgBox Error$, vbOKOnly + vbCritical, "CLWriteDiskEMSA"
ierror = True
Exit Sub

End Sub

Sub CLReadDiskEMSA(datarow As Integer, sample() As TypeSample, tfilename As String, tForm As Form)
' Read an EMSA format spectrum file based on sample and datarow

ierror = False
On Error GoTo CLReadDiskEMSAError

' Add extension
If tfilename$ = vbNullString Then tfilename$ = UserDataDirectory$ & "\untitled"

' Confirm filename
Call IOGetFileName(Int(2), "EMSA", tfilename$, tForm)
If ierror Then Exit Sub

' Export spectrum
Call EMSAReadSpectrum(Int(1), datarow%, sample(), tfilename$)         ' mode = 1 for CL
If ierror Then Exit Sub

' If no error, save UserData directory
UserDataDirectory$ = MiscGetPathOnly$(tfilename$)

Exit Sub

' Errors
CLReadDiskEMSAError:
MsgBox Error$, vbOKOnly + vbCritical, "CLReadDiskEMSA"
ierror = True
Exit Sub

End Sub
