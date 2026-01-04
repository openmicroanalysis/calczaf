Attribute VB_Name = "CodePenepma08d"
' (c) Copyright 1995-2026 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Dim PenepmaTmpSample(1 To 1) As TypeSample

Sub Penepma08ExtractContinuumChannels()
' Extract a range of continuum intensities from the selected Penepma output folder (all sub folders)

ierror = False
On Error GoTo Penepma08ExtractContinuumChannelsError

Dim astring As String, tstring As String
Dim tfilename As String, tMaterialFile As String, tfolder As String

Dim nCount As Long, n As Long, m As Long
Dim sAllFiles() As String

Dim nPoints As Long, xnum As Long
Dim xdata() As Single, ydata() As Single, zdata() As Single
Dim xmin As Single, xmax As Single, xwidth As Single, theta1 As Single, theta2 As Single, phi1 As Single, phi2 As Single
Dim response As Integer

Dim maszbar As Single, zedzbar As Single, zed7zbar As Single

Dim numavg As Long      ' number of channels to average

Const numkeVs& = 10      ' number of output columns (1 to 10 keV)

ReDim atmfrac(1 To MAXCHAN%) As Single
ReDim masfrac(1 To MAXCHAN%) As Single
ReDim zedfrac(1 To MAXCHAN%) As Single

ReDim atemp1(1 To MAXCHAN%) As Integer
ReDim atemp2(1 To MAXCHAN%) As Single

ReDim targetkeVs(1 To numkeVs&) As Single

ReDim averagkeVs(1 To numkeVs&) As Single
ReDim averagInts(1 To numkeVs&) As Single
ReDim averagVars(1 To numkeVs&) As Single

Static tpath As String

icancelauto = False

msg$ = "Do you want to output generated or emitted continuum intensities?" & vbCrLf & vbCrLf
msg$ = msg$ & "CLick Yes for generated continuum intensities or No for emitted continuum intensities."
response% = MsgBox(msg$, vbYesNoCancel + vbQuestion + vbDefaultButton1, "Penepma08ExtractContinuumChannels")
If response% = vbCancel Then Exit Sub

' Browse to a specified folder containing the Penepma output files
tstring$ = "Browse to PENEPMA Output File(s) Folder containing Penepma continuum calculations"
If tpath$ = vbNullString Then tpath$ = PENEPMA_Root$ & "\Penepma\"
tpath$ = IOBrowseForFolderByPath(True, tpath$, tstring$, FormPENEPMA08Batch)
If ierror Then Exit Sub
If Trim$(tpath$) = vbNullString Then Exit Sub

' Get the Penepma generated continuum output files for all files in folder and sub folders
If response% = vbYes Then
Call DirectorySearch("*pe-gen-bremss.dat", tpath$, True, nCount&, sAllFiles$())
Else
Call DirectorySearch("*pe-intens-01.dat", tpath$, True, nCount&, sAllFiles$())
End If
If ierror Then Exit Sub

If nCount& = 0 Then GoTo Penepma08ExtractContinuumChannelsFilesNotFound

' Open continuum intensity output file
Close #Temp2FileNumber%
If response% = vbYes Then
tfilename$ = tpath$ & "\Continuum_Extractions_Generated.dat"
Else
tfilename$ = tpath$ & "\Continuum_Extractions.Emitted.dat"
End If
Open tfilename$ For Output As #Temp2FileNumber%

' Output column labels
astring$ = VbDquote$ & "--Material--" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "Zbar Mass" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "Zbar Z^1.0" & VbDquote$ & vbTab
astring$ = astring$ & VbDquote$ & "Zbar Z^0.7" & VbDquote$ & vbTab

' Intensity labels
For m& = 1 To numkeVs&
targetkeVs!(m&) = CSng(m&)                                                                      ' load values from 1 to 10 keV
astring$ = astring$ & VbDquote$ & Format$(targetkeVs!(m&)) & " keV Intensities" & VbDquote$ & vbTab
Next m&

' Variance labels
For m& = 1 To numkeVs&
targetkeVs!(m&) = CSng(m&)                                                                      ' load values from 1 to 10 keV
astring$ = astring$ & VbDquote$ & Format$(targetkeVs!(m&)) & " keV Variances" & VbDquote$ & vbTab
Next m&

Print #Temp2FileNumber%, astring$
Call IOWriteLog(astring$)

' Loop through all recursively found files
For n& = 1 To nCount&
tfilename$ = sAllFiles$(n&)
If Trim$(tfilename$) <> vbNullString Then

' Get the spectrum intensities (normally only used by PFE for demo mode!)
Screen.MousePointer = vbHourglass
Call Penepma08PenepmaSpectrumRead2(tfilename$, nPoints&, xdata!(), ydata!(), zdata!(), xmin!, xmax!, xnum&, xwidth!, theta1!, theta2!, phi1!, phi2!)
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

' Extract the average intensities for each of the numcol energies
Screen.MousePointer = vbHourglass
Call Penepma08ExtractContinuumChannels2(nPoints&, xdata!(), ydata!(), zdata!(), numkeVs&, targetkeVs!(), averagkeVs!(), averagInts!(), averagVars!())
Screen.MousePointer = vbDefault
If ierror Then Exit Sub

' Get composition from material file
tMaterialFile$ = MiscGetPathOnly$(tfilename$) & "pe-material.dat"
Call Penepma08ReadMaterialFile2(tMaterialFile$, masfrac!(), atmfrac!(), PenepmaTmpSample())
If ierror Then Exit Sub

' Mass fractions
Call StanFormCalculateZbarFrac(Int(1), PenepmaTmpSample(1).LastChan%, atmfrac!(), PenepmaTmpSample(1).AtomicNums%(), atemp1%(), PenepmaTmpSample(1).AtomicWts!(), CSng(1#), masfrac!(), maszbar!)
If ierror Then Exit Sub

' Calculate zed (electron) fractions and zbar
Call StanFormCalculateZbarFrac(Int(0), PenepmaTmpSample(1).LastChan%, atmfrac!(), PenepmaTmpSample(1).AtomicNums%(), PenepmaTmpSample(1).AtomicNums%(), atemp2!(), CSng(1#), zedfrac!(), zedzbar!)
If ierror Then Exit Sub

' Calculate zed^0.7 (electron) fractions and zbar
Call StanFormCalculateZbarFrac(Int(0), PenepmaTmpSample(1).LastChan%, atmfrac!(), PenepmaTmpSample(1).AtomicNums%(), PenepmaTmpSample(1).AtomicNums%(), atemp2!(), CSng(0.7), zedfrac!(), zed7zbar!)
If ierror Then Exit Sub

' Parse out last folder name (add last backslash)
tfolder$ = MiscGetLastFolderOnly$(MiscGetPathOnly$(tfilename$))
If ierror Then Exit Sub

' Load material and Zbars
astring$ = VbDquote$ & tfolder$ & VbDquote$ & vbTab$ & Format$(maszbar!, a80$) & vbTab & Format$(zedzbar!, a80$) & vbTab & Format$(zed7zbar!, a80$) & vbTab

' Load continuum averages
For m& = 1 To numkeVs&
astring$ = astring$ & Format$(averagInts!(m&), a80$) & vbTab
Next m&

' Load continuum averages variances
For m& = 1 To numkeVs&
astring$ = astring$ & Format$(averagVars!(m&), a80$) & vbTab
Next m&

' Write zbars and average continuum intensities to output file
Print #Temp2FileNumber%, astring$
Call IOWriteLog(astring$)

End If
Next n&

Close #Temp2FileNumber%
Screen.MousePointer = vbDefault

Call IOWriteLog("Penepma08ExtractContinuumChannels: Extraction of continuum spectra files in folder " & tpath$ & " is complete!")
Call IOStatusAuto(vbNullString)
Exit Sub

' Errors
Penepma08ExtractContinuumChannelsError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08ExtractContinuumChannels"
ierror = True
Exit Sub

Penepma08ExtractContinuumChannelsFilesNotFound:
Screen.MousePointer = vbDefault
msg$ = "No Penepma spectrum output files found in folder " & tpath$
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08ExtractContinuumChannels"
ierror = True
Exit Sub

End Sub

Sub Penepma08ExtractContinuumChannels2(nPoints As Long, xdata() As Single, ydata() As Single, zdata() As Single, numkeVs As Long, targetkeVs() As Single, averagkeVs() As Single, averagInts() As Single, averagVars() As Single)
' Extract the average intensities based on the target keVs from the passed dataset

ierror = False
On Error GoTo Penepma08ExtractContinuumChannels2Error

Dim n As Long, m As Long, k As Long
Dim avgkeV As Single, avgint As Single, avgvar As Single
Dim temp1 As Single, temp2 As Single, temp3 As Single
Dim target As Single

Const numperside& = 2      ' number of intensities per side to average (total points to average = 2 * 2 + 1)

' Loop through all target keVs
For m& = 1 To numkeVs&
temp1! = 0#
temp2! = 0#
temp3! = 0#

' Find index of closest array to target keV
target! = targetkeVs!(m&) * EVPERKEV#
Call MiscFindClosestMatch&(target!, nPoints&, xdata!(), n&)
If ierror Then Exit Sub

' Now load several points on each side for averaging
If n& > numperside& And n& < nPoints& - numperside& Then
For k& = n& - numperside& To n& + numperside&
temp1! = temp1! + xdata!(k&) / EVPERKEV#
temp2! = temp2! + ydata!(k&)
temp3! = temp3! + zdata!(k&)
Next k&

averagkeVs!(m&) = temp1! / (numperside& * 2 + 1)
averagInts!(m&) = temp2! / (numperside& * 2 + 1)
averagVars!(m&) = temp3! / (numperside& * 2 + 1)
End If

Next m&

Exit Sub

' Errors
Penepma08ExtractContinuumChannels2Error:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08ExtractContinuumChannels2"
ierror = True
Exit Sub

End Sub
