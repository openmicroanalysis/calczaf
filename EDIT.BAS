Attribute VB_Name = "CodeEDIT"
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

' X-ray line files from Johnson and White tables
Dim XrayType As Integer ' 1=Xray, 2=Edge, 3=Flur

Dim XrayLastElement As Integer
Dim XrayLastXray As Integer
Dim XrayLastEdge As Integer

Dim XrayLastElementEmitted As Integer
Dim XrayLastXrayEmitted As Integer
Dim XrayLastElementAbsorbed As Integer

Const MAC_FILE_RECORD_LENGTH% = 2400
Const XRAY_FILE_RECORD_LENGTH% = 188

Sub EditConvertElemInfoDat()
' Convert ELEMINFO.DAT (Armstrong) to XLINE.DAT, XEDGE.DAT, XFLUR.DAT

ierror = False
On Error GoTo EditConvertElemInfoDatError

Dim i As Integer, j As Integer, nrec As Integer
Dim astring As String, eleminfofile As String
Dim sym As String
Dim atnum As Integer
Dim atwts As Single

Dim engrow As TypeEnergy
Dim edgrow As TypeEdge
Dim flurow As TypeFlur

' Check for file
eleminfofile$ = ApplicationCommonAppData$ & "ELEMINFO.DAT"
If Dir$(eleminfofile$) = vbNullString Then GoTo EditConvertElemInfoDatNotFound

' Make CITZAF folder is not found
If Dir$(ApplicationCommonAppData$ & "\CITZAF", vbDirectory) = vbNullString Then
MkDir ApplicationCommonAppData$ & "\CITZAF"
End If

' Open x-ray line file (note CITZAF subfolder!)
Open ApplicationCommonAppData$ & "\CITZAF\XLINE.DAT" For Random Access Write As #XLineFileNumber% Len = XRAY_FILE_RECORD_LENGTH%

' Open x-ray edge file
Open ApplicationCommonAppData$ & "\CITZAF\XEDGE.DAT" For Random Access Write As #XEdgeFileNumber% Len = XRAY_FILE_RECORD_LENGTH%

' Open x-ray flur file
Open ApplicationCommonAppData$ & "\CITZAF\XFLUR.DAT" For Random Access Write As #XFlurFileNumber% Len = XRAY_FILE_RECORD_LENGTH%

' Write all zeros to file first
For i% = 1 To MAXELM% + 2
nrec% = i%
Put #XEdgeFileNumber%, nrec%, edgrow
Put #XLineFileNumber%, nrec%, engrow
Put #XFlurFileNumber%, nrec%, flurow
Next i%

' Open JTA element data file
Open eleminfofile$ For Input As #Temp1FileNumber%

' Skip header
Line Input #Temp1FileNumber%, astring$
Line Input #Temp1FileNumber%, astring$
Line Input #Temp1FileNumber%, astring$
Line Input #Temp1FileNumber%, astring$
Line Input #Temp1FileNumber%, astring$
Line Input #Temp1FileNumber%, astring$
Line Input #Temp1FileNumber%, astring$

' Start writing records to direct file (skip first two records for compatibility)
For i% = 1 To 95    ' JTA tables only go to Am
nrec% = i% + 2

Line Input #Temp1FileNumber%, sym$
Call IOWriteLog(sym$)

' Read atomic number, atomic weight, K & L fluorescent yields
Line Input #Temp1FileNumber%, astring$
atnum% = Val(Mid$(astring$, 24, 4))
atwts! = Val(Mid$(astring$, 28, 9))
flurow.fraction!(1) = Val(Mid$(astring, 37, 9))     ' K alpha fluorescence yield
flurow.fraction!(2) = Val(Mid$(astring, 37, 9))     ' K beta fluorescence yield (same as alpha)
flurow.fraction!(3) = Val(Mid$(astring, 46, 9))     ' L alpha fluorescence yield
flurow.fraction!(4) = Val(Mid$(astring, 46, 9))     ' L beta fluorescence yield (same as alpha)

Call IOWriteLog(Str$(atnum%) & Str$(atwts!) & Str$(flurow.fraction!(1)) & Str$(flurow.fraction!(3)))

Put #XFlurFileNumber%, nrec%, flurow

' Read K, L, and M line energies
Line Input #Temp1FileNumber%, astring$
engrow.energy(1) = Val(Mid$(astring, 28, 9))
engrow.energy!(3) = Val(Mid$(astring, 37, 9))
engrow.energy(5) = Val(Mid$(astring, 46, 9))

Call IOWriteLog(Str$(engrow.energy!(1)) & Str$(engrow.energy!(3)) & Str$(engrow.energy!(5)))

For j% = 1 To MAXRAY% - 1
engrow.energy!(j%) = engrow.energy!(j%) * EVPERKEV#
Next j%
        
Put #XLineFileNumber%, nrec%, engrow

' Read K, L-I,II,III edges
Line Input #Temp1FileNumber%, astring$
edgrow.energy!(1) = Val(Mid$(astring, 19, 9))
edgrow.energy!(2) = Val(Mid$(astring, 28, 9))
edgrow.energy!(3) = Val(Mid$(astring, 37, 9))
edgrow.energy!(4) = Val(Mid$(astring, 46, 9))
Line Input #Temp1FileNumber%, astring$
edgrow.energy!(5) = Val(Mid$(astring, 19, 9))
edgrow.energy!(6) = Val(Mid$(astring, 28, 9))
edgrow.energy!(7) = Val(Mid$(astring, 37, 9))
edgrow.energy!(8) = Val(Mid$(astring, 46, 9))
edgrow.energy!(9) = Val(Mid$(astring, 55, 9))

Call IOWriteLog(Str$(edgrow.energy!(1)) & Str$(edgrow.energy!(2)) & Str$(edgrow.energy!(3)) & Str$(edgrow.energy!(4)))
Call IOWriteLog(Str$(edgrow.energy!(5)) & Str$(edgrow.energy!(6)) & Str$(edgrow.energy!(7)) & Str$(edgrow.energy!(8)) & Str$(edgrow.energy!(9)))

For j% = 1 To MAXEDG%
edgrow.energy!(j%) = edgrow.energy!(j%) * EVPERKEV#
Next j%

Put #XEdgeFileNumber%, nrec, edgrow

Next i%

Close #Temp1FileNumber%
Close #XEdgeFileNumber%
Close #XLineFileNumber%
Close #XFlurFileNumber%

' Inform user
msg$ = "File " & eleminfofile$ & " converted to CITZAF\XLINE.DAT, CITZAF\XEDGE.DAT and CITZAF\XFLUR.DAT"
MsgBox msg$, vbOKOnly + vbInformation, "EditConvertElemInfoDat"

Exit Sub

' Errors
EditConvertElemInfoDatError:
MsgBox Error$, vbOKOnly + vbCritical, "EditConvertElemInfoDat"
ierror = True
Close #Temp1FileNumber%
Close #XEdgeFileNumber%
Close #XLineFileNumber%
Close #XFlurFileNumber%
Exit Sub

EditConvertElemInfoDatNotFound:
msg$ = "File " & eleminfofile$ & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "EditConvertElemInfoDat"
ierror = True
Exit Sub

End Sub

Sub EditConvertMACMatDat()
' Convert MACMAT*.DAT (Armstrong) to CITZMU.DAT

ierror = False
On Error GoTo EditConvertMacMatDatError

Dim i As Integer, j As Integer, k As Integer
Dim nrec As Integer, m As Integer, n As Integer
Dim iz As Integer, ix As Integer, im As Integer
Dim astring As String
Dim macmatkfile As String, macmatlfile As String, macmatmfile As String
Dim ilo As Integer, ihi As Integer, iset As Integer
Dim exportfilenumber As Integer

Const elementsperline% = 8

ReDim ielm(1 To elementsperline%) As Integer
ReDim MACs(1 To elementsperline%) As Single

ReDim inputfilenumbers(1 To 3) As Integer   ' K, L and M files

' Check for files
macmatkfile$ = ApplicationCommonAppData$ & "MACMATK.DAT"
If Dir$(macmatkfile$) = vbNullString Then GoTo EditConvertMacMatDatKNotFound
macmatlfile$ = ApplicationCommonAppData$ & "MACMATL.DAT"
If Dir$(macmatlfile$) = vbNullString Then GoTo EditConvertMacMatDatLNotFound
macmatmfile$ = ApplicationCommonAppData$ & "MACMATM.DAT"
If Dir$(macmatmfile$) = vbNullString Then GoTo EditConvertMacMatDatMNotFound

Dim macrow As TypeMu
        
' Load input file numbers
inputfilenumbers%(1) = Temp1FileNumber%
inputfilenumbers%(2) = Temp2FileNumber%
inputfilenumbers%(3) = Temp3FileNumber%

Open macmatkfile$ For Input As #inputfilenumbers%(1)
Open macmatlfile$ For Input As #inputfilenumbers%(2)
Open macmatmfile$ For Input As #inputfilenumbers%(3)

exportfilenumber% = FreeFile()
Open ApplicationCommonAppData$ & "CITZMU.DAT" For Random Access Read Write As #exportfilenumber% Len = MAC_FILE_RECORD_LENGTH%

' Write all zeros to file first
For i% = 1 To MAXELM%
nrec% = i%
Put #exportfilenumber%, nrec%, macrow
Next i%

' Read unit = 1 first, 2 second, and 3 last
For j% = 1 To 3

If j% = 1 Then
Line Input #inputfilenumbers%(j%), astring
Call IOWriteLog(astring$)
Line Input #inputfilenumbers%(j%), astring
Call IOWriteLog(astring$)
Line Input #inputfilenumbers%(j%), astring
Call IOWriteLog(astring$)
Line Input #inputfilenumbers%(j%), astring
Call IOWriteLog(astring$)
Line Input #inputfilenumbers%(j%), astring
Call IOWriteLog(astring$)

ElseIf j% = 2 Then
Line Input #inputfilenumbers%(j%), astring
Call IOWriteLog(astring$)

ElseIf j% = 3 Then
Line Input #inputfilenumbers%(j%), astring
Call IOWriteLog(astring$)
End If

' Get ilo, ihi
Line Input #inputfilenumbers%(j%), astring
ilo% = Val(Mid$(astring$, 1, 3))
ihi% = Val(Mid$(astring$, 4, 5))
Call IOWriteLog(Str$(ilo%) & Str$(ihi%))

' Calculate number of absorber sets (elementsperline% emitters per set)
iset% = Int(((ihi% - ilo%) + elementsperline%) / elementsperline%)

' Loop on each set of 95 absorbers until done
For m% = 1 To iset%

' Get emitter atomic numbers
Line Input #inputfilenumbers%(j%), astring
For k% = 1 To elementsperline%
ielm%(k%) = Val(Mid$(astring$, k% * elementsperline% - elementsperline% / 2, elementsperline%))
Next k%
Call IOWriteLog(astring$)

' Loop on each absorber
For n% = 1 To 95    ' Armstrong file goes to element 95
Line Input #inputfilenumbers%(j%), astring$
Call IOWriteLog(astring$)

' Load absorber atomic number and macs for each emitter
iz% = Val(Mid$(astring$, 1, 3))
For k% = 1 To elementsperline%
MACs!(k%) = Val(Mid$(astring$, k% * elementsperline% - elementsperline% / 2, elementsperline%))

' Determine emitter record number
nrec% = ielm%(k%)
        
' Calculate record offset
If nrec% >= ilo% And nrec% <= ihi% And iz% <= MAXELM% Then
ix% = j% * 2 - 1
im% = ix% + (iz% - 1) * (MAXRAY% - 1)
Get #exportfilenumber%, nrec%, macrow
macrow.mac!(im%) = MACs!(k%)
Put #exportfilenumber%, nrec%, macrow
End If

Next k%
Next n%

Line Input #inputfilenumbers%(j%), astring
Call IOWriteLog(astring$)
Next m%
Next j%

Close #inputfilenumbers%(1)
Close #inputfilenumbers%(2)
Close #inputfilenumbers%(3)
Close #exportfilenumber%

' Inform user
msg$ = "Files :" & vbCrLf
msg$ = msg$ & macmatkfile$ & vbCrLf
msg$ = msg$ & macmatlfile$ & vbCrLf
msg$ = msg$ & macmatmfile$ & vbCrLf
msg$ = msg$ & "converted to " & ApplicationCommonAppData$ & "CITZMU.DAT"
MsgBox msg$, vbOKOnly + vbInformation, "EditConvertMacMatDat"

Exit Sub

' Errors
EditConvertMacMatDatError:
MsgBox Error$, vbOKOnly + vbCritical, "EditConvertMACMatDat"
ierror = True
Close #inputfilenumbers%(1)
Close #inputfilenumbers%(2)
Close #inputfilenumbers%(3)
Close #exportfilenumber%
Exit Sub

EditConvertMacMatDatKNotFound:
msg$ = "File " & macmatkfile$ & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "EditConvertMacMatDat"
ierror = True
Exit Sub

EditConvertMacMatDatLNotFound:
msg$ = "File " & macmatlfile$ & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "EditConvertMacMatDat"
ierror = True
Exit Sub

EditConvertMacMatDatMNotFound:
msg$ = "File " & macmatmfile$ & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "EditConvertMacMatDat"
ierror = True
Exit Sub

End Sub

Sub EditGetMACData(ielm As Integer, iray As Integer, iabsorb As Integer, temp As Single)
' Get specified MAC value from file

ierror = False
On Error GoTo EditGetMACDataError

Dim nrec As Integer, num As Integer

Dim macrow As TypeMu

' Check for valid
If ielm% < 1 Or ielm% > MAXELM% Then Exit Sub
If iray% < 1 Or iray% > MAXRAY% - 1 Then Exit Sub
If iabsorb% < 1 Or iabsorb% > MAXELM% Then Exit Sub

' Open MAC file
MACFile$ = ApplicationCommonAppData$ & macstring2$(MACTypeFlag%) & ".DAT"
Open MACFile$ For Random Access Read As #MACFileNumber% Len = MAC_FILE_RECORD_LENGTH%

nrec% = ielm%
Get #MACFileNumber%, nrec%, macrow
num% = iray% + (iabsorb% - 1) * (MAXRAY% - 1)
temp! = macrow.mac!(num%)
Close #MACFileNumber%

Exit Sub

' Errors
EditGetMACDataError:
MsgBox Error$, vbOKOnly + vbCritical, "EditGetMACData"
ierror = True
Exit Sub

End Sub

Sub EditGetXrayData(ielm As Integer, irayedg As Integer, temp As Single)
' Get specified data value from file

ierror = False
On Error GoTo EditGetXrayDataError

Dim nrec As Integer

Dim engrow As TypeEnergy
Dim edgrow As TypeEdge
Dim flurow As TypeFlur

' Check for valid
If ielm% < 1 Or ielm% > MAXELM% Then Exit Sub

If XrayType% = 1 Or XrayType% = 3 Then
If irayedg% < 1 Or irayedg% > MAXRAY% Then Exit Sub
Else
If irayedg% < 1 Or irayedg% > MAXEDG% Then Exit Sub
End If

' Open x-ray line file
If XrayType% = 1 Then
Open XLineFile$ For Random Access Read As #XLineFileNumber% Len = XRAY_FILE_RECORD_LENGTH%
nrec% = ielm% + 2
Get #XLineFileNumber%, nrec%, engrow
temp! = engrow.energy!(irayedg%)
Close #XLineFileNumber%

' Open x-ray edge file
ElseIf XrayType% = 2 Then
Open XEdgeFile$ For Random Access Read As #XEdgeFileNumber% Len = XRAY_FILE_RECORD_LENGTH%
nrec% = ielm% + 2
Get #XEdgeFileNumber%, nrec%, edgrow
temp! = edgrow.energy!(irayedg%)
Close #XEdgeFileNumber%

' Open x-ray flur file
Else
Open XFlurFile$ For Random Access Read As #XFlurFileNumber% Len = XRAY_FILE_RECORD_LENGTH%
nrec% = ielm% + 2
Get #XFlurFileNumber%, nrec%, flurow
temp! = flurow.fraction!(irayedg%)
Close #XFlurFileNumber%
End If

Exit Sub

' Errors
EditGetXrayDataError:
MsgBox Error$, vbOKOnly + vbCritical, "EditGetXrayData"
ierror = True
Exit Sub

End Sub

Sub EditMACLoad()
' Load MAC edits

ierror = False
On Error GoTo EditMACLoadError

' Load xray edits
Dim i As Integer

' Warn user about making changes
Call EditWarnExpert
If ierror Then Exit Sub

' Allow user to elect MAC table
Call GetZAFAllLoadMAC
If ierror Then Exit Sub
FormMAC.Caption = "Select an Existing MAC File to Edit"
FormMAC.Show vbModal
If icancelload Then
ierror = True
Exit Sub
End If

' Load the edit MAC form
FormEDITMAC.Frame1.Caption = macstring$(MACTypeFlag%)

' Add the list box items
FormEDITMAC.ComboElement.Clear
For i% = 0 To MAXELM% - 1
FormEDITMAC.ComboElement.AddItem Symup$(i% + 1)
Next i%

FormEDITMAC.ComboXRay.Clear
For i% = 0 To MAXRAY% - 1
FormEDITMAC.ComboXRay.AddItem Xraylo$(i% + 1)
Next i%

FormEDITMAC.ComboAbsorber.Clear
For i% = 0 To MAXELM% - 1
FormEDITMAC.ComboAbsorber.AddItem Symup$(i% + 1)
Next i%

' Set index to last element and x-ray
If XrayLastElementEmitted% > 0 Then
FormEDITMAC.ComboElement.ListIndex = XrayLastElementEmitted%
Else
FormEDITMAC.ComboElement.ListIndex = 8 - 1 ' oxygen
End If

If XrayLastXrayEmitted% > 0 Then
FormEDITMAC.ComboXRay.ListIndex = XrayLastXrayEmitted%
Else
FormEDITMAC.ComboXRay.ListIndex = 0    ' Ka
End If

If XrayLastElementAbsorbed% > 0 Then
FormEDITMAC.ComboAbsorber.ListIndex = XrayLastElementAbsorbed%
Else
FormEDITMAC.ComboAbsorber.ListIndex = 26 - 1 ' iron
End If

Exit Sub

' Errors
EditMACLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "EditMACLoad"
ierror = True
Exit Sub

End Sub

Sub EditMACSave()
' Save MAC edits

ierror = False
On Error GoTo EditMACSaveError

Dim Elm As String, ray As String, absorb As String
Dim ip As Integer, ipp As Integer, ippp As Integer
Dim response As Integer
Dim temp1 As Single, temp2 As Single

Elm$ = FormEDITMAC.ComboElement.Text
ip% = IPOS1(MAXELM%, Elm$, Symlo$())
If ip% = 0 Then GoTo EditMACSaveInvalidElement

ray$ = FormEDITMAC.ComboXRay.Text
ipp% = IPOS1(MAXRAY% - 1, ray$, Xraylo$())
If ipp% = 0 Then GoTo EditMACSaveInvalidXray

absorb$ = FormEDITMAC.ComboAbsorber.Text
ippp% = IPOS1(MAXELM%, absorb$, Symlo$())
If ippp% = 0 Then GoTo EditMACSaveInvalidAbsorber

temp2! = Val(FormEDITMAC.TextDataValue.Text)

' Get current data value
Call EditGetMACData(ip%, ipp%, ippp%, temp1!)
If ierror Then Exit Sub

' Save change
msg$ = "Are you sure you want to change the data value for " & Elm$ & " " & ray$ & " absorbed by " & absorb$ & " from " & Str$(temp1!) & " to " & Str$(temp2!) & "?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "EditMACSave")

If response% = vbCancel Then
ierror = True
Exit Sub
End If

' Set current data value
Call EditSetMACData(ip%, ipp%, ippp%, temp2!)
If ierror Then Exit Sub

XrayLastElementEmitted% = ip% - 1
XrayLastXrayEmitted% = ipp% - 1
XrayLastElementAbsorbed% = ippp% - 1

Exit Sub

' Errors
EditMACSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "EditMACSave"
ierror = True
Exit Sub

EditMACSaveInvalidElement:
msg$ = Elm$ & " is an invalid element"
MsgBox msg$, vbOKOnly + vbExclamation, "EditMACSave"
ierror = True
Exit Sub

EditMACSaveInvalidXray:
msg$ = ray$ & " is an invalid xray"
MsgBox msg$, vbOKOnly + vbExclamation, "EditMACSave"
ierror = True
Exit Sub

EditMACSaveInvalidAbsorber:
msg$ = ray$ & " is an invalid absorber"
MsgBox msg$, vbOKOnly + vbExclamation, "EditMACSave"
ierror = True
Exit Sub

End Sub

Sub EditSetMACData(ielm As Integer, iray As Integer, iabsorb As Integer, temp As Single)
' Set specified MAC value to file

ierror = False
On Error GoTo EditSetMACDataError

Dim nrec As Integer, num As Integer

Dim macrow As TypeMu

' Check for valid
If ielm% < 1 Or ielm% > MAXELM% Then Exit Sub
If iray% < 1 Or iray% > MAXRAY% Then Exit Sub
If iabsorb% < 1 Or iabsorb% > MAXELM% Then Exit Sub

' Open MAC file
MACFile$ = ApplicationCommonAppData$ & macstring2$(MACTypeFlag%) & ".DAT"
Open MACFile$ For Random Access Read Write As #MACFileNumber% Len = MAC_FILE_RECORD_LENGTH%

nrec% = ielm%
Get #MACFileNumber%, nrec%, macrow
num% = iray% + (iabsorb% - 1) * (MAXRAY% - 1)
macrow.mac!(num%) = temp!
Put #MACFileNumber%, nrec%, macrow
Close #MACFileNumber%

Exit Sub

' Errors
EditSetMACDataError:
MsgBox Error$, vbOKOnly + vbCritical, "EditSetMACData"
ierror = True
Exit Sub

End Sub

Sub EditSetXrayData(ielm As Integer, irayedg As Integer, temp As Single)
' Set specified data value to file

ierror = False
On Error GoTo EditSetXrayDataError

Dim nrec As Integer

Dim engrow As TypeEnergy
Dim edgrow As TypeEdge
Dim flurow As TypeFlur

' Check for valid
If ielm% < 1 Or ielm% > MAXELM% Then Exit Sub

If XrayType% = 1 Or XrayType% = 3 Then
If irayedg% < 1 Or irayedg% > MAXRAY% Then Exit Sub
Else
If irayedg% < 1 Or irayedg% > MAXEDG% Then Exit Sub
End If

' Open x-ray line file
If XrayType% = 1 Then
Open XLineFile$ For Random Access Read Write As #XLineFileNumber% Len = XRAY_FILE_RECORD_LENGTH%
nrec% = ielm% + 2
Get #XLineFileNumber%, nrec%, engrow
engrow.energy!(irayedg%) = temp!
Put #XLineFileNumber%, nrec%, engrow
Close #XLineFileNumber%

' Open x-ray edge file
ElseIf XrayType% = 2 Then
Open XEdgeFile$ For Random Access Read Write As #XEdgeFileNumber% Len = XRAY_FILE_RECORD_LENGTH%
nrec% = ielm% + 2
Get #XEdgeFileNumber%, nrec%, edgrow
edgrow.energy!(irayedg%) = temp!
Put #XEdgeFileNumber%, nrec%, edgrow
Close #XEdgeFileNumber%

' Open x-ray flur file
Else
Open XFlurFile$ For Random Access Read Write As #XFlurFileNumber% Len = XRAY_FILE_RECORD_LENGTH%
nrec% = ielm% + 2
Get #XFlurFileNumber%, nrec%, flurow
flurow.fraction!(irayedg%) = temp!
Put #XFlurFileNumber%, nrec%, flurow
Close #XFlurFileNumber%
End If

Exit Sub

' Errors
EditSetXrayDataError:
MsgBox Error$, vbOKOnly + vbCritical, "EditSetXrayData"
ierror = True
Exit Sub


End Sub

Sub EditUpdateDataValue()
' Change event from form, update data value

ierror = False
On Error GoTo EditUpdateDataValueError

Dim Elm As String, ray As String
Dim ip As Integer, ipp As Integer
Dim temp As Single

Elm$ = FormEDITXRAY.ComboElement.Text
ip% = IPOS1(MAXELM%, Elm$, Symlo$())
If ip% = 0 Then Exit Sub

' Emission or fluorescent yield
If XrayType% = 1 Or XrayType% = 3 Then
ray$ = FormEDITXRAY.ComboXRay.Text
ipp% = IPOS1(MAXRAY% - 1, ray$, Xraylo$())
If ipp% = 0 Then Exit Sub

' Edge energy
Else
ray$ = FormEDITXRAY.ComboXRay.Text
ipp% = IPOS1(MAXEDG%, ray$, Edglo$())
If ipp% = 0 Then Exit Sub
End If

' Get data value for valid element and xray
Call EditGetXrayData(ip%, ipp%, temp!)
If ierror Then Exit Sub

FormEDITXRAY.TextDataValue.Text = Str$(temp!)

Exit Sub

' Errors
EditUpdateDataValueError:
MsgBox Error$, vbOKOnly + vbCritical, "EditUpdateDataValue"
ierror = True
Exit Sub

End Sub

Sub EditUpdateMACValue()
' Change event from form, update MAC value

ierror = False
On Error GoTo EditUpdateMACValueError

Dim Elm As String, ray As String, absorb As String
Dim ip As Integer, ipp As Integer, ippp As Integer
Dim temp As Single

Elm$ = FormEDITMAC.ComboElement.Text
ip% = IPOS1(MAXELM%, Elm$, Symlo$())
If ip% = 0 Then Exit Sub

ray$ = FormEDITMAC.ComboXRay.Text
ipp% = IPOS1(MAXRAY% - 1, ray$, Xraylo$())
If ipp% = 0 Then Exit Sub

absorb$ = FormEDITMAC.ComboAbsorber.Text
ippp% = IPOS1(MAXELM%, absorb$, Symlo$())
If ippp% = 0 Then Exit Sub

' Get data value for valid element and xray
Call EditGetMACData(ip%, ipp%, ippp%, temp!)
If ierror Then Exit Sub

FormEDITMAC.TextDataValue.Text = Str$(temp!)

Exit Sub

' Errors
EditUpdateMACValueError:
MsgBox Error$, vbOKOnly + vbCritical, "EditUpdateMACValue"
ierror = True
Exit Sub

End Sub

Sub EditWarnExpert()
' Warn the user about changing the data files

ierror = False
On Error GoTo EditWarnExpertError

Dim response As Integer

msg$ = "Warning- changes to the default x-ray data files should only be made by experienced microprobe operators. "
msg$ = msg$ & "Verify all changes after editing to ensure that the changes do not contain any typographical errors. "
msg$ = msg$ & "Are you sure that you want to make changes to the default x-ray data files?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "EditWarnExpert")

If response% = vbCancel Then
ierror = True
Exit Sub
End If

Exit Sub

' Errors
EditWarnExpertError:
MsgBox Error$, vbOKOnly + vbCritical, "EditWarnExpert"
ierror = True
Exit Sub

End Sub

Sub EditXrayLoad(mode As Integer)
' Load xray edits

ierror = False
On Error GoTo EditXrayLoadError

Dim i As Integer

' Warn user about making changes
Call EditWarnExpert
If ierror Then Exit Sub

' Load passed xray type (1=xray,2=edge,3=flur)
XrayType% = mode%

If XrayType% = 1 Then
FormEDITXRAY.Frame1.Caption = "Edit Xray Emission Energies (eV)"
End If

If XrayType% = 2 Then
FormEDITXRAY.Frame1.Caption = "Edit Xray Edge Energies (eV)"
End If

If XrayType% = 3 Then
FormEDITXRAY.Frame1.Caption = "Edit Xray Fluorescence Yields"
End If

' Add the list box items
FormEDITXRAY.ComboElement.Clear
For i% = 0 To MAXELM% - 1
FormEDITXRAY.ComboElement.AddItem Symup$(i% + 1)
Next i%

If XrayType% = 1 Or XrayType% = 3 Then
FormEDITXRAY.ComboXRay.Clear
For i% = 0 To MAXRAY% - 1
FormEDITXRAY.ComboXRay.AddItem Xraylo$(i% + 1)
Next i%

Else
FormEDITXRAY.ComboXRay.Clear
For i% = 0 To MAXEDG% - 1
FormEDITXRAY.ComboXRay.AddItem Edglo$(i% + 1)
Next i%
End If

' Set index to last element and x-ray
If XrayLastElement% > 0 Then
FormEDITXRAY.ComboElement.ListIndex = XrayLastElement%
Else
FormEDITXRAY.ComboElement.ListIndex = 8 - 1 ' oxygen
End If

If XrayType% = 1 Or XrayType% = 3 Then
If XrayLastXray% > 0 Then
FormEDITXRAY.ComboXRay.ListIndex = XrayLastXray%
Else
FormEDITXRAY.ComboXRay.ListIndex = 0    ' Ka
End If

Else
If XrayLastEdge% > 0 Then
FormEDITXRAY.ComboXRay.ListIndex = XrayLastEdge%
Else
FormEDITXRAY.ComboXRay.ListIndex = 0
End If
End If

Exit Sub

' Errors
EditXrayLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "EditXrayLoad"
ierror = True
Exit Sub

End Sub

Sub EditXraySave()
' Save Xray edits

ierror = False
On Error GoTo EditXraySaveError

Dim Elm As String, ray As String
Dim ip As Integer, ipp As Integer
Dim response As Integer
Dim temp1 As Single, temp2 As Single

Elm$ = FormEDITXRAY.ComboElement.Text
ip% = IPOS1(MAXELM%, Elm$, Symlo$())
If ip% = 0 Then GoTo EditXraySaveInvalidElement

If XrayType% = 1 Or XrayType% = 3 Then
ray$ = FormEDITXRAY.ComboXRay.Text
ipp% = IPOS1(MAXRAY% - 1, ray$, Xraylo$())
If ipp% = 0 Then GoTo EditXraySaveInvalidXray

Else
ray$ = FormEDITXRAY.ComboXRay.Text
ipp% = IPOS1(MAXEDG%, ray$, Edglo$())
If ipp% = 0 Then GoTo EditXraySaveInvalidEdge
End If

temp2! = Val(FormEDITXRAY.TextDataValue.Text)

' Get current data value
Call EditGetXrayData(ip%, ipp%, temp1!)
If ierror Then Exit Sub

' Save change
msg$ = "Are you sure you want to change the data value for " & Elm$ & " " & ray$ & " from " & Str$(temp1!) & " to " & Str$(temp2!) & "?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "EditXraySave")

If response% = vbCancel Then
ierror = True
Exit Sub
End If

' Set current data value
Call EditSetXrayData(ip%, ipp%, temp2!)
If ierror Then Exit Sub

XrayLastElement% = ip% - 1

If XrayType% = 1 Or XrayType% = 3 Then
XrayLastXray% = ipp% - 1
Else
XrayLastEdge% = ipp% - 1
End If

Exit Sub

' Errors
EditXraySaveError:
MsgBox Error$, vbOKOnly + vbCritical, "EditXraySave"
ierror = True
Exit Sub

EditXraySaveInvalidElement:
msg$ = Elm$ & " is an invalid element"
MsgBox msg$, vbOKOnly + vbExclamation, "EditXraySave"
ierror = True
Exit Sub

EditXraySaveInvalidXray:
msg$ = ray$ & " is an invalid xray"
MsgBox msg$, vbOKOnly + vbExclamation, "EditXraySave"
ierror = True
Exit Sub

EditXraySaveInvalidEdge:
msg$ = ray$ & " is an invalid edge"
MsgBox msg$, vbOKOnly + vbExclamation, "EditXraySave"
ierror = True
Exit Sub

End Sub

Sub EditMakeNewMACTable(mode As Integer)
' Create McMaster or MAC30 or MACJTA or FFAST or user defined (copy only) Mac table
' mode = 1 make McMaster
' mode = 2 make MAC30
' mode = 3 make MACJTA
' mode = 4 make FFAST
' mode = 5 make USERMAC (user defined)

ierror = False
On Error GoTo EditMakeNewMACTableError

Dim response As Integer
Dim ielm As Integer, iray As Integer, ip As Integer
Dim nrec As Integer, num As Integer

Dim keV As Single, edg As Single
Dim aelastic As Single, ainelastic As Single, aphoto As Single, atotal As Single
Dim tfilename As String

Dim macrow As TypeMu

ReDim g(3, MAXELM%) As Single
ReDim o(9, MAXELM%) As Single

ReDim lines(1 To 12, 1 To MAXELM%) As Double
ReDim edges(1 To 12, 1 To MAXELM%) As Double

icancelauto = False

' Confirm
If mode% = 1 Then
msg$ = "Are you sure you want to create a new MCMASTER.DAT MAC table?"
End If
If mode% = 2 Then
msg$ = "Are you sure you want to create a new MAC30.DAT MAC table?"
End If
If mode% = 3 Then
msg$ = "Are you sure you want to create a new MACJTA.DAT MAC table?"
End If
If mode% = 4 Then
msg$ = "Are you sure you want to create a new FFAST.DAT MAC table?"
End If
If mode% = 5 Then
msg$ = "Are you sure you want to create a new USERMAC.DAT MAC table based on an existing MAC table?"
End If

response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "EditMakeNewMACTable")
If response% = vbCancel Then
ierror = True
Exit Sub
End If

' If creating user defined MAC table double check if it already exists
If mode% = 5 Then
MACFile$ = ApplicationCommonAppData$ & macstring2$(7) & ".DAT"
If Dir$(MACFile$) <> vbNullString Then
msg$ = "A user defined MAC table (" & MACFile$ & ") already exists. Are you sure you want to create a new USERMAC.DAT MAC table and overwrite any changes that you may have manually made to that MAC table?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "EditMakeNewMACTable")
If response% = vbCancel Then
Exit Sub
End If
End If
End If

' If MAC30, load line and edge energies from LINES2.DAT
If mode% = 2 Then
Call AbsorbLoadLINES2DataFile(lines#(), edges#())
If ierror Then Exit Sub
End If

' If MACJTA, load line and edge energies from LINES.DAT
If mode% = 3 Then
Call AbsorbLoadLINESDataFile(g!(), o!())
If ierror Then Exit Sub
End If

' If FFAST, load MAC values from CHANTLER*.DAT into module level table
If mode% = 4 Then
Call AbsorbLoadCHANTLERDataFile
If ierror Then Exit Sub
End If

' If user defined MAC just copy from existing table. Ask user for MAC file to base new file on
If mode% = 5 Then
Call GetZAFAllLoadMAC
If ierror Then Exit Sub
FormMAC.Caption = "Select an Existing MAC File to Create the New User Defined MAC Table From"
FormMAC.Option6(6).Enabled = False      ' disable the user defined MAC table as a choice
FormMAC.Show vbModal
If icancelload Then Exit Sub

' Copy the file only and exit
tfilename$ = ApplicationCommonAppData$ & macstring2$(MACTypeFlag%) & ".DAT"      ' user selected basis for new user defined MAC table
MACFile$ = ApplicationCommonAppData$ & macstring2$(7) & ".DAT"       ' user defined MAC file
FileCopy tfilename$, MACFile$       ' copy the selected file to the new user defined table

msg$ = MACFile$ & " has been successfully created (based on the existing " & tfilename$ & " file). You may now edit it to add your own user defined MAC values or update it using a user defined update text file with the proper format."
MsgBox msg$, vbOKOnly + vbInformation, "EditMakeNewMACTable"
Exit Sub
End If

' Open output file
If mode% = 1 Then
MACFile$ = ApplicationCommonAppData$ & "MCMASTER.DAT"
End If
If mode% = 2 Then
MACFile$ = ApplicationCommonAppData$ & "MAC30.DAT"
End If
If mode% = 3 Then
MACFile$ = ApplicationCommonAppData$ & "MACJTA.DAT"
End If
If mode% = 4 Then
MACFile$ = ApplicationCommonAppData$ & "FFAST.DAT"
End If

Open MACFile$ For Random Access Write As #MACFileNumber% Len = MAC_FILE_RECORD_LENGTH%
Call IOStatusAuto(vbNullString)

' Loop on each element emitter
For ip% = 1 To MAXELM%
nrec% = AllAtomicNums%(ip%)

' Loop on each absorber
For ielm% = 1 To MAXELM%

msg$ = "Calculating MAC for " & Format$(Symlo$(ip%), a20$) & " absorbed by " & Format$(Symlo$(ielm%), a20$) & "..."
Call IOStatusAuto(msg$)
If icancelauto Then
Call IOStatusAuto(vbNullString)
Close #MACFileNumber%
ierror = True
Exit Sub
End If

' Loop on all xrays
For iray% = 1 To MAXRAY% - 1

' Determine energy of emitter
Call XrayGetEnergy(ip%, iray%, keV!, edg!)
If ierror Then Exit Sub

' Get McMaster value
If keV! > 0# Then
If mode% = 1 Then
Call AbsorbGetMAC(ielm%, keV!, aphoto!, aelastic!, ainelastic!, atotal!)
End If

' Get MAC30 value
If mode% = 2 Then
Call AbsorbGetMAC30(keV!, ielm%, ip%, iray%, lines#(), edges#(), atotal!)
End If

' Get MACJTA value
If mode% = 3 Then
Call AbsorbGetMACJTA(keV!, ielm%, ip%, iray%, g!(), o!(), atotal!)
End If

' Get FFAST value from module level table
If mode% = 4 Then
Call AbsorbGetFFAST(ielm%, ip%, iray%, atotal!)
End If

If ierror Then
Close #MACFileNumber%
Exit Sub
End If

Else
atotal! = 0#
End If

' Calculate position in type and load
num% = iray% + (ielm% - 1) * (MAXRAY% - 1)
macrow.mac!(num%) = atotal!

Next iray%
Next ielm%

' Save this emitter
Put #MACFileNumber%, nrec%, macrow
DoEvents
Next ip%

Call IOStatusAuto(vbNullString)
Close #MACFileNumber%
msg$ = MACFile$ & " has been successfully created"
MsgBox msg$, vbOKOnly + vbInformation, "EditMakeNewMACTable"

Exit Sub

' Errors
EditMakeNewMACTableError:
Close #MACFileNumber%
MsgBox Error$, vbOKOnly + vbCritical, "EditMakeNewMACTable"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub EditGetMACEmitterAbsorber()
' Display all database values for a single emitter absorber pair

ierror = False
On Error GoTo EditGetMACEmitterAbsorberError

Dim ielm As Integer, iray As Integer, iabsorb As Integer, itemp As Integer, i As Integer
Dim temp As Single
Dim astring As String, Elm As String, ray As String, absorb As String

Static emitterstring As String

If emitterstring$ = vbNullString Then emitterstring$ = "Si ka Mg"

' Get string from user
msg$ = "Enter the emitter-xray and absorber pair (e.g., Si ka Mg) that you want to display all database values for"
emitterstring$ = InputBox$(msg$, "EditGetMACEmitterAbsorber", emitterstring$)
If emitterstring$ = vbNullString Then Exit Sub

' Parse
astring$ = emitterstring$
Call MiscParseStringToString(astring$, Elm$)
If ierror Then Exit Sub
Call MiscParseStringToString(astring$, ray$)
If ierror Then Exit Sub
Call MiscParseStringToString(astring$, absorb$)
If ierror Then Exit Sub

ielm% = IPOS1(MAXELM%, Elm$, Symlo$())
If ielm% = 0 Then GoTo EditGetMACEmitterAbsorberInvalidElement

iray% = IPOS1(MAXRAY% - 1, ray$, Xraylo$())
If iray% = 0 Then GoTo EditGetMACEmitterAbsorberInvalidXray

iabsorb% = IPOS1(MAXELM%, absorb$, Symlo$())
If iabsorb% = 0 Then GoTo EditGetMACEmitterAbsorberInvalidAbsorber

' Store original MAC file type
itemp% = MACTypeFlag%

' Open MAC file
Call IOWriteLog(vbNullString)
For i% = 1 To MAXMACTYPE%
MACTypeFlag% = i%
MACFile$ = ApplicationCommonAppData$ & macstring2$(MACTypeFlag%) & ".DAT"

If Dir$(MACFile$) <> vbNullString Then
Call EditGetMACData(ielm%, iray%, iabsorb%, temp!)
msg$ = "MAC value for " & Elm$ & " " & ray$ & " in " & absorb$ & " = " & Format$(Format$(temp!, f102$), a100$) & "  (" & macstring$(MACTypeFlag%) & ")"
Call IOWriteLog(msg$)
End If

Next i%

' Restore MAC file
MACTypeFlag% = itemp%
Exit Sub

' Errors
EditGetMACEmitterAbsorberError:
MsgBox Error$, vbOKOnly + vbCritical, "EditGetMACEmitterAbsorber"
ierror = True
Exit Sub

EditGetMACEmitterAbsorberInvalidElement:
msg$ = Elm$ & " is an invalid emitter element"
MsgBox msg$, vbOKOnly + vbExclamation, "EditGetMACEmitterAbsorber"
ierror = True
Exit Sub

EditGetMACEmitterAbsorberInvalidXray:
msg$ = ray$ & " is an invalid xray"
MsgBox msg$, vbOKOnly + vbExclamation, "EditGetMACEmitterAbsorber"
ierror = True
Exit Sub

EditGetMACEmitterAbsorberInvalidAbsorber:
msg$ = absorb$ & " is an invalid absorber element"
MsgBox msg$, vbOKOnly + vbExclamation, "EditGetMACEmitterAbsorber"
ierror = True
Exit Sub

End Sub

Sub EditConvertTextToDAT(tForm As Form)
' Convert element text file to data file

ierror = False
On Error GoTo EditConvertTextToDATError

Dim i As Integer, nrec As Integer, mode As Integer
Dim astring As String, tfilename1 As String, tfilename2 As String

Dim x6row As TypeEnergy
Dim x9row As TypeEdge

' Check for input file
'tfilename1$ = ApplicationCommonAppData$ & "XRAY_EMISSION.TXT"
'tfilename1$ = ApplicationCommonAppData$ & "XRAY_EDGE.TXT"
tfilename1$ = ApplicationCommonAppData$ & "FLUR_YIELD.TXT"
Call IOGetFileName(Int(2), "TXT", tfilename1$, tForm)
If ierror Then Exit Sub

' Check for output file
'tfilename2$ = ApplicationCommonAppData$ & "XLINE.DAT"
'tfilename2$ = ApplicationCommonAppData$ & "XEDGE.DAT"
tfilename2$ = ApplicationCommonAppData$ & "XFLUR.DAT"
Call IOGetFileName(Int(1), "DAT", tfilename2$, tForm)
If ierror Then Exit Sub

If Dir$(tfilename1$) = vbNullString Then GoTo EditConvertTextToDATNotFound

' Determine mode
If InStr(UCase$(tfilename2$), "EDGE") = 0 Then
mode% = 0                               ' emission line energies or fluorescent yields
Else
mode% = 1                               ' edge energies
End If

' Open x-ray file
Open tfilename2$ For Random Access Write As #Temp2FileNumber% Len = XRAY_FILE_RECORD_LENGTH%

' Write all zeros to file first
For i% = 1 To MAXELM% + 2
nrec% = i%
If mode% = 0 Then                               ' emission line energies or fluorescent yields
Put #Temp2FileNumber%, nrec%, x6row
Else                                            ' edge energies
Put #Temp2FileNumber%, nrec%, x9row
End If
Next i%

' Open input file
Open tfilename1$ For Input As #Temp1FileNumber%

' Skip header
Line Input #Temp1FileNumber%, astring$

' Start writing records to direct file (skip first two records for compatibility)
For i% = 1 To MAXELM%
nrec% = i% + 2

' Read symbol and data
Line Input #Temp1FileNumber%, astring$
If mode% = 0 Then                               ' emission line energies or fluorescent yields
x6row.energy!(1) = Val(Mid$(astring, 11, 9))
x6row.energy!(2) = Val(Mid$(astring, 20, 9))
x6row.energy!(3) = Val(Mid$(astring, 29, 9))
x6row.energy!(4) = Val(Mid$(astring, 38, 9))
x6row.energy!(5) = Val(Mid$(astring, 47, 9))
x6row.energy!(6) = Val(Mid$(astring, 56, 9))
Else                                            ' edge energies
x9row.energy!(1) = Val(Mid$(astring, 11, 9))
x9row.energy!(2) = Val(Mid$(astring, 20, 9))
x9row.energy!(3) = Val(Mid$(astring, 29, 9))
x9row.energy!(4) = Val(Mid$(astring, 38, 9))
x9row.energy!(5) = Val(Mid$(astring, 47, 9))
x9row.energy!(6) = Val(Mid$(astring, 56, 9))
x9row.energy!(7) = Val(Mid$(astring, 65, 9))
x9row.energy!(8) = Val(Mid$(astring, 74, 9))
x9row.energy!(9) = Val(Mid$(astring, 83, 9))
End If

If mode% = 0 Then
Put #Temp2FileNumber%, nrec%, x6row
Else
Put #Temp2FileNumber%, nrec%, x9row
End If

Next i%
        
Close #Temp1FileNumber%
Close #Temp2FileNumber%

' Inform user
msg$ = "File " & tfilename1$ & " converted to " & tfilename2$
MsgBox msg$, vbOKOnly + vbInformation, "EditConvertTextToDAT"

Exit Sub

' Errors
EditConvertTextToDATError:
MsgBox Error$, vbOKOnly + vbCritical, "EditConvertTextToDAT"
ierror = True
Close #Temp1FileNumber%
Close #Temp2FileNumber%
Exit Sub

EditConvertTextToDATNotFound:
msg$ = "File " & tfilename1$ & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "EditConvertTextToDAT"
ierror = True
Exit Sub

End Sub

Sub EditUpdateUserMACFile(tForm As Form)
' Update an existing USERMAC table with values from a USERMAC.TXT file

ierror = False
On Error GoTo EditUpdateUserMACFileError

Dim nrec As Integer, num As Integer
Dim ie As Integer, ia As Integer, ix As Integer
Dim line_number As Long
Dim astring As String, tfilename As String
Dim response As Integer

ReDim mac_values(1 To MAXRAY% - 1) As Single

Dim macrow As TypeMu

' Ask user for update file (default = UserMAC.TXT)
tfilename$ = ApplicationCommonAppData$ & "USERMAC.TXT"
Call IOGetFileName(Int(2), "TXT", tfilename$, tForm)
If ierror Then Exit Sub

' If updating user defined MAC table check that it already exists
MACFile$ = ApplicationCommonAppData$ & macstring2$(7) & ".DAT"
If Dir$(MACFile$) = vbNullString Then
msg$ = "A user defined MAC table (" & MACFile$ & ") does not yet exist. Please first create the file by using the Create New User Defined MAC Table menu."
MsgBox msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "EditUpdateUserMACFile"
ierror = True
Exit Sub
End If
       
msg$ = "Are you sure you want to update the existing user defined MAC table (" & MACFile$ & ") using MAC values read from the user defined update text file (" & tfilename$ & ")?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "EditUpdateUserMACFile")
If response% = vbCancel Then
ierror = True
Exit Sub
End If
       
' Open the input (comma, tab or space delimited) and output files
Open tfilename$ For Input As #Temp1FileNumber%
Open MACFile$ For Random Access Read Write As #MACFileNumber% Len = MAC_FILE_RECORD_LENGTH%

' Read first line of column headings
Input #Temp1FileNumber%, astring
Call IOWriteLog(astring$)

' Loop on entries (ie = emitting Z, ia = absorbing Z)
Call IOStatusAuto(vbNullString)
icancelauto = False
line_number& = 0
Do Until EOF(Temp1FileNumber%)
Input #Temp1FileNumber%, ie%, ia%, mac_values!(1), mac_values!(2), mac_values!(3), mac_values!(4), mac_values!(5), mac_values!(6)
astring$ = "IE=" & Format$(ie%) & ", IA= " & Format$(ia%) & ", MACs (Ka, Kb, La, Lb, Ma, Mb) = " & MiscAutoFormat$(mac_values!(1)) & MiscAutoFormat$(mac_values!(2)) & MiscAutoFormat$(mac_values!(3)) & MiscAutoFormat$(mac_values!(4)) & MiscAutoFormat$(mac_values!(5)) & MiscAutoFormat$(mac_values!(6))
Call IOWriteLog(astring$)
line_number& = line_number& + 1

' Check for valid values
If ie% < 1 Or ie% > MAXELM% Then GoTo EditUpdateUserMACFileBadEmitter
If ia% < 1 Or ia% > MAXELM% Then GoTo EditUpdateUserMACFileBadAbsorber

' Determine emitter record number
nrec% = ie%
        
' Calculate record offset
For ix% = 1 To MAXRAY% - 1
If mac_values!(ix%) > 0# Then

' Calculate position in emitting record for this absorber and x-ray
num% = ix% + (ia% - 1) * (MAXRAY% - 1)
Get #MACFileNumber%, nrec%, macrow
macrow.mac!(num%) = mac_values!(ix%)
Put #MACFileNumber%, nrec%, macrow

End If
Next ix%

Call IOStatusAuto("Processing line " & Format$(line_number&) & " of file " & tfilename$ & "...")
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Close #Temp1FileNumber%
Close #MACFileNumber%
ierror = True
Exit Sub
End If
Loop

Call IOStatusAuto(vbNullString)
Close #Temp1FileNumber%
Close #MACFileNumber%

' Inform user
msg$ = "User defined MAC table (" & MACFile$ & ") was updated using MAC values from the user defined update text file(" & tfilename$ & ")"
MsgBox msg$, vbOKOnly + vbInformation, "EditUpdateUserMACFile"
Exit Sub

' Errors
EditUpdateUserMACFileError:
msg$ = Error$ & ". Please consult the documentation for the format of an example UserMAC.TXT update file and create one using a text editor or exported from Excel as a tab delimited file."
MsgBox msg$, vbOKOnly + vbCritical, "EditUpdateUserMACFile"
Call IOStatusAuto(vbNullString)
ierror = True
Close #Temp1FileNumber%
Close #MACFileNumber%
Exit Sub

EditUpdateUserMACFileBadEmitter:
msg$ = "Invalid emitting atomic number. Please consult the documentation for the format of an example UserMAC.TXT update file and create one using a text editor or exported from Excel as a tab delimited file."
MsgBox msg$, vbOKOnly + vbExclamation, "EditUpdateUserMACFile"
Call IOStatusAuto(vbNullString)
ierror = True
Close #Temp1FileNumber%
Close #MACFileNumber%
Exit Sub

EditUpdateUserMACFileBadAbsorber:
msg$ = "Invalid absorbing atomic number. Please consult the documentation for the format of an example UserMAC.TXT update file and create one using a text editor or exported from Excel as a tab delimited file."
MsgBox msg$, vbOKOnly + vbExclamation, "EditUpdateUserMACFile"
Call IOStatusAuto(vbNullString)
ierror = True
Close #Temp1FileNumber%
Close #MACFileNumber%
Exit Sub

End Sub

Sub EditUpdateXFiles(mode As Integer, tForm As Form)
' Update an existing XLINE, XEDGE or XFLUR .DAT table with values from a XLINE, XEDGE or XFLUR .TXT file (only for non zero input values!)

ierror = False
On Error GoTo EditUpdateXFilesError

Dim ix As Integer, iz As Integer
Dim line_number As Long
Dim astring As String, tfilename As String
Dim response As Integer

Dim tvalues(1 To MAXRAY% - 1) As Single

' Ask user for update file (default = XEDGE.TXT)
If mode% = 1 Then tfilename$ = ApplicationCommonAppData$ & "XLINE.TXT"
If mode% = 2 Then tfilename$ = ApplicationCommonAppData$ & "XEDGE.TXT"
If mode% = 3 Then tfilename$ = ApplicationCommonAppData$ & "XFLUR.TXT"
Call IOGetFileName(Int(2), "TXT", tfilename$, tForm)
If ierror Then Exit Sub
       
If mode% = 1 Then msg$ = "Are you sure you want to update the existing x-ray table (" & XLineFile$ & ") using values read from the specified update text file (" & tfilename$ & ")?"
If mode% = 2 Then msg$ = "Are you sure you want to update the existing x-edge table (" & XEdgeFile$ & ") using values read from the specified update text file (" & tfilename$ & ")?"
If mode% = 3 Then msg$ = "Are you sure you want to update the existing x-flur table (" & XFlurFile$ & ") using values read from the specified update text file (" & tfilename$ & ")?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton2, "EditUpdateXFiles")
If response% = vbCancel Then
ierror = True
Exit Sub
End If
       
' Open the input (comma, tab or space delimited) and output files
Open tfilename$ For Input As #Temp1FileNumber%
If mode% = 1 Then Open XLineFile$ For Random Access Read Write As #Temp2FileNumber% Len = MAC_FILE_RECORD_LENGTH%
If mode% = 2 Then Open XEdgeFile$ For Random Access Read Write As #Temp2FileNumber% Len = MAC_FILE_RECORD_LENGTH%
If mode% = 3 Then Open XFlurFile$ For Random Access Read Write As #Temp2FileNumber% Len = MAC_FILE_RECORD_LENGTH%

' Read first line of column headings
Line Input #Temp1FileNumber%, astring
Call IOWriteLog(astring$)

' Loop on entries (ie = emitting Z, ia = absorbing Z)
Call IOStatusAuto(vbNullString)
icancelauto = False
line_number& = 0
Do Until EOF(Temp1FileNumber%)
Input #Temp1FileNumber%, iz%, tvalues!(1), tvalues!(2), tvalues!(3), tvalues!(4), tvalues!(5), tvalues!(6)
If mode% = 1 Then astring$ = "IZ=" & Format$(iz%) & ", XLINE (Ka, Kb, La, Lb, Ma, Mb) = " & MiscAutoFormat$(tvalues!(1)) & MiscAutoFormat$(tvalues!(2)) & MiscAutoFormat$(tvalues!(3)) & MiscAutoFormat$(tvalues!(4)) & MiscAutoFormat$(tvalues!(5)) & MiscAutoFormat$(tvalues!(6))
If mode% = 2 Then astring$ = "IZ=" & Format$(iz%) & ", XEDGE (Ka, Kb, La, Lb, Ma, Mb) = " & MiscAutoFormat$(tvalues!(1)) & MiscAutoFormat$(tvalues!(2)) & MiscAutoFormat$(tvalues!(3)) & MiscAutoFormat$(tvalues!(4)) & MiscAutoFormat$(tvalues!(5)) & MiscAutoFormat$(tvalues!(6))
If mode% = 3 Then astring$ = "IZ=" & Format$(iz%) & ", XFLUR (Ka, Kb, La, Lb, Ma, Mb) = " & MiscAutoFormat$(tvalues!(1)) & MiscAutoFormat$(tvalues!(2)) & MiscAutoFormat$(tvalues!(3)) & MiscAutoFormat$(tvalues!(4)) & MiscAutoFormat$(tvalues!(5)) & MiscAutoFormat$(tvalues!(6))
Call IOWriteLog(astring$)
line_number& = line_number& + 1

' Check for valid values
If iz% < 1 Or iz% > MAXELM% Then GoTo EditUpdateXFilesBadEmitter
      
' Set updated data value
For ix% = 1 To MAXRAY% - 1
If tvalues!(ix%) > 0# Then
Call EditSetXrayData(iz%, ix%, tvalues!(ix%))
If ierror Then Exit Sub
End If
Next ix%

Call IOStatusAuto("Processing line " & Format$(line_number&) & " of file " & tfilename$ & "...")
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Close #Temp1FileNumber%
Close #Temp2FileNumber%
ierror = True
Exit Sub
End If
Loop

Call IOStatusAuto(vbNullString)
Close #Temp1FileNumber%
Close #Temp2FileNumber%

' Inform user
If mode% = 1 Then msg$ = "Table (" & XLineFile$ & ") was updated using values from the user update text file(" & tfilename$ & ")"
If mode% = 2 Then msg$ = "Table (" & XEdgeFile$ & ") was updated using values from the user update text file(" & tfilename$ & ")"
If mode% = 3 Then msg$ = "Table (" & XFlurFile$ & ") was updated using values from the user update text file(" & tfilename$ & ")"
MsgBox msg$, vbOKOnly + vbInformation, "EditUpdateXFiles"

Exit Sub

' Errors
EditUpdateXFilesError:
msg$ = Error$ & ". Please consult the documentation for the format of an example .TXT update file and create one using a text editor or exported from Excel as a tab delimited file."
MsgBox msg$, vbOKOnly + vbCritical, "EditUpdateXFiles"
Call IOStatusAuto(vbNullString)
ierror = True
Close #Temp1FileNumber%
Close #Temp2FileNumber%
Exit Sub

EditUpdateXFilesBadEmitter:
If mode% = 1 Then msg$ = "Invalid atomic number. Please consult the documentation for the format of an example XLINE.TXT update file and create one using a text editor or exported from Excel as a tab delimited file."
If mode% = 2 Then msg$ = "Invalid atomic number. Please consult the documentation for the format of an example XEDGE.TXT update file and create one using a text editor or exported from Excel as a tab delimited file."
If mode% = 3 Then msg$ = "Invalid atomic number. Please consult the documentation for the format of an example XFLUR.TXT update file and create one using a text editor or exported from Excel as a tab delimited file."
MsgBox msg$, vbOKOnly + vbExclamation, "EditUpdateXFiles"
Call IOStatusAuto(vbNullString)
ierror = True
Close #Temp1FileNumber%
Close #Temp2FileNumber%
Exit Sub

End Sub

