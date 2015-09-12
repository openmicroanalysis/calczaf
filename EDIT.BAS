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

Dim ElementNames(1 To MAXELM%) As String

Sub EditLoadFullElementNames()
' Dang. Never had to do this before!

ierror = False
On Error GoTo EditLoadFullElementNamesError

ElementNames$(1) = "Hydrogen"
ElementNames$(2) = "Helium"
ElementNames$(3) = "Lithium"
ElementNames$(4) = "Beryllium"
ElementNames$(5) = "Boron"
ElementNames$(6) = "Carbon"
ElementNames$(7) = "Nitrogen"
ElementNames$(8) = "Oxygen"
ElementNames$(9) = "Fluorine"
ElementNames$(10) = "Neon"
ElementNames$(11) = "Sodium"
ElementNames$(12) = "Magnesium"
ElementNames$(13) = "Aluminum"
ElementNames$(14) = "Silicon"
ElementNames$(15) = "Phosphorus"
ElementNames$(16) = "Sulfur"
ElementNames$(17) = "Chlorine"
ElementNames$(18) = "Argon"
ElementNames$(19) = "Potassium"
ElementNames$(20) = "Calcium"
ElementNames$(21) = "Scandium"
ElementNames$(22) = "Titanium"
ElementNames$(23) = "Vanadium"
ElementNames$(24) = "Chromium"
ElementNames$(25) = "Manganese"
ElementNames$(26) = "Iron"
ElementNames$(27) = "Cobalt"
ElementNames$(28) = "Nickel"
ElementNames$(29) = "Copper"
ElementNames$(30) = "Zinc"
ElementNames$(31) = "Gallium"
ElementNames$(32) = "Germanium"
ElementNames$(33) = "Arsenic"
ElementNames$(34) = "Selenium"
ElementNames$(35) = "Bromine"
ElementNames$(36) = "Krypton"
ElementNames$(37) = "Rubidium"
ElementNames$(38) = "Strontium"
ElementNames$(39) = "Yttrium"
ElementNames$(40) = "Zirconium"
ElementNames$(41) = "Niobium"
ElementNames$(42) = "Molybdenum"
ElementNames$(43) = "Technetium"
ElementNames$(44) = "Ruthenium"
ElementNames$(45) = "Rhodium"
ElementNames$(46) = "Palladium"
ElementNames$(47) = "Silver"
ElementNames$(48) = "Cadmium"
ElementNames$(49) = "Indium"
ElementNames$(50) = "Tin"
ElementNames$(51) = "Antimony"
ElementNames$(52) = "Tellurium"
ElementNames$(53) = "Iodine"
ElementNames$(54) = "Xenon"
ElementNames$(55) = "Cesium"
ElementNames$(56) = "Barium"
ElementNames$(57) = "Lanthanum"
ElementNames$(58) = "Cerium"
ElementNames$(59) = "Praseodymium"
ElementNames$(60) = "Neodymium"
ElementNames$(61) = "Promethium"
ElementNames$(62) = "Samarium"
ElementNames$(63) = "Europium"
ElementNames$(64) = "Gadolinium"
ElementNames$(65) = "Terbium"
ElementNames$(66) = "Dysprosium"
ElementNames$(67) = "Holmium"
ElementNames$(68) = "Erbium"
ElementNames$(69) = "Thulium"
ElementNames$(70) = "Ytterbium"
ElementNames$(71) = "Lutetium"
ElementNames$(72) = "Hafnium"
ElementNames$(73) = "Tantalum"
ElementNames$(74) = "Tungsten"
ElementNames$(75) = "Rhenium"
ElementNames$(76) = "Osmium"
ElementNames$(77) = "Iridium"
ElementNames$(78) = "Platinum"
ElementNames$(79) = "Gold"
ElementNames$(80) = "Mercury"
ElementNames$(81) = "Thallium"
ElementNames$(82) = "Lead"
ElementNames$(83) = "Bismuth"
ElementNames$(84) = "Polonium"
ElementNames$(85) = "Astatine"
ElementNames$(86) = "Radon"
ElementNames$(87) = "Francium"
ElementNames$(88) = "Radium"
ElementNames$(89) = "Actinium"
ElementNames$(90) = "Thorium"
ElementNames$(91) = "Protactinium"
ElementNames$(92) = "Uranium"
ElementNames$(93) = "Neptunium"
ElementNames$(94) = "Plutonium"
ElementNames$(95) = "Americium"
ElementNames$(96) = "Curium"
ElementNames$(97) = "Berkelium"
ElementNames$(98) = "Californium"
ElementNames$(99) = "Einsteinium"
ElementNames$(100) = "Fermium"

Exit Sub

' Errors
EditLoadFullElementNamesError:
MsgBox Error$, vbOKOnly + vbCritical, "EditLoadFullElementNames"
ierror = True
Exit Sub

End Sub

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
inputfilenumbers%(1) = FreeFile()
Open macmatkfile$ For Input As #inputfilenumbers%(1)
inputfilenumbers%(2) = FreeFile()
Open macmatlfile$ For Input As #inputfilenumbers%(2)
inputfilenumbers%(3) = FreeFile()
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
'   mode% = 0 emission line energies
'   mode% = 1 edge energies
'   mode% = 2 fluorescent yields

ierror = False
On Error GoTo EditConvertTextToDATError

Dim nrec As Integer, mode As Integer
Dim astring As String, tfilename1 As String, tfilename2 As String

Dim x6row As TypeEnergy
Dim x9row As TypeEdge

' Check for input file
'tfilename1$ = ApplicationCommonAppData$ & "XRAY_LINE.TXT"
'tfilename1$ = ApplicationCommonAppData$ & "XRAY_EDGE.TXT"
tfilename1$ = ApplicationCommonAppData$ & "XRAY_FLUR.TXT"
Call IOGetFileName(Int(2), "TXT", tfilename1$, tForm)
If ierror Then Exit Sub

' Check for output file
'tfilename2$ = ApplicationCommonAppData$ & "XLINE_TMP.DAT"
'tfilename2$ = ApplicationCommonAppData$ & "XEDGE_TMP.DAT"
tfilename2$ = ApplicationCommonAppData$ & "XFLUR_TMP.DAT"
Call IOGetFileName(Int(1), "DAT", tfilename2$, tForm)
If ierror Then Exit Sub

If Dir$(tfilename1$) = vbNullString Then GoTo EditConvertTextToDATNotFound

' Determine mode
mode% = -1
If InStr(UCase$(tfilename2$), "LINE") > 0 Then mode% = 0                               ' emission line energies
If InStr(UCase$(tfilename2$), "EDGE") > 0 Then mode% = 1                               ' edge energies
If InStr(UCase$(tfilename2$), "FLUR") > 0 Then mode% = 2                               ' fluorescent yields
If mode% = -1 Then GoTo EditConvertTextToDATUnknownFormat

' Open output x-ray file
Open tfilename2$ For Random Access Write As #Temp2FileNumber% Len = XRAY_FILE_RECORD_LENGTH%

' Write all zeros to file first
For nrec% = 1 To MAXELM% + 2
If mode% = 0 Then                               ' emission line energies
Put #Temp2FileNumber%, nrec%, x6row
ElseIf mode% = 1 Then                           ' edge energies
Put #Temp2FileNumber%, nrec%, x9row
ElseIf mode% = 2 Then                           ' fluorescent yields
Put #Temp2FileNumber%, nrec%, x6row
End If
Next nrec%

' Open input file
Open tfilename1$ For Input As #Temp1FileNumber%

' Skip header line
Line Input #Temp1FileNumber%, astring$

' Start writing records to direct file (skip first two records for compatibility)
For nrec% = 3 To MAXELM% + 2

' Read symbol and data
Line Input #Temp1FileNumber%, astring$
If mode% = 0 Then                               ' emission line energies or fluorescent yields
x6row.energy!(1) = Val(Mid$(astring, 11, 9)) * EVPERKEV#
x6row.energy!(2) = Val(Mid$(astring, 20, 9)) * EVPERKEV#
x6row.energy!(3) = Val(Mid$(astring, 29, 9)) * EVPERKEV#
x6row.energy!(4) = Val(Mid$(astring, 38, 9)) * EVPERKEV#
x6row.energy!(5) = Val(Mid$(astring, 47, 9)) * EVPERKEV#
x6row.energy!(6) = Val(Mid$(astring, 56, 9)) * EVPERKEV#
ElseIf mode% = 1 Then                           ' edge energies
x9row.energy!(1) = Val(Mid$(astring, 11, 9)) * EVPERKEV#
x9row.energy!(2) = Val(Mid$(astring, 20, 9)) * EVPERKEV#
x9row.energy!(3) = Val(Mid$(astring, 29, 9)) * EVPERKEV#
x9row.energy!(4) = Val(Mid$(astring, 38, 9)) * EVPERKEV#
x9row.energy!(5) = Val(Mid$(astring, 47, 9)) * EVPERKEV#
x9row.energy!(6) = Val(Mid$(astring, 56, 9)) * EVPERKEV#
x9row.energy!(7) = Val(Mid$(astring, 65, 9)) * EVPERKEV#
x9row.energy!(8) = Val(Mid$(astring, 74, 9)) * EVPERKEV#
x9row.energy!(9) = Val(Mid$(astring, 83, 9)) * EVPERKEV#
ElseIf mode% = 2 Then                           ' fluorescent yields
x6row.energy!(1) = Val(Mid$(astring, 11, 9))
x6row.energy!(2) = Val(Mid$(astring, 20, 9))
x6row.energy!(3) = Val(Mid$(astring, 29, 9))
x6row.energy!(4) = Val(Mid$(astring, 38, 9))
x6row.energy!(5) = Val(Mid$(astring, 47, 9))
x6row.energy!(6) = Val(Mid$(astring, 56, 9))
End If

If mode% = 0 Then
Put #Temp2FileNumber%, nrec%, x6row
ElseIf mode% = 1 Then
Put #Temp2FileNumber%, nrec%, x9row
ElseIf mode% = 2 Then
Put #Temp2FileNumber%, nrec%, x6row
End If

Next nrec%
        
Close #Temp1FileNumber%
Close #Temp2FileNumber%

' Inform user
msg$ = "File " & tfilename1$ & " converted to " & tfilename2$ & "." & vbCrLf & vbCrLf
msg$ = msg$ & "You may want to rename the new file to the default filename to utilize it in CalcZAF or Probe for EPMA quantitative calculations."
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

EditConvertTextToDATUnknownFormat:
msg$ = "Input/output file format could not be determined. Make sure the filenames contain the string LINE, EDGE or FLUR to identify the data type."
MsgBox msg$, vbOKOnly + vbExclamation, "EditConvertTextToDAT"
ierror = True
Exit Sub

End Sub

Sub EditConvertDATToText(tForm As Form)
' Convert element binary DAT file to text file
'   mode% = 0 emission line energies
'   mode% = 1 edge energies
'   mode% = 2 fluorescent yields

ierror = False
On Error GoTo EditConvertDATToTextError

Dim i As Integer, nrec As Integer, mode As Integer
Dim astring As String, tfilename1 As String, tfilename2 As String

Dim x6row As TypeEnergy
Dim x9row As TypeEdge

' Check for input file
tfilename1$ = ApplicationCommonAppData$ & "XLINE.DAT"
'tfilename1$ = ApplicationCommonAppData$ & "XEDGE.DAT"
'tfilename1$ = ApplicationCommonAppData$ & "XFLUR.DAT"
Call IOGetFileName(Int(2), "DAT", tfilename1$, tForm)
If ierror Then Exit Sub

' Check for output file
tfilename2$ = ApplicationCommonAppData$ & "XRAY_LINE.TXT"
'tfilename2$ = ApplicationCommonAppData$ & "XRAY_EDGE.TXT"
'tfilename2$ = ApplicationCommonAppData$ & "XRAY_FLUR.TXT"
Call IOGetFileName(Int(1), "TXT", tfilename2$, tForm)
If ierror Then Exit Sub

If Dir$(tfilename1$) = vbNullString Then GoTo EditConvertDATToTextNotFound

' Determine mode
mode% = -1
If InStr(UCase$(tfilename1$), "LINE") > 0 Then mode% = 0                               ' emission line energies
If InStr(UCase$(tfilename1$), "EDGE") > 0 Then mode% = 1                               ' edge energies
If InStr(UCase$(tfilename1$), "FLUR") > 0 Then mode% = 2                               ' fluorescent yields
If mode% = -1 Then GoTo EditConvertDATToTextUnknownFormat

' Open binary input DAT file
Open tfilename1$ For Random Access Read As #Temp1FileNumber% Len = XRAY_FILE_RECORD_LENGTH%

' Open x-ray output text file
Open tfilename2$ For Output As #Temp2FileNumber%

' Print text file header (9 spaces for each column)
astring$ = Format$("Element", a90$)
If mode% = 0 Then
For i% = 1 To MAXRAY_OLD%
astring$ = astring$ & Format$(Xraylo$(i%), a90$)
Next i%

ElseIf mode% = 1 Then
For i% = 1 To MAXEDG%
astring$ = astring$ & Format$(Edglo$(i%), a90$)
Next i%

ElseIf mode% = 2 Then
For i% = 1 To MAXRAY_OLD%
astring$ = astring$ & Format$(Xraylo$(i%), a90$)
Next i%
End If
Print #Temp2FileNumber%, astring$

' Get each element record (first two records are just element symbols)
For nrec% = 3 To MAXELM% + 2
If mode% = 0 Then                           ' emission line energies
Get #Temp1FileNumber%, nrec%, x6row
ElseIf mode% = 1 Then                       ' edge energies
Get #Temp1FileNumber%, nrec%, x9row
ElseIf mode% = 2 Then                       ' fluorescent yields
Get #Temp1FileNumber%, nrec%, x6row
End If

' Write element symbol
astring$ = Format$(Symup$(nrec% - 2), a90$)

' Emission lines
If mode% = 0 Then
For i% = 1 To MAXRAY_OLD%
astring$ = astring$ & Format$(MiscAutoFormat$(x6row.energy!(i%) / EVPERKEV#), a90$)
Next i%

' Edge energies
ElseIf mode% = 1 Then
For i% = 1 To MAXEDG%
astring$ = astring$ & Format$(MiscAutoFormat$(x9row.energy!(i%) / EVPERKEV#), a90$)
Next i%

' Fluorescent yields
ElseIf mode% = 2 Then
For i% = 1 To MAXRAY_OLD%
astring$ = astring$ & Format$(MiscAutoFormat$(x6row.energy!(i%)), a90$)
Next i%
End If

' Write to text file
Print #Temp2FileNumber%, astring$
Next nrec%
        
Close #Temp1FileNumber%
Close #Temp2FileNumber%

' Inform user
msg$ = "File " & tfilename1$ & " converted to " & tfilename2$
MsgBox msg$, vbOKOnly + vbInformation, "EditConvertDATToText"

Exit Sub

' Errors
EditConvertDATToTextError:
MsgBox Error$, vbOKOnly + vbCritical, "EditConvertDATToText"
ierror = True
Close #Temp1FileNumber%
Close #Temp2FileNumber%
Exit Sub

EditConvertDATToTextNotFound:
msg$ = "File " & tfilename1$ & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "EditConvertDATToText"
ierror = True
Exit Sub

EditConvertDATToTextUnknownFormat:
msg$ = "Input/output file format could not be determined. Make sure the filenames contain the string LINE, EDGE or FLUR to identify the data type."
MsgBox msg$, vbOKOnly + vbExclamation, "EditConvertDATToText"
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
' mode = 1  XLINE
' mode = 2  XEDGE
' mode = 3  XFLUR

ierror = False
On Error GoTo EditUpdateXFilesError

Dim ix As Integer, iz As Integer, tmax As Integer
Dim line_number As Long
Dim astring As String, tfilename As String
Dim response As Integer

Dim tvalues() As Single

' Dimension temp array depending on mode
If mode% = 1 Or mode% = 2 Then
ReDim tvalues(1 To MAXRAY_OLD%) As Single
Else
ReDim tvalues(1 To MAXEDG%) As Single
End If

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

' Read first line of column headings
Line Input #Temp1FileNumber%, astring
Call IOWriteLog(astring$)

' Loop on element entries
Call IOStatusAuto(vbNullString)
icancelauto = False
line_number& = 0
Do Until EOF(Temp1FileNumber%)
If mode% = 1 Then Input #Temp1FileNumber%, iz%, tvalues!(1), tvalues!(2), tvalues!(3), tvalues!(4), tvalues!(5), tvalues!(6)
If mode% = 2 Then Input #Temp1FileNumber%, iz%, tvalues!(1), tvalues!(2), tvalues!(3), tvalues!(4), tvalues!(5), tvalues!(6), tvalues!(7), tvalues!(8), tvalues!(9)
If mode% = 3 Then Input #Temp1FileNumber%, iz%, tvalues!(1), tvalues!(2), tvalues!(3), tvalues!(4), tvalues!(5), tvalues!(6)

If mode% = 1 Then astring$ = "IZ=" & Format$(iz%) & ", XLINE (Ka, Kb, La, Lb, Ma, Mb) = " & MiscAutoFormat$(tvalues!(1)) & MiscAutoFormat$(tvalues!(2)) & MiscAutoFormat$(tvalues!(3)) & MiscAutoFormat$(tvalues!(4)) & MiscAutoFormat$(tvalues!(5)) & MiscAutoFormat$(tvalues!(6))
If mode% = 2 Then astring$ = "IZ=" & Format$(iz%) & ", XEDGE (K, L-I, L-II, L-III, M-I, M-II, M-III, M-IV, M-V) = " & MiscAutoFormat$(tvalues!(1)) & MiscAutoFormat$(tvalues!(2)) & MiscAutoFormat$(tvalues!(3)) & MiscAutoFormat$(tvalues!(4)) & MiscAutoFormat$(tvalues!(5)) & MiscAutoFormat$(tvalues!(6)) & MiscAutoFormat$(tvalues!(7)) & MiscAutoFormat$(tvalues!(8)) & MiscAutoFormat$(tvalues!(9))
If mode% = 3 Then astring$ = "IZ=" & Format$(iz%) & ", XFLUR (Ka, Kb, La, Lb, Ma, Mb) = " & MiscAutoFormat$(tvalues!(1)) & MiscAutoFormat$(tvalues!(2)) & MiscAutoFormat$(tvalues!(3)) & MiscAutoFormat$(tvalues!(4)) & MiscAutoFormat$(tvalues!(5)) & MiscAutoFormat$(tvalues!(6))
Call IOWriteLog(astring$)
line_number& = line_number& + 1

' Check for valid values
If iz% < 1 Or iz% > MAXELM% Then GoTo EditUpdateXFilesBadEmitter
      
' Set updated data value for XLINE or XFLUR
If mode% = 1 Or mode% = 2 Then
For ix% = 1 To MAXRAY_OLD%
If tvalues!(ix%) > 0# Then
Call EditSetXrayData(iz%, ix%, tvalues!(ix%))
If ierror Then Exit Sub
End If
Next ix%

' Set updated data value for XEDGE
Else
For ix% = 1 To MAXEDG%
If tvalues!(ix%) > 0# Then
Call EditSetXrayData(iz%, ix%, tvalues!(ix%))
If ierror Then Exit Sub
End If
Next ix%
End If

Call IOStatusAuto("Processing line " & Format$(line_number&) & " of file " & tfilename$ & "...")
DoEvents
If icancelauto Then
Call IOStatusAuto(vbNullString)
Close #Temp1FileNumber%
ierror = True
Exit Sub
End If
Loop

Call IOStatusAuto(vbNullString)
Close #Temp1FileNumber%

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
Exit Sub

EditUpdateXFilesBadEmitter:
If mode% = 1 Then msg$ = "Invalid atomic number. Please consult the documentation for the format of an example XLINE.TXT update file and create one using a text editor or exported from Excel as a tab delimited file."
If mode% = 2 Then msg$ = "Invalid atomic number. Please consult the documentation for the format of an example XEDGE.TXT update file and create one using a text editor or exported from Excel as a tab delimited file."
If mode% = 3 Then msg$ = "Invalid atomic number. Please consult the documentation for the format of an example XFLUR.TXT update file and create one using a text editor or exported from Excel as a tab delimited file."
MsgBox msg$, vbOKOnly + vbExclamation, "EditUpdateXFiles"
Call IOStatusAuto(vbNullString)
ierror = True
Close #Temp1FileNumber%
Exit Sub

End Sub

Sub EditUpdateEdgeLineFlurFiles()
' This routine (only called once) updates xray files and add records for elements 95 - 100 (nrec 97-102)

' BE SURE TO COPY UNHASHED XLINE.DAT FROM REGISTER FOLDER
' AFTER STARTING PROGRAM AND THEN RUN UPDATE!!!!!!!!!!!

' From:
' R. B. Firestone, Table of isotopes, 8th Edition,
' Vol II: A+151-272, V.S. Shirley, ed.
' Lawrence Berkeley National Laboratory,
' John Wiley & Sons, Inc, New York

' H. Kleykamp, "Wavelengths of the M X-Rays of Uranium, Plutonium,
' and Americium", Z. Naturforsch., 36a, 1388-1390 (1981)

' H. Kleykamp, "X-ray Emission Wavelengths of Argon, Krypton, Xenon
' and Curium", Z. Naturforsch., 47a, 460-462 (1992).

ierror = False
On Error GoTo EditUpdateEdgeLineFlurFilesError

Dim nrec As Integer, n As Integer, response As Integer

Dim engrow As TypeEnergy
Dim edgrow As TypeEdge
Dim flurow As TypeFlur

msg$ = "Are you sure you want to overwrite the records for elements 95-100 in the x-ray edge (XEDGE.DAT), line (XLINE.DAT) and fluorescent (XFLUR.DAT) data files?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton1, "EditUpdateEdgeLineFlurFiles")
If response% = vbCancel Then Exit Sub

' Open x-ray edge file
Open XEdgeFile$ For Random Access Write As #XEdgeFileNumber% Len = XRAY_FILE_RECORD_LENGTH%

' Open x-ray line file
Open XLineFile$ For Random Access Write As #XLineFileNumber% Len = XRAY_FILE_RECORD_LENGTH%

' Open x-ray flur file
Open XFlurFile$ For Random Access Write As #XFlurFileNumber% Len = XRAY_FILE_RECORD_LENGTH%

' Loop on records
For n% = 95 To 100

If n% = 95 Then
edgrow.energy!(1) = 124.982 * EVPERKEV#  ' K
edgrow.energy!(2) = 23.808 * EVPERKEV#  ' L-I
edgrow.energy!(3) = 22.952 * EVPERKEV#  ' L-II
edgrow.energy!(4) = 18.51 * EVPERKEV#  ' L-III
edgrow.energy!(5) = 6.133 * EVPERKEV#  ' M-I
edgrow.energy!(6) = 5.739 * EVPERKEV#  ' M-II
edgrow.energy!(7) = 4.698 * EVPERKEV#  ' M-III
edgrow.energy!(8) = 4.096 * EVPERKEV#  ' M-IV
edgrow.energy!(9) = 3.89 * EVPERKEV#   ' M-V

engrow.energy!(1) = 106.47 * EVPERKEV#  ' Ka
engrow.energy!(2) = 120.284 * EVPERKEV#  ' Kb
engrow.energy!(3) = 14.62 * EVPERKEV#  ' La
engrow.energy!(4) = 18.856 * EVPERKEV#  ' Lb
engrow.energy!(5) = 3.442 * EVPERKEV#  ' Ma
engrow.energy!(6) = 3.633 * EVPERKEV#  ' Mb

flurow.fraction!(1) = 0.977    ' Ka
flurow.fraction!(2) = 0.977    ' Kb
flurow.fraction!(3) = 0.494    ' La
flurow.fraction!(4) = 0.494    ' Lb
flurow.fraction!(5) = 0#    ' Ma
flurow.fraction!(6) = 0#    ' Mb
End If

If n% = 96 Then
edgrow.energy!(1) = 128.241 * EVPERKEV#  ' K
edgrow.energy!(2) = 24.526 * EVPERKEV#  ' L-I
edgrow.energy!(3) = 23.651 * EVPERKEV#  ' L-II
edgrow.energy!(4) = 18.97 * EVPERKEV#  ' L-III
edgrow.energy!(5) = 6.337 * EVPERKEV#  ' M-I
edgrow.energy!(6) = 5.937 * EVPERKEV#  ' M-II
edgrow.energy!(7) = 4.838 * EVPERKEV#  ' M-III
edgrow.energy!(8) = 4.224 * EVPERKEV#  ' M-IV
edgrow.energy!(9) = 4.009 * EVPERKEV#   ' M-V

engrow.energy!(1) = 109.271 * EVPERKEV#  ' Ka
engrow.energy!(2) = 123.403 * EVPERKEV#  ' Kb
engrow.energy!(3) = 14.961 * EVPERKEV#  ' La
engrow.energy!(4) = 19.427 * EVPERKEV#  ' Lb
engrow.energy!(5) = 3.535 * EVPERKEV#  ' Ma
engrow.energy!(6) = 3.738 * EVPERKEV#  ' Mb

flurow.fraction!(1) = 0.978    ' Ka
flurow.fraction!(2) = 0.978    ' Kb
flurow.fraction!(3) = 0.504    ' La
flurow.fraction!(4) = 0.504    ' Lb
flurow.fraction!(5) = 0#    ' Ma
flurow.fraction!(6) = 0#    ' Mb
End If

If n% = 97 Then
edgrow.energy!(1) = 0#  ' K
edgrow.energy!(2) = 0#  ' L-I
edgrow.energy!(3) = 0#  ' L-II
edgrow.energy!(4) = 0#  ' L-III
edgrow.energy!(5) = 0#  ' M-I
edgrow.energy!(6) = 0#  ' M-II
edgrow.energy!(7) = 0#  ' M-III
edgrow.energy!(8) = 0#  ' M-IV
edgrow.energy!(9) = 0#   ' M-V

engrow.energy!(1) = 0#  ' Ka
engrow.energy!(2) = 0#  ' Kb
engrow.energy!(3) = 0#  ' La
engrow.energy!(4) = 0#  ' Lb
engrow.energy!(5) = 0#  ' Ma
engrow.energy!(6) = 0#  ' Mb

flurow.fraction!(1) = 0#    ' Ka
flurow.fraction!(2) = 0#    ' Kb
flurow.fraction!(3) = 0#    ' La
flurow.fraction!(4) = 0#    ' Lb
flurow.fraction!(5) = 0#    ' Ma
flurow.fraction!(6) = 0#    ' Mb
End If

If n% = 98 Then
End If

If n% = 99 Then
End If

If n% = 100 Then
End If

nrec% = n% + 2
Put #XEdgeFileNumber%, nrec%, edgrow
Put #XLineFileNumber%, nrec%, engrow
Put #XFlurFileNumber%, nrec%, flurow
Next n%

Close #XEdgeFileNumber%
Close #XLineFileNumber%
Close #XFlurFileNumber%

msg$ = "Records updated"
MsgBox msg$, vbOKOnly + vbInformation, "EditUpdateEdgeLineFlurFiles"

Exit Sub

' Errors
EditUpdateEdgeLineFlurFilesError:
MsgBox Error$, vbOKOnly + vbCritical, "EditUpdateEdgeLineFlurFiles"
Close #XEdgeFileNumber%
Close #XLineFileNumber%
Close #XFlurFileNumber%
ierror = True
Exit Sub

End Sub

Sub EditConvertCSVToText(tForm As Form)
' Convert element CSV file to text file for additional x-ray lines (Ln, Lg, Lv, Ll, Mg, Mz) (data from Phil Gopon from DTSA2)
'   mode% = 0 emission line energies
'   mode% = 1 edge energies (not used)
'   mode% = 2 fluorescent yields

ierror = False
On Error GoTo EditConvertCSVToTextError

Dim i As Integer, nrec As Integer, mode As Integer
Dim astring As String, bstring As String, tfilename1 As String, tfilename2 As String

Dim array6(1 To MAXRAY_OLD%) As Single
Dim array9(1 To MAXEDG%) As Single

' Check for new dimensioning
If MAXRAY% <> 13 Then GoTo EditConvertCSVToTextNotNewMAXRAY

' Check for input file
tfilename1$ = "C:\Source\Probewin32-E\Additional Xray Lines\" & "XLINE2.CSV"
'tfilename1$ = "C:\Source\Probewin32-E\Additional Xray Lines\" & "XEDGE2.CSV"         ' not used
'tfilename1$ = "C:\Source\Probewin32-E\Additional Xray Lines\" & "XFLUR2.CSV"
Call IOGetFileName(Int(2), "CSV", tfilename1$, tForm)
If ierror Then Exit Sub

' Check for output file
tfilename2$ = "C:\Source\Probewin32-E\Additional Xray Lines\" & "XRAY_LINE2.TXT"
'tfilename2$ = "C:\Source\Probewin32-E\Additional Xray Lines\"& "XRAY_EDGE2.TXT"       ' not used
'tfilename2$ = "C:\Source\Probewin32-E\Additional Xray Lines\" & "XRAY_FLUR2.TXT"
Call IOGetFileName(Int(1), "TXT", tfilename2$, tForm)
If ierror Then Exit Sub

If Dir$(tfilename1$) = vbNullString Then GoTo EditConvertCSVToTextNotFound

' Determine mode
mode% = -1
If InStr(UCase$(tfilename1$), "LINE") > 0 Then mode% = 0                               ' emission line energies
If InStr(UCase$(tfilename1$), "EDGE") > 0 Then mode% = 1                               ' edge energies (not used)
If InStr(UCase$(tfilename1$), "FLUR") > 0 Then mode% = 2                               ' fluorescent yields
If mode% = -1 Then GoTo EditConvertCSVToTextUnknownFormat

' Open comma delimited input CSV file
Open tfilename1$ For Input As #Temp1FileNumber%

' Open x-ray output text file
Open tfilename2$ For Output As #Temp2FileNumber%

' Print text file header (9 spaces for each column)
astring$ = Format$("Element", a90$)
If mode% = 0 Then
For i% = MAXRAY_OLD% + 1 To MAXRAY% - 1
astring$ = astring$ & Format$(Xraylo$(i%), a90$)
Next i%

ElseIf mode% = 1 Then
For i% = 1 To MAXEDG%
astring$ = astring$ & Format$(Edglo$(i%), a90$)
Next i%

ElseIf mode% = 2 Then
For i% = MAXRAY_OLD% + 1 To MAXRAY% - 1
astring$ = astring$ & Format$(Xraylo$(i%), a90$)
Next i%
End If
Print #Temp2FileNumber%, astring$

' Get each element line (first line is a header)
Line Input #Temp1FileNumber%, astring$
For nrec% = 3 To MAXELM% + 2
Line Input #Temp1FileNumber%, astring$

' Remove atomic number and symbol from string
Call MiscParseStringToStringA(astring$, VbComma, bstring$)
If ierror Then Exit Sub
Call MiscParseStringToStringA(astring$, VbComma, bstring$)
If ierror Then Exit Sub

' Load the values into an array
If mode% = 0 Then                           ' emission line energies
Call InitParseStringToReal(astring$, MAXRAY_OLD%, array6!())
If ierror Then Exit Sub

ElseIf mode% = 1 Then                       ' edge energies
Call InitParseStringToReal(astring$, MAXEDG%, array9!())
If ierror Then Exit Sub

ElseIf mode% = 2 Then                       ' fluorescent yields
Call InitParseStringToReal(astring$, MAXRAY_OLD%, array6!())
If ierror Then Exit Sub
End If

' Write element symbol
astring$ = Format$(Symup$(nrec% - 2), a90$)

' Emission lines in keV
If mode% = 0 Then
For i% = 1 To MAXRAY_OLD%
astring$ = astring$ & Format$(MiscAutoFormat$(array6!(i%)), a90$)
Next i%

' Edge energies in keV (not used)
ElseIf mode% = 1 Then
For i% = 1 To MAXEDG%
astring$ = astring$ & Format$(MiscAutoFormat$(array9!(i%)), a90$)
Next i%

' Fluorescent yields (fraction)
ElseIf mode% = 2 Then
For i% = 1 To MAXRAY_OLD%
astring$ = astring$ & Format$(MiscAutoFormat$(array6!(i%)), a90$)
Next i%
End If

' Write to text file
Print #Temp2FileNumber%, astring$
Next nrec%
        
Close #Temp1FileNumber%
Close #Temp2FileNumber%

' Inform user
msg$ = "File " & tfilename1$ & " converted to " & tfilename2$
MsgBox msg$, vbOKOnly + vbInformation, "EditConvertCSVToText"

Exit Sub

' Errors
EditConvertCSVToTextError:
MsgBox Error$, vbOKOnly + vbCritical, "EditConvertCSVToText"
ierror = True
Close #Temp1FileNumber%
Close #Temp2FileNumber%
Exit Sub

EditConvertCSVToTextNotNewMAXRAY:
msg$ = "The application is not using the new constant for MAXRAY. Check the Global.bas source code and uncomment/comment as necesaary."
MsgBox msg$, vbOKOnly + vbExclamation, "EditConvertCSVToText"
ierror = True
Exit Sub

EditConvertCSVToTextNotFound:
msg$ = "File " & tfilename1$ & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "EditConvertCSVToText"
ierror = True
Exit Sub

EditConvertCSVToTextUnknownFormat:
msg$ = "Input/output file format could not be determined. Make sure the filenames contain the string LINE, EDGE or FLUR to identify the data type."
MsgBox msg$, vbOKOnly + vbExclamation, "EditConvertCSVToText"
ierror = True
Exit Sub

End Sub

Sub EditConvertFFAST2Dat()
' Convert FFAST2*.CSV (Gopon) to FFAST2.DAT for additional x-ray lines (Ln, Lg, Lv, Ll, Mg, Mz)

ierror = False
On Error GoTo EditConvertFFAST2DatError

Dim nrec As Integer, ip As Integer, j As Integer, k As Integer
Dim iemitter As Integer, ixray As Integer, iabsorber As Integer
Dim astring As String, bstring As String
Dim mac_value As Single
Dim exportfilenumber As Integer

Dim macrow As TypeMu

ReDim inputfilenames(1 To MAXRAY_OLD%) As String   ' Ln, Lg, Lv, Ll, Mg, Mz files
ReDim inputfilenumbers(1 To MAXRAY_OLD%) As Integer   ' Ln, Lg, Lv, Ll, Mg, Mz files

' Load full elements names!
Call EditLoadFullElementNames
If ierror Then Exit Sub

' Check for files
inputfilenames$(1) = "C:\Source\Probewin32-E\Additional Xray Lines\" & "FFAST2_Ln.csv"
If Dir$(inputfilenames$(1)) = vbNullString Then GoTo EditConvertFFAST2DatNotFound

inputfilenames$(2) = "C:\Source\Probewin32-E\Additional Xray Lines\" & "FFAST2_Lg.csv"
If Dir$(inputfilenames$(2)) = vbNullString Then GoTo EditConvertFFAST2DatNotFound

inputfilenames$(3) = "C:\Source\Probewin32-E\Additional Xray Lines\" & "FFAST2_Lv.csv"
If Dir$(inputfilenames$(3)) = vbNullString Then GoTo EditConvertFFAST2DatNotFound

inputfilenames$(4) = "C:\Source\Probewin32-E\Additional Xray Lines\" & "FFAST2_Ll.csv"
If Dir$(inputfilenames$(4)) = vbNullString Then GoTo EditConvertFFAST2DatNotFound

inputfilenames$(5) = "C:\Source\Probewin32-E\Additional Xray Lines\" & "FFAST2_Mg.csv"
If Dir$(inputfilenames$(5)) = vbNullString Then GoTo EditConvertFFAST2DatNotFound

inputfilenames$(6) = "C:\Source\Probewin32-E\Additional Xray Lines\" & "FFAST2_Mz.csv"
If Dir$(inputfilenames$(6)) = vbNullString Then GoTo EditConvertFFAST2DatNotFound
      
' Load input file numbers
inputfilenumbers%(1) = FreeFile()
Open inputfilenames$(1) For Input As #inputfilenumbers%(1)

inputfilenumbers%(2) = FreeFile()
Open inputfilenames$(2) For Input As #inputfilenumbers%(2)

inputfilenumbers%(3) = FreeFile()
Open inputfilenames$(3) For Input As #inputfilenumbers%(3)

inputfilenumbers%(4) = FreeFile()
Open inputfilenames$(4) For Input As #inputfilenumbers%(4)

inputfilenumbers%(5) = FreeFile()
Open inputfilenames$(5) For Input As #inputfilenumbers%(5)

inputfilenumbers%(6) = FreeFile()
Open inputfilenames$(6) For Input As #inputfilenumbers%(6)

' Open output file
exportfilenumber% = FreeFile()
Open "C:\Source\Probewin32-E\Additional Xray Lines\" & "FFAST2.DAT" For Random Access Read Write As #exportfilenumber% Len = MAC_FILE_RECORD_LENGTH%

' Write all zeros to file first
For nrec% = 1 To MAXELM%
Put #exportfilenumber%, nrec%, macrow
Next nrec%

' Loop on each input file
For j% = 1 To MAXRAY_OLD%

' Read header info
Line Input #inputfilenumbers%(j%), astring
Line Input #inputfilenumbers%(j%), astring
Line Input #inputfilenumbers%(j%), astring

' Read first element absorber string
Line Input #inputfilenumbers%(j%), astring

' Loop until EOF
Do Until EOF(inputfilenumbers%(j%))

' Determine element absorber Z
Call MiscParseStringToStringA(astring$, VbComma, bstring$)
If ierror Then Exit Sub
ip% = IPOS1%(MAXELM%, bstring$, ElementNames$())
If ip% = 0 Then GoTo EditConvertFFAST2DatBadAbsorber
iabsorber% = ip%

' Read first emitter Z
Line Input #inputfilenumbers%(j%), astring

' Loop while string does not contain two commas (process emitters)
Do While InStr(astring$, ",,") = 0

' Determine element emitter Z
Call MiscParseStringToStringA(astring$, VbSpace, bstring$)
If ierror Then Exit Sub
ip% = IPOS1%(MAXELM%, bstring$, Symlo$())
If ip% = 0 Then GoTo EditConvertFFAST2DatBadEmitter
iemitter% = ip%

' Determine emitter transition (should be constant for each file)
ixray% = j% + MAXRAY_OLD%       ' to be self consistent with ZAFLoadMAC and ZAFLoadMAC2

' Determine mac value
Call MiscParseStringToStringA(astring$, VbComma, bstring$)
If ierror Then Exit Sub

' Sanity check for transition
If j% = 1 Then If bstring$ <> "L2-M1" Then GoTo EditConvertFFAST2DatBadTransition   ' Ln
If j% = 2 Then If bstring$ <> "L2-N4" Then GoTo EditConvertFFAST2DatBadTransition   ' Lg
If j% = 3 Then If bstring$ <> "L2-N6" Then GoTo EditConvertFFAST2DatBadTransition   ' Lv
If j% = 4 Then If bstring$ <> "L3-M1" Then GoTo EditConvertFFAST2DatBadTransition   ' Ll
If j% = 5 Then If bstring$ <> "M3-N5" Then GoTo EditConvertFFAST2DatBadTransition   ' Mg
If j% = 6 Then If bstring$ <> "M5-N3" Then GoTo EditConvertFFAST2DatBadTransition   ' Mz

Call MiscParseStringToStringA(astring$, VbComma, bstring$)
If ierror Then Exit Sub
mac_value! = Val(bstring$)

' Determine emitter record number (1 to MAXELM%)
nrec% = iemitter%
Get #exportfilenumber%, nrec%, macrow

' Calculate record offset and update
k% = (ixray% - MAXRAY_OLD%) + (iabsorber% - 1) * (MAXRAY_OLD%)
macrow.mac!(k%) = mac_value!

' Save edited value
Put #exportfilenumber%, nrec%, macrow

' Read next emitter
Line Input #inputfilenumbers%(j%), astring
Loop

Loop          ' EOF
Next j%       ' next input file

Close #inputfilenumbers%(1)
Close #inputfilenumbers%(2)
Close #inputfilenumbers%(3)
Close #inputfilenumbers%(4)
Close #inputfilenumbers%(5)
Close #inputfilenumbers%(6)

Close #exportfilenumber%

' Inform user
msg$ = "Files :" & vbCrLf
msg$ = msg$ & inputfilenames$(1) & vbCrLf
msg$ = msg$ & inputfilenames$(2) & vbCrLf
msg$ = msg$ & inputfilenames$(3) & vbCrLf
msg$ = msg$ & inputfilenames$(4) & vbCrLf
msg$ = msg$ & inputfilenames$(5) & vbCrLf
msg$ = msg$ & inputfilenames$(6) & vbCrLf
msg$ = msg$ & "converted to " & ApplicationCommonAppData$ & "FFAST2.DAT"
MsgBox msg$, vbOKOnly + vbInformation, "EditConvertFFAST2Dat"

Exit Sub

' Errors
EditConvertFFAST2DatError:
MsgBox Error$, vbOKOnly + vbCritical, "EditConvertFFAST2Dat"
ierror = True
Close #inputfilenumbers%(1)
Close #inputfilenumbers%(2)
Close #inputfilenumbers%(3)
Close #inputfilenumbers%(4)
Close #inputfilenumbers%(5)
Close #inputfilenumbers%(6)
Close #exportfilenumber%
Exit Sub

EditConvertFFAST2DatNotFound:
msg$ = "One of the FFAST2*.csv input files was not found. There should be six of them, e.g. " & inputfilenames$(1) & "."
MsgBox msg$, vbOKOnly + vbExclamation, "EditConvertFFAST2Dat"
ierror = True
Exit Sub

EditConvertFFAST2DatBadAbsorber:
msg$ = "MAC absorber Z could not be determined in string " & bstring$ & ", for input file " & inputfilenames$(j%)
MsgBox msg$, vbOKOnly + vbExclamation, "EditConvertFFAST2Dat"
ierror = True
Exit Sub

EditConvertFFAST2DatBadEmitter:
msg$ = "MAC emitter Z could not be determined in string " & bstring$ & ", for input file " & inputfilenames$(j%)
MsgBox msg$, vbOKOnly + vbExclamation, "EditConvertFFAST2Dat"
ierror = True
Exit Sub

EditConvertFFAST2DatBadTransition:
msg$ = "Invalid x-ray transition in string " & bstring$ & ", for input file " & inputfilenames$(j%)
MsgBox msg$, vbOKOnly + vbExclamation, "EditConvertFFAST2Dat"
ierror = True
Exit Sub


End Sub

