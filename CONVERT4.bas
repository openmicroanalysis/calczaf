Attribute VB_Name = "CodeCONVERT4"
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

Sub ConvertCalculateElectronRange(radius As Single, keV As Single, density As Single, lchan As Integer, syms() As String, wts() As Single)
' Calculate electron Kanaya-Okayama Range (1972)

ierror = False
On Error GoTo ConvertCalculateElectronRangeError

Dim i As Integer, ip As Integer
Dim averageatomicweight As Single
Dim averageatomicnumber As Single

' Calculate average atomic weight
For i% = 1 To lchan%
ip% = IPOS1%(MAXELM%, syms$(i%), Symlo$())
averageatomicweight! = averageatomicweight! + wts!(i%) / 100# * AllAtomicWts!(ip%)
Next i%

' Calculate average atomic number
For i% = 1 To lchan%
ip% = IPOS1%(MAXELM%, syms$(i%), Symlo$())
averageatomicnumber! = averageatomicnumber! + wts!(i%) / 100# * AllAtomicNums%(ip%)
Next i%

' Electron ranges
radius! = (0.0276 * averageatomicweight! * keV! ^ 1.67) / (density! * averageatomicnumber! ^ 0.89)

' Ruste equation gives similar results
'radius! = (0.033 * averageatomicweight! * kev! ^ 1.7) / (density! * averageatomicnumber!)

Exit Sub

ConvertCalculateElectronRangeError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertCalculateElectronRange"
ierror = True
Exit Sub

End Sub

Sub ConvertCalculateElectronEnergy(energy As Single, keV As Single, density As Single, thickness As Single, lchan As Integer, syms() As String, wts() As Single)
' Calculate transmitted electron energy based on Kanaya-Okayama Range (1972)
'   energy (x-ray) in keV
'   density in gm/cm3
'   thickness in microns

ierror = False
On Error GoTo ConvertCalculateElectronEnergyError

Dim i As Integer, ip As Integer
Dim averageatomicweight As Single
Dim averageatomicnumber As Single

' Calculate average atomic weight
For i% = 1 To lchan%
ip% = IPOS1%(MAXELM%, syms$(i%), Symlo$())
averageatomicweight! = averageatomicweight! + wts!(i%) / 100# * AllAtomicWts!(ip%)
Next i%

' Calculate average atomic number
For i% = 1 To lchan%
ip% = IPOS1%(MAXELM%, syms$(i%), Symlo$())
averageatomicnumber! = averageatomicnumber! + wts!(i%) / 100# * AllAtomicNums%(ip%)
Next i%

' Electron energy (final)
energy! = (keV! ^ 1.67 - (density! * thickness! * averageatomicnumber! ^ 0.89) / (0.276 * averageatomicweight!)) ^ (1# / 1.67)
Exit Sub

ConvertCalculateElectronEnergyError:
MsgBox Error$ & ", (make sure specified thickness is not excessive!)", vbOKOnly + vbCritical, "ConvertCalculateElectronEnergy"
ierror = True
Exit Sub

End Sub

Sub ConvertCalculateElectronEnergy2(actualenergy As Single, chan As Integer, sample() As TypeSample)
' Calculate electron energy loss for this sample (standard or unknown coating)

ierror = False
On Error GoTo ConvertCalculateElectronEnergy2Error

Dim averageatomicweight As Single
Dim averageatomicnumber As Single

' Calculate atomic weight
averageatomicweight! = AllAtomicWts!(sample(1).CoatingElement%)

' Calculate atomic number
averageatomicnumber! = AllAtomicNums%(sample(1).CoatingElement%)

' Electron energy loss (calculate angstrom thickness in um for electron energy loss calculation)
actualenergy! = ((sample(1).KilovoltsArray!(chan%)) ^ 1.67 - (sample(1).CoatingDensity! * sample(1).CoatingThickness! / ANGPERMICRON& * averageatomicnumber! ^ 0.89) / (0.276 * averageatomicweight!)) ^ (1# / 1.67)

Exit Sub

' Errors
ConvertCalculateElectronEnergy2Error:
MsgBox Error$ & ", (make sure coating thickness is not excessive!)", vbOKOnly + vbCritical, "ConvertCalculateElectronEnergy2"
ierror = True
Exit Sub

End Sub

Sub ConvertCalculateCoatingElectronAbsorption(ratio As Single, chan As Integer, sample() As TypeSample)
' Calculate intensity change from electron absorption for this sample (standard or unknown coating)
' From Kerrick, et. al., Amer. Min. 58, 920-925 (1973)
' Igen,coated/Igen,uncoated = [100-{(8.3pT)/(eO^2-Ec^2)}]/100 (from JTA)

ierror = False
On Error GoTo ConvertCalculateCoatingElectronAbsorptionError

' Intensity correction from electron energy loss due to coating
ratio! = (100# - ((8.3 * sample(1).CoatingDensity! * sample(1).CoatingThickness! / ANGPERNM&) / (sample(1).KilovoltsArray!(chan%) ^ 2 - (sample(1).LineEdge!(chan%) / EVPERKEV#) ^ 2))) / 100#

' Check for negative correction (coating is too thick)
If ratio! <= 0# Then
msg$ = "ConvertCalculateCoatingElectronAbsorption: Negative electron absorption correction (make sure coating thickness is not excessive!)"
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
ratio! = 1#
End If

Exit Sub

' Errors
ConvertCalculateCoatingElectronAbsorptionError:
MsgBox Error$ & ", (make sure coating thickness is not excessive!)", vbOKOnly + vbCritical, "ConvertCalculateCoatingElectronAbsorption"
ierror = True
Exit Sub

End Sub

Sub ConvertCalculateCoatingXrayTransmission(transmission As Single, chan As Integer, sample() As TypeSample)
' Calculate x-ray absorption loss for this sample coating (standard or unknown coating in angstroms)

ierror = False
On Error GoTo ConvertCalculateCoatingXrayTransmissionError

Dim atotal As Single
Dim sinthickness As Single

' Check energy of emitting x-ray
If sample(1).CoatingElement% = 0 Then GoTo ConvertCalculateCoatingXrayTransmissionBadAbsorber
If sample(1).AtomicNums%(chan%) = 0 Then GoTo ConvertCalculateCoatingXrayTransmissionBadEmitter

' Load table MAC for this absorbing element and the emitting element and x-ray
Call ZAFLoadMac2(sample(1).AtomicNums%(chan%), sample(1).XrayNums%(chan%), sample(1).CoatingElement%, atotal!)
If ierror Then Exit Sub

' Check for bad MAC
If atotal! = 0# Then GoTo ConvertCalculateCoatingXrayTransmissionBadMAC

' Calculate thickness based on takeoff angle
sinthickness! = MathCalculateSinThickness!(sample(1).CoatingThickness!, sample(1).takeoff!)

' Calculate x-ray transmission based on Sin thickness in angstroms
transmission! = NATURALE# ^ (-1# * atotal! * sample(1).CoatingDensity! * sinthickness! * CMPERANGSTROM#)

If VerboseMode Then
Call IOWriteLog(sample(1).Elsyup$(chan%) & " " & sample(1).Xrsyms$(chan%) & ", x-ray energy of " & Format$(sample(1).LineEnergy!(chan%) / EVPERKEV#) & " keV, absorbed by " & Symlo$(sample(1).CoatingElement%) & ", MAC: " & Format$(atotal!) & ", Thick: " & Format$(sample(1).CoatingThickness!) & ", Sin(Thick): " & Format$(sinthickness!) & ", Trans: " & Format$(transmission!))
End If

Exit Sub

' Errors
ConvertCalculateCoatingXrayTransmissionError:
MsgBox Error$ & ", (make sure coating thickness is not excessive!)", vbOKOnly + vbCritical, "ConvertCalculateCoatingXrayTransmission"
ierror = True
Exit Sub

ConvertCalculateCoatingXrayTransmissionBadAbsorber:
msg$ = "Invalid atomic number for coating element"
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertCalculateCoatingXrayTransmission"
ierror = True
Exit Sub

ConvertCalculateCoatingXrayTransmissionBadEmitter:
msg$ = "Invalid atomic number for emitting element"
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertCalculateCoatingXrayTransmission"
ierror = True
Exit Sub

ConvertCalculateCoatingXrayTransmissionBadMAC:
msg$ = "Invalid mass absorption coefficient for coating absorption calculation"
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertCalculateCoatingXrayTransmission"
ierror = True
Exit Sub

End Sub

Sub ConvertCalculateXrayRange(radius As Single, keV As Single, kec As Single, density As Single, syme As String, symx As String, lchan As Integer, syms() As String, wts() As Single)
' Calculate x-ray Kanaya-Okayama Range (in microns for the specified x-ray line)

ierror = False
On Error GoTo ConvertCalculateXrayRangeError

Dim i As Integer, ip As Integer
Dim averageatomicweight As Single
Dim averageatomicnumber As Single

Dim nrec As Integer, i2 As Integer

Dim edgrow As TypeEdge

' Calculate average atomic weight and number
For i% = 1 To lchan%
ip% = IPOS1%(MAXELM%, syms$(i%), Symlo$())
averageatomicweight! = averageatomicweight! + wts!(i%) / 100# * AllAtomicWts!(ip%)
averageatomicnumber! = averageatomicnumber! + wts!(i%) / 100# * AllAtomicNums%(ip%)
Next i%

' Get critical excitation energy for specified element and x-ray
ip% = IPOS1%(MAXELM%, syme$, Symlo$())
If ip% = 0 Then GoTo ConvertCalculateXrayRangeBadSyme
nrec% = AllAtomicNums%(ip%) + 2

' Read from file
Open XEdgeFile$ For Random Access Read As #XEdgeFileNumber% Len = XRAY_FILE_RECORD_LENGTH%
Get #XEdgeFileNumber%, nrec%, edgrow
Close #XEdgeFileNumber%

ip% = IPOS1%(MAXRAY% - 1, symx$, Xraylo$())
If ip% = 0 Then GoTo ConvertCalculateXrayRangeBadSymx

' Calculate edge for each line
If ip% = 1 Then i2% = 1   ' Ka
If ip% = 2 Then i2% = 1   ' Kb
If ip% = 3 Then i2% = 4   ' La
If ip% = 4 Then i2% = 3   ' Lb
If ip% = 5 Then i2% = 9   ' Ma
If ip% = 6 Then i2% = 8   ' Mb

kec! = edgrow.energy!(i2%) / EVPERKEV#

' Electron ranges
radius! = (0.0276 * averageatomicweight! * (keV! ^ 1.67 - kec! ^ 1.67)) / (density! * averageatomicnumber! ^ 0.89)

' Ruste equation gives similar results
'radius! = (0.033 * averageatomicweight! * (kev! ^ 1.7 - kec! ^ 1.7)) / (density! * averageatomicnumber!)

Exit Sub

ConvertCalculateXrayRangeError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertCalculateXrayRange"
Close #XEdgeFileNumber%
ierror = True
Exit Sub

ConvertCalculateXrayRangeBadSyme:
msg$ = "Invalid element symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertCalculateXrayRange"
ierror = True
Exit Sub

ConvertCalculateXrayRangeBadSymx:
msg$ = "Invalid xray symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertCalculateXrayRange"
ierror = True
Exit Sub

End Sub

Sub ConvertCalculateXrayTransmission(transmission As Single, averagemassabsorption As Single, density As Single, thickness As Single, syme As String, symx As String, lchan As Integer, syms() As String, wts() As Single)
' Calculate the x-ray transmission for the specified element and x-ray
' thickness in microns

ierror = False
On Error GoTo ConvertCalculateXrayTransmissionError

Dim i As Integer, ip As Integer, ipp As Integer, iz As Integer, num As Integer

Dim macrow As TypeMu

' Calculate emitter x-ray position
ip% = IPOS1%(MAXELM%, syme$, Symlo$())
If ip% = 0 Then GoTo ConvertCalculateXrayTransmissionBadSyme

ipp% = IPOS1%(MAXRAY% - 1, symx$, Xraylo$())
If ipp% = 0 Then GoTo ConvertCalculateXrayTransmissionBadSymx

' Load MAC filename
If ip% <= MAXRAY_OLD% Then
MACFile$ = ApplicationCommonAppData$ & macstring2$(MACTypeFlag%) & ".DAT"
If Dir$(MACFile$) = vbNullString Then GoTo ConvertCalculateXrayTransmissionNotFound
Open MACFile$ For Random Access Read As #MACFileNumber% Len = MAC_FILE_RECORD_LENGTH%
Get #MACFileNumber%, ip%, macrow
Close #MACFileNumber%

' Load MAC filename for additional x-rays
Else
MACFile$ = ApplicationCommonAppData$ & macstring2$(MACTypeFlag%) & "2.DAT"
If Dir$(MACFile$) = vbNullString Then GoTo ConvertCalculateXrayTransmissionNotFound
Open MACFile$ For Random Access Read As #MACFileNumber% Len = MAC_FILE_RECORD_LENGTH%
Get #MACFileNumber%, ip%, macrow
Close #MACFileNumber%
End If

' Calculate average mass absorption for this emitter
For i% = 1 To lchan%

' Calculate absorber position
iz% = IPOS1%(MAXELM%, syms$(i%), Symlo$())
If iz% = 0 Then GoTo ConvertCalculateXrayTransmissionBadSym
num% = ipp% + (iz% - 1) * (MAXRAY% - 1)

If DebugMode Then
Call IOWriteLog(Symlo$(ip%) & " " & Xraylo$(ipp%) & " absorbed by " & Symlo$(iz%) & " = " & Str$(macrow.mac!(num%)))
End If

averagemassabsorption! = averagemassabsorption! + wts!(i%) / 100# * macrow.mac!(num%)
Next i%

transmission! = NATURALE# ^ (-1# * averagemassabsorption! * density! * thickness! * CMPERMICRON#)
Exit Sub

' Errors
ConvertCalculateXrayTransmissionError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertCalculateXrayTransmission"
Close #MACFileNumber%
ierror = True
Exit Sub

ConvertCalculateXrayTransmissionBadSyme:
msg$ = "Invalid element symbol for emitter"
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertCalculateXrayTransmission"
ierror = True
Exit Sub

ConvertCalculateXrayTransmissionBadSymx:
msg$ = "Invalid xray symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertCalculateXrayTransmission"
ierror = True
Exit Sub

ConvertCalculateXrayTransmissionBadSym:
msg$ = "Invalid element symbol for absorber"
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertCalculateXrayTransmission"
ierror = True
Exit Sub

ConvertCalculateXrayTransmissionNotFound:
msg$ = "File " & MACFile$ & " was not found, please choose another MAC file or create the missing file using the CalcZAF Xray menu items"
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertCalculateXrayTransmission"
ierror = True
Exit Sub

End Sub

Sub ConvertCalculateXrayTransmission2(transmission As Single, averagemassabsorption As Single, density As Single, thickness As Single, energy As Single, lchan As Integer, syms() As String, wts() As Single)
' Calculate the x-ray transmission for an arbitrary x-ray energy in keV (thickness in microns)

ierror = False
On Error GoTo ConvertCalculateXrayTransmission2Error

Dim i As Integer, iz As Integer
Dim aphoto As Single, aelastic As Single, ainelastic As Single, atotal As Single

' Check energy
If energy! < 1# Then GoTo ConvertCalculateXrayTransmission2TooLow

' Calculate average mass absorption for this energy
For i% = 1 To lchan%

' Calculate MAC for this energy (use McMaster)
iz% = IPOS1%(MAXELM%, syms$(i%), Symlo$())
If iz% = 0 Then GoTo ConvertCalculateXrayTransmission2BadSym

Call AbsorbGetMAC(iz%, energy!, aphoto!, aelastic!, ainelastic!, atotal!)
If ierror Then Exit Sub

If DebugMode Then
Call IOWriteLog("X-ray energy of " & Str$(energy!) & " keV, absorbed by " & Symlo$(iz%) & ", MAC = " & Str$(atotal!))
End If

averagemassabsorption! = averagemassabsorption! + wts!(i%) / 100# * atotal!
Next i%

transmission! = NATURALE# ^ (-1# * averagemassabsorption! * density! * thickness! * CMPERMICRON#)
Exit Sub

' Errors
ConvertCalculateXrayTransmission2Error:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertCalculateXrayTransmission2"
ierror = True
Exit Sub

ConvertCalculateXrayTransmission2BadSym:
msg$ = "Invalid element symbol for absorber"
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertCalculateXrayTransmission2"
ierror = True
Exit Sub

ConvertCalculateXrayTransmission2TooLow:
msg$ = "Cannot calculate arbitrary energies below 1 keV"
MsgBox msg$, vbOKOnly + vbExclamation, "ConvertCalculateXrayTransmission2"
ierror = True
Exit Sub

End Sub
