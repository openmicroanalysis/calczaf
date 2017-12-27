Attribute VB_Name = "CodeABSORB"
' (c) Copyright 1995-2018 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Private Const MAXNOFIT% = 13    ' number of elements with no fit data

Dim conv(1 To MAXELM%) As Single ' only MAXELM% element entries in ABSORB.DAT
Dim eed(1 To MAXELM%, 1 To 9) As Single
Dim a0n(1 To MAXELM%) As Single, a1n(1 To MAXELM%) As Single, a0m(1 To MAXELM%) As Single
Dim a1m(1 To MAXELM%) As Single, a0l(1 To MAXELM%) As Single, a1l(1 To MAXELM%) As Single
Dim a2l(1 To MAXELM%) As Single, a0k(1 To MAXELM%) As Single, a1k(1 To MAXELM%) As Single
Dim a2k(1 To MAXELM%) As Single, a3k(1 To MAXELM%) As Single, a0c(1 To MAXELM%) As Single
Dim a1c(1 To MAXELM%) As Single, a2c(1 To MAXELM%) As Single, a3c(1 To MAXELM%) As Single
Dim a0i(1 To MAXELM%) As Single, a1i(1 To MAXELM%) As Single, a2i(1 To MAXELM%) As Single
Dim a3i(1 To MAXELM%) As Single

Dim nofit(1 To MAXNOFIT%) As Integer

Sub AbsorbGetMAC(iz As Integer, energy As Single, aphoto As Single, aelastic As Single, ainelastic As Single, atotal As Single)
' This routine returns the photoelectric, elastic, inelastic and total
' scattering cross sections for element IZ at energy ENERGY (keV).
' Written by Mark Rivers at Brookhaven Nat'l Labs (1988).
' Modified by John Donovan for VB (1995).
' W. H. McMaster, N. Kerr Del Grande, J. H. Mallet and J. H. Hubbell,
' "Compilation of x-ray cross sections ", Lawrence Livermore Lab., 1969."
'   iz is absorber atomic number
'   energy is x-ray energy in keV
'   aphoto is photo absorption cross section
'   aelastic is elastic scattering absorption cross section
'   ainelastic is the inelastic scattering absorption cross section

ierror = False
On Error GoTo AbsorbGetMACError

Dim kk As Integer
Dim l As Integer

Dim ejl1 As Single, ejl2 As Single, ejm10 As Single, ejm11 As Single, ejm20 As Single, ejm21 As Single
Dim ejm30 As Single, ejm31 As Single, ejm40 As Single, ejm41 As Single
Dim wl As Single, aml As Single, am As Single, ac As Single, ai As Single
Dim ip As Integer

Static dataloaded As Integer

aphoto! = 0#
aelastic! = 0#
ainelastic! = 0#
atotal! = 0#

' Check for low energy
If energy! < 1# Then
If VerboseMode Then
msg$ = "WARNING in AbsorbGetMAC- energy " & Format$(Format$(energy!, f83$), a80$) & " is too low for " & MiscAutoUcase$(Symlo$(iz%)) & " absorber MAC calculation."
Call IOWriteLog(msg$)
End If
Exit Sub
End If

' First time through, read in data from ABSORB.DAT
If dataloaded = False Then

' Load elements missing from fit parameters
nofit(1) = 84
nofit(2) = 85
nofit(3) = 87
nofit(4) = 88
nofit(5) = 89
nofit(6) = 91
nofit(7) = 93
nofit(8) = 95
nofit(9) = 96
nofit(10) = 97
nofit(11) = 98
nofit(12) = 99
nofit(13) = 100

' Read data from file
Open AbsorbFile$ For Input As #AbsorbFileNumber%

For l% = 1 To MAXELM%
Input #AbsorbFileNumber%, eed!(l%, 1), eed!(l%, 2), eed!(l%, 3), eed!(l%, 4), eed!(l%, 5), eed!(l%, 6), eed!(l%, 7), eed!(l%, 8), eed!(l%, 9)
Next l%

For l% = 1 To MAXELM%
Input #AbsorbFileNumber%, conv!(l%), a0n!(l%), a1n!(l%), a0m!(l%), a1m!(l%)
Next l%

For l% = 1 To MAXELM%
Input #AbsorbFileNumber%, a0l!(l%), a1l!(l%), a2l!(l%), a0k!(l%), a1k!(l%)
Next l%

For l% = 1 To MAXELM%
Input #AbsorbFileNumber%, a2k!(l%), a3k!(l%), a0c!(l%), a1c!(l%), a2c!(l%)
Next l%

For l% = 1 To MAXELM%
Input #AbsorbFileNumber%, a3c!(l%), a0i!(l%), a1i!(l%), a2i!(l%), a3i!(l%)
Next l%

Close #AbsorbFileNumber%
dataloaded = True
End If

' Check for invalid absorbers
ip% = IPOS2(Int(MAXNOFIT%), iz%, nofit())
If ip% <> 0 Then
If DebugMode Then
msg$ = "WARNING in AbsorbGetMAC- unable to fit " & Format$(Symup$(Int(iz%)), a20$) & " as absorber. MAC for emitting energy (" & Format$(energy!) & " keV), will be set to an arbitrary value of 1000 cm^2/ug."
Call IOWriteLog(msg$)
End If
Exit Sub
End If

' FOLLOWING ARE JUMP FACTORS FOR L & M SUBSHELLS
ejl1! = 1.16
ejl2! = 1.14
ejm10! = 1.0393
ejm11! = 0.00047132
ejm20! = 1.0711
ejm21! = 0.0017851
ejm30! = 1.3809
ejm31! = 0.003106
ejm40! = 2.343
ejm41! = -0.0009287

' Code taken from Barry Gordon's ABSTOT program. Needs to be cleaned up.
      l% = iz%

' DETERMINE SUBSHELL BEING EXCITED
      kk% = 1
529:  If energy! >= eed!(l%, kk%) Then GoTo 530
      kk% = kk% + 1
      If kk% > 9 Then GoTo 530
      GoTo 529
530:  wl! = Log(energy!)
      If kk% = 1 Then GoTo 531
      If kk% = 2 Then GoTo 532
      If kk% = 3 Then GoTo 532
      If kk% = 4 Then GoTo 532
      If kk% = 5 Then GoTo 535
      If kk% = 6 Then GoTo 536
      If kk% = 7 Then GoTo 537
      If kk% = 8 Then GoTo 538
      If kk% = 9 Then GoTo 539
      If kk% = 10 Then GoTo 540
531:  aml! = a0k!(l%) + wl! * a1k!(l%) + wl! ^ 2 * a2k!(l%) + wl! ^ 3 * a3k!(l%)
      am! = Exp(aml!)
      GoTo 541
532:  aml! = a0l!(l%) + wl! * a1l!(l%) + wl! ^ 2 * a2l!(l%)
      am! = Exp(aml!)
      If kk% = 2 Then GoTo 541
      am! = am! / ejl1!
      If kk% = 3 Then GoTo 541
      am! = am! / ejl2!
      GoTo 541
535:  aml! = a0m!(l%) + wl! * a1m!(l%)
      am! = Exp(aml!)
      GoTo 541
536:  aml! = a0m!(l%) + wl! * a1m!(l%)
      am! = Exp(aml!) / (ejm10! + ejm11! * l%)
      GoTo 541
537:  aml! = a0m!(l%) + wl! * a1m!(l%)
      am! = Exp(aml!) / (ejm20! + ejm21! * l%)
      GoTo 541
538:  aml! = a0m!(l%) + wl! * a1m!(l%)
      am! = Exp(aml!) / (ejm30! + ejm31! * l%)
      GoTo 541
539:  aml! = a0m!(l%) + wl! * a1m!(l%)
      am! = Exp(aml!) / (ejm40! + ejm41! * l%)
      GoTo 541
540:  aml! = a0n!(l%) + wl! * a1n!(l%)
      am! = Exp(aml!)
541:  ac! = a0c!(l%) + wl! * a1c!(l%) + wl! ^ 2 * a2c!(l%) + wl! ^ 3 * a3c!(l%)
      ai! = a0i!(l%) + wl! * a1i!(l%) + wl! ^ 2 * a2i!(l%) + wl! ^ 3 * a3i!(l%)
      ac! = Exp(ac!)
      ai! = Exp(ai!)
      aphoto! = am! / conv!(l%)
      aelastic! = ac! / conv!(l%)
      ainelastic! = ai! / conv!(l%)
      atotal! = aphoto! + aelastic! + ainelastic!
Exit Sub

' Errors
AbsorbGetMACError:
MsgBox Error$ & ", calculating MAC for x-ray energy of " & Format$(energy!) & " keV, absorbed by element " & Symup$(iz%), vbOKOnly + vbCritical, "AbsorbGetMAC"
ierror = True
Exit Sub

End Sub
