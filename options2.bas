Attribute VB_Name = "CodeOPTIONS2"
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

Sub OptionsCheckMineral(Index As Integer, sample() As TypeSample)
' Check for minimum elements for mineral calculation

ierror = False
On Error GoTo OptionsCheckMineralError

Dim ip As Integer
Dim sym As String

' None
If Index% = 0 Then Exit Sub

' Olivine (skip disabled quant elements)
If Index% = 1 Then
sym$ = Symup$(12)   ' Mg
ip% = IPOS1DQ(sample(1).LastChan%, sym$, sample(1).Elsyms$(), sample(1).DisableQuantFlag%())
If ip% = 0 Then GoTo OptionCheckMineralMissingElement
If ip% > sample(1).LastElm% And sample(1).CrystalNames$(ip%) <> EDS_CRYSTAL$ Then
msg$ = "Warning in OptionsCheckMineral: " & sym$ & " is not an analyzed element in this sample..."
Call IOWriteLog(msg$)
End If

sym$ = Symup$(26)   ' Fe
ip% = IPOS1DQ(sample(1).LastChan%, sym$, sample(1).Elsyms$(), sample(1).DisableQuantFlag%())
If ip% = 0 Then GoTo OptionCheckMineralMissingElement
If ip% > sample(1).LastElm% And sample(1).CrystalNames$(ip%) <> EDS_CRYSTAL$ Then
msg$ = "Warning in OptionsCheckMineral: " & sym$ & " is not an analyzed element in this sample..."
Call IOWriteLog(msg$)
End If
End If

' Feldspar (skip disabled quant elements)
If Index% = 2 Then
sym$ = Symup$(11)   ' Na
ip% = IPOS1DQ(sample(1).LastChan%, sym$, sample(1).Elsyms$(), sample(1).DisableQuantFlag%())
If ip% = 0 Then GoTo OptionCheckMineralMissingElement
If ip% > sample(1).LastElm% And sample(1).CrystalNames$(ip%) <> EDS_CRYSTAL$ Then
msg$ = "Warning in OptionsCheckMineral: " & sym$ & " is not an analyzed element in this sample..."
Call IOWriteLog(msg$)
End If

sym$ = Symup$(20)   ' Ca
ip% = IPOS1DQ(sample(1).LastChan%, sym$, sample(1).Elsyms$(), sample(1).DisableQuantFlag%())
If ip% = 0 Then GoTo OptionCheckMineralMissingElement
If ip% > sample(1).LastElm% And sample(1).CrystalNames$(ip%) <> EDS_CRYSTAL$ Then
msg$ = "Warning in OptionsCheckMineral: " & sym$ & " is not an analyzed element in this sample..."
Call IOWriteLog(msg$)
End If

sym$ = Symup$(19)   ' K
ip% = IPOS1DQ(sample(1).LastChan%, sym$, sample(1).Elsyms$(), sample(1).DisableQuantFlag%())
If ip% = 0 Then GoTo OptionCheckMineralMissingElement
If ip% > sample(1).LastElm% And sample(1).CrystalNames$(ip%) <> EDS_CRYSTAL$ Then
msg$ = "Warning in OptionsCheckMineral: " & sym$ & " is not an analyzed element in this sample..."
Call IOWriteLog(msg$)
End If
End If

' Pyroxene (skip disabled quant elements)
If Index% = 3 Then
sym$ = Symup$(20)   ' Ca
ip% = IPOS1DQ(sample(1).LastChan%, sym$, sample(1).Elsyms$(), sample(1).DisableQuantFlag%())
If ip% = 0 Then GoTo OptionCheckMineralMissingElement
If ip% > sample(1).LastElm% And sample(1).CrystalNames$(ip%) <> EDS_CRYSTAL$ Then
msg$ = "Warning in OptionsCheckMineral: " & sym$ & " is not an analyzed element in this sample..."
Call IOWriteLog(msg$)
End If

sym$ = Symup$(12)   ' Mg
ip% = IPOS1DQ(sample(1).LastChan%, sym$, sample(1).Elsyms$(), sample(1).DisableQuantFlag%())
If ip% = 0 Then GoTo OptionCheckMineralMissingElement
If ip% > sample(1).LastElm% And sample(1).CrystalNames$(ip%) <> EDS_CRYSTAL$ Then
msg$ = "Warning in OptionsCheckMineral: " & sym$ & " is not an analyzed element in this sample..."
Call IOWriteLog(msg$)
End If

sym$ = Symup$(26)   ' Fe
ip% = IPOS1DQ(sample(1).LastChan%, sym$, sample(1).Elsyms$(), sample(1).DisableQuantFlag%())
If ip% = 0 Then GoTo OptionCheckMineralMissingElement
If ip% > sample(1).LastElm% And sample(1).CrystalNames$(ip%) <> EDS_CRYSTAL$ Then
msg$ = "Warning in OptionsCheckMineral: " & sym$ & " is not an analyzed element in this sample..."
Call IOWriteLog(msg$)
End If
End If

' Garnet (skip disabled quant elements)
If Index% = 4 Then
sym$ = Symup$(20)   ' Ca
ip% = IPOS1DQ(sample(1).LastChan%, sym$, sample(1).Elsyms$(), sample(1).DisableQuantFlag%())
If ip% = 0 Then GoTo OptionCheckMineralMissingElement
If ip% > sample(1).LastElm% And sample(1).CrystalNames$(ip%) <> EDS_CRYSTAL$ Then
msg$ = "Warning in OptionsCheckMineral: " & sym$ & " is not an analyzed element in this sample..."
Call IOWriteLog(msg$)
End If

sym$ = Symup$(12)   ' Mg
ip% = IPOS1DQ(sample(1).LastChan%, sym$, sample(1).Elsyms$(), sample(1).DisableQuantFlag%())
If ip% = 0 Then GoTo OptionCheckMineralMissingElement
If ip% > sample(1).LastElm% And sample(1).CrystalNames$(ip%) <> EDS_CRYSTAL$ Then
msg$ = "Warning in OptionsCheckMineral: " & sym$ & " is not an analyzed element in this sample..."
Call IOWriteLog(msg$)
End If

sym$ = Symup$(26)   ' Fe
ip% = IPOS1DQ(sample(1).LastChan%, sym$, sample(1).Elsyms$(), sample(1).DisableQuantFlag%())
If ip% = 0 Then GoTo OptionCheckMineralMissingElement
If ip% > sample(1).LastElm% And sample(1).CrystalNames$(ip%) <> EDS_CRYSTAL$ Then
msg$ = "Warning in OptionsCheckMineral: " & sym$ & " is not an analyzed element in this sample..."
Call IOWriteLog(msg$)
End If

sym$ = Symup$(25)   ' Mn
ip% = IPOS1DQ(sample(1).LastChan%, sym$, sample(1).Elsyms$(), sample(1).DisableQuantFlag%())
If ip% = 0 Then GoTo OptionCheckMineralMissingElement
If ip% > sample(1).LastElm% And sample(1).CrystalNames$(ip%) <> EDS_CRYSTAL$ Then
msg$ = "Warning in OptionsCheckMineral: " & sym$ & " is not an analyzed element in this sample..."
Call IOWriteLog(msg$)
End If
End If

' Garnet (Al, Fe, Cr) (skip disabled quant elements)
If Index% = 5 Then
sym$ = Symup$(13)   ' Al
ip% = IPOS1DQ(sample(1).LastChan%, sym$, sample(1).Elsyms$(), sample(1).DisableQuantFlag%())
If ip% = 0 Then GoTo OptionCheckMineralMissingElement
If ip% > sample(1).LastElm% And sample(1).CrystalNames$(ip%) <> EDS_CRYSTAL$ Then
msg$ = "Warning in OptionsCheckMineral: " & sym$ & " is not an analyzed element in this sample..."
Call IOWriteLog(msg$)
End If

sym$ = Symup$(26)   ' Fe
ip% = IPOS1DQ(sample(1).LastChan%, sym$, sample(1).Elsyms$(), sample(1).DisableQuantFlag%())
If ip% = 0 Then GoTo OptionCheckMineralMissingElement
If ip% > sample(1).LastElm% And sample(1).CrystalNames$(ip%) <> EDS_CRYSTAL$ Then
msg$ = "Warning in OptionsCheckMineral: " & sym$ & " is not an analyzed element in this sample..."
Call IOWriteLog(msg$)
End If

sym$ = Symup$(24)   ' Cr
ip% = IPOS1DQ(sample(1).LastChan%, sym$, sample(1).Elsyms$(), sample(1).DisableQuantFlag%())
If ip% = 0 Then GoTo OptionCheckMineralMissingElement
If ip% > sample(1).LastElm% And sample(1).CrystalNames$(ip%) <> EDS_CRYSTAL$ Then
msg$ = "Warning in OptionsCheckMineral: " & sym$ & " is not an analyzed element in this sample..."
Call IOWriteLog(msg$)
End If
End If

Exit Sub

' Errors
OptionsCheckMineralError:
MsgBox Error$, vbOKOnly + vbCritical, "OptionsCheckMineral"
ierror = True
Exit Sub

OptionCheckMineralMissingElement:
msg$ = "The specified mineral end-member cannot be calculated since " & sym$ & " is not an analyzed or specified element in this sample." & vbCrLf & vbCrLf
msg$ = msg$ & "If the missing element is not critical for your mineral end member calculations (e.g. Mn or Cr in some garnets), you can add it as a zero concentration specified element (no x-ray line) from the Elements/Cations dialog."
MsgBox msg$, vbOKOnly + vbExclamation, "OptionsCheckMineral"
ierror = True
Exit Sub

End Sub

