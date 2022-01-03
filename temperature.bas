Attribute VB_Name = "CodeTemperature"
' (c) Copyright 1995-2022 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Dim TemperatureKilovolts As Single
Dim TemperatureBeamCurrent As Single
Dim TemperatureBeamSize As Single
Dim TemperatureThermalConductivity As Single

Sub TemperatureLoad()
' Load Temperature calculation dialog

ierror = False
On Error GoTo TemperatureLoadError

Dim i As Integer

' Load density and kev
If TemperatureKilovolts! = 0# Then TemperatureKilovolts! = DefaultKiloVolts!
If TemperatureBeamCurrent! = 0# Then TemperatureBeamCurrent! = DefaultBeamCurrent!
If TemperatureBeamSize! = 0# Then TemperatureBeamSize! = DefaultBeamSize!
If TemperatureBeamSize! = 0# Then TemperatureBeamSize! = 1#

FormTEMPERATURE.TextBeamEnergy.Text = Str$(TemperatureKilovolts!)
FormTEMPERATURE.TextBeamCurrent.Text = Str$(TemperatureBeamCurrent!)
FormTEMPERATURE.TextBeamSize.Text = Str$(TemperatureBeamSize!)
FormTEMPERATURE.TextThermalConductivity.Text = Str$(TemperatureThermalConductivity!)    ' load zero if first time

' Load thermal conductivities in milliwatts per K (need to use long integer for ItemData array)
FormTEMPERATURE.ComboThermalConductivity.Clear

FormTEMPERATURE.ComboThermalConductivity.AddItem "Aluminum = 2.5"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 2.5 * MILLIWATTSPERWATT&
FormTEMPERATURE.ComboThermalConductivity.AddItem "Aluminum Oxide = 0.3"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 0.3 * MILLIWATTSPERWATT&
FormTEMPERATURE.ComboThermalConductivity.AddItem "Calcite = 0.05"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 0.05 * MILLIWATTSPERWATT&

FormTEMPERATURE.ComboThermalConductivity.AddItem "Carbon, Amorphous = 0.016"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 0.016 * MILLIWATTSPERWATT&
FormTEMPERATURE.ComboThermalConductivity.AddItem "Carbon, Graphite (Parallel) = 19.6"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 19.6 * MILLIWATTSPERWATT&
FormTEMPERATURE.ComboThermalConductivity.AddItem "Carbon, Graphite (Perpendicular) = 0.057"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 0.057 * MILLIWATTSPERWATT&
FormTEMPERATURE.ComboThermalConductivity.AddItem "Carbon, Diamond = 23.2"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 23.2 * MILLIWATTSPERWATT&

FormTEMPERATURE.ComboThermalConductivity.AddItem "Cement, Mortar = 0.017"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 0.017 * MILLIWATTSPERWATT&
FormTEMPERATURE.ComboThermalConductivity.AddItem "Cement, Portland = 0.003"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 0.003 * MILLIWATTSPERWATT&
FormTEMPERATURE.ComboThermalConductivity.AddItem "Copper = 4.01"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 4.01 * MILLIWATTSPERWATT&
FormTEMPERATURE.ComboThermalConductivity.AddItem "Epoxy = 0.002"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 0.002 * MILLIWATTSPERWATT&
FormTEMPERATURE.ComboThermalConductivity.AddItem "Fiberglass = 0.001"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 0.001 * MILLIWATTSPERWATT&

FormTEMPERATURE.ComboThermalConductivity.AddItem "Glass, Pyrex = 0.01"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 0.01 * MILLIWATTSPERWATT&
FormTEMPERATURE.ComboThermalConductivity.AddItem "Glass, Window = 0.009"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 0.009 * MILLIWATTSPERWATT&
FormTEMPERATURE.ComboThermalConductivity.AddItem "Gold = 3.1"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 3.1 * MILLIWATTSPERWATT&
FormTEMPERATURE.ComboThermalConductivity.AddItem "Iron Metal = 0.80"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 0.8 * MILLIWATTSPERWATT&
FormTEMPERATURE.ComboThermalConductivity.AddItem "Mica = 0.006"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 0.006 * MILLIWATTSPERWATT&
FormTEMPERATURE.ComboThermalConductivity.AddItem "Obsidian Glass = 0.014"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 0.014 * MILLIWATTSPERWATT&

FormTEMPERATURE.ComboThermalConductivity.AddItem "Polypropylene = 0.002"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 0.002 * MILLIWATTSPERWATT&
FormTEMPERATURE.ComboThermalConductivity.AddItem "Quartz = 0.10"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 0.1 * MILLIWATTSPERWATT&
FormTEMPERATURE.ComboThermalConductivity.AddItem "Steel, Carbon (1%) = 0.32"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 0.32 * MILLIWATTSPERWATT&
FormTEMPERATURE.ComboThermalConductivity.AddItem "Steel, Stainless = 0.16"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 0.16 * MILLIWATTSPERWATT&
FormTEMPERATURE.ComboThermalConductivity.AddItem "Zircon = 0.042"
FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.NewIndex) = 0.042 * MILLIWATTSPERWATT&

' Load first value in combo if text is zero
If Val(FormTEMPERATURE.TextThermalConductivity.Text) = 0 Then
FormTEMPERATURE.ComboThermalConductivity.ListIndex = 0

Else
For i% = 0 To FormTEMPERATURE.ComboThermalConductivity.ListCount - 1
If MiscDifferenceIsSmall(FormTEMPERATURE.ComboThermalConductivity.ItemData(i%), TemperatureThermalConductivity! * MILLIWATTSPERWATT&, 0.0001) Then
FormTEMPERATURE.ComboThermalConductivity.ListIndex = i%
End If
Next i%
End If

Exit Sub

' Errors
TemperatureLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "TemperatureLoad"
ierror = True
Exit Sub

End Sub

Sub TemperatureSave()
' Save Temperature parameters

ierror = False
On Error GoTo TemperatureSaveError

If Val(FormTEMPERATURE.TextBeamEnergy.Text) < MINKILOVOLTS! Or Val(FormTEMPERATURE.TextBeamEnergy.Text) > MAXKILOVOLTS! Then
msg$ = FormTEMPERATURE.TextBeamEnergy.Text & " kilovolts is out of range! (must be between " & Format$(MINKILOVOLTS!) & " and " & Format$(MAXKILOVOLTS!) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "TemperatureSave"
ierror = True
Exit Sub
Else
TemperatureKilovolts! = Val(FormTEMPERATURE.TextBeamEnergy.Text)
End If

If Val(FormTEMPERATURE.TextBeamCurrent.Text) < MINBEAMCURRENT! Or Val(FormTEMPERATURE.TextBeamCurrent.Text) > MAXBEAMCURRENT! Then
msg$ = FormTEMPERATURE.TextBeamCurrent.Text & " beam current is out of range! (must be between " & Format$(MINBEAMCURRENT!) & " and " & Format$(MAXBEAMCURRENT!) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "TemperatureSave"
ierror = True
Exit Sub
Else
TemperatureBeamCurrent! = Val(FormTEMPERATURE.TextBeamCurrent.Text)
End If

If Val(FormTEMPERATURE.TextBeamSize.Text) <= 0# Or Val(FormTEMPERATURE.TextBeamSize.Text) > MAXBEAMSIZE! Then
msg$ = FormTEMPERATURE.TextBeamSize.Text & " beam size is out of range! (must be greater than zero and less than or equal to " & Format$(MAXBEAMSIZE!) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "TemperatureSave"
ierror = True
Exit Sub
Else
TemperatureBeamSize! = Val(FormTEMPERATURE.TextBeamSize.Text)
End If

If Val(FormTEMPERATURE.TextThermalConductivity.Text) < 0.0001 Or Val(FormTEMPERATURE.TextThermalConductivity.Text) > 100# Then
msg$ = FormTEMPERATURE.TextThermalConductivity.Text & " thermal conductivity is out of range! (must be between " & Format$(0.0001) & " and " & Format$(100#) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "TemperatureSave"
ierror = True
Exit Sub
Else
TemperatureThermalConductivity! = Val(FormTEMPERATURE.TextThermalConductivity.Text)
End If

Exit Sub

' Errors
TemperatureSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "TemperatureSave"
ierror = True
Exit Sub

End Sub

Sub TemperatureCalculate()
' Calculate temperature rise

ierror = False
On Error GoTo TemperatureCalculateError

Const TEMPSCALE# = 4.8

Dim temp1 As Single, temp2 As Single

temp1! = TEMPSCALE# * TemperatureKilovolts! * TemperatureBeamCurrent! / NAPERMA#
temp2! = TemperatureThermalConductivity! * TemperatureBeamSize!
If temp2! = 0# Then Exit Sub

msg$ = vbCrLf & "Beam conditions of " & Format$(TemperatureKilovolts!) & " keV, " & Format$(TemperatureBeamCurrent!) & " nA, " & Str$(TemperatureBeamSize!) & " um"
Call IOWriteLog(msg$)
msg$ = "Assuming thermal conductivity of " & Format$(TemperatureThermalConductivity!) & " W/cmK"
Call IOWriteLog(msg$)
msg$ = "Temperature rise = " & Str$(temp1! / temp2!) & " C"
FormTEMPERATURE.LabelTemperatureRise.Caption = msg$
Call IOWriteLog(msg$)

Exit Sub

' Errors
TemperatureCalculateError:
MsgBox Error$, vbOKOnly + vbCritical, "TemperatureCalculate"
ierror = True
Exit Sub

End Sub


