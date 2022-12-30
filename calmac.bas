Attribute VB_Name = "CodeCALMAC"
' (c) Copyright 1995-2023 by John J. Donovan
Option Explicit

Dim CalMACOldSample(1 To 1) As TypeSample

Global MACMode As Integer

Sub CalMacCalculate()
' Calculate the MAC using McMaster

ierror = False
On Error GoTo CalMacCalculateError

Dim i As Integer
Dim energy As Single
Dim aphoto As Single
Dim aelastic As Single
Dim ainelastic As Single
Dim atotal As Single

' Check xray
Call CalMACCheckMAC
If ierror Then Exit Sub

' Get the MAC for this binary
energy! = CalMACOldSample(1).LineEnergy!(1) / 1000#
Call CalMACGetMAC(energy!, aphoto!, aelastic!, ainelastic!, atotal!)
If ierror Then Exit Sub

' Update the display
FormMAIN.MSFlexGrid1.row = 1
FormMAIN.MSFlexGrid1.col = 0: FormMAIN.MSFlexGrid1.Text = Format(CalMACOldSample(1).Elsyms$(1), a90)
FormMAIN.MSFlexGrid1.col = 1: FormMAIN.MSFlexGrid1.Text = Format(CalMACOldSample(1).Xrsyms$(1), a90)
FormMAIN.MSFlexGrid1.col = 2: FormMAIN.MSFlexGrid1.Text = Format$(CalMACOldSample(1).Elsyms$(2), a90)
FormMAIN.MSFlexGrid1.col = 3: FormMAIN.MSFlexGrid1.Text = Format$(Format(energy!, f84), a90)
FormMAIN.MSFlexGrid1.col = 4: FormMAIN.MSFlexGrid1.Text = Format$(Format(aphoto!, e82), a90)
FormMAIN.MSFlexGrid1.col = 5: FormMAIN.MSFlexGrid1.Text = Format$(Format(aelastic!, e82), a90)
FormMAIN.MSFlexGrid1.col = 6: FormMAIN.MSFlexGrid1.Text = Format$(Format(ainelastic!, e82), a90)
FormMAIN.MSFlexGrid1.col = 7: FormMAIN.MSFlexGrid1.Text = Format$(Format(atotal!, e82), a90)

' Write strings to Log Window
msg$ = ""
FormMAIN.MSFlexGrid1.row = 0
For i% = 0 To FormMAIN.MSFlexGrid1.cols - 1
FormMAIN.MSFlexGrid1.col = i
msg$ = msg$ & FormMAIN.MSFlexGrid1.Text
Next i%
Call IOWriteLog(msg$)

msg$ = ""
FormMAIN.MSFlexGrid1.row = 1
For i% = 0 To FormMAIN.MSFlexGrid1.cols - 1
FormMAIN.MSFlexGrid1.col = i
msg$ = msg$ & FormMAIN.MSFlexGrid1.Text
Next i%
Call IOWriteLog(msg$)

Exit Sub

' Errors
CalMacCalculateError:
MsgBox Error$, vbOKOnly + vbCritical, "CalMacCalculate"
ierror = True
Exit Sub

End Sub

Sub CalMacChange()
' Load default x-ray

ierror = False
On Error GoTo CalMacChangeError

Dim ip As Integer
Dim sym As String

sym$ = FormMAIN.ComboElement.Text
ip% = IPOS1(MAXELM%, sym$, Symlo$())
If ip% > 0 Then
If FormMAIN.ComboXRay.Text = "" Then FormMAIN.ComboXRay.Text = Deflin$(ip%)
If sym$ <> CalMACOldSample(1).Elsyms$(1) Then FormMAIN.ComboXRay.Text = Deflin$(ip%)
End If

Exit Sub

' Errors
CalMacChangeError:
MsgBox Error$, vbOKOnly + vbCritical, "CalMacChange"
ierror = True
Exit Sub

End Sub

Sub CalMacLoad()
' Load FormMAIN for CalMac

ierror = False
On Error GoTo CalMacLoadError

Dim i As Integer

' Add the list box items
FormMAIN.ComboElement.Clear
For i% = 0 To MAXELM% - 1
FormMAIN.ComboElement.AddItem Symlo$(i% + 1)
Next i%
FormMAIN.ComboElement.ListIndex = 11    ' Mg

FormMAIN.ComboXRay.Clear
For i% = 0 To MAXRAY% - 2
FormMAIN.ComboXRay.AddItem Xraylo$(i% + 1)
Next i%
FormMAIN.ComboXRay.ListIndex = 0    ' Ka

FormMAIN.ComboAbsorber.Clear
For i% = 0 To MAXELM% - 1
FormMAIN.ComboAbsorber.AddItem Symlo(i% + 1)
Next i%
FormMAIN.ComboAbsorber.ListIndex = 25   ' Fe

FormMAIN.TextKeV.Text = Str$(DefaultKiloVolts!)

' Initialize the Output Grid
FormMAIN.MSFlexGrid1.RowHeight(0) = FormMAIN.MSFlexGrid1.Height / FormMAIN.MSFlexGrid1.rows
FormMAIN.MSFlexGrid1.RowHeight(1) = FormMAIN.MSFlexGrid1.Height / FormMAIN.MSFlexGrid1.rows

FormMAIN.MSFlexGrid1.row = 0
FormMAIN.MSFlexGrid1.col = 0: FormMAIN.MSFlexGrid1.Text = Format$("Element", a90)
FormMAIN.MSFlexGrid1.col = 1: FormMAIN.MSFlexGrid1.Text = Format("X-ray", a90)
FormMAIN.MSFlexGrid1.col = 2: FormMAIN.MSFlexGrid1.Text = Format("Absorb", a90)
FormMAIN.MSFlexGrid1.col = 3: FormMAIN.MSFlexGrid1.Text = Format("Energy", a90)
FormMAIN.MSFlexGrid1.col = 4: FormMAIN.MSFlexGrid1.Text = Format("Photo", a90)
FormMAIN.MSFlexGrid1.col = 5: FormMAIN.MSFlexGrid1.Text = Format("Elastic", a90)
FormMAIN.MSFlexGrid1.col = 6: FormMAIN.MSFlexGrid1.Text = Format("Inelast", a90)
FormMAIN.MSFlexGrid1.col = 7: FormMAIN.MSFlexGrid1.Text = Format("Total", a90)

Exit Sub

' Errors
CalMacLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "CalMacLoad"
ierror = True
Exit Sub

End Sub

Sub CalMacCalculateRange()
' Calculate the MAC range using McMaster

ierror = False
On Error GoTo CalMacCalculateRangeError

Dim i As Integer

Dim energy As Single
Dim aphoto As Single
Dim aelastic As Single
Dim ainelastic As Single
Dim atotal As Single

' Check xray
Call CalMACCheckMAC
If ierror Then Exit Sub

Call IOWriteLog("")
If FormMAIN.menuMethodMcMasterMACs.Checked Then
Call IOWriteLog("McMaster MAC +-100 eV Range")
ElseIf FormMAIN.menuMethodMAC30MACs.Checked Then
Call IOWriteLog("MAC30 MAC +-100 eV Range")
ElseIf FormMAIN.menuMethodJTAMACs.Checked Then
Call IOWriteLog("JTA MAC +-100 eV Range")
End If

' Calculate +- 100 eV on a side
Call IOWriteLog("")
For i% = -100 To 100
energy! = CalMACOldSample(1).LineEnergy!(1) / 1000# + i% / 1000#    ' 1 eV intervals

' Get the MAC for this binary
Call CalMACGetMAC(energy!, aphoto!, aelastic!, ainelastic!, atotal!)
If ierror Then Exit Sub

' Print
msg$ = "Energy= " & MiscAutoFormat$(energy!) & " Angstrom= " & MiscAutoFormat$(ANGKEV! / energy!) & " MAC= " & MiscAutoFormat$(atotal!)
Call IOWriteLog(msg$)
Next i%
Call IOWriteLog("")

Exit Sub

' Errors
CalMacCalculateRangeError:
MsgBox Error$, vbOKOnly + vbCritical, "CalMacCalculateRange"
ierror = True
Exit Sub

End Sub


Sub CalMACGetMAC(energy As Single, aphoto As Single, aelastic As Single, ainelastic As Single, atotal As Single)
' Get the MAC from appropriate Absorb routine

ierror = False
On Error GoTo CalMacGetMACError

Dim iz As Integer
Dim ielm As Integer, iray As Integer

Static initialized1 As Integer, initialized2 As Integer
Static g(3, 95) As Single
Static o(9, 95) As Single

Static lines(1 To 12, 1 To 99) As Double
Static edges(1 To 12, 1 To 99) As Double

' Load the absorber atomic number
iz% = CalMACOldSample(1).AtomicNums%(2)

' McMaster
If MACMode% = 0 Then
Call AbsorbGetMAC(iz%, energy!, aphoto!, aelastic!, ainelastic!, atotal!)
End If

' If MAC30, load line and edge energies from LINES2.DAT
If MACMode% = 1 And Not initialized1 Then
Call AbsorbLoadLINES2DataFile(lines#(), edges#())
If ierror Then Exit Sub
initialized1 = True
End If

' If MACJTA, load line and edge energies from LINES.DAT
If MACMode% = 2 And Not initialized2 Then
Call AbsorbLoadLINESDataFile(g!(), o!())
If ierror Then Exit Sub
initialized2 = True
End If

' Load x-ray and line if not using arbitrary energy
If energy! = 0# Then
ielm% = CalMACOldSample(1).AtomicNums%(1)
iray% = CalMACOldSample(1).XrayNums%(1)
End If

' MAC30
If MACMode% = 1 Then
Call AbsorbGetMAC30(energy!, iz%, ielm%, iray%, lines#(), edges#(), atotal!)
End If

' MACJTA
If MACMode% = 2 Then
Call AbsorbGetMACJTA(energy!, iz%, ielm%, iray%, g!(), o!(), atotal!)
End If

Exit Sub

' Errors
CalMacGetMACError:
MsgBox Error$, vbOKOnly + vbCritical, "CalMacGetMAC"
ierror = True
Exit Sub

End Sub

Sub CalMACCheckMAC()
' Check the arrays

ierror = False
On Error GoTo CalMacCheckMACError

' Get the z absorber number and the energy
CalMACOldSample(1).Elsyms$(1) = FormMAIN.ComboElement.Text
CalMACOldSample(1).Xrsyms$(1) = FormMAIN.ComboXRay.Text
CalMACOldSample(1).Elsyms$(2) = FormMAIN.ComboAbsorber.Text
DefaultKiloVolts! = Val(FormMAIN.TextKeV.Text)

CalMACOldSample(1).kilovolts! = DefaultKiloVolts!
CalMACOldSample(1).takeoff! = DefaultTakeOff!
CalMACOldSample(1).KilovoltsArray!(1) = DefaultKiloVolts!
CalMACOldSample(1).TakeoffArray!(1) = DefaultTakeOff!

CalMACOldSample(1).numcat%(1) = 1
CalMACOldSample(1).numcat%(2) = 1
CalMACOldSample(1).numoxd%(1) = 0
CalMACOldSample(1).numoxd%(2) = 0

If CalMACOldSample(1).Elsyms$(1) = "" Then GoTo CalMacCheckMACNoEmitter
If CalMACOldSample(1).Xrsyms$(1) = "" Then GoTo CalMacCheckMACNoXray
If CalMACOldSample(1).Elsyms$(2) = "" Then GoTo CalMacCheckMACNoAbsorber

' Load the element arrays
CalMACOldSample(1).LastElm% = 1
CalMACOldSample(1).LastChan% = 2

' Check for valid element symbols
Call ElementCheckElement(CalMACOldSample())
If ierror Then Exit Sub

' Fill element arrays
Call ElementLoadArrays(CalMACOldSample())
If ierror Then Exit Sub

' Check for valid xray symbols
Call ElementCheckXray(Int(1), CalMACOldSample())
If ierror Then Exit Sub

Exit Sub

' Errors
CalMacCheckMACError:
MsgBox Error$, vbOKOnly + vbCritical, "CalMacCheckMAC"
ierror = True
Exit Sub

CalMacCheckMACNoEmitter:
msg$ = "No emitting element was entered"
MsgBox msg$, vbOKOnly + vbExclamation, "CalMacCheckMAC"
ierror = True
Exit Sub

CalMacCheckMACNoXray:
msg$ = "No emitting x-ray was entered"
MsgBox msg$, vbOKOnly + vbExclamation, "CalMacCheckMAC"
ierror = True
Exit Sub

CalMacCheckMACNoAbsorber:
msg$ = "No absorbing element was entered"
MsgBox msg$, vbOKOnly + vbExclamation, "CalMacCheckMAC"
ierror = True
Exit Sub

End Sub
