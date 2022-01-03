Attribute VB_Name = "CodeCOND"
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

Type TypeColumn2    ' Demo and Jeol
    ColumnConditionString As String
    takeoff As String
    kilovolts As String
    beamcurrent As String
    beamsize As String
End Type

Sub CondLoad()
' Loads the analytical condition defaults for user to change

ierror = False
On Error GoTo CondLoadError

' Load default
FormCOND.TextKiloVolts.Text = Format$(DefaultKiloVolts!)
FormCOND.TextTakeOff.Text = Format$(DefaultTakeOff!)
FormCOND.TextBeamCurrent.Text = Format$(DefaultBeamCurrent!)
FormCOND.TextBeamSize.Text = Format$(DefaultBeamSize!)

' Disable based on hardware options
If UCase$(app.EXEName) <> UCase$("Calczaf") Then
FormCOND.TextTakeOff.Enabled = False    ' always disabled, except in CalcZAF
End If

Exit Sub

' Errors
CondLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "CondLoad"
ierror = True
Exit Sub

End Sub

Sub CondSave()
' Save changed conditions to defaults

ierror = False
On Error GoTo CondSaveError

' Force reload of afactor arrays
AllAFactorUpdateNeeded = True

If Val(FormCOND.TextKiloVolts.Text) < MINKILOVOLTS! Or Val(FormCOND.TextKiloVolts.Text) > MAXKILOVOLTS! Then
msg$ = FormCOND.TextKiloVolts.Text & " kilovolts is out of range! (must be between " & Format$(MINKILOVOLTS!) & " and " & Format$(MAXKILOVOLTS!) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "CondSave"
ierror = True
Exit Sub
End If

If Val(FormCOND.TextTakeOff.Text) < 1# Or Val(FormCOND.TextTakeOff.Text) > 90# Then
msg$ = FormCOND.TextTakeOff.Text & " takeOff is out of range! (must be between 1 and 90)"
MsgBox msg$, vbOKOnly + vbExclamation, "CondSave"
ierror = True
Exit Sub
End If

If Val(FormCOND.TextBeamCurrent.Text) < 0.01 Or Val(FormCOND.TextBeamCurrent.Text) > 1000# Then
msg$ = "Beam current is out of range! (must be between 0.01 and 1000)"
MsgBox msg$, vbOKOnly + vbExclamation, "CondSave"
ierror = True
Exit Sub
End If

If Val(FormCOND.TextBeamSize.Text) < 0 Or Val(FormCOND.TextBeamSize.Text) > 500# Then
msg$ = "Beam size is out of range! (must be between 0 and 500)"
MsgBox msg$, vbOKOnly + vbExclamation, "CondSave"
ierror = True
Exit Sub
End If

' Save new defaults
DefaultKiloVolts! = Val(FormCOND.TextKiloVolts.Text)
DefaultTakeOff! = Val(FormCOND.TextTakeOff.Text)
DefaultBeamCurrent! = Val(FormCOND.TextBeamCurrent.Text)
DefaultBeamSize! = Val(FormCOND.TextBeamSize.Text)

Exit Sub

' Errors
CondSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "CondSave"
ierror = True
Exit Sub

End Sub

Sub CondBrowseColumnCondition(tForm As Form)
' Browse for a column condition and update analytical conditions text fields if possible

ierror = False
On Error GoTo CondBrowseColumnConditionError

Dim tfilename As String

' Get new filename from user (SX100 or JEOL)
Call IOGetFileName(Int(2), "PCC", tfilename$, tForm)
If ierror Then Exit Sub

' Update column condition field
tForm.TextColumnConditionString.Text = tfilename$

' Extract voltage, current, mag and spot modes


' Update analytical condition fields (see also CondLoadColumn2)


' Update form for using column conditions
tForm.OptionColumnConditionMethod(1).Value = True
Exit Sub

' Errors
CondBrowseColumnConditionError:
MsgBox Error$, vbOKOnly + vbCritical, "CondBrowseColumnCondition"
ierror = True
Exit Sub

End Sub

Sub CondColumnLoad2(columnstring As String, tForm As Form)
' Loads the analytical condition defaults based on column condition string (Demo and Jeol)

ierror = False
On Error GoTo CondColumnLoad2Error

Dim tfilename As String

Dim colrec As TypeColumn2

' Check for string to load
If Trim$(columnstring) = vbNullString Then Exit Sub

tfilename$ = ApplicationCommonAppData$ & "COLUMN2.DAT"
If Dir$(tfilename$) = vbNullString Then Exit Sub
Open tfilename$ For Input As #Temp1FileNumber%

' Loop through file
Do Until EOF(Temp1FileNumber%)
Input #Temp1FileNumber%, colrec.takeoff$, colrec.kilovolts$, colrec.beamcurrent$, colrec.beamsize$, colrec.ColumnConditionString$

' Check for match
If MiscStringsAreSame(colrec.ColumnConditionString$, columnstring$) Then

' Load into text fields
tForm.TextTakeOff.Text = Trim$(colrec.takeoff$)
tForm.TextKiloVolts.Text = Trim$(colrec.kilovolts$)
tForm.TextBeamCurrent.Text = Trim$(colrec.beamcurrent$)
tForm.TextBeamSize.Text = Trim$(colrec.beamsize$)
Close (Temp1FileNumber%)
Exit Sub
End If

Loop

Close (Temp1FileNumber%)
Exit Sub

' Errors
CondColumnLoad2Error:
Close (Temp1FileNumber%)
MsgBox Error$, vbOKOnly + vbCritical, "CondColumnLoad2"
ierror = True
Exit Sub

End Sub

