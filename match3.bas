Attribute VB_Name = "CodeMATCH3"
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit

Sub MatchTypeStandard2(sample() As TypeSample)
' Type out the standard composition (Standard.exe only). Note that "sample()" must be passed for Probewin.exe.

ierror = False
On Error GoTo MatchTypeStandard2Error

Dim number As Integer, i As Integer, stdnum As Integer

' Get standard from listbox
If FormMATCH.ListStandards.ListIndex < 0 Then Exit Sub
number% = FormMATCH.ListStandards.ItemData(FormMATCH.ListStandards.ListIndex)

' Load default match database (normally DHZ)
Call MatchSave(Int(1))
If ierror Then Exit Sub

' Recalculate and display standard data
If number% > 0 Then
Call StanFormCalculate(number%, Int(0))
If ierror Then Exit Sub

' Select clicked standard
For i% = 0 To FormMAIN.ListAvailableStandards.ListCount - 1
stdnum% = FormMAIN.ListAvailableStandards.ItemData(i%)
If stdnum% = number% Then
FormMAIN.ListAvailableStandards.ListIndex = i%
FormMAIN.ListAvailableStandards.Selected(i%) = True
End If
Next i%
End If

' Restore default database (STANDARD.MDB)
Call MatchSave(Int(2))
If ierror Then Exit Sub

Exit Sub

' Errors
MatchTypeStandard2Error:
MsgBox Error$, vbOKOnly + vbCritical, "MatchTypeStandard2"
ierror = True
Exit Sub

End Sub
