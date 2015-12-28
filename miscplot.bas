Attribute VB_Name = "CodeMiscPlot"
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

Sub MiscPlotGetSymbols_GS(nsets As Integer, tGraph As Graph)
' Generate random solid symbols (Graphics Server code)

ierror = False
On Error GoTo MiscPlotGetSymbols_GSError

Dim i As Integer, j As Integer

tGraph.AutoInc = 1  ' for loading symbols
i% = 3
For j% = 1 To nsets%
tGraph.SymbolData = i%
i% = i% + 2
If i% > 13 Then i% = 3
Next j%
tGraph.AutoInc = 0  ' for loading symbols

Exit Sub

' Errors
MiscPlotGetSymbols_GSError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscPlotGetSymbols_GS"
ierror = True
Exit Sub

End Sub

Sub MiscPlotGetSymbols_PE(nsets As Integer, tPesgo As Pesgo)
' Generate random solid symbols (Pro Essentials code)

ierror = False
On Error GoTo MiscPlotGetSymbols_PEError

Dim j As Integer

For j% = 0 To nsets% - 1
If j% Mod 5 = 0 Then
tPesgo.SubsetPointTypes(j%) = PEPT_DOTSOLID&
tPesgo.SubsetLineTypes(j%) = PELT_THIN_SOLID&
ElseIf j% Mod 5 = 1 Then
tPesgo.SubsetPointTypes(j%) = PEPT_SQUARESOLID&
tPesgo.SubsetLineTypes(j%) = PELT_THIN_SOLID&
ElseIf j% Mod 5 = 2 Then
tPesgo.SubsetPointTypes(j%) = PEPT_DIAMONDSOLID&
tPesgo.SubsetLineTypes(j%) = PELT_THIN_SOLID&
ElseIf j% Mod 5 = 3 Then
tPesgo.SubsetPointTypes(j%) = PEPT_UPTRIANGLESOLID&
tPesgo.SubsetLineTypes(j%) = PELT_THIN_SOLID&
ElseIf j% Mod 5 = 4 Then
tPesgo.SubsetPointTypes(j%) = PEPT_DOWNTRIANGLESOLID&
tPesgo.SubsetLineTypes(j%) = PELT_THIN_SOLID&
End If
Next j%

Exit Sub

' Errors
MiscPlotGetSymbols_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscPlotGetSymbols_PE"
ierror = True
Exit Sub

End Sub


