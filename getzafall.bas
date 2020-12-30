Attribute VB_Name = "CodeGETZAFALL"
' (c) Copyright 1995-2021 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub GetZAFAllLoad()
' Load correction method

ierror = False
On Error GoTo GetZAFAllLoadError

Dim i As Integer

' Note that "CorrectionFlag%" is handled different (dimensioned 0 to MAXCORRECTION%) (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
For i% = 0 To MAXCORRECTION%
FormGETZAFALL.Option6(i%).Caption = corstring(i%)
Next i%
FormGETZAFALL.Option6(CorrectionFlag%).value = True

' Load empirical alpha flag
If EmpiricalAlphaFlag% = 1 Then
FormGETZAFALL.CheckEmpiricalAlphaFlag.value = vbUnchecked
Else
FormGETZAFALL.CheckEmpiricalAlphaFlag.value = vbChecked
End If

' Load Penepma K-ratios flag (from binary compositions from Penfluor/Fanal)
If UsePenepmaKratiosFlag% = 1 Then
FormGETZAFALL.CheckUsePenepmaKratios.value = vbUnchecked
FormGETZAFALL.CheckPenepmaKratioLimit.Enabled = False
Else
FormGETZAFALL.CheckUsePenepmaKratios.value = vbChecked
FormGETZAFALL.CheckPenepmaKratioLimit.Enabled = True
End If

FormGETZAFALL.TextPenepmaKratioLimit.Text = Format$(PenepmaKratiosLimitValue!)

If UsePenepmaKratiosLimitFlag Then
FormGETZAFALL.CheckPenepmaKratioLimit.value = vbChecked
FormGETZAFALL.TextPenepmaKratioLimit.Enabled = True
Else
FormGETZAFALL.CheckPenepmaKratioLimit.value = vbUnchecked
FormGETZAFALL.TextPenepmaKratioLimit.Enabled = False
End If

If Penepma12UseKeVRoundingFlag Then
FormGETZAFALL.CheckPenepma12UseKeVRounding.value = vbChecked
Else
FormGETZAFALL.CheckPenepma12UseKeVRounding.value = vbUnchecked
End If

' Load the correction type (event sets the enables)
FormGETZAFALL.Option6(CorrectionFlag%).value = True

Exit Sub

' Errors
GetZAFAllLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "GetZAFAllLoad"
ierror = True
Exit Sub

End Sub

Sub GetZAFAllLoadMAC()
' Load FormMAC

ierror = False
On Error GoTo GETZAFAllLoadMACError

Dim i As Integer

' MACs
For i% = 1 To MAXMACTYPE%
FormMAC.Option6(i% - 1).Caption = macstring(i%)
If i% = MACTypeFlag% Then FormMAC.Option6(i% - 1).value = True
Next i%

Exit Sub

' Errors
GETZAFAllLoadMACError:
MsgBox Error$, vbOKOnly + vbCritical, "GetZAFAllLoadMAC"
ierror = True
Exit Sub

End Sub

Sub GetZAFAllSave()
' Save FormGetZAFAll correction

ierror = False
On Error GoTo GetZAFAllSaveError

Dim i As Integer

' Note that "CorrectionFlag%" is handled different (dimensioned 0 to MAXCORRECTION%) (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
For i% = 0 To MAXCORRECTION%
If FormGETZAFALL.Option6(i%).value Then CorrectionFlag% = i%
Next i%

' Save empirical factors flag
If FormGETZAFALL.CheckEmpiricalAlphaFlag.value = vbUnchecked Then
EmpiricalAlphaFlag% = 1
Else
EmpiricalAlphaFlag% = 2
End If

' Save Penepma K-ratios flag (from binary compositions from Penfluor/Fanal)
If FormGETZAFALL.CheckUsePenepmaKratios.value = vbUnchecked Then
UsePenepmaKratiosFlag% = 1
Else
UsePenepmaKratiosFlag% = 2
End If

If Val(FormGETZAFALL.TextPenepmaKratioLimit.Text) < 50# Or Val(FormGETZAFALL.TextPenepmaKratioLimit.Text) > 99# Then
msg$ = "Penepma kratios limit value out of range (must be between 50 and 99 wt. %)"
MsgBox msg$, vbOKOnly + vbExclamation, "GetZAFAllSave"
Else
PenepmaKratiosLimitValue! = Val(FormGETZAFALL.TextPenepmaKratioLimit.Text)
End If

If FormGETZAFALL.CheckPenepmaKratioLimit.value = vbChecked Then
UsePenepmaKratiosLimitFlag = True
Else
UsePenepmaKratiosLimitFlag = False
End If

If FormGETZAFALL.CheckPenepma12UseKeVRounding.value = vbChecked Then
Penepma12UseKeVRoundingFlag = True
Else
Penepma12UseKeVRoundingFlag = False
End If

' Indicate alpha-factor update
AllAnalysisUpdateNeeded = True
AllAFactorUpdateNeeded = True

Exit Sub

' Errors
GetZAFAllSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "GetZAFAllSave"
ierror = True
Exit Sub

End Sub

Sub GetZAFAllSaveMAC()
' Save FormMAC options

ierror = False
On Error GoTo GetZAFAllSaveMACError

Dim i As Integer
Dim itemp As Integer

' Load value
For i% = 1 To MAXMACTYPE%
If FormMAC.Option6(i% - 1).value Then itemp% = i%
Next i%

' Check that file exists
Call GetZAFAllSaveMAC2(itemp%)
If ierror Then Exit Sub

' Update global MAC file type flag
MACTypeFlag% = itemp%

' Indicate alpha-factor update
AllAnalysisUpdateNeeded = True
AllAFactorUpdateNeeded = True

Exit Sub

' Errors
GetZAFAllSaveMACError:
MsgBox Error$, vbOKOnly + vbCritical, "GetZAFAllSaveMAC"
ierror = True
Exit Sub

End Sub

Sub GetZAFAllOptions()
' Get options

ierror = False
On Error GoTo GetZAFAllOptionsError

' ZAF/Phi-Rho-Z or alpha factors
If FormGETZAFALL.Option6(0).value Or FormGETZAFALL.Option6(1).value Or FormGETZAFALL.Option6(2).value Or FormGETZAFALL.Option6(3).value Then
Call GetZAFLoad
If ierror Then Exit Sub
FormGETZAF.Show vbModal
End If

' Calibration curve
If FormGETZAFALL.Option6(4).value Then
End If

Exit Sub

' Errors
GetZAFAllOptionsError:
MsgBox Error$, vbOKOnly + vbCritical, "GetZAFAllOptions"
ierror = True
Exit Sub

End Sub

Sub GetZAFAllSetEnables()
' Set form enables based on selection

ierror = False
On Error GoTo GetZAFAllSetEnablesError

Dim i As Integer, tIndex As Integer

' Load option index
For i% = 0 To FormGETZAFALL.Option6.count - 1
If FormGETZAFALL.Option6(i%).value = True Then tIndex% = i%
Next i%

' Set enabled property on Empirical alpha factors
If tIndex% < 1 Or tIndex% > 4 Then
FormGETZAFALL.CheckEmpiricalAlphaFlag.Enabled = False
FormGETZAFALL.CheckUsePenepmaKratios.Enabled = False
FormGETZAFALL.CheckPenepmaKratioLimit.Enabled = False
FormGETZAFALL.CheckPenepma12UseKeVRounding.Enabled = False
FormGETZAFALL.TextPenepmaKratioLimit.Enabled = False

Else
FormGETZAFALL.CheckEmpiricalAlphaFlag.Enabled = True
FormGETZAFALL.CheckUsePenepmaKratios.Enabled = True

If FormGETZAFALL.CheckUsePenepmaKratios.value = vbChecked Then
FormGETZAFALL.CheckPenepmaKratioLimit.Enabled = True
FormGETZAFALL.CheckPenepma12UseKeVRounding.Enabled = True
Else
FormGETZAFALL.CheckPenepmaKratioLimit.Enabled = False
FormGETZAFALL.CheckPenepma12UseKeVRounding.Enabled = False
End If

If FormGETZAFALL.CheckUsePenepmaKratios.value = vbChecked And FormGETZAFALL.CheckPenepmaKratioLimit.value = vbChecked Then
FormGETZAFALL.TextPenepmaKratioLimit.Enabled = True
Else
FormGETZAFALL.TextPenepmaKratioLimit.Enabled = False
End If

End If

' Set enabled property for CommandOptions
If tIndex% <= 4 Then
FormGETZAFALL.CommandOptions.Enabled = True
FormGETZAFALL.CommandMACs.Enabled = True
Else
FormGETZAFALL.CommandOptions.Enabled = False
FormGETZAFALL.CommandMACs.Enabled = False
End If

Exit Sub

' Errors
GetZAFAllSetEnablesError:
MsgBox Error$, vbOKOnly + vbCritical, "GetZAFAllSetEnables"
ierror = True
Exit Sub

End Sub
