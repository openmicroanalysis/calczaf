Attribute VB_Name = "CodeEMP2"
' (c) Copyright 1995-2017 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub EmpLoadMACAPF(mode As Integer, emtz As Integer, emtx As Integer, absz As Integer, mac As Single, tstring As String, tfactor As Single, tstandard As String)
' This routine loads the empirical MAC/APF from the empirical arrays
'  mode% = 1  load from "macval" arrays (called by ZAFReadMu/ZAFFlu)
'  mode% = 2  load from "apfval" arrays (called by AnalyzeWeightCorrect)

ierror = False
On Error GoTo EmpLoadMACAPFError

Dim i As Integer

' Default
If mode% = 1 Then mac! = 0#
If mode% = 2 Then mac! = 1#

' Loop through all
For i% = 1 To MAXEMP%

' Find MAC
If mode% = 1 Then
If macez%(i%) = emtz% And macxl%(i%) = emtx% And macaz%(i%) = absz% Then
mac! = macval!(i%)
tstring$ = macstr$(i%)

tfactor! = macrenormfactor!(i%)
tstandard$ = macrenormstandard$(i%)
End If
End If

' Find APF
If mode% = 2 Then
If apfez%(i%) = emtz% And apfxl%(i%) = emtx% And apfaz%(i%) = absz% Then
mac! = apfval!(i%)
tstring$ = apfstr$(i%)

tfactor! = apfrenormfactor!(i%)
tstandard$ = apfrenormstandard$(i%)
End If
End If

Next i%

Exit Sub

' Errors
EmpLoadMACAPFError:
MsgBox Error$, vbOKOnly + vbCritical, "EmpLoadMACAPF"
ierror = True
Exit Sub

End Sub

