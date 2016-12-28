Attribute VB_Name = "CodeGETZAFALL2"
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

Sub GetZAFAllSaveMAC2(itemp As Integer)
' Check for MAC file

ierror = False
On Error GoTo GetZAFAllSaveMAC2Error

' Check that file exists
MACFile$ = ApplicationCommonAppData$ & macstring2$(itemp%) & ".DAT"
If Dir$(MACFile$) = vbNullString Then GoTo GetZAFAllSaveMAC2NotFound

Exit Sub

' Errors
GetZAFAllSaveMAC2Error:
MsgBox Error$, vbOKOnly + vbCritical, "GetZAFAllSaveMAC2"
ierror = True
Exit Sub

GetZAFAllSaveMAC2NotFound:
msg$ = "File " & MACFile$ & " was not found, please choose another MAC file or create the missing file using the CalcZAF Xray menu items"
MsgBox msg$, vbOKOnly + vbExclamation, "GetZAFAllSaveMAC2"
ierror = True
Exit Sub

End Sub
