Attribute VB_Name = "CodePASSWORD"
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

Dim password As String

Sub PasswordLoad(tpassword As String, tFormCaption As String, tFormLabel As String)
' Get the user password

ierror = False
On Error GoTo PasswordLoadError

' Load form
FormMSGBOXPASSWORD.Caption = tFormCaption$
FormMSGBOXPASSWORD.Label1.Caption = tFormLabel$
FormMSGBOXPASSWORD.Show vbModal
If ierror Then Exit Sub

' Return the password
tpassword$ = password$
Exit Sub

' Errors
PasswordLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "PasswordLoad"
ierror = True
Exit Sub

End Sub

Sub PasswordSave()
' Save the user password

ierror = False
On Error GoTo PasswordSaveError

' Save to module level
password$ = Trim$(FormMSGBOXPASSWORD.TextPassword.Text)

Exit Sub

' Errors
PasswordSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "PasswordSave"
ierror = True
Exit Sub

End Sub

