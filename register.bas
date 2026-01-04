Attribute VB_Name = "CodeREGISTER"
' (c) Copyright 1995-2026 by John J. Donovan
Option Explicit

Sub RegisterNow()
' Register the user's name and affiliation

ierror = False
On Error GoTo RegisterNowError

Dim tmousepointer As Integer

' Save original mouse pointer
tmousepointer% = Screen.MousePointer

' Ask for info
Screen.MousePointer = vbDefault
Call MiscAlwaysOnTop(True, FormREGISTER)
FormREGISTER.Show vbModal
Screen.MousePointer = tmousepointer% ' restore original mouse pointer
If ierror Then Exit Sub

Exit Sub

' Errors
RegisterNowError:
MsgBox Error$, vbOKOnly + vbCritical, "RegisterNow"
ierror = True
Exit Sub

End Sub

Sub RegisterSave()
' Save the user's name and affiliation

ierror = False
On Error GoTo RegisterSaveError

Dim tname As String
Dim tinstitution As String

' Check info
tname$ = FormREGISTER.TextName.Text
If Trim$(tname$) = vbNullString Then GoTo RegisterSaveBlankOrShort
If Len(Trim$(tname$)) < 3 Then GoTo RegisterSaveBlankOrShort

tinstitution$ = FormREGISTER.TextInstitution.Text
If Trim$(tinstitution$) = vbNullString Then GoTo RegisterSaveBlankOrShort
If Len(Trim$(tinstitution$)) < 3 Then GoTo RegisterSaveBlankOrShort

' Save to module level
RegistrationName$ = tname$
RegistrationInstitution$ = tinstitution$

Exit Sub

' Errors
RegisterSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "RegisterSave"
ierror = True
Exit Sub

RegisterSaveBlankOrShort:
msg$ = "Please fill out the registration information completely. Both the name and institution fields must be at least 3 characters long."
MsgBox msg$, vbOKOnly + vbExclamation, "RegisterSave"
ierror = True
Exit Sub

End Sub

Sub RegisterLoad2()
' Simple registration

ierror = False
On Error GoTo RegisterLoad2Error

Dim valid As Long

Dim lpAppName As String
Dim lpKeyName As String
Dim lpFileName As String
Dim lpString As String
Dim lpDefault As String

Dim lpReturnString As String * MAXINTEGER%
Dim nSize As Long

' Check for existing PROBEWIN.INI
If Dir$(ProbeWinINIFile$) = vbNullString Then
msg$ = "Unable to open file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "RegisterLoad2"
End
End If

' Use Windows API function to read PROBEWIN.INI
lpFileName$ = ProbeWinINIFile$

' Get strings
nSize& = Len(lpReturnString$)
lpAppName$ = "General"
lpKeyName$ = "RegistrationName"
lpDefault$ = vbNullString
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
If valid& > 0 Then
RegistrationName$ = Left$(lpReturnString$, valid&)
End If

nSize& = Len(lpReturnString$)
lpAppName$ = "General"
lpKeyName$ = "RegistrationInstitution"
lpDefault$ = vbNullString
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
If valid& > 0 Then
RegistrationInstitution$ = Left$(lpReturnString$, valid&)
End If

' Check registration
If Trim$(RegistrationName$) = vbNullString And Trim$(RegistrationInstitution$) = vbNullString Then
FormREGISTER.TextName.Text = MDBUserName$      ' use for early installations
FormREGISTER.TextInstitution.Text = MDBFileTitle$     ' use for early installations

' Allow user to change
Call RegisterNow
If ierror Then Exit Sub

' Save to INI file
lpAppName$ = "General"
lpKeyName$ = "RegistrationName"
lpString$ = RegistrationName$
valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, lpString$, lpFileName$)

lpAppName$ = "General"
lpKeyName$ = "RegistrationInstitution"
lpString$ = RegistrationInstitution$
valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, lpString$, lpFileName$)
End If

Exit Sub

' Errors
RegisterLoad2Error:
MsgBox Error$, vbOKOnly + vbCritical, "RegisterLoad2"
ierror = True
Exit Sub

End Sub
