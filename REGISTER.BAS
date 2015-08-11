Attribute VB_Name = "CodeREGISTER"
' (c) Copyright 1995-2015 by John J. Donovan
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

' Convert globals and write to file
Call RegisterToFile
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

Sub RegisterToFile()
' Convert and write coded string to file

ierror = False
On Error GoTo RegisterToFileError

Dim astring As String * 188
Dim bstring As String * 188
Dim cstring As String * 188
Dim dstring As String * 188

' Save to fixed length strings
astring$ = RegistrationName$
bstring$ = RegistrationInstitution$

' XOR the strings with a key
Call RegisterXOR2(astring$, bstring$, cstring$, dstring$)
If ierror Then Exit Sub

' Set last character to ASCII 170
Mid$(cstring$, 188, 1) = Chr$(170)  ' only use ChrW$ for 0 to 127 ASCII
Mid$(dstring$, 188, 1) = Chr$(170)  ' only use ChrW$ for 0 to 127 ASCII

' Open x-ray line file and save to first and second record
Open XLineFile$ For Random Access Write As #XLineFileNumber% Len = 188
Put #XLineFileNumber%, 1, cstring$
Put #XLineFileNumber%, 2, dstring$
Close #XLineFileNumber%

Exit Sub

' Errors
RegisterToFileError:
MsgBox Error$, vbOKOnly + vbCritical, "RegisterToFile"
Close #XLineFileNumber%
ierror = True
Exit Sub

End Sub

Sub RegisterXOR(astring As String, bstring As String, cstring As String, dstring As String)
' XOR the astring and bstring to cstring and dstring

ierror = False
On Error GoTo RegisterXORError

Dim i As Integer, j As Integer
Dim n As Integer

For n% = 1 To 188
i% = Asc(Mid$(astring$, n%, 1))
j% = i% Xor 170
Mid$(cstring$, n%, 1) = Chr$(j%)  ' only use ChrW$ for 0 to 127 ASCII
i% = Asc(Mid$(bstring$, n%, 1))
j% = i% Xor 170
Mid$(dstring$, n%, 1) = Chr$(j%)  ' only use ChrW$ for 0 to 127 ASCII
Next n%

Exit Sub

' Errors
RegisterXORError:
MsgBox Error$, vbOKOnly + vbCritical, "RegisterXOR"
ierror = True
Exit Sub

End Sub

Sub RegisterXOR2(astring As String, bstring As String, cstring As String, dstring As String)
' XOR the astring and bstring to cstring and dstring (works under unicode)

ierror = False
On Error GoTo RegisterXOR2Error

Dim i As Byte, j As Byte
Dim n As Integer

For n% = 1 To 188
i = Asc(Mid$(astring$, n%, 1))
j = i Xor 170
Mid$(cstring$, n%, 1) = Chr$(j)  ' only use ChrW$ for 0 to 127 ASCII
i = Asc(Mid$(bstring$, n%, 1))
j = i Xor 170
Mid$(dstring$, n%, 1) = Chr$(j)  ' only use ChrW$ for 0 to 127 ASCII
Next n%

Exit Sub

' Errors
RegisterXOR2Error:
MsgBox Error$, vbOKOnly + vbCritical, "RegisterXOR2"
ierror = True
Exit Sub

End Sub

Sub RegisterLoad()
' Load the user's name and affiliation

ierror = False
On Error GoTo RegisterLoadError

Dim astring As String * 188
Dim bstring As String * 188
Dim cstring As String * 188
Dim dstring As String * 188

' Open x-ray line file and read first two records
Open XLineFile$ For Random Access Read As #XLineFileNumber% Len = 188
Get #XLineFileNumber%, 1, astring$
Get #XLineFileNumber%, 2, bstring$
Close #XLineFileNumber%

' Check last byte in first record (if registered it should be a ASCII 170)
If Mid$(astring, 188, 1) <> Chr$(170) Then  ' only use ChrW$ for 0 to 127 ASCII
FormREGISTER.TextName.Text = UCase$(Left$(MDBUserName$, 1)) & Mid$(MDBUserName$, 2)     ' use for early installations
FormREGISTER.TextInstitution.Text = MDBFileTitle$     ' use for early installations

' Allow user to change
Call RegisterNow
If ierror Then Exit Sub

' Save to file
Call RegisterToFile
If ierror Then Exit Sub

' Open x-ray line file and read first two records (again)
Open XLineFile$ For Random Access Read As #XLineFileNumber% Len = 188
Get #XLineFileNumber%, 1, astring$
Get #XLineFileNumber%, 2, bstring$
Close #XLineFileNumber%
End If

' XOR the strings with a key
Call RegisterXOR2(astring$, bstring, cstring$, dstring$)
If ierror Then Exit Sub

' Set last character to ASCII 32
Mid$(cstring$, 188, 1) = VbSpace$
Mid$(dstring$, 188, 1) = VbSpace$

' Save to globals
RegistrationName$ = Trim$(cstring$)
RegistrationInstitution$ = Trim$(dstring$)

Exit Sub

' Errors
RegisterLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "RegisterLoad"
Close #XLineFileNumber%
ierror = True
Exit Sub

End Sub

Sub RegisterLoad2()
' Simple registration (for unicode problems)

ierror = False
On Error GoTo RegisterLoad2Error

Dim astring As String * 188
Dim bstring As String * 188

Dim valid As Long

Dim lpAppName As String
Dim lpKeyName As String
Dim lpFileName As String
Dim lpString As String

' Open x-ray line file and read first two records
Open XLineFile$ For Random Access Read As #XLineFileNumber% Len = 188
Get #XLineFileNumber%, 1, astring$
Get #XLineFileNumber%, 2, bstring$
Close #XLineFileNumber%

' Check last byte in first record (if registered it should be a ASCII 170)
If Mid$(astring, 188, 1) <> Chr$(170) Then  ' only use ChrW$ for 0 to 127 ASCII
FormREGISTER.TextName.Text = MDBUserName$      ' use for early installations
FormREGISTER.TextInstitution.Text = MDBFileTitle$     ' use for early installations

' Allow user to change
Call RegisterNow
If ierror Then Exit Sub

' Check for existing PROBEWIN.INI
If Dir$(ProbeWinINIFile$) = vbNullString Then
msg$ = "Unable to open file " & ProbeWinINIFile$
MsgBox msg$, vbOKOnly + vbExclamation, "RegisterLoad2"
End
End If

' Use Windows API function to read PROBEWIN.INI
lpFileName$ = ProbeWinINIFile$

' Save to INI file
lpAppName$ = "General"
lpKeyName$ = "UserName"
lpString$ = RegistrationName$
valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, lpString$, lpFileName$)

lpAppName$ = "General"
lpKeyName$ = "Title"
lpString$ = RegistrationInstitution$
valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, lpString$, lpFileName$)

MDBUserName$ = RegistrationName$
MDBFileTitle$ = RegistrationInstitution$
Exit Sub
End If

' Load registration
RegistrationName$ = MDBUserName$
RegistrationInstitution$ = MDBFileTitle$

Exit Sub

' Errors
RegisterLoad2Error:
MsgBox Error$, vbOKOnly + vbCritical, "RegisterLoad2"
Close #XLineFileNumber%
ierror = True
Exit Sub

End Sub
