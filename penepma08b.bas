Attribute VB_Name = "CodePenepma08b"
Option Explicit

Function Penepma08CheckPenepmaVersion() As Integer
' Check the Penepma version
' Returns 6 if Penepma06
' Returns 8 if Penepma08
' Returns 12 if Penepma12
' Returns 14 if Penepma14       ' does not officially exist
' Returns 16 if Penepma16
' Returns 0 if path not found

ierror = False
On Error GoTo Penepma08CheckPenepmaVersionError

' Check if path is valid
Penepma08CheckPenepmaVersion% = 0
If Dir$(PENEPMA_Root$, vbDirectory) = vbNullString Then Exit Function
If Dir$(PENEPMA_Path$, vbDirectory) = vbNullString Then Exit Function

' Check that other necessary folders exist
If Dir$(PENEPMA_Root$, vbDirectory) = vbNullString Then GoTo Penepma08CheckPenepmaVersionMissingFolderRoot
If Dir$(PENEPMA_Path$, vbDirectory) = vbNullString Then GoTo Penepma08CheckPenepmaVersionMissingFolderPenepma
If Dir$(PENDBASE_Path$, vbDirectory) = vbNullString Then GoTo Penepma08CheckPenepmaVersionMissingFolderPendbase
If Dir$(PENEPMA_PAR_Path$, vbDirectory) = vbNullString Then GoTo Penepma08CheckPenepmaVersionMissingFolderPenfluor

' Version 06
If InStr(PENEPMA_Root$, "06") > 0 Then
Penepma08CheckPenepmaVersion% = 6
Exit Function

' Version 08
ElseIf InStr(PENEPMA_Root$, "08") > 0 Then
Penepma08CheckPenepmaVersion% = 8
Exit Function

' Version 12
ElseIf InStr(PENEPMA_Root$, "12") > 0 Then
Penepma08CheckPenepmaVersion% = 12
Exit Function

' Version 14 (does not officially exist)
ElseIf InStr(PENEPMA_Root$, "14") > 0 Then
Penepma08CheckPenepmaVersion% = 14
Exit Function

' Version 16
ElseIf InStr(PENEPMA_Root$, "16") > 0 Then
Penepma08CheckPenepmaVersion% = 16
Exit Function

' Unable to determine version
Else
msg$ = "Unable to determine current version Penepma. Please check that the PENDBASE_Path, PENEPMA_Path and PENEPMA_Root strings are properly specified in the " & ProbeWinINIFile$ & " file."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08CheckPenepmaVersion"
ierror = True
Exit Function
End If

Exit Function

' Errors
Penepma08CheckPenepmaVersionError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma08CheckPenepmaVersion"
ierror = True
Exit Function

Penepma08CheckPenepmaVersionMissingFolderRoot:
msg$ = "The specified Penepma root path (" & PENEPMA_Root$ & ") was not found. Please check that the proper Penepma folders and files are present or edit the Penepma paths in the " & ProbeWinINIFile$ & " file."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08CheckVersionNumber"
ierror = True
Exit Function

Penepma08CheckPenepmaVersionMissingFolderPenepma:
msg$ = "The specified Penepma path (" & PENEPMA_Path$ & ") was not found. Please check that the proper Penepma folders and files are present or edit the Penepma paths in the " & ProbeWinINIFile$ & " file."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08CheckVersionNumber"
ierror = True
Exit Function

Penepma08CheckPenepmaVersionMissingFolderPendbase:
msg$ = "The specified Penepma Pendbase path (" & PENDBASE_Path$ & ") was not found. Please check that the proper Penepma folders and files are present or edit the Penepma paths in the " & ProbeWinINIFile$ & " file."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08CheckVersionNumber"
ierror = True
Exit Function

Penepma08CheckPenepmaVersionMissingFolderPenfluor:
msg$ = "The specified Penepma Penfluor path (" & PENEPMA_PAR_Path$ & ") was not found. Please check that the proper Penepma folders and files are present or edit the Penepma paths in the " & ProbeWinINIFile$ & " file."
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma08CheckVersionNumber"
ierror = True
Exit Function

End Function

