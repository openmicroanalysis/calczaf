Attribute VB_Name = "CodeMiscFile"
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Private Type SHFILEOPSTRUCT
   hWnd        As Long
   wFunc       As Long
   pFrom       As String
   pTo         As String
   fFlags      As Integer
   fAborted    As Boolean
   hNameMaps   As Long
   sProgress   As String
 End Type
  
Private Const FO_MOVE As Long = &H1
Private Const FO_COPY As Long = &H2
'Private Const FO_DELETE As Long = &H3
'Private Const FO_RENAME As Long = &H4

Private Const FOF_SILENT As Long = &H4
'Private Const FOF_RENAMEONCOLLISION As Long = &H8
Private Const FOF_NOCONFIRMATION As Long = &H10
Private Const FOF_SIMPLEPROGRESS As Long = &H100
'Private Const FOF_ALLOWUNDO As Long = &H40

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nSize As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHFileOperation Lib "shell32" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
   
'O_FUNC - the File Operation to perform, determined by the type of SHFileOperation action chosen (move/copy)
Dim FO_FUNC As Long

Private Type BROWSEINFO
   hOwner           As Long
   pidlRoot         As Long
   pszDisplayName   As String
   lpszTitle        As String
   ulFlags          As Long
   lpfn             As Long
   lparam           As Long
   iImage           As Long
End Type
   
'Private Const ERROR_SUCCESS As Long = 0
'Private Const CSIDL_DESKTOP As Long = &H0
'Private Const BIF_RETURNONLYFSDIRS As Long = &H1
'Private Const BIF_STATUSTEXT As Long = &H4
'Private Const BIF_RETURNFSANCESTORS As Long = &H8
         
' For ease of reading, constants are substituted for SHFileOperation numbers in code
'Const FileMove As Integer = 1
'Const FileCopy As Integer = 2
  
'Check button index constants
'Const optSilent As Integer = 0
'Const optNoFilenames As Integer = 1
'Const optNoConfirmDialog As Integer = 2
'Const optRenameIfExists As Integer = 3
'Const optPromptMeFirst As Integer = 4

Sub MiscCheckName(mode As Integer, astring As String)
' Check that the user is not using a reserved filename
' mode = 0 check all names
' mode = 1 do not check for "standard" (called from STANDARD.EXE)
' mode = 2 do not check for "user" (called from USER.EXE)

ierror = False
On Error GoTo MiscCheckNameError

' If not called from STANDARD.EXE, check for reserved standard file name
If mode% <> 1 Then
If InStr(astring$, "STANDARD.MDB") > 0 Then GoTo MiscCheckNameReserved
End If

' If not called from USER.EXE, check for reserved user file name
If mode% <> 2 Then
If InStr(astring$, "USER.MDB") > 0 Then GoTo MiscCheckNameReserved
End If

' Check for PROBEWIN reserved names
If InStr(astring$, "SETUP.MDB") > 0 Then GoTo MiscCheckNameReserved
If InStr(astring$, "SETUP2.MDB") > 0 Then GoTo MiscCheckNameReserved
If InStr(astring$, "SETUP3.MDB") > 0 Then GoTo MiscCheckNameReserved
If InStr(astring$, "MATRIX.MDB") > 0 Then GoTo MiscCheckNameReserved
If InStr(astring$, "POSITION.MDB") > 0 Then GoTo MiscCheckNameReserved
If InStr(astring$, "XRAY.MDB") > 0 Then GoTo MiscCheckNameReserved
If InStr(astring$, "TEMP.MDB") > 0 Then GoTo MiscCheckNameReserved

If InStr(astring$, "PROBEWIN.") > 0 Then GoTo MiscCheckNameReserved
If InStr(astring$, "STARTWIN") > 0 Then GoTo MiscCheckNameReserved
If InStr(astring$, "JOYWIN.") > 0 Then GoTo MiscCheckNameReserved
If InStr(astring$, "STAGE.") > 0 Then GoTo MiscCheckNameReserved
If InStr(astring$, "USERWIN.") > 0 Then GoTo MiscCheckNameReserved

If InStr(astring$, "ABSORB.DAT") > 0 Then GoTo MiscCheckNameReserved
If InStr(astring$, "ELEMENTS.DAT") > 0 Then GoTo MiscCheckNameReserved
If InStr(astring$, "CRYSTALS.DAT") > 0 Then GoTo MiscCheckNameReserved
If InStr(astring$, "MOTORS.DAT") > 0 Then GoTo MiscCheckNameReserved
If InStr(astring$, "SCALERS.DAT") > 0 Then GoTo MiscCheckNameReserved

If InStr(astring$, "XLINE.DAT") > 0 Then GoTo MiscCheckNameReserved
If InStr(astring$, "XEDGE.DAT") > 0 Then GoTo MiscCheckNameReserved
If InStr(astring$, "XFLUR.DAT") > 0 Then GoTo MiscCheckNameReserved

If InStr(astring$, "EMPMAC.DAT") > 0 Then GoTo MiscCheckNameReserved
If InStr(astring$, "EMPAPF.DAT") > 0 Then GoTo MiscCheckNameReserved
If InStr(astring$, "EMPFAC.DAT") > 0 Then GoTo MiscCheckNameReserved
If InStr(astring$, "EMPPHA.DAT") > 0 Then GoTo MiscCheckNameReserved

Exit Sub

' Errors
MiscCheckNameError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscCheckName"
ierror = True
Exit Sub

MiscCheckNameReserved:
msg$ = "Filename " & astring$ & " is a reserved filename, please select another filename"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscCheckName"
ierror = True
Exit Sub

End Sub

Sub MiscModifyStringToFilename(astring As String)
' Procedure to modify a given string to make sure that no invalid file name
' characters are present for a given string. Should only be used on a file
' name only (no path).

ierror = False
On Error GoTo MiscModifyStringToFilenameError

If Trim$(astring$) = vbNullString Then GoTo MiscModifyStringToFilenameNoString

' Replace invalid characters
Call MiscReplaceString(astring$, "\", "_")
If ierror Then Exit Sub
Call MiscReplaceString(astring$, "/", "_")
If ierror Then Exit Sub
Call MiscReplaceString(astring$, VbDquote$, "'")
If ierror Then Exit Sub
Call MiscReplaceString(astring$, ":", ";")
If ierror Then Exit Sub
Call MiscReplaceString(astring$, "&", "+")
If ierror Then Exit Sub
Call MiscReplaceString(astring$, "*", "-")
If ierror Then Exit Sub
Call MiscReplaceString(astring$, "|", "-")
If ierror Then Exit Sub
Call MiscReplaceString(astring$, ">", "-")
If ierror Then Exit Sub
Call MiscReplaceString(astring$, "<", "-")
If ierror Then Exit Sub
Call MiscReplaceString(astring$, "?", "!")
If ierror Then Exit Sub

Exit Sub

' Errors
MiscModifyStringToFilenameError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscModifyStringToFilenameError"
ierror = True
Exit Sub

MiscModifyStringToFilenameNoString:
msg$ = "String is blank"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscModifyStringToFilenameError"
ierror = True
Exit Sub

End Sub

Sub MiscChangePath(tpath As String)
' Routine to change the path

ierror = False
On Error GoTo MiscChangePathError

If Trim$(tpath$) = vbNullString Then Exit Sub

If Left$(Trim$(tpath$), 2) <> "\\" Then ChDrive tpath$
If MiscGetPathOnly$(tpath$) <> vbNullString Then
ChDir MiscGetPathOnly$(tpath$)
End If

Exit Sub

' Errors
MiscChangePathError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscChangePath"
ierror = True
Exit Sub

End Sub

Function MiscGetFileNameNoExtension(afilename As String) As String
' Returns (path &) filename without an extension

ierror = False
On Error GoTo MiscGetFileNameNoExtensionError

Dim n As Integer, i As Integer
Dim tfilename As String

' Load copy of filename (do not modify original)
tfilename$ = afilename$

' Loop from end (modified for 32 bit and multiple occurances of ".")
n% = 0
For i% = Len(tfilename$) To 1 Step -1
If Mid$(tfilename$, i%, 1) = "." Then
n% = i%
Exit For    ' exit on first occurance (fixed 04-12-2012)
End If
Next i%

' Remove extension
If n% > 1 Then tfilename$ = Left$(tfilename$, n% - 1)

MiscGetFileNameNoExtension$ = tfilename$
Exit Function

' Errors
MiscGetFileNameNoExtensionError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetFileNameNoExtension"
ierror = True
Exit Function

End Function

Function MiscGetFileNameOnly(afilename As String) As String
' Returns only the filename without a path

ierror = False
On Error GoTo MiscGetFileNameOnlyError

Dim n As Integer
Dim tfilename As String

' Load copy of filename (do not modify original)
tfilename$ = afilename$

' Find last backslash
n% = InStr(tfilename$, "\")
Do While n% > 0
tfilename$ = Mid$(tfilename$, n% + 1, Len(tfilename$))
n% = InStr(tfilename$, "\")
Loop

MiscGetFileNameOnly$ = tfilename$
Exit Function

' Errors
MiscGetFileNameOnlyError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetFileNameOnly"
ierror = True
Exit Function

End Function

Function MiscGetNumberofColumns(tfilename As String) As Integer
' Determine the number of columns in a data file line (does not work for string data, only numeric)

ierror = False
On Error GoTo MiscGetNumberofColumnsError

Dim n As Integer
Dim astring As String

' Get first line of input file
Open tfilename$ For Input As #Temp1FileNumber%
Line Input #Temp1FileNumber%, astring$
Close #Temp1FileNumber%

' Check for empty string
astring$ = Trim$(astring$)
If astring$ = vbNullString Then GoTo MiscGetNumberofColumnsEmpty

' Replace tabs with spaces (bug in VB if tabs with no spaces)
Call MiscReplaceString(astring$, vbTab, VbSpace$)
If ierror Then Exit Function

' Write string to temp file
Open ProgramPath$ & "MODAL.TMP" For Output As #Temp1FileNumber%
Print #Temp1FileNumber%, astring$
Close #Temp1FileNumber%

' Loop until end of temp file
Open ProgramPath$ & "MODAL.TMP" For Input As #Temp1FileNumber%

n% = 0
Do Until EOF(Temp1FileNumber%)
Input #Temp1FileNumber%, astring$
n% = n% + 1
Loop
Close #Temp1FileNumber%

If n% = 0 Then GoTo MiscGetNumberofColumnsZero
MiscGetNumberofColumns% = n%

Exit Function

' Errors
MiscGetNumberofColumnsError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetNumberofColumns"
Close #Temp1FileNumber%
ierror = True
Exit Function

MiscGetNumberofColumnsEmpty:
msg$ = "No column label line in " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "MiscGetNumberofColumns"
Close #Temp1FileNumber%
ierror = True
Exit Function

MiscGetNumberofColumnsZero:
msg$ = "No data columns found in " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "MiscGetNumberofColumns"
Close #Temp1FileNumber%
ierror = True
Exit Function

End Function

Function MiscGetNumberofColumnsA(tfilename As String, achar As String) As Integer
' Determine the number of columns in a data file line using the passed delimiter (works for numeric or text data)

ierror = False
On Error GoTo MiscGetNumberofColumnsAError

Dim n As Integer
Dim astring As String, bstring As String

' Get first line of input file
Open tfilename$ For Input As #Temp1FileNumber%
Line Input #Temp1FileNumber%, astring$
Close #Temp1FileNumber%

' Check for empty string
astring$ = Trim$(astring$)
If astring$ = vbNullString Then GoTo MiscGetNumberofColumnsAEmpty

' Loop until no length
n% = 0
Do Until astring$ = vbNullString
Call MiscParseStringToStringA(astring$, achar$, bstring$)
n% = n% + 1
Loop

If n% = 0 Then GoTo MiscGetNumberofColumnsAZero
MiscGetNumberofColumnsA% = n%

Exit Function

' Errors
MiscGetNumberofColumnsAError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetNumberofColumnsA"
Close #Temp1FileNumber%
ierror = True
Exit Function

MiscGetNumberofColumnsAEmpty:
msg$ = "No column label line in " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "MiscGetNumberofColumnsA"
Close #Temp1FileNumber%
ierror = True
Exit Function

MiscGetNumberofColumnsAZero:
msg$ = "No data columns found in " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "MiscGetNumberofColumnsA"
Close #Temp1FileNumber%
ierror = True
Exit Function

End Function

Function MiscGetNumberofColumnsB(astring As String, achar As String) As Integer
' Determine the number of columns in a string using the passed delimiter (works for numeric or text data)

ierror = False
On Error GoTo MiscGetNumberofColumnsBError

Dim n As Integer
Dim bstring As String

' Check for empty string
astring$ = Trim$(astring$)
If astring$ = vbNullString Then GoTo MiscGetNumberofColumnsBEmpty

' Loop until no length
n% = 0
Do Until astring$ = vbNullString
Call MiscParseStringToStringA(astring$, achar$, bstring$)
n% = n% + 1
Loop

If n% = 0 Then GoTo MiscGetNumberofColumnsBZero
MiscGetNumberofColumnsB% = n%

Exit Function

' Errors
MiscGetNumberofColumnsBError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetNumberofColumnsB"
ierror = True
Exit Function

MiscGetNumberofColumnsBEmpty:
msg$ = "Empty string was passed to function"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscGetNumberofColumnsB"
ierror = True
Exit Function

MiscGetNumberofColumnsBZero:
msg$ = "No data columns (no text delimiters) found in passed string"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscGetNumberofColumnsB"
ierror = True
Exit Function

End Function

Function MiscGetPathOnly(afilename As String) As String
' Returns the path only

ierror = False
On Error GoTo MiscGetPathOnlyError

Dim n As Integer
Dim tfilename As String

If afilename$ = vbNullString Then Exit Function

' Load copy of filename (do not modify original)
tfilename$ = afilename$
If Right$(tfilename$, 1) = "\" Then
MiscGetPathOnly$ = tfilename$           ' just the path was passed already
Exit Function
End If

' Extract path only
n% = InStr(tfilename$, MiscGetFileNameOnly$(tfilename$))
tfilename$ = Left$(tfilename$, n% - 1)

' Return path only
MiscGetPathOnly$ = tfilename$
Exit Function

' Errors
MiscGetPathOnlyError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetPathOnly"
ierror = True
Exit Function

End Function

Function MiscGetPathOnly2(afilename As String) As String
' Returns the path only (without a trailing "\"

ierror = False
On Error GoTo MiscGetPathOnly2Error

Dim tfilename As String

If afilename$ = vbNullString Then Exit Function
tfilename$ = MiscGetPathOnly$(afilename$)

' Remove trailing "\"
If Right$(tfilename$, 1) = "\" Then tfilename$ = Left$(tfilename$, Len(tfilename$) - 1)

' Return path only
MiscGetPathOnly2$ = tfilename$
Exit Function

' Errors
MiscGetPathOnly2Error:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetPathOnly2"
ierror = True
Exit Function

End Function

Function MiscGetLastFolderOnly(afilename As String) As String
' Returns the last folder name before the filename (between two backslashes)

ierror = False
On Error GoTo MiscGetLastFolderOnlyError

Dim tfilename As String
Dim n As Integer

If afilename$ = vbNullString Then Exit Function
tfilename$ = MiscGetPathOnly$(afilename$)

' Remove trailing "\"
If Right$(tfilename$, 1) = "\" Then tfilename$ = Left$(tfilename$, Len(tfilename$) - 1)

' Find next backslash
n% = InStrRev(tfilename$, "\")
If n% > 0 Then tfilename$ = Right$(tfilename$, Len(tfilename$) - n%)

' Return last folder only
MiscGetLastFolderOnly$ = tfilename$
Exit Function

' Errors
MiscGetLastFolderOnlyError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetLastFolderOnly"
ierror = True
Exit Function

End Function

Function MiscGetFileNameExtensionOnly(afilename As String) As String
' Returns filename extension only (with dot).

ierror = False
On Error GoTo MiscGetFileNameExtensionOnlyError

Dim n As Integer, i As Integer
Dim textension As String

' Assume no extension
textension$ = "."

' Loop from end (modified for 32 bit and multiple occurances of ".")
n% = 0
For i% = Len(afilename$) To 1 Step -1
If Mid$(afilename$, i%, 1) = "." Then
n% = i%
Exit For
End If
Next i%

' Save extension only
If n% > 1 Then textension$ = Mid$(afilename$, n%)

MiscGetFileNameExtensionOnly$ = textension$
Exit Function

' Errors
MiscGetFileNameExtensionOnlyError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetFileNameExtensionOnly"
ierror = True
Exit Function

End Function

Sub MiscReadUntilDelimit(tImportDataFileNumber As Integer, astring As String, delimit As String)
' Read a file character by character until the delimit character

ierror = False
On Error GoTo MiscReadUntilDelimitError
    
Dim achar As String
    
astring$ = vbNullString
Do While Not EOF(tImportDataFileNumber%)        ' loop until end of file
    achar$ = Input(1, #tImportDataFileNumber%)  ' get one character
    If achar$ = delimit$ Then Exit Do
    astring$ = astring$ & achar$                ' add to string
Loop

Exit Sub

' Errors
MiscReadUntilDelimitError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "MiscReadUntilDelimit"
Close #tImportDataFileNumber%
ierror = True
Exit Sub

End Sub

Function MiscDeleteLines(from_name As String, to_name As String, tstring As String, istring As String) As Long
' Delete lines containing the specified target string from a text file
'   from_name = file to copy
'   to_name = file copied
'   tstring = target string (skip lines containing this string)
'   istring = ignore string (always copy lines containing this string)

ierror = False
On Error GoTo MiscDeleteLinesError

Dim deleted_lines As Long
Dim astring As String

' Open the input file
If Dir$(from_name$) = vbNullString Then GoTo MiscDeleteLinesNoFromFile
Open from_name For Input As Temp1FileNumber%

' Open the output file
Open to_name For Output As Temp2FileNumber%

' Copy the file skipping lines containing the target
deleted_lines& = 0
Do While Not EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$
   If InStr(astring$, tstring) = 0 Or InStr(astring$, istring) > 0 Then
         Print #Temp2FileNumber%, astring$
   Else
         deleted_lines& = deleted_lines& + 1
   End If
Loop

' Close the files
Close #Temp1FileNumber%
Close #Temp2FileNumber%

MiscDeleteLines& = deleted_lines&
Exit Function

' Errors
MiscDeleteLinesError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscDeleteLines"
Close #Temp1FileNumber%
Close #Temp2FileNumber%
ierror = True
Exit Function

MiscDeleteLinesNoFromFile:
msg$ = "File to copy from (" & from_name$ & ") was not found and could not be copied"
MsgBox Error$, vbOKOnly + vbCritical, "MiscDeleteLines"
ierror = True
Exit Function

End Function

Sub MiscGetFilenameBasis(arraysize As Integer, filenamearray() As String, tfilename As String)
' Determine the filename that is common to all the filenames in the filename array

ierror = False
On Error GoTo MiscMiscGetFilenameBasisError

Dim i As Integer, j As Integer, k As Integer
Dim extension As String

' Check that filename array is more than a single filename
If arraysize% < 2 Then
tfilename$ = filenamearray$(1)
Exit Sub
End If

' Check that file is more than extension
tfilename$ = MiscGetFileNameNoExtension$(filenamearray$(1))
k% = Len(tfilename$)
If k% < 1 Then
tfilename$ = filenamearray$(1)
Exit Sub
End If

' Save extension
extension$ = MiscGetFileNameExtensionOnly$(filenamearray$(1))

' Loop and add characters that are the same in all filenames
tfilename$ = vbNullString
For j% = 1 To k%
For i% = 2 To arraysize%
If Mid$(filenamearray$(1), j%, 1) <> Mid$(filenamearray$(i%), j%, 1) Then GoTo MiscGetFilenameBasisComplete
Next i%
tfilename$ = tfilename$ & Mid$(filenamearray$(1), j%, 1)
Next j%

' Check filename and if no common basis, just return null string
MiscGetFilenameBasisComplete:
If Len(tfilename$) < 1 Then
tfilename$ = vbNullString
Exit Sub
End If

' Return common basis
tfilename$ = tfilename$ & extension$
Exit Sub

' Errors
MiscMiscGetFilenameBasisError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscMiscGetFilenameBasis"
ierror = True
Exit Sub

End Sub

Function MiscFolderMoveOrCopy(mode As Integer, method As Integer, sSource As String, sDestination As String) As Long
' Uses Window SH function to move or copy a folder
' mode = 0 copy, mode = 1 move
' method = 0 silent, method = 1 confirm

ierror = False
On Error GoTo MiscFolderMoveOrCopyError

Dim FOF_FLAGS As Long
Dim SHFileOp As SHFILEOPSTRUCT
   
' Terminate the folder string with a pair of nulls
sSource = sSource & vbNullChar & vbNullChar

' Determine if copy or move
If mode% = 0 Then
FO_FUNC& = FO_COPY&
End If

If mode% = 1 Then
FO_FUNC& = FO_MOVE&
End If
  
' Determine the options selected
FOF_FLAGS& = 0
If method% = 0 Then
FOF_FLAGS& = FOF_FLAGS& Or FOF_SILENT
FOF_FLAGS& = FOF_FLAGS& Or FOF_NOCONFIRMATION
End If

If method% = 1 Then
FOF_FLAGS& = FOF_FLAGS& Or FOF_SIMPLEPROGRESS&
End If
  
' Set up the options
With SHFileOp
      .wFunc = FO_FUNC&
      .pFrom = sSource
      .pTo = sDestination
      .fFlags = FOF_FLAGS&
End With
  
' And perform the chosen copy or move operation
MiscFolderMoveOrCopy = SHFileOperation(SHFileOp)

Exit Function

' Errors
MiscFolderMoveOrCopyError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscFolderMoveOrCopy"
ierror = True
Exit Function

End Function
