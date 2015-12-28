Attribute VB_Name = "CodeDIRECTORY"
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit
  
Public Const MAX_PATH As Long = 260
Public Const INVALID_HANDLE_VALUE As Long = -1
Public Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED As Long = &H800
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
Public Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Public Const FILE_ATTRIBUTE_READONLY As Long = &H1
Public Const FILE_ATTRIBUTE_TEMPORARY As Long = &H100
Public Const FILE_ATTRIBUTE_FLAGS As Long = FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_HIDDEN Or _
                                              FILE_ATTRIBUTE_NORMAL Or FILE_ATTRIBUTE_READONLY

Public Const DRIVE_UNKNOWNTYPE As Long = 1
Public Const DRIVE_REMOVABLE As Long = 2
Public Const DRIVE_FIXED As Long = 3
Public Const DRIVE_REMOTE As Long = 4
Public Const DRIVE_CDROM As Long = 5
Public Const DRIVE_RAMDISK As Long = 6

Public Type FILETIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
dwFileAttributes As Long
ftCreationTime As FILETIME
ftLastAccessTime As FILETIME
ftLastWriteTime As FILETIME
nFileSizeHigh As Long
nFileSizeLow As Long
dwReserved0 As Long
dwReserved1 As Long
cFileName As String * MAX_PATH
cAlternate As String * 14
End Type

' Custom-made UDT for searching
Public Type FILE_PARAMS
bRecurse As Boolean    ' set True to perform a recursive search
bFound As Boolean      ' set only with SearchTreeForFile methods
sFileRoot As String    ' search starting point, ie c:\, c:\winnt\
sFileNameExt As String ' filename/filespec to locate, ie *.dll, notepad.exe
sResult As String      ' path to file. Set only with SearchTreeForFile methods
nFileCount As Long     ' total file count matching filespec. Set in FindXXX only
nFileSize As Double    ' total file size matching filespec. Set in FindXXX only
End Type

Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
     
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
     
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
    (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" _
    (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
        
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" _
    (ByVal nDrive As String) As Long

Public Sub DirectorySearch(textension As String, tpath As String, tRecurse As Integer, nCount As Long, sAllFiles() As String)
' Directory search

ierror = False
On Error GoTo DirectorySearchError

Dim FP As FILE_PARAMS
  
' Load search parameters
With FP
.sFileNameExt = textension$
.sFileRoot = tpath$
.bRecurse = tRecurse%
End With
         
' Dim an array large enough to hold all the returned values
ReDim sAllFiles(1 To 1000000) As String

' Search for files
nCount& = 0
Screen.MousePointer = vbHourglass
Call DirectorySearchForFilesArray(FP, nCount&, sAllFiles$())
Screen.MousePointer = vbDefault

' Strip off the unused allocated array members
If nCount& > 0 Then ReDim Preserve sAllFiles(1 To nCount&)

Exit Sub

' Errors
DirectorySearchError:
MsgBox Error$, vbOKOnly + vbCritical, "DirectorySearch"
ierror = True
Exit Sub
     
End Sub


Public Sub DirectorySearchForFilesArray(FP As FILE_PARAMS, nCount As Long, sAllFiles() As String)
' This routine is primarily interested in the directories, so the file type must be *.*

ierror = False
On Error GoTo DirectorySearchForFilesArrayError
   
Dim WFD As WIN32_FIND_DATA
Dim hFile As Long
Dim sPath As String
Dim sRoot As String
Dim sTmp As String
        
sRoot$ = DirectoryQualifyPath$(FP.sFileRoot)
sPath$ = sRoot & "*.*"
     
' Obtain handle to the first match
hFile& = FindFirstFile(sPath$, WFD)
     
' If valid ...
If hFile& <> INVALID_HANDLE_VALUE Then
     
' DirectoryGetFileInformation function returns the number of files
' matching the filespec (FP.sFileNameExt) in the passed folder
Call DirectoryGetFileInformation(FP, nCount&, sAllFiles$())

Do
        
' If the returned item is a folder...
If FP.bRecurse And (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
              
' Remove trailing nulls
sTmp = DirectoryTrimNull$(WFD.cFileName$)
              
' And if the folder is not the default self and parent folders...
If sTmp$ <> "." And sTmp$ <> ".." Then
              
' Get the file
FP.sFileRoot$ = sRoot$ & sTmp$
     
' This next If..Then just prevents adding extra lines and unneeded paths
' to the array when a file search is performed for a specific file type.
If FP.sFileNameExt$ = "*.*" Then
                 
' Depending on the purpose, you may want to exclude the next 4 optional lines.
' The first two lines adds a blank entry to the array as a separator between new
' directories in the output file. The last two add the directory name alone, before
' listing the files underneath. These four lines can be optionally commented out).
' Obviously, these extra entries will skew the actual file counts.
'nCount& = nCount& + 1
'sAllFiles$(nCount&) = ""
'nCount& = nCount& + 1
'sAllFiles$(nCount&) = FP.sFileRoot$
End If
                 
' Call again
Call DirectorySearchForFilesArray(FP, nCount&, sAllFiles$())
End If
                 
End If
           
' Continue looping until FindNextFile returns 0 (no more matches)
Loop While FindNextFile(hFile&, WFD)
        
' Close the find handle
hFile& = FindClose(hFile&)
     
End If
     
Exit Sub

' Errors
DirectorySearchForFilesArrayError:
MsgBox Error$, vbOKOnly + vbCritical, "DirectorySearchForFilesArray"
ierror = True
Exit Sub
     
End Sub

Private Function DirectoryQualifyPath(sPath As String) As String
' Assures that a passed path ends in a slash

ierror = False
On Error GoTo DirectoryQualifyPathError

If Right$(sPath, 1) <> "\" Then
DirectoryQualifyPath$ = sPath$ & "\"
Else
DirectoryQualifyPath$ = sPath$
End If

Exit Function

' Errors
DirectoryQualifyPathError:
MsgBox Error$, vbOKOnly + vbCritical, "DirectoryQualifyPath"
ierror = True
Exit Function

End Function

Private Sub DirectoryGetFileInformation(FP As FILE_PARAMS, nCount As Long, sAllFiles() As String)
' Gets file info for a folder

ierror = False
On Error GoTo DirectoryGetFileInformationError

' Local working variables
Dim WFD As WIN32_FIND_DATA
Dim hFile As Long
Dim sPath As String
Dim sRoot As String
Dim sTmp As String

' FP.sFileRoot (assigned to sRoot) contains the path to search
sRoot$ = DirectoryQualifyPath(FP.sFileRoot$)

' FP.sFileNameExt (assigned to sPath) contains the full path and filespec
sPath$ = sRoot$ & FP.sFileNameExt$

' Obtain handle to the first filespec match
hFile& = FindFirstFile(sPath$, WFD)

' If valid ...
If hFile& <> INVALID_HANDLE_VALUE Then

Do

' Even though this routine may use a filespec, *.* is still valid and will cause the search
' to return folders as well as files, so a check against folders is still required.
If Not (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then

' Remove trailing nulls
sTmp$ = DirectoryTrimNull(WFD.cFileName$)

' Increment count and add to the array
nCount& = nCount& + 1
sAllFiles$(nCount&) = sRoot$ & sTmp$

End If

Loop While FindNextFile(hFile&, WFD)

' Close the handle
hFile& = FindClose(hFile&)

End If

Exit Sub

' Errors
DirectoryGetFileInformationError:
MsgBox Error$, vbOKOnly + vbCritical, "DirectoryGetFileInformation"
ierror = True
Exit Sub

End Sub

Private Function DirectoryTrimNull(startstr As String) As String
' Returns the string up to the first null, if present, or the passed string

ierror = False
On Error GoTo DirectoryTrimNullError

Dim pos As Integer
     
pos = InStr(startstr, vbNullChar)
     
If pos Then
  DirectoryTrimNull = Left$(startstr, pos - 1)
  Exit Function
End If
    
DirectoryTrimNull = startstr
    
Exit Function

' Errors
DirectoryTrimNullError:
MsgBox Error$, vbOKOnly + vbCritical, "DirectoryTrimNull"
ierror = True
Exit Function

End Function


