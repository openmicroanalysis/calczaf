Attribute VB_Name = "CodeMiscSystem"
' (c) Copyright 1995-2025 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Private Declare Function GetUserDefaultLCID% Lib "kernel32" ()
 
'Public Const LOCALE_ICENTURY& = &H24
'Public Const LOCALE_ICOUNTRY& = &H5
'Public Const LOCALE_ICURRDIGITS& = &H19
'Public Const LOCALE_ICURRENCY& = &H1B
'Public Const LOCALE_IDATE& = &H21
'Public Const LOCALE_IDAYLZERO& = &H26
'Public Const LOCALE_IDEFAULTCODEPAGE& = &HB
'Public Const LOCALE_IDEFAULTCOUNTRY& = &HA
'Public Const LOCALE_IDEFAULTLANGUAGE& = &H9
'Public Const LOCALE_IDIGITS& = &H11
'Public Const LOCALE_IINTLCURRDIGITS& = &H1A
'Public Const LOCALE_ILANGUAGE& = &H1
'Public Const LOCALE_ILDATE& = &H22
'Public Const LOCALE_ILZERO& = &H12
'Public Const LOCALE_IMEASURE& = &HD
'Public Const LOCALE_IMONLZERO& = &H27
'Public Const LOCALE_INEGCURR& = &H1C
'Public Const LOCALE_INEGSEPBYSPACE& = &H57
'Public Const LOCALE_INEGSIGNPOSN& = &H53
'Public Const LOCALE_INEGSYMPRECEDES& = &H56
'Public Const LOCALE_IPOSSEPBYSPACE& = &H55
'Public Const LOCALE_IPOSSIGNPOSN& = &H52
'Public Const LOCALE_IPOSSYMPRECEDES& = &H54
'Public Const LOCALE_ITIME& = &H23
'Public Const LOCALE_ITLZERO& = &H25
'Public Const LOCALE_NOUSEROVERRIDE& = &H80000000
'Public Const LOCALE_S1159& = &H28
'Public Const LOCALE_S2359& = &H29
'Public Const LOCALE_SABBREVCTRYNAME& = &H7
'Public Const LOCALE_SABBREVDAYNAME1& = &H31
'Public Const LOCALE_SABBREVDAYNAME2& = &H32
'Public Const LOCALE_SABBREVDAYNAME3& = &H33
'Public Const LOCALE_SABBREVDAYNAME4& = &H34
'Public Const LOCALE_SABBREVDAYNAME5& = &H35
'Public Const LOCALE_SABBREVDAYNAME6& = &H36
'Public Const LOCALE_SABBREVDAYNAME7& = &H37
'Public Const LOCALE_SABBREVLANGNAME& = &H3
'Public Const LOCALE_SABBREVMONTHNAME1& = &H44
'Public Const LOCALE_SCOUNTRY& = &H6
'Public Const LOCALE_SCURRENCY& = &H14
'Public Const LOCALE_SDATE& = &H1D
'Public Const LOCALE_SDAYNAME1& = &H2A
'Public Const LOCALE_SDAYNAME2& = &H2B
'Public Const LOCALE_SDAYNAME3& = &H2C
'Public Const LOCALE_SDAYNAME4& = &H2D
'Public Const LOCALE_SDAYNAME5& = &H2E
'Public Const LOCALE_SDAYNAME6& = &H2F
'Public Const LOCALE_SDAYNAME7& = &H30
Public Const LOCALE_SDECIMAL& = &HE
'Public Const LOCALE_SENGCOUNTRY& = &H1002
'Public Const LOCALE_SENGLANGUAGE& = &H1001
'Public Const LOCALE_SGROUPING& = &H10
'Public Const LOCALE_SINTLSYMBOL& = &H15
'Public Const LOCALE_SLANGUAGE& = &H2
'Public Const LOCALE_SLIST& = &HC
'Public Const LOCALE_SLONGDATE& = &H20
'Public Const LOCALE_SMONDECIMALSEP& = &H16
'Public Const LOCALE_SMONGROUPING& = &H18
'Public Const LOCALE_SMONTHNAME1& = &H38
'Public Const LOCALE_SMONTHNAME10& = &H41
'Public Const LOCALE_SMONTHNAME11& = &H42
'Public Const LOCALE_SMONTHNAME12& = &H43
'Public Const LOCALE_SMONTHNAME2& = &H39
'Public Const LOCALE_SMONTHNAME3& = &H3A
'Public Const LOCALE_SMONTHNAME4& = &H3B
'Public Const LOCALE_SMONTHNAME5& = &H3C
'Public Const LOCALE_SMONTHNAME6& = &H3D
'Public Const LOCALE_SMONTHNAME7& = &H3E
'Public Const LOCALE_SMONTHNAME8& = &H3F
'Public Const LOCALE_SMONTHNAME9& = &H40
'Public Const LOCALE_SMONTHOUSANDSEP& = &H17
'Public Const LOCALE_SNATIVECTRYNAME& = &H8
'Public Const LOCALE_SNATIVEDIGITS& = &H13
'Public Const LOCALE_SNATIVELANGNAME& = &H4
'Public Const LOCALE_SNEGATIVESIGN& = &H51
'Public Const LOCALE_SPOSITIVESIGN& = &H50
'Public Const LOCALE_SSHORTDATE& = &H1F
Public Const LOCALE_STHOUSAND& = &HF
'Public Const LOCALE_STIME& = &H1E
'Public Const LOCALE_STIMEFORMAT& = &H1003
 
Private Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer

' OS constants
Global Const OS_VERSION_WIN32S& = 0
Global Const OS_VERSION_WIN95& = 1
Global Const OS_VERSION_WINNT& = 2

Global Const OS_VERSION_NT351& = 3
Global Const OS_VERSION_NT4& = 4
Global Const OS_VERSION_XP& = 5
Global Const OS_VERSION_VISTA& = 6
Global Const OS_VERSION_7& = 7

Private Const S_OK& = &H0                ' success
Private Const S_FALSE& = &H1             ' the Folder is valid, but does not exist
Private Const E_INVALIDARG& = &H80070057 ' invalid CSIDL Value

'Private Const CSIDL_DESKTOP                 As Long = &H0   ' <desktop>
'Private Const CSIDL_INTERNET                As Long = &H1   ' Internet Explorer (icon on desktop)
'Private Const CSIDL_PROGRAMS                As Long = &H2   ' Start Menu\Programs
'Private Const CSIDL_CONTROLS                As Long = &H3   ' My Computer\Control Panel
'Private Const CSIDL_PRINTERS                As Long = &H4   ' My Computer\Printers
'Private Const CSIDL_PERSONAL                As Long = &H5   ' My Documents
'Private Const CSIDL_FAVORITES               As Long = &H6   ' <user name>\Favourites
'Private Const CSIDL_STARTUP                 As Long = &H7   ' Start Menu\Programs\Startup
'Private Const CSIDL_RECENT                  As Long = &H8   ' <user name>\Recent
'Private Const CSIDL_SENDTO                  As Long = &H9   ' <user name>\SendTo
'Private Const CSIDL_BITBUCKET               As Long = &HA   ' <desktop>\Recycle Bin
'Private Const CSIDL_STARTMENU               As Long = &HB   ' <user name>\Start Menu
'Private Const CSIDL_MYDOCUMENTS             As Long = &HC   ' logical "My Documents" desktop icon
'Private Const CSIDL_MYMUSIC                 As Long = &HD   ' "My Music" folder
'Private Const CSIDL_MYVIDEO                 As Long = &HE   ' "My Videos" folder
'Private Const CSIDL_DESKTOPDIRECTORY        As Long = &H10  ' <user name>\Desktop
'Private Const CSIDL_DRIVES                  As Long = &H11  ' My Computer
'Private Const CSIDL_NETWORK                 As Long = &H12  ' Network Neighborhood (My Network Places)
'Private Const CSIDL_NETHOOD                 As Long = &H13  ' <user name>\nethood
'Private Const CSIDL_FONTS                   As Long = &H14  ' windows\fonts
'Private Const CSIDL_TEMPLATES               As Long = &H15  ' templates
'Private Const CSIDL_COMMON_STARTMENU        As Long = &H16  ' All Users\Start Menu
'Private Const CSIDL_COMMON_PROGRAMS         As Long = &H17  ' All Users\Start Menu\Programs
'Private Const CSIDL_COMMON_STARTUP          As Long = &H18  ' All Users\Startup
'Private Const CSIDL_COMMON_DESKTOPDIRECTORY As Long = &H19  ' All Users\Desktop
'Private Const CSIDL_APPDATA                 As Long = &H1A  ' <user name>\Application Data
'Private Const CSIDL_PRINTHOOD               As Long = &H1B  ' <user name>\PrintHood
'Private Const CSIDL_LOCAL_APPDATA           As Long = &H1C  ' <user name>\Local Settings\Application Data (non roaming)
'Private Const CSIDL_ALTSTARTUP              As Long = &H1D  ' non localized startup

' Non localized common startup
'Private Const CSIDL_COMMON_ALTSTARTUP       As Long = &H1E
'Private Const CSIDL_COMMON_FAVORITES        As Long = &H1F
'Private Const CSIDL_INTERNET_CACHE          As Long = &H20
'Private Const CSIDL_COOKIES                 As Long = &H21
'Private Const CSIDL_HISTORY                 As Long = &H22

Private Const CSIDL_COMMON_APPDATA          As Long = &H23  ' All Users\Application Data
'Private Const CSIDL_WINDOWS                 As Long = &H24  ' GetWindowsDirectory()
'Private Const CSIDL_SYSTEM                  As Long = &H25  ' GetSystemDirectory()
Private Const CSIDL_PROGRAM_FILES           As Long = &H26  ' C:\Program Files or C:\Program Files (x86)
'Private Const CSIDL_MYPICTURES              As Long = &H27  ' C:\Program Files\My Pictures
'Private Const CSIDL_PROFILE                 As Long = &H28  ' USERPROFILE
'Private Const CSIDL_SYSTEMX86               As Long = &H29  ' x86 system directory on RISC
'Private Const CSIDL_PROGRAM_FILESX86        As Long = &H2A  ' x86 C:\Program Files on RISC
'Private Const CSIDL_PROGRAM_FILES_COMMON    As Long = &H2B  ' C:\Program Files\Common
'Private Const CSIDL_PROGRAM_FILES_COMMONX86 As Long = &H2C  ' x86 Program Files\Common on RISC
'Private Const CSIDL_COMMON_TEMPLATES        As Long = &H2D  ' All Users\Templates
'Private Const CSIDL_COMMON_DOCUMENTS        As Long = &H2E  ' All Users\Documents
'Private Const CSIDL_COMMON_ADMINTOOLS       As Long = &H2F  ' All Users\Start Menu\Programs\Administrative Tools
'Private Const CSIDL_ADMINTOOLS              As Long = &H30  ' <user name>\Start Menu\Programs\Administrative Tools
'Private Const CSIDL_CONNECTIONS             As Long = &H31  ' Network and Dial-up Connections
'Private Const CSIDL_COMMON_MUSIC            As Long = &H35  ' All Users\My Music
'Private Const CSIDL_COMMON_PICTURES         As Long = &H36  ' All Users\My Pictures
'Private Const CSIDL_COMMON_VIDEO            As Long = &H37  ' All Users\My Video
'Private Const CSIDL_RESOURCES               As Long = &H38  ' Resource Directory
'Private Const CSIDL_RESOURCES_LOCALIZED     As Long = &H39  ' Localized Resource Directory
'Private Const CSIDL_COMMON_OEM_LINKS        As Long = &H3A  ' Links to All Users OEM specific apps
'Private Const CSIDL_CDBURN_AREA             As Long = &H3B  ' USERPROFILE\Local Settings\Application Data\Microsoft\CD Burning
' unused                                     As Long = &H3C

'Private Const CSIDL_COMPUTERSNEARME         As Long = &H3D      ' Computers Near Me (computered from Workgroup membership)
'Private Const CSIDL_FLAG_CREATE             As Long = &H8000&   ' combine with CSIDL_ value to force folder creation in SHGetFolderPath()
'Private Const CSIDL_FLAG_DONT_VERIFY        As Long = &H4000    ' combine with CSIDL_ value to return an unverified folder path
'Private Const CSIDL_FLAG_NO_ALIAS           As Long = &H1000    ' combine with CSIDL_ value to insure non-alias versions of the pidl
'Private Const CSIDL_FLAG_PER_USER_INIT      As Long = &H800     ' combine with CSIDL_ value to indicate per-user init (eg. upgrade)
'Private Const CSIDL_FLAG_MASK               As Long = &HFF00&   ' mask for all possible flag values

Private Const SHGFP_TYPE_CURRENT& = 0       ' current location (use this one)
Private Const SHGFP_TYPE_DEFAULT& = 1       ' default location
Private Const MAX_PATH& = 512

' Use for Vista and up
Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long

' Old functions
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Private Declare Function GetWindowsDirectoryB Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectoryB Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal Path As String, ByVal cbBytes As Long) As Long
Private Const MAX_LENGTH& = 512

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
    OSVSize         As Long
    dwVerMajor      As Long
    dwVerMinor      As Long
    dwBuildNumber   As Long
    PlatformID      As Long
    szCSDVersion    As String * 128
End Type
    
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, ByRef bWow64Process As Boolean) As Long

' Constants used by MiscFindWindowPartial
Global Const FWP_STARTSWITH As Long = 0
Global Const FWP_CONTAINS As Long = 1

Declare Function OSSetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Declare Function OSFindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function OSGetWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
'Declare Function OSSetActiveWindow Lib "user32" Alias "SetForegroundWindow" (ByVal hWnd As Long) As Long
Declare Function OSGetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function OSGetParent Lib "user32" Alias "GetParent" (ByVal hWnd As Long) As Long

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Public Type TypeFileInformation
CompanyName As String
FileDescription As String
FileVersion As String
InternalName As String
LegalCopyright As String
OriginalFileName As String
ProductName As String
ProductVersion As String
End Type

' Constant used by OSGetWindow to find next window
Private Const GW_HWNDNEXT As Long = 2

' API for MiscDirectorySort
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
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

Private Type FileInfo
    Filename As String
    Modified As Currency
End Type

Private Const STANDARD_RIGHTS_REQUIRED              As Long = &HF0000
Private Const SYNCHRONIZE                           As Long = &H100000
Private Const MUTANT_QUERY_STATE                    As Long = &H1
Private Const MUTANT_ALL_ACCESS                     As Long = (STANDARD_RIGHTS_REQUIRED& Or SYNCHRONIZE& Or MUTANT_QUERY_STATE&)
'Private Const SECURITY_DESCRIPTOR_REVISION          As Long = 1
'Private Const DACL_SECURITY_INFORMATION             As Long = 4

Public Declare Function OpenMutex Lib "kernel32.dll" Alias "OpenMutexA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal lpName As String) As Long
Public Declare Function ReleaseMutex Lib "kernel32.dll" (ByVal hMutex As Long) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Public Declare Function CreateMutex Lib "kernel32.dll" Alias "CreateMutexA" (ByVal lpMutexAttributes As Any, ByVal bInitialOwner As Boolean, ByVal lpName As String) As Long
Public Declare Function GetLastError Lib "kernel32.dll" () As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nindex As Long) As Long
Private Const GDC_LOGPIXELSX As Long = 88

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, lRemoteNameLength As Long) As Long

Public Function MiscGetWindowsDPI() As Double
' Returns the current screen DPI, *as set in Windows display settings.*  This obviously has no relationship to physical screen DPI,
' which would need to be calculated on a per-monitor basis using something like EDID data.
'
' For convenience, DPI is returned as a float where 1.0 = 96 DPI (the default Windows setting).  2.0 = 200% DPI scaling, etc.

Dim screenDC As Long, logPixelsX As Double
    
ierror = False
On Error GoTo MiscGetWindowsDPIError
    
' Get device context
screenDC& = GetDC(0)
    
' Retrieve logPixelsX via the API; this will be 96 at 100% DPI scaling
logPixelsX# = GetDeviceCaps(screenDC&, GDC_LOGPIXELSX&)
ReleaseDC 0, screenDC&

' Convert that value into a fractional DPI modified (e.g. 1.0 for 100% scaling, 2.0 for 200% scaling)
If logPixelsX# = 0# Then
    MiscGetWindowsDPI# = 1#
    msg$ = "WARNING!  System DPI could not be retrieved; this system may be running in a VM or on Linux via Wine."
    MsgBox msg$, vbOKOnly + vbExclamation, "MiscGetWindowsDPI"
    ierror = True
    Exit Function
Else
    MiscGetWindowsDPI# = logPixelsX# / 96#
End If

Exit Function

' Errors
MiscGetWindowsDPIError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetWindowsDPI"
ierror = True
Exit Function

End Function

Public Function MiscFindWindowPartial(ByVal TitleContains As String, Optional ByVal method As Variant) As Long
' Purpose:     Find an application based on title
'
' Arguments:   TitleContains     Text in title of application
'              Method            Either FWP_STARTSWITH or FWP_CONTAINS
'
' Returns:     hWnd of window, or 0 if window could not be found
'
ierror = False
On Error GoTo MiscFindWindowPartialError

Dim plngThisHWnd As Long, plngResult As Long, plngMethod As Long
Dim pstrThisTitle As String, pstrTitle As String

If IsMissing(method) Then
      plngMethod = FWP_CONTAINS
Else
      plngMethod = method
End If

pstrTitle = UCase$(TitleContains)

' Assume failure
MiscFindWindowPartial = 0
   
' Find first window and loop through all subsequent windows in master window list
plngThisHWnd = OSFindWindow(vbNullString, vbNullString)
Do Until plngThisHWnd = 0

   ' Make sure this window has no parent
   If OSGetParent(plngThisHWnd) = 0 Then
         
      ' Retrieve caption text from current window.
      pstrThisTitle = Space$(256)
      plngResult = OSGetWindowText(plngThisHWnd, pstrThisTitle, Len(pstrThisTitle) - 1)
      If plngResult Then
         
         ' Clean up return string, preparing for case-insensitive comparison
         pstrThisTitle = UCase$(Left$(pstrThisTitle, plngResult))
         
         ' Use appropriate method to determine if current window's caption either starts with or contains passed string
         Select Case plngMethod
            Case FWP_STARTSWITH
               If InStr(pstrThisTitle, pstrTitle) = 1 Then
                  MiscFindWindowPartial = plngThisHWnd
                  Exit Do
               End If
            Case FWP_CONTAINS
               If InStr(pstrThisTitle, pstrTitle) > 0 Then
                  MiscFindWindowPartial = plngThisHWnd
                  Exit Do
               End If
         End Select
      End If
End If

   ' Get next window in master window list and continue.
   plngThisHWnd = OSGetWindow(plngThisHWnd, GW_HWNDNEXT)
Loop

Exit Function

' Errors
MiscFindWindowPartialError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscFindWindowPartial"
ierror = True
Exit Function

End Function

Public Function MiscConvertTimeToClockString(atime As Variant) As String
' Convert the elapsed time to a clock string (hours:minutes:seconds)

ierror = False
On Error GoTo MiscConvertTimeToClockStringError

Dim astring As String
Dim extrahours As Long

MiscConvertTimeToClockString$ = "00:00:00"

' Check for null value (old versions)
If IsNull(atime) Then Exit Function

' Calculate hours greater than 24
extrahours& = Fix(atime) * HOURPERDAY#   ' truncate

' Format string
If Abs(atime) > 4# Then
astring = Format$(extrahours& + Hour(atime), "000") & ":" & Format$(Minute(atime), "00") & ":" & Format$(Second(atime), "00")
ElseIf Abs(atime) > 1# Then
astring = Format$(extrahours& + Hour(atime), "00") & ":" & Format$(Minute(atime), "00") & ":" & Format$(Second(atime), "00")
Else
astring = Format$(Hour(atime), "00") & ":" & Format$(Minute(atime), "00") & ":" & Format$(Second(atime), "00")
End If

' Add negative sign if negative
If atime < 0# Then astring$ = "-" & astring$

MiscConvertTimeToClockString$ = astring$
Exit Function

' Errors
MiscConvertTimeToClockStringError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscConvertTimeToClockString"
ierror = True
Exit Function

End Function

Public Sub MiscGetFileVersionData(ByRef tfilename As String, tmajor As Long, tminor As Long, trevision As Long, tRelease As Long)
' Returns the version number(s) of the passed file (assuming it has a FileInfo structure)

ierror = False
On Error GoTo MiscGetFileVersionDataError

Dim astring As String, bstring As String
Dim tFileInformation As TypeFileInformation

' Check if passed filename is in the path
If Dir$(tfilename$) = vbNullString And Dir$(SystemPath$ & "\" & tfilename$) = vbNullString Then GoTo MiscGetFileVersionDataFileNotFound

' Get the File Information structure
Call MiscGetFileInformation(tfilename$, tFileInformation)
If ierror Then Exit Sub

' Convert the file version string into integers
astring$ = tFileInformation.FileVersion$
Call MiscParseStringToStringA(astring$, ".", bstring$)
tmajor& = Val(bstring$)
Call MiscParseStringToStringA(astring$, ".", bstring$)
tminor& = Val(bstring$)
Call MiscParseStringToStringA(astring$, ".", bstring$)
trevision& = Val(bstring$)
Call MiscParseStringToStringA(astring$, ".", bstring$)
tRelease& = Val(bstring$)

Exit Sub

' Errors
MiscGetFileVersionDataError:
MsgBox Error$ & ", getting version info for " & tfilename$ & " (file must be in the application or system folder)", vbOKOnly + vbCritical, "MiscGetFileVersionData"
ierror = True
Exit Sub

MiscGetFileVersionDataFileNotFound:
msg$ = "The specified file " & tfilename$ & " was not found in the current directory or system path. Please make sure it exists and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "MiscGetFileVersionData"
ierror = True
Exit Sub

End Sub

Public Sub MiscGetFileInformation(ByRef tfilename As String, tFileInfo As TypeFileInformation)
' Returns the file information structure of the passed file (assuming it has a FileInfo structure)

ierror = False
On Error GoTo MiscGetFileInformationError

Dim lBufferLen As Long, lDummy As Long
Dim sBuffer() As Byte
Dim lVerPointer As Long, lRet As Long
Dim Lang_Charset_String As String
Dim HexNumber As Long
Dim i As Integer
Dim strTemp As String

'Clear the Buffer tFileInfo
tFileInfo.CompanyName = vbNullString
tFileInfo.FileDescription = vbNullString
tFileInfo.FileVersion = vbNullString
tFileInfo.InternalName = vbNullString
tFileInfo.LegalCopyright = vbNullString
tFileInfo.OriginalFileName = vbNullString
tFileInfo.ProductName = vbNullString
tFileInfo.ProductVersion = vbNullString

lBufferLen = GetFileVersionInfoSize(tfilename$, lDummy)
If lBufferLen < 1 Then
MsgBox "Error calling OS function GetFileVersionInfoSize", vbOKOnly + vbExclamation, "MiscGetFileInformation"
ierror = True
Exit Sub
End If

ReDim sBuffer(lBufferLen)
lRet = GetFileVersionInfo(tfilename$, 0&, lBufferLen, sBuffer(0))
If lRet = 0 Then
MsgBox "Error calling OS function GetFileVersionInfo", vbOKOnly + vbExclamation, "MiscGetFileInformation"
ierror = True
Exit Sub
End If

lRet = VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", lVerPointer, lBufferLen)
If lRet = 0 Then
MsgBox "Error calling OS function VerQueryValue", vbOKOnly + vbExclamation, "MiscGetFileInformation"
ierror = True
Exit Sub
End If

Dim bytebuffer(255) As Byte
MoveMemory bytebuffer(0), lVerPointer, lBufferLen
HexNumber = bytebuffer(2) + bytebuffer(3) * &H100 + bytebuffer(0) * &H10000 + bytebuffer(1) * &H1000000
Lang_Charset_String = Hex(HexNumber)

'Pull it all apart:
'04------= SUBLANG_ENGLISH_USA
'--09----= LANG_ENGLISH
' ----04E4 = 1252 = Codepage for Windows:Multilingual

Do While Len(Lang_Charset_String) < 8
Lang_Charset_String = "0" & Lang_Charset_String
Loop

Dim strVersionInfo(7) As String
strVersionInfo(0) = "CompanyName"
strVersionInfo(1) = "FileDescription"
strVersionInfo(2) = "FileVersion"
strVersionInfo(3) = "InternalName"
strVersionInfo(4) = "LegalCopyright"
strVersionInfo(5) = "OriginalFileName"
strVersionInfo(6) = "ProductName"
strVersionInfo(7) = "ProductVersion"

Dim buffer As String
For i = 0 To 7
buffer = String(255, 0)
strTemp = "\StringFileInfo\" & Lang_Charset_String & "\" & strVersionInfo(i)
lRet = VerQueryValue(sBuffer(0), strTemp, lVerPointer, lBufferLen)
If lRet = 0 Then
ierror = True
Exit Sub
End If

lstrcpy buffer, lVerPointer
buffer = Mid$(buffer, 1, InStr(buffer, vbNullChar) - 1)
Select Case i
Case 0
tFileInfo.CompanyName = buffer
Case 1
tFileInfo.FileDescription = buffer
Case 2
tFileInfo.FileVersion = buffer
Case 3
tFileInfo.InternalName = buffer
Case 4
tFileInfo.LegalCopyright = buffer
Case 5
tFileInfo.OriginalFileName = buffer
Case 6
tFileInfo.ProductName = buffer
Case 7
tFileInfo.ProductVersion = buffer
End Select
Next i

Exit Sub

' Errors
MiscGetFileInformationError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetFileInformation"
ierror = True
Exit Sub

End Sub

Function MiscSystemGetOSVersionNumber() As Long
' Returns the version number of the operating system

'Global Const OS_VERSION_WIN32S& = 0
'Global Const OS_VERSION_WIN95& = 1
'Global Const OS_VERSION_WINNT& = 2

'Global Const OS_VERSION_NT351& = 3
'Global Const OS_VERSION_NT4& = 4
'Global Const OS_VERSION_XP& = 5
'Global Const OS_VERSION_VISTA& = 6
'Global Const OS_VERSION_7& = 7

ierror = False
On Error GoTo MiscSystemGetOSVersionNumberError

Dim GetWindowsVersion As String
Dim tVersion As Long

Dim osv As OSVERSIONINFO

osv.OSVSize = Len(osv)

If GetVersionEx(osv) = 1 Then
    Select Case osv.PlatformID
        Case VER_PLATFORM_WIN32s
            GetWindowsVersion$ = "Win32s on Windows 3.1": tVersion& = OS_VERSION_WIN32S&
        Case VER_PLATFORM_WIN32_NT
            GetWindowsVersion$ = "Windows NT": tVersion& = OS_VERSION_WINNT&
            
            Select Case osv.dwVerMajor
                Case 3
                    GetWindowsVersion$ = "Windows NT 3.5": tVersion& = OS_VERSION_NT351&
                Case 4
                    GetWindowsVersion$ = "Windows NT 4.0": tVersion& = OS_VERSION_NT4&
                Case 5
                    Select Case osv.dwVerMinor
                        Case 0
                            GetWindowsVersion$ = "Windows 2000": tVersion& = OS_VERSION_XP&
                        Case 1
                            GetWindowsVersion$ = "Windows XP": tVersion& = OS_VERSION_XP&
                        Case 2
                            GetWindowsVersion$ = "Windows Server 2003": tVersion& = OS_VERSION_XP&
                    End Select
                    
                Case 6
                    Select Case osv.dwVerMinor
                        Case 0
                            GetWindowsVersion$ = "Windows Vista/Server 2008": tVersion& = OS_VERSION_VISTA&
                        Case 1
                            GetWindowsVersion$ = "Windows 7/Server 2008 R2": tVersion& = OS_VERSION_7&
                    End Select
                End Select
                    
        Case VER_PLATFORM_WIN32_WINDOWS:
            Select Case osv.dwVerMinor
                Case 0
                    GetWindowsVersion$ = "Windows 95": tVersion& = OS_VERSION_WIN95&
                Case 90
                    GetWindowsVersion$ = "Windows Me": tVersion& = OS_VERSION_WIN95&
                Case Else
                    GetWindowsVersion$ = "Windows 98": tVersion& = OS_VERSION_WIN95&
                End Select
            End Select
        Else
            GetWindowsVersion$ = "Unable to identify your version of Windows": tVersion& = 0
        End If
        
If DebugMode Then
Call IOWriteLog(vbNullString$)
Call IOWriteLog("MiscSystemGetOSVersionNumber: " & GetWindowsVersion$ & ", version= " & tVersion&)
End If

MiscSystemGetOSVersionNumber& = tVersion&
Exit Function

' Errors
MiscSystemGetOSVersionNumberError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSystemGetOSVersionNumber"
ierror = True
Exit Function

End Function

Public Function MiscSystemIsHost64Bit() As Boolean
' Determines if the system is 64 bit or not

ierror = False
On Error GoTo MiscSystemIsHost64BitError

Dim handle As Long
Dim is64Bit As Boolean
 
' Assume initially that this is not a WOW64 process
is64Bit = False
 
' Then try to prove that wrong by attempting to load the IsWow64Process function dynamically
handle& = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
 
' The function exists, so call it
If handle& <> 0 Then
    IsWow64Process GetCurrentProcess(), is64Bit
End If
 
' Return the value
MiscSystemIsHost64Bit = is64Bit

Exit Function

' Errors
MiscSystemIsHost64BitError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSystemIsHost64Bit"
ierror = True
Exit Function

End Function

Public Function MiscSystemGetWindowsDirectory() As String
' Returns the Windows folder path

ierror = False
On Error GoTo MiscSystemGetWindowsDirectoryError

Dim s As String
Dim c As Long

s$ = String$(MAX_LENGTH&, 0)
c& = GetWindowsDirectoryB(s$, MAX_LENGTH&)

If c& > 0 Then
    If c& > Len(s$) Then
        s$ = Space$(c& + 1)
        c& = GetWindowsDirectoryB(s$, MAX_LENGTH&)
       End If
End If
   
MiscSystemGetWindowsDirectory = IIf(c& > 0, Left$(s$, c&), "")
Exit Function
   
' Errors
MiscSystemGetWindowsDirectoryError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSystemGetWindowsDirectory"
ierror = True
Exit Function

End Function

Public Function MiscSystemGetWindowsSystemDirectory() As String
' Returns the Windows system folder path

ierror = False
On Error GoTo MiscSystemGetWindowsSystemDirectoryError

Dim s As String
Dim c As Long

s$ = String$(MAX_LENGTH&, 0)
c& = GetSystemDirectoryB(s$, MAX_LENGTH&)

If c& > 0 Then
    If c& > Len(s$) Then
        s$ = Space$(c& + 1)
        c& = GetSystemDirectoryB(s$, MAX_LENGTH&)
    End If
End If
   
MiscSystemGetWindowsSystemDirectory = IIf(c& > 0, Left$(s$, c&), "")
Exit Function

' Errors
MiscSystemGetWindowsSystemDirectoryError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSystemGetWindowsSystemDirectory"
ierror = True
Exit Function

End Function

Private Function MiscSystemGetTMPPath() As String
' Return the Windows temp folder. It takes two parameters:
'   the length of a fixed- length or pre-initialized string that will contain the path name
'   and the string itself
' GetTempPath returns the length of the path name measured in bytes, or 0 if an error occurs.
' If the return value is greater than the buffer size you specified, then no path information was written to the string.

ierror = False
On Error GoTo MiscSystemGetTMPPathError

Dim sFolder As String   ' name of the folder
Dim lRet As Long        ' return Value

Const MAX_PATH& = 512

sFolder$ = String$(MAX_PATH&, 0)
lRet& = GetTempPath(MAX_PATH&, sFolder)

If lRet& <> 0 Then
    MiscSystemGetTMPPath = Left(sFolder$, InStr(sFolder$, vbNullChar) - 1)
Else
    MiscSystemGetTMPPath = vbNullString
End If

Exit Function

' Errors
MiscSystemGetTMPPathError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSystemGetTMPPath"
ierror = True
Exit Function

End Function

Public Function MiscSystemCreateTempFile(sPrefix As String) As String
' Uses GetTempPath to return fully qualified temp file. Use Kill to delete the file when done.
' The function takes four parameters:
'   the string containing the path for the file
'   a string containing a prefix used to start the unique file name
'   a unique number to construct the temp name
'   a fixed-length or pre-initialized string used to return the fully qualified file name.
' Both the path and prefix strings are required and cannot be empty.
' The unique number can be 0 (NULL), in which case GetTempFileName creates a unique number based on the current system time.

ierror = False
On Error GoTo MiscSystemCreateTempFileError

Dim sTmpPath As String * 512
Dim sTmpName As String * 1024
Dim nRet As Long

nRet& = GetTempPath(512, sTmpPath$)

If (nRet& > 0 And nRet& < 512) Then
nRet& = GetTempFileName(sTmpPath$, sPrefix$, 0, sTmpName$)
            
    If nRet& <> 0 Then
        MiscSystemCreateTempFile$ = Left$(sTmpName$, InStr(sTmpName$, vbNullChar) - 1)
    Else
        MiscSystemCreateTempFile$ = vbNullString
    End If
End If
      
Exit Function

' Errors
MiscSystemCreateTempFileError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSystemCreateTempFile"
ierror = True
Exit Function

End Function

Sub MiscSystemGetFolderPath(mode As Integer, sPath As String)
' Returns a list of all valid system folders
' mode% = 0  get program files path (CSIDL_PROGRAM_FILES)
' mode% = 1  get program data path (CSIDL_COMMON_APPDATA)

ierror = False
On Error GoTo MiscSystemGetFolderPathError

Dim RetVal As Long

' Fill our string buffer
sPath = String(MAX_PATH&, 0)

'RetVal& = SHGetFolderPath(0, CSIDL_LOCAL_APPDATA& Or CSIDL_FLAG_CREATE&, 0, SHGFP_TYPE_CURRENT&, sPath$)
If mode% = 0 Then RetVal& = SHGetFolderPath(0, CSIDL_PROGRAM_FILES&, 0, SHGFP_TYPE_CURRENT&, sPath$)        ' get Programs Files folder
If mode% = 1 Then RetVal& = SHGetFolderPath(0, CSIDL_COMMON_APPDATA&, 0, SHGFP_TYPE_CURRENT&, sPath$)        ' get ProgramData folder

Select Case RetVal&
    Case S_OK&
        ' We retrieved the folder successfully, return the string up to the first null character
        sPath$ = Left(sPath$, InStr(1, sPath$, vbNullChar) - 1)
    Case S_FALSE&
        GoTo MiscSystemGetFolderPathFolderNotFound
    Case E_INVALIDARG&
        GoTo MiscSystemGetFolderPathInvalidPath
End Select

Exit Sub

' Errors
MiscSystemGetFolderPathError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSystemGetFolderPath"
ierror = True
Exit Sub

' The CSIDL in nFolder is valid, but the folder does not exist. Use CSIDL_FLAG_CREATE to have it created automatically
MiscSystemGetFolderPathFolderNotFound:
msg$ = "The folder does not exist"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscSystemGetFolderPath"
ierror = True
Exit Sub

MiscSystemGetFolderPathInvalidPath:
msg$ = "An invalid folder ID was specified"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscSystemGetFolderPath"
ierror = True
Exit Sub

End Sub

Function MiscSystemGetLanguage() As String
' Returns the current language

ierror = False
On Error GoTo MiscSystemGetLanguageError

    Select Case GetSystemDefaultLangID()
    Case &H0: MiscSystemGetLanguage = "Language Neutral"
    Case &H400: MiscSystemGetLanguage = "Process Default Language"
    Case &H401: MiscSystemGetLanguage = "Arabic (Saudi Arabia)"
    Case &H801: MiscSystemGetLanguage = "Arabic(Iraq)"
    Case &HC01: MiscSystemGetLanguage = "Arabic(Egypt)"
    Case &H1001: MiscSystemGetLanguage = "Arabic(Libya)"
    Case &H1401: MiscSystemGetLanguage = "Arabic(Algeria)"
    Case &H1801: MiscSystemGetLanguage = "Arabic(Morocco)"
    Case &H1C01: MiscSystemGetLanguage = "Arabic(Tunisia)"
    Case &H2001: MiscSystemGetLanguage = "Arabic(Oman)"
    Case &H2401: MiscSystemGetLanguage = "Arabic(Yemen)"
    Case &H2801: MiscSystemGetLanguage = "Arabic(Syria)"
    Case &H2C01: MiscSystemGetLanguage = "Arabic(Jordan)"
    Case &H3001: MiscSystemGetLanguage = "Arabic(Lebanon)"
    Case &H3401: MiscSystemGetLanguage = "Arabic(Kuwait)"
    Case &H3801: MiscSystemGetLanguage = "Arabic (U.A.E.)"
    Case &H3C01: MiscSystemGetLanguage = "Arabic(Bahrain)"
    Case &H4001: MiscSystemGetLanguage = "Arabic(Qatar)"
    Case &H402: MiscSystemGetLanguage = "Bulgarian"
    Case &H403: MiscSystemGetLanguage = "Catalan"
    Case &H404: MiscSystemGetLanguage = "Chinese (Taiwan Region)"
    Case &H804: MiscSystemGetLanguage = "Chinese(PRC)"
    Case &HC04: MiscSystemGetLanguage = "Chinese (Hong Kong SAR, PRC)"
    Case &H1004: MiscSystemGetLanguage = "Chinese(Singapore)"
    Case &H405: MiscSystemGetLanguage = "Czech"
    Case &H406: MiscSystemGetLanguage = "Danish"
    Case &H407: MiscSystemGetLanguage = "German(Standard)"
    Case &H807: MiscSystemGetLanguage = "German(Swiss)"
    Case &HC07: MiscSystemGetLanguage = "German(Austrian)"
    Case &H1007: MiscSystemGetLanguage = "German(Luxembourg)"
    Case &H1407: MiscSystemGetLanguage = "German(Liechtenstein)"
    Case &H408: MiscSystemGetLanguage = "Greek"
    Case &H409: MiscSystemGetLanguage = "English (United States)"
    Case &H809: MiscSystemGetLanguage = "English (United Kingdom)"
    Case &HC09: MiscSystemGetLanguage = "English(Australian)"
    Case &H1009: MiscSystemGetLanguage = "English(Canadian)"
    Case &H1409: MiscSystemGetLanguage = "English(New Zealand)"
    Case &H1809: MiscSystemGetLanguage = "English(Ireland)"
    Case &H1C09: MiscSystemGetLanguage = "English (South Africa)"
    Case &H2009: MiscSystemGetLanguage = "English(Jamaica)"
    Case &H2409: MiscSystemGetLanguage = "English(Caribbean)"
    Case &H2809: MiscSystemGetLanguage = "English(Belize)"
    Case &H2C09: MiscSystemGetLanguage = "English(Trinidad)"
    Case &H40A: MiscSystemGetLanguage = "Spanish (Traditional Sort)"
    Case &H80A: MiscSystemGetLanguage = "Spanish(Mexican)"
    Case &HC0A: MiscSystemGetLanguage = "Spanish (Modern Sort)"
    Case &H100A: MiscSystemGetLanguage = "Spanish(Guatemala)"
    Case &H140A: MiscSystemGetLanguage = "Spanish (Costa Rica)"
    Case &H180A: MiscSystemGetLanguage = "Spanish(Panama)"
    Case &H1C0A: MiscSystemGetLanguage = "Spanish (Dominican Republic)"
    Case &H200A: MiscSystemGetLanguage = "Spanish(Venezuela)"
    Case &H240A: MiscSystemGetLanguage = "Spanish(Colombia)"
    Case &H280A: MiscSystemGetLanguage = "Spanish(Peru)"
    Case &H2C0A: MiscSystemGetLanguage = "Spanish(Argentina)"
    Case &H300A: MiscSystemGetLanguage = "Spanish(Ecuador)"
    Case &H340A: MiscSystemGetLanguage = "Spanish(Chile)"
    Case &H380A: MiscSystemGetLanguage = "Spanish(Uruguay)"
    Case &H3C0A: MiscSystemGetLanguage = "Spanish(Paraguay)"
    Case &H400A: MiscSystemGetLanguage = "Spanish(Bolivia)"
    Case &H440A: MiscSystemGetLanguage = "Spanish (El Salvador)"
    Case &H480A: MiscSystemGetLanguage = "Spanish(Honduras)"
    Case &H4C0A: MiscSystemGetLanguage = "Spanish(Nicaragua)"
    Case &H500A: MiscSystemGetLanguage = "Spanish (Puerto Rico)"
    Case &H40B: MiscSystemGetLanguage = "Finnish"
    Case &H40C: MiscSystemGetLanguage = "French(Standard)"
    Case &H80C: MiscSystemGetLanguage = "French(Belgian)"
    Case &HC0C: MiscSystemGetLanguage = "French(Canadian)"
    Case &H100C: MiscSystemGetLanguage = "French(Swiss)"
    Case &H140C: MiscSystemGetLanguage = "French(Luxembourg)"
    Case &H40D: MiscSystemGetLanguage = "Hebrew"
    Case &H40E: MiscSystemGetLanguage = "Hungarian"
    Case &H40F: MiscSystemGetLanguage = "Icelandic"
    Case &H410: MiscSystemGetLanguage = "Italian(Standard)"
    Case &H810: MiscSystemGetLanguage = "Italian(Swiss)"
    Case &H411: MiscSystemGetLanguage = "Japanese"
    Case &H412: MiscSystemGetLanguage = "Korean"
    Case &H812: MiscSystemGetLanguage = "Korean(Johab)"
    Case &H413: MiscSystemGetLanguage = "Dutch(Standard)"
    Case &H813: MiscSystemGetLanguage = "Dutch(Belgian)"
    Case &H414: MiscSystemGetLanguage = "Norwegian(Bokmal)"
    Case &H814: MiscSystemGetLanguage = "Norwegian(Nynorsk)"
    Case &H415: MiscSystemGetLanguage = "Polish"
    Case &H416: MiscSystemGetLanguage = "Portuguese(Brazilian)"
    Case &H816: MiscSystemGetLanguage = "Portuguese(Standard)"
    Case &H418: MiscSystemGetLanguage = "Romanian"
    Case &H419: MiscSystemGetLanguage = "Russian"
    Case &H41A: MiscSystemGetLanguage = "Croatian"
    Case &H81A: MiscSystemGetLanguage = "Serbian(Latin)"
    Case &HC1A: MiscSystemGetLanguage = "Serbian(Cyrillic)"
    Case &H41B: MiscSystemGetLanguage = "Slovak"
    Case &H41C: MiscSystemGetLanguage = "Albanian"
    Case &H41D: MiscSystemGetLanguage = "Swedish"
    Case &H81D: MiscSystemGetLanguage = "Swedish(Finland)"
    Case &H41E: MiscSystemGetLanguage = "Thai"
    Case &H41F: MiscSystemGetLanguage = "Turkish"
    Case &H421: MiscSystemGetLanguage = "Indonesian"
    Case &H422: MiscSystemGetLanguage = "Ukrainian"
    Case &H423: MiscSystemGetLanguage = "Belarusian"
    Case &H424: MiscSystemGetLanguage = "Slovenian"
    Case &H425: MiscSystemGetLanguage = "Estonian"
    Case &H426: MiscSystemGetLanguage = "Latvian"
    Case &H427: MiscSystemGetLanguage = "Lithuanian"
    Case &H429: MiscSystemGetLanguage = "Farsi"
    Case &H42A: MiscSystemGetLanguage = "Vietnamese"
    Case &H42D: MiscSystemGetLanguage = "Basque"
    Case &H436: MiscSystemGetLanguage = "Afrikaans"
    Case &H438: MiscSystemGetLanguage = "Faeroese"
    End Select

Exit Function

' Errors
MiscSystemGetLanguageError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSystemGetLanguage"
ierror = True
Exit Function

End Function

Sub MiscSystemGetRegionalSettings(nLocale As Long, sLocale As String)
' Retrieve the regional setting specified by nLocale (constant)
 
ierror = False
On Error GoTo MiscSystemGetRegionalSettingsError

Dim sSymbol As String
Dim iRet1 As Long
Dim iRet2 As Long
Dim lpLCDataVar As String
Dim pos As Integer
Dim tLocale As Long

' Get user Locale ID
tLocale& = GetUserDefaultLCID()
iRet1& = GetLocaleInfo(tLocale&, nLocale&, lpLCDataVar$, 0)
sSymbol$ = String$(iRet1&, 0)
      
iRet2& = GetLocaleInfo(tLocale&, nLocale&, sSymbol$, iRet1&)
pos% = InStr(sSymbol$, vbNullChar)
If pos% > 0 Then
    sLocale$ = Left$(sSymbol$, pos% - 1)
End If
 
Exit Sub

' Errors
MiscSystemGetRegionalSettingsError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSystemGetRegionalSettings"
ierror = True
Exit Sub

End Sub
 
Sub MiscSystemSetRegionalSettings(nLocale As Long, sLocale As String)
' Change the regional setting specified by nLocale (constant) and sLocale (string)
 
ierror = False
On Error GoTo MiscSystemSetRegionalSettingsError
 
Dim iret As Long
Dim tLocale As Long
      
' Get user Locale ID
tLocale& = GetUserDefaultLCID()
iret& = SetLocaleInfo(nLocale&, nLocale&, sLocale$)
     
Exit Sub

' Errors
MiscSystemSetRegionalSettingsError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSystemSetRegionalSettings"
ierror = True
Exit Sub

End Sub

Sub MiscAlwaysOnTop(mode As Integer, tForm As Form)
' This procedure makes a form "always on top" (leave this code in MiscSystem module for declarations)
' mode =  0 do not make topmost
' mode <> 0 make topmost

ierror = False
On Error GoTo MiscAlwaysOnTopError

' Set some constant values (from WIN32API.TXT)
Const conHwndTopmost& = -1
Const conHwndNoTopmost& = -2
Const conSwpNoActivate& = &H10

Dim pleft As Long, ptop As Long
Dim pwidth As Long, pheight As Long

pleft& = tForm.scaleX(tForm.Left, vbTwips, vbPixels)
ptop& = tForm.scaleX(tForm.Top, vbTwips, vbPixels)
pwidth& = tForm.scaleX(tForm.Width, vbTwips, vbPixels)
pheight& = tForm.scaleX(tForm.Height, vbTwips, vbPixels)

' Turn off the TopMost attribute
If mode% = 0 Then
OSSetWindowPos tForm.hWnd, conHwndNoTopmost&, pleft&, ptop&, pwidth&, pheight&, conSwpNoActivate&

' Turn on the TopMost attribute
Else
OSSetWindowPos tForm.hWnd, conHwndTopmost&, pleft&, ptop&, pwidth&, pheight&, conSwpNoActivate&
End If

Exit Sub

' Errors
MiscAlwaysOnTopError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAlwaysOnTop"
ierror = True
Exit Sub

End Sub

Sub MiscDirectorySort(tpath As String, tfilenames() As String, tfiledates() As Variant)
' Returns an array of filenames sorted by date and time
'
' Shell "cmd /c DIR C:\ /O:-D /B > C:\DirList.txt"
' Note that /O:-D reverses the sort order (Newest to Oldest) and /B, which specifies the bare format removes all information (like date and time) but the filename.

On Error GoTo MiscDirectorySortError

Dim hFind As Long, n As Long, i As Long

Dim WFD As WIN32_FIND_DATA
Dim tFiles() As FileInfo
    
' Create file list
hFind& = FindFirstFile(tpath$, WFD)
    Do
        If (WFD.dwFileAttributes And vbDirectory) = 0 Then
            ReDim Preserve tFiles(n&)
            tFiles(n&).Filename = Left$(WFD.cFileName, lstrlen(WFD.cFileName))
            CopyMemory tFiles(n&).Modified, WFD.ftLastWriteTime, 8
            n& = n& + 1
        End If
    Loop While FindNextFile(hFind&, WFD)
FindClose hFind&
    
' Sort file list by date/time
Dim tFile As FileInfo

For n& = 0 To UBound(tFiles)
        For i& = n& + 1 To UBound(tFiles)
            If tFiles(n&).Modified > tFiles(i&).Modified Then
                tFile = tFiles(i&)
                tFiles(i&) = tFiles(n)
                tFiles(n&) = tFile
            End If
        Next
Next n&

' Return date sorted filenames with date and time (cannot use tFiles().Modified due to format issues)
ReDim tfilenames(1 To UBound(tFiles)) As String
ReDim tfiledates(1 To UBound(tFiles)) As Variant

For n& = 1 To UBound(tFiles)
tfilenames$(n&) = tFiles(n&).Filename$
tfiledates(n&) = FileDateTime(MiscGetPathOnly2$(tpath$) & "\" & tfilenames$(n&))
Next n&

Exit Sub

' Errors
MiscDirectorySortError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscDirectorySort"
ierror = True
Exit Sub

End Sub

Function MiscCheckForMutex(appstring As String) As Boolean
' Check for the mutex to see if an app is running or not

ierror = False
On Error GoTo MiscCheckForMutexError

Dim tMutex As Long

' Get mutex for app string
tMutex& = OpenMutex(MUTANT_ALL_ACCESS&, 0, appstring$)

' Check if app is running
If tMutex& <> 0 Then
    CloseHandle tMutex&
    MiscCheckForMutex = True                            ' app is running
Else
    MiscCheckForMutex = False                           ' app is not running or OS error (use GetLastError to check for ERROR_FILE_NOT_FOUND)
End If

Exit Function

' Errors
MiscCheckForMutexError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscCheckForMutex"
ierror = True
Exit Function

End Function

Public Sub MiscMakePath(ByVal sPath As String)
' Creates multiple (nested) directories as necessary

ierror = False
On Error GoTo MiscMakePathError

  Dim Splits() As String, CurFolder As String
  Dim i As Integer
  
  Splits = Split(sPath$, "\")
  For i% = LBound(Splits) To UBound(Splits)
    CurFolder$ = CurFolder$ & Splits(i%) & "\"
    If Dir$(CurFolder$, vbDirectory) = vbNullString Then MkDir CurFolder$
  Next i%

Exit Sub

' Errors
MiscMakePathError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscMakePath"
ierror = True
Exit Sub

End Sub

Function MiscIsNetworkDrive(driveLetter As String) As Boolean
' Check if the specified drive letter is a network drive or not

ierror = False
On Error GoTo MiscIsNetworkDriveError

  Dim strDrive As String
  Dim strRemoteName As String
  Dim lngReturn As Long
  Dim lngBuffer As Long

  strDrive$ = driveLetter$ & ":\"
  strRemoteName$ = String(255, 0)
  lngBuffer = Len(strRemoteName$)

  lngReturn& = WNetGetConnection(strDrive$, strRemoteName$, lngBuffer)

  If lngReturn& = 0 Then
    MiscIsNetworkDrive = True       ' no error returned, buffer contains network drive information
  Else
    MiscIsNetworkDrive = False      ' error getting network drive information, assume not a network drive
  End If

Exit Function

' Errors
MiscIsNetworkDriveError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscIsNetworkDrive"
ierror = True
Exit Function

End Function



