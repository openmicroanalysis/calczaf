Attribute VB_Name = "CodeIOBROWSE"
' (c) Copyright 1995-2019 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

' AppData folders
Const SpecialFolder_AppData = &H1A        ' for the current Windows user (roaming), on any computer on the network [Windows 98 or later]
Const SpecialFolder_CommonAppData = &H23  ' for all Windows users on this computer [Windows 2000 or later]
Const SpecialFolder_LocalAppData = &H1C   ' for the current Windows user (non roaming), on this computer only [Windows 2000 or later]
Const SpecialFolder_Documents = &H5       ' the Documents folder for the current Windows user

' Browse for folder
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
      
Private Const MAX_PATH& = 260
Private Const WM_USER& = &H400
Private Const BFFM_INITIALIZED& = 1

' Common to both string (path) and IPDL methods
Private Type BROWSEINFO
 hOwner As Long
 pidlRoot As Long
 pszDisplayName As String
 lpszTitle As String
 ulFlags As Long
 lpfn As Long
 lparam As Long
 iImage As Long
End Type

'Private Const BIF_RETURNONLYFSDIRS& = &H1&   ' only return file system directories. If the user selects folders that are not part of the file system, the OK button is grayed.
'Private Const BIF_DONTGOBELOWDOMAIN& = &H2&  ' do not include network folders below the domain level in the dialog box's tree view control.
'Private Const BIF_STATUSTEXT& = &H4&         ' include a status area in the dialog box. The callback function can set the status text by sending messages to the dialog box. This flag is not supported when BIF_NEWDIALOGSTYLE is specified.
'Private Const BIF_RETURNFSANCESTORS& = &H8&  ' only return file system ancestors. An ancestor is a subfolder that is beneath the root folder in the namespace hierarchy. If the user selects an ancestor of the root folder that is not part of the file system, the OK button is grayed.
'Private Const BIF_EDITBOX& = &H10&           ' version 4.71. Include an edit control in the browse dialog box that allows the user to type the name of an item.
'Private Const BIF_VALIDATE& = &H20&          ' version 4.71. If the user types an invalid name into the edit box, the browse dialog box calls the application's BrowseCallbackProc with the BFFM_VALIDATEFAILED message. This flag is ignored if BIF_EDITBOX is not specified.
Private Const BIF_NEWDIALOGSTYLE& = &H40&    ' version 5.0. Use the new user interface. Setting this flag provides the user with a larger dialog box that can be resized. The dialog box has several new capabilities, including: drag-and-drop capability within the dialog box, reordering, shortcut menus, new folders, delete, and other shortcut menu commands. Note  If Component Object Model (COM) is initialized through CoInitializeEx with the COINIT_MULTITHREADED flag set, SHBrowseForFolder fails if BIF_NEWDIALOGSTYLE is passed.
'Private Const BIF_BROWSEINCLUDEURLS& = &H80& ' version 5.0. The browse dialog box can display URLs. The BIF_USENEWUI and BIF_BROWSEINCLUDEFILES flags must also be set. If any of these three flags are not set, the browser dialog box rejects URLs. Even when these flags are set, the browse dialog box displays URLs only if the folder that contains the selected item supports URLs. When the folder's IShellFolder::GetAttributesOf method is called to request the selected item's attributes, the folder must set the SFGAO_FOLDER attribute flag. Otherwise, the browse dialog box will not display the URL.
'Private Const BIF_UAHINT& = &H100&           ' version 6.0. When combined with BIF_NEWDIALOGSTYLE, adds a usage hint to the dialog box, in place of the edit box. BIF_EDITBOX overrides this flag.
Private Const BIF_NONEWFOLDERBUTTON& = &H200&   ' version 6.0. Do not include the New Folder button in the browse dialog box.
'Private Const BIF_NOTRANSLATETARGETS& = &H400&  ' version 6.0. When the selected item is a shortcut, return the PIDL of the shortcut itself rather than its target.
'Private Const BIF_BROWSEFORCOMPUTER& = &H1000&  ' only return computers. If the user selects anything other than a computer, the OK button is grayed.
'Private Const BIF_BROWSEFORPRINTER& = &H2000&   ' only allow the selection of printers. If the user selects anything other than a printer, the OK button is grayed. In Microsoft Windows XP and later systems, the best practice is to use a Windows XP-style dialog, setting the root of the dialog to the Printers and Faxes folder (CSIDL_PRINTERS).
'Private Const BIF_BROWSEINCLUDEFILES& = &H4000& ' version 4.71. The browse dialog box displays files as well as folders.
'Private Const BIF_SHAREABLE& = &H8000&          ' version 5.0. The browse dialog box can display shareable resources on remote systems. This is intended for applications that want to expose remote shares on a local system. The BIF_NEWDIALOGSTYLE flag must also be set.

'Private Const BIF_USENEWUI& = (BIF_EDITBOX Or BIF_NEWDIALOGSTYLE) ' version 5.0. Use the new user interface, including an edit box. This flag is equivalent to BIF_EDITBOX | BIF_NEWDIALOGSTYLE. Note  If COM is initialized through CoInitializeEx with the COINIT_MULTITHREADED flag set, SHBrowseForFolder fails if BIF_USENEWUI is passed.

' Constants ending in 'A' are for Win95 ANSI calls
' Constants ending in 'W' are the wide Unicode calls for NT

' Sets the status text to the null-terminated string specified by the lParam parameter
' wParam is ignored and should be set to 0
'Private Const BFFM_SETSTATUSTEXTA As Long = (WM_USER& + 100)
'Private Const BFFM_SETSTATUSTEXTW As Long = (WM_USER& + 104)

' Selects the specified folder. If the wParam parameter is FALSE, the lParam parameter
' is the PIDL of the folder to select, or it is the path of the folder if wParam is
' the C value TRUE (or 1). Note that after this message is sent, the browse dialog
' receives a subsequent BFFM_SELECTIONCHANGED message.
Private Const BFFM_SETSELECTIONA As Long = (WM_USER& + 102)
'Privte Const BFFM_SETSELECTIONW As Long = (WM_USER& + 103)
  
' If the lParam  parameter is non-zero, enables the OK button, or disables it if
' lParam is zero (docs erroneously said wParam!). wParam is ignored and should be set to 0.
'Private Const BFFM_ENABLEOK As Long = (WM_USER& + 101)

' Specific to the PIDL method (undocumented function) IShellFolder's ParseDisplayName member function should be used instead
Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long

' Specific to the STRING method
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
'Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long

Private Const LMEM_FIXED& = &H0
Private Const LMEM_ZEROINIT& = &H40
Private Const lPtr& = (LMEM_FIXED& Or LMEM_ZEROINIT&)

Public Function BrowseCallbackProcStr(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lparam As Long, ByVal lpData As Long) As Long
' Callback for the Browse STRING method. On initialization, set the dialog's
' pre-selected folder from the pointer to the path allocated as bi.lParam, passed
' back to the callback as lpData param

ierror = False
On Error GoTo BrowseCallbackProcStrError

Select Case uMsg
  Case BFFM_INITIALIZED
    Call SendMessage(hWnd, BFFM_SETSELECTIONA, True, ByVal lpData)
  Case Else:
End Select

Exit Function

' Errors
BrowseCallbackProcStrError:
MsgBox Error$, vbOKOnly + vbCritical, "BrowseCallbackProcStr"
ierror = True
Exit Function

End Function
      
Public Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lparam As Long, ByVal lpData As Long) As Long
' Callback for the Browse PIDL method. On initialization, set the dialog's pre-selected folder using the
'  pidl set as the bi.lParam, and passed back to the callback as lpData param

ierror = False
On Error GoTo BrowseCallbackProcError

Select Case uMsg
  Case BFFM_INITIALIZED
    Call SendMessage(hWnd, BFFM_SETSELECTIONA, False, ByVal lpData)
  Case Else:
End Select

Exit Function

' Errors
BrowseCallbackProcError:
MsgBox Error$, vbOKOnly + vbCritical, "BrowseCallbackProc"
ierror = True
Exit Function

End Function

Public Function BrowseFarProc(pfn As Long) As Long
' A dummy procedure that receives and returns the value of the AddressOf operator
' This workaround is needed as you can't assign AddressOf directly to a member of
' a user-defined type, but you can assign it to another long.

ierror = False
On Error GoTo BrowseFarProcError

BrowseFarProc& = pfn
Exit Function

' Errors
BrowseFarProcError:
MsgBox Error$, vbOKOnly + vbCritical, "BrowseFarProc"
ierror = True
Exit Function
            
End Function
  
Public Function BrowseGetPIDLFromPath(sPath As String) As Long
' Return the pidl to the path supplied by calling the undocumented API #162 (our
' name SHSimpleIDListFromPath). This function is necessary as, unlike documented APIs,
' the API is not implemented in 'A' or 'W' versions.

ierror = False
On Error GoTo BrowseGetPIDLFromPathError

If MiscSystemGetOSVersionNumber&() < OS_VERSION_WINNT& Then
 BrowseGetPIDLFromPath = SHSimpleIDListFromPath(StrConv(sPath, vbUnicode))
Else
  BrowseGetPIDLFromPath = SHSimpleIDListFromPath(sPath)
End If

Exit Function

' Errors
BrowseGetPIDLFromPathError:
MsgBox Error$, vbOKOnly + vbCritical, "BrowseGetPIDLFromPathError"
ierror = True
Exit Function

End Function

Public Function BrowseUnqualifyPath(sPath As String) As String
' Qualifying a path usually involves assuring that its format is valid, including a
' trailing slash ready for a filename. Since the SHBrowseForFolder API will pre-select
' the path if it contains the trailing slash, I call stripping it 'unqualifying the path'.
 
ierror = False
On Error GoTo BrowseUnqualifyPathError
 
If Len(sPath$) > 0 Then
  If Right$(sPath$, 1) = "\" Then
    BrowseUnqualifyPath$ = Left$(sPath$, Len(sPath$) - 1)
    Exit Function
  End If
End If
     
BrowseUnqualifyPath = sPath$
Exit Function

' Errors
BrowseUnqualifyPathError:
MsgBox Error$, vbOKOnly + vbCritical, "BrowseUnqualifyPath"
ierror = True
Exit Function
            
End Function

Public Function IOBrowseForFolderByPath(bMakeNewFolder As Boolean, sSelPath As String, sSelString As String, tForm As Form) As String
' Select folder using start path

ierror = False
On Error GoTo IOBrowseForFolderByPathError

Dim BI As BROWSEINFO
Dim pidl As Long, lpSelPath As Long
Dim sPath As String * MAX_PATH&
    
' The call can not have a trailing slash, so strip it from the path if present
sSelPath$ = BrowseUnqualifyPath(sSelPath$)

' Check for blank path
If Trim$(sSelPath$) = vbNullString Then sSelPath$ = "\"

With BI
        .hOwner = tForm.hWnd
        .pidlRoot = 0
        .lpszTitle = sSelString$
        
    If bMakeNewFolder Then
        .ulFlags& = BIF_NEWDIALOGSTYLE&     ' new dialog style with Make New Folder Button
    Else
        .ulFlags& = BIF_NEWDIALOGSTYLE& Or BIF_NONEWFOLDERBUTTON&     ' new dialog style without Make New Folder button
    End If
  
        .lpfn = BrowseFarProc(AddressOf BrowseCallbackProcStr)
      
    lpSelPath = LocalAlloc(lPtr, Len(sSelPath$) + 1)
    CopyMemory ByVal lpSelPath&, ByVal sSelPath$, Len(sSelPath$) + 1
        .lparam = lpSelPath
End With
      
pidl& = SHBrowseForFolder(BI)
     
If pidl& Then
       
  If SHGetPathFromIDList(pidl&, sPath$) Then
    IOBrowseForFolderByPath$ = Left$(sPath$, InStr(sPath$, vbNullChar) - 1)
  End If
        
Call CoTaskMemFree(pidl&)
     
End If
     
Call LocalFree(lpSelPath&)
Exit Function

' Errors
IOBrowseForFolderByPathError:
MsgBox Error$, vbOKOnly + vbCritical, "IOBrowseForFolderByPath"
ierror = True
Exit Function

End Function

Public Function IOBrowseForFolderByPIDL(sSelPath As String, sSelString As String, tForm As Form) As String
' Select a folder using start PIDL

ierror = False
On Error GoTo IOBrowseForFolderByPIDLError

Dim BI As BROWSEINFO
Dim pidl As Long
Dim sPath As String * MAX_PATH&

' The call can not have a trailing slash, so strip it from the path if present
sSelPath$ = BrowseUnqualifyPath(sSelPath$)

' Check for blank path
If Trim$(sSelPath$) = vbNullString Then sSelPath$ = "\"

With BI
  .hOwner = tForm.hWnd
  .pidlRoot = 0
  .lpszTitle = sSelString$
  .lpfn = BrowseFarProc(AddressOf BrowseCallbackProc)
  .lparam = BrowseGetPIDLFromPath(sSelPath)
End With
    
pidl = SHBrowseForFolder(BI)
    
If pidl Then
  If SHGetPathFromIDList(pidl, sPath) Then
    IOBrowseForFolderByPIDL = Left$(sPath, InStr(sPath, vbNullChar) - 1)
  End If
       
' Free the pidl returned by call to SHBrowseForFolder
Call CoTaskMemFree(pidl)
End If
    
' Free the pidl set in call to BrowseGetPIDLFromPath
Call CoTaskMemFree(BI.lparam)
    
Exit Function

' Errors
IOBrowseForFolderByPIDLError:
MsgBox Error$, vbOKOnly + vbCritical, "IOBrowseForFolderByPIDL"
ierror = True
Exit Function

End Function

Function IOBrowseGetAppDataFolder(ssfParameter As Long) As String
' Returns the AppData folder based on the passed constant

ierror = False
On Error GoTo IOBrowseGetAppDataFolderError

Dim sAppData As String
Dim objShell  As Object
Dim objFolder As Object

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(CLng(ssfParameter&))

If (Not objFolder Is Nothing) Then sAppData$ = objFolder.Self.Path

Set objFolder = Nothing
Set objShell = Nothing

If sAppData$ = vbNullString Then GoTo IOBrowseGetAppdataFolderNotFound
IOBrowseGetAppDataFolder$ = sAppData$

Exit Function

' Errors
IOBrowseGetAppDataFolderError:
MsgBox Error$, vbOKOnly + vbCritical, "IOBrowseGetAppDataFolder"
ierror = True
Exit Function

IOBrowseGetAppdataFolderNotFound:
MsgBox "The application data folder path could not be detected", vbOKOnly + vbCritical, "IOBrowseGetAppDataFolder"
ierror = True
Exit Function

End Function


