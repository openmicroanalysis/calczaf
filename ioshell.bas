Attribute VB_Name = "CodeIoShell"
' (c) Copyright 1995-2021 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" _
  (ByVal hWnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long
   
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Const PROCESS_QUERY_INFORMATION& = &H400&
Private Const STILL_ACTIVE& = &H103&
    
Sub IORunShellExecute(sTopic As String, sFIle As Variant, sParams As Variant, sDirectory As Variant, nShowCmd As Long)
' Execute the passed operation, passing the desktop as the window to receive any error messages

ierror = False
On Error GoTo IORunShellExecuteError

Const ERROR_BAD_FORMAT& = 11&
Const ERROR_FILE_NOT_FOUND& = 2&
Const ERROR_PATH_NOT_FOUND& = 3&
Const SE_ERR_ACCESSDENIED& = 5 ' access denied
Const SE_ERR_ASSOCINCOMPLETE& = 27
Const SE_ERR_DDEBUSY& = 30
Const SE_ERR_DDEFAIL& = 29
Const SE_ERR_DDETIMEOUT& = 28
Const SE_ERR_DLLNOTFOUND& = 32
Const SE_ERR_FNF& = 2 ' file not found
Const SE_ERR_NOASSOC& = 31
Const SE_ERR_OOM& = 8 ' out of memory
Const SE_ERR_PNF& = 3 ' path not found
Const SE_ERR_SHARE& = 26

Dim lngReturn As Long

' Run the OS shell execute function to invoke the default application for this document
'    sTopics include: open, edit, explore, mailto, find, print, properties
lngReturn& = ShellExecute(GetDesktopWindow(), sTopic$, sFIle, sParams, sDirectory, nShowCmd&)
  
    Select Case lngReturn&
        Case 0
            MsgBox "The operating system is out of memory or resources.", vbOKOnly + vbExclamation, "IORunShellExecute"
        Case ERROR_BAD_FORMAT&
            MsgBox "The .exe file is invalid (non-Win32® .exe or error in .exe image).", vbOKOnly + vbExclamation, "IORunShellExecute"
        Case ERROR_FILE_NOT_FOUND&
            MsgBox "The specified file (" & sFIle & ") was not found.", vbOKOnly + vbExclamation, "IORunShellExecute"
        Case ERROR_PATH_NOT_FOUND&
            MsgBox "The specified path was not found.", vbOKOnly + vbExclamation, "IORunShellExecute"
        Case SE_ERR_ACCESSDENIED&
            MsgBox "The operating system denied access to the specified file.", vbOKOnly + vbExclamation, "IORunShellExecute"
        Case SE_ERR_ASSOCINCOMPLETE&
            MsgBox "The file name association is incomplete or invalid.", vbOKOnly + vbExclamation, "IORunShellExecute"
        Case SE_ERR_DDEBUSY&
            MsgBox "The DDE transaction could not be completed because other DDE transactions were being processed.", vbOKOnly + vbExclamation, "IORunShellExecute"
        Case SE_ERR_DDEFAIL&
            MsgBox "The DDE transaction failed.", vbOKOnly + vbExclamation, "IORunShellExecute"
        Case SE_ERR_DDETIMEOUT&
            MsgBox "The DDE transaction could not be completed because the request timed out.", vbOKOnly + vbExclamation, "IORunShellExecute"
        Case SE_ERR_DLLNOTFOUND&
            MsgBox "The specified dynamic-link library was not found. ", vbOKOnly + vbExclamation, "IORunShellExecute"
        Case SE_ERR_FNF&
            MsgBox "The specified file was not found.", vbOKOnly + vbExclamation, "IORunShellExecute"
        Case SE_ERR_NOASSOC&
            MsgBox "There is no application associated with the given file name extension. This error will also be returned if you attempt to print a file that is not printable.", vbOKOnly + vbExclamation, "IORunShellExecute"
        Case SE_ERR_OOM&
            MsgBox "There was not enough memory to complete the operation.", vbOKOnly + vbExclamation, "IORunShellExecute"
        Case SE_ERR_PNF&
            MsgBox "The specified path was not found.", vbOKOnly + vbExclamation, "IORunShellExecute"
        Case SE_ERR_SHARE&
            MsgBox "A sharing violation occurred.", vbOKOnly + vbExclamation, "IORunShellExecute"
    End Select

Exit Sub

' Errors
IORunShellExecuteError:
MsgBox Error$ & ", " & sFIle, vbOKOnly + vbCritical, "IORunShellExecute"
ierror = True
Exit Sub

End Sub

Function IOIsProcessTerminated(currentPID As Variant) As Boolean
' Checks whether a shell process ID is still running (call within doevents loop or from timer event)

ierror = False
On Error GoTo IOIsProcessTerminatedError

Dim ProcHnd As Long, CurECode As Long

' Get the process handle
ProcHnd& = OpenProcess(PROCESS_QUERY_INFORMATION&, True, currentPID)

' Check for exit code
Call GetExitCodeProcess(ProcHnd&, CurECode&)

' Return true
If CurECode& = STILL_ACTIVE& Then
IOIsProcessTerminated = False
Else
IOIsProcessTerminated = True
End If

Exit Function

' Errors
IOIsProcessTerminatedError:
MsgBox Error$, vbOKOnly + vbCritical, "IOIsProcessTerminated"
ierror = True
Exit Function

End Function

Sub IOShellTerminateTask(taskID As Long)
' Terminate the task by process ID
'
' taskkill [/s Computer] [/u Domain\User [/p Password]]] [/fi FilterName] [/pid ProcessID]|[/im ImageName] [/f][/t]
'
' Parameters
'  /s   Computer   : Specifies the name or IP address of a remote computer (do not use backslashes). The default is the local computer.
'  /u   Domain \ User   : Runs the command with the account permissions of the user specified by User or Domain\User. The default is the permissions of the current logged on user on the computer issuing the command.
'  /p   Password   : Specifies the password of the user account that is specified in the /u parameter.
'  /fi   FilterName   : Specifies the types of process(es) to include in or exclude from termination. The following are valid filter names, operators, and values.
'  /pid   ProcessID   : Specifies the process ID of the process to be terminated.
'  /im   ImageName   : Specifies the image name of the process to be terminated. Use the wildcard (*) to specify all image names.
'  /f   : Specifies that process(es) be forcefully terminated. This parameter is ignored for remote processes; all remote processes are forcefully terminated.
'  /t   : Specifies to terminate all child processes along with the parent process, commonly known as a tree kill.
'  /? : Displays help at the command prompt.

ierror = False
On Error GoTo IOShellTerminateTaskError

' Check for task still running
If taskID& = 0 Then Exit Sub

' Run Taskkill.exe  (/k executes but window remains, /c executes but terminates)
If Not IOIsProcessTerminated(taskID&) Then
Shell "cmd.exe /c Taskkill /PID " & Format$(taskID&), vbMinimizedFocus
End If
    
Exit Sub

' Errors
IOShellTerminateTaskError:
MsgBox Error$, vbOKOnly + vbCritical, "IOShellTerminateTask"
ierror = True
Exit Sub

End Sub

