Attribute VB_Name = "CodeExec"
' (c) Copyright 1995-2024 by John J. Donovan
Option Explicit

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
  lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
  ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As _
  PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const NORMAL_PRIORITY_CLASS& = &H20&
Private Const INFINITE& = -1&

Public Sub ExecRun(cmdline As String)
' Determines the OS and calls appropriate command line process

ierror = False
On Error GoTo ExecRunError

' XP, NT or 95/98
If FormMAIN!SysInfo1.OSPlatform < 3 Then
Call ExecRunCommand(cmdline$)
If ierror Then Exit Sub

Else
msg$ = "Unsupported operating system, please contact technical support"
MsgBox msg$, vbOKOnly + vbExclamation, "ExecRun"
ierror = True
Exit Sub
End If

Exit Sub

' Errors
ExecRunError:
MsgBox Error$, vbOKOnly + vbCritical, "ExecRun"
ierror = True
Exit Sub

End Sub

Public Sub ExecRunCommand(cmdline As String)
' Routine to execute a synchronous process for 32 OS

ierror = False
On Error GoTo ExecRunCommandError

Dim ret As Long

Dim proc As PROCESS_INFORMATION
Dim start As STARTUPINFO

' Initialize the STARTUPINFO structure:
start.cb = Len(start)

' Start the shelled application:
ret& = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)

' Wait for the shelled application to finish:
ret& = WaitForSingleObject(proc.hProcess, INFINITE)
ret& = CloseHandle(proc.hProcess)

Exit Sub

' Errors
ExecRunCommandError:
MsgBox Error$, vbOKOnly + vbCritical, "ExecRunCommand"
ierror = True
Exit Sub

End Sub





