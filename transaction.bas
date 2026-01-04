Attribute VB_Name = "CodeTRANSACTION"
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

Dim wrkDefault As Workspace
    
Dim CurrentTransactionStatus As Integer     ' transaction counter (for nested transactions)

Sub TransactionBegin(tprocedurename As String, tdatabasename As String)
' Handles begin transactions (positive numbers indicate existing transactions)

ierror = False
On Error GoTo TransactionBeginError

Dim i As Integer

Set wrkDefault = DBEngine.Workspaces(0)

If CurrentTransactionStatus% > 0 Then
msg$ = "A BEGIN transaction process was attempted by procedure "
msg$ = msg$ & tprocedurename$ & " on database " & tdatabasename$ & " "
msg$ = msg$ & "before the current transaction was committed." & vbCrLf & vbCrLf
msg$ = msg$ & "This is not a good thing generally and may indicate that "
msg$ = msg$ & "nested transactions are invoked. The program will "
msg$ = msg$ & "attempt to commit the current transactions and "
msg$ = msg$ & "continue with the new BEGIN transaction. " & vbCrLf & vbCrLf
msg$ = msg$ & "However it is strongly suggested that this warning "
msg$ = msg$ & "(including the procedure and database names above) "
msg$ = msg$ & "be reported to Probe Software technical support "
msg$ = msg$ & "immediately (if not sooner)."
MsgBox msg$, vbOKOnly + vbExclamation, "TransactionBegin"

' Loop until all transactions are comitted
For i% = 1 To CurrentTransactionStatus%
wrkDefault.CommitTrans  'dbFlushOSCacheWrites
Next i%

CurrentTransactionStatus% = 0
End If

' Start new transaction
wrkDefault.BeginTrans
CurrentTransactionStatus% = CurrentTransactionStatus% + 1

Exit Sub

' Errors
TransactionBeginError:
MsgBox Error$, vbOKOnly + vbCritical, "TransactionBegin"
ierror = True
Exit Sub

End Sub

Sub TransactionCommit(tprocedurename As String, tdatabasename As String)
' Handles commit transactions

ierror = False
On Error GoTo TransactionCommitError

Set wrkDefault = DBEngine.Workspaces(0)

If CurrentTransactionStatus% < 1 Then
msg$ = "A COMMIT transaction process was attempted by procedure "
msg$ = msg$ & tprocedurename$ & " on database " & tdatabasename$ & " "
msg$ = msg$ & "when no current transaction is pending." & vbCrLf & vbCrLf
msg$ = msg$ & "This is not a good thing generally and may indicate that "
msg$ = msg$ & "some database transactions are not protected. The program will "
msg$ = msg$ & "skip the COMMIT transaction and continue. " & vbCrLf & vbCrLf
msg$ = msg$ & "However it is very strongly suggested that this warning "
msg$ = msg$ & "(including the procedure and database names above) "
msg$ = msg$ & "be reported to Probe Software technical support "
msg$ = msg$ & "immediately (if not sooner)."
MsgBox msg$, vbOKOnly + vbExclamation, "TransactionCommit"
Exit Sub
End If

wrkDefault.CommitTrans
CurrentTransactionStatus% = CurrentTransactionStatus% - 1

Exit Sub

' Errors
TransactionCommitError:
MsgBox Error$, vbOKOnly + vbCritical, "TransactionCommit"
ierror = True
Exit Sub

End Sub

Sub TransactionRollback(tprocedurename As String, tdatabasename As String)
' Handles rollback transactions

ierror = False
On Error GoTo TransactionRollbackError

Set wrkDefault = DBEngine.Workspaces(0)

If CurrentTransactionStatus% < 1 Then
msg$ = "A ROLLBACK transaction process was attempted by procedure "
msg$ = msg$ & tprocedurename$ & " on database " & tdatabasename$ & " "
msg$ = msg$ & "when no current transaction is pending." & vbCrLf & vbCrLf
msg$ = msg$ & "This is generally not a serious problem and the ROLLBACK "
msg$ = msg$ & "process will simply be ignored." & vbCrLf & vbCrLf
msg$ = msg$ & "However it is suggested that this warning "
msg$ = msg$ & "(including the procedure and database names above) "
msg$ = msg$ & "be reported to Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "TransactionRollback"
Exit Sub
End If

wrkDefault.Rollback
CurrentTransactionStatus% = CurrentTransactionStatus% - 1

Exit Sub

' Errors
TransactionRollbackError:
MsgBox Error$, vbOKOnly + vbCritical, "TransactionRollback"
ierror = True
Exit Sub

End Sub

Sub TransactionUnload(procedurename As String)
' Checks is there are any pending transactions when the database file is closed or program terminates

ierror = False
On Error GoTo TransactionUnloadError

Dim response As Integer, i As Integer

Set wrkDefault = DBEngine.Workspaces(0)

If CurrentTransactionStatus% > 0 Then
msg$ = "The program is attempting to terminate the application, "
msg$ = msg$ & "from procedure " & procedurename$ & "." & vbCrLf & vbCrLf
msg$ = msg$ & "However there are apparently some pending transactions "
msg$ = msg$ & "that have not been saved. " & vbCrLf & vbCrLf
msg$ = msg$ & "Would you like to save the pending transactions "
msg$ = msg$ & "(recommended)? Otherwise, click No to just terminate the "
msg$ = msg$ & "application without saving the pending transactions. "
msg$ = msg$ & "In any case it is strongly suggested that this warning "
msg$ = msg$ & "(including the procedure name above) be reported to "
msg$ = msg$ & "Probe Software technical support."
response% = MsgBox(msg$, vbYesNo + vbExclamation + vbDefaultButton1, "TransactionUnload")
If response% = vbNo Then Exit Sub

' Loop until all transactions are comitted
For i% = 1 To CurrentTransactionStatus%
wrkDefault.CommitTrans
Next i%

CurrentTransactionStatus% = 0
End If

Exit Sub

' Errors
TransactionUnloadError:
MsgBox Error$, vbOKOnly + vbCritical, "TransactionUnload"
ierror = True
Exit Sub

End Sub
