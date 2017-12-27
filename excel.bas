Attribute VB_Name = "CodeExcel"
' (c) Copyright 1995-2018 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

'Private Const ERR_EXCEL_NOTRUNNING& = 429

' Declare necessary API routines:
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As Long) As Long

Dim xlApp As Object     'Excel.Application
Dim xlBook As Object
Dim xlSheet As Object

Dim ExcelRow As Long, ExcelColumn As Long

Dim ExcelName As String
Dim ExcelNcol As Integer
Dim ExcelLabel(1 To MAXCHAN% + 2) As String

Dim Excel_Version As Single

Sub ExcelCreateSpreadsheet(tForm As Form)
' Create Excel spreadsheet link

ierror = False
On Error GoTo ExcelCreateSpreadsheetError

' Save existing data
Call ExcelCloseSpreadsheet(vbNullString, tForm)
If ierror Then Exit Sub

Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.Workbooks.Add
Set xlSheet = xlBook.Worksheets(1)

' Make Excel visible through the Application object
xlApp.Application.Visible = True

ExcelRow& = 1
ExcelColumn = 0

Exit Sub

' Errors
ExcelCreateSpreadsheetError:
MsgBox Error$, vbOKOnly + vbCritical, "ExcelCreateSpreadsheet"
ierror = True
Exit Sub

End Sub

Sub ExcelCloseSpreadsheet(tfilename As String, tForm As Form)
' Close Excel spreadsheet link

ierror = False
On Error GoTo ExcelCloseSpreadsheetError

Dim response As Integer
Dim oldpath As String

' Check for Excel application running
If Not ExcelAppIsRunning() Then Exit Sub

' Check for link
If xlBook Is Nothing Then Exit Sub
If xlSheet Is Nothing Then Exit Sub

' Save current path
oldpath$ = CurDir$

' Change to new path
If ProbeDataFile$ <> vbNullString Then
Call MiscChangePath(ProbeDataFile$)
If ierror Then Exit Sub
Else
Call MiscChangePath(ApplicationCommonAppData$)
If ierror Then Exit Sub
End If

' Confirm filename
If ExcelRow& > 1 Then

' Ask user whether to save excel data before closing
msg$ = "Do you want to save the changes you made to the Excel spreadsheet?"
response% = MsgBox(msg$, vbYesNoCancel + vbQuestion + vbDefaultButton1 + vbMsgBoxSetForeground, "ExcelCloseSpreadsheet")

If response% = vbCancel Then Exit Sub
If response% = vbNo Then GoTo ExcelCloseSpreadsheetNoSave

' Get filename from user (use XLS for Excel 2003 (version 11) (Office 2003) and earlier)
If Excel_Version! < 12# Then
Call IOGetFileName(Int(1), "XLS", tfilename$, tForm)
If ierror Then Exit Sub

' Use XLSX for Excel 2007 (version 12) (Office 2007) and higher
Else
Call IOGetFileName(Int(1), "XLSX", tfilename$, tForm)
If ierror Then Exit Sub
End If

' Check for open spreadsheet (does not work (FindWindow finds XLMAIN)
'If Not ExcelAppIsRunning() Then
'GoTo ExcelCloseSpreadsheetNoApp
'ExcelNcol% = 0  ' reset label and rows
'ExcelRow& = 1
'End If

' Save
xlBook.SaveAs tfilename$
xlBook.Close
End If

' Close Excel with the Quit method on the Application object
ExcelCloseSpreadsheetNoSave:
xlApp.Application.Quit

' Restore original directories
Call MiscChangePath(oldpath$)
If ierror Then Exit Sub

' Release the object variables
Set xlSheet = Nothing
Set xlBook = Nothing
Set xlApp = Nothing

' Reset label and rows
ExcelNcol% = 0
ExcelRow& = 1

Exit Sub

' Errors
ExcelCloseSpreadsheetError:
MsgBox Error$, vbOKOnly + vbCritical, "ExcelCloseSpreadsheet"
ierror = True
Exit Sub

'ExcelCloseSpreadsheetNoApp:
'msg$ = "Excel is not running, so file must have been saved already or Excel was closed"
'MsgBox msg$, vbOKOnly + vbExclamation, "ExcelCloseSpreadsheet"
'ierror = True
'Exit Sub

End Sub

Sub ExcelSendDataToSpreadsheet(nCol As Integer, temp() As Double)
' Send data to Excel spreadsheet

ierror = False
On Error GoTo ExcelSendDataToSpreadsheetError

Dim i As Integer

' Check for link
If xlSheet Is Nothing Then Exit Sub

' Check for Excel application running
If Not ExcelAppIsRunning() Then GoTo ExcelSendDataToSpreadsheetNoApp

' Load cell range
For i% = 1 To nCol%
xlSheet.Cells(ExcelRow&, ExcelColumn& + i%).value = MiscSetSignificantDigits(10, temp#(i%))
Next i%

' Increment row
ExcelRow& = ExcelRow& + 1
Exit Sub

' Errors
ExcelSendDataToSpreadsheetError:
MsgBox Error$, vbOKOnly + vbCritical, "ExcelSendDataToSpreadsheet"
ierror = True
Exit Sub

ExcelSendDataToSpreadsheetNoApp:
msg$ = "Excel is not running"
MsgBox msg$, vbOKOnly + vbExclamation, "ExcelSendDataToSpreadsheet"
ierror = True
Exit Sub

End Sub

Sub ExcelSendLabelToSpreadsheet(mode As Integer, nCol As Integer, astring As String, tlabel() As String)
' Send labels if changed
' mode = 0 normal check for changed labels
' mode = 1 force label output

ierror = False
On Error GoTo ExcelSendLabelToSpreadsheetError

Dim i As Integer, changed As Integer

' Check for link
If xlSheet Is Nothing Then Exit Sub

' Check for Excel application running
If Not ExcelAppIsRunning() Then GoTo ExcelSendLabelToSpreadsheetNoApp

' Check for change
If mode% = 1 Then changed = True    ' force for Analyze! window
If ExcelNcol% <> nCol% Then changed = True
If ExcelName$ <> astring$ Then changed = True

If Not changed Then
For i% = 1 To nCol%
If ExcelLabel$(i%) <> tlabel$(i%) Then changed = True
Next i%
End If

If Not changed Then Exit Sub

' Load sample name
xlSheet.Cells(ExcelRow&, 1).value = VbDquote$ & astring$ & VbDquote$

' Increment row
ExcelRow& = ExcelRow& + 1

' Load cell range
For i% = 1 To nCol%
xlSheet.Cells(ExcelRow&, ExcelColumn& + i%).value = VbDquote$ & tlabel$(i%) & VbDquote$
Next i%

' Increment row
ExcelRow& = ExcelRow& + 1

' Save for next time
ExcelName$ = astring$
ExcelNcol% = nCol%
For i% = 1 To nCol%
ExcelLabel$(i%) = tlabel$(i%)
Next i%

Exit Sub

' Errors
ExcelSendLabelToSpreadsheetError:
MsgBox Error$, vbOKOnly + vbCritical, "ExcelSendLabelToSpreadsheet"
ierror = True
Exit Sub

ExcelSendLabelToSpreadsheetNoApp:
msg$ = "Excel is not running"
MsgBox msg$, vbOKOnly + vbExclamation, "ExcelSendLabelToSpreadsheet"
ierror = True
Exit Sub

End Sub

Function ExcelSheetIsOpen() As Integer
' Check for Excel spreadsheet (for routines outside module)
' Returns true if Excel link is open

ierror = False
On Error GoTo ExcelSheetIsOpenError

' Check for link
If xlSheet Is Nothing Then
ExcelSheetIsOpen = False
Else
ExcelSheetIsOpen = True
End If

Exit Function

' Errors
ExcelSheetIsOpenError:
MsgBox Error$, vbOKOnly + vbCritical, "ExcelSheetIsOpen"
ierror = True
Exit Function

End Function

Function ExcelAppIsRunning() As Boolean
' Check for Excel application running
' Returns true if Excel is running

ierror = False
On Error GoTo ExcelAppIsRunningError

Dim hWnd As Long

' If Excel is running this API call returns its handle
hWnd = FindWindow("XLMAIN", 0)

' 0 means Excel not running
If hWnd = 0 Then
ExcelAppIsRunning = False
Else
ExcelAppIsRunning = True
End If

' Alternative code (may be required for future versions of Excel)
'On Error Resume Next
'Set xlApp = GetObject(, "Excel.Application")    ' this line causes an error in debug mode only
'If Err = ERR_EXCEL_NOTRUNNING& Then
'ExcelAppIsRunning = False
'Else
'ExcelAppIsRunning = True
'End If

Exit Function

' Errors
ExcelAppIsRunningError:
MsgBox Error$, vbOKOnly + vbCritical, "ExcelAppIsRunning"
ierror = True
Exit Function

End Function

Sub ExcelSendFileToExcelSheet(arraysize As Integer, tfilename As String)
' Send a tab delimited file to an excel spreadsheet

ierror = False
On Error GoTo ExcelSendFileToExcelSheetError

Dim astring As String, bstring As String

' Write filename (and sample) to string
ExcelRow& = 1   ' not used (each paste goes to a separate worksheet)

' Only write filename to spreadsheet if loading more than one file
If arraysize% > 1 Then
bstring$ = tfilename$ & vbCrLf
Call AnalyzeStatusAnal("Please wait while " & tfilename$ & " is exported to Excel...")
Call IOStatusAuto("Please wait while " & tfilename$ & " is exported to Excel...")
DoEvents
End If

' Open file
Open tfilename$ For Input As Temp1FileNumber%
Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$
bstring$ = bstring$ & astring$ & vbCrLf

' Check for cancel
If icancelanal Or icancelauto Then
ierror = True
Exit Sub
End If

DoEvents
Loop
Close (Temp1FileNumber%)

' Send file to sheet
Sleep (200)     ' need for Win7 clipboard issues
Clipboard.Clear
Sleep (200)     ' need for Win7 clipboard issues
Clipboard.SetText bstring$
Sleep (200)     ' need for Win7 clipboard issues

' Load cell range
xlSheet.Paste

Call AnalyzeStatusAnal(vbNullString)
Call IOStatusAuto(vbNullString)
DoEvents
Exit Sub

' Errors
ExcelSendFileToExcelSheetError:
MsgBox Error$, vbOKOnly + vbCritical, "ExcelSendFileToExcelSheet"
Close (Temp1FileNumber%)
Call AnalyzeStatusAnal(vbNullString)
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub ExcelSendFileListToExcel(arraysize As Integer, filenamearray() As String, tForm As Form)
' Send all files to an Excel spreadsheet in separate sheets

ierror = False
On Error GoTo ExcelSendFileListToExcelSheetError

Dim i As Integer
Dim tfilename As String

' Save existing data
Call ExcelCloseSpreadsheet(vbNullString, tForm)
If ierror Then Exit Sub

' Open excel workbook
Set xlApp = CreateObject("Excel.Application")
If xlApp Is Nothing Then GoTo ExcelSendFileListToExcelNotCreatedApp

' Make Excel visible through the Application object
xlApp.Visible = True
Set xlBook = xlApp.Workbooks.Add
If xlBook Is Nothing Then GoTo ExcelSendFileListToExcelNotCreatedBook

' Get Excel version running (Excel 2007 is version 12)
Excel_Version! = ExcelVersionNumber!()
If ierror Then Exit Sub

' Send each file to a separate sheet
For i% = 1 To arraysize%

' Create next sheet
Set xlSheet = xlBook.Worksheets.Add
Set xlSheet = xlBook.ActiveSheet

' Send file
Call ExcelSendFileToExcelSheet(arraysize%, filenamearray$(i%))
If ierror Then Exit Sub
Next i%

' Close excel workbook
ExcelRow& = 2   ' to force ask for filename
If arraysize% = 1 Then

' Change to .xls extension
If Excel_Version! < 12# Then
tfilename$ = MiscGetFileNameNoExtension$(filenamearray$(1)) & ".xls"
Else
tfilename$ = MiscGetFileNameNoExtension$(filenamearray$(1)) & ".xlsx"
End If

Call ExcelCloseSpreadsheet(tfilename$, tForm)
If ierror Then Exit Sub
Else

' Get filename common basis to save all worksheets to
Call MiscGetFilenameBasis(arraysize%, filenamearray$(), tfilename$)
If ierror Then Exit Sub

' Change to .xls extension
If Excel_Version! < 12# Then
tfilename$ = MiscGetFileNameNoExtension$(tfilename$) & ".xls"
Else
tfilename$ = MiscGetFileNameNoExtension$(tfilename$) & ".xlsx"
End If

Call ExcelCloseSpreadsheet(tfilename$, tForm)
If ierror Then Exit Sub
End If

Exit Sub

' Errors
ExcelSendFileListToExcelSheetError:
MsgBox Error$, vbOKOnly + vbCritical, "ExcelSendFileListToExcelSheet"
ierror = True
Exit Sub

ExcelSendFileListToExcelNotCreatedApp:
msg$ = "Excel application could not be created"
MsgBox msg$, vbOKOnly + vbExclamation, "ExcelSendFileListToExcelSheet"
ierror = True
Exit Sub

ExcelSendFileListToExcelNotCreatedBook:
msg$ = "Excel workbook could not be created"
MsgBox msg$, vbOKOnly + vbExclamation, "ExcelSendFileListToExcelSheet"
ierror = True
Exit Sub

End Sub

Function ExcelVersionNumber() As Single
' Returns the Excel application version number

ierror = False
On Error GoTo ExcelVersionNumberError

Dim txlApp As Object

' Make a temporary instance of Excel, just for checking the version number
Set txlApp = CreateObject("Excel.Application")
If txlApp Is Nothing Then GoTo ExcelVersionNumberNotCreatedApp

' Now get Excel version number (version 0 means Excel is not installed)
ExcelVersionNumber! = txlApp.Version

' Close temp instance of Excel
txlApp.Application.Quit
Set txlApp = Nothing

Exit Function

' Errors (if error number is 429, then Excel is not installed)
ExcelVersionNumberError:
MsgBox Error$, vbOKOnly + vbCritical, "ExcelVersionNumber"
ierror = True
Exit Function

ExcelVersionNumberNotCreatedApp:
msg$ = "Excel application could not be created [CreateObject]"
MsgBox msg$, vbOKOnly + vbExclamation, "ExcelVersionNumber"
ierror = True
Exit Function

End Function

