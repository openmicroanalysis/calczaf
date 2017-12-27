Attribute VB_Name = "CodeOUTPUTExcel"
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

Dim Excel_Version As Single

Sub OutputSaveCustom2SendToExcel(MaxNumberofColumnsToOutput As Integer, arraysize As Integer, filenamearray() As String, tForm As Form)
' Send list of data files to excel

ierror = False
On Error GoTo OutputSaveCustom2SendToExcelError

Dim response As Integer

' Check if user wants to send to excel
msg$ = "Do you want to send the custom output data files to Excel?"
response% = MsgBox(msg$, vbYesNoCancel + vbQuestion + vbDefaultButton1, "OutputSaveCustom2SendToExcel")

If response% = vbNo Then
Exit Sub
End If

If response% = vbCancel Then
ierror = True
Exit Sub
End If

' Check Excel version number
Excel_Version! = ExcelVersionNumber!()
If Excel_Version = 0# Then GoTo OutputSaveCustom2SendToExcelNotInstalled

If MaxNumberofColumnsToOutput% > MAX_EXCEL_2003_COLS% And Excel_Version! < 12# Then
msg$ = "Warning: More than " & Format$(MAX_EXCEL_2003_COLS%) & " columns of data to output. This "
msg$ = msg$ & "requires Excel 2007 (version 12), but your Excel version is version " & Format$(Excel_Version!) & ". "
msg$ = msg$ & "Do you want to continue to export your data to Excel anyway?"
response% = MsgBox(msg$, vbYesNoCancel + vbQuestion + vbDefaultButton1, "OutputSaveCustom2SendToExcel")
If response% = vbNo Then
Exit Sub
End If
If response% = vbCancel Then
ierror = True
Exit Sub
End If
End If

' Send all files to excel
Call ExcelSendFileListToExcel(arraysize%, filenamearray$(), tForm)
If ierror Then Exit Sub

Exit Sub

' Errors
OutputSaveCustom2SendToExcelError:
MsgBox Error$, vbOKOnly + vbCritical, "OutputSaveCustom2SendToExcel"
Close (Temp1FileNumber%)
ierror = True
Exit Sub

OutputSaveCustom2SendToExcelNotInstalled:
msg$ = "The Excel application is not installed on this computer."
MsgBox msg$, vbOKOnly + vbExclamation, "OutputSaveCustom2SendToExcel"
Close (Temp1FileNumber%)
ierror = True
Exit Sub

End Sub

