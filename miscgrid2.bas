Attribute VB_Name = "CodeMiscGrid2"
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub MiscCopyGrid2(tGrid As MSFlexGrid)
' Copies data from the grid to the clipboard (for rows with non-blank 1st column)

ierror = False
On Error GoTo MiscCopyGrid2Error

Dim i As Integer, j As Integer
Dim tcols As Integer, trows As Integer

' Determine number of nonblank rows and columns
For j% = 0 To tGrid.rows - 1
tGrid.row = j%
For i% = 0 To tGrid.cols - 1
tGrid.col = i%
If Trim$(tGrid.Text) <> vbNullString Then
trows% = j%
tcols% = i%
End If
Next i%
Next j%

' Select the nonblank cells (assumes that the zeroth row and column are non-blank!)
tGrid.row = 0
tGrid.col = 0
tGrid.RowSel = trows%
tGrid.ColSel = tcols%
tGrid.TopRow = 1

' Copy the selection and put it on the Clipboard
Clipboard.Clear
Sleep (200)     ' need for Win7 clipboard issues
Clipboard.SetText tGrid.Clip

Exit Sub

' Errors
MiscCopyGrid2Error:
MsgBox Error$, vbOKOnly + vbCritical, "MiscCopyGrid2"
ierror = True
Exit Sub

End Sub

