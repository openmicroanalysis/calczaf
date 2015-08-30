Attribute VB_Name = "CodeMiscGrid"
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

Const MAXGEOLOGICAL% = MAXELM%
Dim GeologicalOrder(1 To MAXGEOLOGICAL%) As String

Dim SortFlags() As Integer

Dim SortArray() As String

Sub MiscBlankGrid(tGrid As Grid)
' Blank the grid that was passed

ierror = False
On Error GoTo MiscBlankGridError

Dim i As Integer, j As Integer

Screen.MousePointer = vbHourglass

' First blank the fixed columns
For j% = 0 To tGrid.FixedRows
For i% = 0 To tGrid.cols - 1
tGrid.row = j%
tGrid.col = i%
tGrid.Text = vbNullString
Next i%
Next j%

' Now blank the fixed rows
For j% = 0 To tGrid.FixedCols
For i% = 0 To tGrid.rows - 1
tGrid.col = j%
tGrid.row = i%
tGrid.Text = vbNullString
Next i%
Next j%

' Next blank the non fixed columns and rows
tGrid.SelStartCol = 0 + tGrid.FixedCols
tGrid.SelEndCol = tGrid.cols - 1
tGrid.SelStartRow = 0 + tGrid.FixedRows
tGrid.SelEndRow = tGrid.rows - 1
tGrid.Clip = vbNullString

' Now unselect the grid
tGrid.SelStartCol = 0 + tGrid.FixedCols
tGrid.SelEndCol = tGrid.SelStartCol
tGrid.SelStartRow = 0 + tGrid.FixedRows
tGrid.SelEndRow = tGrid.SelStartRow

Screen.MousePointer = vbDefault
Exit Sub

' Errors
MiscBlankGridError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "MiscBlankGrid"
ierror = True
Exit Sub

End Sub

Sub MiscCopyGrid(mode As Integer, tGrid As Grid)
' Copies data from the grid to the clipboard (for rows with non-blank 1st column)
' mode = 1 copy grid
' mode = 2 copy grid (selected) only

ierror = False
On Error GoTo MiscCopyGridError

Dim i As Integer, j As Integer
Dim tmsg As String

' Copy fixed grid column titles
tmsg$ = vbNullString
tGrid.row = 0
For i% = 0 To tGrid.cols - 1
tGrid.col = i%
If Trim$(tGrid.Text) <> vbNullString Then
tmsg$ = tmsg$ & tGrid.Text & vbTab
Else
'tmsg$ = tmsg$ & "     " & vbTab    ' comment out to avoid adding zeros
End If
Next i%
tmsg$ = tmsg$ & vbCrLf

' Copy Grid
For j% = tGrid.FixedRows To tGrid.rows - 1
tGrid.row = j%

tGrid.col = 0
If Trim$(tGrid.Text) <> vbNullString Then
For i% = 0 To tGrid.cols - 1
tGrid.col = i%

If Trim$(tGrid.Text) <> vbNullString Then
If mode% = 1 Then
tmsg$ = tmsg$ & tGrid.Text & vbTab
ElseIf mode% = 2 And tGrid.CellSelected Then
tmsg$ = tmsg$ & tGrid.Text & vbTab
End If

Else
'tmsg$ = tmsg$ & "  .000" & vbTab   ' comment out to avoid adding zeros
End If

Next i%
tmsg$ = tmsg$ & vbCrLf
End If
Next j%

Clipboard.Clear
Sleep (200)     ' need for Win7 clipboard issues
Clipboard.SetText tmsg$

Exit Sub

' Errors
MiscCopyGridError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscCopyGrid"
ierror = True
Exit Sub

End Sub

Sub MiscCopyGridFromTo(mode As Integer, sArray() As String, row1 As Integer, row2 As Integer, col1 As Integer, col2 As Integer, tGrid As Grid)
' Copies data from the grid to the passed array (rows row1 to row2 and columns col1 to col2 using zero based index)
'  mode = 1 copy from grid to array
'  mode = 2 copy from array to grid

ierror = False
On Error GoTo MiscCopyGridFromToError

Dim i As Integer, j As Integer

' Loop through grid or array
For j% = row1% To row2%
tGrid.row = j%

For i% = col1% To col2%
tGrid.col = i%

' Copy Grid to array
If mode% = 1 Then sArray$(j%, i%) = tGrid.Text

' Copy array to Grid
If mode% = 2 Then tGrid.Text = sArray$(j%, i%)

Next i%
Next j%

Exit Sub

' Errors
MiscCopyGridFromToError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscCopyGridFromTo"
ierror = True
Exit Sub

End Sub

Sub MiscSortGeological(tarray() As String, row1 As Integer, row2 As Integer, col1 As Integer, col2 As Integer, ncols As Integer, sample() As TypeSample)
' Sort the passed array into traditional geological order (column basis)

ierror = False
On Error GoTo MiscSortGeologicalError

Dim i As Integer, j As Integer
Dim m As Integer, n As Integer

' Create temp array for sorting
ReDim SortArray(row1% To row2%, col1% To col2%) As String

' Loop through array and sort in geological order
For j% = row1% To row2%
m% = 0
For i% = col1% To col2%
Call MiscGetNextGeologicalOrder2(sample(1).Elsyms$(), ncols%, m%, n%)
If n% > 0 Then SortArray$(j%, i%) = tarray$(j%, n%)
Next i%
Next j%

' Load sorted array back into passed array
For j% = row1% To row2%
For i% = col1% To col2%
tarray$(j%, i%) = SortArray$(j%, i%)
Next i%
Next j%

Exit Sub

' Errors
MiscSortGeologicalError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSortGeological"
ierror = True
Exit Sub

End Sub

Sub MiscGetNextGeologicalOrder(elementarray() As String, num As Integer, m As Integer, n As Integer)
' Returns the index number in the array of the next element in traditional geological order (does *not* handle duplicate elements)
'  elementarray() = list of element symbols
'  num = size of string array
'  m = last geological order index returned
'  n = next element index returned

ierror = False
On Error GoTo MiscGetNextGeologicalOrderError

Dim i As Integer, ip As Integer

' Search through list for next element
If m% = 0 Then m% = 1
For i% = m% To MAXGEOLOGICAL%
ip% = IPOS1%(num%, GeologicalOrder$(i%), elementarray$())
If ip% > 0 Then GoTo 2000
Next i%

' Geological element was not found in element array, return zero for index
n% = 0
Exit Sub

' Return next starting index (if element is not subsequently duplicated)
2000:
m% = i% + 1

' Return next element index
n% = ip%
Exit Sub

' Errors
MiscGetNextGeologicalOrderError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetNextGeologicalOrder"
ierror = True
Exit Sub

End Sub

Sub MiscGetNextGeologicalOrder2(elementarray() As String, num As Integer, m As Integer, n As Integer)
' Returns the index number in the array of the next element in traditional geological order (handles duplicate elements)
'  elementarray() = list of element symbols
'  num = size of string array
'  m = last geological order index returned
'  n = next element index returned

ierror = False
On Error GoTo MiscGetNextGeologicalOrder2Error

Dim i As Integer, ip As Integer, ipp As Integer

' Init flags
If m% = 0 Then
m% = 1
ReDim SortFlags(1 To num%) As Integer
End If

' Search through list for next element (skip elements already sorted using IPOS1DQ())
For i% = m% To MAXGEOLOGICAL%
ip% = IPOS1DQ%(num%, GeologicalOrder$(i%), elementarray$(), SortFlags%())
If ip% > 0 Then GoTo 2000
Next i%

' Geological element was not found in element array, return zero for index
n% = 0
Exit Sub

' Return next starting index
2000:
If Not MiscIsElementDuplicatedSubsequent(ip%, num%, elementarray$(), ipp%) Then m% = i% + 1

' Return next element index and set already sorted flag
n% = ip%
SortFlags%(ip%) = True
Exit Sub

' Errors
MiscGetNextGeologicalOrder2Error:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetNextGeologicalOrder2"
ierror = True
Exit Sub

End Sub

Sub MiscGetNextGeologicalOrderInit()
' Load the geological order array

ierror = False
On Error GoTo MiscGetNextGeologicalOrderInitError

Dim i As Integer, j As Integer, ip As Integer

' Load geology string
If GeologicalSortOrderFlag% = 1 Then
GeologicalOrder$(1) = "nb"
GeologicalOrder$(2) = "ta"
GeologicalOrder$(3) = "mo"
GeologicalOrder$(4) = "w"

GeologicalOrder$(5) = "si"
GeologicalOrder$(6) = "ge"
GeologicalOrder$(7) = "zr"
GeologicalOrder$(8) = "hf"
GeologicalOrder$(9) = "ti"
GeologicalOrder$(10) = "sn"
GeologicalOrder$(11) = "zn"
GeologicalOrder$(12) = "cd"
GeologicalOrder$(13) = "hg"
GeologicalOrder$(14) = "tl"
GeologicalOrder$(15) = "pb"
GeologicalOrder$(16) = "th"
GeologicalOrder$(17) = "u"

GeologicalOrder$(18) = "b"
GeologicalOrder$(19) = "al"
GeologicalOrder$(20) = "ga"
GeologicalOrder$(21) = "in"
GeologicalOrder$(22) = "v"
GeologicalOrder$(23) = "cr"

GeologicalOrder$(24) = "sc"
GeologicalOrder$(25) = "y"
GeologicalOrder$(26) = "la"
GeologicalOrder$(27) = "ce"
GeologicalOrder$(28) = "pr"
GeologicalOrder$(29) = "nd"
GeologicalOrder$(30) = "pm"
GeologicalOrder$(31) = "sm"
GeologicalOrder$(32) = "eu"
GeologicalOrder$(33) = "gd"
GeologicalOrder$(34) = "tb"
GeologicalOrder$(35) = "dy"
GeologicalOrder$(36) = "ho"
GeologicalOrder$(37) = "er"
GeologicalOrder$(38) = "tm"
GeologicalOrder$(39) = "yb"
GeologicalOrder$(40) = "lu"
GeologicalOrder$(41) = "ac"

GeologicalOrder$(42) = "fe"
GeologicalOrder$(43) = "co"
GeologicalOrder$(44) = "ni"
GeologicalOrder$(45) = "cu"
GeologicalOrder$(46) = "mn"

GeologicalOrder$(47) = "be"
GeologicalOrder$(48) = "mg"
GeologicalOrder$(49) = "ca"
GeologicalOrder$(50) = "sr"
GeologicalOrder$(51) = "ba"
GeologicalOrder$(52) = "ra"

GeologicalOrder$(53) = "li"
GeologicalOrder$(54) = "na"
GeologicalOrder$(55) = "k"
GeologicalOrder$(56) = "rb"
GeologicalOrder$(57) = "cs"
GeologicalOrder$(58) = "fr"

GeologicalOrder$(59) = "p"
GeologicalOrder$(60) = "s"
GeologicalOrder$(61) = "as"
GeologicalOrder$(62) = "se"
GeologicalOrder$(63) = "sb"
GeologicalOrder$(64) = "te"

GeologicalOrder$(65) = "he"
GeologicalOrder$(66) = "ne"
GeologicalOrder$(67) = "ar"
GeologicalOrder$(68) = "kr"
GeologicalOrder$(69) = "xe"
GeologicalOrder$(70) = "rn"

GeologicalOrder$(71) = "i"
GeologicalOrder$(72) = "br"
GeologicalOrder$(73) = "cl"
GeologicalOrder$(74) = "f"

GeologicalOrder$(75) = "c"
GeologicalOrder$(76) = "n"
GeologicalOrder$(77) = "o"
GeologicalOrder$(78) = "h"

' Add in remaining elements
For i% = 79 To MAXELM%
For j% = 1 To MAXELM%
ip% = IPOS1%(i%, Symlo$(j%), GeologicalOrder$())
If ip% = 0 Then
GeologicalOrder$(i%) = Symlo$(j%)
Exit For
End If
Next j%
Next i%

' Low to high Z order
ElseIf GeologicalSortOrderFlag% = 2 Then
For i% = 1 To MAXELM%
GeologicalOrder$(i%) = Symlo$(i%)
Next i%

' High to low Z order
ElseIf GeologicalSortOrderFlag% = 3 Then
For i% = 1 To MAXELM%
GeologicalOrder$(i%) = Symlo$(MAXELM% - (i% - 1))
Next i%
End If

Exit Sub

' Errors
MiscGetNextGeologicalOrderInitError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetNextGeologicalOrderInit"
ierror = True
Exit Sub

End Sub
