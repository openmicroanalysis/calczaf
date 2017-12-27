Attribute VB_Name = "CodeMISC6"
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

Dim tfilenumber As Integer

Sub MiscSaveData_PE(gstring As String, xstring As String, ystring As String, tGraph As Pesgo, tForm As Form)
' Open file to save graph data to disk (Pro Esentials code)

ierror = False
On Error GoTo MiscSaveData_PEError

Dim tfilename As String, tmsg As String

' Get filename to save data
tmsg$ = gstring$
tfilename$ = MiscGetFileNameNoExtension(ProbeDataFile$) & "_" & tmsg$ & ".dat"
Call IOGetFileName(Int(1), "DAT", tfilename$, tForm)
If ierror Then Exit Sub

' Open file
Screen.MousePointer = vbHourglass
DoEvents
tfilenumber% = FreeFile()
Open tfilename$ For Output As #tfilenumber%

' Save all data
Call MiscWriteData_PE(xstring$, ystring$, tGraph)
Close #tfilenumber%

If ierror Then Exit Sub
Exit Sub

' Errors
MiscSaveData_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSaveData_PE"
Close #tfilenumber%
ierror = True
Exit Sub

End Sub

Sub MiscSaveDataSets_PE(tbasename As String, gstring As String, xstring As String, ystring As String, sString() As String, tGraph As Pesgo, tForm As Form)
' Open file to save graph data (multiple sets) to disk (Pro Essentials code)

ierror = False
On Error GoTo MiscSaveDataSets_PEError

Dim tfilename As String, tmsg As String

' Get filename to save data
tmsg$ = gstring$
tfilename$ = MiscGetFileNameNoExtension(tbasename$) & "_" & tmsg$ & ".dat"
Call IOGetFileName(Int(1), "DAT", tfilename$, tForm)
If ierror Then Exit Sub

' Open file
Screen.MousePointer = vbHourglass
DoEvents
tfilenumber% = FreeFile()
Open tfilename$ For Output As #tfilenumber%

' Save all data
Call MiscWriteDataSets_PE(gstring$, xstring$, ystring$, sString$(), tGraph)
Close #tfilenumber%
If ierror Then Exit Sub

Screen.MousePointer = vbDefault
Exit Sub

' Errors
MiscSaveDataSets_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSaveDataSets_PE"
Close #tfilenumber%
ierror = True
Exit Sub

End Sub

Sub MiscWriteData_PE(xstring As String, ystring As String, tGraph As Pesgo)
' Write graph data to disk (Pro Essentials code)

ierror = False
On Error GoTo MiscWriteData_PEError

Dim i As Integer
Dim xdata As Single, ydata As Single

If tGraph.points = 0 Or tGraph.Subsets = 0 Then GoTo MiscWriteData_PENoData

' Write column labels
Screen.MousePointer = vbHourglass
msg$ = VbDquote$ & Format$(xstring$, a80$) & VbDquote$ & vbTab & VbDquote$ & Format$(ystring$, a80$) & VbDquote$
Print #tfilenumber%, msg$

' Loop on graphs
For i% = 1 To tGraph.points
ydata! = tGraph.ydata(0, i% - 1)
xdata! = tGraph.xdata(0, i% - 1)
msg$ = MiscAutoFormat$(xdata!) & vbTab & MiscAutoFormat$(ydata!)

' Write to disk
Print #tfilenumber%, msg$
Next i%

Screen.MousePointer = vbDefault
Exit Sub

' Errors
MiscWriteData_PEError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "MiscWriteData_PE"
Close #tfilenumber%
ierror = True
Exit Sub

MiscWriteData_PENoData:
Screen.MousePointer = vbDefault
msg$ = "No graph data to write to disk"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscWriteData_PE"
ierror = True
Exit Sub

End Sub

Sub MiscWriteDataSets_PE(gstring As String, xstring As String, ystring As String, sString$(), tGraph As Pesgo)
' Write graph data to disk (multiple sets) (Pro Essentials code)

ierror = False
On Error GoTo MiscWriteDataSets_PEError

Dim i As Integer, j As Integer
Dim xdata As Single, ydata As Single

If tGraph.points = 0 Or tGraph.Subsets = 0 Then GoTo MiscWriteDataSets_PENoData

' Write y-data title if multiple data sets
Screen.MousePointer = vbHourglass
If tGraph.Subsets > 1 Then
msg$ = VbDquote$ & Format$(ystring$, a80$) & VbDquote$
Print #tfilenumber%, msg$

' Write column labels
msg$ = vbNullString
For j% = 1 To tGraph.Subsets
msg$ = msg$ & VbDquote$ & Format$(xstring$, a80$) & VbDquote$ & vbTab & VbDquote$ & Format$(sString$(j%), a80$) & VbDquote$ & vbTab
Next j%
Print #tfilenumber%, msg$

' Single data set
Else
msg$ = VbDquote$ & gstring$ & VbDquote$
Print #tfilenumber%, msg$
msg$ = VbDquote$ & Format$(xstring$, a80$) & VbDquote$ & vbTab & VbDquote$ & Format$(ystring$, a80$) & VbDquote$ & vbTab
Print #tfilenumber%, msg$
End If

' Loop on graphs
For i% = 1 To tGraph.points
msg$ = vbNullString

For j% = 1 To tGraph.Subsets
ydata! = tGraph.ydata(j% - 1, i% - 1)
xdata! = tGraph.xdata(j% - 1, i% - 1)
msg$ = msg$ & MiscAutoFormat$(xdata!) & vbTab & MiscAutoFormat$(ydata!) & vbTab
Next j%

' Write to disk
Print #tfilenumber%, msg$
Next i%

Screen.MousePointer = vbDefault
Exit Sub

' Errors
MiscWriteDataSets_PEError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "MiscWriteDataSets_PE"
Close #tfilenumber%
ierror = True
Exit Sub

MiscWriteDataSets_PENoData:
Screen.MousePointer = vbDefault
msg$ = "No graph data to write to disk"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscWriteDataSets_PE"
ierror = True
Exit Sub

End Sub
