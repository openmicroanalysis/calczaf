Attribute VB_Name = "CodeMISC6"
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

Sub MiscSaveData(gstring As String, xstring As String, ystring As String, tGraph As Graph, tForm As Form)
' Open file to save graph data to disk

ierror = False
On Error GoTo MiscSaveDataError

Dim tfilename As String, tmsg As String

' Get filename to save data
tmsg$ = gstring$
tfilename$ = MiscGetFileNameNoExtension(ProbeDataFile$) & "_" & tmsg$ & ".dat"
Call IOGetFileName(Int(1), "DAT", tfilename$, tForm)
If ierror Then Exit Sub

' Open file
Close #Temp1FileNumber%
Screen.MousePointer = vbHourglass
DoEvents
Open tfilename$ For Output As #Temp1FileNumber%

' Save all data
Call MiscWriteData(xstring$, ystring$, tGraph)
Close #Temp1FileNumber%

If ierror Then Exit Sub
Exit Sub

' Errors
MiscSaveDataError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSaveData"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub MiscSaveDataSets(tbasename As String, gstring As String, xstring As String, ystring As String, sString() As String, tGraph As Graph, tForm As Form)
' Open file to save graph data (multiple sets) to disk

ierror = False
On Error GoTo MiscSaveDataSetsError

Dim tfilename As String, tmsg As String

' Get filename to save data
tmsg$ = gstring$
tfilename$ = MiscGetFileNameNoExtension(tbasename$) & "_" & tmsg$ & ".dat"
Call IOGetFileName(Int(1), "DAT", tfilename$, tForm)
If ierror Then Exit Sub

' Open file
Close #Temp1FileNumber%
Screen.MousePointer = vbHourglass
DoEvents
Open tfilename$ For Output As #Temp1FileNumber%

' Save all data
Call MiscWriteDataSets(gstring$, xstring$, ystring$, sString$(), tGraph)
Close #Temp1FileNumber%
If ierror Then Exit Sub

Screen.MousePointer = vbDefault
Exit Sub

' Errors
MiscSaveDataSetsError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSaveDataSets"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub MiscWriteData(xstring As String, ystring As String, tGraph As Graph)
' Write graph data to disk

ierror = False
On Error GoTo MiscWriteDataError

Dim i As Integer
Dim xdata As Single, ydata As Single

If tGraph.NumPoints = 0 Or tGraph.NumSets = 0 Then GoTo MiscWriteDataNoData

' Write column labels
Screen.MousePointer = vbHourglass
msg$ = VbDquote$ & Format$(xstring$, a80$) & VbDquote$ & vbTab & VbDquote$ & Format$(ystring$, a80$) & VbDquote$
Print #Temp1FileNumber%, msg$

' Loop on graphs
For i% = 1 To tGraph.NumPoints
tGraph.ThisSet = 1
tGraph.ThisPoint = i%
ydata! = tGraph.GraphData
tGraph.ThisPoint = i%
xdata! = tGraph.XPosData
msg$ = MiscAutoFormat$(xdata!) & vbTab & MiscAutoFormat$(ydata!)

' Write to disk
Print #Temp1FileNumber%, msg$
Next i%

Screen.MousePointer = vbDefault
Exit Sub

' Errors
MiscWriteDataError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "MiscWriteData"
Close #Temp1FileNumber%
ierror = True
Exit Sub

MiscWriteDataNoData:
Screen.MousePointer = vbDefault
msg$ = "No graph data to write to disk"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscWriteData"
ierror = True
Exit Sub

End Sub

Sub MiscWriteDataSets(gstring As String, xstring As String, ystring As String, sString$(), tGraph As Graph)
' Write graph data to disk (multiple sets)

ierror = False
On Error GoTo MiscWriteDataSetsError

Dim i As Integer, j As Integer
Dim xdata As Single, ydata As Single

If tGraph.NumPoints = 0 Or tGraph.NumSets = 0 Then GoTo MiscWriteDataSetsNoData

' Write y-data title if multiple data sets
Screen.MousePointer = vbHourglass
If tGraph.NumSets > 1 Then
msg$ = VbDquote$ & Format$(ystring$, a80$) & VbDquote$
Print #Temp1FileNumber%, msg$

' Write column labels
msg$ = vbNullString
For j% = 1 To tGraph.NumSets
msg$ = msg$ & VbDquote$ & Format$(xstring$, a80$) & VbDquote$ & vbTab & VbDquote$ & Format$(sString$(j%), a80$) & VbDquote$ & vbTab
Next j%
Print #Temp1FileNumber%, msg$

' Single data set
Else
msg$ = VbDquote$ & gstring$ & VbDquote$
Print #Temp1FileNumber%, msg$
msg$ = VbDquote$ & Format$(xstring$, a80$) & VbDquote$ & vbTab & VbDquote$ & Format$(ystring$, a80$) & VbDquote$ & vbTab
Print #Temp1FileNumber%, msg$
End If

' Loop on graphs
For i% = 1 To tGraph.NumPoints
msg$ = vbNullString
tGraph.ThisPoint = i%

For j% = 1 To tGraph.NumSets
tGraph.ThisSet = j%
ydata! = tGraph.GraphData
xdata! = tGraph.XPosData
msg$ = msg$ & MiscAutoFormat$(xdata!) & vbTab & MiscAutoFormat$(ydata!) & vbTab
Next j%

' Write to disk
Print #Temp1FileNumber%, msg$
Next i%

Screen.MousePointer = vbDefault
Exit Sub

' Errors
MiscWriteDataSetsError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "MiscWriteDataSets"
Close #Temp1FileNumber%
ierror = True
Exit Sub

MiscWriteDataSetsNoData:
Screen.MousePointer = vbDefault
msg$ = "No graph data to write to disk"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscWriteDataSets"
ierror = True
Exit Sub

End Sub


