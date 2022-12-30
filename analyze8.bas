Attribute VB_Name = "CodeAnalyze8"
' (c) Copyright 1995-2023 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

' Array of filenames for output
Dim arraysize As Integer
Dim filenamearray() As String

Dim MatrixAverages(1 To MAXCHAN% + 1, 1 To MAXZAF%) As Single       ' wt%
Dim MatrixAverages2(1 To MAXCHAN% + 1, 1 To MAXZAF%) As Single      ' at %

Sub AnalyzeChangeZAF(analysis As TypeAnalysis, sample() As TypeSample, stdsample() As TypeSample)
' Change ZAF selections for calculating all matrix corrections

ierror = False
On Error GoTo AnalyzeChangeZAFError

' Load individual selections
Call InitGetZAFSetZAF2(izaf%)
If ierror Then Exit Sub

' Print current ZAF selections
Call TypeZAFSelections
If ierror Then Exit Sub

' Load element arrays
Call ElementGetData(sample())
If ierror Then Exit Sub

' Load primary intensities (0 = phi/rho/z, 1,2,3,4 = alpha fits, 5 = calilbration curve, 6 = fundamental parameters)
If CorrectionFlag% <> MAXCORRECTION% Then
Call ZAFSetZAF(sample())
If ierror Then Exit Sub
Else
'Call ZAFSetZAF3(sample())
'If ierror Then Exit Sub
End If

' Update the standard kfacs based on changed conditions
Call UpdateAllStdKfacs(analysis, sample(), stdsample())
If ierror Then Exit Sub

Exit Sub

' Errors
AnalyzeChangeZAFError:
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeChangeZAF"
ierror = True
Exit Sub

End Sub

Sub AnalyzeCalculateAllSave(afilename As String, nsams As Integer, nstring As String, analysis As TypeAnalysis, sample() As TypeSample, tForm As Form)
' Save the results from all matrix corrections

ierror = False
On Error GoTo AnalyzeCalculateAllSaveError

Dim sampleline As Integer
Dim astring As String, lstring As String, tmsg As String
Dim tfilename As String, bstring As String

' Reset array size if first correction
If izaf% = 1 Then arraysize% = 0

' Load filename from database, sample name and zaf correction
tmsg$ = nstring$ & "_" & zafstring$(izaf%)
Call MiscModifyStringToFilename(tmsg$)
If ierror Then Exit Sub

' Create file name for output
tfilename$ = MiscGetFileNameNoExtension$(afilename$) & "-" & tmsg$ & ".dat"

' Store filename for subsequent Excel loading
Close (Temp1FileNumber%)                                ' close in case file was left open
arraysize% = arraysize% + 1
ReDim Preserve filenamearray$(1 To arraysize%)
filenamearray$(arraysize%) = tfilename$                 ' already includes full path
Open tfilename$ For Output As #Temp1FileNumber%

' Write header
Print #Temp1FileNumber%, VbDquote$ & ProbeDataFile$ & VbDquote$
Print #Temp1FileNumber%, VbDquote$ & MDBUserName$ & VbDquote$
Print #Temp1FileNumber%, VbDquote$ & MDBFileTitle$ & VbDquote$
Print #Temp1FileNumber%, VbDquote$ & MDBFileDescription$ & VbDquote$
Print #Temp1FileNumber%, VbDquote$ & "Nominal Beam: " & Format$(NominalBeam!) & VbDquote$

msg$ = macstring$(MACTypeFlag%)
Print #Temp1FileNumber%, VbDquote$ & msg$ & VbDquote$
Print #Temp1FileNumber%, vbNullString

' Output column labels
Call AnalyzeCalculateAllColumnString(lstring$, sample())
If ierror Then Exit Sub
Print #Temp1FileNumber%, lstring$

' Output each line
For sampleline% = 1 To sample(1).Datarows%
If sample(1).LineStatus(sampleline%) Then

' Output column data
Call AnalyzeCalculateAllDataString(astring$, sampleline%, analysis, sample())
If ierror Then Exit Sub
Print #Temp1FileNumber%, astring$

End If
Next sampleline%

' Calculate averages (bstring used below)
Call AnalyzeCalculateAllAverage(astring$, bstring$, analysis, sample())
If ierror Then Exit Sub
Print #Temp1FileNumber%, vbNullString
Print #Temp1FileNumber%, lstring$
Print #Temp1FileNumber%, astring$

Close (Temp1FileNumber%)

' Output just the average compositions only to separate ASCII file
tmsg$ = nstring$ & "_All Corrections"
Call MiscModifyStringToFilename(tmsg$)
If ierror Then Exit Sub

' Create filename for output
tfilename$ = MiscGetFileNameNoExtension$(afilename$) & "-" & tmsg$ & ".dat"

' Check if first time
If izaf% = 1 Then
Open tfilename$ For Output As #Temp1FileNumber%
Else
Open tfilename$ For Append As #Temp1FileNumber%
End If

' Write header
If izaf% = 1 Then
Print #Temp1FileNumber%, VbDquote$ & ProbeDataFile$ & VbDquote$
Print #Temp1FileNumber%, VbDquote$ & MDBUserName$ & VbDquote$
Print #Temp1FileNumber%, VbDquote$ & MDBFileTitle$ & VbDquote$
Print #Temp1FileNumber%, VbDquote$ & MDBFileDescription$ & VbDquote$
Print #Temp1FileNumber%, VbDquote$ & "Nominal Beam: " & Format$(NominalBeam!) & VbDquote$

msg$ = macstring$(MACTypeFlag%)
Print #Temp1FileNumber%, VbDquote$ & msg$ & VbDquote$
Print #Temp1FileNumber%, "All Matrix Correction Averages Only"
Print #Temp1FileNumber%, vbNullString
End If

If izaf% = 1 Then
Print #Temp1FileNumber%, lstring$
End If

' Print numerical results with zafstring (do not need tab before bstring$)
Print #Temp1FileNumber%, VbDquote$ & zafstring(izaf%) & VbDquote$ & bstring$

' Store filename for subsequent Excel loading
If izaf% = MAXZAF% Then
arraysize% = arraysize% + 1
ReDim Preserve filenamearray$(1 To arraysize%)
filenamearray$(arraysize%) = tfilename$ ' already includes full path
End If

Close (Temp1FileNumber%)

' Confirm output files with user
If izaf% = MAXZAF% Then
If nsams% = 1 Then
tmsg$ = "All Matrix Corrections sample data were output to ASCII data files based on sample names"
MsgBox tmsg$, vbOKOnly + vbInformation, "AnalyzeCalculateAllSave"
End If

' Output to excel spreadsheet
Call OutputSaveCustom2SendToExcel(Int(0), arraysize%, filenamearray$(), tForm)
If ierror Then Exit Sub

' Print averages to log window
Call AnalyzeCalculateAllAverageOutput(nstring$, sample())
If ierror Then Exit Sub

End If

Exit Sub

' Errors
AnalyzeCalculateAllSaveError:
Close (Temp1FileNumber%)
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeCalculateAllSave"
ierror = True
Exit Sub

End Sub

Sub AnalyzeCalculateAllAverageOutput(nstring As String, sample() As TypeSample)
' Calculate average and output to log window

ierror = False
On Error GoTo AnalyzeCalculateAllAverageOutputError

Dim j As Integer, chan As Integer

Dim average As TypeAverage

' Print MAC and sample name
msg$ = "Summary of All Calculated (averaged) Matrix Corrections:"
Call IOWriteLog(vbCrLf & msg$)

msg$ = nstring$
Call IOWriteLog(msg$)

msg$ = macstring$(MACTypeFlag%)
Call IOWriteLog(msg$)

' Output wt% averages
msg$ = "Elemental Weight Percents:"
Call IOWriteLog(vbCrLf & msg$)
msg$ = "ELEM: "
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(sample(1).Elsyup$(chan%), a80$)
Next chan%
msg$ = msg$ & Format$("TOTAL", a80$)
Call IOWriteLog(msg$)

For j% = 1 To MAXZAF%
msg$ = Format$(j%, a60$)
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(MatrixAverages!(chan%, j%), f83$), a80$)
Next chan%
msg$ = msg$ & Format$(Format$(MatrixAverages!(sample(1).LastChan% + 1, j%), f83$), a80$)
Call IOWriteLog(msg$ & "   " & zafstring$(j%))
Next j%

' Output standard deviation and standard error
Call MathArrayAverage3(average, MatrixAverages!(), CLng(MAXZAF%), sample(1).LastChan% + 1)
If ierror Then Exit Sub

msg$ = vbCrLf & "AVER: "
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(average.averags!(chan%), f83$), a80$)
Next chan%
msg$ = msg$ & Format$(Format$(average.averags!(sample(1).LastChan% + 1), f83$), a80$)
Call IOWriteLog(msg$)

msg$ = "SDEV: "
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(average.Stddevs!(chan%), f83$), a80$)
Next chan%
msg$ = msg$ & Format$(Format$(average.Stddevs!(sample(1).LastChan% + 1), f83$), a80$)
Call IOWriteLog(msg$)

msg$ = "SERR: "
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(average.Stderrs!(chan%), f83$), a80$)
Next chan%
Call IOWriteLog(msg$)

msg$ = vbCrLf & "MIN:  "
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(average.Minimums!(chan%), f83$), a80$)
Next chan%
msg$ = msg$ & Format$(Format$(average.Minimums!(sample(1).LastChan% + 1), f83$), a80$)
Call IOWriteLog(msg$)

msg$ = "MAX:  "
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(average.Maximums!(chan%), f83$), a80$)
Next chan%
msg$ = msg$ & Format$(Format$(average.Maximums!(sample(1).LastChan% + 1), f83$), a80$)
Call IOWriteLog(msg$)

' Output at% averages
If sample(1).AtomicPercentFlag% Then
msg$ = "Atomic Percents:"
Call IOWriteLog(vbCrLf & msg$)
msg$ = "ELEM: "
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(sample(1).Elsyup$(chan%), a80$)
Next chan%
msg$ = msg$ & Format$("TOTAL", a80$)
Call IOWriteLog(msg$)

For j% = 1 To MAXZAF%
msg$ = Format$(j%, a60$)
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(MatrixAverages2!(chan%, j%), f83$), a80$)
Next chan%
msg$ = msg$ & Format$(Format$(MatrixAverages2!(sample(1).LastChan% + 1, j%), f83$), a80$)
Call IOWriteLog(msg$ & "   " & zafstring$(j%))
Next j%

' Output standard deviation and standard error
Call MathArrayAverage3(average, MatrixAverages2!(), CLng(MAXZAF%), sample(1).LastChan% + 1)
If ierror Then Exit Sub

msg$ = vbCrLf & "AVER: "
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(average.averags!(chan%), f83$), a80$)
Next chan%
msg$ = msg$ & Format$(Format$(average.averags!(sample(1).LastChan% + 1), f83$), a80$)
Call IOWriteLog(msg$)

msg$ = "SDEV: "
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(average.Stddevs!(chan%), f83$), a80$)
Next chan%
msg$ = msg$ & Format$(Format$(average.Stddevs!(sample(1).LastChan% + 1), f83$), a80$)
Call IOWriteLog(msg$)

msg$ = "SERR: "
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(average.Stderrs!(chan%), f83$), a80$)
Next chan%
Call IOWriteLog(msg$)

msg$ = vbCrLf & "MIN:  "
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(average.Minimums!(chan%), f83$), a80$)
Next chan%
msg$ = msg$ & Format$(Format$(average.Minimums!(sample(1).LastChan% + 1), f83$), a80$)
Call IOWriteLog(msg$)

msg$ = "MAX:  "
For chan% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$(average.Maximums!(chan%), f83$), a80$)
Next chan%
msg$ = msg$ & Format$(Format$(average.Maximums!(sample(1).LastChan% + 1), f83$), a80$)
Call IOWriteLog(msg$)
End If

Exit Sub

' Errors
AnalyzeCalculateAllAverageOutputError:
Close (Temp1FileNumber%)
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeCalculateAllAverageOutput"
ierror = True
Exit Sub

End Sub

Sub AnalyzeCalculateAllAverage(astring As String, bstring As String, analysis As TypeAnalysis, sample() As TypeSample)
' Calculate average ZAF for file output (CalculateAllMatrixCorrections)

ierror = False
On Error GoTo AnalyzeCalculateAllAverageError

Dim chan As Integer

Dim average2 As TypeAverage     ' for weight percents
Dim average3 As TypeAverage     ' for atomic percents

' Line number
astring$ = "AVER:"
bstring$ = vbNullString

' Calculate weight percent average
Call MathArrayAverage(average2, analysis.WtsData!(), sample(1).Datarows%, sample(1).LastChan% + 1, sample())
If ierror Then Exit Sub

' Calculate atomic percent average
Call MathArrayAverage(average3, analysis.CalData!(), sample(1).Datarows%, sample(1).LastChan% + 1, sample())
If ierror Then Exit Sub

' Elemental data
For chan% = 1 To sample(1).LastChan%
astring$ = astring$ & vbTab & MiscAutoFormat$(average2.averags!(chan%))
bstring$ = bstring$ & vbTab & MiscAutoFormat$(average2.averags!(chan%))

' Save average for each channel
MatrixAverages!(chan%, izaf%) = average2.averags!(chan%)
Next chan%

' Add total average
astring$ = astring$ & vbTab & MiscAutoFormat$(average2.averags!(sample(1).LastChan% + 1))
bstring$ = bstring$ & vbTab & MiscAutoFormat$(average2.averags!(sample(1).LastChan% + 1))

' Save total average
MatrixAverages!(sample(1).LastChan% + 1, izaf%) = average2.averags!(sample(1).LastChan% + 1)

' Other data
If sample(1).AtomicPercentFlag% Then
For chan% = 1 To sample(1).LastChan%
astring$ = astring$ & vbTab & MiscAutoFormat$(average3.averags!(chan%))
bstring$ = bstring$ & vbTab & MiscAutoFormat$(average3.averags!(chan%))

' Save average
MatrixAverages2!(chan%, izaf%) = average3.averags!(chan%)
Next chan%

' Add total
astring$ = astring$ & vbTab & MiscAutoFormat$(average3.averags!(sample(1).LastChan% + 1))
bstring$ = bstring$ & vbTab & MiscAutoFormat$(average3.averags!(sample(1).LastChan% + 1))

' Save total
MatrixAverages2!(sample(1).LastChan% + 1, izaf%) = average3.averags!(sample(1).LastChan% + 1)
End If

Exit Sub

' Errors
AnalyzeCalculateAllAverageError:
Close (Temp1FileNumber%)
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeCalculateAllAverage"
ierror = True
Exit Sub

End Sub

Sub AnalyzeCalculateAllColumnString(astring As String, sample() As TypeSample)
' Create a column label string

ierror = False
On Error GoTo AnalyzeCalculateAllColumnStringError

Dim i As Integer

' Line number
astring$ = VbDquote$ & Format$("LINE", a80$) & VbDquote$

' Elemental labels
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & vbTab & VbDquote$ & Format$(sample(1).Elsyup$(i%) & " WT%", a80$) & VbDquote$
Next i%

' Add total label
astring$ = astring$ & vbTab & VbDquote$ & Format$("TOTAL", a80$) & VbDquote$

' Atomic labels
If sample(1).AtomicPercentFlag% Then
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & vbTab & VbDquote$ & Format$(sample(1).Elsyup$(i%) & " AT%", a80$) & VbDquote$
Next i%

' Add total label
astring$ = astring$ & vbTab & VbDquote$ & Format$("TOTAL", a80$) & VbDquote$
End If

Exit Sub

' Errors
AnalyzeCalculateAllColumnStringError:
Close (Temp1FileNumber%)
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeCalculateAllColumnString"
ierror = True
Exit Sub

End Sub

Sub AnalyzeCalculateAllDataString(astring As String, sampleline As Integer, analysis As TypeAnalysis, sample() As TypeSample)
' Create a data line string for each line

ierror = False
On Error GoTo AnalyzeCalculateAllDataStringError

Dim i As Integer

' Line number
astring$ = Format$(sample(1).Linenumber&(sampleline%), a80$)

' Elemental data
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & vbTab & MiscAutoFormat$(analysis.WtsData!(sampleline%, i%))
Next i%

' Add total
astring$ = astring$ & vbTab & MiscAutoFormat$(analysis.WtsData!(sampleline%, sample(1).LastChan% + 1))

' Atomic data
If sample(1).AtomicPercentFlag% Then
For i% = 1 To sample(1).LastChan%
astring$ = astring$ & vbTab & MiscAutoFormat$(analysis.CalData!(sampleline%, i%))
Next i%

' Add total
astring$ = astring$ & vbTab & MiscAutoFormat$(analysis.CalData!(sampleline%, sample(1).LastChan% + 1))
End If

Exit Sub

' Errors
AnalyzeCalculateAllDataStringError:
Close (Temp1FileNumber%)
MsgBox Error$, vbOKOnly + vbCritical, "AnalyzeCalculateAllDataString"
ierror = True
Exit Sub

End Sub


