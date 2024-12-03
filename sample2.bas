Attribute VB_Name = "CodeSAMPLE2"
' (c) Copyright 1995-2024 by John J. Donovan
Option Explicit

Sub SampleCheckSampleList(tList As ListBox, nsams As Integer, samplerow As Integer)
' Routine to check that at least one sample in tList is selected. Returns
' the number of samples selected and the row number of the last selected sample.

ierror = False
On Error GoTo SampleCheckSampleListError

Dim i As Integer

' Go through all samples in list
nsams% = 0
samplerow% = 0
For i% = 0 To tList.ListCount - 1
If tList.Selected(i%) Then
nsams% = nsams% + 1
samplerow% = tList.ItemData(i%)
End If
Next i%

' If no sample selected then, give error message
If nsams% = 0 Then
msg$ = "No sample(s) was selected" & vbCrLf & vbCrLf
msg$ = msg$ & "Make sure the Analyze! window is open and one or more samples are selected."
MsgBox msg$, vbOKOnly + vbExclamation, "SampleCheckSampleList"
ierror = True
End If

Exit Sub

' Errors
SampleCheckSampleListError:
MsgBox Error$, vbOKOnly + vbCritical, "SampleCheckSampleList"
ierror = True
Exit Sub

End Sub

Sub SampleCheckSampleList2(tList As ListBox, nsams As Integer, samplerow As Integer)
' Routine to check that at least one sample in tList is selected. Returns
' the number of samples selected and the row number of the first selected sample.

ierror = False
On Error GoTo SampleCheckSampleList2Error

Dim i As Integer

' Go through all samples in list (last to first)
nsams% = 0
samplerow% = 0
For i% = tList.ListCount - 1 To 0 Step -1
If tList.Selected(i%) Then
nsams% = nsams% + 1
samplerow% = tList.ItemData(i%)
End If
Next i%

' If no sample selected then, give error message
If nsams% = 0 Then
msg$ = "No sample(s) was selected" & vbCrLf & vbCrLf
msg$ = msg$ & "Make sure the Analyze! window is open and one or more samples are selected."
MsgBox msg$, vbOKOnly + vbExclamation, "SampleCheckSampleList2"
ierror = True
End If

Exit Sub

' Errors
SampleCheckSampleList2Error:
MsgBox Error$, vbOKOnly + vbCritical, "SampleCheckSampleList2"
ierror = True
Exit Sub

End Sub

Function SampleGetRow(sample() As TypeSample) As Integer
' Returns the sample row of the passed sample setup

ierror = False
On Error GoTo SampleGetRowError

Dim i As Integer

For i% = 1 To NumberofSamples%

If sample(1).number% <> SampleNums%(i%) Then GoTo 1000
If sample(1).Type% <> SampleTyps%(i%) Then GoTo 1000
If sample(1).Set% <> SampleSets%(i%) Then GoTo 1000
SampleGetRow% = i%
Exit Function

1000:  Next i%

SampleGetRow% = 0
Exit Function

' Errors
SampleGetRowError:
MsgBox Error$, vbOKOnly + vbCritical, "SampleGetRow"
ierror = True
Exit Function

End Function

Function SampleGetRow2(inum As Integer, ityp As Integer, iset As Integer) As Integer
' Returns the sample row of the passed sample type (alternative to SampleGetRow())

ierror = False
On Error GoTo SampleGetRow2Error

Dim i As Integer

For i% = 1 To NumberofSamples%

If inum% <> SampleNums%(i%) Then GoTo 2000
If ityp% <> SampleTyps%(i%) Then GoTo 2000
If iset% <> SampleSets%(i%) Then GoTo 2000
SampleGetRow2 = i%
Exit Function

2000:  Next i%

SampleGetRow2 = 0
Exit Function

' Errors
SampleGetRow2Error:
MsgBox Error$, vbOKOnly + vbCritical, "SampleGetRow2"
ierror = True
Exit Function

End Function

Function SampleGetString(samplerow As Integer) As String
' Returns a sample type, number, set and name string based on the sample row number

ierror = False
On Error GoTo SampleGetStringError

Dim tmsg As String, achar As String

' Determine if sample contains no or all deleted lines
If samplerow% > 0 Then
achar$ = " "
If SampleDels%(samplerow%) Then
achar$ = " * "
End If

' Load string based on sample type
SampleGetString = vbNullString
If samplerow% = 0 Then Exit Function
If SampleSets%(samplerow%) > 0 Then
If SampleTyps%(samplerow%) = 1 Then tmsg$ = "St " & Format$(SampleNums%(samplerow%), a40) & " Set " & Format$(SampleSets%(samplerow%), a30) & achar$ & SampleNams$(samplerow%)
Else
If SampleTyps%(samplerow%) = 1 Then tmsg$ = "St " & Format$(SampleNums%(samplerow%), a40) & SampleNams$(samplerow%)
End If
If SampleTyps%(samplerow%) = 2 Then tmsg$ = "Un " & Format$(SampleNums%(samplerow%), a40) & " " & achar$ & SampleNams$(samplerow%)
If SampleTyps%(samplerow%) = 3 Then tmsg$ = "Wa " & Format$(SampleNums%(samplerow%), a40) & " " & achar$ & SampleNams$(samplerow%)
End If

SampleGetString = tmsg$
Exit Function

' Errors
SampleGetStringError:
MsgBox Error$, vbOKOnly + vbCritical, "SampleGetString"
ierror = True
Exit Function

End Function

Function SampleGetString3(samplerow As Integer) As String
' Returns a sample type, number, set (but no name string) based on the sample row number

ierror = False
On Error GoTo SampleGetString3Error

Dim tmsg As String

' Load string based on sample type
SampleGetString3 = vbNullString
If samplerow% = 0 Then Exit Function
If SampleSets%(samplerow%) > 0 Then
If SampleTyps%(samplerow%) = 1 Then tmsg$ = "St " & Format$(SampleNums%(samplerow%), a40) & " Set " & Format$(SampleSets%(samplerow%), a30)
Else
If SampleTyps%(samplerow%) = 1 Then tmsg$ = "St " & Format$(SampleNums%(samplerow%), a40)
End If
If SampleTyps%(samplerow%) = 2 Then tmsg$ = "Un " & Format$(SampleNums%(samplerow%), a40)
If SampleTyps%(samplerow%) = 3 Then tmsg$ = "Wa " & Format$(SampleNums%(samplerow%), a40)

SampleGetString3 = tmsg$
Exit Function

' Errors
SampleGetString3Error:
MsgBox Error$, vbOKOnly + vbCritical, "SampleGetString3"
ierror = True
Exit Function

End Function

Function SampleGetString2(sample() As TypeSample) As String
' Returns a sample type, number, set and name string based on the sample

ierror = False
On Error GoTo SampleGetString2Error

Dim tmsg As String, achar As String

' Determine if sample contains no or all deleted lines
If sample(1).GoodDataRows% > 0 Then
achar$ = " "
Else
achar$ = " * "
End If

' Load string based on sample type
SampleGetString2 = vbNullString
If sample(1).Set% > 0 Then
If sample(1).Type% = 1 Then tmsg$ = "St " & Format$(sample(1).number%, a40) & " Set " & Format$(sample(1).Set%, a30) & achar$ & sample(1).Name$
Else
If sample(1).Type% = 1 Then tmsg$ = "St " & Format$(sample(1).number%, a40) & " " & sample(1).Name$
End If
If sample(1).Type% = 2 Then tmsg$ = "Un " & Format$(sample(1).number%, a40) & " " & achar$ & sample(1).Name$
If sample(1).Type% = 3 Then tmsg$ = "Wa " & Format$(sample(1).number%, a40) & " " & achar$ & sample(1).Name$

SampleGetString2 = tmsg$
Exit Function

' Errors
SampleGetString2Error:
MsgBox Error$, vbOKOnly + vbCritical, "SampleGetString2"
ierror = True
Exit Function

End Function

Function SampleGetLast(mode As Integer) As Integer
' Returns the row number of the last sample of the type specified

ierror = False
On Error GoTo SampleGetLastError

Dim row As Integer

' Loop through all sample starting from last
For row% = NumberofSamples% To 1 Step -1
If mode% = SampleTyps%(row%) Then
SampleGetLast% = row%
Exit Function
End If
Next row%

SampleGetLast% = NumberofSamples%
Exit Function

' Errors
SampleGetLastError:
MsgBox Error$, vbOKOnly + vbCritical, "SampleGetLast"
ierror = True
Exit Function

End Function
