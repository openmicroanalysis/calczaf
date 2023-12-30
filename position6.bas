Attribute VB_Name = "CodePOSITION6"
' (c) Copyright 1995-2024 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub PositionGetXYZ(sampletype As Integer, npts As Long, xdata() As Single, ydata() As Single, zdata() As Single, iData() As Integer, ndata() As Integer, sdata() As Integer, sndata() As String)
' Routine to load position data (x,y,z only) from the POSITION.MDB database based on sample type
'   sampletype = 0 load all
'   sampletype = 1 load standards
'   sampletype = 2 load unknowns
'   sampletype = 3 load wavescans

ierror = False
On Error GoTo PositionGetXYZError

Dim i As Long
Dim SQLQ As String

Dim PoDb As Database
Dim PoRs As Recordset

' Open the database
If PositionDataFile$ = vbNullString Then PositionDataFile$ = ApplicationCommonAppData$ & "POSITION.MDB"
Set PoDb = OpenDatabase(PositionDataFile$, PositionDatabaseNonExclusiveAccess%, dbReadOnly)

' Get position data from "Position" table
If sampletype% > 0 Then
SQLQ$ = "SELECT DISTINCT Position.*, Sample.* from Position, Sample WHERE Types = " & Str$(sampletype%) & " "
SQLQ$ = SQLQ$ & "AND Position.PosToRow = Sample.RowOrder "
SQLQ$ = SQLQ$ & "ORDER by Types, Numbers, PosToRow"
Else
SQLQ$ = "SELECT DISTINCT Position.*, Sample.* from Position, Sample WHERE Types <> 0 "
SQLQ$ = SQLQ$ & "AND Position.PosToRow = Sample.RowOrder "
SQLQ$ = SQLQ$ & "ORDER by Types, Numbers, PosToRow"
End If
Set PoRs = PoDb.OpenRecordset(SQLQ$, dbOpenSnapshot)

' Exit if no data
If PoRs.BOF And PoRs.EOF Then
npts& = 0
Exit Sub
End If

' Load number of points
PoRs.MoveLast
npts& = PoRs.RecordCount
PoRs.MoveFirst

ReDim xdata(1 To npts&) As Single
ReDim ydata(1 To npts&) As Single
ReDim zdata(1 To npts&) As Single
ReDim iData(1 To npts&) As Integer  ' types
ReDim ndata(1 To npts&) As Integer  ' line (row) numbers
ReDim sdata(1 To npts&) As Integer  ' sample numbers
ReDim sndata(1 To npts&) As String  ' sample names

' Load position data
i& = 0
Do Until PoRs.EOF
i& = i& + 1
xdata!(i&) = PoRs("StageX")
ydata!(i&) = PoRs("StageY")
zdata!(i&) = PoRs("StageZ")
iData%(i&) = PoRs("Types")
ndata%(i&) = PoRs("PosOrder")
sdata%(i&) = PoRs("Numbers")
sndata$(i&) = Trim$(vbNullString & PoRs("Names"))
PoRs.MoveNext
Loop

PoRs.Close
PoDb.Close

Exit Sub

' Errors
PositionGetXYZError:
MsgBox Error$, vbOKOnly + vbCritical, "PositionGetXYZ"
ierror = True
Exit Sub

End Sub

Sub PositionGetSampleDataOnly(samplerow As Integer, npts As Integer, xdata() As Single, ydata() As Single, zdata() As Single, iData() As Integer)
' Routine to load position data (x, y, z only) from the POSITION.MDB database

ierror = False
On Error GoTo PositionGetSampleDataOnlyError

Dim i As Integer
Dim SQLQ As String

Dim position As TypePosition

Dim PoDb As Database
Dim PoRs As Recordset

' Open the database
Screen.MousePointer = vbHourglass
If PositionDataFile$ = vbNullString Then PositionDataFile$ = ApplicationCommonAppData$ & "POSITION.MDB"
Set PoDb = OpenDatabase(PositionDataFile$, PositionDatabaseNonExclusiveAccess%, dbReadOnly)

' Get position data from "Position" table
SQLQ$ = "SELECT Position.* from Position WHERE PosToRow = " & Str$(samplerow%) & " "
SQLQ$ = SQLQ$ & "ORDER by PosOrder"
Set PoRs = PoDb.OpenRecordset(SQLQ$, dbOpenSnapshot)

' Update grid for new number of rows
If PoRs.BOF And PoRs.EOF Then GoTo PositionGetSampleDataOnlyNoPositions

' Load number of points
PoRs.MoveLast
npts% = PoRs.RecordCount
PoRs.MoveFirst

ReDim xdata(1 To npts%) As Single
ReDim ydata(1 To npts%) As Single
ReDim zdata(1 To npts%) As Single
ReDim iData(1 To npts%) As Integer

' Load position data
i% = 0
Do Until PoRs.EOF
i% = i% + 1
xdata!(i%) = PoRs("StageX")
ydata!(i%) = PoRs("StageY")
zdata!(i%) = PoRs("StageZ")
iData%(i%) = PoRs("PosOrder")   ' row numbers (may not be consecutive)
PoRs.MoveNext
Loop

PoRs.Close
PoDb.Close

Screen.MousePointer = vbDefault
Exit Sub

' Errors
PositionGetSampleDataOnlyError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "PositionGetSampleDataOnly"
ierror = True
Exit Sub

PositionGetSampleDataOnlyNoPositions:
Screen.MousePointer = vbDefault
msg$ = "No position data for position sample " & Str$(position.samplenumber%) & " in " & PositionDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "PositionGetSampleDataOnly"
ierror = True
Exit Sub

End Sub

