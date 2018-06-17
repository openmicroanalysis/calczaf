Attribute VB_Name = "CodeXray5"
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

Sub XrayGetDatabase()
' Loads the xray database for the user to browse

ierror = False
On Error GoTo XrayGetDatabaseError

Dim klm As Single, keV As Single
Dim xstart As Single, xstop As Single
Dim method As Integer

' Load defaults
If DefaultXrayStart! = 0# Then DefaultXrayStart! = 9.5933
If DefaultXrayStop! = 0# Then DefaultXrayStop! = 10.1867
klm! = DefaultMinimumKLMDisplay!
keV! = DefaultKiloVolts!
xstart! = DefaultXrayStart!
xstop! = DefaultXrayStop!
method% = DefaultAbsorptionEdgeDisplay%

' Load form
Call XrayLoad(Int(2), method%, klm!, keV!, xstart!, xstop!)
If ierror Then Exit Sub

FormXRAY.Show vbModeless
If ierror Then Exit Sub

Exit Sub

' Errors
XrayGetDatabaseError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayGetDatabase"
ierror = True
Exit Sub

End Sub

Sub XrayLoad(mode As Integer, method As Integer, klm As Single, keV As Single, xstart As Single, xstop As Single)
' Calls XrayLoadDatabase to load the ListXray
' mode = 1 enable "Graph Selected" button
' mode = 2 disable "Graph Selected" button
' method = 0 load just x-ray lines
' method = 1 load x-ray lines and absorption edges

ierror = False
On Error GoTo XrayLoadError

Dim i As Integer

' Load xrays from XRAY.MDB database
Call XrayLoadDatabase(method%, klm!, keV!, xstart!, xstop!)
If ierror Then Exit Sub

FormXRAY.TextMinimumKLM.Text = Str$(DefaultMinimumKLMDisplay!)
FormXRAY.TextStart.Text = Str$(xstart!)
FormXRAY.TextStop.Text = Str$(xstop!)
FormXRAY.TextKev.Text = Str$(keV!)

' Load absorption edges checkbox
If DefaultAbsorptionEdgeDisplay% = 1 Then
FormXRAY.CheckAbsorptionEdges.value = vbChecked
Else
FormXRAY.CheckAbsorptionEdges.value = vbUnchecked
End If

' Set "Graph Selected" button ebale
If mode% = 1 Then
FormXRAY.CommandGraphSelected.Enabled = True
Else
FormXRAY.CommandGraphSelected.Enabled = False
End If

' Load combo list for xray element
FormXRAY.ComboElement.Clear
FormXRAY.ComboElm.Clear
For i% = 1 To MAXELM%
FormXRAY.ComboElement.AddItem Symup$(i%)
FormXRAY.ComboElm.AddItem Symup$(i%)
Next i%

FormXRAY.ComboXry.Clear
For i% = 1 To MAXRAY% - 1
FormXRAY.ComboXry.AddItem Xraylo$(i%)
Next i%

FormXRAY.ComboOrder.Clear
For i% = 1 To MAXKLMORDER%
FormXRAY.ComboOrder.AddItem Format$(i%)
Next i%

FormXRAY.ComboMaximumOrder.Clear
For i% = 1 To MAXKLMORDER%
FormXRAY.ComboMaximumOrder.AddItem RomanNum$(i%)
Next i%
FormXRAY.ComboMaximumOrder.ListIndex = DefaultMaximumOrder% - 1

' DoEvents in case Form is just updated
DoEvents
Exit Sub

' Errors
XrayLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayLoad"
ierror = True
Exit Sub

End Sub

Sub XrayLoadDatabase(method As Integer, klm As Single, keV As Single, xstart As Single, xstop As Single)
' Load the list box based on specified kilovolts and range
' method = 0 load just x-ray lines
' method = 1 load x-ray lines and absorption edges
' klm = minimum x-ray intensity (normalized to 100 or 150)
' kev = maximum x-ray energy
' xstart = angstrom start
' xstop = angstrom end

ierror = False
On Error GoTo XrayLoadDatabaseError

Dim i As Integer
Dim tempmin As Single, tempmax As Single
Dim factor As Integer
Dim maxrec As Long
Dim tmsg As String

Dim SQLQ As String
Dim PrDb As Database
Dim PrDs As Recordset

' Sort xstart and xstop
If xstart! < xstop! Then
tempmin! = xstart!
tempmax! = xstop!
Else
tempmin! = xstop!
tempmax! = xstart!
End If

' Open xray database (exclusive and read only)
Screen.MousePointer = vbHourglass
Set PrDb = OpenDatabase(XrayDataFile$, XrayDatabaseNonExclusiveAccess%, dbReadOnly)

' All element x-ray lines in range (limit number of lines loaded)
maxrec& = MAXLISTBOXSIZE% + 1
factor% = 1
Do While maxrec& > MAXLISTBOXSIZE%

' Load just x-ray lines
If GraphWavescanType < 3 Then
If method% = 0 Then
SQLQ$ = "SELECT Xray.* FROM Xray WHERE XrayIntensity >= " & Str$(klm! * factor%) & " AND XrayEnergy <= " & Str$(keV!) & " AND XrayLambda > " & Str$(tempmin!) & " AND XrayLambda < " & Str$(tempmax!) & " AND XrayOrder <= " & Str$(DefaultMaximumOrder%)

' Load x-ray lines and absorption edges
Else
SQLQ$ = "SELECT Xray.* FROM Xray WHERE (XrayIntensity >= " & Str$(klm! * factor%) & " AND XrayEnergy <= " & Str$(keV!) & " AND XrayLambda > " & Str$(tempmin!) & " AND XrayLambda < " & Str$(tempmax!) & " AND XrayOrder <= " & Str$(DefaultMaximumOrder%) & ")"
SQLQ$ = SQLQ$ & " OR (XrayAbsEdge = 'ABS' AND XrayEnergy <= " & Str$(keV!) & " AND XrayLambda > " & Str$(tempmin!) & " AND XrayLambda < " & Str$(tempmax!) & ")"
End If

' Convert back to keV
Else
tempmin! = ANGKEV! / xstart!       ' convert to max keV
tempmax! = ANGKEV! / xstop!       ' convert to min keV
If method% = 0 Then
SQLQ$ = "SELECT Xray.* FROM Xray WHERE XrayIntensity >= " & Str$(klm! * factor%) & " AND XrayEnergy <= " & Str$(keV!) & " AND XrayEnergy > " & Str$(tempmin!) & " AND XrayEnergy < " & Str$(tempmax!) & " AND XrayOrder <= " & Str$(DefaultMaximumOrder%)

' Load x-ray lines and absorption edges
Else
SQLQ$ = "SELECT Xray.* FROM Xray WHERE (XrayIntensity >= " & Str$(klm! * factor%) & " AND XrayEnergy <= " & Str$(keV!) & " AND XrayEnergy > " & Str$(tempmin!) & " AND XrayEnergy < " & Str$(tempmax!) & " AND XrayOrder <= " & Str$(DefaultMaximumOrder%) & ")"
SQLQ$ = SQLQ$ & " OR (XrayAbsEdge = 'ABS' AND XrayEnergy <= " & Str$(keV!) & " AND XrayEnergy > " & Str$(tempmin!) & " AND XrayEnergy < " & Str$(tempmax!) & ")"
End If
End If

' Add skip for forbidden elements
If NumberofForbiddenElements% > 0 Then
For i% = 1 To NumberofForbiddenElements%
SQLQ$ = SQLQ$ & " AND XraySymbol <> '" & Trim$(Symup$(ForbiddenElements%(i%))) & "'"
Next i%
End If

' Query database
Set PrDs = PrDb.OpenRecordset(SQLQ$, dbOpenSnapshot)
If Not PrDs.EOF Then PrDs.MoveLast
maxrec& = PrDs.RecordCount
factor% = factor% * 2
Loop
If Not PrDs.BOF Then PrDs.MoveFirst

' Load list boxes
FormXRAY.ListXray.Clear
Do Until PrDs.EOF
msg$ = Space$(MAXKLMORDERCHAR%)     ' make place holder for order = 1 lines
tmsg$ = String$(MAXKLMORDERCHAR%, 64)   ' use "@" format specifier for higher orders
If PrDs("XrayOrder") > 1 Then msg$ = Format$(RomanNum$(PrDs("XrayOrder")), tmsg$)
msg$ = PrDs("XraySymbol") & " " & PrDs("XrayLine") & " " & msg$ & vbTab & MiscAutoFormat$(PrDs("XrayLambda")) & " " & MiscAutoFormat$(PrDs("XrayEnergy")) & " " & MiscAutoFormat$(PrDs("XrayIntensity"))
msg$ = msg$ & " " & Format$(PrDs("XrayAbsEdge"), "@@@") & vbTab & PrDs("XrayReference")

FormXRAY.ListXray.AddItem msg$
FormXRAY.ListXray.ItemData(FormXRAY.ListXray.NewIndex) = PrDs("XrayAtomNum")
PrDs.MoveNext
Loop

If DebugMode And VerboseMode Then
Call IOWriteLog("XrayLoadDatabase: " & Str$(FormXRAY.ListXray.ListCount) & " xray lines loaded...")
End If
PrDs.Close
PrDb.Close

Screen.MousePointer = vbDefault
Exit Sub

' Errors
XrayLoadDatabaseError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "XrayLoadDatabase"
ierror = True
Exit Sub

End Sub

Sub XrayHighlightElement()
' Highlight element in FormXRAY list

ierror = False
On Error GoTo XrayHighlightElementError

Dim i As Integer

' Unselect and select selected element
For i% = 0 To FormXRAY.ListXray.ListCount - 1
FormXRAY.ListXray.Selected(i%) = False
If MiscStringsAreSame(Left$(FormXRAY.ListXray.List(i%), 2), FormXRAY.ComboElement.Text) Then
FormXRAY.ListXray.Selected(i%) = True
End If
Next i%

Exit Sub

' Errors
XrayHighlightElementError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayHighlightElement"
ierror = True
Exit Sub

End Sub

Sub XrayLoadNewRange()
' Re-load xray list box based on new values

ierror = False
On Error GoTo XrayLoadNewRangeError

Call XraySave
If ierror Then Exit Sub

' Load list box
Call XrayLoadDatabase(DefaultAbsorptionEdgeDisplay%, DefaultMinimumKLMDisplay!, DefaultKiloVolts!, DefaultXrayStart!, DefaultXrayStop!)
If ierror Then Exit Sub

Exit Sub

' Errors
XrayLoadNewRangeError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayLoadNewRange"
ierror = True
Exit Sub

End Sub

Sub XraySave()
' Save form parameters

ierror = False
On Error GoTo XraySaveError

Dim method As Integer
Dim klm As Single, keV As Single
Dim xstart As Single, xstop As Single

' Load user new range and check for valid values
If Val(FormXRAY.TextMinimumKLM.Text) < 0.005 Or Val(FormXRAY.TextMinimumKLM.Text) > 150# Then
msg$ = "Minimum KLM intensity out of range (must be between 0.005 and 150)"
MsgBox msg$, vbOKOnly + vbExclamation, "XraySave"
ierror = True
Exit Sub
Else
klm! = Val(FormXRAY.TextMinimumKLM.Text)
End If

If Val(FormXRAY.TextKev.Text) < MINKILOVOLTS! Or Val(FormXRAY.TextKev.Text) > MAXKILOVOLTS! Then
msg$ = "Kilovolts value is out of range (must be between " & Format$(MINKILOVOLTS!) & " and " & Format$(MAXKILOVOLTS!) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "XraySave"
ierror = True
Exit Sub
Else
keV! = Val(FormXRAY.TextKev.Text)
End If

If Val(FormXRAY.TextStart.Text) < 0.5 Or Val(FormXRAY.TextStart.Text) > 240# Then
msg$ = "Minimum angstroms value is out of range (must be between 0.5 and 240)"
MsgBox msg$, vbOKOnly + vbExclamation, "XraySave"
ierror = True
Exit Sub
Else
xstart! = Val(FormXRAY.TextStart.Text)
End If

If Val(FormXRAY.TextStop.Text) < 0.5 Or Val(FormXRAY.TextStop.Text) > 240# Then
msg$ = "Maximum angstroms value is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "XraySave"
ierror = True
Exit Sub
Else
xstop! = Val(FormXRAY.TextStop.Text)
End If

' Load x-ray absorption edges flag
If FormXRAY.CheckAbsorptionEdges.value = vbChecked Then
method% = 1
Else
method% = 0
End If

DefaultMaximumOrder% = FormXRAY.ComboMaximumOrder.ListIndex + 1

' Save defaults
DefaultMinimumKLMDisplay! = klm!
If Not RealTimeMode Then DefaultKiloVolts! = keV!
DefaultXrayStart! = xstart!
DefaultXrayStop! = xstop!
DefaultAbsorptionEdgeDisplay = method%

Exit Sub

' Errors
XraySaveError:
MsgBox Error$, vbOKOnly + vbCritical, "XraySave"
ierror = True
Exit Sub

End Sub

Sub XrayExtractListString(tmsg As String, esym As String, xsym As String, order As Integer, lambda As Single, intens As Single, xlabel As String, xedge As String)
' Routine to extract information from the FormXRAY list string passed

ierror = False
On Error GoTo XrayExtractListStringError

esym$ = Left$(tmsg$, 2)            ' element symbol
xsym$ = Mid$(tmsg$, 4, 8)          ' xray symbol

If Trim$(Mid$(tmsg$, 13, 5)) = vbNullString Then
order% = 1  ' assume blank is Bragg first order line
Else
order% = XrayRomanTo%(Mid$(tmsg$, 13, 5))   ' actual Bragg order line (change to 5 character Roman numneral order)
End If

lambda! = Val(Mid$(tmsg$, 18, 8))  ' lambda
intens! = Val(Mid$(tmsg$, 37, 8))  ' intensity
xedge$ = Mid$(tmsg$, 46, 3)        ' absorption edge flag
xlabel$ = " " & Left$(tmsg$, 17)   ' label

Exit Sub

' Errors
XrayExtractListStringError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayExtractListString"
ierror = True
Exit Sub

End Sub

Sub XrayOpenNewMDB()
' This routine reads file XRAY.ALL and converts it to XRAY.MDB file (use MakXray.exe to convert XrayData5_8_99.txt to xray.dat, then eventually to xray.all)

ierror = False
On Error GoTo XrayOpenNewMDBError

Dim nrec As Long, lastrec As Long
Dim response As Integer
Dim xrayfile As String

Dim StDb As Database
Dim StDt As Recordset
Dim XrayRow As TypeXray

' Specify the xray database variables
Dim Xray As TableDef

' If file already exists, warn user
If Dir$(XrayDataFile$) <> vbNullString Then
msg$ = "Xray Database: " & vbCrLf
msg$ = msg$ & XrayDataFile$ & vbCrLf
msg$ = msg$ & " already exists, are you sure you want to overwrite it?"
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton2, "XrayOpenNewMDB")
If response% = vbNo Then
ierror = True
Exit Sub
End If
End If

' Check for XRAY.ALL (See MakeXray.Exe) (XRAY.ALL is not distributed, so do not use ProgramPath$ or ApplicationCommonAppData$)
xrayfile$ = app.Path & "\XRAY.ALL"
If Dir$(xrayfile$) = vbNullString Then GoTo XrayOpenNewMDBNoXrayFile

' If xray database exists, delete it
If Dir$(XrayDataFile$) <> vbNullString Then
Kill XrayDataFile$

' Else inform user
Else
msg$ = "Creating a new xray database: " & XrayDataFile$
MsgBox msg$, vbOKOnly + vbInformation, "XrayOpenNewMDB"
End If

' Open a new database by copying from existing MDB template
Call FileInfoCreateDatabase(XrayDataFile$)
If ierror Then Exit Sub

' Open as existing database
Set StDb = OpenDatabase(XrayDataFile$, DatabaseExclusiveAccess%, False)

' Specify the xray database "Xray" table
Set Xray = StDb.CreateTableDef("NewTableDef")
Xray.Name = "Xray"

With Xray
.Fields.Append .CreateField("XrayAtomNum", dbInteger)
.Fields.Append .CreateField("XrayOrder", dbInteger)
.Fields.Append .CreateField("XraySymbol", dbText, DbTextXrayStringLength%)
.Fields("XraySymbol").AllowZeroLength = True
.Fields.Append .CreateField("XrayLine", dbText, 8)
.Fields("XrayLine").AllowZeroLength = True
.Fields.Append .CreateField("XrayAbsEdge", dbText, 3)
.Fields("XrayAbsEdge").AllowZeroLength = True
.Fields.Append .CreateField("XrayReference", dbText, 5)
.Fields("XrayReference").AllowZeroLength = True
.Fields.Append .CreateField("XrayLambda", dbSingle)
.Fields.Append .CreateField("XrayEnergy", dbSingle)
.Fields.Append .CreateField("XrayIntensity", dbSingle)
End With

StDb.TableDefs.Append Xray

' Add new records to "Xray" table
Set StDt = StDb.OpenRecordset("Xray", dbOpenTable)

' Open old binary xray file and get number of records
Open xrayfile$ For Random Access Read As #Temp1FileNumber% Len = 38
Get #Temp1FileNumber%, 1, lastrec&

' Load all records from old file
For nrec& = 2 To lastrec&
Get #Temp1FileNumber%, nrec&, XrayRow

StDt.AddNew
StDt("XrayAtomNum") = XrayRow.atnum%
StDt("XraySymbol") = Left$(XrayRow.syme$, DbTextElementStringLength%)
StDt("XrayOrder") = XrayRow.n%
StDt("XrayLine") = Left$(XrayRow.xline$, 8)
StDt("XrayAbsEdge") = Left$(XrayRow.abedg$, 3)
StDt("XrayLambda") = XrayRow.xwave#
StDt("XrayEnergy") = ANGKEV! / XrayRow.xwave# * XrayRow.n%
StDt("XrayIntensity") = XrayRow.xints#
StDt("XrayReference") = Left$(XrayRow.refer$, 5)
StDt.Update

Call IOStatusAuto("Converting xray record " & Format$(nrec&) & "...")
Next nrec&
StDt.Close
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)

' Close the setup database
StDb.Close
Screen.MousePointer = vbDefault

' Create new File table for xray database
Call FileInfoMakeNewTable(Int(6), vbNullString)
If ierror Then Exit Sub

msg$ = "Output completed to XRAY.MDB"
MsgBox msg$, vbOKOnly + vbInformation, "XrayOpenNewMDB"

Exit Sub

' Errors
XrayOpenNewMDBError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "XrayOpenNewMDB"
Call IOStatusAuto(vbNullString)
Close #Temp1FileNumber%
ierror = True
Exit Sub

XrayOpenNewMDBNoXrayFile:
msg$ = "File " & xrayfile$ & " was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "XrayOpenNewMDB"
ierror = True
Exit Sub

End Sub

Function XrayRomanTo(astring As String) As Integer
' Convert Roman numeral to number (integer)

ierror = False
On Error GoTo XrayRomanToError

Dim n As Integer

n% = 0
If Trim$(astring$) = "I" Then n% = 1
If Trim$(astring$) = "II" Then n% = 2
If Trim$(astring$) = "III" Then n% = 3
If Trim$(astring$) = "IV" Then n% = 4
If Trim$(astring$) = "V" Then n% = 5
If Trim$(astring$) = "VI" Then n% = 6
If Trim$(astring$) = "VII" Then n% = 7
If Trim$(astring$) = "VIII" Then n% = 8
If Trim$(astring$) = "IX" Then n% = 9
If Trim$(astring$) = "X" Then n% = 10

If Trim$(astring$) = "XI" Then n% = 11
If Trim$(astring$) = "XII" Then n% = 12
If Trim$(astring$) = "XIII" Then n% = 13
If Trim$(astring$) = "XIV" Then n% = 14
If Trim$(astring$) = "XV" Then n% = 15
If Trim$(astring$) = "XVI" Then n% = 16
If Trim$(astring$) = "XVII" Then n% = 17
If Trim$(astring$) = "XVIII" Then n% = 18
If Trim$(astring$) = "XIX" Then n% = 19
If Trim$(astring$) = "XX" Then n% = 20
If n% = 0 Then GoTo XrayRomanToNotFound

XrayRomanTo% = n%
Exit Function

' Errors
XrayRomanToError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "XrayRomanTo"
ierror = True
Exit Function

XrayRomanToNotFound:
Screen.MousePointer = vbDefault
msg$ = "Unable to convert Roman number " & astring$ & " to a integer numeric value (should not occur)"
MsgBox msg$, vbOKOnly + vbExclamation, "XrayRomanTo"
ierror = True
Exit Function

End Function

Sub XraySpecifyRange(mode As Integer)
' Specify spectrum range in FormXRAY list
' mode = 1 element change
' mode = 2 xray change
' mode = 3 order change

ierror = False
On Error GoTo XraySpecifyRangeError

Dim order As Integer
Dim sym As String, ray As String
Dim keV As Single, lam As Single
Dim xstart As Single, xstop As Single
Dim ip As Integer, ipp As Integer

sym$ = FormXRAY.ComboElm.Text
ip% = IPOS1(MAXELM%, sym$, Symlo$())
If ip% = 0 Then Exit Sub

' Load default x-ray if element changed
If mode% = 1 Then
FormXRAY.ComboXry.Text = Deflin$(ip%)
FormXRAY.ComboOrder.Text = Format$(1)   ' assume first order
End If

ray$ = FormXRAY.ComboXry.Text
ipp% = IPOS1(MAXRAY% - 1, ray$, Xraylo$())
If ipp% = 0 Then Exit Sub

order% = Val(FormXRAY.ComboOrder.Text)
If order% = 0 Then Exit Sub

' Get angstroms
Call XrayGetKevLambda(sym$, ray$, keV!, lam!)
If ierror Then Exit Sub

' Convert according to order
lam! = lam! * order%

' Calculate start and stop
xstart! = lam! - lam! * 0.03
xstop! = lam! + lam! * 0.03

FormXRAY.TextStart.Text = Str$(xstart!)
FormXRAY.TextStop.Text = Str$(xstop!)

' Load new range
Call XrayLoadNewRange
If ierror Then Exit Sub

Exit Sub

' Errors
XraySpecifyRangeError:
MsgBox Error$, vbOKOnly + vbCritical, "XraySpecifyRange"
ierror = True
Exit Sub

End Sub

Sub XrayLoadDatabase2(minintensity As Single, atnum As Integer, xstart As Single, xstop As Single, num As Long, sarray() As String, xarray() As String, iarray() As Single, earray() As Single)
' Load array of lines (1st order only) and intensities for the elements specified for the passed range (in keV)
' minintensity = minimum intensity (normalized to 100)
' atnum = atomic number of element
' xstart = kev start
' xstop = kev end
' sarray() = element symbol string
' xarray() = xray symbol string
' iarray() = intensity
' earray() = energy in kev

ierror = False
On Error GoTo XrayLoadDatabase2Error

Dim n As Long
Dim tempmin As Single, tempmax As Single

Dim SQLQ As String
Dim PrDb As Database
Dim PrDs As Recordset

' Sort xstart and xstop
If xstart! < xstop! Then
tempmin! = xstart!
tempmax! = xstop!
Else
tempmin! = xstop!
tempmax! = xstart!
End If

' Open xray database (exclusive and read only)
Screen.MousePointer = vbHourglass
Set PrDb = OpenDatabase(XrayDataFile$, XrayDatabaseNonExclusiveAccess%, dbReadOnly)

' All element x-ray lines in range
SQLQ$ = "SELECT Xray.* FROM Xray WHERE XrayAtomNum = " & Str$(atnum%) & " AND XrayIntensity >= " & Str$(minintensity!) & " AND XrayOrder = 1" & " AND XrayEnergy > " & Str$(tempmin!) & " AND XrayEnergy < " & Str$(tempmax!) '& " AND XrayEdge <> ABS"

num& = 0
Set PrDs = PrDb.OpenRecordset(SQLQ$, dbOpenSnapshot)
If PrDs.EOF Then
Screen.MousePointer = vbDefault
Exit Sub
End If

PrDs.MoveLast
num& = PrDs.RecordCount
PrDs.MoveFirst
ReDim sarray(1 To num&) As String
ReDim xarray(1 To num&) As String
ReDim iarray(1 To num&) As Single
ReDim earray(1 To num&) As Single

' Load list array
n& = 0
Do Until PrDs.EOF
n& = n& + 1
sarray$(n&) = PrDs("XraySymbol")
xarray$(n&) = PrDs("XrayLine")
iarray!(n&) = PrDs("XrayIntensity")
earray!(n&) = PrDs("XrayEnergy")
PrDs.MoveNext
Loop

PrDs.Close
PrDb.Close

Screen.MousePointer = vbDefault
Exit Sub

' Errors
XrayLoadDatabase2Error:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "XrayLoadDatabase2"
ierror = True
Exit Sub

End Sub

Sub XrayGetKLMElements2(elmarray() As Boolean)
' Loads the periodic table for the user to browse for KLM elements

ierror = False
On Error GoTo XrayGetKLMElements2Error

Dim i As Integer, j As Integer

' Pass user selected elements
Call Periodic2To(elmarray())
If ierror Then Exit Sub

' Load form
icancelload = False
Call Periodic2Load
If icancelload = True Then Exit Sub

' Get selected elements
Call Periodic2Return(elmarray())
If ierror Then Exit Sub

' Unselect and select selected elements
For i% = 0 To FormXRAY.ListXray.ListCount - 1
FormXRAY.ListXray.Selected(i%) = False
For j% = 1 To MAXELM%
If elmarray(j%) Then
If MiscStringsAreSame(Left$(FormXRAY.ListXray.List(i%), 2), Symlo$(j%)) Then
FormXRAY.ListXray.Selected(i%) = True
End If
End If
Next j%
Next i%

Exit Sub

' Errors
XrayGetKLMElements2Error:
MsgBox Error$, vbOKOnly + vbCritical, "XrayGetKLMElements2"
ierror = True
Exit Sub

End Sub

Sub XrayGetKLMElements0()
' Loads the periodic table for the user to browse for KLM elements (only used by CalcZAF and Standard)

ierror = False
On Error GoTo XrayGetKLMElements0Error

Dim i As Integer, j As Integer
Dim elmarray(1 To MAXELM%) As Boolean

' Load form
icancelload = False
Call Periodic2Load
If icancelload = True Then Exit Sub

' Get selected elements
Call Periodic2Return(elmarray())
If ierror Then Exit Sub

' Unselect and select selected element
For i% = 0 To FormXRAY.ListXray.ListCount - 1
FormXRAY.ListXray.Selected(i%) = False
For j% = 1 To MAXELM%
If elmarray(j%) Then
If MiscStringsAreSame(Left$(FormXRAY.ListXray.List(i%), 2), Symlo$(j%)) Then
FormXRAY.ListXray.Selected(i%) = True
End If
End If
Next j%
Next i%

Exit Sub

' Errors
XrayGetKLMElements0Error:
MsgBox Error$, vbOKOnly + vbCritical, "XrayGetKLMElements0"
ierror = True
Exit Sub

End Sub
