Attribute VB_Name = "CodeSTANDARD4"
' (c) Copyright 1995-2023 by John J. Donovan
Option Explicit

Dim StandardTmpSample(1 To 1) As TypeSample

Sub StandardOpenDATFile(mode As Integer, tfilename As String, tForm As Form)
' Open a ASCII DAT file for standard composition import
' mode = 1 export
' mode = 2 import
' mode = 3 export (single row format)
' mode = 4 import (single row format)

ierror = False
On Error GoTo StandardOpenDATFileError

' Make sure that a .MDB standard database is already open to import into
If StandardDataFile$ = vbNullString Then
Call StandardOpenMDBFile(tfilename$, tForm)
If ierror Then Exit Sub
End If

' Open the DAT file
If mode% = 1 Or mode% = 2 Then
tfilename$ = "standard.dat"
Else
tfilename$ = "standard2.dat"
End If

' Get file name from user
If mode% = 2 Or mode% = 4 Then
Call IOGetFileName(Int(2), "DAT", tfilename$, tForm)
If ierror Then Exit Sub

Else
Call IOGetFileName(Int(1), "DAT", tfilename$, tForm)
If ierror Then Exit Sub
End If

' No error, load file name
ImportDataFile$ = tfilename$

' Open the ASCII file and import the data to the standard database
If mode% = 2 Or mode% = 4 Then
Open ImportDataFile$ For Input As #ImportDataFileNumber%
Else
Open ImportDataFile$ For Output As #ImportDataFileNumber%
End If

Exit Sub

' Errors
StandardOpenDATFileError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardOpenDATFile"
ierror = True
Exit Sub

End Sub

Sub StandardReadDATFile(mode As Integer)
' The routine imports standard composition data from an ASCII file to the standard database
' mode = 1 read normal ASCII
' mode = 2 read single row format
' mode = 3 read all_weights-modified.txt

ierror = False
On Error GoTo StandardReadDATFileError

Dim linecount As Long, standardcount As Long
Dim ip As Integer

icancelauto = False

' If single row format skip first two lines
If mode% = 2 Then
Line Input #ImportDataFileNumber%, msg$
If Mid$(msg$, 2, 6) <> "Number" Then Line Input #ImportDataFileNumber%, msg$  ' for Excel files with filename as first line
End If

' Loop on standard import file
linecount& = 0
standardcount& = 0
Do While Not EOF(ImportDataFileNumber%)

' Read data from ASCII standard file
If mode% = 1 Then
Call StandardReadDATSample(linecount&, StandardTmpSample())
If ierror Then Exit Sub

ElseIf mode% = 2 Then
Call StandardReadDATSample2(linecount&, StandardTmpSample())
If ierror Then Exit Sub

Else
Call StandardReadDATSample3(standardcount&, linecount&, StandardTmpSample())
If ierror Then Exit Sub
If standardcount& = 0 Then Exit Sub     ' end of file reached
End If

' Update status form
Call IOStatusAuto("Importing Standard " & Str$(StandardTmpSample(1).number%) & " " & StandardTmpSample(1).Name$)
DoEvents
If icancelauto Then
ierror = True
Exit Sub
End If

' Check if standard already exists, ask whether to replace
ip% = StandardGetRow%(StandardTmpSample(1).number%)
If ip% > 0 Then
Call StandardReplaceRecord(StandardTmpSample())
If ierror Then Exit Sub

' Else add new standard composition in MDB database
Else
Call StandardAddRecord(StandardTmpSample())
If ierror Then Exit Sub

' Update available standard list
If NumberOfAvailableStandards% + 1 > MAXINDEX% Then GoTo StandardReadDATFileTooMany
NumberOfAvailableStandards% = NumberOfAvailableStandards% + 1
StandardIndexNumbers%(NumberOfAvailableStandards%) = StandardTmpSample(1).number%
StandardIndexNames$(NumberOfAvailableStandards%) = StandardTmpSample(1).Name$
StandardIndexDescriptions$(NumberOfAvailableStandards%) = StandardTmpSample(1).Description$
StandardIndexDensities!(NumberOfAvailableStandards%) = StandardTmpSample(1).SampleDensity!
StandardIndexMaterialTypes$(NumberOfAvailableStandards%) = StandardTmpSample(1).MaterialType$
StandardIndexMountNames$(NumberOfAvailableStandards%) = StandardTmpSample(1).MountNames$
End If

Loop

Exit Sub

' Errors
StandardReadDATFileError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardReadDATFile"
ierror = True
Exit Sub

StandardReadDATFileTooMany:
msg$ = "Too many standards were found in " & ImportDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardReadDATFile"
ierror = True
Exit Sub

End Sub

Sub StandardReadDATSample(linecount As Long, sample() As TypeSample)
' Read standard compositions from STANDARD.DAT file. Called by StandardReadDATFile.

ierror = False
On Error GoTo StandardReadDATSampleError

Dim ip As Integer, i As Integer

' Initialize
Call InitSample(sample())
If ierror Then Exit Sub

' Read standard number and name
Input #ImportDataFileNumber%, sample(1).number%, sample(1).Name$
linecount& = linecount& + 1

' Read standard description (new in v. 2.48)
Input #ImportDataFileNumber%, sample(1).Description$
linecount& = linecount& + 1

' Read DisplayAsOxideFlag flag and number of elements in standard
Input #ImportDataFileNumber%, sample(1).DisplayAsOxideFlag%, sample(1).LastChan%
linecount& = linecount& + 1

' Read standard density
Input #ImportDataFileNumber%, sample(1).SampleDensity!
linecount& = linecount& + 1

' Load defaults
sample(1).LastElm% = sample(1).LastChan%
sample(1).takeoff! = DefaultTakeOff!
sample(1).kilovolts! = DefaultKiloVolts!
sample(1).beamcurrent! = DefaultBeamCurrent!
sample(1).beamsize! = DefaultBeamSize!

' Load kilovolts array
For i% = 1 To sample(1).LastChan%
sample(1).TakeoffArray!(i%) = sample(1).takeoff!
sample(1).KilovoltsArray!(i%) = sample(1).kilovolts!
sample(1).BeamCurrentArray(i%) = sample(1).beamcurrent!
sample(1).BeamSizeArray(i%) = sample(1).beamsize!
Next i%

' Confirm on screen if debug
If DebugMode Then
msg$ = Str$(sample(1).number%) & " " & sample(1).Name$ & " " & Str$(sample(1).LastChan%)
Call IOWriteLog(msg$)
End If

' Check for invalid TakeOff, KiloVolts and DisplayAsOxideFlag flag
If sample(1).takeoff! < 15# Or sample(1).takeoff! > 75# Then GoTo StandardReadDATSampleInvalidTakeOff
If sample(1).kilovolts! < 1# Or sample(1).kilovolts! > 100# Then GoTo StandardReadDATSampleInvalidKiloVolts
If sample(1).DisplayAsOxideFlag% <> 0 And sample(1).DisplayAsOxideFlag% <> -1 Then GoTo StandardReadDATSampleBadDisplayAsOxide

' Check for invalid standard number or number of elements in standard
If sample(1).number% <= 0 Then GoTo StandardReadDATSampleBadStandardNumber
If sample(1).LastChan% > MAXCHAN% Then GoTo StandardReadDATSampleTooManyElements

' Read standard elements symbols
For i% = 1 To sample(1).LastChan%
Input #ImportDataFileNumber%, sample(1).Elsyms$(i%)
Next i%
linecount& = linecount& + 1

' Check for valid element symbols and load default x-ray lines
msg$ = vbNullString
For i% = 1 To sample(1).LastChan%
ip% = IPOS1(MAXELM%, sample(1).Elsyms$(i%), Symlo$())
If ip% = 0 Then GoTo StandardReadDATSampleInvalidElement
sample(1).Xrsyms$(i%) = Deflin$(ip%)
msg$ = msg$ & Format$(sample(1).Elsyms$(i%), a80$)
Next i%
If DebugMode Then Call IOWriteLog(msg$)

' Read standard cations and oxygens
msg$ = vbNullString
For i% = 1 To sample(1).LastChan%
Input #ImportDataFileNumber%, sample(1).numcat%(i%)
msg$ = msg$ & Format$(sample(1).numcat%(i%), a80$)
Next i%
linecount& = linecount& + 1
If DebugMode Then Call IOWriteLog(msg$)

msg$ = vbNullString
For i% = 1 To sample(1).LastChan%
Input #ImportDataFileNumber%, sample(1).numoxd%(i%)
msg$ = msg$ & Format$(sample(1).numoxd%(i%), a80$)
Next i%
linecount& = linecount& + 1
If DebugMode Then Call IOWriteLog(msg$)

' Read standard composition in elemental weight percents
msg$ = vbNullString
For i% = 1 To sample(1).LastChan%
Input #ImportDataFileNumber%, sample(1).ElmPercents!(i%)
msg$ = msg$ + Format$(Format$(sample(1).ElmPercents!(i%), f83$), a80$)
Next i%
linecount& = linecount& + 1
If DebugMode Then Call IOWriteLog(msg$)

' Read material types, formula flags, etc.
Input #ImportDataFileNumber%, sample(1).MaterialType$, sample(1).FormulaElementFlag, sample(1).FormulaRatio!, sample(1).FormulaElement$, sample(1).MountNames$
msg$ = VbDquote$ & sample(1).MaterialType$ & VbDquote$ & ", " & sample(1).FormulaElementFlag & ", " & Str$(sample(1).FormulaRatio!) & ", " & VbDquote$ & sample(1).FormulaElement$ & VbDquote$ & ", " & VbDquote$ & sample(1).MountNames$ & VbDquote$
linecount& = linecount& + 1
If DebugMode Then Call IOWriteLog(msg$)

Exit Sub

' Errors
StandardReadDATSampleError:
msg$ = Error$ & " in file " & ImportDataFile$ & " on line " & Str$(linecount&)
MsgBox msg$, vbOKOnly + vbCritical, "StandardReadDATSample"
ierror = True
Exit Sub

StandardReadDATSampleInvalidTakeOff:
msg$ = "TakeOff is invalid in " & ImportDataFile$ & " on line " & Str$(linecount&)
MsgBox msg$, vbOKOnly + vbExclamation, "StandardReadDATSample"
ierror = True
Exit Sub

StandardReadDATSampleInvalidKiloVolts:
msg$ = "KiloVolts is invalid in " & ImportDataFile$ & " on line " & Str$(linecount&)
MsgBox msg$, vbOKOnly + vbExclamation, "StandardReadDATSample"
ierror = True
Exit Sub

StandardReadDATSampleBadDisplayAsOxide:
msg$ = "DisplayAsOxideFlag flag is invalid in " & ImportDataFile$ & " on line " & Str$(linecount&)
MsgBox msg$, vbOKOnly + vbExclamation, "StandardReadDATSample"
ierror = True
Exit Sub

StandardReadDATSampleBadStandardNumber:
msg$ = "Standard " & Str$(sample(1).number%) & " is invalid in " & ImportDataFile$ & " on line " & Str$(linecount&)
MsgBox msg$, vbOKOnly + vbExclamation, "StandardReadDATSample"
ierror = True
Exit Sub

StandardReadDATSampleTooManyElements:
msg$ = "Too many Elements in " & ImportDataFile$ & " on line " & Str$(linecount&)
MsgBox msg$, vbOKOnly + vbExclamation, "StandardReadDATSample"
ierror = True
Exit Sub

StandardReadDATSampleInvalidElement:
msg$ = "Invalid Element in " & ImportDataFile$ & " on line " & Str$(linecount&)
MsgBox msg$, vbOKOnly + vbExclamation, "StandardReadDATSample"
ierror = True
Exit Sub

End Sub

Sub StandardReadDATSample2(linecount As Long, sample() As TypeSample)
' Read standard compositions from STANDARD2.DAT file (single row format). Called by StandardReadDATFile.

ierror = False
On Error GoTo StandardReadDATSample2Error

Dim itemp As Integer, i As Integer, n As Integer
Dim tmsg As String, achar As String

' Initialize
Call InitSample(sample())
If ierror Then Exit Sub

' Read standard number and name
Input #ImportDataFileNumber%, sample(1).number%, sample(1).Name$

' Read standard description (new in v. 2.48)
Input #ImportDataFileNumber%, sample(1).Description$

' Read DisplayAsOxideFlag flag in standard
Input #ImportDataFileNumber%, sample(1).DisplayAsOxideFlag%

' Read standard density
Input #ImportDataFileNumber%, sample(1).SampleDensity!

' Load defaults
sample(1).LastElm% = sample(1).LastChan%
sample(1).takeoff! = DefaultTakeOff!
sample(1).kilovolts! = DefaultKiloVolts!
sample(1).beamcurrent! = DefaultBeamCurrent!
sample(1).beamsize! = DefaultBeamSize!

' Confirm on screen if debug
If DebugMode Then
msg$ = "Loading standard " & Str$(sample(1).number%) & " " & sample(1).Name$ & "..."
Call IOWriteLog(msg$)
End If

' Check for invalid TakeOff, KiloVolts and DisplayAsOxideFlag flag
If sample(1).takeoff! < 15# Or sample(1).takeoff! > 75# Then GoTo StandardReadDATSample2InvalidTakeOff
If sample(1).kilovolts! < 1# Or sample(1).kilovolts! > 100# Then GoTo StandardReadDATSample2InvalidKiloVolts
If sample(1).DisplayAsOxideFlag% <> 0 And sample(1).DisplayAsOxideFlag% <> -1 Then GoTo StandardReadDATSample2BadDisplayAsOxide

' Check for invalid standard number or number of elements in standard
If sample(1).number% <= 0 Then GoTo StandardReadDATSample2BadStandardNumber

' Read standard element wt percents
sample(1).LastChan% = 0
For n% = 1 To MAXELM%
achar$ = vbNullString
tmsg$ = vbNullString
Do Until achar$ = vbTab
achar$ = Input(1, #ImportDataFileNumber%)
If achar$ <> vbTab Then tmsg$ = tmsg$ & achar$
Loop

If tmsg$ <> "n.a." Then
sample(1).LastChan% = sample(1).LastChan% + 1

' Check for valid number of elements
If sample(1).LastChan% > MAXCHAN% Then GoTo StandardReadDATSample2TooManyElements
sample(1).Elsyms$(sample(1).LastChan%) = Symlo$(n%)
sample(1).Xrsyms$(sample(1).LastChan%) = Deflin$(n%)
sample(1).ElmPercents!(sample(1).LastChan%) = Val(tmsg$)
End If
Next n%

' Read standard cations
For n% = 1 To MAXELM%
Input #ImportDataFileNumber%, itemp%
For i% = 1 To sample(1).LastChan%
If MiscStringsAreSame(sample(1).Elsyms$(i%), Symlo$(n%)) Then
sample(1).numcat%(i%) = itemp%
End If
Next i%
Next n%

' Read standard oxygens
For n% = 1 To MAXELM%
Input #ImportDataFileNumber%, itemp%
For i% = 1 To sample(1).LastChan%
If MiscStringsAreSame(sample(1).Elsyms$(i%), Symlo$(n%)) Then
sample(1).numoxd%(i%) = itemp%
End If
Next i%
Next n%

' Read material type, formula flags, etc.
Input #ImportDataFileNumber%, sample(1).MaterialType$, sample(1).FormulaElementFlag, sample(1).FormulaRatio!, sample(1).FormulaElement$, sample(1).MountNames$

' Load kilovolts array
For i% = 1 To sample(1).LastChan%
sample(1).TakeoffArray!(i%) = sample(1).takeoff!
sample(1).KilovoltsArray!(i%) = sample(1).kilovolts!
sample(1).BeamCurrentArray(i%) = sample(1).beamcurrent!
sample(1).BeamSizeArray(i%) = sample(1).beamsize!
Next i%

linecount& = linecount& + 1

Exit Sub

' Errors
StandardReadDATSample2Error:
msg$ = Error$ & " in file " & ImportDataFile$ & " on line " & Str$(linecount&)
MsgBox msg$, vbOKOnly + vbCritical, "StandardReadDATSample2"
ierror = True
Exit Sub

StandardReadDATSample2InvalidTakeOff:
msg$ = "TakeOff is invalid in " & ImportDataFile$ & " on line " & Str$(linecount&)
MsgBox msg$, vbOKOnly + vbExclamation, "StandardReadDATSample2"
ierror = True
Exit Sub

StandardReadDATSample2InvalidKiloVolts:
msg$ = "KiloVolts is invalid in " & ImportDataFile$ & " on line " & Str$(linecount&)
MsgBox msg$, vbOKOnly + vbExclamation, "StandardReadDATSample2"
ierror = True
Exit Sub

StandardReadDATSample2BadDisplayAsOxide:
msg$ = "DisplayAsOxideFlag flag is invalid in " & ImportDataFile$ & " on line " & Str$(linecount&)
MsgBox msg$, vbOKOnly + vbExclamation, "StandardReadDATSample2"
ierror = True
Exit Sub

StandardReadDATSample2BadStandardNumber:
msg$ = "Standard " & Str$(sample(1).number%) & " is invalid in " & ImportDataFile$ & " on line " & Str$(linecount&)
MsgBox msg$, vbOKOnly + vbExclamation, "StandardReadDATSample2"
ierror = True
Exit Sub

StandardReadDATSample2TooManyElements:
msg$ = "Too many Elements in " & ImportDataFile$ & " on line " & Str$(linecount&)
MsgBox msg$, vbOKOnly + vbExclamation, "StandardReadDATSample2"
ierror = True
Exit Sub

StandardReadDATSample2InvalidElement:
msg$ = "Invalid Element in " & ImportDataFile$ & " on line " & Str$(linecount&)
MsgBox msg$, vbOKOnly + vbExclamation, "StandardReadDATSample2"
ierror = True
Exit Sub

End Sub

Sub StandardReadDATSample3(standardcount As Long, linecount As Long, sample() As TypeSample)
' Read standard compositions from all_weights-modified.txt file. Called by StandardReadDATFile.

ierror = False
On Error GoTo StandardReadDATSample3Error

Dim ip As Integer, n As Integer, i As Integer
Dim astring As String, bstring As String
Dim achar As String, sym As String

' Initialize
Call InitSample(sample())
If ierror Then Exit Sub

' Read standard number and name
astring$ = vbNullString
achar$ = vbNullString
Do Until achar$ = vbLf
achar$ = Input(1, #ImportDataFileNumber%)
If achar$ <> vbLf Then astring$ = astring$ & achar$
Loop
linecount& = linecount& + 1

' Check for end of text file
If astring$ = "END" Then
standardcount& = 0      ' indicate end of file
Exit Sub
End If

' Load next standard name
sample(1).Name$ = Trim$(Left$(astring$, InStr(astring$, VbSpace)))
sample(1).Description$ = Trim$(Mid$(astring$, InStr(astring$, VbSpace)))

' Load defaults
sample(1).DisplayAsOxideFlag% = False
sample(1).SampleDensity! = DEFAULTDENSITY!
sample(1).takeoff! = DefaultTakeOff!
sample(1).kilovolts! = DefaultKiloVolts!
sample(1).beamcurrent! = DefaultBeamCurrent!
sample(1).beamsize! = DefaultBeamSize!

' Generate standard number
standardcount& = standardcount& + 1
sample(1).number% = standardcount&

' Confirm on screen if debug
If DebugMode Then
msg$ = "Loading standard " & Str$(sample(1).number%) & " " & sample(1).Name$ & "..."
Call IOWriteLog(msg$)
End If

' Read standard element and wt percents
sample(1).LastChan% = 0
For n% = 1 To MAXELM%
astring$ = vbNullString
achar$ = vbNullString
Do Until achar$ = vbLf
achar$ = Input(1, #ImportDataFileNumber%)
If achar$ <> vbLf Then astring$ = astring$ & achar$
Loop
linecount& = linecount& + 1

' Parse into element and wt%
If Trim$(astring$) = vbNullString Then Exit For
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub

sym$ = Trim$(bstring$)
If UCase$(sym$) = UCase$("d") Then sym$ = "H"   ' deuterium should be hydrogen
ip% = IPOS1%(MAXELM%, sym$, Symlo$())
If ip% = 0 Then GoTo StandardReadDATSample3InvalidElement

' Check for valid number of elements
sample(1).LastChan% = sample(1).LastChan% + 1
If sample(1).LastChan% > MAXCHAN% Then GoTo StandardReadDATSample3TooManyElements

sample(1).Elsyms$(sample(1).LastChan%) = Trim$(sym$)
sample(1).Xrsyms$(sample(1).LastChan%) = Deflin$(ip%)
sample(1).ElmPercents!(sample(1).LastChan%) = Val(Trim$(astring$))
sample(1).numcat%(sample(1).LastChan%) = AllCat%(ip%)
sample(1).numoxd%(sample(1).LastChan%) = AllOxd%(ip%)

If ip% = 8 And sample(1).ElmPercents!(sample(1).LastChan%) > 5# Then    ' display as oxide if oxygen is present
sample(1).DisplayAsOxideFlag% = True
End If
Next n%

sample(1).LastElm% = sample(1).LastChan%

' Load kilovolts array
For i% = 1 To sample(1).LastChan%
sample(1).TakeoffArray!(i%) = sample(1).takeoff!
sample(1).KilovoltsArray!(i%) = sample(1).kilovolts!
sample(1).BeamCurrentArray(i%) = sample(1).beamcurrent!
sample(1).BeamSizeArray(i%) = sample(1).beamsize!
Next i%

Exit Sub

' Errors
StandardReadDATSample3Error:
msg$ = Error$ & " in file " & ImportDataFile$ & " on line " & Str$(linecount&) & " in standard " & sample(1).Name$
MsgBox msg$, vbOKOnly + vbCritical, "StandardReadDATSample3"
ierror = True
Exit Sub

StandardReadDATSample3TooManyElements:
msg$ = "Too many Elements in " & ImportDataFile$ & " on line " & Str$(linecount&) & " in standard " & sample(1).Name$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardReadDATSample3"
ierror = True
Exit Sub

StandardReadDATSample3InvalidElement:
msg$ = sym$ & " is an invalid element in " & ImportDataFile$ & " on line " & Str$(linecount&) & " in standard " & sample(1).Name$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardReadDATSample3"
ierror = True
Exit Sub

End Sub

Sub StandardWriteDATFile(mode As Integer)
' The routine exports standard composition data to an ASCII file from the standard database
' mode = 1 write normal ASCII
' mode = 2 write single row format

ierror = False
On Error GoTo StandardWriteDATFileError

Dim i As Integer, n As Integer
Dim bstring As String

ReDim outarray(1 To NumberOfAvailableStandards%) As Integer
ReDim arrayindex(1 To NumberOfAvailableStandards%) As Integer

icancelauto = False

' Sort by standard number
Call MiscSortIntegerArray(NumberOfAvailableStandards%, StandardIndexNumbers%(), outarray%(), arrayindex%())
If ierror Then Exit Sub

' Write column labels (if single row format)
If mode% = 2 Then
bstring$ = VbDquote & "Number" & VbDquote & vbTab & VbDquote & "Name" & VbDquote & vbTab
bstring$ = bstring$ & VbDquote & "Description" & VbDquote & vbTab
bstring$ = bstring$ & VbDquote & "DisplayAsOxideFlag" & VbDquote & vbTab

bstring$ = bstring$ & VbDquote & "Density" & VbDquote & vbTab    ' added 9-14-2012

For n% = 1 To MAXELM%
bstring$ = bstring$ & VbDquote & Symup$(n%) & " WT%" & VbDquote & vbTab
Next n%
For n% = 1 To MAXELM%
bstring$ = bstring$ & VbDquote & Symup$(n%) & " CAT" & VbDquote & vbTab
Next n%
For n% = 1 To MAXELM%
bstring$ = bstring$ & VbDquote & Symup$(n%) & " OXD" & VbDquote & vbTab
Next n%

bstring$ = bstring$ & VbDquote$ & "Material Type" & VbDquote$ & vbTab
bstring$ = bstring$ & VbDquote$ & "Formula Element Flag" & VbDquote$ & vbTab
bstring$ = bstring$ & VbDquote$ & "Formula Ratio" & VbDquote$ & vbTab
bstring$ = bstring$ & VbDquote$ & "Formula Element" & VbDquote$ & vbTab
bstring$ = bstring$ & VbDquote$ & "Mount Names" & VbDquote & vbTab

Print #ImportDataFileNumber%, bstring$
End If

' Loop on standards
For i% = 1 To NumberOfAvailableStandards%

' Get standard from database
Call StandardGetMDBStandard(StandardIndexNumbers%(arrayindex%(i%)), StandardTmpSample())
If ierror Then Exit Sub

Call IOStatusAuto("Exporting Standard " & Str$(StandardIndexNumbers%(arrayindex%(i%))) & " (" & Str$(i%) & " of " & Str$(NumberOfAvailableStandards%) & ")")
DoEvents
If icancelauto Then
ierror = True
Exit Sub
End If

' Write data to ASCII standard file
If mode% = 1 Then
Call StandardWriteDATSample(StandardTmpSample())
If ierror Then Exit Sub

Else
Call StandardWriteDATSample2(StandardTmpSample())
If ierror Then Exit Sub
End If

Next i%

' Confirm with user
msg$ = "Export file " & ImportDataFile$ & " containing standard compositions was exported."
MsgBox msg$, vbOKOnly + vbInformation, "StandardWriteDATFile"

Exit Sub

' Errors
StandardWriteDATFileError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardWriteDATFile"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub StandardWriteDATFile2(tForm As Form)
' This routine sends the exported data to an Excel file

ierror = False
On Error GoTo StandardWriteDATFile2Error

Dim response As Integer
Dim Excel_Version As Single

Dim filenamearray(1 To 1) As String

' Get Excel version
Excel_Version! = ExcelVersionNumber!()
If ierror Then Exit Sub

' Check if user wants to send to excel
If Excel_Version! < 12# Then
msg$ = "Do you want to send the export file to Excel? Note that because of the 256 column limit in older versions of Excel, Excel 12 or later is required for this export feature."
Else
msg$ = "Do you want to send the export file to Excel?"
End If
response% = MsgBox(msg$, vbYesNo + vbQuestion + vbDefaultButton1, "StandardWriteDATFile2")
If response% = vbNo Then Exit Sub

' Send all files to excel
filenamearray$(1) = ImportDataFile$
Call ExcelSendFileListToExcel(Int(1), filenamearray$(), tForm)
If ierror Then Exit Sub

Exit Sub

' Errors
StandardWriteDATFile2Error:
MsgBox Error$, vbOKOnly + vbCritical, "StandardWriteDATFile2"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub StandardWriteDATSample(sample() As TypeSample)
' Write standard compositions to STANDARD.DAT file. Called by StandardWriteDATFile

ierror = False
On Error GoTo StandardWriteDATSampleError

Dim i As Integer

' Write standard number and name
sample(1).Name$ = Replace$(sample(1).Name$, VbDquote$, VbSquote$)
Print #ImportDataFileNumber%, sample(1).number%, VbDquote$ & sample(1).Name$ & VbDquote$

' Write standard description (beginning in v. 2.48)
sample(1).Description$ = Replace$(sample(1).Description$, VbDquote$, VbSquote$)
Print #ImportDataFileNumber%, VbDquote$ & sample(1).Description$ & VbDquote$

' Write DisplayAsOxideFlag flag and number of elements in standard
Print #ImportDataFileNumber%, sample(1).DisplayAsOxideFlag%, sample(1).LastChan%

' Write standard density
Print #ImportDataFileNumber%, sample(1).SampleDensity!

' Write standard elements symbols
msg$ = vbNullString
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(VbDquote$ & sample(1).Elsyms$(i%) & VbDquote$ & " ", a80$)
Next i%
Print #ImportDataFileNumber%, msg$

' Write standard cations and oxygens
msg$ = vbNullString
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(sample(1).numcat%(i%), a80$)
Next i%
Print #ImportDataFileNumber%, msg$

msg$ = vbNullString
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(sample(1).numoxd%(i%), a80$)
Next i%
Print #ImportDataFileNumber%, msg$

' Write standard composition in elemental weight percents
msg$ = vbNullString
For i% = 1 To sample(1).LastChan%
msg$ = msg$ & Format$(Format$((sample(1).ElmPercents!(i%)), f83$), a80$)
Next i%
Print #ImportDataFileNumber%, msg$

' Write material types, formula flags, etc.
Print #ImportDataFileNumber%, VbDquote$ & sample(1).MaterialType$ & VbDquote$, sample(1).FormulaElementFlag, sample(1).FormulaRatio!, VbDquote$ & sample(1).FormulaElement$ & VbDquote$, VbDquote$ & sample(1).MountNames$ & VbDquote

Exit Sub

' Errors
StandardWriteDATSampleError:
MsgBox Error$, vbOKOnly + vbCritical, "StandardWriteDATSample"
ierror = True
Exit Sub

End Sub

Sub StandardWriteDATSample2(sample() As TypeSample)
' Write standard compositions to STANDARD2.DAT file (single row format). Called by StandardWriteDATFile

ierror = False
On Error GoTo StandardWriteDATSample2Error

Dim i As Integer, n As Integer
Dim tmsg As String
Dim astring As String, bstring As String
Dim cstring As String, dstring As String
Dim estring As String

' Write standard number and name
sample(1).Name$ = Replace$(sample(1).Name$, VbDquote$, VbSquote$)
astring$ = sample(1).number% & vbTab & VbDquote$ & sample(1).Name$ & VbDquote$ & vbTab

' Write standard description (beginning in v. 2.48)
sample(1).Description$ = Replace$(sample(1).Description$, VbDquote$, VbSquote$)
astring$ = astring$ & VbDquote$ & sample(1).Description$ & VbDquote$ & vbTab

' Write DisplayAsOxideFlag flag in standard
astring$ = astring$ & sample(1).DisplayAsOxideFlag% & vbTab

' Write standard density
astring$ = astring$ & sample(1).SampleDensity! & vbTab

' Write standard elements weight percents in atomic number order
For n% = 1 To MAXELM%
tmsg$ = "n.a."
For i% = 1 To sample(1).LastChan%
If MiscStringsAreSame(sample(1).Elsyms$(i%), Symlo$(n%)) Then
tmsg$ = MiscAutoFormat$(sample(1).ElmPercents!(i%))
End If
Next i%
bstring$ = bstring$ & tmsg$ & vbTab
Next n%

' Write standard cations and oxygens
For n% = 1 To MAXELM%
tmsg$ = Format$(AllCat%(n%))
For i% = 1 To sample(1).LastChan%
If MiscStringsAreSame(sample(1).Elsyms$(i%), Symlo$(n%)) Then
tmsg$ = Format$(sample(1).numcat%(i%))
End If
Next i%
cstring$ = cstring$ & tmsg$ & vbTab
Next n%

For n% = 1 To MAXELM%
tmsg$ = Format$(AllOxd%(n%))
For i% = 1 To sample(1).LastChan%
If MiscStringsAreSame(sample(1).Elsyms$(i%), Symlo$(n%)) Then
tmsg$ = Format$(sample(1).numoxd%(i%))
End If
Next i%
dstring$ = dstring$ & tmsg$ & vbTab
Next n%

' Write material type, flormula flags, etc
estring$ = VbDquote$ & sample(1).MaterialType$ & VbDquote$ & vbTab & sample(1).FormulaElementFlag & vbTab & sample(1).FormulaRatio! & vbTab & VbDquote$ & sample(1).FormulaElement$ & VbDquote$ & vbTab & VbDquote$ & sample(1).MountNames$ & VbDquote & vbTab

' Write the line to the file
Print #ImportDataFileNumber%, astring$ & bstring$ & cstring$ & dstring$ & estring$
Exit Sub

' Errors
StandardWriteDATSample2Error:
MsgBox Error$, vbOKOnly + vbCritical, "StandardWriteDATSample2"
ierror = True
Exit Sub

End Sub

Sub StandardImportCameca(tForm As Form)
' Import selected Cameca PeakSight sx.mdb file and add to currently open standard composition database.

ierror = False
On Error GoTo StandardImportCamecaError

Dim StDb As Database
Dim StDs1 As Recordset, StDs2 As Recordset, StDs3 As Recordset
Dim SQLQ As String, tfilename As String

' Initialize the Tmp sample
Call InitSample(StandardTmpSample())
If ierror Then Exit Sub

msg$ = "Browse to the folder containing the Cameca standard sx.mdb file"
MsgBox msg$, vbOKOnly + vbInformation, "StandardImportCameca"

' Get path to sx.mdb (using DAO 3.6 supports Access 2000 database format which is required for Cameca PeakSight)
tfilename$ = "sx.mdb"
Call IOGetMDBFileName(Int(8), tfilename$, tForm)
If ierror Then Exit Sub

' Open the sx.mdb database
Screen.MousePointer = vbHourglass
Set StDb = OpenDatabase(tfilename$, DatabaseNonExclusiveAccess%, dbReadOnly)
Set StDs1 = StDb.OpenRecordset("Labels", dbOpenSnapshot)

Do Until StDs1.EOF
If StDs1("LabelID") <> 0 And Trim$(vbNullString & StDs1("Label")) <> vbNullString Then

StandardTmpSample(1).number% = StDs1("LabelID")                                 ' use "LabelID" so it matches .POS file standard numbers
StandardTmpSample(1).Name$ = Trim$(vbNullString & StDs1("Label"))
StandardTmpSample(1).Description$ = Trim$(vbNullString & StDs1("Comment"))
StandardTmpSample(1).DisplayAsOxideFlag% = False
StandardTmpSample(1).SampleDensity! = DEFAULTDENSITY!                           ' StDs1("Density") does not work!

' Get standard composition data for specified standard from standard database
SQLQ$ = "SELECT Elements.* FROM Elements WHERE Elements.LabelID = " & Str$(StandardTmpSample(1).number%)
Set StDs2 = StDb.OpenRecordset(SQLQ$, dbOpenSnapshot, dbReadOnly)

' Load all elements from "Elements" table that matched the standard number
StandardTmpSample(1).LastChan% = 0
Do Until StDs2.EOF
If StDs2("AtomNum") >= 1 And StDs2("AtomNum") <= MAXELM% Then

' Check Phase.PhaseNum <> -2 for this Elements.PhaseID (coating element)
SQLQ$ = "SELECT Phases.PhaseNum FROM Phases WHERE Phases.PhaseID = " & StDs2("PhaseID")
Set StDs3 = StDb.OpenRecordset(SQLQ$, dbOpenSnapshot, dbReadOnly)
If StDs3("PhaseNum") = 0 Then

If StandardTmpSample(1).LastChan% + 1 > MAXCHAN% Then GoTo StandardImportCamecaTooManyElements
StandardTmpSample(1).LastChan% = StandardTmpSample(1).LastChan% + 1
StandardTmpSample(1).Elsyms$(StandardTmpSample(1).LastChan%) = Symlo$(StDs2("AtomNum"))
StandardTmpSample(1).ElmPercents!(StandardTmpSample(1).LastChan%) = StDs2("Weight") * 100#
StandardTmpSample(1).numcat%(StandardTmpSample(1).LastChan%) = AllCat%(StDs2("AtomNum"))
StandardTmpSample(1).numoxd%(StandardTmpSample(1).LastChan%) = AllOxd%(StDs2("AtomNum"))
If StDs2("AtomNum") = 8 And StDs2("Weight") > 0.05 Then StandardTmpSample(1).DisplayAsOxideFlag% = True    ' assume oxide display if oxygen > 5%
End If

End If
StDs2.MoveNext
Loop
StDs2.Close

' Add the standard to the database (based on LabelID)
If StandardTmpSample(1).LastChan% > 0 Then
Call StandardAddRecord(StandardTmpSample())
If ierror Then Exit Sub

' Update available standard list
If NumberOfAvailableStandards% + 1 > MAXINDEX% Then GoTo StandardImportCamecaTooMany
NumberOfAvailableStandards% = NumberOfAvailableStandards% + 1
StandardIndexNumbers%(NumberOfAvailableStandards%) = StandardTmpSample(1).number%
StandardIndexNames$(NumberOfAvailableStandards%) = StandardTmpSample(1).Name$
StandardIndexDescriptions$(NumberOfAvailableStandards%) = StandardTmpSample(1).Description$
StandardIndexDensities!(NumberOfAvailableStandards%) = StandardTmpSample(1).SampleDensity!
StandardIndexMaterialTypes$(NumberOfAvailableStandards%) = StandardTmpSample(1).MaterialType$
StandardIndexMountNames$(NumberOfAvailableStandards%) = StandardTmpSample(1).MountNames$
End If

End If
StDs1.MoveNext
Loop

' Close the SX.MDB standard database
StDs1.Close
StDb.Close

Screen.MousePointer = vbDefault

msg$ = "Standard compositions imported from Cameca PeakSight standard database " & tfilename$ & vbCrLf & vbCrLf
msg$ = msg$ & "Next, (optionally) use the Stage application to import the Cameca stage positions for each mount and save to .POS files."
MsgBox msg$, vbOKOnly + vbInformation, "StandardImportCameca"
Exit Sub

' Errors
StandardImportCamecaError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardImportCameca"
ierror = True
Exit Sub

StandardImportCamecaTooManyElements:
Screen.MousePointer = vbDefault
msg$ = "Too many elements in standard number " & Str$(StandardTmpSample(1).number%) & " in " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardImportCameca"
ierror = True
Exit Sub

StandardImportCamecaTooMany:
Screen.MousePointer = vbDefault
msg$ = "Too many standards were found in " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardImportCameca"
ierror = True
Exit Sub

End Sub

Sub StandardImportJEOL(tForm As Form)
' Import selected JEOL 8200/8500 import file(s) and add to currently open standard composition database (needs rework)

ierror = False
On Error GoTo StandardImportJEOLError

Dim tfilename As String, tpath As String, tstring As String

Dim nCount As Long, n As Long, m As Long
Dim sAllFiles() As String
Dim arrayfilenames(1 To 1) As String

' Initialize the Tmp sample
Call InitSample(StandardTmpSample())
If ierror Then Exit Sub

' Get path to standard folder (each sub folder is a standard block/mount)
tpath$ = "C:\"
tstring$ = "Browse to Folder Containing JEOL 8200/8500 Standard Block/Mount Files"
tpath$ = IOBrowseForFolderByPath(False, tpath$, tstring$, tForm)
If ierror Then Exit Sub
If Trim$(tpath$) = vbNullString Then Exit Sub

' Get all JEOL standard files recursively
Call DirectorySearch("*.cmp", tpath$, True, nCount&, sAllFiles$())                  ' what extension does 8200/8500 use?
If ierror Then Exit Sub
If nCount& < 1 Then GoTo StandardImportJEOLNoFiles

Screen.MousePointer = vbHourglass
For n& = 1 To nCount&
tfilename$ = sAllFiles$(n&)

' Open the ASCII file and import the data to the standard database
Open tfilename$ For Input As #ImportDataFileNumber%

' Parse data
Call StandardImportJEOLParseFile(ImportDataFileNumber%, tfilename$, StandardTmpSample())
If ierror Then
Close #ImportDataFileNumber%
Exit Sub
End If

' Check if standard contains elements
If StandardTmpSample(1).LastChan% > 0 Then

' Get next available standard number
StandardTmpSample(1).number% = StandardGetNumber()
If ierror Then
Close #ImportDataFileNumber%
Exit Sub
End If

' Add the standard to the database
Call StandardAddRecord(StandardTmpSample())
If ierror Then
Close #ImportDataFileNumber%
Exit Sub
End If

' Update available standard list
If NumberOfAvailableStandards% + 1 > MAXINDEX% Then GoTo StandardImportJEOLTooMany
NumberOfAvailableStandards% = NumberOfAvailableStandards% + 1
StandardIndexNumbers%(NumberOfAvailableStandards%) = StandardTmpSample(1).number%
StandardIndexNames$(NumberOfAvailableStandards%) = StandardTmpSample(1).Name$
StandardIndexDescriptions$(NumberOfAvailableStandards%) = StandardTmpSample(1).Description$
StandardIndexDensities!(NumberOfAvailableStandards%) = StandardTmpSample(1).SampleDensity!
StandardIndexMaterialTypes$(NumberOfAvailableStandards%) = StandardTmpSample(1).MaterialType$
StandardIndexMountNames$(NumberOfAvailableStandards%) = StandardTmpSample(1).MountNames$
End If

' Close the import file
Close #ImportDataFileNumber%
Next n&

Screen.MousePointer = vbDefault
Exit Sub

' Errors
StandardImportJEOLError:
Close #ImportDataFileNumber%
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardImportJEOL"
ierror = True
Exit Sub

StandardImportJEOLNoFiles:
msg$ = "No standard files were found in any folders or sub folders"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardImportJEOL"
ierror = True
Exit Sub

StandardImportJEOLTooMany:
Close #ImportDataFileNumber%
Screen.MousePointer = vbDefault
msg$ = "Too many standards were found in " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardImportJEOL"
ierror = True
Exit Sub

End Sub

Sub StandardImportJEOLParseFile(tImportDataFileNumber%, tfilename As String, sample() As TypeSample)
' Parse JEOL 8200/8500 standard import format

ierror = False
On Error GoTo StandardImportJEOLParseFileError

Dim i As Integer
Dim ielm As Integer
Dim astring As String, bstring As String

ReDim temp1(1 To MAXCHAN%) As Single
ReDim temp2(1 To MAXCHAN%) As Single






Exit Sub

' Errors
StandardImportJEOLParseFileError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardImportJEOLParseFile"
ierror = True
Exit Sub

StandardImportJEOLParseFileTooManyElements:
Screen.MousePointer = vbDefault
msg$ = "Too many elements in standard number " & Str$(sample(1).number%) & " in " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardImportJEOLParseFile"
ierror = True
Exit Sub

StandardImportJEOLParseFileBadDataFormat:
Screen.MousePointer = vbDefault
msg$ = "Invalid data format value in standard number " & Str$(sample(1).number%) & " in " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardImportJEOLParseFile"
ierror = True
Exit Sub

StandardImportJEOLParseFileBadElement:
Screen.MousePointer = vbDefault
msg$ = "Invalid element number in standard number " & Str$(sample(1).number%) & " in " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardImportJEOLParseFile"
ierror = True
Exit Sub

End Sub

Sub StandardImportJEOL8x30(tForm As Form)
' Import JEOL 8230/8530 standard files and add to currently open standard composition database

ierror = False
On Error GoTo StandardImportJEOL8x30Error

Dim response As Integer
Dim tstring As String
Dim tfilename As String, tfilename2 As String

Dim nCount As Long, n As Long, m As Long
Dim sAllFiles() As String
Dim arrayfilenames(1 To 1) As String

Static tpath As String

' Initialize the Tmp sample
Call InitSample(StandardTmpSample())
If ierror Then Exit Sub

' Get path to standard folder (each sub folder is a standard block/mount)
tpath$ = "C:\"
tstring$ = "Browse to the folder containing the JEOL 8230/8530/iSP/iHP200F standard block/mount folders (usually StdCmp)"
tpath$ = IOBrowseForFolderByPath(False, tpath$, tstring$, tForm)
If ierror Then Exit Sub
If Trim$(tpath$) = vbNullString Then Exit Sub

' Get all JEOL Cmp files recursively
Call DirectorySearch("*.cmp", tpath$, True, nCount&, sAllFiles$())
If ierror Then Exit Sub
If nCount& < 1 Then GoTo StandardImportJEOL8x30NoFiles

' Ask user whether to include block/mount folder in standard name
msg$ = "Would you like to include the block/mount folder name in the standard name?" & vbCrLf & vbCrLf
msg$ = msg$ & "Click Yes to include the mount name in the standard name, click No to specify the mount name as the mount name parameter."
response% = MsgBox(msg$, vbYesNoCancel + vbQuestion + vbDefaultButton2, "StandardImportJEOL8x30")
If response% = vbCancel Then Exit Sub

Screen.MousePointer = vbHourglass
For n& = 1 To nCount&
tfilename$ = sAllFiles$(n&)

' Open the ASCII file and import the data to the standard database
Open tfilename$ For Input As #ImportDataFileNumber%

' Parse data
Call StandardImportJEOL8x30ParseFile(ImportDataFileNumber%, tfilename$, response%, StandardTmpSample())
If ierror Then
Close #ImportDataFileNumber%
Exit Sub
End If

' Check if standard contains elements
If StandardTmpSample(1).LastChan% > 0 Then

' Get next available standard number
StandardTmpSample(1).number% = StandardGetNumber()
If ierror Then
Close #ImportDataFileNumber%
Exit Sub
End If

' Add the standard to the database
Call StandardAddRecord(StandardTmpSample())
If ierror Then
Close #ImportDataFileNumber%
Exit Sub
End If

' Update available standard list
If NumberOfAvailableStandards% + 1 > MAXINDEX% Then GoTo StandardImportJEOL8x30TooMany
NumberOfAvailableStandards% = NumberOfAvailableStandards% + 1
StandardIndexNumbers%(NumberOfAvailableStandards%) = StandardTmpSample(1).number%
StandardIndexNames$(NumberOfAvailableStandards%) = StandardTmpSample(1).Name$
StandardIndexDescriptions$(NumberOfAvailableStandards%) = StandardTmpSample(1).Description$
StandardIndexDensities!(NumberOfAvailableStandards%) = StandardTmpSample(1).SampleDensity!
StandardIndexMaterialTypes$(NumberOfAvailableStandards%) = StandardTmpSample(1).MaterialType$
StandardIndexMountNames$(NumberOfAvailableStandards%) = StandardTmpSample(1).MountNames$
End If

' Close the import file
Close #ImportDataFileNumber%

' Now write a small text file with the standard number and stage positions (even for standards that have no elements)
tfilename2$ = MiscGetFileNameNoExtension(tfilename$) & ".txt"
Open tfilename2$ For Output As #ExportDataFileNumber%
Write #ExportDataFileNumber%, StandardTmpSample(1).number%, StandardTmpSample(1).Name$, StandardTmpSample(1).StagePositions!(1, 1), StandardTmpSample(1).StagePositions!(1, 2), StandardTmpSample(1).StagePositions!(1, 3)
Close #ExportDataFileNumber%

Next n&

Screen.MousePointer = vbDefault
msg$ = "All JEOL standard compositions saved to the current standard composition database (" & StandardDataFile$ & ")." & vbCrLf & vbCrLf
msg$ = msg$ & "Next, (optionally) use the Stage application to import the JEOL stage positions for each mount and save to .POS files."
MsgBox msg$, vbOKOnly + vbInformation, "StandardImportJEOL8x30"
Exit Sub

' Errors
StandardImportJEOL8x30Error:
Close #ImportDataFileNumber%
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardImportJEOL8x30"
ierror = True
Exit Sub

StandardImportJEOL8x30NoFiles:
msg$ = "No .cmp files were found in any folders or sub folders"
MsgBox msg$, vbOKOnly + vbExclamation, "StandardImportJEOL8x30"
ierror = True
Exit Sub

StandardImportJEOL8x30TooMany:
Close #ImportDataFileNumber%
Screen.MousePointer = vbDefault
msg$ = "Too many standards were found in " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardImportJEOL8x30"
ierror = True
Exit Sub

End Sub

Sub StandardImportJEOL8x30ParseFile(tImportDataFileNumber%, tfilename As String, response As Integer, sample() As TypeSample)
' Parse JEOL 8230/8530 standard import format .Cmp files
'   mode = vbNo do not use block/mount folder in name
'   mode = vbYes do use block folder/mount in name

ierror = False
On Error GoTo StandardImportJEOL8x30ParseFileError

Dim materialflag As Integer, ratioflag As Integer, ivalence As Integer
Dim chan As Integer, ielm As Integer
Dim astring As String, bstring As String
Dim tparameter As String, tfolder As String

ReDim temp1(1 To MAXCHAN%) As Single
ReDim temp2(1 To MAXCHAN%) As Single

' Parse out the folder name
tfolder$ = MiscGetPathOnly$(tfilename$)
tfolder$ = MiscGetLastFolderName$(tfolder$)
If ierror Then Exit Sub

' Read standard name
tparameter$ = "$XM_CMP_STD_SAMPLE_NAME"
Call StandardImportJEOL8x30ParseFile2(tfilename$, tparameter$, astring$)
If ierror Then Exit Sub

sample(1).Name$ = Trim$(astring$)

' Read standard comment
tparameter$ = "$XM_CMP_COMMENT"
Call StandardImportJEOL8x30ParseFile2(tfilename$, tparameter$, astring$)
If ierror Then Exit Sub

sample(1).Description$ = Trim$(astring$)

' Create name from name and comment and block/mount folder (if specified)
If response% = vbYes Then
If sample(1).Description$ <> vbNullString Then
sample(1).Name$ = tfolder$ & "_" & Trim(sample(1).Name$ & "_" & sample(1).Description$)
Else
sample(1).Name$ = tfolder$ & "_" & Trim(sample(1).Name$)
End If

Else
If sample(1).Description$ <> vbNullString Then
sample(1).Name$ = Trim(sample(1).Name$ & "_" & sample(1).Description$)
Else
sample(1).Name$ = Trim(sample(1).Name$)
End If
sample(1).MountNames$ = tfolder$
End If

' Make sure that sample name does not contain any embedded double quotes
sample(1).Name$ = Replace$(sample(1).Name$, VbDquote$, VbSquote$)
If ierror Then Exit Sub

' Read material type flag (0 = elemental, 1 = oxide)
tparameter$ = "$XM_CMP_MATERIAL"
Call StandardImportJEOL8x30ParseFile2(tfilename$, tparameter$, astring$)
If ierror Then Exit Sub

If Val(astring$) = 0 Then
materialflag% = 0               ' elemental data
Else
materialflag% = 1               ' oxide data
End If
sample(1).OxideOrElemental% = 2         ' standards are always stored as elemental compositions

' Get ratio type flag (0 = elemental or oxide percents, 1 = formula atoms)?
tparameter$ = "$XM_CMP_RATIO_TYPE"
Call StandardImportJEOL8x30ParseFile2(tfilename$, tparameter$, astring$)
If ierror Then Exit Sub
ratioflag% = Val(astring$)

' Get stage positions for subsequent saving in .POS files (See StagForm.bas in Stage app)
tparameter$ = "$XM_CMP_STAGE_POS"
Call StandardImportJEOL8x30ParseFile2(tfilename$, tparameter$, astring$)
If ierror Then Exit Sub

' Parse stage data
Call MiscParseStringToString(astring$, bstring$)    ' get X position
sample(1).StagePositions!(1, 1) = Val(bstring$)
Call MiscParseStringToString(astring$, bstring$)    ' get Y position
sample(1).StagePositions!(1, 2) = Val(bstring$)
Call MiscParseStringToString(astring$, bstring$)    ' get Z position
sample(1).StagePositions!(1, 3) = Val(bstring$)

' Number of elements
tparameter$ = "$XM_CMP_NUMBER_OF_DATA"
Call StandardImportJEOL8x30ParseFile2(tfilename$, tparameter$, astring$)
If ierror Then Exit Sub

' Check for a valid number of elements
sample(1).LastChan% = Val(astring$)
If sample(1).LastChan% = 0 Then
msg$ = "No elements defined for standard " & sample(1).Name$ & " in " & tfilename$ & vbCrLf & vbCrLf
msg$ = msg$ & "The standard will not be imported into the standard composition database."
MsgBox msg$, vbOKOnly + vbExclamation, "StandardImportJEOL8x30ParseFile"
Exit Sub
End If
If sample(1).LastChan% > MAXCHAN% Then GoTo StandardImportJEOL8x30ParseFileTooManyElements

' Loop on elements
For chan% = 1 To sample(1).LastChan%

tparameter$ = "$XM_CMP_COMPOSITION%" & Format$(chan% - 1)
Call StandardImportJEOL8x30ParseFile2(tfilename$, tparameter$, astring$)
If ierror Then Exit Sub

' Parse element, valence, and concentration
Call MiscParseStringToString(astring$, bstring$)    ' parse element symbol
ielm% = IPOS1%(MAXELM%, bstring$, Symlo$())
If ielm% = 0 Then GoTo StandardImportJEOL8x30ParseFileBadElement
sample(1).Elsyms$(chan%) = Symlo$(ielm%)

Call MiscParseStringToString(astring$, bstring$)    ' parse valence
ivalence% = Val(bstring$)                           ' not utilized yet!

Call MiscParseStringToString(astring$, bstring$)    ' parse concentration/atoms
temp1!(chan%) = Val(bstring$)

' Load other parameters
sample(1).numcat%(chan%) = AllCat%(ielm%)
sample(1).numoxd%(chan%) = AllOxd%(ielm%)

Next chan%

' Load default density
sample(1).SampleDensity! = DEFAULTDENSITY!

' If oxide data, add oxygen to elements
If materialflag% = 1 Then
sample(1).LastChan% = sample(1).LastChan% + 1
sample(1).Elsyms$(sample(1).LastChan%) = Symlo$(ATOMIC_NUM_OXYGEN%)
ivalence% = -2                                      ' not utilized yet!

' Load other parameters
sample(1).numcat%(sample(1).LastChan%) = AllCat%(ATOMIC_NUM_OXYGEN%)
sample(1).numoxd%(sample(1).LastChan%) = AllOxd%(ATOMIC_NUM_OXYGEN%)
End If

' Convert atoms to elemental concentrations
If ratioflag% = 1 Then

' Load temp array for atom to elemental conversion
For chan% = 1 To sample(1).LastChan%
temp2(chan%) = temp1(chan%)
Next chan%

' Convert atoms to weight percent
For chan% = 1 To sample(1).LastChan%
temp1!(chan%) = ConvertAtomToWeight(sample(1).LastChan%, chan%, temp2!(), sample(1).Elsyms$())
If ierror Then Exit Sub
Next chan%
End If

' Convert all to elemental concentrations
For chan% = 1 To sample(1).LastChan%

' Elemental material, just load concentrations
If materialflag% = 0 Then
temp2!(chan%) = temp1!(chan%)

' Oxide material, convert to elemental
ElseIf materialflag% = 1 Then
temp2!(chan%) = ConvertOxdToElm!(temp1!(chan%), sample(1).Elsyms$(chan%), sample(1).numcat%(chan%), sample(1).numoxd%(chan%))
End If

' Load elemental concentrations
sample(1).ElmPercents!(chan%) = temp2!(chan%)

' Assume oxide display if oxygen > 5%
If UCase$(Trim$(sample(1).Elsyms$(chan%))) = UCase$(Trim$(Symlo$(ATOMIC_NUM_OXYGEN%))) And sample(1).ElmPercents!(chan%) > 5# Then sample(1).DisplayAsOxideFlag% = True
Next chan%

' Now that oxides are converted to elemental concentrations, if oxide material, calculate oxygen from cations
If materialflag% = 1 Then
sample(1).OxygenChannel% = sample(1).LastChan%
sample(1).ElmPercents!(sample(1).OxygenChannel%) = ConvertOxygenFromCations!(sample())
If ierror Then Exit Sub
End If

Exit Sub

' Errors
StandardImportJEOL8x30ParseFileError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardImportJEOL8x30ParseFile"
ierror = True
Exit Sub

StandardImportJEOL8x30ParseFileTooManyElements:
Screen.MousePointer = vbDefault
msg$ = "Too many elements defined in " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardImportJEOL8x30ParseFile"
ierror = True
Exit Sub

StandardImportJEOL8x30ParseFileBadElement:
Screen.MousePointer = vbDefault
msg$ = "Invalid element symbol (" & astring$ & ") in " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardImportJEOL8x30ParseFile"
ierror = True
Exit Sub

End Sub

Sub StandardImportJEOL8x30ParseFile2(tfilename As String, tparameter As String, treturn As String)
' Read the 8x30 JEOL .cmp file and load specified parameter

ierror = False
On Error GoTo StandardImportJEOL8x30ParseFile2Error

Dim n As Integer
Dim tfilenumber As Integer
Dim astring As String, bstring As String

tfilenumber% = FreeFile()
treturn$ = vbNullString

' Read each parameter by looping through the .cnd file
Open tfilename$ For Input As tfilenumber%

Do Until EOF(tfilenumber%)
Line Input #tfilenumber%, astring$

If InStr(astring$, tparameter$) > 0 Then
treturn$ = Trim$(Mid$(astring$, Len(tparameter$) + 2))
Close tfilenumber%
Exit Sub
End If

Loop

Close tfilenumber%

' If we get to here something is wrong because the parameter was not found
msg$ = "Unable to find parameter (" & tparameter$ & ") in file " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardImportJEOL8x30ParseFile2"
ierror = True

Exit Sub

' Errors
StandardImportJEOL8x30ParseFile2Error:
Close tfilenumber%
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "StandardImportJEOL8x30ParseFile2"
ierror = True
Exit Sub

End Sub

Sub StandardInputCalcZAFStandardFormatKratios(tForm As Form)
' Input k-ratios from a CalcZAF standard format file (*.dat)

ierror = False
On Error GoTo StandardInputCalcZAFStandardFormatKratiosError

Dim i As Integer, ip As Integer
Dim stdnum As Integer, stdrow As Integer
Dim response As Integer

Dim tfilename As String
Dim tfilenumber As Integer
Dim tlinenumber As Integer

Dim tKratios(1 To MAXCHAN%) As Single
Dim tPercents(1 To MAXCHAN%) As Single

' Ask user for file name
Call IOGetFileName(Int(2), "DAT", tfilename$, tForm)
If ierror Then Exit Sub

tfilenumber% = FreeFile()
Open tfilename$ For Input As #tfilenumber%

' Loop on all lines
tlinenumber% = 0
Do Until EOF(tfilenumber%)
tlinenumber% = tlinenumber% + 1

' Read name, kilovolts and takeoff
Input #tfilenumber%, StandardTmpSample(1).Name$, StandardTmpSample(1).takeoff!, StandardTmpSample(1).kilovolts!

' Check if standard name and number exists in standard database
ip% = IPOS1%(NumberOfAvailableStandards%, StandardTmpSample(1).Name$, StandardIndexNames$())
If ip% = 0 Then GoTo StandardInputCalcZAFStandardFormatKratiosNameNotFound

' Check standard number and confirm with user
stdnum% = StandardIndexNumbers%(ip%)

' Get standard table row
stdrow% = StandardGetRow%(stdnum%)
If ierror Then Exit Sub

' Read oxide flag, number of elements
Input #tfilenumber%, StandardTmpSample(1).OxideOrElemental%, StandardTmpSample(1).LastChan%
If StandardTmpSample(1).LastChan% < 1 Or StandardTmpSample(1).LastChan% > MAXCHAN% Then GoTo StandardInputCalcZAFStandardFormatKratiosBadLastChan

' Loop on each element
For i% = 1 To StandardTmpSample(1).LastChan%
Input #tfilenumber%, StandardTmpSample(1).Elsyms$(i%)
Next i%

' Loop on each xray
For i% = 1 To StandardTmpSample(1).LastChan%
Input #tfilenumber%, StandardTmpSample(1).Xrsyms$(i%)
Next i%

' Loop on raw k-ratios
For i% = 1 To StandardTmpSample(1).LastChan%
Input #tfilenumber%, tKratios!(i%)
Next i%

' Loop on standard weight% (published)
For i% = 1 To StandardTmpSample(1).LastChan%
Input #tfilenumber%, tPercents!(i%)
Next i%

' Loop on std assignments
For i% = 1 To StandardTmpSample(1).LastChan%
Input #tfilenumber%, StandardTmpSample(1).StdAssigns%(i%)
Next i%

' Cations/oxides
For i% = 1 To StandardTmpSample(1).LastChan%
Input #tfilenumber%, StandardTmpSample(1).numcat%(i%)
Next i%

For i% = 1 To StandardTmpSample(1).LastChan%
Input #tfilenumber%, StandardTmpSample(1).numoxd%(i%)
Next i%

' Confirm with user
msg$ = "Are you sure that you want to add these k-ratios (input line: " & Format$(tlinenumber%) & ")" & vbCrLf
msg$ = msg$ & "to standard: " & Format$(stdnum%) & ", " & StandardTmpSample(1).Name$ & vbCrLf
msg$ = msg$ & "in: " & StandardDataFile$ & "?"
response% = MsgBox(msg$, vbYesNoCancel + vbQuestion + vbDefaultButton2, "StandardInputCalcZAFStandardFormatKratios")
If response% = vbCancel Then Exit Sub

' Save input parameters to k-ratio table in standard database
If response% = vbYes Then





End If
Loop

Close #tfilenumber%

' Confirm input
msg$ = "All standard k-ratios in input file, " & tfilename$ & ", were imported into standard database: " & StandardDataFile$
MsgBox msg$, vbOKOnly + vbInformation, "StandardInputCalcZAFStandardFormatKratios"

Exit Sub

' Errors
StandardInputCalcZAFStandardFormatKratiosError:
Close #tfilenumber%
MsgBox Error$, vbOKOnly + vbCritical, "StandardInputCalcZAFStandardFormatKratios"
Call AnalyzeStatusAnal(vbNullString)
ierror = True
Exit Sub

StandardInputCalcZAFStandardFormatKratiosNameNotFound:
Close #tfilenumber%
msg$ = "Standard name, " & StandardTmpSample(1).Name$ & ", was not found in " & StandardDataFile$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardInputCalcZAFStandardFormatKratios"
ierror = True
Exit Sub

StandardInputCalcZAFStandardFormatKratiosBadLastChan:
Close #tfilenumber%
msg$ = "Standard name, " & StandardTmpSample(1).Name$ & ", has an invalid number of channels in input file " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "StandardInputCalcZAFStandardFormatKratios"
ierror = True
Exit Sub

End Sub

