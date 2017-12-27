Attribute VB_Name = "CodeFILEINFO2"
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

Const DbTextFileInfoTypeFieldLength% = 8
Const DbTextFileInfoTitleFieldLength% = 64
Const DbTextFileInfoUserFieldLength% = 64
Const DbTextFileInfoCustomFieldLength% = 64
Const DbTextFileInfoDescriptionFieldLength% = 128

Function FileInfoGetVersion(tfilename As String, tfiletype As String) As Single
' Return version number for the specified database filename and filetype

ierror = False
On Error GoTo FileInfoGetVersionError

Dim fileattributes As Integer, attributetest As Integer

Dim FiDb As Database
Dim FiDt As Recordset

' Check for valid database name
If Trim$(tfilename$) = vbNullString Then GoTo FileInfoGetVersionNoFileName

' Check for read only on passed database file
fileattributes% = GetAttr(tfilename$)
attributetest% = fileattributes% And vbReadOnly
If attributetest% > 0 Then GoTo FileInfoGetVersionReadOnly

' Open the passed database for read only
Screen.MousePointer = vbHourglass
Set FiDb = OpenDatabase(tfilename$, DatabaseNonExclusiveAccess%, dbReadOnly)
Set FiDt = FiDb.OpenRecordset("File", dbOpenSnapshot)

' Check file type
If tfiletype$ <> FiDt("Type") Then GoTo FileInfoGetVersionWrongType

' Open the File table, check for newer version number
If FiDt("Version") > ProgramVersionNumber! Then GoTo FileInfoGetVersionOldVersion

' Check for older VB3 16 bit database versions
If FiDt("Version") < 4# Then
If tfiletype$ = "PROBE" Then
'If RealTimeMode Then GoTo FileInfoGetVersion16BitVersionProbe
Else
GoTo FileInfoGetVersion16BitVersion
End If
End If

' Return version number of database file
FileInfoGetVersion! = FiDt("Version")

FiDt.Close
FiDb.Close

Screen.MousePointer = vbDefault
Exit Function

' Errors
FileInfoGetVersionError:
Screen.MousePointer = vbDefault
MsgBox Error$ & ", accessing file " & tfilename$ & ", type: " & tfiletype$, vbOKOnly + vbCritical, "FileInfoGetVersion"
ierror = True
Exit Function

FileInfoGetVersionNoFileName:
Screen.MousePointer = vbDefault
msg$ = "Database file name is blank"
MsgBox msg$, vbOKOnly + vbExclamation, "FileInfoGetVersion"
ierror = True
Exit Function

FileInfoGetVersionReadOnly:
Screen.MousePointer = vbDefault
msg$ = "File " & tfilename$ & " is a read only file (" & Format$(fileattributes%) & ") and cannot be opened for updating. Please check the Read-only file attributes in the Properties tab by right clicking on the file."
MsgBox msg$, vbOKOnly + vbExclamation, "FileInfoGetVersion"
ierror = True
Exit Function

FileInfoGetVersionWrongType:
Screen.MousePointer = vbDefault
msg$ = "File " & tfilename$ & " is not a " & tfiletype$ & " database"
MsgBox msg$, vbOKOnly + vbExclamation, "FileInfoGetVersion"
ierror = True
Exit Function

FileInfoGetVersionOldVersion:
Screen.MousePointer = vbDefault
msg$ = "The database file (" & tfilename$ & "), type " & tfiletype$ & ", version number v. " & Str$(FiDt("Version")) & " is "
msg$ = msg$ & "newer than the program version number (v. " & Str$(ProgramVersionNumber!) & ")." & vbCrLf & vbCrLf
msg$ = msg$ & "Please upgrade the Probe for EPMA application files, (specifically " & app.EXEName & ".exe) to v. " & Str$(FiDt("Version")) & " or higher and try again." & vbCrLf
msg$ = msg$ & vbCrLf & "To update the Probe for EPMA application files simply use the Help | Update Probe for EPMA menu and automatically download updated application files if you have an Internet connection. "
msg$ = msg$ & "Otherwise download from another computer which is conected to the Internet using the Help menu or from the Probe Software website. Or contact Probe Software for assistance."
MsgBox msg$, vbOKOnly + vbExclamation, "FileInfoGetVersion"
ierror = True
Exit Function

FileInfoGetVersion16BitVersion:
Screen.MousePointer = vbDefault
msg$ = "The database file version number (v. " & Str$(FiDt("Version")) & ") indicates that " & tfilename$ & " is a 16 bit database." & vbCrLf & vbCrLf
msg$ = msg$ = "Please convert the database file to the new 32 bit database format by using the appropriate Export/Import feature."
MsgBox msg$, vbOKOnly + vbExclamation, "FileInfoGetVersion"
ierror = True
Exit Function

'FileInfoGetVersion16BitVersionProbe:
'Screen.MousePointer = vbDefault
'msg$ = "The database file version number (v. " & Str$(FiDt("Version")) & ") indicates "
'msg$ = msg$ & "that " & tfilename$ & " is a 16 bit database. This database file may only be opened "
'msg$ = msg$ & " for off-line data processing with the 32 bit version."
'MsgBox msg$, vbOKOnly + vbExclamation, "FileInfoGetVersion"
'ierror = True
'Exit Function

End Function

Sub FileInfoLoadData(mode As Integer, tfilename As String)
' Load data from file table
' mode = 1  load standard database File table
' mode = 2  load probe database File table
' mode = 3  load setup database File table (primary standards)
' mode = 4  load user database File table
' mode = 5  load position database File table
' mode = 6  load xray database File table
' mode = 7  load setup database File table (MAN standards)
' mode = 8  load setup database File table (interference standards)
' mode = 9  matrix database File table (Penepma matrix k-ratios and factors)
' mode = 10 boundary database File table (Penepma boundary k-ratios and factors)
' mode = 11 pure element database File table (Penepma pure element intensities)

ierror = False
On Error GoTo FileInfoLoadDataError

Dim FiDb As Database
Dim FiDt As Recordset

Dim versionnumber As Single

' Get correct database filename
Call FileInfoGetMDBFileName(mode%, tfilename$)
If ierror Then Exit Sub

' Check that passed drive letter is valid
If Not InitIsDriveMediaPresent(tfilename$) Then  ' check if drive exists
msg$ = "The specified drive letter " & Left$(tfilename$, 2) & " does not exist, either insert the drive and/or media (if removable) containing the MDB data file and try again."
MsgBox msg$, vbOKOnly + vbExclamation, "FileInfoLoadData"
ierror = True
Exit Sub
End If

' Check that file exists
If Dir$(tfilename$) = vbNullString Then
If mode% = 1 Then               ' standard
GoTo FileInfoLoadDataNotFound
ElseIf mode% = 2 Then           ' probe
GoTo FileInfoLoadDataNotFound
ElseIf mode% = 3 Then           ' setup (primary)
GoTo FileInfoLoadDataNotFound
ElseIf mode% = 4 Then           ' user
Call UserOpenNewFile(tfilename$)
ElseIf mode% = 5 Then           ' position
GoTo FileInfoLoadDataNotFound
ElseIf mode% = 6 Then           ' xray
GoTo FileInfoLoadDataNotFound
ElseIf mode% = 7 Then           ' setup (MAN)
GoTo FileInfoLoadDataNotFound
ElseIf mode% = 8 Then           ' setup (interference)
GoTo FileInfoLoadDataNotFound
ElseIf mode% = 9 Then           ' matrix (Penepm12 k-ratios and alpha factors)
GoTo FileInfoLoadDataNotFound
ElseIf mode% = 10 Then          ' boundary (Penepm12 k-ratios and alpha factors)
GoTo FileInfoLoadDataNotFound
ElseIf mode% = 11 Then          ' pure (Penepm12 pure element intensities)
GoTo FileInfoLoadDataNotFound
End If
End If

' Check data file version number
If mode% = 1 Then
versionnumber! = FileInfoGetVersion!(tfilename$, "STANDARD")
If ierror Then Exit Sub
ElseIf mode% = 2 Then
versionnumber! = FileInfoGetVersion!(tfilename$, "PROBE")
If ierror Then Exit Sub
ElseIf mode% = 3 Then
versionnumber! = FileInfoGetVersion!(tfilename$, "SETUP")
If ierror Then Exit Sub
ElseIf mode% = 4 Then
versionnumber! = FileInfoGetVersion!(tfilename$, "USER")
If ierror Then Exit Sub
ElseIf mode% = 5 Then
versionnumber! = FileInfoGetVersion!(tfilename$, "POSITION")
If ierror Then Exit Sub
ElseIf mode% = 6 Then
versionnumber! = FileInfoGetVersion!(tfilename$, "XRAY")
If ierror Then Exit Sub
ElseIf mode% = 7 Then
versionnumber! = FileInfoGetVersion!(tfilename$, "SETUP")
If ierror Then Exit Sub
ElseIf mode% = 8 Then
versionnumber! = FileInfoGetVersion!(tfilename$, "SETUP")
If ierror Then Exit Sub
ElseIf mode% = 9 Then
versionnumber! = FileInfoGetVersion!(tfilename$, "MATRIX")
If ierror Then Exit Sub
ElseIf mode% = 10 Then
versionnumber! = FileInfoGetVersion!(tfilename$, "BOUNDARY")
If ierror Then Exit Sub
ElseIf mode% = 11 Then
versionnumber! = FileInfoGetVersion!(tfilename$, "PURE")
If ierror Then Exit Sub
End If

' Open the passed database
Screen.MousePointer = vbHourglass
Set FiDb = OpenDatabase(tfilename$, DatabaseNonExclusiveAccess%, dbReadOnly)
Set FiDt = FiDb.OpenRecordset("File", dbOpenTable)    ' use dbOpenTable for .DateCreated and .DateUpdated properties below

' Load type and version number
DataFileVersionNumber! = FiDt("Version")
MDBFileType$ = Trim$(vbNullString & FiDt("Type"))

' Load file info in globals
MDBFileTitle$ = Trim$(vbNullString & FiDt("Title"))
MDBUserName$ = Trim$(vbNullString & FiDt("User"))
MDBFileDescription$ = Trim$(vbNullString & FiDt("Description"))

If DataFileVersionNumber! >= 2.05 Then
CustomText1$ = Trim$(vbNullString & FiDt("Custom1"))
CustomText2$ = Trim$(vbNullString & FiDt("Custom2"))
CustomText3$ = Trim$(vbNullString & FiDt("Custom3"))
End If

' Save date created, update and last modified
MDBFileCreated$ = FiDt.DateCreated
MDBFileUpdated$ = FiDt.LastUpdated
MDBFileModified$ = FileDateTime(tfilename$)

FiDt.Close
FiDb.Close

' Load to Probe data file global if loading Probe data file
If mode% = 2 Then
ProbeDataFileVersionNumber! = DataFileVersionNumber!
End If

Screen.MousePointer = vbDefault
Exit Sub

' Errors
FileInfoLoadDataError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "FileInfoLoadData"
ierror = True
Exit Sub

FileInfoLoadDataNotFound:
msg$ = "The selected database file " & tfilename$ & " (mode " & Format$(mode%) & ") was not found"
MsgBox msg$, vbOKOnly + vbExclamation, "FileInfoLoadData"
ierror = True
Exit Sub

End Sub

Sub FileInfoMakeNewTable(mode As Integer, tfilename As String)
' Creates a new File table for the passed filename
' mode = 1  standard database File table
' mode = 2  probe database File table
' mode = 3  setup database File table (primary standards)
' mode = 4  user database File table
' mode = 5  position database File table
' mode = 6  xray database File table
' mode = 7  setup database File table (MAN standards)
' mode = 8  setup database File table (interference standards)
' mode = 9  matrix database File table (Penepma matrix k-ratios and factors)
' mode = 10 boundary database File table (Penepma boundary k-ratios and factors)
' mode = 11 pure database File table (Penepma pure element intensities)

ierror = False
On Error GoTo FileInfoMakeNewTableError

Dim FiDb As Database
Dim FiDt As Recordset

' Specify the probe database variables, first File table
Dim File As New TableDef

Dim FileType As New Field
Dim FileVersion As New Field
Dim FileTitle As New Field
Dim FileUser As New Field
Dim FileDescription As New Field

Dim FileCustom1 As New Field
Dim FileCustom2 As New Field
Dim FileCustom3 As New Field

' Get correct database filename
Call FileInfoGetMDBFileName(mode%, tfilename$)
If ierror Then Exit Sub

' Open the specified database file
Screen.MousePointer = vbHourglass
Set FiDb = OpenDatabase(tfilename$, DatabaseExclusiveAccess%, False)

' Specify the "File" table
File.Name = "File"

FileType.Name = "Type"
FileType.Type = dbText
FileType.Size = DbTextFileInfoTypeFieldLength%
FileType.AllowZeroLength = True
File.Fields.Append FileType

FileVersion.Name = "Version"
FileVersion.Type = dbSingle
File.Fields.Append FileVersion

FileTitle.Name = "Title"
FileTitle.Type = dbText
FileTitle.Size = DbTextFileInfoTitleFieldLength%
FileTitle.AllowZeroLength = True
File.Fields.Append FileTitle

FileUser.Name = "User"
FileUser.Type = dbText
FileUser.Size = DbTextFileInfoUserFieldLength%
FileUser.AllowZeroLength = True
File.Fields.Append FileUser

FileDescription.Name = "Description"
FileDescription.Type = dbText
FileDescription.Size = DbTextFileInfoDescriptionFieldLength%
FileDescription.AllowZeroLength = True
File.Fields.Append FileDescription

FileCustom1.Name = "Custom1"
FileCustom1.Type = dbText
FileCustom1.Size = DbTextFileInfoCustomFieldLength%
FileCustom1.AllowZeroLength = True
File.Fields.Append FileCustom1

FileCustom2.Name = "Custom2"
FileCustom2.Type = dbText
FileCustom2.Size = DbTextFileInfoCustomFieldLength%
FileCustom2.AllowZeroLength = True
File.Fields.Append FileCustom2

FileCustom3.Name = "Custom3"
FileCustom3.Type = dbText
FileCustom3.Size = DbTextFileInfoCustomFieldLength%
FileCustom3.AllowZeroLength = True
File.Fields.Append FileCustom3

FiDb.TableDefs.Append File

' Save file specific info based on type
If mode% = 1 Then
MDBFileType$ = "STANDARD"
MDBFileTitle$ = "Default Standard Database"
MDBFileDescription$ = "Standard Composition (Probe for EPMA)"

ElseIf mode% = 2 Then
MDBFileType$ = "PROBE"

ElseIf mode% = 3 Then
MDBFileType$ = "SETUP"
MDBFileDescription$ = "Element Setup (Probe for EPMA)"

ElseIf mode% = 4 Then
MDBFileType$ = "USER"
MDBFileTitle$ = "Default User Database"
MDBFileDescription$ = "Hourly Usage (Probe for EPMA)"

ElseIf mode% = 5 Then
MDBFileType$ = "POSITION"
MDBFileDescription$ = "Digitized Positions (Probe for EPMA)"

ElseIf mode% = 6 Then
MDBFileType$ = "XRAY"
MDBFileTitle$ = "Default Xray Database"
MDBFileDescription$ = "NIST X-ray Lines (Probe for EPMA)"

ElseIf mode% = 7 Then       ' MAN intensity database
MDBFileType$ = "SETUP"
MDBFileDescription$ = "MAN Intensities (Probe for EPMA)"

ElseIf mode% = 8 Then       ' Interference intensity database
MDBFileType$ = "SETUP"
MDBFileDescription$ = "Interference Intensities (Probe for EPMA)"

ElseIf mode% = 9 Then       ' Matrix database
MDBFileType$ = "MATRIX"
MDBFileDescription$ = "Penepma Matrix K-Ratios and Factors (Probe for EPMA)"

ElseIf mode% = 10 Then       ' Boundary database
MDBFileType$ = "BOUNDARY"
MDBFileDescription$ = "Penepma Boundary K-Ratios and Factors (Probe for EPMA)"

ElseIf mode% = 11 Then       ' Pure database
MDBFileType$ = "PURE"
MDBFileDescription$ = "Penepma Pure Element Intensity (Probe for EPMA)"
End If

' Now add new record to "File" table
Set FiDt = FiDb.OpenRecordset("File", dbOpenTable)

' Update global
DataFileVersionNumber! = ProgramVersionNumber!

FiDt.AddNew
FiDt("Type") = Left$(MDBFileType$, DbTextFileInfoTypeFieldLength%)
FiDt("Version") = DataFileVersionNumber!
FiDt("Title") = Left$(MDBFileTitle$, DbTextFileInfoTitleFieldLength%)
FiDt("User") = Left$(MDBUserName$, DbTextFileInfoUserFieldLength%)
FiDt("Description") = Left$(MDBFileDescription$, DbTextFileInfoDescriptionFieldLength%)

FiDt("Custom1") = Left$(CustomText1$, DbTextFileInfoCustomFieldLength%)
FiDt("Custom2") = Left$(CustomText2$, DbTextFileInfoCustomFieldLength%)
FiDt("Custom3") = Left$(CustomText3$, DbTextFileInfoCustomFieldLength%)
FiDt.Update
FiDt.Close

FiDb.Close
Screen.MousePointer = vbDefault

Exit Sub

' Errors
FileInfoMakeNewTableError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "FileInfoMakeNewTable"
ierror = True
Exit Sub

End Sub

Sub FileInfoSaveData(filename As String)
' Save data to MDB "File" table

ierror = False
On Error GoTo FileInfoSaveDataError

Dim FiDb As Database
Dim FiDt As Recordset

' Open the specified database file
Screen.MousePointer = vbHourglass
Set FiDb = OpenDatabase(filename$, DatabaseExclusiveAccess%, False)

' Open file table
Set FiDt = FiDb.OpenRecordset("File", dbOpenTable)

' Save user, name, and description fields
FiDt.Edit
FiDt("User") = Left$(MDBUserName$, DbTextFileInfoUserFieldLength%)
FiDt("Title") = Left$(MDBFileTitle$, DbTextFileInfoTitleFieldLength%)
FiDt("Description") = Left$(MDBFileDescription$, DbTextFileInfoDescriptionFieldLength%)

If DataFileVersionNumber! >= 2.05 Then
FiDt("Custom1") = Left$(CustomText1$, DbTextFileInfoCustomFieldLength%)
FiDt("Custom2") = Left$(CustomText2$, DbTextFileInfoCustomFieldLength%)
FiDt("Custom3") = Left$(CustomText3$, DbTextFileInfoCustomFieldLength%)
End If

FiDt.Update

' Close the database
FiDt.Close
FiDb.Close
Screen.MousePointer = vbDefault
Exit Sub

' Errors
FileInfoSaveDataError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "FileInfoSaveData"
ierror = True
Exit Sub

End Sub

Sub FileInfoGetMDBFileName(mode As Integer, tfilename As String)
' Loads the correct filename based on mode
' mode = 1  standard database File table
' mode = 2  probe database File table
' mode = 3  setup database File table (primary standards)
' mode = 4  user database File table
' mode = 5  position database File table
' mode = 6  xray database File table
' mode = 7  setup database File table (MAN standards)
' mode = 8  setup database File table (interference standards)
' mode = 9  matrix database File table (Penepma matrix k-ratios and factors)
' mode = 10 boundary database File table (Penepma boundary k-ratios and factors)

ierror = False
On Error GoTo FileInfoGetMDBFileNameError

' Load default filename if blank
If mode% = 1 And tfilename$ = vbNullString Then
tfilename$ = StandardDataFile$                  ' might be using non-standard data file

ElseIf mode% = 2 And tfilename$ = vbNullString Then
GoTo FileInfoGetMDBFileNameBlankProbeFile

ElseIf mode% = 3 And tfilename$ = vbNullString Then
tfilename$ = ApplicationCommonAppData$ & "SETUP.MDB"

ElseIf mode% = 4 And tfilename$ = vbNullString Then
tfilename$ = ApplicationCommonAppData$ & "USER.MDB"

ElseIf mode% = 5 And tfilename$ = vbNullString Then
tfilename$ = ApplicationCommonAppData$ & "POSITION.MDB"

ElseIf mode% = 6 And tfilename$ = vbNullString Then
tfilename$ = ApplicationCommonAppData$ & "XRAY.MDB"

ElseIf mode% = 7 And tfilename$ = vbNullString Then
tfilename$ = ApplicationCommonAppData$ & "SETUP2.MDB"

ElseIf mode% = 8 And tfilename$ = vbNullString Then
tfilename$ = ApplicationCommonAppData$ & "SETUP3.MDB"

ElseIf mode% = 9 And tfilename$ = vbNullString Then
tfilename$ = ApplicationCommonAppData$ & "MATRIX.MDB"

ElseIf mode% = 10 And tfilename$ = vbNullString Then
tfilename$ = ApplicationCommonAppData$ & "BOUNDARY.MDB"

ElseIf mode% = 11 And tfilename$ = vbNullString Then
tfilename$ = ApplicationCommonAppData$ & "PURE.MDB"
End If

Exit Sub

' Errors
FileInfoGetMDBFileNameError:
MsgBox Error$, vbOKOnly + vbCritical, "FileInfoGetMDBFileName"
ierror = True
Exit Sub

FileInfoGetMDBFileNameBlankProbeFile:
msg$ = "Probe database file name is blank"
MsgBox msg$, vbOKOnly + vbExclamation, "FileInfoGetMDBFileName"
ierror = True
Exit Sub

End Sub

Sub FileInfoCreateDatabase(tfilename As String)
' Creates a new MDB database file from an existing MDB template

ierror = False
On Error GoTo FileInfoCreateDatabaseError

' Check that passed filename is not blank
If Trim$(tfilename$) = vbNullString Then GoTo FileInfoCreateDatabaseBlankFilename

' Check if file already exists, if so, delete it first
If Dir$(tfilename$) <> vbNullString Then
Kill tfilename$
DoEvents
End If

FileCopy MDB_Template$, tfilename$

Exit Sub

' Errors
FileInfoCreateDatabaseError:
MsgBox Error$, vbOKOnly + vbCritical, "FileInfoCreateDatabase"
ierror = True
Exit Sub

FileInfoCreateDatabaseBlankFilename:
msg$ = "Passed database file name is blank"
MsgBox msg$, vbOKOnly + vbExclamation, "FileInfoCreateDatabase"
ierror = True
Exit Sub

End Sub

