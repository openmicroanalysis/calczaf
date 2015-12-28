Attribute VB_Name = "CodeUSER"
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

Type TypeUser
versionnumber As Single
DataFile As String
UserName As String
FileTitle As String
FileDescription As String
CustomText1 As String
CustomText2 As String
CustomText3 As String
StartDateTime As Variant
StopDateTime As Variant
FileCreated As String
FileModified As String
TotalAcquisitionTime As Variant
End Type

Sub UserGetLastRecord(tusername As String, tForm As Form)
' Returns the last match to the user name field

ierror = False
On Error GoTo UserGetLastRecordError

Dim Version As Single
Dim SQLQ As String

Dim UsDb As Database
Dim UsDt As Recordset
Dim UsDs As Recordset

' Only get last record if passed string is at least characters in length
If Len(tusername$) < 3 Then Exit Sub

' Check for user database
If Trim$(UserDataFile$) = vbNullString Then GoTo UserGetLastRecordBlankName
If Dir$(UserDataFile$) = vbNullString Then
Call UserOpenNewFile(UserDataFile$)
If ierror Then Exit Sub
End If

' Check USER.MDB for for new fields
Call UserUpdateMDBFile
If ierror Then Exit Sub

' Open user database
Screen.MousePointer = vbHourglass
Set UsDb = OpenDatabase(UserDataFile$, UserDatabaseNonExclusiveAccess%, dbReadOnly)

' First check file version number
Set UsDt = UsDb.OpenRecordset("File", dbOpenTable)
Version! = UsDt("Version")
UsDt.Close

' Replace single quotes (for Irish surnames)
tusername$ = Replace$(tusername$, "'", "?")

' Define search based on user name
SQLQ$ = "SELECT User.* FROM User WHERE User.Names Like '" & Trim$(tusername$) & "*' "
SQLQ$ = SQLQ$ & " ORDER BY Rows"

' Create a dynaset
Set UsDs = UsDb.OpenRecordset(SQLQ$, dbOpenDynaset, dbReadOnly)

' If records found, load last record
If Not UsDs.BOF And Not UsDs.EOF Then
UsDs.MoveLast

tForm.TextTitle.Text = Trim$(vbNullString & UsDs("Titles"))
tForm.TextDescription.Text = Trim$(vbNullString & UsDs("Descriptions"))

If Version! >= 2.05 Then
tForm.TextCustom1.Text = Trim$(vbNullString & UsDs("Custom1"))
tForm.TextCustom2.Text = Trim$(vbNullString & UsDs("Custom2"))
tForm.TextCustom3.Text = Trim$(vbNullString & UsDs("Custom3"))
End If
End If

UsDs.Close
UsDb.Close

Screen.MousePointer = vbDefault
Exit Sub

' Errors
UserGetLastRecordError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "UserGetLastRecord"
ierror = True
Exit Sub

UserGetLastRecordBlankName:
msg$ = "User data file name is blank"
MsgBox msg$, vbOKOnly + vbExclamation, "UserGetLastRecord"
ierror = True
Exit Sub

End Sub

Sub UserOpenNewFile(tfilename As String)
' The routine opens a new user database file (USER.MDB)

ierror = False
On Error GoTo UserOpenNewFileError

Dim UsDb As Database

' Specify the user database variables
Dim UserTable As New TableDef

Dim UserRows As New Field
Dim UserNames As New Field
Dim UserFiles As New Field
Dim UserCreatedDates As New Field
Dim UserModifiedDates As New Field
Dim UserVersions As New Field
Dim UserTitles As New Field
Dim UserDescriptions As New Field
Dim UserStartDateTimes As New Field
Dim UserStopdateTimes As New Field

Dim UserCustom1 As New Field
Dim UserCustom2 As New Field
Dim UserCustom3 As New Field

Dim UserTotalAcquisitionTime As New Field

Dim NameIndex As New Index

' Note that because of DAO 3.5 compatibility issues, we cannot create new SETUP*.MDB files at this time (DB data control compatibility)
msg$ = "Unable to create a new user database: " & UserDataFile$ & " due to DAO 3.50 compatibility issues." & vbCrLf & vbCrLf
msg$ = msg$ & "To obtain a new user database file please download the default USER.MDB files from our ftp servers or contact Probe Software."
MsgBox msg$, vbOKOnly + vbInformation, "UserOpenNewFile"
Exit Sub

' Check for blank file
If Trim$(tfilename$) = vbNullString Then GoTo UserOpenNewFileBlankName

' If file exists, erase it
If Dir$(tfilename$) <> vbNullString Then
Kill tfilename$

' Else inform user
Else
msg$ = "Creating a new User database: " & tfilename$
MsgBox msg$, vbOKOnly + vbInformation, "UserOpenNewFile"
End If

' Open the new database and create the tables and index
Screen.MousePointer = vbHourglass
'Set UsDb = CreateDatabase(tfilename$, dbLangGeneral)
'If UsDb Is Nothing Or Err <> 0 Then GoTo UserOpenNewFileError

' Open a new database by copying from existing MDB template
Call FileInfoCreateDatabase(tfilename$)
If ierror Then Exit Sub

' Open as existing database
Set UsDb = OpenDatabase(tfilename$, DatabaseExclusiveAccess%, False)

' Specify the user database "User" table
UserTable.Name = "User"

UserRows.Name = "Rows"
UserRows.Type = dbLong
UserTable.Fields.Append UserRows

UserNames.Name = "Names"
UserNames.Type = dbText
UserNames.Size = DbTextUserNameLength%
UserNames.AllowZeroLength = True
UserTable.Fields.Append UserNames

UserFiles.Name = "Files"
UserFiles.Type = dbText
UserFiles.Size = DbTextFilenameLength%
UserFiles.AllowZeroLength = True
UserTable.Fields.Append UserFiles

UserCreatedDates.Name = "CreatedDates"
UserCreatedDates.Type = dbDate
UserTable.Fields.Append UserCreatedDates

UserModifiedDates.Name = "ModifiedDates"
UserModifiedDates.Type = dbDate
UserTable.Fields.Append UserModifiedDates

UserVersions.Name = "Versions"
UserVersions.Type = dbSingle
UserTable.Fields.Append UserVersions

UserTitles.Name = "Titles"
UserTitles.Type = dbText
UserTitles.Size = DbTextDescriptionLength%
UserTitles.AllowZeroLength = True
UserTable.Fields.Append UserTitles

UserDescriptions.Name = "Descriptions"
UserDescriptions.Type = dbText
UserDescriptions.Size = DbTextDescriptionLength%
UserDescriptions.AllowZeroLength = True
UserTable.Fields.Append UserDescriptions

UserStartDateTimes.Name = "StartDateTimes"
UserStartDateTimes.Type = dbDate
UserTable.Fields.Append UserStartDateTimes

UserStopdateTimes.Name = "StopDateTimes"
UserStopdateTimes.Type = dbDate
UserTable.Fields.Append UserStopdateTimes

' Add custom fields
UserCustom1.Name = "Custom1"
UserCustom1.Type = dbText
UserCustom1.Size = DbTextNameLength%
UserCustom1.AllowZeroLength = True
UserTable.Fields.Append UserCustom1

UserCustom2.Name = "Custom2"
UserCustom2.Type = dbText
UserCustom2.Size = DbTextNameLength%
UserCustom2.AllowZeroLength = True
UserTable.Fields.Append UserCustom2

UserCustom3.Name = "Custom3"
UserCustom3.Type = dbText
UserCustom3.Size = DbTextNameLength%
UserCustom3.AllowZeroLength = True
UserTable.Fields.Append UserCustom3

' New to version 7.45 (see UserUpdateMDBFile)
UserTotalAcquisitionTime.Name = "TotalAcquisitionTime"
UserTotalAcquisitionTime.Type = dbDouble
UserTable.Fields.Append UserTotalAcquisitionTime

' Specify the user database "NameIndex" index
NameIndex.Name = "User Names"
NameIndex.Fields = "Names"
UserTable.Indexes.Append NameIndex

UsDb.TableDefs.Append UserTable

' Close the user database
UsDb.Close
Screen.MousePointer = vbDefault

' Create new File table for user database
Call FileInfoMakeNewTable(Int(4), tfilename$)
If ierror Then Exit Sub

' No errors, load file name
UserDataFile$ = tfilename$

Exit Sub

' Errors
UserOpenNewFileError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "UserOpenNewFile"
ierror = True
Exit Sub

UserOpenNewFileBlankName:
msg$ = "The user database file name is blank"
MsgBox msg$, vbOKOnly + vbExclamation, "UserOpenNewFile"
ierror = True
Exit Sub

End Sub

Sub UserOpenOldFile(Filename As String)
' Open existing user file

ierror = False
On Error GoTo UserFormOpenOldFileError

' Get data fields
Call FileInfoLoadData(Int(4), Filename$)
If ierror Then Exit Sub

' No errors, load filename
UserDataFile$ = Filename$
MDBUserName$ = app.EXEName      ' for InitWindow

Exit Sub

' Errors
UserFormOpenOldFileError:
MsgBox Error$, vbOKOnly + vbCritical, "UserFormOpenOldFile"
ierror = True
Exit Sub

End Sub

Sub UserSaveRecord(user As TypeUser)
' Save record to user database

ierror = False
On Error GoTo UserSaveRecordError

Dim rownumber As Long
Dim fileversionnumber As Single

Dim UsDb As Database
Dim UsDt As Recordset

' Check for user database
If Trim$(UserDataFile$) = vbNullString Then GoTo UserSaveRecordBlankName
If Dir$(UserDataFile$) = vbNullString Then
Call UserOpenNewFile(UserDataFile$)
If ierror Then Exit Sub
End If

' Check USER.MDB for for new fields
Call UserUpdateMDBFile
If ierror Then Exit Sub

' Open user database
Screen.MousePointer = vbHourglass
Set UsDb = OpenDatabase(UserDataFile$, UserDatabaseExclusiveAccess%, False)

' First check file version number
Set UsDt = UsDb.OpenRecordset("File", dbOpenTable)
fileversionnumber! = UsDt("Version")
UsDt.Close

' Open user table
Set UsDt = UsDb.OpenRecordset("User", dbOpenTable)

' Go to end of table to get last row number
If UsDt.BOF And UsDt.EOF Then
rownumber& = 1
Else
UsDt.MoveLast
rownumber& = UsDt("Rows") + 1
End If

' Save to user database
UsDt.AddNew
UsDt("Rows") = rownumber&
UsDt("Versions") = user.versionnumber!
UsDt("Files") = Left$(user.DataFile$, DbTextFilenameLength%)
UsDt("Names") = Left$(user.UserName$, DbTextUserNameLength%)
UsDt("Titles") = Left$(user.FileTitle$, DbTextTitleStringLength%)
UsDt("Descriptions") = Left$(user.FileDescription$, DbTextDescriptionLength%)
UsDt("StartDateTimes") = user.StartDateTime
UsDt("StopDateTimes") = user.StopDateTime
UsDt("CreatedDates") = user.FileCreated$
UsDt("ModifiedDates") = user.FileModified$

If fileversionnumber! >= 2.05 Then
UsDt("Custom1") = Left$(user.CustomText1$, DbTextNameLength%)
UsDt("Custom2") = Left$(user.CustomText2$, DbTextNameLength%)
UsDt("Custom3") = Left$(user.CustomText3$, DbTextNameLength%)
End If

If fileversionnumber! >= 7.45 Then
UsDt("TotalAcquisitionTime") = CDbl(user.TotalAcquisitionTime)
End If

UsDt.Update

UsDt.Close
UsDb.Close

Screen.MousePointer = vbDefault
Exit Sub

' Errors
UserSaveRecordError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "UserSaveRecord"
ierror = True
Exit Sub

UserSaveRecordBlankName:
msg$ = "User data file name is blank"
MsgBox msg$, vbOKOnly + vbExclamation, "UserSaveRecord"
ierror = True
Exit Sub

End Sub

Sub UserUpdateMDBFile()
' Routine to update the USER.MDB file automatically for new fields

ierror = False
On Error GoTo UserUpdateMDBFileError

Dim UsDb As Database
Dim UsRs As Recordset

Dim versionnumber As Single
Dim updated As Integer

Dim UserTotalAcquisitionTime As New Field

' First check that a user database exists
If Trim$(UserDataFile$) = vbNullString Then UserDataFile$ = ApplicationCommonAppData$ & "USER.MDB"
If Dir$(UserDataFile$) = vbNullString Then Exit Sub

' Get version number
versionnumber! = FileInfoGetVersion!(UserDataFile$, "USER")
If ierror Then Exit Sub

' Open the database
Screen.MousePointer = vbHourglass
Set UsDb = OpenDatabase(UserDataFile$, UserDatabaseExclusiveAccess%, False)

Call TransactionBegin("UserUpdateMDBFile", UserDataFile$)
If ierror Then Exit Sub

' Flag file as not updated
updated = False

' Add total acquisition time field and records
If versionnumber! < 7.45 Then

' Add total acquisition time field to User table
UserTotalAcquisitionTime.Name = "TotalAcquisitionTime"
UserTotalAcquisitionTime.Type = dbDouble
UsDb.TableDefs("User").Fields.Append UserTotalAcquisitionTime

' Open user table in user database
Set UsRs = UsDb.OpenRecordset("User", dbOpenTable)

' Add default total acquisition time fields to all records
Do Until UsRs.EOF
UsRs.Edit
UsRs("TotalAcquisitionTime") = 0#
UsRs.Update
UsRs.MoveNext
Loop

UsRs.Close
updated = True
End If

' Add new fields and records based on "versionnumber"




' Open "File" table and update data file version number
If updated Then
Set UsRs = UsDb.OpenRecordset("File", dbOpenTable)
UsRs.Edit
UsRs("Version") = ProgramVersionNumber!
UsRs.Update
UsRs.Close
End If

UsDb.Close

Call TransactionCommit("UserUpdateMDBFile", UserDataFile$)
If ierror Then Exit Sub

Screen.MousePointer = vbDefault
Exit Sub

' Errors
UserUpdateMDBFileError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "UserUpdateMDBFile"
Call TransactionRollback("UserUpdateMDBFile", UserDataFile$)
ierror = True
Exit Sub

End Sub

