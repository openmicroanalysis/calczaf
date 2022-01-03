Attribute VB_Name = "CodeFILEINFO"
' (c) Copyright 1995-2022 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub FileInfoLoad(mode As Integer, tfilename As String)
' Loads the FormFileInfo form with the .MDB "File" table info
' mode = 1  load standard database File table
' mode = 2  load probe database File table
' mode = 3  load setup database File table (primary standards)
' mode = 4  load user database File table
' mode = 5  load position database File table
' mode = 6  load xray database File table
' mode = 7  load setup database File table (MAN standards)
' mode = 8  setup database File table (interference standards)
' mode = 9  matrix database File table (Penepma matrix k-ratios and factors)
' mode = 10 boundary database File table (Penepma boundary k-ratios and factors)
' mode = 11 pure database File table (Penepma pure element intensities)

ierror = False
On Error GoTo FileInfoLoadError

' Get data fields
Call FileInfoLoadData(mode%, tfilename$)
If ierror Then Exit Sub

' Load File table info
FormFILEINFO.LabelFileName.Caption = tfilename$
FormFILEINFO.LabelVersion.Caption = Format$(DataFileVersionNumber!)
FormFILEINFO.LabelType.Caption = MDBFileType$

FormFILEINFO.TextUser.Text = MDBUserName$
FormFILEINFO.TextTitle.Text = MDBFileTitle$
FormFILEINFO.TextDescription.Text = MDBFileDescription$

FormFILEINFO.LabelDateCreated.Caption = MDBFileCreated$
FormFILEINFO.LabelDateModified.Caption = MDBFileModified$
FormFILEINFO.LabelLastUpdated.Caption = MDBFileUpdated$

' Load default custom labels
FormFILEINFO.LabelCustom1.Caption = CustomLabel1$
FormFILEINFO.LabelCustom2.Caption = CustomLabel2$
FormFILEINFO.LabelCustom3.Caption = CustomLabel3$

' Load custom text from globals
FormFILEINFO.TextCustom1.Text = CustomText1$
FormFILEINFO.TextCustom2.Text = CustomText2$
FormFILEINFO.TextCustom3.Text = CustomText3$

Exit Sub

' Errors
FileInfoLoadError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "FileInfoLoad"
ierror = True
Exit Sub

End Sub

Sub FileInfoSave()
' Saves any user changes to the .MDB "File" table "Name" and "Description" fields

ierror = False
On Error GoTo FileInfoSaveError

Dim tfilename As String

' Check for blank filename
If FormFILEINFO.LabelFileName.Caption = vbNullString Then GoTo FileInfoSaveBadFileName
tfilename$ = FormFILEINFO.LabelFileName.Caption

' Get user, title and and description file information
MDBUserName$ = FormFILEINFO.TextUser.Text
MDBFileTitle$ = FormFILEINFO.TextTitle.Text
MDBFileDescription$ = FormFILEINFO.TextDescription.Text

' Save custom fields to globals
CustomText1$ = FormFILEINFO.TextCustom1.Text
CustomText2$ = FormFILEINFO.TextCustom2.Text
CustomText3$ = FormFILEINFO.TextCustom3.Text

' Save data to file table
Call FileInfoSaveData(tfilename$)
If ierror Then Exit Sub

Exit Sub

' Errors
FileInfoSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "FileInfoSave"
ierror = True
Exit Sub

FileInfoSaveBadFileName:
msg$ = "The database file name is blank"
MsgBox msg$, vbOKOnly + vbExclamation, "FileInfoSave"
ierror = True
Exit Sub

End Sub
