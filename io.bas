Attribute VB_Name = "CodeIO"
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

Dim LogString As String

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenFilename As Type_CDL_OpenFileName) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenFilename As Type_CDL_OpenFileName) As Long

Private Type Type_CDL_OpenFileName
    lStructSize As Long          ' the size of this struct (use the Len function)
    hwndOwner As Long            ' the hWnd of the owner window. The dialog will be modal to this window
    hInstance As Long            ' the instance of the calling thread. You can use the App.hInstance here.
    lpstrFilter As String        ' use this to filter what files are showen in the dialog. Separate each filter with ChrW$(0). The string also has to end with a Chr(0).
    lpstrCustomFilter As String  ' pattern the user has choosed is saved here if you pass a non empty string. I never use this one
    nMaxCustFilter As Long       ' maximum saved custom filters. Since I never use the lpstrCustomFilter I always pass 0 to this.
    nFilterIndex As Long         ' what filter (of lpstrFilter) is showed when the user opens the dialog.
    lpstrFile As String          ' path and name of the file. This must be at least MAX_PATH character long.
    nMaxFile As Long             ' length of lpstrFile
    lpstrFileTitle As String     ' name only of the file returned. Should be MAX_PATH character long
    nMaxFileTitle As Long        ' length of lpstrFileTitle
    lpstrInitialDir As String    ' path to the initial folder. If you pass an empty string the initial path is the current path.
    lpstrTitle As String         ' caption of the dialog
    flags As Long                ' flags. See the values in MSDN Library (you can look at the flags property of the common dialog control)
    nFileOffset As Integer       ' points to the what character in lpstrFile where the actual filename begins (zero based)
    nFileExtension As Integer    ' same as nFileOffset except that it points to the file extention.
    lpstrDefExt As String        ' can contain the extention Windows should add to a file if the user doesn't provide one (used with the GetSaveFileName API function)
    lCustData As Long            ' only used if you provide a Hook procedure (Making a Hook procedure is pretty messy in VB.
    lpfnHook As Long             ' pointer to the hook procedure.
    lpTemplateName As String     ' a string that contains a dialog template resource name. Only used with the hook procedure.
End Type

'Private Const OFN_READONLY = &H1
'Private Const OFN_OVERWRITEPROMPT = &H2
'Private Const OFN_HIDEREADONLY = &H4
'Private Const OFN_NOCHANGEDIR = &H8
'Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
'Private Const OFN_ENABLETEMPLATE = &H40
'Private Const OFN_ENABLETEMPLATEHANDLE = &H80
'Private Const OFN_NOVALIDATE = &H100
'Private Const OFN_ALLOWMULTISELECT = &H200
'Private Const OFN_EXTENSIONDIFFERENT = &H400
'Private Const OFN_PATHMUSTEXIST = &H800
'Private Const OFN_FILEMUSTEXIST = &H1000
'Private Const OFN_CREATEPROMPT = &H2000
'Private Const OFN_SHAREAWARE = &H4000
'Private Const OFN_NOREADONLYRETURN = &H8000
'Private Const OFN_NOTESTFILECREATE = &H10000
'Private Const OFN_NONETWORKBUTTON = &H20000
'Private Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
Private Const OFN_EXPLORER = &H80000                         '  new look commdlg
'Private Const OFN_NODEREFERENCELINKS = &H100000
'Private Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
Private Const OFN_ENABLESIZING = &H800000
'Private Const OFN_SHAREFALLTHROUGH = 2
'Private Const OFN_SHARENOWARN = 1
'Private Const OFN_SHAREWARN = 0

Sub IOCloseOUTFile()
' This routine closes the .OUT file

ierror = False
On Error GoTo IOCloseOUTFileError

' Close file and update flag
Close #OutputDataFileNumber%
SaveToDisk = False

' Update menu item
FormMAIN.menuOutputSaveToDiskLog.Checked = False

Exit Sub

' Errors
IOCloseOUTFileError:
MsgBox Error$, vbOKOnly + vbCritical, "IOCloseOUTFile"
Close #OutputDataFileNumber%
ierror = True
Exit Sub

End Sub

Sub IOGetFileName(mode As Integer, ioextension As String, iofilename As String, tForm As Form)
' This routine returns a filename based on passed extension
'  mode = 0 save new file (do not check for existing file to allow append mode)
'  mode = 1 save new file
'  mode = 2 open old file

ierror = False
On Error GoTo IOGetFileNameError

' Set filters and dialog titles for .DAT files
If UCase$(ioextension$) = "DAT" Then
FormMAIN.CMDialog1.Filter = "ASCII Data Files (*.DAT)|*.DAT|All Files (*.*)|*.*|"
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.dat"

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Save ASCII Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Read ASCII Data From"
End If

' Position files
ElseIf UCase$(ioextension$) = "POS" Then
FormMAIN.CMDialog1.Filter = "ASCII Position Files (*.POS)|*.POS|All Files (*.*)|*.*|"
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.pos"

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Export Position Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Import Position Data From"
End If

' (Canadian Geological Survey, Custom Format #1) Position files
ElseIf UCase$(ioextension$) = "LEP" Then
FormMAIN.CMDialog1.Filter = "LEP Position Files (*.LEP)|*.LEP|All Files (*.*)|*.*|"
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.lep"

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Export LEP Position Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Import LEP Position Data From"
End If

' Text files
ElseIf UCase$(ioextension$) = "TXT" Then
FormMAIN.CMDialog1.Filter = "Probe Text Files (*.TXT)|*.TXT|All Files (*.*)|*.*|"
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.txt"

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Export Probe Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Import Probe Data From"
End If

' Output files
ElseIf UCase$(ioextension$) = "OUT" Then
FormMAIN.CMDialog1.Filter = "Probe Output Files (*.OUT)|*.OUT|All Files (*.*)|*.*|"
If Trim$(iofilename$) = vbNullString Then
If ProbeDataFile$ <> vbNullString Then
iofilename$ = MiscGetFileNameNoExtension$(ProbeDataFile$) & ".out"
Else
iofilename$ = UserDataDirectory$ & "\" & app.EXEName$ & ".out"
End If
End If
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.out"

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output Probe Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input Probe Data From"
End If

' Excel files
ElseIf UCase$(ioextension$) = "XLS" Then    ' Excel 2003 (version 11) (Office 2003) and earlier
FormMAIN.CMDialog1.Filter = "Excel Files (*.XLS)|*.XLS|All Files (*.*)|*.*|"
If Trim$(iofilename$) = vbNullString And ProbeDataFile$ <> vbNullString Then
iofilename$ = MiscGetFileNameNoExtension$(ProbeDataFile$) & ".xls"
End If
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.xls"

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output Excel Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input Excel Data From"
End If

ElseIf UCase$(ioextension$) = "XLSX" Then   ' Excel 2007 (version 12) (Office 2007) and later
FormMAIN.CMDialog1.Filter = "Excel Files (*.XLSX)|*.XLSX|All Files (*.*)|*.*|"
If Trim$(iofilename$) = vbNullString And ProbeDataFile$ <> vbNullString Then
iofilename$ = MiscGetFileNameNoExtension$(ProbeDataFile$) & ".xlsx"
End If
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.xlsx"

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output Excel Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input Excel Data From"
End If

' Grid files
ElseIf UCase$(ioextension$) = "GRD" Then
FormMAIN.CMDialog1.Filter = "Grid Files (*.GRD)|*.GRD|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.grd"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output Grid Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input Grid Data From"
End If

' Edax spectrum files
ElseIf UCase$(ioextension$) = "SPC" Then
FormMAIN.CMDialog1.Filter = "Edax Spectrum Files (*.SPC)|*.SPC|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.spc"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output Spectrum Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input Spectrum Data From"
End If

' Penepma Input files
ElseIf UCase$(ioextension$) = "IN" Then
FormMAIN.CMDialog1.Filter = "Penepma Input Files (*.IN)|*.IN|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.in"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output Penepma Input Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input Penepma Input Data From"
End If

' Thermo spectrum files
ElseIf UCase$(ioextension$) = "EMSA" Then
'FormMAIN.CMDialog1.Filter = "EMSA (Thermo) Spectrum Files (*.EMSA)|*.EMSA|All Files (*.*)|*.*|"
FormMAIN.CMDialog1.Filter = "EMSA/MSA) Spectrum Files (*.EMSA;*.MSA)|*.EMSA;*.MSA|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.emsa"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output Spectrum Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input Spectrum Data From"
End If

' Bruker spectrum files
ElseIf UCase$(ioextension$) = "SPX" Then
FormMAIN.CMDialog1.Filter = "Bruker Spectrum Files (*.SPX)|*.SPX|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.spx"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output Spectrum Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input Spectrum Data From"
End If

' Alpha-factor and condition files
ElseIf UCase$(ioextension$) = "AFA" Then
FormMAIN.CMDialog1.Filter = "Probe for EPMA Condition Files (*.AFA)|*.AFA|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.afa"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Save Condition Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input Condition Data From"
End If

' CalcImage project files
ElseIf UCase$(ioextension$) = "CIP" Then
FormMAIN.CMDialog1.Filter = "CalcImage Project Files (*.CIP)|*.CIP|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.cip"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output Project Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input Project Data From"
End If

' BMP image files
ElseIf UCase$(ioextension$) = "BMP" Then
FormMAIN.CMDialog1.Filter = "Bitmap Image Files (*.BMP)|*.BMP|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.bmp"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output Bitmap Image To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input Bitmap Image From"
End If

' JPG image files
ElseIf UCase$(ioextension$) = "JPG" Then
FormMAIN.CMDialog1.Filter = "JPEG Image Files (*.JPG)|*.JPG|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.jpg"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output JPEG Image To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input JPEG Image From"
End If

' GIF image files
ElseIf UCase$(ioextension$) = "GIF" Then
FormMAIN.CMDialog1.Filter = "GIF Image Files (*.GIF)|*.GIF|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.gif"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output GIF Image To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input GIF Image From"
End If

' TIF image files
ElseIf UCase$(ioextension$) = "TIF" Then
FormMAIN.CMDialog1.Filter = "Tiff Image Files (*.TIF)|*.TIF|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.tif"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output Tiff Image To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input Tiff Image From"
End If

' PCC Probe column condition files
ElseIf UCase$(ioextension$) = "PCC" Then
FormMAIN.CMDialog1.Filter = "Probe Column Condition Files (*.PCC)|*.PCC|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.pcc"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output Probe Column Condition To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input Probe Column Condition From"
End If

' MAT Penepma material files (see Standard.exe Analytical menu)
ElseIf UCase$(ioextension$) = "MAT" Then
FormMAIN.CMDialog1.Filter = "PENEPMA Material Files (*.MAT)|*.MAT|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "material.mat"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output PENEPMA Material Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input PENEPMA Material Data From"
End If

' GEO Penepma geometry files (see Standard.exe Analytical menu)
ElseIf UCase$(ioextension$) = "GEO" Then
FormMAIN.CMDialog1.Filter = "PENEPMA Geometry Files (*.GEO)|*.GEO|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "bulk.geo"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output PENEPMA Geometry Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input PENEPMA Geometry Data From"
End If

ElseIf UCase$(ioextension$) = "RAW" Then
FormMAIN.CMDialog1.Filter = "Lispix RAW Datacube Files (*.RAW)|*.RAW|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "spectrum.raw"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output Lispix RAW Datacube Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input Lispix RAW Datacube Data From"
End If

' PAR Penepma parameter files (see Standard.exe Analytical menu)
ElseIf UCase$(ioextension$) = "PAR" Then
FormMAIN.CMDialog1.Filter = "PENEPMA Parameter Files (*.PAR)|*.PAR|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "material.par"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output PENEPMA Parameter Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input PENEPMA Parameter Data From"
End If

' PrbImg ProbeImage image files
ElseIf UCase$(ioextension$) = UCase$("PrbImg") Then
FormMAIN.CMDialog1.Filter = "ProbeImage Image Files (*.PrbImg)|*.PrbImg|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.PrbImg"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output ProbeImage Image To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input ProbeImage Image From"
End If

' Thermo SI spectrum image files (CalcImage)
ElseIf UCase$(ioextension$) = "SI" Then
FormMAIN.CMDialog1.Filter = "Thermo Spectrum Image Files (*.SI)|*.SI|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.si"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output Spectrum Image Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input Spectrum Image Data From"
End If

' All valid image files, Graphic Files (*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg|All Files (*.*)|*.*
ElseIf UCase$(ioextension$) = "IMG" Then
FormMAIN.CMDialog1.Filter = "Graphic Files (*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg|All Files (*.*)|*.*"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.bmp"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output Image Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input Image Data From"
End If

' JEOL .LUT palette files
ElseIf UCase$(ioextension$) = "LUT" Then
FormMAIN.CMDialog1.Filter = "JEOL Palette Files (*.LUT)|*.LUT|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.lut"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output JEOL Palette Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input JEOL Palette Data From"
End If

' CalcImage .FC palette files
ElseIf UCase$(ioextension$) = "FC" Then
FormMAIN.CMDialog1.Filter = "CalcImage Palette Files (*.FC)|*.FC|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.fc"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output CalcImage Palette Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input CalcImage Palette Data From"
End If

' Microbeam Services, Custom Format #2 Position files
ElseIf UCase$(ioextension$) = "DCD" Then
FormMAIN.CMDialog1.Filter = "DCD Position Files (*.DCD)|*.DCD|All Files (*.*)|*.*|"
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.dcd"

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Export DCD Position Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Import DCD Position Data From"
End If

' Golden Software .CLF color palette files
ElseIf UCase$(ioextension$) = "CLR" Then
FormMAIN.CMDialog1.Filter = "CLR Color Files (*.CLR)|*.CLR|All Files (*.*)|*.*|"
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.clr"

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Export CLR Color Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Import CLR Color Data From"
End If

' PrbAcq ProbeImage acquisition files
ElseIf UCase$(ioextension$) = UCase$("PrbAcq") Then
FormMAIN.CMDialog1.Filter = "ProbeImage Acquisition Files (*.PrbAcq)|*.PrbAcq|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.PrbAcq"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output ProbeImage Acquisition To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input ProbeImage Acquisition From"
End If

' CSV comma delimited files
ElseIf UCase$(ioextension$) = UCase$("CSV") Then
FormMAIN.CMDialog1.Filter = "CSV Comma Delimited Data Files (*.CSV)|*.CSV|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.csv"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output Comma Delimited Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input Comma Delimited Data From"
End If

' CND JEOL Mapping files
ElseIf UCase$(ioextension$) = UCase$("CND") Then
FormMAIN.CMDialog1.Filter = "CND JEOL Mapping Data (*.CND)|*.CND|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.cnd"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output JEOL Mapping Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input JEOL Mapping Data From"
End If

' ImpDAT Cameca Mapping files
ElseIf UCase$(ioextension$) = UCase$("ImpDAT") Then
FormMAIN.CMDialog1.Filter = "ImpDAT Cameca Mapping Data (*.ImpDAT)|*.ImpDAT|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.ImpDAT"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output Cameca Mapping Data To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input Cameca Mapping Data From"
End If

' ACQ calibration files
ElseIf UCase$(ioextension$) = UCase$("ACQ") Then
FormMAIN.CMDialog1.Filter = "ACQ Calibration Files (*.ACQ)|*.ACQ|All Files (*.*)|*.*|"
If mode% < 2 Then
If Trim$(iofilename$) = vbNullString Then iofilename$ = "untitled.acq"
End If

If mode% < 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open File To Output ACQ Calibration To"
Else
FormMAIN.CMDialog1.DialogTitle = "Open File To Input ACQ Calibration From"
End If

End If

' Specify default filter
FormMAIN.CMDialog1.FilterIndex = 1

' Specify OFN Flags
If mode% = 0 Then
FormMAIN.CMDialog1.flags = cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn Or cdlOFNPathMustExist
ElseIf mode% = 1 Then
FormMAIN.CMDialog1.flags = cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt
ElseIf mode% = 2 Then
FormMAIN.CMDialog1.flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist Or cdlOFNPathMustExist
End If

' To fix Win 7 bug for long file names (need to use system call to GetOpenFileName or GetSaveFileName)
FormMAIN.CMDialog1.flags = FormMAIN.CMDialog1.flags Or OFN_EXPLORER Or OFN_ENABLEHOOK Or OFN_ENABLESIZING

' Specify initial directory
FormMAIN.CMDialog1.InitDir = UserDataDirectory$
If UCase$(ioextension$) = UCase$("DAT") And MiscStringsAreSame(app.EXEName, "CalcZAF") Then FormMAIN.CMDialog1.InitDir = CalcZAFDATFileDirectory$
If UCase$(ioextension$) = UCase$("POS") Then FormMAIN.CMDialog1.InitDir = StandardPOSFileDirectory$
If UCase$(ioextension$) = UCase$("LEP") Then FormMAIN.CMDialog1.InitDir = StandardPOSFileDirectory$
If UCase$(ioextension$) = UCase$("DCD") Then FormMAIN.CMDialog1.InitDir = StandardPOSFileDirectory$
If UCase$(ioextension$) = UCase$("PCC") Then FormMAIN.CMDialog1.InitDir = ColumnPCCFileDirectory$
If UCase$(ioextension$) = UCase$("PrbImg") Then FormMAIN.CMDialog1.InitDir = UserImagesDirectory$
If UCase$(ioextension$) = UCase$("MAT") Then FormMAIN.CMDialog1.InitDir = PENEPMA_Path$
If UCase$(ioextension$) = UCase$("GEO") Then FormMAIN.CMDialog1.InitDir = PENEPMA_Root$
If UCase$(ioextension$) = UCase$("IN") Then FormMAIN.CMDialog1.InitDir = PENEPMA_Path$
If UCase$(ioextension$) = UCase$("SI") Then FormMAIN.CMDialog1.InitDir = UserEDSDirectory$
If UCase$(ioextension$) = UCase$("LUT") Then FormMAIN.CMDialog1.InitDir = UserDataDirectory$
If UCase$(ioextension$) = UCase$("FC") Then FormMAIN.CMDialog1.InitDir = UserDataDirectory$
If UCase$(ioextension$) = UCase$("IMG") Then FormMAIN.CMDialog1.InitDir = UserDataDirectory$
If UCase$(ioextension$) = UCase$("CLR") Then FormMAIN.CMDialog1.InitDir = SurferAppDirectory$
If UCase$(ioextension$) = UCase$("PrbAcq") Then FormMAIN.CMDialog1.InitDir = UserImagesDirectory$
If UCase$(ioextension$) = UCase$("CSV") Then FormMAIN.CMDialog1.InitDir = UserDataDirectory$
If UCase$(ioextension$) = UCase$("CND") Then FormMAIN.CMDialog1.InitDir = UserImagesDirectory$
If UCase$(ioextension$) = UCase$("ImpDAT") Then FormMAIN.CMDialog1.InitDir = UserImagesDirectory$
If UCase$(ioextension$) = UCase$("ACQ") Then FormMAIN.CMDialog1.InitDir = UserDataDirectory$

' Specify default extension
FormMAIN.CMDialog1.DefaultExt = ioextension$

' Common dialog action
FormMAIN.CMDialog1.CancelError = True
FormMAIN.CMDialog1.filename = iofilename$

If mode% < 2 Then
'FormMAIN.CMDialog1.ShowSave
'iofilename$ = FormMAIN.CMDialog1.filename
Call IOGetFileName2(mode%, tForm.hWnd, FormMAIN.CMDialog1.DialogTitle, iofilename$, FormMAIN.CMDialog1.Filter, FormMAIN.CMDialog1.flags, FormMAIN.CMDialog1.InitDir, FormMAIN.CMDialog1.DefaultExt)
If ierror Then Exit Sub
Else
'FormMAIN.CMDialog1.ShowOpen
'iofilename$ = FormMAIN.CMDialog1.filename
Call IOGetFileName2(mode%, tForm.hWnd, FormMAIN.CMDialog1.DialogTitle, iofilename$, FormMAIN.CMDialog1.Filter, FormMAIN.CMDialog1.flags, FormMAIN.CMDialog1.InitDir, FormMAIN.CMDialog1.DefaultExt)
If ierror Then Exit Sub
End If

' Check for different extension
If FormMAIN.CMDialog1.flags And cdlOFNExtensionDifferent Then GoTo IOGetFileNameBadExtension    ' works with IOGetFileName2

' Check for illegal filename
If mode% < 2 Then
Call MiscCheckName(Int(0), iofilename$)
If ierror Then Exit Sub
End If

' If saving file and file already exists, kill file
If mode% = 1 And Dir$(iofilename$) <> vbNullString Then
Kill iofilename$
End If

Exit Sub

' Errors
IOGetFileNameError:
If Err <> cdlCancel Then MsgBox Error$, vbOKOnly + vbCritical, "IOGetFileName"
ierror = True
Exit Sub

IOGetFileNameBadExtension:
msg$ = "Missing or wrong ." & ioextension$ & " extension in data file name"
MsgBox msg$, vbOKOnly + vbExclamation, "IOGetFileName"
ierror = True
Exit Sub

End Sub

Function IOGetHelpContextID(formstring As String) As Integer
' Returns the help context ID for the specified application and form from the PROBEHLP.INI file

ierror = False
On Error GoTo IOGetHelpContextIDError

Dim valid As Integer, nDefault As Integer
Dim lpAppName As String
Dim lpKeyName As String
Dim lpFileName As String

' In case loading FormMAIN (before Activate event for InitFiles is called)
If ProgramPath$ = vbNullString Then ProgramPath$ = app.Path & "\"
lpFileName$ = ProgramPath$ & "\PROBEHLP.INI"

' Check for existing PROBEHLP.INI   ' just return if not found
IOGetHelpContextID% = 1 ' default is Help Contents
If Dir$(lpFileName$) = vbNullString Then Exit Function

' Get FormMAIN as default first in case no specific form entry exists
lpAppName$ = app.EXEName
lpKeyName$ = "FormMAIN"
nDefault% = 1
valid% = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault%, lpFileName$)
IOGetHelpContextID% = valid%

' Next try to get the specific form
If formstring$ = vbNullString Then Exit Function
lpKeyName$ = formstring$
nDefault% = valid%  ' load previous FormMAIN context ID as default
valid% = GetPrivateProfileInt(lpAppName$, lpKeyName$, nDefault%, lpFileName$)
IOGetHelpContextID% = valid%

Exit Function

' Errors
IOGetHelpContextIDError:
MsgBox Error$, vbOKOnly + vbCritical, "IOGetHelpContextID"
ierror = True
Exit Function

End Function

Sub IOGetMDBFileName(mode As Integer, mdbfilename$, tForm As Form)
' Get mdb filename from user
' mode = 1 open new probe database file
' mode = 2 open old probe database file
' mode = 3 open new standard database file
' mode = 4 open old standard database file
' mode = 5 open new userwin database file
' mode = 6 open old userwin database file
' mode = 7 open new sx.mdb database file (Cameca PeakSight, mode is not used)
' mode = 8 open old sx.mdb database file (Cameca PeakSight)

ierror = False
On Error GoTo IOGetMDBFileNameError

Dim mode2 As Integer

' Set filters
FormMAIN.CMDialog1.Filter = "*.MDB (*.MDB)|*.MDB|All Files (*.*)|*.*|"

' Specify default filter
FormMAIN.CMDialog1.FilterIndex = 1

' Specify OFN Flags
If mode% = 1 Or mode% = 3 Or mode% = 5 Or mode% = 7 Then
FormMAIN.CMDialog1.flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn Or cdlOFNPathMustExist
Else
FormMAIN.CMDialog1.flags = cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn Or cdlOFNFileMustExist
End If

' Add Explorer dialog flag
FormMAIN.CMDialog1.flags = FormMAIN.CMDialog1.flags Or cdlOFNExplorer

' Specify inital directory (probe data files use Userdata directory, other use ApplicationCommonAppData$)
If mode% = 1 Or mode% = 2 Then
FormMAIN.CMDialog1.InitDir = UserDataDirectory$
Else
FormMAIN.CMDialog1.InitDir = ApplicationCommonAppData$
End If

' Specify dialog title
If mode% = 1 Then
FormMAIN.CMDialog1.DialogTitle = "Open New Probe Database File"
ElseIf mode% = 3 Then
FormMAIN.CMDialog1.DialogTitle = "Open New Standard Database File"
ElseIf mode% = 5 Then
FormMAIN.CMDialog1.DialogTitle = "Open New User Database File"
ElseIf mode% = 7 Then
FormMAIN.CMDialog1.DialogTitle = "Open New Cameca SX.MDB Database File"

ElseIf mode% = 2 Then
FormMAIN.CMDialog1.DialogTitle = "Open Old Probe Database File"
ElseIf mode% = 4 Then
FormMAIN.CMDialog1.DialogTitle = "Open Old Standard Database File"
ElseIf mode% = 6 Then
FormMAIN.CMDialog1.DialogTitle = "Open Old User Database File"
ElseIf mode% = 8 Then
FormMAIN.CMDialog1.DialogTitle = "Open Old Cameca SX.MDB Database File"
End If

' Specify default extension
FormMAIN.CMDialog1.DefaultExt = "MDB"

' Specify default if not blank
If mdbfilename$ <> vbNullString Then
FormMAIN.CMDialog1.filename = mdbfilename$
Else
FormMAIN.CMDialog1.filename = "*.mdb"
End If

' Get COMMON DIALOG Filename
If mode% = 1 Or mode% = 3 Or mode% = 5 Or mode% = 7 Then
'FormMAIN.CMDialog1.ShowSave
'mdbfilename$ = FormMAIN.CMDialog1.filename
mode2% = 1
Call IOGetFileName2(mode2%, tForm.hWnd, FormMAIN.CMDialog1.DialogTitle, mdbfilename$, FormMAIN.CMDialog1.Filter, FormMAIN.CMDialog1.flags, FormMAIN.CMDialog1.InitDir, FormMAIN.CMDialog1.DefaultExt)
If ierror Then Exit Sub
Else
'FormMAIN.CMDialog1.ShowOpen
'mdbfilename$ = FormMAIN.CMDialog1.filename
mode2% = 2
Call IOGetFileName2(mode2%, tForm.hWnd, FormMAIN.CMDialog1.DialogTitle, mdbfilename$, FormMAIN.CMDialog1.Filter, FormMAIN.CMDialog1.flags, FormMAIN.CMDialog1.InitDir, FormMAIN.CMDialog1.DefaultExt)
If ierror Then Exit Sub
End If

' If no file name returned, just exit sub
If mdbfilename$ = vbNullString Then Exit Sub

' Save path to new userdata directory (if opening probe database)
If mode% = 1 Or mode% = 2 Then
UserDataDirectory$ = MiscGetPathOnly2$(mdbfilename$)

' Set PictureSnap folder in case in demo mode
If InterfaceType% = 0 Then DemoImagesDirectory$ = OriginalDemoImagesDirectory$
End If

' Check for extension
If FormMAIN.CMDialog1.flags And cdlOFNExtensionDifferent Then GoTo IOGetMDBFileNameBadExtension

' Check for reserved file name (if new file)
If mode% = 1 Then
Call MiscCheckName(Int(0), mdbfilename$)
If ierror Then Exit Sub
ElseIf mode% = 3 Then
Call MiscCheckName(Int(1), mdbfilename$)
If ierror Then Exit Sub
ElseIf mode% = 5 Then
Call MiscCheckName(Int(2), mdbfilename$)
If ierror Then Exit Sub
ElseIf mode% = 7 Then
Call MiscCheckName(Int(0), mdbfilename$)
If ierror Then Exit Sub
End If

' If new file and file exists, then kill existing file
If Dir$(mdbfilename$) <> vbNullString And (mode% = 1 Or mode% = 3 Or mode% = 5 Or mode% = 7) Then
Kill mdbfilename$
End If

Exit Sub

' Errors
IOGetMDBFileNameError:
If Err <> cdlCancel Then MsgBox Error$, vbOKOnly + vbCritical, "IOGetMDBFileName"
ierror = True
Exit Sub

IOGetMDBFileNameBadExtension:
msg$ = "Missing or wrong .MDB extension on database file name"
MsgBox msg$, vbOKOnly + vbExclamation, "IOGetMDBFileName"
ierror = True
Exit Sub

End Sub

Sub IOLogFont()
' Change the Log window font

ierror = False
On Error GoTo IOLogFontError

' Set font flags
FormMAIN.CMDialog1.CancelError = True
FormMAIN.CMDialog1.flags = cdlCFBoth

' Load values for Log window
FormMAIN.CMDialog1.FontName = LogWindowFontName$
FormMAIN.CMDialog1.FontSize = LogWindowFontSize%
FormMAIN.CMDialog1.FontBold = LogWindowFontBold
FormMAIN.CMDialog1.FontItalic = LogWindowFontItalic
FormMAIN.CMDialog1.FontUnderline = LogWindowFontUnderline
FormMAIN.CMDialog1.FontStrikethru = LogWindowFontStrikeThru

' Call Fonts common dialog
FormMAIN.CMDialog1.ShowFont

' Update the Log Window text properties
If FormMAIN.CMDialog1.FontName <> vbNullString Then FormMAIN.TextLog.SelFontName = FormMAIN.CMDialog1.FontName
FormMAIN.TextLog.SelFontSize = FormMAIN.CMDialog1.FontSize
FormMAIN.TextLog.SelBold = FormMAIN.CMDialog1.FontBold
FormMAIN.TextLog.SelItalic = FormMAIN.CMDialog1.FontItalic
FormMAIN.TextLog.SelUnderline = FormMAIN.CMDialog1.FontUnderline
FormMAIN.TextLog.SelStrikeThru = FormMAIN.CMDialog1.FontStrikethru

' Store in global variables
If FormMAIN.TextLog.SelFontName <> vbNullString Then LogWindowFontName$ = FormMAIN.TextLog.SelFontName
LogWindowFontSize% = FormMAIN.TextLog.SelFontSize
LogWindowFontBold = FormMAIN.TextLog.SelBold
LogWindowFontItalic = FormMAIN.TextLog.SelItalic
LogWindowFontUnderline = FormMAIN.TextLog.SelUnderline
LogWindowFontStrikeThru = FormMAIN.TextLog.SelStrikeThru
Exit Sub

' Errors
IOLogFontError:
If Err <> cdlCancel Then MsgBox Error$, vbOKOnly + vbCritical, "IOLogFont"
ierror = True
Exit Sub

End Sub

Sub IOOpenOUTFile(tForm As Form)
' This routine opens the .OUT file for saving output to disk

ierror = False
On Error GoTo IOOpenOUTFileError

Dim tfilename As String
Dim response As Integer

' Get filename (do not check for existing file)
Call IOGetFileName(Int(0), "OUT", tfilename$, tForm)
If ierror Then Exit Sub

' User did not click cancel, check if it already exists
If Dir$(tfilename$) <> vbNullString Then
msg$ = "Output File: " & vbCrLf
msg$ = msg$ & tfilename$ & vbCrLf
msg$ = msg$ & " already exists, are you sure you want to overwrite it (click No to append)?"
response% = MsgBox(msg$, vbYesNoCancel + vbQuestion + vbDefaultButton2, "IOOpenOUTFile")

If response% = vbCancel Then
ierror = True
Exit Sub
End If

' Since user wants to open file make sure it is closed since it exists
Close (OutputDataFileNumber%)

' If user selects overwrite, erase it
If response% = vbYes Then
Kill tfilename$
End If
End If

' Since user did not cancel, set true and open .OUT file
If response% = vbYes Then
Open tfilename$ For Output As #OutputDataFileNumber%
Else
Open tfilename$ For Append As #OutputDataFileNumber%
End If

' No error, load filename
OutputDataFile$ = tfilename$

' Flag menu
SaveToDisk = True
FormMAIN.menuOutputSaveToDiskLog.Checked = True

Exit Sub

' Errors
IOOpenOUTFileError:
If Err <> cdlCancel Then MsgBox Error$, vbOKOnly + vbCritical, "IOOpenOUTFile"
ierror = True
Exit Sub

End Sub

Sub IOPrintLog()
' Print the log window text

ierror = False
On Error GoTo IOPrintLogError

' CancelError is True
FormMAIN.CMDialog1.CancelError = True

' Check for text
If FormMAIN.TextLog.Text = vbNullString Then GoTo IOPrintLogNoText

' Set flags
FormMAIN.CMDialog1.flags = cdlPDReturnDC + cdlPDNoPageNums
If FormMAIN.TextLog.SelLength = 0 Then
FormMAIN.CMDialog1.flags = FormMAIN.CMDialog1.flags + cdlPDAllPages
Else
FormMAIN.CMDialog1.flags = FormMAIN.CMDialog1.flags + cdlPDSelection
End If
    
FormMAIN.CMDialog1.ShowPrinter

' Print text
Screen.MousePointer = vbHourglass
'Printer.Print vbNullString
FormMAIN.TextLog.SelPrint FormMAIN.CMDialog1.hDC
'FormMAIN.TextLog.SelPrint FormMAIN.CMDialog1.hDC, SelPrintStartDocFlag% ' VB6 only
Printer.EndDoc
Screen.MousePointer = vbDefault

Exit Sub

' Errors
IOPrintLogError:
Screen.MousePointer = vbDefault
If Err <> cdlCancel Then MsgBox Error$, vbOKOnly + vbCritical, "IOPrintLog"
ierror = True
Exit Sub

IOPrintLogNoText:
msg$ = "No text in log window to print"
MsgBox msg$, vbOKOnly + vbExclamation, "IOPrintLog"
ierror = True
Exit Sub

End Sub

Sub IOPrintSetup()
' Open a Print Setup Dialog box

ierror = False
On Error GoTo IOPrintSetupError

' CancelError is True
FormMAIN.CMDialog1.CancelError = True

FormMAIN.CMDialog1.flags = cdlPDPrintSetup
FormMAIN.CMDialog1.PrinterDefault = True
FormMAIN.CMDialog1.ShowPrinter
Exit Sub

' Errors
IOPrintSetupError:
If Err <> cdlCancel Then MsgBox Error$, vbOKOnly + vbCritical, "IOPrintSetup"
ierror = True
Exit Sub

End Sub

Sub IOViewLog()
' This routine is called to view the current .OUT file

ierror = False
On Error GoTo IOViewLogError

Dim taskID As Long  ' for 32 bit OS

' Check for current .OUT file
If Trim$(OutputDataFile$) = vbNullString Then GoTo IOViewLogNoFile

' If "SaveToDisk" is true, first close .OUT file
If SaveToDisk Then
Close #OutputDataFileNumber%
End If

' Use default file viewer app to view copy of file
FileCopy OutputDataFile$, UserDataDirectory$ & "\temp.out"

' If "SaveToDisk" is true, open .OUT file
If SaveToDisk Then
Open OutputDataFile$ For Append As #OutputDataFileNumber%
End If

' View file
msg$ = VbDquote$ & FileViewer$ & VbDquote$ & " " & VbDquote$ & UserDataDirectory$ & "\temp.out" & VbDquote$
taskID& = Shell(msg$, vbNormalFocus)

Exit Sub

' Errors
IOViewLogError:
msg$ = vbCrLf & "Make sure that the file viewer executable " & FileViewer$ & ", is correctly specified."
MsgBox Error$ & msg$, vbOKOnly + vbCritical, "IOViewLog"
Close #OutputDataFileNumber%
ierror = True
Exit Sub

IOViewLogNoFile:
msg$ = "There is no current disk log file"
MsgBox msg$, vbOKOnly + vbExclamation, "IOViewLog"
ierror = True
Exit Sub

End Sub

Sub IOViewFile(ioextension As String, tForm As Form)
' This routine is called to view a file with the specified extension

ierror = False
On Error GoTo IOViewFileError

Dim taskID As Long  ' for 32 bit OS
Dim tfilename As String

Call IOGetFileName(Int(2), ioextension$, tfilename$, tForm)
If ierror Then Exit Sub

' View file
msg$ = VbDquote$ & FileViewer$ & VbDquote$ & " " & VbDquote$ & tfilename$ & VbDquote$
taskID& = Shell(msg$, vbNormalFocus)

Exit Sub

' Errors
IOViewFileError:
msg$ = vbCrLf & "Make sure that the file viewer executable " & FileViewer$ & ", is correctly specified."
MsgBox Error$ & msg$, vbOKOnly + vbCritical, "IOViewFile"
ierror = True
Exit Sub

End Sub

Sub IODisplayVersionInfo()
' Routine to display the version.txt file

ierror = False
On Error GoTo IODisplayVersionInfoError

Dim taskID As Long  ' for 32 bit OS
Dim tfilename As String

' View version file
tfilename$ = ApplicationCommonAppData$ & "version.txt"
If Dir$(tfilename$) = vbNullString Then GoTo IODisplayVersionInfoNoFile
msg$ = VbDquote$ & FileViewer$ & VbDquote$ & " " & VbDquote$ & tfilename$ & VbDquote$
taskID& = Shell(msg$, vbNormalFocus)

Exit Sub

' Errors
IODisplayVersionInfoError:
msg$ = vbCrLf & "Make sure that the file viewer executable " & FileViewer$ & ", is correctly specified."
MsgBox Error$ & msg$, vbOKOnly + vbCritical, "IODisplayVersionInfo"
ierror = True
Exit Sub

IODisplayVersionInfoNoFile:
msg$ = "Missing VERSION.TXT file"
MsgBox msg$, vbOKOnly + vbExclamation, "IODisplayVersionInfo"
ierror = True
Exit Sub

End Sub

Sub IOWriteLog(astring As String)
' This routine concatenates the passed string to the "LogString"
' global variables. When the TimerLogWindow event occurs, routine
' IODumpLog is called to write the string to the Log window.

ierror = False
On Error GoTo IOWriteLogError

' If time stamp mode then add time of day
If TimeStampMode Then astring$ = Time$ & ": " & astring$

' Concatenate string
LogString$ = LogString$ & astring$ & vbCrLf

' Check for string length and dump if getting too long
If Len(LogString$) > 500 Then Call IODumpLog

Exit Sub

' Errors
IOWriteLogError:
MsgBox Error$, vbOKOnly + vbCritical, "IOWriteLog"
ierror = True
Exit Sub

End Sub

Sub IODumpLog()
' This routine is called by "TimerLogWindow" event to write the contents of the "LogString" to the Log window.
' If "SaveToDisk" is true then also write to .OUT file
' If "SaveToText" is true then also write to .TXT file

ierror = False
On Error GoTo IODumpLogError

' Append to Log window text if not empty string
If LogString$ = vbNullString Then Exit Sub

' Check that the text buffer is filled up, if so clip from end
If Len(FormMAIN.TextLog.Text) > LogWindowBufferSize& Then
FormMAIN.TextLog.SelStart = 0
FormMAIN.TextLog.SelLength = 4000   ' remove 4K at a time
FormMAIN.TextLog.SelText = vbNullString
End If

' Set the insertion point
Call IODumpLogFont
If ierror Then Exit Sub

' Append the new text
FormMAIN.TextLog.SelText = LogString$

' Reset the insertion point
Call IODumpLogFont
If ierror Then Exit Sub

' If "SaveToDisk" is true then also write to .OUT file (use semi-colon to prevent an extra <cr>)
If SaveToDisk Then
Print #OutputDataFileNumber%, LogString$;
End If

' If "SaveToText" is true then also write to .TXT file (use semi-colon to prevent an extra <cr>)
If SaveToText Then
Print #OutputReportFileNumber%, LogString$;
End If

' Clear the "LogString"
LogString$ = vbNullString
Exit Sub

' Errors
IODumpLogError:
Screen.MousePointer = vbDefault
MsgBox Error$ & ", file handle= " & Format$(OutputDataFileNumber%), vbOKOnly + vbCritical, "IODumpLog"
ierror = True
SaveToDisk = False
SaveToText = False
Exit Sub

End Sub

Sub IOSendLog(ichar As Integer)
' This routine sends directly to the log file if specified.

ierror = False
On Error GoTo IOSendLogError

If SaveToDisk Then
Print #OutputDataFileNumber%, Chr$(ichar%);  ' only use ChrW$ for 0 to 127 ASCII
If ichar% = 13 Then
Print #OutputDataFileNumber%, vbLf;
End If
End If

Exit Sub

' Errors
IOSendLogError:
MsgBox Error$, vbOKOnly + vbCritical, "IOSendLog"
ierror = True
Exit Sub

End Sub

Sub IODumpLogFont()
' Sets insertion point and restores log window font

ierror = False
On Error GoTo IODumpLogFontError

' Set the insertion point
FormMAIN.TextLog.SelStart = Len(FormMAIN.TextLog.Text)
FormMAIN.TextLog.SelLength = 0

' Restore font
FormMAIN.TextLog.SelFontName = LogWindowFontName$
FormMAIN.TextLog.SelFontSize = LogWindowFontSize%
FormMAIN.TextLog.SelColor = vbBlack
FormMAIN.TextLog.SelBold = LogWindowFontBold
FormMAIN.TextLog.SelItalic = LogWindowFontItalic
FormMAIN.TextLog.SelUnderline = LogWindowFontUnderline
FormMAIN.TextLog.SelStrikeThru = LogWindowFontStrikeThru
FormMAIN.TextLog.SelIndent = 0

Exit Sub

' Errors
IODumpLogFontError:
MsgBox Error$, vbOKOnly + vbCritical, "IODumpLogFont"
ierror = True
Exit Sub

End Sub

Sub IODumpText(astring As String, tText As TextBox)
' Simple append text to passed textbox

ierror = False
On Error GoTo IODumpTextError

' Set the insertion point
tText.SelStart = Len(tText.Text)
tText.SelLength = 0
tText.SelText = astring$ & vbCrLf

Exit Sub

' Errors
IODumpTextError:
MsgBox Error$, vbOKOnly + vbCritical, "IODumpText"
ierror = True
Exit Sub

End Sub

Sub IOShellProcess(tstring As String)
' Shell a separate process for string

ierror = False
On Error GoTo IOShellProcessError

Dim taskID As Long

' Shell process
taskID& = Shell(tstring$, vbNormalFocus)

Exit Sub

' Errors
IOShellProcessError:
MsgBox Error$, vbOKOnly + vbCritical, "IOShellProcess"
ierror = True
Exit Sub

End Sub

Sub IOTextViewer()
' This routine is called to run the file viewer

ierror = False
On Error GoTo IOTextViewerError

Dim taskID As Long  ' for 32 bit OS

' Open file viewer
msg$ = VbDquote$ & FileViewer$ & VbDquote$
taskID& = Shell(msg$, vbNormalFocus)

Exit Sub

' Errors
IOTextViewerError:
msg$ = vbCrLf & "Make sure that the file viewer executable " & FileViewer$ & ", is correctly specified."
MsgBox Error$ & msg$, vbOKOnly + vbCritical, "IOTextViewer"
ierror = True
Exit Sub

End Sub

Sub IOTextViewer2(tfilename As String)
' This routine is called to run the file viewer and load the passed file

ierror = False
On Error GoTo IOTextViewer2Error

Dim taskID As Long  ' for 32 bit OS

' Open file viewer
If Trim$(tfilename$) = vbNullString Then Exit Sub
If Dir$(tfilename$) = vbNullString$ Then Exit Sub

msg$ = VbDquote$ & FileViewer$ & VbDquote$ & " " & VbDquote$ & tfilename$ & VbDquote$
taskID& = Shell(msg$, vbNormalFocus)

Exit Sub

' Errors
IOTextViewer2Error:
msg$ = vbCrLf & "Make sure that the file viewer executable " & FileViewer$ & ", is correctly specified."
MsgBox Error$ & msg$, vbOKOnly + vbCritical, "IOTextViewer2"
ierror = True
Exit Sub

End Sub

Sub IOWriteLogRichText(astring As String, tFontName As String, tFontSize As Integer, tFontColor As Long, tformat As Integer, tIndent As Integer)
' This is a version of the IOWriteLog procedure that handles Rich Text strings

ierror = False
On Error GoTo IOWriteLogRichTextError

' Send any existing text to the control
Call IODumpLog

' Set the insertion point
FormMAIN.TextLog.SelStart = Len(FormMAIN.TextLog.Text)
FormMAIN.TextLog.SelLength = 0

' Set font attributes
If tFontName$ = vbNullString Then
FormMAIN.TextLog.SelFontName = LogWindowFontName$
Else
FormMAIN.TextLog.SelFontName = tFontName$
End If

' Set font size attributes
If tFontSize% = 0 Then
FormMAIN.TextLog.SelFontSize = LogWindowFontSize%
Else
FormMAIN.TextLog.SelFontSize = tFontSize%
End If

' Set font color attributes
If tFontColor& = 0 Then
FormMAIN.TextLog.SelColor = vbBlack
Else
FormMAIN.TextLog.SelColor = tFontColor&
End If

' Set font format attributes
If Not tformat% And 1 Then
FormMAIN.TextLog.SelBold = LogWindowFontBold
Else
FormMAIN.TextLog.SelBold = True
End If

If Not tformat% And 2 Then
FormMAIN.TextLog.SelItalic = LogWindowFontItalic
Else
FormMAIN.TextLog.SelItalic = True
End If

If Not tformat% And 4 Then
FormMAIN.TextLog.SelUnderline = LogWindowFontUnderline
Else
FormMAIN.TextLog.SelUnderline = True
End If

If Not tformat% And 8 Then
FormMAIN.TextLog.SelStrikeThru = LogWindowFontStrikeThru
Else
FormMAIN.TextLog.SelStrikeThru = True
End If

' Set indent attributes
If tIndent% = 0 Then
FormMAIN.TextLog.SelIndent = 0
Else
FormMAIN.TextLog.SelIndent = tIndent%
End If

' Output Rich Text string
FormMAIN.TextLog.SelText = astring$ & vbCrLf

' Reset the insertion point
Call IODumpLogFont
If ierror Then Exit Sub

' If "SaveToDisk" is true then also write to .OUT file (won't save rich text attributes)
If SaveToDisk Then
Print #OutputDataFileNumber%, astring$
End If

Exit Sub

' Errors
IOWriteLogRichTextError:
MsgBox Error$, vbOKOnly + vbCritical, "IOWriteLogRichText"
ierror = True
Exit Sub

End Sub

Sub IOWriteLogRichText2(astring As String, tFontName As String, tFontSize As Integer, tFontColor As Long, tformat As Integer, tIndent As Integer, tSubScript() As Boolean, tSuperScript() As Boolean)
' This is a version of the IOWriteLogRichText procedure that handles sub and superscripts (does not send vbCrLf)
' Note: tSubScript = array that contains flags in the character positions of characters to be subscripted
' Note: tSuperScript = array that contains flags in the character positions of characters to be superscripted

' Note: routine is not finished!!!!

ierror = False
On Error GoTo IOWriteLogRichText2Error

' Send any existing text to the control
Call IODumpLog

' Set insertion point
FormMAIN.TextLog.SelStart = Len(FormMAIN.TextLog.Text)
FormMAIN.TextLog.SelLength = 0

' Set font attributes



' Output Rich Text string
FormMAIN.TextLog.SelText = astring$ & vbCrLf

' Reset the insertion point
Call IODumpLogFont
If ierror Then Exit Sub

' If "SaveToDisk" is true then also write to .OUT file (won't save rich text attributes)
If SaveToDisk Then
Print #OutputDataFileNumber%, astring$
End If

Exit Sub

' Errors
IOWriteLogRichText2Error:
MsgBox Error$, vbOKOnly + vbCritical, "IOWriteLogRichText2"
ierror = True
Exit Sub

End Sub

Sub IOTextLoadFile(tfilename As String, tText As RichTextBox)
' This routine is called to load a text file to the passed text box (usually the log window)

ierror = False
On Error GoTo IOTextLoadFileError

Dim astring As String, bstring As String

bstring$ = vbNullString
Open tfilename$ For Input As #Temp1FileNumber%
Do Until EOF(Temp1FileNumber%)
Line Input #Temp1FileNumber%, astring$
bstring$ = bstring$ & astring$ & vbCrLf
Loop
Close #Temp1FileNumber%

tText.SelText = bstring$ & vbCrLf
Exit Sub

' Errors
IOTextLoadFileError:
MsgBox Error$, vbOKOnly + vbCritical, "IOTextLoadFile"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub IOTextSaveFile(tfilename As String, tText As RichTextBox)
' This routine is called to save a text file from the passed text box (usually the log window)

ierror = False
On Error GoTo IOTextSaveFileError

Dim astring As String

tText.SetFocus
tText.SelStart = 0
tText.SelLength = Len(tText.Text)
astring$ = tText.SelText

Open tfilename$ For Output As #Temp1FileNumber%
Print #Temp1FileNumber%, astring$
Close #Temp1FileNumber%

Exit Sub

' Errors
IOTextSaveFileError:
MsgBox Error$, vbOKOnly + vbCritical, "IOTextSaveFile"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub IOGetFileName2(mode As Integer, tFormhWnd As Long, tTitle As String, tfilename As String, tFilter As String, tFlags As Long, tInitDir As String, textension As String)
' Alternative code to call API for Windows 7 CCL offset default filename bug
'  mode = 0 save new file (do not check for existing file to allow append mode)
'  mode = 1 save new file
'  mode = 2 open old file
  
ierror = False
On Error GoTo IOGetFileName2Error

Dim n As Long

Dim cdl_OpenFileName As Type_CDL_OpenFileName

Const cdl_MaxBuffer = 512

  With cdl_OpenFileName
    .hwndOwner = tFormhWnd
    
    .hInstance = app.hInstance
    .lpstrTitle = tTitle$
    .lpstrInitialDir = tInitDir$
    .flags = tFlags&
    
    ' Replace "|" with null character
    tFilter$ = Replace$(tFilter$, "|", vbNullChar)
    .lpstrFilter = tFilter$
    .nFilterIndex = 1
    
    .lpstrDefExt = textension$
        
    .lpstrFile = tfilename$ & String(cdl_MaxBuffer, 0) ' default file (must add extra buffer for Win 7 cdl bug)
    .nMaxFile = Len(.lpstrFile)
    
    .lpstrFileTitle = String(cdl_MaxBuffer, 0)     ' filename only
    .nMaxFileTitle = Len(.lpstrFileTitle)
       
    .lStructSize = Len(cdl_OpenFileName)
  End With

' Save file
If mode% < 2 Then
    If GetSaveFileName(cdl_OpenFileName) <> 0 Then        ' file was selected
        tfilename$ = Left$(cdl_OpenFileName.lpstrFile, cdl_OpenFileName.nMaxFile)
        n& = InStr(tfilename$, vbNullChar)
        tfilename$ = Left$(tfilename$, n& - 1)     ' remove trailing nulls
        tFlags& = cdl_OpenFileName.flags      ' return flags
    Else
        ierror = True                                       ' CANCEL button was pressed
        Exit Sub
    End If

' Open file
Else
    If GetOpenFileName(cdl_OpenFileName) <> 0 Then          ' file was selected
        tfilename$ = Left$(cdl_OpenFileName.lpstrFile, cdl_OpenFileName.nMaxFile)
        n& = InStr(tfilename$, vbNullChar)
        tfilename$ = Left$(tfilename$, n& - 1)     ' remove trailing nulls
        tFlags& = cdl_OpenFileName.flags      ' return flags
    Else
        ierror = True                                       ' CANCEL button was pressed
        Exit Sub
    End If
End If
    
Exit Sub

' Errors
IOGetFileName2Error:
MsgBox Error$, vbOKOnly + vbCritical, "IOGetFileName2"
ierror = True
Exit Sub

End Sub

Sub IOBrowseHTTP(method As Integer, sURL As String)
' This routine is called to handle Internet requests (WWW or DVD)
'  method% = 0  Use WWW
'  method% = 1  Use DVD
'  Probe Software web:     file:///F:/Probe%20Software%20Web%20Site/probesoftware.com/index.html
'  Probe Software Forum:   file:///F:/Probe%20Software%20Web%20Site/smf.probesoftware.com/index.html

ierror = False
On Error GoTo IOBrowseHTTPError

Dim aURL As String, astring As String
Dim n As Integer

Static sDVD As String

' Always check for empty passed string
If sURL$ = vbNullString Then GoTo IOBrowseHTTPNoURL

' Using WWW
If method% = 0 Then

' Load into local variable
aURL$ = sURL$

' Using DVD
Else

' Check for empty DVD path
If sDVD$ = vbNullString Then

' Ask user for DVD drive location
sDVD$ = IOBrowseForFolderByPath(False, sDVD$ & "\", "Please browse to the DVD drive containing the Probe Software Web Site disk", FormMAIN)
If ierror Then Exit Sub

' Double check that DVD path is valid
If sDVD$ = vbNullString Then GoTo IOBrowseHTTPDriveNotFound

' Check for first backslash
n% = InStr(sDVD$, "\")
If n% <> 3 Then GoTo IOBrowseHTTPFolderNotFound

' Trim just in case
sDVD$ = Left$(sDVD$, 3)

' Check for the proper folder
If Dir$(sDVD$ & "Probe Software Web Site", vbDirectory) = vbNullString Then GoTo IOBrowseHTTPFolderNotFound
End If

' Remove https"/
astring$ = Replace$(sURL$, "https:/", vbNullString)

' Load specific DVD path
aURL$ = "file:///" & sDVD$ & "Probe%20Software%20Web%20Site" & astring$

' Change backslashes to forward slashes
aURL$ = Replace$(aURL$, "\", "/")

' Change index.php to index.html (no php support from DVD to resolve URLs to specific boards and topics)
aURL$ = Replace$(aURL$, "index.php", "index.html")

' Temp code to fix local browse issue for smf links
n% = InStr(aURL$, "/smf/")
If n% > 0 Then
sDVD$ = Left$(sDVD$, 2) & "/"   ' replace backslash with forward slash
If aURL$ <> "file:///" & sDVD$ & "Probe%20Software%20Web%20Site/smf.probesoftware.com/index.html" Then
If InStr(UCase$(app.EXEName), UCase$("CalcZAF")) > 0 Then aURL$ = "file:///" & sDVD$ & "Probe%20Software%20Web%20Site/smf.probesoftware.com/index8b25.html?board=7.0"
If InStr(UCase$(app.EXEName), UCase$("Standard")) > 0 Then aURL$ = "file:///" & sDVD$ & "Probe%20Software%20Web%20Site/smf.probesoftware.com/index8b25.html?board=7.0"
If InStr(UCase$(app.EXEName), UCase$("Probewin")) > 0 Then aURL$ = "file:///" & sDVD$ & "Probe%20Software%20Web%20Site/smf.probesoftware.com/index9c2d.html?board=2.0"
If InStr(UCase$(app.EXEName), UCase$("CalcImage")) > 0 Then aURL$ = "file:///" & sDVD$ & "Probe%20Software%20Web%20Site/smf.probesoftware.com/indexfc47.html?board=4.0"
End If
End If

' Actual local file URLs in browser for How To Do Quant Mapping Part I, II and III
'aURL$ = "file:///F:/Probe%20Software%20Web%20Site/smf.probesoftware.com/index782e.html?topic=106.0"
'aURL$ = "file:///F:/Probe%20Software%20Web%20Site/smf.probesoftware.com/indexe8c6.html?topic=141.0"
'aURL$ = "file:///F:/Probe%20Software%20Web%20Site/smf.probesoftware.com/index5c04.html?topic=146.0"
End If

' Default is use Internet
Call IORunShellExecute("open", aURL$, 0&, 0&, SW_SHOWNORMAL&)   ' open the URL using the default browser
If ierror Then Exit Sub

Exit Sub

' Errors
IOBrowseHTTPError:
MsgBox Error$, vbOKOnly + vbCritical, "IOBrowseHTTP"
sDVD$ = vbNullString
ierror = True
Exit Sub

IOBrowseHTTPNoURL:
msg$ = "No URL was passed. This error should not occur. Please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "IOBrowseHTTP"
sDVD$ = vbNullString
ierror = True
Exit Sub

IOBrowseHTTPDriveNotFound:
msg$ = "The DVD drive was not specified. Please make sure a DVD disk is loaded into the proper DVD drive."
MsgBox msg$, vbOKOnly + vbExclamation, "IOBrowseHTTP"
sDVD$ = vbNullString
ierror = True
Exit Sub

IOBrowseHTTPFolderNotFound:
msg$ = "The Probe Software Web Site folder was not found in the specified drive path (" & sDVD$ & ")." & vbCrLf & vbCrLf
msg$ = msg$ & "Please make sure the Probe Software Web DVD is loaded into the proper DVD drive and that the root folder only is selected."
MsgBox msg$, vbOKOnly + vbExclamation, "IOBrowseHTTP"
sDVD$ = vbNullString
ierror = True
Exit Sub

End Sub


