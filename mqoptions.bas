Attribute VB_Name = "CodeMQOPTIONS"
' (c) Copyright 1995-2021 by John J. Donovan
Option Explicit

Const MAXOUTPUT% = 4
Const MAXZBAR% = 8

' MQ options
Dim MassZedDiffMinFlag As Integer
Dim MassZedDiffMaxFlag As Integer
Dim MassZedDiffMax As Single
Dim MassZedDiffMin As Single

Dim FilmDensity As Single
Dim FilmThickness As Single

Dim SubstrateAtomicNumber As Integer
Dim SubstrateXrayLine As String
Dim SubstrateDensity As Single
Dim SubstrateThickness As Single

Dim NumberofTrajectories As Long
Dim HistogramRange As Single

Dim SecondaryEnergy As Single

Dim ElementPath As String

Dim MQOptionsTmpSample(1 To 1) As TypeSample

Sub MqOptionsLoad()
' Load the form

ierror = False
On Error GoTo MqOptionsLoadError

' Init
Call MqOptionsInit
If ierror Then Exit Sub

' Load conditions
FormMQOPTIONS.LabelTakeOff.Caption = Str$(DefaultTakeOff!)
FormMQOPTIONS.LabelKiloVolts.Caption = Str$(DefaultKiloVolts!)

' Load options
If MassZedDiffMinFlag Then
FormMQOPTIONS.CheckMassZedDiffMin.Value = vbChecked
Else
FormMQOPTIONS.CheckMassZedDiffMin.Value = vbUnchecked
End If

If MassZedDiffMaxFlag Then
FormMQOPTIONS.CheckMassZedDiffMax.Value = vbChecked
Else
FormMQOPTIONS.CheckMassZedDiffMax.Value = vbUnchecked
End If

FormMQOPTIONS.TextMassZedDiffMin.Text = Str$(MassZedDiffMin!)
FormMQOPTIONS.TextMassZedDiffMax.Text = Str$(MassZedDiffMax!)

FormMQOPTIONS.TextFilmDensity.Text = Str$(FilmDensity!)
FormMQOPTIONS.TextFilmThickness.Text = Str$(FilmThickness!)

FormMQOPTIONS.TextSubstrateAtomicNumber.Text = Str$(SubstrateAtomicNumber%)
FormMQOPTIONS.TextSubstrateXrayLine.Text = SubstrateXrayLine$

FormMQOPTIONS.TextSubstrateDensity.Text = Str$(SubstrateDensity!)
FormMQOPTIONS.TextSubstrateThickness.Text = Str$(SubstrateThickness!)

FormMQOPTIONS.TextNumberofTrajectories.Text = Str$(NumberofTrajectories&)
FormMQOPTIONS.TextHistogramRange.Text = Str$(HistogramRange!)

FormMQOPTIONS.TextSecondaryEnergy.Text = Str$(SecondaryEnergy!)

Exit Sub

' Errors
MqOptionsLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "MqOptionsLoad"
ierror = True
Exit Sub

End Sub

Sub MqOptionsSave()
' Save the form

ierror = False
On Error GoTo MqOptionsSaveError

' Save options
If FormMQOPTIONS.CheckMassZedDiffMin.Value = vbChecked Then
MassZedDiffMinFlag = True
Else
MassZedDiffMinFlag = False
End If
If FormMQOPTIONS.CheckMassZedDiffMax.Value = vbChecked Then
MassZedDiffMaxFlag = True
Else
MassZedDiffMaxFlag = False
End If

If Val(FormMQOPTIONS.TextMassZedDiffMin.Text) > 0# And Val(FormMQOPTIONS.TextMassZedDiffMin.Text) < 50 Then
MassZedDiffMin! = Val(FormMQOPTIONS.TextMassZedDiffMin.Text)
Else
msg$ = "Mass-Zed minimum difference is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "MqOptionsSave"
ierror = True
Exit Sub
End If

If Val(FormMQOPTIONS.TextMassZedDiffMax.Text) > 0# And Val(FormMQOPTIONS.TextMassZedDiffMax.Text) < 50 Then
MassZedDiffMax! = Val(FormMQOPTIONS.TextMassZedDiffMax.Text)
Else
msg$ = "Mass-Zed maximum difference is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "MqOptionsSave"
ierror = True
Exit Sub
End If

If Val(FormMQOPTIONS.TextFilmDensity.Text) > 0.1 And Val(FormMQOPTIONS.TextFilmDensity.Text) <= 100# Then
FilmDensity! = Val(FormMQOPTIONS.TextFilmDensity.Text)
Else
msg$ = "Film density is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "MqOptionsSave"
ierror = True
Exit Sub
End If

If Val(FormMQOPTIONS.TextFilmThickness.Text) > 0.001 And Val(FormMQOPTIONS.TextFilmThickness.Text) <= 1000000# Then
FilmThickness! = Val(FormMQOPTIONS.TextFilmThickness.Text)
Else
msg$ = "Film thickness is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "MqOptionsSave"
ierror = True
Exit Sub
End If

If Val(FormMQOPTIONS.TextSubstrateAtomicNumber.Text) > 2 And Val(FormMQOPTIONS.TextSubstrateAtomicNumber.Text) <= MAXELM% Then
SubstrateAtomicNumber% = Val(FormMQOPTIONS.TextSubstrateAtomicNumber.Text)
Else
msg$ = "Substrate atomic number is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "MqOptionsSave"
ierror = True
Exit Sub
End If

If FormMQOPTIONS.TextSubstrateXrayLine.Text = "K" Or FormMQOPTIONS.TextSubstrateXrayLine.Text = "L" Or FormMQOPTIONS.TextSubstrateXrayLine.Text = "M" Then
SubstrateXrayLine$ = FormMQOPTIONS.TextSubstrateXrayLine.Text
Else
msg$ = "Substrate x-ray line is invalid"
MsgBox msg$, vbOKOnly + vbExclamation, "MqOptionsSave"
ierror = True
Exit Sub
End If

If Val(FormMQOPTIONS.TextSubstrateDensity.Text) > 0.1 And Val(FormMQOPTIONS.TextSubstrateDensity.Text) <= 100# Then
SubstrateDensity! = Val(FormMQOPTIONS.TextSubstrateDensity.Text)
Else
msg$ = "Substrate density is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "MqOptionsSave"
ierror = True
Exit Sub
End If

If Val(FormMQOPTIONS.TextSubstrateThickness.Text) > 0.001 And Val(FormMQOPTIONS.TextSubstrateThickness.Text) <= 1000000# Then
SubstrateThickness! = Val(FormMQOPTIONS.TextSubstrateThickness.Text)
Else
msg$ = "Substrate thickness is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "MqOptionsSave"
ierror = True
Exit Sub
End If

If Val(FormMQOPTIONS.TextNumberofTrajectories.Text) > 0 And Val(FormMQOPTIONS.TextNumberofTrajectories.Text) <= 10000000 Then
NumberofTrajectories& = Val(FormMQOPTIONS.TextNumberofTrajectories.Text)
Else
msg$ = "Number of trajectories is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "MqOptionsSave"
ierror = True
Exit Sub
End If
 
If Val(FormMQOPTIONS.TextHistogramRange.Text) > 0.001 And Val(FormMQOPTIONS.TextHistogramRange.Text) <= 1000# Then
HistogramRange! = Val(FormMQOPTIONS.TextHistogramRange.Text)
Else
msg$ = "Histogram range is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "MqOptionsSave"
ierror = True
Exit Sub
End If

If Val(FormMQOPTIONS.TextSecondaryEnergy.Text) > 0# And Val(FormMQOPTIONS.TextSecondaryEnergy.Text) <= DefaultKiloVolts! Then
SecondaryEnergy! = Val(FormMQOPTIONS.TextSecondaryEnergy.Text)
Else
msg$ = "Secondary energy is out of range"
MsgBox msg$, vbOKOnly + vbExclamation, "MqOptionsSave"
ierror = True
Exit Sub
End If

Exit Sub

' Errors
MqOptionsSaveError:
MsgBox Error$, vbOKOnly + vbCritical, "MqOptionsSave"
ierror = True
Exit Sub

End Sub

Sub MqOptionsSendCompToFile(maszbar As Single, zedzbar As Single, bfilename As String, sample() As TypeSample)
' Output the passed data to a MQ (monte-carlo) input file

ierror = False
On Error GoTo MqOptionsSendCompToFileError

Dim i As Integer
Dim xsym As String
Dim tfilename As String, astring As String

Static initialized(0 To 2) As Boolean

' Init
Call MqOptionsInit
If ierror Then Exit Sub

' Check if mass-electron difference is greater than tolerance
If maszbar! > 0# And MassZedDiffMinFlag Then
If Abs((maszbar! - zedzbar!) / maszbar! * 100#) < MassZedDiffMin! Then Exit Sub
End If

If maszbar! > 0# And MassZedDiffMaxFlag Then
If Abs((maszbar! - zedzbar!) / maszbar! * 100#) > MassZedDiffMax! Then Exit Sub
End If

' Check for sub-directory
astring$ = Dir$(UserDataDirectory$ & "\MQData", vbDirectory)
If Not MiscStringsAreSame(astring$, "MQData") Then
MkDir UserDataDirectory$ & "\MQData"
End If

' Load file name based on standard name
astring$ = Trim$(Format$(sample(1).kilovolts!)) & "-" & Trim$(sample(1).Name$) & ".INP"

' Check for invalid filename characters
Call MiscModifyStringToFilename(astring$)
If ierror Then Exit Sub

' Open file
tfilename$ = UserDataDirectory$ & "\MQData\" & astring$
Open tfilename$ For Output As #Temp1FileNumber%

' Output configuration
Print #Temp1FileNumber%, "u"    ' units
Print #Temp1FileNumber%, "1"    ' film on substrate (vs. embedded object)
Print #Temp1FileNumber%, Format$(sample(1).LastChan%) & VbComma$ & Trim$(Str$(FilmDensity!))  ' filem number of elements and density
Print #Temp1FileNumber%, Trim$(Str$(FilmThickness!))  ' use 100 microns thick for infinitly thick

' Output composition of film
For i% = 1 To sample(1).LastChan%
xsym$ = UCase$(Left$(sample(1).Xrsyms$(i%), 1))
If Trim$(xsym$) = vbNullString Then xsym$ = "K"  ' H and He
astring$ = Format$(sample(1).AtomicNums%(i%)) & VbComma$ & Trim$(MiscAutoFormat$(sample(1).ElmPercents!(i%) / 100#)) & VbComma$ & xsym$
Print #Temp1FileNumber%, astring$
Next i%

Print #Temp1FileNumber%, "1" & VbComma$ & Trim$(Str$(SubstrateDensity!))    ' substrate number of elements and density of substrate
Print #Temp1FileNumber%, Trim$(Str$(SubstrateThickness!))      ' 100 microns thick for infinitly thick
Print #Temp1FileNumber%, Trim$(Str$(SubstrateAtomicNumber%)) & VbComma$ & "1" & VbComma$ & Trim$(SubstrateXrayLine$)    ' at. #, wt. frac. and x-ray line of substrate"
Print #Temp1FileNumber%, Trim$(Str$(90 - sample(1).takeoff!))       ' angle from electron beam to detector
Print #Temp1FileNumber%, "90"       ' horizontal angle from detector to tilt axis
Print #Temp1FileNumber%, "5021"     ' random number
Print #Temp1FileNumber%, Trim$(Left$(Format$(sample(1).kilovolts!), 4))     ' operating volatge
Print #Temp1FileNumber%, "1"        ' Gaussian beam spread
Print #Temp1FileNumber%, ".1"       ' beam diameter
Print #Temp1FileNumber%, "0"        ' specimen tilt (zero = normal)
Print #Temp1FileNumber%, Trim$(Str$(NumberofTrajectories&))
Print #Temp1FileNumber%, Trim$(Str$(HistogramRange!))     ' (microns)
Print #Temp1FileNumber%, Trim$(Str$(SecondaryEnergy!))                 ' minimum energy for secondary generation

Close #Temp1FileNumber%

' Check if first time, if so, delete file
If InStr(bfilename$, "STANDARD.BAT") And Not initialized(0) Then
On Error Resume Next            ' next line causes error if "limited" Windows account
If Dir$(bfilename$) <> vbNullString Then Kill bfilename$
initialized(0) = True
On Error GoTo MqOptionsSendCompToFileError
End If

If InStr(bfilename$, "ELEMENT.BAT") And Not initialized(1) Then
On Error Resume Next            ' next line causes error if "limited" Windows account
If Dir$(bfilename$) <> vbNullString Then Kill bfilename$
initialized(1) = True
On Error GoTo MqOptionsSendCompToFileError
End If

If InStr(bfilename$, "BINARY.BAT") And Not initialized(2) Then
On Error Resume Next            ' next line causes error if "limited" Windows account
If Dir$(bfilename$) <> vbNullString Then Kill bfilename$
initialized(2) = True
On Error GoTo MqOptionsSendCompToFileError
End If

' Save call to MCARLO batch file
Open bfilename$ For Append As #Temp1FileNumber%
astring$ = "call ..\mcarlo " & VbDquote$ & MiscGetFileNameNoExtension$(MiscGetFileNameOnly$(tfilename$)) & VbDquote$
Print #Temp1FileNumber%, astring$
Close #Temp1FileNumber%

Exit Sub

' Errors
MqOptionsSendCompToFileError:
MsgBox Error$, vbOKOnly + vbCritical, "MqOptionsSendCompToFile"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub MqOptionsInit()
' Initialize the variables

ierror = False
On Error GoTo MqOptionsInitError

Static lastkev As Single

If MassZedDiffMin! = 0 Then MassZedDiffMin! = 3#
If MassZedDiffMax! = 0 Then MassZedDiffMax! = 0.1

If FilmDensity! = 0# Then FilmDensity! = 2.7
If FilmThickness! = 0# Then FilmThickness! = 10000#

If SubstrateAtomicNumber% = 0 Then SubstrateAtomicNumber% = 14
If SubstrateXrayLine$ = vbNullString Then SubstrateXrayLine$ = "K"

If SubstrateDensity! = 0# Then SubstrateDensity! = 2.7
If SubstrateThickness! = 0# Then SubstrateThickness! = 100

If NumberofTrajectories& = 0 Then NumberofTrajectories& = 10000
If HistogramRange! = 0# Then HistogramRange! = 5#

' Update secondary minimum energy if voltage is changed to suppress
' creation of fast secondaries by default.
If lastkev! <> DefaultKiloVolts! Or SecondaryEnergy! = 0# Then
SecondaryEnergy! = DefaultKiloVolts! / 2#
lastkev! = DefaultKiloVolts!
End If

Exit Sub

' Errors
MqOptionsInitError:
MsgBox Error$, vbOKOnly + vbCritical, "MqOptionsInit"
ierror = True
Exit Sub

End Sub

Sub MQOptionsCalculateAll()
' Create a set of input files for all pure elements based on current conditions

ierror = False
On Error GoTo MqOptionsCalculateAllError

Dim i As Integer
Dim bfilename As String

' Init sample
Call InitSample(MQOptionsTmpSample())
If ierror Then Exit Sub

' Init MQ options
Call MqOptionsInit
If ierror Then Exit Sub

' Specify a single pure element
MQOptionsTmpSample(1).LastElm% = 1
MQOptionsTmpSample(1).LastChan% = 1
MQOptionsTmpSample(1).kilovolts = DefaultKiloVolts!
MQOptionsTmpSample(1).takeoff = DefaultTakeOff!
MQOptionsTmpSample(1).ElmPercents!(1) = 100#  ' pure element (100%)

' Loop on all elements
For i% = 1 To MAXELM%

' Load element and default line
MQOptionsTmpSample(1).Name$ = Trim$(Symup$(i%))
MQOptionsTmpSample(1).Elsyms$(1) = Symlo$(i%)
MQOptionsTmpSample(1).Xrsyms$(1) = Deflin$(i%)

Call ElementLoadArrays(MQOptionsTmpSample())
If ierror Then Exit Sub

Call ElementCheckXray(Int(0), MQOptionsTmpSample())
If ierror Then
ierror = False
msg$ = "Skipping " & Symup$(i%) & " " & Deflin$(i%) & " at " & Str$(DefaultKiloVolts!) & "KeV..."
Call IOWriteLog(msg$)
GoTo 2000:
End If

' Output the input file
If FormMQOPTIONS.Visible Then
bfilename$ = UserDataDirectory$ & "\MQData\" & Trim$(Format$(MQOptionsTmpSample(1).kilovolts!)) & "-" & "ELEMENT.BAT"
Call MqOptionsSendCompToFile(0#, 0#, bfilename$, MQOptionsTmpSample())
If ierror Then Exit Sub
End If

2000:
Next i%

msg$ = "Pure element files created and saved to " & UserDataDirectory$ & "\MQData folder. " & vbCrLf
msg$ = msg$ & "Batch file saved to " & bfilename$ & ". Run this batch file to calculate all "
msg$ = msg$ & "pure element calculations (this could take a while depending on the number of trajectories specified)."
MsgBox msg$, vbOKOnly + vbInformation, "MqOptionsCalculateAll"

Exit Sub

' Errors
MqOptionsCalculateAllError:
MsgBox Error$, vbOKOnly + vbCritical, "MqOptionsCalculateAll"
ierror = True
Exit Sub

End Sub

Sub MQOptionsExtract(tForm As Form)
' Get the pure element folder for data extractions

ierror = False
On Error GoTo MqOptionsExtractError

Dim tfilename As String

msg$ = "Please select a path for the pure element MQ data files"
MsgBox msg$, vbOKOnly + vbInformation, "MqOptionsExtract"

' Get new filename for pure element files
tfilename$ = vbNullString
Call IOGetFileName(Int(2), "DAT", tfilename$, tForm)
If ierror Then Exit Sub

' Save pure element path
ElementPath$ = MiscGetPathOnly$(tfilename$)

Exit Sub

' Errors
MqOptionsExtractError:
MsgBox Error$, vbOKOnly + vbCritical, "MqOptionsExtract"
ierror = True
Exit Sub

End Sub
Sub MQOptionsExtractData(tForm As Form)
' Extract BS and x-ray intensities for a specified MQ output file

ierror = False
On Error GoTo MqOptionsExtractDataError

Dim tfilenumber As Integer

Dim tfilename As String     ' input file
Dim tfilename1 As String    ' input file (compound)
Dim tfilename2 As String    ' input file (pure element)
Dim tfilename3 As String    ' output files

Dim i As Integer, j As Integer

Dim electron(1 To MAXCHAN%) As Single
Dim ratio(1 To MAXCHAN%) As Single
Dim merror(1 To MAXCHAN%) As Single
Dim eerror(1 To MAXCHAN%) As Single
Dim syms(1 To MAXCHAN%) As String

Dim massbse As Single
Dim elecbse As Single
Dim masserror As Single
Dim elecerror As Single

Dim filmnumberofelements As Integer
Dim substratenumberofelements As Integer

Dim tfilmnumberofelements As Integer
Dim tsubstratenumberofelements As Integer

Dim atomicnumber(1 To MAXCHAN%) As Integer
Dim tatomicnumber(1 To MAXCHAN%) As Single
Dim AtomicWeight(1 To MAXCHAN%) As Single
Dim concentration(1 To MAXCHAN%) As Single

Dim backscatter As Single
Dim generated(1 To MAXCHAN%) As Single
Dim emitted(1 To MAXCHAN%) As Single

Dim atomicnumberA(1 To MAXCHAN%) As Integer
Dim atomicweightA(1 To MAXCHAN%) As Single
Dim concentrationA(1 To MAXCHAN%) As Single

Dim backscatterA As Single
Dim generatedA(1 To MAXCHAN%) As Single
Dim emittedA(1 To MAXCHAN%) As Single

Dim atomicnumberB(1 To MAXCHAN%) As Integer
Dim atomicweightB(1 To MAXCHAN%) As Single
Dim concentrationB(1 To MAXCHAN%) As Single

Dim backscatterB As Single
Dim generatedB(1 To MAXCHAN%) As Single
Dim emittedB(1 To MAXCHAN%) As Single

Dim atomicnumberC(1 To MAXCHAN%) As Integer
Dim atomicweightC(1 To MAXCHAN%) As Single
Dim concentrationC(1 To MAXCHAN%) As Single

Dim backscatterC As Single
Dim generatedC(1 To MAXCHAN%) As Single
Dim emittedC(1 To MAXCHAN%) As Single

Dim atomicnumberD(1 To MAXCHAN%) As Integer
Dim atomicweightD(1 To MAXCHAN%) As Single
Dim concentrationD(1 To MAXCHAN%) As Single

Dim backscatterD As Single
Dim generatedD(1 To MAXCHAN%) As Single
Dim emittedD(1 To MAXCHAN%) As Single

Dim columnlabels As String
Dim aexp As Single

Dim maszbar As Single    ' z fraction
Dim zedzbar As Single    ' z fraction

ReDim atemp1(1 To MAXCHAN%) As Integer
ReDim atemp2(1 To MAXCHAN%) As Single

ReDim zedfrac(1 To MAXCHAN%) As Single
ReDim masfrac(1 To MAXCHAN%) As Single
ReDim atmfrac(1 To MAXCHAN%) As Single

ReDim masabars(1 To MAXZBAR%) As Single
ReDim masaexps(1 To MAXZBAR%) As Single
ReDim masfracs(1 To MAXZBAR%, 1 To MAXCHAN%) As Single

ReDim zedzbars(1 To MAXZBAR%) As Single
ReDim zedzexps(1 To MAXZBAR%) As Single
ReDim zedfracs(1 To MAXZBAR%, 1 To MAXCHAN%) As Single

Static initializedbse As Boolean
Static initializedxry As Boolean

If Trim$(ElementPath$) = vbNullString Then
msg$ = "Please select a path for the pure element MQ data files first"
MsgBox msg$, vbOKOnly + vbExclamation, "MqOptionsExtractData"
ierror = True
Exit Sub
End If

' Get compound filenamne
msg$ = "Please select a compound MQ data file (using the same keV conditions as the pure elements!)"
MsgBox msg$, vbOKOnly + vbInformation, "MqOptionsExtractData"

' Get new filename for compound file
Call IOGetFileName(Int(2), "DAT", tfilename1$, tForm)
If ierror Then Exit Sub

' Save default directory (for next call)
UserDataDirectory$ = MiscGetPathOnly2$(tfilename1$)

' Open compound file and read composition
Call MQOptionsExtractDataFile(tfilename1$, filmnumberofelements%, substratenumberofelements%, atomicnumber%(), AtomicWeight!(), concentration!(), backscatter!, generated!(), emitted!())
If ierror Then Exit Sub

' Open individual element files
tfilename$ = MiscGetFileNameOnly$(tfilename1$)

' Check for more than 4 elements in compound
If filmnumberofelements% > MAXOUTPUT% Then
msg$ = "Only properties for the first 4 elements in the compound will be extracted"
MsgBox msg$, vbOKOnly + vbInformation, "MqOptionsExtractData"
End If

If filmnumberofelements% > 0 Then
tfilename2$ = Trim$(Str$(Val(tfilename$))) & "-" & Trim$(Symup$(atomicnumber%(1))) & ".DAT"
Call MQOptionsExtractDataFile(ElementPath$ & tfilename2$, tfilmnumberofelements%, tsubstratenumberofelements%, atomicnumberA%(), atomicweightA!(), concentrationA!(), backscatterA!, generatedA!(), emittedA!())
If ierror Then Exit Sub
End If

If filmnumberofelements% > 1 Then
tfilename2$ = Trim$(Str$(Val(tfilename$))) & "-" & Trim$(Symup$(atomicnumber%(2))) & ".DAT"
Call MQOptionsExtractDataFile(ElementPath$ & tfilename2$, tfilmnumberofelements%, tsubstratenumberofelements%, atomicnumberB%(), atomicweightB!(), concentrationB!(), backscatterB!, generatedB!(), emittedB!())
If ierror Then Exit Sub
End If

If filmnumberofelements% > 2 Then
tfilename2$ = Trim$(Str$(Val(tfilename$))) & "-" & Trim$(Symup$(atomicnumber%(3))) & ".DAT"
Call MQOptionsExtractDataFile(ElementPath$ & tfilename2$, tfilmnumberofelements%, tsubstratenumberofelements%, atomicnumberC%(), atomicweightC!(), concentrationC!(), backscatterC!, generatedC!(), emittedC!())
If ierror Then Exit Sub
End If

If filmnumberofelements% > 3 Then
tfilename2$ = Trim$(Str$(Val(tfilename$))) & "-" & Trim$(Symup$(atomicnumber%(4))) & ".DAT"
Call MQOptionsExtractDataFile(ElementPath$ & tfilename2$, tfilmnumberofelements%, tsubstratenumberofelements%, atomicnumberD%(), atomicweightD!(), concentrationD!(), backscatterD!, generatedD!(), emittedD!())
If ierror Then Exit Sub
End If

' Load single precision atomic weights for conversion routine
For i% = 1 To MAXOUTPUT%
tatomicnumber!(i%) = atomicnumber%(i%)
Next i%

' Calculate atomic fractions
Call ConvertWeightToAtomic(filmnumberofelements%, AtomicWeight!(), concentration!(), atmfrac!())
If ierror Then Exit Sub

' Calculate various mass and zed fractions with range of exponents
aexp! = 0.4
For j% = 1 To MAXZBAR%
aexp! = aexp! + 0.1

' Mass fractions
Call StanFormCalculateZbarFrac(Int(1), MAXOUTPUT%, atmfrac!(), atomicnumber%(), atemp1%(), AtomicWeight!(), aexp!, masfrac!(), maszbar!)
If ierror Then Exit Sub

For i% = 1 To MAXOUTPUT%
masfracs!(j%, i%) = masfrac!(i%)
Next i%

' Zed fractions
Call StanFormCalculateZbarFrac(Int(0), MAXOUTPUT%, atmfrac!(), atomicnumber%(), atomicnumber%(), atemp2!(), aexp!, zedfrac!(), zedzbar!)
If ierror Then Exit Sub

For i% = 1 To MAXOUTPUT%
zedfracs!(j%, i%) = zedfrac!(i%)
Next i%

' Z-bars
masaexps!(j%) = aexp!
zedzexps!(j%) = aexp!
masabars!(j%) = maszbar!
zedzbars!(j%) = zedzbar!
Next j%

' Calculate electron fractions (exponent = 1.0)
Call ConvertWeightToElectron(filmnumberofelements%, tatomicnumber!(), AtomicWeight!(), concentration!(), electron!())
If ierror Then Exit Sub

' Calculate compound backscatter (property averaged)
massbse! = 0#
massbse! = massbse! + concentration!(1) * backscatterA!
massbse! = massbse! + concentration!(2) * backscatterB!
massbse! = massbse! + concentration!(3) * backscatterC!
massbse! = massbse! + concentration!(4) * backscatterD!

elecbse! = 0#
elecbse! = elecbse! + electron!(1) * backscatterA!
elecbse! = elecbse! + electron!(2) * backscatterB!
elecbse! = elecbse! + electron!(3) * backscatterC!
elecbse! = elecbse! + electron!(4) * backscatterD!

masserror! = (massbse! - backscatter!) / backscatter!
elecerror! = (elecbse! - backscatter!) / backscatter!

' Calculate k-ratios
If atomicnumber%(1) <> 0 Then syms$(1) = Symup$(atomicnumber%(1))
If atomicnumber%(2) <> 0 Then syms$(2) = Symup$(atomicnumber%(2))
If atomicnumber%(3) <> 0 Then syms$(3) = Symup$(atomicnumber%(3))
If atomicnumber%(4) <> 0 Then syms$(4) = Symup$(atomicnumber%(4))

If generatedA!(1) <> 0# Then ratio!(1) = generated!(1) / generatedA!(1)
If generatedB!(1) <> 0# Then ratio!(2) = generated!(2) / generatedB!(1)
If generatedC!(1) <> 0# Then ratio!(3) = generated!(3) / generatedC!(1)
If generatedD!(1) <> 0# Then ratio!(4) = generated!(4) / generatedD!(1)

If ratio!(1) <> 0# Then merror!(1) = (concentration!(1) - ratio!(1)) / ratio!(1)
If ratio!(2) <> 0# Then merror!(2) = (concentration!(2) - ratio!(2)) / ratio!(2)
If ratio!(3) <> 0# Then merror!(3) = (concentration!(3) - ratio!(3)) / ratio!(3)
If ratio!(4) <> 0# Then merror!(4) = (concentration!(4) - ratio!(4)) / ratio!(4)

If ratio!(1) <> 0# Then eerror!(1) = (electron!(1) - ratio!(1)) / ratio!(1)
If ratio!(2) <> 0# Then eerror!(2) = (electron!(2) - ratio!(2)) / ratio!(2)
If ratio!(3) <> 0# Then eerror!(3) = (electron!(3) - ratio!(3)) / ratio!(3)
If ratio!(4) <> 0# Then eerror!(4) = (electron!(4) - ratio!(4)) / ratio!(4)

' Create output file name for BSE data
tfilename3$ = MiscGetPathOnly$(tfilename1$) & "extractbse.dat"

' Check if first time and delete if so, then write new column labels
If Not initializedbse Then
If Dir$(tfilename3$) <> vbNullString Then Kill tfilename3$
columnlabels$ = VbDquote$ & "Compound Name" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "A symbols" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "B symbols" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "C symbols" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "D symbols" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "BSE (MQ calc., " & Left$(MiscGetFileNameOnly$(tfilename1$), 2) & " KeV)" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "BSE (mass fraction prediction)" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "BSE (electron fraction prediction)" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "Error (BSE mass fraction prediction)" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "Error (BSE electron fraction prediction)" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "A pure BSE" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "B pure BSE" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "C pure BSE" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "D pure BSE" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "n=.5 mass zbar" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "n=.6 mass zbar" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "n=.7 mass zbar" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "n=.8 mass zbar" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "n=.9 mass zbar" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "n=1.0 mass zbar" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "n=1.1 mass zbar" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "n=1.2 mass zbar" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "n=.5 elec zbar" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "n=.6 elec zbar" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "n=.7 elec zbar" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "n=.8 elec zbar" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "n=.9 elec zbar" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "n=1.0 elec zbar" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "n=1.1 elec zbar" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "n=1.2 elec zbar" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "A mass fraction" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "B mass fraction" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "C mass fraction" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "D mass fraction" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "A electron fraction" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "B electron fraction" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "C electron fraction" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "D electron fraction" & VbDquote$ & VbComma$
Open tfilename3$ For Append As tfilenumber%
Print #tfilenumber%, columnlabels$
Close tfilenumber%
initializedbse = True
End If

' Write extracted (and calculated) data for BSE
Open tfilename3$ For Append As tfilenumber%
tfilename$ = MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(tfilename1$))
Print #tfilenumber%, VbDquote$ & Mid$(tfilename$, 4) & VbDquote$ & VbComma$ _
& syms$(1) & VbComma$ & syms$(2) & VbComma$ & syms$(3) & VbComma$ & syms$(4) & VbComma$ _
& MQOptionsFormat$(backscatter!) & VbComma$ & MQOptionsFormat$(massbse!) & VbComma$ & MQOptionsFormat$(elecbse!) & VbComma$ _
& MQOptionsFormat$(masserror!) & VbComma$ & MQOptionsFormat$(elecerror!) & VbComma$ _
& MQOptionsFormat$(backscatterA!) & VbComma$ & MQOptionsFormat$(backscatterB!) & VbComma$ _
& MQOptionsFormat$(backscatterC!) & VbComma$ & MQOptionsFormat$(backscatterD!) & VbComma$ _
& MQOptionsFormat$(masabars!(1)) & VbComma$ & MQOptionsFormat$(masabars!(2)) & VbComma$ & MQOptionsFormat$(masabars!(3)) & VbComma$ & MQOptionsFormat$(masabars!(4)) & VbComma$ _
& MQOptionsFormat$(masabars!(5)) & VbComma$ & MQOptionsFormat$(masabars!(6)) & VbComma$ & MQOptionsFormat$(masabars!(7)) & VbComma$ & MQOptionsFormat$(masabars!(8)) & VbComma$ _
& MQOptionsFormat$(zedzbars!(1)) & VbComma$ & MQOptionsFormat$(zedzbars!(2)) & VbComma$ & MQOptionsFormat$(zedzbars!(3)) & VbComma$ & MQOptionsFormat$(zedzbars!(4)) & VbComma$ _
& MQOptionsFormat$(zedzbars!(5)) & VbComma$ & MQOptionsFormat$(zedzbars!(6)) & VbComma$ & MQOptionsFormat$(zedzbars!(7)) & VbComma$ & MQOptionsFormat$(zedzbars!(8)) & VbComma$ _
& MQOptionsFormat$(concentration!(1)) & VbComma$ & MQOptionsFormat$(concentration!(2)) & VbComma$ _
& MQOptionsFormat$(concentration!(3)) & VbComma$ & MQOptionsFormat$(concentration!(4)) & VbComma$ _
& MQOptionsFormat$(electron!(1)) & VbComma$ & MQOptionsFormat$(electron!(2)) & VbComma$ _
& MQOptionsFormat$(electron!(3)) & VbComma$ & MQOptionsFormat$(electron!(4)) & VbComma$

Close tfilenumber%

Call IOWriteLog("MQ data from " & tfilename1$ & " extracted to " & tfilename3$)

' Create output file name for xray
tfilename3$ = MiscGetPathOnly$(tfilename1$) & "extractxry.dat"

' Check if first time and delete if so, then write new column labels
If Not initializedxry Then
If Dir$(tfilename3$) <> vbNullString Then Kill tfilename3$
columnlabels$ = VbDquote$ & "Compound Name" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "A symbols" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "B symbols" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "C symbols" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "D symbols" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "A pure (gen.) intensities" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "B pure (gen.) intensities" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "C pure (gen.) intensities" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "D pure (gen.) intensities)" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "A compound (gen.) intensities" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "B compound (gen.) intensities" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "C compound (gen.) intensities" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "D compound (gen.) intensities" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "A k-ratio (gen.)" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "B k-ratio (gen.)" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "C k-ratio (gen.)" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "D k-ratio (gen.)" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "A mass fraction" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "B mass fraction" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "C mass fraction" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "D mass fraction" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "A electron fraction" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "B electron fraction" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "C electron fraction" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "D electron fraction" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "A mass fraction prediction error" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "B mass fraction prediction error" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "C mass fraction prediction error" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "D mass fraction prediction error" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "A electron fraction prediction error" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "B electron fraction prediction error" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "C electron fraction prediction error" & VbDquote$ & VbComma$
columnlabels$ = columnlabels$ & VbDquote$ & "D electron fraction prediction error" & VbDquote$ & VbComma$
Open tfilename3$ For Append As tfilenumber%
Print #tfilenumber%, columnlabels$
Close tfilenumber%
initializedxry = True
End If

' Write extracted (and calculated) data
Open tfilename3$ For Append As tfilenumber%
tfilename$ = MiscGetFileNameOnly$(MiscGetFileNameNoExtension$(tfilename1$))
Print #tfilenumber%, VbDquote$ & Mid$(tfilename$, 4) & VbDquote$ & VbComma$ _
& syms$(1) & VbComma$ & syms$(2) & VbComma$ & syms$(3) & VbComma$ & syms$(4) & VbComma$ _
& MQOptionsFormat$(generatedA!(1)) & VbComma$ & MQOptionsFormat$(generatedB!(1)) & VbComma$ _
& MQOptionsFormat$(generatedC!(1)) & VbComma$ & MQOptionsFormat$(generatedD!(1)) & VbComma$ _
& MQOptionsFormat$(generated!(1)) & VbComma$ & MQOptionsFormat$(generated!(2)) & VbComma$ _
& MQOptionsFormat$(generated!(3)) & VbComma$ & MQOptionsFormat$(generated!(4)) & VbComma$ _
& MQOptionsFormat$(ratio!(1)) & VbComma$ & MQOptionsFormat$(ratio!(2)) & VbComma$ _
& MQOptionsFormat$(ratio!(3)) & VbComma$ & MQOptionsFormat$(ratio!(4)) & VbComma$ _
& MQOptionsFormat$(concentration!(1)) & VbComma$ & MQOptionsFormat$(concentration!(2)) & VbComma$ _
& MQOptionsFormat$(concentration!(3)) & VbComma$ & MQOptionsFormat$(concentration!(4)) & VbComma$ _
& MQOptionsFormat$(electron!(1)) & VbComma$ & MQOptionsFormat$(electron!(2)) & VbComma$ _
& MQOptionsFormat$(electron!(3)) & VbComma$ & MQOptionsFormat$(electron!(4)) & VbComma$ _
& MQOptionsFormat$(merror!(1)) & VbComma$ & MQOptionsFormat$(merror!(2)) & VbComma$ _
& MQOptionsFormat$(merror!(3)) & VbComma$ & MQOptionsFormat$(merror!(4)) & VbComma$ _
& MQOptionsFormat$(eerror!(1)) & VbComma$ & MQOptionsFormat$(eerror!(2)) & VbComma$ _
& MQOptionsFormat$(eerror!(3)) & VbComma$ & MQOptionsFormat$(eerror!(4)) & VbComma$

Close tfilenumber%

Call IOWriteLog("MQ data from " & tfilename1$ & " extracted to " & tfilename3$)
Exit Sub

' Errors
MqOptionsExtractDataError:
Close tfilenumber%
MsgBox Error$, vbOKOnly + vbCritical, "MqOptionsExtractData"
ierror = True
Exit Sub

End Sub

Sub MQOptionsExtractDataFile(tfilename As String, filmnumberofelements As Integer, substratenumberofelements As Integer, atomicnumber() As Integer, AtomicWeight() As Single, concentration() As Single, backscatter As Single, generated() As Single, emitted() As Single)
' Parse the MQ output file

ierror = False
On Error GoTo MqOptionsExtractDataFileError

Dim tfilenumber As Integer
Dim astring As String, bstring As String
Dim i As Integer

' Check for file
If Trim$(tfilename$) = vbNullString Then GoTo MQOptionsExtractDataFileNoName
If Dir$(tfilename$) = vbNullString Then GoTo MQOptionsExtractDataFileNoFile

tfilenumber% = FreeFile()
Open tfilename$ For Input As tfilenumber%

Input #tfilenumber%, astring$    ' read title
Input #tfilenumber%, astring$    ' read unit type
Line Input #tfilenumber%, astring$    ' read number of elements and density (film)
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
filmnumberofelements% = Val(bstring$)

Line Input #tfilenumber%, astring$    ' read film thickness

For i% = 1 To filmnumberofelements%
Line Input #tfilenumber%, astring$    ' read at. #, at wt., conc. and line
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
atomicnumber%(i%) = Val(bstring$)
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
AtomicWeight!(i%) = Val(bstring$)
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
concentration!(i%) = Val(bstring$)
Next i%

Line Input #tfilenumber%, astring$    ' read number of elements and density (substrate)
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
substratenumberofelements% = Val(bstring$)

Line Input #tfilenumber%, astring$    ' read substrate thickness
For i% = 1 To substratenumberofelements%
Line Input #tfilenumber%, astring$    ' read at. #, at wt., conc. and line
Next i%

Line Input #tfilenumber%, astring$    ' read angles
Line Input #tfilenumber%, astring$    ' read random seed
Line Input #tfilenumber%, astring$    ' read KeV
Line Input #tfilenumber%, astring$    ' read angles
Line Input #tfilenumber%, astring$    ' number of trajectories
Line Input #tfilenumber%, astring$    ' read histogram range
Line Input #tfilenumber%, astring$    ' read minimum energy for secondaries

For i% = 1 To filmnumberofelements% + substratenumberofelements%
Line Input #tfilenumber%, astring$    ' read edges
Line Input #tfilenumber%, astring$    ' read edges
Next i%

Line Input #tfilenumber%, astring$    ' read ZAV, AAV
Line Input #tfilenumber%, astring$    ' read Kanaya-Okayama range
Line Input #tfilenumber%, astring$    ' read "ELECTRONS"

Line Input #tfilenumber%, astring$    ' read BSE
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
Call MiscParseStringToString(astring$, bstring$)
If ierror Then Exit Sub
backscatter! = Val(bstring$)

Line Input #tfilenumber%, astring$    ' read transmission
Line Input #tfilenumber%, astring$    ' read average x,y,z
Line Input #tfilenumber%, astring$    ' read secondaries

For i% = 1 To filmnumberofelements% + substratenumberofelements%
Line Input #tfilenumber%, astring$    ' read x-rays
generated!(i%) = Mid$(astring$, 33, 13)
emitted!(i%) = Mid$(astring$, 58, 13)
Next i%

For i% = 1 To filmnumberofelements%
Line Input #tfilenumber%, astring$    ' read phichi
Next i%

Close (tfilenumber%)

Exit Sub

' Errors
MqOptionsExtractDataFileError:
Close (tfilenumber%)
MsgBox Error$, vbOKOnly + vbCritical, "MqOptionsExtractDataFile"
ierror = True
Exit Sub

MQOptionsExtractDataFileNoName:
msg$ = "File name is blank"
MsgBox msg$, vbOKOnly + vbCritical, "MqOptionsExtractDataFile"
ierror = True
Exit Sub

MQOptionsExtractDataFileNoFile:
msg$ = "File " & tfilename$ & " was not found"
MsgBox msg$, vbOKOnly + vbCritical, "MqOptionsExtractDataFile"
ierror = True
Exit Sub

End Sub

Function MQOptionsFormat(temp As Single) As String
' Return a blank string if zero

ierror = False
On Error GoTo MQOptionsFormatError

If temp! = 0# Then
MQOptionsFormat$ = vbNullString
Else
MQOptionsFormat$ = Format$(temp!)
End If

Exit Function

' Errors
MQOptionsFormatError:
MsgBox Error$, vbOKOnly + vbCritical, "MqOptionsFormat"
ierror = True
Exit Function

End Function
