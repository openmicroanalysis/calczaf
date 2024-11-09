Attribute VB_Name = "CodeSECONDARY"
' (c) Copyright 1995-2024 by John J. Donovan
Option Explicit

' Browse image file
Global BoundaryImageFileName As String          ' define as global for code in SecondaryAllLoadImage
Dim BoundaryImageNumber As Integer

' Horizontal filed width for non-graphical methods (in microns)
Dim ImageHFW As Single

' Graphical boundary string
Dim bmsg As String

' Form level variables for a single element
Dim SecondaryFluorescenceFlag As Boolean
Dim SecondaryFluorescenceDistanceMethod As Integer   ' 0 = specified distance, 1 = calculated from boundary (specified using several methods but stored as two X,Y pairs)

' Specified distance
Dim SpecifiedDistance As Single

' Point and angle
Dim XStageCoordinate As Single, YStageCoordinate As Single
Dim BoundaryAngle As Single

' Two points
Dim X1StageCoordinate As Single, Y1StageCoordinate As Single
Dim X2StageCoordinate As Single, Y2StageCoordinate As Single

' Two points (graphical)
Dim XX1StageCoordinate As Single, YY1StageCoordinate As Single
Dim XX2StageCoordinate As Single, YY2StageCoordinate As Single

' Browse k-ratio DAT file
Dim BoundaryKratiosDATFile As String
Dim BoundaryKratiosDATFileLine1 As String    ' characteristic line
Dim BoundaryKratiosDATFileLine2 As String    ' electron energy
Dim BoundaryKratiosDATFileLine3 As String    ' column labels (modified k-ratios.dat file from modified Fanal.exe)

' Sample conditions
Dim Boundary_keV As Single, Boundary_mag As Single, Boundary_scanrota As Single

' Parameters to pass to secondary2.bas
Dim Boundary_X_Pos1 As Single, Boundary_Y_Pos1 As Single    ' start boundary line coordinates (stage units)
Dim Boundary_X_Pos2 As Single, Boundary_Y_Pos2 As Single    ' end boundary line coordinates (stage units)

Dim MatA_String As String, MatB_String As String, MatBStd_String As String

Dim SecondaryMatASample(1 To 1) As TypeSample
Dim SecondaryMatBSample(1 To 1) As TypeSample

Sub SecondaryLoad()
' Load FormSECONDARY

ierror = False
On Error GoTo SecondaryLoadError

Dim i As Integer, ip As Integer
Dim tmsg As String

Static initialized As Boolean

' Initialze module level variables
If Not initialized Then
ImageHFW! = 400#             ' assume 400 um HFW

If SpecifiedDistance! = 0# Then SpecifiedDistance! = 25#    ' assume 25 microns for first point of Wark test data file

' Assume vertical boundary at current position (0 degrees equals vertical and 180 equals horizontal)
If UCase$(app.EXEName) = UCase$("CalcZAF") Then
If Not MiscMotorInBounds(XMotor%, RealTimeMotorPositions!(XMotor%)) Then RealTimeMotorPositions!(XMotor%) = MotLoLimits!(XMotor%) + (MotHiLimits!(XMotor%) - MotLoLimits!(XMotor%)) / 2#
If Not MiscMotorInBounds(YMotor%, RealTimeMotorPositions!(YMotor%)) Then RealTimeMotorPositions!(YMotor%) = MotLoLimits!(YMotor%) + (MotHiLimits!(YMotor%) - MotLoLimits!(YMotor%)) / 2#
XStageCoordinate! = RealTimeMotorPositions!(XMotor%)
YStageCoordinate! = RealTimeMotorPositions!(XMotor%)
BoundaryAngle! = 0#

' Assume vertical boundary at stage center (Y points at +/- 0.8 of HFW
X1StageCoordinate! = XStageCoordinate!
Y1StageCoordinate! = YStageCoordinate! + (ImageHFW! / 2# * 0.8) / MotUnitsToAngstromMicrons!(YMotor%)
X2StageCoordinate! = XStageCoordinate!
Y2StageCoordinate! = YStageCoordinate! - (ImageHFW! / 2# * 0.8) / MotUnitsToAngstromMicrons!(YMotor%)
End If

SecondaryFluorescenceFlag = True    ' always true for CalcZAF (single element)

initialized = True
End If

' Load form
If SecondaryFluorescenceFlag Then
FormSECONDARY.CheckUseSecondaryFluorescenceCorrection.value = vbChecked     ' this control is invisible in CalcZAF
Else
FormSECONDARY.CheckUseSecondaryFluorescenceCorrection.value = vbUnchecked
End If

FormSECONDARY.TextHFW.Text = Format$(ImageHFW!)                        ' in um
FormSECONDARY.TextSpecifiedDistance.Text = Format$(SpecifiedDistance!)  ' in um

FormSECONDARY.TextXStageCoordinate.Text = Format$(XStageCoordinate!)
FormSECONDARY.TextYStageCoordinate.Text = Format$(YStageCoordinate!)
FormSECONDARY.TextBoundaryAngle.Text = Format$(BoundaryAngle!)

FormSECONDARY.TextX1StageCoordinate.Text = Format$(X1StageCoordinate!)
FormSECONDARY.TextY1StageCoordinate.Text = Format$(Y1StageCoordinate!)
FormSECONDARY.TextX2StageCoordinate.Text = Format$(X2StageCoordinate!)
FormSECONDARY.TextY2StageCoordinate.Text = Format$(Y2StageCoordinate!)

' Load default material B
Call StandardGetMDBIndex
If ierror Then Exit Sub

' Check for standard
If NumberOfAvailableStandards% <= 0 Then GoTo SecondaryLoadNoStandards

' Load default Mat B composition
If SecondaryMatBSample(1).LastChan% = 0 Then
ip% = 0
For i% = 1 To NumberOfAvailableStandards%
If MiscStringsAreSimilar("TiO2", StandardIndexNames$(i%)) Then  ' search for TiO2
ip% = StandardIndexNumbers%(i%)
Exit For
End If
Next i%

' Load the current Mat B
If ip% > 0 Then
Call StandardGetMDBStandard(ip%, SecondaryMatBSample())   ' load TiO2
If ierror Then Exit Sub
Else
Call StandardGetMDBStandard(StandardIndexNumbers%(1), SecondaryMatBSample())   ' just load first standard
If ierror Then Exit Sub
End If
End If

' Load distance option
FormSECONDARY.OptionDistanceMethod(SecondaryFluorescenceDistanceMethod%).value = True

' Load k-ratio file if already specified
If BoundaryKratiosDATFile$ <> vbNullString Then
FormSECONDARY.LabelKratiosDATFile.Caption = BoundaryKratiosDATFile$
End If

' If image is specified, go ahead and reload
If SecondaryFluorescenceDistanceMethod% = 3 And BoundaryImageFileName$ <> vbNullString Then
FormSECONDARY.LabelImageBMPFile.Caption = BoundaryImageFileName$
Call SecondaryLoadImage(BoundaryImageFileName$)
If ierror Then Exit Sub
End If

Exit Sub

' Errors
SecondaryLoadError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryLoad"
ierror = True
Exit Sub

SecondaryLoadNoStandards:
msg$ = "No standard compositions were found in the standard database file. please select a valid standard database."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryLoad"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub SecondarySave()
' Save options

ierror = False
On Error GoTo SecondarySaveError

Dim radians As Single

If FormSECONDARY.CheckUseSecondaryFluorescenceCorrection.value = vbChecked Then    ' this control is invisible in CalcZAF
SecondaryFluorescenceFlag = True                            ' module level flag
UseSecondaryBoundaryFluorescenceCorrectionFlag = True       ' set global flag if any element boundary fluorescence flag is true (PFE only)
Else
SecondaryFluorescenceFlag = False
End If

' Save parameters
If Val(FormSECONDARY.TextHFW.Text) <= 0# Or Val(FormSECONDARY.TextHFW.Text) > 10000# Then
msg$ = "Horizontal field width of " & FormSECONDARY.TextHFW.Text & " microns is out of range! (must be greater than 0 and less than 10,000)"
MsgBox msg$, vbOKOnly + vbExclamation, "SecondarySave"
ierror = True
Exit Sub
Else
ImageHFW! = Val(FormSECONDARY.TextHFW.Text)
End If

' Save distance option
SecondaryFluorescenceDistanceMethod% = 0
If FormSECONDARY.OptionDistanceMethod(1).value = True Then
SecondaryFluorescenceDistanceMethod% = 1
ElseIf FormSECONDARY.OptionDistanceMethod(2).value = True Then
SecondaryFluorescenceDistanceMethod% = 2
ElseIf FormSECONDARY.OptionDistanceMethod(3).value = True Then
SecondaryFluorescenceDistanceMethod% = 3
End If

' Specified distance
If SecondaryFluorescenceDistanceMethod% = 0 Then
If Val(FormSECONDARY.TextSpecifiedDistance.Text) <= 0# Or Val(FormSECONDARY.TextSpecifiedDistance.Text) > 10000# Then
msg$ = "Specified distance of " & FormSECONDARY.TextSpecifiedDistance.Text & " microns is out of range! (must be greater than 0 and less than 10,000)"
MsgBox msg$, vbOKOnly + vbExclamation, "SecondarySave"
ierror = True
Exit Sub
Else
SpecifiedDistance! = Val(FormSECONDARY.TextSpecifiedDistance.Text)
End If
End If

' One point and angle
If SecondaryFluorescenceDistanceMethod% = 1 Then
If Val(FormSECONDARY.TextXStageCoordinate.Text) < MotLoLimits!(XMotor%) Or Val(FormSECONDARY.TextXStageCoordinate.Text) > MotHiLimits!(XMotor%) Then
msg$ = "Stage X coordinate of " & FormSECONDARY.TextXStageCoordinate.Text & " is out of range! (must be greater than " & Format$(MotLoLimits!(XMotor%)) & " and less than " & Format$(MotHiLimits!(XMotor%)) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "SecondarySave"
ierror = True
Exit Sub
Else
XStageCoordinate! = Val(FormSECONDARY.TextXStageCoordinate.Text)
End If

If Val(FormSECONDARY.TextYStageCoordinate.Text) < MotLoLimits!(YMotor%) Or Val(FormSECONDARY.TextYStageCoordinate.Text) > MotHiLimits!(YMotor%) Then
msg$ = "Stage Y coordinate of " & FormSECONDARY.TextYStageCoordinate.Text & " is out of range! (must be greater than " & Format$(MotLoLimits!(YMotor%)) & " and less than " & Format$(MotHiLimits!(YMotor%)) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "SecondarySave"
ierror = True
Exit Sub
Else
YStageCoordinate! = Val(FormSECONDARY.TextYStageCoordinate.Text)
End If

If Val(FormSECONDARY.TextBoundaryAngle.Text) < 0# Or Val(FormSECONDARY.TextBoundaryAngle.Text) > 180# Then
msg$ = FormSECONDARY.TextBoundaryAngle.Text & " degrees is out of range! (must be greater than 0 and less than 180)"
MsgBox msg$, vbOKOnly + vbExclamation, "SecondarySave"
ierror = True
Exit Sub
Else
BoundaryAngle! = Val(FormSECONDARY.TextBoundaryAngle.Text)
End If
End If

' Two points
If SecondaryFluorescenceDistanceMethod% = 2 Then
If Val(FormSECONDARY.TextX1StageCoordinate.Text) < MotLoLimits!(XMotor%) Or Val(FormSECONDARY.TextX1StageCoordinate.Text) > MotHiLimits!(XMotor%) Then
msg$ = "Stage X1 coordinate of " & FormSECONDARY.TextX1StageCoordinate.Text & " is out of range! (must be greater than " & Format$(MotLoLimits!(XMotor%)) & " and less than " & Format$(MotHiLimits!(XMotor%)) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "SecondarySave"
ierror = True
Exit Sub
Else
X1StageCoordinate! = Val(FormSECONDARY.TextX1StageCoordinate.Text)
End If

If Val(FormSECONDARY.TextY1StageCoordinate.Text) < MotLoLimits!(YMotor%) Or Val(FormSECONDARY.TextY1StageCoordinate.Text) > MotHiLimits!(YMotor%) Then
msg$ = "Stage Y1 coordinate of " & FormSECONDARY.TextY1StageCoordinate.Text & " is out of range! (must be greater than " & Format$(MotLoLimits!(YMotor%)) & " and less than " & Format$(MotHiLimits!(YMotor%)) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "SecondarySave"
ierror = True
Exit Sub
Else
Y1StageCoordinate! = Val(FormSECONDARY.TextY1StageCoordinate.Text)
End If

If Val(FormSECONDARY.TextX2StageCoordinate.Text) < MotLoLimits!(XMotor%) Or Val(FormSECONDARY.TextX2StageCoordinate.Text) > MotHiLimits!(XMotor%) Then
msg$ = "Stage X2 coordinate of " & FormSECONDARY.TextX2StageCoordinate.Text & " is out of range! (must be greater than " & Format$(MotLoLimits!(XMotor%)) & " and less than " & Format$(MotHiLimits!(XMotor%)) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "SecondarySave"
ierror = True
Exit Sub
Else
X2StageCoordinate! = Val(FormSECONDARY.TextX2StageCoordinate.Text)
End If

If Val(FormSECONDARY.TextY2StageCoordinate.Text) < MotLoLimits!(YMotor%) Or Val(FormSECONDARY.TextY2StageCoordinate.Text) > MotHiLimits!(YMotor%) Then
msg$ = "Stage Y2 coordinate of " & FormSECONDARY.TextY2StageCoordinate.Text & " is out of range! (must be greater than " & Format$(MotLoLimits!(YMotor%)) & " and less than " & Format$(MotHiLimits!(YMotor%)) & ")"
MsgBox msg$, vbOKOnly + vbExclamation, "SecondarySave"
ierror = True
Exit Sub
Else
Y2StageCoordinate! = Val(FormSECONDARY.TextY2StageCoordinate.Text)
End If
End If

' Graphical method is saved during mouse up/mouse down events (just check for inbounds below)

' Now calculate the 2 point boundary coordinates for distance calculation
If SecondaryFluorescenceDistanceMethod% = 0 Then
' Specified distance is calculated based on sample coordinate in SecondaryInitLine which is called by CalcZAFCalculate or other analysis code

' 1 point and angle (assume +/- 50 um)
ElseIf SecondaryFluorescenceDistanceMethod% = 1 Then
radians! = BoundaryAngle! * PI! / 180#
Boundary_X_Pos1! = XStageCoordinate! + Sin(radians!) * (ImageHFW! / 2# * 0.8) / MotUnitsToAngstromMicrons!(XMotor%)
Boundary_Y_Pos1! = YStageCoordinate! + Cos(radians!) * (ImageHFW! / 2# * 0.8) / MotUnitsToAngstromMicrons!(YMotor%)

Boundary_X_Pos2! = XStageCoordinate! - Sin(radians!) * (ImageHFW! / 2# * 0.8) / MotUnitsToAngstromMicrons!(XMotor%)
Boundary_Y_Pos2! = YStageCoordinate! - Cos(radians!) * (ImageHFW! / 2# * 0.8) / MotUnitsToAngstromMicrons!(YMotor%)

' 2 points (just load form coordinates)
ElseIf SecondaryFluorescenceDistanceMethod% = 2 Then
Boundary_X_Pos1! = X1StageCoordinate!
Boundary_Y_Pos1! = Y1StageCoordinate!

Boundary_X_Pos2! = X2StageCoordinate!
Boundary_Y_Pos2! = Y2StageCoordinate!

' Graphical
ElseIf SecondaryFluorescenceDistanceMethod% = 3 Then
If XX1StageCoordinate! < MotLoLimits!(XMotor%) Or XX1StageCoordinate! > MotHiLimits!(XMotor%) Then GoTo SecondarySaveBadGraphicalBoundary
If YY1StageCoordinate! < MotLoLimits!(YMotor%) Or YY1StageCoordinate! > MotHiLimits!(YMotor%) Then GoTo SecondarySaveBadGraphicalBoundary
If XX2StageCoordinate! < MotLoLimits!(XMotor%) Or XX2StageCoordinate! > MotHiLimits!(XMotor%) Then GoTo SecondarySaveBadGraphicalBoundary
If YY2StageCoordinate! < MotLoLimits!(YMotor%) Or YY2StageCoordinate! > MotHiLimits!(YMotor%) Then GoTo SecondarySaveBadGraphicalBoundary

Boundary_X_Pos1! = XX1StageCoordinate!
Boundary_Y_Pos1! = YY1StageCoordinate!

Boundary_X_Pos2! = XX2StageCoordinate!
Boundary_Y_Pos2! = YY2StageCoordinate!
End If

' Save k-ratio file
BoundaryKratiosDATFile$ = Trim$(FormSECONDARY.LabelKratiosDATFile.Caption)

' Save image file name
BoundaryImageFileName$ = Trim$(FormSECONDARY.LabelImageBMPFile.Caption)

Exit Sub

' Errors
SecondarySaveError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondarySave"
ierror = True
Exit Sub

SecondarySaveBadGraphicalBoundary:
msg$ = "One or more graphical boundary coordinates are out of stage X or Y limits! Be sure that you have properly specified a boundary by clicking and dragging the mouse on the loaded image."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondarySave"
ierror = True
Exit Sub

End Sub

Sub SecondaryBrowseFile(mode As Integer, tForm As Form)
' Browse to a file for secondary calculations
'  mode = 0 browse for k-ratio .DAT file
'  mode = 1 browse for image .BMP/JPG/GIF file

ierror = False
On Error GoTo SecondaryBrowseFileError

Dim tfilename As String, ioextension As String, tfilename2 As String
Dim tfolder As String, astring As String, bstring As String
Dim takeoff As Single, keV As Single, dval As Single
Dim eval As Integer, xval As Integer
Dim esym As String, xsym As String

' K-ratio file
If mode% = 0 Then
If Trim$(BoundaryKratiosDATFile$) = vbNullString Then BoundaryKratiosDATFile$ = PENEPMA_Root$ & "\Fanal\Couple\" & "k-ratios.dat"
tfilename$ = BoundaryKratiosDATFile$
ioextension$ = "DAT"
Call IOGetFileName(Int(2), ioextension$, tfilename$, tForm)
If ierror Then Exit Sub

' Check that file is k-ratios.dat, not k-ratios2.dat
tfolder$ = MiscGetPathOnly$(tfilename$)
If Not MiscStringsAreSame("k-ratios.dat", MiscGetFileNameOnly$(tfilename$)) Then GoTo SecondaryBrowseFileNotKRATIOSDAT

' Confirm with user based on fanal.txt (as of 10/05/2014)
tfilename2$ = tfolder$ & "fanal.txt"
If Dir$(tfilename2$) = vbNullString Then GoTo SecondaryBrowseFileFANALTXTNotFound

Open tfilename2$ For Input As Temp2FileNumber%
Input #Temp2FileNumber%, MatA_String$
Input #Temp2FileNumber%, MatB_String$
Input #Temp2FileNumber%, MatBStd_String$
Input #Temp2FileNumber%, takeoff!
Input #Temp2FileNumber%, keV!

Input #Temp2FileNumber%, eval%
If eval% < 1 Or eval% > MAXELM% Then GoTo SecondaryBrowseFileElementNotFound
esym$ = Symup$(eval%)

Input #Temp2FileNumber%, xval%
If xval% < 1 Or xval% > MAXRAY% - 1 Then GoTo SecondaryBrowseFileXrayNotFound
xsym$ = Xraylo$(xval%)

Close #Temp2FileNumber%
DoEvents

' Read the first three lines of k-ratios.dat and store in module level string variables
Open tfilename$ For Input As #Temp2FileNumber%

' Read characteristic line
Line Input #Temp2FileNumber%, BoundaryKratiosDATFileLine1$

' Read electron energy line
Line Input #Temp2FileNumber%, BoundaryKratiosDATFileLine2$

' Read column labels line
Line Input #Temp2FileNumber%, BoundaryKratiosDATFileLine3$

' Now determine last distance value
Do Until EOF(Temp2FileNumber%)
Line Input #Temp2FileNumber%, astring$
Loop

Close #Temp2FileNumber%

' Extract last distance
Call MiscParseStringToStringA(astring$, " ", bstring$)
If ierror Then Exit Sub
Call MiscParseStringToStringA(astring$, " ", bstring$)
If ierror Then Exit Sub
dval! = Val(bstring$)

msg$ = "K-ratio boundary fluorescence data was found for " & esym$ & " " & xsym$ & " in " & MatA_String$ & " adjacent to " & MatB_String$ & " at " & Format$(takeoff!) & " deg, " & Format$(keV!) & " keV, using " & MatBStd_String$ & " as the primary standard." & vbCrLf & vbCrLf
msg$ = msg$ & "The total secondary fluorescence distance modeled was " & Format$(dval!) & " microns."
MsgBox msg$, vbOKOnly + vbInformation, "SecondaryBrowseFile"

BoundaryKratiosDATFile$ = tfilename$
FormSECONDARY.LabelKratiosDATFile.Caption = BoundaryKratiosDATFile$
End If

' Image file
If mode% = 1 Then
tfilename$ = BoundaryImageFileName$
ioextension$ = "IMG"
tfilename$ = MiscGetFileNameNoExtension$(tfilename$)    ' remove extension so all image files are visible
Call IOGetFileName(Int(2), ioextension$, tfilename$, tForm)
If ierror Then Exit Sub

' Check that image file has an ACQ stage calibration associated with it
If Dir$(MiscGetFileNameNoExtension$(tfilename$) & ".acq") = vbNullString Then GoTo SecondaryBrowseFileNotCalibrated

' Load the image
Call SecondaryLoadImage(tfilename$)
If ierror Then Exit Sub

BoundaryImageNumber% = 0        ' re-set boundary image number to idicate that user changed boundary image file
BoundaryImageFileName$ = tfilename$
FormSECONDARY.LabelImageBMPFile.Caption = BoundaryImageFileName$
End If

Exit Sub

' Errors
SecondaryBrowseFileError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryBrowseFile"
Close #Temp2FileNumber%
ierror = True
Exit Sub

SecondaryBrowseFileNotKRATIOSDAT:
msg$ = "The specified kratio file (" & MiscGetFileNameOnly$(tfilename$) & ") is not " & VbDquote$ & "k-ratios.dat" & VbDquote$ & ". Please choose a k-ratios.dat file for secondary fluorescence corrections."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryBrowseFile"
ierror = True
Exit Sub

SecondaryBrowseFileFANALTXTNotFound:
msg$ = "The specified couple folder (" & MiscGetPathOnly2$(tfilename$) & ") does not contain a FANAL.TXT file." & vbCrLf & vbCrLf
msg$ = msg$ & "Please re-calculate the necessary secondary fluorescence couple in Standard.exe and try the secondary fluorescence correction again in CalcZAF."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryBrowseFile"
ierror = True
Exit Sub

SecondaryBrowseFileElementNotFound:
msg$ = "The folder (" & MiscGetPathOnly2$(tfilename$) & ") contains an extracted element atomic number " & Format$(eval%) & ", that is not a valid atomic number. Please choose a valid Fanal/Couple folder for the k-ratio data file." & vbCrLf & vbCrLf
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryBrowseFile"
ierror = True
Exit Sub

SecondaryBrowseFileXrayNotFound:
msg$ = "The folder (" & MiscGetPathOnly2$(tfilename$) & ") contains an extracted element x-ray number " & Format$(xval%) & ", that is not a valid x-ray number. Please choose a valid Fanal/Couple folder for the k-ratio data file."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryBrowseFile"
ierror = True
Exit Sub

SecondaryBrowseFileNotCalibrated:
msg$ = "The specified image file, " & tfilename$ & " is not associated with a stage calibration (.ACQ) file. Please select an image file that is stage calibrated."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryBrowseFile"
ierror = True
Exit Sub

End Sub

Sub SecondaryInit(sample() As TypeSample)
' Initialize sample variables for the SF calculation

ierror = False
On Error GoTo SecondaryInitError

' Make sure k-ratio variables are initialized
Call SecondaryInit2
If ierror Then Exit Sub

Exit Sub

' Errors
SecondaryInitError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryInit"
ierror = True
Exit Sub

End Sub

Sub SecondaryInitLine(sampleline As Integer, sample() As TypeSample)
' Initialize sample line variables for the CalcZAF/Probe for EPMA calculation

ierror = False
On Error GoTo SecondaryInitLineError

' Fixed specified distance (generate boundary coordinates based on fixed distance)
If SecondaryFluorescenceDistanceMethod% = 0 Then
Boundary_X_Pos1! = sample(1).StagePositions!(sampleline%, 1) + SpecifiedDistance! / MotUnitsToAngstromMicrons!(XMotor%)
Boundary_Y_Pos1! = sample(1).StagePositions!(sampleline%, 2) + SpecifiedDistance! / MotUnitsToAngstromMicrons!(YMotor%)

Boundary_X_Pos2! = sample(1).StagePositions!(sampleline%, 1) + SpecifiedDistance! / MotUnitsToAngstromMicrons!(XMotor%)
Boundary_Y_Pos2! = sample(1).StagePositions!(sampleline%, 2) - SpecifiedDistance! / MotUnitsToAngstromMicrons!(YMotor%)
End If

' Calculate distance from boundary line for this stage coordinate
Call SecondaryCalculateDistance(Boundary_X_Pos1!, Boundary_Y_Pos1!, Boundary_X_Pos2, Boundary_Y_Pos2, sampleline%, sample())
If ierror Then Exit Sub

Exit Sub

' Errors
SecondaryInitLineError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryInitLine"
ierror = True
Exit Sub

End Sub

Sub SecondaryInitChan(chan As Integer, sample() As TypeSample)
' Load the k-ratios for the specified channel from the k-ratio DAT file

ierror = False
On Error GoTo SecondaryInitChanError

' Check for valid k-ratio file
If Trim$(BoundaryKratiosDATFile$) = vbNullString Then GoTo SecondaryInitChanNoFilename

' Read k-ratio values (and set sample flags) (CalcZAF only works with a single Kratio file for one element)
Call SecondaryReadKratiosDATFile(BoundaryKratiosDATFile$, sample())
If ierror Then Exit Sub

' Process the k-ratio values just read in
Call SecondaryProcessKratiosDAT(chan%)
If ierror Then Exit Sub

Exit Sub

' Errors
SecondaryInitChanError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryInitChan"
ierror = True
Exit Sub

SecondaryInitChanNoFilename:
msg$ = "No k-ratio data file was specified. Please browse for an appropriate k-ratio couple data file for secondary boundary fluorescence corrections."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryInitChan"
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub SecondaryLoadImage(tfilename As String)
' Load the image into image control (and picturebox for printing)

ierror = False
On Error GoTo SecondaryLoadImageError

Dim hfw As Single, mag As Single

Static OriginalImageHeight As Single

' Load initial image height (square image control)
If OriginalImageHeight! = 0# Then OriginalImageHeight! = FormSECONDARY.Image1.Height

' Check for filename
If Trim$(tfilename$) = vbNullString Then GoTo SecondaryLoadImageNoFile
If Dir$(tfilename$) = vbNullString Then GoTo SecondaryLoadImageNotFound

Screen.MousePointer = vbHourglass
Set FormSECONDARY.Image1 = LoadPicture(tfilename$)
Set FormSECONDARY.Picture1 = LoadPicture(tfilename$)        ' for printing image and graphics to printer
Set FormSECONDARY.Picture2 = LoadPicture(tfilename$)        ' for PictureSnapConvert and Copy To Clipboard (with graphics objects)
Screen.MousePointer = vbDefault

' Rescale form to image aspect
If FormSECONDARY.Image1.Picture.Type > 0 Then   ' bitmap
If FormSECONDARY.Image1.Picture.Width <> 0# Then
FormSECONDARY.Image1.Height = OriginalImageHeight! * FormSECONDARY.Image1.Picture.Height / FormSECONDARY.Image1.Picture.Width
End If
End If

' Load the image in PictureSnap for proper form calibration
Call PictureSnapFileOpen(Int(0), tfilename$, FormSECONDARY)
If ierror Then Exit Sub

' Load the text field for horizontal field width
Call PictureSnapCalculateHFW(hfw!, mag!)
If ierror Then Exit Sub
FormSECONDARY.TextHFW.Text = Format$(hfw!)

PictureSnapFilename$ = tfilename$
PictureSnapCalibrated = True
Exit Sub

' Errors
SecondaryLoadImageError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryLoadImage"
ierror = True
Exit Sub

SecondaryLoadImageNoFile:
msg$ = "No image file was specified."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryLoadImage"
ierror = True
Exit Sub

SecondaryLoadImageNotFound:
msg$ = "Image file " & tfilename$ & " was not found."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryLoadImage"
ierror = True
Exit Sub

End Sub

Sub SecondaryLoadBoundary()
' Loads all distance method controls based on stored 2 point boundary

ierror = False
On Error GoTo SecondaryLoadBoundaryError

Dim radians As Double
Dim tadjacent As Double, topposite As Double

' Calculate opposite and adjacent lengths
tadjacent# = Boundary_X_Pos2! - Boundary_X_Pos1!
topposite# = Boundary_Y_Pos2! - Boundary_Y_Pos1!

' Check for vertical boundary
If tadjacent# <> 0# Then

' Calculate angle from two points
radians# = topposite# / tadjacent#
BoundaryAngle! = (90# - Atn(radians#) * 180# / PID#)

' Vertical boundary is zero angle
Else
BoundaryAngle! = 0#
End If

XStageCoordinate! = Boundary_X_Pos1! + (Boundary_X_Pos2! - Boundary_X_Pos1!) / 2#
YStageCoordinate! = Boundary_Y_Pos1! + (Boundary_Y_Pos2! - Boundary_Y_Pos1!) / 2#

' 2 points (just load form coordinates)
X1StageCoordinate! = Boundary_X_Pos1!
Y1StageCoordinate! = Boundary_Y_Pos1!

X2StageCoordinate! = Boundary_X_Pos2!
Y2StageCoordinate! = Boundary_Y_Pos2!

' Graphical
XX1StageCoordinate! = Boundary_X_Pos1!
YY1StageCoordinate! = Boundary_Y_Pos1!

XX2StageCoordinate! = Boundary_X_Pos2!
YY2StageCoordinate! = Boundary_Y_Pos2!

Exit Sub

' Errors
SecondaryLoadBoundaryError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryLoadBoundary"
ierror = True
Exit Sub

End Sub

Sub SecondaryUpdateCursor(tX As Single, tY As Single, tForm As Form)
' Updates the stage cursor

ierror = False
On Error GoTo SecondaryUpdateCursorError

Dim mode As Integer
Dim sx As Single, sy As Single, sz As Single
Dim fractionx As Single, fractiony As Single

' Get current distance mode
mode% = 0
If FormSECONDARY.OptionDistanceMethod(1).value = True Then
mode% = 1
ElseIf FormSECONDARY.OptionDistanceMethod(2).value = True Then
mode% = 2
ElseIf FormSECONDARY.OptionDistanceMethod(3).value = True Then
mode% = 3
End If

' Check for proper mode
If mode% = 0 Then Exit Sub
If mode% = 3 And Trim$(BoundaryImageFileName$) = vbNullString Then Exit Sub

' If not calibrated just calculate pixels and exit
If Not PictureSnapCalibrated Then
sx! = tX! / Screen.TwipsPerPixelX
sy! = tY! / Screen.TwipsPerPixelY

' Convert to stage coordinates and back to form coordinates to get fractional distance
Else

' Convert from Image1 control "stretched" units to FormPICTURESNAP.Picture2 "unstretched" units
Call PictureSnapUnStretch(Int(0), tX!, tY!, tForm.Image1)
If ierror Then Exit Sub

Call PictureSnapConvert(Int(1), tX!, tY!, CSng(0#), sx!, sy!, sz!, fractionx!, fractiony!)
If ierror Then Exit Sub
End If

FormSECONDARY.LabelCursorPosition.Caption = MiscAutoFormat$(sx!) & ", " & MiscAutoFormat$(sy!)
Exit Sub

' Errors
SecondaryUpdateCursorError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryUpdateCursor"
ierror = True
Exit Sub

End Sub

Sub SecondaryGetBoundary(mode As Integer, X1 As Single, Y1 As Single, x2 As Single, y2 As Single, tForm As Form)
' Get the boundary coordinates from the user's click and drag amd convert to stage positions and store
' mode = 0 first boundary points (mouse down)
' mode = 1 second boundary points (mouse up)

ierror = False
On Error GoTo SecondaryGetBoundaryError

Dim dmode As Integer

Static scx1 As Single, scy1 As Single
Static scx2 As Single, scy2 As Single

Static stx1 As Single, sty1 As Single, stz1 As Single
Static stx2 As Single, sty2 As Single, stz2 As Single

Dim fractionx1 As Single, fractiony1 As Single
Dim fractionx2 As Single, fractiony2 As Single

' Check if proper mode and image is loaded
dmode% = 0
If FormSECONDARY.OptionDistanceMethod(1).value = True Then
dmode% = 1
ElseIf FormSECONDARY.OptionDistanceMethod(2).value = True Then
dmode% = 2
ElseIf FormSECONDARY.OptionDistanceMethod(3).value = True Then
dmode% = 3
End If

If dmode% <> 3 Or Trim$(BoundaryImageFileName$) = vbNullString Then Exit Sub

' Store
If mode% = 1 Then
scx1! = X1!
scy1! = Y1!
Exit Sub
End If

scx2! = x2!
scy2! = y2!

' Check if zero distance (double click)
If scx1! = scx2! And scy1! = scy2! Then GoTo SecondaryGetBoundaryPointsSame

If VerboseMode Then
msg$ = "X1,Y1= " & Format$(scx1!) & "," & Format$(scy1!) & ", X2,Y2= " & Format$(scx2!) & "," & Format$(scy2!)
Call IOWriteLog("SecondaryGetBoundary: Control units: " & msg$)
End If

' Convert from "stretched" units to "unstretched" units
Call PictureSnapUnStretch(Int(0), scx1!, scy1!, tForm.Image1)
If ierror Then Exit Sub

' Convert from "stretched" units to "unstretched" units
Call PictureSnapUnStretch(Int(0), scx2!, scy2!, tForm.Image1)
If ierror Then Exit Sub

If VerboseMode Then
msg$ = "X1,Y1= " & Format$(scx1!) & "," & Format$(scy1!) & ", X2,Y2= " & Format$(scx2!) & "," & Format$(scy2!)
Call IOWriteLog("SecondaryGetBoundary: UnStretched units: " & msg$)
End If

' Convert to stage coordinates
Call PictureSnapConvert(Int(1), scx1!, scy1!, CSng(0#), stx1!, sty1!, stz1!, fractionx1!, fractiony1!)
If ierror Then Exit Sub
Call PictureSnapConvert(Int(1), scx2!, scy2!, CSng(0#), stx2!, sty2!, stz2!, fractionx2!, fractiony2!)
If ierror Then Exit Sub

' Load coordinates to label
bmsg$ = "X1,Y1= " & MiscAutoFormat$(stx1!) & "," & MiscAutoFormat$(sty1!) & ", X2,Y2= " & MiscAutoFormat$(stx2!) & "," & MiscAutoFormat$(sty2!)
FormSECONDARY.LabelBoundaryCoordinates.Caption = bmsg$

If VerboseMode Then
Call IOWriteLog("SecondaryGetBoundary: Stage units: " & bmsg$)
End If

' Draw the boundary using stage coordinates
Call SecondaryDrawBoundary(stx1!, sty1!, stx2!, sty2!, tForm)
If ierror Then Exit Sub

' Save graphical boundary stage coordinates to module level
XX1StageCoordinate! = stx1!
YY1StageCoordinate! = sty1!

XX2StageCoordinate! = stx2!
YY2StageCoordinate! = sty2!

Exit Sub

' Errors
SecondaryGetBoundaryError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryGetBoundary"
ierror = True
Exit Sub

SecondaryGetBoundaryPointsSame:
msg$ = "The specified boundary coordinates are the same. Please specify two discrete points to define the boundary (use mouse click and drag if using graphical method)."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryGetBoundary"
ierror = True
Exit Sub

End Sub

Sub SecondaryDrawBoundary(stx1 As Single, sty1 As Single, stx2 As Single, sty2 As Single, tForm As Form)
' Draw boundary on image using passed stage coordinates

ierror = False
On Error GoTo SecondaryDrawBoundaryError

Dim tcolor As Long

Dim scx1 As Single, scy1 As Single, scz1 As Single
Dim scx2 As Single, scy2 As Single, scz2 As Single
Dim fractionx1 As Single, fractiony1 As Single
Dim fractionx2 As Single, fractiony2 As Single

' First reload the image to clean up previously drawn boundaries
tForm.Refresh
tForm.Image1.Refresh
tForm.Picture1.Picture = tForm.Image1       ' to clear previous boundary (refresh doesn't work!)

' Check if a calibration is loaded
If Not PictureSnapCalibrated Then Exit Sub

' Display items for debugging
If DebugMode And VerboseMode Then
msg$ = "X1,Y1= " & MiscAutoFormat$(stx1!) & "," & MiscAutoFormat$(sty1!) & ", X2,Y2= " & MiscAutoFormat$(stx2!) & "," & MiscAutoFormat$(sty2!)
Call IOWriteLog(vbCrLf & "SecondaryDrawBoundary: Stage units: " & msg$)
End If

If DebugMode And VerboseMode Then
msg$ = "Left,Top= " & Format$(tForm.Image1.Left) & "," & Format$(tForm.Image1.Top) & ", Width,Height= " & Format$(tForm.Image1.Width) & "," & Format$(tForm.Image1.Height)
Call IOWriteLog("SecondaryDrawBoundary: Image1 (FormUnits): " & msg$)
End If

' Convert to image coordinates
Call PictureSnapConvert(Int(2), scx1!, scy1!, scz1!, stx1!, sty1!, CSng(0#), fractionx1!, fractiony1!)
If ierror Then Exit Sub
Call PictureSnapConvert(Int(2), scx2!, scy2!, scz2!, stx2!, sty2!, CSng(0#), fractionx2!, fractiony2!)
If ierror Then Exit Sub

If DebugMode And VerboseMode Then
msg$ = "X1,Y1= " & Format$(scx1!) & "," & Format$(scy1!) & ", X2,Y2= " & Format$(scx2!) & "," & Format$(scy2!)
Call IOWriteLog("SecondaryDrawBoundary: Image units: " & msg$)
End If

' Convert from "unstretched" units to "stretched" units
Call PictureSnapUnStretch(Int(1), scx1!, scy1!, tForm.Image1)
If ierror Then Exit Sub

Call PictureSnapUnStretch(Int(1), scx2!, scy2!, tForm.Image1)
If ierror Then Exit Sub

If DebugMode And VerboseMode Then
msg$ = "X1,Y1= " & Format$(scx1!) & "," & Format$(scy1!) & ", X2,Y2= " & Format$(scx2!) & "," & Format$(scy2!)
Call IOWriteLog("SecondaryDrawBoundary: Stretched units: " & msg$)
End If

If DebugMode And VerboseMode Then
msg$ = "Left,Top= " & Format$(tForm.ScaleLeft) & "," & Format$(tForm.ScaleTop) & ", Width,Height= " & Format$(tForm.ScaleWidth) & "," & Format$(tForm.ScaleHeight)
Call IOWriteLog("SecondaryDrawBoundary: tForm (FormUnits): " & msg$)
End If

' Adjust for form coordinate scaling
scx1! = tForm.Image1.Left + scx1!
scy1! = tForm.Image1.Top + scy1!
scx2! = tForm.Image1.Left + scx2!
scy2! = tForm.Image1.Top + scy2!

If DebugMode And VerboseMode Then
msg$ = "X1,Y1= " & Format$(scx1!) & "," & Format$(scy1!) & ", X2,Y2= " & Format$(scx2!) & "," & Format$(scy2!)
Call IOWriteLog("SecondaryDrawBoundary: Offset Screen : " & msg$)
End If

' Draw boundary line on form
tcolor& = RGB(255, 0, 0)
tForm.DrawWidth = 2
tForm.Line (scx1!, scy1!)-(scx2!, scy2!), tcolor&

' Draw on FormSECONDARY.Picture2 for clipboard
tForm.Picture2.DrawWidth = 2
scx1! = tForm.Picture2.ScaleWidth * fractionx1!
scy1! = tForm.Picture2.ScaleHeight * fractiony1!
scx2! = tForm.Picture2.ScaleWidth * fractionx2!
scy2! = tForm.Picture2.ScaleHeight * fractiony2!
tForm.Picture2.Line (scx1!, scy1!)-(scx2!, scy2!), tcolor&

' Draw also on FormSECONDARY.Picture1 for printer
tForm.Picture1.DrawWidth = 2
scx1! = tForm.Picture1.ScaleWidth * fractionx1!
scy1! = tForm.Picture1.ScaleHeight * fractiony1!
scx2! = tForm.Picture1.ScaleWidth * fractionx2!
scy2! = tForm.Picture1.ScaleHeight * fractiony2!
tForm.Picture1.Line (scx1!, scy1!)-(scx2!, scy2!), tcolor&

Exit Sub

' Errors
SecondaryDrawBoundaryError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryDrawBoundary"
ierror = True
Exit Sub

End Sub

Sub SecondaryUpdateBoundary()
' Redraw the boundary for the indicated distance method

ierror = False
On Error GoTo SecondaryUpdateBoundaryError

Dim dmode As Integer
Dim radians As Single

' Dimension coordinates (assume using three points and XYZ)
Dim cpoint1(1 To 3) As Single, cpoint2(1 To 3) As Single, cpoint3(1 To 3) As Single
Dim apoint1(1 To 3) As Single, apoint2(1 To 3) As Single, apoint3(1 To 3) As Single

' Get current distance mode
dmode% = 0
If FormSECONDARY.OptionDistanceMethod(1).value = True Then
dmode% = 1
ElseIf FormSECONDARY.OptionDistanceMethod(2).value = True Then
dmode% = 2
ElseIf FormSECONDARY.OptionDistanceMethod(3).value = True Then
dmode% = 3
End If

' Save the horizontal field width to module level
If Val(FormSECONDARY.TextHFW.Text) > 0# And Val(FormSECONDARY.TextHFW.Text) <= 10000# Then
ImageHFW! = Val(FormSECONDARY.TextHFW.Text)
End If

' Specified distance (nothing to do)
If dmode% = 0 Then
FormSECONDARY.LabelCursorPosition.Caption = vbNullString
FormSECONDARY.Refresh

' Point and angle
ElseIf dmode% = 1 Then
If Val(FormSECONDARY.TextBoundaryAngle.Text) >= 0 And Val(FormSECONDARY.TextBoundaryAngle.Text) <= 180 Then
radians! = Val(FormSECONDARY.TextBoundaryAngle.Text) * PI! / 180#

If Val(FormSECONDARY.TextXStageCoordinate.Text) >= MotLoLimits!(XMotor%) And Val(FormSECONDARY.TextXStageCoordinate.Text) <= MotHiLimits!(XMotor%) Then
If Val(FormSECONDARY.TextYStageCoordinate.Text) >= MotLoLimits!(YMotor%) And Val(FormSECONDARY.TextYStageCoordinate.Text) <= MotHiLimits!(YMotor%) Then
Boundary_X_Pos1! = Val(FormSECONDARY.TextXStageCoordinate.Text) + (Sin(radians!) * (ImageHFW! / 2# * 0.8) / MotUnitsToAngstromMicrons!(XMotor%))
Boundary_Y_Pos1! = Val(FormSECONDARY.TextYStageCoordinate.Text) + (Cos(radians!) * (ImageHFW! / 2# * 0.8) / MotUnitsToAngstromMicrons!(YMotor%))

Boundary_X_Pos2! = Val(FormSECONDARY.TextXStageCoordinate.Text) - (Sin(radians!) * (ImageHFW! / 2# * 0.8) / MotUnitsToAngstromMicrons!(XMotor%))
Boundary_Y_Pos2! = Val(FormSECONDARY.TextYStageCoordinate.Text) - (Cos(radians!) * (ImageHFW! / 2# * 0.8) / MotUnitsToAngstromMicrons!(YMotor%))

' Now load (fake) image coordinates (using 2 point mode)
cpoint1!(1) = FormSECONDARY.Image1.Width
cpoint1!(2) = FormSECONDARY.Image1.Height
cpoint2!(1) = 0
cpoint2!(2) = 0

' Convert from Image1 control "stretched" units to FormPICTURESNAP.Picture2 "unstretched" units
Call PictureSnapUnStretch(Int(0), cpoint1!(1), cpoint1!(2), FormSECONDARY.Image1)
If ierror Then Exit Sub
Call PictureSnapUnStretch(Int(0), cpoint2!(1), cpoint2!(2), FormSECONDARY.Image1)
If ierror Then Exit Sub

If MiscIsInstrumentStage("CAMECA") Then
apoint1!(1) = Val(FormSECONDARY.TextXStageCoordinate.Text) + (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(XMotor%)
apoint1!(2) = Val(FormSECONDARY.TextYStageCoordinate.Text) - (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(YMotor%)
apoint2!(1) = Val(FormSECONDARY.TextXStageCoordinate.Text) - (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(XMotor%)
apoint2!(2) = Val(FormSECONDARY.TextYStageCoordinate.Text) + (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(YMotor%)
Else
apoint1!(1) = Val(FormSECONDARY.TextXStageCoordinate.Text) - (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(XMotor%)
apoint1!(2) = Val(FormSECONDARY.TextYStageCoordinate.Text) + (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(YMotor%)
apoint2!(1) = Val(FormSECONDARY.TextXStageCoordinate.Text) + (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(XMotor%)
apoint2!(2) = Val(FormSECONDARY.TextYStageCoordinate.Text) - (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(YMotor%)
End If

Call PictureSnapSendCalibration(Int(0), cpoint1!(), cpoint2!(), cpoint3!(), apoint1!(), apoint2!(), apoint3!(), Boundary_keV!, Boundary_mag!, Boundary_scanrota!)
If ierror Then Exit Sub
End If
End If
End If

' 2 points
ElseIf dmode% = 2 Then
If Val(FormSECONDARY.TextX1StageCoordinate.Text) >= MotLoLimits!(XMotor%) And Val(FormSECONDARY.TextX1StageCoordinate.Text) <= MotHiLimits!(XMotor%) Then
If Val(FormSECONDARY.TextY1StageCoordinate.Text) >= MotLoLimits!(YMotor%) And Val(FormSECONDARY.TextY1StageCoordinate.Text) <= MotHiLimits!(YMotor%) Then
If Val(FormSECONDARY.TextX2StageCoordinate.Text) >= MotLoLimits!(XMotor%) And Val(FormSECONDARY.TextX2StageCoordinate.Text) <= MotHiLimits!(XMotor%) Then
If Val(FormSECONDARY.TextY2StageCoordinate.Text) >= MotLoLimits!(YMotor%) And Val(FormSECONDARY.TextY2StageCoordinate.Text) <= MotHiLimits!(YMotor%) Then
Boundary_X_Pos1! = Val(FormSECONDARY.TextX1StageCoordinate.Text)
Boundary_Y_Pos1! = Val(FormSECONDARY.TextY1StageCoordinate.Text)
Boundary_X_Pos2! = Val(FormSECONDARY.TextX2StageCoordinate.Text)
Boundary_Y_Pos2! = Val(FormSECONDARY.TextY2StageCoordinate.Text)

' Now load (fake) image coordinates (using 2 point mode)
cpoint1!(1) = FormSECONDARY.Image1.Width
cpoint1!(2) = FormSECONDARY.Image1.Height
cpoint2!(1) = 0
cpoint2!(2) = 0

' Convert from Image1 control "stretched" units to FormPICTURESNAP.Picture2 "unstretched" units
Call PictureSnapUnStretch(Int(0), cpoint1!(1), cpoint1!(2), FormSECONDARY.Image1)
If ierror Then Exit Sub
Call PictureSnapUnStretch(Int(0), cpoint2!(1), cpoint2!(2), FormSECONDARY.Image1)
If ierror Then Exit Sub

If MiscIsInstrumentStage("CAMECA") Then
If Boundary_X_Pos1! < Boundary_X_Pos2! Then
apoint1!(1) = Boundary_X_Pos1! + (Boundary_X_Pos2! - Boundary_X_Pos1!) / 2# + (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(XMotor%)
apoint2!(1) = Boundary_X_Pos1! + (Boundary_X_Pos2! - Boundary_X_Pos1!) / 2# - (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(XMotor%)
Else
apoint1!(1) = Boundary_X_Pos2! + (Boundary_X_Pos1! - Boundary_X_Pos2!) / 2# + (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(XMotor%)
apoint2!(1) = Boundary_X_Pos2! + (Boundary_X_Pos1! - Boundary_X_Pos2!) / 2# - (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(XMotor%)
End If

If Val(FormSECONDARY.TextY1StageCoordinate.Text) < Val(FormSECONDARY.TextY2StageCoordinate.Text) Then
apoint1!(2) = Boundary_Y_Pos1! + (Boundary_Y_Pos2! - Boundary_Y_Pos1!) / 2# - (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(YMotor%)
apoint2!(2) = Boundary_Y_Pos1! + (Boundary_Y_Pos2! - Boundary_Y_Pos1!) / 2# + (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(YMotor%)
Else
apoint1!(2) = Boundary_Y_Pos2! + (Boundary_Y_Pos1! - Boundary_Y_Pos2!) / 2# - (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(YMotor%)
apoint2!(2) = Boundary_Y_Pos2! + (Boundary_Y_Pos1! - Boundary_Y_Pos2!) / 2# + (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(YMotor%)
End If

Else
If Boundary_X_Pos1! < Boundary_X_Pos2! Then
apoint1!(1) = Boundary_X_Pos1! + (Boundary_X_Pos2! - Boundary_X_Pos1!) / 2# - (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(XMotor%)
apoint2!(1) = Boundary_X_Pos1! + (Boundary_X_Pos2! - Boundary_X_Pos1!) / 2# + (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(XMotor%)
Else
apoint1!(1) = Boundary_X_Pos2! + (Boundary_X_Pos1! - Boundary_X_Pos2!) / 2# - (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(XMotor%)
apoint2!(1) = Boundary_X_Pos2! + (Boundary_X_Pos1! - Boundary_X_Pos2!) / 2# + (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(XMotor%)
End If

If Val(FormSECONDARY.TextY1StageCoordinate.Text) < Val(FormSECONDARY.TextY2StageCoordinate.Text) Then
apoint1!(2) = Boundary_Y_Pos1! + (Boundary_Y_Pos2! - Boundary_Y_Pos1!) / 2# + (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(YMotor%)
apoint2!(2) = Boundary_Y_Pos1! + (Boundary_Y_Pos2! - Boundary_Y_Pos1!) / 2# - (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(YMotor%)
Else
apoint1!(2) = Boundary_Y_Pos2! + (Boundary_Y_Pos1! - Boundary_Y_Pos2!) / 2# + (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(YMotor%)
apoint2!(2) = Boundary_Y_Pos2! + (Boundary_Y_Pos1! - Boundary_Y_Pos2!) / 2# - (ImageHFW! / 2#) / MotUnitsToAngstromMicrons!(YMotor%)
End If
End If

Call PictureSnapSendCalibration(Int(0), cpoint1!(), cpoint2!(), cpoint3!(), apoint1!(), apoint2!(), apoint3!(), Boundary_keV!, Boundary_mag!, Boundary_scanrota!)
If ierror Then Exit Sub
End If
End If
End If
End If

' Graphical method (calibration already loaded from ACQ file)
ElseIf dmode% = 3 Then
If BoundaryImageFileName$ <> vbNullString Then
Boundary_X_Pos1! = XX1StageCoordinate!
Boundary_Y_Pos1! = YY1StageCoordinate!
Boundary_X_Pos2! = XX2StageCoordinate!
Boundary_Y_Pos2! = YY2StageCoordinate!
Else
FormSECONDARY.LabelCursorPosition.Caption = vbNullString
End If
End If

' Draw boundary on form
If (dmode% > 0 And dmode% < 3) Or (dmode% = 3 And Trim$(BoundaryImageFileName$) <> vbNullString) Then
Call SecondaryDrawBoundary(Boundary_X_Pos1!, Boundary_Y_Pos1!, Boundary_X_Pos2!, Boundary_Y_Pos2!, FormSECONDARY)
If ierror Then Exit Sub
End If

Exit Sub

' Errors
SecondaryUpdateBoundaryError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryUpdateBoundary"
ierror = True
Exit Sub

End Sub

Sub SecondaryDrawPoints(tForm As Form)
' Draw analyzed points on image

ierror = False
On Error GoTo SecondaryDrawPointsError

Dim i As Long, tcolor As Long

Dim scx As Single, scy As Single, scz As Single
Dim stx As Single, sty As Single, stz As Single
Dim fractionx As Single, fractiony As Single
Dim radius As Single

Dim n As Long
Dim X() As Single, Y() As Single, Z() As Single

' Check if a calibration is loaded
If Not PictureSnapCalibrated Then Exit Sub

' Get coordinate points
Call SecondaryGetCoordinates(n&, X!(), Y!(), Z!())
If ierror Then Exit Sub

' Check for valid points to plot
If n& < 1 Then Exit Sub

' Loop on all points
For i& = 1 To n&
stx! = X!(i&)
sty! = Y!(i&)
stz! = Z!(i&)

' Convert stage to image (form) coordinates and obtain fractional position
Call PictureSnapConvert(Int(2), scx!, scy!, scz!, stx!, sty!, stz!, fractionx!, fractiony!)
If ierror Then Exit Sub

' Convert from "unstretched" units to "stretched" units
Call PictureSnapUnStretch(Int(1), scx!, scy!, tForm.Image1)
If ierror Then Exit Sub

' Adjust for form coordinate scaling
scx! = tForm.Image1.Left + scx!
scy! = tForm.Image1.Top + scy!

' Calculate a radius
If tForm.Image1.Width = 0# Then Exit Sub
radius! = tForm.Image1.Width / 150#

' Display two calibration points
tForm.DrawWidth = 2
tcolor& = RGB(255, 0, 0)
tForm.Circle (scx!, scy!), radius!, tcolor&

' Draw also on FormSECONDARY.Picture1 for clipboard
If tForm.Picture2.ScaleWidth = 0# Then Exit Sub
radius! = tForm.Picture2.Width / 150#
tForm.Picture2.DrawWidth = 2
scx! = tForm.Picture2.ScaleWidth * fractionx!
scy! = tForm.Picture2.ScaleHeight * fractiony!
tForm.Picture2.Circle (scx!, scy!), radius!, tcolor&

Next i&

Exit Sub

' Errors
SecondaryDrawPointsError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryDrawPoints"
ierror = True
Exit Sub

End Sub

Sub SecondaryLoadMode(Index As Integer)
' Load controls based on distance method

ierror = False
On Error GoTo SecondaryLoadModeError

' Fixed distance
If Index% = 0 Then
FormSECONDARY.TextSpecifiedDistance.Enabled = True
FormSECONDARY.TextXStageCoordinate.Enabled = False
FormSECONDARY.TextYStageCoordinate.Enabled = False
FormSECONDARY.TextBoundaryAngle.Enabled = False
FormSECONDARY.TextX1StageCoordinate.Enabled = False
FormSECONDARY.TextY1StageCoordinate.Enabled = False
FormSECONDARY.TextX2StageCoordinate.Enabled = False
FormSECONDARY.TextY2StageCoordinate.Enabled = False
FormSECONDARY.TextHFW.Enabled = False
FormSECONDARY.CommandCopyToClipboard.Enabled = False
FormSECONDARY.CommandPrintImage.Enabled = False
FormSECONDARY.Image1.Enabled = False
FormSECONDARY.Image1.Picture = LoadPicture()
FormSECONDARY.LabelBoundaryCoordinates.Caption = vbNullString

' X,Y and angle
ElseIf Index% = 1 Then
FormSECONDARY.TextSpecifiedDistance.Enabled = False
FormSECONDARY.TextXStageCoordinate.Enabled = True
FormSECONDARY.TextYStageCoordinate.Enabled = True
FormSECONDARY.TextBoundaryAngle.Enabled = True
FormSECONDARY.TextX1StageCoordinate.Enabled = False
FormSECONDARY.TextY1StageCoordinate.Enabled = False
FormSECONDARY.TextX2StageCoordinate.Enabled = False
FormSECONDARY.TextY2StageCoordinate.Enabled = False
FormSECONDARY.TextHFW.Enabled = True
FormSECONDARY.CommandCopyToClipboard.Enabled = False
FormSECONDARY.CommandPrintImage.Enabled = False
FormSECONDARY.Image1.Enabled = True
FormSECONDARY.Image1.Picture = LoadPicture()
FormSECONDARY.LabelBoundaryCoordinates.Caption = vbNullString

' X,Y pair
ElseIf Index% = 2 Then
FormSECONDARY.TextSpecifiedDistance.Enabled = False
FormSECONDARY.TextXStageCoordinate.Enabled = False
FormSECONDARY.TextYStageCoordinate.Enabled = False
FormSECONDARY.TextBoundaryAngle.Enabled = False
FormSECONDARY.TextX1StageCoordinate.Enabled = True
FormSECONDARY.TextY1StageCoordinate.Enabled = True
FormSECONDARY.TextX2StageCoordinate.Enabled = True
FormSECONDARY.TextY2StageCoordinate.Enabled = True
FormSECONDARY.TextHFW.Enabled = True
FormSECONDARY.CommandCopyToClipboard.Enabled = False
FormSECONDARY.CommandPrintImage.Enabled = False
FormSECONDARY.Image1.Enabled = True
FormSECONDARY.Image1.Picture = LoadPicture()
FormSECONDARY.LabelBoundaryCoordinates.Caption = vbNullString

' Graphical boundary
Else
FormSECONDARY.TextSpecifiedDistance.Enabled = False
FormSECONDARY.TextXStageCoordinate.Enabled = False
FormSECONDARY.TextYStageCoordinate.Enabled = False
FormSECONDARY.TextBoundaryAngle.Enabled = False
FormSECONDARY.TextX1StageCoordinate.Enabled = False
FormSECONDARY.TextY1StageCoordinate.Enabled = False
FormSECONDARY.TextX2StageCoordinate.Enabled = False
FormSECONDARY.TextY2StageCoordinate.Enabled = False
FormSECONDARY.TextHFW.Enabled = False
FormSECONDARY.CommandCopyToClipboard.Enabled = True
FormSECONDARY.CommandPrintImage.Enabled = True
FormSECONDARY.Image1.Enabled = True
If BoundaryImageFileName$ <> vbNullString Then
Call SecondaryLoadImage(BoundaryImageFileName$)
If ierror Then Exit Sub
FormSECONDARY.LabelBoundaryCoordinates.Caption = bmsg$
End If
End If

Exit Sub

' Errors
SecondaryLoadModeError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryLoadMode"
ierror = True
Exit Sub

End Sub

Sub SecondaryCopyToClipboard()
' Copy the image (Picture1) to clipboard with graphics objects

ierror = False
On Error GoTo SecondaryCopyToClipboardError

' Load again
Call SecondaryLoadImage(BoundaryImageFileName$)
If ierror Then Exit Sub

' Redraw graphics objects
Call SecondaryUpdateBoundary
If ierror Then Exit Sub
Call SecondaryDrawPoints(FormSECONDARY)
If ierror Then Exit Sub

' Copy image and graphics objects to clipboard
Call BMPCopyEntirePicture(FormSECONDARY.Picture2)
If ierror Then Exit Sub

Exit Sub

' Errors
SecondaryCopyToClipboardError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryCopyToClipboard"
ierror = True
Exit Sub

End Sub

Sub SecondaryPrintImage(tForm As Form)
' Print the object picture property to the printer

ierror = False
On Error GoTo SecondaryPrintImageError

' Print image in Picture1 (Picture3 is for temporary use)
Call BMPPrintDiagram(tForm.Picture1, tForm.Picture3, CSng(0.5), CSng(0.5), CSng(7 * ImageInterfaceImageIxIy!), CSng(7#))
If ierror Then Exit Sub

Exit Sub

' Errors
SecondaryPrintImageError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryPrintImage"
ierror = True
Exit Sub

End Sub

Sub SecondarySampleLoadFrom(chan As Integer, sample() As TypeSample)
' Load the module parameters from the passed sample

ierror = False
On Error GoTo SecondarySampleLoadFromError

' Load the module level parameters from the passed sample
If sample(1).SecondaryFluorescenceBoundaryFlag(chan%) Then
SecondaryFluorescenceFlag = True
Else
SecondaryFluorescenceFlag = False
End If

' Material A and B (and B std) could be swapped for some elements
MatA_String$ = sample(1).SecondaryFluorescenceBoundaryMatA_String$(chan%)
MatB_String$ = sample(1).SecondaryFluorescenceBoundaryMatB_String$(chan%)
MatBStd_String$ = sample(1).SecondaryFluorescenceBoundaryMatBStd_String$(chan%)

' Load selected k-ratio DAT file parameters
BoundaryKratiosDATFile$ = sample(1).SecondaryFluorescenceBoundaryKratiosDATFile$(chan%)
BoundaryKratiosDATFileLine1$ = sample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine1$(chan%)
BoundaryKratiosDATFileLine2$ = sample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine2$(chan%)
BoundaryKratiosDATFileLine3$ = sample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine3$(chan%)

' Load image parameters (same for all elements)
BoundaryImageFileName$ = sample(1).SecondaryFluorescenceBoundaryImageFileName$
BoundaryImageNumber% = sample(1).SecondaryFluorescenceBoundaryImageNumber%

' Load distance method (same for all elements)
SecondaryFluorescenceDistanceMethod% = sample(1).SecondaryFluorescenceBoundaryDistanceMethod%
SpecifiedDistance! = sample(1).SecondaryFluorescenceBoundarySpecifiedDistance!

' Load 2 point boundary coordinates (same for all elements)
Boundary_X_Pos1! = sample(1).SecondaryFluorescenceBoundaryCoordinateX1!
Boundary_X_Pos2! = sample(1).SecondaryFluorescenceBoundaryCoordinateX2!
Boundary_Y_Pos1! = sample(1).SecondaryFluorescenceBoundaryCoordinateY1!
Boundary_Y_Pos2! = sample(1).SecondaryFluorescenceBoundaryCoordinateY2!

' Load conditions
Boundary_keV! = sample(1).kilovolts!
Boundary_mag! = sample(1).magnificationimaging!
Boundary_scanrota! = DefaultScanRotation!

Exit Sub

' Errors
SecondarySampleLoadFromError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondarySampleLoadFrom"
ierror = True
Exit Sub

End Sub

Sub SecondarySampleSaveTo(chan As Integer, tImageHFW As Single, sample() As TypeSample)
' Saves the FormSECONDARY to the passed sample

ierror = False
On Error GoTo SecondarySampleSaveToError

' Save the module level parameters to the passed sample
If SecondaryFluorescenceFlag = True Then
sample(1).SecondaryFluorescenceBoundaryFlag(chan%) = True
Else
sample(1).SecondaryFluorescenceBoundaryFlag(chan%) = False
End If

' Material A and B (and B std) could be swapped for some elements
sample(1).SecondaryFluorescenceBoundaryMatA_String$(chan%) = MatA_String$
sample(1).SecondaryFluorescenceBoundaryMatB_String$(chan%) = MatB_String$
sample(1).SecondaryFluorescenceBoundaryMatBStd_String$(chan%) = MatBStd_String$

' Save selected k-ratio DAT file parameters
sample(1).SecondaryFluorescenceBoundaryKratiosDATFile$(chan%) = BoundaryKratiosDATFile$
sample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine1$(chan%) = BoundaryKratiosDATFileLine1$
sample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine2$(chan%) = BoundaryKratiosDATFileLine2$
sample(1).SecondaryFluorescenceBoundaryKratiosDATFileLine3$(chan%) = BoundaryKratiosDATFileLine3$

' Save selected image file (same for all elements)
sample(1).SecondaryFluorescenceBoundaryImageFileName$ = BoundaryImageFileName$
sample(1).SecondaryFluorescenceBoundaryImageNumber% = BoundaryImageNumber%

' Save distance method (same for all elements)
sample(1).SecondaryFluorescenceBoundaryDistanceMethod% = SecondaryFluorescenceDistanceMethod%
sample(1).SecondaryFluorescenceBoundarySpecifiedDistance! = SpecifiedDistance!

' Save boundary (all distance methods save two points)
sample(1).SecondaryFluorescenceBoundaryCoordinateX1! = Boundary_X_Pos1!
sample(1).SecondaryFluorescenceBoundaryCoordinateY1! = Boundary_Y_Pos1!
sample(1).SecondaryFluorescenceBoundaryCoordinateX2! = Boundary_X_Pos2!
sample(1).SecondaryFluorescenceBoundaryCoordinateY2! = Boundary_Y_Pos2!

' In case user changed ImageHFW in FormSECONDARY
tImageHFW! = ImageHFW!

Exit Sub

' Errors
SecondarySampleSaveToError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondarySampleSaveTo"
ierror = True
Exit Sub

End Sub

Sub SecondaryUpdatePositions(mode As Integer)
' Update the boundary coordinates based on current stage position
' mode = 0 update boundary coordinate (and angle)
' mode = 1 update boundary coordinate (first pair)
' mode = 2 update boundary coordinate (second pair)

ierror = False
On Error GoTo SecondaryUpdatePositionsError

If Not RealTimeMode Then GoTo SecondaryUpdatePositionsNotRealTime

' Get current positions
If mode% = 0 Then
FormSECONDARY.TextXStageCoordinate.Text = MiscAutoFormat$(RealTimeMotorPositions!(XMotor%))
FormSECONDARY.TextYStageCoordinate.Text = MiscAutoFormat$(RealTimeMotorPositions!(YMotor%))
FormSECONDARY.OptionDistanceMethod(1).value = True

ElseIf mode% = 1 Then
FormSECONDARY.TextX1StageCoordinate.Text = MiscAutoFormat$(RealTimeMotorPositions!(XMotor%))
FormSECONDARY.TextY1StageCoordinate.Text = MiscAutoFormat$(RealTimeMotorPositions!(YMotor%))
FormSECONDARY.OptionDistanceMethod(2).value = True

ElseIf mode% = 2 Then
FormSECONDARY.TextX2StageCoordinate.Text = MiscAutoFormat$(RealTimeMotorPositions!(XMotor%))
FormSECONDARY.TextY2StageCoordinate.Text = MiscAutoFormat$(RealTimeMotorPositions!(YMotor%))
FormSECONDARY.OptionDistanceMethod(2).value = True
End If

Exit Sub

' Errors
SecondaryUpdatePositionsError:
MsgBox Error$, vbOKOnly + vbCritical, "SecondaryUpdatePositions"
ierror = True
Exit Sub

SecondaryUpdatePositionsNotRealTime:
msg$ = "The software is not currently connected to an instrument.  Please make a connection to the instrument and try again, or enter the boundary coordinates manually."
MsgBox msg$, vbOKOnly + vbExclamation, "SecondaryUpdatePositions"
ierror = True
Exit Sub

End Sub

