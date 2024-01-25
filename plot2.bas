Attribute VB_Name = "CodePLOT2"
' (c) Copyright 1995-2024 by John J. Donovan
Option Explicit

Sub PlotGetRelativeHours(i As Integer, lasttime As Double, nexttime As Double, elapsedhours As Single, sample() As TypeSample)
' Calculates the relative change in hours

ierror = False
On Error GoTo PlotGetRelativeHoursError

If lasttime# <> 0# Then
nexttime# = sample(1).DateTimes(i%) - lasttime#
elapsedhours! = elapsedhours! + nexttime# * 24#
End If

' Save last time
lasttime# = sample(1).DateTimes(i%)

Exit Sub

' Errors
PlotGetRelativeHoursError:
MsgBox Error$, vbOKOnly + vbCritical, "PlotGetRelativeHours"
ierror = True
Exit Sub

End Sub

Sub PlotAddItems(n As Integer, astring As String)
' Loads the list string into all three axis list boxes

ierror = False
On Error GoTo PlotAddItemsError

FormPLOT.ListXAxis.AddItem astring
FormPLOT.ListXAxis.ItemData(FormPLOT.ListXAxis.NewIndex) = n%   ' composite number containing x-ray, motor, crystal indexes

FormPLOT.ListYAxis.AddItem astring
FormPLOT.ListYAxis.ItemData(FormPLOT.ListYAxis.NewIndex) = n%

FormPLOT.ListZAxis.AddItem astring
FormPLOT.ListZAxis.ItemData(FormPLOT.ListZAxis.NewIndex) = n%

Exit Sub

' Errors
PlotAddItemsError:
MsgBox Error$, vbOKOnly + vbCritical, "PlotAddItems"
ierror = True
Exit Sub

End Sub

Function PlotGetSelectedType(i As Integer, astring As String) As Integer
' Returns the calculation type and axis label
'  returns 0 if raw data
'  returns 1 if elemental percents or total calculation
'  returns 2 if oxide percents or total calculation
'  returns 3 if atomic percents or total calculation
'  returns 4 if formula atoms or total calculation
'  returns 5 if raw k-ratio calculation
'  returns 6 if detection limit calculation
'  returns 7 if percent error calculation
' Also returns the list string for the axis label

ierror = False
On Error GoTo PlotGetSelectedTypeError

' Return the list box string (all axis list boxes are the same)
PlotGetSelectedType% = 0
astring$ = FormPLOT.ListXAxis.List(i%)

' Load stage coordinate labels
If InStr(astring$, "X Stage Coordinates") > 0 Then Exit Function
If InStr(astring$, "Y Stage Coordinates") > 0 Then Exit Function
If InStr(astring$, "Z Stage Coordinates") > 0 Then Exit Function

If InStr(astring$, "Elemental Percents") > 0 Then PlotGetSelectedType% = 1
If InStr(astring$, "Oxide Percents") > 0 Then PlotGetSelectedType% = 2
If InStr(astring$, "Atomic Percents") > 0 Then PlotGetSelectedType% = 3
If InStr(astring$, "Formula Atoms") > 0 Then PlotGetSelectedType% = 4
If InStr(astring$, "Raw K-Ratios") > 0 Then PlotGetSelectedType% = 5
If InStr(astring$, "Detection Limits") > 0 Then PlotGetSelectedType% = 6
If InStr(astring$, "Percent Errors") > 0 Then PlotGetSelectedType% = 7

If InStr(astring$, "Elemental Totals") > 0 Then PlotGetSelectedType% = 1
If InStr(astring$, "Oxide Totals") > 0 Then PlotGetSelectedType% = 2
If InStr(astring$, "Atomic Totals") > 0 Then PlotGetSelectedType% = 3
If InStr(astring$, "Formula Totals") > 0 Then PlotGetSelectedType% = 4

If InStr(astring$, "On Counts (P+B)") > 0 Then PlotGetSelectedType% = 1     ' calculate MAN background intensities
If InStr(astring$, "On Counts (P-B)") > 0 Then PlotGetSelectedType% = 1     ' calculate MAN background intensities

Exit Function

' Errors
PlotGetSelectedTypeError:
MsgBox Error$, vbOKOnly + vbCritical, "PlotGetSelectedType"
ierror = True
Exit Function

End Function

Sub PlotOutputBASFile(basfile As String, sampletitle As String, nsets As Integer, xlabel As String, zlabel As String, ldata() As String)
' Output a surfer .BAS script file for creating 3D surface and contour plots
' using Golden Software's SURFER program.

ierror = False
On Error GoTo PlotOutputBASFileError

Dim i As Integer
Dim OS64bitMode As Boolean
Dim astring As String
Dim tfilename As String, filedir As String

' Get complete path only
filedir$ = MiscGetPathOnly$(basfile$)

' Get filename only without extension
tfilename$ = MiscGetFileNameNoExtension$(MiscGetFileNameOnly$(basfile$))

' Open the file
Open basfile$ For Output As #Temp1FileNumber%

astring$ = VbSquote$ & " This routine is created by Probe for EPMA"
Print #Temp1FileNumber%, astring$
Print #Temp1FileNumber%, vbNullString

' Bug in Win 7 (64) causes current directory to change- comment out CurDir statement if so
OS64bitMode = MiscSystemIsHost64Bit()

' Output variables (based on SurferOutputVersionNumber)
astring$ = "Directory$ = " & VbDquote$ & filedir$ & VbDquote$           ' use this to document original data folder
Print #Temp1FileNumber%, astring$

If OS64bitMode Then
astring$ = "'Directory$ = CurDir$() & " & VbDquote$ & "\" & VbDquote$                ' use this for actual variable
Else
astring$ = "Directory$ = CurDir$() & " & VbDquote$ & "\" & VbDquote$                ' use this for actual variable
End If
Print #Temp1FileNumber%, astring$

astring$ = "If Command$() <> " & VbDquote$ & VbDquote$ & " Then Directory$ = Command$()    ' see if script path was passed as a command line argument"
Print #Temp1FileNumber%, astring$
Print #Temp1FileNumber%, vbNullString

astring$ = "File$ = " & VbDquote$ & tfilename$ & VbDquote$
Print #Temp1FileNumber%, astring$
astring$ = "Sample$ = " & VbDquote$ & sampletitle$ & VbDquote$   ' leave blank for now
Print #Temp1FileNumber%, astring$

' Output number of columns line
astring$ = "MaxCol% = " & Str$(nsets%)
Print #Temp1FileNumber%, astring$
astring$ = "Dim ZLabel$(MaxCol%) as String"
Print #Temp1FileNumber%, astring$

' Output column labels
astring$ = "XLabel$ = " & VbDquote$ & xlabel$ & VbDquote$
Print #Temp1FileNumber%, astring$
astring$ = "YLabel$ = " & VbDquote$ & zlabel$ & VbDquote$
Print #Temp1FileNumber%, astring$

' Z labels. Note that array for Surfer is dimensioned from zero.
astring$ = "Dim ZLabel$(63) as String"
For i% = 1 To nsets%
astring$ = "ZLabel$(" & Format$(i% - 1) & ") = " & VbDquote$ & ldata$(i%) & VbDquote$
Print #Temp1FileNumber%, astring$
Next i%
Print #Temp1FileNumber%, vbNullString

' Add code to change back to data folder since UAC changes this in Win 7
astring$ = VbSquote$ & " Change back to data folder (need for Win 7 64 bit)"
Print #Temp1FileNumber%, astring$
astring$ = "ChDrive Directory$"
Print #Temp1FileNumber%, astring$
astring$ = "ChDir Directory$"
Print #Temp1FileNumber%, astring$
Print #Temp1FileNumber%, vbNullString

' Now concatanate the .BAS file with GRIDBB.BAS
Open GRIDBB_BAS_File$ For Input As #Temp2FileNumber%

Do While Not EOF(Temp2FileNumber%)
Line Input #Temp2FileNumber%, astring$
Print #Temp1FileNumber%, astring$
Loop

Close #Temp1FileNumber%
Close #Temp2FileNumber%

Exit Sub

' Errors
PlotOutputBASFileError:
MsgBox Error$, vbOKOnly + vbCritical, "PlotOutputBASFile"
Close #Temp1FileNumber%
Close #Temp2FileNumber%
ierror = True
Exit Sub

End Sub

Sub PlotOutputBASFile2(basfile As String, sampletitle As String, nsets As Integer, xlabel As String, zlabel As String, ldata() As String)
' Output a surfer .BAS script file for creating 3D surface and contour plots
' using Golden Software's SURFER (version 7.0 or later) program.

ierror = False
On Error GoTo PlotOutputBASFile2Error

Dim i As Integer
Dim OS64bitMode As Boolean
Dim astring As String
Dim tfilename As String, filedir As String

Dim gX_Polarity As Integer, gY_Polarity As Integer
Dim gStage_Units As String

' Get complete path only
filedir$ = MiscGetPathOnly$(basfile$)

' Get filename only without extension
tfilename$ = MiscGetFileNameNoExtension$(MiscGetFileNameOnly$(basfile$))

' Open the file
Open basfile$ For Output As #Temp1FileNumber%

' Output module level declarations
astring$ = "Option Explicit"
Print #Temp1FileNumber%, astring$
Print #Temp1FileNumber%, vbNullString

astring$ = "Dim SurferApp As Object"
Print #Temp1FileNumber%, astring$
astring$ = "Dim SurferWks As Object"
Print #Temp1FileNumber%, astring$
astring$ = "Dim SurferDoc As Object"
Print #Temp1FileNumber%, astring$
astring$ = "Dim SurferPlot As Object"
Print #Temp1FileNumber%, astring$
astring$ = "Dim SurferPageSetup as Object"
Print #Temp1FileNumber%, astring$
astring$ = "Dim SurferShapes As Object"
Print #Temp1FileNumber%, astring$
astring$ = "Dim SurferMapFrame As Object"
Print #Temp1FileNumber%, astring$
astring$ = "Dim SurferAxes As Object"
Print #Temp1FileNumber%, astring$
astring$ = "Dim SurferAxis As Object"
Print #Temp1FileNumber%, astring$
astring$ = "Dim SurferSelection As Object"
Print #Temp1FileNumber%, astring$
astring$ = "Dim SurferText As Object"
Print #Temp1FileNumber%, astring$
astring$ = "Dim SurferFontFormat As Object"
Print #Temp1FileNumber%, astring$

astring$ = "Dim SurferImageMap As Object"       ' new 01/16/2010
Print #Temp1FileNumber%, astring$
astring$ = "Dim SurferColorMap As Object"
Print #Temp1FileNumber%, astring$

Print #Temp1FileNumber%, vbNullString

' Create sub main routine
astring$ = "Sub Main"
Print #Temp1FileNumber%, astring$
astring$ = VbSquote$ & " This routine is created by Probe for EPMA"
Print #Temp1FileNumber%, astring$
Print #Temp1FileNumber%, vbNullString

astring$ = "Dim Directory As String"
Print #Temp1FileNumber%, astring$
astring$ = "Dim File As String"
Print #Temp1FileNumber%, astring$
astring$ = "Dim Sample As String"
Print #Temp1FileNumber%, astring$
astring$ = "Dim MaxCol As Integer"
Print #Temp1FileNumber%, astring$
astring$ = "Dim XLabel As String"
Print #Temp1FileNumber%, astring$
astring$ = "Dim YLabel As String"
Print #Temp1FileNumber%, astring$
Print #Temp1FileNumber%, vbNullString

' Output JEOL stage invert flags
astring$ = "Dim XInvert As Integer"
Print #Temp1FileNumber%, astring$
astring$ = "Dim YInvert As Integer"
Print #Temp1FileNumber%, astring$
Print #Temp1FileNumber%, vbNullString

' Bug in Win 7 (64) causes current directory to change- comment out CurDir statement if so
OS64bitMode = MiscSystemIsHost64Bit()

' Output run specific strings
astring$ = "Directory$ = " & VbDquote$ & filedir$ & VbDquote$           ' use this to document original data folder
Print #Temp1FileNumber%, astring$

If OS64bitMode Then
astring$ = "'Directory$ = CurDir$() & " & VbDquote$ & "\" & VbDquote$                ' use this for actual variable
Else
astring$ = "Directory$ = CurDir$() & " & VbDquote$ & "\" & VbDquote$                ' use this for actual variable
End If
Print #Temp1FileNumber%, astring$

astring$ = "If Command$() <> " & VbDquote$ & VbDquote$ & " Then Directory$ = Command$()    ' see if script path was passed as a command line argument"
Print #Temp1FileNumber%, astring$
Print #Temp1FileNumber%, vbNullString

astring$ = "File$ = " & VbDquote$ & tfilename$ & VbDquote$
Print #Temp1FileNumber%, astring$
astring$ = "Sample$ = " & VbDquote$ & sampletitle$ & VbDquote$   ' leave blank for now
Print #Temp1FileNumber%, astring$

' Output column labels
astring$ = "XLabel$ = " & VbDquote$ & xlabel$ & VbDquote$
Print #Temp1FileNumber%, astring$
astring$ = "YLabel$ = " & VbDquote$ & zlabel$ & VbDquote$
Print #Temp1FileNumber%, astring$

' Get stage polarity
Call GridCheckGRDInfo(tfilename$ & ".grd", gX_Polarity%, gY_Polarity%, gStage_Units$)
If ierror Then Exit Sub

' Write axis invert flags (for JEOL "anti-cartesian" display in Surfer)
astring$ = "XInvert% = " & Format$(gX_Polarity%)
Print #Temp1FileNumber%, astring$
astring$ = "YInvert% = " & Format$(gY_Polarity%)
Print #Temp1FileNumber%, astring$
Print #Temp1FileNumber%, vbNullString

' Output number of columns line
astring$ = "MaxCol% = " & Str$(nsets%)
Print #Temp1FileNumber%, astring$
astring$ = "ReDim ZLabel(1 to MaxCol%) as String"
Print #Temp1FileNumber%, astring$

' Z labels. Note that array for Surfer is dimensioned from 1.
For i% = 1 To nsets%
astring$ = "ZLabel$(" & Format$(i%) & ") = " & VbDquote$ & ldata$(i%) & VbDquote$
Print #Temp1FileNumber%, astring$
Next i%
Print #Temp1FileNumber%, vbNullString

' Add code to change back to data folder since UAC changes this in Win 7
astring$ = VbSquote$ & " Change back to data folder (need for Win 7 64 bit)"
Print #Temp1FileNumber%, astring$
astring$ = "ChDrive Directory$"
Print #Temp1FileNumber%, astring$
astring$ = "ChDir Directory$"
Print #Temp1FileNumber%, astring$
Print #Temp1FileNumber%, vbNullString

astring$ = VbSquote$ & " Call output routine"
Print #Temp1FileNumber%, astring$
astring$ = "Call GridAll(Directory$, File$, Sample$, XInvert%, YInvert%, MaxCol%, XLabel$, YLabel$, ZLabel$())"
Print #Temp1FileNumber%, astring$
Print #Temp1FileNumber%, vbNullString

astring$ = "End Sub"
Print #Temp1FileNumber%, astring$
Print #Temp1FileNumber%, vbNullString

' Now concatanate the .BAS file with GRIDCC.BAS
Open GRIDCC_BAS_File$ For Input As #Temp2FileNumber%

Do While Not EOF(Temp2FileNumber%)
Line Input #Temp2FileNumber%, astring$
Print #Temp1FileNumber%, astring$
Loop

Close #Temp1FileNumber%
Close #Temp2FileNumber%

Exit Sub

' Errors
PlotOutputBASFile2Error:
MsgBox Error$, vbOKOnly + vbCritical, "PlotOutputBASFile2"
Close #Temp1FileNumber%
Close #Temp2FileNumber%
ierror = True
Exit Sub

End Sub

