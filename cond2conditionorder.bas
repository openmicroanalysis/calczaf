Attribute VB_Name = "CodeCOND2ConditionOrder"
' (c) Copyright 1995-2026 by John J. Donovan
Option Explicit

' Column condition arrays for ordering
Dim ConditionsNumberOf As Integer
Dim Condition_Takeoffs() As Single
Dim Condition_Kilovolts() As Single
Dim Condition_BeamCurrents() As Single
Dim Condition_BeamSizes() As Single

Dim Condition_ColMethods() As Integer
Dim Condition_ColStrings() As String

Sub Cond2ConditionDefaultOrder(sample() As TypeSample)
' Load default combined condition orders for new or modified sample

ierror = False
On Error GoTo Cond2ConditionDefaultOrderError

Dim n As Integer, nn As Integer
Dim chan As Integer, tchan As Integer
Dim bMatched As Boolean, tAddCond As Boolean

' Check for at least one element
If sample(1).LastElm% <= 0 Then Exit Sub

' Re-set check for matched conditions (load channel 1 as default)
bMatched = Cond2ConditionMatched(Int(0), Int(1), nn%, sample())

' Loop on sample until all conditions loaded
bMatched = False
Do Until bMatched
tAddCond = False

' Check for this condition in existing conditions
For chan% = 1 To sample(1).LastElm%
bMatched = Cond2ConditionMatched(Int(1), chan%, nn%, sample())
If Not bMatched Then
tAddCond = True
tchan% = chan%
Exit For
End If
Next chan%

' Found new condition for this sample, so load new condition
If tAddCond Then
n% = ConditionsNumberOf% + 1

' Increase array
If n% > ConditionsNumberOf% Then
ReDim Preserve Condition_Takeoffs(1 To n%) As Single
ReDim Preserve Condition_Kilovolts(1 To n%) As Single
ReDim Preserve Condition_BeamCurrents(1 To n%) As Single
ReDim Preserve Condition_BeamSizes(1 To n%) As Single
ReDim Preserve Condition_ColMethods(1 To n%) As Integer
ReDim Preserve Condition_ColStrings(1 To n%) As String
End If

' Save condition to module level variables
Condition_Takeoffs!(n%) = sample(1).TakeoffArray!(tchan%)
Condition_Kilovolts!(n%) = sample(1).KilovoltsArray!(tchan%)
Condition_BeamCurrents!(n%) = sample(1).BeamCurrentArray!(tchan%)
Condition_BeamSizes!(n%) = sample(1).BeamSizeArray!(tchan%)
Condition_ColMethods%(n%) = sample(1).ColumnConditionMethodArray%(tchan%)
Condition_ColStrings$(n%) = sample(1).ColumnConditionStringArray$(tchan%)

ConditionsNumberOf% = n%
End If

Loop

' Now reload all condition orders based on new conditions
For chan% = 1 To sample(1).LastElm%
bMatched = Cond2ConditionMatched(Int(1), chan%, nn%, sample())
If Not bMatched Then GoTo Cond2ConditionDefaultOrderNotFound
sample(1).ConditionNumbers%(chan%) = nn%
Next chan%

' Load (for now) sample combined condition acquisition order as condition step
For n% = 1 To Cond2ConditionGetMaxCondition%(sample())
sample(1).ConditionOrders%(n%) = n%
Next n%

Exit Sub

' Errors
Cond2ConditionDefaultOrderError:
MsgBox Error$, vbOKOnly + vbCritical, "Cond2ConditionDefaultOrder"
ierror = True
Exit Sub

Cond2ConditionDefaultOrderNotFound:
msg$ = "Channel " & Format$(chan%) & "(" & sample(1).Elsyms$(chan%) & " " & sample(1).Xrsyms$(chan%) & ") could not be matched to an existing condition. This error should not occur, please contact Probe Software technical support."
MsgBox msg$, vbOKOnly + vbExclamation, "Cond2ConditionDefaultOrder"
ierror = True
Exit Sub

End Sub

Function Cond2ConditionMatched(mode As Integer, chan As Integer, nn As Integer, sample() As TypeSample) As Integer
' Checks to see if the sample conditions that are different from the parameters since the last time called.
'  mode = 0 reset condition
'  mode = 1  check for changed condition
'  chan is the element channel to check the conditions for
'  nn is the column condition number that was matched (0 = no match)

ierror = False
On Error GoTo Cond2ConditionMatchedError

Dim n As Integer

' Set default as changed
Cond2ConditionMatched = False

' Check for just a re-set
If mode% = 0 Then
ConditionsNumberOf% = 1
ReDim Condition_Takeoffs(1 To ConditionsNumberOf%) As Single
ReDim Condition_Kilovolts(1 To ConditionsNumberOf%) As Single
ReDim Condition_BeamCurrents(1 To ConditionsNumberOf%) As Single
ReDim Condition_BeamSizes(1 To ConditionsNumberOf%) As Single
ReDim Condition_ColMethods(1 To ConditionsNumberOf%) As Integer
ReDim Condition_ColStrings(1 To ConditionsNumberOf%) As String

Condition_Takeoffs!(ConditionsNumberOf%) = sample(1).TakeoffArray!(chan%)     ' load passed element as default
Condition_Kilovolts!(ConditionsNumberOf%) = sample(1).KilovoltsArray!(chan%)
Condition_BeamCurrents!(ConditionsNumberOf%) = sample(1).BeamCurrentArray!(chan%)
Condition_BeamSizes!(ConditionsNumberOf%) = sample(1).BeamSizeArray!(chan%)
Condition_ColMethods%(ConditionsNumberOf%) = sample(1).ColumnConditionMethodArray%(chan%)
Condition_ColStrings$(ConditionsNumberOf%) = sample(1).ColumnConditionStringArray$(chan%)
nn% = 0
Exit Function
End If

' Using analytical conditions (analytical *and* column condition strings must match!)
For n% = 1 To ConditionsNumberOf%

' Using analytical conditions
If sample(1).ColumnConditionMethodArray%(chan%) = 0 Then
If sample(1).TakeoffArray!(chan%) = Condition_Takeoffs!(n%) And _
sample(1).KilovoltsArray!(chan%) = Condition_Kilovolts!(n%) And _
sample(1).BeamCurrentArray!(chan%) = Condition_BeamCurrents!(n%) And _
sample(1).BeamSizeArray!(chan%) = Condition_BeamSizes!(n%) And _
sample(1).ColumnConditionStringArray$(chan%) = Condition_ColStrings$(n%) Then
Cond2ConditionMatched = True
nn% = n%
End If

' Using column conditions
Else
If sample(1).ColumnConditionStringArray$(chan%) = Condition_ColStrings$(n%) Then
Cond2ConditionMatched = True
nn% = n%
End If
End If
Next n%

Exit Function

' Errors
Cond2ConditionMatchedError:
MsgBox Error$, vbOKOnly + vbCritical, "Cond2ConditionMatched"
ierror = True
Exit Function

End Function

Sub Cond2ConditionDisplayGrid(tForm As Form)
' Load combined condition to the display grid

ierror = False
On Error GoTo Cond2ConditionDisplayGridError

Dim i As Integer

' Load size
tForm.GridConditions.cols = 7
tForm.GridConditions.rows = ConditionsNumberOf% + 1

' Blank grid
tForm.GridConditions.Clear

' Resize and label fixed row
For i% = 0 To tForm.GridConditions.cols - 1
tForm.GridConditions.ColWidth(i%) = 400
Next i%

' Make last column very long for filename
tForm.GridConditions.ColWidth(tForm.GridConditions.cols - 1) = 4000

' Load fixed column labels
tForm.GridConditions.row = 0
tForm.GridConditions.col = 0
tForm.GridConditions.Text = "Num"

tForm.GridConditions.col = 1
tForm.GridConditions.Text = "TO"

tForm.GridConditions.col = 2
tForm.GridConditions.Text = "keV"

tForm.GridConditions.col = 3
tForm.GridConditions.Text = "nA"

tForm.GridConditions.col = 4
tForm.GridConditions.Text = "um"

tForm.GridConditions.col = 5
tForm.GridConditions.Text = "Col"

tForm.GridConditions.col = 6
tForm.GridConditions.Text = "File"

' Load the conditions to the grid
For i% = 1 To ConditionsNumberOf%
tForm.GridConditions.row = i%

tForm.GridConditions.col = 0
tForm.GridConditions.Text = Format$(i%)

tForm.GridConditions.col = 1
tForm.GridConditions.Text = Format$(Condition_Takeoffs!(i%))

tForm.GridConditions.col = 2
tForm.GridConditions.Text = Format$(Condition_Kilovolts!(i%))

tForm.GridConditions.col = 3
tForm.GridConditions.Text = Format$(Condition_BeamCurrents!(i%))

tForm.GridConditions.col = 4
tForm.GridConditions.Text = Format$(Condition_BeamSizes!(i%))

tForm.GridConditions.col = 5
tForm.GridConditions.Text = Format$(Condition_ColMethods%(i%))

tForm.GridConditions.col = 6
tForm.GridConditions.Text = MiscGetFileNameOnly$(Condition_ColStrings$(i%))
Next i%

Exit Sub

' Errors
Cond2ConditionDisplayGridError:
MsgBox Error$, vbOKOnly + vbCritical, "Cond2ConditionDisplayGrid"
ierror = True
Exit Sub

End Sub

Function Cond2ConditionGetNextChannel(motor As Integer, tOrder As Integer, tCond As Integer, sample() As TypeSample) As Integer
' Returns the next channel for this spectrometer number, spectrometer order and condition order

ierror = False
On Error GoTo Cond2ConditionGetNextChannelError

Dim chan As Integer, nextchan As Integer

' Find next spectro and element for this condition order
nextchan% = 0
For chan% = 1 To sample(1).LastElm%
If sample(1).ConditionNumbers%(chan%) = tCond% Then
If sample(1).MotorNumbers%(chan%) = motor% Then

If sample(1).OrderNumbers%(chan%) = tOrder% Then
nextchan% = chan%
Exit For
End If

End If
End If
Next chan%

Cond2ConditionGetNextChannel% = nextchan%
Exit Function

' Errors
Cond2ConditionGetNextChannelError:
MsgBox Error$, vbOKOnly + vbCritical, "Cond2ConditionGetNextChannel"
ierror = True
Exit Function

End Function

Function Cond2ConditionGetMaxCondition(sample() As TypeSample) As Integer
' Returns the maximum condition order for the sample

ierror = False
On Error GoTo Cond2ConditionGetMaxConditionError

Dim chan As Integer, max_order As Integer

' Search element for max condition number
max_order% = MININTEGER%
For chan% = 1 To sample(1).LastElm%
If sample(1).ConditionNumbers%(chan%) > max_order% Then max_order% = sample(1).ConditionNumbers%(chan%)
Next chan%

Cond2ConditionGetMaxCondition% = max_order%
Exit Function

' Errors
Cond2ConditionGetMaxConditionError:
MsgBox Error$, vbOKOnly + vbCritical, "Cond2ConditionGetMaxCondition"
ierror = True
Exit Function

End Function

Function Cond2ConditionGetMaxCondition2(motor As Integer, sample() As TypeSample) As Integer
' Returns the maximum condition order for the specified spectrometer for the sample

ierror = False
On Error GoTo Cond2ConditionGetMaxCondition2Error

Dim chan As Integer, max_order As Integer

' Search element for max condition number
max_order% = 0
For chan% = 1 To sample(1).LastElm%
If sample(1).MotorNumbers%(chan%) = motor% Then
If sample(1).ConditionNumbers%(chan%) > max_order% Then max_order% = sample(1).ConditionNumbers%(chan%)
End If
Next chan%

Cond2ConditionGetMaxCondition2% = max_order%
Exit Function

' Errors
Cond2ConditionGetMaxCondition2Error:
MsgBox Error$, vbOKOnly + vbCritical, "Cond2ConditionGetMaxCondition2"
ierror = True
Exit Function

End Function

Function Cond2ConditionGetChannel(tCond As Integer, sample() As TypeSample) As Integer
' Returns first channel for the specified condition number (skips disable acquisition elements)

ierror = False
On Error GoTo Cond2ConditionGetChannelError

Dim chan As Integer, nchan As Integer

' Search element for specified condition number
nchan% = 0
For chan% = 1 To sample(1).LastElm%
If sample(1).DisableAcqFlag%(chan%) = 0 Then
If sample(1).ConditionNumbers%(chan%) = tCond% Then
nchan% = chan%
Exit For
End If
End If
Next chan%

Cond2ConditionGetChannel% = nchan%
Exit Function

' Errors
Cond2ConditionGetChannelError:
MsgBox Error$, vbOKOnly + vbCritical, "Cond2ConditionGetChannel"
ierror = True
Exit Function

End Function

Function Cond2ConditionCheckCondition(condition As Integer, wdsdone() As Boolean, sample() As TypeSample) As Boolean
' Returns true if all elements using the specified condition are complete

ierror = False
On Error GoTo Cond2ConditionCheckConditionError

Dim chan As Integer

Cond2ConditionCheckCondition = True

' Loop on channels
For chan% = 1 To sample(1).LastElm%
If sample(1).ConditionNumbers%(chan%) = condition% And Not wdsdone(chan%) Then
Cond2ConditionCheckCondition = False
Exit For
End If
Next chan%

Exit Function

' Errors
Cond2ConditionCheckConditionError:
MsgBox Error$, vbOKOnly + vbCritical, "Cond2ConditionCheckCondition"
ierror = True
Exit Function

End Function

Function CondConditionGetAcquisitionOrder(conditionorder As Integer, sample() As TypeSample) As Integer
' Return the condition number for the specified acquisition condition order

ierror = False
On Error GoTo CondConditionGetAcquisitionOrderError

Dim n As Integer

CondConditionGetAcquisitionOrder% = 1

' Loop on conditions
For n% = 1 To MAXCOND%
If sample(1).ConditionOrders%(n%) = conditionorder% Then
CondConditionGetAcquisitionOrder% = n%
Exit For
End If
Next n%

Exit Function

' Errors
CondConditionGetAcquisitionOrderError:
MsgBox Error$, vbOKOnly + vbCritical, "CondConditionGetAcquisitionOrder"
ierror = True
Exit Function

End Function

Function Cond2ConditionCheckConditionPeaking(tPeakStandardNumber As Integer, condition As Integer, peakingdone() As Boolean, sample() As TypeSample) As Boolean
' Returns true if all peaking elements are peaked for the specified standard (Automate! window)

ierror = False
On Error GoTo Cond2ConditionCheckConditionPeakingError

Dim chan As Integer

Cond2ConditionCheckConditionPeaking = True

' Loop on channels
For chan% = 1 To sample(1).LastElm%
If sample(1).ConditionNumbers%(chan%) = condition% Then
If Not AcquisitionOnAutomate Or (AcquisitionOnAutomate And Not PeakOnAssignedStandardsFlag) Or (AcquisitionOnAutomate And tPeakStandardNumber% = sample(1).StdAssigns%(chan%)) Then
If WavePeakCenterFlags(chan%) And Not peakingdone(chan%) Then
Cond2ConditionCheckConditionPeaking = False
Exit For
End If
End If
End If
Next chan%

Exit Function

' Errors
Cond2ConditionCheckConditionPeakingError:
MsgBox Error$, vbOKOnly + vbCritical, "Cond2ConditionCheckConditionPeaking"
ierror = True
Exit Function

End Function

Function Cond2ConditionIsMotorUsed(condition As Integer, motor As Integer, sample() As TypeSample) As Boolean
' Returns true if this spectrometer is used in the specified condition

ierror = False
On Error GoTo Cond2ConditionIsMotorUsedError

Dim chan As Integer

Cond2ConditionIsMotorUsed = False

' Loop on channels
For chan% = 1 To sample(1).LastElm%
If sample(1).ConditionNumbers%(chan%) = condition% Then
If sample(1).MotorNumbers%(chan%) = motor% Then
Cond2ConditionIsMotorUsed = True
Exit For
End If
End If
Next chan%

Exit Function

' Errors
Cond2ConditionIsMotorUsedError:
MsgBox Error$, vbOKOnly + vbCritical, "Cond2ConditionIsMotorUsed"
ierror = True
Exit Function

End Function

