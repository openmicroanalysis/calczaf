Attribute VB_Name = "CodePenepma12A"
' (c) Copyright 1995-2026 by John J. Donovan
Option Explicit

Const MAXTRIES% = 10

Dim nnum() As Integer        ' binary or pure element sequence number
Dim nrequest() As Integer    ' request flag T/F
Dim nconfirm() As Integer    ' confirm flag T/F
Dim ncomplete() As Integer   ' completion flag T/F

Dim Penepma_TmpSample(1 To 1) As TypeSample

Sub Penepma12CalculateHemisphere(nPoints As Long, xdist() As Double, yinte() As Double, yhemi() As Single)
' This routine calculates the total intensity for a specified radius hemisphere from the constant intensity planer surface interval
'  npoints = the number of points in the Fanal calculation
'  xdist() = the passed array of distances
'  yinte() = the passed array of intensities
'  yhemi() = the returned total intensity for hemispheres of the radii specified in xdist()

ierror = False
On Error GoTo Penepma12CalculateHemisphereError

Dim m As Long

Dim area() As Double
Dim aveinte() As Double

' Just exit for now
Exit Sub

ReDim area(1 To nPoints&) As Double    ' creates array to hold the area of the half plane between xdist(n+1) and xdist(n)
ReDim aveinte(1 To nPoints&) As Double  ' creates an array to hold the average intensity at a point on the plane between xdist(n+1) and xdist(n)

' Set all entries of aveinte to 0
For m& = 1 To nPoints&
aveinte#(m&) = 0#
Next m&

' We need for all x  area(x) = (PID#/2)(xdist(x-1)^2-xdist(x)^2)


' We need for all k aveinte(k) = (yinte(k)-yinte(0) - area(x)*aveinte(x))/area(k) & x must take on all values of area and aveinte arrays

Exit Sub

' Errors
Penepma12CalculateHemisphereError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CalculateHemisphere"
ierror = True
Exit Sub

End Sub

Sub Penepma12CalculateAlphaFactors2(tBinaryRanges() As Single, tBinary_Kratios() As Double, tBinary_Factors() As Single)
' Calculate the alpha factors only for the passed k-ratios
'  tBinaryRanges!(1 to MAXBINARY%)  are the compositional binaries in weight percent
'  tBinary_Kratios#(1 to MAXBINARY%)  are the k-ratios for each x-ray and binary composition
'  tBinary_Factors!(1 to MAXBINARY%)  are the alpha factors for each x-ray and binary composition, alpha = (C/K - C)/(1 - C)

ierror = False
On Error GoTo Penepma12CalculateAlphaFactors2Error

Dim n As Integer
Dim k As Single, c As Single

' Check for reasonable k-ratios (if too far from boundary, garbage intensities are generated)
For n% = 1 To MAXBINARY%
If tBinary_Kratios#(n%) > 0# And tBinary_Kratios#(n%) < 1000# Then

' Calculate alpha factor for this binary composition
c! = tBinaryRanges!(n%) / 100#
k! = tBinary_Kratios#(n%) / 100#    ' k-ratios are in k-ratio percent
tBinary_Factors!(n%) = ((c! / k!) - c!) / (1 - c!)        ' calculate binary alpha factors

If DebugMode Then
If n% = 1 Then
msg$ = vbCrLf & "Calculated alpha factors..."
Call IOWriteLog(msg$)
End If
msg$ = "P=" & Format$(n%) & ", C=" & Format$(c!, f84$) & ", K=" & Format$(k!, f84$) & ", Alpha=" & Format$(tBinary_Factors!(n%), f84$)
Call IOWriteLog(msg$)
End If

' No valid data to calculate, just zero the alpha factor (and k-ratio)
Else
tBinary_Kratios#(n%) = 0#
tBinary_Factors!(n%) = 0#
End If
Next n%

Exit Sub

' Errors
Penepma12CalculateAlphaFactors2Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CalculateAlphaFactors2"
Close #Temp1FileNumber%
Call IOStatusAuto(vbNullString)
ierror = True
Exit Sub

End Sub

Sub Penepma12CalculateGetComposition(mode As Integer, sample() As TypeSample)
' Get a composition
' mode = 1 get formula
' mode = 2 get weight
' mode = 3 get standard composition

ierror = False
On Error GoTo Penepma12CalculateGetCompositionError

Dim i As Integer, ip As Integer
Dim astring As String

ReDim outarray(1 To MAXCHAN%) As Integer
ReDim arrayindex(1 To MAXCHAN%) As Integer

' Write space to log window for new composition
Call IOWriteLog(vbNullString)

' Init sample
Call InitSample(sample())
If ierror Then Exit Sub

' Get formula or weight from user
If mode% = 1 Then FormFORMULA.Show vbModal
If mode% = 2 Then FormWEIGHT.Show vbModal
If mode% = 3 Then FormSTDCOMP.Show vbModal

' If error, just clear and exit
If ierror Then
Call InitSample(sample())
Exit Sub
End If

' Return modified sample
Call FormulaReturnSample(sample())
If ierror Then Exit Sub

' Load atomic numbers
For i% = 1 To sample(1).LastChan%
ip% = IPOS1%(MAXELM%, sample(1).Elsyms$(i%), Symlo$())
If ip% = 0 Then GoTo Penepma12CalculateGetCompositionNotValid
sample(1).AtomicNums%(i%) = AllAtomicNums%(ip%)
Next i%

' Sort elements by atomic number order (for binary calculations)
Call MiscSortIntegerArray(sample(1).LastChan%, sample(1).AtomicNums%(), outarray%(), arrayindex%())
If ierror Then Exit Sub

' Laod into tmp sample
Penepma_TmpSample(1) = sample(1)

' Re-sort
For i% = 1 To sample(1).LastChan%
sample(1).Elsyms$(i%) = Penepma_TmpSample(1).Elsyms$(arrayindex%(i%))
sample(1).ElmPercents!(i%) = Penepma_TmpSample(1).ElmPercents!(arrayindex%(i%))
sample(1).AtomicNums%(i%) = Penepma_TmpSample(1).AtomicNums%(arrayindex%(i%))
Next i%

' Load string
Call IOWriteLog(vbCrLf & vbCrLf & "Material to Calculate: " & Penepma_TmpSample(1).Name$)

For i% = 1 To sample(1).LastChan%
astring$ = astring$ & sample(1).Elsyms$(i%) & MiscAutoFormat$(sample(1).ElmPercents!(i%)) & " "
Next i%

Call IOWriteLog("Composition to Calculate (wt.%): " & astring$)

Exit Sub

' Errors
Penepma12CalculateGetCompositionError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CalculateGetComposition"
ierror = True
Exit Sub

Penepma12CalculateGetCompositionNotValid:
msg$ = sample(1).Elsyms$(i%) & " is not a valid element symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "Penepma12CalculateGetComposition"
ierror = True
Exit Sub

End Sub

Sub Penepma12CalculateBinaryBeta(iray As Integer, ibin As Integer, tRanges() As Single, tZAF_Coeffs() As Single, tZAF_Betas() As Single)
' This routine calculates a beta factor for the specified binary composition
'  tRanges(1 to MAXBINARY%) (always 99 to 1 wt%)
'  tZAF_Coeffs(1 To MAXRAY% - 1, 1 To MAXCOEFF%)
'  tZAF_Betas(1 To MAXRAY% - 1, 1 To MAXBINARY%)

ierror = False
On Error GoTo Penepma12CalculateBinaryBetaError

Dim betafrac As Single

ReDim frac(1 To 2) As Single

' Convert to weight fractions
frac!(1) = tRanges!(ibin%) / 100#
frac!(2) = (100# - tRanges!(ibin%)) / 100#

' Calculate for the emitter only
betafrac! = frac!(1) + (tZAF_Coeffs(iray%, 1) + frac!(2) * tZAF_Coeffs(iray%, 2) + (frac!(2) ^ 2) * tZAF_Coeffs(iray%, 3)) * frac!(2)

' Return beta
tZAF_Betas!(iray%, ibin%) = betafrac!

Exit Sub

' Errors
Penepma12CalculateBinaryBetaError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CalculateBinaryBeta"
ierror = True
Exit Sub

End Sub
