Attribute VB_Name = "CodePenepma12A"
' (c) Copyright 1995-2023 by John J. Donovan
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

Sub Penepma12CalculateRandomCheck(method As Integer, l As Integer, m As Integer, CalculateRandomTable() As Integer, im As Integer, mm As Integer, done As Boolean)
' Check for next available binary or pure element to calculate, request, confirm and complete
' method = 0, create new shared index file (only if file does not exist already, delete manually)
' method = 1, check for next available sequence number
' method = 2  update completion flag as calculation complete
' l = size of first array index for CalculateRandomTable(), always 3 or 2
' m = size of second array index for CalculateRandomTable(), always 4095 or 100
' CalculateRandomTable() = look up table (l = 3)
'   CalculateRandomTable(1, m) = binary 1 index
'   CalculateRandomTable(2, m) = binary 2 index
'   CalculateRandomTable(1, m) = binary number sequence
' CalculateRandomTable() = look up table (l = 2)
'   CalculateRandomTable(1, m) = pure element index
'   CalculateRandomTable(1, m) = pure element sequence number
' im = selected or specified binary sequence or pure element sequence number (returned)
' mm = total number of binary pairs or pure elements already completed (returned)
' done = flag for all sequence numbers indicated as completed or not (returned)

ierror = False
On Error GoTo Penepma12CalculateRandomCheckError

Dim i As Integer
Dim n As Integer

' Create shared index file
If method% = 0 Then
Call Penepma12CalculateRandomCheck2(Int(0), m%, nnum%(), nrequest%(), nconfirm(), ncomplete%(), im%, mm%, done)
If ierror Then Exit Sub
End If

' Check shared index file for pure element or binary compositions and return next available (random) sequence number
If method% = 1 Then
Penepma12CalculateRandomCheckTryAgain:
Call Penepma12CalculateRandomCheck2(Int(1), m%, nnum%(), nrequest%(), nconfirm(), ncomplete%(), im%, mm%, done)
If ierror Then Exit Sub
If done Then Exit Sub

' Not done, so now calculate a random sequence number and see if it is available
im% = 0
For i% = 1 To MAXINTEGER%
n% = Int((m% - 1 + 1) * Rnd() + 1)
If Not ncomplete(n%) Then
im% = n%
GoTo Penepma12CalculateRandomCheckRequest
End If
Next i%

' If we get here it couldn't find an uncompleted calculation at random (just do a straight search)
im% = 0
For n% = 1 To m%
If Not ncomplete(n%) Then
im% = n%
End If
Next n%
If im% > 0 Then GoTo Penepma12CalculateRandomCheckRequest

' If we get here there was a problem
msg$ = "Warning in Penepma12CalculateRandomCheck: the PAR share file shows calculations not completed, but no sequence is available for calculation."
Call IOWriteLog(msg$)
GoTo Penepma12CalculateRandomCheckTryAgain

' Now try to get a confirm on the selected calculation
Penepma12CalculateRandomCheckRequest:
Call Penepma12CalculateRandomCheck2(Int(2), m%, nnum%(), nrequest%(), nconfirm(), ncomplete%(), im%, mm%, done)
If ierror Then Exit Sub
If done Then Exit Sub

' Could not get a confirm on that number, try again
If im% = 0 Then GoTo Penepma12CalculateRandomCheckTryAgain

' Now write a confirm to lock the calculation and return selected calculation
Call Penepma12CalculateRandomCheck2(Int(3), m%, nnum%(), nrequest%(), nconfirm(), ncomplete%(), im%, mm%, done)
If ierror Then Exit Sub
If done Then Exit Sub
If im% > 0 Then Exit Sub

' Could not get a lock on that number, try again
GoTo Penepma12CalculateRandomCheckTryAgain
End If

' Update completion flag for specified sequence number (im%)
If method% = 2 Then
Call Penepma12CalculateRandomCheck2(Int(4), m%, nnum%(), nrequest%(), nconfirm(), ncomplete%(), im%, mm%, done)
If ierror Then Exit Sub
If done Then Exit Sub
End If

Exit Sub

' Errors
Penepma12CalculateRandomCheckError:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CalculateRandomCheck"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma12CalculateRandomCheck2(mode As Integer, m As Integer, nnum() As Integer, nrequest() As Integer, nconfirm() As Integer, ncomplete() As Integer, im As Integer, mm As Integer, done As Boolean)
' Performs various file operations for PAR share file
'  mode = 0  create PAR share file
'  mode = 1  read all values (return done status) (all modes except zero check this)
'  mode = 2  check request and write request if still available
'  mode = 3  check confirm and write confirm if still available
'  mode = 4  write complete

ierror = False
On Error GoTo Penepma12CalculateRandomCheck2Error

Dim n As Integer
Dim tfilename As String
Dim astring As String

Dim irequest As Integer    ' request flag T/F
Dim iconfirm As Integer    ' confirm flag T/F
Dim icomplete As Integer   ' completion flag T/F

ReDim nnum(1 To m%) As Integer        ' sequence number
ReDim nrequest(1 To m%) As Integer    ' request flag T/F
ReDim nconfirm(1 To m%) As Integer    ' confirm flag T/F
ReDim ncomplete(1 To m%) As Integer   ' completion flag T/F

tfilename$ = PENEPMA_PAR_Path$ & "\PAR_share.txt"

' First create the PAR share file if not found
If mode% = 0 Then
If Dir$(tfilename$) = vbNullString Then
Open tfilename$ For Output As #Temp1FileNumber%

' Loop on all sequence numbers
For n% = 1 To m%

' Check PAR share path for exiting PAR files and mark those as complete if found
Call Penepma12CalculateRandomScanPath(n%, irequest%, iconfirm, icomplete%)
If ierror Then Exit Sub

' Create an output string
astring$ = Format$(n%) & vbTab           ' sequence flag
astring$ = astring$ & Format$(irequest%) & vbTab    ' request flag
astring$ = astring$ & Format$(iconfirm%) & vbTab    ' confirm flag
astring$ = astring$ & Format$(icomplete%) & vbTab   ' completion flag

Print #Temp1FileNumber%, astring$
Next n%

Close #Temp1FileNumber%

Else
msg$ = vbCrLf & "The PAR Share file (" & tfilename$ & ") already exists and PAR file calculations may be ongoing. Please check that all programs are terminated or all calculations are completed, then manually delete the PAR Share file and try again."
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbMagenta, Int(FONT_REGULAR%), Int(0))
End If

Exit Sub
End If

' First go through file and load available sequence numbers (return done status)
If mode% >= 1 Then
Call Penepma12CalculateRandomCheckOpen(Int(0), tfilename$)
If ierror Then Exit Sub

For n% = 1 To m%
Input #Temp1FileNumber%, nnum%(n%), nrequest%(n%), nconfirm(n%), ncomplete%(n%)
Next n%
Close #Temp1FileNumber%

' Check if all calculations are complete
done = True
mm% = 0
For n% = 1 To m%
If Not ncomplete(n%) Then
done = False
Else
mm% = mm% + 1
End If
Next n%

' Do not exit as this allows check for completion each time
End If

' Check for next available calculation and write request
If mode% = 2 Then

' Check passed sequence number
If Not nrequest(im%) Then
Call Penepma12CalculateRandomCheckOpen(Int(1), tfilename$)
If ierror Then Exit Sub

nrequest%(im%) = True
For n% = 1 To m%
Print #Temp1FileNumber%, nnum%(n%), nrequest%(n%), nconfirm(n%), ncomplete%(n%)
Next n%
Close #Temp1FileNumber%

' Request not available
Else
im% = 0
End If

Exit Sub
End If


' Wait, then check for confirm and write confirm if ok
If mode% = 3 Then
If Not nconfirm(im%) Then
Call Penepma12CalculateRandomCheckOpen(Int(1), tfilename$)
If ierror Then Exit Sub

nconfirm%(im%) = True
For n% = 1 To m%
Print #Temp1FileNumber%, nnum%(n%), nrequest%(n%), nconfirm(n%), ncomplete%(n%)
Next n%
Close #Temp1FileNumber%

' Confirm not available
Else
im% = 0
End If

Exit Sub
End If

' Write completion flag
If mode% = 4 Then
Call Penepma12CalculateRandomCheckOpen(Int(1), tfilename$)
If ierror Then Exit Sub

ncomplete%(im%) = True
For n% = 1 To m%
Print #Temp1FileNumber%, nnum%(n%), nrequest%(n%), nconfirm(n%), ncomplete%(n%)
Next n%
Close #Temp1FileNumber%

Exit Sub
End If

Exit Sub

' Errors
Penepma12CalculateRandomCheck2Error:
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CalculateRandomCheck2"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma12CalculateRandomCheckOpen(method As Integer, tfilename As String)
' Special routine to open a file and try again if locked
' method = 0  open for Input Shared
' method = 1 open for Output Lock Read Write

ierror = False
On Error GoTo Penepma12CalculateRandomCheckOpenError

Dim ntry As Integer

' Use an On Error statement for trapping if tfilename is already open for exclusive use
On Error GoTo Penepma12CalculateRandomCheckOpenWait

' Open probe database
Penepma12CalculateRandomCheckOpenTryAgain:
ntry% = ntry% + 1
If method% = 0 Then
Open tfilename$ For Input Shared As #Temp1FileNumber%
Else
Open tfilename$ For Output Lock Read Write As #Temp1FileNumber%
End If
GoTo Penepma12CalculateRandomCheckOpenProceed

Penepma12CalculateRandomCheckOpenWait:
Call MiscDelay3(FormMAIN.StatusBarAuto, "for open data file...", CDbl(3#), Now)      ' wait 3 seconds and try again
If ierror Then Exit Sub

' Check for too many tries
If ntry% > MAXTRIES% Then
msg$ = vbCrLf & "Penepma12CalculateRandomCheckOpen: Unable to open file " & tfilename$ & " for method " & Format$(method%) & " after " & Format$(ntry%) & " attempts."
Call IOWriteLogRichText(msg$, vbNullString, Int(LogWindowFontSize%), vbRed, Int(FONT_REGULAR%), Int(0))
ierror = True
Exit Sub
End If

' Try again
Resume Penepma12CalculateRandomCheckOpenTryAgain

' Database opened, go ahead and exit normally
Penepma12CalculateRandomCheckOpenProceed:
On Error GoTo Penepma12CalculateRandomCheckOpenError

FormMAIN.StatusBarAuto.Panels(1).Text = vbNullString
Exit Sub

' Errors
Penepma12CalculateRandomCheckOpenError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CalculateRandomCheckOpen"
Close #Temp1FileNumber%
ierror = True
Exit Sub

End Sub

Sub Penepma12CalculateRandomScanPath(n As Integer, irequest As Integer, iconfirm As Integer, icomplete As Integer)
' Check the path for the indicated PAR file

ierror = False
On Error GoTo Penepma12CalculateRandomScanPathError

Dim i As Integer, j As Integer
Dim tfilename As String

' Create filensame to check
i% = CalculateRandomTable%(1, n%)   ' BinaryElement1
j% = CalculateRandomTable%(2, n%)   ' BinaryElement2

tfilename$ = PENEPMA_PAR_Path$ & "\" & Trim$(Symup$(i%)) & "-" & Trim$(Symup$(j%)) & "_*.par"
If Dir$(tfilename$) <> vbNullString Then
irequest% = True
iconfirm% = True
icomplete% = True
Call IOWriteLog("One or more PAR files for the " & Trim$(Symup$(i%)) & "-" & Trim$(Symup$(j%)) & " binary already exists in the folder " & PENEPMA_PAR_Path$ & " and the calculations will be skipped...")
DoEvents
Else
irequest% = False
iconfirm% = False
icomplete% = False
End If

Exit Sub

' Errors
Penepma12CalculateRandomScanPathError:
Screen.MousePointer = vbDefault
MsgBox Error$, vbOKOnly + vbCritical, "Penepma12CalculateRandomScanPath"
Close #Temp1FileNumber%
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
