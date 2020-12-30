Attribute VB_Name = "CodeLeast3"
' (c) Copyright 1995-2021 by John J. Donovan
Option Explicit

Sub LeastSmoothSavitzkyGolay(k As Integer, npts As Integer, aydata() As Single, tydata() As Single)
' Smooth passed data using Savitsky-Golay smoothing
' k% is the number of points on each side to smooth (must be an even number)
' npts% = number of data to be smoothed
' aydata() = input data
' tydata() = smoothed data

ierror = False
On Error GoTo LeastSmoothSavitzkyGolayError

Const MAXPOWER2% = 32

Dim i As Integer

' Returned filter coefficients
Dim np As Long   ' number of returned filter coefficients (nl% + nr% + 1)
Dim nl As Integer, nr As Integer    ' number of left and rightward data points used
Dim ld As Integer   ' order of the derivative (ld = 0 for smooth function)
Dim m As Integer    ' order of the smoothing polynomial (usually 2 or 4)

Dim n As Long, l As Long
Dim isign As Long    ' 1 for convolve, -1 for deconvolve

Dim respns() As Double, ans() As Double

' Check that k% is even
If k% Mod 2 <> 0 Then GoTo LeastSmoothSavitzkyGolayBadM

' Dimension filter coefficient array
ReDim respns(1 To npts%) As Double

' Calculate smoothing coefficients
nl% = k%                 ' number of points on left side
nr% = k%                 ' number of points on right side
ld% = 0                 ' order of derivative desired (0 for smoothed function)
m% = 2                  ' order of the smoothing polynomial (usually 2 or 4)
np& = nl% + nr% + 1     ' should be an odd number

' Check for minimum number of points
If npts% < np& Then GoTo LeastSmoothSavitzkyGolayBadPoints

' Get smoothing coefficients
Call LeastSmoothSavitzkyGolay2(respns#(), np&, nl%, nr%, ld%, m%)
If ierror Then Exit Sub

' Print out SG coefficients
If DebugMode And VerboseMode Then
msg$ = "Stavizky-Golay coefficients (m=" & Str$(m%) & ", nl=" & Str$(nl%) & ", nl=" & Str$(nr%) & "): "
For i% = 1 To np&
msg$ = msg$ & Str$(respns#(i%))
If i% <> np& Then msg$ = msg$ & ", "
Next i%
Call IOWriteLog(msg$)
End If

' Calculate closest power of two for array size
For l& = 1 To MAXPOWER2%
If npts% < 2 ^ l& Then Exit For
Next l&
If l& = MAXPOWER2% Then GoTo LeastSmoothSavitzkyGolayBadPower
n& = 2 ^ l&

' Dimension data to be convolved into a power of two size array
ReDim tydata(1 To n&) As Single
ReDim ans(1 To n&) As Double
ReDim Preserve respns(1 To n&) As Double

' Load data into power of 2 sized temp array
For i% = 1 To npts%
tydata!(i%) = aydata!(i%)
Next i%

' Smooth passed data
isign& = 1
Call LeastSmoothConvolve(tydata!(), n&, respns#(), np&, isign&, ans#())
If ierror Then Exit Sub

' Load smoothed data back into return array (do not overload end points)
For i% = 1 To npts%
If i% > k% And i% < npts% - k% Then tydata!(i%) = ans#(i%)
Next i%

Exit Sub

' Errors
LeastSmoothSavitzkyGolayError:
MsgBox Error$, vbOK + vbCritical, "LeastSmoothSavitzkyGolay"
ierror = True
Exit Sub

LeastSmoothSavitzkyGolayBadM:
msg$ = "Number of points on each side to smooth must be an even number"
MsgBox msg$, vbOK + vbExclamation, "LeastSmoothSavitzkyGolay"
ierror = True
Exit Sub

LeastSmoothSavitzkyGolayBadPower:
msg$ = "Bad power of two in convolution initialization"
MsgBox msg$, vbOK + vbExclamation, "LeastSmoothSavitzkyGolay"
ierror = True
Exit Sub

LeastSmoothSavitzkyGolayBadPoints:
msg$ = "Insufficient number of data points to obtain smoothing coefficients for based on specified Savizky-Golay smoothing parameters"
MsgBox msg$, vbOK + vbExclamation, "LeastSmoothSavitzkyGolay"
ierror = True
Exit Sub

End Sub

Sub LeastSmoothSavitzkyGolay2(respns() As Double, np As Long, nl As Integer, nr As Integer, ld As Integer, m As Integer)
' Perform the actual Savitsky-Golay coefficient calculation
' From Numerical Recipes

ierror = False
On Error GoTo LeastSmoothSavitzkyGolay2Error

Const mmax% = 6

Dim imj As Integer, ipj As Integer
Dim j As Integer, k As Integer, kk As Integer, mm As Integer
Dim fac As Double, sum As Double
Dim d As Double

ReDim indx(mmax + 1) As Integer
ReDim a(mmax + 1, mmax + 1) As Double, b(mmax + 1) As Double

' Check for valid dimensions
If np < nl + nr + 1 Or nl < 0 Or nr < 0 Or ld > m Or m > mmax Or nl + nr < m Then
MsgBox "Invalid integer arguments", vbOKCancel + vbExclamation, "LeastSmoothSaviskyGolay2"
ierror = True
Exit Sub
End If
     
' Set up normal equations of desired least squares fit
For ipj% = 0 To 2 * m%
sum# = 0#
If ipj = 0 Then sum# = 1#

    For k% = 1 To nr%
    sum# = sum# + CDbl(k%) ^ ipj%
    Next k%
    
    For k% = 1 To nl%
    sum# = sum# + CDbl(-k%) ^ ipj%
    Next k%
    
    If ipj% < 2 * m% - ipj% Then
    mm% = ipj%
    Else
    mm% = 2 * m% - ipj%
    End If
    
    For imj% = -mm% To mm% Step 2
    a#(1 + (ipj% + imj%) / 2, 1 + (ipj% - imj%) / 2) = sum#
    Next imj%
Next ipj%

' Solve them
Call Plan3dLUDCMP(a#(), m% + 1, Int(mmax% + 1), indx%(), d#)
If ierror Then Exit Sub

For j% = 1 To m% + 1
b#(j%) = 0#
Next j%
    
' Right hand side vector is unit vector depending on which derivative we want
b#(ld% + 1) = 1#
    
' Back substitute, giving one row of the inverse matrix
Call Plan3dLUBKSB(a#(), m% + 1, Int(mmax% + 1), indx%(), b#())
If ierror Then Exit Sub

' Zero the output array (it may be bigger than the number of coefficients)
For kk% = 1 To np&
respns#(kk%) = 0#
Next kk%

' Each Savitzky-Golay coefficient is the dot product of power of an integer with the inverse matrix row
For k% = -nl% To nr%
sum# = b#(1)
fac# = 1#

For mm% = 1 To m%
fac# = fac# * k%
sum# = sum# + b#(mm% + 1) * fac#
Next mm%

' Store in wrap around order
kk% = ((np& - k%) Mod np&) + 1
respns#(kk%) = sum#
Next k%

Exit Sub

' Errors
LeastSmoothSavitzkyGolay2Error:
MsgBox Error$, vbOK + vbCritical, "LeastSmoothSavitzkyGolay2"
ierror = True
Exit Sub

End Sub

Sub LeastSmoothConvolve(tydata() As Single, n As Long, respns() As Double, m As Long, isign As Long, ans() As Double)
' Convolve (smooth) the data
'  tydata!() to be convolved
'  n& = size of data to be convolved (n must be an integer power of two- zero padded)
'  respns#() smoothing filter coefficients
'  m& = wrap around array size for respns!()
'  isign& = flag (1 = convolution, -1 = deconvolution)
'  ans#() = returned (convolved or deconvolved data)
' From Numerical Recipes

ierror = False
On Error GoTo LeastSmoothConvolveError

Dim i As Long, no2 As Long
Dim dum As Double, ans1 As Double

ReDim FFT(2 * n) As Double
ReDim ans(2 * n) As Double

' Pur respns in an array of length n
For i = 1 To Int((m - 1) / 2)
  respns(n + 1 - i) = respns(m + 1 - i)
Next i

' Pad with zeros
For i = (m + 3) / 2 To n - (m - 1) / 2
  respns(i) = 0!
Next i

' FFT both at once
Call LeastFFTTWOFFT(tydata!(), respns#(), FFT#(), ans#(), n&)
If ierror Then Exit Sub

' Multiple FFTs to convolve, divide to deconvolve
no2 = Int(n / 2)
For i = 1 To no2 + 1
  If isign = 1 Then
    dum = ans(2 * i - 1)
    ans(2 * i - 1) = (FFT(2 * i - 1) * dum - FFT(2 * i) * ans(2 * i)) / no2
    ans(2 * i) = (FFT(2 * i - 1) * ans(2 * i) + FFT(2 * i) * dum) / no2
  
  ElseIf isign = -1 Then
    If dum = 0! And ans(2 * i) = 0! Then
    MsgBox "Deconvolving at response zero", vbOKCancel + vbExclamation, "LeastSmoothConvolve"
    ierror = True
    Exit Sub
    End If
    
    ans1 = FFT(2 * i - 1) * dum + FFT(2 * i) * ans(2 * i)
    ans(2 * i - 1) = ans1 / (dum * dum + ans(2 * i) * ans(2 * i)) / no2
    ans1 = FFT(2 * i) * dum - FFT(2 * i - 1) * ans(2 * i)
    ans(2 * i) = ans1 / (dum * dum + ans(2 * i) * ans(2 * i)) / no2
  
  Else
    MsgBox "No meaning for parameter isign (not defined)", vbOKCancel + vbExclamation, "LeastSmoothConvolve"
    ierror = True
    Exit Sub
  End If
Next i

' Calculate the FFT(pack last element with first for REALFT)
ans(2) = ans(2 * no2 + 1)
Call LeastFFTREALFT(ans(), no2, CLng(-1))
If ierror Then Exit Sub

Exit Sub

' Errors
LeastSmoothConvolveError:
MsgBox Error$, vbOK + vbCritical, "LeastSmoothConvolve"
ierror = True
Exit Sub

End Sub

Sub LeastFFTTWOFFT(DATA1() As Single, DATA2() As Double, FFT1() As Double, FFT2() As Double, n As Long)
' Routine to calculate two FFTs
' From Numerical Recipes

ierror = False
On Error GoTo LeastFFTTWOFFTError

Dim j As Long, n2 As Long, j2 As Long
Dim C1R As Double, C1I As Double, C2R As Double, C2I As Double
Dim CONJR As Double, CONJI As Double
Dim H1R As Double, H1I As Double, H2R As Double, H2I As Double

C1R = 0.5
C1I = 0!
C2R = 0!
C2I = -0.5
For j = 1 To n
  FFT1(2 * j - 1) = DATA1(j)
  FFT1(2 * j) = DATA2(j)
Next j

Call LeastFFTFOUR1(FFT1#(), n&, CLng(1))
If ierror Then Exit Sub

FFT2(1) = FFT1(2)
FFT2(2) = 0!
FFT1(2) = 0!
n2 = 2 * (n + 2)
For j = 2 To n / 2 + 1
  j2 = 2 * j
  CONJR = FFT1(n2 - j2 - 1)
  CONJI = -FFT1(n2 - j2)
  H1R = C1R * (FFT1(j2 - 1) + CONJR) - C1I * (FFT1(j2) + CONJI)
  H1I = C1I * (FFT1(j2 - 1) + CONJR) + C1R * (FFT1(j2) + CONJI)
  H2R = C2R * (FFT1(j2 - 1) - CONJR) - C2I * (FFT1(j2) - CONJI)
  H2I = C2I * (FFT1(j2 - 1) - CONJR) + C2R * (FFT1(j2) - CONJI)
  FFT1(j2 - 1) = H1R
  FFT1(j2) = H1I
  FFT1(n2 - j2 - 1) = H1R
  FFT1(n2 - j2) = -H1I
  
  FFT2(j2 - 1) = H2R
  FFT2(j2) = H2I
  FFT2(n2 - j2 - 1) = H2R
  FFT2(n2 - j2) = -H2I
Next j

Exit Sub

' Errors
LeastFFTTWOFFTError:
MsgBox Error$, vbOK + vbCritical, "LeastFFTTWOFFT"
ierror = True
Exit Sub

End Sub

Sub LeastFFTFOUR1(DATQ() As Double, nn As Long, isign As Long)
' From Numerical Recipes

ierror = False
On Error GoTo LeastFFTFOUR1Error

Dim mmax As Long
Dim n As Long, i As Long, m As Long, j As Long, istep As Long
Dim tempr As Double, tempi As Double
Dim theta As Double, wtemp As Double
Dim WPR As Double, WPI As Double, WR As Double, WI As Double

n = 2 * nn
j = 1
For i = 1 To n Step 2
  If j > i Then
    tempr = DATQ(j)
    tempi = DATQ(j + 1)
    DATQ(j) = DATQ(i)
    DATQ(j + 1) = DATQ(i + 1)
    DATQ(i) = tempr
    DATQ(i + 1) = tempi
  End If
  m = Int(n / 2)
  While m >= 2 And j > m
    j = j - m
    m = Int(m / 2)
  Wend
  j = j + m
Next i

mmax = 2
While n > mmax
  istep = 2 * mmax
  theta# = 6.28318530717959 / (isign * mmax)
  WPR# = -2# * Sin(0.5 * theta#) ^ 2
  WPI# = Sin(theta#)
  WR# = 1#
  WI# = 0#
  For m = 1 To mmax Step 2
    For i = m To n Step istep
      j = i + mmax
      tempr = WR# * DATQ(j) - WI# * DATQ(j + 1)
      tempi = WR# * DATQ(j + 1) + WI# * DATQ(j)
      DATQ(j) = DATQ(i) - tempr
      DATQ(j + 1) = DATQ(i + 1) - tempi
      DATQ(i) = DATQ(i) + tempr
      DATQ(i + 1) = DATQ(i + 1) + tempi
    Next i
    wtemp# = WR#
    WR# = WR# * WPR# - WI# * WPI# + WR#
    WI# = WI# * WPR# + wtemp# * WPI# + WI#
  Next m
  mmax = istep
Wend

Exit Sub

' Errors
LeastFFTFOUR1Error:
MsgBox Error$, vbOK + vbCritical, "LeastFFTFOUR1"
ierror = True
Exit Sub

End Sub

Sub LeastFFTREALFT(DATQ() As Double, n As Long, isign As Long)
' From Numerical Recipes

ierror = False
On Error GoTo LeastFFTREALFTError

Dim N2P3 As Long, i As Long
Dim i1 As Long, i2 As Long, i3 As Long, i4 As Long
Dim c1 As Double, c2 As Double
Dim theta As Double, wtemp As Double
Dim WPR As Double, WPI As Double, WR As Double, WI As Double
Dim H1R As Double, H1I As Double, H2R As Double, H2I As Double
Dim WRS As Double, WIS As Double

theta# = 3.14159265358979 / CDbl(n)
c1 = 0.5
If isign = 1 Then
  c2 = -0.5
  Call LeastFFTFOUR1(DATQ#(), n&, CLng(1))
  If ierror Then Exit Sub
Else
  c2 = 0.5
  theta# = -theta#
End If

WPR# = -2# * Sin(0.5 * theta#) ^ 2
WPI# = Sin(theta#)
WR# = 1# + WPR#
WI# = WPI#
N2P3 = 2 * n + 3
For i = 2 To Int(n / 2)
  i1 = 2 * i - 1
  i2 = i1 + 1
  i3 = N2P3 - i2
  i4 = i3 + 1
  WRS# = WR#
  WIS# = WI#
  H1R = c1 * (DATQ(i1) + DATQ(i3))
  H1I = c1 * (DATQ(i2) - DATQ(i4))
  H2R = -c2 * (DATQ(i2) + DATQ(i4))
  H2I = c2 * (DATQ(i1) - DATQ(i3))
  DATQ(i1) = H1R + WRS# * H2R - WIS# * H2I
  DATQ(i2) = H1I + WRS# * H2I + WIS# * H2R
  DATQ(i3) = H1R - WRS# * H2R + WIS# * H2I
  DATQ(i4) = -H1I + WRS# * H2I + WIS# * H2R
  wtemp# = WR#
  WR# = WR# * WPR# - WI# * WPI# + WR#
  WI# = WI# * WPR# + wtemp# * WPI# + WI#
Next i

If isign = 1 Then
  H1R = DATQ(1)
  DATQ(1) = H1R + DATQ(2)
  DATQ(2) = H1R - DATQ(2)
Else
  H1R = DATQ(1)
  DATQ(1) = c1 * (H1R + DATQ(2))
  DATQ(2) = c1 * (H1R - DATQ(2))
  
  Call LeastFFTFOUR1(DATQ#(), n&, CLng(-1))
  If ierror Then Exit Sub
End If

Exit Sub

' Errors
LeastFFTREALFTError:
MsgBox Error$, vbOK + vbCritical, "LeastFFTREALFT"
ierror = True
Exit Sub

End Sub

