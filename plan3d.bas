Attribute VB_Name = "CodePLAN3D"
' (c) Copyright 1995-2023 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub Plan3dCalculate(nPoints As Integer, dxdata() As Single, dydata() As Single, dzdata() As Single, acoeff() As Single, d As Double)
' This routine will derive an equation for a plane fit, through a matrix
' of points in 3 dimensions. This routine is called from Test3d. Array "acoeff(1 to 3)",
' is the returned equation of fit for the plane.
'  zpos  =  acoeff(1) + xpos * acoeff(2) + ypos * acoeff(3)

ierror = False
On Error GoTo Plan3dCalculateError

Dim i As Integer, j As Integer, k As Integer

ReDim f(1 To MAXDIM%, 1 To MAXDIM%) As Double
ReDim FInv(1 To MAXDIM%, 1 To MAXDIM%) As Double
ReDim x(1 To nPoints%, 1 To MAXDIM%) As Double
ReDim xt(1 To MAXDIM%, 1 To nPoints%) As Double
ReDim temp(1 To MAXDIM%, 1 To nPoints%) As Double

' Load dxdata and dydata into design matrix
If nPoints% < 3 Then GoTo Plan3dcalculateTooFewPoints
For i% = 1 To nPoints%
x#(i%, 1) = 1#
x#(i%, 2) = dxdata!(i%)
x#(i%, 3) = dydata!(i%)
Next i%

' Transpose design matrix
For i% = 1 To nPoints%
For j% = 1 To MAXDIM%
xt#(j%, i%) = x#(i%, j%)
Next j%
Next i%

' Multiply array "xt" times array "x" and load into array "f"
For i% = 1 To MAXDIM%
For j% = 1 To MAXDIM%
f#(i%, j%) = 0#
For k% = 1 To nPoints%
f#(i%, j%) = f#(i%, j%) + xt#(i%, k%) * x#(k%, j%)
Next k%
Next j%
Next i%

' Get inverse of matrix array "f"
Call Plan3dInvertMatrix(f#(), MAXDIM%, MAXDIM%, FInv#(), d#)
If ierror Then Exit Sub

' Multiply array "finv" times array "xt" and load into array "temp"
For i% = 1 To MAXDIM%
For j% = 1 To nPoints%
temp#(i%, j%) = 0#
For k% = 1 To MAXDIM%
temp#(i%, j%) = temp#(i%, j%) + FInv#(i%, k%) * xt#(k%, j%)
Next k%
Next j%
Next i%

' Multiply array "temp" times array "dzdata" and get coefficients for "acoeff"
For i% = 1 To MAXDIM%
acoeff!(i%) = 0#
For k% = 1 To nPoints%
acoeff!(i%) = acoeff!(i%) + temp#(i%, k%) * dzdata!(k%)
Next k%
Next i%

Exit Sub

' Errors
Plan3dCalculateError:
MsgBox Error$, vbOKOnly + vbCritical, "Plan3dCalculate"
ierror = True
Exit Sub

Plan3dcalculateTooFewPoints:
msg$ = "Too few points to calculate a plane fit"
MsgBox msg$, vbOKOnly + vbExclamation, "Plan3dCalculate"
ierror = True
Exit Sub

End Sub

Function Plan3dCalculateSD(nPoints As Integer, dxdata() As Single, dydata() As Single, dzdata() As Single, acoeff() As Single, d As Double) As String
' Return a text string of standard deviation fit

ierror = False
On Error GoTo Plan3dCalculateSDError

Dim i As Integer
Dim sd As Single, zpos As Single
Dim tmsg As String

' Load fit into text field
Plan3dCalculateSD$ = vbNullString
tmsg$ = "Determinate: " & Str$(d#) & vbCrLf
tmsg$ = tmsg$ & "Fit Coefficients: " & Str$(acoeff(1)) & " " & Str$(acoeff(2)) & " " & Str$(acoeff(3))

' Calculate standard deviation
sd! = 0#
For i% = 1 To nPoints%
zpos! = acoeff(1) + dxdata!(i%) * acoeff!(2) + dydata!(i%) * acoeff!(3)
sd! = sd! + (dzdata!(i%) - zpos!) * (dzdata!(i%) - zpos!)
Next i%
sd! = sd! / (nPoints% - 1)

tmsg$ = tmsg$ & vbCrLf & "Standard deviation: " & Str$(sd!)
Plan3dCalculateSD$ = tmsg$

Exit Function

' Errors
Plan3dCalculateSDError:
MsgBox Error$, vbOKOnly + vbCritical, "Plan3dCalculateSD"
ierror = True
Exit Function

End Function

Sub Plan3dCalculateTilt(acoeff() As Single, tilt As Single, astring As String)
' Routine to calculate the tilt of a specimen based on 3 or more points. Called to warn user of excessive specimen tilt for
' polygon defined or fiducial referenced digitized samples.

ierror = False
On Error GoTo Plan3dCalculateTiltError

Const dpr! = 0.017453292                    ' degrees per radian (2*pi/360)

Dim thetax As Single, thetay As Single, theta As Single

' Set defaults
tilt! = 0#
astring$ = vbNullString

' Calculate the maximum tilt in degrees. For small angles, assume that: theta = Sqr(thetax**2 + thetay**2)
If acoeff!(2) < -1# Or acoeff!(2) > 1# Or acoeff!(3) < -1# Or acoeff!(3) > 1# Then
msg$ = "Warning: Zero or infinite tilt, cannot calculate specimen tilt for given equation"
MsgBox msg$, vbOKOnly + vbExclamation, "Plan3dCalculateTilt"
Exit Sub
End If

thetax! = Sin(acoeff!(2))
thetay! = Sin(acoeff!(3))
theta! = Sqr(thetax! ^ 2 + thetay! ^ 2)
tilt! = theta! / dpr!

astring$ = "Specimen tilt in radians: " & vbCrLf
astring$ = astring$ & "ThetaX = " & Str$(thetax!) & " ThetaY= " & Str$(thetay!) & " Theta= " & Str$(theta!) & vbCrLf & vbCrLf
astring$ = astring$ & "Specimen tilt in degrees: " & vbCrLf
astring$ = astring$ & "ThetaX = " & Str$(thetax! / dpr!) & " ThetaY= " & Str$(thetay! / dpr!) & " Theta= " & Str$(tilt!)
If Abs(tilt!) > 0.5 Then
astring$ = astring$ & vbCrLf & vbCrLf & "WARNING: Specimen tilt exceeds 0.5 degree"
End If

Exit Sub

' Errors
Plan3dCalculateTiltError:
MsgBox Error$, vbOKOnly + vbCritical, "Plan3dCalculateTilt"
ierror = True
Exit Sub

End Sub

Function Plan3dCRS(x11 As Single, y11 As Single, x12 As Single, y12 As Single, x21 As Single, y21 As Single, x22 As Single, y22 As Single, xt As Single, yt As Single) As Integer
' This routine determines if two segments cross. It receives the four
' points that define the two segments.

' If two segments have the same slope, even if they have the
' same intercept, they are assumed not to cross.

' The segments are closed at one end (the first point)
' and open at the other so that if a segment goes
' exactly through an end point it is only counted once
' on successive calculations.

' Segment endpoint nomenclature:
'  e.g., x21 = x-coordinate of first point of second segment
        
ierror = False
On Error GoTo Plan3dCRSError

Dim a1 As Single, b1 As Single, a2 As Single, b2 As Single
Dim x As Single, y As Single
Dim f1 As Single, f2 As Single

a1! = 0#
b1! = 0#
a2! = 0#
b2! = 0#

x! = 0#
y! = 0#

f1! = 0#
f2! = 0#

' Normalized position of intersection along segments:
'  f = 0.0 means intersection on first point
'  f = 1.0 means intersection on second point
'  0.0 <= f < 1.0 means intersection within segment

If x11! = x12! And x21! = x22! Then GoTo 50
If x11! = x12! Then GoTo 100
If x21! = x22! Then GoTo 110

' Neither slope is infinite:  no pathologies
If x11! - x12! = 0# Then GoTo 50
a1! = (y11! - y12!) / (x11! - x12!)
b1! = y11! - a1! * x11!

If x21! - x22! = 0# Then GoTo 50
a2! = (y21! - y22!) / (x21! - x22!)
b2! = y21! - a2! * x21!

If a1! - a2! = 0# Then GoTo 50
x! = (b2! - b1!) / (a1! - a2!)

f1! = (x! - x11!) / (x12! - x11!)
f2! = (x! - x21!) / (x22! - x21!)

Plan3dCRS = ((f1! < 1#) And (f1! >= 0#) And (f2! < 1#) And (f2! >= 0#))

' Calculate intersection of segments
xt! = -(b1! - b2!) / (a1! - a2!)
yt! = a1! * xt! + b1!

Exit Function

' Slopes are equal  - no intersection even if intercepts are equal
50:
Plan3dCRS = False
Exit Function

' Slope of first segment is infinite
100:
x! = x11!
If x21! - x22! = 0# Then GoTo 50
a2! = (y21! - y22!) / (x21! - x22!)
b2! = y21! - a2! * x21!
y! = a2! * x! + b2!

If y12! - y11! = 0# Then GoTo 50
f1! = (y! - y11!) / (y12! - y11!)
If x22! - x21! = 0# Then GoTo 50
f2! = (x! - x21!) / (x22! - x21!)

Plan3dCRS = (f1! < 1#) And (f1! >= 0#) And (f2! < 1#) And (f2! >= 0#)

' Calculate intersection of segments
xt! = x11!
yt! = a1! * xt! + b1!
Exit Function

' Slope of second segment is infinite
110:
x! = x21!
If x11! - x12! = 0# Then GoTo 50
a1! = (y11! - y12!) / (x11! - x12!)
b1! = y11! - a1! * x11!

y! = a1! * x! + b1!
If x12! - x11! = 0# Then GoTo 50
f1! = (x! - x11!) / (x12! - x11!)
If y22! - y21! = 0# Then GoTo 50
f2! = (y! - y21!) / (y22! - y21!)

Plan3dCRS = (f1! < 1#) And (f1! >= 0#) And (f2! < 1#) And (f2! >= 0#)

' Calculate intersection of segments
xt! = x21!
yt! = a1! * xt! + b1!

Exit Function

' Errors
Plan3dCRSError:
MsgBox Error$, vbOKOnly + vbCritical, "Plan3dCRS"
ierror = True
Exit Function

End Function

Sub Plan3dGetExteriorPoints(nPoints As Integer, dxdata() As Single, dydata() As Single, xext As Single, yext As Single)
' Determine an exterior x and y coordinate

ierror = False
On Error GoTo Plan3dGetExteriorPointsError

Dim i As Integer
Dim xmin As Single, ymin As Single
Dim xmax As Single, ymax As Single

' Find min and max
xmin! = 1E+20
ymin! = 1E+20
xmax! = -1E+20
ymax! = -1E+20
For i% = 1 To nPoints%
If dxdata!(i%) < xmin! Then xmin! = dxdata!(i%)
If dydata!(i%) < ymin! Then ymin! = dydata!(i%)
If dxdata!(i%) > xmax! Then xmax! = dxdata!(i%)
If dydata!(i%) > ymax! Then ymax! = dydata!(i%)
Next i%

' Add 10 times the X and Y distance to obtain an exterior point
xext! = xmax! + (xmax! - xmin!) * 10#
yext! = ymax! + (ymax! - ymin!) * 10#

Exit Sub

' Errors
Plan3dGetExteriorPointsError:
MsgBox Error$, vbOKOnly + vbCritical, "Plan3dGetExteriorPoints"
ierror = True
Exit Sub

End Sub

Function Plan3DIsOutSide(xpos As Single, ypos As Single, xext As Single, yext As Single, np As Integer, tX() As Single, tY() As Single) As Integer
' This routine determines if a point is inside or outside of an area defined by a list of segments in tx!(np%), ty!(np%).
'  xpos and ypos is the position to be tested
'  xext and yext is an X/Y coordinate that is required to be outside the polygon

ierror = False
On Error GoTo Plan3dIsOutSideError

Dim ncross As Integer, i As Integer
Dim xt As Single, yt As Single

' Check for valid polygon
If np% < 3 Then GoTo Plan3dIsOutSideBadPolygon

' Calculate number of segment crossings of line from grid point (xpos, ypos) to exterior point (xext, yext)
ncross% = 0
For i% = 1 To np% - 1

If Plan3dCRS(xpos!, ypos!, xext!, yext!, tX!(i%), tY!(i%), tX!(i% + 1), tY!(i% + 1), xt!, yt!) Then ncross% = ncross% + 1
If ierror Then Exit Function

Next i%

If Plan3dCRS(xpos!, ypos!, xext!, yext!, tX!(np%), tY!(np%), tX!(1), tY!(1), xt!, yt!) Then ncross% = ncross% + 1
If ierror Then Exit Function

' Check for odd or even number of crossings (odd = inside and even = outside)
If ncross% Mod 2 <> 0 Then
Plan3DIsOutSide% = False
Else
Plan3DIsOutSide% = True
End If

If DebugMode Then
msg$ = vbCrLf & "Plan3dIsOutside Crossings = " & Str$(ncross%) & Str$(Plan3DIsOutSide%)
Call IOWriteLog(msg$)
msg$ = "Test point = " & Str$(xpos!) & Str$(ypos!)
Call IOWriteLog(msg$)
msg$ = "Exterior point = " & Str$(xext!) & Str$(yext!)
Call IOWriteLog(msg$)
For i% = 1 To np%
msg$ = "Polygon coordinate " & Str$(i%) & " = " & Str$(tX!(i%)) & Str$(tY!(i%))
Call IOWriteLog(msg$)
Next i%
End If

Exit Function

' Errors
Plan3dIsOutSideError:
MsgBox Error$, vbOKOnly + vbCritical, "Plan3dIsOutSide"
ierror = True
Exit Function

Plan3dIsOutSideBadPolygon:
msg$ = "Need more than two points to define a polygon"
MsgBox msg$, vbOKOnly + vbExclamation, "Plan3dIsOutSide"
ierror = True
Exit Function

End Function

Sub Plan3dCalculateHyper(nPoints As Integer, dxdata() As Single, dydata() As Single, dzdata() As Single, acoeff() As Single, d As Double)
' This routine will derive an equation for a hyperbolic fit, through a matrix of points in 3 dimensions. This routine is called
' from Test3d. Array "acoeff()", is the returned equation of fit for the hyperbolic surface.

ierror = False
On Error GoTo Plan3dCalculateHyperError

Dim i As Integer, j As Integer, k As Integer

ReDim f(1 To MAXCOEFF9%, 1 To MAXCOEFF9%) As Double
ReDim FInv(1 To MAXCOEFF9%, 1 To MAXCOEFF9%) As Double
ReDim x(1 To nPoints%, 1 To MAXCOEFF9%) As Double
ReDim xt(1 To MAXCOEFF9%, 1 To nPoints%) As Double
ReDim temp(1 To MAXCOEFF9%, 1 To nPoints%) As Double

' Load dxdata and dydata into design matrix
If nPoints% < MAXCOEFF9% Then GoTo Plan3dCalculateHyperTooFewPoints

' Use hyperbolic fit method
For i% = 1 To nPoints%
x#(i%, 1) = 1#
x#(i%, 2) = dxdata!(i%)
x#(i%, 3) = dydata!(i%)
x#(i%, 4) = dxdata!(i%) ^ 2
x#(i%, 5) = dydata!(i%) ^ 2
x#(i%, 6) = dxdata!(i%) ^ 3
x#(i%, 7) = dydata!(i%) ^ 3
x#(i%, 8) = (dxdata!(i%) * dydata!(i%)) ^ 2
x#(i%, 9) = (dxdata!(i%) * dydata!(i%)) ^ 3
Next i%

' Transpose design matrix
For i% = 1 To nPoints%
For j% = 1 To MAXCOEFF9%
xt#(j%, i%) = x#(i%, j%)
Next j%
Next i%

' Multiply array "xt" times array "x" and load into array "f"
For i% = 1 To MAXCOEFF9%
For j% = 1 To MAXCOEFF9%
f#(i%, j%) = 0#
For k% = 1 To nPoints%
f#(i%, j%) = f#(i%, j%) + xt#(i%, k%) * x#(k%, j%)
Next k%
Next j%
Next i%

' Get inverse of matrix array "f"
Call Plan3dInvertMatrix(f#(), MAXCOEFF9%, MAXCOEFF9%, FInv#(), d#)
If ierror Then Exit Sub

' Multiply array "finv" times array "xt" and load into array "temp"
For i% = 1 To MAXCOEFF9%
For j% = 1 To nPoints%
temp#(i%, j%) = 0#
For k% = 1 To MAXCOEFF9%
temp#(i%, j%) = temp#(i%, j%) + FInv#(i%, k%) * xt#(k%, j%)
Next k%
Next j%
Next i%

' Multiply array "temp" times array "dzdata" and get coefficients for "acoeff"
For i% = 1 To MAXCOEFF9%
acoeff!(i%) = 0#
For k% = 1 To nPoints%
acoeff!(i%) = acoeff!(i%) + temp#(i%, k%) * dzdata!(k%)
Next k%
Next i%

Exit Sub

' Errors
Plan3dCalculateHyperError:
MsgBox Error$, vbOKOnly + vbCritical, "Plan3dCalculateHyper"
ierror = True
Exit Sub

Plan3dCalculateHyperTooFewPoints:
msg$ = "Too few points to calculate a hyperbolic polynomial fit"
MsgBox msg$, vbOKOnly + vbExclamation, "Plan3dCalculateHyper"
ierror = True
Exit Sub

End Sub

Function Plan3DCalculateEffectiveTakeOff(tSpectrometerOrientation As Single, acoeff() As Single, tEffectiveTakeOff As Single) As Single
' Function to calculate a modified effective take off angle based on the passed
'  spectrometer orientation (0 degrees = north and going clockwise looking down on the top of the instrument)
'  stage orientation (0 = cartesian, -1 = anti-cartesian)
'  sample tilt coefficients based on fit to stage X/Y/Z point coordinates
'
'  returns the modified effective takeoff angle
'  zpos  =  acoeff(1) + xpos * acoeff(2) + ypos * acoeff(3)
        
    ierror = False
    On Error GoTo Plan3DCalculateEffectiveTakeOffError
    
    Dim v_x As Single, v_y As Single, v_z As Single 'direction vector of the line (v) sustained by the spectrometer
    Dim n_x As Single, n_y As Single, n_z As Single 'normal vector of the plane (n)
    Dim phi_in_deg As Single
    Dim theta_in_rad As Single
    Dim phi_in_rad As Single
    Dim v_norm As Single
    Dim n_norm As Single
    Dim v_dot_n As Single
    Dim alpha_in_rad As Single ' alpha is the angle between n and v.
    Const PI As Single = 3.14159274
    
    ' Return unmodified takeoff angle as default
    Plan3DCalculateEffectiveTakeOff! = tEffectiveTakeOff!

    ' Calculate modified effective take off angle based on sample tilt
    n_x! = -acoeff(2)
    n_y! = -acoeff(3)
    n_z! = 1#

    ' For testing only
    'n_x = 1
    'n_y = 0
    'n_z = 1
    'tEffectiveTakeOff = 40
    'tSpectrometerOrientation = 90
    
    phi_in_rad! = (90# - tSpectrometerOrientation!) * PI / 180#
    theta_in_rad! = (90# - tEffectiveTakeOff!) * PI / 180#

    v_x! = Sin(theta_in_rad!) * Cos(phi_in_rad!)
    v_y! = Sin(theta_in_rad!) * Sin(phi_in_rad!)
    v_z! = Cos(theta_in_rad!)

    v_norm! = Sqr(v_x! * v_x! + v_y! * v_y! + v_z! * v_z!) 'calculates the norm  of v. It should always be 1.
    
    n_norm! = Sqr(n_x! * n_x! + n_y! * n_y! + n_z! * n_z!) 'calculates the norm  of n.
    
    v_dot_n! = v_x! * n_x! + v_y! * n_y! + v_z! * n_z! 'calculates the dot product of v and n.

    If v_norm! = 0 Or v_norm! = 0 Then GoTo Plan3DCalculateEffectiveTakeOffError 'Check to not divide by 0.
    
    alpha_in_rad! = MathArcCos2(v_dot_n! / (v_norm! * n_norm!))

    Plan3DCalculateEffectiveTakeOff! = 90# - (alpha_in_rad! * 180# / PI)

Exit Function

' Errors
Plan3DCalculateEffectiveTakeOffError:
MsgBox Error$, vbOKOnly + vbCritical, "Plan3DCalculateEffectiveTakeOff"
ierror = True
Exit Function

End Function

