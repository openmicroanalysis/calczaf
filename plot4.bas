Attribute VB_Name = "CodePLOT4"
' (c) Copyright 1995-2026 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub PlotGetRelativeMicrons(firstline As Boolean, sampleline As Integer, xydist As Single, sample() As TypeSample)
' Calculates the relative change in stage coordinates in microns

ierror = False
On Error GoTo PlotGetRelativeMicronsError

Dim incr As Single

Static lastsamplename As String
Static xold As Single
Static yold As Single

' Check for "continued" sample name, if not re-set relative micron distance variables
If firstline Or (sample(1).Name$ <> lastsamplename$ And sample(1).Name$ <> CONTINUED$) Then
xold! = sample(1).StagePositions!(sampleline%, 1)
yold! = sample(1).StagePositions!(sampleline%, 2)
xydist! = 0#
firstline = False
End If

' Calculate x and y distance (ignore Z)
incr! = (xold - sample(1).StagePositions!(sampleline%, 1)) ^ 2 + (yold - sample(1).StagePositions!(sampleline%, 2)) ^ 2
incr! = Sqr(incr!)
xydist! = xydist! + incr! * MotUnitsToAngstromMicrons!(XMotor%)

' Store last position for next increment
xold! = sample(1).StagePositions!(sampleline%, 1)
yold! = sample(1).StagePositions!(sampleline%, 2)

' Store the current name (store even if "continued" to force next relative calculation)
lastsamplename$ = sample(1).Name$

Exit Sub

' Errors
PlotGetRelativeMicronsError:
MsgBox Error$, vbOKOnly + vbCritical, "PlotGetRelativeMicrons"
ierror = True
Exit Sub

End Sub


