Attribute VB_Name = "CodeXRAYPlot"
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub XrayPlotGetGraphExtents_PE(tForm As Form, xstart As Single, xstop As Single)
' Return the graph extents for the passed form (Pro Essentials code)

ierror = False
On Error GoTo XrayPlotGetGraphExtents_PEError

xstart! = tForm.Pesgo1.ManualMinX
xstop! = tForm.Pesgo1.ManualMaxX

Exit Sub

' Errors
XrayPlotGetGraphExtents_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayPlotGetGraphExtents_PE"
ierror = True
Exit Sub

End Sub

Sub XrayPlotShowDateStamp_PE(tForm As Form, sample() As TypeSample)
' Show a date time file stamp in the corner (Pro Essentials code)

ierror = False
On Error GoTo XrayPlotShowDateStamp_PEError

Dim r As Long
Dim xpos As Double, ypos As Double

' Display the filename and current time
xpos# = (tForm.Graph1.SDKInfo(7) + tForm.Graph1.SDKInfo(5)) / 2
ypos# = tForm.Graph1.SDKInfo(8) + tForm.Graph1.SDKInfo(6) + 50
r& = GSRText(xpos#, ypos#, 2, GSR_TXMID% + GSR_TXTOP% + GSR_TXTRANS%, 0, ProbeDataFile$ & ", " & Format$(CDate(sample(1).DateTimes(1))))

Exit Sub

' Errors
XrayPlotShowDateStamp_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayPlotShowDateStamp_PE"
ierror = True
Exit Sub

End Sub


