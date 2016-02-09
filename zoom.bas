Attribute VB_Name = "CodeZoom"
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

Sub ZoomPrintGraph_PE(tGraph As Pesgo)
' Print the graph at current zoom (Pro Essentials code)

ierror = False
On Error GoTo ZoomPrintGraph_PEError

Dim bstatus As Boolean

' Launch print dialog
'bstatus = tGraph.PEprintgraph(CLng(0), CLng(0), CLng(0))      ' printer default
bstatus = tGraph.PEprintgraph(CLng(0), CLng(0), CLng(1))      ' print landscape
'bstatus = tGraph.PEprintgraph(CLng(0), CLng(0), CLng(2))      ' print portrait

Exit Sub

' Errors
ZoomPrintGraph_PEError:
MsgBox Error$, vbOKOnly + vbCritical, "ZoomPrintGraph_PE"
ierror = True
Exit Sub

End Sub

Sub ZoomTrack(mode As Integer, X As Single, Y As Single, fX As Double, fY As Double, tGraph As Pesgo)
' Convert track data for Pro Essentials
'  mode = 0 for entire graph control
'  mode = 1 for just the plot area

ierror = False
On Error GoTo ZoomTrackError

Dim nA As Long, nX As Long, nY As Long
Dim nLeft As Integer, nTop As Integer
Dim nRight As Integer, nBottom As Integer
Dim pX As Integer, pY As Integer
    
' Get last mouse location within control
tGraph.GetLastMouseMove pX%, pY%
    
' Test to see if this is within grid area
tGraph.GetRectGraph nLeft%, nTop%, nRight%, nBottom%
If mode% = 0 Or (mode% = 1 And pX% > nLeft% And pX% < nRight% And pY% > nTop% And pY% < nBottom%) Then
   nA& = 0              ' initialize subset to use (if using OverlapMultiAxes)
   nX& = CLng(pX%)      ' initialize nX and nY with mouse location
   nY& = CLng(pY%)
   tGraph.PEconvpixeltograph nA&, nX&, nY&, fX#, fY#, 0, 0, 0
End If

Exit Sub

' Errors
ZoomTrackError:
MsgBox Error$, vbOKOnly + vbCritical, "ZoomTrack"
ierror = True
Exit Sub

End Sub
