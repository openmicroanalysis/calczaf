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

Private Declare Function GSBox2D Lib "GSWDLL32.DLL" (ByVal fxOrg#, ByVal fyOrg#, ByVal fWid#, ByVal fHt#, ByVal nPatt&, ByVal nClr&) As Long
Private Declare Function GSSetROP Lib "GSWDLL32.DLL" (ByVal nROP&) As Long

Dim newbox As Integer, mousedown As Integer
Dim x1 As Double, y1 As Double
Dim X2 As Double, Y2 As Double
Dim xmax As Double, ymax As Double
Dim xmin As Double, ymin As Double

Sub ZoomSDKPress(PressStatus As Integer, PressX As Double, PressY As Double, PressDataX As Double, PressDataY As Double, mode As Integer, tForm As Form)
' Handle graph mouse press events
' mode = 0  force x and y minimum to zero if negative
' mode = 1  normal zoom
' mode = 2  force y minimum to zero if negative

ierror = False
On Error GoTo ZoomSDKPressError

Dim r As Long
Dim txmax As Double, tymax As Double
Dim txmin As Double, tymin As Double

' Get and store beginning coordinates for box
    If PressStatus = 1 Then
        newbox = True
        mousedown = True
        ' Save form coordinates
        x1# = PressX#
        y1# = PressY#
        ' Save the intial position
        xmin# = PressDataX#
        ymax# = PressDataY#
        
' Mouse up, raw new zoom
    Else
        If Not mousedown Then Exit Sub
        mousedown = False
        'Erase last box drawn
        Call ZoomDrawBox(x1#, y1#, X2#, Y2#)
        If ierror Then Exit Sub
        xmax# = PressDataX#
        ymin# = PressDataY#
        
        ' Check for too much zoom
        tymax# = tForm.Graph1.YAxisMax
        tymin# = tForm.Graph1.YAxisMin
        If tymax# = 0# And tymin# = 0# Then tymax# = 1  ' (when initial YAxisStyle = 1)
        txmax# = tForm.Graph1.XAxisMax
        txmin# = tForm.Graph1.XAxisMin
        If txmax# = 0# And txmin# = 0# Then txmax# = 1  ' (when initial XAxisStyle = 1)
        If txmax# - txmin# = 0 Then txmin# = 0
        If Abs((ymax# - ymin#) / (tymax# - tymin#)) > 0.01 Then
        If Abs((xmax# - xmin#) / (txmax# - txmin#)) > 0.01 Then

        ' Rescale axes
        If Not MiscDifferenceIsSmall(CSng(ymax#), CSng(ymin#), 0.00001) Then
        tForm.Graph1.YAxisMax = ymax#
        If (mode% = 0 Or mode% = 2) And ymin# < 0# Then
        tForm.Graph1.YAxisMin = 0#
        Else
        tForm.Graph1.YAxisMin = ymin#
        End If
        tForm.Graph1.YAxisStyle = 2
        End If

        If Not MiscDifferenceIsSmall(CSng(xmax#), CSng(xmin#), 0.00001) Then
        tForm.Graph1.XAxisMax = xmax#
        If mode% = 0 And xmin# < 0# Then
        tForm.Graph1.XAxisMin = 0#
        Else
        tForm.Graph1.XAxisMin = xmin#
        End If
        tForm.Graph1.XAxisStyle = 2
        End If
        
        End If
        End If
  
    r& = GSSetROP(0)         ' Set the raster OP back to normal
    tForm.Graph1.DrawMode = 2      ' refresh the graph
    End If

Exit Sub

' Errors
ZoomSDKPressError:
MsgBox Error$, vbOKOnly + vbCritical, "ZoomSDKPress"
ierror = True
Exit Sub

End Sub

Sub ZoomDrawBox(x1 As Double, y1 As Double, X2 As Double, Y2 As Double)
' Draw a rectangle

ierror = False
On Error GoTo ZoomDrawBoxError

Dim nX As Double, nY As Double, nWidth As Double, nHeight As Double
Dim nPatt As Integer, nColor As Integer
Dim r As Long

nPatt% = 1
nColor% = 15

If X2# < x1# Then
nX# = X2#
Else
nX# = x1#
End If

If Y2# < y1# Then
nY# = Y2#
Else
nY# = y1#
End If

nWidth# = Abs(X2# - x1#)
nHeight# = Abs(Y2# - y1#)

r& = GSBox2D(nX#, nY#, nWidth#, nHeight#, nPatt%, nColor%)
Exit Sub

' Errors
ZoomDrawBoxError:
MsgBox Error$, vbOKOnly + vbCritical, "ZoomDrawBox"
ierror = True
Exit Sub

End Sub

Sub ZoomSDKTrack(TrackX As Double, TrackY As Double)
' Draw rectangle

ierror = False
On Error GoTo ZoomSDKTrackError

Dim r As Long

If mousedown Then
    If newbox Then
        ' Draw new box
        X2# = TrackX#
        Y2# = TrackY#
        Call ZoomDrawBox(x1#, y1#, X2#, Y2#)
        newbox = 0
    Else
        ' Redraw previous box
        Call ZoomDrawBox(x1#, y1#, X2#, Y2#)
        If ierror Then Exit Sub
        r& = GSSetROP(2)
        ' Draw new box
        X2# = TrackX#
        Y2# = TrackY#
        Call ZoomDrawBox(x1#, y1#, X2#, Y2#)
        If ierror Then Exit Sub
    End If
End If

Exit Sub

' Errors
ZoomSDKTrackError:
MsgBox Error$, vbOKOnly + vbCritical, "ZoomSDKTrack"
ierror = True
Exit Sub

End Sub

Sub ZoomPrintGraph_GS(tGraph As Graph)
' Print the graph at current zoom (Graphics Server code)

ierror = False
On Error GoTo ZoomPrintGraph_GSError

If tGraph.NumSets < 1 Or tGraph.NumPoints < 1 Then Exit Sub

tGraph.PrintInfo(11) = 1    ' landscape
tGraph.PrintInfo(12) = 1    ' fit to page

tGraph.PrintInfo(6) = Printer.ScaleLeft
tGraph.PrintInfo(7) = Printer.ScaleTop
tGraph.PrintInfo(8) = Printer.ScaleWidth
tGraph.PrintInfo(9) = Printer.ScaleHeight

' Check if color printer is default and forcing B and W
tGraph.PrintStyle = 3       ' print color with border
tGraph.DrawMode = 5         ' print
Exit Sub

' Errors
ZoomPrintGraph_GSError:
MsgBox Error$, vbOKOnly + vbCritical, "ZoomPrintGraph_GS"
ierror = True
Exit Sub

End Sub

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
