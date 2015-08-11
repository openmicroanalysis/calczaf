Attribute VB_Name = "CodeXRAY4"
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub XrayPlotConvert(tForm As Form, dx As Double, dy As Double, convdx As Double, convdy As Double)
' Convert data units to Graph SDK units for drawing stage coordinates on graph

ierror = False
On Error GoTo XrayPlotConvertError

Dim xorg As Double, xlen As Double  ' in graph units
Dim xmin As Double, xmax As Double  ' in data units
Dim xoffset As Double, xscale As Double ' from data to graph units

Dim yorg As Double, ylen As Double  ' in graph units
Dim ymin As Double, ymax As Double  ' in data units
Dim yoffset As Double, yscale As Double ' from data to graph units

xorg# = tForm.Graph1.SDKInfo(7)
xlen# = tForm.Graph1.SDKInfo(5)
xmin# = tForm.Graph1.SDKInfo(2)
xmax# = tForm.Graph1.SDKInfo(1)

yorg# = tForm.Graph1.SDKInfo(8)
ylen# = tForm.Graph1.SDKInfo(6)
ymin# = tForm.Graph1.SDKInfo(4)
ymax# = tForm.Graph1.SDKInfo(3)

' Calculate scale factors
If xmax# - xmin# = 0 Then Exit Sub
If ymax# - ymin# = 0 Then Exit Sub
xscale# = xlen# / (xmax# - xmin#)
yscale# = ylen# / (ymax# - ymin#)

' X-axis is all negative
If xmax# < 0# Then
xoffset# = xorg# - (xmax# * xscale#)

' X-axis is all positive
ElseIf xmin# > 0# Then
xoffset# = xorg# - (xmin# * xscale#)

' Normal
Else
xoffset# = xorg#
End If

' Y-axis is all negative
If ymax# < 0# Then
yoffset# = yorg# - (ymax# * yscale#)

' Y-axis is all positive
ElseIf ymin# > 0# Then
yoffset# = yorg# - (ymin# * yscale#)

' Normal
Else
yoffset# = yorg#
End If

' Calculate converted values
convdx# = dx# * xscale# + xoffset#
convdy# = dy# * yscale# + yoffset#

Exit Sub

' Errors
XrayPlotConvertError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayPlotConvert"
ierror = True
Exit Sub

End Sub

Sub XrayPlotConvert2(tForm As Form, dx As Double, dy As Double, convdx As Double, convdy As Double)
' Convert Graph SDK units to data units

ierror = False
On Error GoTo XrayPlotConvert2Error

Dim xorg As Double, xlen As Double  ' in graph units
Dim xmin As Double, xmax As Double  ' in data units
Dim xoffset As Double, xscale As Double ' from data to graph units

Dim yorg As Double, ylen As Double  ' in graph units
Dim ymin As Double, ymax As Double  ' in data units
Dim yoffset As Double, yscale As Double ' from data to graph units

xorg# = tForm.Graph1.SDKInfo(7)
xlen# = tForm.Graph1.SDKInfo(5)
xmin# = tForm.Graph1.SDKInfo(2)
xmax# = tForm.Graph1.SDKInfo(1)

yorg# = tForm.Graph1.SDKInfo(8)
ylen# = tForm.Graph1.SDKInfo(6)
ymin# = tForm.Graph1.SDKInfo(4)
ymax# = tForm.Graph1.SDKInfo(3)

' Calculate scale factors
If xmax# - xmin# = 0 Then Exit Sub
If ymax# - ymin# = 0 Then Exit Sub
xscale# = xlen# / (xmax# - xmin#)
yscale# = ylen# / (ymax# - ymin#)

' X-axis is all negative
If xmax# <= 0# Then
xoffset# = xorg# - (xmax# * xscale#)

' X-axis is all positive
ElseIf xmin# > 0# Then
xoffset# = xorg# - (xmin# * xscale#)

' Normal
Else
xoffset# = xorg#
End If

' Y-axis is all negative
If ymax# <= 0# Then
yoffset# = yorg# - (ymax# * yscale#)

' Y-axis is all positive
ElseIf ymin# > 0# Then
yoffset# = yorg# - (ymin# * yscale#)

' Normal
Else
yoffset# = yorg#
End If

' Calculate converted values
If xscale# <> 0# Then convdx# = (dx# - xoffset#) / xscale#
If yscale# <> 0# Then convdy# = (dy# - yoffset#) / yscale#

Exit Sub

' Errors
XrayPlotConvert2Error:
MsgBox Error$, vbOKOnly + vbCritical, "XrayPlotConvert2"
ierror = True
Exit Sub

End Sub

Sub XrayPlotConvertLog(tForm As Form, dx As Double, dy As Double, xflag As Boolean, yflag As Boolean, convdx As Single, convdy As Single)
' Convert linear data units to log data units

ierror = False
On Error GoTo XrayPlotConvertLogError

Dim ndecadesx As Integer, ndecadesy As Integer

Dim xmin As Double, xmax As Double  ' in data units
Dim ymin As Double, ymax As Double  ' in data units

Dim x_equals_one As Double, y_equals_one As Double

' Graph must be linlog or loglin
If tForm.Graph1.GraphType = graphScatter Then Exit Sub

' Check for linear values greater than zero
If xflag And dx# <= 0# Then Exit Sub
If xflag And dx# <= 0# Then Exit Sub

xmin# = tForm.Graph1.SDKInfo(2)
xmax# = tForm.Graph1.SDKInfo(1)

ymin# = tForm.Graph1.SDKInfo(4)
ymax# = tForm.Graph1.SDKInfo(3)

' Calculate number of decades in graph
If xflag Then ndecadesx% = MiscConvertLog10#(CLng(xmax# / xmin#))
If yflag Then ndecadesy% = MiscConvertLog10#(CLng(ymax# / ymin#))

' Calculate linear data value for 1
If xflag And xmax# > 1# Then x_equals_one# = xmax# - xmax# / ndecadesx%
If yflag And ymax# > 1# Then y_equals_one# = ymax# - ymax# / ndecadesy%
If xflag And xmax# = 1# Then x_equals_one# = xmax#
If yflag And ymax# = 1# Then y_equals_one# = ymax#
If xflag And xmax# < 1# Then x_equals_one# = xmax# + xmax# / ndecadesx%
If yflag And ymax# < 1# Then y_equals_one# = ymax# + ymax# / ndecadesy%

' Calculate log value
convdx! = dx#
If xflag Then convdx! = 10 ^ ((dx# / xmax# - x_equals_one# / xmax#) * ndecadesx%)
convdy! = dy#
If yflag Then convdy! = 10 ^ ((dy# / ymax# - y_equals_one# / ymax#) * ndecadesy%)

Exit Sub

' Errors
XrayPlotConvertLogError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayPlotConvertLog"
ierror = True
Exit Sub

End Sub

Sub XrayPlotConvertGraph(tGraph As Graph, dx As Double, dy As Double, convdx As Double, convdy As Double)
' Convert data units to Graph SDK units for drawing stage coordinates on graph

ierror = False
On Error GoTo XrayPlotConvertGraphError

Dim xorg As Double, xlen As Double  ' in graph units
Dim xmin As Double, xmax As Double  ' in data units
Dim xoffset As Double, xscale As Double ' from data to graph units

Dim yorg As Double, ylen As Double  ' in graph units
Dim ymin As Double, ymax As Double  ' in data units
Dim yoffset As Double, yscale As Double ' from data to graph units

xorg# = tGraph.SDKInfo(7)
xlen# = tGraph.SDKInfo(5)
xmin# = tGraph.SDKInfo(2)
xmax# = tGraph.SDKInfo(1)

yorg# = tGraph.SDKInfo(8)
ylen# = tGraph.SDKInfo(6)
ymin# = tGraph.SDKInfo(4)
ymax# = tGraph.SDKInfo(3)

' Calculate scale factors
If xmax# - xmin# = 0 Then Exit Sub
If ymax# - ymin# = 0 Then Exit Sub
xscale# = xlen# / (xmax# - xmin#)
yscale# = ylen# / (ymax# - ymin#)

' X-axis is all negative
If xmax# < 0# Then
xoffset# = xorg# - (xmax# * xscale#)

' X-axis is all positive
ElseIf xmin# > 0# Then
xoffset# = xorg# - (xmin# * xscale#)

' Normal
Else
xoffset# = xorg#
End If

' Y-axis is all negative
If ymax# < 0# Then
yoffset# = yorg# - (ymax# * yscale#)

' Y-axis is all positive
ElseIf ymin# > 0# Then
yoffset# = yorg# - (ymin# * yscale#)

' Normal
Else
yoffset# = yorg#
End If

' Calculate converted values
convdx# = dx# * xscale# + xoffset#
convdy# = dy# * yscale# + yoffset#

Exit Sub

' Errors
XrayPlotConvertGraphError:
MsgBox Error$, vbOKOnly + vbCritical, "XrayPlotConvertGraph"
ierror = True
Exit Sub

End Sub
