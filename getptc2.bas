Attribute VB_Name = "CodeGETPTC2"
' (c) Copyright 1995-2019 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub GetPTCDefaults(zafinit As Integer, zaf As TypeZAF)
' Load PTC defaults

ierror = False
On Error GoTo GetPTCDefaultsError

' Special code for pure element calculations
If zafinit% = 0 Then
zaf.imodel% = 1
zaf.models%(1) = 1
zaf.idiam% = 1
zaf.diams!(1) = 10000# ' in microns

zaf.model% = 1      ' thick polished
zaf.diam! = 10000#  ' diameter in microns

zaf.d! = 1#         ' diameter in cm (thick pure element)
zaf.rho! = 1#       ' density in g/cm^3
zaf.j9! = 1#        ' thickness factor
zaf.X1! = 0.00001   ' numerical integration step in g/cm^2
Exit Sub
End If

' Load current model
If PTCModel% = 0 Then PTCModel% = 1
zaf.models%(1) = PTCModel%
zaf.model% = zaf.models%(1)

' Load current diameter (in microns)
If PTCDiameter! = 0# Then PTCDiameter! = 10000#
zaf.diams!(1) = PTCDiameter!
zaf.diam! = zaf.diams!(1)

' Load density (in g/cm^3)
If PTCDensity! = 0# Then PTCDensity! = 3#
zaf.rho! = PTCDensity!

' Load thickness factor
If PTCThicknessFactor! = 0# Then PTCThicknessFactor! = 1#
zaf.j9! = PTCThicknessFactor!

' Load numerical integration (in g/cm^2)
If PTCNumericalIntegrationStep! = 0# Then PTCNumericalIntegrationStep! = 0.00001
zaf.X1! = PTCNumericalIntegrationStep!

Exit Sub

' Errors
GetPTCDefaultsError:
MsgBox Error$, vbOKOnly + vbCritical, "GetPTCDefaults"
ierror = True
Exit Sub

End Sub

Sub GetPTCGetModelDiameter(imodel As Integer, idiam As Integer, zaf As TypeZAF)
' Load PTC current model and diameter

ierror = False
On Error GoTo GetPTCGetModelDiameterError

' Load current model
zaf.model% = zaf.models%(imodel%)

' Load current diameter (in microns)
zaf.diam! = zaf.diams!(idiam%)

Exit Sub

' Errors
GetPTCGetModelDiameterError:
MsgBox Error$, vbOKOnly + vbCritical, "GetPTCGetModelDiameter"
ierror = True
Exit Sub

End Sub


