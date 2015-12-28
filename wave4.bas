Attribute VB_Name = "CodeWAVE4"
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

Sub WaveCalibrateReadWrite2(cpcMot As Integer, cpcXtal As String, cpcNumOf As Integer, cpcElm() As String, cpcXray() As String, cpcPosT() As Single, cpcPosA() As Single, cpcStd() As Integer, cpcCoeff() As Single)
' Reads the multiple peak calibration file for the next spectrometer crystal combination

ierror = False
On Error GoTo WaveCalibrateReadWrite2Error

Dim n As Integer

' Read calibration file
Input #Temp1FileNumber%, cpcMot%
Input #Temp1FileNumber%, cpcXtal$
Input #Temp1FileNumber%, cpcNumOf%

For n% = 1 To cpcNumOf%
Input #Temp1FileNumber%, cpcElm$(n%)
Input #Temp1FileNumber%, cpcXray$(n%)
Input #Temp1FileNumber%, cpcPosT!(n%)
Input #Temp1FileNumber%, cpcPosA!(n%)
Input #Temp1FileNumber%, cpcStd%(n%)
Next n%

Input #Temp1FileNumber%, cpcMot%
Input #Temp1FileNumber%, cpcCoeff!(1)
Input #Temp1FileNumber%, cpcCoeff!(2)
Input #Temp1FileNumber%, cpcCoeff!(3)

Exit Sub

' Errors
WaveCalibrateReadWrite2Error:
msg$ = "Note: if you received this error because the number of spectrometers or number crystals in your configuration has changed, you might have to edit or delete your PROBEWIN-KA.CAL, PROBEWIN-KB.CAL, etc. spectrometer multiple peak calibration files and restart the program."
MsgBox Error$ & vbCrLf & vbCrLf & msg$, vbOKOnly + vbCritical, "WaveCalibrateReadWrite2"
Close (Temp1FileNumber%)
ierror = True
Exit Sub

End Sub

