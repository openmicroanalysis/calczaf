Attribute VB_Name = "CodeCONVERT7"
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

Sub ConvertUpdateEdgeLineFlurFiles()
' This routine (only called once) updates xray files and add records for elements 95 - 100 (nrec 97-102)

' BE SURE TO COPY UNHASHED XLINE.DAT FROM REGISTER FOLDER
' AFTER STARTING PROGRAM AND THEN RUN UPDATE!!!!!!!!!!!

' From:
' R. B. Firestone, Table of isotopes, 8th Edition,
' Vol II: A+151-272, V.S. Shirley, ed.
' Lawrence Berkeley National Laboratory,
' John Wiley & Sons, Inc, New York

' H. Kleykamp, "Wavelengths of the M X-Rays of Uranium, Plutonium,
' and Americium", Z. Naturforsch., 36a, 1388-1390 (1981)

' H. Kleykamp, "X-ray Emission Wavelengths of Argon, Krypton, Xenon
' and Curium", Z. Naturforsch., 47a, 460-462 (1992).

ierror = False
On Error GoTo ConvertUpdateEdgeLineFlurFilesError

Dim nrec As Integer, n As Integer, response As Integer

Dim engrow As TypeEnergy
Dim edgrow As TypeEdge
Dim flurow As TypeFlur

msg$ = "Are you sure you want to overwrite the records for elements 95-100 in the x-ray edge (XEDGE.DAT), line (XLINE.DAT) and fluorescent (XFLUR.DAT) data files?"
response% = MsgBox(msg$, vbOKCancel + vbQuestion + vbDefaultButton1, "ConvertUpdateEdgeLineFlurFiles")
If response% = vbCancel Then Exit Sub

' Open x-ray edge file
Open XEdgeFile$ For Random Access Write As #XEdgeFileNumber% Len = 188

' Open x-ray line file
Open XLineFile$ For Random Access Write As #XLineFileNumber% Len = 188

' Open x-ray flur file
Open XFlurFile$ For Random Access Write As #XFlurFileNumber% Len = 188

' Loop on records
For n% = 95 To 100

If n% = 95 Then
edgrow.energy!(1) = 124.982 * EVPERKEV#  ' K
edgrow.energy!(2) = 23.808 * EVPERKEV#  ' L-I
edgrow.energy!(3) = 22.952 * EVPERKEV#  ' L-II
edgrow.energy!(4) = 18.51 * EVPERKEV#  ' L-III
edgrow.energy!(5) = 6.133 * EVPERKEV#  ' M-I
edgrow.energy!(6) = 5.739 * EVPERKEV#  ' M-II
edgrow.energy!(7) = 4.698 * EVPERKEV#  ' M-III
edgrow.energy!(8) = 4.096 * EVPERKEV#  ' M-IV
edgrow.energy!(9) = 3.89 * EVPERKEV#   ' M-V

engrow.energy!(1) = 106.47 * EVPERKEV#  ' Ka
engrow.energy!(2) = 120.284 * EVPERKEV#  ' Kb
engrow.energy!(3) = 14.62 * EVPERKEV#  ' La
engrow.energy!(4) = 18.856 * EVPERKEV#  ' Lb
engrow.energy!(5) = 3.442 * EVPERKEV#  ' Ma
engrow.energy!(6) = 3.633 * EVPERKEV#  ' Mb

flurow.fraction!(1) = 0.977    ' Ka
flurow.fraction!(2) = 0.977    ' Kb
flurow.fraction!(3) = 0.494    ' La
flurow.fraction!(4) = 0.494    ' Lb
flurow.fraction!(5) = 0#    ' Ma
flurow.fraction!(6) = 0#    ' Mb
End If

If n% = 96 Then
edgrow.energy!(1) = 128.241 * EVPERKEV#  ' K
edgrow.energy!(2) = 24.526 * EVPERKEV#  ' L-I
edgrow.energy!(3) = 23.651 * EVPERKEV#  ' L-II
edgrow.energy!(4) = 18.97 * EVPERKEV#  ' L-III
edgrow.energy!(5) = 6.337 * EVPERKEV#  ' M-I
edgrow.energy!(6) = 5.937 * EVPERKEV#  ' M-II
edgrow.energy!(7) = 4.838 * EVPERKEV#  ' M-III
edgrow.energy!(8) = 4.224 * EVPERKEV#  ' M-IV
edgrow.energy!(9) = 4.009 * EVPERKEV#   ' M-V

engrow.energy!(1) = 109.271 * EVPERKEV#  ' Ka
engrow.energy!(2) = 123.403 * EVPERKEV#  ' Kb
engrow.energy!(3) = 14.961 * EVPERKEV#  ' La
engrow.energy!(4) = 19.427 * EVPERKEV#  ' Lb
engrow.energy!(5) = 3.535 * EVPERKEV#  ' Ma
engrow.energy!(6) = 3.738 * EVPERKEV#  ' Mb

flurow.fraction!(1) = 0.978    ' Ka
flurow.fraction!(2) = 0.978    ' Kb
flurow.fraction!(3) = 0.504    ' La
flurow.fraction!(4) = 0.504    ' Lb
flurow.fraction!(5) = 0#    ' Ma
flurow.fraction!(6) = 0#    ' Mb
End If

If n% = 97 Then
edgrow.energy!(1) = 0#  ' K
edgrow.energy!(2) = 0#  ' L-I
edgrow.energy!(3) = 0#  ' L-II
edgrow.energy!(4) = 0#  ' L-III
edgrow.energy!(5) = 0#  ' M-I
edgrow.energy!(6) = 0#  ' M-II
edgrow.energy!(7) = 0#  ' M-III
edgrow.energy!(8) = 0#  ' M-IV
edgrow.energy!(9) = 0#   ' M-V

engrow.energy!(1) = 0#  ' Ka
engrow.energy!(2) = 0#  ' Kb
engrow.energy!(3) = 0#  ' La
engrow.energy!(4) = 0#  ' Lb
engrow.energy!(5) = 0#  ' Ma
engrow.energy!(6) = 0#  ' Mb

flurow.fraction!(1) = 0#    ' Ka
flurow.fraction!(2) = 0#    ' Kb
flurow.fraction!(3) = 0#    ' La
flurow.fraction!(4) = 0#    ' Lb
flurow.fraction!(5) = 0#    ' Ma
flurow.fraction!(6) = 0#    ' Mb
End If

If n% = 98 Then
End If

If n% = 99 Then
End If

If n% = 100 Then
End If

nrec% = n% + 2
Put #XEdgeFileNumber%, nrec%, edgrow
Put #XLineFileNumber%, nrec%, engrow
Put #XFlurFileNumber%, nrec%, flurow
Next n%

Close #XEdgeFileNumber%
Close #XLineFileNumber%
Close #XFlurFileNumber%

msg$ = "Records updated"
MsgBox msg$, vbOKOnly + vbInformation, "ConvertUpdateEdgeLineFlurFiles"

Exit Sub

' Errors
ConvertUpdateEdgeLineFlurFilesError:
MsgBox Error$, vbOKOnly + vbCritical, "ConvertUpdateEdgeLineFlurFiles"
Close #XEdgeFileNumber%
Close #XLineFileNumber%
Close #XFlurFileNumber%
ierror = True
Exit Sub

End Sub
