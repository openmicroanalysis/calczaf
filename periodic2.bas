Attribute VB_Name = "CodePERIODIC2"
' (c) Copyright 1995-2025 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Dim PeriodicElement(1 To MAXELM%) As Boolean

Sub Periodic2Load()
' Load the periodic table form for multiple element selection

ierror = False
On Error GoTo Periodic2LoadError

Dim i As Integer

' Load element symbols based on current selections
For i% = 1 To MAXELM%
FormPERIODIC2.CommandElement(i% - 1).Caption = Symup$(i%)

If PeriodicElement(i%) Then
FormPERIODIC2.CommandElement(i% - 1).BackColor = vbRed
FormPERIODIC2.CommandElement(i% - 1).tag = True
Else
FormPERIODIC2.CommandElement(i% - 1).BackColor = vbButtonFace
FormPERIODIC2.CommandElement(i% - 1).tag = False
End If

Next i%

' Load form
FormPERIODIC2.Show vbModal
Exit Sub

' Errors
Periodic2LoadError:
MsgBox Error$, vbOKOnly + vbCritical, "Periodic2Load"
ierror = True
Exit Sub

End Sub

Sub Periodic2Save()
' Save the periodic selections

ierror = False
On Error GoTo Periodic2SaveError

Dim i As Integer

' Save element symbols based on current selections
For i% = 1 To MAXELM%
If FormPERIODIC2.CommandElement(i% - 1).tag = True Then
PeriodicElement(i%) = True
Else
PeriodicElement(i%) = False
End If
Next i%

Unload FormPERIODIC2
Exit Sub

' Errors
Periodic2SaveError:
MsgBox Error$, vbOKOnly + vbCritical, "Periodic2Save"
ierror = True
Exit Sub

End Sub

Sub Periodic2SelectElement(ielm As Integer)
' Toggle selected element

ierror = False
On Error GoTo Periodic2SelectElementError

' Toggle element
PeriodicElement(ielm%) = Not PeriodicElement(ielm%)

' Update form
If PeriodicElement(ielm%) Then
FormPERIODIC2.CommandElement(ielm% - 1).BackColor = vbRed
FormPERIODIC2.CommandElement(ielm% - 1).tag = True
Else
FormPERIODIC2.CommandElement(ielm% - 1).BackColor = vbButtonFace
FormPERIODIC2.CommandElement(ielm% - 1).tag = False
End If

Exit Sub

' Errors
Periodic2SelectElementError:
MsgBox Error$, vbOKOnly + vbCritical, "Periodic2SelectElement"
ierror = True
Exit Sub

End Sub

Sub Periodic2Return(elmarray() As Boolean)
' Return the periodic selections to calling procedure

ierror = False
On Error GoTo Periodic2ReturnError

Dim i As Integer

' Save element symbols based on current selections
For i% = 1 To MAXELM%
elmarray(i%) = PeriodicElement(i%)
Next i%

Exit Sub

' Errors
Periodic2ReturnError:
MsgBox Error$, vbOKOnly + vbCritical, "Periodic2Return"
ierror = True
Exit Sub

End Sub

Sub Periodic2To(elmarray() As Boolean)
' Loads the periodic selections from calling procedure

ierror = False
On Error GoTo Periodic2ToError

Dim i As Integer

' Load element symbols based on passed array
For i% = 1 To MAXELM%
PeriodicElement(i%) = elmarray(i%)
Next i%

Exit Sub

' Errors
Periodic2ToError:
MsgBox Error$, vbOKOnly + vbCritical, "Periodic2To"
ierror = True
Exit Sub

End Sub

Sub Periodic2Clear()
' Clear the periodic table form for multiple element selection

ierror = False
On Error GoTo Periodic2ClearError

Dim i As Integer

' Load element symbols based on current selections
For i% = 1 To MAXELM%
FormPERIODIC2.CommandElement(i% - 1).BackColor = vbButtonFace
FormPERIODIC2.CommandElement(i% - 1).tag = False
Next i%

Exit Sub

' Errors
Periodic2ClearError:
MsgBox Error$, vbOKOnly + vbCritical, "Periodic2Clear"
ierror = True
Exit Sub

End Sub

