Attribute VB_Name = "CodeMW"
' (c) Copyright 1995-2024 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Sub MWCalculate(astring As String, numelms As Integer, elems() As String, fatoms() As Single, weight As Single)
' Parse elements and calculate the molecular weight of the passed formula

ierror = False
On Error GoTo MWCalculateError

Dim i As Integer

Dim AtomName As String, AtomNum As Integer
Dim AtomWeight As Single, NumAtoms As Single
Dim MolFormula As String, MolWeight As Single

Dim BlockOn As Integer, BlockWeight As Single
Dim NumBlocks As Integer

Dim ElemFormula As String, MolPercent As String
Dim ElemKinds As Integer

Dim tstring As String, fstring As String

ReDim NumElems(1 To MAXELM%) As Integer
ReDim NumElemsInBlock(1 To MAXELM%) As Integer

' Check for blank string
fstring$ = astring$
If Trim$(fstring$) = vbNullString Then GoTo MWCalculateEmptyString
    
Do Until fstring$ = vbNullString
AtomName$ = vbNullString
AtomNum% = 0
NumAtoms! = 0
tstring$ = Left$(fstring$, 1)
    
' Select on each character
    Select Case tstring$
    
' Alphabetic character
    Case "A" To "Z", "a" To "z"
      AtomName$ = UCase$(tstring$)
      fstring$ = Mid$(fstring$, 2): tstring$ = Left$(fstring$, 1)
      If tstring$ >= "a" And tstring$ <= "z" Then
        AtomName$ = AtomName$ & tstring$
      End If
      For i = MAXELM% To 1 Step -1
        If AtomName$ = RTrim$(Symup$(i%)) Then AtomNum% = i%: Exit For
      Next i%
      If AtomNum% = 0 Then
        If Len(AtomName$) = 2 Then
          AtomName$ = Left$(AtomName$, 1)
          For i% = MAXELM% To 1 Step -1
            If AtomName$ = RTrim$(Symup$(i%)) Then AtomNum% = i%: Exit For
          Next i%
          If AtomNum% = 0 Then GoTo MWCalculateNoElement
        Else
          GoTo MWCalculateNoElement
        End If
      Else
        If Len(AtomName$) = 2 Then fstring$ = Mid$(fstring$, 2)
      End If
      AtomWeight! = AllAtomicWts!(AtomNum%)
      NumAtoms! = MWNumber!(fstring$)
      
      ' Add in number of atoms
      If Not BlockOn% Then
        NumElems%(AtomNum%) = NumElems%(AtomNum%) + NumAtoms!
        MolWeight! = MolWeight! + AtomWeight! * NumAtoms!
        
      ' Add in number of atoms in block "()"
      Else
        BlockWeight! = BlockWeight! + AtomWeight! * NumAtoms!
        NumElemsInBlock%(AtomNum%) = NumElemsInBlock%(AtomNum%) + NumAtoms!
      End If
      
      If MolFormula$ <> vbNullString And Right$(MolFormula$, 1) <> "(" Then MolFormula$ = MolFormula$ + "-"
      MolFormula$ = MolFormula$ + AtomName$
      If NumAtoms! <> 1 Then MolFormula$ = MolFormula$ + LTrim$(Str$(NumAtoms!))
      
' Open parentheses
    Case "("
      If Not BlockOn% Then
        BlockOn% = True
        If MolFormula$ <> vbNullString Then MolFormula$ = MolFormula$ + "-"
        MolFormula$ = MolFormula$ + "("
        fstring$ = Mid$(fstring$, 2)
      Else
        GoTo MWCalculateSyntaxError
      End If
      
' Close parentheses
    Case ")"
      If BlockOn% Then
        BlockOn% = False
        fstring$ = Mid$(fstring$, 2)
        NumBlocks% = MWNumber!(fstring$)
        MolWeight! = MolWeight! + BlockWeight! * NumBlocks%
        
        ' Add in number of atoms in block
        For i% = 1 To MAXELM%
            NumElems%(i%) = NumElems%(i%) + NumElemsInBlock%(i%) * NumBlocks%
            NumElemsInBlock%(i%) = 0
        Next i%
        
        BlockWeight! = 0
        MolFormula$ = MolFormula$ + ")"
        If NumBlocks% <> 1 Then MolFormula$ = MolFormula$ + LTrim$(Str$(NumBlocks%))
      Else
        GoTo MWCalculateSyntaxError
      End If
      
' Delimiters
    Case " ", "-", "_", ",", ";": fstring$ = Mid$(fstring$, 2)
    
' All other
    Case Else
      GoTo MWCalculateSyntaxError
    End Select

Loop

' Load return arrays
  If MolFormula$ <> vbNullString Then
    ElemFormula$ = vbNullString
    MolPercent$ = vbNullString
    ElemKinds% = 0
    For i% = MAXELM% To 1 Step -1   ' start from high atomic numbers
      If NumElems%(i%) <> 0 Then
        ElemFormula$ = ElemFormula$ + RTrim$(Symup$(i%))
        If NumElems%(i%) <> 0 Then ElemFormula$ = ElemFormula$ + LTrim$(Str$(NumElems%(i%)))
        If MolPercent$ <> vbNullString Then MolPercent$ = MolPercent$ + "  "
        MolPercent$ = MolPercent$ + RTrim$(Symup$(i%))
        MolPercent$ = MolPercent$ + Str$(CInt((AllAtomicWts!(i%) * NumElems%(i%) * 100 / MolWeight!) * 100) / 100) + "%"
        ElemKinds% = ElemKinds% + 1
        
        elems$(ElemKinds%) = Symup$(i%)
        fatoms!(ElemKinds%) = NumElems%(i%)
        If fatoms!(ElemKinds%) = 0# Then fatoms!(ElemKinds%) = 1#
        
      End If
    Next i%
  End If
  
' Check for elements
If ElemKinds% < 1 Then GoTo MWCalculateNoElements
    
' Return parameters
numelms% = ElemKinds%
weight! = MolWeight!

' DebugMode results
If DebugMode Then
msg$ = MolFormula$ & " = "
msg$ = msg$ & ElemFormula$ & " = "
msg$ = msg$ & Str$(MolWeight!) & "g/mol, "
msg$ = msg$ & MolPercent$
Call IOWriteLog(msg$)
End If
    
Exit Sub

' Errors
MWCalculateError:
MsgBox Error$, vbOKOnly + vbCritical, "MWCalculate"
ierror = True
Exit Sub

MWCalculateEmptyString:
msg$ = "Empty formula string"
MsgBox msg$, vbOKOnly + vbExclamation, "MWCalculate"
ierror = True
Exit Sub

MWCalculateNoElement:
msg$ = AtomName$ & " is not a valid element"
MsgBox msg$, vbOKOnly + vbExclamation, "MWCalculate"
ierror = True
Exit Sub

MWCalculateSyntaxError:
msg$ = tstring$ & " <- Syntax error!"
MsgBox msg$, vbOKOnly + vbExclamation, "MWCalculate"
ierror = True
Exit Sub

MWCalculateNoElements:
msg$ = "No elements found in string " & astring$
MsgBox msg$, vbOKOnly + vbExclamation, "MWCalculate"
ierror = True
Exit Sub

End Sub

Function MWNumber(cstring As String) As Single
' Convert numeric characters to value
Dim sString As String, tstring As String

ierror = False
On Error GoTo MWNumberError

  sString$ = vbNullString
  Do Until cstring$ = vbNullString
    tstring$ = Left$(cstring$, 1)
    Select Case tstring$
    Case "0" To "9", ".": sString$ = sString$ + tstring$: cstring$ = Mid$(cstring$, 2)
    Case Else: Exit Do
    End Select
  Loop
  If sString$ = vbNullString Then MWNumber! = 1 Else MWNumber! = Val(sString$)

Exit Function

' Errors
MWNumberError:
MsgBox Error$, vbOKOnly + vbCritical, "MWNumber"
ierror = True
Exit Function

End Function

