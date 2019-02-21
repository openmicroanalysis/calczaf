Attribute VB_Name = "CodeMISC"
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

Function IPOS1(ByVal n As Integer, ByVal sym As String, symarray() As String) As Integer
' This routine returns as its value a pointer to the first occurance
' of 'sym' in the character array 'symarray'.  The first 'n' positions
' in  'symarray' are searched.  If 'sym' does not occur in those positions
' IPOS1 is equal to 0. Example:
'  n = 4
'  sym    = "li"
'  symarray = "h","he","li","be"
'  IPOS1    will be set to 3

ierror = False
On Error GoTo IPOS1Error

Dim i As Integer

If n% <= 0 Then GoTo Fail1
For i% = 1 To n%
If Trim$(LCase$(symarray$(i%))) = Trim$(LCase$(sym$)) Then GoTo Found1
Next i%

Fail1:
IPOS1 = 0
Exit Function

Found1:
IPOS1 = i%
Exit Function

' Errors
IPOS1Error:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS1"
ierror = True
Exit Function

End Function

Function IPOS2(ByVal n As Integer, ByVal num As Integer, iarray() As Integer) As Integer
' This routine returns as its value a pointer to the first occurance
' of 'num' in the integer array 'iarray'.  The first 'n' positions
' in  'iarray' are searched.  If 'num' does not occur in those positions
' IPOS2 is equal to 0. Example:
'  n = 4
'  num    = 22
'  iarray = 16,23,22,24
'  IPOS2    will be set to 3

ierror = False
On Error GoTo IPOS2Error

Dim i As Integer

If n% <= 0 Then GoTo Fail2
For i% = 1 To n%
If iarray%(i%) = num% Then GoTo Found2
Next i%

Fail2:
IPOS2 = 0
Exit Function

Found2:
IPOS2 = i%
Exit Function

' Errors
IPOS2Error:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS2"
ierror = True
Exit Function

End Function

Function IPOS2A(ByVal n As Long, ByVal num1 As Integer, ByVal num2 As Integer, array1() As Long, array2() As Long) As Integer
' This routine returns as its value a pointer to the first occurance of 'num1' and 'num2' in the integer
' arrays 'array1' and 'array2'.  The first 'n' positions are searched.  If 'num2' and 'num2' do not occur in
' those positions then IPOS2A is equal to 0. Example:
'  n = 4
'  num1    = 22
'  num2    = 2
'  array1() = 16,22,22,24
'  array2() = 1,1,2,4
'  IPOS2A    will be set to 3

ierror = False
On Error GoTo IPOS2AError

Dim i As Integer

If n& <= 0 Then GoTo Fail2A
For i% = 1 To n&
If array1&(i%) = num1% And array2&(i%) = num2% Then GoTo Found2A
Next i%

Fail2A:
IPOS2A = 0
Exit Function

Found2A:
IPOS2A = i%
Exit Function

' Errors
IPOS2AError:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS2A"
ierror = True
Exit Function

End Function

Function IPOS22(ByVal n As Long, ByVal num As Long, narray() As Long) As Long
' This routine returns as its value a pointer to the first occurance
' of 'num' in the LONG integer array 'narray'.  The first 'n' positions
' in  'narray' are searched.  If 'num' does not occur in those positions
' IPOS22 is equal to 0. Example:
'  n = 4
'  num    = 22
'  narray = 16,23,22,24
'  IPOS22    will be set to 3

ierror = False
On Error GoTo IPOS22Error

Dim i As Long

If n& <= 0 Then GoTo Fail22
For i& = 1 To n&
If narray&(i&) = num& Then GoTo Found22
Next i&

Fail22:
IPOS22 = 0
Exit Function

Found22:
IPOS22 = i&
Exit Function

' Errors
IPOS22Error:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS22"
ierror = True
Exit Function

End Function

Function IPOS3(ByVal n As Integer, ByVal temp As Single, rarray() As Single) As Integer
' This routine returns as its value a pointer to the first occurance
' of 'temp' in the real array 'rarray'.  The first 'n' positions
' in  'rarray' are searched.  If 'temp' does not occur in those positions
' IPOS3 is equal to 0. Example:
'  n = 4
'  temp    = 22.3
'  iarray = 16.1,23.78.,22.3,24.12
'  IPOS3    will be set to 3

ierror = False
On Error GoTo IPOS3Error

Dim i As Integer

If n% <= 0 Then GoTo Fail3
For i% = 1 To n%
If rarray!(i%) = temp! Then GoTo Found3
Next i%

Fail3:
IPOS3 = 0
Exit Function

Found3:
IPOS3 = i%
Exit Function

' Errors
IPOS3Error:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS3"
ierror = True
Exit Function

End Function

Function IPOS4(ByVal n As Integer, ByVal sym As String, symray() As String) As Integer
' This routine returns as its value a pointer to the first occurance
' of 'sym' in the character array 'symray'.  The first 'n' positions
' in  'symray' are searched but only the first character is checked!
' If 'sym' does not occur in those positions IPOS4 is equal to 0. Example:
'  n = 6
'  sym    = "kb"
'  symray = "ka","kb","la","lb","ma",mb"
'  IPOS4    will be set to 1

' Use this formula to convert 1, 3, 5 to 1, 2, 3
' i% = i% - (i% - 1) + (i% - 1) / 2

ierror = False
On Error GoTo IPOS4Error

Dim i As Integer

If n% <= 0 Then GoTo Fail4
For i% = 1 To n%
If Left$(Trim$(LCase$(symray$(i%))), 1) = Left$(Trim$(LCase$(sym$)), 1) Then GoTo Found4
Next i%

Fail4:
IPOS4 = 0
Exit Function

Found4:
IPOS4 = i%
Exit Function

' Errors
IPOS4Error:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS4"
ierror = True
Exit Function

End Function

Sub MiscAddCRToText(tText As TextBox)
' Add a <cr> to text box

ierror = False
On Error GoTo MiscAddCRToTextError

tText.Text = tText.Text & vbCrLf
tText.SetFocus

' Let got focus event complete
DoEvents

' Set the insertion point at the end of the text
tText.SelStart = Len(tText.Text)

Exit Sub

' Errors
MiscAddCRToTextError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAddCRToText"
ierror = True
Exit Sub

End Sub

Function MiscAutoFormat(ByVal treal As Single) As String
' Function to return an automatically formatted real number

ierror = False
On Error GoTo MiscAutoFormatError

Dim astring As String

' Negative numbers
If treal! < 0# Then
astring$ = f85$
If Abs(treal!) >= 1# Then astring$ = f84$
If Abs(treal!) >= 10# Then astring$ = f83$
If Abs(treal!) >= 100# Then astring$ = f82$
If Abs(treal!) >= 1000# Then astring$ = f81$
If Abs(treal!) >= 10000# Then astring$ = f80$

' Positive numbers
Else
astring$ = f86$
If treal! >= 1# Then astring$ = f85$
If treal! >= 10# Then astring$ = f84$
If treal! >= 100# Then astring$ = f83$
If treal! >= 1000# Then astring$ = f82$
If treal! >= 10000# Then astring$ = f81$
If treal! >= 100000# Then astring$ = f80$
End If

' Format number
MiscAutoFormat$ = Format$(Format$(treal!, astring$), a80$)

Exit Function

' Errors
MiscAutoFormatError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAutoFormat"
ierror = True
Exit Function

End Function

Function MiscAutoFormatA(ByVal treal As Single) As String
' Function to return an automatically formatted real number (maximum 4 decimals)

ierror = False
On Error GoTo MiscAutoFormatAError

Dim astring As String

' Negative numbers
If treal! < 0# Then
astring$ = f85$
If Abs(treal!) >= 1# Then astring$ = f84$
If Abs(treal!) >= 10# Then astring$ = f83$
If Abs(treal!) >= 100# Then astring$ = f82$
If Abs(treal!) >= 1000# Then astring$ = f81$
If Abs(treal!) >= 10000# Then astring$ = f80$

' Positive numbers
Else
astring$ = f85$
If treal! >= 1# Then astring$ = f85$
If treal! >= 10# Then astring$ = f84$
If treal! >= 100# Then astring$ = f83$
If treal! >= 1000# Then astring$ = f82$
If treal! >= 10000# Then astring$ = f81$
If treal! >= 100000# Then astring$ = f80$
End If

' Format number
MiscAutoFormatA$ = Format$(Format$(treal!, astring$), a80$)

Exit Function

' Errors
MiscAutoFormatAError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAutoFormatA"
ierror = True
Exit Function

End Function

Function MiscAutoFormatB(ByVal treal As Single) As String
' Function to return an automatically formatted real number (maximum 3 decimals)

ierror = False
On Error GoTo MiscAutoFormatBError

Dim astring As String

' Negative numbers
If treal! < 0# Then
astring$ = f83$
If Abs(treal!) >= 1# Then astring$ = f82$
If Abs(treal!) >= 10# Then astring$ = f81$
If Abs(treal!) >= 100# Then astring$ = f80$
If Abs(treal!) >= 1000# Then astring$ = f80$
If Abs(treal!) >= 10000# Then astring$ = f80$

' Positive numbers
Else
astring$ = f83$
If treal! >= 1# Then astring$ = f82$
If treal! >= 10# Then astring$ = f81$
If treal! >= 100# Then astring$ = f80$
If treal! >= 1000# Then astring$ = f80$
If treal! >= 10000# Then astring$ = f80$
If treal! >= 100000# Then astring$ = f80$
End If

' Format number
MiscAutoFormatB$ = Format$(Format$(treal!, astring$), a80$)

Exit Function

' Errors
MiscAutoFormatBError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAutoFormatB"
ierror = True
Exit Function

End Function

Function MiscAutoFormatBB(ByVal treal As Single) As String
' Function to return an automatically formatted real number (maximum 4 decimals)

ierror = False
On Error GoTo MiscAutoFormatBBError

Dim astring As String

' Negative numbers
If treal! < 0# Then
astring$ = f84$
If Abs(treal!) >= 0.1 Then astring$ = f83$
If Abs(treal!) >= 1# Then astring$ = f82$
If Abs(treal!) >= 10# Then astring$ = f81$
If Abs(treal!) >= 100# Then astring$ = f80$
If Abs(treal!) >= 1000# Then astring$ = f80$
If Abs(treal!) >= 10000# Then astring$ = f80$

' Positive numbers
Else
astring$ = f84$
If treal! >= 0.1 Then astring$ = f83$
If treal! >= 1# Then astring$ = f82$
If treal! >= 10# Then astring$ = f81$
If treal! >= 100# Then astring$ = f80$
If treal! >= 1000# Then astring$ = f80$
If treal! >= 10000# Then astring$ = f80$
If treal! >= 100000# Then astring$ = f80$
End If

' Format number
MiscAutoFormatBB$ = Format$(Format$(treal!, astring$), a80$)

Exit Function

' Errors
MiscAutoFormatBBError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAutoFormatBB"
ierror = True
Exit Function

End Function

Function MiscAutoFormatC(ByVal treal As Single) As String
' Function to return an automatically formatted real number (maximum decimals)

ierror = False
On Error GoTo MiscAutoFormatCError

Dim astring As String

' Negative numbers
If treal! < 0# Then
astring$ = f87$
If Abs(treal!) >= 0.01 Then astring$ = f86$
If Abs(treal!) >= 0.1 Then astring$ = f85$
If Abs(treal!) >= 1# Then astring$ = f84$
If Abs(treal!) >= 10# Then astring$ = f83$
If Abs(treal!) >= 100# Then astring$ = f82$
If Abs(treal!) >= 1000# Then astring$ = f81$
If Abs(treal!) >= 10000# Then astring$ = f80$

' Positive numbers
Else
astring$ = f87$
If treal! >= 0.1 Then astring$ = f86$
If treal! >= 1# Then astring$ = f85$
If treal! >= 10# Then astring$ = f84$
If treal! >= 100# Then astring$ = f83$
If treal! >= 1000# Then astring$ = f82$
If treal! >= 10000# Then astring$ = f81$
If treal! >= 100000# Then astring$ = f80$
End If

' Format number
MiscAutoFormatC$ = Format$(Format$(treal!, astring$), a80$)

Exit Function

' Errors
MiscAutoFormatCError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAutoFormatC"
ierror = True
Exit Function

End Function

Function MiscAutoFormatI(ByVal itemp As Integer) As String
' Function to return an automatically formatted integer (8 characters)

ierror = False
On Error GoTo MiscAutoFormatIError

' Format number
MiscAutoFormatI$ = Format$(itemp%, a80$)

Exit Function

' Errors
MiscAutoFormatIError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAutoFormatI"
ierror = True
Exit Function

End Function

Function MiscAutoFormatL(ByVal ntemp As Long) As String
' Function to return an automatically formatted long integer (8 characters)

ierror = False
On Error GoTo MiscAutoFormatLError

' Format number
MiscAutoFormatL$ = Format$(ntemp&, a80$)

Exit Function

' Errors
MiscAutoFormatLError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAutoFormatL"
ierror = True
Exit Function

End Function

Function MiscAutoFormatM(ByVal treal As Single) As String
' Function to return an automatically formatted real number (maximum decimals) in 10 characters

ierror = False
On Error GoTo MiscAutoFormatMError

Dim astring As String

' Negative numbers
If treal! < 0# Then
astring$ = f109$
If Abs(treal!) >= 0.01 Then astring$ = f108$
If Abs(treal!) >= 0.1 Then astring$ = f107$
If Abs(treal!) >= 1# Then astring$ = f106$
If Abs(treal!) >= 10# Then astring$ = f105$
If Abs(treal!) >= 100# Then astring$ = f104$
If Abs(treal!) >= 1000# Then astring$ = f103$
If Abs(treal!) >= 10000# Then astring$ = f102$
If Abs(treal!) >= 100000# Then astring$ = f101$
If Abs(treal!) >= 1000000# Then astring$ = f100$

' Positive numbers
Else
astring$ = f109$
If treal! >= 0.01 Then astring$ = f108$
If treal! >= 0.1 Then astring$ = f107$
If treal! >= 1# Then astring$ = f106$
If treal! >= 10# Then astring$ = f105$
If treal! >= 100# Then astring$ = f104$
If treal! >= 1000# Then astring$ = f103$
If treal! >= 10000# Then astring$ = f102$
If treal! >= 100000# Then astring$ = f101$
If treal! >= 1000000# Then astring$ = f100$
End If

' Format number
MiscAutoFormatM$ = Format$(Format$(treal!, astring$), a100$)

Exit Function

' Errors
MiscAutoFormatMError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAutoFormatM"
ierror = True
Exit Function

End Function

Function MiscAutoFormatN(ByVal treal As Single, ByVal n As Integer) As String
' Function to return an automatically formatted real number (n=decimals)

ierror = False
On Error GoTo MiscAutoFormatNError

Dim astring As String

' Format number
astring$ = f80$
If n% = 1 Then astring$ = f81$
If n% = 2 Then astring$ = f82$
If n% = 3 Then astring$ = f83$
If n% = 4 Then astring$ = f84$
If n% = 5 Then astring$ = f85$
If n% = 6 Then astring$ = f86$
If n% = 7 Then astring$ = f87$
MiscAutoFormatN$ = Format$(Format$(treal!, astring$), a80$)

Exit Function

' Errors
MiscAutoFormatNError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAutoFormatN"
ierror = True
Exit Function

End Function

Function MiscDifferenceIsSmall(ByVal temp1 As Single, ByVal temp2 As Single, ByVal toler As Single) As Boolean
' Checks if the difference between two numbers is small realtive to a tolerance
' toler in fraction (0.01 = 1%)

ierror = False
On Error GoTo MiscDifferenceIsSmallError

' Assume large difference
MiscDifferenceIsSmall = False

' If temp1! is not zero, check
If temp1! <> 0# Then
If Abs((temp1! - temp2!) / temp1!) < toler! Then
MiscDifferenceIsSmall = True
Exit Function
End If

' If temp2! is not zero, check
ElseIf temp2! <> 0# Then
If Abs((temp1! - temp2!) / temp2!) < toler! Then
MiscDifferenceIsSmall = True
Exit Function
End If

' If they are zero, see if they are equal
ElseIf temp1! = temp2! Then
MiscDifferenceIsSmall = True
Exit Function
End If

Exit Function

' Errors
MiscDifferenceIsSmallError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscDifferenceIsSmall"
ierror = True
Exit Function

End Function

Sub MiscParseStringToString(astring As String, bstring As String)
' Parse the first substring (space delimited) in "astring" and return in "bstring"

ierror = False
On Error GoTo MiscParseStringToStringError

Dim n As Long

' Check for empty string
astring$ = Trim$(astring$)
bstring$ = vbNullString
If astring$ = vbNullString Then GoTo MiscParseStringToStringEmpty

' Parse out based on first space character
n& = InStr(astring$, VbSpace$)

' Load substring
If n& > 0 Then
bstring$ = Left$(astring$, n& - 1)
Else
bstring$ = Trim$(astring$)
End If

' Return remainder
If n& > 0 Then
astring$ = Mid$(astring$, n&)
Else
astring$ = Mid$(astring$, n& + 1)
End If

'  If end of string, blank returned string
If astring$ = bstring$ Then astring$ = vbNullString

' Strip double quotes from ends (if present)
bstring$ = Trim$(bstring$)
If Left$(bstring$, 1) = VbDquote$ Then bstring$ = Mid$(bstring$, 2)
If Right$(bstring$, 1) = VbDquote$ Then bstring$ = Left$(bstring$, Len(bstring$) - 1)

Exit Sub

' Errors
MiscParseStringToStringError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscParseStringToString"
ierror = True
Exit Sub

MiscParseStringToStringEmpty:
msg$ = "Empty string"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscParseStringToString"
ierror = True
Exit Sub

End Sub

Sub MiscParseStringToStringA(astring As String, achar As String, bstring As String)
' Parse the first substring (achar delimited) in "astring" and return in "bstring"

ierror = False
On Error GoTo MiscParseStringToStringAError

Dim n As Long

' Check for more than one character in delimiter
If Len(achar$) > 1 Then GoTo MiscParseStringToStringANotSingleChar

' Check for empty string
astring$ = Trim$(astring$)
bstring$ = vbNullString
If astring$ = vbNullString Then GoTo MiscParseStringToStringAEmpty

' Parse out based on first delimiting character
n& = InStr(astring$, achar$)

' Load substring
If n& > 0 Then
bstring$ = Left$(astring$, n& - 1)
Else
bstring$ = Trim$(astring$)
End If

' Return remainder (without delimiting character)
If n& > 0 Then
astring$ = Mid$(astring$, n& + 1)
Else
astring$ = vbNullString   ' end of string
End If

' Strip double quotes from ends (if present)
bstring$ = Trim$(bstring$)
If Left$(bstring$, 1) = VbDquote$ Then bstring$ = Mid$(bstring$, 2)
If Right$(bstring$, 1) = VbDquote$ Then bstring$ = Left$(bstring$, Len(bstring$) - 1)

Exit Sub

' Errors
MiscParseStringToStringAError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscParseStringToStringA"
ierror = True
Exit Sub

MiscParseStringToStringAEmpty:
msg$ = "Empty string"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscParseStringToStringA"
ierror = True
Exit Sub

MiscParseStringToStringANotSingleChar:
msg$ = "Delimiting character string (" & achar$ & ") is more than one character in length"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscParseStringToStringA"
ierror = True
Exit Sub

End Sub

Sub MiscParseStringToStringB(astring As String, cstring As String, bstring As String)
' Parse the first substring (cstring delimited) in "astring" and return in "bstring"
' String cstring$ can be more than one character in length

ierror = False
On Error GoTo MiscParseStringToStringBError

Dim n As Long

' Check for bad delimiter
If Len(cstring$) = 0 Then GoTo MiscParseStringToStringBNoString

' Check for empty string
astring$ = Trim$(astring$)
bstring$ = vbNullString
If astring$ = vbNullString Then GoTo MiscParseStringToStringBEmpty

' Parse out based on delimiting string
n& = InStr(astring$, cstring$)

' Load substring
If n& > 0 Then
bstring$ = Left$(astring$, n& - 1)
Else
bstring$ = Trim$(astring$)
End If

' Return remainder (without delimiting character)
If n& > 0 Then
astring$ = Mid$(astring$, n& + Len(cstring$))
Else
astring$ = vbNullString   ' end of string
End If

' Strip double quotes from ends (if present)
bstring$ = Trim$(bstring$)
If Left$(bstring$, 1) = VbDquote$ Then bstring$ = Mid$(bstring$, 2)
If Right$(bstring$, 1) = VbDquote$ Then bstring$ = Left$(bstring$, Len(bstring$) - 1)

Exit Sub

' Errors
MiscParseStringToStringBError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscParseStringToStringB"
ierror = True
Exit Sub

MiscParseStringToStringBEmpty:
msg$ = "Empty string"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscParseStringToStringB"
ierror = True
Exit Sub

MiscParseStringToStringBNoString:
msg$ = "Delimiting character string (" & cstring$ & ") is less than one character in length"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscParseStringToStringB"
ierror = True
Exit Sub

End Sub

Sub MiscParseStringToStringT(astring As String, bstring As String)
' Parse the first substring (space delimited) in "astring" and return in "bstring". Same as MiscParseStringToString
' but with timed modal MgsBox.

ierror = False
On Error GoTo MiscParseStringToStringTError

Dim n As Long

' Check for empty string
astring$ = Trim$(astring$)
bstring$ = vbNullString
If astring$ = vbNullString Then GoTo MiscParseStringToStringTEmpty

' Parse out based on first space character
n& = InStr(astring$, VbSpace$)

' Load substring
If n& > 0 Then
bstring$ = Left$(astring$, n& - 1)
Else
bstring$ = Trim$(astring$)
End If

' Return remainder
If n& > 0 Then
astring$ = Mid$(astring$, n&)
Else
astring$ = Mid$(astring$, n& + 1)
End If

'  If end of string, blank returned string
If astring$ = bstring$ Then astring$ = vbNullString

' Strip double quotes from ends (if present)
bstring$ = Trim$(bstring$)
If Left$(bstring$, 1) = VbDquote$ Then bstring$ = Mid$(bstring$, 2)
If Right$(bstring$, 1) = VbDquote$ Then bstring$ = Left$(bstring$, Len(bstring$) - 1)

Exit Sub

' Errors
MiscParseStringToStringTError:
Call MiscMsgBoxTim(FormMSGBOXTIME, "MiscParseStringToStringT", Error$, CSng(20#))
ierror = True
Exit Sub

MiscParseStringToStringTEmpty:
msg$ = "Empty string"
Call MiscMsgBoxTim(FormMSGBOXTIME, "MiscParseStringToStringT", msg$, CSng(20#))
ierror = True
Exit Sub

End Sub

Sub MiscReplaceString(astring As String, achar As String, bchar As String)
' Replace all occurances of "achar$" with "bchar$" in a string (obsolete, use Replace$ function instead)

ierror = False
On Error GoTo MiscReplaceStringError

Dim k As Integer

' If "astring$" is empty just return
If astring$ = vbNullString Then Exit Sub

' If "achar$" equals "bchar$" just return
If achar$ = bchar$ Then Exit Sub

' Check that "achar$" and "bchar$" are equal in length
If Len(achar$) <> Len(bchar$) Then GoTo MiscReplaceStringDifferentSize

k% = 1
Do Until k% = 0
k% = InStr(astring$, achar$)  ' check for "achar$"
If k% > 0 Then Mid$(astring$, k%, Len(achar$)) = bchar$ ' replace with "bchar$"
Loop

Exit Sub

' Errors
MiscReplaceStringError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscReplaceString"
ierror = True
Exit Sub

MiscReplaceStringDifferentSize:
msg$ = "Strings are different length"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscReplaceString"
ierror = True
Exit Sub

End Sub

Sub MiscReplaceStringA(astring As String, achar As String, bchar As String)
' Replace all occurances of "achar$" with "bchar$" in a string (obsolete, use Replace$ function instead)

ierror = False
On Error GoTo MiscReplaceStringAError

Dim k As Integer

' If "astring$" is empty just return
If astring$ = vbNullString Then Exit Sub

' If "achar$" equals "bchar$" just return
If achar$ = bchar$ Then Exit Sub

' Check that "achar$" and "bchar$" are equal in length
If Len(achar$) <> Len(bchar$) Then GoTo MiscReplaceStringADifferentSize

k% = 1
Do Until k% = 0
k% = InStr(astring$, achar$)  ' check for "achar$"
If k% > 0 Then Mid$(astring$, k%, Len(achar$)) = bchar$ ' replace with "bchar$"
Loop

Exit Sub

' Errors
MiscReplaceStringAError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscReplaceStringA"
ierror = True
Exit Sub

MiscReplaceStringADifferentSize:
msg$ = "Strings are different length"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscReplaceStringA"
ierror = True
Exit Sub

End Sub

Sub MiscSelectText(tText As Control)
' Select the current text (called from GotFocus event)

ierror = False
On Error GoTo MiscSelectTextError

If tText Is Nothing Then Exit Sub

If Not TypeOf tText Is TextBox Then Exit Sub
tText.SelStart = 0
tText.SelLength = Len(tText)

Exit Sub

' Errors
MiscSelectTextError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSelectText"
ierror = True
Exit Sub

End Sub

Sub MiscSelectText2(mode As Integer, tText As Control)
' Select the current text (called from GotFocus and LostFocus events)
' mode = 1 make background yellow
' mode = 0 make background white

ierror = False
On Error GoTo MiscSelectText2Error

If tText Is Nothing Then Exit Sub

If Not TypeOf tText Is TextBox Then Exit Sub
If mode% = 1 Then tText.BackColor = vbYellow
If mode% = 0 Then tText.BackColor = vbWhite

Exit Sub

' Errors
MiscSelectText2Error:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSelectText2"
ierror = True
Exit Sub

End Sub

Function MiscSetSignificantDigits(inum As Integer, temp As Double) As Double
' Force the specified number of significant digits

ierror = False
On Error GoTo MiscSetSignificantDigitsError

Dim astring As String
Dim m As Integer, n As Integer
Dim notzero As Integer

' Load default
MiscSetSignificantDigits# = temp#

' Load into string
astring$ = Trim$(MiscAutoFormatD$(temp#))

' Loop and replace
notzero = False
n% = 0
For m% = 1 To Len(astring$)
If Mid$(astring$, m%, 1) <> "." And Mid$(astring$, m%, 1) <> "-" Then
If Mid$(astring$, m%, 1) <> "0" Then notzero = True
If notzero Then n% = n% + 1
If n% > inum% Then Mid$(astring$, m%, 1) = "0"
End If
Next m%

' Return value
MiscSetSignificantDigits# = Val(astring$)

Exit Function

' Errors
MiscSetSignificantDigitsError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSetSignificantDigits"
ierror = True
Exit Function

End Function

Sub MiscSortIntegerArray(ByVal n As Integer, inarray() As Integer, outarray() As Integer, arrayindex() As Integer)
' This routine accepts an integer array 'inarray' of length 'n'.  It returns
' in  'outarray' the array sorted in increasing order.  It returns in 'arrayindex'
' an indexing array which contains pointers to the elements of 'inarray' in
' increasing order. Example:
'  n        =  5
'  inarray  =  6,-1, 5, 3, 6
'  outarray = -1, 3, 5, 6, 6
'  arrayindex    =  2, 4, 3, 1, 5

ierror = False
On Error GoTo MiscSortIntegerArrayError

Dim itemp As Integer, i As Integer, j As Integer

For i% = 1 To n%
outarray%(i%) = inarray(i%)
arrayindex%(i%) = i%
Next i%

For i% = 1 To n% - 1
For j% = i% + 1 To n%
If outarray%(j%) >= outarray%(i%) Then GoTo 800
itemp% = outarray%(j%)
outarray%(j%) = outarray%(i%)
outarray%(i%) = itemp%

itemp% = arrayindex%(j%)
arrayindex%(j%) = arrayindex%(i%)
arrayindex%(i%) = itemp%
800:  Next j%
Next i%

Exit Sub

' Errors
MiscSortIntegerArrayError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSortIntegerArray"
ierror = True
Exit Sub

End Sub

Sub MiscSortRealArray(ByVal mode As Integer, ByVal n As Integer, inintarray() As Integer, outintarray() As Integer, insinarray() As Single, outsinarray() As Single)
' This routine sorts a real number array. The integer array is just along for the ride as an index
'  mode% = 1  sort by increasing real number order
'  mode% = 2  sort by decreasing real number order

ierror = False
On Error GoTo MiscSortRealArrayError

Dim i As Integer, j As Integer
Dim itemp As Integer
Dim temp As Single

For i% = 1 To n%
outintarray%(i%) = inintarray%(i%)
outsinarray!(i%) = insinarray!(i%)
Next i%

For i% = 1 To n% - 1
For j% = i% + 1 To n%
If mode% = 1 And outsinarray!(j%) >= outsinarray!(i%) Then GoTo 400
If mode% = 2 And outsinarray!(j%) <= outsinarray!(i%) Then GoTo 400
temp! = outsinarray!(j%)
outsinarray!(j%) = outsinarray!(i%)
outsinarray!(i%) = temp!

itemp% = outintarray%(j%)
outintarray%(j%) = outintarray%(i%)
outintarray%(i%) = itemp%

400:  Next j%
Next i%

Exit Sub

' Errors
MiscSortRealArrayError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSortRealArray"
ierror = True
Exit Sub

End Sub

Sub MiscSortStringArray(ByVal mode As Integer, ByVal n As Integer, instrarray() As String, outstrarray() As String, insinarray() As Single, outsinarray() As Single)
' This routine sorts a real number array; The string array is just along for the ride
'  mode% = 1  sort by increasing real number order
'  mode% = 2  sort by decreasing real number order

ierror = False
On Error GoTo MiscSortStringArrayError

Dim i As Integer, j As Integer
Dim atemp As String
Dim temp As Single

For i% = 1 To n%
outstrarray$(i%) = instrarray$(i%)
outsinarray!(i%) = insinarray!(i%)
Next i%

For i% = 1 To n% - 1
For j% = i% + 1 To n%
If mode% = 1 And outsinarray!(j%) >= outsinarray!(i%) Then GoTo 600
If mode% = 2 And outsinarray!(j%) <= outsinarray!(i%) Then GoTo 600
temp! = outsinarray!(j%)
outsinarray!(j%) = outsinarray!(i%)
outsinarray!(i%) = temp!

atemp$ = outstrarray$(j%)
outstrarray$(j%) = outstrarray$(i%)
outstrarray$(i%) = atemp$

600:  Next j%
Next i%

Exit Sub

' Errors
MiscSortStringArrayError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSortStringArray"
ierror = True
Exit Sub

End Sub

Function MiscStringsAreSame(ByVal astring As String, ByVal bstring As String) As Integer
' Checks if the two passed strings are the same (not case sensitive)

ierror = False
On Error GoTo MiscStringsAreSameError

' Assume strings are not the same
MiscStringsAreSame% = False

' Compare
If UCase$(Trim$(astring$)) = UCase$(Trim$(bstring$)) Then
MiscStringsAreSame% = True
End If

Exit Function

' Errors
MiscStringsAreSameError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscStringsAreSame"
ierror = True
Exit Function

End Function

Function MiscStringsAreSimilar(ByVal astring As String, ByVal bstring As String) As Integer
' Checks if the two passed strings are similar (same or contain one another)

ierror = False
On Error GoTo MiscStringsAreSimilarError

' Assume strings are not the similar
MiscStringsAreSimilar% = False

' Compare
If UCase$(Trim$(astring$)) = UCase$(Trim$(bstring$)) Then
MiscStringsAreSimilar% = True
End If

If InStr(UCase$(Trim$(astring$)), UCase$(Trim$(bstring$))) > 0 Then
MiscStringsAreSimilar% = True
End If

If InStr(UCase$(Trim$(bstring$)), UCase$(Trim$(astring$))) > 0 Then
MiscStringsAreSimilar% = True
End If

Exit Function

' Errors
MiscStringsAreSimilarError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscStringsAreSimilar"
ierror = True
Exit Function

End Function

Sub MiscDoEvents(ByVal numberofloops As Integer)
' Performs a specified number of Doevents

ierror = False
On Error GoTo MiscDoEventsError

Dim i As Integer

For i% = 1 To numberofloops%
DoEvents
Next i%

Exit Sub

' Errors
MiscDoEventsError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscDoEvents"
ierror = True
Exit Sub

End Sub

Function MiscAutoUcase(ByVal sym As String) As String
' Make the element or x-ray symbol upper case (1st character only)

ierror = False
On Error GoTo MiscAutoUcaseError

If Len(sym$) = 0 Then
Exit Function
ElseIf Len(sym$) = 1 Then
MiscAutoUcase$ = UCase$(Left$(sym$, 1))
Else
MiscAutoUcase$ = UCase$(Left$(sym$, 1)) & Right$(sym$, 1)
End If

Exit Function

' Errors
MiscAutoUcaseError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAutoUcase"
ierror = True
Exit Function

End Function

Function MiscAutoFormatD(ByVal treal As Double) As String
' Function to return an automatically formatted real number (double precision)

ierror = False
On Error GoTo MiscAutoFormatDError

Dim astring As String

' Negative numbers
If treal# < 0# Then
astring$ = f125$
If Abs(treal#) >= 1000# Then astring$ = f125$
If Abs(treal#) >= 10000# Then astring$ = f124$
If Abs(treal#) >= 100000# Then astring$ = f123$
If Abs(treal#) >= 1000000# Then astring$ = f122$
If Abs(treal#) >= 10000000# Then astring$ = f121$
If Abs(treal#) >= 100000000# Then astring$ = f120$

' Positive numbers
Else
astring$ = f126$
If treal# >= 1000# Then astring$ = f126$
If treal# >= 10000# Then astring$ = f125$
If treal# >= 100000# Then astring$ = f124$
If treal# >= 1000000# Then astring$ = f123$
If treal# >= 10000000# Then astring$ = f122$
If treal# >= 100000000# Then astring$ = f121$
If treal# >= 1000000000# Then astring$ = f120$
End If

' Format number
MiscAutoFormatD$ = Format$(Format$(treal#, astring$), a12$)

Exit Function

' Errors
MiscAutoFormatDError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAutoFormatD"
ierror = True
Exit Function

End Function

Function MiscAutoFormatQ(ByVal precision As Single, ByVal detectionlimit As Single, ByVal treal As Single) As String
' Function to return an automatically formatted real number (based on passed percent precision and detection limit)

ierror = False
On Error GoTo MiscAutoFormatQError

Dim bstring As String
Dim temp As Single

' Check for zero precision or detection limit
If precision! = 0# Or detectionlimit! = 0# Then
MiscAutoFormatQ$ = Format$(Format$(treal!, f83$), a80$)
Exit Function
End If

' Round the float value based on the percent precision
temp! = MiscAutoFormatZ!(precision!, treal!)
If ierror Then Exit Function

bstring$ = Format$(temp!, "General Number")

' Set to "n.d.", if less than detection limit
If treal! < detectionlimit! Then bstring$ = "n.d."

' Format number
MiscAutoFormatQ$ = Format$(bstring$, a80$)

Exit Function

' Errors
MiscAutoFormatQError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAutoFormatQ"
ierror = True
Exit Function

End Function

Function MiscAutoFormatZ(ByVal precision As Single, ByVal treal As Single) As Single
' Function to return an automatically rounded real number (based on passed percent error)

ierror = False
On Error GoTo MiscAutoFormatZError

Dim nChar As Integer, exponent As Integer
Dim ntemp As Long
Dim mantissa As Single, temp As Single
Dim astring As String

' Check for zero
If treal! = 0# Then Exit Function

' Determine number significant digits to save
nChar% = 1
If Abs(precision!) <= 100# Then nChar% = 1
If Abs(precision!) <= 10# Then nChar% = 2
If Abs(precision!) <= 1# Then nChar% = 3
If Abs(precision!) <= 0.1 Then nChar% = 4
If Abs(precision!) <= 0.01 Then nChar% = 5
If Abs(precision!) <= 0.001 Then nChar% = 6
If Abs(precision!) <= 0.0001 Then nChar% = 7

If treal! >= 100# Then nChar% = nChar% + 1

' Convert number to normalized format between 0 and 1
astring$ = Format$(treal!, e137$)   ' note: e137$ = "+.0000000e+00;-.0000000e+00"
mantissa! = Mid$(astring$, 1, 9)
exponent% = Mid$(astring$, 11, 3)

' Calculate as integer based on number of significant digits
temp! = mantissa! * 10 ^ nChar%

' Round to nearest integer and truncate (positive and negative numbers)
temp! = temp! + 0.5

ntemp& = Int(temp!)
temp! = ntemp&

' Calculate back from integer significant digits
temp! = temp! / 10 ^ nChar%

' Apply saved exponent from normalization to recover
temp! = temp! * 10 ^ exponent

' Format number
MiscAutoFormatZ! = Round(temp!, nChar%)                     ' need to perform an additional rounding here because of rare issues for scientific notation

' If above rounding causes a zero, load without additional rounding
If MiscAutoFormatZ! = 0# Then MiscAutoFormatZ! = temp!

Exit Function

' Errors
MiscAutoFormatZError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAutoFormatZ"
ierror = True
Exit Function

End Function

Function MiscElementToNumber(ByVal sym As String) As Integer
' Convert the element to an atomic number

ierror = False
On Error GoTo MiscElementToNumberError

Dim ip As Integer

ip% = IPOS1%(MAXELM%, sym$, Symlo$())
If ip% = 0 Then GoTo MiscElementToNumberBadSym
MiscElementToNumber = ip%

Exit Function

' Errors
MiscElementToNumberError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscElementToNumber"
ierror = True
Exit Function

MiscElementToNumberBadSym:
msg$ = "Invalid element symbol"
MsgBox msg$, vbOKOnly + vbExclamation, "MiscElementToNumber"
ierror = True
Exit Function

End Function

Function MiscGetSymbolFromString(astring As String) As String
' Return the first one or two character element chemical symbol

ierror = False
On Error GoTo MiscGetSymbolFromStringError

Dim i As Integer
Dim sym As String

' Load first two character match (case sensitive)
For i% = 1 To MAXELM%
If Left$(astring, 2) = Left$(Symup$(i%), 2) Then
sym$ = Left$(astring$, 2)
GoTo Found
End If
Next i%

' Load first single character match (case sensitive)
For i% = 1 To MAXELM%
If Left$(astring$, 1) = Left$(Symup$(i%), 1) Then
sym$ = Left$(astring$, 1)
GoTo Found
End If
Next i%

Found:
MiscGetSymbolFromString$ = sym$
Exit Function

' Errors
MiscGetSymbolFromStringError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetSymbolFromString"
ierror = True
Exit Function

End Function

Function MiscReplaceStringSub(astring As String, bstring As String, cstring As String) As String
' Replace all occurances of bstring$ with cstring$ in astring$ (need not be same length) (obsolete, use Replace$ function instead)

ierror = False
On Error GoTo MiscReplaceStringSubError

Dim i As Integer, k As Integer
Dim tstring As String

' If "astring$" is empty just return
If astring$ = vbNullString Then Exit Function

' If "bstring$" equals "cstring$" just return
If bstring$ = cstring$ Then Exit Function

k% = Len(bstring$)
i% = 0
tstring$ = vbNullString
Do Until i% > Len(astring$)
i% = i% + 1
If Mid$(astring$, i%, k%) = bstring$ Then
tstring$ = tstring$ & cstring$
i% = i% + (Len(bstring$) - 1)
Else
tstring$ = tstring$ & Mid$(astring$, i%, 1)
End If
Loop

MiscReplaceStringSub$ = tstring$

Exit Function

' Errors
MiscReplaceStringSubError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscReplaceStringSub"
ierror = True
Exit Function

End Function

Function IPOSDQ(ByVal n As Integer, ByVal sym1 As String, ByVal sym2 As String, sym1array() As String, sym2array() As String, dq() As Integer) As Integer
' This routine returns as its value a pointer to the first occurance (interference corrections only)
' of 'sym1' and 'sym2' in the character array 'sym1array' and 'sym2array' that are not disabled.
' The first 'n' positions in  'symarray' are searched.  If 'sym1' and 'sym2' does not occur in those
' positions IPOSDQ is equal to 0. However, if sym2 is empty, it is ignored. Example:
'  n = 4
'  sym1    = "zn"
'  sym2    = "la"
'  sym1array = "zn","zn","na","pb"
'  sym2array = "ka","la","ka","ma"
'  IPOSDQ    will be set to 2
'
'  n = 4
'  sym1    = "zn"
'  sym2    = ""
'  sym1array = "zn","zn","na","pb"
'  sym2array = "","","",""
'  IPOSDQ    will be set to 1

ierror = False
On Error GoTo IPOSDQError

Dim i As Integer

If n% <= 0 Then GoTo FailDQ
For i% = 1 To n%

' Check interfering element (and x-ray) (ProbeDataFileVersionNumber > 6.41)
If sym2$ <> vbNullString Then
If Trim$(LCase$(sym1array$(i%))) = Trim$(LCase$(sym1$)) And Trim$(LCase$(sym2array$(i%))) = Trim$(LCase$(sym2$)) And dq%(i%) = 0 Then GoTo FoundDQ

' Check only interfering element
Else
If Trim$(LCase$(sym1array$(i%))) = Trim$(LCase$(sym1$)) And dq%(i%) = 0 Then GoTo FoundDQ
End If
Next i%

FailDQ:
IPOSDQ = 0
Exit Function

FoundDQ:
IPOSDQ = i%
Exit Function

' Errors
IPOSDQError:
MsgBox Error$, vbOKOnly + vbCritical, "IPOSDQ"
ierror = True
Exit Function

End Function

Function IPOS1B(ByVal m As Integer, ByVal n As Integer, ByVal sym As String, symray() As String) As Integer
' This routine returns as its value a pointer to the first occurance
' of 'sym' in the character array 'symray'.  The 'm' through 'n' positions
' in  'symray' are searched.  If 'sym' does not occur in those positions
' IPOS1B is equal to 0. Example:
'  m = 3
'  n = 4
'  sym    = "si"
'  symray = "si","al","si","si"
'  IPOS1B    will be set to 3

ierror = False
On Error GoTo IPOS1BError

Dim i As Integer

If n% <= 0 Then GoTo Fail1B
For i% = m% To n%
If Trim$(LCase$(symray$(i%))) = Trim$(LCase$(sym$)) Then GoTo Found1B
Next i%

Fail1B:
IPOS1B = 0
Exit Function

Found1B:
IPOS1B = i%
Exit Function

' Errors
IPOS1BError:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS1B"
ierror = True
Exit Function

End Function

Function IPOS1DQ(ByVal n As Integer, ByVal sym As String, symarray() As String, dqarray() As Integer) As Integer
' This routine returns as its value a pointer to the first occurance
' of 'sym' in the character array 'symarray'.  The first 'n' positions
' in  'symarray' are searched and the disable flag is checked.
' If 'sym' does not occur in those positions IPOS1DQ is equal to 0. Example:
'  n = 4
'  sym    = "f"
'  symarray = "f","ca","f","fe"
'  dqarray = 1,0,0,0
'  IPOS1    will be set to 3

ierror = False
On Error GoTo IPOS1DQError

Dim i As Integer

If n% <= 0 Then GoTo Fail1DQ
For i% = 1 To n%
If Trim$(LCase$(symarray$(i%))) = Trim$(LCase$(sym$)) And dqarray%(i%) = 0 Then GoTo Found1DQ
Next i%

Fail1DQ:
IPOS1DQ = 0
Exit Function

Found1DQ:
IPOS1DQ = i%
Exit Function

' Errors
IPOS1DQError:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS1DQ"
ierror = True
Exit Function

End Function

Function MiscAutoFormat4(ByVal treal As Single) As String
' Function to return an automatically formatted real number in 4 characters

ierror = False
On Error GoTo MiscAutoFormat4Error

Dim astring As String

' Negative numbers
If treal! < 0# Then
astring$ = f42$
If Abs(treal!) >= 1# Then astring$ = f41$
If Abs(treal!) >= 10# Then astring$ = f40$

' Positive numbers
Else
astring$ = f43$
If treal! >= 1 Then astring$ = f42$
If treal! >= 10# Then astring$ = f41$
If treal! >= 100# Then astring$ = f40$
End If

' Format number
MiscAutoFormat4$ = Format$(Format$(treal!, astring$), a40$)

Exit Function

' Errors
MiscAutoFormat4Error:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAutoFormat4"
ierror = True
Exit Function

End Function

Function MiscAutoFormat6(ByVal treal As Single) As String
' Function to return an automatically formatted real number in 6 characters

ierror = False
On Error GoTo MiscAutoFormat6Error

Dim astring As String

' Negative numbers
If treal! < 0# Then
astring$ = f64$
If Abs(treal!) >= 1# Then astring$ = f63$
If Abs(treal!) >= 10# Then astring$ = f62$
If Abs(treal!) >= 100# Then astring$ = f61$
If Abs(treal!) >= 1000# Then astring$ = f60$

' Positive numbers
Else
astring$ = f65$
If treal! >= 1 Then astring$ = f64$
If treal! >= 10# Then astring$ = f63$
If treal! >= 100# Then astring$ = f62$
If treal! >= 1000# Then astring$ = f61$
If treal! >= 10000# Then astring$ = f60$
End If

' Format number
MiscAutoFormat6$ = Format$(Format$(treal!, astring$), a60$)

Exit Function

' Errors
MiscAutoFormat6Error:
MsgBox Error$, vbOKOnly + vbCritical, "MiscAutoFormat6"
ierror = True
Exit Function

End Function

Function IPOS1A(ByVal n As Integer, ByVal sym1 As String, ByVal sym2 As String, sym1array() As String, sym2array() As String) As Integer
' This routine returns as its value a pointer to the first occurance of 'sym1' and 'sym2' in the character
' arrays 'sym1array' and 'sym2array'. The first 'n' positions in  'sym1array' and sym2array' are searched.
' If 'sym1' and 'sym2' does not occur in those positions IPOS1A is equal to 0. Example:
'  n = 4
'  sym1    = "zn"
'  sym2    = "la"
'  sym1array = "zr","zn","na","pb"
'  sym2array = "la","la","ka","ma"
'  IPOS1A    will be set to 2

ierror = False
On Error GoTo IPOS1AError

Dim i As Integer

If n% <= 0 Then GoTo Fail1A
For i% = 1 To n%
If Trim$(LCase$(sym1array$(i%))) = Trim$(LCase$(sym1$)) And Trim$(LCase$(sym2array$(i%))) = Trim$(LCase$(sym2$)) Then GoTo Found1A
Next i%

Fail1A:
IPOS1A = 0
Exit Function

Found1A:
IPOS1A = i%
Exit Function

' Errors
IPOS1AError:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS1A"
ierror = True
Exit Function

End Function

Function IPOS1EDS(n As Integer, sym1 As String, sym2 As String, sym1array() As String, sym2array() As String, sample() As TypeSample) As Integer
' This routine returns as its value a pointer to the first occurance of an EDS element and x-ray in the sample() array
' If 'sym1' and and 'sym2' does not occur in those positions, IPOS1EDS is equal to 0. Example:
'  n = 4
'  sym1    = "zn"
'  sym2    = "la"
'  sym1array = "zr","zn","na","pb"
'  sym2array = "la","la","ka","ma"
'  IPOS1EDS    will be set to 2

ierror = False
On Error GoTo IPOS1EDSError

Dim i As Integer

' Check for vaild array
If n% <= 0 Then GoTo Fail1EDS

For i% = 1 To n%
If sample(1).DisableQuantFlag%(i%) = 0 Then
If sample(1).CrystalNames$(i%) = EDS_CRYSTAL$ Then
If Trim$(LCase$(sym1array$(i%))) = Trim$(LCase$(sym1$)) And Trim$(LCase$(sym2array$(i%))) = Trim$(LCase$(sym2$)) Then GoTo Found1EDS
End If
End If
Next i%

Fail1EDS:
IPOS1EDS = 0
Exit Function

Found1EDS:
IPOS1EDS = i%
Exit Function

' Errors
IPOS1EDSError:
MsgBox Error$, vbOKOnly + vbCritical, "IPOS1EDS"
ierror = True
Exit Function

End Function

Sub MiscParsePrivateProfileString(astring As String, nchars As Long, tcomment As String)
' Remove all comment characters after first semi-colon (including semi-colon)

ierror = False
On Error GoTo MiscParsePrivateProfileStringError

Dim n As Long
Dim bstring As String

' Check for empty string
tcomment$ = vbNullString
astring$ = Trim$(astring$)
If astring$ = vbNullString Or nchars& = 0 Then Exit Sub

' Parse out based on first semi-colon
n& = InStr(astring$, ";")

' Load substring
If n& > 0 Then
bstring$ = Left$(astring$, n& - 1)
bstring$ = Left$(bstring$, nchars&)
tcomment$ = Space$(20) & Right$(astring$, Len(astring$) - (n& - 1))    ' including semi-colon and extra 20 spaces before
Else
bstring$ = Left$(astring$, nchars&)
tcomment$ = vbNullString
End If

' Replace all nulls, tabs and quotes with spaces
bstring$ = Replace$(bstring$, vbNullChar, VbSpace$)
bstring$ = Replace$(bstring$, vbTab, VbSpace$)
bstring$ = Replace$(bstring$, VbDquote$, VbSpace$)

' Return string
astring$ = Trim$(bstring$)
nchars& = Len(astring$)
Exit Sub

' Errors
MiscParsePrivateProfileStringError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscParsePrivateProfileString"
ierror = True
Exit Sub

End Sub

Function MiscSetRounding(tvalue As Single, tsignificance As Integer) As Single
' Function to round a value to a specified significance (for numbers greater than 1)
'  e.g., MiscSetRounding(CSng(5248), int(100)) = 5200

ierror = False
On Error GoTo MiscSetRoundingError

Dim ntemp As Double

ntemp# = tvalue! / tsignificance%
ntemp# = CInt(ntemp#)
ntemp# = ntemp# * tsignificance%
    
MiscSetRounding! = CSng(ntemp#)
Exit Function

' Errors
MiscSetRoundingError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSetRounding"
ierror = True
Exit Function

End Function

Function MiscSetRounding2(tvalue As Single, ndigits As Integer) As Single
' Function to asymmetrically round a value to a specified number of decimal digits.
'  e.g., MiscSetRounding2(CSng(5.248), CInt(2)) = 5.25
'  e.g., MiscSetRounding2(CSng(-2.5), CInt(0)) = -2

ierror = False
On Error GoTo MiscSetRounding2Error

Dim nFactor As Double
Dim ntemp As Double

nFactor# = 10 ^ ndigits%
ntemp# = (tvalue! * nFactor#) + 0.5
    
MiscSetRounding2! = Int(CDec(ntemp#)) / nFactor#
Exit Function

' Errors
MiscSetRounding2Error:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSetRounding2"
ierror = True
Exit Function

End Function

Function MiscSetRounding3(tvalue As Single, ndigits As Integer) As Single
' Function to symmetrically round a value to a specified number of decimal digits
'  e.g., MiscSetRounding3(CSng(5.248), CInt(2)) = 5.25
'  e.g., MiscSetRounding3(CSng(-2.5), CInt(0)) = -3

ierror = False
On Error GoTo MiscSetRounding3Error

MiscSetRounding3! = Fix(tvalue! * (10 ^ ndigits%) + 0.5 * Sgn(tvalue!)) / (10 ^ ndigits%)
    
Exit Function

' Errors
MiscSetRounding3Error:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSetRounding3"
ierror = True
Exit Function

End Function

Function MiscSetRounding4(tvalue As Double, ndigits As Integer) As Double
' Function to symmetrically round a value to a specified double precision number of decimal digits
'  e.g., MiscSetRounding4(CDbl(5.248), CInt(2)) = 5.25
'  e.g., MiscSetRounding4(CDbl(-2.5), CInt(0)) = -3

ierror = False
On Error GoTo MiscSetRounding4Error

MiscSetRounding4# = Fix(tvalue# * (10 ^ ndigits%) + 0.5 * Sgn(tvalue#)) / (10 ^ ndigits%)
    
Exit Function

' Errors
MiscSetRounding4Error:
MsgBox Error$, vbOKOnly + vbCritical, "MiscSetRounding4"
ierror = True
Exit Function

End Function

Function MiscConvertBytesToString(barray() As Byte) As String
' Converts a byte array into a string array

ierror = False
On Error GoTo MiscConvertBytesToStringError

Dim n As Long
Dim astring As String

' Loop and convert
MiscConvertBytesToString$ = vbNullString
astring$ = vbNullString
For n& = 0 To UBound(barray)
astring$ = astring$ & Chr$(barray(n&))  ' only use ChrW$ for 0 to 127 ASCII
Next n&

MiscConvertBytesToString$ = astring$
Exit Function

' Errors
MiscConvertBytesToStringError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscConvertBytesToString"
ierror = True
Exit Function

End Function

Function MiscInstr(astring As String, bstring As String) As Integer
' Returns the starting position of the last occurance of bstring in astring

ierror = False
On Error GoTo MiscInstrError

Dim n As Integer

If Len(bstring) > Len(astring$) Then Exit Function
For n% = Len(astring$) To 1 Step -1
If Mid$(astring$, n%, Len(bstring$)) = bstring$ Then Exit For
Next n%

MiscInstr% = n%
Exit Function

' Errors
MiscInstrError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscInstr"
ierror = True
Exit Function

End Function

Function IPOSBoolean(n As Integer, tBoolean As Boolean, barray() As Boolean) As Integer
' This routine returns as its value a pointer to the first occurance
' of True in the boolean array 'barray'.  The first 'n' positions
' in  'barray' are searched.  If True does not occur in those positions
' IPOSBoolean is equal to 0. Example:
'  n = 4
'  tBoolean = True
'  iarray = False, False, True, False
'  IPOSBoolean    will be set to 3

ierror = False
On Error GoTo IPOSBooleanError

Dim i As Integer

If n% <= 0 Then GoTo Fail2
For i% = 1 To n%
If barray(i%) = tBoolean Then GoTo Found2
Next i%

Fail2:
IPOSBoolean = 0
Exit Function

Found2:
IPOSBoolean = i%
Exit Function

' Errors
IPOSBooleanError:
MsgBox Error$, vbOKOnly + vbCritical, "IPOSBoolean"
ierror = True
Exit Function

End Function

