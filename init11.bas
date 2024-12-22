Attribute VB_Name = "CodeInit11"
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

Sub InitINIReadWriteArray(mode As Integer, tfilename As String, tSection As String, tkeyname As String, n As Integer, tarray() As Single)
' Open the passed INI type file and reads or writes an array of real numbers
'  mode = 1 read array
'  mode = 2 write array
'  n = number of points in tarray

ierror = False
On Error GoTo InitINIReadWriteArrayError

Dim i As Integer
Dim valid As Long, nSize As Long

Dim lpAppName As String
Dim lpKeyName As String
Dim lpString As String
Dim lpDefault As String
Dim lpFileName As String
Dim lpReturnString As String * 255

Dim astring As String, tcomment As String

' Check for existing file if mode = 1
If mode% = 1 Then
If Dir$(tfilename$) = vbNullString Then
msg$ = "Unable to open file " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIReadWriteArray"
ierror = True
Exit Sub
End If
End If

' Use Windows API function to read INI style file
lpFileName$ = tfilename$
nSize& = Len(lpReturnString$)
lpAppName$ = tSection$
lpKeyName$ = tkeyname$

' Read array (load default with array values)
If mode% = 1 Then
For i% = 1 To n%
If i% = 1 Then
lpDefault$ = Format$(tarray!(i%))
Else
lpDefault$ = lpDefault$ & "," & Format$(tarray!(i%))
End If
Next i%
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
astring$ = Left$(lpReturnString$, valid&)
If astring$ <> vbNullString Then
Call InitParseStringToReal(astring$, n%, tarray!())
If ierror Then Exit Sub
End If
End If

' Write array
If mode% = 2 Then
astring$ = vbNullString
For i% = 1 To n%
If i% = 1 Then
astring$ = astring$ & Trim$(Str$(tarray!(i%)))
Else
astring$ = astring$ & "," & Trim$(Str$(tarray!(i%)))
End If
Next i%
lpString$ = VbDquote$ & astring$ & VbDquote
valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, lpString$, lpFileName$)
End If

Exit Sub

' Errors
InitINIReadWriteArrayError:
MsgBox Error$, vbOKOnly + vbCritical, "InitINIReadWriteArray"
ierror = True
Exit Sub

End Sub

Sub InitINIReadWriteScaler(mode As Integer, tfilename As String, tSection As String, tkeyname As String, tvalue As Single)
' Open the passed INI type file and reads or writes a scaler value
'  mode = 1 read scaler value
'  mode = 2 write scaler value

ierror = False
On Error GoTo InitINIReadWriteScalerError

Dim valid As Long, tValid As Long, nSize As Long

Dim lpAppName As String
Dim lpKeyName As String
Dim lpString As String
Dim lpDefault As String
Dim lpFileName As String
Dim lpReturnString As String * 255

Dim astring As String, tcomment As String

' Check for existing file if mode = 1
If mode% = 1 Then
If Dir$(tfilename$) = vbNullString Then
msg$ = "Unable to open file " & tfilename$
MsgBox msg$, vbOKOnly + vbExclamation, "InitINIReadWriteScaler"
ierror = True
Exit Sub
End If
End If

' Use Windows API function to read INI style file
lpFileName$ = tfilename$
nSize& = Len(lpReturnString$)
lpAppName$ = tSection$
lpKeyName$ = tkeyname$

' Read scaler value
If mode% = 1 Then
lpDefault$ = tvalue!    '       assume passed value is default
tValid& = GetPrivateProfileString(lpAppName$, lpKeyName$, vbNullString, lpReturnString$, nSize&, lpFileName$)   ' check for keyword without default value
valid& = GetPrivateProfileString(lpAppName$, lpKeyName$, lpDefault$, lpReturnString$, nSize&, lpFileName$)
Call MiscParsePrivateProfileString(lpReturnString$, valid&, tcomment$)
astring$ = Left$(lpReturnString$, valid&)
tvalue! = Val(astring$)

If Left$(lpReturnString$, valid&) = vbNullString Then tValid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpDefault$ & VbDquote$ & tcomment$, lpFileName$)
End If

' Write scaler value
If mode% = 2 Then
astring$ = Format$(tvalue!)
lpString$ = astring$
valid& = WritePrivateProfileString(lpAppName$, lpKeyName$, VbDquote$ & lpString$ & VbDquote$, lpFileName$)
End If

Exit Sub

' Errors
InitINIReadWriteScalerError:
MsgBox Error$, vbOKOnly + vbCritical, "InitINIReadWriteScaler"
ierror = True
Exit Sub

End Sub

Sub InitParseStringToRealDelimit(astring As String, icount As Integer, realarray() As Single, tdelimit As String)
' Parse a string to a single precision array (based on delimiting character)

ierror = False
On Error GoTo InitParseStringToRealDelimitError

Dim i As Integer, n As Integer
Dim tstring As String

' Check for empty string
If astring$ = vbNullString Then GoTo InitParseStringToRealDelimitEmpty
If icount% < 1 Then GoTo InitParseStringToRealDelimitNoCount

' Parse out sub-strings based on delimit placement
n% = 1
For i% = 1 To Len(astring$)
If Mid$(astring$, i%, 1) <> tdelimit$ Then
tstring$ = tstring$ & Mid$(astring$, i%, 1)
Else
realarray!(n%) = Val(tstring$)
tstring$ = vbNullString
n% = n% + 1
If n% > icount% Then Exit Sub
End If
Next i%

' Load last sub-string
realarray!(n%) = Val(tstring$)

Exit Sub

' Errors
InitParseStringToRealDelimitError:
MsgBox Error$, vbOKOnly + vbCritical, "InitParseStringToRealDelimit"
ierror = True
Exit Sub

InitParseStringToRealDelimitEmpty:
msg$ = "Empty string"
MsgBox msg$, vbOKOnly + vbExclamation, "InitParseStringToRealDelimit"
ierror = True
Exit Sub

InitParseStringToRealDelimitNoCount:
msg$ = "Array count is less than one"
MsgBox msg$, vbOKOnly + vbExclamation, "InitParseStringToRealDelimit"
ierror = True
Exit Sub

End Sub

Sub InitINIReadWriteCommentString(mode As Integer, tfilename As String, tSection As String, tkeyname As String, tcomment As String)
' Read or write the comment string from or to the existing INI keyword
'   mode = 0 read
'   mode = 1 write

ierror = False
On Error GoTo InitINIReadWriteCommentStringError

Dim valid As Long
Dim nSize As Long

Dim lpDefault As String
Dim lpReturnString As String * 255

Dim m As Long, n As Long, nn As Long
Dim astring As String

' Init blank comment string if reading
If mode% = 0 Then
tcomment$ = vbNullString

' Get the whole INI entry
nSize& = Len(lpReturnString$)
lpDefault$ = vbNullString
valid& = GetPrivateProfileString(tSection$, tkeyname$, lpDefault$, lpReturnString$, nSize&, tfilename$)
If valid& <= 0 Then Exit Sub

' Search for semi colon character and check for no comment string
n& = InStr(lpReturnString$, ";")
If n& <= 0 Then Exit Sub

' Load commant string (with semi-colon)
tcomment$ = Mid$(lpReturnString$, n&, valid& - n& + 1)
End If

' Write comment string
If mode% = 1 Then

' Get the whole INI entry
nSize& = Len(lpReturnString$)
lpDefault$ = vbNullString
valid& = GetPrivateProfileString(tSection$, tkeyname$, lpDefault$, lpReturnString$, nSize&, tfilename$)
If valid& <= 0 Then Exit Sub

' Search for semi colon character in return string
astring$ = Left$(lpReturnString$, valid&)
astring$ = Trim$(astring$)
m& = InStr(astring$, ";")
If m& <= 0 Then m& = Len(astring$)

' Search for semi colon character in passed comment string
n& = InStr(tcomment$, ";")

' If no comment string found, just add todays date
If n& <= 0 Then
tcomment$ = " ; last modified " & Now
astring$ = astring$ & tcomment$

' If comment string found, replace
Else

' First check for "last modified" string in comment string
nn& = InStr(tcomment$, "last modified")
If nn& <= 0 Then
tcomment$ = " " & tcomment$ & ", last modified " & Now
Else
tcomment$ = " " & Left$(tcomment$, nn& - 1) & "last modified " & Now
End If
astring$ = Left$(astring$, m&) & tcomment$
End If

' Now write back to INI file
valid& = WritePrivateProfileString(tSection$, tkeyname$, astring$, tfilename$)
End If

Exit Sub

' Errors
InitINIReadWriteCommentStringError:
MsgBox Error$, vbOKOnly + vbCritical, "InitINIReadWriteCommentString"
ierror = True
Exit Sub

End Sub

Sub InitINIReadWriteString(mode As Integer, tfilename As String, tSection As String, tkeyname As String, tstring As String, tcomment As String)
' Read or write a string value from or to the existing INI keyword. If reading, it returns the string entry and comment string
' separately. If writing a string, it writes both the passed string and passed comment string.
'   mode = 0 read
'   mode = 1 write

ierror = False
On Error GoTo InitINIReadWriteStringError

Dim n As Long

Dim valid As Long
Dim nSize As Long

Dim lpDefault As String
Dim lpReturnString As String * 255
Dim astring As String

' Init blank comment string if reading
If mode% = 0 Then
tstring$ = vbNullString

' Get the INI entry
nSize& = Len(lpReturnString$)
lpDefault$ = vbNullString
valid& = GetPrivateProfileString(tSection$, tkeyname$, lpDefault$, lpReturnString$, nSize&, tfilename$)
If valid& <= 0 Then
tstring$ = vbNullString
tcomment$ = vbNullString
Exit Sub
End If

' Check for comment string
n& = InStr(lpReturnString$, ";")
If n& > 0 Then
tcomment$ = Trim$(Mid$(lpReturnString$, n&, valid& - n& + 1)) ' return with semi-colon as first character
Else
tcomment$ = vbNullString        ' if no semi-colon assume empty comment string
End If

' Load string (without comment string)
If n& > 0 Then
tstring$ = Trim$(Left$(lpReturnString$, n& - 1))
Else
tstring$ = Trim$(Left$(lpReturnString$, valid&))
End If

' Remove trailing tabs from string
n& = InStr(lpReturnString$, vbTab)
If n& > 0 Then
tstring$ = Left$(tstring$, n& - 1)
End If

' Remove double quotes (since not using MiscParsePrivateProfileString)
If Left$(tstring$, 1) = VbDquote Then tstring$ = Mid$(tstring$, 2)
If Right$(tstring$, 1) = VbDquote Then tstring$ = Left$(tstring$, Len(tstring$) - 1)
End If

' Write string
If mode% = 1 Then

' Search for semi colon character in comment string
If tcomment$ <> vbNullString Then
n& = InStr(tcomment$, ";")
If n& > 0 Then
tcomment$ = Mid$(tcomment$, n&)   ' save with semi-colon as first character
Else
tcomment$ = "; " & tcomment$    ' add semi-colon as first character
End If
End If

' Combine string with comment
If tcomment$ <> vbNullString Then
astring$ = tstring$ & " " & tcomment$
Else
astring$ = tstring$
End If

' Now write back to INI file
valid& = WritePrivateProfileString(tSection$, tkeyname$, astring$, tfilename$)
End If

Exit Sub

' Errors
InitINIReadWriteStringError:
MsgBox Error$, vbOKOnly + vbCritical, "InitINIReadWriteString"
ierror = True
Exit Sub

End Sub

Sub InitParseStringToInteger(astring As String, icount As Integer, integerarray() As Integer)
' Parse a string to a 2 byte (short) integer array

ierror = False
On Error GoTo InitParseStringToIntegerError

Dim i As Integer, n As Integer
Dim tstring As String

' Check for empty string
If astring$ = vbNullString Then GoTo InitParseStringToIntegerEmpty
If icount% < 1 Then GoTo InitParseStringToIntegerNoCount

' Parse out sub-strings based on comma placement
n% = 1
For i% = 1 To Len(astring$)
If Mid$(astring, i%, 1) <> VbComma$ Then
tstring$ = tstring$ & Mid$(astring, i%, 1)
Else
integerarray%(n%) = Val(tstring$)
tstring$ = vbNullString
n% = n% + 1
If n% > icount% Then Exit Sub
End If
Next i%

' Load last sub-string
integerarray%(n%) = Val(tstring$)

Exit Sub

' Errors
InitParseStringToIntegerError:
MsgBox Error$, vbOKOnly + vbCritical, "InitParseStringToInteger"
ierror = True
Exit Sub

InitParseStringToIntegerEmpty:
msg$ = "Empty string"
MsgBox msg$, vbOKOnly + vbExclamation, "InitParseStringToInteger"
ierror = True
Exit Sub

InitParseStringToIntegerNoCount:
msg$ = "Array count is less than one"
MsgBox msg$, vbOKOnly + vbExclamation, "InitParseStringToInteger"
ierror = True
Exit Sub

End Sub

Sub InitParseStringToLong(astring As String, nCount As Long, longarray() As Long)
' Parse a string to a 4 byte long array

ierror = False
On Error GoTo InitParseStringToLongError

Dim i As Long, n As Long
Dim tstring As String

' Check for empty string
If astring$ = vbNullString Then GoTo InitParseStringToLongEmpty
If nCount& < 1 Then GoTo InitParseStringToLongNoCount

' Parse out sub-strings based on comma placement
n& = 1
For i& = 1 To Len(astring$)
If Mid$(astring, i&, 1) <> VbComma$ Then
tstring$ = tstring$ & Mid$(astring, i&, 1)
Else
longarray&(n&) = Val(tstring$)
tstring$ = vbNullString
n& = n& + 1
If n& > nCount& Then Exit Sub
End If
Next i&

' Load last sub-string
longarray&(n&) = Val(tstring$)

Exit Sub

' Errors
InitParseStringToLongError:
MsgBox Error$, vbOKOnly + vbCritical, "InitParseStringToLong"
ierror = True
Exit Sub

InitParseStringToLongEmpty:
msg$ = "Empty string"
MsgBox msg$, vbOKOnly + vbExclamation, "InitParseStringToLong"
ierror = True
Exit Sub

InitParseStringToLongNoCount:
msg$ = "Array count is less than one"
MsgBox msg$, vbOKOnly + vbExclamation, "InitParseStringToLong"
ierror = True
Exit Sub

End Sub

Sub InitParseStringToString(astring As String, icount As Integer, stringarray() As String)
' Parse a comma delimited string to a string array

ierror = False
On Error GoTo InitParseStringToStringError

Dim i As Integer, n As Integer
Dim tstring As String

' Check for empty string
If astring$ = vbNullString Then GoTo InitParseStringToStringEmpty
If icount% < 1 Then GoTo InitParseStringToStringNoCount

' Parse out sub-strings based on comma placement
n% = 1
For i% = 1 To Len(astring$)
If Mid$(astring, i%, 1) <> VbComma$ Then
tstring$ = tstring$ & Mid$(astring, i%, 1)
Else
stringarray$(n%) = tstring$
tstring$ = vbNullString
n% = n% + 1
If n% > icount% Then Exit Sub
End If
Next i%

' Load last sub-string
stringarray$(n%) = tstring$

Exit Sub

' Errors
InitParseStringToStringError:
MsgBox Error$, vbOKOnly + vbCritical, "InitParseStringToString"
ierror = True
Exit Sub

InitParseStringToStringEmpty:
msg$ = "Empty string"
MsgBox msg$, vbOKOnly + vbExclamation, "InitParseStringToString"
ierror = True
Exit Sub

InitParseStringToStringNoCount:
msg$ = "Array count is less than one"
MsgBox msg$, vbOKOnly + vbExclamation, "InitParseStringToString"
ierror = True
Exit Sub

End Sub

Sub InitParseStringToString2(astring As String, nn As Integer, icount As Integer, stringarray() As String)
' Parse a comma delimited string to a string array (two dimensional array)

ierror = False
On Error GoTo InitParseStringToString2Error

Dim i As Integer, n As Integer
Dim tstring As String

' Check for empty string
If astring$ = vbNullString Then GoTo InitParseStringToString2Empty
If icount% < 1 Then GoTo InitParseStringToString2NoCount

' Parse out sub-strings based on comma placement
n% = 1
For i% = 1 To Len(astring$)
If Mid$(astring, i%, 1) <> VbComma$ Then
tstring$ = tstring$ & Mid$(astring, i%, 1)
Else
stringarray$(nn%, n%) = tstring$
tstring$ = vbNullString
n% = n% + 1
If n% > icount% Then Exit Sub
End If
Next i%

' Load last sub-string
stringarray$(nn%, n%) = tstring$

Exit Sub

' Errors
InitParseStringToString2Error:
MsgBox Error$, vbOKOnly + vbCritical, "InitParseStringToString2"
ierror = True
Exit Sub

InitParseStringToString2Empty:
msg$ = "Empty string"
MsgBox msg$, vbOKOnly + vbExclamation, "InitParseStringToString2"
ierror = True
Exit Sub

InitParseStringToString2NoCount:
msg$ = "Array count is less than one"
MsgBox msg$, vbOKOnly + vbExclamation, "InitParseStringToString2"
ierror = True
Exit Sub

End Sub

Sub InitParseStringToStringCount(astring As String, icount As Integer, stringarray() As String)
' Parse a comma delimited string to a string array and determine number of sub-strings

ierror = False
On Error GoTo InitParseStringToStringCountError

Dim i As Integer, n As Integer
Dim tstring As String

' Check for empty string
If astring$ = vbNullString Then GoTo InitParseStringToStringCountEmpty

' Dimension for single string to begin with
ReDim stringarray(1 To 1) As String

' Parse out sub-strings based on comma placement
n% = 1
For i% = 1 To Len(astring$)
If Mid$(astring, i%, 1) <> VbComma$ Then
tstring$ = tstring$ & Mid$(astring, i%, 1)
Else
stringarray$(n%) = Trim$(tstring$)
tstring$ = vbNullString
n% = n% + 1
If n% > 1 Then ReDim Preserve stringarray(1 To n%) As String
End If
Next i%

' Load last sub-string
stringarray$(n%) = Trim$(tstring$)
icount% = n%
Exit Sub

' Errors
InitParseStringToStringCountError:
MsgBox Error$, vbOKOnly + vbCritical, "InitParseStringToStringCount"
ierror = True
Exit Sub

InitParseStringToStringCountEmpty:
msg$ = "Empty string"
MsgBox msg$, vbOKOnly + vbExclamation, "InitParseStringToStringCount"
ierror = True
Exit Sub

End Sub

Sub InitParseStringToStringCount2(astring As String, icount As Integer, stringarray() As String)
' Parse a tab delimited string to a string array and determine number of sub-strings

ierror = False
On Error GoTo InitParseStringToStringCount2Error

Dim i As Integer, n As Integer
Dim tstring As String

' Check for empty string
If astring$ = vbNullString Then GoTo InitParseStringToStringCount2Empty

' Dimension for single string to begin with
ReDim stringarray(1 To 1) As String

' Parse out sub-strings based on tab placement
n% = 1
For i% = 1 To Len(astring$)
If Mid$(astring, i%, 1) <> vbTab$ Then
tstring$ = tstring$ & Mid$(astring, i%, 1)
Else
stringarray$(n%) = Trim$(tstring$)
tstring$ = vbNullString
n% = n% + 1
If n% > 1 Then ReDim Preserve stringarray(1 To n%) As String
End If
Next i%

' Load last sub-string
stringarray$(n%) = Trim$(tstring$)
icount% = n%
Exit Sub

' Errors
InitParseStringToStringCount2Error:
MsgBox Error$, vbOKOnly + vbCritical, "InitParseStringToStringCount2"
ierror = True
Exit Sub

InitParseStringToStringCount2Empty:
msg$ = "Empty string"
MsgBox msg$, vbOKOnly + vbExclamation, "InitParseStringToStringCount2"
ierror = True
Exit Sub

End Sub

Sub InitParseStringToReal(astring As String, icount As Integer, realarray() As Single)
' Parse a string to a single precision array

ierror = False
On Error GoTo InitParseStringToRealError

Dim i As Integer, n As Integer
Dim tstring As String

' Check for empty string
If astring$ = vbNullString Then GoTo InitParseStringToRealEmpty
If icount% < 1 Then GoTo InitParseStringToRealNoCount

' Remove comment
If InStr(astring$, ";") > 0 Then
astring$ = Left$(astring$, InStr(astring$, ";") - 1)
End If

astring$ = Trim$(astring$)

' Remove double quotes if found
If Left(astring$, 1) = VbDquote Then astring$ = Mid$(astring$, 2)
If Right(astring$, 1) = VbDquote Then astring$ = Left$(astring$, Len(astring$) - 1)

' Parse out sub-strings based on comma placement
n% = 1
For i% = 1 To Len(astring$)
If Mid$(astring$, i%, 1) <> VbComma$ Then
tstring$ = tstring$ & Mid$(astring$, i%, 1)
Else
realarray!(n%) = Val(tstring$)
tstring$ = vbNullString
n% = n% + 1
If n% > icount% Then Exit Sub
End If
Next i%

' Load last sub-string
realarray!(n%) = Val(tstring$)

Exit Sub

' Errors
InitParseStringToRealError:
MsgBox Error$, vbOKOnly + vbCritical, "InitParseStringToReal"
ierror = True
Exit Sub

InitParseStringToRealEmpty:
msg$ = "Empty string"
MsgBox msg$, vbOKOnly + vbExclamation, "InitParseStringToReal"
ierror = True
Exit Sub

InitParseStringToRealNoCount:
msg$ = "Array count is less than one"
MsgBox msg$, vbOKOnly + vbExclamation, "InitParseStringToReal"
ierror = True
Exit Sub

End Sub
