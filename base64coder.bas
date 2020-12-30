Attribute VB_Name = "CodeBase64Coder"
' (c) Copyright 1995-2021 by John J. Donovan
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Private InitDone  As Boolean
Private Map1(0 To 63)  As Byte
Private Map2(0 To 127) As Byte

Public Function Base64EncodeString(ByVal s As String) As String
' Encodes a string into Base64 format. No blanks or line breaks are inserted.
'   Parameters
'       S      a String to be encoded.
'   Returns    a String with the Base64 encoded data.
   
ierror = False
On Error GoTo Base64EncodeStringError
   
Base64EncodeString$ = Base64Encode(Base64ConvertStringToBytes(s$))
   
Exit Function

' Errors
Base64EncodeStringError:
MsgBox Error$, vbOKOnly + vbCritical, "Base64EncodeString"
ierror = True
Exit Function

End Function

Public Function Base64Encode(InData() As Byte)
' Encodes a byte array into Base64 format. No blanks or line breaks are inserted.
'   Parameters
'       InData    an array containing the data bytes to be encoded.
'   Returns      a string with the Base64 encoded data.
   
ierror = False
On Error GoTo Base64EncodeError
   
Base64Encode = Base64Encode2(InData, UBound(InData) - LBound(InData) + 1)

Exit Function

' Errors
Base64EncodeError:
MsgBox Error$, vbOKOnly + vbCritical, "Base64Encode"
ierror = True
Exit Function

End Function

Public Function Base64Encode2(InData() As Byte, ByVal InLen As Long) As String
' Encodes a byte array into Base64 format.
' No blanks or line breaks are inserted.
' Parameters:
'   InData    an array containing the data bytes to be encoded.
'   InLen     number of bytes to process in InData.
' Returns:    a string with the Base64 encoded data.
   
ierror = False
On Error GoTo Base64Encode2Error

If Not InitDone Then Base64Init

If InLen = 0 Then Base64Encode2 = "": Exit Function

Dim ODataLen As Long: ODataLen = (InLen * 4 + 2) \ 3     ' output length without padding
Dim OLen As Long: OLen = ((InLen + 2) \ 3) * 4           ' output length including padding
Dim Out() As Byte
   
ReDim Out(0 To OLen - 1) As Byte

Dim ip0 As Long: ip0 = LBound(InData)
Dim ip As Long
Dim op As Long
   
   Do While ip < InLen
      Dim i0 As Byte: i0 = InData(ip0 + ip): ip = ip + 1
      Dim i1 As Byte: If ip < InLen Then i1 = InData(ip0 + ip): ip = ip + 1 Else i1 = 0
      Dim i2 As Byte: If ip < InLen Then i2 = InData(ip0 + ip): ip = ip + 1 Else i2 = 0
      Dim o0 As Byte: o0 = i0 \ 4
      Dim o1 As Byte: o1 = ((i0 And 3) * &H10) Or (i1 \ &H10)
      Dim o2 As Byte: o2 = ((i1 And &HF) * 4) Or (i2 \ &H40)
      Dim o3 As Byte: o3 = i2 And &H3F
      Out(op) = Map1(o0): op = op + 1
      Out(op) = Map1(o1): op = op + 1
      Out(op) = IIf(op < ODataLen, Map1(o2), Asc("=")): op = op + 1
      Out(op) = IIf(op < ODataLen, Map1(o3), Asc("=")): op = op + 1
   Loop

Base64Encode2 = Base64ConvertBytesToString(Out)
   
Exit Function

' Errors
Base64Encode2Error:
MsgBox Error$, vbOKOnly + vbCritical, "Base64Encode2"
ierror = True
Exit Function

End Function

Public Function Base64DecodeString(ByVal s As String) As String
' Decodes a string from Base64 format.
' Parameters:
'    s        a Base64 String to be decoded.
' Returns     a String containing the decoded data.

ierror = False
On Error GoTo Base64DecodeStringError

If s$ = "" Then Base64DecodeString = "": Exit Function
   
Base64DecodeString$ = Base64ConvertBytesToString(Base64Decode(s$))
   
Exit Function

' Errors
Base64DecodeStringError:
MsgBox Error$, vbOKOnly + vbCritical, "Base64DecodeString"
ierror = True
Exit Function

End Function

Public Function Base64Decode(ByVal s As String) As Byte()
' Decodes a byte array from Base64 format.
' Parameters
'   s         a Base64 String to be decoded.
' Returns:    an array containing the decoded data bytes.

ierror = False
On Error GoTo Base64DecodeError

If Not InitDone Then Base64Init
   
Dim IBuf() As Byte: IBuf = Base64ConvertStringToBytes(s$)
Dim ILen As Long: ILen = UBound(IBuf) + 1
   
If ILen Mod 4 <> 0 Then GoTo Base64DecodeNotMultipleOf4
   
   Do While ILen > 0
      If IBuf(ILen - 1) <> Asc("=") Then Exit Do
      ILen = ILen - 1
   Loop
   
Dim OLen As Long: OLen = (ILen * 3) \ 4
Dim Out() As Byte
   
ReDim Out(0 To OLen - 1) As Byte
   
Dim ip As Long
Dim op As Long
   
   Do While ip < ILen
      Dim i0 As Byte: i0 = IBuf(ip): ip = ip + 1
      Dim i1 As Byte: i1 = IBuf(ip): ip = ip + 1
      Dim i2 As Byte: If ip < ILen Then i2 = IBuf(ip): ip = ip + 1 Else i2 = Asc("A")
      Dim i3 As Byte: If ip < ILen Then i3 = IBuf(ip): ip = ip + 1 Else i3 = Asc("A")
      If i0 > 127 Or i1 > 127 Or i2 > 127 Or i3 > 127 Then GoTo Base64DecodeInvalidCharacter
      
      Dim b0 As Byte: b0 = Map2(i0)
      Dim b1 As Byte: b1 = Map2(i1)
      Dim b2 As Byte: b2 = Map2(i2)
      Dim b3 As Byte: b3 = Map2(i3)
      If b0 > 63 Or b1 > 63 Or b2 > 63 Or b3 > 63 Then GoTo Base64DecodeInvalidCharacter
         
      Dim o0 As Byte: o0 = (b0 * 4) Or (b1 \ &H10)
      Dim o1 As Byte: o1 = ((b1 And &HF) * &H10) Or (b2 \ 4)
      Dim o2 As Byte: o2 = ((b2 And 3) * &H40) Or b3
      Out(op) = o0: op = op + 1
      If op < OLen Then Out(op) = o1: op = op + 1
      If op < OLen Then Out(op) = o2: op = op + 1
   Loop

Base64Decode = Out

Exit Function

' Errors
Base64DecodeError:
MsgBox Error$, vbOKOnly + vbCritical, "Base64Decode"
ierror = True
Exit Function

Base64DecodeNotMultipleOf4:
msg$ = "Length of Base64 encoded input string is not a multiple of 4."
MsgBox msg$, vbOKOnly + vbExclamation, "Base64Decode"
ierror = True
Exit Function

Base64DecodeInvalidCharacter:
msg$ = "Illegal character in Base64 encoded data."
MsgBox msg$, vbOKOnly + vbExclamation, "Base64Decode"
ierror = True
Exit Function

End Function

Private Sub Base64Init()
' Initialize the byte maps
   
ierror = False
On Error GoTo Base64InitError
   
Dim c As Integer, i As Integer

' Set Map1
   i = 0
   For c = Asc("A") To Asc("Z"): Map1(i) = c: i = i + 1: Next
   For c = Asc("a") To Asc("z"): Map1(i) = c: i = i + 1: Next
   For c = Asc("0") To Asc("9"): Map1(i) = c: i = i + 1: Next
   
   Map1(i) = Asc("+"): i = i + 1
   Map1(i) = Asc("/"): i = i + 1
   
' Set Map2
   For i = 0 To 127: Map2(i) = 255: Next
   For i = 0 To 63: Map2(Map1(i)) = i: Next
   
InitDone = True

Exit Sub

' Errors
Base64InitError:
MsgBox Error$, vbOKOnly + vbCritical, "Base64Init"
ierror = True
Exit Sub

End Sub

Private Function Base64ConvertStringToBytes(ByVal s As String) As Byte()
' Convert string to bytes

ierror = False
On Error GoTo Base64ConvertStringToBytesError

Dim b1() As Byte: b1 = s
Dim l As Long: l = (UBound(b1) + 1) \ 2
   
If l = 0 Then Base64ConvertStringToBytes = b1: Exit Function
   
Dim b2() As Byte
   
ReDim b2(0 To l - 1) As Byte

Dim p As Long
   
   For p = 0 To l - 1
      Dim c As Long: c = b1(2 * p) + 256 * CLng(b1(2 * p + 1))
      If c >= 256 Then c = Asc("?")
      b2(p) = c
   Next
   
Base64ConvertStringToBytes = b2

Exit Function

' Errors
Base64ConvertStringToBytesError:
MsgBox Error$, vbOKOnly + vbCritical, "Base64ConvertStringToBytes"
ierror = True
Exit Function

End Function

Private Function Base64ConvertBytesToString(b() As Byte) As String
' Convert bytes to string

ierror = False
On Error GoTo Base64ConvertBytesToStringError

Dim l As Long: l = UBound(b) - LBound(b) + 1
Dim b2() As Byte
   
ReDim b2(0 To (2 * l) - 1) As Byte

Dim p0 As Long: p0 = LBound(b)
Dim p As Long
   
For p = 0 To l - 1: b2(2 * p) = b(p0 + p): Next
   
Dim s As String: s = b2
   
Base64ConvertBytesToString = s
   
Exit Function

' Errors
Base64ConvertBytesToStringError:
MsgBox Error$, vbOKOnly + vbCritical, "Base64ConvertBytesToString"
ierror = True
Exit Function

End Function
