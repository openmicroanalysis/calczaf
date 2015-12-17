Attribute VB_Name = "CodeBMP4"
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

Private Declare Function BMPSendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long

Private Const TwipFactor& = 1440
Private Const WM_PAINT& = &HF
Private Const WM_PRINT& = &H317
Private Const PRF_CLIENT& = &H4&    ' Draw the window's client area
Private Const PRF_CHILDREN& = &H10& ' Draw all visible child windows
Private Const PRF_OWNED& = &H20&    ' Draw all owned windows

Sub BMPPrintDiagram(pic1 As PictureBox, pic2 As PictureBox, xp As Single, yp As Single, pcWidth As Single, pcHeight As Single)
' Prints the bitmap and graphics objects to the printer
' pic1 = picturebox to be printed
' pic2 = temporary picturebox
' xp, yp, pcWidth, pcHeight are the margins and dimensions of the image in inches

ierror = False
On Error GoTo BMPPrintDiagramError

Dim rv As Long

With pic2
   .Top = 0
   .Left = 0
   .Width = pic1.Width
   .Height = pic1.Height
End With

pic2.AutoRedraw = True
rv& = BMPSendMessage(pic1.hWnd, WM_PAINT, pic2.hdc, 0)
rv& = BMPSendMessage(pic1.hWnd, WM_PRINT, pic2.hdc, PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)

' Make pic2's image permanent
pic2.Picture = pic2.Image
pic2.AutoRedraw = False
Printer.Orientation = vbPRORLandscape
Printer.PaintPicture pic2.Picture, xp! * TwipFactor&, yp! * TwipFactor&, pcWidth! * TwipFactor, pcHeight * TwipFactor&
Printer.EndDoc

Exit Sub

' Errors
BMPPrintDiagramError:
msg$ = ". There is probably not enough video memory on the system video board. Try reducing the bit depth of the video display (Desktop | Properties | Settings) from 32 to 16 and try again."
MsgBox Error$ & msg$, vbOKOnly + vbCritical, "BMPPrintDiagram"
ierror = True
Exit Sub

End Sub


