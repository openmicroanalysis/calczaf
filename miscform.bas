Attribute VB_Name = "CodeMISCForm"
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

' MiscMsgBoxAll globals
'vbOK       1   OK button pressed
'vbCancel   2   Cancel button pressed
'vbAbort    3   Abort button pressed
'vbRetry    4   Retry button pressed
'vbIgnore   5   Ignore button pressed
'vbYes      6   Yes button pressed
'vbNo       7   No button pressed
Global Const vbYesToAll% = 8        ' not specified by VB5

Global MsgBoxAllReturnValue As Integer

Private Declare Function HtmlHelp Lib "HHCtrl.ocx" Alias "HtmlHelpA" (ByVal hWndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwdata As Any) As Long

'Const HH_DISPLAY_TOPIC& = &H0
Const HH_HELP_CONTEXT& = &HF         ' Display mapped numeric value in dwData

'Const HH_SET_WIN_TYPE& = &H4
'Const HH_GET_WIN_TYPE& = &H5
'Const HH_GET_WIN_HANDLE& = &H6
'Const HH_DISPLAY_TEXT_POPUP& = &HE   ' Display string resource ID or text in a pop-up window
'Const HH_TP_HELP_CONTEXTMENU& = &H10 ' Text pop-up help, similar to WinHelp's HELP_CONTEXTMENU
'Const HH_TP_HELP_WM_HELP& = &H11     ' text pop-up help, similar to WinHelp's HELP_WM_HELP

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nindex As Long) As Long

Private Const SM_CXICON& = 11
Private Const SM_CYICON& = 12

Private Const SM_CXSMICON& = 49
Private Const SM_CYSMICON& = 50
   
Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" ( _
      ByVal hInst As Long, _
      ByVal lpsz As String, _
      ByVal uType As Long, _
      ByVal cxDesired As Long, _
      ByVal cyDesired As Long, _
      ByVal fuLoad As Long _
   ) As Long
   
'Private Const LR_DEFAULTCOLOR& = &H0
'Private Const LR_MONOCHROME& = &H1
'Private Const LR_COLOR& = &H2
'Private Const LR_COPYRETURNORG& = &H4
'Private Const LR_COPYDELETEORG& = &H8
'Private Const LR_LOADFROMFILE& = &H10
'Private Const LR_LOADTRANSPARENT& = &H20
'Private Const LR_DEFAULTSIZE& = &H40
'Private Const LR_VGACOLOR& = &H80
'Private Const LR_LOADMAP3DCOLORS& = &H1000
'Private Const LR_CREATEDIBSECTION& = &H2000
'Private Const LR_COPYFROMRESOURCE& = &H4000
Private Const LR_SHARED& = &H8000&

Private Const IMAGE_ICON& = 1
Private Const WM_SETICON& = &H80

Private Const ICON_SMALL& = 0
Private Const ICON_BIG& = 1
Private Const GW_OWNER& = 4

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Function MiscMsgBoxAll(tForm As Form, astring As String, bstring As String) As Integer
' "Yes to All" MsgBox function
' tForm is modal form
' astring is the form caption
' bstring is the label caption

ierror = False
On Error GoTo MiscMsgBoxAllError

tForm.Caption = astring$
tForm.Label1.Caption = bstring$

tForm.Show vbModal
MiscMsgBoxAll% = MsgBoxAllReturnValue%
Exit Function

' Errors
MiscMsgBoxAllError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscMsgBoxAll"
ierror = True
Exit Function

End Function

Sub MiscMsgBoxTim(tForm As Form, formstring As String, labelstring As String, timeinterval As Single)
' "Timed" MsgBox procedure
' tForm is modal form
' formstring is the form caption
' labelstring is the label caption
' timeinterval is time delay in seconds

ierror = False
On Error GoTo MiscMsgBoxTimError

' 65535 is maximum milliseconds for interval property
If timeinterval! > BIT16& / MSECPERSEC# Then timeinterval! = BIT16& / MSECPERSEC#

tForm.Caption = formstring$
tForm.Label1.Caption = labelstring$
tForm.Timer1.Interval = timeinterval! * MSECPERSEC#
tForm.Timer1.Enabled = True

tForm.Show vbModal
Exit Sub

' Errors
MiscMsgBoxTimError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscMsgBoxTim"
ierror = True
Exit Sub

End Sub

Sub MiscLoadIcon(tForm As Form)
' Load an icon for the form (uses .RES files)

ierror = False
On Error GoTo MiscLoadIconError

Dim App_Icon_Name As String

' Load icon based on application name
App_Icon_Name$ = vbNullString
If UCase$(app.EXEName) = UCase$("CalcImage") Then App_Icon_Name$ = "CalcImage_Icon"
If UCase$(app.EXEName) = UCase$("CalcZAF") Then App_Icon_Name$ = "CalcZAF_Icon"
If UCase$(app.EXEName) = UCase$("Probewin") Then App_Icon_Name$ = "Probewin_Icon"
If UCase$(app.EXEName) = UCase$("Stage") Then App_Icon_Name$ = "Stage_Icon"
If UCase$(app.EXEName) = UCase$("Standard") Then App_Icon_Name$ = "Standard_Icon"
If UCase$(app.EXEName) = UCase$("Startwin") Then App_Icon_Name$ = "Startwin_Icon"

If UCase$(app.EXEName) = UCase$("Calibrate") Then App_Icon_Name$ = "Calibrate_Icon"
If UCase$(app.EXEName) = UCase$("CalMAC") Then App_Icon_Name$ = "CalMAC_Icon"
If UCase$(app.EXEName) = UCase$("Coat") Then App_Icon_Name$ = "Coat_Icon"
If UCase$(app.EXEName) = UCase$("ConvertToPrbImg") Then App_Icon_Name$ = "ConvertToPrbImg_Icon"
If UCase$(app.EXEName) = UCase$("Drift") Then App_Icon_Name$ = "Drift_Icon"
If UCase$(app.EXEName) = UCase$("Evaluate") Then App_Icon_Name$ = "Evaluate_Icon"
If UCase$(app.EXEName) = UCase$("Faraday") Then App_Icon_Name$ = "Faraday_Icon"
If UCase$(app.EXEName) = UCase$("GunAlign") Then App_Icon_Name$ = "GunAlign_Icon"
If UCase$(app.EXEName) = UCase$("Matrix") Then App_Icon_Name$ = "Matrix_Icon"
If UCase$(app.EXEName) = UCase$("Monitor") Then App_Icon_Name$ = "Monitor_Icon"
If UCase$(app.EXEName) = UCase$("PenPFE") Then App_Icon_Name$ = "PenPFE_Icon"
If UCase$(app.EXEName) = UCase$("ProbeUserWizard") Then App_Icon_Name$ = "ProbeUserWizard_Icon"
If UCase$(app.EXEName) = UCase$("Remote") Then App_Icon_Name$ = "Remote_Icon"
If UCase$(app.EXEName) = UCase$("Search") Then App_Icon_Name$ = "Search_Icon"
If UCase$(app.EXEName) = UCase$("StripChart") Then App_Icon_Name$ = "StripChart_Icon"
If UCase$(app.EXEName) = UCase$("TestSX100") Then App_Icon_Name$ = "TestSX100_Icon"
If UCase$(app.EXEName) = UCase$("TestEDS") Then App_Icon_Name$ = "TestEDS_Icon"
If UCase$(app.EXEName) = UCase$("TestFid") Then App_Icon_Name$ = "TestFid_Icon"
If UCase$(app.EXEName) = UCase$("TestImage") Then App_Icon_Name$ = "TestImage_Icon"
If UCase$(app.EXEName) = UCase$("TestJEOL") Then App_Icon_Name$ = "TestJEOL_Icon"
If UCase$(app.EXEName) = UCase$("TestMatrix") Then App_Icon_Name$ = "TestMatrix_Icon"
If UCase$(app.EXEName) = UCase$("TestMonitor") Then App_Icon_Name$ = "TestMonitor_Icon"
If UCase$(app.EXEName) = UCase$("TestRemote") Then App_Icon_Name$ = "TestRemote_Icon"
If UCase$(app.EXEName) = UCase$("TestStage") Then App_Icon_Name$ = "TestStage_Icon"
If UCase$(app.EXEName) = UCase$("TestThermo") Then App_Icon_Name$ = "TestThermo_Icon"
If UCase$(app.EXEName) = UCase$("TestType") Then App_Icon_Name$ = "TestType_Icon"
If UCase$(app.EXEName) = UCase$("Userwin") Then App_Icon_Name$ = "Userwin_Icon"
If UCase$(app.EXEName) = UCase$("Vacuum") Then App_Icon_Name$ = "Vacuum_Icon"
If UCase$(app.EXEName) = UCase$("Wizard") Then App_Icon_Name$ = "Wizard_Icon"

' If nothing loaded use default icon
If App_Icon_Name$ = vbNullString Then Exit Sub

' Load form and application icon
Call MiscFormSetIcon(tForm.hWnd, App_Icon_Name$, True)
If ierror Then Exit Sub

Exit Sub

' Errors
MiscLoadIconError:
MsgBox Error$ & ", loading application icon " & App_Icon_Name$ & " for " & tForm.Name, vbOKOnly + vbCritical, "MiscLoadIcon"
ierror = True
Exit Sub

End Sub

Sub MiscCenterForm(tForm As Form)
' Center the passed form

ierror = False
On Error GoTo MiscCenterFormError

' Check if maxmimized or minimized
If tForm.WindowState = vbMaximized Then Exit Sub
If tForm.WindowState = vbMinimized Then Exit Sub

' Center the form
tForm.Left = (Screen.Width - tForm.Width) / 2
tForm.Top = (Screen.Height - tForm.Height) / 2

Exit Sub

' Errors
MiscCenterFormError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscCenterForm"
ierror = True
Exit Sub

End Sub

Sub MiscCopyList(mode As Integer, tList As ListBox)
' Copies data from the list to the clipboard
' mode = 1 copy list
' mode = 2 copy list (selected) only (use Selected property)

ierror = False
On Error GoTo MiscCopyListError

Dim i As Integer

' Copy List
msg$ = vbNullString
For i% = 0 To tList.ListCount - 1
If mode% = 1 Then
msg$ = msg$ & tList.List(i%) & vbCrLf
ElseIf mode% = 2 And tList.Selected(i%) Then
msg$ = msg$ & tList.List(i%) & vbCrLf
End If
Next i%

Clipboard.Clear
Sleep (200)     ' need for Win7 clipboard issues
Clipboard.SetText msg$

Exit Sub

' Errors
MiscCopyListError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscCopyList"
ierror = True
Exit Sub

End Sub

Function MiscGetClipboardFormat() As Integer
' Determine the clipboard format

ierror = False
On Error GoTo MiscGetClipboardFormatError

Dim clipformat As Integer

' Define bitmap formats
If Clipboard.GetFormat(vbCFText) Then clipformat% = clipformat% + 1
If Clipboard.GetFormat(vbCFBitmap) Then clipformat% = clipformat% + 2
If Clipboard.GetFormat(vbCFDIB) Then clipformat% = clipformat% + 4
If Clipboard.GetFormat(vbCFRTF) Then clipformat% = clipformat% + 8

Select Case clipformat%
        Case 1
            msg$ = "The Clipboard contains only text"
        Case 2, 4, 6
            msg$ = "The Clipboard contains only a bitmap"
        Case 3, 5, 7
            msg$ = "The Clipboard contains text and a bitmap"
        Case 8, 9
            msg$ = "The Clipboard contains only rich text"
        Case Else
            msg$ = "There is nothing on the Clipboard"
End Select
'MsgBox msg$, vbOKOnly, "MiscGetClipboardFormat"

MiscGetClipboardFormat = clipformat%
Exit Function

' Errors
MiscGetClipboardFormatError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscGetClipboardFormat"
ierror = True
Exit Function

End Function

Function MiscIsFormLoaded(fstring As String) As Boolean
' Checks to see if a form is already loaded (whether visible or not)

ierror = False
On Error GoTo MiscIsFormLoadedError

Dim objForm As Form
    
' Need to pass a string because referencing tForm.Name causes a load event in the passed form
MiscIsFormLoaded = False
For Each objForm In VB.Forms
  If Trim$(objForm.Name) = Trim$(fstring$) Then
      MiscIsFormLoaded = True
      Exit For
  End If
Next

Exit Function

' Errors
MiscIsFormLoadedError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscIsFormLoaded"
ierror = True
Exit Function

End Function

Public Sub MiscFormSetIcon(ByVal hWnd As Long, ByVal sIconResName As String, Optional ByVal bSetAsAppIcon As Boolean = True)
' Load the specified icon to the form

ierror = False
On Error GoTo MiscFormSetIconError

Dim lhWndTop As Long
Dim lhWnd As Long
Dim cX As Long
Dim cY As Long
Dim hIconLarge As Long
Dim hIconSmall As Long
      
' Find VB's hidden parent window:
    If (bSetAsAppIcon) Then
        lhWnd& = hWnd&
        lhWndTop& = lhWnd&
        Do While Not (lhWnd& = 0)
            lhWnd& = GetWindow(lhWnd&, GW_OWNER&)
            If Not (lhWnd& = 0) Then
            lhWndTop& = lhWnd&
        End If
        Loop
    End If
   
    cX& = GetSystemMetrics(SM_CXICON&)
    cY& = GetSystemMetrics(SM_CYICON&)
    
    hIconLarge& = LoadImageAsString(app.hInstance, sIconResName$, IMAGE_ICON&, cX&, cY&, LR_SHARED&)
    
    If (bSetAsAppIcon) Then
        SendMessageLong lhWndTop&, WM_SETICON&, ICON_BIG&, hIconLarge&
    End If
    
    SendMessageLong hWnd&, WM_SETICON&, ICON_BIG&, hIconLarge&
   
    cX& = GetSystemMetrics(SM_CXSMICON&)
    cY& = GetSystemMetrics(SM_CYSMICON&)
    hIconSmall& = LoadImageAsString(app.hInstance, sIconResName$, IMAGE_ICON&, cX&, cY&, LR_SHARED&)
         
    If (bSetAsAppIcon) Then
        SendMessageLong lhWndTop&, WM_SETICON&, ICON_SMALL&, hIconSmall&
    End If
    
    SendMessageLong hWnd&, WM_SETICON&, ICON_SMALL&, hIconSmall&
   
Exit Sub

' Errors
MiscFormSetIconError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscFormSetIcon"
ierror = True
Exit Sub

End Sub

Sub MiscFormLoadHelp(helpindex As Long)
' Load the .CHM Help file for the specified help index

ierror = False
On Error GoTo MiscFormLoadHelpError

Dim istatus As Long

If DebugMode Then
Call IOWriteLog("Help file : " & app.HelpFile & ", Help Context : " & Format$(helpindex&))
End If

istatus = HtmlHelp(FormMAIN.hWnd, app.HelpFile, HH_HELP_CONTEXT&, helpindex&)
'istatus& = HtmlHelp(FormMAIN.hWnd, app, HelpFile, HH_DISPLAY_TOPIC&, ByVal "Sample.html")
            
Exit Sub

' Errors
MiscFormLoadHelpError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscFormLoadHelp"
ierror = True
Exit Sub

End Sub

Function MiscFormControlHasIndex(ControlArray As Object, ByVal Index As Integer) As Boolean
' Checks if a specific index of a control array is loaded

ierror = False
On Error GoTo MiscFormControlHasIndexError

MiscFormControlHasIndex = (VarType(ControlArray(Index)) <> vbObject)

Exit Function

' Errors
MiscFormControlHasIndexError:
MsgBox Error$, vbOKOnly + vbCritical, "MiscFormControlHasIndex"
ierror = True
Exit Function

End Function
