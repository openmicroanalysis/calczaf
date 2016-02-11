VERSION 5.00
Object = "{6E5043E8-C452-4A6A-B011-9B5687112610}#1.0#0"; "Pesgo32f.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormEDSDISPLAY3 
   AutoRedraw      =   -1  'True
   Caption         =   "EDS Spectrum Display"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   12345
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextStopkeV 
      Height          =   285
      Left            =   3120
      TabIndex        =   14
      Text            =   "20"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TextStartkeV 
      Height          =   285
      Left            =   1680
      TabIndex        =   13
      Text            =   "0"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CommandCopyToClipboard 
      Caption         =   "Copy To Clipboard"
      Height          =   495
      Left            =   4920
      TabIndex        =   10
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Frame FrameKLM 
      Caption         =   "KLM Markers"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7440
      TabIndex        =   4
      Top             =   4920
      Width           =   4815
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   2520
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Increment atomic number of specified KLM element"
         Top             =   600
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.ComboBox ComboSpecificElement 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton OptionKLM 
         Caption         =   "Specific Element"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Display KLM markers for a specific element"
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton OptionKLM 
         Caption         =   "All Elements"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Display KLM markers for all elements (you're just kiddin' right?)"
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton OptionKLM 
         Caption         =   "Analyzed Elements"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Display KLM markers for the specified EDS elements"
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton OptionKLM 
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Do not display any KLM markers"
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton CommandZoom 
      Caption         =   "Zoom Full"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H00008000&
      Caption         =   "Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   1095
   End
   Begin Pesgo32fLib.Pesgo Pesgo1 
      Height          =   4335
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   8895
      _Version        =   65536
      _ExtentX        =   15690
      _ExtentY        =   7646
      _StockProps     =   96
      _AllProps       =   "EDSDisplay3.frx":0000
   End
   Begin VB.Label LabelTrack 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   4695
   End
   Begin VB.Label LabelSpectrumName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   4695
   End
End
Attribute VB_Name = "FormEDSDISPLAY3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

Private Sub ComboSpecificElement_Change()
If Not DebugMode Then On Error Resume Next
Call EDSDisplayRedraw(FormEDSDISPLAY3)
If ierror Then Exit Sub
End Sub

Private Sub ComboSpecificElement_Click()
If Not DebugMode Then On Error Resume Next
Call EDSDisplayRedraw(FormEDSDISPLAY3)
If ierror Then Exit Sub
End Sub

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormEDSDISPLAY3
End Sub

Private Sub CommandCopyToClipboard_Click()
If Not DebugMode Then On Error Resume Next
FormEDSDISPLAY3.Pesgo1.DpiX = 300       ' might choose to use the PE global export functionality? press X on any open graph to see it!
FormEDSDISPLAY3.Pesgo1.DpiY = 300
FormEDSDISPLAY3.Pesgo1.ExportImageLargeFont = False ' = True for large font
Call FormEDSDISPLAY3.Pesgo1.PEcopybitmaptoclipboard(1400, 600)
End Sub

Private Sub CommandZoom_Click()
If Not DebugMode Then On Error Resume Next
Call EDSZoomFull(FormEDSDISPLAY3)
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
icancelload = False
Call InitWindow(Int(2), MDBUserName$, Me)
' Load form and application icon
Call MiscLoadIcon(Me)
End Sub

Private Sub Form_Resize()
If Not DebugMode Then On Error Resume Next
Dim temp As Single
Const TopOfButtons% = 1200

FormEDSDISPLAY3.LabelSpectrumName.Top = FormEDSDISPLAY3.ScaleHeight - TopOfButtons%
FormEDSDISPLAY3.LabelTrack.Top = FormEDSDISPLAY3.ScaleHeight - (TopOfButtons% - 300)

FormEDSDISPLAY3.CommandZoom.Top = FormEDSDISPLAY3.ScaleHeight - TopOfButtons%
FormEDSDISPLAY3.CommandClose.Top = FormEDSDISPLAY3.ScaleHeight - TopOfButtons%

FormEDSDISPLAY3.FrameKLM.Top = FormEDSDISPLAY3.ScaleHeight - TopOfButtons%
FormEDSDISPLAY3.CommandCopyToClipboard.Top = FormEDSDISPLAY3.ScaleHeight - (TopOfButtons% - 600)

' Make EDS display full size of window (PE)
FormEDSDISPLAY3.Pesgo1.Width = FormEDSDISPLAY3.ScaleWidth - FormEDSDISPLAY3.Pesgo1.Left * 2#
temp! = FormEDSDISPLAY3.ScaleHeight - (TopOfButtons% + WINDOWBORDERWIDTH%)
If temp! > 0# Then FormEDSDISPLAY3.Pesgo1.Height = temp!
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
If ierror Then Exit Sub
End Sub

Private Sub OptionKLM_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next
If Index% = 3 Then
FormEDSDISPLAY3.UpDown1.Enabled = True
FormEDSDISPLAY3.ComboSpecificElement.Enabled = True
Else
FormEDSDISPLAY3.UpDown1.Enabled = False
FormEDSDISPLAY3.ComboSpecificElement.Enabled = False
End If

' Redraw if data
Call EDSDisplayRedraw(FormEDSDISPLAY3)
If ierror Then Exit Sub

' Save the selection
Call EDSSaveKLM(FormEDSDISPLAY3)
If ierror Then Exit Sub
End Sub

Private Sub Pesgo1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not DebugMode Then On Error Resume Next
Dim astring As String
Dim fX As Double, fY As Double      ' last mouse position

' Get mouse position in data units
Call MiscPlotTrack(Int(1), X!, Y!, fX#, fY#, FormEDSDISPLAY3.Pesgo1)
If ierror Then Exit Sub
   
' Format graph mouse position
If fX# <> 0# And fY# <> 0# Then
   astring$ = FormEDSDISPLAY3.Pesgo1.XAxisLabel & " " & (CSng(fX#)) & ", " & FormEDSDISPLAY3.Pesgo1.YAxisLabel & " " & MiscAutoFormat$(CSng(fY#))
   FormEDSDISPLAY3.LabelTrack.Caption = astring$
Else
   FormEDSDISPLAY3.LabelTrack.Caption = vbNullString
End If
End Sub

Private Sub Pesgo1_ZoomIn()
If Not DebugMode Then On Error Resume Next
Call EDSRePlotKLM(FormEDSDISPLAY3)
If ierror Then Exit Sub
End Sub

Private Sub Pesgo1_ZoomOut()
If Not DebugMode Then On Error Resume Next
Call EDSRePlotKLM(FormEDSDISPLAY3)
If ierror Then Exit Sub
End Sub

Private Sub UpDown1_DownClick()
If Not DebugMode Then On Error Resume Next
Dim ip As Integer

' Increment element
If FormEDSDISPLAY3.ComboSpecificElement.ListCount < 1 Then Exit Sub

ip% = FormEDSDISPLAY3.ComboSpecificElement.ListIndex - 1
If ip% < 0 Then ip% = FormEDSDISPLAY3.ComboSpecificElement.ListCount - 1

' Change list element
FormEDSDISPLAY3.ComboSpecificElement.ListIndex = ip%
'FormEDSDISPLAY3.ComboSpecificElement.Refresh
End Sub

Private Sub UpDown1_UpClick()
If Not DebugMode Then On Error Resume Next
Dim ip As Integer

' Increment element
If FormEDSDISPLAY3.ComboSpecificElement.ListCount < 1 Then Exit Sub

ip% = FormEDSDISPLAY3.ComboSpecificElement.ListIndex + 1
If ip% > FormEDSDISPLAY3.ComboSpecificElement.ListCount - 1 Then ip% = 0

' Change list element
FormEDSDISPLAY3.ComboSpecificElement.ListIndex = ip%
'FormEDSDISPLAY3.ComboSpecificElement.Refresh
End Sub
