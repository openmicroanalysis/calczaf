VERSION 5.00
Object = "{827E9F53-96A4-11CF-823E-000021570103}#1.0#0"; "graphs32.ocx"
Begin VB.Form FormCLDISPLAY 
   AutoRedraw      =   -1  'True
   Caption         =   "CL Spectrum Display"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   12345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   12345
   StartUpPosition =   3  'Windows Default
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
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton CommandCopyToClipboard 
      Caption         =   "Copy To Clipboard"
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton CommandZoom 
      Caption         =   "Zoom Full"
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   4920
      Width           =   1095
   End
   Begin GraphsLib.Graph Graph1 
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   12135
      _Version        =   327680
      _ExtentX        =   21405
      _ExtentY        =   8281
      _StockProps     =   96
      BorderStyle     =   1
      Background      =   "15~-1~-1~-1~-1~-1~-1"
      ColorData       =   "1~1~1~1~1~1~1~1~1~1"
      GraphType       =   0
      NumPoints       =   10
      RandomData      =   0
      LabelYFormat    =   ""
   End
   Begin VB.Label LabelTrack 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   4
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
Attribute VB_Name = "FormCLDISPLAY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormCLDISPLAY
End Sub

Private Sub CommandCopyToClipboard_Click()
If Not DebugMode Then On Error Resume Next
FormCLDISPLAY.Graph1.DrawMode = 3     ' bit blit (DrawMode = 3) to put in bitmap format (otherwise metafile format does not work correctly)
DoEvents
FormCLDISPLAY.Graph1.DrawMode = 4     ' copy to clipboard
End Sub

Private Sub CommandZoom_Click()
If Not DebugMode Then On Error Resume Next
Call CLZoomFull(FormCLDISPLAY)
If ierror Then Exit Sub
Call CLSetBinSize(FormCLDISPLAY)
If ierror Then Exit Sub
Call CLDisplayRedraw
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
Dim temp As Single, TopOfButtons As Single

' Make EDS display full size of window
FormCLDISPLAY.Graph1.Width = FormCLDISPLAY.ScaleWidth - FormCLDISPLAY.Graph1.Left * 2#

TopOfButtons! = 600
FormCLDISPLAY.LabelSpectrumName.Top = FormCLDISPLAY.ScaleHeight - TopOfButtons!
FormCLDISPLAY.LabelTrack.Top = FormCLDISPLAY.ScaleHeight - (TopOfButtons! - 300)

FormCLDISPLAY.CommandZoom.Top = FormCLDISPLAY.ScaleHeight - TopOfButtons!
FormCLDISPLAY.CommandCopyToClipboard.Top = FormCLDISPLAY.ScaleHeight - TopOfButtons!

FormCLDISPLAY.CommandClose.Top = FormCLDISPLAY.ScaleHeight - TopOfButtons!
FormCLDISPLAY.CommandClose.Left = FormCLDISPLAY.ScaleWidth - (FormCLDISPLAY.CommandClose.Width + 200)

temp! = FormCLDISPLAY.ScaleHeight - (TopOfButtons! + WINDOWBORDERWIDTH%)
If temp! > 0# Then FormCLDISPLAY.Graph1.Height = temp!
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub Graph1_SDKPress(PressStatus As Integer, PressX As Double, PressY As Double, PressDataX As Double, PressDataY As Double)
If Not DebugMode Then On Error Resume Next
Call CLZoomGraph(PressStatus%, PressX#, PressY#, PressDataX#, PressDataY#, Int(0), FormCLDISPLAY)
If ierror Then Exit Sub
End Sub

Private Sub Graph1_SDKTrack(TrackX As Double, TrackY As Double, TrackDataX As Double, TrackDataY As Double)
If Not DebugMode Then On Error Resume Next

Dim astring As String
Dim tempx As Single, tempy As Single

tempx! = TrackDataX#
tempy! = TrackDataY#
astring$ = MiscAutoFormat$(tempx!) & ", " & MiscAutoFormat$(tempy!)
FormCLDISPLAY.LabelTrack.Caption = astring$

' Zoom rectangle
Call ZoomSDKTrack(TrackX#, TrackY#)
If ierror Then Exit Sub

End Sub

