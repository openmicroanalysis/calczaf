VERSION 5.00
Object = "{6E5043E8-C452-4A6A-B011-9B5687112610}#1.0#0"; "Pesgo32f.ocx"
Begin VB.Form FormCLDISPLAY 
   AutoRedraw      =   -1  'True
   Caption         =   "CL Spectrum Display"
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   13575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   13575
   StartUpPosition =   3  'Windows Default
   Begin Pesgo32fLib.Pesgo Pesgo1 
      Height          =   3855
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   9255
      _Version        =   65536
      _ExtentX        =   16325
      _ExtentY        =   6800
      _StockProps     =   96
      _AllProps       =   "CLDisplay.frx":0000
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
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton CommandCopyToClipboard 
      Caption         =   "Copy To Clipboard"
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton CommandZoom 
      Caption         =   "Zoom Full"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
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
Attribute VB_Name = "FormCLDISPLAY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormCLDISPLAY
End Sub

Private Sub CommandCopyToClipboard_Click()
If Not DebugMode Then On Error Resume Next
FormCLDISPLAY.Pesgo1.DpiX = 300       ' might choose to use the PE global export functionality? press X on any open graph to see it!
FormCLDISPLAY.Pesgo1.DpiY = 300
FormCLDISPLAY.Pesgo1.ExportImageLargeFont = False ' = True for large font
Call FormCLDISPLAY.Pesgo1.PEcopybitmaptoclipboard(1400, 600)
End Sub

Private Sub CommandZoom_Click()
If Not DebugMode Then On Error Resume Next
FormCLDISPLAY.Pesgo1.PEactions = UNDO_ZOOM&
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
TopOfButtons! = 600
FormCLDISPLAY.LabelSpectrumName.Top = FormCLDISPLAY.ScaleHeight - TopOfButtons!
FormCLDISPLAY.LabelTrack.Top = FormCLDISPLAY.ScaleHeight - (TopOfButtons! - 300)

FormCLDISPLAY.CommandZoom.Top = FormCLDISPLAY.ScaleHeight - TopOfButtons!
FormCLDISPLAY.CommandCopyToClipboard.Top = FormCLDISPLAY.ScaleHeight - TopOfButtons!

FormCLDISPLAY.CommandClose.Top = FormCLDISPLAY.ScaleHeight - TopOfButtons!
FormCLDISPLAY.CommandClose.Left = FormCLDISPLAY.ScaleWidth - (FormCLDISPLAY.CommandClose.Width + 200)

' Make CL display full size of window (PE)
FormCLDISPLAY.Pesgo1.Width = FormCLDISPLAY.ScaleWidth - FormCLDISPLAY.Pesgo1.Left * 2#
temp! = FormCLDISPLAY.ScaleHeight - (TopOfButtons! + WINDOWBORDERWIDTH%)
If temp! > 0# Then FormCLDISPLAY.Pesgo1.Height = temp!
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub Pesgo1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not DebugMode Then On Error Resume Next
Dim astring As String
Dim fX As Double, fY As Double      ' last mouse position

' Get mouse position in data units
Call MiscPlotTrack(Int(1), X!, Y!, fX#, fY#, FormCLDISPLAY.Pesgo1)
If ierror Then Exit Sub
   
' Format graph mouse position
If fX# <> 0# And fY# <> 0# Then
   astring$ = FormCLDISPLAY.Pesgo1.XAxisLabel & " " & (CSng(fX#)) & ", " & FormCLDISPLAY.Pesgo1.YAxisLabel & " " & MiscAutoFormat$(CSng(fY#))
   FormCLDISPLAY.LabelTrack.Caption = astring$
Else
   FormCLDISPLAY.LabelTrack.Caption = vbNullString
End If
End Sub
