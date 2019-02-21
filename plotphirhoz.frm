VERSION 5.00
Object = "{6E5043E8-C452-4A6A-B011-9B5687112610}#1.0#0"; "Pesgo32f.ocx"
Begin VB.Form FormPlotPhiRhoZ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plot Phi-Rhi-Z Curves"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   13080
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptionDepthMassOrMicrons 
      Caption         =   "Plot Micron Depth"
      Height          =   255
      Index           =   1
      Left            =   10560
      TabIndex        =   9
      Top             =   3600
      Width           =   2295
   End
   Begin VB.OptionButton OptionDepthMassOrMicrons 
      Caption         =   "Plot Mass Depth"
      Height          =   255
      Index           =   0
      Left            =   10560
      TabIndex        =   8
      Top             =   3360
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H00C0FFC0&
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CommandClipBoard 
      Caption         =   "Copy To ClipBoard"
      Height          =   375
      Left            =   10560
      TabIndex        =   4
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton CommandPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   10560
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton CommandZoomFull 
      BackColor       =   &H0080FFFF&
      Caption         =   "Zoom Full"
      Height          =   375
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton CommandSaveData 
      Caption         =   "Export Data"
      Height          =   375
      Left            =   10560
      TabIndex        =   1
      Top             =   2520
      Width           =   2295
   End
   Begin Pesgo32fLib.Pesgo Pesgo1 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      _Version        =   65536
      _ExtentX        =   18018
      _ExtentY        =   14631
      _StockProps     =   96
      _AllProps       =   "PlotPhiRhoZ.frx":0000
   End
   Begin VB.Label LabelXPos 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   10440
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label LabelYPos 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   11760
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "FormPlotPhiRhoZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2019 by John J. Donovan
Option Explicit

Private Sub CommandClipboard_Click()
If Not DebugMode Then On Error Resume Next
FormPlotPhiRhoZ.Pesgo1.AllowExporting = True
FormPlotPhiRhoZ.Pesgo1.ExportImageLargeFont = False
FormPlotPhiRhoZ.Pesgo1.ExportImageDpi = 400
Call FormPlotPhiRhoZ.Pesgo1.PEcopybitmaptoclipboard(FormPlotPhiRhoZ.Pesgo1.Width / 10, FormPlotPhiRhoZ.Pesgo1.Height / 10)
End Sub

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormPlotPhiRhoZ
End Sub

Private Sub CommandPrint_Click()
If Not DebugMode Then On Error Resume Next
Call MiscPlotPrintGraph_PE(FormPlotPhiRhoZ.Pesgo1)
If ierror Then Exit Sub
End Sub

Private Sub CommandSaveData_Click()
Call PlotPhiRhoZCurvesExport(FormPlotPhiRhoZ)
If ierror Then Exit Sub
End Sub

Private Sub CommandZoomFull_Click()
If Not DebugMode Then On Error Resume Next
FormPlotPhiRhoZ.Pesgo1.PEactions = UNDO_ZOOM&
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormPlotPhiRhoZ)
HelpContextID = IOGetHelpContextID("FormPlotPhiRhoZ")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub OptionDepthMassOrMicrons_Click(Index As Integer)
Call CalcZAFCalculate
If ierror Then Exit Sub
End Sub

Private Sub Pesgo1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not DebugMode Then On Error Resume Next
Dim fX As Double, fY As Double      ' mouse position in graph coordinates

' Get mouse position in data units
Call MiscPlotTrack(Int(1), x!, y!, fX#, fY#, FormPlotPhiRhoZ.Pesgo1)
If ierror Then Exit Sub
   
' Format graph mouse position
If fX# <> 0# And fY# <> 0# Then
   FormPlotPhiRhoZ.LabelXPos.Caption = MiscAutoFormat$(CSng(fX#))
   FormPlotPhiRhoZ.LabelYPos.Caption = MiscAutoFormat$(CSng(fY#))
Else
   FormPlotPhiRhoZ.LabelXPos.Caption = vbNullString
   FormPlotPhiRhoZ.LabelYPos.Caption = vbNullString
End If
End Sub

Private Sub Pesgo1_ZoomOut()
If Not DebugMode Then On Error Resume Next
FormPlotPhiRhoZ.Pesgo1.ManualScaleControlX = PEMSC_NONE&        ' automatically control X Axis
FormPlotPhiRhoZ.Pesgo1.ManualScaleControlY = PEMSC_NONE&        ' automatically control Y Axis
End Sub

