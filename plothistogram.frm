VERSION 5.00
Object = "{6E5043E8-C452-4A6A-B011-9B5687112610}#1.0#0"; "Pesgo32f.ocx"
Begin VB.Form FormPlotHistoConc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Concentration Histogram"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CommandZoomFull 
      Caption         =   "Zoom Full"
      Height          =   495
      Left            =   9000
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton CommandClipboard 
      Caption         =   "Copy To Clipboard"
      Height          =   615
      Left            =   9000
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H00C0FFC0&
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin Pesgo32fLib.Pesgo Pesgo1 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      _Version        =   65536
      _ExtentX        =   15478
      _ExtentY        =   13573
      _StockProps     =   96
      _AllProps       =   "PlotHistogram.frx":0000
   End
End
Attribute VB_Name = "FormPlotHistoConc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2024 by John J. Donovan
Option Explicit

Private Sub CommandClipboard_Click()
If Not DebugMode Then On Error Resume Next
FormPlotHistoConc.Pesgo1.AllowExporting = True
FormPlotHistoConc.Pesgo1.ExportImageLargeFont = False
FormPlotHistoConc.Pesgo1.ExportImageDpi = 400
Call FormPlotHistoConc.Pesgo1.PEcopybitmaptoclipboard(FormPlotHistoConc.Pesgo1.Width / 10, FormPlotHistoConc.Pesgo1.Height / 10)
End Sub

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormPlotHistoConc
End Sub

Private Sub CommandZoomFull_Click()
If Not DebugMode Then On Error Resume Next
FormPlotHistoConc.Pesgo1.PEactions = UNDO_ZOOM&
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormPlotHistoConc)
HelpContextID = IOGetHelpContextID("FormPlotHistoConc")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

