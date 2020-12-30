VERSION 5.00
Object = "{6E5043E8-C452-4A6A-B011-9B5687112610}#1.0#0"; "Pesgo32f.ocx"
Begin VB.Form FormPLOTHISTO_PE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histogram Plot"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11025
   ControlBox      =   0   'False
   FillColor       =   &H80000013&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox CheckLabels 
      Caption         =   "Plot Labels"
      Height          =   255
      Left            =   9720
      TabIndex        =   15
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton CommandPlotConcentrationHistogram 
      BackColor       =   &H0080FFFF&
      Caption         =   "Plot Concentration Histogram"
      Height          =   855
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Plot a concentration histogram (a la Carpenter)"
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton CommandOptions 
      Caption         =   "Histogram Options"
      Height          =   495
      Left            =   9720
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to modify the histogram plot options"
      Top             =   4680
      Width           =   1215
   End
   Begin Pesgo32fLib.Pesgo Pesgo1 
      Height          =   7575
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   9495
      _Version        =   65536
      _ExtentX        =   16748
      _ExtentY        =   13361
      _StockProps     =   96
      _AllProps       =   "PLOTHISTO_PE.frx":0000
   End
   Begin VB.CommandButton CommandCopyToClipboard 
      Caption         =   "Copy To Clipboard"
      Height          =   495
      Left            =   9720
      TabIndex        =   11
      ToolTipText     =   "Click this button to copy the graph to the system clipboard"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.OptionButton OptionColumnNumber 
      Caption         =   "Element B"
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
      Left            =   9720
      TabIndex        =   2
      ToolTipText     =   "Plot the second element (B)"
      Top             =   960
      Width           =   1335
   End
   Begin VB.OptionButton OptionColumnNumber 
      Caption         =   "Element A"
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
      Left            =   9720
      TabIndex        =   1
      ToolTipText     =   "Plot the first element (A)"
      Top             =   720
      Value           =   -1  'True
      Width           =   1335
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
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Maximum"
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
      Left            =   9720
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label LabelMaximum 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9720
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Minimum"
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
      Left            =   9720
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label LabelMinimum 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9720
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label LabelStdDev 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9720
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "StdDev"
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
      Left            =   9720
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label LabelAverage 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9720
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Average"
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
      Left            =   9720
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "FormPLOTHISTO_PE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2021 by John J. Donovan
Option Explicit

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormPLOTHISTO_PE
End Sub

Private Sub CommandCopyToClipboard_Click()
If Not DebugMode Then On Error Resume Next
FormPLOTHISTO_PE.Pesgo1.AllowExporting = True
FormPLOTHISTO_PE.Pesgo1.ExportImageLargeFont = False
FormPLOTHISTO_PE.Pesgo1.ExportImageDpi = 400
Call FormPLOTHISTO_PE.Pesgo1.PEcopybitmaptoclipboard(FormPLOTHISTO_PE.Pesgo1.Width / 10, FormPLOTHISTO_PE.Pesgo1.Height / 10)
End Sub

Private Sub CommandOptions_Click()
If Not DebugMode Then On Error Resume Next
' Load the options
Call CalcZAFHistogramLoad
If ierror Then Exit Sub
FormHISTO.Show vbModal
End Sub

Private Sub CommandPlotConcentrationHistogram_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFPlotHistogramConcentration
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormPLOTHISTO_PE)
HelpContextID = IOGetHelpContextID("FormPLOTHISTO")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub OptionColumnNumber_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call CalcZAFPlotHistogram(Int(0))
If ierror Then Exit Sub
End Sub
