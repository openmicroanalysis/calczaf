VERSION 5.00
Object = "{827E9F53-96A4-11CF-823E-000021570103}#1.0#0"; "graphs32.ocx"
Begin VB.Form FormPLOTHISTO_GS 
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
   Begin VB.CommandButton CommandOptions 
      Caption         =   "Histogram Options"
      Height          =   495
      Left            =   9720
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to modify the histogram plot options"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton CommandCopyToClipboard 
      Caption         =   "Copy To Clipboard"
      Height          =   495
      Left            =   9720
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to copy the graph to the system clipboard"
      Top             =   4320
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
      TabIndex        =   3
      ToolTipText     =   "Plot the second element (B)"
      Top             =   7440
      Width           =   1215
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
      TabIndex        =   2
      ToolTipText     =   "Plot the first element (A)"
      Top             =   7200
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H0000C000&
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
   End
   Begin GraphsLib.Graph Graph1 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _Version        =   327680
      _ExtentX        =   16748
      _ExtentY        =   13361
      _StockProps     =   96
      BorderStyle     =   1
      Background      =   "15~-1~-1~-1~-1~-1~-1"
      GraphType       =   3
      GridStyle       =   3
      LabelEvery      =   5
      NumPoints       =   50
      Palette         =   1
      PrintStyle      =   1
      Toolbar         =   2
      XAxisStyle      =   2
      YAxisMinorTicks =   "1~0"
      Bar2DGap        =   1
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
      TabIndex        =   11
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label LabelMaximum 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9720
      TabIndex        =   10
      Top             =   3120
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
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label LabelMinimum 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9720
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label LabelStdDev 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9720
      TabIndex        =   7
      Top             =   1680
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
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label LabelAverage 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9720
      TabIndex        =   5
      Top             =   960
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
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "FormPLOTHISTO_GS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormPLOTHISTO_GS
End Sub

Private Sub CommandCopyToClipboard_Click()
If Not DebugMode Then On Error Resume Next
FormPLOTHISTO_GS.Graph1.DrawMode = 3     ' blit to put in bitmap format (otherwise metafile format does not work correctly)
DoEvents
FormPLOTHISTO_GS.Graph1.DrawMode = 4     ' copy to clipboard
End Sub

Private Sub CommandOptions_Click()
If Not DebugMode Then On Error Resume Next
' Load the options
Call CalcZAFHistogramLoad
If ierror Then Exit Sub
FormHISTO.Show vbModal
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormPLOTHISTO_GS)
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
