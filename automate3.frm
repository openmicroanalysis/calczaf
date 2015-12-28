VERSION 5.00
Begin VB.Form FormAUTOMATE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hidden form"
   ClientHeight    =   6345
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CheckBox CheckDoNotBlankBeam 
      Height          =   195
      Left            =   840
      TabIndex        =   7
      ToolTipText     =   "Check this box to not blank the beam for stage moves using the Automate! window Go button"
      Top             =   4200
      Width           =   135
   End
   Begin VB.ComboBox ComboFiducial 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   "Select all samples with the specified fiducial set number"
      Top             =   360
      Width           =   510
   End
   Begin VB.CheckBox CheckBeamDeflection 
      Caption         =   "On Stds"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   5
      ToolTipText     =   "Acquire standard position samples using beam deflection"
      Top             =   3360
      Width           =   975
   End
   Begin VB.CheckBox CheckBeamDeflection 
      Caption         =   "On Unks"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   4
      ToolTipText     =   "Acquire unknown position samples using beam deflection"
      Top             =   3360
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox CheckBeamDeflection 
      Caption         =   "On Wavs"
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   3
      ToolTipText     =   "Acquire wavescan position samples using beam deflection"
      Top             =   3360
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.ListBox ListDigitize 
      Height          =   1425
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label LabelCurrentRow 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label LabelCurrentPosition 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   3975
   End
End
Attribute VB_Name = "FormAUTOMATE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormAUTOMATE)
HelpContextID = IOGetHelpContextID("FormAUTOMATE")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

