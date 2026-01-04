VERSION 5.00
Begin VB.Form FormMAC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mass Absorption Coefficients"
   ClientHeight    =   2520
   ClientLeft      =   1575
   ClientTop       =   1800
   ClientWidth     =   7455
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2520
   ScaleWidth      =   7455
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame6 
      Caption         =   "Mass Absorption Coefficients File Source"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Default USERMAC/USERMAC2 files contain actinide MACs from Poeml/Wright"
         Top             =   1920
         Width           =   5415
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Contains MACs for additional lines Ln, Lg, Lv, Ll, Mg, Mz "
         Top             =   1560
         Width           =   5415
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1320
         Width           =   5415
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1080
         Width           =   5415
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   840
         Width           =   5415
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   360
         Width           =   5415
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   600
         Width           =   5415
      End
   End
End
Attribute VB_Name = "FormMAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2026 by John J. Donovan
Option Explicit

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormMAC
icancelload = True
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call GetZAFAllSaveMAC
If ierror Then Exit Sub
Unload FormMAC
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
icancelload = False
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormMAC)
HelpContextID = IOGetHelpContextID("FormMAC")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub
