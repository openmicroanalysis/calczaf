VERSION 5.00
Begin VB.Form FormCOND 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Analytical Conditions"
   ClientHeight    =   2400
   ClientLeft      =   1650
   ClientTop       =   2820
   ClientWidth     =   5025
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2400
   ScaleWidth      =   5025
   Begin VB.Frame Frame1 
      Caption         =   "Enter Default Analytical Conditions"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3495
      Begin VB.TextBox TextBeamSize 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         ToolTipText     =   "Beam size in microns"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox TextBeamCurrent 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Beam current in nA"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox TextKiloVolts 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         ToolTipText     =   "Operating voltage in kilovolt units"
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox TextTakeOff 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Take off angle (normally fixed)"
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Beam Size (um)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Beam Current (nA)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Kilovolts"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Take Off"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H0000C000&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "FormCOND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormCOND
icancelload = True
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call CondSave
If ierror Then Exit Sub
Unload FormCOND
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
icancelload = False
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormCOND)
HelpContextID = IOGetHelpContextID("FormCOND")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub TextBeamCurrent_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextBeamSize_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextKilovolts_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextTakeOff_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

