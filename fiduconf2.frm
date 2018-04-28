VERSION 5.00
Begin VB.Form FormFIDUCONF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modify Fiducial Positions"
   ClientHeight    =   2370
   ClientLeft      =   2580
   ClientTop       =   3750
   ClientWidth     =   8775
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2370
   ScaleWidth      =   8775
   Begin VB.Frame Frame1 
      Caption         =   "Enter Approximate Fiducial Positions For Fiducial Set "
      ClipControls    =   0   'False
      Height          =   2175
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7095
      Begin VB.TextBox TextDescription 
         Height          =   285
         Left            =   2160
         TabIndex        =   0
         Top             =   480
         Width           =   4695
      End
      Begin VB.TextBox TextZ 
         Height          =   285
         Index           =   2
         Left            =   4560
         TabIndex        =   10
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox TextZ 
         Height          =   285
         Index           =   1
         Left            =   4560
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox TextY 
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   8
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox TextY 
         Height          =   285
         Index           =   1
         Left            =   3360
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox TextX 
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox TextX 
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox TextZ 
         Height          =   285
         Index           =   0
         Left            =   4560
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TextY 
         Height          =   285
         Index           =   0
         Left            =   3360
         TabIndex        =   2
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TextX 
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   1
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         Caption         =   "Fiducial Description"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label LabelPoint 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label LabelPoint 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label LabelPoint 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Point#"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Z"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Y"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3360
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "X"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7440
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "FormFIDUCONF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2018 by John J. Donovan
Option Explicit

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormFIDUCONF
ierror = True
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call TestFidFiducialSaveConfirm
If ierror Then Exit Sub
Unload FormFIDUCONF
ierror = False
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormFIDUCONF)
HelpContextID = IOGetHelpContextID("FormFIDUCONF")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub TextDescription_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextX_GotFocus(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextY_GotFocus(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextZ_GotFocus(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

