VERSION 5.00
Begin VB.Form FormMODAL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modal Analysis"
   ClientHeight    =   7680
   ClientLeft      =   2160
   ClientTop       =   975
   ClientWidth     =   7410
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7680
   ScaleWidth      =   7410
   Begin VB.Frame Frame4 
      Caption         =   "Phase Options"
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   3720
      TabIndex        =   26
      Top             =   3240
      Width           =   3615
      Begin VB.CommandButton Command11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Update Phase"
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2775
      End
      Begin VB.OptionButton OptionEndMember 
         Caption         =   "Garnet"
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   29
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton OptionEndMember 
         Caption         =   "Pyroxene"
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   33
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton OptionEndMember 
         Caption         =   "Feldspar"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   32
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton OptionEndMember 
         Caption         =   "Olivine"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   31
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton OptionEndMember 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox TextMinimumVector 
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         Caption         =   "Minimum Vector"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Start"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Data Files"
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   7215
      Begin VB.TextBox TextOutputDataFile 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   5655
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Browse"
         Height          =   375
         Left            =   6000
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Browse"
         Height          =   375
         Left            =   6000
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox TextInputDataFile 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Ouput Data File"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Input Data File"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Group Options"
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   3495
      Begin VB.CheckBox CheckWeight 
         Caption         =   "Weight Concentrations For Fit"
         Height          =   252
         Left            =   120
         TabIndex        =   37
         Top             =   1440
         Width           =   3252
      End
      Begin VB.CheckBox CheckNormalize 
         Caption         =   "Normalize Concentrations For Fit"
         Height          =   252
         Left            =   120
         TabIndex        =   36
         Top             =   1200
         Width           =   3252
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Update Group"
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2775
      End
      Begin VB.CheckBox CheckDoEndMembers 
         Caption         =   "Do End-Member Calculations"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox TextMinimumTotal 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         Caption         =   "Minimum Total for Input"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000C000&
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   615
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Group Definitions"
      ForeColor       =   &H00FF0000&
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.CommandButton Command10 
         Caption         =   "Remove"
         Height          =   255
         Left            =   3960
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Add"
         Height          =   255
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Delete"
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "New"
         Height          =   255
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Delete"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "New"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1815
      End
      Begin VB.ListBox ListStandards 
         Height          =   1620
         Left            =   3960
         Sorted          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   600
         Width           =   1812
      End
      Begin VB.ListBox ListPhases 
         Height          =   1620
         Left            =   2040
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   600
         Width           =   1812
      End
      Begin VB.ListBox ListGroups 
         Height          =   1620
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   600
         Width           =   1812
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   "Standards"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         TabIndex        =   22
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         Caption         =   "Phases"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         Caption         =   "Groups"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "FormMODAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

Private Sub Command1_Click()
If Not DebugMode Then On Error Resume Next
Call ModalGetInputDataFile(FormMODAL)
If ierror Then Exit Sub
End Sub

Private Sub Command10_Click()
If Not DebugMode Then On Error Resume Next
Call ModalStandardRemove
If ierror Then Exit Sub
End Sub

Private Sub Command11_Click()
If Not DebugMode Then On Error Resume Next
Call ModalSaveOptionsPhase
If ierror Then Exit Sub
End Sub

Private Sub Command12_Click()
If Not DebugMode Then On Error Resume Next
Call ModalSaveOptionsGroup
If ierror Then Exit Sub
End Sub

Private Sub Command2_Click()
If Not DebugMode Then On Error Resume Next
Call ModalSaveForm
If ierror Then Exit Sub
Unload FormMODAL
End Sub

Private Sub Command3_Click()
If Not DebugMode Then On Error Resume Next
Call ModalSaveForm
If ierror Then Exit Sub
Call ModalStartModal
If ierror Then Exit Sub
End Sub

Private Sub Command4_Click()
If Not DebugMode Then On Error Resume Next
Call ModalGetOutputDataFile(FormMODAL)
If ierror Then Exit Sub
End Sub

Private Sub Command5_Click()
If Not DebugMode Then On Error Resume Next
Call ModalGroupNew
If ierror Then Exit Sub
End Sub

Private Sub Command6_Click()
If Not DebugMode Then On Error Resume Next
Call ModalGroupDelete
If ierror Then Exit Sub
End Sub

Private Sub Command7_Click()
If Not DebugMode Then On Error Resume Next
Call ModalPhaseNew
If ierror Then Exit Sub
End Sub

Private Sub Command8_Click()
If Not DebugMode Then On Error Resume Next
Call ModalPhaseDelete
If ierror Then Exit Sub
End Sub

Private Sub Command9_Click()
If Not DebugMode Then On Error Resume Next
Call ModalStandardAdd
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormMODAL)
HelpContextID = IOGetHelpContextID("FormMODAL")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub ListGroups_Click()
If Not DebugMode Then On Error Resume Next
Call ModalUpdatePhases
If ierror Then Exit Sub
End Sub

Private Sub ListPhases_Click()
If Not DebugMode Then On Error Resume Next
Call ModalUpdateStandards
If ierror Then Exit Sub
End Sub

Private Sub TextInputDataFile_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMinimumTotal_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMinimumVector_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextOutputDataFile_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

