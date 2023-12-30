VERSION 5.00
Begin VB.Form FormEffective 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculate K-ratios for a Range of Effective Takeoff Angles"
   ClientHeight    =   5265
   ClientLeft      =   750
   ClientTop       =   4290
   ClientWidth     =   12225
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
   ScaleHeight     =   5265
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CommandCalculate 
      BackColor       =   &H0080FFFF&
      Caption         =   "Calculate K-Ratios"
      Height          =   735
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Calculates k-ratios for a range of effective takeoff angles"
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Standard Compositions"
      ForeColor       =   &H00FF0000&
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9375
      Begin VB.CommandButton CommandFindNextString2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Next Match"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton CommandFindNextNumber2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Next Match"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton CommandFindNextString1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Next Match"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton CommandFindNextNumber1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Next Match"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox TextStandardString2 
         Height          =   285
         Left            =   5280
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Type a few characters of the standard name and the program will automatically select it"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox TextStandardNumber2 
         Height          =   285
         Left            =   7200
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Type a few characters of the standard name and the program will automatically select it"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox TextStandardString1 
         Height          =   285
         Left            =   600
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Type a few characters of the standard name and the program will automatically select it"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox TextStandardNumber1 
         Height          =   285
         Left            =   2520
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Type a few characters of the standard name and the program will automatically select it"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.ListBox ListStandardSecondary 
         Height          =   1815
         ItemData        =   "Effective.frx":0000
         Left            =   4800
         List            =   "Effective.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   600
         Width           =   4335
      End
      Begin VB.ListBox ListStandardPrimary 
         Height          =   1815
         ItemData        =   "Effective.frx":0004
         Left            =   240
         List            =   "Effective.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label12 
         Caption         =   "Enter String To Find:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   27
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Enter Number To Find:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   26
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Enter String To Find:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Enter Number To Find:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   22
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Select the secondary standard composition"
         Height          =   255
         Left            =   4800
         TabIndex        =   7
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Select the primary standard composition"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Analytical Conditions"
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   9375
      Begin VB.TextBox TextBeamEnergy 
         Height          =   285
         Left            =   7800
         TabIndex        =   19
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TextTakeoffIncrement 
         Height          =   285
         Left            =   5880
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TextTakeoffHigh 
         Height          =   285
         Left            =   4440
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TextTakeoffLow 
         Height          =   285
         Left            =   3000
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox ComboXRay 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox ComboElement 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Increment"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5760
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "High Takeoff"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Beam Energy (keV)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7440
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Low Takeoff"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "X-Ray Line"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Element"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   735
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   $"Effective.frx":0008
      Height          =   3015
      Left            =   9720
      TabIndex        =   32
      Top             =   1320
      Width           =   2295
   End
End
Attribute VB_Name = "FormEffective"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2024 by John J. Donovan
Option Explicit

Private Sub ComboElement_Change()
If Not DebugMode Then On Error Resume Next
Call EffectiveTakeoffAngleElementUpdate
If ierror Then Exit Sub
End Sub

Private Sub ComboElement_Click()
If Not DebugMode Then On Error Resume Next
Call EffectiveTakeoffAngleElementUpdate
If ierror Then Exit Sub
End Sub

Private Sub CommandCalculate_Click()
If Not DebugMode Then On Error Resume Next
Call EffectiveTakeoffAngleKRatiosSave
If ierror Then Exit Sub
Call EffectiveTakeoffAngleKRatiosCalculate
If ierror Then Exit Sub
End Sub

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Call EffectiveTakeoffAngleKRatiosSave
If ierror Then Exit Sub
Unload FormEffective
End Sub

Private Sub CommandFindNextNumber1_Click()
If Not DebugMode Then On Error Resume Next
Call StandardFindNumber(Int(1), FormEffective.TextStandardNumber1.Text, FormEffective.ListStandardPrimary)
If ierror Then Exit Sub
End Sub

Private Sub CommandFindNextNumber2_Click()
If Not DebugMode Then On Error Resume Next
Call StandardFindNumber(Int(1), FormEffective.TextStandardNumber2.Text, FormEffective.ListStandardSecondary)
If ierror Then Exit Sub
End Sub

Private Sub CommandFindNextString1_Click()
If Not DebugMode Then On Error Resume Next
Call StandardFindString(Int(1), FormEffective.TextStandardString1.Text, FormEffective.ListStandardPrimary)
If ierror Then Exit Sub
End Sub

Private Sub CommandFindNextString2_Click()
If Not DebugMode Then On Error Resume Next
Call StandardFindString(Int(1), FormEffective.TextStandardString2.Text, FormEffective.ListStandardSecondary)
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormEffective)
HelpContextID = IOGetHelpContextID("FormEFFECTIVE")
Call EffectiveTakeoffAngleKRatiosLoad
If ierror Then Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub ListStandardPrimary_Click()
Call EffectiveTakeoffAngleLoadElements
If ierror Then Exit Sub
End Sub

Private Sub TextBeamEnergy_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextStandardNumber1_Change()
If Not DebugMode Then On Error Resume Next
Call StandardFindNumber(Int(0), FormEffective.TextStandardNumber1.Text, FormEffective.ListStandardPrimary)
If ierror Then Exit Sub
End Sub

Private Sub TextStandardNumber2_Change()
If Not DebugMode Then On Error Resume Next
Call StandardFindNumber(Int(0), FormEffective.TextStandardNumber2.Text, FormEffective.ListStandardSecondary)
If ierror Then Exit Sub
End Sub

Private Sub TextStandardString1_Change()
If Not DebugMode Then On Error Resume Next
Call StandardFindString(Int(0), FormEffective.TextStandardString1.Text, FormEffective.ListStandardPrimary)
If ierror Then Exit Sub
End Sub

Private Sub TextStandardString2_Change()
If Not DebugMode Then On Error Resume Next
Call StandardFindString(Int(0), FormEffective.TextStandardString2.Text, FormEffective.ListStandardSecondary)
If ierror Then Exit Sub
End Sub

Private Sub TextTakeoffHigh_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextTakeoffIncrement_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextTakeoffLow_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub
