VERSION 5.00
Begin VB.Form FormXRAY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "X-Ray Database"
   ClientHeight    =   8400
   ClientLeft      =   1710
   ClientTop       =   1530
   ClientWidth     =   8265
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
   ScaleHeight     =   8400
   ScaleWidth      =   8265
   Begin VB.ComboBox ComboMaximumOrder 
      Height          =   315
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Specify Range"
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   6360
      TabIndex        =   24
      Top             =   2880
      Width           =   1815
      Begin VB.ComboBox ComboOrder 
         Height          =   315
         Left            =   240
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Display a range of the x-ray  database centered on the specified element and x-ray"
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox ComboElm 
         Height          =   315
         Left            =   120
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Display a range of the x-ray  database centered on the specified element and x-ray"
         Top             =   360
         Width           =   735
      End
      Begin VB.ComboBox ComboXry 
         Height          =   315
         Left            =   960
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Display a range of the x-ray  database centered on the specified element and x-ray"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Bragg Order"
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
         Left            =   240
         TabIndex        =   31
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.CommandButton CommandGraphSelected 
      Caption         =   "Graph Selected"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Graph the selected x-ray line in the wavescan graph"
      Top             =   600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox CheckAbsorptionEdges 
      Caption         =   "Absorption Edges"
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
      Left            =   6360
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Toggle display display of x-ray absorption edges"
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Highlight Element"
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   6360
      TabIndex        =   19
      Top             =   1320
      Width           =   1815
      Begin VB.CommandButton CommandPeriodic 
         Caption         =   "Periodic Table"
         Height          =   495
         Left            =   240
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox ComboElement 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Select x-ray lines of the specified element"
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.TextBox TextMinimumKLM 
      Height          =   285
      Left            =   6480
      TabIndex        =   0
      ToolTipText     =   "Minimum x-ray intensity to display (default = 0.1)"
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "NIST X-Ray Lines (multi-select)"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   8175
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton CommandCopySelectedToClipboard 
         Caption         =   "Copy Selected to Clipboard"
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
         Left            =   3360
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   7680
         Width           =   2655
      End
      Begin VB.CommandButton CommandCopyToClipboard 
         Caption         =   "Copy to Clipboard"
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
         Left            =   360
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   7680
         Width           =   2655
      End
      Begin VB.ListBox ListXray 
         Height          =   6105
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   600
         Width           =   5895
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   $"Xray2.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   23
         Top             =   6720
         Width           =   5895
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Reference"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Intensity"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3360
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Energy"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Angstroms"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "X-Ray Line"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton CommandLoadNewRange 
      Caption         =   "Load New Range"
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Load the new x-ray database range based on the start and stop angstroms"
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox TextStart 
      Height          =   285
      Left            =   6480
      TabIndex        =   1
      ToolTipText     =   "Specify start of range for x-ray database display"
      Top             =   6840
      Width           =   1575
   End
   Begin VB.TextBox TextStop 
      Height          =   285
      Left            =   6480
      TabIndex        =   2
      ToolTipText     =   "Specify end of range for x-ray database display"
      Top             =   7440
      Width           =   1575
   End
   Begin VB.TextBox TextKev 
      Height          =   285
      Left            =   6480
      TabIndex        =   3
      ToolTipText     =   "Operating voltage for calculating higher Bragg order lines"
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H00C0FFC0&
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label LabelMaximumOrder 
      Alignment       =   2  'Center
      Caption         =   "Maximum Order"
      Height          =   255
      Left            =   6480
      TabIndex        =   28
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Minimum Intensity"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   13
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Start Angstroms"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   7
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Stop Angstroms"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   6
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "KeV"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   5
      Top             =   7800
      Width           =   1575
   End
End
Attribute VB_Name = "FormXRAY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2021 by John J. Donovan
Option Explicit

Private Sub ComboElement_Change()
If Not DebugMode Then On Error Resume Next
Call XrayHighlightElement
If ierror Then Exit Sub
End Sub

Private Sub ComboElement_Click()
If Not DebugMode Then On Error Resume Next
Call XrayHighlightElement
If ierror Then Exit Sub
End Sub

Private Sub ComboElm_Change()
If Not DebugMode Then On Error Resume Next
Call XraySpecifyRange(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub ComboElm_Click()
If Not DebugMode Then On Error Resume Next
Call XraySpecifyRange(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub ComboOrder_Change()
If Not DebugMode Then On Error Resume Next
Call XraySpecifyRange(Int(3))
If ierror Then Exit Sub
End Sub

Private Sub ComboOrder_Click()
If Not DebugMode Then On Error Resume Next
Call XraySpecifyRange(Int(3))
If ierror Then Exit Sub
End Sub

Private Sub ComboXry_Change()
If Not DebugMode Then On Error Resume Next
Call XraySpecifyRange(Int(2))
If ierror Then Exit Sub
End Sub

Private Sub ComboXry_Click()
If Not DebugMode Then On Error Resume Next
Call XraySpecifyRange(Int(2))
If ierror Then Exit Sub
End Sub

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormXRAY
End Sub

Private Sub CommandCopySelectedToClipboard_Click()
If Not DebugMode Then On Error Resume Next
Call MiscCopyList(Int(2), FormXRAY.ListXray)
If ierror Then Exit Sub
End Sub

Private Sub CommandCopyToClipboard_Click()
If Not DebugMode Then On Error Resume Next
Call MiscCopyList(Int(1), FormXRAY.ListXray)
If ierror Then Exit Sub
End Sub

Private Sub CommandGraphSelected_Click()
If Not DebugMode Then On Error Resume Next
' Nothing to do here in CalcZAF/Standard
End Sub

Private Sub CommandLoadNewRange_Click()
If Not DebugMode Then On Error Resume Next
Call XrayLoadNewRange
If ierror Then Exit Sub
End Sub

Private Sub CommandPeriodic_Click()
If Not DebugMode Then On Error Resume Next
Call XrayGetKLMElements0
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormXRAY)
HelpContextID = IOGetHelpContextID("FormXRAY")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub TextKev_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMinimumKLM_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextStart_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextStop_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

