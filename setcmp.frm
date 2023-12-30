VERSION 5.00
Begin VB.Form FormSETCMP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Element Properties"
   ClientHeight    =   3480
   ClientLeft      =   1410
   ClientTop       =   1770
   ClientWidth     =   7305
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
   ScaleHeight     =   3480
   ScaleWidth      =   7305
   Begin VB.Frame Frame2 
      Caption         =   "Enter Other Element Properties (not saved to standard database)"
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   5895
      Begin VB.ComboBox ComboCrystal 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Crystal (default)"
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
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton CommandClear 
      Caption         =   "Clear"
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
      Left            =   6120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter Element Properties and Weight Percent For"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5895
      Begin VB.TextBox TextAtomicWtsStd 
         Height          =   285
         Left            =   4440
         TabIndex        =   19
         ToolTipText     =   "Enter the atomic weight for this element for enriched isotope standards (default is natural abundance)"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox TextAtomicChargesStd 
         Height          =   285
         Left            =   3000
         TabIndex        =   17
         ToolTipText     =   "Enter the ionic charge for charge balance calculations"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox TextComposition 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   2775
      End
      Begin VB.ComboBox ComboOxygens 
         Height          =   315
         Left            =   4440
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox ComboCations 
         Height          =   315
         Left            =   3000
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox ComboXRay 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox ComboElement 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Atomic Weight"
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
         Left            =   4440
         TabIndex        =   20
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Charge"
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
         Left            =   3000
         TabIndex        =   18
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label LabelComposition 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Enter Composition in "
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Oxygens"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Cations"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "X-Ray (default)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Element"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FormSETCMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2024 by John J. Donovan
Option Explicit

Private Sub ComboElement_Change()
If Not DebugMode Then On Error Resume Next
' Update
Call GetCmpSetCmpUpdateCombo
If ierror Then Exit Sub
End Sub

Private Sub ComboElement_Click()
If Not DebugMode Then On Error Resume Next
Call GetCmpSetCmpUpdateCombo
If ierror Then Exit Sub
End Sub

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormSETCMP
End Sub

Private Sub CommandClear_Click()
If Not DebugMode Then On Error Resume Next
Call GetCmpSetCmpClear
If ierror Then Exit Sub
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call GetCmpSetCmpSave
If ierror Then Exit Sub
Unload FormSETCMP
Call GetCmpSave
If ierror Then Exit Sub
Call GetCmpLoadGrid
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormSETCMP)
HelpContextID = IOGetHelpContextID("FormSETCMP")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub TextAtomicChargesStd_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextAtomicWtsStd_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextComposition_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

