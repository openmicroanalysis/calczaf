VERSION 5.00
Begin VB.Form FormSTANDARDCOATING 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Standard Parameters"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Edit Individual Standard Conductive Coating Parameters"
      ForeColor       =   &H00FF0000&
      Height          =   4575
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   6495
      Begin VB.CheckBox CheckCoatingFlag 
         Caption         =   "Use Conductive Coating"
         Height          =   735
         Left            =   4680
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Uncheck this box for no conductive coating for the selected standard"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton CommandApply 
         Caption         =   "Apply Parameters To Selected Standard"
         Height          =   975
         Left            =   4800
         TabIndex        =   19
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox TextCoatingThickness 
         Height          =   285
         Left            =   4800
         TabIndex        =   17
         ToolTipText     =   "Enter the thickness of the elemental coating (in angstroms) for the selected standard"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox TextCoatingDensity 
         Height          =   285
         Left            =   4800
         TabIndex        =   15
         ToolTipText     =   "Enter the density of the elemental coating (in gm/cm3) for the selected standard"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox ComboCoatingElement 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4800
         TabIndex        =   13
         ToolTipText     =   "Select the element coating material for the selected standard"
         Top             =   600
         Width           =   1335
      End
      Begin VB.ListBox ListStandard 
         Height          =   3960
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Thickness (A)"
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
         Left            =   4800
         TabIndex        =   18
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Density"
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
         Left            =   4800
         TabIndex        =   16
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Element"
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
         Left            =   4800
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6720
      TabIndex        =   9
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame6 
      Caption         =   "Conductive Coating Parameters For All Standards"
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   6495
      Begin VB.CommandButton CommandAssignToAll 
         BackColor       =   &H0080FFFF&
         Caption         =   "Assign To All Standards"
         Height          =   735
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox TextCoatingThicknessAll 
         Height          =   285
         Left            =   2640
         TabIndex        =   4
         ToolTipText     =   "Enter the thickness of the elemental coating (in angstroms) for all standards"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TextCoatingDensityAll 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         ToolTipText     =   "Enter the density of the elemental coating (in gm/cm3) for all standards"
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox ComboCoatingElementAll 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Select the element coating material for all standards"
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox CheckCoatingFlagAll 
         Caption         =   "Use Conductive Coating On All Standards"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Uncheck this box for no conductive coating for all standards"
         Top             =   1080
         Width           =   4215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Element"
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
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Thickness (A)"
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
         Left            =   2640
         TabIndex        =   6
         ToolTipText     =   "Enter coating thickness in angstroms"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Density"
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
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label LabelTurnOn 
      Alignment       =   2  'Center
      Caption         =   "To enable coating corrections, please explicitly turn on coating options in Analytical | Analysis Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6720
      TabIndex        =   21
      Top             =   1680
      Width           =   1575
   End
End
Attribute VB_Name = "FormSTANDARDCOATING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2025 by John J. Donovan
Option Explicit

Private Sub CommandApply_Click()
If Not DebugMode Then On Error Resume Next
Call StandardCoatingApply
If ierror Then Exit Sub
End Sub

Private Sub CommandAssignToAll_Click()
If Not DebugMode Then On Error Resume Next
Call StandardCoatingAssignToAll
If ierror Then Exit Sub
End Sub

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormSTANDARDCOATING
icancelload = True
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
' Save changes
Call StandardCoatingSave
If ierror Then Exit Sub
Unload FormSTANDARDCOATING
icancelload = False
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormSTANDARDCOATING)
HelpContextID = IOGetHelpContextID("FormSTANDARDCOATING")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub ListStandard_Click()
If Not DebugMode Then On Error Resume Next
Call StandardCoatingSelect
If ierror Then Exit Sub
End Sub

Private Sub TextCoatingDensity_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextCoatingDensityAll_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextCoatingThickness_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextCoatingThicknessAll_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub
