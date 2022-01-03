VERSION 5.00
Begin VB.Form FormZAFELM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Element Parameters"
   ClientHeight    =   4320
   ClientLeft      =   825
   ClientTop       =   4905
   ClientWidth     =   6480
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
   ScaleHeight     =   4320
   ScaleWidth      =   6480
   Begin VB.Frame Frame2 
      Caption         =   "Standard Parameters"
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   120
      TabIndex        =   18
      Top             =   2640
      Width           =   6255
      Begin VB.CommandButton CommandAddStandardsToRun 
         BackColor       =   &H0080FFFF&
         Caption         =   "Add/Remove Standards To/From Run"
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
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1080
         Width           =   3855
      End
      Begin VB.TextBox TextIntensityStd 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3720
         TabIndex        =   22
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox ComboStandard 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Intensity"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3720
         TabIndex        =   21
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Assigned Standard"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.CommandButton CommandDelete 
      Caption         =   "Delete"
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
      Left            =   5280
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Unknown Parameters"
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5055
      Begin VB.OptionButton OptionSpecified 
         Caption         =   "Specified"
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Select this option for specified elemental data (by difference, stoichiometry or fixed concentrations)"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton OptionAnalyzed 
         Caption         =   "Analyzed"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Select this option for measured elemental data"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox TextIntensity 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2640
         TabIndex        =   5
         Top             =   1800
         Width           =   2295
      End
      Begin VB.ComboBox ComboElement 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox ComboXRay 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox ComboCations 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2640
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox ComboOxygens 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3840
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox TextWeight 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label LabelIntensity 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Intensity"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   14
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Element"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "X-Ray Line"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Cations"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Oxygens"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label LabelWeight 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Weight Percent"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   2295
      End
   End
End
Attribute VB_Name = "FormZAFELM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2022 by John J. Donovan
Option Explicit

Private Sub ComboElement_Change()
' Update default x-ray and cations
If Not DebugMode Then On Error Resume Next
Call CalcZAFElementUpdate
If ierror Then Exit Sub
End Sub

Private Sub ComboElement_Click()
' Update default x-ray and cations
If Not DebugMode Then On Error Resume Next
Call CalcZAFElementUpdate
If ierror Then Exit Sub
End Sub

Private Sub ComboXray_Change()
If Not DebugMode Then On Error Resume Next
If FormZAFELM.ComboXRay.Text <> vbNullString Then
FormZAFELM.OptionAnalyzed.Value = True
Else
FormZAFELM.OptionSpecified.Value = True
End If
End Sub

Private Sub ComboXray_Click()
If Not DebugMode Then On Error Resume Next
If FormZAFELM.ComboXRay.Text <> vbNullString Then
FormZAFELM.OptionAnalyzed.Value = True
Else
FormZAFELM.OptionSpecified.Value = True
End If
End Sub

Private Sub CommandAddStandardsToRun_Click()
' Add standards to the current run
If Not DebugMode Then On Error Resume Next
Call AddStdLoad
If ierror Then Exit Sub
FormADDSTD.Show vbModal
If ierror Then Exit Sub
' Re-load standards for FormZAFELM
Call CalcZAFElementLoad2
If ierror Then Exit Sub
End Sub

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormZAFELM
End Sub

Private Sub CommandDelete_Click()
' Delete element fields
If Not DebugMode Then On Error Resume Next
Call CalcZAFElementDelete
If ierror Then Exit Sub
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFElementSave
If ierror Then Exit Sub
Unload FormZAFELM
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormZAFELM)
HelpContextID = IOGetHelpContextID("FormZAFELM")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub OptionAnalyzed_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFElementUpdate
If ierror Then Exit Sub
If FormZAF.OptionCalculate(0).Value = False Then
FormZAFELM.TextWeight.Enabled = False
FormZAFELM.TextIntensity.Enabled = True
FormZAFELM.ComboStandard.Enabled = True
FormZAFELM.TextIntensityStd.Enabled = True
End If
End Sub

Private Sub OptionSpecified_Click()
If Not DebugMode Then On Error Resume Next
If FormZAF.OptionCalculate(0).Value = False Then
FormZAFELM.ComboXRay.Text = vbNullString
FormZAFELM.TextWeight.Enabled = True
FormZAFELM.TextIntensity.Enabled = False
FormZAFELM.ComboStandard.Enabled = False
FormZAFELM.TextIntensityStd.Enabled = False
End If
End Sub

Private Sub TextIntensity_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextIntensityStd_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextWeight_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

