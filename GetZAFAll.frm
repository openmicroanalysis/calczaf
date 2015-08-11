VERSION 5.00
Begin VB.Form FormGETZAFALL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Matrix Correction Methods"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CommandOptions 
      BackColor       =   &H0000FFFF&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Select the ZAF or Phi-Rho-Z matrix correction procedure"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton CommandMACs 
      BackColor       =   &H0080FFFF&
      Caption         =   "MACs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Select the mass absorption coefficient lookup table"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame6 
      Caption         =   "Correction Method"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Use a rigorous fundamental parameters analytical matrix correction method (under development)"
         Top             =   3360
         Width           =   5055
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Use a multi-standard calibration curve matrix correction similar to XRF (useful for trace carbon in steel for example)"
         Top             =   3000
         Width           =   5055
      End
      Begin VB.TextBox TextPenepmaKratioLimit 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   14
         ToolTipText     =   "Enter the maximum concentration of the emitting element to load k-ratios from Penepma (to minimize precision errors)"
         Top             =   2520
         Width           =   975
      End
      Begin VB.CheckBox CheckPenepmaKratioLimit 
         Caption         =   "Use Penepma K-Ratio Limits"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         ToolTipText     =   "Select this option to limit the use of Penepma K-ratios for concentrations of the emitting element less than the specified limit"
         Top             =   2520
         Width           =   3015
      End
      Begin VB.CheckBox CheckUsePenepmaKratios 
         Caption         =   "Use PENEPMA Alpha Factors"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   $"GetZAFAll.frx":0000
         Top             =   2280
         Width           =   2895
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Use ZAF or Phi-Rho-Z selections for matrix correction"
         Top             =   360
         Width           =   4335
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
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
         Index           =   1
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Use single term alpha correction factors (50:50 constant fit from Ogilvie, Albee and Ray)"
         Top             =   840
         Width           =   4335
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
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
         Index           =   2
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Use two term alpha correction factors (linear fit from Rivers)"
         Top             =   1080
         Width           =   4335
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Use three term alpha correction factors (polynomial fit from Armstrong)"
         Top             =   1320
         Width           =   4335
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Use four term alpha correction factors (non-linear fit from Donovan)"
         Top             =   1560
         Width           =   5055
      End
      Begin VB.CheckBox CheckEmpiricalAlphaFlag 
         Caption         =   "Use Empirical Alpha Factors"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Overload the calculated Bence-Albee correction factors using empirical correction factors from the EMPFAC.DAT file"
         Top             =   1920
         Width           =   2895
      End
      Begin VB.Label LabelUsePenepmaAlphaFactors 
         Alignment       =   2  'Center
         Caption         =   "Penepma alpha factors includes fluorescence by beta lines and also the continuum."
         Height          =   615
         Left            =   3360
         TabIndex        =   12
         Top             =   1920
         Width           =   2295
      End
   End
End
Attribute VB_Name = "FormGETZAFALL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit

Private Sub CheckPenepmaKratioLimit_Click()
If Not DebugMode Then On Error Resume Next
If FormGETZAFALL.CheckPenepmaKratioLimit.Value = vbChecked Then
FormGETZAFALL.TextPenepmaKratioLimit.Enabled = True
Else
FormGETZAFALL.TextPenepmaKratioLimit.Enabled = False
End If
End Sub

Private Sub CheckUsePenepmaKratios_Click()
If FormGETZAFALL.CheckUsePenepmaKratios.Value = vbChecked Then
FormGETZAFALL.CheckPenepmaKratioLimit.Enabled = True
Else
FormGETZAFALL.CheckPenepmaKratioLimit.Enabled = False
End If
End Sub

Private Sub Command1_Click()
' User clicked OK in form GETZAFALL
If Not DebugMode Then On Error Resume Next
' Save the current correction method
Call GetZAFAllSave
If ierror Then Exit Sub
Unload FormGETZAFALL
End Sub

Private Sub Command2_Click()
If Not DebugMode Then On Error Resume Next
Unload FormGETZAFALL
icancelload = True
End Sub

Private Sub CommandMACs_Click()
' Get MAC file selections
If Not DebugMode Then On Error Resume Next
Call GetZAFAllLoadMAC
If ierror Then Exit Sub
FormMAC.Show vbModal
End Sub

Private Sub CommandOptions_Click()
' Get options
If Not DebugMode Then On Error Resume Next
Call GetZAFAllOptions
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
icancelload = False
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormGETZAFALL)
HelpContextID = IOGetHelpContextID("FormGETZAFALL")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub Option6_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next

' Set enabled property on Empirical alpha factors
If Index% < 1 Or Index% > 4 Then
FormGETZAFALL.CheckEmpiricalAlphaFlag.Enabled = False
FormGETZAFALL.CheckUsePenepmaKratios.Enabled = False
FormGETZAFALL.CheckPenepmaKratioLimit.Enabled = False

FormGETZAFALL.TextPenepmaKratioLimit.Enabled = False

Else
FormGETZAFALL.CheckEmpiricalAlphaFlag.Enabled = True
FormGETZAFALL.CheckUsePenepmaKratios.Enabled = True
FormGETZAFALL.CheckPenepmaKratioLimit.Enabled = True

If FormGETZAFALL.CheckUsePenepmaKratios.Value = vbChecked Then
FormGETZAFALL.CheckPenepmaKratioLimit.Enabled = True
Else
FormGETZAFALL.CheckPenepmaKratioLimit.Enabled = False
End If

If FormGETZAFALL.CheckPenepmaKratioLimit.Value = vbChecked Then
FormGETZAFALL.TextPenepmaKratioLimit.Enabled = True
End If
End If

' Set enabled property for CommandOptions
If Index% <= 4 Then
FormGETZAFALL.CommandOptions.Enabled = True
FormGETZAFALL.CommandMACs.Enabled = True
Else
FormGETZAFALL.CommandOptions.Enabled = False
FormGETZAFALL.CommandMACs.Enabled = False
End If

End Sub

Private Sub TextPenepmaKratioLimit_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub