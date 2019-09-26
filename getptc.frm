VERSION 5.00
Begin VB.Form FormGETPTC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Particle and Thin Film (single-layer) Calculations"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CommandCancel 
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
      Left            =   10800
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H00C0FFC0&
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
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Particle/Thin Film Options"
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
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin VB.CheckBox CheckPTCDoNotNormalizeSpecifiedFlag 
         Caption         =   "Do Not Normalize Specified (Fixed) Element Concentrations For Normalized Output"
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
         Left            =   240
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "To force the application to normalize specified elements (fixed concentrations) along with the other element concentration"
         Top             =   3120
         Width           =   7935
      End
      Begin VB.CommandButton CommandHelpOnThinFilms 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Help On Thin Films"
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
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton CommandHelpOnParticles 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Help On Particles"
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
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   480
         Width           =   2055
      End
      Begin VB.CheckBox CheckUsePTC 
         Caption         =   "Use Particle or Thin Film Calculations With Selected Phi-Rho-Z Corrections"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Use particle or thin film corrections for the selected samples"
         Top             =   360
         Width           =   8535
      End
      Begin VB.ComboBox ComboPTCModel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Select the particle or thin film model"
         Top             =   840
         Width           =   7575
      End
      Begin VB.TextBox TextPTCNumericalIntegrationStep 
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
         Left            =   2640
         TabIndex        =   10
         ToolTipText     =   "Enter the numerical integration step size (use smaller values for smaller particles or thinner films)"
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox TextPTCThicknessFactor 
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
         Left            =   2640
         TabIndex        =   9
         ToolTipText     =   "Enter the particle thickness to diameter or length ratio"
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox TextPTCDensity 
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
         Left            =   2640
         TabIndex        =   8
         ToolTipText     =   "Enter the particle density"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox TextPTCDiameter 
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
         Left            =   2640
         TabIndex        =   7
         ToolTipText     =   "Enter the particle diameter or thin film thickness in microns"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "0.00001 g/cm^2 (15 keV, 10 um)   (Use smaller value for higher keV, smaller size, thinner)"
         Height          =   255
         Left            =   3840
         TabIndex        =   15
         ToolTipText     =   $"GETPTC.frx":0000
         Top             =   2760
         Width           =   6495
      End
      Begin VB.Label Label9 
         Caption         =   "Sample thickness to diameter ratio = 1 (typical for equant particles or thin films)"
         Height          =   255
         Left            =   3840
         TabIndex        =   14
         Top             =   2280
         Width           =   5535
      End
      Begin VB.Label Label8 
         Caption         =   "3 g/cm^3  is typical for silicate minerals (irrelevant for thick specimens)"
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         ToolTipText     =   "For thin film or particle calculations an accurate density is required"
         Top             =   1800
         Width           =   5655
      End
      Begin VB.Label Label7 
         Caption         =   "Numerical Integration Step"
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
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Use 10000 (mg/cm^2) for thick specimens or enter thickness in um for thin films or particles"
         Height          =   255
         Left            =   3840
         TabIndex        =   11
         ToolTipText     =   $"GETPTC.frx":009D
         Top             =   1320
         Width           =   6495
      End
      Begin VB.Label Label5 
         Caption         =   "Particle Thickness Factor"
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
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Particle/Film Density"
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
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Particle Diameter/Thickness"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Particle/Thin Film Model"
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
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   2415
      End
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Use overscanned or defocussed beam for particles."
      Height          =   615
      Left            =   10800
      TabIndex        =   20
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Only unsupported thin films (TEM) or low-Z substrates are allowed."
      Height          =   855
      Left            =   10800
      TabIndex        =   19
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Thanks to John Armstrong"
      Height          =   495
      Left            =   11160
      TabIndex        =   18
      Top             =   3120
      Width           =   1095
   End
End
Attribute VB_Name = "FormGETPTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2019 by John J. Donovan
Option Explicit

Dim TextChanged As Boolean

Private Sub ComboPTCModel_Change()
If Not DebugMode Then On Error Resume Next
TextChanged = True
End Sub

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormGETPTC
icancelload = True
End Sub

Private Sub CommandHelpOnParticles_Click()
If Not DebugMode Then On Error Resume Next
Call IOBrowseHTTP(ProbeSoftwareInternetBrowseMethod%, "https://probesoftware.com/smf/index.php?topic=281.0")
If ierror Then Exit Sub
End Sub

Private Sub CommandHelpOnThinFilms_Click()
If Not DebugMode Then On Error Resume Next
Call IOBrowseHTTP(ProbeSoftwareInternetBrowseMethod%, "https://probesoftware.com/smf/index.php?topic=111.0")
If ierror Then Exit Sub
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
If TextChanged Then FormGETPTC.CheckUsePTC.value = vbChecked
Call GetPTCSave
If ierror Then Exit Sub
Unload FormGETPTC
End Sub

Private Sub Form_Activate()
If Not DebugMode Then On Error Resume Next
TextChanged = False
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
icancelload = False
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormGETPTC)
HelpContextID = IOGetHelpContextID("FormGETPTC")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub TextPTCDensity_Change()
If Not DebugMode Then On Error Resume Next
TextChanged = True
End Sub

Private Sub TextPTCDensity_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextPTCDiameter_Change()
If Not DebugMode Then On Error Resume Next
TextChanged = True
End Sub

Private Sub TextPTCDiameter_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextPTCNumericalIntegrationStep_Change()
If Not DebugMode Then On Error Resume Next
TextChanged = True
End Sub

Private Sub TextPTCNumericalIntegrationStep_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextPTCThicknessFactor_Change()
If Not DebugMode Then On Error Resume Next
TextChanged = True
End Sub

Private Sub TextPTCThicknessFactor_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub
