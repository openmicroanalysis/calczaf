VERSION 5.00
Begin VB.Form FormZAFOPT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculation Options"
   ClientHeight    =   6345
   ClientLeft      =   1440
   ClientTop       =   3480
   ClientWidth     =   9120
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
   ScaleHeight     =   6345
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TextDensity 
      Height          =   285
      Left            =   8040
      TabIndex        =   35
      Top             =   1800
      Width           =   975
   End
   Begin VB.Frame Frame6 
      Caption         =   "Sample Conductive Coating (need to explicitly turn on in Analytical menu)"
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   120
      TabIndex        =   26
      Top             =   5160
      Width           =   7815
      Begin VB.TextBox TextCoatingThickness 
         Height          =   285
         Left            =   2640
         TabIndex        =   30
         ToolTipText     =   "Enter the thickness of the elemental coating (in angstroms)"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TextCoatingDensity 
         Height          =   285
         Left            =   1440
         TabIndex        =   29
         ToolTipText     =   "Enter the density of the elemental coating (in gm/cm3)"
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox ComboCoatingElement 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   240
         TabIndex        =   28
         ToolTipText     =   "Select the element coating material for the sample(s)"
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox CheckCoatingFlag 
         Caption         =   "Use Conductive Coating"
         Height          =   255
         Left            =   4320
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Uncheck this box for no conductive coating on the selected sample(s)"
         Top             =   600
         Width           =   2655
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Formula Options"
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      TabIndex        =   20
      Top             =   3960
      Width           =   7815
      Begin VB.ComboBox ComboFormula 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Specify the formula element basis (e.g., for Mg2SiO4 use 2 Mg or 1 Si or 4 oxygen)"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox TextFormula 
         Height          =   285
         Left            =   3720
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Number of atoms for the formula basis"
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox CheckFormula 
         Caption         =   "Calculate Formula Based On"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Perform a formula atom calculation"
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Atoms Of"
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
         Left            =   4440
         TabIndex        =   24
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8040
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Calculation Options"
      ForeColor       =   &H00FF0000&
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.CheckBox CheckHydrogenStoichiometry 
         Caption         =   "Hydrogen Stoichiometry To Excess Oxygen"
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
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   $"Zafopt.frx":0000
         Top             =   3120
         Width           =   4095
      End
      Begin VB.TextBox TextHydrogenStoichiometry 
         Height          =   285
         Left            =   5400
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Ratio of hydrogen to oxygen atoms (1 = OH and 2 = H2O)"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox TextDifferenceFormula 
         Height          =   285
         Left            =   3480
         TabIndex        =   37
         ToolTipText     =   "Enter the formula by difference (not saved for export/import)"
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CheckBox CheckDifferenceFormula 
         Caption         =   "Formula By Difference (e.g. Li2B4O7):"
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
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "Specify a formula by difference in the sample analysis"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.CheckBox CheckAtomicPercents 
         Caption         =   "Calculate Atomic Percents"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Also calculate the atomic percent composition"
         Top             =   600
         Width           =   3855
      End
      Begin VB.CheckBox CheckDisplayAsOxide 
         Caption         =   "Display Results As Oxide Formulas"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "Display the results in oxides formulas"
         Top             =   360
         Width           =   3855
      End
      Begin VB.CheckBox CheckUseOxygenFromHalogensCorrection 
         Caption         =   "Use Oxygen From Halogens Correction"
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
         TabIndex        =   19
         ToolTipText     =   "Subtract oxygen equivalent of halogens from oxygen"
         Top             =   960
         Width           =   3855
      End
      Begin VB.CheckBox CheckCalculateElectronandXrayRanges 
         Caption         =   "Calculate Electron and Xray Ranges"
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
         TabIndex        =   18
         ToolTipText     =   "Calculate the electron and x-ray ranges for the unknown composition"
         Top             =   1200
         Width           =   3855
      End
      Begin VB.CheckBox CheckDifference 
         Caption         =   "Element By Difference"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Calculate an element by difference from 100%"
         Top             =   1560
         Width           =   3255
      End
      Begin VB.CheckBox CheckStoichiometry 
         Caption         =   "Stoichiometry To Oxygen"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2280
         Width           =   3255
      End
      Begin VB.CheckBox CheckRelative 
         Caption         =   "Stoichiometry To Another Element"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Calculate an element by stoichiometry to another element"
         Top             =   2640
         Width           =   3375
      End
      Begin VB.ComboBox ComboStoichiometry 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2280
         Width           =   735
      End
      Begin VB.ComboBox ComboRelativeTo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox TextStoichiometry 
         Height          =   285
         Left            =   3480
         TabIndex        =   13
         Top             =   2280
         Width           =   855
      End
      Begin VB.ComboBox ComboDifference 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.ComboBox ComboRelative 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox TextRelative 
         Height          =   285
         Left            =   3480
         TabIndex        =   10
         Top             =   2640
         Width           =   855
      End
      Begin VB.OptionButton OptionElemental 
         Caption         =   "Calculate as Elemental"
         Height          =   255
         Left            =   4080
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Calculate the composition as elemental (with stoichiometric oxygen)"
         Top             =   600
         Width           =   3255
      End
      Begin VB.OptionButton OptionOxide 
         Caption         =   "Calculate with Stoichiometric Oxygen"
         Height          =   255
         Left            =   4080
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Calculate the composition with oxygen by stoichiometry added to the matrix correction"
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label9 
         Caption         =   "(OH = 1, H2O = 2)"
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
         TabIndex        =   41
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Oxygen"
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
         Left            =   6840
         TabIndex        =   17
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Atoms Of"
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
         Left            =   4320
         TabIndex        =   3
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Atoms Of"
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
         Left            =   4320
         TabIndex        =   4
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "To"
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
         Left            =   6240
         TabIndex        =   5
         Top             =   2400
         Width           =   495
      End
   End
   Begin VB.Label Label6 
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
      Left            =   8040
      TabIndex        =   34
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "FormZAFOPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2018 by John J. Donovan
Option Explicit

Private Sub CheckHydrogenStoichiometry_Click()
If Not DebugMode Then On Error Resume Next
If FormZAFOPT.CheckHydrogenStoichiometry.value = vbChecked Then
Call ZAFOptionCheckForExcessOxygen
End If
End Sub

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormZAFOPT
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call ZAFOptionSave
If ierror Then Exit Sub
Unload FormZAFOPT
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormZAFOPT)
HelpContextID = IOGetHelpContextID("FormZAFOPT")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub OptionOxide_Click()
If Not DebugMode Then On Error Resume Next
Call ZAFOptionOxygen
If ierror Then Exit Sub
End Sub

Private Sub TextCoatingDensity_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextCoatingThickness_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextDifferenceFormula_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextFormula_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextHydrogenStoichiometry_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextRelative_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextStoichiometry_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub
