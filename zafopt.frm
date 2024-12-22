VERSION 5.00
Begin VB.Form FormZAFOPT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculation Options"
   ClientHeight    =   6030
   ClientLeft      =   1440
   ClientTop       =   3480
   ClientWidth     =   13155
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
   ScaleHeight     =   6030
   ScaleWidth      =   13155
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CommandDynamicElements 
      Caption         =   "Calculate UnAnalyzed Elements Dynamically"
      Height          =   975
      Left            =   9960
      TabIndex        =   63
      ToolTipText     =   "Calculate elements by difference and/or by stocihiometry dynamically based on pixel k-ratio values"
      Top             =   840
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "EDS Spectral Data And Quant Calculations"
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   8760
      TabIndex        =   40
      Top             =   2040
      Width           =   4215
      Begin VB.CommandButton CommandHelpOnEDSWDS 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Help"
         Height          =   255
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Click this button to get detailed help from our on-line user forum"
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton CommandSelectQuantMethodOrProject 
         Caption         =   "EDS Quant Method"
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox CheckUseEDSSpectra 
         Caption         =   "Use EDS Spectra For Quant"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label LabelQuantMethodOrProject 
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   44
         Top             =   1200
         Width           =   3975
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Sample Conductive Coating"
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   8760
      TabIndex        =   26
      Top             =   3960
      Width           =   4215
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
         Left            =   840
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Uncheck this box for no conductive coating on the selected sample(s)"
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "See standard coating options under Analytical menu and global coating correction options in Analysis Options dialog"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   45
         Top             =   1200
         Width           =   3735
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
      Top             =   5040
      Width           =   8415
      Begin VB.ComboBox ComboFormula 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Specify the formula element basis (e.g., for Mg2SiO4 use 2 Mg or 1 Si or 4 oxygen)"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox TextFormula 
         Height          =   285
         Left            =   3480
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Number of atoms for the formula basis"
         Top             =   360
         Width           =   855
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
         Left            =   4320
         TabIndex        =   24
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   11160
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Calculation Options"
      ForeColor       =   &H00FF0000&
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.CheckBox CheckUseOxygenFromSulfurCorrection 
         Caption         =   "Use Oxygen From Sulfur Correction"
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
         Left            =   3360
         TabIndex        =   62
         ToolTipText     =   "Subtract oxygen equivalent of sulfur from oxygen (must be negative charge valence)"
         Top             =   960
         Width           =   3255
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         TabIndex        =   53
         Top             =   3840
         Width           =   8175
         Begin VB.OptionButton OptionFerrousFerricOption 
            Caption         =   "Li Amphibole"
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
            Index           =   6
            Left            =   5760
            TabIndex        =   61
            ToolTipText     =   "Select Droop option. Use Calcic Amphibole for 13 cations exclusive of Ca, Na and K (suitable for many calcic amphiboles)."
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton OptionFerrousFerricOption 
            Caption         =   "Oxo Amphibole"
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
            Index           =   5
            Left            =   3960
            TabIndex        =   60
            ToolTipText     =   "Select Droop option. Use Calcic Amphibole for 13 cations exclusive of Ca, Na and K (suitable for many calcic amphiboles)."
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton OptionFerrousFerricOption 
            Caption         =   "Fe-Mg-Mn Amphibole"
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
            Index           =   4
            Left            =   1920
            TabIndex        =   59
            ToolTipText     =   "Select Droop option. Use Calcic Amphibole for 13 cations exclusive of Ca, Na and K (suitable for many calcic amphiboles)."
            Top             =   360
            Width           =   2055
         End
         Begin VB.OptionButton OptionFerrousFerricOption 
            Caption         =   "Na-Ca Amphibole"
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
            Index           =   3
            Left            =   240
            TabIndex        =   58
            ToolTipText     =   "Select Droop option. Use Calcic Amphibole for 13 cations exclusive of Ca, Na and K (suitable for many calcic amphiboles)."
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton OptionFerrousFerricOption 
            Caption         =   "Mineral"
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
            Index           =   0
            Left            =   240
            TabIndex        =   56
            ToolTipText     =   "Select Droop option. Use Mineral for simple minerals, e.g. Fe-Ti oxides, etc."
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptionFerrousFerricOption 
            Caption         =   "Sodic Amphibole"
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
            Left            =   1200
            TabIndex        =   55
            ToolTipText     =   $"ZAFOPT.frx":0000
            Top             =   120
            Width           =   1575
         End
         Begin VB.OptionButton OptionFerrousFerricOption 
            Caption         =   "Calcic Amphibole"
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
            Left            =   2880
            TabIndex        =   54
            ToolTipText     =   "Select Droop option. Use Calcic Amphibole for 13 cations exclusive of Ca, Na and K (suitable for many calcic amphiboles)."
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Caption         =   "Contributed by Joy, Locock and Moy"
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
            Left            =   4680
            TabIndex        =   57
            Top             =   0
            Width           =   3615
         End
      End
      Begin VB.CheckBox CheckFerrousFerricCalculation 
         Caption         =   "Calculate Excess Oxygen From Ferrous/Ferric Ratio"
         Height          =   315
         Left            =   120
         TabIndex        =   50
         ToolTipText     =   "Calculate the excess oxygen from ferric iron based on total charge balance (Droop, 1987)"
         Top             =   3480
         Width           =   4815
      End
      Begin VB.TextBox TextFerrousFerricTotalCations 
         Height          =   285
         Left            =   5640
         TabIndex        =   49
         ToolTipText     =   "Enter total formula cations for this iron bearing mineral (e.g., ilmenite = 2)"
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox TextFerrousFerricTotalOxygens 
         Height          =   285
         Left            =   7200
         TabIndex        =   48
         ToolTipText     =   "Enter total formula oxygens for this iron bearing mineral (e.g., ilmenite = 3)"
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox TextDensity 
         Height          =   285
         Left            =   7080
         TabIndex        =   46
         Top             =   1200
         Width           =   975
      End
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
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   $"ZAFOPT.frx":0093
         Top             =   3120
         Width           =   4095
      End
      Begin VB.TextBox TextHydrogenStoichiometry 
         Height          =   285
         Left            =   5640
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Ratio of hydrogen to oxygen atoms (1 = OH and 2 = H2O)"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox TextDifferenceFormula 
         Height          =   285
         Left            =   3600
         TabIndex        =   35
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
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Specify a formula by difference in the sample analysis"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.CheckBox CheckAtomicPercents 
         Caption         =   "Calculate Atomic Percents"
         Height          =   255
         Left            =   120
         TabIndex        =   34
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
         Width           =   3375
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
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2280
         Width           =   735
      End
      Begin VB.ComboBox ComboRelativeTo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox TextStoichiometry 
         Height          =   285
         Left            =   3600
         TabIndex        =   13
         Top             =   2280
         Width           =   855
      End
      Begin VB.ComboBox ComboDifference 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1440
         Width           =   735
      End
      Begin VB.ComboBox ComboRelative 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox TextRelative 
         Height          =   285
         Left            =   3600
         TabIndex        =   10
         Top             =   2640
         Width           =   855
      End
      Begin VB.OptionButton OptionElemental 
         Caption         =   "Calculate as Elemental"
         Height          =   255
         Left            =   4320
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Calculate the composition as elemental (with stoichiometric oxygen)"
         Top             =   600
         Width           =   3255
      End
      Begin VB.OptionButton OptionOxide 
         Caption         =   "Calculate with Stoichiometric Oxygen"
         Height          =   255
         Left            =   4320
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Calculate the composition with oxygen by stoichiometry added to the matrix correction"
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label15 
         Caption         =   "Cations"
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
         Left            =   5040
         TabIndex        =   52
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Oxygens"
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
         Left            =   6480
         TabIndex        =   51
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Density"
         Height          =   255
         Left            =   7080
         TabIndex        =   47
         Top             =   960
         Width           =   975
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
         Left            =   6480
         TabIndex        =   39
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
         Left            =   7080
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
         Left            =   4440
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
         Left            =   4440
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
         Left            =   6480
         TabIndex        =   5
         Top             =   2400
         Width           =   495
      End
   End
End
Attribute VB_Name = "FormZAFOPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2025 by John J. Donovan
Option Explicit

Private Sub CheckFerrousFerricCalculation_Click()
If Not DebugMode Then On Error Resume Next
If FormZAFOPT.CheckFerrousFerricCalculation.Value = vbChecked Then
If FormZAFOPT.OptionFerrousFerricOption(0).Value = True Then
FormZAFOPT.TextFerrousFerricTotalCations.Enabled = True
FormZAFOPT.TextFerrousFerricTotalOxygens.Enabled = True
FormZAFOPT.TextFerrousFerricTotalOxygens.Enabled = True
FormZAFOPT.OptionFerrousFerricOption(0).Enabled = True
FormZAFOPT.OptionFerrousFerricOption(1).Enabled = True
FormZAFOPT.OptionFerrousFerricOption(2).Enabled = True
FormZAFOPT.OptionFerrousFerricOption(3).Enabled = True
FormZAFOPT.OptionFerrousFerricOption(4).Enabled = True
FormZAFOPT.OptionFerrousFerricOption(5).Enabled = True
FormZAFOPT.OptionFerrousFerricOption(6).Enabled = True
Else
FormZAFOPT.TextFerrousFerricTotalCations.Enabled = False
FormZAFOPT.TextFerrousFerricTotalOxygens.Enabled = False
FormZAFOPT.OptionFerrousFerricOption(0).Enabled = False
FormZAFOPT.OptionFerrousFerricOption(1).Enabled = False
FormZAFOPT.OptionFerrousFerricOption(2).Enabled = False
FormZAFOPT.OptionFerrousFerricOption(3).Enabled = False
FormZAFOPT.OptionFerrousFerricOption(4).Enabled = False
FormZAFOPT.OptionFerrousFerricOption(5).Enabled = False
FormZAFOPT.OptionFerrousFerricOption(6).Enabled = False
End If
Else
FormZAFOPT.TextFerrousFerricTotalCations.Enabled = False
FormZAFOPT.TextFerrousFerricTotalOxygens.Enabled = False
FormZAFOPT.OptionFerrousFerricOption(0).Enabled = False
FormZAFOPT.OptionFerrousFerricOption(1).Enabled = False
FormZAFOPT.OptionFerrousFerricOption(2).Enabled = False
FormZAFOPT.OptionFerrousFerricOption(3).Enabled = False
FormZAFOPT.OptionFerrousFerricOption(4).Enabled = False
FormZAFOPT.OptionFerrousFerricOption(5).Enabled = False
FormZAFOPT.OptionFerrousFerricOption(6).Enabled = False
End If
End Sub

Private Sub CheckHydrogenStoichiometry_Click()
If Not DebugMode Then On Error Resume Next
If FormZAFOPT.CheckHydrogenStoichiometry.Value = vbChecked Then
Call ZAFOptionCheckForExcessOxygen
End If
End Sub

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormZAFOPT
End Sub

Private Sub CommandDynamicElements_Click()
Call ZAFOptionLoadDynamicElements
If ierror Then Exit Sub
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

Private Sub OptionFerrousFerricOption_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next
If FormZAFOPT.OptionFerrousFerricOption(0).Value = True Then
FormZAFOPT.TextFerrousFerricTotalCations.Enabled = True
FormZAFOPT.TextFerrousFerricTotalOxygens.Enabled = True
Else
FormZAFOPT.TextFerrousFerricTotalCations.Enabled = False
FormZAFOPT.TextFerrousFerricTotalOxygens.Enabled = False
End If
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

Private Sub TextDensity_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextDifferenceFormula_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextFerrousFerricTotalCations_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextFerrousFerricTotalOxygens_GotFocus()
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
