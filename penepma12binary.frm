VERSION 5.00
Begin VB.Form FormPenepma12Binary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculate and Extract Binary Compositions"
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptionMinimumOvervoltagePercent 
      Caption         =   "Use 2% Minimum Overvoltage For Binary K-Ratio Fanal Extractions"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   8760
      Width           =   5415
   End
   Begin VB.OptionButton OptionMinimumOvervoltagePercent 
      Caption         =   "Use 40% Minimum Overvoltage For Binary K-Ratio Fanal Extractions"
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   9480
      Width           =   5415
   End
   Begin VB.OptionButton OptionMinimumOvervoltagePercent 
      Caption         =   "Use 20% Minimum Overvoltage For Binary K-Ratio Fanal Extractions"
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   9240
      Width           =   5415
   End
   Begin VB.OptionButton OptionMinimumOvervoltagePercent 
      Caption         =   "Use 10% Minimum Overvoltage For Binary K-Ratio Fanal Extractions"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   9000
      Width           =   5415
   End
   Begin VB.CommandButton CommandCalculateKratios 
      BackColor       =   &H0080FFFF&
      Caption         =   "Calculate and Compare to Experimental Kratios"
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   62
      TabStop         =   0   'False
      ToolTipText     =   "Read binary k-ratio input .DAT file and calculate all necessary PAR files for all compositions (e.g., Pouchou2.dat)"
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Calculate Binary Density"
      ForeColor       =   &H00FF0000&
      Height          =   2415
      Left            =   5160
      TabIndex        =   50
      Top             =   4680
      Width           =   2055
      Begin VB.CommandButton CommandCalculateDensity 
         Caption         =   "Calculate Density"
         Height          =   375
         Left            =   240
         TabIndex        =   60
         TabStop         =   0   'False
         ToolTipText     =   "Calculate the pure element normalized density based on the above binary composition"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox TextDensityElementA 
         Height          =   285
         Left            =   120
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "Enter symbol for element A to calculate pure element normalized density"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TextDensityElementB 
         Height          =   285
         Left            =   1080
         TabIndex        =   56
         TabStop         =   0   'False
         ToolTipText     =   "Enter symbol for element B to calculate pure element normalized density"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TextConcAA 
         Height          =   285
         Left            =   120
         TabIndex        =   53
         TabStop         =   0   'False
         ToolTipText     =   "Enter elemental weight percent for element A to calculate pure element normalized density"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox TextConcBB 
         Height          =   285
         Left            =   1080
         TabIndex        =   52
         TabStop         =   0   'False
         ToolTipText     =   "Enter elemental weight percent for element B to calculate pure element normalized density"
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "SymA"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "SymB"
         Height          =   255
         Left            =   1080
         TabIndex        =   58
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "ConcA"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "ConcB"
         Height          =   255
         Left            =   1080
         TabIndex        =   54
         Top             =   960
         Width           =   855
      End
      Begin VB.Label LabelDensity 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   480
         TabIndex        =   51
         Top             =   2040
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Calculate Alpha Factors"
      ForeColor       =   &H00FF0000&
      Height          =   2775
      Left            =   5160
      TabIndex        =   38
      Top             =   1440
      Width           =   2055
      Begin VB.TextBox TextKratB 
         Height          =   285
         Left            =   1080
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   "Enter elemental kratio (relative to pure element intensity) for element B"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox TextKratA 
         Height          =   285
         Left            =   120
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   "Enter elemental kratio (relative to pure element intensity) for element A"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox TextConcB 
         Height          =   285
         Left            =   1080
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   "Enter elemental concentration for element B"
         Top             =   1200
         Width           =   855
      End
      Begin VB.OptionButton OptionEnter 
         Caption         =   "Enter Percentage"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "Enter elemental concentrations (and k-ratios) in weight fraction (and intensity fraction)"
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton OptionEnter 
         Caption         =   "Enter Fractional"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Enter elemental concentrations (and k-ratios) in weight percent (and intensity percent)"
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox TextConcA 
         Height          =   285
         Left            =   120
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Enter elemental concentration for element A"
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton CommandCalculateAlphas 
         Caption         =   "Calculate Alpha Factors"
         Height          =   495
         Left            =   360
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Calculate alpha factors based on composition and k-ratios"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "KratB"
         Height          =   255
         Left            =   1080
         TabIndex        =   48
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "KratA"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "ConcB"
         Height          =   255
         Left            =   1080
         TabIndex        =   44
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "ConcA"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.CommandButton CommandTestLoadPenepmaAtomicWeights 
      Caption         =   "Test Load Penepma Atomic Weights"
      Height          =   495
      Left            =   5160
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Read Penepma atomic weights from Pendbase\pdfiles\pdcompos.p08 (for self consistency in calculations)"
      Top             =   7920
      Width           =   2055
   End
   Begin VB.CommandButton CommandPenPFE 
      BackColor       =   &H0080FFFF&
      Caption         =   "Open PenPFE Control Dialog"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "For creating multiple shared instances of PenPFE"
      Top             =   8040
      Width           =   4335
   End
   Begin VB.Frame Frame5 
      Caption         =   "Extract Binary Boundary or Matrix K-Ratios"
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
      TabIndex        =   16
      Top             =   4200
      Width           =   4935
      Begin VB.CommandButton CommandExtractRandom 
         Caption         =   "Extract Random K-Ratios Using TXT Share Folder"
         Height          =   375
         Left            =   240
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3240
         Width           =   4335
      End
      Begin VB.CommandButton CommandCheckPenfluorInputFiles 
         Caption         =   "Check Penfluor Input Files"
         Height          =   375
         Left            =   240
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   2880
         Width           =   2175
      End
      Begin VB.CommandButton CommandExtract2 
         Caption         =   "Extract Binary K-Ratios From Specified Composition"
         Height          =   495
         Left            =   1920
         TabIndex        =   68
         TabStop         =   0   'False
         ToolTipText     =   "Extract all k-ratios from PAR files based on the specified formula"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.CheckBox CheckDoNotOverwriteTXT 
         Caption         =   "Do Not Overwrite Existing .TXT Files"
         Height          =   255
         Left            =   1800
         TabIndex        =   61
         ToolTipText     =   "Skip extractions for existing .TXT files"
         Top             =   720
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.ComboBox ComboExtractMatrixA1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Select the first element in the boundary binary"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox ComboExtractMatrixA2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Select the second element in the boundary binary"
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton CommandOutputPlotData 
         Caption         =   "Output Plot Data"
         Height          =   495
         Left            =   240
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Output k-ratio data to plot file for emitting element, matrix and beam energy"
         Top             =   2400
         Width           =   2175
      End
      Begin VB.CommandButton CommandExtract 
         BackColor       =   &H0080FFFF&
         Caption         =   "Extract Binary K-Ratios"
         Height          =   375
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Extract binary boundary or matrix fluorescence component k-ratio files for the measured element"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.OptionButton OptionExtractMethod 
         Caption         =   "Matrix Only"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Extract matrix (self) fluorescence without boundary effect"
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptionExtractMethod 
         Caption         =   "Boundary"
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Extract boundary fluorescence effects for the measured element and the matrix element"
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox ComboExtractElement 
         Height          =   315
         Left            =   240
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Select the emitting element to extract boundary or matrix fluorescence"
         Top             =   600
         Width           =   615
      End
      Begin VB.ComboBox ComboExtractMatrix 
         Height          =   315
         Left            =   960
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Select the  matrix element of the emitting element in the beam incident material"
         Top             =   600
         Width           =   615
      End
      Begin VB.CheckBox CheckExtractForSpecifiedRange 
         Caption         =   "Extract K-Ratios For Matrix Range"
         Height          =   315
         Left            =   1800
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Extract k-ratios for all binaries between the specified Emitter and Matrix elements"
         Top             =   960
         Width           =   3015
      End
      Begin VB.ComboBox ComboExtractMatrixB2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Select the second element in the boundary binary"
         Top             =   1920
         Width           =   615
      End
      Begin VB.ComboBox ComboExtractMatrixB1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Select the first element in the boundary binary"
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton CommandExtractBinary 
         Caption         =   "Extract From Input File"
         Height          =   375
         Left            =   2400
         TabIndex        =   72
         TabStop         =   0   'False
         ToolTipText     =   "Extract binary (alpha) element intensities using Fanal from an input file, e.g., Pouchou2.dat"
         Top             =   2880
         Width           =   2175
      End
      Begin VB.CommandButton CommandExtractPureElementIntensities 
         Caption         =   "Extract Pure Element Intensities"
         Height          =   495
         Left            =   2400
         TabIndex        =   69
         TabStop         =   0   'False
         ToolTipText     =   "Extract pure element intensities using Fanal (always uses element range mode)"
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label LabelExtractMatrixA1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Matrix1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label LabelExtractMatrixA2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Matrix2"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   36
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label LabelExtractElement 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Emitter"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   615
      End
      Begin VB.Label LabelExtractMatrix 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Matrix"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   28
         Top             =   360
         Width           =   615
      End
      Begin VB.Label LabelExtractMatrixB2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Bound2"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   840
         TabIndex        =   27
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label LabelExtractMatrixB1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Bound1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1680
         Width           =   615
      End
   End
   Begin VB.CommandButton CommandClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
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
      Left            =   5280
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   840
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   240
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Calculate Binary or Pure Element PAR Files"
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
      Top             =   240
      Width           =   4935
      Begin VB.CheckBox CheckOverwriteHigherMinimumEnergyPAR 
         Caption         =   "Only Overwrite Higher Minimum Energy PAR Files"
         Height          =   255
         Left            =   840
         TabIndex        =   67
         Top             =   1080
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.CommandButton CommandCalculateRandom 
         Caption         =   "Calculate Random Binaries Using PAR Share Folder"
         Height          =   375
         Left            =   240
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3240
         Width           =   4335
      End
      Begin VB.OptionButton OptionBinaryMethod 
         Caption         =   "Binary Composition"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Calculate parameter files over a range of 11 binary compositions (1 to 99 percent) "
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton OptionBinaryMethod 
         Caption         =   "Pure Elements"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Calculate pure element parameter files"
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox ComboBinaryElement2 
         Height          =   315
         Left            =   720
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Select the ending element for binary fluorescence calculations"
         Top             =   1800
         Width           =   615
      End
      Begin VB.ComboBox ComboBinaryElement1 
         Height          =   315
         Left            =   720
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Select the starting element for binary fluorescence calculations"
         Top             =   1440
         Width           =   615
      End
      Begin VB.CheckBox CheckCalculateForMatrixRange 
         Caption         =   "Calculate PAR Files For Element Range"
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1800
         Width           =   3255
      End
      Begin VB.CommandButton CommandBinaryCalculate 
         BackColor       =   &H0080FFFF&
         Caption         =   "Calculate Binary Parameter Files"
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Calculate binary or pure element  parameter (PAR) files for the selected element range"
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Frame Frame6 
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Width           =   4575
         Begin VB.OptionButton OptionFromFormula 
            Caption         =   "From Formula"
            Height          =   255
            Left            =   720
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Calculate complete binary composition PAR files for a system specified by a formula"
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton OptionFromStandard 
            Caption         =   "From Standard"
            Height          =   255
            Left            =   2520
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Calculate complete binary composition PAR files for a system specified by a standard composition"
            Top             =   120
            Width           =   1575
         End
         Begin VB.CommandButton CommandCalculateComposition 
            Caption         =   "Calculate Binary PAR Files From Specified Composition"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Calculate all necessary binary or pure element parameter (PAR) files based on the specified formula or standard composition"
            Top             =   480
            Width           =   4335
         End
      End
      Begin VB.CheckBox CheckDoNotOverwritePAR 
         Caption         =   "Do Not Overwrite Existing .PAR Files"
         Height          =   255
         Left            =   840
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Skip calculations for existing PAR files"
         Top             =   600
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.CheckBox CheckOverwriteLowerPrecisionPAR 
         Caption         =   "Only Overwrite Lower Precision PAR Files"
         Height          =   255
         Left            =   840
         TabIndex        =   66
         Top             =   840
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.Label LabelToAnd 
         Alignment       =   2  'Center
         Caption         =   "Matrix2"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label LabelFrom 
         Alignment       =   2  'Center
         Caption         =   "Matrix1"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   495
      End
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7200
      Y1              =   8640
      Y2              =   8640
   End
End
Attribute VB_Name = "FormPenepma12Binary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2023 by John J. Donovan
Option Explicit

Private Sub CommandCalculateAlphas_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
Call Penepma12BinarySave
If ierror Then Exit Sub
Call Penepma12BinaryCalculateAlphaFactor
If ierror Then Exit Sub
End Sub

Private Sub CommandCalculateDensity_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12BinarySave
If ierror Then Exit Sub
Call Penepma12CalculateBinaryDensity
If ierror Then Exit Sub
End Sub

Private Sub CommandCalculateKratios_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
Call Penepma12BinarySave
If ierror Then Exit Sub
Call Penepma12CalculateKratios(FormPenepma12Binary)
If ierror Then Exit Sub
End Sub

Private Sub CommandCalculateRandom_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
Call Penepma12BinarySave
If ierror Then Exit Sub
Call Penepma12CalculateRandom
If ierror Then Exit Sub
End Sub

Private Sub CommandBinaryCalculate_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
Call Penepma12BinarySave
If ierror Then Exit Sub
Call Penepma12Calculate
If ierror Then Exit Sub
End Sub

Private Sub CommandCalculateComposition_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
Call Penepma12BinarySave
If ierror Then Exit Sub
Call Penepma12CalculateComposition
If ierror Then Exit Sub
End Sub

Private Sub CommandCheckPenfluorInputFiles_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12CheckPenfluorInputFiles
If ierror Then Exit Sub
End Sub

Private Sub CommandExtract_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
Call Penepma12BinarySave
If ierror Then Exit Sub
Call Penepma12Extract
If ierror Then Exit Sub
End Sub

Private Sub CommandExtract2_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
Call Penepma12BinarySave
If ierror Then Exit Sub
Call Penepma12Extract2
If ierror Then Exit Sub
End Sub

Private Sub CommandExtractBinary_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12ExtractBinary(FormPenepma12Binary)
If ierror Then Exit Sub
End Sub

Private Sub CommandExtractPureElementIntensities_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
Call Penepma12BinarySave
If ierror Then Exit Sub
Call Penepma12ExtractPure
If ierror Then Exit Sub
End Sub

Private Sub CommandExtractRandom_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
Call Penepma12BinarySave
If ierror Then Exit Sub
Call Penepma12ExtractRandom
If ierror Then Exit Sub
End Sub

Private Sub CommandOutputPlotData_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
Call Penepma12BinarySave
If ierror Then Exit Sub
Call Penepma12OutputPlotData
If ierror Then Exit Sub
End Sub

Private Sub CommandPenPFE_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
Call Penepma12BinarySave
If ierror Then Exit Sub
Call Penepma12Random        ' open the PenPFE control dialog
If ierror Then Exit Sub
End Sub

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormPenepma12Binary
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12BinarySave
If ierror Then Exit Sub
Unload FormPenepma12Binary
End Sub

Private Sub CommandTestLoadPenepmaAtomicWeights_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12AtomicWeights
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormPenepma12Binary)
HelpContextID = IOGetHelpContextID("FormPENEPMA12BINARY")
Call Penepma12BinaryLoad
If ierror Then Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub OptionBinaryMethod_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next
If Index% = 0 Then
FormPenepma12Binary.CommandBinaryCalculate.Caption = "Create Binary Parameter Files"
FormPenepma12Binary.CommandCalculateComposition.Caption = "Calculate Binary PAR Files From Specified Composition"
Else
FormPenepma12Binary.CommandBinaryCalculate.Caption = "Create Pure Element Parameter Files"
FormPenepma12Binary.CommandCalculateComposition.Caption = "Calculate Element PAR Files From Specified Composition"
End If
End Sub

Private Sub OptionExtractMethod_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next
If Index% = 0 Then
FormPenepma12Binary.ComboExtractElement.Enabled = True
FormPenepma12Binary.ComboExtractMatrix.Enabled = False

FormPenepma12Binary.ComboExtractMatrixA1.Enabled = True
FormPenepma12Binary.ComboExtractMatrixA2.Enabled = True
FormPenepma12Binary.ComboExtractMatrixB1.Enabled = True
FormPenepma12Binary.ComboExtractMatrixB2.Enabled = True

FormPenepma12Binary.CheckExtractForSpecifiedRange.Value = vbUnchecked
FormPenepma12Binary.CheckExtractForSpecifiedRange.Enabled = False
End If

If Index% = 1 Then
FormPenepma12Binary.ComboExtractElement.Enabled = True
FormPenepma12Binary.ComboExtractMatrix.Enabled = True

FormPenepma12Binary.ComboExtractMatrixA1.Enabled = False
FormPenepma12Binary.ComboExtractMatrixA2.Enabled = False
FormPenepma12Binary.ComboExtractMatrixB1.Enabled = False
FormPenepma12Binary.ComboExtractMatrixB2.Enabled = False
FormPenepma12Binary.CheckExtractForSpecifiedRange.Enabled = True
End If
End Sub

Private Sub TextConcA_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextConcAA_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextConcB_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextConcBB_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextDensityElementA_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextDensityElementB_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextKratA_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextKratB_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

