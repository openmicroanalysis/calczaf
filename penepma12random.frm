VERSION 5.00
Begin VB.Form FormPenepma12Random 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perform PenPFE Calculations"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   14385
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "Pure Element Intensities"
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
      Height          =   2295
      Left            =   11160
      TabIndex        =   75
      Top             =   2760
      Width           =   3135
      Begin VB.CommandButton CommandScanPure 
         Caption         =   "Scan Input Files and Write To Pure.MDB"
         Height          =   615
         Left            =   240
         TabIndex        =   77
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CommandButton CommandCreatePure 
         Caption         =   "Create New (Empty) Pure.MDB"
         Height          =   615
         Left            =   240
         TabIndex        =   76
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Create, Update Or Read Boundary.MDB K-Ratio Database for Penepma Boundary Correction Calculations"
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
      Height          =   3615
      Left            =   120
      TabIndex        =   34
      Top             =   5400
      Width           =   14175
      Begin VB.CommandButton CommandUpdateBoundary2 
         Caption         =   $"Penepma12Random.frx":0000
         Enabled         =   0   'False
         Height          =   1095
         Left            =   4320
         TabIndex        =   74
         Top             =   2400
         Width           =   3255
      End
      Begin VB.TextBox TextBoundaryB2 
         Height          =   285
         Left            =   9120
         TabIndex        =   73
         ToolTipText     =   "Enter concentration of boundary element B2"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox TextBoundaryB1 
         Height          =   285
         Left            =   7680
         TabIndex        =   71
         ToolTipText     =   "Enter concentration of boundary element B1"
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox TextMatrixA1 
         Height          =   285
         Left            =   7680
         TabIndex        =   69
         ToolTipText     =   "Enter concentration of matrix element A1"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox TextMatrixA2 
         Height          =   285
         Left            =   9120
         TabIndex        =   67
         ToolTipText     =   "Enter concentration of matrix element A2"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.ComboBox ComboBoundaryB1 
         Height          =   315
         Left            =   10920
         TabIndex        =   61
         TabStop         =   0   'False
         ToolTipText     =   "Select the first element in the boundary binary"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox ComboBoundaryB2 
         Height          =   315
         Left            =   12480
         TabIndex        =   60
         TabStop         =   0   'False
         ToolTipText     =   "Select the second element in the boundary binary"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox ComboMatrixA2 
         Height          =   315
         Left            =   12480
         TabIndex        =   59
         TabStop         =   0   'False
         ToolTipText     =   "Select the second element in the boundary binary"
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox ComboMatrixA1 
         Height          =   315
         Left            =   10920
         TabIndex        =   58
         TabStop         =   0   'False
         ToolTipText     =   "Select the first element in the boundary binary"
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton CommandUpdateBoundary 
         Caption         =   "Interpolate K-Ratios Based On TakeOff, Emitting Element, Matrix Binary and Boundary Binary At Specified Mass Distance (Only)"
         Height          =   495
         Left            =   4320
         TabIndex        =   57
         Top             =   1680
         Width           =   6135
      End
      Begin VB.TextBox TextDensityB 
         Height          =   285
         Left            =   12480
         TabIndex        =   56
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox TextDensityA 
         Height          =   285
         Left            =   10920
         TabIndex        =   54
         Top             =   3120
         Width           =   1335
      End
      Begin VB.OptionButton OptionDistance 
         Caption         =   "Use Mass"
         Height          =   255
         Index           =   1
         Left            =   12600
         TabIndex        =   52
         Top             =   1920
         Width           =   1335
      End
      Begin VB.OptionButton OptionDistance 
         Caption         =   "Use Microns"
         Height          =   255
         Index           =   0
         Left            =   11040
         TabIndex        =   51
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox TextDistanceMass 
         Height          =   285
         Left            =   12480
         TabIndex        =   50
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox TextDistanceMicrons 
         Height          =   285
         Left            =   10920
         TabIndex        =   48
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton CommandCreateBoundary 
         Caption         =   "Create New (Empty) Boundary.MDB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1200
         Width           =   3255
      End
      Begin VB.CommandButton CommandScanBoundary 
         Caption         =   "Scan Input Files and Write To Boundary.MDB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   2040
         Width           =   3255
      End
      Begin VB.CommandButton CommandReadBoundary 
         Caption         =   "Read Boundary.MDB (for specified energy, emitter, x-ray and matrix/boundary binary)"
         Height          =   375
         Left            =   4320
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   960
         Width           =   6135
      End
      Begin VB.ComboBox ComboEmitterElement2 
         Height          =   315
         Left            =   7440
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "Select the measured element to profile fluorescence"
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox ComboEmitterXRay2 
         Height          =   315
         Left            =   8400
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Select the measured x-ray to profile fluorescence"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TextBeamTakeoff2 
         Height          =   285
         Left            =   5160
         TabIndex        =   36
         ToolTipText     =   "Enter beam takeoff angle in degrees (40 for JEOL and Cameca)"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TextBeamEnergy2 
         Height          =   285
         Left            =   6240
         TabIndex        =   35
         ToolTipText     =   "Enter beam energy in electron volts (eV)"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Bound2"
         Height          =   255
         Left            =   9120
         TabIndex        =   72
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Bound1"
         Height          =   255
         Left            =   7680
         TabIndex        =   70
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Matrix1"
         Height          =   255
         Left            =   7680
         TabIndex        =   68
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Matrix2"
         Height          =   255
         Left            =   9120
         TabIndex        =   66
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label LabelExtractMatrixB1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Bound1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10920
         TabIndex        =   65
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label LabelExtractMatrixB2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Bound2"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12480
         TabIndex        =   64
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label LabelExtractMatrixA2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Matrix2"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12480
         TabIndex        =   63
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label LabelExtractMatrixA1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Matrix1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10920
         TabIndex        =   62
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Boundary Density"
         Height          =   255
         Left            =   12480
         TabIndex        =   55
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Incident Density"
         Height          =   255
         Left            =   10920
         TabIndex        =   53
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Distance (ug/cm2)"
         Height          =   255
         Left            =   12480
         TabIndex        =   49
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Distance (um)"
         Height          =   255
         Left            =   10920
         TabIndex        =   47
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Element"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7440
         TabIndex        =   46
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "X-Ray"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8400
         TabIndex        =   45
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "KeV"
         Height          =   255
         Left            =   6240
         TabIndex        =   44
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Take-off"
         Height          =   255
         Left            =   5160
         TabIndex        =   43
         Top             =   360
         Width           =   855
      End
      Begin VB.Label LabelBoundaryDisplay 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4320
         TabIndex        =   42
         Top             =   1320
         Width           =   6135
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Create, Update Or Read Matrix.MDB K-Ratio Database for Penepma Matrix Correction Calculations"
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
      Height          =   2415
      Left            =   120
      TabIndex        =   19
      Top             =   2760
      Width           =   10935
      Begin VB.CommandButton CommandAddMatrix 
         Caption         =   "Add Penepma K-Ratios To Matrix.MDB"
         Height          =   495
         Left            =   2640
         TabIndex        =   80
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton CommandCheckDeviations 
         Caption         =   "Check Database Alpha Fit Deviations"
         Height          =   375
         Left            =   360
         TabIndex        =   79
         Top             =   1920
         Width           =   4215
      End
      Begin VB.CommandButton CommandCheckKratios 
         Caption         =   "Check Database Kratios Against CalcZAF"
         Height          =   375
         Left            =   360
         TabIndex        =   78
         Top             =   1560
         Width           =   4215
      End
      Begin VB.TextBox TextBeamEnergy 
         Height          =   285
         Left            =   6240
         TabIndex        =   30
         ToolTipText     =   "Enter beam energy in electron volts (eV)"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TextBeamTakeoff 
         Height          =   285
         Left            =   5160
         TabIndex        =   29
         ToolTipText     =   "Enter beam takeoff angle in degrees (40 for JEOL and Cameca)"
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox ComboMatrixElement 
         Height          =   315
         Left            =   9720
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Select the boundary or matrix element"
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox ComboEmitterXRay 
         Height          =   315
         Left            =   8400
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Select the measured x-ray to profile fluorescence"
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox ComboEmitterElement 
         Height          =   315
         Left            =   7440
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Select the measured element to profile fluorescence"
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton CommandReadMatrix 
         BackColor       =   &H0080FFFF&
         Caption         =   "Read Matrix.MDB (for specified energy, emitter, x-ray and matrix)"
         Height          =   495
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1080
         Width           =   5655
      End
      Begin VB.CommandButton CommandScanMatrix 
         Caption         =   "Scan Input Files and Write To Matrix.MDB"
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
         Left            =   360
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton CommandCreateMatrix 
         Caption         =   "Create New (Empty) Matrix.MDB"
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
         Left            =   360
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label LabelMatrixDisplay 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   4920
         TabIndex        =   33
         Top             =   1680
         Width           =   5655
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Take-off"
         Height          =   255
         Left            =   5160
         TabIndex        =   32
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Caption         =   "KeV"
         Height          =   255
         Left            =   6240
         TabIndex        =   31
         Top             =   360
         Width           =   735
      End
      Begin VB.Label LabelBoundaryMatrix 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Matrix"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9720
         TabIndex        =   28
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "X-Ray"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8400
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Element"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7440
         TabIndex        =   25
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H00C0FFC0&
      Cancel          =   -1  'True
      Caption         =   "Close"
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
      Height          =   615
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   240
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Perform Multiple Instance Penfluor/Fitall/Fanal Calculations"
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
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10935
      Begin VB.TextBox TextPenPFEPath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   360
         Width           =   7935
      End
      Begin VB.CommandButton CommandBrowsePenPFE 
         Caption         =   "Browse"
         Height          =   255
         Left            =   9720
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox TextPARSharePath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   720
         Width           =   7935
      End
      Begin VB.CommandButton CommandBrowsePARShare 
         Caption         =   "Browse"
         Height          =   255
         Left            =   9720
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox TextNumberOfInstances 
         Height          =   285
         Left            =   3600
         TabIndex        =   9
         Top             =   1200
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Caption         =   "PAR File Calculations"
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
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   5175
         Begin VB.OptionButton OptionBinaryMethod 
            Caption         =   "Pure Elements"
            Height          =   195
            Index           =   1
            Left            =   2400
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Calculate pure element compositions over the range of elements specified"
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton OptionBinaryMethod 
            Caption         =   "Binary Composition"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Calculate a range of binary compositions for the two elements specified"
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton CommandCalculate 
            Caption         =   "Calculate"
            Height          =   375
            Left            =   4080
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "K-Ratio Extractions"
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
         Height          =   735
         Left            =   5640
         TabIndex        =   1
         Top             =   1560
         Width           =   5175
         Begin VB.OptionButton OptionExtractMethod 
            Caption         =   "Boundary"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Extract boundary fluorescence effects for the measured element and the matrix element"
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton OptionExtractMethod 
            Caption         =   "Matrix Only"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Extract matrix (self) fluorescence without boundary effect"
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton CommandExtract 
            Caption         =   "Extract"
            Height          =   375
            Left            =   4080
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Label Label1 
         Caption         =   "PenPFE Path"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "PAR Share Path"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Number of Penfluor/Fanal Instances To Create"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   3495
      End
   End
   Begin VB.Label LabelCopying 
      Alignment       =   2  'Center
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
      Height          =   1695
      Left            =   11160
      TabIndex        =   18
      Top             =   1080
      Width           =   3135
   End
End
Attribute VB_Name = "FormPenepma12Random"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2017 by John J. Donovan
Option Explicit

Private Sub CommandAddMatrix_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12MatrixAddPenepmaKRatios
If ierror Then Exit Sub
End Sub

Private Sub CommandBrowsePARShare_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12RandomBrowseFolder(Int(2))
If ierror Then Exit Sub
End Sub

Private Sub CommandBrowsePenPFE_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12RandomBrowseFolder(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub CommandCalculate_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12RandomSave
If ierror Then Exit Sub
FormPenepma12Random.LabelCopying.Caption = "Please wait for application and Penepma12 folder copy operations to complete!"
Call Penepma12RandomRun
FormPenepma12Random.LabelCopying.Caption = vbNullString
If ierror Then Exit Sub
Call Penepma12RandomRunPenfluor
If ierror Then Exit Sub
End Sub

Private Sub CommandCheckDeviations_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12MatrixCheckDeviations(CSng(40#))
If ierror Then Exit Sub
Call Penepma12MatrixCheckDeviations(CSng(52.5))
If ierror Then Exit Sub
Call Penepma12MatrixCheckDeviations(CSng(75#))
If ierror Then Exit Sub
End Sub

Private Sub CommandCheckKratios_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12MatrixCheckKratios
If ierror Then Exit Sub
End Sub

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12RandomSave
If ierror Then Exit Sub
Unload FormPenepma12Random
End Sub

Private Sub CommandCreateBoundary_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12BoundaryNewMDB
If ierror Then Exit Sub
End Sub

Private Sub CommandCreateMatrix_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12MatrixNewMDB
If ierror Then Exit Sub
End Sub

Private Sub CommandCreatePure_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12PureNewMDB
If ierror Then Exit Sub
End Sub

Private Sub CommandExtract_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12RandomSave
If ierror Then Exit Sub
FormPenepma12Random.LabelCopying.Caption = "Please wait for application and Penepma12 folder copy operations to complete!"
Call Penepma12RandomRun
FormPenepma12Random.LabelCopying.Caption = vbNullString
If ierror Then Exit Sub
Call Penepma12RandomRunFanal
If ierror Then Exit Sub
End Sub

Private Sub CommandReadBoundary_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12RandomSave
If ierror Then Exit Sub
Call Penepma12RandomReadBoundary
If ierror Then Exit Sub
End Sub

Private Sub CommandReadMatrix_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12RandomSave
If ierror Then Exit Sub
Call Penepma12RandomReadMatrix
If ierror Then Exit Sub
End Sub

Private Sub CommandScanBoundary_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12BoundaryScanMDB
If ierror Then Exit Sub
End Sub

Private Sub CommandScanMatrix_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12MatrixScanMDB
If ierror Then Exit Sub
End Sub

Private Sub CommandScanPure_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12PureScanMDB
If ierror Then Exit Sub
End Sub

Private Sub CommandUpdateBoundary_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12RandomSave
If ierror Then Exit Sub
Call Penepma12RandomBoundaryInterpolate
If ierror Then Exit Sub
End Sub

Private Sub CommandUpdateBoundary2_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12RandomSave
If ierror Then Exit Sub

End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
icancelload = False
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormPenepma12Random)
HelpContextID = IOGetHelpContextID("FormPenepma12Random")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub OptionDistance_Click(Index As Integer)
If Index% = 0 Then
FormPenepma12Random.TextDistanceMicrons.Enabled = True
FormPenepma12Random.TextDistanceMass.Enabled = False
Else
FormPenepma12Random.TextDistanceMicrons.Enabled = False
FormPenepma12Random.TextDistanceMass.Enabled = True
End If
End Sub

Private Sub TextBeamEnergy_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextBeamEnergy2_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextBeamTakeoff_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextBeamTakeoff2_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextBoundaryB1_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextBoundaryB2_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextDensityA_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextDensityB_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextDistanceMass_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextDistanceMicrons_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMatrixA1_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMatrixA2_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextNumberOfInstances_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextPARSharePath_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextPenPFEPath_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub
