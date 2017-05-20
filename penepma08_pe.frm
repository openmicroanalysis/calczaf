VERSION 5.00
Object = "{6E5043E8-C452-4A6A-B011-9B5687112610}#1.0#0"; "Pesgo32f.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormPENEPMA08_PE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create PENEPMA Material and Input Files"
   ClientHeight    =   12480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13950
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
   ScaleHeight     =   12480
   ScaleWidth      =   13950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CommandPlot 
      BackColor       =   &H0080FFFF&
      Caption         =   "Plot Spectrum"
      Height          =   495
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   103
      TabStop         =   0   'False
      ToolTipText     =   "Plot a spectrum from a previous calculation"
      Top             =   5160
      Width           =   1695
   End
   Begin Pesgo32fLib.Pesgo Pesgo1 
      Height          =   3615
      Left            =   120
      TabIndex        =   97
      Top             =   8760
      Width           =   11055
      _Version        =   65536
      _ExtentX        =   19500
      _ExtentY        =   6376
      _StockProps     =   96
      _AllProps       =   "PENEPMA08_PE.frx":0000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Graph Display"
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   11280
      TabIndex        =   93
      Top             =   10200
      Width           =   2535
      Begin VB.OptionButton OptionDisplayGraph 
         Caption         =   "Total X-ray Spectrum"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   96
         TabStop         =   0   'False
         ToolTipText     =   "Display current total x-ray spectrum"
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton OptionDisplayGraph 
         Caption         =   "Char X-ray Spectrum"
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
         TabIndex        =   95
         TabStop         =   0   'False
         ToolTipText     =   "Display characteristic x-ray spectrum"
         Top             =   960
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.OptionButton OptionDisplayGraph 
         Caption         =   "BSE Energy Spectrum"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   94
         TabStop         =   0   'False
         ToolTipText     =   "Display backscatter energy spectrum"
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.CommandButton CommandHelp 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Help"
      Height          =   495
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   84
      TabStop         =   0   'False
      ToolTipText     =   "Click this button to get detailed help from our on-line user forum"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton CommandClipboard 
      Caption         =   "Clipboard"
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
      Left            =   12720
      TabIndex        =   80
      TabStop         =   0   'False
      ToolTipText     =   "Copy the graph to the system clipboard"
      Top             =   12000
      Width           =   1095
   End
   Begin VB.CommandButton CommandBatch 
      BackColor       =   &H0080FFFF&
      Caption         =   "Batch Mode"
      Height          =   495
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   76
      TabStop         =   0   'False
      ToolTipText     =   "Run a number of PENEPMA simulations in ""batch"" mode"
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CheckBox CheckUseLogScale 
      Caption         =   "Use Log Scale"
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
      Left            =   11520
      TabIndex        =   75
      TabStop         =   0   'False
      ToolTipText     =   "Use log scale for Y axis"
      Top             =   11640
      Width           =   2055
   End
   Begin VB.Frame Frame4 
      Caption         =   "Output Parameters"
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   11280
      TabIndex        =   67
      Top             =   8520
      Width           =   2535
      Begin VB.TextBox TextElapsedTime 
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   960
         TabIndex        =   70
         TabStop         =   0   'False
         ToolTipText     =   "Current elapsed time"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox TextElapsedShowers 
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   960
         TabIndex        =   69
         TabStop         =   0   'False
         ToolTipText     =   "Current number of simulated showers"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TextElapsedBSE 
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   960
         TabIndex        =   68
         TabStop         =   0   'False
         ToolTipText     =   "Current electron backscatter coefficient"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Time"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Showers"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "BSE Frac."
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
         TabIndex        =   71
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.CommandButton CommandZoomFull 
      Caption         =   "Zoom Full"
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
      Left            =   11280
      TabIndex        =   66
      TabStop         =   0   'False
      ToolTipText     =   "Rescale the graph to display all the spectrum data"
      Top             =   12000
      Width           =   1215
   End
   Begin VB.CheckBox CheckUseGridLines 
      Caption         =   "Use Grid Lines"
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
      Left            =   11520
      TabIndex        =   65
      TabStop         =   0   'False
      ToolTipText     =   "Show grid lines on graph"
      Top             =   11400
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Left            =   11280
      Top             =   7440
   End
   Begin VB.CommandButton CommandRunPENEPMA 
      BackColor       =   &H0080FFFF&
      Caption         =   "Run Input File In PENEPMA"
      Default         =   -1  'True
      Height          =   615
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   47
      TabStop         =   0   'False
      ToolTipText     =   "Run the current (last created) Input file in PENEPMA (results will be displayed below during the simulation)"
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton CommandEditInput 
      Caption         =   "Edit Input File"
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
      Left            =   12120
      TabIndex        =   46
      TabStop         =   0   'False
      ToolTipText     =   "Edit the current PENEPMA input file using the default text editor"
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton CommandDeleteDumpFiles 
      Caption         =   "Delete Dump Files"
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
      Left            =   12120
      TabIndex        =   45
      TabStop         =   0   'False
      ToolTipText     =   "Delete the temporary ""dump"" files for starting an interrupted simulation over from the beginning"
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton CommandPENEPMAPrompt 
      Caption         =   "PENEPMA Prompt"
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
      Left            =   12120
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Create a command prompt in the PENEPMA directory. Type: penepma < ""input file name"""
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton CommandPENDBASEPrompt 
      Caption         =   "PENDBASE Prompt"
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
      Left            =   12120
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Create a command prompt in the PENDBASE directory"
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Height          =   495
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton CommandClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   12120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "PENEPMA Input File (*.INI)"
      ForeColor       =   &H00FF0000&
      Height          =   7815
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   7095
      Begin VB.TextBox TextEnergyRangeMinMaxNumber 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6120
         TabIndex        =   102
         ToolTipText     =   "Enter the number of energy channels for the Penepma simulation"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox TextEnergyRangeMinMaxNumber 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   5280
         TabIndex        =   101
         ToolTipText     =   "Enter the maximum energy range for the Penepma simulation"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox TextEnergyRangeMinMaxNumber 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4440
         TabIndex        =   100
         ToolTipText     =   "Enter the minimum energy range for the Penepma simulation"
         Top             =   1800
         Width           =   735
      End
      Begin MSComCtl2.UpDown UpDownXray 
         Height          =   375
         Index           =   0
         Left            =   5640
         TabIndex        =   98
         TabStop         =   0   'False
         ToolTipText     =   "Increment/decrement x-ray line for optimization"
         Top             =   4320
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Frame Frame5 
         Caption         =   "Detector Geometry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   3960
         TabIndex        =   87
         Top             =   5640
         Width           =   3015
         Begin VB.OptionButton OptionDetectorGeometry 
            Caption         =   "West"
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
            Left            =   2160
            TabIndex        =   92
            TabStop         =   0   'False
            ToolTipText     =   "Select the annular detector geometry option for greatest sensitivity"
            Top             =   480
            Width           =   735
         End
         Begin VB.OptionButton OptionDetectorGeometry 
            Caption         =   "South"
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
            Left            =   1440
            TabIndex        =   91
            TabStop         =   0   'False
            ToolTipText     =   "Select the annular detector geometry option for greatest sensitivity"
            Top             =   480
            Width           =   735
         End
         Begin VB.OptionButton OptionDetectorGeometry 
            Caption         =   "East"
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
            Left            =   840
            TabIndex        =   90
            TabStop         =   0   'False
            ToolTipText     =   "Select the annular detector geometry option for greatest sensitivity"
            Top             =   480
            Width           =   615
         End
         Begin VB.OptionButton OptionDetectorGeometry 
            Caption         =   "North"
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
            TabIndex        =   89
            TabStop         =   0   'False
            ToolTipText     =   "Select the annular detector geometry option for greatest sensitivity"
            Top             =   480
            Width           =   735
         End
         Begin VB.OptionButton OptionDetectorGeometry 
            Caption         =   "Annular (0 to 360 degrees)"
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
            Left            =   360
            TabIndex        =   88
            TabStop         =   0   'False
            ToolTipText     =   "Select the annular detector geometry option for greatest sensitivity"
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.CommandButton CommandBrowseInputFiles 
         Caption         =   "Browse"
         Height          =   375
         Left            =   120
         TabIndex        =   83
         TabStop         =   0   'False
         ToolTipText     =   "Browse to an existing Penepma input file and modify and save again if desired"
         Top             =   7320
         Width           =   975
      End
      Begin VB.OptionButton OptionProduction 
         Caption         =   "Optimize Production of Thin Film X-rays (use bilayer *.geo files)"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   79
         TabStop         =   0   'False
         ToolTipText     =   "Based on template file Bilayer.in"
         Top             =   3600
         Width           =   5895
      End
      Begin VB.TextBox TextBeamTakeoff 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   62
         ToolTipText     =   "Enter beam takeoff angle in degrees (40 for JEOL and Cameca)"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox TextDumpPeriod 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   61
         ToolTipText     =   "Enter the time interval for the dump files to be updated (for live display)"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton CommandElement 
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
         Index           =   1
         Left            =   6480
         Picture         =   "PENEPMA08_PE.frx":6114
         Style           =   1  'Graphical
         TabIndex        =   60
         TabStop         =   0   'False
         ToolTipText     =   "Select a specific element for optimization"
         Top             =   5040
         Width           =   495
      End
      Begin VB.CommandButton CommandAdjust 
         Caption         =   "Adjust"
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
         Index           =   0
         Left            =   5880
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   $"PENEPMA08_PE.frx":6818
         Top             =   4320
         Width           =   615
      End
      Begin VB.CommandButton CommandElement 
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
         Index           =   0
         Left            =   6480
         Picture         =   "PENEPMA08_PE.frx":68B3
         Style           =   1  'Graphical
         TabIndex        =   59
         TabStop         =   0   'False
         ToolTipText     =   "Select a specific element for optimization"
         Top             =   4320
         Width           =   495
      End
      Begin VB.CommandButton CommandAdjust 
         Caption         =   "Adjust"
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
         Index           =   1
         Left            =   5880
         TabIndex        =   58
         TabStop         =   0   'False
         ToolTipText     =   $"PENEPMA08_PE.frx":6FB7
         Top             =   5040
         Width           =   615
      End
      Begin VB.TextBox TextEABS2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4800
         TabIndex        =   54
         ToolTipText     =   "Minimum photon absorption energy for this material"
         Top             =   5040
         Width           =   855
      End
      Begin VB.TextBox TextEABS1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   53
         ToolTipText     =   "Minimum electron interaction energy for this material"
         Top             =   5040
         Width           =   855
      End
      Begin VB.TextBox TextEABS2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4800
         TabIndex        =   52
         ToolTipText     =   "Minimum photon absorption energy for this material"
         Top             =   4320
         Width           =   855
      End
      Begin VB.TextBox TextEABS1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3960
         TabIndex        =   51
         ToolTipText     =   "Minimum electron interaction energy for this material"
         Top             =   4320
         Width           =   855
      End
      Begin VB.TextBox TextMaterialFiles 
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
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   50
         ToolTipText     =   $"PENEPMA08_PE.frx":7052
         Top             =   5040
         Width           =   2775
      End
      Begin VB.CommandButton CommandBrowseMaterialFiles 
         Caption         =   "Browse"
         Height          =   375
         Index           =   1
         Left            =   3000
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   "Browse for a previously created material file"
         Top             =   4920
         Width           =   855
      End
      Begin VB.TextBox TextBeamDirection 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4440
         TabIndex        =   44
         ToolTipText     =   "Enter beam direction phi (azimuthal angle in degrees)"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox TextInputTitle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   43
         ToolTipText     =   "Enter title for this input file (up to 120 characters)"
         Top             =   360
         Width           =   5055
      End
      Begin VB.OptionButton OptionProduction 
         Caption         =   "Optimize Secondary Fluorescent X-rays (use couple or sphere *.geo files)"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Based on template file CuFe_sec.in"
         Top             =   3360
         Width           =   6615
      End
      Begin VB.OptionButton OptionProduction 
         Caption         =   "Optimize Production of Continuum X-rays"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   40
         ToolTipText     =   "Based on template file Cu_cont.in"
         Top             =   3120
         Width           =   4575
      End
      Begin VB.OptionButton OptionProduction 
         Caption         =   "Optimize Production of Backscatter Electrons"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   39
         ToolTipText     =   "Based on template file Cu_back.in"
         Top             =   2880
         Width           =   4575
      End
      Begin VB.OptionButton OptionProduction 
         Caption         =   "Optimize Production of Characteristic X-rays"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   38
         ToolTipText     =   "Based on template file Cu_cha.in"
         Top             =   2640
         Width           =   4575
      End
      Begin VB.TextBox TextSimulationTimePeriod 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   36
         ToolTipText     =   "Enter the total simulation time period in seconds"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox TextNumberSimulatedShowers 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   34
         ToolTipText     =   "Enter the number of simulated showers (incident electrons)"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox TextInputFile 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   31
         ToolTipText     =   "Enter the file name of the target PENEPMA Input file"
         Top             =   6960
         Width           =   6855
      End
      Begin VB.TextBox TextBeamAperture 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   30
         ToolTipText     =   "Enter beam aperture (in degrees)"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox TextBeamDirection 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   29
         ToolTipText     =   "Enter beam direction theta (polar angle in degrees)"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton CommandBrowseGeometry 
         Caption         =   "Browse"
         Height          =   375
         Left            =   3000
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Browse for a specified PENEPMA geometry file"
         Top             =   6000
         Width           =   855
      End
      Begin VB.CommandButton CommandBrowseMaterialFiles 
         Caption         =   "Browse"
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Browse for a previously created material file"
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox TextGeometryFile 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Enter (or browse to) the full path and file name of the Penepma Geometry file (*.GEO)"
         Top             =   6120
         Width           =   2775
      End
      Begin VB.TextBox TextMaterialFiles 
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
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   $"PENEPMA08_PE.frx":70E5
         Top             =   4320
         Width           =   2775
      End
      Begin VB.TextBox TextBeamPosition 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   5760
         TabIndex        =   11
         ToolTipText     =   "Enter beam position z in cm units (working distance)"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox TextBeamPosition 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4440
         TabIndex        =   10
         ToolTipText     =   "Enter beam position y in cm units"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TextBeamPosition 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   9
         ToolTipText     =   "Enter beam position x in cm units (1e-3 cm = 10 um)"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox TextBeamEnergy 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   8
         ToolTipText     =   "Enter beam energy in electron volts (eV)"
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton CommandOutputInputFile 
         BackColor       =   &H0080FFFF&
         Caption         =   "Create PENEPMA Input File"
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Create the specified penepma input file (make sure the material and geometry files are correctly specified)"
         Top             =   7320
         Width           =   4095
      End
      Begin MSComCtl2.UpDown UpDownXray 
         Height          =   375
         Index           =   1
         Left            =   5640
         TabIndex        =   99
         TabStop         =   0   'False
         ToolTipText     =   "Increment/decrement x-ray line for optimization"
         Top             =   5040
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label LabelMaterialFiles 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   82
         Top             =   4800
         Width           =   2775
      End
      Begin VB.Label LabelMaterialFiles 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   81
         Top             =   4080
         Width           =   2775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Minimum Electron/Photon Energy (eV)"
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
         Left            =   3960
         TabIndex        =   55
         Top             =   4080
         Width           =   2895
      End
      Begin VB.Label Label16 
         Caption         =   "Input File Title"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   1815
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   6960
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label14 
         Caption         =   "Dump Time (sec), Range (min, max, num)"
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
         TabIndex        =   35
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label Label11 
         Caption         =   "Number Showers, Simulation Time"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   2160
         Width           =   3015
      End
      Begin VB.Line Line4 
         X1              =   120
         X2              =   6960
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   6960
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6960
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Label Label9 
         Caption         =   "File Name of Target Input File (for above specified parameters). Create in PENEPMA_Path"
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
         TabIndex        =   32
         Top             =   6720
         Width           =   6735
      End
      Begin VB.Label Label8 
         Caption         =   "Beam Theta, Phi, Aperture (deg)"
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
         TabIndex        =   28
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Geometry File (*.GEO)"
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
         TabIndex        =   25
         Top             =   5880
         Width           =   4095
      End
      Begin VB.Label Label12 
         Caption         =   "Beam Position (um) (X, Y, Z)"
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
         TabIndex        =   7
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label13 
         Caption         =   "Take-off, Beam Energy (eV)"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   3015
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "PENEPMA Materials File (*.MAT)"
      ForeColor       =   &H00FF0000&
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton CommandOutputWeight 
         BackColor       =   &H0080FFFF&
         Caption         =   "Create PENEPMA Material From Weight"
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   86
         TabStop         =   0   'False
         ToolTipText     =   "Create the specified material based on the entered weight percent string (no decimals)"
         Top             =   7200
         Width           =   3615
      End
      Begin VB.CommandButton CommandOutputFormula 
         BackColor       =   &H0080FFFF&
         Caption         =   "Create PENEPMA Material From Formula"
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   77
         TabStop         =   0   'False
         ToolTipText     =   "Create the specified material based on the entered chemical formula"
         Top             =   6720
         Width           =   3615
      End
      Begin VB.CheckBox CheckMaterialLoadOrder 
         Caption         =   "Reverse Material Load Order In Input File"
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
         Left            =   360
         TabIndex        =   56
         TabStop         =   0   'False
         ToolTipText     =   "Load the material files in reverse order, e.g., CuFe instead of FeCu"
         Top             =   5760
         Width           =   4095
      End
      Begin VB.TextBox TextMaterialWcb 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   19
         ToolTipText     =   "Enter the oscillator energy (0 = use default)"
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox TextMaterialFcb 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         ToolTipText     =   "Enter the oscillator strength  (0 = use default)"
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox TextMaterialDensity 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   15
         ToolTipText     =   "Enter the material density in gm/cm^3"
         Top             =   4920
         Width           =   1215
      End
      Begin VB.ListBox ListAvailableStandards 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3960
         Left            =   120
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Standard compositions available for output"
         Top             =   600
         Width           =   4335
      End
      Begin VB.CommandButton CommandOutputMaterial 
         BackColor       =   &H0080FFFF&
         Caption         =   "Create PENEPMA Material From List"
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Create the specified materials file for penempa calculations"
         Top             =   6120
         Width           =   3615
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   4440
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Oscillator strength (Fcb) and Oscillator energy (Wcb) of the plasmon should be zero for insulators"
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
         TabIndex        =   20
         Top             =   5280
         Width           =   3855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Oscillator Energy"
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
         Top             =   4680
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Oscillator Strength"
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
         Left            =   1560
         TabIndex        =   16
         Top             =   4680
         Width           =   1335
      End
      Begin VB.Label LabelMaterialDensity 
         Alignment       =   2  'Center
         Caption         =   "Material Density"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Select 1 or 2 Material(s) for Output"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.TextBox TextTemp 
      Height          =   285
      Left            =   12120
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   7320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label LabelElapsedTime 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   12120
      TabIndex        =   78
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Label LabelProgress 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   11280
      TabIndex        =   74
      Top             =   8040
      Width           =   2535
   End
   Begin VB.Label LabelYPos 
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
      TabIndex        =   64
      Top             =   8400
      Width           =   975
   End
   Begin VB.Label LabelXPos 
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
      TabIndex        =   63
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   $"PENEPMA08_PE.frx":7171
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
      Left            =   1320
      TabIndex        =   48
      Top             =   8040
      Width           =   9615
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   $"PENEPMA08_PE.frx":72C3
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   12120
      TabIndex        =   37
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "FormPENEPMA08_PE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2017 by John J. Donovan
Option Explicit

Dim ProductionIndex As Integer

Private Sub CheckUseGridlines_Click()
If Not DebugMode Then On Error Resume Next
If FormPENEPMA08_PE.CheckUseGridLines.Value = vbChecked Then
FormPENEPMA08_PE.Pesgo1.GridLineControl = PEGLC_BOTH& ' show x and y grid
FormPENEPMA08_PE.Pesgo1.GridBands = True ' adds colour banding on background
Else
FormPENEPMA08_PE.Pesgo1.GridLineControl = PEGLC_NONE&
FormPENEPMA08_PE.Pesgo1.GridBands = False ' removes colour banding on background
End If
End Sub

Private Sub CheckUseLogScale_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08PlotLog
If ierror Then Exit Sub
Call Penepma08GraphUpdate(GraphDisplayOption%)
If ierror Then Exit Sub
End Sub

Private Sub CommandAdjust_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next
ProductionIndex% = ProductionIndex% + 1
If ProductionIndex% > MAXPRODUCTION% Then ProductionIndex% = 0
Call Penepma08AdjustEABS(ProductionIndex%, Index%, FormPENEPMA08_PE)
If ierror Then Exit Sub
End Sub

Private Sub CommandBatch_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08BatchLoad
If ierror Then Exit Sub
FormPENEPMA08Batch.Show vbModeless
End Sub

Private Sub CommandBrowseGeometry_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08BrowseGeometryFile(FormPENEPMA08_PE)
If ierror Then Exit Sub
End Sub

Private Sub CommandBrowseInputFiles_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08BrowseInputFile(FormPENEPMA08_PE)
If ierror Then Exit Sub
End Sub

Private Sub CommandBrowseMaterialFiles_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call Penepma08BrowseMaterialFile(Index% + 1, FormPENEPMA08_PE)
If ierror Then Exit Sub
Call Penepma08AdjustEABS(ProductionIndex%, Index%, FormPENEPMA08_PE)
If ierror Then Exit Sub
End Sub

Private Sub CommandClipboard_Click()
If Not DebugMode Then On Error Resume Next
FormPENEPMA08_PE.Pesgo1.AllowExporting = True
'FormPENEPMA08_PE.Pesgo1.ExportImageLargeFont = False
'FormPENEPMA08_PE.Pesgo1.ExportImageDpi = 450
Call FormPENEPMA08_PE.Pesgo1.PEcopybitmaptoclipboard(1200, 600)
End Sub

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
icancelload = True
Unload FormPENEPMA08Batch
Unload FormPENEPMA08_PE
End Sub

Private Sub CommandDeleteDumpFiles_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08DeleteDumpFiles
If ierror Then Exit Sub
End Sub

Private Sub CommandEditInput_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08EditInput
If ierror Then Exit Sub
End Sub

Private Sub CommandElement_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call Penepma08AdjustEABS(Int(MAXPRODUCTION% + 1), Index%, FormPENEPMA08_PE)
If ierror Then Exit Sub
End Sub

Private Sub CommandHelp_Click()
If Not DebugMode Then On Error Resume Next
Call IOBrowseHTTP(ProbeSoftwareInternetBrowseMethod%, "http://probesoftware.com/smf/index.php?topic=59.msg221#msg221")
If ierror Then Exit Sub
End Sub

Private Sub CommandOutputFormula_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08SaveMaterial(FormPENEPMA08_PE)
If ierror Then Exit Sub
Call Penepma08CreateMaterialFormula(Int(1), FormPENEPMA08_PE)
If ierror Then Exit Sub
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08SaveInput(FormPENEPMA08_PE)
If ierror Then Exit Sub
Call Penepma08SaveMaterial(FormPENEPMA08_PE)
If ierror Then Exit Sub
Call Penepma08SaveDisplay(FormPENEPMA08_PE)
If ierror Then Exit Sub
Unload FormPENEPMA08Batch
Unload FormPENEPMA08_PE
End Sub

Private Sub CommandOutputInputFile_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08SaveInput(FormPENEPMA08_PE)
If ierror Then Exit Sub
Call Penepma08CreateInput(Int(0))
If ierror Then Exit Sub
End Sub

Private Sub CommandOutputMaterial_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08SaveMaterial(FormPENEPMA08_PE)
If ierror Then Exit Sub
Call Penepma08CreateMaterial(FormPENEPMA08_PE)
If ierror Then Exit Sub
End Sub

Private Sub CommandOutputWeight_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08SaveMaterial(FormPENEPMA08_PE)
If ierror Then Exit Sub
Call Penepma08CreateMaterialFormula(Int(2), FormPENEPMA08_PE)
If ierror Then Exit Sub
End Sub

Private Sub CommandPENDBASEPrompt_Click()
If Not DebugMode Then On Error Resume Next
Dim taskID As Long
ChDrive PENDBASE_Path$
taskID& = Shell("cmd.exe /k cd " & VbDquote$ & PENDBASE_Path$ & VbDquote$, vbNormalFocus)
End Sub

Private Sub CommandPENEPMAPrompt_Click()
If Not DebugMode Then On Error Resume Next
Dim taskID As Long
ChDrive PENEPMA_Path$
taskID& = Shell("cmd.exe /k cd " & VbDquote$ & PENEPMA_Path$ & VbDquote$, vbNormalFocus)
End Sub

Private Sub CommandPlot_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08PlotSpectra
If ierror Then Exit Sub
End Sub

Private Sub CommandRunPENEPMA_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08SaveDisplay(FormPENEPMA08_PE)
If ierror Then Exit Sub
Call Penepma08RunPenepma(FormPENEPMA08_PE)
If ierror Then Exit Sub
End Sub

Private Sub CommandZoomFull_Click()
If Not DebugMode Then On Error Resume Next
Pesgo1.PEactions = UNDO_ZOOM&
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormPENEPMA08_PE)
HelpContextID = IOGetHelpContextID("FormPENEPMA08")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub ListAvailableStandards_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08ListStandard(Int(0), FormPENEPMA08_PE)
If ierror Then Exit Sub
End Sub

Private Sub ListAvailableStandards_DblClick()
If Not DebugMode Then On Error Resume Next
Call Penepma08ListStandard(Int(1), FormPENEPMA08_PE)
If ierror Then Exit Sub
End Sub

Private Sub OptionDisplayGraph_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call Penepma08GraphUpdate(Index%)
If ierror Then Exit Sub
End Sub

Private Sub OptionProduction_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call Penepma08LoadProduction(Index%, FormPENEPMA08_PE)
If ierror Then Exit Sub
ProductionIndex% = Index%

If Index% = 0 Then  ' optimize x-rays
FormPENEPMA08_PE.LabelMaterialFiles(0).Caption = "Bulk Beam Incident Material"
FormPENEPMA08_PE.LabelMaterialFiles(1).Caption = vbNullString
FormPENEPMA08_PE.TextMaterialFiles(1).Text = vbNullString

ElseIf Index% = 1 Then  ' optimize backscatter
FormPENEPMA08_PE.LabelMaterialFiles(0).Caption = "Bulk Beam Incident Material"
FormPENEPMA08_PE.LabelMaterialFiles(1).Caption = vbNullString
FormPENEPMA08_PE.TextMaterialFiles(1).Text = vbNullString

ElseIf Index% = 2 Then  ' optimize continuum
FormPENEPMA08_PE.LabelMaterialFiles(0).Caption = "Bulk Beam Incident Material"
FormPENEPMA08_PE.LabelMaterialFiles(1).Caption = vbNullString
FormPENEPMA08_PE.TextMaterialFiles(1).Text = vbNullString

ElseIf Index% = 3 Then  ' optimize couple or hemisphere
FormPENEPMA08_PE.LabelMaterialFiles(0).Caption = "Beam Incident Material (X>0)"
FormPENEPMA08_PE.LabelMaterialFiles(1).Caption = "Adjacent Phase or Matrix"

ElseIf Index% = 4 Then  ' optimize bilayer (thin film)
FormPENEPMA08_PE.LabelMaterialFiles(0).Caption = "Thin Film Material"
FormPENEPMA08_PE.LabelMaterialFiles(1).Caption = "Substrate Material"
End If

' Set production control enables
Call Penepma08SetOptionProductionEnables(Index%)
If ierror Then Exit Sub

End Sub

Private Sub Pesgo1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not DebugMode Then On Error Resume Next
Dim fX As Double, fY As Double      ' last mouse position

' Get mouse position in data units
Call MiscPlotTrack(Int(1), X!, Y!, fX#, fY#, FormPENEPMA08_PE.Pesgo1)
If ierror Then Exit Sub
   
' Format graph mouse position
If fX# <> 0# And fY# <> 0# Then
   FormPENEPMA08_PE.LabelXPos.Caption = MiscAutoFormat$(CSng(fX#))
   FormPENEPMA08_PE.LabelYPos.Caption = MiscAutoFormat$(CSng(fY#))
Else
   FormPENEPMA08_PE.LabelXPos.Caption = vbNullString
   FormPENEPMA08_PE.LabelYPos.Caption = vbNullString
End If
End Sub

Private Sub TextBeamDirection_GotFocus(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextEnergyRangeMinMaxNumber_GotFocus(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub UpDownXray_DownClick(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call Penepma08XrayAdjust(Int(-1), Index% + 1)
If ierror Then Exit Sub
Call Penepma08AdjustEABS(ProductionIndex%, Index%, FormPENEPMA08_PE)
If ierror Then Exit Sub
End Sub

Private Sub UpDownXray_UpClick(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call Penepma08XrayAdjust(Int(1), Index% + 1)
If ierror Then Exit Sub
Call Penepma08AdjustEABS(ProductionIndex%, Index%, FormPENEPMA08_PE)
If ierror Then Exit Sub
End Sub

Private Sub TextBeamAperture_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextBeamEnergy_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextBeamPosition_GotFocus(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextBeamTakeoff_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextDumpPeriod_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextEABS1_GotFocus(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextEABS2_GotFocus(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextGeometryFile_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextInputFile_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextInputTitle_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMaterialDensity_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMaterialFcb_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMaterialFiles_GotFocus(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMaterialWcb_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextNumberSimulatedShowers_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextSimulationTimePeriod_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub Timer1_Timer()
If Not DebugMode Then On Error Resume Next
Call Penepma08LoadPenepmaDAT(FormPENEPMA08_PE)
If ierror Then Exit Sub
Call Penepma08GraphUpdate(GraphDisplayOption%)
If ierror Then Exit Sub
Call Penepma08CheckTermination(FormPENEPMA08_PE)
If ierror Then Exit Sub
End Sub

