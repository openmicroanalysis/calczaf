VERSION 5.00
Object = "{6E5043E8-C452-4A6A-B011-9B5687112610}#1.0#0"; "Pesgo32f.ocx"
Begin VB.Form FormPENEPMA12 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculate Penepma 2012 Fluorescence Couple Profiles"
   ClientHeight    =   12495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   12495
   ScaleWidth      =   17160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CommandHelp 
      BackColor       =   &H00FF8080&
      Caption         =   "Help"
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   99
      ToolTipText     =   "Click this button to get detailed help from our on-line user forum"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton CommandBinary 
      Caption         =   "Binary Calculations"
      Height          =   495
      Left            =   13680
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Fanal Prompt"
      Height          =   495
      Left            =   12600
      TabIndex        =   77
      TabStop         =   0   'False
      ToolTipText     =   "Opens a command prompt in the Penepma12 Fanal folder (e.g., type fanal.exe < fanal.in)"
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Penfluor Prompt"
      Height          =   495
      Left            =   11520
      TabIndex        =   76
      TabStop         =   0   'False
      ToolTipText     =   "Opens a command prompt in the Penepma12\Penfluor folder (e.g., type penfluor.exe or fitall.exe)"
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "PENEPMA Prompt"
      Height          =   255
      Left            =   9600
      TabIndex        =   75
      TabStop         =   0   'False
      ToolTipText     =   "Opens a command prompt in the Penepma12 Penepma folder"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command10 
      Caption         =   "PENDBASE Prompt"
      Height          =   255
      Left            =   9600
      TabIndex        =   74
      TabStop         =   0   'False
      ToolTipText     =   "Opens a command prompt in the Penepma12 Pendbase folder (e.g., type material.exe < material1.inp)"
      Top             =   960
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      Caption         =   "Calculate Secondary Fluorescence Profiles For the Specified Element and X-ray"
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
      Height          =   10695
      Left            =   9480
      TabIndex        =   56
      Top             =   1680
      Width           =   7575
      Begin Pesgo32fLib.Pesgo Pesgo1 
         Height          =   7695
         Left            =   120
         TabIndex        =   103
         Top             =   2880
         Width           =   7335
         _Version        =   65536
         _ExtentX        =   12938
         _ExtentY        =   13573
         _StockProps     =   96
         _AllProps       =   "PENEPMA12_PE.frx":0000
      End
      Begin VB.CheckBox CheckUseGridLines 
         Caption         =   "Use Grid Lines"
         Height          =   255
         Left            =   4680
         TabIndex        =   97
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox CheckUseLogScale 
         Caption         =   "Use Log Scale"
         Height          =   255
         Left            =   3120
         TabIndex        =   96
         ToolTipText     =   "Display intensities on log scale"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton CommandCopy 
         Caption         =   "Copy To Clipboard"
         Height          =   495
         Left            =   6480
         TabIndex        =   94
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to copy the current plot to the clipboard"
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton CommandZoomFull 
         Caption         =   "Zoom Full"
         Height          =   375
         Left            =   4560
         TabIndex        =   79
         TabStop         =   0   'False
         ToolTipText     =   "Zoom to full graph extents"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CheckBox CheckSendToExcel 
         Caption         =   "Send To Excel"
         Height          =   255
         Left            =   3120
         TabIndex        =   78
         TabStop         =   0   'False
         ToolTipText     =   "Check this box to send the modified output with ""apparent"" concentrations and matrix correction factors to Excel"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton CommandRunFanal 
         BackColor       =   &H0000FFFF&
         Caption         =   "Run Fanal (generate k-ratio file for couple boundary)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   65
         TabStop         =   0   'False
         ToolTipText     =   $"PENEPMA12_PE.frx":6104
         Top             =   2040
         Width           =   2775
      End
      Begin VB.ComboBox ComboElementStd 
         Height          =   315
         Left            =   960
         TabIndex        =   64
         TabStop         =   0   'False
         ToolTipText     =   "Select the measured element to profile fluorescence"
         Top             =   1680
         Width           =   735
      End
      Begin VB.ComboBox ComboXRayStd 
         Height          =   315
         Left            =   2520
         TabIndex        =   63
         TabStop         =   0   'False
         ToolTipText     =   "Select the measured x-ray to profile fluorescence"
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton CommandBrowseParBStd 
         BackColor       =   &H0080FFFF&
         Caption         =   "Browse"
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
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   62
         TabStop         =   0   'False
         ToolTipText     =   "Browse material B Std parameter file (for matrix calculations please select a pure element standard)"
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton CommandBrowseParB 
         BackColor       =   &H0080FFFF&
         Caption         =   "Browse"
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
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   61
         TabStop         =   0   'False
         ToolTipText     =   "Browse material B parameter file"
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton CommandBrowseParA 
         BackColor       =   &H0080FFFF&
         Caption         =   "Browse"
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
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   60
         TabStop         =   0   'False
         ToolTipText     =   "Browse material A parameter file"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TextParameterFileBStd 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   59
         TabStop         =   0   'False
         ToolTipText     =   $"PENEPMA12_PE.frx":6197
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox TextParameterFileB 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   58
         TabStop         =   0   'False
         ToolTipText     =   "Specify the boundary material"
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox TextParameterFileA 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "Specify the beam incident material (all electrons come to rest within this material)"
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox TextMeasuredMicrons 
         Height          =   285
         Left            =   4080
         TabIndex        =   11
         ToolTipText     =   $"PENEPMA12_PE.frx":6225
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox TextMeasuredPoints 
         Height          =   285
         Left            =   5400
         TabIndex        =   12
         ToolTipText     =   "Enter number of points to calculate for fluorescence profile"
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label LabelXPos 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6480
         TabIndex        =   82
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label LabelYPos 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6480
         TabIndex        =   81
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Element"
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
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "X-Ray"
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
         Height          =   255
         Left            =   1800
         TabIndex        =   72
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label36 
         Caption         =   "Material B Std (primary std)"
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
         TabIndex        =   71
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label29 
         Caption         =   "Material B (boundary)"
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
         TabIndex        =   70
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label30 
         Caption         =   "Material A (beam incident)"
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
         TabIndex        =   69
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Caption         =   "Parameter Files are created in PENEPMA_Root\Penfluor"
         Height          =   255
         Left            =   1080
         TabIndex        =   68
         Top             =   330
         Width           =   4695
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Caption         =   "Microns"
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
         Left            =   3360
         TabIndex        =   67
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Points"
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
         Left            =   4800
         TabIndex        =   66
         Top             =   1680
         Width           =   615
      End
   End
   Begin VB.TextBox TextBeamTakeoff 
      Height          =   285
      Left            =   12120
      TabIndex        =   0
      ToolTipText     =   "Enter beam takeoff angle in degrees (40 for JEOL and Cameca)"
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox TextBeamEnergy 
      Height          =   285
      Left            =   14280
      TabIndex        =   1
      ToolTipText     =   "Enter beam energy in electron volts (eV)"
      Top             =   360
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Primary Intensity Calculations (create .PAR files for one or all .MAT files) (~10 hours each at 3600 sec)"
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
      Height          =   3135
      Left            =   120
      TabIndex        =   34
      Top             =   7680
      Width           =   9135
      Begin VB.CheckBox CheckAutoAdjustMinimumEnergy 
         Caption         =   "Automatically Adjust Electron/Photon Minimum Energy for Z < 11 (Na)"
         Height          =   255
         Left            =   240
         TabIndex        =   101
         ToolTipText     =   "Automatically lower electron/photon minimum energy for materials containing elements with Z < 11 (Na)"
         Top             =   2520
         Value           =   1  'Checked
         Width           =   5895
      End
      Begin VB.TextBox TextPenepmaMinimumElectronEnergy 
         Height          =   285
         Left            =   3120
         TabIndex        =   100
         ToolTipText     =   "Specify the Penepma minimum electron energy that was utilized for the Penfluor binary calculations (usually 1.0)"
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox TextSimulationShowers 
         Height          =   285
         Left            =   4440
         TabIndex        =   91
         ToolTipText     =   "Enter number of simulation trajectories per each of 10 voltages (1 to 50 keV) "
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox TextSimulationTime 
         Height          =   285
         Left            =   4440
         TabIndex        =   89
         ToolTipText     =   "Enter simulation time per each of 10 voltages (1 to 50 keV)  in seconds"
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton CommandRunPenfluorBStd 
         BackColor       =   &H0080FFFF&
         Caption         =   "Run Penfluor/Fitall for Material B Std Only"
         Height          =   495
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   53
         TabStop         =   0   'False
         ToolTipText     =   "Run Penfluor and Fitall on material B Std to create new .par file for secondary fluorescence calculations"
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CommandButton CommandRunPenfluorB 
         BackColor       =   &H0080FFFF&
         Caption         =   "Run Penfluor/Fitall for Material B Only"
         Height          =   495
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   52
         TabStop         =   0   'False
         ToolTipText     =   "Run Penfluor and Fitall on material B to create new .par file for secondary fluorescence calculations"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CommandButton CommandRunPenfluorA 
         BackColor       =   &H0080FFFF&
         Caption         =   "Run Penfluor/Fitall for Material A Only"
         Height          =   495
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   51
         TabStop         =   0   'False
         ToolTipText     =   "Run Penfluor and Fitall on material A to create new .par file for secondary fluorescence calculations"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton CommandRunPenfluorRunAll 
         BackColor       =   &H0000FFFF&
         Caption         =   "Run Penfluor and Fitall for ALL three materials (generate .PAR files)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "Run Penfluor and Fitall on all three specified materials to create new .par files for secondary fluorescence calculations"
         Top             =   360
         Width           =   2535
      End
      Begin VB.CommandButton CommandBrowseMatBStd 
         Caption         =   "Browse"
         Height          =   255
         Left            =   5400
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "Browse material B Std material file"
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton CommandBrowseMatB 
         Caption         =   "Browse"
         Height          =   255
         Left            =   5400
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "Browse material B material file"
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton CommandBrowseMatA 
         Caption         =   "Browse"
         Height          =   255
         Left            =   5400
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Browse material A material file"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TextMaterialFileBStd 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Specify the primary standard for Penfluor calculations"
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox TextMaterialFileB 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "Specify the boundary material for Penfluor calculations"
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox TextMaterialFileA 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Specify the beam incident material for Penfluor calculations"
         Top             =   600
         Width           =   4215
      End
      Begin VB.Label Label20 
         Caption         =   "Minimum Electron Energy (keV)"
         Height          =   255
         Left            =   720
         TabIndex        =   102
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label Label31 
         Caption         =   "Number of Simulations (in electrons, x10)"
         Height          =   255
         Left            =   1080
         TabIndex        =   92
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label22 
         Caption         =   "Time (in seconds per simulation, x10)"
         Height          =   255
         Left            =   1080
         TabIndex        =   90
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Create Material Files in PENEPMA_Path"
         Height          =   255
         Left            =   1080
         TabIndex        =   45
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label18 
         Caption         =   "Mat. B Std"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Mat. B"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Mat. A"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Timer Timer1 
      Left            =   9000
      Top             =   120
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
      Left            =   15480
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H00008000&
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
      Height          =   615
      Left            =   15480
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "PENPEMA Material Files (create .MAT files)"
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
      Height          =   7335
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   9135
      Begin VB.ListBox ListAtomicDensitiesBStd 
         Height          =   1425
         Left            =   6120
         TabIndex        =   84
         TabStop         =   0   'False
         ToolTipText     =   "Double click to load the selected density to the electron range density field"
         Top             =   5760
         Width           =   2895
      End
      Begin VB.ListBox ListAtomicDensitiesB 
         Height          =   1425
         Left            =   3120
         TabIndex        =   83
         TabStop         =   0   'False
         ToolTipText     =   "Double click to load the selected density to the electron range density field"
         Top             =   5760
         Width           =   2895
      End
      Begin VB.ListBox ListAtomicDensitiesA 
         Height          =   1425
         Left            =   120
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "Double click to load the selected density to the electron range density field"
         Top             =   5760
         Width           =   2895
      End
      Begin VB.CommandButton CommandOutputMaterialBStd 
         BackColor       =   &H0080FFFF&
         Caption         =   "Create PENEPMA Mat. B Std From List"
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
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   $"PENEPMA12_PE.frx":62C6
         Top             =   4440
         Width           =   2895
      End
      Begin VB.ListBox ListAvailableStandardsBStd 
         Height          =   2400
         Left            =   6120
         Sorted          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   "Standard compositions available for output for Material B Std (for matrix calculation please select a pure element standard)"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CommandButton CommandOutputFormulaBStd 
         BackColor       =   &H0000FFFF&
         Caption         =   "Create PENEPMA Material B Std From Formula"
         Height          =   495
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   $"PENEPMA12_PE.frx":635B
         Top             =   5160
         Width           =   2895
      End
      Begin VB.TextBox TextMaterialWcbBStd 
         Height          =   285
         Left            =   8040
         TabIndex        =   10
         ToolTipText     =   "Enter the oscillator energy (0 = use default)"
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox TextMaterialFcbBStd 
         Height          =   285
         Left            =   7080
         TabIndex        =   9
         ToolTipText     =   "Enter the oscillator strength  (0 = use default)"
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox TextMaterialDensityBStd 
         Height          =   285
         Left            =   6120
         TabIndex        =   8
         ToolTipText     =   "Enter the material density in gm/cm^3"
         Top             =   3840
         Width           =   975
      End
      Begin VB.ListBox ListAvailableStandardsA 
         Height          =   2400
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Standard compositions available for output for Material A"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CommandButton CommandOutputMaterialB 
         BackColor       =   &H0080FFFF&
         Caption         =   "Create PENEPMA Material B From List"
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   $"PENEPMA12_PE.frx":63F8
         Top             =   4440
         Width           =   2895
      End
      Begin VB.ListBox ListAvailableStandardsB 
         Height          =   2400
         Left            =   3120
         Sorted          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Standard compositions available for output for Material B"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox TextMaterialDensityB 
         Height          =   285
         Left            =   3120
         TabIndex        =   5
         ToolTipText     =   "Enter the material density in gm/cm^3"
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox TextMaterialFcbB 
         Height          =   285
         Left            =   4080
         TabIndex        =   6
         ToolTipText     =   "Enter the oscillator strength  (0 = use default)"
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox TextMaterialWcbB 
         Height          =   285
         Left            =   5040
         TabIndex        =   7
         ToolTipText     =   "Enter the oscillator energy (0 = use default)"
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton CommandOutputFormulaB 
         BackColor       =   &H0000FFFF&
         Caption         =   "Create PENEPMA Material B From Formula"
         Height          =   495
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   $"PENEPMA12_PE.frx":6489
         Top             =   5160
         Width           =   2895
      End
      Begin VB.CommandButton CommandOutputMaterialA 
         BackColor       =   &H0080FFFF&
         Caption         =   "Create PENEPMA Material A From List"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   $"PENEPMA12_PE.frx":6540
         Top             =   4440
         Width           =   2895
      End
      Begin VB.TextBox TextMaterialDensityA 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Enter the material density in gm/cm^3"
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox TextMaterialFcbA 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         ToolTipText     =   "Enter the oscillator strength  (0 = use default)"
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox TextMaterialWcbA 
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         ToolTipText     =   "Enter the oscillator energy (0 = use default)"
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton CommandOutputFormulaA 
         BackColor       =   &H0000FFFF&
         Caption         =   "Create PENEPMA Material A From Formula"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   $"PENEPMA12_PE.frx":65D1
         Top             =   5160
         Width           =   2895
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "(boundary material)"
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
         Left            =   3120
         TabIndex        =   98
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "(primary standard)"
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
         Left            =   6600
         TabIndex        =   88
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "(beam incident material)"
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
         TabIndex        =   87
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Osc. Strength"
         Height          =   255
         Left            =   7080
         TabIndex        =   32
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "(must contain the measured element)"
         Height          =   255
         Left            =   6120
         TabIndex        =   54
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "Select Material B Std"
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
         Left            =   6120
         TabIndex        =   48
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Osc. Energy"
         Height          =   255
         Left            =   8040
         TabIndex        =   33
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Density"
         Height          =   255
         Left            =   6120
         TabIndex        =   31
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Select Material B"
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
         Left            =   3240
         TabIndex        =   29
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Density"
         Height          =   255
         Left            =   3120
         TabIndex        =   28
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Osc. Strength"
         Height          =   255
         Left            =   4080
         TabIndex        =   27
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Osc. Energy"
         Height          =   255
         Left            =   5040
         TabIndex        =   26
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Select Material A"
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
         Left            =   600
         TabIndex        =   20
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label LabelMaterialDensity 
         Alignment       =   2  'Center
         Caption         =   "Density"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Osc. Strength"
         Height          =   255
         Left            =   1080
         TabIndex        =   18
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Osc. Energy"
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Oscillator strength (Fcb) and Oscillator energy (Wcb) of the plasmon should be zero for insulators"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   4200
         Width           =   8655
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   9000
         Y1              =   5040
         Y2              =   5040
      End
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      Caption         =   "Kilovolts"
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
      Left            =   13440
      TabIndex        =   93
      Top             =   360
      Width           =   855
   End
   Begin VB.Label LabelProgress 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1080
      TabIndex        =   85
      Top             =   10920
      Width           =   7215
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Caption         =   $"PENEPMA12_PE.frx":6688
      Height          =   855
      Left            =   240
      TabIndex        =   80
      Top             =   11520
      Width           =   9015
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Take-off"
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
      Left            =   11280
      TabIndex        =   50
      Top             =   360
      Width           =   855
   End
   Begin VB.Label LabelRemainingTime 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2520
      TabIndex        =   86
      Top             =   11160
      Width           =   4335
   End
End
Attribute VB_Name = "FormPENEPMA12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

Dim ShiftDown As Integer, CtrlDown As Integer, AltDown As Integer

Private Sub CheckUseGridlines_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12PlotGrid
If ierror Then Exit Sub
Call Penepma12PlotUpdate_PE(Int(1), FormPENEPMA12)
If ierror Then Exit Sub
End Sub

Private Sub CheckUseLogScale_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12PlotLog
If ierror Then Exit Sub
Call Penepma12PlotUpdate_PE(Int(1), FormPENEPMA12)
If ierror Then Exit Sub
End Sub

Private Sub CommandHelp_Click()
If Not DebugMode Then On Error Resume Next
Call IOBrowseHTTP(ProbeSoftwareInternetBrowseMethod%, "http://probesoftware.com/smf/index.php?topic=58.msg214#msg214")
If ierror Then Exit Sub
End Sub

Private Sub CommandRunFanal_Click()
If Not DebugMode Then On Error Resume Next
Screen.MousePointer = vbHourglass
Call Penepma12Save
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
Screen.MousePointer = vbHourglass
Call Penepma12RunFanal
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
Screen.MousePointer = vbHourglass
Call Penepma12RunFanalOutput(FormPENEPMA12)
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Not DebugMode Then On Error Resume Next
ShiftDown = (Shift And vbShiftMask) > 0
CtrlDown = (Shift And vbCtrlMask) > 0
AltDown = (Shift And vbAltMask) > 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If Not DebugMode Then On Error Resume Next
ShiftDown = 0
CtrlDown = 0
AltDown = 0
End Sub

Private Sub ComboElementStd_Change()
If Not DebugMode Then On Error Resume Next
Call Penepma12UpdateCombo
If ierror Then Exit Sub
End Sub

Private Sub ComboElementStd_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12UpdateCombo
If ierror Then Exit Sub
End Sub

Private Sub Command10_Click()
If Not DebugMode Then On Error Resume Next
Dim taskID As Long
ChDrive PENDBASE_Path$
taskID& = Shell("cmd.exe /k cd " & VbDquote$ & PENDBASE_Path$ & VbDquote$, vbNormalFocus)
End Sub

Private Sub Command11_Click()
If Not DebugMode Then On Error Resume Next
Dim taskID As Long
ChDrive PENEPMA_Path$
taskID& = Shell("cmd.exe /k cd " & VbDquote$ & PENEPMA_Path$ & VbDquote$, vbNormalFocus)
End Sub

Private Sub Command12_Click()
If Not DebugMode Then On Error Resume Next
Dim taskID As Long
ChDrive PENEPMA_Root$
taskID& = Shell("cmd.exe /k cd " & VbDquote$ & PENEPMA_Root$ & "\Penfluor" & VbDquote$, vbNormalFocus)
End Sub

Private Sub Command13_Click()
If Not DebugMode Then On Error Resume Next
Dim taskID As Long
ChDrive PENEPMA_Root$
taskID& = Shell("cmd.exe /k cd " & VbDquote$ & PENEPMA_Root$ & "\Fanal" & VbDquote$, vbNormalFocus)
End Sub

Private Sub CommandBinary_Click()
If Not DebugMode Then On Error Resume Next
' Check if secret keyboard combination is present. If so load calculate binary form
If ShiftDown And CtrlDown And AltDown Then
FormPenepma12Binary.Show vbModeless
ShiftDown = 0   ' to prevent next click from causing it to automatically reload
CtrlDown = 0
AltDown = 0
Else
msg$ = "Please contact Probe Software, to access the Penepma binary calculations area"
MsgBox msg$, vbOKOnly + vbInformation, "FormPenepma12"
End If
End Sub

Private Sub CommandBrowseMatA_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12BrowseMaterialFile(Int(1), FormPENEPMA12)
If ierror Then Exit Sub
End Sub

Private Sub CommandBrowseMatB_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12BrowseMaterialFile(Int(2), FormPENEPMA12)
If ierror Then Exit Sub
End Sub

Private Sub CommandBrowseMatBStd_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12BrowseMaterialFile(Int(3), FormPENEPMA12)
If ierror Then Exit Sub
End Sub

Private Sub CommandBrowseParA_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12BrowseParameterFile(Int(1), FormPENEPMA12)
If ierror Then Exit Sub
Call Penepma12PlotLoad_PE
If ierror Then Exit Sub
End Sub

Private Sub CommandBrowseParB_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12BrowseParameterFile(Int(2), FormPENEPMA12)
If ierror Then Exit Sub
Call Penepma12PlotLoad_PE
If ierror Then Exit Sub
End Sub

Private Sub CommandBrowseParBStd_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12BrowseParameterFile(Int(3), FormPENEPMA12)
If ierror Then Exit Sub
Call Penepma12PlotLoad_PE
If ierror Then Exit Sub
End Sub

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormPENEPMA12
End Sub

Private Sub CommandCopy_Click()
If Not DebugMode Then On Error Resume Next
FormPENEPMA12.Pesgo1.AllowExporting = True
'FormPENEPMA12.Pesgo1.ExportImageLargeFont = False
'FormPENEPMA12.Pesgo1.ExportImageDpi = 450
Call FormPENEPMA12.Pesgo1.PEcopybitmaptoclipboard(600, 600)
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
Unload FormPENEPMA12
End Sub

Private Sub CommandOutputFormulaA_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
Call Penepma12CreateMaterialFormula(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub CommandOutputFormulaB_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
Call Penepma12CreateMaterialFormula(Int(2))
If ierror Then Exit Sub
End Sub

Private Sub CommandOutputFormulaBStd_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
Call Penepma12CreateMaterialFormula(Int(3))
If ierror Then Exit Sub
End Sub

Private Sub CommandOutputMaterialA_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
Call Penepma12CreateMaterial(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub CommandOutputMaterialB_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
Call Penepma12CreateMaterial(Int(2))
If ierror Then Exit Sub
End Sub

Private Sub CommandOutputMaterialBStd_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
Call Penepma12CreateMaterial(Int(3))
If ierror Then Exit Sub
End Sub

Private Sub CommandRunPenfluorA_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
TotalNumberOfSimulations& = 1
CurrentSimulationsNumber& = 1
' Check with user
Call Penepma12RunPenfluorCheck(Int(1))
If ierror Then Exit Sub
Call Penepma12RunPenfluorCheck2(Int(1))
If ierror Then Exit Sub
' Run Penfluor and Fitall on material A
Call Penepma12RunPenFluor(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub CommandRunPenfluorB_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
TotalNumberOfSimulations& = 1
CurrentSimulationsNumber& = 1
' Check with user
Call Penepma12RunPenfluorCheck(Int(2))
If ierror Then Exit Sub
Call Penepma12RunPenfluorCheck2(Int(2))
If ierror Then Exit Sub
' Run Penfluor and Fitall on material B
Call Penepma12RunPenFluor(Int(2))
If ierror Then Exit Sub
End Sub

Private Sub CommandRunPenfluorBStd_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
TotalNumberOfSimulations& = 1
CurrentSimulationsNumber& = 1
' Check with user
Call Penepma12RunPenfluorCheck(Int(3))
If ierror Then Exit Sub
Call Penepma12RunPenfluorCheck2(Int(3))
If ierror Then Exit Sub
' Run Penfluor and Fitall on material B Std
Call Penepma12RunPenFluor(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub CommandRunPenfluorRunAll_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12Save
If ierror Then Exit Sub
TotalNumberOfSimulations& = 3
' Check with user
Call Penepma12RunPenfluorCheck(Int(0))
If ierror Then Exit Sub
Call Penepma12RunPenfluorCheck2(Int(0))
If ierror Then Exit Sub
' Run Penfluor and Fitall on all three materials
CurrentSimulationsNumber& = 1
Call Penepma12RunPenFluor(Int(1))
If ierror Then Exit Sub
CurrentSimulationsNumber& = 2
Call Penepma12RunPenFluor(Int(2))
If ierror Then Exit Sub
CurrentSimulationsNumber& = 3
Call Penepma12RunPenFluor(Int(3))
If ierror Then Exit Sub
End Sub

Private Sub CommandZoomFull_Click()
If Not DebugMode Then On Error Resume Next
Pesgo1.PEactions = UNDO_ZOOM&
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormPENEPMA12)
HelpContextID = IOGetHelpContextID("FormPENEPMA12")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub ListAtomicDensitiesA_DblClick()
If Not DebugMode Then On Error Resume Next
FormPENEPMA12.TextMaterialDensityA.Text = MiscAutoFormat$(AllAtomicDensities!(FormPENEPMA12.ListAtomicDensitiesA.ListIndex + 1))
End Sub

Private Sub ListAtomicDensitiesB_DblClick()
If Not DebugMode Then On Error Resume Next
FormPENEPMA12.TextMaterialDensityB.Text = MiscAutoFormat$(AllAtomicDensities!(FormPENEPMA12.ListAtomicDensitiesB.ListIndex + 1))
End Sub

Private Sub ListAtomicDensitiesBStd_DblClick()
If Not DebugMode Then On Error Resume Next
FormPENEPMA12.TextMaterialDensityBStd.Text = MiscAutoFormat$(AllAtomicDensities!(FormPENEPMA12.ListAtomicDensitiesBStd.ListIndex + 1))
End Sub

Private Sub ListAvailableStandardsA_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12GetStandard(Int(0), Int(1))
If ierror Then Exit Sub
End Sub

Private Sub ListAvailableStandardsA_DblClick()
If Not DebugMode Then On Error Resume Next
Call Penepma12GetStandard(Int(1), Int(1))
If ierror Then Exit Sub
End Sub

Private Sub ListAvailableStandardsB_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12GetStandard(Int(0), Int(2))
If ierror Then Exit Sub
End Sub

Private Sub ListAvailableStandardsB_DblClick()
If Not DebugMode Then On Error Resume Next
Call Penepma12GetStandard(Int(1), Int(2))
If ierror Then Exit Sub
End Sub

Private Sub ListAvailableStandardsBStd_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma12GetStandard(Int(0), Int(3))
If ierror Then Exit Sub
End Sub

Private Sub ListAvailableStandardsBStd_DblClick()
If Not DebugMode Then On Error Resume Next
Call Penepma12GetStandard(Int(1), Int(3))
If ierror Then Exit Sub
End Sub

Private Sub Pesgo1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Not DebugMode Then On Error Resume Next
Dim fX As Double, fY As Double      ' last mouse position

' Get mouse position in data units
Call ZoomTrack(Int(1), x!, Y!, fX#, fY#, FormPENEPMA12.Pesgo1)
If ierror Then Exit Sub
   
' Format graph mouse position
If fX# <> 0# And fY# <> 0# Then
   FormPENEPMA12.LabelXPos.Caption = MiscAutoFormat$(CSng(fX#))
   FormPENEPMA12.LabelYPos.Caption = MiscAutoFormat$(CSng(fY#))
Else
   FormPENEPMA12.LabelXPos.Caption = vbNullString
   FormPENEPMA12.LabelYPos.Caption = vbNullString
End If
End Sub

Private Sub Pesgo1_ZoomIn()
If Not DebugMode Then On Error Resume Next
CommandZoomFull.Enabled = True
End Sub

Private Sub Pesgo1_ZoomOut()
If Not DebugMode Then On Error Resume Next
CommandZoomFull.Enabled = False
End Sub

Private Sub TextBeamEnergy_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextBeamTakeoff_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMaterialDensityA_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMaterialDensityB_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMaterialDensityBStd_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMaterialFcbA_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMaterialFcbB_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMaterialFcbBStd_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMaterialWcbA_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMaterialWcbB_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMaterialWcbBStd_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMeasuredMicrons_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMeasuredPoints_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextPenepmaMinimumElectronEnergy_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextSimulationShowers_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextSimulationTime_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub Timer1_Timer()
If Not DebugMode Then On Error Resume Next
Call Penepma12CheckTermination
If ierror Then Exit Sub
End Sub
