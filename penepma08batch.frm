VERSION 5.00
Begin VB.Form FormPENEPMA08Batch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PENEPMA- Batch Mode"
   ClientHeight    =   13140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   13140
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Generate Bulk Pure Element Input Files Based On Specified Range"
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
      Height          =   1215
      Left            =   120
      TabIndex        =   33
      Top             =   11880
      Width           =   7215
      Begin VB.CommandButton CommandRename 
         Caption         =   "Copy and Rename PARs"
         Height          =   315
         Left            =   4920
         TabIndex        =   46
         ToolTipText     =   "Copy Penepma pure element files and rename to synthetic spectrum folder"
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton CommandCreatePure 
         Caption         =   "Create Bulk Pure Element Input Files For Penepma"
         Height          =   495
         Left            =   4920
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   360
         Width           =   2055
      End
      Begin VB.CheckBox CheckDoNotOverwriteExisting 
         Caption         =   "Do Not Overwrite Existing Input Files"
         Height          =   255
         Left            =   1800
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   360
         Width           =   3015
      End
      Begin VB.ComboBox ComboPureElement1 
         Height          =   315
         Left            =   960
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Select the first element for pure element calculations"
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox ComboPureElement2 
         Height          =   315
         Left            =   960
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Select the second element for pure element calculations"
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Electron beam energy, etc.  is based on value in main Penepma window!"
         Height          =   375
         Left            =   1920
         TabIndex        =   40
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Element1"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Element2"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.CheckBox CheckSortByDate 
      Caption         =   "Sort By Date"
      Height          =   195
      Left            =   6000
      TabIndex        =   32
      Top             =   4200
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Run Penepma batch Calculations For The Selected Input Files"
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
      Height          =   2175
      Left            =   120
      TabIndex        =   24
      Top             =   7920
      Width           =   7215
      Begin VB.TextBox TextBatchFolder 
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   41
         Top             =   600
         Width           =   4575
      End
      Begin VB.CommandButton CommandExtractKratios2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Extract K-ratios"
         Height          =   255
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   $"PENEPMA08Batch.frx":0000
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ComboBox ComboXray 
         Height          =   315
         Left            =   6480
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1440
         Width           =   615
      End
      Begin VB.ComboBox ComboElm 
         Height          =   315
         Left            =   5760
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton CommandRunBatch 
         BackColor       =   &H0080FFFF&
         Caption         =   "Run Selected Input Files In Batch Mode"
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
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Run the selected PENEPMA Input files in batch mode"
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton CommandBrowseBatchFolder 
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
         Height          =   735
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Browse to specify the output folder for all selected batch calculations"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Batch Project Folder For Storing Batch Results"
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
         TabIndex        =   28
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label LabelCurrentInputFile 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   5295
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Generate Penepma Binary Composition Input Files (11 binaries from 1 to 99%) "
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
      Height          =   1215
      Left            =   120
      TabIndex        =   17
      Top             =   10440
      Width           =   7215
      Begin VB.CommandButton CommandCreateBinaries 
         Caption         =   "Create Binary Composition Input Files For Penepma"
         Height          =   735
         Left            =   2280
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Create binary composition files from 1 to 99%"
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox ComboBinaryElement2 
         Height          =   315
         Left            =   960
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Select the second element for binary fluorescence calculations"
         Top             =   720
         Width           =   615
      End
      Begin VB.ComboBox ComboBinaryElement1 
         Height          =   315
         Left            =   960
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Select the first element for binary fluorescence calculations"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton CommandExtractKRatios 
         Caption         =   "Extract K-Ratios From Binary Compositions"
         Height          =   735
         Left            =   5040
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Extract the k-ratios from the folder containing the binary composition calculations"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label LabelToAnd 
         Alignment       =   2  'Center
         Caption         =   "Element2"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   735
      End
      Begin VB.Label LabelFrom 
         Alignment       =   2  'Center
         Caption         =   "Element1"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton CommandReload 
      Caption         =   "Reload List"
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Frame FrameInputFileParameters 
      Caption         =   "Selected File Properties"
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
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Width           =   7215
      Begin VB.TextBox TextEnergyRangeMinMaxNumber 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   5880
         TabIndex        =   44
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox TextEnergyRangeMinMaxNumber 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   4440
         TabIndex        =   43
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox TextEnergyRangeMinMaxNumber 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   42
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox TextEABS1 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Electron absorption energy for this material"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox TextEABS2 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Photon absorption energy for this material"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox TextDumpPeriod 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Enter the time interval for the dump files to be updated (for live display)"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox TextNumberSimulatedShowers 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Enter the number of simulated showers (incident electrons)"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox TextSimulationTimePeriod 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Enter the total simulation time period in seconds"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox TextBeamEnergy 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Enter beam energy in electron volts (eV)"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox TextInputTitle 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Enter title for this input file (up to 120 characters)"
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label Label14 
         Caption         =   "Energy Range (min, max, num)"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Electron/Photon Absorption (eV)"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label11 
         Caption         =   "Number of Showers, Simulation Time"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label13 
         Caption         =   "Beam Energy (eV), Dump Period (in sec)"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label16 
         Caption         =   "Input File Title"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H00C0FFC0&
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Close the PENEPMA batch mode window"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "PENEPMA Batch Mode Processing"
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
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.ListBox ListInputFiles 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4350
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         ToolTipText     =   "Select the PENEPMA Input files to run in batch mode"
         Top             =   360
         Width           =   5535
      End
   End
   Begin VB.Label LabelPENEPMA08Batch 
      Alignment       =   2  'Center
      Caption         =   $"PENEPMA08Batch.frx":0096
      Height          =   3015
      Left            =   6000
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "FormPENEPMA08Batch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2017 by John J. Donovan
Option Explicit

Private Sub CommandBrowseBatchFolder_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08BatchBrowseFolder
If ierror Then Exit Sub
End Sub

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormPENEPMA08Batch
End Sub

Private Sub CommandCreateBinaries_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08SaveInput(FormPENEPMA08_PE)
If ierror Then Exit Sub
Call Penepma08BatchBinaryCreate(FormPENEPMA08_PE)
If ierror Then Exit Sub
Call Penepma08BatchLoad
If ierror Then Exit Sub
End Sub

Private Sub CommandCreatePure_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08SaveInput(FormPENEPMA08_PE)
If ierror Then Exit Sub
Call Penepma08BatchBulkPureElementCreate(FormPENEPMA08_PE)
If ierror Then Exit Sub
Call Penepma08BatchLoad
If ierror Then Exit Sub
End Sub

Private Sub CommandExtractKRatios_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08BatchBinaryExtract
If ierror Then Exit Sub
End Sub

Private Sub CommandExtractKratios2_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08BatchExtractKratios(FormPENEPMA08_PE)
If ierror Then Exit Sub
End Sub

Private Sub CommandReload_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08BatchLoad
If ierror Then Exit Sub
End Sub

Private Sub CommandRename_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08BatchCopyRename
If ierror Then Exit Sub
End Sub

Private Sub CommandRunBatch_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08RunPenepmaBatch(FormPENEPMA08_PE)
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormPENEPMA08Batch)
HelpContextID = IOGetHelpContextID("FormPENEPMA08Batch")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub ListInputFiles_Click()
If Not DebugMode Then On Error Resume Next
Call Penepma08BatchGetInputParameters(FormPENEPMA08Batch.ListInputFiles.ListIndex)
If ierror Then Exit Sub
End Sub

Private Sub TextBatchFolder_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

