VERSION 5.00
Begin VB.Form FormMQOPTIONS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MQ (Monte-Carlo Input File Options)"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CommandSpecifyFolderForPureElementOutput 
      Caption         =   "Specify Folder For Pure Element MQ Output Files"
      Height          =   255
      Left            =   4320
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   5040
      Width           =   3735
   End
   Begin VB.CommandButton CommandExtractCompoundDataFromMQOutputFile 
      Caption         =   "Extract  Compound Data From MQ Output File"
      Height          =   255
      Left            =   4320
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5400
      Width           =   3735
   End
   Begin VB.CommandButton CommandCreateMQInputFilesForAllBinaries 
      Caption         =   "Create MQ Input Files for All Binaries (1-94)"
      Height          =   375
      Left            =   240
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5280
      Width           =   3735
   End
   Begin VB.Frame Frame3 
      Caption         =   "Conditions (change from Analytical menu)"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   27
      Top             =   1080
      Width           =   3975
      Begin VB.Label LabelKiloVolts 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2520
         TabIndex        =   31
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Operating Voltage (KeV)"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label LabelTakeOff 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2520
         TabIndex        =   29
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Take Off Angle (Degrees)"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.CommandButton CommandCreateMQInputFilesForAllElements 
      BackColor       =   &H0080FFFF&
      Caption         =   "Create MQ Input Files For All Elements (1-94)"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4800
      Width           =   3735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Output Parameters"
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
      Height          =   3855
      Left            =   4200
      TabIndex        =   15
      Top             =   1080
      Width           =   3975
      Begin VB.TextBox TextSecondaryEnergy 
         Height          =   285
         Left            =   2640
         TabIndex        =   10
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox TextHistogramRange 
         Height          =   285
         Left            =   2640
         TabIndex        =   9
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox TextNumberofTrajectories 
         Height          =   285
         Left            =   2640
         TabIndex        =   8
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox TextSubstrateXrayLine 
         Height          =   285
         Left            =   2640
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox TextSubstrateAtomicNumber 
         Height          =   285
         Left            =   2640
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox TextSubstrateDensity 
         Height          =   285
         Left            =   2640
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox TextSubstrateThickness 
         Height          =   285
         Left            =   2640
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox TextFilmDensity 
         Height          =   285
         Left            =   2640
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TextFilmThickness 
         Height          =   285
         Left            =   2640
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Minimum Secondary Energy (KeV)"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label9 
         Caption         =   "Histogram Range (microns)"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label8 
         Caption         =   "Number of Trajectories"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "Substrate X-ray Line (K, L, M)"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Substrate Atomic Number"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Substrate Density (gm/cm3)"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Substrate Thickness (microns)"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Film Density (gm/cm3)"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Film Thickness (microns)"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   2415
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
      Left            =   7200
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   720
      Width           =   975
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
      Height          =   495
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Output Options"
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
      Height          =   1815
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   3975
      Begin VB.CheckBox CheckMassZedDiffMax 
         Caption         =   "Use Maximum Mass - Electron Zbar Difference"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3615
      End
      Begin VB.CheckBox CheckMassZedDiffMin 
         Caption         =   "Use Minimum Mass - Electron Zbar Difference"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox TextMassZedDiffMax 
         Height          =   285
         Left            =   2640
         TabIndex        =   1
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox TextMassZedDiffMin 
         Height          =   285
         Left            =   2640
         TabIndex        =   0
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Maximum Percent Difference"
         Height          =   255
         Left            =   480
         TabIndex        =   34
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Minimum Percent Difference"
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   $"MQOptions.frx":0000
      Height          =   855
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "FormMQOPTIONS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
ierror = True
Unload FormMQOPTIONS
End Sub

Private Sub CommandCreateMQInputFilesForAllBinaries_Click()
If Not DebugMode Then On Error Resume Next
Call MqOptionsSave
If ierror Then Exit Sub
Call StanFormCalculateBinary
If ierror Then Exit Sub
End Sub

Private Sub CommandCreateMQInputFilesForAllElements_Click()
If Not DebugMode Then On Error Resume Next
Call MqOptionsSave
If ierror Then Exit Sub
Screen.MousePointer = vbHourglass
Call MQOptionsCalculateAll
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End Sub

Private Sub CommandExtractCompoundDataFromMQOutputFile_Click()
If Not DebugMode Then On Error Resume Next
Call MQOptionsExtractData(FormMQOPTIONS)
If ierror Then Exit Sub
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call MqOptionsSave
If ierror Then Exit Sub
Unload FormMQOPTIONS
End Sub

Private Sub CommandSpecifyFolderForPureElementOutput_Click()
If Not DebugMode Then On Error Resume Next
Call MQOptionsExtract(FormMQOPTIONS)
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormMQOPTIONS)
HelpContextID = IOGetHelpContextID("FormMQOPTIONS")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub TextFilmDensity_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextFilmThickness_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextHistogramRange_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMassZedDiffMax_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMassZedDiffMin_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextNumberofTrajectories_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextSecondaryEnergy_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextSubstrateAtomicNumber_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextSubstrateDensity_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextSubstrateThickness_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextSubstrateXrayLine_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub
