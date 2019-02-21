VERSION 5.00
Begin VB.Form FormRANGE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculate Electron and Xray Ranges"
   ClientHeight    =   11385
   ClientLeft      =   1245
   ClientTop       =   5580
   ClientWidth     =   10110
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
   ScaleHeight     =   11385
   ScaleWidth      =   10110
   Begin VB.CommandButton CommandHelpRange 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Help"
      Height          =   615
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Click this button to get detailed help from our on-line user forum"
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame6 
      Caption         =   "Electron Energy Loss (for low overvoltage situations)"
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   120
      TabIndex        =   35
      Top             =   9600
      Width           =   5535
      Begin VB.CommandButton CommandCalculateElectronEnergyTransmitted 
         Caption         =   "Calculate Electron Energy Transmitted"
         Height          =   495
         Left            =   1800
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Calculate the x-ray transmission of the specified x-ray energy in the current matrix at the current thickness"
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label LabelElectronEnergyFinal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   240
         TabIndex        =   37
         Top             =   960
         Width           =   5055
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Uses density, thickness and composition from above fields"
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
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.ListBox ListAtomicDensities 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   6000
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Double click to load the selected density to the electron range density field"
      Top             =   960
      Width           =   3735
   End
   Begin VB.Frame Frame5 
      Caption         =   "Xray Transmission (at an arbitrary energy)"
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   120
      TabIndex        =   29
      Top             =   7680
      Width           =   5535
      Begin VB.CommandButton CommandCalculateXrayTransmission2 
         Caption         =   "Calculate X-Ray Transmission"
         Height          =   495
         Left            =   1800
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Calculate the x-ray transmission of the specified x-ray energy in the current matrix at the current thickness"
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox TextXrayEnergy 
         Height          =   285
         Left            =   120
         TabIndex        =   30
         ToolTipText     =   "Enter the x-ray energy (in keV)"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label LabelXrayTransmission2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   240
         TabIndex        =   33
         Top             =   1080
         Width           =   5055
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Energy (keV)"
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
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Xray Transmission (at electron range specified above)"
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   120
      TabIndex        =   18
      Top             =   5760
      Width           =   5535
      Begin VB.TextBox TextThickness 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton CommandCalculateXrayTransmission 
         Caption         =   "Calculate X-Ray Transmission"
         Height          =   495
         Left            =   1800
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Calculate the x-ray transmission of a given emitter and x-ray in the current matrix at the specified thickness"
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label LabelXrayTransmission 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   5055
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Thickness (um)"
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
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Matrix or Film Composition"
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton CommandEnterCompositionAsStandard 
         Caption         =   "Enter Composition as Standard"
         Height          =   255
         Left            =   600
         TabIndex        =   28
         Top             =   840
         Width           =   4095
      End
      Begin VB.CommandButton CommandEnterCompositionAsWeightString 
         Caption         =   "Enter Composition as Weight String"
         Height          =   255
         Left            =   600
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   600
         Width           =   4095
      End
      Begin VB.CommandButton CommandEnterCompositionAsFormulaString 
         BackColor       =   &H0080FFFF&
         Caption         =   "Enter Composition as Formula String"
         Height          =   255
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label LabelComposition 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   4935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Electron Depth Range Radius"
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   5535
      Begin VB.CommandButton CommandCalculateElectronRange 
         BackColor       =   &H0080FFFF&
         Caption         =   "Calculate Electron Range"
         Height          =   495
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Calculate the ultimate electron range for the current matrix and density at a given electron energy"
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox TextKev 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         ToolTipText     =   "Enter the electron beam primary energy"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TextDensity 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Enter the matrix density (in gm/cm^2)"
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label LabelElectronRange 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   5055
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Electron keV"
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
         Left            =   1440
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Xray Production Depth Range Radius"
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   5535
      Begin VB.CommandButton CommandCalculateXrayRange 
         BackColor       =   &H0080FFFF&
         Caption         =   "Calculate X-Ray Range"
         Height          =   495
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Calculate the x-ray emission range for the given emitter in the current matrix and density"
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox ComboElement 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Text            =   "ComboElement"
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox ComboXray 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Text            =   "ComboXray"
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label LabelXrayRange 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Width           =   5055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "X-Ray"
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
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   615
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1335
   End
   Begin VB.OLE OLE4 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   1125
      Left            =   5880
      OleObjectBlob   =   "Range.frx":0000
      SizeMode        =   3  'Zoom
      TabIndex        =   38
      Top             =   9930
      Width           =   4005
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   $"Range.frx":1218
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6000
      TabIndex        =   27
      Top             =   5520
      Width           =   3495
   End
   Begin VB.OLE OLE3 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   1335
      Left            =   6960
      OleObjectBlob   =   "Range.frx":12F6
      SizeMode        =   3  'Zoom
      TabIndex        =   26
      Top             =   7080
      Width           =   1815
   End
   Begin VB.OLE OLE1 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   975
      Left            =   6000
      OleObjectBlob   =   "Range.frx":230E
      TabIndex        =   25
      Top             =   4170
      Width           =   3735
   End
   Begin VB.OLE OLE2 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   975
      Left            =   6360
      OleObjectBlob   =   "Range.frx":3526
      TabIndex        =   24
      Top             =   2490
      Width           =   3015
   End
End
Attribute VB_Name = "FormRANGE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2019 by John J. Donovan
Option Explicit

Private Sub CommandCalculateElectronEnergyTransmitted_Click()
If Not DebugMode Then On Error Resume Next
Call RangeSave
If ierror Then Exit Sub
Call RangeCalculate(Int(5))
If ierror Then Exit Sub
End Sub

Private Sub CommandCalculateElectronRange_Click()
If Not DebugMode Then On Error Resume Next
Call RangeSave
If ierror Then Exit Sub
Call RangeCalculate(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub CommandCalculateXrayRange_Click()
If Not DebugMode Then On Error Resume Next
Call RangeSave
If ierror Then Exit Sub
Call RangeCalculate(Int(2))
If ierror Then Exit Sub
End Sub

Private Sub CommandCalculateXrayTransmission_Click()
If Not DebugMode Then On Error Resume Next
Call RangeSave
If ierror Then Exit Sub
Call RangeCalculate(Int(3))
If ierror Then Exit Sub
End Sub

Private Sub CommandCalculateXrayTransmission2_Click()
If Not DebugMode Then On Error Resume Next
Call RangeSave
If ierror Then Exit Sub
Call RangeCalculate(Int(4))
If ierror Then Exit Sub
End Sub

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormRANGE
End Sub

Private Sub CommandEnterCompositionAsFormulaString_Click()
If Not DebugMode Then On Error Resume Next
Call RangeGetComposition(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub CommandEnterCompositionAsStandard_Click()
If Not DebugMode Then On Error Resume Next
Call RangeGetComposition(Int(3))
If ierror Then Exit Sub
End Sub

Private Sub CommandEnterCompositionAsWeightString_Click()
If Not DebugMode Then On Error Resume Next
Call RangeGetComposition(Int(2))
If ierror Then Exit Sub
End Sub

Private Sub CommandHelpRange_Click()
If Not DebugMode Then On Error Resume Next
Call IOBrowseHTTP(ProbeSoftwareInternetBrowseMethod%, "https://probesoftware.com/smf/index.php?topic=86.0")
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormRANGE)
HelpContextID = IOGetHelpContextID("FormRANGE")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub ListAtomicDensities_DblClick()
If Not DebugMode Then On Error Resume Next
FormRANGE.TextDensity.Text = Format$(AllAtomicDensities!(FormRANGE.ListAtomicDensities.ListIndex + 1))
End Sub

Private Sub TextDensity_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextKev_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextThickness_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextXrayEnergy_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub
