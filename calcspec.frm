VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormCALCSPEC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculate Spectrometer Position"
   ClientHeight    =   3840
   ClientLeft      =   1770
   ClientTop       =   1230
   ClientWidth     =   11280
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
   ScaleHeight     =   3840
   ScaleWidth      =   11280
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   6360
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Adjust the analyzing crystal refractive index"
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.TextBox TextKIndex 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5160
      TabIndex        =   20
      ToolTipText     =   $"CalcSpec.frx":0000
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton CommandHelp 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Help"
      Height          =   375
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Spectrometer Types"
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   6495
      Begin VB.OptionButton OptionInstrument 
         Caption         =   "Cameca 180mm Spectrometer (same reading as 160mm spectro!)"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1080
         Width           =   6255
      End
      Begin VB.OptionButton OptionInstrument 
         Caption         =   "JEOL 100mm Spectrometer (same reading as 140mm spectro!)"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   840
         Width           =   6255
      End
      Begin VB.OptionButton OptionInstrument 
         Caption         =   "JEOL 140mm Spectrometer"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   360
         Width           =   4095
      End
      Begin VB.OptionButton OptionInstrument 
         Caption         =   "Cameca 160mm Spectrometer"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   600
         Width           =   4095
      End
   End
   Begin VB.CheckBox CheckUseRefractiveIndex 
      Caption         =   "Use Refractive Index For Bragg Calculation"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   840
      Width           =   4095
   End
   Begin VB.CommandButton CommandCalculate 
      BackColor       =   &H0080FFFF&
      Caption         =   "Calculate Spectrometer Position"
      Height          =   1335
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox TextCalcSpec 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   3000
      Width           =   8415
   End
   Begin VB.ComboBox ComboCrystal 
      Height          =   315
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Select the bragg reflection order for the indicated x-ray line"
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox ComboElement 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Enter the analyzed or specified element"
      Top             =   360
      Width           =   1335
   End
   Begin VB.ComboBox ComboXRay 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "Enter the x-ray line for analyzed elements or no x-ray line for specified elements"
      Top             =   360
      Width           =   1335
   End
   Begin VB.ComboBox ComboBraggOrder 
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Select the bragg reflection order for the indicated x-ray line"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   1575
   End
   Begin VB.OLE OLE3 
      BackStyle       =   0  'Transparent
      Class           =   "Equation.3"
      Enabled         =   0   'False
      Height          =   600
      Left            =   6810
      OleObjectBlob   =   "CalcSpec.frx":00C2
      SizeMode        =   1  'Stretch
      TabIndex        =   27
      Top             =   600
      Width           =   4395
   End
   Begin VB.OLE OLE1 
      BackStyle       =   0  'Transparent
      Class           =   "Equation.3"
      Enabled         =   0   'False
      Height          =   855
      Left            =   8610
      OleObjectBlob   =   "CalcSpec.frx":12DA
      SizeMode        =   1  'Stretch
      TabIndex        =   26
      Top             =   2880
      Width           =   2535
   End
   Begin VB.OLE OLE2 
      BackStyle       =   0  'Transparent
      Class           =   "Equation.3"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   8640
      OleObjectBlob   =   "CalcSpec.frx":22F2
      SizeMode        =   1  'Stretch
      TabIndex        =   25
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "K-index"
      Height          =   255
      Left            =   4440
      TabIndex        =   21
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Crystal"
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Element"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "X-Ray Line"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Bragg Order"
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label LabelQuickScanHiPeaks 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label LabelQuickScanLoPeaks 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6000
      TabIndex        =   5
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label LabelWaveScanLoPeaks 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label LabelWaveScanHiPeaks 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "FormCALCSPEC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2017 by John J. Donovan
Option Explicit

Private Sub CheckUseRefractiveIndex_Click()
If Not DebugMode Then On Error Resume Next
If FormCALCSPEC.CheckUseRefractiveIndex.Value = vbChecked Then
FormCALCSPEC.TextKIndex.Enabled = True
FormCALCSPEC.UpDown1.Enabled = True
Else
FormCALCSPEC.TextKIndex.Enabled = False
FormCALCSPEC.UpDown1.Enabled = False
End If
End Sub

Private Sub ComboCrystal_Change()
If Not DebugMode Then On Error Resume Next
If FormCALCSPEC.ComboCrystal.ListIndex > -1 Then
FormCALCSPEC.TextKIndex.Text = Format$(AllCrystalKs!(FormCALCSPEC.ComboCrystal.ListIndex + 1), f96$)
End If
End Sub

Private Sub ComboCrystal_Click()
If Not DebugMode Then On Error Resume Next
If FormCALCSPEC.ComboCrystal.ListIndex > -1 Then
FormCALCSPEC.TextKIndex.Text = Format$(AllCrystalKs!(FormCALCSPEC.ComboCrystal.ListIndex + 1), f96$)
End If
End Sub

Private Sub CommandCalculate_Click()
If Not DebugMode Then On Error Resume Next
Call CalcSpecCalculate
If ierror Then Exit Sub
End Sub

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Call CalcSpecSave
If ierror Then Exit Sub
Unload FormCALCSPEC
End Sub

Private Sub CommandHelp_Click()
If Not DebugMode Then On Error Resume Next
Call IOBrowseHTTP(ProbeSoftwareInternetBrowseMethod%, "http://probesoftware.com/smf/index.php?topic=375.0")
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
icancelload = False
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormCALCSPEC)
HelpContextID = IOGetHelpContextID("FormCALCSPEC")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub TextKIndex_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub UpDown1_DownClick()
If Not DebugMode Then On Error Resume Next
AllCrystalKs!(FormCALCSPEC.ComboCrystal.ListIndex + 1) = Val(FormCALCSPEC.TextKIndex.Text)
AllCrystalKs!(FormCALCSPEC.ComboCrystal.ListIndex + 1) = AllCrystalKs!(FormCALCSPEC.ComboCrystal.ListIndex + 1) * 0.8
FormCALCSPEC.TextKIndex.Text = Format$(AllCrystalKs!(FormCALCSPEC.ComboCrystal.ListIndex + 1), f96$)
End Sub

Private Sub UpDown1_UpClick()
If Not DebugMode Then On Error Resume Next
AllCrystalKs!(FormCALCSPEC.ComboCrystal.ListIndex + 1) = Val(FormCALCSPEC.TextKIndex.Text)
AllCrystalKs!(FormCALCSPEC.ComboCrystal.ListIndex + 1) = AllCrystalKs!(FormCALCSPEC.ComboCrystal.ListIndex + 1) * 1.2
FormCALCSPEC.TextKIndex.Text = Format$(AllCrystalKs!(FormCALCSPEC.ComboCrystal.ListIndex + 1), f96$)
End Sub
