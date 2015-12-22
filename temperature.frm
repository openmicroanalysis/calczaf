VERSION 5.00
Begin VB.Form FormTEMPERATURE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculate Sample Temperature Rise From Electron Dose"
   ClientHeight    =   3480
   ClientLeft      =   1245
   ClientTop       =   5580
   ClientWidth     =   10575
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
   ScaleHeight     =   3480
   ScaleWidth      =   10575
   Begin VB.Frame Frame2 
      Caption         =   "Beam Conditions"
      ForeColor       =   &H00FF0000&
      Height          =   3255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6015
      Begin VB.ComboBox ComboThermalConductivity 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Selected thermal conductivities for a few materials"
         Top             =   1440
         Width           =   3855
      End
      Begin VB.TextBox TextBeamCurrent 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         ToolTipText     =   "Enter the electron beam current in nano-amps"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox TextThermalConductivity 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Enter the thermal conductivity (in watts/cmK)"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080FFFF&
         Caption         =   "Calculate Temperature Rise"
         Height          =   615
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Calculate the deltaT, temperature rise in the material under the given beam conditions"
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox TextBeamSize 
         Height          =   285
         Left            =   3840
         TabIndex        =   2
         ToolTipText     =   "Enter the electron beam diameter in microns"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TextBeamEnergy 
         Height          =   285
         Left            =   240
         TabIndex        =   0
         ToolTipText     =   "Enter the primary beam energy in kilovolts"
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Thermal Conductivities (W/cmK)"
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
         Left            =   1800
         TabIndex        =   14
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Beam Current (nA)"
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
         Left            =   1800
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Conductivity"
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
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label LabelTemperatureRise 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   2640
         Width           =   5055
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Beam Diameter (um)"
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
         Left            =   3840
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00008000&
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "1 W/(mK) = 1 W/(mC) = 0.85984 kcal/(hr mC) = 0.5779 Btu/(ft hr F) "
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
      Left            =   6480
      TabIndex        =   15
      Top             =   1200
      Width           =   3735
   End
   Begin VB.OLE OLE2 
      BackStyle       =   0  'Transparent
      Class           =   "Equation.3"
      Enabled         =   0   'False
      Height          =   1455
      Left            =   6360
      OleObjectBlob   =   "Temperature.frx":0000
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1800
      Width           =   3975
   End
End
Attribute VB_Name = "FormTEMPERATURE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2015 by John J. Donovan
Option Explicit

Private Sub ComboThermalConductivity_Click()
If Not DebugMode Then On Error Resume Next
FormTEMPERATURE.TextThermalConductivity.Text = FormTEMPERATURE.ComboThermalConductivity.ItemData(FormTEMPERATURE.ComboThermalConductivity.ListIndex) / MILLIWATTSPERWATT&
End Sub

Private Sub Command1_Click()
If Not DebugMode Then On Error Resume Next
Call TemperatureSave
If ierror Then Exit Sub
Unload FormTEMPERATURE
End Sub

Private Sub Command5_Click()
If Not DebugMode Then On Error Resume Next
Call TemperatureSave
If ierror Then Exit Sub
Call TemperatureCalculate
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormTEMPERATURE)
HelpContextID = IOGetHelpContextID("FormTEMPERATURE")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub TextBeamCurrent_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextBeamEnergy_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextBeamSize_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextThermalConductivity_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub
