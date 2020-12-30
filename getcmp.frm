VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FormGETCMP 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Standard Composition"
   ClientHeight    =   9705
   ClientLeft      =   1545
   ClientTop       =   1245
   ClientWidth     =   9840
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9705
   ScaleWidth      =   9840
   Begin VB.CommandButton CommandMemoText 
      Caption         =   "Memo Text"
      Height          =   375
      Left            =   8040
      TabIndex        =   49
      ToolTipText     =   "Edit memo text for additional standard information"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton CommandEnterAtomFormula 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "Enter Atom Formula Composition"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Enter the standard composition as a formula (e.g., Fe2SiO4 or MgCaSi2O6)"
      Top             =   6600
      Width           =   4695
   End
   Begin VB.ComboBox ComboFormula 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   47
      TabStop         =   0   'False
      ToolTipText     =   $"GETCMP.frx":0000
      Top             =   7080
      Width           =   735
   End
   Begin VB.TextBox TextFormula 
      Height          =   285
      Left            =   3000
      TabIndex        =   46
      TabStop         =   0   'False
      ToolTipText     =   "Number of atoms for the formula basis"
      Top             =   7080
      Width           =   615
   End
   Begin VB.CheckBox CheckFormula 
      Caption         =   "Calculate Formula Based On"
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
      TabIndex        =   45
      TabStop         =   0   'False
      ToolTipText     =   "Perform a formula atom calculation"
      Top             =   7080
      Width           =   2655
   End
   Begin VB.TextBox TextMaterialType 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Enter an optional material type description for this standard (for filtering the standard list in Add/Remove Standards)"
      Top             =   7680
      Width           =   5175
   End
   Begin VB.CommandButton CommandDeleteCLSpectrum 
      Caption         =   "Delete CL Spectrum"
      Height          =   495
      Left            =   5040
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   9120
      Width           =   1935
   End
   Begin VB.CommandButton CommandDeleteEDSSpectrum 
      Caption         =   "Delete EDS Spectrum"
      Height          =   495
      Left            =   2880
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   9120
      Width           =   2055
   End
   Begin VB.CommandButton CommandDisplayCLSpectrum 
      BackColor       =   &H0080FFFF&
      Caption         =   "Display CL Spectrum"
      Height          =   495
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CommandButton CommandDisplayEDSSpectrum 
      BackColor       =   &H0080FFFF&
      Caption         =   "Display EDS Spectrum"
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   8640
      Width           =   2055
   End
   Begin VB.ListBox ListCLSpectra 
      Height          =   1425
      Left            =   7080
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   8160
      Width           =   2655
   End
   Begin VB.ListBox ListEDSSpectra 
      Height          =   1425
      Left            =   120
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   8160
      Width           =   2655
   End
   Begin VB.CommandButton CommandImportCLSpectra 
      Caption         =   "Import CL Spectrum"
      Height          =   495
      Left            =   5040
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton CommandImportEDSSpectra 
      Caption         =   "Import EDS Spectrum"
      Height          =   495
      Left            =   2880
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   8160
      Width           =   2055
   End
   Begin VB.CommandButton CommandCalculateDensity 
      Caption         =   "Calculate Density"
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
      Height          =   375
      Left            =   8040
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox TextDensity 
      Height          =   285
      Left            =   8040
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton CommandUpdateExcess 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "Update Excess"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Update the composition for the entered excess oxygen"
      Top             =   7560
      Width           =   1575
   End
   Begin VB.TextBox TextExcessOxygen 
      Height          =   285
      Left            =   8520
      TabIndex        =   4
      ToolTipText     =   "Enter any excess oxygen here (e.g., from Fe2O3) and then click the Update Excess button"
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Frame Frame5 
      Caption         =   "Sample Number, Name and Description"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   7575
      Begin VB.CommandButton CommandAddCR 
         Caption         =   "Add <cr>"
         Height          =   255
         Left            =   6480
         TabIndex        =   42
         ToolTipText     =   "Add a carriage return to the text description (place cursor and hit Add <cr>)"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TextNumber 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox TextDescription 
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   840
         Width           =   7335
      End
      Begin VB.TextBox TextName 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   5775
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Enter Composition In"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   240
      TabIndex        =   10
      Top             =   5640
      Width           =   2295
      Begin VB.OptionButton OptionEnterElemental 
         Caption         =   "Elemental Percent"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Click this option to enter the standard composition in elemental weight percents"
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton OptionEnterOxide 
         Caption         =   "Oxide Percent"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Click this option to enter the standard composition in oxide weight percents"
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Display Composition As"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   2760
      TabIndex        =   9
      Top             =   5640
      Width           =   2415
      Begin VB.OptionButton OptionNotDisplayAsOxide 
         Caption         =   "Elemental Standard"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Click this option to display the standard composition in elemental weight percents"
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton OptionDisplayAsOxide 
         Caption         =   "Oxide Standard"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Click this option to display the standard composition in oxide weight percents"
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton CommandCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7920
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton CommandOK 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Click Element Row to Edit Element Composition and/or Cations (click empty row to add)"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   9615
      Begin MSFlexGridLib.MSFlexGrid GridElementList 
         Height          =   2415
         Left            =   120
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   360
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   4260
         _Version        =   393216
         Rows            =   73
         Cols            =   8
      End
   End
   Begin VB.Line Line2 
      X1              =   5400
      X2              =   9600
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label Label11 
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
      Left            =   3600
      TabIndex        =   48
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Optional Material Type (e.g., silicate, alloy, feldspar, etc.)"
      Height          =   255
      Left            =   240
      TabIndex        =   44
      Top             =   7440
      Width           =   5175
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9720
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Density (gm/cm3"
      Height          =   255
      Left            =   8040
      TabIndex        =   32
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label LabelHalogenCorrectedOxygen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8520
      TabIndex        =   31
      ToolTipText     =   "Total oxygen minus oxygen equivalent from halogens"
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Halogen Corrected Oxygen (wt.%)"
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
      Left            =   5400
      TabIndex        =   30
      Top             =   7200
      Width           =   3135
   End
   Begin VB.Label LabelOxygenFromHalogens 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8520
      TabIndex        =   29
      ToolTipText     =   "Total oxygen equivalent from halogens (F, Cl, Br, I)"
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Total Oxygen from Halogens (wt.%)"
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
      Left            =   5400
      TabIndex        =   28
      Top             =   6960
      Width           =   3135
   End
   Begin VB.Label LabelOxygenFromCations 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8520
      TabIndex        =   27
      ToolTipText     =   "Total oxygen calculated from oxide stoichiometry (enter as elemental oxygen when entering composition in oxide percents)"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      Caption         =   "Total Oxygen From Cations (wt.%)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   26
      Top             =   6600
      Width           =   3135
   End
   Begin VB.Label LabelAtomic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8520
      TabIndex        =   25
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label LabelOxide 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6960
      TabIndex        =   24
      ToolTipText     =   "Oxide Total (sum)"
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label LabelElemental 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   23
      ToolTipText     =   "Elemental total (sum)"
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Excess Oxygen"
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
      Left            =   7200
      TabIndex        =   20
      Top             =   7560
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Atomic"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8520
      TabIndex        =   19
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Oxide"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6960
      TabIndex        =   18
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Elemental"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   17
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Current Column Totals"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   16
      Top             =   5640
      Width           =   4215
   End
End
Attribute VB_Name = "FormGETCMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2021 by John J. Donovan
Option Explicit

Private Sub CheckFormula_Click()
If Not DebugMode Then On Error Resume Next
If FormGETCMP.CheckFormula.Value = vbChecked Then
FormGETCMP.TextFormula.Enabled = True
FormGETCMP.ComboFormula.Enabled = True
Else
FormGETCMP.TextFormula.Enabled = False
FormGETCMP.ComboFormula.Enabled = False
End If
End Sub

Private Sub CommandAddCR_Click()
If Not DebugMode Then On Error Resume Next
Call MiscAddCRToText(FormGETCMP.TextDescription)
If ierror Then Exit Sub
End Sub

Private Sub CommandCalculateDensity_Click()
If Not DebugMode Then On Error Resume Next
Call GetCmpCalculateDensity
If ierror Then Exit Sub
End Sub

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormGETCMP
ierror = True
End Sub

Private Sub CommandDeleteCLSpectrum_Click()
If Not DebugMode Then On Error Resume Next
Call GetCmpDeleteSpectrum(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub CommandDeleteEDSSpectrum_Click()
If Not DebugMode Then On Error Resume Next
Call GetCmpDeleteSpectrum(Int(0))
If ierror Then Exit Sub
End Sub

Private Sub CommandDisplayCLSpectrum_Click()
If Not DebugMode Then On Error Resume Next
Call GetCmpDisplaySpectrum(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub CommandDisplayEDSSpectrum_Click()
If Not DebugMode Then On Error Resume Next
Call GetCmpDisplaySpectrum(Int(0))
If ierror Then Exit Sub
End Sub

Private Sub CommandEnterAtomFormula_Click()
If Not DebugMode Then On Error Resume Next
' Load FORMULA form and get user formula
Call GetCmpLoadFormula
If ierror Then Exit Sub
End Sub

Private Sub CommandImportCLSpectra_Click()
If Not DebugMode Then On Error Resume Next
Call GetCmpImportSpectrum(Int(1), FormGETCMP)
If ierror Then Exit Sub
End Sub

Private Sub CommandImportEDSSpectra_Click()
If Not DebugMode Then On Error Resume Next
Call GetCmpImportSpectrum(Int(0), FormGETCMP)
If ierror Then Exit Sub
End Sub

Private Sub CommandMemoText_Click()
If Not DebugMode Then On Error Resume Next
Call GetCmpMemoTextLoad
If ierror Then Exit Sub
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call GetCmpSaveAll
If ierror Then Exit Sub
End Sub

Private Sub CommandUpdateExcess_Click()
If Not DebugMode Then On Error Resume Next
Call GetCmpChangedExcess
If ierror Then Exit Sub
End Sub

Private Sub Form_Activate()
If Not DebugMode Then On Error Resume Next
FormGETCMP.TextName.SetFocus
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormGETCMP)
HelpContextID = IOGetHelpContextID("FormGETCMP")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub GridElementList_Click()
If Not DebugMode Then On Error Resume Next
Dim elementrow As Integer

' Determine current element row number
elementrow% = FormGETCMP.GridElementList.row
If elementrow% < 1 Or elementrow% > MAXCHAN% Then Exit Sub

' Load element row for FormSETCMP
Call GetCmpSetCmpLoadElement(elementrow%)
If ierror Then Exit Sub
End Sub

Private Sub OptionDisplayAsOxide_Click()
If Not DebugMode Then On Error Resume Next
' Reload composition
Call GetCmpSave
If ierror Then Exit Sub
' Reload the entire grid
Call GetCmpLoadGrid
If ierror Then Exit Sub
FormGETCMP.CommandUpdateExcess.Enabled = True
End Sub

Private Sub OptionEnterElemental_Click()
If Not DebugMode Then On Error Resume Next
FormGETCMP.LabelOxygenFromCations.ForeColor = vbBlack
End Sub

Private Sub OptionEnterOxide_Click()
If Not DebugMode Then On Error Resume Next
FormGETCMP.LabelOxygenFromCations.ForeColor = vbRed
FormGETCMP.OptionDisplayAsOxide.Value = True
End Sub

Private Sub OptionNotDisplayAsOxide_Click()
If Not DebugMode Then On Error Resume Next
' Reload composition
Call GetCmpSave
If ierror Then Exit Sub
' Reload the entire grid
Call GetCmpLoadGrid
If ierror Then Exit Sub
FormGETCMP.CommandUpdateExcess.Enabled = False
End Sub

Private Sub TextDensity_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextDescription_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextExcessOxygen_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextFormula_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMaterialType_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextName_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextNumber_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

