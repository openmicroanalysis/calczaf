VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form FormZAF 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculate ZAF and Phi-Rho-Z Corrections"
   ClientHeight    =   6015
   ClientLeft      =   1740
   ClientTop       =   2955
   ClientWidth     =   9570
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
   Icon            =   "ZAF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6015
   ScaleWidth      =   9570
   Begin VB.Frame FrameElementList 
      Caption         =   "Element List (click element row to edit)"
      ForeColor       =   &H00FF0000&
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      Begin VB.CheckBox CheckPlotPhiRhoZCurves 
         Caption         =   "Plot Phi-Rho-Z Curves"
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
         Left            =   6240
         TabIndex        =   19
         Top             =   4200
         Width           =   2775
      End
      Begin MSFlexGridLib.MSFlexGrid GridElementList 
         Height          =   2895
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   360
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   5106
         _Version        =   393216
         Rows            =   73
         Cols            =   8
      End
      Begin VB.CommandButton CommandHelpCalcZAF 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Help"
         Height          =   375
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to get detailed help from our on-line user forum"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CheckBox CheckUseAllMatrixCorrections 
         Caption         =   "Use All Matrix Corrections"
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
         Left            =   6240
         TabIndex        =   16
         ToolTipText     =   "Calculate results using all available matrix corrections"
         Top             =   3960
         Width           =   2535
      End
      Begin VB.CommandButton CommandExcel 
         Caption         =   ">>Excel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4080
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Output the current data results to the currently open Excel sheet (see Output menu to open a link to Excel)"
         Top             =   3960
         Width           =   1455
      End
      Begin VB.CommandButton CommandExcelOptions 
         Caption         =   "Excel Options"
         Height          =   495
         Left            =   4080
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Specify which calculation results to output to the currently open Excel spreadsheet (see Output menu to open an Excel link)"
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton CommandCombinedConditions 
         Caption         =   "Combined Conditions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Modify the analytical conditions for each element (TKCS)"
         Top             =   5160
         Width           =   1935
      End
      Begin VB.CommandButton CommandCompositionStandard 
         Caption         =   "Enter Composition From Database"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Specify a sample composition based on a standard composition from the standard composition database"
         Top             =   4080
         Width           =   3495
      End
      Begin VB.CommandButton CommandCompositionWeight 
         Caption         =   "Enter Composition as Weight String"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Specify a sample composition based on a weight fraction formula"
         Top             =   3720
         Width           =   3495
      End
      Begin VB.CommandButton CommandCompositionAtom 
         Caption         =   "Enter Composition as Formula String"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Specify a sample composition based on an atomic fraction formula"
         Top             =   3360
         Width           =   3495
      End
      Begin VB.CommandButton CommandNext 
         BackColor       =   &H0080FFFF&
         Caption         =   "Load Next Dataset From Input File"
         Enabled         =   0   'False
         Height          =   615
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Load the next data set from the CalcZAF input file"
         Top             =   4440
         Width           =   3255
      End
      Begin VB.CommandButton CommandCopyToClipboard 
         Caption         =   "Copy Grid To Clipboard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7920
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Copy the current analysis results to the Windows clipboard"
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton CommandZAFOption 
         BackColor       =   &H0080FFFF&
         Caption         =   "Calculation Options"
         Height          =   495
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Specify different calculation options (oxygen by stoichiometry, formula, etc.)"
         Top             =   3360
         Width           =   1335
      End
      Begin VB.OptionButton OptionCalculate 
         Caption         =   "Calculate Weight Concentrations From Intensities (k-ratio)"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Calculate composition from the current sample using elemental (normalized) k-ratios (no standard is specified)"
         Top             =   5400
         Width           =   5415
      End
      Begin VB.OptionButton OptionCalculate 
         Caption         =   "Calculate Weight Concentrations From Intensities (k-raw)"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Calculate composition from the current sample using raw k-ratios and a specified standard"
         Top             =   5160
         Width           =   5415
      End
      Begin VB.OptionButton OptionCalculate 
         Caption         =   "Calculate Weight Concentrations From Intensities (counts)"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Calculate composition from the current sample using both standard and unknown intensities (in cps/nA)"
         Top             =   4920
         Width           =   5415
      End
      Begin VB.CommandButton CommandClose 
         BackColor       =   &H00C0FFC0&
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   495
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Close this window"
         Top             =   3360
         Width           =   735
      End
      Begin VB.OptionButton OptionCalculate 
         Caption         =   "Calculate Intensities From Weight Concentrations"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Calculate k-ratios from the current sample concentrations"
         Top             =   4680
         Width           =   5415
      End
      Begin VB.CommandButton CommandCalculate 
         BackColor       =   &H0080FFFF&
         Caption         =   "Calculate"
         Height          =   495
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Calculate the k-ratio or composition for the current sample"
         Top             =   3360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FormZAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2025 by John J. Donovan
Option Explicit

Private Sub CommandCalculate_Click()
' Calculate weight or intensity
If Not DebugMode Then On Error Resume Next
If FormZAF.CheckPlotPhiRhoZCurves.Value = vbChecked Then
CalculatePhiRhoZPlotCurves = True
Else
CalculatePhiRhoZPlotCurves = False
End If
If FormZAF.CheckUseAllMatrixCorrections.Value = vbChecked Then
CalculateAllMatrixCorrections = True
Call CalcZAFCalculateAll(FormZAF)
CalculateAllMatrixCorrections = False
If ierror Then Exit Sub
Else
CalculateAllMatrixCorrections = False
Call CalcZAFCalculate
If ierror Then Exit Sub
End If
End Sub

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormZAF
End Sub

Private Sub CommandCombinedConditions_Click()
' Calculate weight or intensity
If Not DebugMode Then On Error Resume Next
Call CalcZAFCombinedConditions
If ierror Then Exit Sub
End Sub

Private Sub CommandCompositionAtom_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFGetComposition(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub CommandCompositionStandard_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFGetComposition(Int(3))
If ierror Then Exit Sub
End Sub

Private Sub CommandCompositionWeight_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFGetComposition(Int(2))
If ierror Then Exit Sub
End Sub

Private Sub CommandCopyToClipboard_Click()
If Not DebugMode Then On Error Resume Next
Call MiscCopyGrid2(FormZAF.GridElementList)
If ierror Then Exit Sub
End Sub

Private Sub CommandExcel_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFGetExcel
If ierror Then Exit Sub
End Sub

Private Sub CommandExcelOptions_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFExcelOptionsLoad
If ierror Then Exit Sub
FormEXCELOPTIONS.Show vbModal
End Sub

Private Sub CommandHelpCalcZAF_Click()
If Not DebugMode Then On Error Resume Next
Call IOBrowseHTTP(ProbeSoftwareInternetBrowseMethod%, "https://smf.probesoftware.com/index.php?topic=81.0")
If ierror Then Exit Sub
End Sub

Private Sub CommandNext_Click()
If Not DebugMode Then On Error Resume Next
Call CalcZAFImportNext
If ierror Then Exit Sub
End Sub

Private Sub CommandZAFOption_Click()
If Not DebugMode Then On Error Resume Next

' Get options
Call CalcZAFOption
If ierror Then Exit Sub

' Sort elements
Call CalcZAFSave
If ierror Then Exit Sub

' Update element list
Call CalcZAFLoadList
If ierror Then Exit Sub

End Sub

Private Sub Form_Activate()
If Not DebugMode Then On Error Resume Next
If ExcelSheetIsOpen() Then
FormZAF.CommandExcel.Enabled = True
FormZAF.CommandExcelOptions.Enabled = True
Else
FormZAF.CommandExcel.Enabled = False
FormZAF.CommandExcelOptions.Enabled = False
End If
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormZAF)
HelpContextID = IOGetHelpContextID("FormZAF")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub GridElementList_Click()
Dim elementrow As Integer

' Determine current element row number
elementrow% = FormZAF.GridElementList.row
If elementrow% < 1 Or elementrow% > MAXCHAN% Then Exit Sub

' Load element row for FormZAFELM
Call CalcZAFElementLoad(elementrow%)
If ierror Then Exit Sub

' Sort elements
Call CalcZAFSave
If ierror Then Exit Sub

' Update element list
Call CalcZAFLoadList
If ierror Then Exit Sub
End Sub

Private Sub OptionCalculate_Click(Index As Integer)
If Not DebugMode Then On Error Resume Next
Call CalcZAFGetMode(Index%)
If ierror Then Exit Sub
End Sub

