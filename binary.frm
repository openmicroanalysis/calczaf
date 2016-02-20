VERSION 5.00
Begin VB.Form FormBINARY 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Binary Calculation Options"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Z-bar Output Filters"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   27
      Top             =   5160
      Width           =   6495
      Begin VB.TextBox TextBinaryOutputMaximumZbarDiff 
         Height          =   285
         Left            =   360
         TabIndex        =   7
         ToolTipText     =   "Specify maxnimum difference in mass fraction and electron zbar for output"
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox CheckBinaryOutputMaximumZbar 
         Caption         =   "Use Maximum Mass-Electron Zbar Difference Output"
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
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   $"Binary.frx":0000
         Top             =   1080
         Width           =   5895
      End
      Begin VB.TextBox TextBinaryOutputMinimumZbarDiff 
         Height          =   285
         Left            =   360
         TabIndex        =   6
         ToolTipText     =   "Specify minimum difference in mass fraction and electron zbar for output"
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox CheckBinaryOutputMinimumZbar 
         Caption         =   "Use Minimum Mass-Electron Zbar Difference Output"
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
         TabStop         =   0   'False
         ToolTipText     =   $"Binary.frx":0089
         Top             =   360
         Width           =   5895
      End
      Begin VB.Label Label8 
         Caption         =   "Maximum Percent Difference, < ABS(MZbar-EZbar)/MZbar * 100"
         Height          =   255
         Left            =   1200
         TabIndex        =   31
         Top             =   1320
         Width           =   5175
      End
      Begin VB.Label Label3 
         Caption         =   "Minimum Percent Difference, > ABS(MZbar-EZbar)/MZbar * 100"
         Height          =   255
         Left            =   1200
         TabIndex        =   29
         Top             =   600
         Width           =   5175
      End
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Left            =   5400
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H0000C000&
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "First Approximation Options"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   17
      Top             =   7080
      Width           =   6495
      Begin VB.CheckBox CheckFirstApproximationApplyAbsorption 
         Caption         =   "Apply Absorption Correction To First Approximation Intensities"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   360
         Width           =   4935
      End
      Begin VB.CheckBox CheckFirstApproximationApplyFluorescence 
         Caption         =   "Apply Fluorescence Correction To First Approximation Intensities"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   720
         Width           =   4935
      End
      Begin VB.CheckBox CheckFirstApproximationApplyAtomicNumber 
         Caption         =   "Apply Atomic Number Correction To First Approximation Intensities"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1080
         Width           =   4935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Output Filter Options"
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
      TabIndex        =   9
      Top             =   120
      Width           =   5175
      Begin VB.TextBox TextMaximumAtomicNumberCorrection 
         Height          =   285
         Left            =   360
         TabIndex        =   5
         ToolTipText     =   "Specify maximum atomic number correction for output"
         Top             =   4320
         Width           =   735
      End
      Begin VB.CheckBox CheckUseMaximumAtomicNumberCorrectionOutput 
         Caption         =   "Use Maximum Atomic Number Correction Output"
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
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Check to limit output to only binaries with small atomic number corrections"
         Top             =   4080
         Width           =   4815
      End
      Begin VB.TextBox TextMaximumFluorescenceCorrection 
         Height          =   285
         Left            =   360
         TabIndex        =   4
         ToolTipText     =   "Specify maximum fluorescence correction for output"
         Top             =   3600
         Width           =   735
      End
      Begin VB.CheckBox CheckUseMaximumFluorescenceCorrectionOutput 
         Caption         =   "Use Maximum Fluorescence Correction Output"
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
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Check to limit output to only binaries with small fluorescence corrections"
         Top             =   3360
         Width           =   4695
      End
      Begin VB.CheckBox CheckUseMinimumAbsorptionCorrectionOutput 
         Caption         =   "Use Minimum Absorption Correction Output"
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
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Check to limit output to only binaries with large absorption corrections"
         Top             =   360
         Width           =   4815
      End
      Begin VB.CheckBox CheckUseMinimumFluorescenceCorrectionOutput 
         Caption         =   "Use Minimum Fluorescence Correction Output"
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
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Check to limit output to only binaries with large fluorescence corrections"
         Top             =   1080
         Width           =   4695
      End
      Begin VB.CheckBox CheckUseMinimumAtomicNumberCorrectionOutput 
         Caption         =   "Use Minimum Atomic Number Correction Output"
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
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Check to limit output to only binaries with large atomic number corrections"
         Top             =   1800
         Width           =   4695
      End
      Begin VB.CheckBox CheckUseMaximumAbsorptionCorrectionOutput 
         Caption         =   "Use Maximum Absorption Correction Output"
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
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Check to limit output to only binaries with small absorption corrections"
         Top             =   2640
         Width           =   4935
      End
      Begin VB.TextBox TextMinimumAbsorptionCorrection 
         Height          =   285
         Left            =   360
         TabIndex        =   0
         ToolTipText     =   "Specify minimum absorption correction for output"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TextMinimumFluorescenceCorrection 
         Height          =   285
         Left            =   360
         TabIndex        =   1
         ToolTipText     =   "Specify minimum fluorescence correction for output"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox TextMinimumAtomicNumberCorrection 
         Height          =   285
         Left            =   360
         TabIndex        =   2
         ToolTipText     =   "Specify minimum atomic number correction for output"
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox TextMaximumAbsorptionCorrection 
         Height          =   285
         Left            =   360
         TabIndex        =   3
         ToolTipText     =   "Specify maximum absorption correction for output"
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Maximum Atomic Number Correction, < ABS(Zed-1.0)"
         Height          =   255
         Left            =   1200
         TabIndex        =   26
         Top             =   4320
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Maximum Fluorescence Correction, < ABS(Flu-1.0)"
         Height          =   255
         Left            =   1200
         TabIndex        =   24
         Top             =   3600
         Width           =   3855
      End
      Begin VB.Label Label4 
         Caption         =   "Minimum Absorption Correction, > ABS(Abs-1.0)"
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label5 
         Caption         =   "Minimum Fluorescence Correction, > ABS(Flu-1.0)"
         Height          =   255
         Left            =   1200
         TabIndex        =   16
         Top             =   1320
         Width           =   3855
      End
      Begin VB.Label Label6 
         Caption         =   "Minimum Atomic Number Correction, > ABS(Zed-1.0)"
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   2040
         Width           =   3855
      End
      Begin VB.Label Label7 
         Caption         =   "Maximum Absorption Correction, < ABS(Abs-1.0)"
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   2880
         Width           =   3855
      End
   End
End
Attribute VB_Name = "FormBINARY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2016 by John J. Donovan
Option Explicit

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormBINARY
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
' Save the options
Call CalcZAFBinarySave
If ierror Then Exit Sub
Unload FormBINARY
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormBINARY)
HelpContextID = IOGetHelpContextID("FormBINARY")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub TextBinaryOutputMaximumZbarDiff_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextBinaryOutputMinimumZbarDiff_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMaximumAbsorptionCorrection_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMaximumAtomicNumberCorrection_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMaximumFluorescenceCorrection_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMinimumAbsorptionCorrection_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMinimumAtomicNumberCorrection_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMinimumFluorescenceCorrection_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

