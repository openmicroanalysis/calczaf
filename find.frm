VERSION 5.00
Begin VB.Form FormFIND 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Standards"
   ClientHeight    =   6465
   ClientLeft      =   3240
   ClientTop       =   1170
   ClientWidth     =   6345
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6465
   ScaleWidth      =   6345
   Begin VB.CommandButton CommandFilterStandardList 
      Caption         =   "Filter Std List"
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
      Left            =   4200
      TabIndex        =   14
      ToolTipText     =   "Filter Fin Standards list using above criteria"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.OptionButton OptionGreaterOrLess 
      Caption         =   "> 50% Absorption"
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
      Left            =   4320
      TabIndex        =   13
      ToolTipText     =   "Find all matrix corrections at least 10% greater than 1.0 (large absorption)"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.OptionButton OptionGreaterOrLess 
      Caption         =   "> 5% Fluorescence"
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
      Left            =   4320
      TabIndex        =   12
      ToolTipText     =   "Find all matrix corrections 5% or less than 1.0 (large fluorescence)"
      Top             =   1080
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Standards Found (double-click to see composition data)"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   4215
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   6135
      Begin VB.CommandButton CommandSaveStandardsToClipboard 
         Caption         =   "Save Standards Found to Clipboard"
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
         Left            =   720
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Save standards found to clipboard"
         Top             =   3720
         Width           =   4695
      End
      Begin VB.ListBox ListStandards 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3210
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   5895
      End
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H00C0FFC0&
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter Search Element and Range"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton CommandFindStandards 
         BackColor       =   &H0080FFFF&
         Caption         =   "Find Standards"
         Default         =   -1  'True
         Height          =   495
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Find all standards containing the specified element in the given range"
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox TextHigh 
         Height          =   285
         Left            =   2520
         TabIndex        =   3
         ToolTipText     =   "High limit in weight percent"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TextLow 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         ToolTipText     =   "Low limit in weight percent"
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox ComboElement 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Select element to find"
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Element"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Low Limit"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "High Limit"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FormFIND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2019 by John J. Donovan
Option Explicit

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormFIND
End Sub

Private Sub CommandFilterStandardList_Click()
If Not DebugMode Then On Error Resume Next
Screen.MousePointer = vbHourglass
Call FindStandardsFilter
Screen.MousePointer = vbDefault
If ierror Then Exit Sub
End Sub

Private Sub CommandFindStandards_Click()
If Not DebugMode Then On Error Resume Next
Call FindStandards(FormFIND.ListStandards)
If ierror Then Exit Sub
End Sub

Private Sub CommandSaveStandardsToClipboard_Click()
If Not DebugMode Then On Error Resume Next
Call MiscCopyList(Int(1), FormFIND.ListStandards)
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormFIND)
HelpContextID = IOGetHelpContextID("FormFIND")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub ListStandards_DblClick()
If Not DebugMode Then On Error Resume Next
Dim number As Integer

' Get standard from listbox
If FormFIND.ListStandards.ListIndex < 0 Then Exit Sub
number% = FormFIND.ListStandards.ItemData(FormFIND.ListStandards.ListIndex)

' Recalculate and display standard data
If number% > 0 Then Call StanFormCalculate(number%, Int(0))
If ierror Then Exit Sub

End Sub

Private Sub TextHigh_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextLow_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

