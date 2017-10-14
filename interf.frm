VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormINTERF 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculate Nominal Interferences"
   ClientHeight    =   6015
   ClientLeft      =   1170
   ClientTop       =   2100
   ClientWidth     =   6105
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
   ScaleHeight     =   6015
   ScaleWidth      =   6105
   Begin VB.Frame Frame2 
      Caption         =   "Interference Options"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   2415
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   5895
      Begin VB.CommandButton CommandCalculate 
         BackColor       =   &H0080FFFF&
         Caption         =   "Calculate"
         Default         =   -1  'True
         Height          =   495
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Calculate the specified interferences"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox TextPHADiscrimination 
         Height          =   285
         Left            =   4680
         TabIndex        =   2
         ToolTipText     =   "PHA discrimination factor. Use larger numbers for more high energy discrimination for fewer high order interferences."
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox TextMinimumOverlap 
         Height          =   285
         Left            =   4680
         TabIndex        =   1
         ToolTipText     =   "Minimum overlap tolerance. Use larger numbers for fewer interferences."
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox TextLiFPeakWidth 
         Height          =   285
         Left            =   4680
         TabIndex        =   0
         ToolTipText     =   "LiF peak width. Use larger numbers for more overlap."
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton OptionInterferencePeak 
         Caption         =   "Low Off Peak Interferences"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Calculate low off-peak interferences"
         Top             =   840
         Width           =   3135
      End
      Begin VB.OptionButton OptionInterferencePeak 
         Caption         =   "High Off Peak Interferences"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Calculate high off-peak interferences"
         Top             =   600
         Width           =   3255
      End
      Begin VB.OptionButton OptionInterferencePeak 
         Caption         =   "On Peak Interferences"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Calculate on-peak interferences"
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         Caption         =   "PHA Discrimination for High Orders (1 to 100)"
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
         TabIndex        =   18
         Top             =   2040
         Width           =   4335
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Minimum Overlap Tolerance in Percent (0.01 to 50)"
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
         TabIndex        =   17
         ToolTipText     =   "Minimum overlap tolerance. Use smaller numbers for more interferences"
         Top             =   1680
         Width           =   4455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "LiF Peak Width (typical) in Angstroms (0.01 to 0.5)"
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
         TabIndex        =   16
         Top             =   1320
         Width           =   4455
      End
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H00C0FFC0&
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selected Sample Composition"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   3135
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4695
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   615
         Left            =   2880
         TabIndex        =   22
         ToolTipText     =   "Change the selected element"
         Top             =   1440
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1085
         _Version        =   393216
      End
      Begin VB.ComboBox ComboXray 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   21
         ToolTipText     =   "Calculate a specific x-ray line for the selected element"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox TextRangeFraction 
         Height          =   285
         Left            =   3960
         TabIndex        =   6
         ToolTipText     =   $"INTERF.frx":0000
         Top             =   2760
         Width           =   615
      End
      Begin VB.OptionButton OptionSelected 
         Caption         =   "Selected Element"
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Calculate interferences for only the selected element in the composition"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton OptionAll 
         Caption         =   "All Elements"
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Cac=lculate interferences for all elements in the composition"
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox ComboElement 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Select a specific element to calculate interferences for"
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton CommandEnterUnknown 
         Caption         =   "Enter Unknown"
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Default is the currebtly selected standard composition, or enter a composition using a weight percent string"
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton CommandLoadXrayDatabase 
         Caption         =   "Load Xray Database"
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
         Left            =   2520
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Load the NIST x-ray database to search for specific interferences"
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox TextComposition 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Range Fraction"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Top             =   2760
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FormINTERF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2017 by John J. Donovan
Option Explicit

Private Sub ComboElement_Change()
If Not DebugMode Then On Error Resume Next
Call InterfUpdateElement
If ierror Then Exit Sub
End Sub

Private Sub ComboElement_Click()
If Not DebugMode Then On Error Resume Next
Call InterfUpdateElement
If ierror Then Exit Sub
End Sub

Private Sub CommandCalculate_Click()
If Not DebugMode Then On Error Resume Next
Call InterfSave
If ierror Then Exit Sub
End Sub

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormINTERF
End Sub

Private Sub CommandEnterUnknown_Click()
If Not DebugMode Then On Error Resume Next
Call InterfLoadWeight
If ierror Then Exit Sub
End Sub

Private Sub CommandLoadXrayDatabase_Click()
If Not DebugMode Then On Error Resume Next
' Load element xray range
Call InterfLoadXrayDatabase
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormINTERF)
HelpContextID = IOGetHelpContextID("FormINTERF")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub OptionAll_Click()
If Not DebugMode Then On Error Resume Next
FormINTERF.UpDown1.Enabled = False
FormINTERF.ComboElement.Enabled = False
FormINTERF.ComboXray.Enabled = False
FormINTERF.CommandLoadXrayDatabase.Enabled = False
End Sub

Private Sub OptionSelected_Click()
If Not DebugMode Then On Error Resume Next
FormINTERF.UpDown1.Enabled = True
FormINTERF.ComboElement.Enabled = True
FormINTERF.ComboXray.Enabled = True
FormINTERF.CommandLoadXrayDatabase.Enabled = True
End Sub

Private Sub TextLiFPeakWidth_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextMinimumOverlap_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextPHADiscrimination_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextRangeFraction_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub UpDown1_DownClick()
If Not DebugMode Then On Error Resume Next
Dim ip As Integer

' Increment element
If FormINTERF.ComboElement.ListCount < 0 Then Exit Sub
ip% = FormINTERF.ComboElement.ListIndex + 1
If ip% > FormINTERF.ComboElement.ListCount - 1 Then ip% = 0

' Change list element
FormINTERF.ComboElement.ListIndex = ip%
FormINTERF.ComboElement.Refresh
End Sub

Private Sub UpDown1_UpClick()
If Not DebugMode Then On Error Resume Next
Dim ip As Integer

' Increment element
If FormINTERF.ComboElement.ListCount < 0 Then Exit Sub
ip% = FormINTERF.ComboElement.ListIndex - 1
If ip% < 0 Then ip% = FormINTERF.ComboElement.ListCount - 1

' Change list element
FormINTERF.ComboElement.ListIndex = ip%
FormINTERF.ComboElement.Refresh
End Sub
