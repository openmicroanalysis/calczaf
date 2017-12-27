VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormCOND2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Combined Analytical Conditions"
   ClientHeight    =   5130
   ClientLeft      =   1650
   ClientTop       =   2820
   ClientWidth     =   8985
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
   ScaleHeight     =   5130
   ScaleWidth      =   8985
   Begin VB.CommandButton CommandHelpCombinedConditions 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Help"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Click this button to get detailed help from our on-line user forum"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Modify Channel Order"
      ForeColor       =   &H00FF0000&
      Height          =   2895
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   7335
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   2415
         Left            =   6960
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Click to change the selected element channel order"
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   4260
         _Version        =   393216
      End
      Begin VB.ListBox ListElements 
         Height          =   2400
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Changing the channel order affects the order in which the combined condition elements are acquired"
         Top             =   360
         Width           =   6735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Modify (Combined) Analytical Conditions For Each Element"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7335
      Begin VB.CommandButton CommandApply 
         BackColor       =   &H0080FFFF&
         Caption         =   "Apply Conditions To Selected Element"
         Height          =   1215
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox ComboElementXraySpectrometerCrystal 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Specify the element setup"
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox TextKiloVolts 
         Height          =   285
         Left            =   3720
         TabIndex        =   1
         ToolTipText     =   "Specify the operating voltage"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox TextTakeOff 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   0
         ToolTipText     =   "Specify the take-off angle (normally fixed)"
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Select one element at a time from the list, edit the conditions and click Apply Conditions"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Element, Xray, Spectrometer, Crystal"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Kilovolts"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3720
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Take Off"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3720
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton CommandOk 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7560
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "FormCOND2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2018 by John J. Donovan
Option Explicit

Private Sub ComboElementXraySpectrometerCrystal_Change()
If Not DebugMode Then On Error Resume Next
Call Cond2LoadField
If ierror Then Exit Sub
End Sub

Private Sub ComboElementXraySpectrometerCrystal_Click()
If Not DebugMode Then On Error Resume Next
Call Cond2LoadField
If ierror Then Exit Sub
End Sub

Private Sub CommandApply_Click()
If Not DebugMode Then On Error Resume Next
Call Cond2Apply
If ierror Then Exit Sub
End Sub

Private Sub CommandHelpCombinedConditions_Click()
If Not DebugMode Then On Error Resume Next
Call IOBrowseHTTP(ProbeSoftwareInternetBrowseMethod%, "http://probesoftware.com/smf/index.php?topic=5.0")
If ierror Then Exit Sub
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
' Save default analytical conditions
Call Cond2SaveField
If ierror Then Exit Sub
Unload FormCOND2
End Sub

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormCOND2
icancelload = True
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
icancelload = False
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormCOND2)
HelpContextID = IOGetHelpContextID("FormCOND2")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub ListElements_Click()
If Not DebugMode Then On Error Resume Next
FormCOND2.ComboElementXraySpectrometerCrystal.ListIndex = FormCOND2.ListElements.ListIndex
End Sub

Private Sub TextKilovolts_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub TextTakeOff_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

Private Sub UpDown1_DownClick()
If Not DebugMode Then On Error Resume Next
Call Cond2Sort(Int(1))
If ierror Then Exit Sub
End Sub

Private Sub UpDown1_UpClick()
If Not DebugMode Then On Error Resume Next
Call Cond2Sort(Int(2))
If ierror Then Exit Sub
End Sub
