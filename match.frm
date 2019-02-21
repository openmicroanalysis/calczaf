VERSION 5.00
Begin VB.Form FormMATCH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Match Standards"
   ClientHeight    =   6960
   ClientLeft      =   3015
   ClientTop       =   2145
   ClientWidth     =   5880
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
   ScaleHeight     =   6960
   ScaleWidth      =   5880
   Begin VB.TextBox TextMinimumVector 
      Height          =   285
      Left            =   4680
      TabIndex        =   9
      ToolTipText     =   "Enter the minimum vector fit for a match"
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton CommandChange 
      BackColor       =   &H0080FFFF&
      Caption         =   "Change Match Database"
      Height          =   735
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Change the match database to another standard database"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input Composition"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton CommandMatchStandards 
         BackColor       =   &H0080FFFF&
         Caption         =   "Match Standards"
         Default         =   -1  'True
         Height          =   495
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Match the current composition to the match database"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton CommandEnterUnknown 
         Caption         =   "Enter Unknown"
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Enter an unknown composition to match to"
         Top             =   240
         Width           =   1695
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
         Height          =   1935
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label LabelUnknown 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Left            =   2400
         TabIndex        =   11
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H00C0FFC0&
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Standards Found (double-click to see composition data)"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   5655
      Begin VB.CommandButton CommandCopyStandardsFoundToClipboard 
         Caption         =   "Copy Standards Found to Clipboard"
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
         Left            =   240
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Copy the match standards to the clipboard"
         Top             =   3600
         Width           =   5175
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
         Height          =   3000
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Minimum Vector"
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
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "FormMATCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2019 by John J. Donovan
Option Explicit

Private Sub CommandChange_Click()
If Not DebugMode Then On Error Resume Next
Call MatchOpenDatabase(FormMATCH)
If ierror Then Exit Sub
End Sub

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormMATCH
End Sub

Private Sub CommandCopyStandardsFoundToClipboard_Click()
If Not DebugMode Then On Error Resume Next
Call MiscCopyList(Int(1), FormMATCH.ListStandards)
If ierror Then Exit Sub
End Sub

Private Sub CommandEnterUnknown_Click()
If Not DebugMode Then On Error Resume Next
Call MatchLoadWeight
If ierror Then Exit Sub
End Sub

Private Sub CommandMatchStandards_Click()
If Not DebugMode Then On Error Resume Next
Call IOStatusAuto(vbNullString)
Call MatchSample
Call IOStatusAuto(vbNullString)
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormMATCH)
HelpContextID = IOGetHelpContextID("FormMATCH")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub ListStandards_DblClick()
If Not DebugMode Then On Error Resume Next
Call MatchTypeStandard
If ierror Then Exit Sub
End Sub

Private Sub TextMinimumVector_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub
