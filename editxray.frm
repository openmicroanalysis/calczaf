VERSION 5.00
Begin VB.Form FormEDITXRAY 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit X-Ray|X-Edge|X-Flur Tables"
   ClientHeight    =   1305
   ClientLeft      =   2670
   ClientTop       =   2835
   ClientWidth     =   6975
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
   ScaleHeight     =   1305
   ScaleWidth      =   6975
   Begin VB.Frame Frame1 
      Caption         =   "Edit"
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      Begin VB.TextBox TextDataValue 
         Height          =   285
         Left            =   4320
         TabIndex        =   8
         ToolTipText     =   "Enter the new or modified x-ray energy (in eV) or fluorescent yield (in fraction)"
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox ComboXray 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2040
         TabIndex        =   4
         Text            =   "ComboXray"
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox ComboElement 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Text            =   "ComboElement"
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Data Value"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "X-Ray"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Element"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "FormEDITXRAY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2018 by John J. Donovan
Option Explicit

Private Sub ComboElement_Change()
If Not DebugMode Then On Error Resume Next
Call EditUpdateDataValue
If ierror Then Exit Sub
End Sub

Private Sub ComboElement_Click()
If Not DebugMode Then On Error Resume Next
Call EditUpdateDataValue
If ierror Then Exit Sub
End Sub

Private Sub ComboXray_Change()
If Not DebugMode Then On Error Resume Next
Call EditUpdateDataValue
If ierror Then Exit Sub
End Sub

Private Sub ComboXray_Click()
If Not DebugMode Then On Error Resume Next
Call EditUpdateDataValue
If ierror Then Exit Sub
End Sub

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormEDITXRAY
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call EditXraySave
If ierror Then Exit Sub
Unload FormEDITXRAY
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call MiscCenterForm(FormEDITXRAY)
Call MiscLoadIcon(FormEDITXRAY)
HelpContextID = IOGetHelpContextID("FormEDITXRAY")
End Sub

Private Sub TextDataValue_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

