VERSION 5.00
Begin VB.Form FormWEIGHT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Weight Percent Entry"
   ClientHeight    =   1545
   ClientLeft      =   750
   ClientTop       =   4290
   ClientWidth     =   9225
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
   ScaleHeight     =   1545
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CommandCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter Weight Percent String For Match"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7575
      Begin VB.TextBox TextWeightPercentString 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         ToolTipText     =   "Enter the unknown composition as a weight percent string (spaces and parentheses are ok)"
         Top             =   480
         Width           =   6015
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "For example : ""mg57ca3o40"", or ""fe55si14o31"""
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
         TabIndex        =   5
         Top             =   840
         Width           =   7215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Wt. % String"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FormWEIGHT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2018 by John J. Donovan
Option Explicit

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormWEIGHT
icancel = True
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call FormulaSaveWeight
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
icancel = False
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormWEIGHT)
HelpContextID = IOGetHelpContextID("FormWEIGHT")
Call FormulaLoad(Int(1))
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub TextWeightPercentString_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

