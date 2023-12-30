VERSION 5.00
Begin VB.Form FormFORMULA 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   2130
   ClientLeft      =   2055
   ClientTop       =   1950
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
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2130
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CommandCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton CommandOK 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter Formula String For"
      ClipControls    =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6135
      Begin VB.TextBox TextFormulaString 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Enter a formula string (spaces and parentheses are ok)"
         Top             =   360
         Width           =   5895
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "For example : ""fe2sio4"", ""h2o"", ""ch2ch3oh"" or ""ca2mg5si8o22(oh)2"""
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
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   5895
      End
   End
End
Attribute VB_Name = "FormFORMULA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2024 by John J. Donovan
Option Explicit

Private Sub CommandCancel_Click()
If Not DebugMode Then On Error Resume Next
Unload FormFORMULA
icancel = True
End Sub

Private Sub CommandOK_Click()
If Not DebugMode Then On Error Resume Next
Call FormulaSaveFormula
If ierror Then Exit Sub
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
icancel = False
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormFORMULA)
HelpContextID = IOGetHelpContextID("FormFORMULA")
Call FormulaLoad(Int(0))
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub TextFormulaString_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub

