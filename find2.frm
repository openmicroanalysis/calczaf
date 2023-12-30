VERSION 5.00
Begin VB.Form FormFIND2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search To Find Standard Name"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CommandNext 
      BackColor       =   &H0080FFFF&
      Caption         =   "Next Match"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton CommandClose 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Close"
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox TextStandardString 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Type the first few characters of the standard name and the program will go to that standard"
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Type Standard Name to Find"
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
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "FormFIND2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (c) Copyright 1995-2024 by John J. Donovan
Option Explicit

Private Sub CommandClose_Click()
If Not DebugMode Then On Error Resume Next
Unload FormFIND2
ierror = True
End Sub

Private Sub CommandNext_Click()
If Not DebugMode Then On Error Resume Next
Call StandardFindString(Int(1), FormFIND2.TextStandardString, FormMAIN.ListAvailableStandards)
End Sub

Private Sub Form_Activate()
If Not DebugMode Then On Error Resume Next
FormFIND2.TextStandardString.SetFocus
End Sub

Private Sub Form_Load()
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(2), MDBUserName$, Me)
Call MiscLoadIcon(FormFIND2)
HelpContextID = IOGetHelpContextID("FormFIND2")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode Then On Error Resume Next
Call InitWindow(Int(1), MDBUserName$, Me)
End Sub

Private Sub TextStandardString_Change()
If Not DebugMode Then On Error Resume Next
Call StandardFindString(Int(0), FormFIND2.TextStandardString, FormMAIN.ListAvailableStandards)
End Sub

Private Sub TextStandardString_GotFocus()
If Not DebugMode Then On Error Resume Next
Call MiscSelectText(Screen.ActiveForm.ActiveControl)
End Sub
